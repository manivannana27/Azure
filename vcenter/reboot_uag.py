#!/usr/bin/env python3
"""
Reboot UAG VMs listed in Excel across three vCenters on a timed schedule.

Excel columns (header names are matched case-insensitively):
  - Server / VM / Name / Hostname  -> VM name in vCenter
  - Component / Role / Type        -> e.g. CS, UAG  (only UAG is rebooted)
  - Time / Schedule / RebootTime   -> 12, 12:10, or an Excel time
  - vCenter (optional)             -> hostname/IP if the VM should be looked up
                                     on a specific vCenter only

Credential XML is created with Export-VCenterCredential.ps1 (Get-Credential |
Export-Clixml). Usernames stored as DOMAIN\\user are converted to user@domain
(UPN) before login.

Example:
  python reboot_uag.py --excel servers.xlsx --cred vcenter_credential.xml
  python reboot_uag.py --excel servers.xlsx --cred vcenter_credential.xml --dry-run
"""

from __future__ import annotations

import argparse
import logging
import re
import ssl
import subprocess
import sys
import time
import xml.etree.ElementTree as ET
from collections import defaultdict
from datetime import datetime, time as dt_time, timedelta
from pathlib import Path
from typing import Any

# ---------------------------------------------------------------------------
# Hard-code your three vCenters here
# ---------------------------------------------------------------------------
VCENTERS = [
    "vcenter1.example.com",
    "vcenter2.example.com",
    "vcenter3.example.com",
]

# Used when CLIXML stores NETBIOS\\user (no dot in the domain part).
# DOMAIN\\jsmith  ->  jsmith@<UPN_SUFFIX>
# corp.example.com\\jsmith  ->  jsmith@corp.example.com  (suffix not used)
UPN_SUFFIX = "example.com"

COMPONENT_FILTER = "UAG"
WAIT_BETWEEN_SETS_MINUTES = 10
VCENTER_PORT = 443
# ---------------------------------------------------------------------------

try:
    from openpyxl import load_workbook
except ImportError:  # pragma: no cover
    load_workbook = None  # type: ignore[assignment]

try:
    from pyVim.connect import Disconnect, SmartConnect
    from pyVmomi import vim
except ImportError:  # pragma: no cover
    Disconnect = None  # type: ignore[assignment]
    SmartConnect = None  # type: ignore[assignment]
    vim = None  # type: ignore[assignment]


SERVER_HEADERS = ("server", "vm", "vmname", "name", "hostname", "servername", "server name")
COMPONENT_HEADERS = ("component", "role", "type", "comp")
TIME_HEADERS = ("time", "schedule", "reboottime", "reboot time", "starttime")
VCENTER_HEADERS = ("vcenter", "vc", "vcentername")

LOG = logging.getLogger("reboot_uag")


def configure_logging(verbose: bool) -> None:
    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(
        level=level,
        format="%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )


def to_upn(username: str, upn_suffix: str = UPN_SUFFIX) -> str:
    """Convert DOMAIN\\user (CLIXML) to user@domain (vCenter UPN)."""
    value = (username or "").strip().strip("'\"")
    if not value:
        raise ValueError("Credential username is empty")
    if "@" in value:
        return value
    if "\\" in value:
        domain, user = value.split("\\", 1)
        domain = domain.strip()
        user = user.strip()
        if not user:
            raise ValueError(f"Could not parse username from {username!r}")
        if "." in domain:
            return f"{user}@{domain}"
        suffix = (upn_suffix or "").lstrip("@")
        if not suffix:
            raise ValueError(
                "NETBIOS domain in credential XML requires UPN_SUFFIX "
                "(e.g. example.com) so login can use user@domain.com"
            )
        return f"{user}@{suffix}"
    if upn_suffix and "." not in value:
        return f"{value}@{upn_suffix.lstrip('@')}"
    return value


def _local_tag(tag: str) -> str:
    return tag.split("}")[-1] if tag else tag


def username_from_clixml(xml_path: Path) -> str | None:
    """Read UserName from a PowerShell Export-Clixml file without decrypting."""
    try:
        root = ET.parse(xml_path).getroot()
    except (ET.ParseError, OSError):
        return None
    for el in root.iter():
        name = el.attrib.get("N") or el.attrib.get("n")
        if name == "UserName" and (el.text or "").strip():
            return el.text.strip()
        if _local_tag(el.tag) in {"UserName", "username"} and (el.text or "").strip():
            return el.text.strip()
    return None


def _run_powershell(script: str) -> str:
    commands = (
        ["powershell", "-NoProfile", "-NonInteractive", "-Command", script],
        ["pwsh", "-NoProfile", "-NonInteractive", "-Command", script],
    )
    last_error = ""
    for cmd in commands:
        try:
            completed = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                check=False,
            )
        except FileNotFoundError:
            continue
        if completed.returncode == 0:
            return (completed.stdout or "").strip()
        last_error = (completed.stderr or completed.stdout or "").strip()
    raise RuntimeError(
        "Could not import CLIXML credentials via PowerShell. "
        "Run this script on the Windows account that created the XML. "
        f"Detail: {last_error or 'PowerShell not found'}"
    )


def load_ps_credential(xml_path: Path, upn_suffix: str = UPN_SUFFIX) -> tuple[str, str]:
    """Load username/password from Export-Clixml (DPAPI, same Windows user)."""
    path = xml_path.resolve()
    if not path.is_file():
        raise FileNotFoundError(f"Credential XML not found: {path}")

    ps_path = str(path).replace("'", "''")
    script = f"""
$ErrorActionPreference = 'Stop'
$cred = Import-Clixml -Path '{ps_path}'
Write-Output $cred.UserName
Write-Output '---PASSWORD---'
$bstr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($cred.Password)
[System.Runtime.InteropServices.Marshal]::PtrToStringAuto($bstr)
"""
    output = _run_powershell(script)
    if "---PASSWORD---" not in output:
        # Fall back to XML username + fail clearly if password missing
        xml_user = username_from_clixml(path)
        raise RuntimeError(
            "CLIXML import did not return a password. "
            f"Parsed username={xml_user!r}"
        )
    raw_user, password = output.split("---PASSWORD---", 1)
    raw_user = raw_user.strip() or (username_from_clixml(path) or "")
    password = password.strip("\r\n")
    if not password:
        raise RuntimeError("Imported credential password is empty")
    return to_upn(raw_user, upn_suffix=upn_suffix), password


def parse_reboot_time(value: Any) -> dt_time | None:
    """Parse Excel time cells: 12, 12:10, datetime, or Excel fraction of a day."""
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return value.time().replace(second=0, microsecond=0)
    if hasattr(value, "hour") and hasattr(value, "minute") and not isinstance(value, datetime):
        return dt_time(value.hour, value.minute)
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if 0 <= float(value) < 1:
            seconds = int(round(float(value) * 24 * 3600))
            hours, rem = divmod(seconds, 3600)
            minutes, _ = divmod(rem, 60)
            return dt_time(hours % 24, minutes)
        if 0 <= int(value) <= 23 and float(value).is_integer():
            return dt_time(int(value), 0)
    text = str(value).strip()
    if not text:
        return None
    text = text.replace(".", ":")
    for fmt in ("%H:%M:%S", "%H:%M", "%H"):
        try:
            return datetime.strptime(text, fmt).time().replace(second=0, microsecond=0)
        except ValueError:
            continue
    match = re.fullmatch(r"(\d{1,2})(?::(\d{1,2}))?", text)
    if match:
        hour = int(match.group(1))
        minute = int(match.group(2) or 0)
        if 0 <= hour <= 23 and 0 <= minute <= 59:
            return dt_time(hour, minute)
    raise ValueError(f"Unrecognized reboot time: {value!r}")


def _norm_header(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip().lower())


def _header_index(headers: list[str], aliases: tuple[str, ...]) -> int | None:
    for idx, header in enumerate(headers):
        if header in aliases:
            return idx
    return None


def load_excel_rows(excel_path: Path) -> list[dict[str, Any]]:
    if load_workbook is None:
        raise RuntimeError("openpyxl is required. Install with: pip install openpyxl")
    wb = load_workbook(excel_path, data_only=True, read_only=True)
    try:
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
    finally:
        wb.close()
    if not rows:
        raise ValueError(f"Excel file is empty: {excel_path}")
    headers = [_norm_header(h) for h in rows[0]]
    server_i = _header_index(headers, SERVER_HEADERS)
    component_i = _header_index(headers, COMPONENT_HEADERS)
    time_i = _header_index(headers, TIME_HEADERS)
    vcenter_i = _header_index(headers, VCENTER_HEADERS)
    if server_i is None or component_i is None or time_i is None:
        raise ValueError(
            "Excel must have Server, Component, and Time columns. "
            f"Found headers: {rows[0]!r}"
        )
    records: list[dict[str, Any]] = []
    for row in rows[1:]:
        if row is None or all(c is None or str(c).strip() == "" for c in row):
            continue
        records.append(
            {
                "server": str(row[server_i]).strip() if row[server_i] is not None else "",
                "component": str(row[component_i]).strip() if row[component_i] is not None else "",
                "time": parse_reboot_time(row[time_i]),
                "vcenter": (
                    str(row[vcenter_i]).strip()
                    if vcenter_i is not None and row[vcenter_i] is not None
                    else ""
                ),
            }
        )
    return records


def filter_component(rows: list[dict[str, Any]], component: str) -> list[dict[str, Any]]:
    wanted = component.strip().upper()
    selected = []
    for row in rows:
        if not row["server"]:
            continue
        if (row["component"] or "").strip().upper() != wanted:
            continue
        if row["time"] is None:
            LOG.warning("Skipping %s: no reboot time in Excel", row["server"])
            continue
        selected.append(row)
    return selected


def group_by_time(rows: list[dict[str, Any]]) -> list[tuple[dt_time, list[dict[str, Any]]]]:
    groups: dict[dt_time, list[dict[str, Any]]] = defaultdict(list)
    for row in rows:
        groups[row["time"]].append(row)
    return sorted(groups.items(), key=lambda item: item[0])


def scheduled_datetime(clock: dt_time, now: datetime | None = None) -> datetime:
    now = now or datetime.now()
    return datetime.combine(now.date(), clock)


def seconds_until(target: datetime, now: datetime | None = None) -> float:
    now = now or datetime.now()
    return (target - now).total_seconds()


def wait_until(target: datetime, dry_run: bool) -> None:
    remaining = seconds_until(target)
    if remaining <= 0:
        return
    LOG.info("Waiting until %s (%s)", target.strftime("%H:%M"), timedelta(seconds=int(remaining)))
    if dry_run:
        LOG.info("Dry-run: not sleeping")
        return
    while True:
        remaining = seconds_until(target)
        if remaining <= 0:
            return
        time.sleep(min(remaining, 30))


def wait_minutes(minutes: float, dry_run: bool, reason: str) -> None:
    LOG.info("Waiting %s minute(s) %s", minutes, reason)
    if dry_run:
        LOG.info("Dry-run: not sleeping")
        return
    time.sleep(max(0.0, minutes * 60))


def connect_vcenters(hosts: list[str], username: str, password: str, port: int) -> dict[str, Any]:
    if SmartConnect is None:
        raise RuntimeError("pyvmomi is required. Install with: pip install pyvmomi")
    context = ssl._create_unverified_context()
    sessions: dict[str, Any] = {}
    errors: list[str] = []
    for host in hosts:
        try:
            LOG.info("Connecting to vCenter %s as %s", host, username)
            si = SmartConnect(
                host=host,
                user=username,
                pwd=password,
                port=port,
                sslContext=context,
            )
            sessions[host] = si
        except Exception as exc:  # noqa: BLE001 - surface every login failure
            errors.append(f"{host}: {exc}")
            LOG.error("Failed to login to %s: %s", host, exc)
    if not sessions:
        raise RuntimeError("Could not log in to any vCenter. " + "; ".join(errors))
    return sessions


def disconnect_all(sessions: dict[str, Any]) -> None:
    for host, si in sessions.items():
        try:
            if Disconnect is not None:
                Disconnect(si)
        except Exception as exc:  # noqa: BLE001
            LOG.debug("Disconnect %s: %s", host, exc)


def _iter_vms(si: Any):
    content = si.RetrieveContent()
    view = content.viewManager.CreateContainerView(content.rootFolder, [vim.VirtualMachine], True)
    try:
        for vm in view.view:
            yield vm
    finally:
        view.Destroy()


def find_vm(sessions: dict[str, Any], server: str, preferred_vcenter: str = "") -> tuple[str, Any] | None:
    targets = list(sessions)
    if preferred_vcenter:
        exact = [h for h in targets if h.lower() == preferred_vcenter.lower()]
        if exact:
            targets = exact + [h for h in targets if h not in exact]
        else:
            LOG.warning(
                "Excel vCenter %s is not in the hardcoded list; searching all vCenters for %s",
                preferred_vcenter,
                server,
            )
    wanted = server.lower()
    for host in targets:
        si = sessions[host]
        for vm in _iter_vms(si):
            name = (vm.name or "")
            if name.lower() == wanted:
                return host, vm
    return None


def reboot_vm(vm: Any, dry_run: bool) -> str:
    name = vm.name
    if dry_run:
        return f"dry-run: would reboot {name}"
    runtime = vm.runtime
    if runtime.powerState != vim.VirtualMachinePowerState.poweredOn:
        return f"skipped {name}: power state is {runtime.powerState}"
    tools = getattr(runtime, "toolsRunningStatus", None)
    try:
        if tools == "guestToolsRunning":
            vm.RebootGuest()
            return f"guest reboot issued for {name}"
        vm.ResetVM_Task()
        return f"reset task issued for {name} (VMware Tools not running)"
    except Exception as exc:  # noqa: BLE001
        raise RuntimeError(f"Failed to reboot {name}: {exc}") from exc


def reboot_set(
    rows: list[dict[str, Any]],
    sessions: dict[str, Any] | None,
    dry_run: bool,
) -> None:
    for row in rows:
        server = row["server"]
        if dry_run and not sessions:
            LOG.info("Dry-run: would reboot %s (component=%s)", server, row["component"])
            continue
        found = find_vm(sessions or {}, server, row.get("vcenter") or "")
        if not found:
            LOG.error("VM not found in any connected vCenter: %s", server)
            continue
        host, vm = found
        try:
            result = reboot_vm(vm, dry_run=dry_run)
            LOG.info("[%s] %s", host, result)
        except Exception as exc:  # noqa: BLE001
            LOG.error("%s", exc)


def run_schedule(
    groups: list[tuple[dt_time, list[dict[str, Any]]]],
    sessions: dict[str, Any] | None,
    wait_minutes_between: float,
    dry_run: bool,
    skip_past: bool,
    now: datetime | None = None,
) -> None:
    now = now or datetime.now()
    for index, (clock, rows) in enumerate(groups):
        start_at = scheduled_datetime(clock, now=now)
        names = ", ".join(r["server"] for r in rows)
        LOG.info(
            "Set %s/%s at %s (%s host(s)): %s",
            index + 1,
            len(groups),
            clock.strftime("%H:%M"),
            len(rows),
            names,
        )
        if skip_past and start_at < now and seconds_until(start_at, now=now) < -60:
            LOG.warning("Skipping past set at %s", clock.strftime("%H:%M"))
            continue
        wait_until(start_at, dry_run=dry_run)
        reboot_set(rows, sessions, dry_run=dry_run)
        if index < len(groups) - 1:
            wait_minutes(
                wait_minutes_between,
                dry_run=dry_run,
                reason="before the next UAG set",
            )


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Reboot UAG VMs from Excel via three vCenters")
    parser.add_argument("--excel", required=True, type=Path, help="Excel workbook with server list")
    parser.add_argument(
        "--cred",
        required=True,
        type=Path,
        help="CLIXML from Export-VCenterCredential.ps1 / Get-Credential",
    )
    parser.add_argument("--component", default=COMPONENT_FILTER, help="Component to reboot (default: UAG)")
    parser.add_argument(
        "--wait-minutes",
        type=float,
        default=WAIT_BETWEEN_SETS_MINUTES,
        help="Wait after each set before the next (default: 10)",
    )
    parser.add_argument("--dry-run", action="store_true", help="Parse/plan only; do not login or reboot")
    parser.add_argument("--skip-past", action="store_true", help="Do not reboot sets whose time already passed")
    parser.add_argument("--verbose", action="store_true")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    configure_logging(args.verbose)
    rows = load_excel_rows(args.excel)
    selected = filter_component(rows, args.component)
    if not selected:
        LOG.error("No %s servers with a reboot time were found in %s", args.component, args.excel)
        return 1
    groups = group_by_time(selected)
    LOG.info(
        "Loaded %s %s server(s) in %s timed set(s) from %s",
        len(selected),
        args.component.upper(),
        len(groups),
        args.excel,
    )

    sessions = None
    try:
        if args.dry_run:
            LOG.info("Dry-run: skipping vCenter login (username will still be converted from XML if readable)")
            xml_user = username_from_clixml(args.cred)
            if xml_user:
                LOG.info("XML username %s -> UPN %s", xml_user, to_upn(xml_user))
        else:
            username, password = load_ps_credential(args.cred)
            LOG.info("Using vCenter username %s", username)
            sessions = connect_vcenters(VCENTERS, username, password, VCENTER_PORT)
        run_schedule(
            groups,
            sessions,
            wait_minutes_between=args.wait_minutes,
            dry_run=args.dry_run,
            skip_past=args.skip_past,
        )
    finally:
        if sessions:
            disconnect_all(sessions)
    return 0


if __name__ == "__main__":
    sys.exit(main())
