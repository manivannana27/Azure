#!/usr/bin/env python3
"""
Read vCenter and VM details from CSV (ServerName|Role|Site|Environment|Reboot).

- Log in to every row whose Role is a vCenter (Prod-vCenter, DMZ-vCenter, …)
- Reboot VMs whose Reboot column is 12 (full or short VM name)
- Wait 60 seconds, check power status, write a TXT log

Credentials come from a PowerShell Get-Credential Export-Clixml file.
DOMAIN\\user in that XML is converted to user@domain.com before login.

Example:
  python reboot_uag.py
  python reboot_uag.py --csv inventory.csv --cred vcenter_credential.xml --dry-run
"""

from __future__ import annotations

import argparse
import csv
import io
import ssl
import socket
import subprocess
import sys
import time
import xml.etree.ElementTree as ET
from datetime import datetime
from pathlib import Path
from typing import Any

# ------- CONFIG (edit paths / suffix) -------
CLIXML_PATH = r"vcenter_credential.xml"
CSV_PATH = r"inventory.csv"
OUTPUT_DIR = r"."
UPN_SUFFIX = "example.com"
TARGET_REBOOT = 12
STATUS_WAIT_SECONDS = 60
# Explicit names plus any Role that contains "vcenter"
VCENTER_ROLES = ("prod-vcenter", "dmz-vcenter", "vcenter")
# ----------------------------------------

try:
    from pyVim.connect import Disconnect, SmartConnect
    try:
        from pyVim.connect import SmartConnectNoSSL
    except Exception:  # pragma: no cover
        SmartConnectNoSSL = None
    from pyVmomi import vim
except ImportError:  # pragma: no cover
    Disconnect = None  # type: ignore[assignment]
    SmartConnect = None  # type: ignore[assignment]
    SmartConnectNoSSL = None  # type: ignore[assignment]
    vim = None  # type: ignore[assignment]


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
            raise ValueError("NETBIOS domain in credential XML requires UPN_SUFFIX")
        return f"{user}@{suffix}"
    if upn_suffix and "." not in value:
        return f"{value}@{upn_suffix.lstrip('@')}"
    return value


def _local_tag(tag: str) -> str:
    return tag.split("}")[-1] if tag else tag


def username_from_clixml(xml_path: Path) -> str | None:
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
            completed = subprocess.run(cmd, capture_output=True, text=True, check=False)
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


def read_credentials_from_clixml(xml_path: str | Path, upn_suffix: str = UPN_SUFFIX) -> tuple[str, str]:
    path = Path(xml_path).resolve()
    if not path.is_file():
        raise FileNotFoundError(f"CLIXML file not found: {xml_path}")

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
        raise RuntimeError("CLIXML import did not return a password.")
    raw_user, password = output.split("---PASSWORD---", 1)
    raw_user = raw_user.strip() or (username_from_clixml(path) or "")
    password = password.strip("\r\n")
    if not password:
        raise RuntimeError("Imported credential password is empty")
    return to_upn(raw_user, upn_suffix=upn_suffix), password


def normalize_row(row: dict[str, Any]) -> dict[str, str]:
    out: dict[str, str] = {}
    for key, value in (row or {}).items():
        name = str(key or "").replace("\ufeff", "").strip().lower()
        out[name] = "" if value is None else str(value).strip()
    return out


def row_get(row: dict[str, str], *names: str) -> str:
    for name in names:
        if name in row and row[name]:
            return row[name]
    return ""


def parse_reboot_value(value: str) -> int | None:
    text = (value or "").strip()
    if not text:
        return None
    if ":" in text:
        hour = text.split(":", 1)[0].strip()
        minute = text.split(":", 1)[1].strip()
        try:
            if int(minute) == 0:
                return int(hour)
        except ValueError:
            return None
        return None
    try:
        number = float(text)
    except ValueError:
        return None
    if number.is_integer():
        return int(number)
    return None


def is_vcenter_role(role: str, vcenter_roles: tuple[str, ...] = VCENTER_ROLES) -> bool:
    """Match Prod-vCenter / DMZ-vCenter regardless of CSV casing."""
    value = (role or "").strip().lower()
    if not value:
        return False
    known = {r.strip().lower() for r in vcenter_roles}
    if value in known:
        return True
    return "vcenter" in value


def read_csv_rows(path: str | Path) -> list[dict[str, str]]:
    p = Path(path)
    if not p.exists():
        raise FileNotFoundError(f"CSV file not found: {path}")

    raw = p.read_text(encoding="utf-8-sig", errors="ignore")
    text_lines = raw.splitlines()
    if not text_lines:
        return []

    sample = "\n".join(text_lines[:10])
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",|\t;")
    except csv.Error:
        dialect = csv.get_dialect("excel")

    reader = csv.DictReader(io.StringIO("\n".join(text_lines)), dialect=dialect)
    return [normalize_row(row) for row in reader if any(str(v or "").strip() for v in row.values())]


def classify_rows(
    rows: list[dict[str, str]],
    target_reboot: int = TARGET_REBOOT,
) -> tuple[list[str], list[dict[str, str]]]:
    """Return (vCenter hostnames, VM rows to reboot)."""
    vcenters: list[str] = []
    reboot_vms: list[dict[str, str]] = []
    seen_vc: set[str] = set()

    for row in rows:
        role = row_get(row, "role")
        server = row_get(row, "servername", "server", "vm", "name", "hostname")
        reboot_val = parse_reboot_value(row_get(row, "reboot"))

        if is_vcenter_role(role) and server:
            key = server.lower()
            if key not in seen_vc:
                vcenters.append(server)
                seen_vc.add(key)
            continue

        if server and reboot_val == target_reboot:
            reboot_vms.append(row)

    return vcenters, reboot_vms


def build_search_index(reboot_vm_rows: list[dict[str, str]]) -> tuple[set[str], dict[str, list[str]]]:
    search_keys: set[str] = set()
    key_to_csvnames: dict[str, list[str]] = {}
    for row in reboot_vm_rows:
        full = row_get(row, "servername", "server", "vm", "name", "hostname")
        if not full:
            continue
        short = full.split(".", 1)[0] if "." in full else full
        for key in {full.lower(), short.lower()}:
            if not key:
                continue
            search_keys.add(key)
            names = key_to_csvnames.setdefault(key, [])
            if full not in names:
                names.append(full)
    return search_keys, key_to_csvnames


def connect_vcenter(host: str, user: str, pwd: str, insecure: bool = True):
    if SmartConnect is None:
        raise RuntimeError("pyvmomi is required. Install with: pip install pyvmomi")
    try:
        socket.getaddrinfo(host, 443)
    except Exception as exc:
        raise RuntimeError(f"Host resolution failed for '{host}': {exc}") from exc

    if insecure:
        if SmartConnectNoSSL:
            return SmartConnectNoSSL(host=host, user=user, pwd=pwd)
        ctx = ssl.create_default_context()
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE
        return SmartConnect(host=host, user=user, pwd=pwd, sslContext=ctx)
    ctx = ssl.create_default_context()
    return SmartConnect(host=host, user=user, pwd=pwd, sslContext=ctx)


def get_vm_matches_for_vcenter(si: Any, search_keys: set[str], key_to_csvnames: dict[str, list[str]]):
    content = si.RetrieveContent()
    view = content.viewManager.CreateContainerView(content.rootFolder, [vim.VirtualMachine], True)
    results: dict[str, list[Any]] = {}
    try:
        for vm in view.view:
            vm_name = (vm.name or "").strip()
            if not vm_name:
                continue
            key = vm_name.lower()
            if key not in search_keys:
                continue
            for csv_name in key_to_csvnames.get(key, []):
                results.setdefault(csv_name, []).append(vm)
    finally:
        view.Destroy()
    return results


def request_reboot(vm: Any, dry_run: bool = False) -> str:
    if dry_run:
        return "dry-run"
    try:
        try:
            vm.RebootGuest()
            return "RebootGuest"
        except Exception as inner:
            try:
                state = str(vm.runtime.powerState)
            except Exception:
                state = "unknown"
            if state == "poweredOn":
                try:
                    vm.ResetVM_Task()
                    return f"ResetVM_Task (fallback from RebootGuest: {type(inner).__name__})"
                except Exception as reset_err:
                    return f"RebootFailed_ResetError({type(reset_err).__name__})"
            return f"NoReboot_VMState={state}"
    except Exception as exc:
        return f"RebootFailed({type(exc).__name__})"


def get_vm_power_state(vm: Any) -> str:
    try:
        return str(vm.runtime.powerState)
    except Exception:
        return "unknown"


def write_results(log_path: Path, results: list[dict[str, Any]], target_reboot: int = TARGET_REBOOT) -> None:
    log_path.parent.mkdir(parents=True, exist_ok=True)
    with log_path.open("w", encoding="utf-8") as handle:
        handle.write(f"Reboot Results (Reboot=={target_reboot} VMs)\n")
        handle.write(f"Generated at: {datetime.now()}\n\n")
        for entry in results:
            handle.write(
                f"CSV_VM={entry['csv_name']}, "
                f"vCenter={entry['host']}, "
                f"VM={entry['vm_name']}, "
                f"action={entry['action']}, "
                f"final_power={entry['final_power']}\n"
            )


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Reboot CSV Reboot==12 VMs via Prod/DMZ vCenters")
    parser.add_argument("--csv", default=CSV_PATH, help="Inventory CSV path")
    parser.add_argument("--cred", default=CLIXML_PATH, help="CLIXML credential path")
    parser.add_argument("--output-dir", default=OUTPUT_DIR, help="Directory for result TXT files")
    parser.add_argument("--reboot", type=int, default=TARGET_REBOOT, help="Reboot column value to select (default: 12)")
    parser.add_argument("--wait-seconds", type=int, default=STATUS_WAIT_SECONDS)
    parser.add_argument("--dry-run", action="store_true", help="Classify/plan only; do not login or reboot")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)

    try:
        rows = read_csv_rows(args.csv)
    except Exception as exc:
        print("ERROR reading CSV:", exc, file=sys.stderr)
        return 3

    if not rows:
        print("CSV is empty. Exiting.")
        return 1

    headers = list(rows[0].keys())
    if "role" not in headers:
        print(
            "CSV is missing the Role column (heading must be Role, not Roles). "
            f"Found headers: {headers}"
        )
        return 1

    prod_vcenters, reboot_vm_rows = classify_rows(rows, target_reboot=args.reboot)
    print(f"Found {len(prod_vcenters)} vCenter row(s): {', '.join(prod_vcenters) or '(none)'}")
    print(f"Found {len(reboot_vm_rows)} VM row(s) with Reboot=={args.reboot}")

    if not prod_vcenters:
        print(
            "No vCenter entries in CSV. Role must be Prod-vCenter, DMZ-vCenter, "
            "or another value containing 'vCenter' (any case)."
        )
        return 1

    if not reboot_vm_rows:
        print(f"No VMs with Reboot=={args.reboot} in CSV. Exiting.")
        return 1

    search_keys, key_to_csvnames = build_search_index(reboot_vm_rows)

    if args.dry_run:
        xml_user = username_from_clixml(Path(args.cred))
        if xml_user:
            print(f"XML username {xml_user} -> UPN {to_upn(xml_user)}")
        print("Dry-run: would log in to:", ", ".join(prod_vcenters))
        for row in reboot_vm_rows:
            print("Dry-run: would reboot", row_get(row, "servername", "server", "vm", "name"))
        return 0

    try:
        username, password = read_credentials_from_clixml(args.cred)
    except Exception as exc:
        print("ERROR reading CLIXML:", exc, file=sys.stderr)
        return 2

    print(f"Using vCenter username {username}")

    vc_sessions: dict[str, Any] = {}
    global_matches: dict[str, dict[str, list[Any]]] = {}

    try:
        for host in prod_vcenters:
            print(f"[{host}] attempting login...")
            try:
                si = connect_vcenter(host, username, password, insecure=True)
                vc_sessions[host] = si
                print(f"[{host}] logged in successfully")
            except Exception as exc:
                print(f"[{host}] login FAILED: {type(exc).__name__}: {exc}")

        if not vc_sessions:
            print("No vCenters logged in successfully. Exiting.")
            return 1

        for host, si in vc_sessions.items():
            print(f"[{host}] scanning VMs for Reboot=={args.reboot} list...")
            try:
                vc_results = get_vm_matches_for_vcenter(si, search_keys, key_to_csvnames)
            except Exception as exc:
                print(f"[{host}] error while scanning VMs: {exc}")
                continue
            for csv_name, vm_list in vc_results.items():
                host_map = global_matches.setdefault(csv_name, {})
                host_map.setdefault(host, []).extend(vm_list)

        targets = []
        for csv_name, host_map in global_matches.items():
            for host, vms in host_map.items():
                for vm in vms:
                    targets.append({"csv_name": csv_name, "host": host, "vm": vm})

        if not targets:
            print("No matching VMs found in any vCenter for Reboot==12 rows.")
            return 1

        print("\n=== Sending reboot commands to matching VMs ===")
        results: list[dict[str, Any]] = []
        for target in targets:
            vm = target["vm"]
            vm_name = (vm.name or "").strip()
            action = request_reboot(vm, dry_run=False)
            print(
                f"Reboot request: CSV VM='{target['csv_name']}', "
                f"vCenter='{target['host']}', VM='{vm_name}', action='{action}'"
            )
            results.append(
                {
                    "csv_name": target["csv_name"],
                    "host": target["host"],
                    "vm_name": vm_name,
                    "action": action,
                    "final_power": None,
                }
            )

        print(f"\nWaiting {args.wait_seconds} seconds before checking status...")
        time.sleep(max(0, args.wait_seconds))

        print("\n=== Checking power status after reboot ===")
        for entry in results:
            vm_obj = None
            for vm in global_matches.get(entry["csv_name"], {}).get(entry["host"], []):
                if (vm.name or "").strip() == entry["vm_name"]:
                    vm_obj = vm
                    break
            final_state = "unknown_vm_lost" if vm_obj is None else get_vm_power_state(vm_obj)
            entry["final_power"] = final_state
            print(
                f"Status: CSV VM='{entry['csv_name']}', "
                f"vCenter='{entry['host']}', VM='{entry['vm_name']}', "
                f"action='{entry['action']}', final_power='{final_state}'"
            )

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_path = Path(args.output_dir) / f"reboot_results_{args.reboot}_{timestamp}.txt"
        write_results(log_path, results, target_reboot=args.reboot)
        print(f"\nReboot results saved to: {log_path}")
        return 0
    finally:
        for host, si in vc_sessions.items():
            try:
                if Disconnect is not None:
                    Disconnect(si)
                print(f"[{host}] logged out successfully")
            except Exception as exc:
                print(f"[{host}] logout error: {exc}")


if __name__ == "__main__":
    sys.exit(main())
