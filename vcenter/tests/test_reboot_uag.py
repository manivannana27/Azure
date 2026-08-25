import sys
from datetime import datetime, time
from pathlib import Path

from openpyxl import Workbook

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from reboot_uag import (  # noqa: E402
    filter_component,
    group_by_time,
    load_excel_rows,
    parse_reboot_time,
    scheduled_datetime,
    seconds_until,
    to_upn,
    username_from_clixml,
)


def test_to_upn_converts_netbios_slash_to_user_at_domain():
    assert to_upn(r"CONTOSO\jsmith", upn_suffix="contoso.com") == "jsmith@contoso.com"


def test_to_upn_converts_dns_domain_slash():
    assert to_upn(r"corp.example.com\jsmith") == "jsmith@corp.example.com"


def test_to_upn_keeps_existing_upn():
    assert to_upn("jsmith@contoso.com") == "jsmith@contoso.com"


def test_to_upn_bare_user_gets_suffix():
    assert to_upn("jsmith", upn_suffix="contoso.com") == "jsmith@contoso.com"


def test_parse_time_hour_only():
    assert parse_reboot_time(12) == time(12, 0)
    assert parse_reboot_time("12") == time(12, 0)


def test_parse_time_hour_minute():
    assert parse_reboot_time("12:10") == time(12, 10)


def test_parse_excel_fraction_of_day():
    value = 12 / 24 + 10 / 1440
    assert parse_reboot_time(value) == time(12, 10)


def test_filter_only_uag():
    rows = [
        {"server": "server1", "component": "cs", "time": time(12, 0)},
        {"server": "server2", "component": "uag", "time": time(12, 0)},
        {"server": "server3", "component": "UAG", "time": time(12, 10)},
    ]
    selected = filter_component(rows, "UAG")
    assert [r["server"] for r in selected] == ["server2", "server3"]


def test_group_by_time_orders_sets():
    rows = [
        {"server": "b", "component": "UAG", "time": time(12, 10)},
        {"server": "a", "component": "UAG", "time": time(12, 0)},
        {"server": "c", "component": "UAG", "time": time(12, 0)},
    ]
    groups = group_by_time(rows)
    assert [t.strftime("%H:%M") for t, _ in groups] == ["12:00", "12:10"]
    assert [r["server"] for r in groups[0][1]] == ["a", "c"]


def test_seconds_until_future():
    now = datetime(2026, 8, 25, 11, 50, 0)
    target = scheduled_datetime(time(12, 0), now=now)
    assert seconds_until(target, now=now) == 600


def test_username_from_clixml(tmp_path: Path):
    xml = tmp_path / "cred.xml"
    xml.write_text(
        """<?xml version="1.0"?>
<Objs Version="1.1.0.1" xmlns="http://schemas.microsoft.com/powershell/2004/04">
  <Obj RefId="0">
    <Props>
      <S N="UserName">CONTOSO\\jsmith</S>
      <SS N="Password">01000000deadbeef</SS>
    </Props>
  </Obj>
</Objs>
""",
        encoding="utf-8",
    )
    assert username_from_clixml(xml) == r"CONTOSO\jsmith"


def test_load_excel_rows(tmp_path: Path):
    path = tmp_path / "servers.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.append(["Server", "Component", "Time"])
    ws.append(["server1", "cs", 12])
    ws.append(["server2", "uag", "12:10"])
    wb.save(path)
    rows = load_excel_rows(path)
    assert rows[0]["server"] == "server1"
    assert rows[1]["component"] == "uag"
    assert rows[1]["time"] == time(12, 10)
