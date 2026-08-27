import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from reboot_uag import (  # noqa: E402
    build_search_index,
    classify_rows,
    is_vcenter_role,
    parse_reboot_value,
    read_csv_rows,
    to_upn,
    username_from_clixml,
)


def test_to_upn_converts_netbios_slash_to_user_at_domain():
    assert to_upn(r"CONTOSO\jsmith", upn_suffix="contoso.com") == "jsmith@contoso.com"


def test_to_upn_converts_dns_domain_slash():
    assert to_upn(r"corp.example.com\jsmith") == "jsmith@corp.example.com"


def test_to_upn_keeps_existing_upn():
    assert to_upn("jsmith@contoso.com") == "jsmith@contoso.com"


def test_is_vcenter_role_matches_prod_and_dmz_any_case():
    assert is_vcenter_role("Prod-vCenter")
    assert is_vcenter_role("prod-vcenter")
    assert is_vcenter_role("DMZ-vCenter")
    assert is_vcenter_role("dmz-vcenter")
    assert not is_vcenter_role("UAG")
    assert not is_vcenter_role("CS")


def test_old_case_bug_would_miss_dmz():
    # role.lower() in {"DMZ-vCenter"} is False; the helper must still match.
    assert "dmz-vcenter" not in {"DMZ-vCenter"}
    assert is_vcenter_role("DMZ-vCenter")


def test_parse_reboot_value():
    assert parse_reboot_value("12") == 12
    assert parse_reboot_value("12:00") == 12
    assert parse_reboot_value("12:10") is None
    assert parse_reboot_value("") is None


def test_classify_rows_finds_vcenters_and_reboot_12(tmp_path: Path):
    csv_path = tmp_path / "inv.csv"
    csv_path.write_text(
        "ServerName,Role,Site,Environment,Reboot\n"
        "vc1.example.com,Prod-vCenter,A,Prod,\n"
        "vc2.example.com,DMZ-vCenter,A,Prod,\n"
        "server1,CS,A,Prod,12\n"
        "server2,UAG,A,Prod,12\n"
        "server3,UAG,A,Prod,12:10\n",
        encoding="utf-8",
    )
    rows = read_csv_rows(csv_path)
    vcenters, vms = classify_rows(rows, target_reboot=12)
    assert vcenters == ["vc1.example.com", "vc2.example.com"]
    assert [r["servername"] for r in vms] == ["server1", "server2"]


def test_classify_pipe_delimited_and_roles_header(tmp_path: Path):
    csv_path = tmp_path / "inv.csv"
    csv_path.write_text(
        "ServerName|Role|Site|Environment|Reboot\n"
        "vc1|prod-vcenter|A|Prod|\n"
        "uag01|UAG|A|Prod|12\n",
        encoding="utf-8",
    )
    rows = read_csv_rows(csv_path)
    vcenters, vms = classify_rows(rows)
    assert vcenters == ["vc1"]
    assert vms[0]["servername"] == "uag01"


def test_vcenter_row_with_reboot_12_is_not_a_guest_target():
    rows = [
        {"servername": "vc1", "role": "Prod-vCenter", "reboot": "12"},
        {"servername": "vm1", "role": "UAG", "reboot": "12"},
    ]
    vcenters, vms = classify_rows(rows)
    assert vcenters == ["vc1"]
    assert [r["servername"] for r in vms] == ["vm1"]


def test_build_search_index_includes_short_name():
    rows = [{"servername": "host.corp.example.com", "role": "UAG", "reboot": "12"}]
    keys, mapping = build_search_index(rows)
    assert "host.corp.example.com" in keys
    assert "host" in keys
    assert mapping["host"] == ["host.corp.example.com"]


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
