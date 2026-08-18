# Azure

Collection of Azure operational scripts.

## Azure Virtual Desktop inventory (`export_avd_inventory.py`)

Python script that signs in to **one Azure subscription**, inventories every Azure Virtual Desktop (AVD) host pool, and writes a single Excel workbook with multiple sheets.

### What is collected

| Sheet | Contents |
| --- | --- |
| **HostPools_SessionHosts** | Host pool details, type (Pooled/Personal), and every associated session host |
| **AppGroups_Access** | Host pool, type, application groups, published applications/desktops, and who is assigned access |
| HostPools | One row per host pool |
| SessionHosts | One row per session host |
| ApplicationGroups | Application groups with assignment and application summaries |
| Assignments | Role assignments on each application group (users, groups, service principals) |
| Applications | Published RemoteApps / session desktops and the principals that have access |
| Workspaces | Workspaces and linked application groups |
| Summary | Counts |

Application-group access is read from Azure RBAC on the application group (typically **Desktop Virtualization User**). Microsoft Graph is used when available to resolve object IDs to display names and UPNs.

### Setup

```bash
python -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt
az login
az account set --subscription "<subscription-id>"
```

### Run

```bash
python export_avd_inventory.py --subscription-id "<subscription-id>"
```

Authentication options:

```bash
# Azure CLI / environment / managed identity / browser (default)
python export_avd_inventory.py --subscription-id "<subscription-id>" --auth default

# Device code (headless / no browser on this machine)
python export_avd_inventory.py --subscription-id "<subscription-id>" --auth device --tenant-id "<tenant-id>"

# Interactive browser
python export_avd_inventory.py --subscription-id "<subscription-id>" --auth interactive --tenant-id "<tenant-id>"
```

Write to a specific file:

```bash
python export_avd_inventory.py --subscription-id "<subscription-id>" --output avd-inventory.xlsx
```

Include role assignments inherited from the resource group or subscription:

```bash
python export_avd_inventory.py --subscription-id "<subscription-id>" --include-inherited-assignments
```

### Permissions

- **Desktop Virtualization Reader** (or higher) on the subscription
- `Microsoft.Authorization/roleAssignments/read`
- Optional Microsoft Graph: `Directory.Read.All` (or equivalent) so assignment object IDs resolve to names
