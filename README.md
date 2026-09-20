# FortiGate Policy to Excel Exporter

A tool that parses FortiGate firewall backup configuration files (`.conf`) and automatically converts and splits policies and objects by VDOM into highly readable Excel (`.xlsx`) workbooks.

<img width="802" height="607" alt="image" src="https://github.com/user-attachments/assets/01b1390f-9840-4ad8-abbe-5a12f589af31" />


---

## 📌 Table of Contents
1. [Key Features](#-key-features)
2. [System Requirements](#-system-requirements)
3. [Installing Python and Required Modules](#-installing-python-and-required-modules)
4. [How to Run](#-how-to-run)
5. [Excel Sheet Structure & Design Highlights](#-excel-sheet-structure--design-highlights)
6. [Folder and File Structure](#-folder-and-file-structure)
7. [Frequently Asked Questions (FAQ & Troubleshooting)](#-frequently-asked-questions-faq--troubleshooting)

---

## 🌟 Key Features

### 1. Comprehensive Extraction & Dedicated Sheets for 11 Core Policies, Routing, VPN & Interface Configs
* **Firewall Policy**: Standard firewall policies (Action, Schedule, NAT, IP Pool, Inspection Mode, UTM Profile, Log, etc. with 33 detailed columns)
* **Local-in Policy**: Firewall appliance self-access control policies
* **Central-NAT**: Central SNAT Map (Original IP/Port ↔ Translated IP/Port mapping, explicit NAT enable status)
* **DNAT (VIP)**: Virtual IP and port-forwarding mapping (External IP/Port ↔ Mapped IP/Port)
* **DoS Policy**: Denial of Service (DoS/DDoS) defense policies
* **Static Route**: Static routing entries (Destination network with automatic CIDR conversion, Gateway, Interface, Distance, Priority, special flags, and disabled row shading)
* **Policy Route**: Policy-Based Routing (PBR) (Incoming/Outgoing Interface, Gateway, Action, Protocol, Port Range, address objects multi-row flattening & cell merging)
* **OSPF**: Dynamic OSPF routing configuration (Structured section tables for Global Settings/Router ID, config network, config ospf-interface, config redistribute, and Route-Map & Access-List Filter Details)
* **IPsec VPN**: IPsec VPN Phase 1 and Phase 2 tunnel configurations (P1 vs P2 header color distinction, 1:N hierarchical tunnel mapping, Proposals, P2 DH Group default '14 5', local/remote subnets default '0.0.0.0/0', and center-aligned merged cells)
* **Network Interface**: Network interface configurations per VDOM parsed from `config system interface` (Status UP/DOWN, Name, Alias, Type, IP/Subnet, VLAN ID, Parent Interface, VRF, Mode, Admin Access, Speed, Description with 14 columns)
* **External Resource**: External threat intelligence feeds (`config system external-resource`) with URL, refresh rate, source IP, enabled status, etc.

### 2. Multi-row Object/Group Expansion with Real IPs, Ports & Comments (Multi-row Flattening)
* Unlike conventional tools that simply display object or group names, this tool automatically traces **member objects, actual IP addresses, subnet ranges, and port numbers** within groups and expands them row by row.
* **Unified IP Formatting**:
  * Single Host: `x.x.x.x/32`
  * Subnets (e.g., Class C): `x.x.x.x/24`
  * IP Range: `1.1.1.1-1.1.1.10`
* **Isolated Object Comments**: Extracts comments from each object and group, placing them into dedicated `Src Comment`, `Dst Comment`, and `Svc Comment` columns, completely separating them from policy-level comments (`Policy Comment`).
* **Internet Service & External Resource Mapping**:
  * Internet services (e.g. `Fortinet-DNS`, `Dropbox-Web`) are accurately classified as source/destination objects (`Src/Dst OBJ Name`) with type `internet-service`, keeping service port columns clean.
  * External resource feed objects (e.g. `AbuseIPDB`) are resolved to their declared resource types (`address`, `domain`, `malware`, etc.).
  * Destination VIP objects automatically resolve `Dst Type` to their actual VIP type (`static-nat`, etc.).

### 3. Unified Security Profiles (Sec Profile) & 1:1 Profile Comments
* Consolidates all assigned security profiles (SSL/SSH Inspection, IPS Sensor, Web Filter, Antivirus, DNS Filter, Application Control, File Filter, etc.) into a single multi-line **`Sec Profile`** column.
* Traces comments defined in each security profile and maps them 1:1 in the adjacent **`Sec Profile Comment`** column.
* Displays VDOM-wide / policy inspection mode (**`Inspection Mode`**: `Flow-based` / `Proxy-based`).

### 4. Stable Sort by Source Interface (Ascending)
* Groups policies cleanly by sorting ascending by **`Src Interface`**.
* Strictly maintains the firewall rule evaluation order (First-Match Precedence) within each interface using **Python Timsort (Stable Sort)**.

### 5. Granular Log Traffic & UTM Status Display
* **Log Traffic**: Differentiates `disable`, `all`, and `utm` states, appending **`session-start`** on a new line when `set logtraffic-start enable` is active (e.g., `all\nsession-start`, `utm\nsession-start`).
* **UTM Status**: Clearly outputs **`disable`** when UTM is disabled or unset.

### 6. Smart Vertical Cell Merging & Top Alignment
* When a policy spans multiple rows due to multiple objects, **common policy attributes (Seq, VDOM, Enable, ID, Name, Action, Schedule, etc.)** and **identical group member areas** are automatically merged vertically.
* Sets text vertical alignment to **`top`** across all data and merged cells to maximize readability.

### 7. Enterprise-Grade Visual Styling
* **Header & Font Color Coding**:
  * Source (Src) section: Blue header + blue font
  * Destination (Dst) section: Red header + red font
  * Service section: Blue-gray header
* **Zebra Striping**: Alternating row shading across all columns ensures smooth visual scanning.
* **Disabled Policy Shading**: Disabled policies (`Enable == N`) are highlighted with a dark gray background, making them instantly distinguishable from active rules.
* **Red Tab Highlighting for Empty Sheets**: Worksheets that contain no policy data (i.e. only column headers in row 1) automatically have their sheet tab color set to **Red (`FFFF0000`)**, allowing engineers and auditors to spot unconfigured policy categories at a glance.

### 8. Hostname Directory Creation & Per-VDOM Excel Splitting
* Prevents performance lag and freezing typically caused by dumping large-scale configuration policies into a single spreadsheet.
* Automatically creates a folder named after the device hostname (e.g., `JBNU_SVF-FW1`), and generates **independent Excel workbooks (`<vdom_name>.xlsx`) for each VDOM**.
* Provides a master summary file (**`_TOTAL_SUMMARY.xlsx`**) summarizing policy and object counts across all VDOMs at a glance.

### 9. Built-in GUI with Dynamic Auto Output Directory
* Intuitive file browser dialog (`Select File...`)
* **Dynamic Output Directory**: Automatically updates the output directory to the selected `.conf` file's parent folder upon every file selection.
* Real-time progress bar and syntax-highlighted terminal console log
* Windows High-DPI / ClearType support for crisp, blur-free text rendering on QHD/4K monitors
* Automatic notification in the bottom status bar and direct folder opening upon export completion without cumbersome popup dialogs

---

## 💻 System Requirements

| Item | Recommended Specification |
|---|---|
| **Operating System** | Windows 10 / Windows 11 (64-bit) *(macOS / Linux supported via CLI mode)* |
| **Python Version** | **Python 3.8 or higher** (Fully compatible with Python 3.10 ~ 3.13) |
| **Display** | 1920×1080 (FHD) or higher (Fully supports scaled QHD & 4K environments) |
| **Required Library** | `openpyxl` |

---

## 📦 Installing Python and Required Modules

### 1. Install Python
1. Download Python 3.10 or later from the [official Python website](https://www.python.org/downloads/).
2. On the first installation screen, make sure to **check the `[✔] Add python.exe to PATH`** checkbox before proceeding with the installation.

### 2. Install Required Library (`openpyxl`)
Open Command Prompt (CMD) or PowerShell and run:

```bash
pip install openpyxl
```

---

## 🚀 How to Run

### Method 1: Double-click from Explorer (Recommended)
1. Double-click the **`start.bat`** file in the program folder.
2. The GUI window will open.
3. Click the **[Select File...]** button and choose the FortiGate `.conf` file you want to convert.
4. Click the **[Start Export to Excel]** button.
5. Once the export is complete, the destination folder will open automatically.

### Method 2: Direct Execution with Python (GUI Mode)
Run from the terminal without arguments to launch GUI mode:

```bash
python fortigate_policy_to_excel.py
```

### Method 3: Command-Line Execution (CLI Mode / Script Automation)
Run directly from batch scripts or in the background using command-line arguments without opening the GUI:

```bash
# Basic usage: python fortigate_policy_to_excel.py <config_file_path> [output_directory]
python fortigate_policy_to_excel.py "C:\backup\my_firewall.conf"
```

---

## 📊 Excel Sheet Structure & Design Highlights

### 1. Firewall Policy Sheet Column Structure (33 Columns)

```
[Policy Basic Information]
  Col 1: Seq                 - Sequence number (Sorted ascending by Src Interface)
  Col 2: VDOM                - VDOM name
  Col 3: Enable              - Enabled status (Y / N)
  Col 4: ID                  - Policy ID
  Col 5: Name                - Policy name
  Col 6: Action              - accept / deny

[Source Section (Src) - Blue header & font]
  Col 7: Src Interface       - Source interface
  Col 8: Src Group OBJ       - Source group name
  Col 9: Src OBJ Name        - Member object name (Address / Group / Internet Service / External Resource)
  Col 10: Src Type           - ipmask / iprange / fqdn / internet-service, etc.
  Col 11: Src IP             - Actual IP / Subnet / Range
  Col 12: Src Comment        - Source object comment

[Destination Section (Dst) - Red header & font]
  Col 13: Dst Interface      - Destination interface
  Col 14: Dst Group OBJ      - Destination group name
  Col 15: Dst OBJ Name       - Member object name (Address / Group / Internet Service / External Resource)
  Col 16: Dst Type           - ipmask / iprange / fqdn / internet-service / static-nat(VIP), etc.
  Col 17: Dst IP             - Actual IP / Subnet / Range
  Col 18: Dst Comment        - Destination object comment

[Service Section (Service) - Blue-gray header]
  Col 19: Svc Group OBJ      - Service group name
  Col 20: Svc OBJ Name       - Service object name
  Col 21: Svc Port           - Port number (IP/ALL displayed as ALL, blank for Internet Service)
  Col 22: Svc Comment        - Service object comment

[General & Security Profiles]
  Col 23: Schedule           - Schedule
  Col 24: NAT                - NAT status (enable / disable)
  Col 25: IP Pool            - IP Pool status
  Col 26: Pool Name          - Pool name
  Col 27: Pool IP            - Assigned Pool IP range
  Col 28: Inspection Mode    - Flow-based / Proxy-based
  Col 29: UTM Status         - UTM applied status (enable / disable)
  Col 30: Sec Profile        - Applied security profiles (SSL/SSH, IPS, AV, WebFilter, etc. combined with newlines)
  Col 31: Sec Profile Comment- Comments from each security profile (1:1 mapping with newlines)
  Col 32: Log Traffic        - Traffic logging (all / utm / disable combined with session-start)
  Col 33: Comment            - Policy-level comments
```

### 2. Local-in Policy Sheet Column Structure (22 Columns)

```
[Policy Basic Information]
  Col 1: Seq                 - Sequence number
  Col 2: VDOM                - VDOM name
  Col 3: Enable              - Enabled status (Y / N)
  Col 4: ID                  - Policy ID
  Col 5: Interface           - Inbound listening interface

[Source Section (Src)]
  Col 6: Src Group OBJ       - Source group name
  Col 7: Src OBJ Name        - Source object name
  Col 8: Src Type            - Address object type (ipmask / fqdn, etc.)
  Col 9: Src IP              - Source IP / subnet range
  Col 10: Src Comment        - Source object comment

[Destination Section (Dst)]
  Col 11: Dst Group OBJ      - Destination group name
  Col 12: Dst OBJ Name       - Destination object name
  Col 13: Dst Type           - Address object type
  Col 14: Dst IP             - Destination IP / subnet range
  Col 15: Dst Comment        - Destination object comment

[Action & Service]
  Col 16: Action             - accept / deny
  Col 17: Svc Group OBJ      - Service group name
  Col 18: Svc OBJ Name       - Service object name
  Col 19: Svc Port           - Service port number
  Col 20: Svc Comment        - Service object comment

[General]
  Col 21: Schedule           - Schedule
  Col 22: Comment            - Policy comments
```

### 3. Central-NAT Sheet Column Structure (21 Columns)

```
[Policy Basic Information]
  Col 1: Seq                 - Sequence number
  Col 2: VDOM                - VDOM name
  Col 3: Enable              - Enabled status (Y / N)
  Col 4: ID                  - NAT rule ID
  Col 5: Src Interface       - Source interface
  Col 6: Dst Interface       - Destination interface

[Original Source Section (Orig Src)]
  Col 7: Orig Group OBJ      - Original source group name
  Col 8: Orig OBJ Name       - Original source object name
  Col 9: Orig Type           - Object type (ipmask, etc.)
  Col 10: Orig IP            - Original source IP / subnet
  Col 11: Orig Comment       - Original source comment

[Destination Section (Dst)]
  Col 12: Dst Group OBJ      - Destination group name
  Col 13: Dst OBJ Name       - Destination object name
  Col 14: Dst Type           - Object type
  Col 15: Dst IP             - Destination IP / subnet
  Col 16: Dst Comment        - Destination comment

[NAT IP Pool & Status]
  Col 17: NAT IP Pool Name   - Target NAT IP Pool name
  Col 18: NAT Pool IP        - Translated NAT Pool IP range
  Col 19: NAT Pool Type      - IP Pool type (overload, one-to-one, etc.)
  Col 20: NAT                - NAT enabled status (enable / disable)
  Col 21: Comment            - NAT rule comments
```

### 4. DNAT (VIP) Sheet Column Structure (18 Columns)

```
[VIP Basic Information]
  Col 1: Seq                 - Sequence number
  Col 2: VDOM                - VDOM name
  Col 3: Name                - VIP object name
  Col 4: Type                - VIP type (static-nat, server-load-balance, etc.)

[External & Mapped IP / Interface]
  Col 5: External IP         - External listening IP (extip)
  Col 6: Mapped IP           - Internal mapped IP (mappedip)
  Col 7: External Interface  - External interface (extintf)

[Port Forwarding Settings]
  Col 8: Port Forward        - Port forwarding status (enable / disable)
  Col 9: Protocol            - Protocol (tcp / udp / sctp, etc.)
  Col 10: External Port      - External port range (extport)
  Col 11: Mapped Port        - Internal mapped port range (mappedport)

[Server Load Balancing (SLB)]
  Col 12: Server Type        - Server type (http, https, ip, etc.)
  Col 13: LDB Method         - Load balancing method (round-robin, weighted, etc.)
  Col 14: Monitor            - Health check monitor name
  Col 15: Real Server IP     - Real server backend IP
  Col 16: Real Server Port   - Real server service port
  Col 17: Real Server Weight - Real server weight

[General]
  Col 18: Comment            - VIP object comment
```

### 5. DoS Policy Sheet Column Structure (26 Columns)

```
[Policy Basic Information]
  Col 1: Seq                 - Sequence number
  Col 2: VDOM                - VDOM name
  Col 3: Enable              - Enabled status (Y / N)
  Col 4: ID                  - DoS policy ID
  Col 5: Interface           - Protected interface

[Source Section (Src)]
  Col 6: Src Group OBJ       - Source group name
  Col 7: Src OBJ Name        - Source object name
  Col 8: Src Type            - Address object type
  Col 9: Src IP              - Source IP / subnet range
  Col 10: Src Comment        - Source object comment

[Destination Section (Dst)]
  Col 11: Dst Group OBJ      - Destination group name
  Col 12: Dst OBJ Name       - Destination object name
  Col 13: Dst Type           - Address object type
  Col 14: Dst IP             - Destination IP / subnet range
  Col 15: Dst Comment        - Destination object comment

[Service Section (Service)]
  Col 16: Svc Group OBJ      - Service group name
  Col 17: Svc OBJ Name       - Service object name
  Col 18: Svc Port           - Service port number
  Col 19: Svc Comment        - Service object comment

[Anomaly Detection]
  Col 20: Anomaly Name       - Attack / anomaly pattern name (e.g., tcp_syn_flood)
  Col 21: Anomaly Status     - Anomaly defense status (enable / disable)
  Col 22: Anomaly Log        - Logging status (enable / disable)
  Col 23: Anomaly Quarant    - Quarantine action (attacker / disable)
  Col 24: Anomaly Action     - Mitigation action (pass / block / disable)
  Col 25: Anomaly Threshold  - Trigger threshold (packets/second)

[Policy Comment]
  Col 26: Comment            - DoS policy comments
```

### 6. ACL Policy Sheet Column Structure (20 Columns)

```
[Policy Basic Info]
  Col 1: Seq                 - Sequence number
  Col 2: VDOM                - VDOM name
  Col 3: Enable              - Enabled status (Y / N, N and dark gray shaded if disabled)
  Col 4: ID                  - ACL Rule ID (edit number)
  Col 5: Interface           - Inbound receiving interface

[Source Columns (Src) - Blue Style]
  Col 6: Src Group OBJ       - Source group name (Recursively resolved when group objects are present)
  Col 7: Src OBJ Name        - Source object name
  Col 8: Src Type            - Address object type (ipmask / geography, etc.)
  Col 9: Src IP              - Resolved source IP subnet / range
  Col 10: Src Comment        - Source object comment

[Destination Columns (Dst) - Red Style]
  Col 11: Dst Group OBJ      - Destination group name
  Col 12: Dst OBJ Name       - Destination object name
  Col 13: Dst Type           - Address object type
  Col 14: Dst IP             - Resolved destination IP subnet / range
  Col 15: Dst Comment        - Destination object comment

[Service Columns (Service)]
  Col 16: Svc Group OBJ      - Service group name
  Col 17: Svc OBJ Name       - Service object name
  Col 18: Svc Port           - Service port
  Col 19: Svc Comment        - Service object comment

[Policy Comment]
  Col 20: Comment            - ACL policy comment
※ Multiple address objects and group members configured under firewall acl are completely expanded into individual rows with common columns merged.
```

### 7. Static Route Sheet Column Structure (11 Columns)

```
[Routing Basic Info]
  Col 1: Seq                 - Sequence number
  Col 2: VDOM                - VDOM name
  Col 3: Enable              - Enabled status (Y / N, N and dark gray shaded if disabled)
  Col 4: ID                  - Static Route ID (edit number)

[Destination & Path]
  Col 5: Destination         - Destination network (Auto-converted to IP/CIDR, 0.0.0.0/0 if omitted)
  Col 6: Gateway             - Gateway IP address
  Col 7: Interface           - Outbound interface (device)

[Metric & Options]
  Col 8: Distance            - Administrative distance (Default: 10)
  Col 9: Priority            - Route priority
  Col 10: Options            - Summary of special flags (Blackhole, Dynamic-GW, BFD, Link-Mon-Exempt)
  Col 11: Comment            - Route comment
```

### 8. Policy Route Sheet Column Structure (21 Columns)

```
[Policy Route Basic Info]
  Col 1: Seq                 - Sequence number
  Col 2: VDOM                - VDOM name
  Col 3: Enable              - Enabled status (Y / N, N and dark gray shaded if disabled)
  Col 4: ID                  - Policy Route ID
  Col 5: Incoming Intf       - Inbound receiving interface (input-device)
  Col 6: Outgoing Intf       - Outbound transmitting interface (output-device)
  Col 7: Gateway             - Gateway IP address
  Col 8: Action              - permit / deny
  Col 9: Protocol            - ALL / TCP(6) / UDP(17) / ICMP(1), etc.
  Col 10: Port Range         - Port range (start-port ~ end-port, ALL if omitted)

[Source Area (Src)]
  Col 11: Src Group OBJ      - Source group name
  Col 12: Src OBJ Name       - Detailed object name (Address / Group / Direct subnet)
  Col 13: Src Type           - Object type (ipmask / fqdn, etc.)
  Col 14: Src IP             - Actual source IP / Subnet (Recursively resolved)
  Col 15: Src Comment        - Source object comment

[Destination Area (Dst)]
  Col 16: Dst Group OBJ      - Destination group name
  Col 17: Dst OBJ Name       - Detailed object name
  Col 18: Dst Type           - Object type
  Col 19: Dst IP             - Actual destination IP / Subnet (Recursively resolved)
  Col 20: Dst Comment        - Destination object comment

[Other]
  Col 21: Comment            - Policy route comment
```

### 9. OSPF Sheet Structure (5 Structured Section Tables)

```
[1. OSPF Global Settings]
  - vDOM, Router ID, Total Networks, Total Interfaces, Total Areas
[2. OSPF Networks (config network)]
  - Seq, ID, Prefix (IP/CIDR), Area ID
[3. OSPF Interfaces (config ospf-interface)]
  - Seq, Name, Interface, Cost, Dead / Hello Interval, Network Type, Priority (Defaults to 1), Authentication (Defaults to none)
[4. OSPF Redistribution (config redistribute)]
  - Seq, Protocol (Connected, Static, RIP, BGP, ISIS), Status, Route-Map, Metric, Metric Type
[5. Route-Map & Filter Details (config router route-map / access-list)]
  - Seq, Route-Map, Rule, Route-Map Action (PERMIT/DENY), Match Target, Filtered Prefix (Per-rule row splitting), Action (permit/deny), Exact Match (enable/disable), ACL Comment
※ Worksheets with unconfigured OSPF display header rows only and have their Excel tab automatically highlighted in Red.
※ Section title banner fills strictly match the width of each underlying section table (5, 4, 8, 6, 9 columns) to ensure clean visual presentation.
```

### 10. IPsec VPN Sheet Column Structure (30 Columns)

```
[Common Info - Default Header (#2F5496)]
  Col 1: Seq                 - Sequence number (Center aligned & merged)
  Col 2: vDOM                - VDOM name (Center aligned & merged)

[Phase 1 Tunnel & Network/Advanced Settings - Navy Blue Header (#1F4E78)]
  Col 3: P1 Name             - Phase 1 tunnel name (Left aligned & merged)
  Col 4: Interface           - Physical bound interface (Center aligned & merged)
  Col 5: Remote Gateway      - Remote peer public IP / (Dialup/Dynamic) (Center aligned & merged)
  Col 6: Local Gateway       - Local binding IP (Center aligned & merged)
  Col 7: IKE Version         - IKE version (v1 / v2) (Center aligned & merged)
  Col 8: P1 Proposal         - Phase 1 encryption/integrity proposals (Left aligned & merged)
  Col 9: P1 DH Group         - Phase 1 Diffie-Hellman groups (Defaults to '14 5' if omitted in config)
  Col 10: NAT Traversal      - NAT Traversal mode (Defaults to enable)
  Col 11: Keepalive Frequency- Keepalive interval in seconds (Defaults to 10)
  Col 12: Dead Peer Detection- DPD mode (disable / on-idle / on-demand, Defaults to on-demand)
  Col 13: DPD Retry Count    - DPD retry count (Defaults to 3)
  Col 14: DPD Retry Interval - DPD retry interval in seconds (Number only, Defaults to 20)
  Col 15: FEC Egress         - Forward Error Correction Egress (Defaults to disable)
  Col 16: FEC Ingress        - Forward Error Correction Ingress (Defaults to disable)
  Col 17: Add Route          - Add gateway route automatically (add-gw-route, Defaults to enable)
  Col 18: Auto Discovery Sender   - Dynamic tunnel sender discovery (Defaults to disable)
  Col 19: Auto Discovery Receiver - Dynamic tunnel receiver discovery (Defaults to disable)
  Col 20: Exchange Interface IP   - Exchange interface IP (Defaults to disable)
  Col 21: Device Creation    - Tunnel virtual interface creation (net-device, Defaults to disable)
  Col 22: P1 Comment         - Phase 1 comment (Left aligned & merged)

[Phase 2 Child Tunnel Info - Dark Teal Header (#2A5C5A)]
  Col 23: P2 Name            - Phase 2 child tunnel name
  Col 24: P2 Proposal        - Phase 2 encryption/integrity proposals
  Col 25: P2 DH Group        - Phase 2 Diffie-Hellman groups (Defaults to '14 5' if omitted in config)
  Col 26: Local Subnet / Src - Local subnet or address object (Defaults to '0.0.0.0/0' if omitted)
  Col 27: Remote Subnet / Dst- Remote subnet or address object (Defaults to '0.0.0.0/0' if omitted)
  Col 28: Auto Negotiate     - Auto-negotiation status (enable / disable)
  Col 29: Keepalive          - Keepalive status (enable / disable)
  Col 30: P2 Comment         - Phase 2 comment
※ Columns with short values but long titles use two-line headers to optimize column widths and readability.
```

### 11. Network Interface Sheet Column Structure (16 Columns)

```
[Interface Basic Info]
  Col 1: Seq                 - Sequence number
  Col 2: vDOM                - VDOM name
  Col 3: Status              - Link status (UP / DOWN, highlighted in red if DOWN)
  Col 4: Name                - Interface name (Bold)
  Col 5: Alias               - Interface alias
  Col 6: Type                - Interface type (physical, vlan, loopback, aggregate, tunnel, etc.)

[Network Address & Hierarchy Binding]
  Col 7: Primary IP          - Primary assigned IP address and subnet mask (Defaults to '0.0.0.0/0' if unconfigured, Blue font)
  Col 8: Secondary IP        - Secondary IP subnet ranges (Parsed from config secondaryip, multi-line)
  Col 9: Remote IP (Tunnel)  - Remote tunnel peer IP (Parsed from set remote-ip for tunnel interfaces)
  Col 10: VLAN ID            - VLAN tag ID
  Col 11: Parent / Member Interface - Physical parent interface (VLAN/Tunnel) or member interfaces (Aggregate/Redundant)
  Col 12: VRF                - Assigned VRF ID (Defaults to '0' if omitted)
  Col 13: Addressing Mode    - IP assignment mode (static, dhcp, pppoe, etc.)

[Security & Management]
  Col 14: Administrative Access - Allowed management access protocols (ping, https, ssh, snmp, fgfm, etc.)
  Col 15: Speed / Duplex     - Port speed & duplex settings (1000 / auto, 10000 / full, etc. formatted for readability)
  Col 16: Description        - Interface description / comment
```

### 12. External Resource Sheet Column Structure (9 Columns)

```
  Col 1: Seq                 - Sequence number
  Col 2: VDOM                - VDOM name
  Col 3: Enable              - Enabled status (Y / N, N if set status disable)
  Col 4: Name                - External resource name (e.g., AbuseIPDB_Blacklist_Score-75)
  Col 5: Type                - Resource type (address / domain / malware, etc.)
  Col 6: Resource URL        - External feed download URL
  Col 7: Refresh Rate (min)  - Refresh interval in minutes
  Col 8: Source IP           - Source IP interface used for external connection
  Col 9: Comment             - Resource comments
```

---

## 📁 Folder and File Structure

```
fortigate_policy_to_excel/
│
├── fortigate_policy_to_excel.py   # [Core] Unified script containing conversion engine & GUI
├── README_ko.md                   # [Docs] Korean user manual & guide
├── README_en.md                   # [Docs] English user manual & guide
├── Start.bat                      # [Run] Windows quick-start batch script
│
└── <hostname>/                    # [Output] Output directory created upon conversion
    ├── _TOTAL_SUMMARY.xlsx        # Master summary of total policies, routing, VPN, interfaces & objects per VDOM
    ├── root.xlsx                  # root VDOM workbook
    └── ...                        # Independent Excel workbooks for each VDOM
```

---

## ❓ Frequently Asked Questions (FAQ & Troubleshooting)

### Q1. Error: "Python executable not found"
* **Cause**: Python is not installed on your PC, or the "Add Python to PATH" option was not checked during installation.
* **Solution**: Re-run the Python installer from the official website, choose **`Modify`**, check **`Add Python to PATH`**, and complete the installation.

### Q2. Error: "Requires openpyxl: pip install openpyxl"
* **Cause**: The Excel handling library is not installed.
* **Solution**: Open Command Prompt (CMD) or PowerShell and run `pip install openpyxl`.

### Q3. Does text look blurry on high-resolution displays?
* The application features built-in Windows **Per-Monitor High-DPI awareness** and **ClearType rendering**, delivering razor-sharp text on QHD and 4K displays even under custom scaling (125%, 150%, 175%).

### Q4. Can it process large configuration files with over 100,000 lines?
* Yes. Parsing and Excel generation run in background threads (`threading.Thread`), enabling configuration files with 120,000+ lines from enterprise firewalls to be exported reliably within seconds without freezing the UI (Not Responding).

---
