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

### 1. Comprehensive Extraction & Dedicated Sheets for 5 Core Policies
* **Firewall Policy**: Standard firewall policies (Action, Schedule, NAT, IP Pool, UTM Profile, Log, etc.)
* **Local-in Policy**: Firewall appliance self-access control policies
* **Central-NAT**: Central SNAT Map (Original IP/Port ↔ Translated IP/Port mapping)
* **DNAT (VIP)**: Virtual IP and port-forwarding mapping (External IP/Port ↔ Mapped IP/Port)
* **DoS Policy**: Denial of Service (DoS/DDoS) defense policies

### 2. Multi-row Object/Group Expansion with Real IPs, Ports & Comments (Multi-row Flattening)
* Unlike conventional tools that simply display object or group names, this tool automatically traces **member objects, actual IP addresses, subnet ranges, and port numbers** within groups and expands them row by row.
* **Unified IP Formatting**:
  * Single Host: `x.x.x.x/32`
  * Subnets (e.g., Class C): `x.x.x.x/24`
  * IP Range: `1.1.1.1-1.1.1.10`
* **Isolated Object Comments**: Extracts comments from each object and group, placing them into dedicated `Src Comment`, `Dst Comment`, and `Svc Comment` columns, completely separating them from policy-level comments (`Comments`).

### 3. Smart Vertical Cell Merging
* When a policy spans multiple rows due to multiple objects, **common policy attributes (Seq, VDOM, Enable, ID, Name, Action, Schedule, etc.)** and **identical group member areas** are automatically merged vertically to maximize readability.

### 4. Enterprise-Grade Visual Styling
* **Header & Font Color Coding**:
  * Source (Src) section: Blue header + blue font
  * Destination (Dst) section: Red header + red font
  * Service section: Blue-gray header
* **Zebra Striping**: Alternating row shading across all columns ensures smooth visual scanning.
* **Disabled Policy Shading**: Disabled policies (`Enable == N`) are highlighted with a dark gray background, making them instantly distinguishable from active rules.

### 5. Hostname Directory Creation & Per-VDOM Excel Splitting
* Prevents performance lag and freezing typically caused by dumping large-scale configuration policies into a single spreadsheet.
* Automatically creates a folder named after the device hostname (e.g., `JBNU_SVF-FW1`), and generates **independent Excel workbooks (`<vdom_name>.xlsx`) for each VDOM**.
* Provides a master summary file (**`_TOTAL_SUMMARY.xlsx`**) summarizing policy and object counts across all VDOMs at a glance.

### 6. Built-in GUI
* Intuitive file browser dialog (`Select File...`)
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

### Firewall Policy Sheet Column Structure (32 Columns)

```
[Policy Basic Information]
  Col 1: Seq            - Sequence number
  Col 2: VDOM           - VDOM name
  Col 3: Enable         - Enabled status (Y / N)
  Col 4: ID             - Policy ID
  Col 5: Name           - Policy name
  Col 6: Action         - accept / deny

[Source Section (Src) - Blue header & font]
  Col 7: Src Interface  - Source interface
  Col 8: Src Group OBJ  - Source group name
  Col 9: Src OBJ Name   - Member object name
  Col 10: Src Type      - ipmask / iprange / fqdn, etc.
  Col 11: Src IP        - Actual IP / Subnet / Range
  Col 12: Src Comment   - Source object comment

[Destination Section (Dst) - Red header & font]
  Col 13: Dst Interface - Destination interface
  Col 14: Dst Group OBJ - Destination group name
  Col 15: Dst OBJ Name  - Member object name
  Col 16: Dst Type      - ipmask / iprange / fqdn, etc.
  Col 17: Dst IP        - Actual IP / Subnet / Range
  Col 18: Dst Comment   - Destination object comment

[Service Section (Service) - Blue-gray header]
  Col 19: Svc Group OBJ - Service group name
  Col 20: Svc OBJ Name  - Service object name
  Col 21: Svc Port      - Port number (IP/ALL displayed as ALL)
  Col 22: Svc Comment   - Service object comment

[General & Security Profiles]
  Col 23: Schedule      - Schedule
  Col 24: NAT           - NAT status (enable / disable)
  Col 25: IP Pool       - IP Pool status
  Col 26: Pool Name     - Pool name
  Col 27: Pool IP       - Assigned Pool IP range
  Col 28: UTM Status    - UTM applied status
  Col 29: SSL/SSH Profile
  Col 30: IPS Sensor
  Col 31: Log Traffic   - all / utm / disable
  Col 32: Comments      - Policy-level comments
```

---

## 📁 Folder and File Structure

```
fortigate_policy_to_excel/
│
├── fortigate_policy_to_excel.py   # [Core] Unified script containing the conversion engine and GUI
├── README.md                      # [Docs] User manual and guide
│
└── <hostname>/                    # [Output] Output directory created upon conversion
    ├── _TOTAL_SUMMARY.xlsx        # Master summary of total policies and objects per VDOM
    ├── root.xlsx                  # root VDOM workbook (individual sheets for 5 core policies)
    ├── DMZ.xlsx                   # DMZ VDOM workbook
    ├── IDC.xlsx                   # IDC VDOM workbook
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
