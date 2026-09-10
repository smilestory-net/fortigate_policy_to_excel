# FortiGate Policy to Excel Exporter v1.0

A tool that parses FortiGate firewall backup configuration files (`.conf` or `.txt`), analyzes policies and objects per VDOM, and automatically exports them into highly readable, partitioned Excel (`.xlsx`) workbooks.

---

## 📌 Table of Contents
1. [Key Features](#-key-features)
2. [System Requirements](#-system-requirements)
3. [Python & Dependency Installation](#-python--dependency-installation)
4. [How to Run](#-how-to-run)
5. [Excel Sheet Structure & Design Highlights](#-excel-sheet-structure--design-highlights)
6. [Directory & File Structure](#-directory--file-structure)
7. [Frequently Asked Questions (FAQ & Troubleshooting)](#-frequently-asked-questions-faq--troubleshooting)

---

## 🌟 Key Features

### 1. Full Extraction of 5 Core Policy Types into Separate Sheets
* **Firewall Policy**: Standard firewall policies (Action, Schedule, NAT, IP Pool, UTM Profile, Log, etc.).
* **Local-in Policy**: Firewall device self-access control policies.
* **Central-NAT**: Central SNAT Map (Original IP/Port ↔ Translated IP/Port mapping).
* **DNAT (VIP)**: Virtual IP and port-forwarding mapping (External IP/Port ↔ Mapped IP/Port).
* **DoS Policy**: Denial of Service (DoS/DDoS) anomaly defense policies.

### 2. Multi-Row Flattening of Objects and Groups with Real IPs, Ports & Comments
* Rather than simply displaying raw object or group names, this tool recursively resolves every address/service group to reveal **member names, actual IP addresses, subnet prefixes, and port numbers** row-by-row.
* **Unified IP Format**:
  * Single host: `x.x.x.x/32`
  * Subnet network: `x.x.x.x/24`
  * IP range: `1.1.1.1-1.1.1.10`
* **Isolated Object Comments**: Object comments are extracted and populated into dedicated `Src Comment`, `Dst Comment`, and `Svc Comment` columns, completely separated from policy comments (`Comments`).

### 3. Smart Vertical Cell Merging
* When a policy spans multiple rows due to multiple source/destination/service members, **common policy attributes (Seq, VDOM, Enable, ID, Name, Action, Schedule, etc.)** and **identical group member spans** are merged vertically for maximum readability.

### 4. Enterprise-Grade Visual Styling
* **Header & Text Color Hierarchy**:
  * Source (Src) Area: Soft blue header fill + blue text.
  * Destination (Dst) Area: Soft red header fill + red text.
  * Service (Service) Area: Blue-gray header fill.
* **Zebra Striping**: Alternating odd/even row fills across all columns ensure comfortable visual scanning.
* **Disabled Policy Highlighting**: Disabled policies (`Enable == N`) are shaded with a darker gray background, allowing instant distinction from active policies.

### 5. Hostname Directory Creation & Per-VDOM Excel Splitting
* Eliminates workbook lag and freeze issues caused by packing all policies into a single giant Excel file.
* Automatically creates an output directory named after the firewall hostname (e.g., `JBNU_SVF-FW1`), and generates **individual Excel files for each VDOM (`<vdom_name>.xlsx`)**.
* Simultaneously creates a comprehensive master summary workbook (**`_TOTAL_SUMMARY.xlsx`**) aggregating policy counts and object totals across all VDOMs.

### 6. Built-in GUI
* Intuitive file explorer browser (`Browse...`) provided.
* Real-time progress bar and syntax-highlighted terminal log console.
* Native Windows **High-DPI (Per-Monitor v2) & ClearType** support for crisp, razor-sharp font rendering on QHD and 4K monitors.
* Automatic result folder launch upon completion without intrusive modal popups.

---

## 💻 System Requirements

| Item | Requirement / Recommendation |
|---|---|
| **Operating System** | Windows 10 / Windows 11 (64-bit) *(macOS / Linux supported via CLI mode)* |
| **Python Version** | **Python 3.8 or higher** (Python 3.10 ~ 3.13 fully supported) |
| **Display** | 1920×1080 (FHD) or higher (QHD, 4K display scaling fully supported) |
| **Required Library** | `openpyxl` |

---

## 📦 Python & Dependency Installation

### 1. Python Installation
1. Download Python 3.10+ from the official [Python website](https://www.python.org/downloads/).
2. On the first installation screen, make sure to check the **`[✔] Add python.exe to PATH`** checkbox before proceeding with the installation.

### 2. Installing Required Dependencies (`openpyxl`)
Open Command Prompt (CMD) or PowerShell and run:

```bash
pip install openpyxl
```

> **Note**: GUI and system modules such as `tkinter`, `threading`, `ctypes`, and `re` are included in Python's standard library and require no extra installation.

---

## 🚀 How to Run

### Method 1. Double-Click Launcher (Recommended for Windows)
1. Double-click the **`실행하기.bat`** file in the project folder.
2. The dark theme GUI window will open.
3. Click **[Browse...]** to select your FortiGate `.conf` (or `.txt`) file.
4. Click the green **`[▶ 엑셀 변환 실행 (Start Conversion)]`** button.
5. Once conversion is complete, the destination folder opens automatically.

### Method 2. Running via Python (GUI Mode)
Run the script without arguments in your terminal to launch the GUI:

```bash
python fortigate_policy_to_excel.py
```

### Method 3. Command-Line Interface (CLI Mode / Batch Automation)
Run headless in CLI mode without opening a GUI window:

```bash
# Syntax: python fortigate_policy_to_excel.py <config_file_path> [output_directory]
python fortigate_policy_to_excel.py "C:\backup\my_firewall.conf"
```

---

## 📊 Excel Sheet Structure & Design Highlights

### Firewall Policy Sheet Columns (32 Columns Total)

```
[Common Policy Attributes]
  Col 1: Seq            - Sequence number
  Col 2: VDOM           - VDOM name
  Col 3: Enable         - Enabled status (Y / N)
  Col 4: ID             - Policy ID
  Col 5: Name           - Policy name
  Col 6: Action         - accept / deny

[Source (Src) Area - Blue Header & Font]
  Col 7: Src Interface  - Source interface
  Col 8: Src Group OBJ  - Source group name
  Col 9: Src OBJ Name   - Resolved object name
  Col 10: Src Type      - ipmask / iprange / fqdn, etc.
  Col 11: Src IP        - Actual IP / Subnet prefix / Range
  Col 12: Src Comment   - Source object comment

[Destination (Dst) Area - Red Header & Font]
  Col 13: Dst Interface - Destination interface
  Col 14: Dst Group OBJ - Destination group name
  Col 15: Dst OBJ Name  - Resolved object name
  Col 16: Dst Type      - ipmask / iprange / fqdn, etc.
  Col 17: Dst IP        - Actual IP / Subnet prefix / Range
  Col 18: Dst Comment   - Destination object comment

[Service (Service) Area - Blue-Gray Header]
  Col 19: Svc Group OBJ - Service group name
  Col 20: Svc OBJ Name  - Resolved service object name
  Col 21: Svc Port      - Port definition (IP/ALL normalized to ALL)
  Col 22: Svc Comment   - Service object comment

[Security Profiles & Settings]
  Col 23: Schedule      - Schedule configuration
  Col 24: NAT           - NAT status (enable / disable)
  Col 25: IP Pool       - IP Pool status
  Col 26: Pool Name     - IP Pool name
  Col 27: Pool IP       - Assigned IP pool address range
  Col 28: UTM Status    - UTM profile active status
  Col 29: SSL/SSH Profile
  Col 30: IPS Sensor
  Col 31: Log Traffic   - all / utm / disable
  Col 32: Comments      - Original policy comment
```

---

## 📁 Directory & File Structure

```
fortigate_policy_to_excel/
│
├── fortigate_policy_to_excel.py   # [Core] All-in-one parsing engine and GUI script
├── README.md                      # [Documentation] User manual and technical reference
│
└── <Hostname>/                    # [Output] Directory generated upon conversion
    ├── _TOTAL_SUMMARY.xlsx        # Master summary of all VDOM policies & object counts
    ├── root.xlsx                  # root VDOM workbook (5 policy types in separate sheets)
    ├── DMZ.xlsx                   # DMZ VDOM workbook
    ├── IDC.xlsx                   # IDC VDOM workbook
    └── ...                        # Individual workbooks for each VDOM
```

---

## ❓ Frequently Asked Questions (FAQ & Troubleshooting)

### Q1. Running `실행하기.bat` outputs "Python 실행 경로를 찾을 수 없습니다" (Cannot find Python path).
* **Cause**: Python is not installed, or the "Add Python to PATH" option was not checked during installation.
* **Solution**: Re-run the Python installer from [python.org](https://www.python.org/), choose **`Modify`**, check the **`Add Python to PATH`** checkbox, and finish installation.

### Q2. "openpyxl 필요: pip install openpyxl" error appears.
* **Cause**: The Excel manipulation library `openpyxl` is missing.
* **Solution**: Open Command Prompt (CMD) or PowerShell and run `pip install openpyxl`.

### Q3. Does the application look blurry on high-resolution monitors?
* No. The application features built-in Windows **Per-Monitor High-DPI (v2) and ClearType subpixel rendering**. Even on QHD or 4K displays with 125%, 150%, or 175% scaling, all fonts, icons, and borders render natively with razor-sharp clarity.

### Q4. Can it handle very large config files (> 100,000 lines)?
* Yes. Parsing and Excel exports run on an asynchronous background thread (`threading.Thread`). Config files exceeding 120,000 lines are parsed and generated in seconds without UI freezing or "(Not Responding)" states.
