#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
FortiGate Configuration -> Excel Exporter & GUI (Unified Single-File Edition)
=============================================================================
FortiGate 방화벽 설정 파일(.conf)을 분석하여:
  1. 호스트네임 디렉토리 생성 및 각 vDOM별 엑셀 파일(<vdom_name>.xlsx) 분할 생성
  2. 전체 vDOM 총괄 요약 파일(_TOTAL_SUMMARY.xlsx) 동시 생성
  3. 10대 핵심 정책, 라우팅(Static/Policy/OSPF) 및 IPsec VPN 개별 시트 완벽 분리 수록
  4. 객체/그룹의 실제 IP, 서브넷, 포트, 코멘트를 다중 행 전개 및 스마트 셀 세로 병합(Merge)
  5. Svc Protocol 컬럼 삭제, Svc Port 'ALL' 표기, Src/Dst/Svc Comment 분리 수록
  6. 출발지(파랑), 목적지(빨강) 가독성 컬러 스타일링 및 비활성화 정책(진한 회색) 음영 처리
  7. 모던 다크 테마 GUI 및 커맨드라인(CLI) 모드 완벽 통합 지원


Parses FortiGate firewall backup configuration files (.conf / .txt) to:
  1. Create a hostname-based directory with partitioned Excel files per vDOM (<vdom_name>.xlsx)
  2. Simultaneously generate a master summary workbook (_TOTAL_SUMMARY.xlsx) across all vDOMs
  3. Fully extract 10 core policies, routing (Static/Policy/OSPF), and IPsec VPN into dedicated sheets
  4. Recursively resolve objects/groups to actual IPs/ports/comments with multi-row flattening & cell merging
  5. Optimize service ports ('ALL') and provide dedicated Src/Dst/Svc Comment columns
  6. Apply professional visual styling (Src blue, Dst red, zebra striping, disabled policy shading)
  7. Provide an integrated modern Dark Theme GUI and headless CLI execution in a single file

Usage / 사용법:
  - GUI Mode: python fortigate_policy_to_excel.py (Run without args or double-click / 인자 없이 실행)
  - CLI Mode: python fortigate_policy_to_excel.py <config_file> [output_dir]
"""

import sys
import os
import re
import threading
import subprocess
import ctypes
import webbrowser
import base64
import tempfile
from collections import OrderedDict
from datetime import datetime

# Windows 고해상도(High-DPI) 화면에서 흐림 방지 및 선명한 ClearType 렌더링 활성화 / Enable Windows High-DPI (Per-Monitor v2) & ClearType rendering
if sys.platform == 'win32':
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
    except Exception:
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            try:
                ctypes.windll.user32.SetProcessDPIAware()
            except Exception:
                pass

try:
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox, scrolledtext
    HAS_TKINTER = True
except ImportError:
    HAS_TKINTER = False

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.cell_range import CellRange
except ImportError:
    print("openpyxl 필요: pip install openpyxl")
    sys.exit(1)


# ================================================================
#  1. 유틸리티 / Utilities
# ================================================================

_QUOTED_RE = re.compile(r'"([^"]*)"')


def parse_quoted_values(line):
    """'set field "val1" "val2" ...' -> ["val1", "val2"]"""
    parts = line.strip().split(None, 2)
    if len(parts) < 3:
        return []
    rest = parts[2]
    if '"' not in rest:
        return rest.split()
    quoted = _QUOTED_RE.findall(rest)
    if quoted:
        return quoted
    return rest.split()


def parse_set_value(line):
    """'set field value' -> value (따옴표 제거 / Strip surrounding quotes)"""
    parts = line.strip().split(None, 2)
    if len(parts) < 3:
        return ""
    val = parts[2].strip()
    if val.startswith('"') and val.endswith('"'):
        val = val[1:-1]
    return val


def get_field_name(line):
    parts = line.strip().split(None, 2)
    return parts[1] if len(parts) >= 2 else ""


def parse_edit_id(line):
    parts = line.strip().split(None, 1)
    if len(parts) >= 2:
        val = parts[1].strip()
        if val.startswith('"') and val.endswith('"'):
            val = val[1:-1]
        return val
    return ""


def mask_to_prefix(mask_str):
    """255.255.255.0 -> 24 (서브넷 마스크를 CIDR Prefix 숫자로 변환 / Convert subnet mask to CIDR prefix)"""
    try:
        parts = mask_str.split('.')
        bits = ''.join(format(int(p), '08b') for p in parts)
        return bits.count('1')
    except:
        return 32


def format_subnet(ip, mask):
    """IP와 마스크를 기반으로 (IP, Prefix) 반환 / Return (IP, prefix) tuple from IP and subnet mask"""
    prefix = mask_to_prefix(mask)
    return ip, str(prefix)


def format_ip_str(obj):
    """
    객체 타입별 IP 문자열 단일 포맷팅:
      - 단일 호스트: 192.168.1.1/32
      - 네트워크 서브넷 대역: 192.168.10.0/24, 0.0.0.0/0 등
      - IP 범위 (Range): 1.1.1.1-1.1.1.10
      - FQDN, 지리(국가), 동적 객체 등

    Unify IP display format based on object type:
      - Single host: 192.168.1.1/32
      - Subnet network: 192.168.10.0/24, 0.0.0.0/0, etc.
      - IP Range: 1.1.1.1-1.1.1.10
      - FQDN, Geography, Dynamic, etc.
    """
    if not obj:
        return ""
    t = obj.get('type', '')
    if t == 'ipmask':
        ip = obj.get('ip', '')
        prefix = obj.get('prefix', '')
        if ip:
            if prefix != '':
                return f"{ip}/{prefix}"
            return f"{ip}/32"
        return obj.get('display', '')
    elif t == 'iprange':
        start = obj.get('start-ip', '')
        end = obj.get('end-ip', '')
        if start and end:
            if start == end:
                return f"{start}/32"
            return f"{start}-{end}"
        return obj.get('display', '')
    elif t == 'fqdn':
        return obj.get('fqdn', '')
    elif t == 'wildcard-fqdn':
        return obj.get('wildcard-fqdn', '')
    elif t == 'geography':
        return f"Country: {obj.get('country', '')}"
    elif t == 'dynamic':
        return "(dynamic)"
    elif t == 'ipv6':
        return obj.get('ip', '') or '(IPv6)'
    return obj.get('display', '')


def parse_hostname(lines, default_name="FortiGate"):
    """config 파일에서 방화벽 hostname 추출 / Extract firewall hostname from config file"""
    for line in lines[:300]:
        s = line.strip()
        if s.startswith("set hostname "):
            val = parse_set_value(line)
            clean_name = re.sub(r'[\\/*?:"<>|]', "", val).strip()
            if clean_name:
                return clean_name
    return default_name


# ================================================================
#  2. vDOM 경계 파싱 / vDOM Boundary Parsing
# ================================================================

def find_vdom_boundaries(lines):
    vdom_config_starts = [i for i, l in enumerate(lines) if l.strip() == "config vdom"]
    n = len(lines)
    if len(vdom_config_starts) < 2:
        return [("root", 0, n - 1)]

    sections = []
    for i in range(1, len(vdom_config_starts)):
        start = vdom_config_starts[i]
        vdom_name = parse_edit_id(lines[start + 1]) if start + 1 < n else ""
        end = vdom_config_starts[i + 1] - 1 if i + 1 < len(vdom_config_starts) else n - 1
        sections.append((vdom_name, start, end))
    return sections


def find_section_range(lines, start, end, section_name):
    results = []
    i = start
    target = f"config {section_name}"
    while i <= end:
        if lines[i].strip() == target:
            sec_start = i
            depth = 1
            j = i + 1
            while j <= end and depth > 0:
                s = lines[j].strip()
                if s.startswith("config "):
                    depth += 1
                elif s == "end":
                    depth -= 1
                j += 1
            results.append((sec_start, j - 1))
            i = j
        else:
            i += 1
    return results


# ================================================================
#  3. 객체 파서 / Object Parsers (Address, AddrGrp, Service, SvcGrp, IPPool)
# ================================================================

def parse_address_objects(lines, vdom_start, vdom_end):
    """config firewall address -> dict[name] = {type, ip, prefix, comment, ...}"""
    addrs = {}
    builtin = {
        'all': {'type': 'ipmask', 'ip': '0.0.0.0', 'prefix': '0', 'display': '0.0.0.0/0', 'comment': ''},
        'none': {'type': 'ipmask', 'ip': '0.0.0.0', 'prefix': '32', 'display': '0.0.0.0/32', 'comment': ''},
        'FIREWALL_AUTH_PORTAL_ADDRESS': {'type': 'ipmask', 'ip': '', 'prefix': '', 'display': '(auth-portal)', 'comment': ''},
        'FABRIC_DEVICE': {'type': 'ipmask', 'ip': '', 'prefix': '', 'display': '(fabric-device)', 'comment': ''},
    }
    addrs.update(builtin)

    ranges = find_section_range(lines, vdom_start, vdom_end, "firewall address")
    for sr, er in ranges:
        i = sr + 1
        while i <= er:
            if lines[i].strip().startswith("edit "):
                name = parse_edit_id(lines[i])
                obj = {'type': 'ipmask', 'ip': '', 'prefix': '', 'fqdn': '',
                       'start-ip': '', 'end-ip': '', 'country': '',
                       'wildcard-fqdn': '', 'display': '', 'comment': ''}
                subnet_ip = ''
                subnet_mask = ''
                has_data = False
                i += 1
                while i <= er:
                    s = lines[i].strip()
                    if s in ("next", "end"):
                        break
                    if s.startswith("set "):
                        f = get_field_name(lines[i])
                        v = parse_set_value(lines[i])
                        if f == 'type':
                            obj['type'] = v
                            has_data = True
                        elif f == 'subnet':
                            parts = v.split()
                            if len(parts) == 2:
                                subnet_ip, subnet_mask = parts
                            elif '/' in v:
                                subnet_ip = v.split('/')[0]
                                subnet_mask = v.split('/')[1]
                            has_data = True
                        elif f == 'fqdn':
                            obj['fqdn'] = v
                            has_data = True
                        elif f == 'start-ip':
                            obj['start-ip'] = v
                            has_data = True
                        elif f == 'end-ip':
                            obj['end-ip'] = v
                            has_data = True
                        elif f == 'country':
                            obj['country'] = v
                            has_data = True
                        elif f == 'wildcard-fqdn':
                            obj['wildcard-fqdn'] = v
                            has_data = True
                        elif f in ('comment', 'comments'):
                            obj['comment'] = v
                    i += 1

                if name in builtin and not has_data:
                    continue

                if obj['type'] == 'ipmask' and subnet_ip:
                    ip, pf = format_subnet(subnet_ip, subnet_mask)
                    obj['ip'] = ip
                    obj['prefix'] = pf
                    obj['display'] = f"{ip}/{pf}" if pf else ip
                elif obj['type'] == 'fqdn':
                    obj['display'] = obj['fqdn']
                elif obj['type'] == 'iprange':
                    obj['ip'] = obj['start-ip']
                    obj['display'] = f"{obj['start-ip']}-{obj['end-ip']}"
                elif obj['type'] == 'geography':
                    obj['display'] = f"Country: {obj['country']}"
                elif obj['type'] == 'wildcard-fqdn':
                    obj['display'] = obj['wildcard-fqdn']
                elif obj['type'] == 'dynamic':
                    obj['display'] = '(dynamic)'
                elif not obj['display']:
                    if subnet_ip:
                        ip, pf = format_subnet(subnet_ip, subnet_mask)
                        obj['ip'] = ip
                        obj['prefix'] = pf
                        obj['display'] = f"{ip}/{pf}" if pf else ip

                addrs[name] = obj
            i += 1

    ranges6 = find_section_range(lines, vdom_start, vdom_end, "firewall address6")
    for sr, er in ranges6:
        i = sr + 1
        while i <= er:
            if lines[i].strip().startswith("edit "):
                name = parse_edit_id(lines[i])
                obj = {'type': 'ipv6', 'ip': '', 'prefix': '', 'display': '(IPv6)', 'comment': ''}
                i += 1
                while i <= er:
                    s = lines[i].strip()
                    if s in ("next", "end"):
                        break
                    if s.startswith("set "):
                        f = get_field_name(lines[i])
                        v = parse_set_value(lines[i])
                        if f == 'ip6':
                            obj['ip'] = v
                            obj['display'] = v
                        elif f in ('comment', 'comments'):
                            obj['comment'] = v
                    i += 1
                if name not in addrs and name not in builtin:
                    addrs[name] = obj
            i += 1
    return addrs


def parse_addrgrp_objects(lines, vdom_start, vdom_end):
    """config firewall addrgrp -> dict[name] = {'members': [...], 'comment': '...'}"""
    groups = {}
    for sec_name in ("firewall addrgrp", "firewall addrgrp6"):
        ranges = find_section_range(lines, vdom_start, vdom_end, sec_name)
        for sr, er in ranges:
            i = sr + 1
            while i <= er:
                if lines[i].strip().startswith("edit "):
                    name = parse_edit_id(lines[i])
                    members = []
                    comment = ""
                    i += 1
                    while i <= er:
                        s = lines[i].strip()
                        if s in ("next", "end"):
                            break
                        if s.startswith("set "):
                            f = get_field_name(lines[i])
                            if f == "member":
                                members = parse_quoted_values(lines[i])
                            elif f in ("comment", "comments"):
                                comment = parse_set_value(lines[i])
                        i += 1
                    groups[name] = {'members': members, 'comment': comment}
                i += 1
    return groups


def parse_service_objects(lines, vdom_start, vdom_end):
    """config firewall service custom -> dict[name] = {protocol, tcp_port, comment, ...}"""
    svcs = {}
    svcs['ALL'] = {'protocol': 'ALL', 'tcp_port': '', 'udp_port': '', 'display': 'ALL', 'comment': ''}
    svcs['ALL_TCP'] = {'protocol': 'TCP', 'tcp_port': '1-65535', 'udp_port': '', 'display': 'TCP/1-65535', 'comment': ''}
    svcs['ALL_UDP'] = {'protocol': 'UDP', 'tcp_port': '', 'udp_port': '1-65535', 'display': 'UDP/1-65535', 'comment': ''}
    svcs['ALL_ICMP'] = {'protocol': 'ICMP', 'tcp_port': '', 'udp_port': '', 'display': 'ICMP', 'comment': ''}
    svcs['ALL_ICMP6'] = {'protocol': 'ICMP6', 'tcp_port': '', 'udp_port': '', 'display': 'ICMP6', 'comment': ''}

    ranges = find_section_range(lines, vdom_start, vdom_end, "firewall service custom")
    for sr, er in ranges:
        i = sr + 1
        while i <= er:
            if lines[i].strip().startswith("edit "):
                name = parse_edit_id(lines[i])
                obj = {'protocol': 'TCP/UDP/SCTP', 'tcp_port': '', 'udp_port': '',
                       'sctp_port': '', 'icmptype': '', 'icmpcode': '',
                       'protocol_number': '', 'display': '', 'comment': ''}
                i += 1
                while i <= er:
                    s = lines[i].strip()
                    if s in ("next", "end"):
                        break
                    if s.startswith("set "):
                        f = get_field_name(lines[i])
                        v = parse_set_value(lines[i])
                        if f == 'protocol':
                            obj['protocol'] = v
                        elif f == 'tcp-portrange':
                            obj['tcp_port'] = v
                        elif f == 'udp-portrange':
                            obj['udp_port'] = v
                        elif f == 'sctp-portrange':
                            obj['sctp_port'] = v
                        elif f == 'icmptype':
                            obj['icmptype'] = v
                        elif f == 'icmpcode':
                            obj['icmpcode'] = v
                        elif f == 'protocol-number':
                            obj['protocol_number'] = v
                        elif f in ('comment', 'comments'):
                            obj['comment'] = v
                    i += 1

                parts = []
                if obj['protocol'] == 'ICMP':
                    parts.append('ICMP')
                    if obj['icmptype']:
                        parts[-1] += f" type={obj['icmptype']}"
                elif obj['protocol'] == 'ICMP6':
                    parts.append('ICMP6')
                elif obj['protocol'] == 'IP':
                    if obj['protocol_number']:
                        parts.append(f"IP/{obj['protocol_number']}")
                    else:
                        parts.append('ALL')  # IP/ALL 대신 ALL 간소화 표기 / Normalize IP/ALL to ALL
                else:
                    if obj['tcp_port']:
                        parts.append(f"TCP/{obj['tcp_port']}")
                    if obj['udp_port']:
                        parts.append(f"UDP/{obj['udp_port']}")
                    if obj['sctp_port']:
                        parts.append(f"SCTP/{obj['sctp_port']}")
                obj['display'] = ', '.join(parts) if parts else ('ALL' if name == 'ALL' else name)
                svcs[name] = obj
            i += 1
    return svcs


def parse_service_groups(lines, vdom_start, vdom_end):
    """config firewall service group -> dict[name] = {'members': [...], 'comment': '...'}"""
    groups = {}
    ranges = find_section_range(lines, vdom_start, vdom_end, "firewall service group")
    for sr, er in ranges:
        i = sr + 1
        while i <= er:
            if lines[i].strip().startswith("edit "):
                name = parse_edit_id(lines[i])
                members = []
                comment = ""
                i += 1
                while i <= er:
                    s = lines[i].strip()
                    if s in ("next", "end"):
                        break
                    if s.startswith("set "):
                        f = get_field_name(lines[i])
                        if f == "member":
                            members = parse_quoted_values(lines[i])
                        elif f in ("comment", "comments"):
                            comment = parse_set_value(lines[i])
                    i += 1
                groups[name] = {'members': members, 'comment': comment}
            i += 1
    return groups


def parse_ippool_objects(lines, vdom_start, vdom_end):
    """config firewall ippool -> dict[name] = {startip, endip, type, display}"""
    pools = {}
    ranges = find_section_range(lines, vdom_start, vdom_end, "firewall ippool")
    for sr, er in ranges:
        i = sr + 1
        while i <= er:
            if lines[i].strip().startswith("edit "):
                name = parse_edit_id(lines[i])
                obj = {'startip': '', 'endip': '', 'type': 'overload', 'display': '', 'comment': ''}
                i += 1
                while i <= er:
                    s = lines[i].strip()
                    if s in ("next", "end"):
                        break
                    if s.startswith("set "):
                        f = get_field_name(lines[i])
                        v = parse_set_value(lines[i])
                        if f in ('startip', 'endip', 'type'):
                            obj[f] = v
                        elif f in ('comment', 'comments'):
                            obj['comment'] = v
                    i += 1
                if obj['startip'] == obj['endip']:
                    obj['display'] = f"{obj['startip']}/32" if obj['startip'] else ''
                else:
                    obj['display'] = f"{obj['startip']}-{obj['endip']}"
                pools[name] = obj
            i += 1
    return pools


def format_schedule_time(obj):
    """
    스케줄 객체(recurring / onetime)의 실제 시간 문자열 포맷팅
    Format actual schedule time string for recurring or onetime schedules
    """
    if not obj:
        return ""
    stype = obj.get('type', '')
    if stype == 'recurring':
        days = obj.get('day', [])
        start = obj.get('start', '')
        end = obj.get('end', '')

        day_map = {
            'sunday': 'Sun', 'monday': 'Mon', 'tuesday': 'Tue',
            'wednesday': 'Wed', 'thursday': 'Thu', 'friday': 'Fri', 'saturday': 'Sat',
            'none': 'None'
        }
        all_7_days = {'sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday'}
        day_set = set(d.lower() for d in days)

        if day_set >= all_7_days:
            day_str = "매일 (Sun-Sat)"
        elif days:
            day_str = ', '.join(day_map.get(d.lower(), d) for d in days)
        else:
            day_str = "매일 (Sun-Sat)"

        has_time = False
        time_str = ""
        if start or end:
            s_val = start if start else "00:00"
            e_val = end if end else "00:00"
            if not (s_val == "00:00" and e_val == "00:00"):
                time_str = f"{s_val}-{e_val}"
                has_time = True

        if has_time:
            return f"{day_str} {time_str}"
        else:
            if day_set >= all_7_days:
                return "Always"
            return f"{day_str} (All Day)"

    elif stype == 'onetime':
        start_raw = obj.get('start', '')
        end_raw = obj.get('end', '')

        def _fmt_dt(dt_str):
            if not dt_str:
                return ""
            parts = dt_str.strip().split()
            if len(parts) == 2:
                # time date -> date time (e.g., '00:00 2026/07/16' -> '2026/07/16 00:00')
                if '/' in parts[1] or '-' in parts[1]:
                    return f"{parts[1]} {parts[0]}"
            return dt_str

        s_fmt = _fmt_dt(start_raw)
        e_fmt = _fmt_dt(end_raw)
        if s_fmt and e_fmt:
            return f"{s_fmt} ~ {e_fmt}"
        elif s_fmt:
            return f"{s_fmt} ~"
        elif e_fmt:
            return f"~ {e_fmt}"
        return ""

    return ""


def parse_schedule_recurring(lines, vdom_start, vdom_end):
    """config firewall schedule recurring -> dict[name] = {'type': 'recurring', 'day': [...], 'start': '...', 'end': '...', 'comment': '...'}"""
    scheds = {}
    builtin = {
        'always': {
            'type': 'recurring',
            'day': ['sunday', 'monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday'],
            'start': '', 'end': '', 'comment': 'Always active'
        },
        'none': {
            'type': 'recurring',
            'day': ['none'],
            'start': '', 'end': '', 'comment': 'Never active'
        }
    }
    scheds.update(builtin)

    ranges = find_section_range(lines, vdom_start, vdom_end, "firewall schedule recurring")
    for sr, er in ranges:
        i = sr + 1
        while i <= er:
            if lines[i].strip().startswith("edit "):
                name = parse_edit_id(lines[i])
                obj = {'type': 'recurring', 'day': [], 'start': '', 'end': '', 'comment': ''}
                i += 1
                while i <= er:
                    s = lines[i].strip()
                    if s in ("next", "end"):
                        break
                    if s.startswith("set "):
                        f = get_field_name(lines[i])
                        v = parse_set_value(lines[i])
                        if f == 'day':
                            obj['day'] = v.split()
                        elif f == 'start':
                            obj['start'] = v
                        elif f == 'end':
                            obj['end'] = v
                        elif f in ('comment', 'comments'):
                            obj['comment'] = v
                    i += 1
                scheds[name] = obj
            i += 1
    return scheds


def parse_schedule_onetime(lines, vdom_start, vdom_end):
    """config firewall schedule onetime -> dict[name] = {'type': 'onetime', 'start': '...', 'end': '...', 'comment': '...'}"""
    scheds = {}
    ranges = find_section_range(lines, vdom_start, vdom_end, "firewall schedule onetime")
    for sr, er in ranges:
        i = sr + 1
        while i <= er:
            if lines[i].strip().startswith("edit "):
                name = parse_edit_id(lines[i])
                obj = {'type': 'onetime', 'start': '', 'end': '', 'expiration-days': '', 'comment': ''}
                i += 1
                while i <= er:
                    s = lines[i].strip()
                    if s in ("next", "end"):
                        break
                    if s.startswith("set "):
                        f = get_field_name(lines[i])
                        v = parse_set_value(lines[i])
                        if f == 'start':
                            obj['start'] = v
                        elif f == 'end':
                            obj['end'] = v
                        elif f == 'expiration-days':
                            obj['expiration-days'] = v
                        elif f in ('comment', 'comments'):
                            obj['comment'] = v
                    i += 1
                scheds[name] = obj
            i += 1
    return scheds


def parse_schedule_group(lines, vdom_start, vdom_end):
    """config firewall schedule group -> dict[name] = {'type': 'group', 'members': [...], 'comment': '...'}"""
    groups = {}
    ranges = find_section_range(lines, vdom_start, vdom_end, "firewall schedule group")
    for sr, er in ranges:
        i = sr + 1
        while i <= er:
            if lines[i].strip().startswith("edit "):
                name = parse_edit_id(lines[i])
                obj = {'type': 'group', 'members': [], 'comment': ''}
                i += 1
                while i <= er:
                    s = lines[i].strip()
                    if s in ("next", "end"):
                        break
                    if s.startswith("set "):
                        f = get_field_name(lines[i])
                        if f == 'member':
                            obj['members'] = parse_quoted_values(lines[i])
                        elif f in ('comment', 'comments'):
                            obj['comment'] = parse_set_value(lines[i])
                    i += 1
                groups[name] = obj
            i += 1
    return groups


def parse_security_profile_comments(lines, start, end):
    """
    FortiGate 보안 프로파일들의 코멘트 수집
    """
    profile_comments = {}
    sec_names = [
        "firewall ssl-ssh-profile", "ips sensor", "antivirus profile",
        "webfilter profile", "dnsfilter profile", "application list",
        "file-filter profile", "emailfilter profile", "dlp sensor",
        "dlp profile", "waf profile", "casb profile", "videofilter profile",
        "sctp-filter profile", "cifs profile", "virtual-patch profile",
        "voip profile"
    ]
    for sec in sec_names:
        for sr, er in find_section_range(lines, start, end, sec):
            i = sr + 1
            cur_prof = None
            depth = 1
            while i <= er:
                line = lines[i]
                s = line.strip()
                if s.startswith("config "):
                    depth += 1
                elif s == "end":
                    depth -= 1
                elif depth == 1:
                    if s.startswith("edit "):
                        cur_prof = parse_edit_id(line)
                    elif s.startswith("set "):
                        f = get_field_name(line)
                        if f in ('comment', 'comments') and cur_prof:
                            profile_comments[cur_prof] = parse_set_value(line)
                    elif s == "next":
                        cur_prof = None
                i += 1
    return profile_comments


def parse_external_resources(lines, start, end):
    """
    'config system external-resource' 파싱
    Returns: dict of {name: {'name': ..., 'type': ..., 'comments': ..., 'resource': ..., 'refresh-rate': ..., 'source-ip': ...}}
    """
    resources = OrderedDict()
    for sr, er in find_section_range(lines, start, end, "system external-resource"):
        i = sr + 1
        cur_obj = None
        while i <= er:
            s = lines[i].strip()
            if s.startswith("edit "):
                name = parse_edit_id(lines[i])
                cur_obj = {
                    'name': name,
                    'status': 'enable',
                    'type': 'address',
                    'comments': '',
                    'resource': '',
                    'refresh-rate': '',
                    'source-ip': '',
                    'category': ''
                }
            elif s.startswith("set ") and cur_obj:
                f = get_field_name(lines[i])
                v = parse_set_value(lines[i])
                if f in ('comment', 'comments'):
                    cur_obj['comments'] = v
                else:
                    cur_obj[f] = v
            elif s in ("next", "end"):
                if cur_obj:
                    resources[cur_obj['name']] = cur_obj
                    cur_obj = None
            i += 1
    return resources


def parse_vdom_inspection_mode(lines, start, end):
    """
    vDOM 설정의 inspection-mode (flow 또는 proxy) 파싱
    """
    for sr, er in find_section_range(lines, start, end, "system settings"):
        for i in range(sr + 1, er + 1):
            s = lines[i].strip()
            if s.startswith("set inspection-mode"):
                return parse_set_value(lines[i])
    return "flow"


# ================================================================
#  4. 객체 리졸버 (그룹 확장 + IP/코멘트 매핑) / Object Resolver (Recursive Group & Comment Resolution)
# ================================================================

class ObjectResolver:
    def __init__(self, addr_dict, addrgrp_dict, svc_dict, svcgrp_dict, ippool_dict, ext_resources=None,
                 sched_recurring_dict=None, sched_onetime_dict=None, sched_group_dict=None):
        self.addrs = addr_dict
        self.addrgrps = addrgrp_dict
        self.svcs = svc_dict
        self.svcgrps = svcgrp_dict
        self.ippools = ippool_dict
        self.ext_resources = ext_resources or {}
        self.sched_recurring = sched_recurring_dict or {}
        self.sched_onetime = sched_onetime_dict or {}
        self.sched_groups = sched_group_dict or {}
        self._addr_cache = {}
        self._svc_cache = {}
        self._sched_cache = {}

    def resolve_address(self, name, _visited=None):
        """
        주소 객체 또는 그룹을 재귀적으로 확장하여 멤버별 세부 정보 반환
        Recursively resolve address object or group into individual member entries
        Returns: list of (group_name, obj_name, obj_type, formatted_ip, comment)
        """
        top_level = (_visited is None)
        if top_level:
            if name in self._addr_cache:
                return self._addr_cache[name]
            _visited = set()
        if name in _visited:
            return [(None, name, 'ref-loop', '', '')]
        _visited.add(name)

        if name in self.addrgrps:
            result = []
            grp_data = self.addrgrps[name]
            members = grp_data['members'] if isinstance(grp_data, dict) else grp_data
            grp_comment = grp_data.get('comment', '') if isinstance(grp_data, dict) else ''

            for m in members:
                sub = self.resolve_address(m, _visited.copy())
                for item in sub:
                    c = item[4] or grp_comment
                    result.append((name, item[1], item[2], item[3], c))
            if not result:
                result.append((None, name, 'group(empty)', '', grp_comment))
            if top_level:
                self._addr_cache[name] = result
            return result

        if name in self.addrs:
            obj = self.addrs[name]
            f_ip = format_ip_str(obj)
            res = [(None, name, obj.get('type', 'ipmask'), f_ip, obj.get('comment', ''))]
            if top_level:
                self._addr_cache[name] = res
            return res

        if name in self.ext_resources:
            ext = self.ext_resources[name]
            ext_type = ext.get('type', 'external-resource')
            res = [(None, name, ext_type, ext.get('resource', ''), ext.get('comments', ''))]
            if top_level:
                self._addr_cache[name] = res
            return res

        res = [(None, name, 'unknown', '', '')]
        if top_level:
            self._addr_cache[name] = res
        return res

    def resolve_service(self, name, _visited=None):
        """
        서비스 객체 또는 그룹을 재귀적으로 확장하여 포트 및 코멘트 정보 반환
        Recursively resolve service object or group into individual service/port entries
        Returns: list of (group_name, svc_name, port_display, comment)
        """
        top_level = (_visited is None)
        if top_level:
            if name in self._svc_cache:
                return self._svc_cache[name]
            _visited = set()
        if name in _visited:
            return [(None, name, '', '')]
        _visited.add(name)

        if name in self.svcgrps:
            result = []
            grp_data = self.svcgrps[name]
            members = grp_data['members'] if isinstance(grp_data, dict) else grp_data
            grp_comment = grp_data.get('comment', '') if isinstance(grp_data, dict) else ''

            for m in members:
                sub = self.resolve_service(m, _visited.copy())
                for item in sub:
                    c = item[3] or grp_comment
                    result.append((name, item[1], item[2], c))
            res = result if result else [(None, name, '', grp_comment)]
            if top_level:
                self._svc_cache[name] = res
            return res

        if name in self.svcs:
            obj = self.svcs[name]
            proto = obj.get('protocol', '')
            port_parts = []
            if obj.get('tcp_port'):
                port_parts.append(f"TCP/{obj['tcp_port']}")
            if obj.get('udp_port'):
                port_parts.append(f"UDP/{obj['udp_port']}")
            if obj.get('sctp_port'):
                port_parts.append(f"SCTP/{obj['sctp_port']}")
            if proto == 'ICMP':
                port_parts.append('ICMP')
            elif proto == 'ICMP6':
                port_parts.append('ICMP6')
            elif proto == 'IP':
                pn = obj.get('protocol_number', '')
                if pn:
                    port_parts.append(f"IP/{pn}")
                else:
                    port_parts.append('ALL')
            port_display = ', '.join(port_parts) if port_parts else ''
            if port_display in ('IP/ALL', ''):
                if name == 'ALL':
                    port_display = 'ALL'
            res = [(None, name, port_display, obj.get('comment', ''))]
            if top_level:
                self._svc_cache[name] = res
            return res

        res = [(None, name, '', '')]
        if top_level:
            self._svc_cache[name] = res
        return res

    def resolve_ippool(self, name):
        """IP Pool 이름 -> (pool_name, pool_ip_display, pool_type) / Resolve IP pool name to display tuple"""
        if name in self.ippools:
            obj = self.ippools[name]
            return (name, obj.get('display', ''), obj.get('type', ''))
        return (name, '', '')

    def resolve_schedule(self, name, _visited=None):
        """
        스케줄 객체 또는 그룹을 재귀적으로 확장하여 멤버별 세부 시간 정보 반환
        Recursively resolve schedule object or group into individual member schedule entries
        Returns: list of (group_name, obj_name, sched_type, time_display, comment)
        """
        if not name:
            return [(None, 'always', 'recurring', 'Always', '')]

        top_level = (_visited is None)
        if top_level:
            if name in self._sched_cache:
                return self._sched_cache[name]
            _visited = set()
        if name in _visited:
            return [(None, name, 'ref-loop', '', '')]
        _visited.add(name)

        if name in self.sched_groups:
            result = []
            grp_data = self.sched_groups[name]
            members = grp_data['members'] if isinstance(grp_data, dict) else []
            grp_comment = grp_data.get('comment', '') if isinstance(grp_data, dict) else ''

            for m in members:
                sub = self.resolve_schedule(m, _visited.copy())
                for item in sub:
                    c = item[4] or grp_comment
                    result.append((name, item[1], item[2], item[3], c))
            res = result if result else [(name, '(empty)', 'group(empty)', '', grp_comment)]
            if top_level:
                self._sched_cache[name] = res
            return res

        if name in self.sched_recurring:
            obj = self.sched_recurring[name]
            time_disp = format_schedule_time(obj)
            res = [(None, name, 'recurring', time_disp, obj.get('comment', ''))]
            if top_level:
                self._sched_cache[name] = res
            return res

        if name in self.sched_onetime:
            obj = self.sched_onetime[name]
            time_disp = format_schedule_time(obj)
            res = [(None, name, 'onetime', time_disp, obj.get('comment', ''))]
            if top_level:
                self._sched_cache[name] = res
            return res

        nl = name.lower()
        if nl == 'always':
            res = [(None, 'always', 'recurring', 'Always', '')]
        elif nl == 'none':
            res = [(None, 'none', 'recurring', 'None', '')]
        else:
            res = [(None, name, 'unknown', '', '')]

        if top_level:
            self._sched_cache[name] = res
        return res



# ================================================================
#  5. 정책 파서 / Policy Parsers (Firewall, Local-in, Central-NAT, VIP, DoS)
# ================================================================

def parse_firewall_policy(lines, sec_start, sec_end):
    policies = []
    i = sec_start + 1
    while i <= sec_end:
        if lines[i].strip().startswith("edit "):
            p = {}
            p['id'] = parse_edit_id(lines[i])
            p['status'] = 'enable'
            p['name'] = ''
            p['action'] = 'deny'
            p['srcintf'] = []
            p['dstintf'] = []
            p['srcaddr'] = []
            p['dstaddr'] = []
            p['internet-service'] = ''
            p['internet-service-name'] = []
            p['internet-service-src'] = ''
            p['internet-service-src-name'] = []
            p['schedule'] = ''
            p['service'] = []
            p['utm-status'] = ''
            p['ssl-ssh-profile'] = ''
            p['ips-sensor'] = ''
            p['av-profile'] = ''
            p['webfilter-profile'] = ''
            p['application-list'] = ''
            p['nat'] = ''
            p['ippool'] = ''
            p['poolname'] = []
            p['logtraffic'] = ''
            p['logtraffic-start'] = ''
            p['comments'] = ''
            i += 1
            while i <= sec_end:
                s = lines[i].strip()
                if s in ("next", "end"):
                    break
                if s.startswith("set "):
                    f = get_field_name(lines[i])
                    if f in ('srcaddr', 'dstaddr', 'service', 'srcintf', 'dstintf',
                             'poolname', 'internet-service-name', 'internet-service-src-name'):
                        p[f] = parse_quoted_values(lines[i])
                    elif f == 'action':
                        p['action'] = parse_set_value(lines[i])
                    else:
                        p[f] = parse_set_value(lines[i])
                i += 1
            policies.append(p)
        i += 1
    return policies


def parse_local_in_policy(lines, sec_start, sec_end):
    policies = []
    i = sec_start + 1
    while i <= sec_end:
        if lines[i].strip().startswith("edit "):
            p = {}
            p['id'] = parse_edit_id(lines[i])
            p['intf'] = ''
            p['srcaddr'] = []
            p['dstaddr'] = []
            p['action'] = 'deny'
            p['service'] = []
            p['schedule'] = ''
            p['status'] = 'enable'
            p['comments'] = ''
            i += 1
            while i <= sec_end:
                s = lines[i].strip()
                if s in ("next", "end"):
                    break
                if s.startswith("set "):
                    f = get_field_name(lines[i])
                    if f in ('srcaddr', 'dstaddr', 'service'):
                        p[f] = parse_quoted_values(lines[i])
                    elif f == 'action':
                        p['action'] = parse_set_value(lines[i])
                    else:
                        p[f] = parse_set_value(lines[i])
                i += 1
            policies.append(p)
        i += 1
    return policies


def parse_central_snat(lines, sec_start, sec_end):
    entries = []
    i = sec_start + 1
    while i <= sec_end:
        if lines[i].strip().startswith("edit "):
            e = {}
            e['id'] = parse_edit_id(lines[i])
            e['status'] = 'enable'
            e['srcintf'] = []
            e['dstintf'] = []
            e['orig-addr'] = []
            e['dst-addr'] = []
            e['nat-ippool'] = []
            e['nat'] = 'enable'
            e['comments'] = ''
            i += 1
            while i <= sec_end:
                s = lines[i].strip()
                if s in ("next", "end"):
                    break
                if s.startswith("set "):
                    f = get_field_name(lines[i])
                    if f in ('orig-addr', 'dst-addr', 'nat-ippool', 'srcintf', 'dstintf'):
                        e[f] = parse_quoted_values(lines[i])
                    else:
                        e[f] = parse_set_value(lines[i])
                i += 1
            entries.append(e)
        i += 1
    return entries


def parse_vip(lines, sec_start, sec_end):
    entries = []
    i = sec_start + 1
    while i <= sec_end:
        if lines[i].strip().startswith("edit "):
            e = {}
            e['name'] = parse_edit_id(lines[i])
            e['uuid'] = ''
            e['comment'] = ''
            e['type'] = 'static-nat'
            e['extip'] = ''
            e['mappedip'] = ''
            e['extintf'] = ''
            e['portforward'] = ''
            e['protocol'] = ''
            e['extport'] = ''
            e['mappedport'] = ''
            e['server-type'] = ''
            e['ldb-method'] = ''
            e['monitor'] = ''
            e['service'] = []
            e['arp-reply'] = 'enable'
            e['nat-source-vip'] = 'disable'
            e['realservers'] = []
            i += 1
            in_rs = False
            rs = None
            while i <= sec_end:
                s = lines[i].strip()
                if in_rs:
                    if s == "end":
                        if rs:
                            e['realservers'].append(rs)
                            rs = None
                        in_rs = False
                        i += 1
                        continue
                    if s.startswith("edit "):
                        if rs:
                            e['realservers'].append(rs)
                        rs = {'id': parse_edit_id(lines[i]), 'ip': '', 'port': '', 'weight': '1'}
                    elif s.startswith("set ") and rs:
                        f = get_field_name(lines[i])
                        rs[f] = parse_set_value(lines[i])
                    elif s == "next":
                        if rs:
                            e['realservers'].append(rs)
                            rs = None
                    i += 1
                    continue
                if s in ("next", "end"):
                    break
                if s == "config realservers":
                    in_rs = True
                    i += 1
                    continue
                if s.startswith("set "):
                    f = get_field_name(lines[i])
                    if f == 'mappedip':
                        e['mappedip'] = '\n'.join(parse_quoted_values(lines[i]))
                    elif f == 'monitor':
                        e['monitor'] = '\n'.join(parse_quoted_values(lines[i]))
                    elif f == 'service':
                        e['service'] = parse_quoted_values(lines[i])
                    elif f == 'arp-reply':
                        val = parse_set_value(lines[i])
                        e['arp-reply'] = val if val else 'enable'
                    elif f == 'nat-source-vip':
                        val = parse_set_value(lines[i])
                        e['nat-source-vip'] = val if val else 'disable'
                    else:
                        e[f] = parse_set_value(lines[i])
                i += 1
            entries.append(e)
        i += 1
    return entries


def parse_dos_policy(lines, sec_start, sec_end):
    policies = []
    i = sec_start + 1
    while i <= sec_end:
        if lines[i].strip().startswith("edit "):
            p = {}
            p['id'] = parse_edit_id(lines[i])
            p['interface'] = ''
            p['srcaddr'] = []
            p['dstaddr'] = []
            p['service'] = []
            p['status'] = 'enable'
            p['comments'] = ''
            p['anomalies'] = []
            i += 1
            in_anomaly = False
            anomaly = None
            while i <= sec_end:
                s = lines[i].strip()
                if in_anomaly:
                    if s == "end":
                        if anomaly:
                            p['anomalies'].append(anomaly)
                            anomaly = None
                        in_anomaly = False
                        i += 1
                        continue
                    if s.startswith("edit "):
                        if anomaly:
                            p['anomalies'].append(anomaly)
                        anomaly = {'name': parse_edit_id(lines[i]),
                                   'status': 'disable', 'log': 'disable',
                                   'action': 'disable', 'quarantine': 'disable',
                                   'threshold': ''}
                    elif s.startswith("set ") and anomaly:
                        f_name = get_field_name(lines[i])
                        f_val = parse_set_value(lines[i])
                        if f_name == 'quarantine':
                            anomaly['quarantine'] = 'attacker' if f_val.lower() == 'attacker' else 'disable'
                        else:
                            anomaly[f_name] = f_val
                    elif s == "next":
                        if anomaly:
                            p['anomalies'].append(anomaly)
                            anomaly = None
                    i += 1
                    continue
                if s in ("next", "end"):
                    break
                if s == "config anomaly":
                    in_anomaly = True
                    i += 1
                    continue
                if s.startswith("set "):
                    f = get_field_name(lines[i])
                    if f in ('srcaddr', 'dstaddr', 'service'):
                        p[f] = parse_quoted_values(lines[i])
                    else:
                        p[f] = parse_set_value(lines[i])
                i += 1
            policies.append(p)
        i += 1
    return policies


def parse_firewall_acl(lines, sec_start, sec_end):
    """
    config firewall acl 파싱
    Returns list of ACL policy dicts
    """
    policies = []
    i = sec_start + 1
    while i <= sec_end:
        if lines[i].strip().startswith("edit "):
            p = {
                'id': parse_edit_id(lines[i]),
                'status': 'enable',
                'interface': '',
                'srcaddr': [],
                'dstaddr': [],
                'service': [],
                'comments': ''
            }
            i += 1
            while i <= sec_end:
                s = lines[i].strip()
                if s in ("next", "end"):
                    break
                if s.startswith("set "):
                    f = get_field_name(lines[i])
                    if f in ('srcaddr', 'dstaddr', 'service'):
                        p[f] = parse_quoted_values(lines[i])
                    elif f in ('comments', 'comment'):
                        p['comments'] = parse_set_value(lines[i])
                    else:
                        p[f] = parse_set_value(lines[i])
                i += 1
            policies.append(p)
        i += 1
    return policies


def parse_router_static(lines, vdom_start, vdom_end):
    """config router static -> list of static route dicts"""
    routes = []
    ranges = find_section_range(lines, vdom_start, vdom_end, "router static")
    for sr, er in ranges:
        i = sr + 1
        while i <= er:
            if lines[i].strip().startswith("edit "):
                route_id = parse_edit_id(lines[i])
                r = {
                    'id': route_id,
                    'status': 'enable',
                    'dst': '0.0.0.0/0',
                    'dst_raw': '',
                    'gateway': '',
                    'device': '',
                    'distance': '10',
                    'priority': '1',
                    'comment': '',
                    'blackhole': 'disable',
                    'dynamic-gateway': 'disable',
                    'link-monitor-exempt': 'disable',
                    'bfd': 'disable',
                }
                has_dst = False
                has_distance = False
                has_priority = False
                i += 1
                while i <= er:
                    s = lines[i].strip()
                    if s in ("next", "end"):
                        break
                    if s.startswith("set "):
                        f = get_field_name(lines[i])
                        v = parse_set_value(lines[i])
                        if f == 'dst':
                            has_dst = True
                            parts = v.split()
                            if len(parts) == 2:
                                ip, pf = format_subnet(parts[0], parts[1])
                                r['dst'] = f"{ip}/{pf}" if pf else ip
                            elif '/' in v:
                                r['dst'] = v
                            else:
                                r['dst'] = v
                            r['dst_raw'] = v
                        elif f == 'gateway':
                            r['gateway'] = v
                        elif f == 'device':
                            r['device'] = v
                        elif f == 'distance':
                            r['distance'] = v
                            has_distance = True
                        elif f == 'priority':
                            r['priority'] = v
                            has_priority = True
                        elif f == 'status':
                            r['status'] = v
                        elif f in ('comment', 'comments'):
                            r['comment'] = v
                        elif f == 'blackhole':
                            r['blackhole'] = v
                        elif f == 'dynamic-gateway':
                            r['dynamic-gateway'] = v
                        elif f == 'link-monitor-exempt':
                            r['link-monitor-exempt'] = v
                        elif f == 'bfd':
                            r['bfd'] = v
                    i += 1
                if not has_distance:
                    r['distance'] = '10'
                if not has_priority:
                    r['priority'] = '1'
                routes.append(r)
            i += 1
    return routes


def parse_router_policy(lines, vdom_start, vdom_end):
    """config router policy -> list of policy route dicts"""
    policies = []
    ranges = find_section_range(lines, vdom_start, vdom_end, "router policy")
    for sr, er in ranges:
        i = sr + 1
        while i <= er:
            if lines[i].strip().startswith("edit "):
                pol_id = parse_edit_id(lines[i])
                p = {
                    'id': pol_id,
                    'status': 'enable',
                    'seq-num': '',
                    'input-device': [],
                    'output-device': [],
                    'srcaddr': [],
                    'src': '',
                    'dstaddr': [],
                    'dst': '',
                    'action': 'permit',
                    'protocol': '0',
                    'start-port': '',
                    'end-port': '',
                    'gateway': '',
                    'comments': ''
                }
                i += 1
                while i <= er:
                    s = lines[i].strip()
                    if s in ("next", "end"):
                        break
                    if s.startswith("set "):
                        f = get_field_name(lines[i])
                        if f in ('input-device', 'output-device'):
                            p[f] = parse_quoted_values(lines[i])
                        elif f in ('srcaddr', 'dstaddr'):
                            p[f] = parse_quoted_values(lines[i])
                        else:
                            v = parse_set_value(lines[i])
                            if f == 'src':
                                parts = v.split()
                                if len(parts) == 2:
                                    ip, pf = format_subnet(parts[0], parts[1])
                                    p['src'] = f"{ip}/{pf}" if pf else ip
                                else:
                                    p['src'] = v
                            elif f == 'dst':
                                parts = v.split()
                                if len(parts) == 2:
                                    ip, pf = format_subnet(parts[0], parts[1])
                                    p['dst'] = f"{ip}/{pf}" if pf else ip
                                else:
                                    p['dst'] = v
                            else:
                                p[f] = v
                    i += 1
                policies.append(p)
            i += 1
    return policies


def parse_router_ospf(lines, vdom_start, vdom_end):
    """config router ospf -> dict containing router-id, networks, interfaces, redistribute, areas"""
    data = {
        'router-id': '',
        'networks': [],
        'interfaces': [],
        'redistribute': [],
        'areas': []
    }
    ranges = find_section_range(lines, vdom_start, vdom_end, "router ospf")
    if not ranges:
        return data

    for sr, er in ranges:
        i = sr + 1
        while i <= er:
            s = lines[i].strip()
            if s.startswith("set router-id "):
                data['router-id'] = parse_set_value(lines[i])
            elif s == "config network":
                i += 1
                while i <= er:
                    s2 = lines[i].strip()
                    if s2 == "end":
                        break
                    if s2.startswith("edit "):
                        net_id = parse_edit_id(lines[i])
                        net_entry = {'id': net_id, 'prefix': '', 'area': '0.0.0.0'}
                        i += 1
                        while i <= er:
                            s3 = lines[i].strip()
                            if s3 in ("next", "end"):
                                break
                            if s3.startswith("set prefix "):
                                v = parse_set_value(lines[i])
                                parts = v.split()
                                if len(parts) == 2:
                                    ip, pf = format_subnet(parts[0], parts[1])
                                    net_entry['prefix'] = f"{ip}/{pf}" if pf else ip
                                else:
                                    net_entry['prefix'] = v
                            elif s3.startswith("set area "):
                                net_entry['area'] = parse_set_value(lines[i])
                            i += 1
                        data['networks'].append(net_entry)
                    i += 1
            elif s == "config ospf-interface":
                i += 1
                while i <= er:
                    s2 = lines[i].strip()
                    if s2 == "end":
                        break
                    if s2.startswith("edit "):
                        intf_name = parse_edit_id(lines[i])
                        intf_entry = {
                            'name': intf_name, 'interface': '', 'cost': '',
                            'priority': '', 'dead-interval': '', 'hello-interval': '',
                            'network-type': '', 'authentication': ''
                        }
                        i += 1
                        while i <= er:
                            s3 = lines[i].strip()
                            if s3 in ("next", "end"):
                                break
                            if s3.startswith("set "):
                                f = get_field_name(lines[i])
                                intf_entry[f] = parse_set_value(lines[i])
                            i += 1
                        data['interfaces'].append(intf_entry)
                    i += 1
            elif s.startswith("config redistribute "):
                proto = parse_edit_id(s)
                redist_entry = {'protocol': proto, 'status': 'disable', 'routemap': '', 'metric': '', 'metric-type': ''}
                i += 1
                while i <= er:
                    s2 = lines[i].strip()
                    if s2 == "end":
                        break
                    if s2.startswith("set "):
                        f = get_field_name(lines[i])
                        redist_entry[f] = parse_set_value(lines[i])
                    i += 1
                data['redistribute'].append(redist_entry)
            elif s == "config area":
                i += 1
                while i <= er:
                    s2 = lines[i].strip()
                    if s2 == "end":
                        break
                    if s2.startswith("edit "):
                        area_id = parse_edit_id(lines[i])
                        area_entry = {'id': area_id, 'type': 'regular'}
                        i += 1
                        while i <= er:
                            s3 = lines[i].strip()
                            if s3 in ("next", "end"):
                                break
                            if s3.startswith("set type "):
                                area_entry['type'] = parse_set_value(lines[i])
                            i += 1
                        data['areas'].append(area_entry)
                    i += 1
            i += 1

    acls, rmaps = parse_router_route_map_and_acl(lines, vdom_start, vdom_end)
    data['access_lists'] = acls
    data['route_maps'] = rmaps
    return data


def parse_router_route_map_and_acl(lines, vdom_start, vdom_end):
    """
    config router route-map 및 config router access-list 파싱
    Returns (acls_dict, route_maps_dict)
    """
    acls = {}
    acl_ranges = find_section_range(lines, vdom_start, vdom_end, "router access-list")
    for sr, er in acl_ranges:
        i = sr + 1
        while i <= er:
            if lines[i].strip().startswith("edit "):
                name = parse_edit_id(lines[i])
                acl = {'name': name, 'comments': '', 'rules': []}
                i += 1
                while i <= er:
                    s = lines[i].strip()
                    if s in ("next", "end"):
                        break
                    if s.startswith("set comments ") or s.startswith("set comment "):
                        acl['comments'] = parse_set_value(lines[i])
                    elif s == "config rule":
                        i += 1
                        while i <= er:
                            s2 = lines[i].strip()
                            if s2 == "end":
                                break
                            if s2.startswith("edit "):
                                rid = parse_edit_id(lines[i])
                                rule = {'id': rid, 'action': 'permit', 'prefix': '', 'exact_match': 'disable'}
                                i += 1
                                while i <= er:
                                    s3 = lines[i].strip()
                                    if s3 in ("next", "end"):
                                        break
                                    if s3.startswith("set action "):
                                        rule['action'] = parse_set_value(lines[i])
                                    elif s3.startswith("set prefix "):
                                        v = parse_set_value(lines[i])
                                        parts = v.split()
                                        if len(parts) == 2:
                                            ip, pf = format_subnet(parts[0], parts[1])
                                            rule['prefix'] = f"{ip}/{pf}" if pf else ip
                                        else:
                                            rule['prefix'] = v
                                    elif s3.startswith("set exact-match "):
                                        rule['exact_match'] = parse_set_value(lines[i])
                                    i += 1
                                acl['rules'].append(rule)
                            i += 1
                    i += 1
                acls[name] = acl
            i += 1

    rmaps = {}
    rm_ranges = find_section_range(lines, vdom_start, vdom_end, "router route-map")
    for sr, er in rm_ranges:
        i = sr + 1
        while i <= er:
            if lines[i].strip().startswith("edit "):
                name = parse_edit_id(lines[i])
                rm = {'name': name, 'comments': '', 'rules': []}
                i += 1
                while i <= er:
                    s = lines[i].strip()
                    if s in ("next", "end"):
                        break
                    if s.startswith("set comments ") or s.startswith("set comment "):
                        rm['comments'] = parse_set_value(lines[i])
                    elif s == "config rule":
                        i += 1
                        while i <= er:
                            s2 = lines[i].strip()
                            if s2 == "end":
                                break
                            if s2.startswith("edit "):
                                rid = parse_edit_id(lines[i])
                                rrule = {'id': rid, 'action': 'permit', 'match_ip': '', 'set_actions': []}
                                i += 1
                                while i <= er:
                                    s3 = lines[i].strip()
                                    if s3 in ("next", "end"):
                                        break
                                    if s3.startswith("set action "):
                                        rrule['action'] = parse_set_value(lines[i])
                                    elif s3.startswith("set match-ip-address "):
                                        rrule['match_ip'] = parse_set_value(lines[i])
                                    elif s3.startswith("set "):
                                        rrule['set_actions'].append(lines[i].strip())
                                    i += 1
                                rm['rules'].append(rrule)
                            i += 1
                    i += 1
                rmaps[name] = rm
            i += 1

    return acls, rmaps


def parse_system_interfaces(lines):
    """
    config system interface 파싱
    Returns dict: vdom_name -> list of interface dicts
    """
    vdom_intfs = {}
    ranges = find_section_range(lines, 0, len(lines)-1, "system interface")
    for sr, er in ranges:
        i = sr + 1
        while i <= er:
            if lines[i].strip().startswith("edit "):
                name = parse_edit_id(lines[i])
                intf = {
                    'name': name,
                    'vdom': 'root',
                    'ip': '',
                    'prefix': '',
                    'display': '',
                    'secondary_ips': [],
                    'remote-ip': '',
                    'type': 'physical',
                    'vlanid': '',
                    'interface': '',
                    'member': [],
                    'status': 'up',
                    'mode': 'static',
                    'vrf': '0',
                    'allowaccess': '',
                    'alias': '',
                    'description': '',
                    'speed': '',
                    'mtu': ''
                }
                i += 1
                while i <= er:
                    s = lines[i].strip()
                    if s in ("next", "end"):
                        break
                    if s == "config secondaryip":
                        i += 1
                        while i <= er:
                            s2 = lines[i].strip()
                            if s2 == "end":
                                break
                            if s2.startswith("edit "):
                                i += 1
                                while i <= er:
                                    s3 = lines[i].strip()
                                    if s3 in ("next", "end"):
                                        break
                                    if s3.startswith("set ip "):
                                        vip = parse_set_value(lines[i])
                                        parts = vip.split()
                                        if len(parts) == 2:
                                            sip, spf = format_subnet(parts[0], parts[1])
                                            sec_str = f"{sip}/{spf}" if spf else sip
                                        else:
                                            sec_str = vip
                                        intf['secondary_ips'].append(sec_str)
                                    i += 1
                            i += 1
                    elif s.startswith("set "):
                        f = get_field_name(lines[i])
                        v = parse_set_value(lines[i])
                        if f == 'ip':
                            parts = v.split()
                            if len(parts) == 2:
                                ip, pf = format_subnet(parts[0], parts[1])
                                intf['ip'] = ip
                                intf['prefix'] = pf
                                intf['display'] = f"{ip}/{pf}" if pf else ip
                            elif '/' in v:
                                intf['display'] = v
                            else:
                                intf['display'] = v
                        elif f == 'remote-ip':
                            parts = v.split()
                            if len(parts) == 2:
                                rip, rpf = format_subnet(parts[0], parts[1])
                                intf['remote-ip'] = f"{rip}/{rpf}" if rpf else rip
                            elif '/' in v:
                                intf['remote-ip'] = v
                            else:
                                intf['remote-ip'] = v
                        elif f == 'allowaccess':
                            intf['allowaccess'] = ' '.join(parse_quoted_values(lines[i])) or v
                        elif f == 'member':
                            intf['member'] = parse_quoted_values(lines[i])
                        else:
                            intf[f] = v
                    i += 1
                vdom = intf.get('vdom', 'root')
                vdom_intfs.setdefault(vdom, []).append(intf)
            i += 1
    return vdom_intfs


def parse_ipsec_vpn(lines, vdom_start, vdom_end):
    """
    config vpn ipsec phase1-interface & phase2-interface 파싱
    Returns list of dicts: each containing phase1 info + list of phase2 entries
    """
    p1_list = []
    p1_dict = OrderedDict()

    # 1. Phase 1 파싱
    p1_ranges = find_section_range(lines, vdom_start, vdom_end, "vpn ipsec phase1-interface")
    for sr, er in p1_ranges:
        i = sr + 1
        while i <= er:
            if lines[i].strip().startswith("edit "):
                p1_name = parse_edit_id(lines[i])
                p1 = {
                    'name': p1_name,
                    'type': 'static',
                    'interface': '',
                    'ike-version': '1',
                    'local-gw': '',
                    'remote-gw': '',
                    'peertype': '',
                    'proposal': '',
                    'dhgrp': '14 5',
                    'nattraversal': 'enable',
                    'keepalive': '10',
                    'dpd': 'on-demand',
                    'dpd-retrycount': '3',
                    'dpd-retryinterval': '20',
                    'fec-egress': 'disable',
                    'fec-ingress': 'disable',
                    'add-gw-route': 'enable',
                    'auto-discovery-sender': 'disable',
                    'auto-discovery-receiver': 'disable',
                    'exchange-interface-ip': 'disable',
                    'net-device': 'disable',
                    'comments': '',
                    'peerid': '',
                    'phase2_list': []
                }
                i += 1
                while i <= er:
                    s = lines[i].strip()
                    if s in ("next", "end"):
                        break
                    if s.startswith("set "):
                        f = get_field_name(lines[i])
                        if f == 'proposal':
                            p1['proposal'] = ' '.join(parse_quoted_values(lines[i])) or parse_set_value(lines[i])
                        else:
                            p1[f] = parse_set_value(lines[i])
                    i += 1
                p1_dict[p1_name] = p1
                p1_list.append(p1)
            i += 1

    # 2. Phase 2 파싱
    p2_ranges = find_section_range(lines, vdom_start, vdom_end, "vpn ipsec phase2-interface")
    for sr, er in p2_ranges:
        i = sr + 1
        while i <= er:
            if lines[i].strip().startswith("edit "):
                p2_name = parse_edit_id(lines[i])
                p2 = {
                    'name': p2_name,
                    'phase1name': '',
                    'proposal': '',
                    'dhgrp': '14 5',
                    'src-addr-type': 'subnet',
                    'src-subnet': '',
                    'src-name': '',
                    'dst-addr-type': 'subnet',
                    'dst-subnet': '',
                    'dst-name': '',
                    'auto-negotiate': 'disable',
                    'keepalive': 'disable',
                    'keylifeseconds': '',
                    'comments': ''
                }
                has_dhgrp = False
                i += 1
                while i <= er:
                    s = lines[i].strip()
                    if s in ("next", "end"):
                        break
                    if s.startswith("set "):
                        f = get_field_name(lines[i])
                        if f == 'dhgrp':
                            p2['dhgrp'] = parse_set_value(lines[i])
                            has_dhgrp = True
                        elif f == 'proposal':
                            p2['proposal'] = ' '.join(parse_quoted_values(lines[i])) or parse_set_value(lines[i])
                        elif f in ('src-subnet', 'dst-subnet'):
                            v = parse_set_value(lines[i])
                            parts = v.split()
                            if len(parts) == 2:
                                ip, pf = format_subnet(parts[0], parts[1])
                                p2[f] = f"{ip}/{pf}" if pf else ip
                            else:
                                p2[f] = v
                        else:
                            p2[f] = parse_set_value(lines[i])
                    i += 1
                p1_target = p2.get('phase1name', '')
                if p1_target in p1_dict:
                    p1_dict[p1_target]['phase2_list'].append(p2)
                else:
                    orphan_p1 = {
                        'name': p1_target or '(Unlinked)',
                        'type': '',
                        'interface': '',
                        'ike-version': '',
                        'local-gw': '',
                        'remote-gw': '',
                        'peertype': '',
                        'proposal': '',
                        'dhgrp': '',
                        'dpd': '',
                        'comments': '',
                        'peerid': '',
                        'phase2_list': [p2]
                    }
                    p1_dict[orphan_p1['name']] = orphan_p1
                    p1_list.append(orphan_p1)
            i += 1

    return p1_list


# ================================================================
#  6. Excel 스타일 & 셀 병합 유틸리티 / Excel Styling & Cell Formatting Utilities
# ================================================================

# --- 컬럼 헤더 배색 / Column Header Fills ---
HDR_DEFAULT_FILL = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid") # 기본 진한 남색 / Default Dark Navy
HDR_SRC_FILL     = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid") # 진한 파랑 (출발지) / Deep Blue (Source)
HDR_DST_FILL     = PatternFill(start_color="843C39", end_color="843C39", fill_type="solid") # 진한 빨강 (목적지) / Deep Red (Destination)
HDR_SVC_FILL     = PatternFill(start_color="415A77", end_color="415A77", fill_type="solid") # 블루그레이 (서비스) / Blue-Gray (Service)
HDR_SCHED_FILL   = PatternFill(start_color="2A5C5A", end_color="2A5C5A", fill_type="solid") # 딥 틸 (스케줄) / Deep Teal (Schedule)
HDR_P1_FILL      = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid") # 딥 네이비 블루 (IPsec Phase 1) / Deep Navy Blue
HDR_P2_FILL      = PatternFill(start_color="2A5C5A", end_color="2A5C5A", fill_type="solid") # 딥 틸 (IPsec Phase 2) / Deep Teal
HDR_FONT         = Font(name="맑은 고딕", size=10, bold=True, color="FFFFFF")

# --- 데이터 행 배경색 (행 전체 열에 일괄 적용) / Data Row Fills (Zebra Striping across all columns) ---
ODD_ROW_FILL     = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid") # 홀수 행 (흰색) / Odd rows (White)
EVEN_ROW_FILL    = PatternFill(start_color="F2F4F8", end_color="F2F4F8", fill_type="solid") # 짝수 행 (소프트 그레이) / Even rows (Soft Gray)
DIS_ROW_FILL     = PatternFill(start_color="D5D8DC", end_color="D5D8DC", fill_type="solid") # 비활성화 행 (진한 회색) / Disabled rows (Dark Gray)

# --- 데이터 폰트 색상 / Data Font Colors ---
FONT_DEFAULT      = Font(name="맑은 고딕", size=9, color="000000")
FONT_DEFAULT_BOLD = Font(name="맑은 고딕", size=9, bold=True, color="000000")

# 출발지 열 폰트 (파란 계열) / Source Column Font (Blue Palette)
FONT_SRC          = Font(name="맑은 고딕", size=9, color="003366")
FONT_SRC_GRP      = Font(name="맑은 고딕", size=9, bold=True, color="002060")

# 목적지 열 폰트 (빨간 계열) / Destination Column Font (Red Palette)
FONT_DST          = Font(name="맑은 고딕", size=9, color="800000")
FONT_DST_GRP      = Font(name="맑은 고딕", size=9, bold=True, color="9C0006")

# 스케줄 열 폰트 (틸 계열) / Schedule Column Font (Teal Palette)
FONT_SCHED        = Font(name="맑은 고딕", size=9, color="1B4D4B")
FONT_SCHED_GRP    = Font(name="맑은 고딕", size=9, bold=True, color="113634")

# 비활성화 행 폰트 / Disabled Row Font Colors
FONT_DIS          = Font(name="맑은 고딕", size=9, color="495057")
FONT_DIS_SRC      = Font(name="맑은 고딕", size=9, color="1B365D")
FONT_DIS_DST      = Font(name="맑은 고딕", size=9, color="6B1D1D")
FONT_DIS_SCHED    = Font(name="맑은 고딕", size=9, color="3D5554")

# Action 및 vDOM 폰트 / Action (Accept / Deny) & vDOM Fonts
ACCEPT_FONT      = Font(name="맑은 고딕", size=9, color="008000", bold=True)
DENY_FONT        = Font(name="맑은 고딕", size=9, color="C00000", bold=True)
SUBHDR_FILL      = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")

THIN_BORDER = Border(
    left=Side(style='thin', color='C0D0E0'),
    right=Side(style='thin', color='C0D0E0'),
    top=Side(style='thin', color='C0D0E0'),
    bottom=Side(style='thin', color='C0D0E0'))
WRAP = Alignment(wrap_text=True, vertical='top')
CENTER = Alignment(horizontal='center', vertical='top', wrap_text=True)


def get_cell_style(col_category, seq, is_disabled, is_group=False, is_action=False, action_val=''):
    """
    행 전체 열에 일관된 배경색(Zebra 격행 / 비활성화 진한 회색)을 적용하고, 영역별(출발지/목적지/액션) 글씨색을 지정
    Apply consistent background fill (Zebra striping / Disabled dark gray) across the entire row and assign specific font colors
    """
    if is_disabled:
        fill = DIS_ROW_FILL
    elif seq % 2 == 0:
        fill = EVEN_ROW_FILL
    else:
        fill = ODD_ROW_FILL

    if is_action:
        if action_val.lower() == 'accept':
            font = ACCEPT_FONT
        elif action_val.lower() in ('deny', 'block', 'drop'):
            font = DENY_FONT
        else:
            font = FONT_DEFAULT
    elif is_disabled:
        if col_category == 'src':
            font = FONT_DIS_SRC
        elif col_category == 'dst':
            font = FONT_DIS_DST
        elif col_category == 'sched':
            font = FONT_DIS_SCHED
        else:
            font = FONT_DIS
    else:
        if col_category == 'src':
            font = FONT_SRC_GRP if is_group else FONT_SRC
        elif col_category == 'dst':
            font = FONT_DST_GRP if is_group else FONT_DST
        elif col_category == 'sched':
            font = FONT_SCHED_GRP if is_group else FONT_SCHED
        else:
            font = FONT_DEFAULT_BOLD if is_group else FONT_DEFAULT

    return fill, font


def sc(ws, r, c, val, font=FONT_DEFAULT, fill=None, align=WRAP):
    """Set cell with styling (High-Performance Cached Style Engine)."""
    cell = ws.cell(row=r, column=c, value=val)
    key = (id(font), id(fill), id(align))
    if not hasattr(ws, '_sc_cache'):
        ws._sc_cache = {}
    st = ws._sc_cache.get(key)
    if st is not None:
        cell._style = st
    else:
        cell.font = font
        cell.alignment = align
        cell.border = THIN_BORDER
        if fill:
            cell.fill = fill
        ws._sc_cache[key] = cell._style
    return cell


def write_styled_header(ws, row, headers_with_cat):
    has_multiline = any('\n' in str(h) for h, _ in headers_with_cat)
    for col, (h, cat) in enumerate(headers_with_cat, 1):
        if cat == 'src':
            fill = HDR_SRC_FILL
        elif cat == 'dst':
            fill = HDR_DST_FILL
        elif cat == 'svc':
            fill = HDR_SVC_FILL
        elif cat == 'sched':
            fill = HDR_SCHED_FILL
        elif cat == 'p1':
            fill = HDR_P1_FILL
        elif cat == 'p2':
            fill = HDR_P2_FILL
        else:
            fill = HDR_DEFAULT_FILL
        sc(ws, row, col, h, font=HDR_FONT, fill=fill, align=CENTER)
    if has_multiline:
        ws.row_dimensions[row].height = 28
    ws.freeze_panes = ws.cell(row=row + 1, column=1).coordinate


_TOP_ALIGN_CACHE = {}


def _get_top_alignment(h_align):
    al = _TOP_ALIGN_CACHE.get(h_align)
    if al is None:
        al = Alignment(horizontal=h_align, vertical='top', wrap_text=True)
        _TOP_ALIGN_CACHE[h_align] = al
    return al


def merge_row_range(ws, start_row, end_row, col, h_align=None):
    if end_row > start_row:
        ws.merged_cells.ranges.add(CellRange(min_row=start_row, min_col=col, max_row=end_row, max_col=col))
        cell = ws.cell(row=start_row, column=col)
        cur_h = h_align or (cell.alignment.horizontal if cell.alignment else 'center')
        cell.alignment = _get_top_alignment(cur_h)


def merge_group_spans(ws, p_start, expanded_list, col, h_align='left'):
    if not expanded_list or len(expanded_list) <= 1:
        return
    cur_grp = expanded_list[0][0]
    s_idx = 0
    for i in range(1, len(expanded_list)):
        grp = expanded_list[i][0]
        if grp != cur_grp:
            if cur_grp and (i - 1) > s_idx:
                merge_row_range(ws, p_start + s_idx, p_start + i - 1, col, h_align=h_align)
            cur_grp = grp
            s_idx = i
    if cur_grp and (len(expanded_list) - 1) > s_idx:
        merge_row_range(ws, p_start + s_idx, p_start + len(expanded_list) - 1, col, h_align=h_align)


def auto_fit(ws, min_w=6, max_w=45):
    # 다중 열 병합 셀(섹션 타이틀 등)을 수집하여 단일 열 너비 과다 확장 방지
    multi_col_merged = set()
    for rng in ws.merged_cells.ranges:
        if rng.min_col < rng.max_col:
            for r in range(rng.min_row, rng.max_row + 1):
                for c in range(rng.min_col, rng.max_col + 1):
                    multi_col_merged.add((r, c))

    for col_cells in ws.columns:
        mx = 0
        cl = get_column_letter(col_cells[0].column)
        col_idx = col_cells[0].column
        for cell in col_cells:
            if (cell.row, col_idx) in multi_col_merged:
                continue
            val = cell.value
            if val is not None:
                s_val = str(val)
                lines = s_val.split('\n') if '\n' in s_val else (s_val,)
                for line in lines:
                    if not line:
                        continue
                    if line.isascii():
                        line_w = len(line)
                    else:
                        line_w = len(line) + sum(1 for ch in line if ord(ch) > 127)
                    if line_w > mx:
                        mx = line_w
        ws.column_dimensions[cl].width = min(max(mx + 2, min_w), max_w)


TAB_COLOR_EMPTY = "FFFF0000"  # 빈 시트 빨간색 탭 색상 (Solid Red ARGB) / Red Tab Color for Empty Sheets


def check_and_mark_empty_sheet_tabs(wb, empty_tab_color=TAB_COLOR_EMPTY):
    """
    컬럼 헤더(1행) 외에 데이터 내용이 전혀 없는 빈 시트의 탭 색상을 빨간색으로 표시
    Mark sheet tab color as red for sheets that have no data content beyond the header row
    """
    for ws in wb.worksheets:
        if ws.title in ("Summary", "vDOM Total Summary"):
            continue
        has_data = False
        if ws.max_row > 1:
            for row_vals in ws.iter_rows(min_row=2, values_only=True):
                if any(v is not None for v in row_vals):
                    has_data = True
                    break
        if not has_data:
            ws.sheet_properties.tabColor = empty_tab_color


def safe_save_workbook(wb, filepath, log_fn=None):
    """
    엑셀 파일 저장 시 Excel 프로그램 등에서 파일이 열려 있어 PermissionError(WinError 32 / [Errno 13])가
    발생하는 경우 사용자에게 명확하고 친절한 안내 메시지를 제공하고 대체 타임스탬프 파일명으로 안전하게 보존합니다.
    """
    try:
        wb.save(filepath)
        return filepath
    except PermissionError as e:
        import time
        ts = time.strftime("%Y%m%d_%H%M%S")
        base, ext = os.path.splitext(filepath)
        alt_path = f"{base}_{ts}{ext}"
        fname = os.path.basename(filepath)
        msg_en = (
            f"[WARNING] '{fname}' is currently locked by Excel or another program and cannot be overwritten.\n"
            f"          Preserving data by saving to timestamped file: '{os.path.basename(alt_path)}'"
        )
        if log_fn:
            log_fn(msg_en)
        else:
            print(msg_en)
        try:
            wb.save(alt_path)
            return alt_path
        except Exception:
            raise PermissionError(
                f"'{fname}' is locked by another process (e.g. Microsoft Excel). "
                f"Please close '{fname}' in Excel and retry. (Error: {e})"
            )



# ================================================================
#  7. 시트 작성 - Firewall Policy / Sheet Builder - Firewall Policy
# ================================================================

def write_fw_policy_sheet(ws, policies, resolver, vdom_name="", dn_ents=None, profile_comments=None, vdom_inspection_mode="flow"):
    headers_with_cat = [
        ("Seq", "base"), ("vDOM", "base"), ("Enable", "base"),
        ("ID", "base"), ("Name", "base"), ("Action", "base"),
        # 출발지 열 (7~12) / Source columns (7-12)
        ("Src Interface", "src"), ("Src Group OBJ", "src"),
        ("Src OBJ Name", "src"), ("Src Type", "src"), ("Src IP", "src"),
        ("Src Comment", "src"),
        # 목적지 열 (13~18) / Destination columns (13-18)
        ("Dst Interface", "dst"), ("Dst Group OBJ", "dst"),
        ("Dst OBJ Name", "dst"), ("Dst Type", "dst"), ("Dst IP", "dst"),
        ("Dst Comment", "dst"),
        # 서비스 열 (19~22) - Protocol 제외 / Service columns (19-22) - Protocol omitted
        ("Svc Group OBJ", "svc"), ("Svc OBJ Name", "svc"),
        ("Svc Port", "svc"), ("Svc Comment", "svc"),
        # 스케줄 열 (23~27) / Schedule columns (23-27)
        ("Sched Group OBJ", "sched"), ("Sched OBJ Name", "sched"),
        ("Sched Type", "sched"), ("Sched Time", "sched"),
        ("Sched Comment", "sched"),
        # 기타 필드 열 (28~37) / Miscellaneous columns (28-37)
        ("NAT", "base"), ("IP Pool", "base"),
        ("Pool Name", "base"), ("Pool IP", "base"),
        ("Inspection Mode", "base"), ("UTM Status", "base"),
        ("Sec Profile", "base"), ("Sec Profile Comment", "base"),
        ("Log Traffic", "base"), ("Comment", "base")
    ]
    write_styled_header(ws, 1, headers_with_cat)
    ncol = len(headers_with_cat)

    vip_type_map = {}
    if dn_ents:
        for v in dn_ents:
            if isinstance(v, dict) and v.get('name'):
                vip_type_map[v['name']] = v.get('type', 'static-nat')

    row = 2
    seq = 0
    # Src Interface 기준 오름차순 안정 정렬 (동일 인터페이스 내 원래 순서 유지) / Stable sort by Src Interface
    sorted_policies = sorted(policies, key=lambda p: ', '.join(p.get('srcintf') or []))

    for p in sorted_policies:
        seq += 1
        is_dis = p.get('status', 'enable') == 'disable'
        action = p.get('action', 'deny')

        b_fill, b_font = get_cell_style('base', seq, is_dis)
        s_fill, s_font = get_cell_style('src', seq, is_dis)
        sg_fill, sg_font = get_cell_style('src', seq, is_dis, is_group=True)
        d_fill, d_font = get_cell_style('dst', seq, is_dis)
        dg_fill, dg_font = get_cell_style('dst', seq, is_dis, is_group=True)
        v_fill, v_font = get_cell_style('svc', seq, is_dis)
        vg_fill, vg_font = get_cell_style('svc', seq, is_dis, is_group=True)
        sc_fill, sc_font = get_cell_style('sched', seq, is_dis)
        scg_fill, scg_font = get_cell_style('sched', seq, is_dis, is_group=True)
        act_fill, act_font = get_cell_style('base', seq, is_dis, is_action=True, action_val=action)

        src_expanded = []
        for addr_name in (p.get('srcaddr') or []):
            src_expanded.extend(resolver.resolve_address(addr_name))
        for isn in (p.get('internet-service-src-name') or []):
            src_expanded.append((None, isn, 'internet-service', '', ''))
        if not src_expanded:
            src_expanded = [(None, 'all', 'ipmask', '0.0.0.0/0', '')]

        dst_expanded = []
        for addr_name in (p.get('dstaddr') or []):
            dst_expanded.extend(resolver.resolve_address(addr_name))
        for isn in (p.get('internet-service-name') or []):
            dst_expanded.append((None, isn, 'internet-service', '', ''))
        if not dst_expanded:
            dst_expanded = [(None, 'all', 'ipmask', '0.0.0.0/0', '')]

        svc_expanded = []
        for svc_name in (p.get('service') or []):
            svc_expanded.extend(resolver.resolve_service(svc_name))

        sched_name = p.get('schedule', '')
        sched_expanded = resolver.resolve_schedule(sched_name)

        pool_items = []
        for pn in (p.get('poolname') or []):
            pool_items.append(resolver.resolve_ippool(pn))

        sec_profiles = []
        sec_profile_comments = []
        for pk in ['ips-sensor', 'av-profile', 'webfilter-profile', 'dnsfilter-profile',
                   'application-list', 'file-filter-profile', 'emailfilter-profile',
                   'dlp-profile', 'dlp-sensor', 'waf-profile', 'casb-profile',
                   'videofilter-profile', 'sctp-filter-profile', 'cifs-profile',
                   'virtual-patch-profile', 'voip-profile']:
            val = p.get(pk)
            if val:
                sec_profiles.append(val)
                comm = profile_comments.get(val, '') if profile_comments else ''
                sec_profile_comments.append(comm)

        ssl_prof = p.get('ssl-ssh-profile')
        if ssl_prof:
            sec_profiles.append(ssl_prof)
            comm = profile_comments.get(ssl_prof, '') if profile_comments else ''
            sec_profile_comments.append(comm)
        elif action != 'deny':
            sec_profiles.append('no-inspection')
            comm = profile_comments.get('no-inspection', '') if profile_comments else ''
            sec_profile_comments.append(comm)

        sec_profile_str = '\n'.join(sec_profiles)
        sec_profile_comment_str = '\n'.join(sec_profile_comments)

        raw_insp = p.get('inspection-mode') or vdom_inspection_mode or 'flow'
        insp_mode_display = 'Proxy-based' if raw_insp.lower() == 'proxy' else 'Flow-based'

        lt = p.get('logtraffic', '')
        lt_start = p.get('logtraffic-start', '')
        if lt == 'disable':
            log_traffic_str = 'disable'
        else:
            base_log = 'all' if lt == 'all' else 'utm'
            logs = [base_log]
            if lt_start == 'enable':
                logs.append('session-start')
            log_traffic_str = '\n'.join(logs)

        max_rows = max(len(src_expanded), len(dst_expanded),
                       len(svc_expanded), len(pool_items),
                       len(sched_expanded), 1)
        p_start = row
        p_end = row + max_rows - 1

        for ri in range(max_rows):
            cur_r = row + ri

            if ri == 0:
                sc(ws, cur_r, 1, seq, font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 2, vdom_name, font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 3, 'N' if is_dis else 'Y', font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 4, p.get('id', ''), font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 5, p.get('name', ''), font=b_font, fill=b_fill)
                sc(ws, cur_r, 6, action.upper(), font=act_font, fill=act_fill, align=CENTER)
            else:
                for c in range(1, 7):
                    sc(ws, cur_r, c, None, font=b_font, fill=b_fill)

            # 출발지 열 (7~12) / Source columns (7-12)
            if ri == 0:
                sc(ws, cur_r, 7, '\n'.join(p.get('srcintf') or []), font=s_font, fill=s_fill)
            else:
                sc(ws, cur_r, 7, None, font=s_font, fill=s_fill)

            if ri < len(src_expanded):
                grp, obj_name, obj_type, ip, comm = src_expanded[ri]
                sc(ws, cur_r, 8, grp or '', font=sg_font if grp else s_font, fill=s_fill)
                sc(ws, cur_r, 9, obj_name, font=s_font, fill=s_fill)
                sc(ws, cur_r, 10, obj_type, font=s_font, fill=s_fill, align=CENTER)
                sc(ws, cur_r, 11, ip, font=s_font, fill=s_fill)
                sc(ws, cur_r, 12, comm, font=s_font, fill=s_fill)
            else:
                for c in range(8, 13):
                    sc(ws, cur_r, c, None, font=s_font, fill=s_fill)

            # 목적지 열 (13~18) / Destination columns (13-18)
            if ri == 0:
                sc(ws, cur_r, 13, '\n'.join(p.get('dstintf') or []), font=d_font, fill=d_fill)
            else:
                sc(ws, cur_r, 13, None, font=d_font, fill=d_fill)

            if ri < len(dst_expanded):
                grp, obj_name, obj_type, ip, comm = dst_expanded[ri]
                if vip_type_map and obj_name in vip_type_map:
                    obj_type = vip_type_map[obj_name]
                sc(ws, cur_r, 14, grp or '', font=dg_font if grp else d_font, fill=d_fill)
                sc(ws, cur_r, 15, obj_name, font=d_font, fill=d_fill)
                sc(ws, cur_r, 16, obj_type, font=d_font, fill=d_fill, align=CENTER)
                sc(ws, cur_r, 17, ip, font=d_font, fill=d_fill)
                sc(ws, cur_r, 18, comm, font=d_font, fill=d_fill)
            else:
                for c in range(14, 19):
                    sc(ws, cur_r, c, None, font=d_font, fill=d_fill)

            # 서비스 열 (19~22) / Service columns (19-22)
            if ri < len(svc_expanded):
                sgrp, sname, sport, scomm = svc_expanded[ri]
                sc(ws, cur_r, 19, sgrp or '', font=vg_font if sgrp else v_font, fill=v_fill)
                sc(ws, cur_r, 20, sname, font=v_font, fill=v_fill)
                sc(ws, cur_r, 21, sport, font=v_font, fill=v_fill)
                sc(ws, cur_r, 22, scomm, font=v_font, fill=v_fill)
            else:
                for c in range(19, 23):
                    sc(ws, cur_r, c, None, font=v_font, fill=v_fill)

            # 스케줄 열 (23~27) / Schedule columns (23-27)
            if ri < len(sched_expanded):
                scgrp, scname, sctype, sctime, sccomm = sched_expanded[ri]
                sc(ws, cur_r, 23, scgrp or '', font=scg_font if scgrp else sc_font, fill=sc_fill)
                sc(ws, cur_r, 24, scname, font=sc_font, fill=sc_fill)
                sc(ws, cur_r, 25, sctype, font=sc_font, fill=sc_fill, align=CENTER)
                sc(ws, cur_r, 26, sctime, font=sc_font, fill=sc_fill)
                sc(ws, cur_r, 27, sccomm, font=sc_font, fill=sc_fill)
            else:
                for c in range(23, 28):
                    sc(ws, cur_r, c, None, font=sc_font, fill=sc_fill)

            # 기타 필드 열 (28~37) / Miscellaneous columns (28-37)
            if ri == 0:
                utm_val = 'enable' if p.get('utm-status') == 'enable' else 'disable'
                sc(ws, cur_r, 28, p.get('nat', ''), font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 29, p.get('ippool', ''), font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 32, insp_mode_display, font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 33, utm_val, font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 34, sec_profile_str, font=b_font, fill=b_fill)
                sc(ws, cur_r, 35, sec_profile_comment_str, font=b_font, fill=b_fill)
                sc(ws, cur_r, 36, log_traffic_str, font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 37, p.get('comments', ''), font=b_font, fill=b_fill)
            else:
                for c in [28, 29, 32, 33, 34, 35, 36, 37]:
                    sc(ws, cur_r, c, None, font=b_font, fill=b_fill)

            if ri < len(pool_items):
                pname, pip, ptype = pool_items[ri]
                sc(ws, cur_r, 30, pname, font=b_font, fill=b_fill)
                sc(ws, cur_r, 31, pip, font=b_font, fill=b_fill)
            else:
                sc(ws, cur_r, 30, None, font=b_font, fill=b_fill)
                sc(ws, cur_r, 31, None, font=b_font, fill=b_fill)

        if p_end > p_start:
            common_cols = [1, 2, 3, 4, 5, 6, 7, 13, 28, 29, 32, 33, 34, 35, 36, 37]
            for c in common_cols:
                merge_row_range(ws, p_start, p_end, c)
            if len(pool_items) <= 1:
                merge_row_range(ws, p_start, p_end, 30)
                merge_row_range(ws, p_start, p_end, 31)

            merge_group_spans(ws, p_start, src_expanded, 8, h_align='left')
            merge_group_spans(ws, p_start, dst_expanded, 14, h_align='left')
            merge_group_spans(ws, p_start, svc_expanded, 19, h_align='left')
            merge_group_spans(ws, p_start, sched_expanded, 23, h_align='left')

        row += max_rows

    auto_fit(ws)
    return ws


# ================================================================
#  8. 시트 작성 - Local-in Policy / Sheet Builder - Local-in Policy
# ================================================================

def write_local_in_sheet(ws, policies, resolver, vdom_name=""):
    headers_with_cat = [
        ("Seq", "base"), ("vDOM", "base"), ("Enable", "base"),
        ("ID", "base"), ("Interface", "base"),
        # 출발지 열 (6~10) / Source columns (6-10)
        ("Src Group OBJ", "src"), ("Src OBJ Name", "src"),
        ("Src Type", "src"), ("Src IP", "src"), ("Src Comment", "src"),
        # 목적지 열 (11~15) / Destination columns (11-15)
        ("Dst Group OBJ", "dst"), ("Dst OBJ Name", "dst"),
        ("Dst Type", "dst"), ("Dst IP", "dst"), ("Dst Comment", "dst"),
        ("Action", "base"),
        # 서비스 열 (17~20) - Protocol 제외 / Service columns (17-20) - Protocol omitted
        ("Svc Group OBJ", "svc"), ("Svc OBJ Name", "svc"),
        ("Svc Port", "svc"), ("Svc Comment", "svc"),
        # 스케줄 열 (21~25) / Schedule columns (21-25)
        ("Sched Group OBJ", "sched"), ("Sched OBJ Name", "sched"),
        ("Sched Type", "sched"), ("Sched Time", "sched"),
        ("Sched Comment", "sched"),
        ("Comment", "base")
    ]
    write_styled_header(ws, 1, headers_with_cat)
    ncol = len(headers_with_cat)

    row = 2
    seq = 0
    for p in policies:
        seq += 1
        is_dis = p.get('status', 'enable') == 'disable'
        action = p.get('action', 'deny')

        b_fill, b_font = get_cell_style('base', seq, is_dis)
        s_fill, s_font = get_cell_style('src', seq, is_dis)
        sg_fill, sg_font = get_cell_style('src', seq, is_dis, is_group=True)
        d_fill, d_font = get_cell_style('dst', seq, is_dis)
        dg_fill, dg_font = get_cell_style('dst', seq, is_dis, is_group=True)
        v_fill, v_font = get_cell_style('svc', seq, is_dis)
        vg_fill, vg_font = get_cell_style('svc', seq, is_dis, is_group=True)
        sc_fill, sc_font = get_cell_style('sched', seq, is_dis)
        scg_fill, scg_font = get_cell_style('sched', seq, is_dis, is_group=True)
        act_fill, act_font = get_cell_style('base', seq, is_dis, is_action=True, action_val=action)

        src_expanded = []
        for a in (p.get('srcaddr') or []):
            src_expanded.extend(resolver.resolve_address(a))
        dst_expanded = []
        for a in (p.get('dstaddr') or []):
            dst_expanded.extend(resolver.resolve_address(a))
        svc_expanded = []
        for s in (p.get('service') or []):
            svc_expanded.extend(resolver.resolve_service(s))
        sched_name = p.get('schedule', '')
        sched_expanded = resolver.resolve_schedule(sched_name)

        max_rows = max(len(src_expanded), len(dst_expanded), len(svc_expanded), len(sched_expanded), 1)
        p_start = row
        p_end = row + max_rows - 1

        for ri in range(max_rows):
            cur_r = row + ri
            if ri == 0:
                sc(ws, cur_r, 1, seq, font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 2, vdom_name, font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 3, 'N' if is_dis else 'Y', font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 4, p.get('id', ''), font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 5, p.get('intf', ''), font=b_font, fill=b_fill)
                sc(ws, cur_r, 16, action.upper(), font=act_font, fill=act_fill, align=CENTER)
                sc(ws, cur_r, 26, p.get('comments', ''), font=b_font, fill=b_fill)
            else:
                for c in [1, 2, 3, 4, 5, 16, 26]:
                    sc(ws, cur_r, c, None, font=b_font, fill=b_fill)

            # 출발지 / Src (Source)
            if ri < len(src_expanded):
                grp, n, t, ip, comm = src_expanded[ri]
                sc(ws, cur_r, 6, grp or '', font=sg_font if grp else s_font, fill=s_fill)
                sc(ws, cur_r, 7, n, font=s_font, fill=s_fill)
                sc(ws, cur_r, 8, t, font=s_font, fill=s_fill, align=CENTER)
                sc(ws, cur_r, 9, ip, font=s_font, fill=s_fill)
                sc(ws, cur_r, 10, comm, font=s_font, fill=s_fill)
            else:
                for c in range(6, 11):
                    sc(ws, cur_r, c, None, font=s_font, fill=s_fill)

            # 목적지 / Dst (Destination)
            if ri < len(dst_expanded):
                grp, n, t, ip, comm = dst_expanded[ri]
                sc(ws, cur_r, 11, grp or '', font=dg_font if grp else d_font, fill=d_fill)
                sc(ws, cur_r, 12, n, font=d_font, fill=d_fill)
                sc(ws, cur_r, 13, t, font=d_font, fill=d_fill, align=CENTER)
                sc(ws, cur_r, 14, ip, font=d_font, fill=d_fill)
                sc(ws, cur_r, 15, comm, font=d_font, fill=d_fill)
            else:
                for c in range(11, 16):
                    sc(ws, cur_r, c, None, font=d_font, fill=d_fill)

            # 서비스 / Service
            if ri < len(svc_expanded):
                sgrp, sn, sport, scomm = svc_expanded[ri]
                sc(ws, cur_r, 17, sgrp or '', font=vg_font if sgrp else v_font, fill=v_fill)
                sc(ws, cur_r, 18, sn, font=v_font, fill=v_fill)
                sc(ws, cur_r, 19, sport, font=v_font, fill=v_fill)
                sc(ws, cur_r, 20, scomm, font=v_font, fill=v_fill)
            else:
                for c in range(17, 21):
                    sc(ws, cur_r, c, None, font=v_font, fill=v_fill)

            # 스케줄 / Schedule
            if ri < len(sched_expanded):
                scgrp, scname, sctype, sctime, sccomm = sched_expanded[ri]
                sc(ws, cur_r, 21, scgrp or '', font=scg_font if scgrp else sc_font, fill=sc_fill)
                sc(ws, cur_r, 22, scname, font=sc_font, fill=sc_fill)
                sc(ws, cur_r, 23, sctype, font=sc_font, fill=sc_fill, align=CENTER)
                sc(ws, cur_r, 24, sctime, font=sc_font, fill=sc_fill)
                sc(ws, cur_r, 25, sccomm, font=sc_font, fill=sc_fill)
            else:
                for c in range(21, 26):
                    sc(ws, cur_r, c, None, font=sc_font, fill=sc_fill)

        if p_end > p_start:
            common_cols = [1, 2, 3, 4, 5, 16, 26]
            for c in common_cols:
                merge_row_range(ws, p_start, p_end, c)
            merge_group_spans(ws, p_start, src_expanded, 6, h_align='left')
            merge_group_spans(ws, p_start, dst_expanded, 11, h_align='left')
            merge_group_spans(ws, p_start, svc_expanded, 17, h_align='left')
            merge_group_spans(ws, p_start, sched_expanded, 21, h_align='left')

        row += max_rows

    auto_fit(ws)
    return ws


# ================================================================
#  9. 시트 작성 - Central-NAT / Sheet Builder - Central-NAT
# ================================================================

def write_central_nat_sheet(ws, entries, resolver, vdom_name=""):
    headers_with_cat = [
        ("Seq", "base"), ("vDOM", "base"), ("Enable", "base"), ("ID", "base"),
        ("Src Interface", "src"), ("Dst Interface", "dst"),
        # 원래 출발지 열 (7~11) / Original Source columns (7-11)
        ("Orig Group OBJ", "src"), ("Orig OBJ Name", "src"),
        ("Orig Type", "src"), ("Orig IP", "src"), ("Orig Comment", "src"),
        # 목적지 열 (12~16) / Destination columns (12-16)
        ("Dst Group OBJ", "dst"), ("Dst OBJ Name", "dst"),
        ("Dst Type", "dst"), ("Dst IP", "dst"), ("Dst Comment", "dst"),
        # IP 풀 및 기타 열 (17~21) / IP Pool & Other columns (17-21)
        ("NAT IP Pool Name", "base"), ("NAT Pool IP", "base"),
        ("NAT Pool Type", "base"), ("NAT", "base"), ("Comment", "base")
    ]
    write_styled_header(ws, 1, headers_with_cat)
    ncol = len(headers_with_cat)

    row = 2
    seq = 0
    for e in entries:
        seq += 1
        is_dis = e.get('status', 'enable') == 'disable'

        b_fill, b_font = get_cell_style('base', seq, is_dis)
        s_fill, s_font = get_cell_style('src', seq, is_dis)
        sg_fill, sg_font = get_cell_style('src', seq, is_dis, is_group=True)
        d_fill, d_font = get_cell_style('dst', seq, is_dis)
        dg_fill, dg_font = get_cell_style('dst', seq, is_dis, is_group=True)

        orig_expanded = []
        for a in (e.get('orig-addr') or []):
            orig_expanded.extend(resolver.resolve_address(a))
        dst_expanded = []
        for a in (e.get('dst-addr') or []):
            dst_expanded.extend(resolver.resolve_address(a))
        pool_items = []
        for pn in (e.get('nat-ippool') or []):
            pool_items.append(resolver.resolve_ippool(pn))

        max_rows = max(len(orig_expanded), len(dst_expanded), len(pool_items), 1)
        p_start = row
        p_end = row + max_rows - 1

        for ri in range(max_rows):
            cur_r = row + ri
            if ri == 0:
                sc(ws, cur_r, 1, seq, font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 2, vdom_name, font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 3, 'N' if is_dis else 'Y', font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 4, e.get('id', ''), font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 5, '\n'.join(e.get('srcintf') or []), font=s_font, fill=s_fill)
                sc(ws, cur_r, 6, '\n'.join(e.get('dstintf') or []), font=d_font, fill=d_fill)
                nat_val = 'disable' if e.get('nat') == 'disable' else 'enable'
                sc(ws, cur_r, 20, nat_val, font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 21, e.get('comments', ''), font=b_font, fill=b_fill)
            else:
                for c in [1, 2, 3, 4, 20, 21]:
                    sc(ws, cur_r, c, None, font=b_font, fill=b_fill)
                sc(ws, cur_r, 5, None, font=s_font, fill=s_fill)
                sc(ws, cur_r, 6, None, font=d_font, fill=d_fill)

            # 원래 출발지 / Orig (Source)
            if ri < len(orig_expanded):
                grp, n, t, ip, comm = orig_expanded[ri]
                sc(ws, cur_r, 7, grp or '', font=sg_font if grp else s_font, fill=s_fill)
                sc(ws, cur_r, 8, n, font=s_font, fill=s_fill)
                sc(ws, cur_r, 9, t, font=s_font, fill=s_fill, align=CENTER)
                sc(ws, cur_r, 10, ip, font=s_font, fill=s_fill)
                sc(ws, cur_r, 11, comm, font=s_font, fill=s_fill)
            else:
                for c in range(7, 12):
                    sc(ws, cur_r, c, None, font=s_font, fill=s_fill)

            # 목적지 / Dst (Destination)
            if ri < len(dst_expanded):
                grp, n, t, ip, comm = dst_expanded[ri]
                sc(ws, cur_r, 12, grp or '', font=dg_font if grp else d_font, fill=d_fill)
                sc(ws, cur_r, 13, n, font=d_font, fill=d_fill)
                sc(ws, cur_r, 14, t, font=d_font, fill=d_fill, align=CENTER)
                sc(ws, cur_r, 15, ip, font=d_font, fill=d_fill)
                sc(ws, cur_r, 16, comm, font=d_font, fill=d_fill)
            else:
                for c in range(12, 17):
                    sc(ws, cur_r, c, None, font=d_font, fill=d_fill)

            # IP 풀 / IP Pool
            if ri < len(pool_items):
                pname, pip, ptype = pool_items[ri]
                sc(ws, cur_r, 17, pname, font=b_font, fill=b_fill)
                sc(ws, cur_r, 18, pip, font=b_font, fill=b_fill)
                sc(ws, cur_r, 19, ptype, font=b_font, fill=b_fill, align=CENTER)
            else:
                for c in range(17, 20):
                    sc(ws, cur_r, c, None, font=b_font, fill=b_fill)

        if p_end > p_start:
            common_cols = [1, 2, 3, 4, 5, 6, 20, 21]
            for c in common_cols:
                merge_row_range(ws, p_start, p_end, c)
            if len(pool_items) <= 1:
                merge_row_range(ws, p_start, p_end, 17)
                merge_row_range(ws, p_start, p_end, 18)
                merge_row_range(ws, p_start, p_end, 19)
            merge_group_spans(ws, p_start, orig_expanded, 7, h_align='left')
            merge_group_spans(ws, p_start, dst_expanded, 12, h_align='left')

        row += max_rows

    auto_fit(ws)
    return ws


# ================================================================
# 10. 시트 작성 - DNAT (VIP) / Sheet Builder - DNAT (VIP)
# ================================================================

def write_dnat_sheet(ws, entries, resolver=None, vdom_name=""):
    headers_with_cat = [
        ("Seq", "base"), ("vDOM", "base"), ("Name", "base"), ("Type", "base"),
        ("External IP", "src"), ("Mapped IP", "dst"), ("External Interface", "src"),
        ("Port Forward", "base"), ("Protocol", "base"),
        ("External Port", "src"), ("Mapped Port", "dst"),
        # 서비스 열 (12~15) / Service columns (12-15)
        ("Svc Group OBJ", "svc"), ("Svc OBJ Name", "svc"),
        ("Svc Port", "svc"), ("Svc Comment", "svc"),
        ("ARP Reply", "base"), ("NAT Source VIP", "base"),
        ("Server Type", "base"), ("LDB Method", "base"), ("Monitor", "base"),
        ("Real Server IP", "dst"), ("Real Server Port", "dst"), ("Real Server Weight", "dst"),
        ("Comment", "base")
    ]
    write_styled_header(ws, 1, headers_with_cat)
    ncol = len(headers_with_cat)

    row = 2
    seq = 0
    for e in entries:
        seq += 1
        b_fill, b_font = get_cell_style('base', seq, False)
        s_fill, s_font = get_cell_style('src', seq, False)
        d_fill, d_font = get_cell_style('dst', seq, False)
        v_fill, v_font = get_cell_style('svc', seq, False)
        vg_fill, vg_font = get_cell_style('svc', seq, False, is_group=True)

        # Service expansion via resolver
        svc_expanded = []
        raw_services = e.get('service') or []
        if resolver:
            for sname in raw_services:
                svc_expanded.extend(resolver.resolve_service(sname))
        else:
            for sname in raw_services:
                svc_expanded.append((None, sname, '', ''))

        rs_list = e.get('realservers', [])
        max_rows = max(len(rs_list), len(svc_expanded), 1)
        p_start = row
        p_end = row + max_rows - 1

        for ri in range(max_rows):
            cur_r = row + ri
            if ri == 0:
                sc(ws, cur_r, 1, seq, font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 2, vdom_name, font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 3, e.get('name', ''), font=b_font, fill=b_fill)
                sc(ws, cur_r, 4, e.get('type', 'static-nat'), font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 5, e.get('extip', ''), font=s_font, fill=s_fill)
                sc(ws, cur_r, 6, e.get('mappedip', ''), font=d_font, fill=d_fill)
                sc(ws, cur_r, 7, e.get('extintf', ''), font=s_font, fill=s_fill)
                sc(ws, cur_r, 8, e.get('portforward', ''), font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 9, e.get('protocol', ''), font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 10, e.get('extport', ''), font=s_font, fill=s_fill, align=CENTER)
                sc(ws, cur_r, 11, e.get('mappedport', ''), font=d_font, fill=d_fill, align=CENTER)
                sc(ws, cur_r, 16, e.get('arp-reply', 'enable'), font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 17, e.get('nat-source-vip', 'disable'), font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 18, e.get('server-type', ''), font=b_font, fill=b_fill)
                sc(ws, cur_r, 19, e.get('ldb-method', ''), font=b_font, fill=b_fill)
                sc(ws, cur_r, 20, e.get('monitor', ''), font=b_font, fill=b_fill)
                sc(ws, cur_r, 24, e.get('comment', ''), font=b_font, fill=b_fill)
            else:
                for c in [1, 2, 3, 4, 8, 9, 16, 17, 18, 19, 20, 24]:
                    sc(ws, cur_r, c, None, font=b_font, fill=b_fill)
                for c in [5, 7, 10]:
                    sc(ws, cur_r, c, None, font=s_font, fill=s_fill)
                for c in [6, 11]:
                    sc(ws, cur_r, c, None, font=d_font, fill=d_fill)

            # Service columns (12~15)
            if ri < len(svc_expanded):
                sgrp, sname, sport, scomm = svc_expanded[ri]
                sc(ws, cur_r, 12, sgrp or '', font=vg_font if sgrp else v_font, fill=v_fill)
                sc(ws, cur_r, 13, sname, font=v_font, fill=v_fill)
                sc(ws, cur_r, 14, sport, font=v_font, fill=v_fill)
                sc(ws, cur_r, 15, scomm, font=v_font, fill=v_fill)
            else:
                for c in range(12, 16):
                    sc(ws, cur_r, c, None, font=v_font, fill=v_fill)

            # Real Server columns (21~23)
            if ri < len(rs_list):
                rs = rs_list[ri]
                sc(ws, cur_r, 21, rs.get('ip', ''), font=d_font, fill=d_fill)
                sc(ws, cur_r, 22, rs.get('port', ''), font=d_font, fill=d_fill, align=CENTER)
                sc(ws, cur_r, 23, rs.get('weight', '1'), font=d_font, fill=d_fill, align=CENTER)
            else:
                for c in [21, 22, 23]:
                    sc(ws, cur_r, c, None, font=d_font, fill=d_fill)

        if p_end > p_start:
            common_cols = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 16, 17, 18, 19, 20, 24]
            for c in common_cols:
                merge_row_range(ws, p_start, p_end, c)
            merge_group_spans(ws, p_start, svc_expanded, 12, h_align='left')

        row += max_rows

    auto_fit(ws)
    return ws


# ================================================================
# 11. 시트 작성 - DoS Policy / Sheet Builder - DoS Policy
# ================================================================

def write_dos_sheet(ws, policies, resolver, vdom_name=""):
    headers_with_cat = [
        ("Seq", "base"), ("vDOM", "base"), ("Enable", "base"),
        ("ID", "base"), ("Interface", "base"),
        # 출발지 열 (6~10) / Source columns (6-10)
        ("Src Group OBJ", "src"), ("Src OBJ Name", "src"),
        ("Src Type", "src"), ("Src IP", "src"), ("Src Comment", "src"),
        # 목적지 열 (11~15) / Destination columns (11-15)
        ("Dst Group OBJ", "dst"), ("Dst OBJ Name", "dst"),
        ("Dst Type", "dst"), ("Dst IP", "dst"), ("Dst Comment", "dst"),
        # 서비스 열 (16~19) - Protocol 제외 / Service columns (16-19) - Protocol omitted
        ("Svc Group OBJ", "svc"), ("Svc OBJ Name", "svc"),
        ("Svc Port", "svc"), ("Svc Comment", "svc"),
        # 아노말리 변칙 탐지 열 (20~25) / Anomaly detection columns (20-25)
        ("Anomaly Name", "base"), ("Anomaly Status", "base"),
        ("Anomaly Log", "base"), ("Anomaly Quarant", "base"),
        ("Anomaly Action", "base"), ("Anomaly Threshold", "base"),
        ("Comment", "base")
    ]
    write_styled_header(ws, 1, headers_with_cat)
    ncol = len(headers_with_cat)

    row = 2
    seq = 0
    for p in policies:
        seq += 1
        is_dis = p.get('status', 'enable') == 'disable'

        b_fill, b_font = get_cell_style('base', seq, is_dis)
        s_fill, s_font = get_cell_style('src', seq, is_dis)
        sg_fill, sg_font = get_cell_style('src', seq, is_dis, is_group=True)
        d_fill, d_font = get_cell_style('dst', seq, is_dis)
        dg_fill, dg_font = get_cell_style('dst', seq, is_dis, is_group=True)
        v_fill, v_font = get_cell_style('svc', seq, is_dis)
        vg_fill, vg_font = get_cell_style('svc', seq, is_dis, is_group=True)

        src_expanded = []
        for a in (p.get('srcaddr') or []):
            src_expanded.extend(resolver.resolve_address(a))
        dst_expanded = []
        for a in (p.get('dstaddr') or []):
            dst_expanded.extend(resolver.resolve_address(a))
        svc_expanded = []
        for s in (p.get('service') or []):
            svc_expanded.extend(resolver.resolve_service(s))
        anomalies = p.get('anomalies', [])

        max_rows = max(len(src_expanded), len(dst_expanded),
                       len(svc_expanded), len(anomalies), 1)
        p_start = row
        p_end = row + max_rows - 1

        for ri in range(max_rows):
            cur_r = row + ri
            if ri == 0:
                sc(ws, cur_r, 1, seq, font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 2, vdom_name, font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 3, 'N' if is_dis else 'Y', font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 4, p.get('id', ''), font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 5, p.get('interface', ''), font=b_font, fill=b_fill)
                sc(ws, cur_r, 26, p.get('comments', ''), font=b_font, fill=b_fill)
            else:
                for c in [1, 2, 3, 4, 5, 26]:
                    sc(ws, cur_r, c, None, font=b_font, fill=b_fill)

            # 출발지 / Src (Source)
            if ri < len(src_expanded):
                grp, n, t, ip, comm = src_expanded[ri]
                sc(ws, cur_r, 6, grp or '', font=sg_font if grp else s_font, fill=s_fill)
                sc(ws, cur_r, 7, n, font=s_font, fill=s_fill)
                sc(ws, cur_r, 8, t, font=s_font, fill=s_fill, align=CENTER)
                sc(ws, cur_r, 9, ip, font=s_font, fill=s_fill)
                sc(ws, cur_r, 10, comm, font=s_font, fill=s_fill)
            else:
                for c in range(6, 11):
                    sc(ws, cur_r, c, None, font=s_font, fill=s_fill)

            # 목적지 / Dst (Destination)
            if ri < len(dst_expanded):
                grp, n, t, ip, comm = dst_expanded[ri]
                sc(ws, cur_r, 11, grp or '', font=dg_font if grp else d_font, fill=d_fill)
                sc(ws, cur_r, 12, n, font=d_font, fill=d_fill)
                sc(ws, cur_r, 13, t, font=d_font, fill=d_fill, align=CENTER)
                sc(ws, cur_r, 14, ip, font=d_font, fill=d_fill)
                sc(ws, cur_r, 15, comm, font=d_font, fill=d_fill)
            else:
                for c in range(11, 16):
                    sc(ws, cur_r, c, None, font=d_font, fill=d_fill)

            # 서비스 / Service
            if ri < len(svc_expanded):
                sgrp, sn, sport, scomm = svc_expanded[ri]
                sc(ws, cur_r, 16, sgrp or '', font=vg_font if sgrp else v_font, fill=v_fill)
                sc(ws, cur_r, 17, sn, font=v_font, fill=v_fill)
                sc(ws, cur_r, 18, sport, font=v_font, fill=v_fill)
                sc(ws, cur_r, 19, scomm, font=v_font, fill=v_fill)
            else:
                for c in range(16, 20):
                    sc(ws, cur_r, c, None, font=v_font, fill=v_fill)

            # 변칙 탐지 / Anomaly
            if ri < len(anomalies):
                a = anomalies[ri]
                aact = a.get('action') or 'disable'
                if aact.lower() in ('block', 'deny'):
                    afont = DENY_FONT
                elif aact.lower() in ('pass', 'accept'):
                    afont = ACCEPT_FONT
                else:
                    afont = b_font

                astatus = a.get('status') or 'disable'
                alog = a.get('log') or 'disable'
                aquar = a.get('quarantine') or 'disable'
                if aquar.lower() != 'attacker':
                    aquar = 'disable'
                athresh = a.get('threshold', '')

                sc(ws, cur_r, 20, a.get('name', ''), font=b_font, fill=b_fill)
                sc(ws, cur_r, 21, astatus, font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 22, alog, font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 23, aquar, font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 24, aact, font=afont, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 25, athresh, font=b_font, fill=b_fill, align=CENTER)
            else:
                for c in range(20, 26):
                    sc(ws, cur_r, c, None, font=b_font, fill=b_fill)

        if p_end > p_start:
            common_cols = [1, 2, 3, 4, 5, 26]
            for c in common_cols:
                merge_row_range(ws, p_start, p_end, c)
            merge_group_spans(ws, p_start, src_expanded, 6, h_align='left')
            merge_group_spans(ws, p_start, dst_expanded, 11, h_align='left')
            merge_group_spans(ws, p_start, svc_expanded, 16, h_align='left')

        row += max_rows

    auto_fit(ws)
    return ws


# ================================================================
# 12. 시트 작성 - ACL Policy / Sheet Builder - ACL Policy (config firewall acl)
# ================================================================

def write_acl_sheet(ws, policies, resolver, vdom_name=""):
    """
    'ACL Policy' 시트 작성 / Sheet Builder - ACL Policy (config firewall acl)
    출발지/목적지 그룹 객체 자동 전개, 서비스 그룹 전개 지원
    """
    headers_with_cat = [
        ("Seq", "base"), ("vDOM", "base"), ("Enable", "base"),
        ("ID", "base"), ("Interface", "base"),
        # 출발지 열 (6~10) / Source columns (6-10)
        ("Src Group OBJ", "src"), ("Src OBJ Name", "src"),
        ("Src Type", "src"), ("Src IP", "src"), ("Src Comment", "src"),
        # 목적지 열 (11~15) / Destination columns (11-15)
        ("Dst Group OBJ", "dst"), ("Dst OBJ Name", "dst"),
        ("Dst Type", "dst"), ("Dst IP", "dst"), ("Dst Comment", "dst"),
        # 서비스 열 (16~19) / Service columns (16-19)
        ("Svc Group OBJ", "svc"), ("Svc OBJ Name", "svc"),
        ("Svc Port", "svc"), ("Svc Comment", "svc"),
        # 코멘트 열 (20) / Comment column (20)
        ("Comment", "base")
    ]
    write_styled_header(ws, 1, headers_with_cat)
    if not policies:
        return ws

    row = 2
    seq = 0
    for p in policies:
        seq += 1
        is_dis = p.get('status', 'enable') == 'disable'

        b_fill, b_font = get_cell_style('base', seq, is_dis)
        s_fill, s_font = get_cell_style('src', seq, is_dis)
        sg_fill, sg_font = get_cell_style('src', seq, is_dis, is_group=True)
        d_fill, d_font = get_cell_style('dst', seq, is_dis)
        dg_fill, dg_font = get_cell_style('dst', seq, is_dis, is_group=True)
        v_fill, v_font = get_cell_style('svc', seq, is_dis)
        vg_fill, vg_font = get_cell_style('svc', seq, is_dis, is_group=True)

        src_expanded = []
        for a in (p.get('srcaddr') or []):
            src_expanded.extend(resolver.resolve_address(a))
        dst_expanded = []
        for a in (p.get('dstaddr') or []):
            dst_expanded.extend(resolver.resolve_address(a))
        svc_expanded = []
        for s in (p.get('service') or []):
            svc_expanded.extend(resolver.resolve_service(s))

        max_rows = max(len(src_expanded), len(dst_expanded), len(svc_expanded), 1)
        p_start = row
        p_end = row + max_rows - 1

        for ri in range(max_rows):
            cur_r = row + ri
            if ri == 0:
                sc(ws, cur_r, 1, seq, font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 2, vdom_name, font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 3, 'N' if is_dis else 'Y', font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 4, p.get('id', ''), font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 5, p.get('interface', ''), font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 20, p.get('comments', ''), font=b_font, fill=b_fill)
            else:
                for c in [1, 2, 3, 4, 5, 20]:
                    sc(ws, cur_r, c, None, font=b_font, fill=b_fill)

            # 출발지 / Src (Source)
            if ri < len(src_expanded):
                grp, n, t, ip, comm = src_expanded[ri]
                sc(ws, cur_r, 6, grp or '', font=sg_font if grp else s_font, fill=s_fill)
                sc(ws, cur_r, 7, n, font=s_font, fill=s_fill)
                sc(ws, cur_r, 8, t, font=s_font, fill=s_fill, align=CENTER)
                sc(ws, cur_r, 9, ip, font=s_font, fill=s_fill)
                sc(ws, cur_r, 10, comm, font=s_font, fill=s_fill)
            else:
                for c in range(6, 11):
                    sc(ws, cur_r, c, None, font=s_font, fill=s_fill)

            # 목적지 / Dst (Destination)
            if ri < len(dst_expanded):
                grp, n, t, ip, comm = dst_expanded[ri]
                sc(ws, cur_r, 11, grp or '', font=dg_font if grp else d_font, fill=d_fill)
                sc(ws, cur_r, 12, n, font=d_font, fill=d_fill)
                sc(ws, cur_r, 13, t, font=d_font, fill=d_fill, align=CENTER)
                sc(ws, cur_r, 14, ip, font=d_font, fill=d_fill)
                sc(ws, cur_r, 15, comm, font=d_font, fill=d_fill)
            else:
                for c in range(11, 16):
                    sc(ws, cur_r, c, None, font=d_font, fill=d_fill)

            # 서비스 / Service
            if ri < len(svc_expanded):
                sgrp, sn, sport, scomm = svc_expanded[ri]
                sc(ws, cur_r, 16, sgrp or '', font=vg_font if sgrp else v_font, fill=v_fill)
                sc(ws, cur_r, 17, sn, font=v_font, fill=v_fill)
                sc(ws, cur_r, 18, sport, font=v_font, fill=v_fill)
                sc(ws, cur_r, 19, scomm, font=v_font, fill=v_fill)
            else:
                for c in range(16, 20):
                    sc(ws, cur_r, c, None, font=v_font, fill=v_fill)

        if p_end > p_start:
            for c in [1, 2, 3, 4, 5]:
                merge_row_range(ws, p_start, p_end, c, h_align='center')
            merge_row_range(ws, p_start, p_end, 20, h_align='left')
            merge_group_spans(ws, p_start, src_expanded, 6, h_align='left')
            merge_group_spans(ws, p_start, dst_expanded, 11, h_align='left')
            merge_group_spans(ws, p_start, svc_expanded, 16, h_align='left')

        row += max_rows

    auto_fit(ws)
    return ws


def write_external_resource_sheet(ws, resources, vdom_name=""):
    """
    'External Resource' 시트 작성 / Sheet Builder - External Resource
    """
    headers_with_cat = [
        ("Seq", "base"), ("vDOM", "base"), ("Enable", "base"), ("Name", "base"),
        ("Type", "base"), ("Resource URL", "base"),
        ("Refresh Rate (min)", "base"), ("Source IP", "base"),
        ("Comment", "base")
    ]
    write_styled_header(ws, 1, headers_with_cat)
    ncol = len(headers_with_cat)

    row = 2
    res_list = list(resources.values()) if isinstance(resources, dict) else (resources or [])
    for seq, r in enumerate(res_list, 1):
        is_dis = (r.get('status') == 'disable')
        b_fill, b_font = get_cell_style('base', seq, is_dis)
        s_fill, s_font = get_cell_style('src', seq, is_dis)

        enable_val = 'N' if is_dis else 'Y'

        sc(ws, row, 1, seq, font=b_font, fill=b_fill, align=CENTER)
        sc(ws, row, 2, vdom_name, font=b_font, fill=b_fill, align=CENTER)
        sc(ws, row, 3, enable_val, font=b_font, fill=b_fill, align=CENTER)
        sc(ws, row, 4, r.get('name', ''), font=b_font, fill=b_fill)
        sc(ws, row, 5, r.get('type', ''), font=b_font, fill=b_fill, align=CENTER)
        sc(ws, row, 6, r.get('resource', ''), font=s_font, fill=s_fill)
        sc(ws, row, 7, r.get('refresh-rate', ''), font=b_font, fill=b_fill, align=CENTER)
        sc(ws, row, 8, r.get('source-ip', ''), font=b_font, fill=b_fill, align=CENTER)
        sc(ws, row, 9, r.get('comments', ''), font=b_font, fill=b_fill)
        row += 1

    auto_fit(ws)
    return ws


# ================================================================
#  12. 시트 작성 - Static Route / Sheet Builder - Static Route
# ================================================================

def write_static_route_sheet(ws, static_routes, vdom_name=""):
    """
    'Static Route' 시트 작성 / Sheet Builder - Static Route
    """
    headers_with_cat = [
        ("Seq", "base"), ("vDOM", "base"), ("Enable", "base"), ("ID", "base"),
        ("Destination", "dst"), ("Gateway", "src"), ("Interface", "base"),
        ("Distance", "base"), ("Priority", "base"), ("Options", "base"),
        ("Comment", "base")
    ]
    write_styled_header(ws, 1, headers_with_cat)
    if not static_routes:
        return ws

    row = 2
    for seq, r in enumerate(static_routes, 1):
        is_dis = (r.get('status') == 'disable')
        b_fill, b_font = get_cell_style('base', seq, is_dis)
        s_fill, s_font = get_cell_style('src', seq, is_dis)
        d_fill, d_font = get_cell_style('dst', seq, is_dis)

        enable_val = 'N' if is_dis else 'Y'

        opts = []
        if r.get('blackhole') == 'enable':
            opts.append('Blackhole')
        if r.get('dynamic-gateway') == 'enable':
            opts.append('Dynamic-GW')
        if r.get('bfd') == 'enable':
            opts.append('BFD')
        if r.get('link-monitor-exempt') == 'enable':
            opts.append('Link-Mon-Exempt')
        options_str = ', '.join(opts)

        sc(ws, row, 1, seq, font=b_font, fill=b_fill, align=CENTER)
        sc(ws, row, 2, vdom_name, font=b_font, fill=b_fill, align=CENTER)
        sc(ws, row, 3, enable_val, font=b_font, fill=b_fill, align=CENTER)
        sc(ws, row, 4, r.get('id', ''), font=b_font, fill=b_fill, align=CENTER)
        sc(ws, row, 5, r.get('dst', '0.0.0.0/0'), font=d_font, fill=d_fill)
        sc(ws, row, 6, r.get('gateway', ''), font=s_font, fill=s_fill, align=CENTER)
        sc(ws, row, 7, r.get('device', ''), font=b_font, fill=b_fill, align=CENTER)
        sc(ws, row, 8, r.get('distance', ''), font=b_font, fill=b_fill, align=CENTER)
        sc(ws, row, 9, r.get('priority', ''), font=b_font, fill=b_fill, align=CENTER)
        sc(ws, row, 10, options_str, font=b_font, fill=b_fill, align=CENTER)
        sc(ws, row, 11, r.get('comment', ''), font=b_font, fill=b_fill)
        row += 1

    auto_fit(ws)
    return ws


# ================================================================
#  13. 시트 작성 - Policy Route / Sheet Builder - Policy Route
# ================================================================

def write_policy_route_sheet(ws, policy_routes, resolver, vdom_name=""):
    """
    'Policy Route' 시트 작성 / Sheet Builder - Policy Route
    """
    headers_with_cat = [
        ("Seq", "base"), ("vDOM", "base"), ("Enable", "base"), ("ID", "base"),
        ("Incoming Intf", "src"), ("Outgoing Intf", "dst"), ("Gateway", "base"),
        ("Action", "base"), ("Protocol", "base"), ("Port Range", "base"),
        # 출발지 열 (11~15)
        ("Src Group OBJ", "src"), ("Src OBJ Name", "src"),
        ("Src Type", "src"), ("Src IP", "src"), ("Src Comment", "src"),
        # 목적지 열 (16~20)
        ("Dst Group OBJ", "dst"), ("Dst OBJ Name", "dst"),
        ("Dst Type", "dst"), ("Dst IP", "dst"), ("Dst Comment", "dst"),
        # 코멘트 (21)
        ("Comment", "base")
    ]
    write_styled_header(ws, 1, headers_with_cat)
    if not policy_routes:
        return ws

    proto_names = {
        '0': 'ALL', '1': 'ICMP(1)', '6': 'TCP(6)', '17': 'UDP(17)',
        '47': 'GRE(47)', '50': 'ESP(50)', '51': 'AH(51)', '89': 'OSPF(89)'
    }

    row = 2
    for seq, p in enumerate(policy_routes, 1):
        is_dis = (p.get('status') == 'disable')
        action_val = p.get('action', 'permit')
        b_fill, b_font = get_cell_style('base', seq, is_dis)
        s_fill, s_font = get_cell_style('src', seq, is_dis)
        sg_fill, sg_font = get_cell_style('src', seq, is_dis, is_group=True)
        d_fill, d_font = get_cell_style('dst', seq, is_dis)
        dg_fill, dg_font = get_cell_style('dst', seq, is_dis, is_group=True)
        _, act_font = get_cell_style('base', seq, is_dis, is_action=True, action_val='accept' if action_val == 'permit' else 'deny')

        # 출발지 전개
        src_expanded = []
        for a in (p.get('srcaddr') or []):
            src_expanded.extend(resolver.resolve_address(a))
        if not src_expanded and p.get('src'):
            src_expanded = [('', '(subnet)', 'ipmask', p.get('src'), '')]

        # 목적지 전개
        dst_expanded = []
        for a in (p.get('dstaddr') or []):
            dst_expanded.extend(resolver.resolve_address(a))
        if not dst_expanded and p.get('dst'):
            dst_expanded = [('', '(subnet)', 'ipmask', p.get('dst'), '')]

        max_rows = max(len(src_expanded), len(dst_expanded), 1)
        p_start = row
        p_end = row + max_rows - 1

        proto_val = str(p.get('protocol', '0'))
        proto_str = proto_names.get(proto_val, f"Proto({proto_val})")
        start_p = p.get('start-port', '')
        end_p = p.get('end-port', '')
        if start_p and end_p:
            port_str = f"{start_p}" if start_p == end_p else f"{start_p}-{end_p}"
        elif start_p:
            port_str = f"{start_p}"
        else:
            port_str = "ALL"

        in_intf = '\n'.join(p.get('input-device') or [])
        out_intf = '\n'.join(p.get('output-device') or [])

        for ri in range(max_rows):
            cur_r = row + ri
            if ri == 0:
                sc(ws, cur_r, 1, seq, font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 2, vdom_name, font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 3, 'N' if is_dis else 'Y', font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 4, p.get('id', ''), font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 5, in_intf, font=s_font, fill=s_fill)
                sc(ws, cur_r, 6, out_intf, font=d_font, fill=d_fill)
                sc(ws, cur_r, 7, p.get('gateway', ''), font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 8, action_val, font=act_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 9, proto_str, font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 10, port_str, font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 21, p.get('comments', ''), font=b_font, fill=b_fill)
            else:
                for c in [1, 2, 3, 4, 7, 8, 9, 10, 21]:
                    sc(ws, cur_r, c, None, font=b_font, fill=b_fill)
                sc(ws, cur_r, 5, None, font=s_font, fill=s_fill)
                sc(ws, cur_r, 6, None, font=d_font, fill=d_fill)

            # 출발지
            if ri < len(src_expanded):
                grp, n, t, ip, comm = src_expanded[ri]
                sc(ws, cur_r, 11, grp or '', font=sg_font if grp else s_font, fill=s_fill)
                sc(ws, cur_r, 12, n, font=s_font, fill=s_fill)
                sc(ws, cur_r, 13, t, font=s_font, fill=s_fill, align=CENTER)
                sc(ws, cur_r, 14, ip, font=s_font, fill=s_fill)
                sc(ws, cur_r, 15, comm, font=s_font, fill=s_fill)
            else:
                for c in range(11, 16):
                    sc(ws, cur_r, c, None, font=s_font, fill=s_fill)

            # 목적지
            if ri < len(dst_expanded):
                grp, n, t, ip, comm = dst_expanded[ri]
                sc(ws, cur_r, 16, grp or '', font=dg_font if grp else d_font, fill=d_fill)
                sc(ws, cur_r, 17, n, font=d_font, fill=d_fill)
                sc(ws, cur_r, 18, t, font=d_font, fill=d_fill, align=CENTER)
                sc(ws, cur_r, 19, ip, font=d_font, fill=d_fill)
                sc(ws, cur_r, 20, comm, font=d_font, fill=d_fill)
            else:
                for c in range(16, 21):
                    sc(ws, cur_r, c, None, font=d_font, fill=d_fill)

        # 병합
        if max_rows > 1:
            for c in [1, 2, 3, 4, 7, 8, 9, 10]:
                merge_row_range(ws, p_start, p_end, c, h_align='center')
            for c in [5, 6, 21]:
                merge_row_range(ws, p_start, p_end, c, h_align='left')
            merge_group_spans(ws, p_start, src_expanded, 11, h_align='left')
            merge_group_spans(ws, p_start, dst_expanded, 16, h_align='left')

        row += max_rows

    auto_fit(ws)
    return ws


# ================================================================
#  14. 시트 작성 - OSPF / Sheet Builder - OSPF
# ================================================================

def write_ospf_sheet(ws, ospf_data, vdom_name=""):
    """
    'OSPF' 시트 작성 / Sheet Builder - OSPF
    """
    has_ospf = bool(ospf_data and (
        ospf_data.get('router-id') or
        ospf_data.get('networks') or
        ospf_data.get('interfaces') or
        any(r.get('status') == 'enable' for r in ospf_data.get('redistribute', []))
    ))

    if not has_ospf:
        # 비어있는 경우 표준 헤더만 1행에 배치하여 자동 빨간색 탭 마킹 지원
        headers_with_cat = [
            ("Seq", "base"), ("vDOM", "base"), ("Router ID", "base"),
            ("Area", "base"), ("Prefix", "base"), ("Interface", "base"), ("Cost", "base")
        ]
        write_styled_header(ws, 1, headers_with_cat)
        return ws

    SEC_TITLE_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    SEC_TITLE_FONT = Font(name="맑은 고딕", size=11, bold=True, color="FFFFFF")

    def write_sec_title(ws, r, title, end_col):
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=end_col)
        for c in range(1, end_col + 1):
            sc(ws, r, c, title if c == 1 else None, font=SEC_TITLE_FONT, fill=SEC_TITLE_FILL)

    cur_r = 1
    # 1. Global Info (5열 구성)
    write_sec_title(ws, cur_r, " 1. OSPF Global Settings", 5)
    cur_r += 1

    sc(ws, cur_r, 1, "vDOM", font=HDR_FONT, fill=HDR_DEFAULT_FILL, align=CENTER)
    sc(ws, cur_r, 2, "Router ID", font=HDR_FONT, fill=HDR_DEFAULT_FILL, align=CENTER)
    sc(ws, cur_r, 3, "Total Networks", font=HDR_FONT, fill=HDR_DEFAULT_FILL, align=CENTER)
    sc(ws, cur_r, 4, "Total Interfaces", font=HDR_FONT, fill=HDR_DEFAULT_FILL, align=CENTER)
    sc(ws, cur_r, 5, "Total Areas", font=HDR_FONT, fill=HDR_DEFAULT_FILL, align=CENTER)
    cur_r += 1

    sc(ws, cur_r, 1, vdom_name, font=FONT_DEFAULT, fill=ODD_ROW_FILL, align=CENTER)
    sc(ws, cur_r, 2, ospf_data.get('router-id') or '(Not Configured)', font=FONT_DEFAULT_BOLD, fill=ODD_ROW_FILL, align=CENTER)
    sc(ws, cur_r, 3, len(ospf_data.get('networks', [])), font=FONT_DEFAULT, fill=ODD_ROW_FILL, align=CENTER)
    sc(ws, cur_r, 4, len(ospf_data.get('interfaces', [])), font=FONT_DEFAULT, fill=ODD_ROW_FILL, align=CENTER)
    sc(ws, cur_r, 5, len(ospf_data.get('areas', [])), font=FONT_DEFAULT, fill=ODD_ROW_FILL, align=CENTER)
    cur_r += 2

    # 2. Networks (4열 구성)
    write_sec_title(ws, cur_r, " 2. OSPF Networks (config network)", 4)
    cur_r += 1

    net_headers = [("Seq", "base"), ("ID", "base"), ("Prefix", "dst"), ("Area", "src")]
    for col, (h, cat) in enumerate(net_headers, 1):
        sc(ws, cur_r, col, h, font=HDR_FONT, fill=HDR_DST_FILL if cat == 'dst' else (HDR_SRC_FILL if cat == 'src' else HDR_DEFAULT_FILL), align=CENTER)
    cur_r += 1

    networks = ospf_data.get('networks', [])
    if networks:
        for seq, net in enumerate(networks, 1):
            fill = EVEN_ROW_FILL if seq % 2 == 0 else ODD_ROW_FILL
            sc(ws, cur_r, 1, seq, font=FONT_DEFAULT, fill=fill, align=CENTER)
            sc(ws, cur_r, 2, net.get('id', ''), font=FONT_DEFAULT, fill=fill, align=CENTER)
            sc(ws, cur_r, 3, net.get('prefix', ''), font=FONT_DST, fill=fill)
            sc(ws, cur_r, 4, net.get('area', ''), font=FONT_SRC, fill=fill, align=CENTER)
            cur_r += 1
    else:
        sc(ws, cur_r, 1, "(No networks configured)", font=FONT_DEFAULT, fill=ODD_ROW_FILL, align=CENTER)
        cur_r += 1
    cur_r += 1

    # 3. Interfaces (8열 구성 - Priority 기본값 1, Authentication 기본값 none)
    write_sec_title(ws, cur_r, " 3. OSPF Interfaces (config ospf-interface)", 8)
    cur_r += 1

    intf_headers = ["Seq", "Name", "Interface", "Cost", "Dead / Hello Interval", "Network Type", "Priority", "Authentication"]
    for col, h in enumerate(intf_headers, 1):
        sc(ws, cur_r, col, h, font=HDR_FONT, fill=HDR_DEFAULT_FILL, align=CENTER)
    cur_r += 1

    interfaces = ospf_data.get('interfaces', [])
    if interfaces:
        for seq, inf in enumerate(interfaces, 1):
            fill = EVEN_ROW_FILL if seq % 2 == 0 else ODD_ROW_FILL
            interval_str = f"Dead: {inf.get('dead-interval', '-')}, Hello: {inf.get('hello-interval', '-')}"
            sc(ws, cur_r, 1, seq, font=FONT_DEFAULT, fill=fill, align=CENTER)
            sc(ws, cur_r, 2, inf.get('name', ''), font=FONT_DEFAULT_BOLD, fill=fill)
            sc(ws, cur_r, 3, inf.get('interface', ''), font=FONT_SRC, fill=fill, align=CENTER)
            sc(ws, cur_r, 4, inf.get('cost', ''), font=FONT_DEFAULT, fill=fill, align=CENTER)
            sc(ws, cur_r, 5, interval_str, font=FONT_DEFAULT, fill=fill, align=CENTER)
            sc(ws, cur_r, 6, inf.get('network-type', ''), font=FONT_DEFAULT, fill=fill, align=CENTER)
            sc(ws, cur_r, 7, inf.get('priority') or '1', font=FONT_DEFAULT, fill=fill, align=CENTER)
            sc(ws, cur_r, 8, inf.get('authentication') or 'none', font=FONT_DEFAULT, fill=fill, align=CENTER)
            cur_r += 1
    else:
        sc(ws, cur_r, 1, "(No ospf-interfaces configured)", font=FONT_DEFAULT, fill=ODD_ROW_FILL, align=CENTER)
        cur_r += 1
    cur_r += 1

    # 4. Redistribution (6열 구성)
    write_sec_title(ws, cur_r, " 4. OSPF Redistribution (config redistribute)", 6)
    cur_r += 1

    redist_headers = ["Seq", "Protocol", "Status", "Route-Map", "Metric", "Metric Type"]
    for col, h in enumerate(redist_headers, 1):
        sc(ws, cur_r, col, h, font=HDR_FONT, fill=HDR_DEFAULT_FILL, align=CENTER)
    cur_r += 1

    redists = [r for r in ospf_data.get('redistribute', []) if r.get('status') == 'enable' or r.get('routemap')]
    if not redists:
        redists = ospf_data.get('redistribute', [])

    if redists:
        for seq, rd in enumerate(redists, 1):
            is_dis = (rd.get('status') != 'enable')
            b_fill, b_font = get_cell_style('base', seq, is_dis)
            sc(ws, cur_r, 1, seq, font=b_font, fill=b_fill, align=CENTER)
            sc(ws, cur_r, 2, rd.get('protocol', '').upper(), font=b_font, fill=b_fill, align=CENTER)
            sc(ws, cur_r, 3, rd.get('status', 'disable'), font=b_font, fill=b_fill, align=CENTER)
            sc(ws, cur_r, 4, rd.get('routemap', ''), font=b_font, fill=b_fill)
            sc(ws, cur_r, 5, rd.get('metric', ''), font=b_font, fill=b_fill, align=CENTER)
            sc(ws, cur_r, 6, rd.get('metric-type', ''), font=b_font, fill=b_fill, align=CENTER)
            cur_r += 1
    else:
        sc(ws, cur_r, 1, "(No redistribution configured)", font=FONT_DEFAULT, fill=ODD_ROW_FILL, align=CENTER)
        cur_r += 1
    cur_r += 1

    # 5. Route-Map & Filter Details (9열 구성 - match-ip 서브넷별 행 분리, Action, Exact Match 분리, Set Actions 열 삭제)
    write_sec_title(ws, cur_r, " 5. Route-Map & Filter Details (config router route-map / access-list)", 9)
    cur_r += 1

    rm_headers = ["Seq", "Route-Map", "Rule", "Route-Map Action", "Match Target", "Filtered Prefix", "Action", "Exact Match", "ACL Comment"]
    for col, h in enumerate(rm_headers, 1):
        sc(ws, cur_r, col, h, font=HDR_FONT, fill=HDR_DEFAULT_FILL, align=CENTER)
    cur_r += 1

    rmaps = ospf_data.get('route_maps', {})
    acls = ospf_data.get('access_lists', {})

    rm_entries = []
    for rm_name, rm in rmaps.items():
        for rule in rm.get('rules', []):
            rm_entries.append((rm_name, rule))

    if rm_entries:
        for seq, (rm_name, rule) in enumerate(rm_entries, 1):
            fill = EVEN_ROW_FILL if seq % 2 == 0 else ODD_ROW_FILL
            match_ip = rule.get('match_ip', '')
            match_target = f"match-ip: {match_ip}" if match_ip else "(match all)"

            acl = acls.get(match_ip, {}) if match_ip else {}
            acl_comm = acl.get('comments', '')
            acl_rules = acl.get('rules', [])

            max_rows = max(len(acl_rules), 1)
            p_start = cur_r
            p_end = cur_r + max_rows - 1

            act_val = rule.get('action', 'permit').upper()
            act_font = DENY_FONT if act_val == 'DENY' else ACCEPT_FONT

            for ri in range(max_rows):
                r_now = cur_r + ri
                if ri == 0:
                    sc(ws, r_now, 1, seq, font=FONT_DEFAULT, fill=fill, align=CENTER)
                    sc(ws, r_now, 2, rm_name, font=FONT_DEFAULT_BOLD, fill=fill)
                    sc(ws, r_now, 3, rule.get('id', ''), font=FONT_DEFAULT, fill=fill, align=CENTER)
                    sc(ws, r_now, 4, act_val, font=act_font, fill=fill, align=CENTER)
                    sc(ws, r_now, 5, match_target, font=FONT_DEFAULT, fill=fill)
                    sc(ws, r_now, 9, acl_comm, font=FONT_DEFAULT, fill=fill)
                else:
                    for c in [1, 3, 4]:
                        sc(ws, r_now, c, None, font=FONT_DEFAULT, fill=fill, align=CENTER)
                    for c in [2, 5, 9]:
                        sc(ws, r_now, c, None, font=FONT_DEFAULT, fill=fill)

                if ri < len(acl_rules):
                    ar = acl_rules[ri]
                    p_val = ar.get('prefix', '')
                    a_val = ar.get('action', 'permit')
                    em_val = ar.get('exact_match', 'disable')
                    a_font = DENY_FONT if a_val.upper() == 'DENY' else ACCEPT_FONT
                    sc(ws, r_now, 6, p_val, font=FONT_SRC, fill=fill)
                    sc(ws, r_now, 7, a_val, font=a_font, fill=fill, align=CENTER)
                    sc(ws, r_now, 8, em_val, font=FONT_DEFAULT, fill=fill, align=CENTER)
                else:
                    pref_fallback = "-" if match_ip else "(all)"
                    sc(ws, r_now, 6, pref_fallback, font=FONT_DEFAULT, fill=fill)
                    sc(ws, r_now, 7, "-", font=FONT_DEFAULT, fill=fill, align=CENTER)
                    sc(ws, r_now, 8, "-", font=FONT_DEFAULT, fill=fill, align=CENTER)

            if p_end > p_start:
                merge_row_range(ws, p_start, p_end, 1, h_align='center')
                merge_row_range(ws, p_start, p_end, 2, h_align='left')
                merge_row_range(ws, p_start, p_end, 3, h_align='center')
                merge_row_range(ws, p_start, p_end, 4, h_align='center')
                merge_row_range(ws, p_start, p_end, 5, h_align='left')
                merge_row_range(ws, p_start, p_end, 9, h_align='left')

            cur_r += max_rows
    else:
        sc(ws, cur_r, 1, "(No route-maps configured)", font=FONT_DEFAULT, fill=ODD_ROW_FILL, align=CENTER)
        cur_r += 1

    auto_fit(ws)
    return ws


# ================================================================
#  15. 시트 작성 - IPsec VPN / Sheet Builder - IPsec VPN
# ================================================================

def write_ipsec_vpn_sheet(ws, vpn_data, resolver, vdom_name=""):
    """
    'IPsec VPN' 시트 작성 / Sheet Builder - IPsec VPN
    P1 / P2 헤더 배색 구분, Network & Advanced 설정 완벽 지원, P1/P2 DH Group 기본값(14 5), 서브넷 기본값(0.0.0.0/0)
    헤더 가독성 향상(2줄 줄바꿈), DPD Retry Interval 숫자 전용 표기
    """
    headers_with_cat = [
        ("Seq", "base"), ("vDOM", "base"),
        # Phase 1 및 Network/Advanced 컬럼 (3~22) - 딥 네이비 헤더
        ("P1 Name", "p1"), ("Interface", "p1"), ("Remote\nGateway", "p1"), ("Local\nGateway", "p1"),
        ("IKE\nVersion", "p1"), ("P1 Proposal", "p1"), ("P1 DH\nGroup", "p1"),
        ("NAT\nTraversal", "p1"), ("Keepalive\nFrequency", "p1"),
        ("Dead Peer\nDetection (DPD)", "p1"), ("DPD Retry\nCount", "p1"), ("DPD Retry\nInterval", "p1"),
        ("FEC\nEgress", "p1"), ("FEC\nIngress", "p1"),
        ("Add\nRoute", "p1"),
        ("Auto Discovery\nSender", "p1"), ("Auto Discovery\nReceiver", "p1"),
        ("Exchange\nInterface IP", "p1"), ("Device\nCreation", "p1"),
        ("P1 Comment", "p1"),
        # Phase 2 컬럼 (23~30) - 딥 틸 헤더
        ("P2 Name", "p2"), ("P2 Proposal", "p2"), ("P2 DH\nGroup", "p2"),
        ("Local Subnet\n/ Src", "p2"), ("Remote Subnet\n/ Dst", "p2"),
        ("Auto\nNegotiate", "p2"), ("Keepalive", "p2"), ("P2 Comment", "p2")
    ]
    write_styled_header(ws, 1, headers_with_cat)
    if not vpn_data:
        return ws

    row = 2
    for seq, p1 in enumerate(vpn_data, 1):
        b_fill, b_font = get_cell_style('base', seq, False)
        s_fill, s_font = get_cell_style('src', seq, False)
        d_fill, d_font = get_cell_style('dst', seq, False)

        p2_list = p1.get('phase2_list', [])
        max_rows = max(len(p2_list), 1)
        p_start = row
        p_end = row + max_rows - 1

        p1_dhgrp = p1.get('dhgrp') or '14 5'
        p1_dpd = p1.get('dpd') or 'on-demand'
        p1_retry_cnt = p1.get('dpd-retrycount') or '3'
        p1_retry_int = str(p1.get('dpd-retryinterval') or '20')
        p1_natt = p1.get('nattraversal') or 'enable'
        p1_keepalive = p1.get('keepalive') or '10'
        p1_fec_egress = p1.get('fec-egress') or 'disable'
        p1_fec_ingress = p1.get('fec-ingress') or 'disable'
        p1_add_route = p1.get('add-gw-route') or 'enable'
        p1_ad_sender = p1.get('auto-discovery-sender') or 'disable'
        p1_ad_receiver = p1.get('auto-discovery-receiver') or 'disable'
        p1_ex_intf_ip = p1.get('exchange-interface-ip') or 'disable'
        p1_dev_creation = p1.get('net-device') or 'disable'

        for ri in range(max_rows):
            cur_r = row + ri
            if ri == 0:
                sc(ws, cur_r, 1, seq, font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 2, vdom_name, font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 3, p1.get('name', ''), font=FONT_DEFAULT_BOLD, fill=b_fill)
                sc(ws, cur_r, 4, p1.get('interface', ''), font=s_font, fill=s_fill, align=CENTER)
                sc(ws, cur_r, 5, p1.get('remote-gw', '') or '(Dialup/Dynamic)', font=d_font, fill=d_fill, align=CENTER)
                sc(ws, cur_r, 6, p1.get('local-gw', '') or '-', font=s_font, fill=s_fill, align=CENTER)
                sc(ws, cur_r, 7, f"v{p1.get('ike-version', '1')}", font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 8, p1.get('proposal', ''), font=b_font, fill=b_fill)
                sc(ws, cur_r, 9, p1_dhgrp, font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 10, p1_natt, font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 11, p1_keepalive, font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 12, p1_dpd, font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 13, p1_retry_cnt, font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 14, p1_retry_int, font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 15, p1_fec_egress, font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 16, p1_fec_ingress, font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 17, p1_add_route, font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 18, p1_ad_sender, font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 19, p1_ad_receiver, font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 20, p1_ex_intf_ip, font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 21, p1_dev_creation, font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 22, p1.get('comments', ''), font=b_font, fill=b_fill)
            else:
                for c in [1, 2, 7, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21]:
                    sc(ws, cur_r, c, None, font=b_font, fill=b_fill, align=CENTER)
                for c in [4, 6]:
                    sc(ws, cur_r, c, None, font=s_font, fill=s_fill, align=CENTER)
                sc(ws, cur_r, 5, None, font=d_font, fill=d_fill, align=CENTER)
                for c in [3, 8, 22]:
                    sc(ws, cur_r, c, None, font=b_font, fill=b_fill)

            # Phase 2
            if ri < len(p2_list):
                p2 = p2_list[ri]
                dhgrp_val = p2.get('dhgrp') or '14 5'
                src_val = p2.get('src-subnet') or p2.get('src-name') or '0.0.0.0/0'
                dst_val = p2.get('dst-subnet') or p2.get('dst-name') or '0.0.0.0/0'

                sc(ws, cur_r, 23, p2.get('name', ''), font=FONT_DEFAULT_BOLD, fill=b_fill)
                sc(ws, cur_r, 24, p2.get('proposal', ''), font=b_font, fill=b_fill)
                sc(ws, cur_r, 25, dhgrp_val, font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 26, src_val, font=s_font, fill=s_fill)
                sc(ws, cur_r, 27, dst_val, font=d_font, fill=d_fill)
                sc(ws, cur_r, 28, p2.get('auto-negotiate', 'disable'), font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 29, p2.get('keepalive', 'disable'), font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 30, p2.get('comments', ''), font=b_font, fill=b_fill)
            else:
                for c in [23, 24, 30]:
                    sc(ws, cur_r, c, None, font=b_font, fill=b_fill)
                for c in [25, 28, 29]:
                    sc(ws, cur_r, c, None, font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 26, None, font=s_font, fill=s_fill)
                sc(ws, cur_r, 27, None, font=d_font, fill=d_fill)

        if max_rows > 1:
            # 가운데 정렬할 Phase 1 열: Seq(1), vDOM(2), Interface(4), Remote GW(5), Local GW(6), IKE Ver(7), DH Group(9), NAT-T(10), Keepalive(11), DPD(12), Retry Cnt(13), Retry Int(14), FEC-E(15), FEC-I(16), Add-Route(17), AD-S(18), AD-R(19), Ex-Intf-IP(20), Dev-Create(21)
            for c in [1, 2, 4, 5, 6, 7, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21]:
                merge_row_range(ws, p_start, p_end, c, h_align='center')
            # 왼쪽 정렬할 Phase 1 열: P1 Name(3), P1 Proposal(8), P1 Comment(22)
            for c in [3, 8, 22]:
                merge_row_range(ws, p_start, p_end, c, h_align='left')

        row += max_rows

    auto_fit(ws)
    return ws


# ================================================================
#  16. 시트 작성 - Network Interface / Sheet Builder - Network Interface
# ================================================================

def format_interface_speed(val):
    """
    인터페이스 Speed/Duplex 포맷팅 (예: 1000auto -> 1000 / auto, 10000full -> 10000 / full)
    """
    if not val:
        return ''
    s = str(val).strip()
    if '/' in s or s.lower() == 'auto':
        return s
    m = re.match(r'^(\d+[A-Za-z]?)(auto|full|half)$', s, re.IGNORECASE)
    if m:
        return f"{m.group(1)} / {m.group(2)}"
    return s


def write_interface_sheet(ws, interfaces, vdom_name=""):
    """
    'Network Interface' 시트 작성 / Sheet Builder - Network Interface
    """
    headers_with_cat = [
        ("Seq", "base"), ("vDOM", "base"), ("Status", "base"), ("Name", "base"),
        ("Alias", "base"), ("Type", "base"), ("Primary IP", "src"),
        ("Secondary IP", "src"), ("Remote IP (Tunnel)", "src"),
        ("VLAN ID", "base"), ("Parent / Member Interface", "base"),
        ("VRF", "base"), ("Addressing Mode", "base"), ("Administrative Access", "base"),
        ("Speed / Duplex", "base"), ("Description", "base")
    ]
    write_styled_header(ws, 1, headers_with_cat)
    if not interfaces:
        return ws

    row = 2
    for seq, intf in enumerate(interfaces, 1):
        is_dis = (intf.get('status') == 'down')
        b_fill, b_font = get_cell_style('base', seq, is_dis)
        s_fill, s_font = get_cell_style('src', seq, is_dis)

        status_str = 'DOWN' if is_dis else 'UP'
        pip_val = intf.get('display') or '0.0.0.0/0'
        sec_ips_str = '\n'.join(intf.get('secondary_ips', [])) if intf.get('secondary_ips') else None
        remote_ip_str = intf.get('remote-ip') or None
        if intf.get('member'):
            parent_member_str = ', '.join(intf.get('member', []))
        elif intf.get('interface'):
            parent_member_str = intf.get('interface')
        else:
            parent_member_str = None
        vrf_val = intf.get('vrf')
        vrf_str = str(vrf_val) if vrf_val is not None and str(vrf_val).strip() != '' else '0'

        sc(ws, row, 1, seq, font=b_font, fill=b_fill, align=CENTER)
        sc(ws, row, 2, vdom_name, font=b_font, fill=b_fill, align=CENTER)
        sc(ws, row, 3, status_str, font=DENY_FONT if is_dis else ACCEPT_FONT, fill=b_fill, align=CENTER)
        sc(ws, row, 4, intf.get('name', ''), font=FONT_DEFAULT_BOLD, fill=b_fill)
        sc(ws, row, 5, intf.get('alias', ''), font=b_font, fill=b_fill)
        sc(ws, row, 6, intf.get('type', ''), font=b_font, fill=b_fill, align=CENTER)
        sc(ws, row, 7, pip_val, font=s_font, fill=s_fill, align=CENTER)
        sc(ws, row, 8, sec_ips_str, font=s_font, fill=s_fill, align=CENTER)
        sc(ws, row, 9, remote_ip_str, font=s_font, fill=s_fill, align=CENTER)
        sc(ws, row, 10, intf.get('vlanid', '') or None, font=b_font, fill=b_fill, align=CENTER)
        sc(ws, row, 11, parent_member_str, font=b_font, fill=b_fill, align=CENTER if parent_member_str else None)
        sc(ws, row, 12, vrf_str, font=b_font, fill=b_fill, align=CENTER)
        sc(ws, row, 13, intf.get('mode', '') or None, font=b_font, fill=b_fill, align=CENTER)
        sc(ws, row, 14, intf.get('allowaccess', '') or None, font=b_font, fill=b_fill)
        sc(ws, row, 15, format_interface_speed(intf.get('speed', '')) or None, font=b_font, fill=b_fill, align=CENTER)
        sc(ws, row, 16, intf.get('description', '') or None, font=b_font, fill=b_fill)
        row += 1

    auto_fit(ws)
    return ws


# ================================================================
# 17. 개별 vDOM 엑셀 생성 함수 / Per-vDOM Excel Generation Function
# ================================================================

def export_single_vdom_excel(vdom_name, fw_pols, li_pols, cn_ents, dn_ents, dos_pols, resolver, obj_counts, filepath, profile_comments=None, vdom_inspection_mode="flow", ext_resources=None, static_routes=None, policy_routes=None, ospf_data=None, vpn_data=None, interfaces=None, acl_pols=None, log_fn=None):
    wb = Workbook()

    # 1. 요약 시트 / Summary Sheet
    ws_sum = wb.active
    ws_sum.title = "Summary"
    sc(ws_sum, 1, 1, "Item", font=HDR_FONT, fill=HDR_DEFAULT_FILL, align=CENTER)
    sc(ws_sum, 1, 2, "Count", font=HDR_FONT, fill=HDR_DEFAULT_FILL, align=CENTER)

    items = [
        ("vDOM Name", vdom_name),
        ("Firewall Policy", len(fw_pols)),
        ("Local-in Policy", len(li_pols)),
        ("Central-NAT", len(cn_ents)),
        ("DNAT (VIP)", len(dn_ents)),
        ("DoS Policy", len(dos_pols)),
        ("ACL Policy", len(acl_pols) if acl_pols else 0),
        ("Static Route", len(static_routes) if static_routes else 0),
        ("Policy Route", len(policy_routes) if policy_routes else 0),
        ("OSPF Networks", len(ospf_data.get('networks', [])) if ospf_data else 0),
        ("IPsec VPN (P1)", len(vpn_data) if vpn_data else 0),
        ("Network Interfaces", len(interfaces) if interfaces else 0),
        ("External Resources", len(ext_resources) if ext_resources else 0),
        ("Address Objects", obj_counts[0]),
        ("Address Groups", obj_counts[1]),
        ("Service Objects", obj_counts[2]),
        ("Service Groups", obj_counts[3]),
        ("IP Pools", obj_counts[4]),
        ("Schedule Recurring", obj_counts[5] if len(obj_counts) > 5 else 0),
        ("Schedule Onetime", obj_counts[6] if len(obj_counts) > 6 else 0),
        ("Schedule Groups", obj_counts[7] if len(obj_counts) > 7 else 0),
    ]

    for idx, (label, cnt) in enumerate(items, 2):
        fill = EVEN_ROW_FILL if idx % 2 == 0 else ODD_ROW_FILL
        sc(ws_sum, idx, 1, label, font=FONT_DEFAULT_BOLD if idx == 2 else FONT_DEFAULT, fill=fill)
        sc(ws_sum, idx, 2, cnt, font=FONT_DEFAULT_BOLD if idx == 2 else FONT_DEFAULT, fill=fill, align=CENTER)
    auto_fit(ws_sum, min_w=15)

    # 2. 방화벽 정책 시트 / Firewall Policy Sheet
    ws_fw = wb.create_sheet("Firewall Policy")
    write_fw_policy_sheet(ws_fw, fw_pols, resolver, vdom_name, dn_ents=dn_ents,
                          profile_comments=profile_comments, vdom_inspection_mode=vdom_inspection_mode)

    # 3. 로컬 인 정책 시트 / Local-in Policy Sheet
    ws_li = wb.create_sheet("Local-in Policy")
    write_local_in_sheet(ws_li, li_pols, resolver, vdom_name)

    # 4. Central-NAT 정책 시트 / Central-NAT Policy Sheet
    ws_cn = wb.create_sheet("Central-NAT")
    write_central_nat_sheet(ws_cn, cn_ents, resolver, vdom_name)

    # 5. DNAT (VIP) 정책 시트 / DNAT (VIP) Sheet
    ws_dn = wb.create_sheet("DNAT (VIP)")
    write_dnat_sheet(ws_dn, dn_ents, resolver, vdom_name)

    # 6. DoS 정책 시트 / DoS Policy Sheet
    ws_dos = wb.create_sheet("DoS Policy")
    write_dos_sheet(ws_dos, dos_pols, resolver, vdom_name)

    # 7. ACL 정책 시트 / ACL Policy Sheet
    ws_acl = wb.create_sheet("ACL Policy")
    write_acl_sheet(ws_acl, acl_pols or [], resolver, vdom_name)

    # 7. 정적 라우팅 시트 / Static Route Sheet
    ws_sr = wb.create_sheet("Static Route")
    write_static_route_sheet(ws_sr, static_routes or [], vdom_name)

    # 8. 정책 라우팅 시트 / Policy Route Sheet
    ws_pr = wb.create_sheet("Policy Route")
    write_policy_route_sheet(ws_pr, policy_routes or [], resolver, vdom_name)

    # 9. OSPF 동적 라우팅 시트 / OSPF Sheet
    ws_ospf = wb.create_sheet("OSPF")
    write_ospf_sheet(ws_ospf, ospf_data or {}, vdom_name)

    # 10. IPsec VPN 시트 / IPsec VPN Sheet
    ws_vpn = wb.create_sheet("IPsec VPN")
    write_ipsec_vpn_sheet(ws_vpn, vpn_data or [], resolver, vdom_name)

    # 11. 네트워크 인터페이스 시트 / Network Interface Sheet
    ws_intf = wb.create_sheet("Network Interface")
    write_interface_sheet(ws_intf, interfaces or [], vdom_name)

    # 12. 외부 리소스 시트 / External Resource Sheet
    ws_ext = wb.create_sheet("External Resource")
    write_external_resource_sheet(ws_ext, ext_resources or {}, vdom_name)

    # 13. 내용 없는 빈 시트 탭 색상 빨간색으로 지정 / Highlight empty sheet tabs with red
    check_and_mark_empty_sheet_tabs(wb, TAB_COLOR_EMPTY)

    actual_saved = safe_save_workbook(wb, filepath, log_fn=log_fn)
    return actual_saved


# ================================================================
# 13. 모던 다크 테마 GUI 컴포넌트 / Modern Dark Theme GUI Components
# ================================================================

# ================================================================
#  다크 테마 색상 팔레트 / Dark Theme Color Palette
# ================================================================
C_BG_APP = "#1e1e1e"          # 메인 에디터 배경색 / Main Editor Background
C_BG_SIDEBAR = "#252526"      # 사이드바 & 카드 배경색 / Sidebar & Card Background
C_BG_PANEL = "#181818"        # 터미널 콘솔 배경색 / Terminal Console Background
C_BORDER = "#333333"          # 구분선 및 테두리 / Border & Separator
C_BORDER_LIGHT = "#3c3c3c"    # 입력창 테두리 색상 / Input Box Border
C_INPUT_BG = "#3c3c3c"        # 입력창 배경색 / Input Field Background
C_INPUT_FG = "#cccccc"        # 입력창 글자색 / Input Field Text Color
C_TEXT_MAIN = "#d4d4d4"       # 기본 텍스트 (밝은 회색) / Main Foreground Text
C_TEXT_MUTED = "#858585"      # 보조 텍스트 (중간 회색) / Muted Secondary Text
C_ACCENT_BLUE = "#007acc"         # 액센트 블루 / Accent Blue
C_ACCENT_BLUE_HOVER = "#1f8ad2"   # 액센트 블루 호버 / Accent Blue Hover
C_BTN_PRIMARY = "#0e639c"         # 기본 버튼 (블루) / Primary Button
C_BTN_PRIMARY_HOVER = "#1177bb"
C_BTN_SECONDARY = "#3a3d41"   # 보조 버튼 (다크 그레이) / Secondary Button
C_BTN_SECONDARY_HOVER = "#45494e"
C_BTN_DISABLED = "#2d2d2d"    # 비활성화 버튼 / Disabled Button
C_BTN_DISABLED_FG = "#555555"
C_TAG_vDOM = "#dcdcaa"        # vDOM 로그 태그 / vDOM Tag
C_TAG_VDOM_NAME = "#81c995"    # vDOM 이름 녹색 강조 태그 / vDOM Name Green Tag
C_TAG_INFO = "#569cd6"        # 안내 정보 태그 (파란색) / Info Tag (Blue)
C_TAG_SUCCESS = "#4ec9b0"     # 완료/성공 태그 (민트색) / Success Tag (Mint)
C_TAG_ERROR = "#f14c4c"       # 에러 태그 (빨간색) / Error Tag (Red)
C_TAG_COMMENT = "#6a9955"     # 세부 항목 태그 (초록색) / Detail Tag (Green)
C_STATUSBAR_BG = "#007acc"    # 하단 상태표시줄 배경색 / Status Bar Background
C_STATUSBAR_FG = "#ffffff"    # 하단 상태표시줄 글자색 / Status Bar Text Color


# ==============================================================================
# 커스텀 둥근 위젯 (타원형 버튼 및 일체형 입력창) / Custom Rounded Widgets
# ==============================================================================
class RoundedButton(tk.Canvas):
    """
    모던 타원형(캡슐형) 버튼 위젯 / Modern Oval / Capsule (Pill) Button Widget
    부드러운 곡면 타원형 디자인과 호버, 클릭 피드백을 지원합니다.
    Provides smooth oval capsule aesthetics with responsive hover and click animations.
    """
    def __init__(self, parent, text, bg, hover_bg, cmd=None, fg='#ffffff', font=('Segoe UI', 10),
                 height=34, width=None, radius=None, border_color=None, border_width=1,
                 disabled_bg=C_BTN_DISABLED, disabled_fg=C_BTN_DISABLED_FG, icon_type=None):
        self._target_h = height
        self._font = font
        self._text = text
        self._cmd = cmd
        self._bg = bg
        self._orig_bg = bg
        self._hover_bg = hover_bg
        self._disabled_bg = disabled_bg
        self._disabled_fg = disabled_fg
        self._normal_fg = fg
        self._border_color = border_color
        self._border_width = border_width if border_color else 0
        self._radius = radius if radius is not None else height // 2
        self._state = 'normal'
        self._is_hovered = False
        self._is_pressed = False
        self._icon_type = icon_type

        if width is None:
            dummy = tk.Label(parent, text=text, font=font)
            text_w = dummy.winfo_reqwidth()
            dummy.destroy()
            extra_icon = 28 if icon_type else 0
            self._target_w = text_w + height + 16 + extra_icon
        else:
            self._target_w = width

        parent_bg = parent.cget('bg')
        super().__init__(
            parent, width=self._target_w, height=self._target_h,
            bg=parent_bg, bd=0, highlightthickness=0, relief=tk.FLAT
        )

        self.bind('<Configure>', self._on_resize)
        self.bind('<Enter>', self._on_enter)
        self.bind('<Leave>', self._on_leave)
        self.bind('<Button-1>', self._on_press)
        self.bind('<ButtonRelease-1>', self._on_release)

    def _draw_capsule(self, x1, y1, x2, y2, r, fill, outline=''):
        w = x2 - x1
        h = y2 - y1
        r = min(r, h // 2, w // 2)
        if r <= 0:
            self.create_rectangle(x1, y1, x2, y2, fill=fill, outline=outline, tags='shape')
            return
        self.create_oval(x1, y1, x1 + 2 * r, y2, fill=fill, outline=outline, tags='shape')
        self.create_oval(x2 - 2 * r, y1, x2, y2, fill=fill, outline=outline, tags='shape')
        self.create_rectangle(x1 + r, y1, x2 - r, y2, fill=fill, outline=outline, tags='shape')

    def _redraw(self):
        self.delete('all')
        w = self.winfo_width()
        h = self.winfo_height()
        if w <= 1 or h <= 1:
            w = self._target_w
            h = self._target_h

        if self._state == 'disabled':
            fill_color = self._disabled_bg
            text_color = self._disabled_fg
            border_col = None
            self.config(cursor='arrow')
        else:
            fill_color = self._hover_bg if self._is_hovered else self._bg
            text_color = self._normal_fg
            border_col = self._border_color
            self.config(cursor='hand2')

        r = min(h // 2, self._radius)

        if border_col and self._border_width > 0:
            self._draw_capsule(0, 0, w, h, r, border_col, '')
            bw = self._border_width
            self._draw_capsule(bw, bw, w - bw, h - bw, max(1, r - bw), fill_color, '')
        else:
            self._draw_capsule(0, 0, w, h, r, fill_color, '')

        offset_y = 1 if (self._is_pressed and self._state != 'disabled') else 0

        if self._icon_type == 'xlsx':
            # 문서 (.xlsx) 벡터 아이콘 렌더링 / Draw Vector XLSX Document Icon
            dummy = tk.Label(self, text=self._text, font=self._font)
            tw = dummy.winfo_reqwidth()
            dummy.destroy()

            icon_w = 20
            gap = 10
            total_content_w = icon_w + gap + tw
            start_x = (w - total_content_w) // 2

            ix = start_x + icon_w // 2
            iy = (h // 2) + offset_y

            # 1. 문서 종이 시트 (상단 우측 모서리 접힘) / Document sheet with folded corner
            pw = 14
            ph = 18
            px1 = ix - pw // 2
            py1 = iy - ph // 2
            px2 = px1 + pw
            py2 = py1 + ph
            fold = 4

            if self._state == 'disabled':
                sheet_color = '#606064'
                fold_color = '#4a4a4e'
                line_color = '#454548'
                badge_color = '#3e4e42'
                x_color = '#888888'
            else:
                sheet_color = '#ffffff'
                fold_color = '#b8d6fc'
                line_color = '#b0c4de'
                badge_color = '#107c41'  # Office Excel Green
                x_color = '#ffffff'

            pts = [
                px1, py1,
                px2 - fold, py1,
                px2, py1 + fold,
                px2, py2,
                px1, py2
            ]
            self.create_polygon(pts, fill=sheet_color, outline='', tags='icon')
            self.create_polygon(px2 - fold, py1, px2 - fold, py1 + fold, px2, py1 + fold, fill=fold_color, outline='', tags='icon')

            # 스프레드시트 눈금선 / Spreadsheet grid lines
            for ly in [py1 + 6, py1 + 10, py1 + 14]:
                self.create_line(px1 + 3, ly, px2 - 3, ly, fill=line_color, width=1, tags='icon')

            # 2. 엑셀 녹색 'X' 뱃지 / Green Excel 'X' Badge
            gw, gh = 11, 11
            gx = px1 - 3
            gy = iy - 2
            self.create_rectangle(gx, gy, gx + gw, gy + gh, fill=badge_color, outline='', tags='icon')
            self.create_text(gx + gw // 2 + 1, gy + gh // 2, text='X', fill=x_color, font=('Segoe UI', 7, 'bold'), tags='icon')

            # 3. 텍스트 레이블 / Text Label
            text_x = start_x + icon_w + gap + tw // 2
            self.create_text(text_x, iy, text=self._text, fill=text_color, font=self._font, tags='label')
        else:
            cy = (h // 2) + offset_y
            self.create_text(w // 2, cy, text=self._text, fill=text_color, font=self._font, tags='label')

    def _on_resize(self, event):
        self._redraw()

    def _on_enter(self, event):
        if self._state != 'disabled':
            self._is_hovered = True
            self._redraw()

    def _on_leave(self, event):
        self._is_hovered = False
        self._is_pressed = False
        self._redraw()

    def _on_press(self, event):
        if self._state != 'disabled':
            self._is_pressed = True

    def _on_release(self, event):
        if self._state != 'disabled' and self._is_pressed:
            self._is_pressed = False
            if 0 <= event.x <= self.winfo_width() and 0 <= event.y <= self.winfo_height():
                if self._cmd:
                    self._cmd()
            self._redraw()

    def set_state(self, state):
        self._state = state
        self._redraw()

    def set_text(self, text):
        self._text = text
        self._redraw()

    def config(self, **kwargs):
        if 'state' in kwargs:
            st = kwargs.pop('state')
            self._state = 'disabled' if str(st) == str(tk.DISABLED) else 'normal'
            self._redraw()
        if 'text' in kwargs:
            self._text = kwargs.pop('text')
            self._redraw()
        if kwargs:
            super().config(**kwargs)


class RoundedEntry(tk.Frame):
    """
    모던 둥근 모서리(타원형) 텍스트 입력창 위젯 / Modern Rounded Capsule Text Entry Widget
    버튼과 동일한 높이(일체감) 및 둥근 타원형 테두리를 제공합니다.
    Maintains identical height and rounded capsule style with adjacent buttons for visual harmony.
    """
    def __init__(self, parent, font=('Consolas', 10), bg=C_INPUT_BG, fg=C_INPUT_FG,
                 border_color=C_BORDER_LIGHT, focus_border=C_ACCENT_BLUE, height=34, radius=None):
        parent_bg = parent.cget('bg')
        super().__init__(parent, bg=parent_bg, height=height)
        self.pack_propagate(False)
        self.grid_propagate(False)

        self._target_h = height
        self._radius = radius if radius is not None else height // 2
        self._bg = bg
        self._border_color = border_color
        self._focus_border = focus_border
        self._is_focused = False

        self.canvas = tk.Canvas(self, bg=parent_bg, bd=0, highlightthickness=0, relief=tk.FLAT, height=height)
        self.canvas.pack(fill=tk.BOTH, expand=True)

        self.entry = tk.Entry(
            self.canvas, font=font, bg=bg, fg=fg, insertbackground='#ffffff',
            bd=0, relief=tk.FLAT, highlightthickness=0
        )
        self.win_id = self.canvas.create_window(0, 0, window=self.entry, anchor='w')

        self.entry.bind('<FocusIn>', self._on_focus_in)
        self.entry.bind('<FocusOut>', self._on_focus_out)
        self.canvas.bind('<Configure>', self._on_resize)
        self.canvas.bind('<Button-1>', lambda e: self.entry.focus_set())

    def _draw_capsule(self, x1, y1, x2, y2, r, fill, outline=''):
        w = x2 - x1
        h = y2 - y1
        r = min(r, h // 2, w // 2)
        if r <= 0:
            self.canvas.create_rectangle(x1, y1, x2, y2, fill=fill, outline=outline, tags='bg_shape')
            return
        self.canvas.create_oval(x1, y1, x1 + 2 * r, y2, fill=fill, outline=outline, tags='bg_shape')
        self.canvas.create_oval(x2 - 2 * r, y1, x2, y2, fill=fill, outline=outline, tags='bg_shape')
        self.canvas.create_rectangle(x1 + r, y1, x2 - r, y2, fill=fill, outline=outline, tags='bg_shape')

    def _redraw(self):
        self.canvas.delete('bg_shape')
        w = self.canvas.winfo_width()
        h = self.canvas.winfo_height()
        if w <= 1 or h <= 1:
            h = self._target_h
            w = 300

        r = min(h // 2, self._radius)
        b_col = self._focus_border if self._is_focused else self._border_color

        # 외곽 테두리 렌더링 / Render Outer Border
        self._draw_capsule(0, 0, w, h, r, b_col, '')
        # 내부 배경 채우기 / Render Inner Background Fill
        bw = 1
        self._draw_capsule(bw, bw, w - bw, h - bw, max(1, r - bw), self._bg, '')

        pad_x = max(14, r)
        entry_w = max(10, w - 2 * pad_x)
        self.canvas.coords(self.win_id, pad_x, h // 2)
        self.canvas.itemconfigure(self.win_id, width=entry_w)

    def _on_resize(self, event):
        self._redraw()

    def _on_focus_in(self, event):
        self._is_focused = True
        self._redraw()

    def _on_focus_out(self, event):
        self._is_focused = False
        self._redraw()

    # Entry 위젯 표준 메서드 위임 / Delegate Standard Entry Widget Methods
    def get(self):
        return self.entry.get()

    def insert(self, index, string):
        return self.entry.insert(index, string)

    def delete(self, first, last=None):
        return self.entry.delete(first, last)

    def icursor(self, index):
        return self.entry.icursor(index)

    def xview(self, *args):
        return self.entry.xview(*args)


class RoundedBadge(tk.Canvas):
    """
    정적 타원형 정보 뱃지 위젯 / Static Rounded Capsule Badge Widget
    마우스 클릭이나 호버 반응이 없는 순수 시각적 정보 뱃지입니다.
    A static visual indicator badge with no click, hover, or button interactions.
    """
    def __init__(self, parent, text, bg, fg, font=('Segoe UI', 9), height=22):
        dummy = tk.Label(parent, text=text, font=font)
        text_w = dummy.winfo_reqwidth()
        dummy.destroy()
        w = text_w + 18

        parent_bg = parent.cget('bg')
        super().__init__(
            parent, width=w, height=height, bg=parent_bg, bd=0, highlightthickness=0, relief=tk.FLAT
        )
        r = height // 2
        # 좌우 반원 및 중앙 직사각형 / Left/Right Arcs and Center Rectangle
        self.create_oval(0, 0, 2 * r, height, fill=bg, outline='')
        self.create_oval(w - 2 * r, 0, w, height, fill=bg, outline='')
        self.create_rectangle(r, 0, w - r, height, fill=bg, outline='')
        self.create_text(w // 2, height // 2, text=text, fill=fg, font=font)


class ModernCheckbox(tk.Frame):
    """
    모던 다크 테마 커스텀 체크박스 위젯 / Modern Dark Theme Custom Checkbox Widget
    표준 Tkinter 체크박스보다 글자 크기에 맞춰 시각적으로 균형 잡힌 크기(18x18px)와 깔끔한 체크마크를 제공합니다.
    """
    def __init__(self, parent, text, variable, bg=C_BG_SIDEBAR, fg=C_TEXT_MAIN,
                 box_size=18, font=('Segoe UI', 10), active_fg='#ffffff'):
        super().__init__(parent, bg=bg, cursor='hand2')
        self.var = variable
        self.box_size = box_size
        self.fg = fg
        self.active_fg = active_fg
        self.bg_color = bg

        self.canvas = tk.Canvas(self, width=box_size, height=box_size, bg=bg, bd=0, highlightthickness=0)
        self.canvas.pack(side=tk.LEFT, padx=(0, 8))

        self.lbl = tk.Label(self, text=text, font=font, fg=fg, bg=bg, cursor='hand2')
        self.lbl.pack(side=tk.LEFT)

        self.bind('<Button-1>', self._toggle)
        self.canvas.bind('<Button-1>', self._toggle)
        self.lbl.bind('<Button-1>', self._toggle)

        self._is_hovered = False
        self.bind('<Enter>', self._on_enter)
        self.bind('<Leave>', self._on_leave)
        self.canvas.bind('<Enter>', self._on_enter)
        self.canvas.bind('<Leave>', self._on_leave)
        self.lbl.bind('<Enter>', self._on_enter)
        self.lbl.bind('<Leave>', self._on_leave)

        # 변수 값 변경 시 자동 재렌더링 / Auto-redraw on variable change
        try:
            self.var.trace_add('write', lambda *args: self._draw())
        except AttributeError:
            self.var.trace('w', lambda *args: self._draw())

        self._draw()

    def _on_enter(self, event=None):
        self._is_hovered = True
        self.lbl.config(fg=self.active_fg)
        self._draw()

    def _on_leave(self, event=None):
        self._is_hovered = False
        self.lbl.config(fg=self.fg)
        self._draw()

    def _toggle(self, event=None):
        self.var.set(not self.var.get())

    def _draw(self):
        self.canvas.delete('all')
        s = self.box_size
        val = self.var.get()

        if val:
            fill = C_ACCENT_BLUE if not self._is_hovered else C_ACCENT_BLUE_HOVER
            border = fill
        else:
            fill = C_INPUT_BG
            border = '#6e6e72' if self._is_hovered else '#4e4e52'

        p = 1
        self.canvas.create_rectangle(p, p, s - p, s - p, fill=fill, outline=border, width=1.5)

        if val:
            pts = [
                p + s * 0.22, p + s * 0.50,
                p + s * 0.44, p + s * 0.72,
                p + s * 0.80, p + s * 0.28
            ]
            self.canvas.create_line(pts, fill='#ffffff', width=2, capstyle=tk.ROUND, joinstyle=tk.ROUND)



# ==============================================================================
# 내장 Fortinet 아이콘 (Base64) / Embedded Fortinet Icon (Self-contained)
# ==============================================================================
FORTINET_ICO_BASE64 = """
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
0/RILx38/SuU/D8/osQ3bqY9l32CWlsaqL+xRvfxo7JenU1t5qwCQMreOyDNsYjubytq9zfyU0XAYF2M
Oi75S2pjc67z5uup7/bbaIj72iP8tPVqmDW09Q4a2pKjrXf4P9fvB3O+7+ou6rnxaure0MhPwDpuAPb+
8EjRLsBwUxV1XPZxGrj/Dup/eCv1PbRFWjOD2+5jbaWBbVt8GhSv21kPhUf98nUrJa6/knqbVeN3AVB4
m6wMDhuros7P/A0N3r3ZUx7qddib994y2ZrzuTxW3wkqG3Gd/s23Uf+tN1P7dV+gPZ++mPZc0EztfH89
/PuDcs1J1LnvpDPaDwCE0wKI22a12/jT0uyM0khcVz5+sow2RmmgUUQfqqYebpBB6uRzO0RftESJ8/vi
QlEpMYhnX6tPdk2U2WtrMbfQIh2iK9PL3xWzHB36tZstiEpTH9/3SLzWZ/I7kM7xkOO85/QPiQhRojHq
8mjlz7v4tauMskmIPSQFylion681GIvQcKOIRaGc0Yp6Ym/kGdf3KT4blqP9UQAg9IOEMXtKpkbFo2/y
r9BSwRly/bPnSldIrbSunN7PfIrZOxddeQcqk840UbTM6c6IE0DCfYIWt4K8CvrM2WodC0hTgc9yV7cV
/F4saC486k+3SI+YAVkkMEbKUw5i4VfSydecMtGgD7qfdE5e5O7Pt8sorWeWUvLeInrZeY1nbj+3zgAA
IdxCbPcrVaORJhtTPqlHn+2tmekS12cXbOwFtPjehtIrT9JeSy43D1U5fhHGnDXmwUqeQckGF/OCTn+e
s0BGNbKqIgCI6hmDGg/Uo87e+lS5ZaPvy/vqgirq5OuYHtlPOSP67qxEKuSOPQCAPIeikZy5ZZf2lTFy
61on5TuUiJah6pN0bqH57+gJpD96SseLTufvAQAQBAEAEAQBABAEAQAQBAEAEAQBABAEAQAQBAEAEAQB
ABAEAQAQBAEAEAQBABAEAQAQBAEAEAQBABAEAQAQBAEAEAQBABAEAQAQBAEAEAQBABAEAQAQBAAgEyAI
AIAgCACAIAgAgCAIAIAgCACAIAgAgCAIAIAgCACAIAgAgCAIAIAgCACAIAgAgCAIAIAgCACAIAgAgCAI
AIAgCACAIAgAgCAIAIAgCACAIAgAgCAIAIAgCACAIAgAgCAIAIAgCACAIAgAgCAIAIAgCACAIAgAgCAI
AIAgCACAIAgAgCAIAIAgCACAIAgAgCAAAJkAQQAABEEAAARBAAAEQctA/x/97mDHN/dlXgAAAABJRU5E
rkJggg==
"""


def get_fortinet_icon_path():
    """
    내장된 Base64 아이콘 데이터를 임시 폴더에 캐싱하여 경로 반환
    Caches embedded Base64 icon into temp directory and returns its file path.
    """
    try:
        temp_dir = tempfile.gettempdir()
        target_path = os.path.join(temp_dir, "fortinet.ico")
        if not os.path.exists(target_path) or os.path.getsize(target_path) != 10864:
            raw_data = base64.b64decode(FORTINET_ICO_BASE64.strip())
            with open(target_path, "wb") as f_out:
                f_out.write(raw_data)
        return target_path
    except Exception:
        return ""


class FortiGateGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("FortiGate Policy to Excel Exporter  v1.5")
        self.root.minsize(860, 480)
        self.root.configure(bg=C_BG_APP)

        # 실행한 모니터의 작업 영역 정중앙에 창 배치 / Center window on active monitor
        self._center_window(1000, 720)

        # 창 제목 좌측의 기본 아이콘(깃털) 완전 제거 및 작업표시줄 아이콘 설정 / Completely remove title bar icon & configure taskbar icon
        self._remove_title_icon()

        self.last_target_dir = ""
        self.is_running = False

        self._setup_styles()
        self._create_widgets()
        # 윈도우 매핑 후 타이틀바 아이콘 제거 재확정 / Re-confirm title icon removal after window mapping
        self.root.after(50, self._remove_title_icon)

    def _center_window(self, width=1000, height=720):
        """
        현재 프로그램이 실행된 모니터의 작업 영역 정중앙에 창을 배치
        Centers the window precisely in the work area of the active monitor.
        """
        positioned = False
        x, y = 0, 0

        if sys.platform == 'win32':
            try:
                from ctypes import windll, wintypes, Structure, byref

                class POINT(Structure):
                    _fields_ = [('x', wintypes.LONG), ('y', wintypes.LONG)]

                class RECT(Structure):
                    _fields_ = [('left', wintypes.LONG), ('top', wintypes.LONG),
                                ('right', wintypes.LONG), ('bottom', wintypes.LONG)]

                class MONITORINFO(Structure):
                    _fields_ = [('cbSize', wintypes.DWORD),
                                ('rcMonitor', RECT),
                                ('rcWork', RECT),
                                ('dwFlags', wintypes.DWORD)]

                user32 = windll.user32
                pt = POINT()
                user32.GetCursorPos(byref(pt))

                MONITOR_DEFAULTTONEAREST = 2
                h_mon = user32.MonitorFromPoint(pt, MONITOR_DEFAULTTONEAREST)
                if h_mon:
                    mi = MONITORINFO()
                    mi.cbSize = ctypes.sizeof(MONITORINFO)
                    if user32.GetMonitorInfoW(h_mon, byref(mi)):
                        work_w = mi.rcWork.right - mi.rcWork.left
                        work_h = mi.rcWork.bottom - mi.rcWork.top
                        x = mi.rcWork.left + max(0, (work_w - width) // 2)
                        y = mi.rcWork.top + max(0, (work_h - height) // 2)
                        positioned = True
            except Exception:
                pass

        if not positioned:
            self.root.update_idletasks()
            sw = self.root.winfo_screenwidth()
            sh = self.root.winfo_screenheight()
            x = max(0, (sw - width) // 2)
            y = max(0, (sh - height) // 2)

        self.root.geometry(f"{width}x{height}+{x}+{y}")

    def _remove_title_icon(self):
        """
        창 제목 좌측의 기본 아이콘(깃털)은 완전히 숨기고, 작업표시줄에는 fortinet.ico를 표시
        - Windows DWM 상에서 1x1 완전 투명 HICON 핸들을 ICON_SMALL에 적용 -> 창 제목줄 아이콘 완전 미표시
        - fortinet.ico가 존재할 경우 ICON_BIG 및 Window Class Icon에 적용 -> 작업표시줄에는 Fortinet 아이콘 표시
        - SetCurrentProcessExplicitAppUserModelID 설정으로 작업표시줄 앱 분리
        """
        if sys.platform != 'win32':
            return
        try:
            from ctypes import windll, wintypes, Structure, byref

            # 작업표시줄 AppUserModelID 등록 / Register AppUserModelID for Taskbar separation
            try:
                windll.shell32.SetCurrentProcessExplicitAppUserModelID(
                    'fortinet.fortigate.policytoexcel.exporter.1.0'
                )
            except Exception:
                pass

            self.root.update_idletasks()
            user32 = windll.user32
            gdi32 = windll.gdi32

            hwnd = user32.GetParent(self.root.winfo_id())
            if not hwnd:
                hwnd = self.root.winfo_id()

            # 1. 1x1 완전 투명 마스크 및 비트맵 생성 (창 제목줄 소형 아이콘 제거용) / 1. Create 1x1 fully transparent mask & bitmap (hides title bar icon)
            hbmMask = gdi32.CreateBitmap(1, 1, 1, 1, None)
            mask_bytes = (ctypes.c_ubyte * 1)(0xFF)
            gdi32.SetBitmapBits(hbmMask, 1, mask_bytes)
            hbmColor = gdi32.CreateBitmap(1, 1, 1, 32, None)

            class ICONINFO(Structure):
                _fields_ = [
                    ('fIcon', wintypes.BOOL),
                    ('xHotspot', wintypes.DWORD),
                    ('yHotspot', wintypes.DWORD),
                    ('hbmMask', wintypes.HBITMAP),
                    ('hbmColor', wintypes.HBITMAP)
                ]

            ii = ICONINFO()
            ii.fIcon = True
            ii.xHotspot = 0
            ii.yHotspot = 0
            ii.hbmMask = hbmMask
            ii.hbmColor = hbmColor

            self._transparent_hicon = user32.CreateIconIndirect(byref(ii))
            gdi32.DeleteObject(hbmMask)
            gdi32.DeleteObject(hbmColor)

            # 2. 작업표시줄용 fortinet.ico 아이콘 로드 (내장 Base64 데이터에서 자동 추출) / Load embedded fortinet.ico for Taskbar
            fortinet_ico_path = get_fortinet_icon_path()
            hIcon_big = self._transparent_hicon

            if fortinet_ico_path and os.path.isfile(fortinet_ico_path):
                IMAGE_ICON = 1
                LR_LOADFROMFILE = 0x00000010
                LR_DEFAULTSIZE = 0x00000040
                loaded_icon = user32.LoadImageW(
                    None, fortinet_ico_path, IMAGE_ICON, 0, 0, LR_LOADFROMFILE | LR_DEFAULTSIZE
                )
                if loaded_icon:
                    hIcon_big = loaded_icon
                    self._fortinet_hicon = loaded_icon

            # 3. WM_SETICON 적용: SMALL은 투명, BIG은 fortinet.ico / 3. Apply WM_SETICON: Transparent for SMALL (title bar), fortinet.ico for BIG (taskbar)
            WM_SETICON = 0x80
            ICON_SMALL = 0
            ICON_BIG = 1
            user32.SendMessageW(hwnd, WM_SETICON, ICON_SMALL, self._transparent_hicon)
            user32.SendMessageW(hwnd, WM_SETICON, ICON_BIG, hIcon_big)

            # 작업표시줄 클래스 대형 아이콘 등록 / Register Class Icon for Taskbar
            if hIcon_big != self._transparent_hicon:
                GCLP_HICON = -14
                SetClassLongPtr = getattr(user32, 'SetClassLongPtrW', user32.SetClassLongW)
                SetClassLongPtr(hwnd, GCLP_HICON, hIcon_big)

            SWP_NOSIZE = 0x0001
            SWP_NOMOVE = 0x0002
            SWP_NOZORDER = 0x0004
            SWP_FRAMECHANGED = 0x0020
            user32.SetWindowPos(hwnd, 0, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_FRAMECHANGED)
        except Exception:
            pass

    def _setup_styles(self):
        self.style = ttk.Style()
        try:
            self.style.theme_use('clam')
        except:
            pass

        # 프로그레스바 스타일 (블루) / Progressbar Style (Accent Blue)
        self.style.configure(
            "VSCode.Horizontal.TProgressbar",
            troughcolor=C_BG_SIDEBAR,
            background=C_ACCENT_BLUE,
            darkcolor=C_ACCENT_BLUE,
            lightcolor=C_ACCENT_BLUE,
            bordercolor=C_BG_SIDEBAR,
            thickness=6
        )

    def _create_widgets(self):
        # -------------------------------------------------------------
        # 1. 상단 타이틀 바 / 1. Top Title Bar (Breadcrumb Style)
        # -------------------------------------------------------------
        header_frame = tk.Frame(self.root, bg=C_BG_SIDEBAR, bd=0, highlightthickness=1, highlightbackground=C_BORDER)
        header_frame.pack(fill=tk.X, side=tk.TOP)

        header_content = tk.Frame(header_frame, bg=C_BG_SIDEBAR, padx=16, pady=10)
        header_content.pack(fill=tk.X)

        title_left = tk.Frame(header_content, bg=C_BG_SIDEBAR)
        title_left.pack(side=tk.LEFT, fill=tk.Y)

        title_lbl = tk.Label(
            title_left,
            text="FORTIGATE POLICY TO EXCEL EXPORTER",
            font=('Segoe UI', 13, 'bold'),
            fg="#ffffff",
            bg=C_BG_SIDEBAR
        )
        title_lbl.pack(anchor=tk.W)

        subtitle_lbl = tk.Label(
            title_left,
            text="FortiOS Firewall Policy & Object Exporter System",
            font=('Segoe UI', 10),
            fg=C_TEXT_MUTED,
            bg=C_BG_SIDEBAR
        )
        subtitle_lbl.pack(anchor=tk.W, pady=(2, 0))

        # 우측 정적 정보 뱃지 / Right-side Static Info Badges
        badges_frame = tk.Frame(header_content, bg=C_BG_SIDEBAR)
        badges_frame.pack(side=tk.RIGHT, anchor=tk.E)

        self._create_badge(badges_frame, "FortiOS 6.x / 7.x", "#264f78", "#9cdcfe")
        self._create_badge(badges_frame, "Multi-vDOM", "#37373d", "#cccccc")

        # -------------------------------------------------------------
        # 2. 하단 상태 표시줄 / 2. Bottom Status Bar (Always Pinned to Bottom)
        # -------------------------------------------------------------
        status_bar = tk.Frame(self.root, bg=C_STATUSBAR_BG, height=28)
        status_bar.pack(fill=tk.X, side=tk.BOTTOM)

        self.status_lbl = tk.Label(
            status_bar,
            text="● Ready",
            font=('Segoe UI', 10),
            fg=C_STATUSBAR_FG,
            bg=C_STATUSBAR_BG,
            anchor='w',
            justify=tk.LEFT
        )
        self.status_lbl.pack(side=tk.LEFT, padx=10, pady=3, anchor=tk.W)

        status_right = tk.Label(
            status_bar,
            text="github.com/smilestory-net",
            font=('Segoe UI', 10),
            fg=C_STATUSBAR_FG,
            bg=C_STATUSBAR_BG,
            cursor='hand2'
        )
        status_right.pack(side=tk.RIGHT, padx=14, pady=3)
        status_right.bind("<Button-1>", lambda e: webbrowser.open("https://github.com/smilestory-net"))

        # -------------------------------------------------------------
        # 3. 메인 중앙 컨테이너 / 3. Main Center Container (Flexible Remaining Area)
        # -------------------------------------------------------------
        main_container = tk.Frame(self.root, bg=C_BG_APP, padx=16, pady=10)
        main_container.pack(fill=tk.BOTH, expand=True)

        # 설정 카드 (경로 및 옵션) / Configuration Card (Paths & Options)
        config_card = tk.Frame(main_container, bg=C_BG_SIDEBAR, bd=0, highlightthickness=1, highlightbackground=C_BORDER)
        config_card.pack(fill=tk.X, pady=(0, 10))

        # 카드 헤더 바 / Card Header Bar
        card_header = tk.Frame(config_card, bg=C_BG_SIDEBAR, padx=14, pady=8)
        card_header.pack(fill=tk.X)
        tk.Label(
            card_header,
            text="CONFIGURATION & PATHS",
            font=('Segoe UI', 10, 'bold'),
            fg=C_TEXT_MUTED,
            bg=C_BG_SIDEBAR
        ).pack(anchor=tk.W)

        # 구분선 / Separator Line
        tk.Frame(config_card, bg=C_BORDER, height=1).pack(fill=tk.X)

        card_body = tk.Frame(config_card, bg=C_BG_SIDEBAR, padx=14, pady=12)
        card_body.pack(fill=tk.X)

        # 설정 파일 선택 행 / Config File Path Row
        tk.Label(
            card_body, text="Config File (.conf) :",
            font=('Segoe UI', 10), fg=C_TEXT_MAIN, bg=C_BG_SIDEBAR, anchor=tk.W
        ).grid(row=0, column=0, sticky=tk.W, pady=6)

        self.entry_conf = RoundedEntry(card_body, height=34)
        self.entry_conf.grid(row=0, column=1, sticky=tk.EW, padx=(10, 10), pady=6)

        self.btn_browse_conf = RoundedButton(
            card_body, "Select File...", C_BTN_SECONDARY, C_BTN_SECONDARY_HOVER, self._browse_conf,
            font=('Segoe UI', 10), height=34, width=160
        )
        self.btn_browse_conf.grid(row=0, column=2, pady=6)

        # 출력 디렉터리 선택 행 / Output Directory Path Row
        tk.Label(
            card_body, text="Output Directory :",
            font=('Segoe UI', 10), fg=C_TEXT_MAIN, bg=C_BG_SIDEBAR, anchor=tk.W
        ).grid(row=1, column=0, sticky=tk.W, pady=6)

        self.entry_out = RoundedEntry(card_body, height=34)
        self.entry_out.grid(row=1, column=1, sticky=tk.EW, padx=(10, 10), pady=6)

        self.btn_browse_out = RoundedButton(
            card_body, "Select Folder...", C_BTN_SECONDARY, C_BTN_SECONDARY_HOVER, self._browse_out,
            font=('Segoe UI', 10), height=34, width=160
        )
        self.btn_browse_out.grid(row=1, column=2, pady=6)

        card_body.columnconfigure(1, weight=1)

        # 옵션 체크박스 / Option Checkboxes (좌측 라벨과 동일하게 열 0부터 시작하여 정렬)
        opt_frame = tk.Frame(card_body, bg=C_BG_SIDEBAR)
        opt_frame.grid(row=2, column=0, columnspan=3, sticky=tk.W, pady=(12, 4))

        self.var_open_folder = tk.BooleanVar(value=True)
        self.chk_open = ModernCheckbox(
            opt_frame,
            "Automatically open result directory upon completion",
            self.var_open_folder,
            bg=C_BG_SIDEBAR,
            fg=C_TEXT_MAIN,
            box_size=18,
            font=('Segoe UI', 10)
        )
        self.chk_open.pack(side=tk.LEFT)

        # -------------------------------------------------------------
        # 4. 액션 바 (변환 실행 및 결과 폴더 열기) / 4. Action Bar (Execution & Results)
        # -------------------------------------------------------------
        action_frame = tk.Frame(main_container, bg=C_BG_APP)
        action_frame.pack(fill=tk.X, pady=(4, 12))

        # 두 버튼의 비율을 이상적인 7:3 비율로 배분 / Balanced Grid Layout (7:3 Ratio)
        action_frame.columnconfigure(0, weight=7) # 엑셀 변환 실행 / Export Policy to Excel
        action_frame.columnconfigure(1, weight=3) # 결과 폴더 열기 / Open Result Folder

        # 왼쪽 '엑셀 정책 변환 실행' 버튼 (문서 .xlsx 벡터 아이콘 적용) / Left-side 'Start Export to Excel' Button with XLSX Document Icon
        self.btn_run = RoundedButton(
            action_frame,
            "Start Export to Excel",
            C_BTN_PRIMARY,
            C_BTN_PRIMARY_HOVER,
            self._start_conversion,
            font=('Segoe UI', 11, 'bold'),
            height=42,
            icon_type='xlsx'
        )
        self.btn_run.grid(row=0, column=0, sticky=tk.EW, padx=(0, 6))

        # 오른쪽 '결과 폴더 열기' 버튼 / Right-side 'Open Result Folder' Button
        self.btn_open_res = RoundedButton(
            action_frame,
            "📁  Open Result Folder",
            C_BTN_SECONDARY,
            C_BTN_SECONDARY_HOVER,
            self._open_result_folder,
            font=('Segoe UI', 10),
            height=42
        )
        self.btn_open_res.grid(row=0, column=1, sticky=tk.EW, padx=(6, 0))
        self._set_btn_state(self.btn_open_res, False)

        # -------------------------------------------------------------
        # 4. 슬림 프로그레스 바 (청색 게이지) / 4. Slim Progress Bar (Blue Gauge)
        # -------------------------------------------------------------
        self.progressbar = ttk.Progressbar(
            main_container,
            mode='determinate',
            style="VSCode.Horizontal.TProgressbar"
        )
        self.progressbar.pack(fill=tk.X, pady=(0, 6))

        # -------------------------------------------------------------
        # 5. 터미널 콘솔 패널 (인터랙티브 탭 시스템) / 5. Terminal Console Panel (Interactive Tab System)
        # -------------------------------------------------------------
        terminal_panel = tk.Frame(
            main_container, bg=C_BG_PANEL, bd=0, highlightthickness=1, highlightbackground=C_BORDER
        )
        terminal_panel.pack(fill=tk.BOTH, expand=True)

        # 탭 헤더 바 / Tab Header Bar
        tab_bar = tk.Frame(terminal_panel, bg=C_BG_SIDEBAR, height=30)
        tab_bar.pack(fill=tk.X, side=tk.TOP)

        # 우측 로그 지우기 버튼 / Right-side Clear Button
        btn_clear = RoundedButton(
            tab_bar,
            text="⊘ Clear",
            bg=C_BG_SIDEBAR,
            hover_bg=C_BORDER,
            cmd=self._clear_log,
            fg=C_TEXT_MUTED,
            font=('Segoe UI', 9),
            height=24,
            width=70,
            border_color=C_BORDER,
            border_width=1
        )
        btn_clear.pack(side=tk.RIGHT, padx=10, pady=3)

        # 인터랙티브 탭 시스템 구축 / Build Interactive Tab System
        self._tabs = {}
        self._tab_views = {}
        self._active_tab = "OUTPUT"
        self._problem_count = 0

        tab_defs = [
            ("PROBLEMS", "PROBLEMS  0"),
            ("OUTPUT", "OUTPUT"),
            ("DEBUG CONSOLE", "DEBUG CONSOLE"),
            ("TERMINAL", "TERMINAL")
        ]

        view_container = tk.Frame(terminal_panel, bg=C_BG_PANEL)
        view_container.pack(fill=tk.BOTH, expand=True)

        for tab_id, label_text in tab_defs:
            txt = scrolledtext.ScrolledText(
                view_container,
                wrap=tk.WORD,
                font=('Consolas', 10),
                bg=C_BG_PANEL,
                fg=C_TEXT_MAIN,
                insertbackground="#ffffff",
                selectbackground="#264f78",
                relief=tk.FLAT,
                bd=0,
                padx=10,
                pady=8
            )
            txt.tag_config('info', foreground=C_TAG_INFO)
            txt.tag_config('vdom', foreground=C_TAG_vDOM, font=('Consolas', 10, 'bold'))
            txt.tag_config('vdom_name', foreground=C_TAG_VDOM_NAME, font=('Consolas', 10, 'bold'))
            txt.tag_config('success', foreground=C_TAG_SUCCESS, font=('Consolas', 10, 'bold'))
            txt.tag_config('error', foreground=C_TAG_ERROR, font=('Consolas', 10, 'bold'))
            txt.tag_config('warning', foreground="#fdd663", font=('Consolas', 10, 'bold'))
            txt.tag_config('debug', foreground="#8ab4f8", font=('Consolas', 10))
            txt.tag_config('trace', foreground="#d2a8ff", font=('Consolas', 9))
            txt.tag_config('comment', foreground=C_TAG_COMMENT)
            txt.tag_config('muted', foreground=C_TEXT_MUTED)

            self._tab_views[tab_id] = txt

            # 탭 헤더 라벨 및 밑줄 / Tab Header Label & Underline Indicator
            t_frame = tk.Frame(tab_bar, bg=C_BG_SIDEBAR, cursor='hand2')
            t_frame.pack(side=tk.LEFT, padx=6, fill=tk.Y)

            lbl = tk.Label(
                t_frame, text=label_text, font=('Segoe UI', 10),
                fg=C_TEXT_MUTED, bg=C_BG_SIDEBAR, pady=5, cursor='hand2'
            )
            lbl.pack()

            underline = tk.Frame(t_frame, bg=C_BG_SIDEBAR, height=2)
            underline.pack(fill=tk.X, side=tk.BOTTOM)

            self._tabs[tab_id] = {
                'frame': t_frame,
                'label': lbl,
                'underline': underline,
                'base_text': label_text
            }

            def make_handler(tid=tab_id):
                return lambda e: self._switch_tab(tid)

            t_frame.bind('<Button-1>', make_handler())
            lbl.bind('<Button-1>', make_handler())
            underline.bind('<Button-1>', make_handler())

            def make_hover_enter(tid=tab_id):
                return lambda e: self._on_tab_hover(tid, True)
            def make_hover_leave(tid=tab_id):
                return lambda e: self._on_tab_hover(tid, False)

            t_frame.bind('<Enter>', make_hover_enter())
            t_frame.bind('<Leave>', make_hover_leave())
            lbl.bind('<Enter>', make_hover_enter())
            lbl.bind('<Leave>', make_hover_leave())

        self.log_text = self._tab_views["OUTPUT"]
        self.problems_text = self._tab_views["PROBLEMS"]
        self.debug_text = self._tab_views["DEBUG CONSOLE"]
        self.terminal_text = self._tab_views["TERMINAL"]

        self._switch_tab("OUTPUT")
        self._init_tab_contents()

    # -------------------------------------------------------------
    # 헬퍼 메서드: 위젯 커스텀 생성 / Helper Methods: Custom Widget Creation
    # -------------------------------------------------------------
    def _create_badge(self, parent, text, bg, fg):
        """
        정적 타원형 정보 뱃지 생성 / Create Static Oval/Capsule Info Badge
        """
        badge = RoundedBadge(parent, text=text, bg=bg, fg=fg, font=('Segoe UI', 9), height=22)
        badge.pack(side=tk.LEFT, padx=3)

    def _switch_tab(self, tab_id):
        """
        인터랙티브 탭 전환 핸들러 / Interactive Tab Switch Handler
        """
        self._active_tab = tab_id
        for tid, tab in self._tabs.items():
            if tid == tab_id:
                tab['label'].config(fg="#ffffff", font=('Segoe UI', 10, 'bold'))
                tab['underline'].config(bg=C_ACCENT_BLUE)
            else:
                tab['label'].config(fg=C_TEXT_MUTED, font=('Segoe UI', 10))
                tab['underline'].config(bg=C_BG_SIDEBAR)

        for tid, view in self._tab_views.items():
            if tid == tab_id:
                view.pack(fill=tk.BOTH, expand=True)
            else:
                view.pack_forget()

    def _on_tab_hover(self, tab_id, is_enter):
        if self._active_tab != tab_id:
            color = "#ffffff" if is_enter else C_TEXT_MUTED
            self._tabs[tab_id]['label'].config(fg=color)

    def _init_tab_contents(self):
        """
        각 탭의 초기 안내 정보 초기화 / Initialize Tab Contents
        """
        # OUTPUT 탭 초기 메시지 / Initial OUTPUT Tab Message
        self._log_raw("FortiGate Config -> Excel Converter initialized.\n", 'muted')
        self._log_raw("Select a .conf file and click 'Start Conversion' to begin.\n\n", 'muted')

        # PROBLEMS 탭 초기 메시지 / Initial PROBLEMS Tab Message
        self.problems_text.insert(tk.END, "No problems have been detected in the workspace.\n", 'muted')

        # DEBUG CONSOLE 탭 초기 메시지 / Initial DEBUG CONSOLE Tab Message
        self.debug_text.insert(tk.END, "[Debug Console : Policy Parser Diagnostics]\n", 'info')
        self.debug_text.insert(tk.END, "Ready to capture vDOM breakdown, policy counts, and object mapping metrics.\n", 'muted')

        # TERMINAL 환경 및 진단 정보 / TERMINAL Environment & Diagnostics Information
        import platform
        try:
            import openpyxl
            openpyxl_ver = openpyxl.__version__
        except Exception:
            openpyxl_ver = "Unknown"

        py_ver = platform.python_version()
        sys_os = platform.system() + " " + platform.release()
        script_dir = os.path.dirname(os.path.abspath(__file__))

        term_info = (
            f"FortiGate Policy to Excel - Environment Console\n"
            f"--------------------------------------------------\n"
            f"• Python Version     : {py_ver} ({sys.executable})\n"
            f"• Operating System   : {sys_os} (High-DPI Aware v2)\n"
            f"• OpenPyXL Engine    : {openpyxl_ver}\n"
            f"• Workspace Root     : {script_dir}\n"
            f"\n"
            f"[CLI Execution Syntax]\n"
            f"  python fortigate_policy_to_excel.py <config_path> [output_dir]\n"
            f"  python fortigate_policy_to_excel.py --gui\n"
            f"\n"
            f"[Features Active]\n"
            f"  ✔ Recursive Address / Service Group Resolution\n"
            f"  ✔ 32-Column Firewall Policy Multi-row Flattening\n"
            f"  ✔ Individual vDOM Workbooks + TOTAL_SUMMARY.xlsx\n"
        )
        self.terminal_text.insert(tk.END, term_info, 'muted')

    def _create_button(self, parent, text, bg, hover_bg, cmd, font=('Segoe UI', 10), height=34, width=None):
        """
        타원형 버튼 생성 헬퍼 / Oval Button Helper
        """
        return RoundedButton(
            parent, text=text, bg=bg, hover_bg=hover_bg, cmd=cmd,
            font=font, height=height, width=width
        )

    def _set_btn_state(self, btn, enabled):
        """
        버튼 활성화/비활성화 상태 설정 / Set Button Enabled/Disabled State
        """
        if hasattr(btn, 'set_state'):
            btn.set_state('normal' if enabled else 'disabled')
        else:
            if enabled:
                btn.config(state=tk.NORMAL, bg=getattr(btn, '_orig_bg', C_BTN_PRIMARY), fg="#ffffff", cursor='hand2')
            else:
                btn.config(state=tk.DISABLED, bg=C_BTN_DISABLED, fg=C_BTN_DISABLED_FG, cursor='arrow')

    # -------------------------------------------------------------
    # 이벤트 핸들러 / Event Handlers
    # -------------------------------------------------------------
    def _browse_conf(self):
        fpath = filedialog.askopenfilename(
            title="Select FortiGate Configuration File",
            filetypes=[
                ("FortiGate Config", "*.conf"),
                ("All Files", "*.*")
            ]
        )
        if fpath:
            fpath = os.path.normpath(fpath)
            self.entry_conf.delete(0, tk.END)
            self.entry_conf.insert(0, fpath)

            dir_path = os.path.dirname(fpath)
            self.entry_out.delete(0, tk.END)
            self.entry_out.insert(0, dir_path)

            fname = os.path.basename(fpath)
            self._set_status(f"✔ File loaded: {fname}", 0)

    def _browse_out(self):
        dpath = filedialog.askdirectory(title="Select Output Directory")
        if dpath:
            dpath = os.path.normpath(dpath)
            self.entry_out.delete(0, tk.END)
            self.entry_out.insert(0, dpath)

    def _clear_log(self):
        """
        현재 활성화된 탭의 텍스트 영역 비우기 / Clear Current Tab Log Area
        """
        active_view = self._tab_views.get(self._active_tab)
        if active_view:
            active_view.delete('1.0', tk.END)

    def _log_raw(self, text, tag=None):
        def append():
            if tag:
                self.log_text.insert(tk.END, text, tag)
            else:
                self.log_text.insert(tk.END, text)
            self.log_text.see(tk.END)
        self.root.after(0, append)

    def _log(self, text):
        def append():
            line = text + "\n"
            stripped = text.strip()
            if stripped.startswith("[*]"):
                self.log_text.insert(tk.END, line, 'info')
            elif stripped.startswith("[ERROR]") or "[오류]" in stripped or stripped.startswith("Traceback"):
                self.log_text.insert(tk.END, line, 'error')
                self._record_problem(line)
            elif stripped.startswith("[WARNING]") or "WARNING:" in stripped:
                self.log_text.insert(tk.END, line, 'warning')
            elif stripped.startswith("[DEBUG]"):
                self.log_text.insert(tk.END, line, 'debug')
            elif "File " in stripped and ", line " in stripped:
                self.log_text.insert(tk.END, line, 'trace')
            elif stripped.startswith("[") and ("vDOM" in stripped or "/" in stripped[:6]):
                m_vdom = re.search(r"(Parsing vDOM\s+['\"])(.*?)(['\"])", line)
                if m_vdom:
                    self.log_text.insert(tk.END, line[:m_vdom.start(2)], 'vdom')
                    self.log_text.insert(tk.END, m_vdom.group(2), 'vdom_name')
                    self.log_text.insert(tk.END, line[m_vdom.end(2):], 'vdom')
                else:
                    self.log_text.insert(tk.END, line, 'vdom')
            elif "[SUCCESS]" in stripped or "[SUMMARY COMPLETE]" in stripped or "[생성 완료]" in stripped or "[성공]" in stripped or "[총괄 요약 완료]" in stripped:
                self.log_text.insert(tk.END, line, 'success')
            elif stripped.startswith("-") or stripped.startswith("->"):
                self.log_text.insert(tk.END, line, 'comment')
            else:
                self.log_text.insert(tk.END, line)
            self.log_text.see(tk.END)
        self.root.after(0, append)

    def _record_problem(self, text):
        """
        오류 발생 시 PROBLEMS 탭에 기록 및 카운트 갱신 / Record Issue to PROBLEMS Tab
        """
        if self._problem_count == 0:
            self.problems_text.delete('1.0', tk.END)
        self._problem_count += 1
        self.problems_text.insert(tk.END, f"[{self._problem_count}] {text}", 'error')
        self._tabs["PROBLEMS"]['label'].config(text=f"PROBLEMS  {self._problem_count}")

    def _set_status(self, text, progress_pct=None):
        def update():
            self.status_lbl.config(text=text, anchor='w')
            if progress_pct is not None:
                self.progressbar['value'] = progress_pct
        self.root.after(0, update)

    # -------------------------------------------------------------
    # 비동기 변환 실행 스레드 / Asynchronous Conversion Execution Thread
    # -------------------------------------------------------------
    def _start_conversion(self):
        if self.is_running:
            return

        conf_file = self.entry_conf.get().strip()
        out_dir = self.entry_out.get().strip()

        if not conf_file:
            messagebox.showwarning("Input Required", "Please select a FortiGate configuration file (.conf).")
            return
        if not os.path.isfile(conf_file):
            messagebox.showerror("File Not Found", f"The specified configuration file does not exist:\n{conf_file}")
            return

        if not out_dir:
            out_dir = os.path.dirname(conf_file)
            self.entry_out.insert(0, out_dir)

        self.is_running = True
        self._set_btn_state(self.btn_run, False)
        self._set_btn_state(self.btn_open_res, False)
        self.log_text.delete('1.0', tk.END)
        self.progressbar['value'] = 0

        t = threading.Thread(target=self._run_process, args=(conf_file, out_dir), daemon=True)
        t.start()

    def _run_process(self, conf_file, out_dir):
        try:
            self.last_target_dir = execute_conversion(
                conf_file, out_dir,
                log_fn=self._log,
                status_fn=self._set_status
            )
            self.root.after(0, self._on_success)
        except Exception as ex:
            self._set_status(f"✖ Error occurred: {str(ex)}", 0)
            self._log(f"\n[ERROR] An error occurred during export:\n{str(ex)}")
            self.root.after(0, lambda: messagebox.showerror("Export Error", f"An error occurred during export:\n{str(ex)}"))
        finally:
            self.is_running = False
            self.root.after(0, lambda: self._set_btn_state(self.btn_run, True))

    def _on_success(self):
        self._set_btn_state(self.btn_open_res, True)
        if self.var_open_folder.get() and self.last_target_dir:
            self._open_result_folder()

    def _open_result_folder(self):
        if self.last_target_dir and os.path.isdir(self.last_target_dir):
            if sys.platform == 'win32':
                os.startfile(self.last_target_dir)
            elif sys.platform == 'darwin':
                subprocess.Popen(['open', self.last_target_dir])
            else:
                subprocess.Popen(['xdg-open', self.last_target_dir])


# ================================================================
# 14. 핵심 변환 파이프라인 및 실행 진입점 / Core Conversion Pipeline & Entry Points
# ================================================================

def execute_conversion(config_file, base_dir=None, log_fn=print, status_fn=None):
    """
    FortiGate 설정 파일(.conf)을 분석하여 각 vDOM별 엑셀 파일 및 총괄 요약 파일 생성
    Core conversion pipeline used by both PyWebView GUI, Tkinter GUI, and CLI mode.
    """
    if base_dir is None:
        base_dir = os.path.dirname(config_file) or '.'

    if status_fn:
        status_fn("⟳ Loading configuration file...", 5)
    log_fn(f"[*] Config loading: {config_file}")

    if not os.path.isfile(config_file):
        raise FileNotFoundError(f"Configuration file not found: '{config_file}'")

    if os.path.getsize(config_file) == 0:
        raise ValueError(f"Configuration file is empty (0 bytes): '{config_file}'")

    with open(config_file, 'r', encoding='utf-8', errors='replace') as f:
        lines = f.read().split('\n')

    if not lines or (len(lines) == 1 and not lines[0].strip()):
        raise ValueError("Configuration file contains no readable lines or content.")

    log_fn(f"    - Read {len(lines):,} lines successfully")

    # Check for valid FortiGate configuration syntax
    has_fg_syntax = any(
        l.strip().startswith(('config ', '#config-version=', '#build', 'set ', 'edit '))
        for l in lines[:300]
    )
    if not has_fg_syntax:
        log_fn("[WARNING] The file does not appear to contain standard FortiGate configuration syntax.")
        log_fn("[DEBUG] Expected directives like 'config ...', '#config-version', or 'set ...' were not found in header.")

    # 1. 호스트네임 추출 / 1. Extract Hostname
    hostname = parse_hostname(lines, default_name="FortiGate")
    target_dir = os.path.join(base_dir, hostname)
    os.makedirs(target_dir, exist_ok=True)
    log_fn(f"[*] Hostname: {hostname}")
    log_fn(f"[*] Output Directory: {target_dir}")

    # 2. vDOM 경계 분석 / 2. Analyze vDOM Boundaries
    if status_fn:
        status_fn("⟳ Analyzing vDOM boundaries...", 10)
    log_fn("[*] vDOM parsing...")
    vdom_sections = find_vdom_boundaries(lines)
    num_vdoms = len(vdom_sections)
    log_fn(f"    {num_vdoms} vDOM(s) detected: {', '.join(v[0] for v in vdom_sections)}")

    all_profile_comments = parse_security_profile_comments(lines, 0, len(lines)-1)
    all_ext_resources = parse_external_resources(lines, 0, len(lines)-1)
    all_sched_recur = parse_schedule_recurring(lines, 0, len(lines)-1)
    all_sched_onetime = parse_schedule_onetime(lines, 0, len(lines)-1)
    all_sched_grp = parse_schedule_group(lines, 0, len(lines)-1)
    all_vdom_interfaces = parse_system_interfaces(lines)

    # 3. vDOM별 파싱 및 엑셀 개별 파일 생성 / 3. Parse per vDOM & Generate Individual Excel Files
    summary_list = []
    vdom_obj_counts = {}

    for idx, (vdom_name, vs, ve) in enumerate(vdom_sections, 1):
        pct = 10 + int((idx / max(1, num_vdoms)) * 80)
        if status_fn:
            status_fn(f"⟳ Processing ({idx}/{num_vdoms}): {vdom_name}", pct)
        log_fn(f"\n[{idx}/{num_vdoms}] Parsing vDOM '{vdom_name}' (lines {vs+1:,} ~ {ve+1:,})...")

        vdom_insp_mode = parse_vdom_inspection_mode(lines, vs, ve)
        vdom_ext_res = parse_external_resources(lines, vs, ve)
        merged_ext_res = OrderedDict(all_ext_resources)
        merged_ext_res.update(vdom_ext_res)

        addr_dict = parse_address_objects(lines, vs, ve)
        addrgrp_dict = parse_addrgrp_objects(lines, vs, ve)
        svc_dict = parse_service_objects(lines, vs, ve)
        svcgrp_dict = parse_service_groups(lines, vs, ve)
        ippool_dict = parse_ippool_objects(lines, vs, ve)

        merged_sched_recur = dict(all_sched_recur)
        merged_sched_recur.update(parse_schedule_recurring(lines, vs, ve))
        merged_sched_onetime = dict(all_sched_onetime)
        merged_sched_onetime.update(parse_schedule_onetime(lines, vs, ve))
        merged_sched_grp = dict(all_sched_grp)
        merged_sched_grp.update(parse_schedule_group(lines, vs, ve))

        resolver = ObjectResolver(
            addr_dict, addrgrp_dict, svc_dict, svcgrp_dict, ippool_dict,
            ext_resources=merged_ext_res,
            sched_recurring_dict=merged_sched_recur,
            sched_onetime_dict=merged_sched_onetime,
            sched_group_dict=merged_sched_grp
        )
        obj_counts = [
            len(addr_dict), len(addrgrp_dict), len(svc_dict), len(svcgrp_dict), len(ippool_dict),
            len(merged_sched_recur), len(merged_sched_onetime), len(merged_sched_grp)
        ]
        vdom_obj_counts[vdom_name] = obj_counts

        fw_pols = []
        for sr, er in find_section_range(lines, vs, ve, "firewall policy"):
            fw_pols.extend(parse_firewall_policy(lines, sr, er))

        li_pols = []
        for sr, er in find_section_range(lines, vs, ve, "firewall local-in-policy"):
            li_pols.extend(parse_local_in_policy(lines, sr, er))

        cn_ents = []
        for sr, er in find_section_range(lines, vs, ve, "firewall central-snat-map"):
            cn_ents.extend(parse_central_snat(lines, sr, er))

        dn_ents = []
        for sr, er in find_section_range(lines, vs, ve, "firewall vip"):
            dn_ents.extend(parse_vip(lines, sr, er))

        dos_pols = []
        for sr, er in find_section_range(lines, vs, ve, "firewall DoS-policy"):
            dos_pols.extend(parse_dos_policy(lines, sr, er))

        acl_pols = []
        for sr, er in find_section_range(lines, vs, ve, "firewall acl"):
            acl_pols.extend(parse_firewall_acl(lines, sr, er))

        static_routes = parse_router_static(lines, vs, ve)
        policy_routes = parse_router_policy(lines, vs, ve)
        ospf_data = parse_router_ospf(lines, vs, ve)
        vpn_data = parse_ipsec_vpn(lines, vs, ve)
        vdom_interfaces = all_vdom_interfaces.get(vdom_name, [])

        counts = [
            len(fw_pols), len(li_pols), len(cn_ents), len(dn_ents), len(dos_pols), len(acl_pols),
            len(static_routes), len(policy_routes), len(ospf_data.get('networks', [])), len(vpn_data),
            len(vdom_interfaces)
        ]
        summary_list.append((vdom_name, counts))

        clean_vdom_filename = re.sub(r'[\\/*?:"<>|]', "_", vdom_name) + ".xlsx"
        vdom_file_path = os.path.join(target_dir, clean_vdom_filename)
        actual_path = export_single_vdom_excel(vdom_name, fw_pols, li_pols, cn_ents, dn_ents, dos_pols,
                                               resolver, obj_counts, vdom_file_path,
                                               profile_comments=all_profile_comments,
                                               vdom_inspection_mode=vdom_insp_mode,
                                               ext_resources=merged_ext_res,
                                               static_routes=static_routes,
                                               policy_routes=policy_routes,
                                               ospf_data=ospf_data,
                                               vpn_data=vpn_data,
                                               interfaces=vdom_interfaces,
                                               acl_pols=acl_pols,
                                               log_fn=log_fn)
        saved_name = os.path.basename(actual_path)
        log_fn(f"    -> [Saved] {saved_name} (Policy: {counts[0]}, LocalIn: {counts[1]}, CNAT: {counts[2]}, VIP: {counts[3]}, DoS: {counts[4]}, ACL: {counts[5]}, StaticRt: {counts[6]}, PolicyRt: {counts[7]}, OSPF: {counts[8]}, IPsec: {counts[9]}, Intf: {counts[10]})")

    # 4. 전체 vDOM 통합 요약 파일 생성 (_TOTAL_SUMMARY.xlsx) / 4. Generate Total Summary Excel
    if status_fn:
        status_fn("⟳ Building total summary workbook...", 95)
    total_summary_path = os.path.join(target_dir, "_TOTAL_SUMMARY.xlsx")
    wb_tot = Workbook()
    ws_tot = wb_tot.active
    ws_tot.title = "vDOM Total Summary"
    tot_headers = ["vDOM", "Firewall Policy", "Local-in Policy",
                   "Central-NAT", "DNAT (VIP)", "DoS Policy", "ACL Policy",
                   "Static Route", "Policy Route", "OSPF Networks", "IPsec VPN",
                   "Network Interfaces",
                   "Address Objects", "Addr Groups",
                   "Service Objects", "Svc Groups", "IP Pools",
                   "Sched Recurring", "Sched Onetime", "Sched Groups"]
    for col, h in enumerate(tot_headers, 1):
        sc(ws_tot, 1, col, h, font=HDR_FONT, fill=HDR_DEFAULT_FILL, align=CENTER)
    ws_tot.freeze_panes = "A2"

    for ri, (vdom, counts) in enumerate(summary_list, 2):
        fill = EVEN_ROW_FILL if ri % 2 == 0 else ODD_ROW_FILL
        oc = vdom_obj_counts.get(vdom, [0, 0, 0, 0, 0, 0, 0, 0])
        vals = [vdom] + counts + oc
        for col, v in enumerate(vals, 1):
            f = Font(name="맑은 고딕", size=10, bold=(col == 1))
            sc(ws_tot, ri, col, v, font=f, fill=fill, align=CENTER)

    tr = len(summary_list) + 2
    sc(ws_tot, tr, 1, "Total", font=Font(name="맑은 고딕", size=10, bold=True), fill=SUBHDR_FILL, align=CENTER)
    for col in range(2, len(tot_headers) + 1):
        total = sum(
            (summary_list[r][1][col - 2] if col <= 12 else
             vdom_obj_counts.get(summary_list[r][0], [0]*8)[col - 13])
            for r in range(len(summary_list))
        )
        sc(ws_tot, tr, col, total, font=Font(name="맑은 고딕", size=10, bold=True), fill=SUBHDR_FILL, align=CENTER)

    auto_fit(ws_tot, min_w=12)
    saved_summary_path = safe_save_workbook(wb_tot, total_summary_path, log_fn=log_fn)
    log_fn(f"\n[*] [SUMMARY COMPLETE] {saved_summary_path}")

    fmt = "  {:<20} {:>7} {:>7} {:>6} {:>5} {:>4} {:>4} {:>8} {:>8} {:>5} {:>6} {:>6} | {:>5} {:>5} {:>5} {:>5} {:>5}"
    hdr_str = fmt.format("vDOM", "Policy", "LocalIn", "C-NAT", "DNAT", "DoS", "ACL", "StaticRt", "PolicyRt", "OSPF", "IPsec", "Intf",
                         "Addr", "AGrp", "Svc", "SGrp", "Pool")
    sep_w = len(hdr_str)

    log_fn("\n" + "=" * sep_w)
    log_fn(f"  FortiGate [{hostname}] vDOM Export Summary")
    log_fn("=" * sep_w)
    log_fn(hdr_str)
    log_fn("-" * sep_w)
    totals = [0] * 16
    for vdom, counts in summary_list:
        oc = vdom_obj_counts.get(vdom, [0]*5)
        log_fn(fmt.format(vdom, *counts, *oc[:5]))
        for i in range(11):
            totals[i] += counts[i]
        for i in range(5):
            totals[11 + i] += oc[i]
    log_fn("-" * sep_w)
    log_fn(fmt.format("Total", *totals))
    log_fn("=" * sep_w)
    log_fn(f"\n[Finished] All {len(vdom_sections)} vDOM Excel files created in directory: '{target_dir}'")
    if status_fn:
        status_fn("✔ All Excel workbooks generated successfully", 100)

    return target_dir


def run_gui(force_tk=False):
    """
    GUI 실행:
      1. 기본적으로 최신 Glassmorphism 테마의 PyWebView GUI 실행 (첨부 이미지 스타일 1:1 완벽 구현)
      2. pywebview 미설치 또는 force_tk=True인 경우 기존 Tkinter 모던 다크 테마 GUI로 안전하게 폴백
    """
    if not force_tk:
        try:
            import webview_ui
            webview_ui.launch_gui(execute_conversion)
            return
        except Exception as e:
            print(f"[*] Note: Modern PyWebView GUI not started ({e}). Falling back to Tkinter GUI...")

    if not HAS_TKINTER:
        print("[ERROR] Neither PyWebView nor Tkinter is available. GUI cannot start.")
        print("Usage: python fortigate_policy_to_excel.py <config_file> [output_dir]")
        sys.exit(1)

    # Windows 작업표시줄 독립 앱 ID 설정 / Set AppUserModelID for independent taskbar icon
    if sys.platform == 'win32':
        try:
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
                'fortinet.fortigate.policytoexcel.exporter.1.0'
            )
        except Exception:
            pass

    root = tk.Tk()
    app = FortiGateGUI(root)
    root.mainloop()


def run_cli(config_file=None, base_dir=None):
    """커맨드라인(CLI) 모드 실행 / Launch Headless Command-Line Interface (CLI)"""
    if config_file is None:
        if len(sys.argv) < 2:
            print("Usage: python fortigate_policy_to_excel.py <config_file> [output_dir]")
            sys.exit(1)
        config_file = sys.argv[1]
    if base_dir is None:
        base_dir = sys.argv[2] if len(sys.argv) >= 3 else os.path.dirname(config_file) or '.'

    execute_conversion(config_file, base_dir, log_fn=print)


def main():
    # 인자가 없거나 --gui 플래그인 경우 최신 Glassmorphism GUI 모드로 실행
    if len(sys.argv) < 2 or (len(sys.argv) >= 2 and sys.argv[1] in ('--gui', '-g')):
        run_gui(force_tk=False)
    elif len(sys.argv) >= 2 and sys.argv[1] in ('--tk', '--tkinter'):
        run_gui(force_tk=True)
    elif len(sys.argv) >= 2 and sys.argv[1] in ('--help', '-h', '/?'):
        print("FortiGate Policy to Excel Exporter v2.1")
        print("Usage:")
        print("  Modern GUI Mode  : python fortigate_policy_to_excel.py [--gui]")
        print("  Classic GUI Mode : python fortigate_policy_to_excel.py --tk")
        print("  CLI Mode         : python fortigate_policy_to_excel.py <config_file> [output_dir]")
        sys.exit(0)
    else:
        run_cli()


if __name__ == "__main__":
    main()

