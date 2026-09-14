#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
FortiGate Configuration -> Excel Exporter & GUI (Unified Single-File Edition)
=============================================================================
[한국어]
FortiGate 방화벽 설정 파일(.conf)을 분석하여:
  1. 호스트네임 디렉토리 생성 및 각 VDOM별 엑셀 파일(<vdom_name>.xlsx) 분할 생성
  2. 전체 VDOM 총괄 요약 파일(_TOTAL_SUMMARY.xlsx) 동시 생성
  3. 객체/그룹의 실제 IP, 서브넷, 포트, 코멘트를 다중 행 전개 및 스마트 셀 세로 병합(Merge)
  4. Svc Protocol 컬럼 삭제, Svc Port 'ALL' 표기, Src/Dst/Svc Comment 분리 수록
  5. 출발지(파랑), 목적지(빨강) 가독성 컬러 스타일링 및 비활성화 정책(진한 회색) 음영 처리
  6. 모던 다크 테마 GUI 및 커맨드라인(CLI) 모드 완벽 통합 지원

[English]
Parses FortiGate firewall backup configuration files (.conf / .txt) to:
  1. Create a hostname-based directory with partitioned Excel files per VDOM (<vdom_name>.xlsx)
  2. Simultaneously generate a master summary workbook (_TOTAL_SUMMARY.xlsx) across all VDOMs
  3. Recursively resolve objects/groups to actual IPs/ports/comments with multi-row flattening & cell merging
  4. Optimize service ports ('ALL') and provide dedicated Src/Dst/Svc Comment columns
  5. Apply professional visual styling (Src blue, Dst red, zebra striping, disabled policy shading)
  6. Provide an integrated modern Dark Theme GUI and headless CLI execution in a single file

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
except ImportError:
    print("openpyxl 필요: pip install openpyxl")
    sys.exit(1)


# ================================================================
#  1. 유틸리티 / Utilities
# ================================================================

def parse_quoted_values(line):
    """'set field "val1" "val2" ...' -> ["val1", "val2"]"""
    parts = line.strip().split(None, 2)
    if len(parts) < 3:
        return []
    rest = parts[2]
    quoted = re.findall(r'"([^"]*)"', rest)
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
    [한국어] 객체 타입별 IP 문자열 단일 포맷팅:
      - 단일 호스트: 192.168.1.1/32
      - 네트워크 서브넷 대역: 192.168.10.0/24, 0.0.0.0/0 등
      - IP 범위 (Range): 1.1.1.1-1.1.1.10
      - FQDN, 지리(국가), 동적 객체 등

    [English] Unify IP display format based on object type:
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
#  2. VDOM 경계 파싱 / VDOM Boundary Parsing
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
    while i <= end:
        if lines[i].strip() == f"config {section_name}":
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
    VDOM 설정의 inspection-mode (flow 또는 proxy) 파싱
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
    def __init__(self, addr_dict, addrgrp_dict, svc_dict, svcgrp_dict, ippool_dict, ext_resources=None):
        self.addrs = addr_dict
        self.addrgrps = addrgrp_dict
        self.svcs = svc_dict
        self.svcgrps = svcgrp_dict
        self.ippools = ippool_dict
        self.ext_resources = ext_resources or {}

    def resolve_address(self, name, _visited=None):
        """
        [한국어] 주소 객체 또는 그룹을 재귀적으로 확장하여 멤버별 세부 정보 반환
        [English] Recursively resolve address object or group into individual member entries
        Returns: list of (group_name, obj_name, obj_type, formatted_ip, comment)
        """
        if _visited is None:
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
            return result

        if name in self.addrs:
            obj = self.addrs[name]
            f_ip = format_ip_str(obj)
            return [(None, name, obj.get('type', 'ipmask'), f_ip, obj.get('comment', ''))]

        if name in self.ext_resources:
            ext = self.ext_resources[name]
            ext_type = ext.get('type', 'external-resource')
            return [(None, name, ext_type, ext.get('resource', ''), ext.get('comments', ''))]

        return [(None, name, 'unknown', '', '')]

    def resolve_service(self, name, _visited=None):
        """
        [한국어] 서비스 객체 또는 그룹을 재귀적으로 확장하여 포트 및 코멘트 정보 반환
        [English] Recursively resolve service object or group into individual service/port entries
        Returns: list of (group_name, svc_name, port_display, comment)
        """
        if _visited is None:
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
            return result if result else [(None, name, '', grp_comment)]

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
            return [(None, name, port_display, obj.get('comment', ''))]

        return [(None, name, '', '')]

    def resolve_ippool(self, name):
        """IP Pool 이름 -> (pool_name, pool_ip_display, pool_type) / Resolve IP pool name to display tuple"""
        if name in self.ippools:
            obj = self.ippools[name]
            return (name, obj.get('display', ''), obj.get('type', ''))
        return (name, '', '')


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
                                   'status': '', 'log': '', 'action': '', 'threshold': ''}
                    elif s.startswith("set ") and anomaly:
                        anomaly[get_field_name(lines[i])] = parse_set_value(lines[i])
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


# ================================================================
#  6. Excel 스타일 & 셀 병합 유틸리티 / Excel Styling & Cell Formatting Utilities
# ================================================================

# --- 컬럼 헤더 배색 / Column Header Fills ---
HDR_DEFAULT_FILL = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid") # 기본 진한 남색 / Default Dark Navy
HDR_SRC_FILL     = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid") # 진한 파랑 (출발지) / Deep Blue (Source)
HDR_DST_FILL     = PatternFill(start_color="843C39", end_color="843C39", fill_type="solid") # 진한 빨강 (목적지) / Deep Red (Destination)
HDR_SVC_FILL     = PatternFill(start_color="415A77", end_color="415A77", fill_type="solid") # 블루그레이 (서비스) / Blue-Gray (Service)
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

# 비활성화 행 폰트 / Disabled Row Font Colors
FONT_DIS          = Font(name="맑은 고딕", size=9, color="495057")
FONT_DIS_SRC      = Font(name="맑은 고딕", size=9, color="1B365D")
FONT_DIS_DST      = Font(name="맑은 고딕", size=9, color="6B1D1D")

# Action 및 VDOM 폰트 / Action (Accept / Deny) & VDOM Fonts
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
    [한국어] 행 전체 열에 일관된 배경색(Zebra 격행 / 비활성화 진한 회색)을 적용하고, 영역별(출발지/목적지/액션) 글씨색을 지정
    [English] Apply consistent background fill (Zebra striping / Disabled dark gray) across the entire row and assign specific font colors
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
        else:
            font = FONT_DIS
    else:
        if col_category == 'src':
            font = FONT_SRC_GRP if is_group else FONT_SRC
        elif col_category == 'dst':
            font = FONT_DST_GRP if is_group else FONT_DST
        else:
            font = FONT_DEFAULT_BOLD if is_group else FONT_DEFAULT

    return fill, font


def sc(ws, r, c, val, font=FONT_DEFAULT, fill=None, align=WRAP):
    """Set cell with styling."""
    cell = ws.cell(row=r, column=c, value=val)
    cell.font = font
    cell.alignment = align
    cell.border = THIN_BORDER
    if fill:
        cell.fill = fill
    return cell


def write_styled_header(ws, row, headers_with_cat):
    for col, (h, cat) in enumerate(headers_with_cat, 1):
        if cat == 'src':
            fill = HDR_SRC_FILL
        elif cat == 'dst':
            fill = HDR_DST_FILL
        elif cat == 'svc':
            fill = HDR_SVC_FILL
        else:
            fill = HDR_DEFAULT_FILL
        sc(ws, row, col, h, font=HDR_FONT, fill=fill, align=CENTER)
    ws.freeze_panes = ws.cell(row=row + 1, column=1).coordinate


def merge_row_range(ws, start_row, end_row, col, h_align=None):
    if end_row > start_row:
        ws.merge_cells(start_row=start_row, start_column=col,
                       end_row=end_row, end_column=col)
        cell = ws.cell(row=start_row, column=col)
        cur_h = h_align or (cell.alignment.horizontal if cell.alignment else 'center')
        cell.alignment = Alignment(horizontal=cur_h, vertical='top', wrap_text=True)


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
    for col_cells in ws.columns:
        mx = 0
        cl = get_column_letter(col_cells[0].column)
        for cell in col_cells:
            if cell.value:
                for line in str(cell.value).split('\n'):
                    mx = max(mx, len(line))
        ws.column_dimensions[cl].width = min(max(mx + 2, min_w), max_w)


# ================================================================
#  7. 시트 작성 - Firewall Policy / Sheet Builder - Firewall Policy
# ================================================================

def write_fw_policy_sheet(ws, policies, resolver, vdom_name="", dn_ents=None, profile_comments=None, vdom_inspection_mode="flow"):
    headers_with_cat = [
        ("Seq", "base"), ("VDOM", "base"), ("Enable", "base"),
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
        # 기타 필드 열 (23~33) / Miscellaneous columns (23-33)
        ("Schedule", "base"), ("NAT", "base"), ("IP Pool", "base"),
        ("Pool Name", "base"), ("Pool IP", "base"),
        ("Inspection Mode", "base"), ("UTM Status", "base"),
        ("Sec Profile", "base"), ("Sec Profile Comment", "base"),
        ("Log Traffic", "base"), ("Policy Comment", "base")
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
                       len(svc_expanded), len(pool_items), 1)
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

            # 기타 필드 열 (23~33) / Miscellaneous columns (23-33)
            if ri == 0:
                utm_val = 'enable' if p.get('utm-status') == 'enable' else 'disable'
                sc(ws, cur_r, 23, p.get('schedule', ''), font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 24, p.get('nat', ''), font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 25, p.get('ippool', ''), font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 28, insp_mode_display, font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 29, utm_val, font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 30, sec_profile_str, font=b_font, fill=b_fill)
                sc(ws, cur_r, 31, sec_profile_comment_str, font=b_font, fill=b_fill)
                sc(ws, cur_r, 32, log_traffic_str, font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 33, p.get('comments', ''), font=b_font, fill=b_fill)
            else:
                for c in [23, 24, 25, 28, 29, 30, 31, 32, 33]:
                    sc(ws, cur_r, c, None, font=b_font, fill=b_fill)

            if ri < len(pool_items):
                pname, pip, ptype = pool_items[ri]
                sc(ws, cur_r, 26, pname, font=b_font, fill=b_fill)
                sc(ws, cur_r, 27, pip, font=b_font, fill=b_fill)
            else:
                sc(ws, cur_r, 26, None, font=b_font, fill=b_fill)
                sc(ws, cur_r, 27, None, font=b_font, fill=b_fill)

        if p_end > p_start:
            common_cols = [1, 2, 3, 4, 5, 6, 7, 13, 23, 24, 25, 28, 29, 30, 31, 32, 33]
            for c in common_cols:
                merge_row_range(ws, p_start, p_end, c)
            if len(pool_items) <= 1:
                merge_row_range(ws, p_start, p_end, 26)
                merge_row_range(ws, p_start, p_end, 27)

            merge_group_spans(ws, p_start, src_expanded, 8, h_align='left')
            merge_group_spans(ws, p_start, dst_expanded, 14, h_align='left')
            merge_group_spans(ws, p_start, svc_expanded, 19, h_align='left')

        row += max_rows

    auto_fit(ws)
    return ws


# ================================================================
#  8. 시트 작성 - Local-in Policy / Sheet Builder - Local-in Policy
# ================================================================

def write_local_in_sheet(ws, policies, resolver, vdom_name=""):
    headers_with_cat = [
        ("Seq", "base"), ("VDOM", "base"), ("Enable", "base"),
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
        ("Schedule", "base"), ("Comments", "base")
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
                sc(ws, cur_r, 5, p.get('intf', ''), font=b_font, fill=b_fill)
                sc(ws, cur_r, 16, action.upper(), font=act_font, fill=act_fill, align=CENTER)
                sc(ws, cur_r, 21, p.get('schedule', ''), font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 22, p.get('comments', ''), font=b_font, fill=b_fill)
            else:
                for c in [1, 2, 3, 4, 5, 16, 21, 22]:
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

        if p_end > p_start:
            common_cols = [1, 2, 3, 4, 5, 16, 21, 22]
            for c in common_cols:
                merge_row_range(ws, p_start, p_end, c)
            merge_group_spans(ws, p_start, src_expanded, 6, h_align='left')
            merge_group_spans(ws, p_start, dst_expanded, 11, h_align='left')
            merge_group_spans(ws, p_start, svc_expanded, 17, h_align='left')

        row += max_rows

    auto_fit(ws)
    return ws


# ================================================================
#  9. 시트 작성 - Central-NAT / Sheet Builder - Central-NAT
# ================================================================

def write_central_nat_sheet(ws, entries, resolver, vdom_name=""):
    headers_with_cat = [
        ("Seq", "base"), ("VDOM", "base"), ("Enable", "base"), ("ID", "base"),
        ("Src Interface", "src"), ("Dst Interface", "dst"),
        # 원래 출발지 열 (7~11) / Original Source columns (7-11)
        ("Orig Group OBJ", "src"), ("Orig OBJ Name", "src"),
        ("Orig Type", "src"), ("Orig IP", "src"), ("Orig Comment", "src"),
        # 목적지 열 (12~16) / Destination columns (12-16)
        ("Dst Group OBJ", "dst"), ("Dst OBJ Name", "dst"),
        ("Dst Type", "dst"), ("Dst IP", "dst"), ("Dst Comment", "dst"),
        # IP 풀 및 기타 열 (17~21) / IP Pool & Other columns (17-21)
        ("NAT IP Pool Name", "base"), ("NAT Pool IP", "base"),
        ("NAT Pool Type", "base"), ("NAT", "base"), ("Comments", "base")
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

def write_dnat_sheet(ws, entries, vdom_name=""):
    headers_with_cat = [
        ("Seq", "base"), ("VDOM", "base"), ("Name", "base"), ("Type", "base"),
        ("External IP", "src"), ("Mapped IP", "dst"), ("External Interface", "src"),
        ("Port Forward", "base"), ("Protocol", "base"),
        ("External Port", "src"), ("Mapped Port", "dst"),
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

        rs_list = e.get('realservers', [])
        max_rows = max(len(rs_list), 1)
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
                sc(ws, cur_r, 12, e.get('server-type', ''), font=b_font, fill=b_fill)
                sc(ws, cur_r, 13, e.get('ldb-method', ''), font=b_font, fill=b_fill)
                sc(ws, cur_r, 14, e.get('monitor', ''), font=b_font, fill=b_fill)
                sc(ws, cur_r, 18, e.get('comment', ''), font=b_font, fill=b_fill)
            else:
                for c in [1, 2, 3, 4, 8, 9, 12, 13, 14, 18]:
                    sc(ws, cur_r, c, None, font=b_font, fill=b_fill)
                for c in [5, 7, 10]:
                    sc(ws, cur_r, c, None, font=s_font, fill=s_fill)
                for c in [6, 11]:
                    sc(ws, cur_r, c, None, font=d_font, fill=d_fill)

            if ri < len(rs_list):
                rs = rs_list[ri]
                sc(ws, cur_r, 15, rs.get('ip', ''), font=d_font, fill=d_fill)
                sc(ws, cur_r, 16, rs.get('port', ''), font=d_font, fill=d_fill, align=CENTER)
                sc(ws, cur_r, 17, rs.get('weight', '1'), font=d_font, fill=d_fill, align=CENTER)
            else:
                for c in [15, 16, 17]:
                    sc(ws, cur_r, c, None, font=d_font, fill=d_fill)

        if p_end > p_start:
            common_cols = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 18]
            for c in common_cols:
                merge_row_range(ws, p_start, p_end, c)

        row += max_rows

    auto_fit(ws)
    return ws


# ================================================================
# 11. 시트 작성 - DoS Policy / Sheet Builder - DoS Policy
# ================================================================

def write_dos_sheet(ws, policies, resolver, vdom_name=""):
    headers_with_cat = [
        ("Seq", "base"), ("VDOM", "base"), ("Enable", "base"),
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
        ("Comments", "base"),
        # 아노말리 변칙 탐지 열 (21~25) / Anomaly detection columns (21-25)
        ("Anomaly Name", "base"), ("Anomaly Status", "base"),
        ("Anomaly Log", "base"), ("Anomaly Action", "base"),
        ("Anomaly Threshold", "base")
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

            # 변칙 탐지 / Anomaly
            if ri < len(anomalies):
                a = anomalies[ri]
                aact = a.get('action', '')
                afont = DENY_FONT if aact == 'block' else b_font
                sc(ws, cur_r, 21, a.get('name', ''), font=b_font, fill=b_fill)
                sc(ws, cur_r, 22, a.get('status', ''), font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 23, a.get('log', ''), font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 24, aact, font=afont, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 25, a.get('threshold', ''), font=b_font, fill=b_fill, align=CENTER)
            else:
                for c in range(21, 26):
                    sc(ws, cur_r, c, None, font=b_font, fill=b_fill)

        if p_end > p_start:
            common_cols = [1, 2, 3, 4, 5, 20]
            for c in common_cols:
                merge_row_range(ws, p_start, p_end, c)
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
        ("Seq", "base"), ("VDOM", "base"), ("Enable", "base"), ("Name", "base"),
        ("Type", "base"), ("Resource URL", "base"),
        ("Refresh Rate (min)", "base"), ("Source IP", "base"),
        ("Comments", "base")
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
# 12. 개별 VDOM 엑셀 생성 함수 / Per-VDOM Excel Generation Function
# ================================================================

def export_single_vdom_excel(vdom_name, fw_pols, li_pols, cn_ents, dn_ents, dos_pols, resolver, obj_counts, filepath, profile_comments=None, vdom_inspection_mode="flow", ext_resources=None):
    wb = Workbook()

    # 1. 요약 시트 / Summary Sheet
    ws_sum = wb.active
    ws_sum.title = "Summary"
    sc(ws_sum, 1, 1, "Item", font=HDR_FONT, fill=HDR_DEFAULT_FILL, align=CENTER)
    sc(ws_sum, 1, 2, "Count", font=HDR_FONT, fill=HDR_DEFAULT_FILL, align=CENTER)

    items = [
        ("VDOM Name", vdom_name),
        ("Firewall Policy", len(fw_pols)),
        ("Local-in Policy", len(li_pols)),
        ("Central-NAT", len(cn_ents)),
        ("DNAT (VIP)", len(dn_ents)),
        ("DoS Policy", len(dos_pols)),
        ("Address Objects", obj_counts[0]),
        ("Address Groups", obj_counts[1]),
        ("Service Objects", obj_counts[2]),
        ("Service Groups", obj_counts[3]),
        ("IP Pools", obj_counts[4]),
    ]
    if ext_resources:
        items.append(("External Resources", len(ext_resources)))

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
    write_dnat_sheet(ws_dn, dn_ents, vdom_name)

    # 6. DoS 정책 시트 / DoS Policy Sheet
    ws_dos = wb.create_sheet("DoS Policy")
    write_dos_sheet(ws_dos, dos_pols, resolver, vdom_name)

    # 7. 외부 리소스 시트 / External Resource Sheet
    if ext_resources:
        ws_ext = wb.create_sheet("External Resource")
        write_external_resource_sheet(ws_ext, ext_resources, vdom_name)

    wb.save(filepath)


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
C_TAG_VDOM = "#dcdcaa"        # VDOM 로그 태그 (노란색) / VDOM Tag (Yellow)
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
        target_path = os.path.join(temp_dir, "_fortinet_embedded_v1.ico")
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
        self.root.title("FortiGate Policy to Excel Exporter  v1.2")
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
            bg=C_STATUSBAR_BG
        )
        self.status_lbl.pack(side=tk.LEFT, padx=10, pady=3)

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
            txt.tag_config('vdom', foreground=C_TAG_VDOM, font=('Consolas', 10, 'bold'))
            txt.tag_config('success', foreground=C_TAG_SUCCESS, font=('Consolas', 10, 'bold'))
            txt.tag_config('error', foreground=C_TAG_ERROR, font=('Consolas', 10, 'bold'))
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
        self.debug_text.insert(tk.END, "Ready to capture VDOM breakdown, policy counts, and object mapping metrics.\n", 'muted')

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
            f"  ✔ Individual VDOM Workbooks + TOTAL_SUMMARY.xlsx\n"
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
            elif stripped.startswith("[") and ("VDOM" in stripped or "/" in stripped[:6]):
                self.log_text.insert(tk.END, line, 'vdom')
            elif "[SUCCESS]" in stripped or "[SUMMARY COMPLETE]" in stripped or "[생성 완료]" in stripped or "[성공]" in stripped or "[총괄 요약 완료]" in stripped:
                self.log_text.insert(tk.END, line, 'success')
            elif "[ERROR]" in stripped or "[오류]" in stripped:
                self.log_text.insert(tk.END, line, 'error')
                self._record_problem(line)
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
            self.status_lbl.config(text=text)
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
            self._set_status("⟳ Loading configuration file...", 5)
            self._log(f"[*] Loading FortiGate configuration: {conf_file}")

            with open(conf_file, 'r', encoding='utf-8', errors='replace') as f:
                lines = f.read().split('\n')
            self._log(f"    - Read {len(lines):,} lines successfully")

            # 1. 호스트네임 추출 / 1. Extract Hostname
            hostname = parse_hostname(lines, default_name="FortiGate")
            target_dir = os.path.join(out_dir, hostname)
            os.makedirs(target_dir, exist_ok=True)
            self.last_target_dir = target_dir
            self._log(f"[*] Hostname: {hostname}")
            self._log(f"[*] Target Directory: {target_dir}")

            # 2. VDOM 경계 분석 / 2. Analyze VDOM Boundaries
            self._set_status("⟳ Analyzing VDOM boundaries...", 10)
            vdom_sections = find_vdom_boundaries(lines)
            num_vdoms = len(vdom_sections)
            self._log(f"[*] {num_vdoms} VDOM(s) detected: {', '.join(v[0] for v in vdom_sections)}")

            all_profile_comments = parse_security_profile_comments(lines, 0, len(lines)-1)
            all_ext_resources = parse_external_resources(lines, 0, len(lines)-1)

            # 3. VDOM별 파싱 및 엑셀 개별 파일 생성 / 3. Parse per VDOM & Generate Individual Excel Files
            summary_list = []
            vdom_obj_counts = {}

            for idx, (vdom_name, vs, ve) in enumerate(vdom_sections, 1):
                pct = 10 + int((idx / num_vdoms) * 80)
                self._set_status(f"⟳ Processing ({idx}/{num_vdoms}): {vdom_name}", pct)
                self._log(f"\n[{idx}/{num_vdoms}] Parsing VDOM '{vdom_name}' (lines {vs+1:,} ~ {ve+1:,})...")

                vdom_insp_mode = parse_vdom_inspection_mode(lines, vs, ve)
                vdom_ext_res = parse_external_resources(lines, vs, ve)
                merged_ext_res = OrderedDict(all_ext_resources)
                merged_ext_res.update(vdom_ext_res)

                addr_dict = parse_address_objects(lines, vs, ve)
                addrgrp_dict = parse_addrgrp_objects(lines, vs, ve)
                svc_dict = parse_service_objects(lines, vs, ve)
                svcgrp_dict = parse_service_groups(lines, vs, ve)
                ippool_dict = parse_ippool_objects(lines, vs, ve)

                resolver = ObjectResolver(addr_dict, addrgrp_dict, svc_dict, svcgrp_dict, ippool_dict, ext_resources=merged_ext_res)
                obj_counts = [len(addr_dict), len(addrgrp_dict), len(svc_dict), len(svcgrp_dict), len(ippool_dict)]
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

                counts = [len(fw_pols), len(li_pols), len(cn_ents), len(dn_ents), len(dos_pols)]
                summary_list.append((vdom_name, counts))

                clean_vdom_filename = re.sub(r'[\\/*?:"<>|]', "_", vdom_name) + ".xlsx"
                vdom_file_path = os.path.join(target_dir, clean_vdom_filename)
                export_single_vdom_excel(vdom_name, fw_pols, li_pols, cn_ents, dn_ents, dos_pols,
                                         resolver, obj_counts, vdom_file_path,
                                         profile_comments=all_profile_comments,
                                         vdom_inspection_mode=vdom_insp_mode,
                                         ext_resources=merged_ext_res)
                self._log(f"    -> [SUCCESS] {clean_vdom_filename} (Policy: {counts[0]}, LocalIn: {counts[1]}, CNAT: {counts[2]}, VIP: {counts[3]}, DoS: {counts[4]})")

            # 4. 전체 요약 엑셀 생성 / 4. Generate Total Summary Excel
            self._set_status("⟳ Building total summary workbook...", 95)
            total_summary_path = os.path.join(target_dir, "_TOTAL_SUMMARY.xlsx")
            wb_tot = Workbook()
            ws_tot = wb_tot.active
            ws_tot.title = "VDOM Total Summary"
            tot_headers = ["VDOM", "Firewall Policy", "Local-in Policy",
                           "Central-NAT", "DNAT (VIP)", "DoS Policy",
                           "Address Objects", "Addr Groups",
                           "Service Objects", "Svc Groups", "IP Pools"]
            for col, h in enumerate(tot_headers, 1):
                sc(ws_tot, 1, col, h, font=HDR_FONT, fill=HDR_DEFAULT_FILL, align=CENTER)
            ws_tot.freeze_panes = "A2"

            for ri, (vdom, counts) in enumerate(summary_list, 2):
                fill = EVEN_ROW_FILL if ri % 2 == 0 else ODD_ROW_FILL
                oc = vdom_obj_counts.get(vdom, [0, 0, 0, 0, 0])
                vals = [vdom] + counts + oc
                for col, v in enumerate(vals, 1):
                    f = Font(name="맑은 고딕", size=10, bold=(col == 1))
                    sc(ws_tot, ri, col, v, font=f, fill=fill, align=CENTER)

            tr = len(summary_list) + 2
            sc(ws_tot, tr, 1, "Total", font=Font(name="맑은 고딕", size=10, bold=True), fill=SUBHDR_FILL, align=CENTER)
            for col in range(2, len(tot_headers) + 1):
                total = sum(
                    (summary_list[r][1][col - 2] if col <= 6 else
                     vdom_obj_counts.get(summary_list[r][0], [0]*5)[col - 7])
                    for r in range(len(summary_list))
                )
                sc(ws_tot, tr, col, total, font=Font(name="맑은 고딕", size=10, bold=True), fill=SUBHDR_FILL, align=CENTER)

            auto_fit(ws_tot, min_w=12)
            wb_tot.save(total_summary_path)
            self._log(f"\n[*] [SUMMARY COMPLETE] _TOTAL_SUMMARY.xlsx")

            self._set_status("✔ All Excel workbooks generated successfully", 100)
            self._log(f"\n[SUCCESS] Total {num_vdoms} VDOM Excel files exported to '{target_dir}'.")

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
# 14. 실행 진입점 (GUI 및 CLI 통합) / Execution Entry Points (Unified GUI & CLI)
# ================================================================

def run_gui():
    """모던 다크 테마 GUI 실행 / Launch Modern Dark Theme GUI"""
    if not HAS_TKINTER:
        print("[ERROR] Tkinter module is not available. GUI cannot start.")
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

    print(f"[*] Config loading: {config_file}")
    with open(config_file, 'r', encoding='utf-8', errors='replace') as f:
        lines = f.read().split('\n')
    print(f"    {len(lines):,} lines loaded")

    hostname = parse_hostname(lines, default_name="FortiGate")
    target_dir = os.path.join(base_dir, hostname)
    os.makedirs(target_dir, exist_ok=True)
    print(f"[*] Hostname: {hostname}")
    print(f"[*] Output Directory: {target_dir}")

    print("[*] VDOM parsing...")
    vdom_sections = find_vdom_boundaries(lines)
    print(f"    {len(vdom_sections)} VDOMs: {', '.join(v[0] for v in vdom_sections)}")

    all_profile_comments = parse_security_profile_comments(lines, 0, len(lines)-1)
    all_ext_resources = parse_external_resources(lines, 0, len(lines)-1)

    summary_list = []
    vdom_obj_counts = {}

    for vdom_name, vs, ve in vdom_sections:
        print(f"\n[*] Processing VDOM '{vdom_name}' ({vs+1}~{ve+1})...")

        vdom_insp_mode = parse_vdom_inspection_mode(lines, vs, ve)
        vdom_ext_res = parse_external_resources(lines, vs, ve)
        merged_ext_res = OrderedDict(all_ext_resources)
        merged_ext_res.update(vdom_ext_res)

        addr_dict = parse_address_objects(lines, vs, ve)
        addrgrp_dict = parse_addrgrp_objects(lines, vs, ve)
        svc_dict = parse_service_objects(lines, vs, ve)
        svcgrp_dict = parse_service_groups(lines, vs, ve)
        ippool_dict = parse_ippool_objects(lines, vs, ve)

        resolver = ObjectResolver(addr_dict, addrgrp_dict, svc_dict, svcgrp_dict, ippool_dict, ext_resources=merged_ext_res)
        obj_counts = [len(addr_dict), len(addrgrp_dict), len(svc_dict), len(svcgrp_dict), len(ippool_dict)]
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

        counts = [len(fw_pols), len(li_pols), len(cn_ents), len(dn_ents), len(dos_pols)]
        summary_list.append((vdom_name, counts))

        clean_vdom_filename = re.sub(r'[\\/*?:"<>|]', "_", vdom_name) + ".xlsx"
        vdom_file_path = os.path.join(target_dir, clean_vdom_filename)
        export_single_vdom_excel(vdom_name, fw_pols, li_pols, cn_ents, dn_ents, dos_pols,
                                 resolver, obj_counts, vdom_file_path,
                                 profile_comments=all_profile_comments,
                                 vdom_inspection_mode=vdom_insp_mode,
                                 ext_resources=merged_ext_res)
        print(f"    -> [Saved] {clean_vdom_filename} (Policy:{counts[0]}, LocalIn:{counts[1]}, CNAT:{counts[2]}, VIP:{counts[3]}, DoS:{counts[4]})")

    # 전체 VDOM 통합 요약 파일 생성 (_TOTAL_SUMMARY.xlsx) / Generate Total Summary Excel Across All VDOMs (_TOTAL_SUMMARY.xlsx)
    total_summary_path = os.path.join(target_dir, "_TOTAL_SUMMARY.xlsx")
    wb_tot = Workbook()
    ws_tot = wb_tot.active
    ws_tot.title = "VDOM Total Summary"
    tot_headers = ["VDOM", "Firewall Policy", "Local-in Policy",
                   "Central-NAT", "DNAT (VIP)", "DoS Policy",
                   "Address Objects", "Addr Groups",
                   "Service Objects", "Svc Groups", "IP Pools"]
    for col, h in enumerate(tot_headers, 1):
        sc(ws_tot, 1, col, h, font=HDR_FONT, fill=HDR_DEFAULT_FILL, align=CENTER)
    ws_tot.freeze_panes = "A2"

    for ri, (vdom, counts) in enumerate(summary_list, 2):
        fill = EVEN_ROW_FILL if ri % 2 == 0 else ODD_ROW_FILL
        oc = vdom_obj_counts.get(vdom, [0, 0, 0, 0, 0])
        vals = [vdom] + counts + oc
        for col, v in enumerate(vals, 1):
            f = Font(name="맑은 고딕", size=10, bold=(col == 1))
            sc(ws_tot, ri, col, v, font=f, fill=fill, align=CENTER)

    tr = len(summary_list) + 2
    sc(ws_tot, tr, 1, "Total", font=Font(name="맑은 고딕", size=10, bold=True), fill=SUBHDR_FILL, align=CENTER)
    for col in range(2, len(tot_headers) + 1):
        total = sum(
            (summary_list[r][1][col - 2] if col <= 6 else
             vdom_obj_counts.get(summary_list[r][0], [0]*5)[col - 7])
            for r in range(len(summary_list))
        )
        sc(ws_tot, tr, col, total, font=Font(name="맑은 고딕", size=10, bold=True), fill=SUBHDR_FILL, align=CENTER)

    auto_fit(ws_tot, min_w=12)
    wb_tot.save(total_summary_path)
    print(f"\n[OK] Total Summary Saved: {total_summary_path}")

    print("\n" + "=" * 80)
    print(f"  FortiGate [{hostname}] VDOM Export Summary")
    print("=" * 80)
    fmt = "  {:<20} {:>8} {:>8} {:>8} {:>8} {:>8}  | {:>5} {:>5} {:>5} {:>5} {:>5}"
    print(fmt.format("VDOM", "Policy", "LocalIn", "C-NAT", "DNAT", "DoS",
                      "Addr", "AGrp", "Svc", "SGrp", "Pool"))
    print("-" * 80)
    totals = [0] * 10
    for vdom, counts in summary_list:
        oc = vdom_obj_counts.get(vdom, [0]*5)
        print(fmt.format(vdom, *counts, *oc))
        for i in range(5):
            totals[i] += counts[i]
            totals[i + 5] += oc[i]
    print("-" * 80)
    print(fmt.format("Total", *totals))
    print("=" * 80)
    print(f"\n[Finished] All {len(vdom_sections)} VDOM Excel files created in directory: '{target_dir}'")


def main():
    # 인자가 없거나 --gui 플래그인 경우 GUI 모드로 실행 / Run in GUI mode if no args or --gui flag is provided
    if len(sys.argv) < 2 or (len(sys.argv) >= 2 and sys.argv[1] in ('--gui', '-g')):
        run_gui()
    elif len(sys.argv) >= 2 and sys.argv[1] in ('--help', '-h', '/?'):
        print("FortiGate Policy to Excel Exporter v1.2")
        print("Usage:")
        print("  GUI Mode : python fortigate_policy_to_excel.py [--gui]")
        print("  CLI Mode : python fortigate_policy_to_excel.py <config_file> [output_dir]")
        sys.exit(0)
    else:
        run_cli()


if __name__ == "__main__":
    main()
