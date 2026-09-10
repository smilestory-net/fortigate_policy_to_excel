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


# ================================================================
#  4. 객체 리졸버 (그룹 확장 + IP/코멘트 매핑) / Object Resolver (Recursive Group & Comment Resolution)
# ================================================================

class ObjectResolver:
    def __init__(self, addr_dict, addrgrp_dict, svc_dict, svcgrp_dict, ippool_dict):
        self.addrs = addr_dict
        self.addrgrps = addrgrp_dict
        self.svcs = svc_dict
        self.svcgrps = svcgrp_dict
        self.ippools = ippool_dict

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
                             'poolname', 'internet-service-name'):
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
            e['nat'] = ''
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
        cell.alignment = Alignment(horizontal=cur_h, vertical='center', wrap_text=True)


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

def write_fw_policy_sheet(ws, policies, resolver, vdom_name=""):
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
        # 기타 필드 열 (23~32) / Miscellaneous columns (23-32)
        ("Schedule", "base"), ("NAT", "base"), ("IP Pool", "base"),
        ("Pool Name", "base"), ("Pool IP", "base"),
        ("UTM Status", "base"), ("SSL/SSH Profile", "base"), ("IPS Sensor", "base"),
        ("Log Traffic", "base"), ("Comments", "base")
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
        for addr_name in (p.get('srcaddr') or ['all']):
            src_expanded.extend(resolver.resolve_address(addr_name))

        dst_expanded = []
        for addr_name in (p.get('dstaddr') or ['all']):
            dst_expanded.extend(resolver.resolve_address(addr_name))

        svc_expanded = []
        for svc_name in (p.get('service') or []):
            svc_expanded.extend(resolver.resolve_service(svc_name))
        if not svc_expanded and p.get('internet-service-name'):
            for isn in p['internet-service-name']:
                svc_expanded.append((None, isn, 'internet-service', ''))

        pool_items = []
        for pn in (p.get('poolname') or []):
            pool_items.append(resolver.resolve_ippool(pn))

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

            # 기타 필드 열 (23~32) / Miscellaneous columns (23-32)
            if ri == 0:
                sc(ws, cur_r, 23, p.get('schedule', ''), font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 24, p.get('nat', ''), font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 25, p.get('ippool', ''), font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 28, p.get('utm-status', ''), font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 29, p.get('ssl-ssh-profile', ''), font=b_font, fill=b_fill)
                sc(ws, cur_r, 30, p.get('ips-sensor', ''), font=b_font, fill=b_fill)
                sc(ws, cur_r, 31, p.get('logtraffic', ''), font=b_font, fill=b_fill, align=CENTER)
                sc(ws, cur_r, 32, p.get('comments', ''), font=b_font, fill=b_fill)
            else:
                for c in [23, 24, 25, 28, 29, 30, 31, 32]:
                    sc(ws, cur_r, c, None, font=b_font, fill=b_fill)

            if ri < len(pool_items):
                pname, pip, ptype = pool_items[ri]
                sc(ws, cur_r, 26, pname, font=b_font, fill=b_fill)
                sc(ws, cur_r, 27, pip, font=b_font, fill=b_fill)
            else:
                sc(ws, cur_r, 26, None, font=b_font, fill=b_fill)
                sc(ws, cur_r, 27, None, font=b_font, fill=b_fill)

        if p_end > p_start:
            common_cols = [1, 2, 3, 4, 5, 6, 7, 13, 23, 24, 25, 28, 29, 30, 31, 32]
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
                sc(ws, cur_r, 20, e.get('nat', ''), font=b_font, fill=b_fill, align=CENTER)
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


# ================================================================
# 12. 개별 VDOM 엑셀 생성 함수 / Per-VDOM Excel Generation Function
# ================================================================

def export_single_vdom_excel(vdom_name, fw_pols, li_pols, cn_ents, dn_ents, dos_pols, resolver, obj_counts, filepath):
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
    for idx, (label, cnt) in enumerate(items, 2):
        fill = EVEN_ROW_FILL if idx % 2 == 0 else ODD_ROW_FILL
        sc(ws_sum, idx, 1, label, font=FONT_DEFAULT_BOLD if idx == 2 else FONT_DEFAULT, fill=fill)
        sc(ws_sum, idx, 2, cnt, font=FONT_DEFAULT_BOLD if idx == 2 else FONT_DEFAULT, fill=fill, align=CENTER)
    auto_fit(ws_sum, min_w=15)

    # 2. 방화벽 정책 시트 / Firewall Policy Sheet
    ws_fw = wb.create_sheet("Firewall Policy")
    write_fw_policy_sheet(ws_fw, fw_pols, resolver, vdom_name)

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
C_ACCENT_BLUE = "#007acc"     # 액센트 블루 / Accent Blue
C_BTN_PRIMARY = "#0e639c"     # 기본 버튼 (블루) / Primary Button
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


class FortiGateGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("FortiGate Policy to Excel Exporter v1.0")
        self.root.geometry("1000x720")
        self.root.minsize(860, 480)
        self.root.configure(bg=C_BG_APP)

        self.last_target_dir = ""
        self.is_running = False

        self._setup_styles()
        self._create_widgets()

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
            text="FORTIGATE CONFIG EXPORTER",
            font=('Segoe UI', 13, 'bold'),
            fg="#ffffff",
            bg=C_BG_SIDEBAR
        )
        title_lbl.pack(anchor=tk.W)

        subtitle_lbl = tk.Label(
            title_left,
            text="FortiOS Policy & Object Parser",
            font=('Segoe UI', 10),
            fg=C_TEXT_MUTED,
            bg=C_BG_SIDEBAR
        )
        subtitle_lbl.pack(anchor=tk.W, pady=(2, 0))

        # 우측 정보 배지 / Right-side Info Badges
        badges_frame = tk.Frame(header_content, bg=C_BG_SIDEBAR)
        badges_frame.pack(side=tk.RIGHT, anchor=tk.E)

        self._create_badge(badges_frame, "FortiOS 6.x / 7.x", "#264f78", "#9cdcfe")
        self._create_badge(badges_frame, "Multi-VDOM", "#37373d", "#cccccc")

        # -------------------------------------------------------------
        # 2. 하단 상태 표시줄 / 2. Bottom Status Bar (Always Pinned to Bottom)
        # -------------------------------------------------------------
        status_bar = tk.Frame(self.root, bg=C_STATUSBAR_BG, height=28)
        status_bar.pack(fill=tk.X, side=tk.BOTTOM)

        self.status_lbl = tk.Label(
            status_bar,
            text="● 준비 완료 (대기 중)",
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
            text="EXPLORER : CONFIGURATION & PATHS",
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
            card_body, text="Config File (.conf):",
            font=('Segoe UI', 10), fg=C_TEXT_MAIN, bg=C_BG_SIDEBAR, width=17, anchor=tk.W
        ).grid(row=0, column=0, sticky=tk.W, pady=4)

        self.entry_conf = tk.Entry(
            card_body,
            font=('Consolas', 10),
            bg=C_INPUT_BG,
            fg=C_INPUT_FG,
            insertbackground="#ffffff",
            relief=tk.FLAT,
            bd=0,
            highlightthickness=1,
            highlightbackground=C_BORDER_LIGHT,
            highlightcolor=C_ACCENT_BLUE
        )
        self.entry_conf.grid(row=0, column=1, sticky=tk.EW, padx=(6, 8), ipady=5)

        self.btn_browse_conf = self._create_button(
            card_body, "Browse...", C_BTN_SECONDARY, C_BTN_SECONDARY_HOVER, self._browse_conf, font=('Segoe UI', 10), width=10
        )
        self.btn_browse_conf.grid(row=0, column=2, pady=4)

        # 출력 디렉터리 선택 행 / Output Directory Path Row
        tk.Label(
            card_body, text="Output Directory:",
            font=('Segoe UI', 10), fg=C_TEXT_MAIN, bg=C_BG_SIDEBAR, width=17, anchor=tk.W
        ).grid(row=1, column=0, sticky=tk.W, pady=4)

        self.entry_out = tk.Entry(
            card_body,
            font=('Consolas', 10),
            bg=C_INPUT_BG,
            fg=C_INPUT_FG,
            insertbackground="#ffffff",
            relief=tk.FLAT,
            bd=0,
            highlightthickness=1,
            highlightbackground=C_BORDER_LIGHT,
            highlightcolor=C_ACCENT_BLUE
        )
        self.entry_out.grid(row=1, column=1, sticky=tk.EW, padx=(6, 8), ipady=5)

        self.btn_browse_out = self._create_button(
            card_body, "Browse...", C_BTN_SECONDARY, C_BTN_SECONDARY_HOVER, self._browse_out, font=('Segoe UI', 10), width=10
        )
        self.btn_browse_out.grid(row=1, column=2, pady=4)

        card_body.columnconfigure(1, weight=1)

        # 옵션 체크박스 / Option Checkboxes
        opt_frame = tk.Frame(card_body, bg=C_BG_SIDEBAR)
        opt_frame.grid(row=2, column=1, columnspan=2, sticky=tk.W, pady=(8, 2))

        self.var_open_folder = tk.BooleanVar(value=True)
        chk_open = tk.Checkbutton(
            opt_frame,
            text=" 변환 완료 후 결과 폴더 자동으로 열기 (Auto-open result directory)",
            variable=self.var_open_folder,
            bg=C_BG_SIDEBAR,
            fg=C_TEXT_MAIN,
            selectcolor=C_INPUT_BG,
            activebackground=C_BG_SIDEBAR,
            activeforeground="#ffffff",
            highlightthickness=0,
            bd=0,
            font=('Segoe UI', 10)
        )
        chk_open.pack(side=tk.LEFT)

        # -------------------------------------------------------------
        # 4. 액션 바 (변환 실행 및 결과 폴더 열기) / 4. Action Bar (Execution & Results)
        # -------------------------------------------------------------
        action_frame = tk.Frame(main_container, bg=C_BG_APP)
        action_frame.pack(fill=tk.X, pady=(2, 10))

        self.btn_run = self._create_button(
            action_frame,
            "▶  엑셀 변환 실행 (Start Conversion)",
            C_BTN_PRIMARY,
            C_BTN_PRIMARY_HOVER,
            self._start_conversion,
            font=('Segoe UI', 11, 'bold'),
            pad=(16, 9)
        )
        self.btn_run.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.btn_open_res = self._create_button(
            action_frame,
            "📁 결과 폴더 열기 (Open Folder)",
            C_BTN_SECONDARY,
            C_BTN_SECONDARY_HOVER,
            self._open_result_folder,
            font=('Segoe UI', 10),
            pad=(16, 9)
        )
        self.btn_open_res.pack(side=tk.RIGHT, padx=(10, 0))
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
        # 5. 터미널 콘솔 패널 (통합 터미널 스타일) / 5. Terminal Console Panel (Integrated Terminal Style)
        # -------------------------------------------------------------
        terminal_panel = tk.Frame(
            main_container, bg=C_BG_PANEL, bd=0, highlightthickness=1, highlightbackground=C_BORDER
        )
        terminal_panel.pack(fill=tk.BOTH, expand=True)

        # 탭 헤더 바 / Tab Header Bar
        tab_bar = tk.Frame(terminal_panel, bg=C_BG_SIDEBAR, height=30)
        tab_bar.pack(fill=tk.X, side=tk.TOP)

        # 탭 라벨 목록 / Tab Labels
        self._create_tab(tab_bar, "PROBLEMS  0", active=False)
        self._create_tab(tab_bar, "OUTPUT", active=True)
        self._create_tab(tab_bar, "DEBUG CONSOLE", active=False)
        self._create_tab(tab_bar, "TERMINAL", active=False)

        # 우측 로그 지우기 버튼 / Right-side Clear Button
        btn_clear = tk.Button(
            tab_bar,
            text="⊘ Clear",
            bg=C_BG_SIDEBAR,
            fg=C_TEXT_MUTED,
            activebackground=C_BG_SIDEBAR,
            activeforeground="#ffffff",
            bd=0,
            relief=tk.FLAT,
            font=('Segoe UI', 10),
            cursor='hand2',
            command=self._clear_log
        )
        btn_clear.pack(side=tk.RIGHT, padx=10)

        # 로그 텍스트 영역 / Log Text Area
        self.log_text = scrolledtext.ScrolledText(
            terminal_panel,
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
        self.log_text.pack(fill=tk.BOTH, expand=True)

        # 태그 색상 설정 (구문 강조) / Tag Color Configurations (Syntax Highlighting)
        self.log_text.tag_config('info', foreground=C_TAG_INFO)
        self.log_text.tag_config('vdom', foreground=C_TAG_VDOM, font=('Consolas', 10, 'bold'))
        self.log_text.tag_config('success', foreground=C_TAG_SUCCESS, font=('Consolas', 10, 'bold'))
        self.log_text.tag_config('error', foreground=C_TAG_ERROR, font=('Consolas', 10, 'bold'))
        self.log_text.tag_config('comment', foreground=C_TAG_COMMENT)
        self.log_text.tag_config('muted', foreground=C_TEXT_MUTED)

        # 초기 환영 메시지 / Initial Welcome Message
        self._log_raw("FortiGate Config -> Excel Converter initialized.\n", 'muted')
        self._log_raw("Select a .conf file and click 'Start Conversion' to begin.\n\n", 'muted')

    # -------------------------------------------------------------
    # 헬퍼 메서드: 위젯 커스텀 생성 / Helper Methods: Custom Widget Creation
    # -------------------------------------------------------------
    def _create_badge(self, parent, text, bg, fg):
        lbl = tk.Label(
            parent, text=text, font=('Segoe UI', 10),
            bg=bg, fg=fg, padx=7, pady=3
        )
        lbl.pack(side=tk.LEFT, padx=3)

    def _create_tab(self, parent, text, active=False):
        frame = tk.Frame(parent, bg=C_BG_SIDEBAR)
        frame.pack(side=tk.LEFT, padx=8, fill=tk.Y)

        fg = "#ffffff" if active else C_TEXT_MUTED
        lbl = tk.Label(frame, text=text, font=('Segoe UI', 10, 'bold' if active else 'normal'),
                       fg=fg, bg=C_BG_SIDEBAR, pady=4)
        lbl.pack()

        if active:
            underline = tk.Frame(frame, bg=C_ACCENT_BLUE, height=2)
            underline.pack(fill=tk.X, side=tk.BOTTOM)

    def _create_button(self, parent, text, bg, hover_bg, cmd, font=('Segoe UI', 10), pad=(10, 4), width=None):
        btn = tk.Button(
            parent,
            text=text,
            bg=bg,
            fg="#ffffff",
            activebackground=hover_bg,
            activeforeground="#ffffff",
            relief=tk.FLAT,
            bd=0,
            font=font,
            cursor='hand2',
            padx=pad[0],
            pady=pad[1],
            command=cmd
        )
        if width:
            btn.config(width=width)

        btn._orig_bg = bg
        btn._hover_bg = hover_bg

        def on_enter(e):
            if str(btn['state']) != 'disabled':
                btn.config(bg=btn._hover_bg)

        def on_leave(e):
            if str(btn['state']) != 'disabled':
                btn.config(bg=btn._orig_bg)

        btn.bind("<Enter>", on_enter)
        btn.bind("<Leave>", on_leave)
        return btn

    def _set_btn_state(self, btn, enabled):
        if enabled:
            btn.config(state=tk.NORMAL, bg=btn._orig_bg, fg="#ffffff", cursor='hand2')
        else:
            btn.config(state=tk.DISABLED, bg=C_BTN_DISABLED, fg=C_BTN_DISABLED_FG, cursor='arrow')

    # -------------------------------------------------------------
    # 이벤트 핸들러 / Event Handlers
    # -------------------------------------------------------------
    def _browse_conf(self):
        fpath = filedialog.askopenfilename(
            title="FortiGate 설정 파일 선택",
            filetypes=[
                ("FortiGate Config", "*.conf;*.txt;*.cfg"),
                ("모든 파일", "*.*")
            ]
        )
        if fpath:
            fpath = os.path.normpath(fpath)
            self.entry_conf.delete(0, tk.END)
            self.entry_conf.insert(0, fpath)

            if not self.entry_out.get().strip():
                dir_path = os.path.dirname(fpath)
                self.entry_out.delete(0, tk.END)
                self.entry_out.insert(0, dir_path)

            fname = os.path.basename(fpath)
            self._set_status(f"✔ 파일 로드됨: {fname}", 0)

    def _browse_out(self):
        dpath = filedialog.askdirectory(title="출력 폴더 선택")
        if dpath:
            dpath = os.path.normpath(dpath)
            self.entry_out.delete(0, tk.END)
            self.entry_out.insert(0, dpath)

    def _clear_log(self):
        self.log_text.delete('1.0', tk.END)

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
            elif "[생성 완료]" in stripped or "[성공]" in stripped or "[총괄 요약 완료]" in stripped:
                self.log_text.insert(tk.END, line, 'success')
            elif "[ERROR]" in stripped or "[오류]" in stripped:
                self.log_text.insert(tk.END, line, 'error')
            elif stripped.startswith("-") or stripped.startswith("->"):
                self.log_text.insert(tk.END, line, 'comment')
            else:
                self.log_text.insert(tk.END, line)
            self.log_text.see(tk.END)
        self.root.after(0, append)

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
            messagebox.showwarning("입력 필요", "FortiGate Config 파일(.conf)을 선택해주세요.")
            return
        if not os.path.isfile(conf_file):
            messagebox.showerror("오류", f"지정된 파일이 존재하지 않습니다:\n{conf_file}")
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
            self._set_status("⟳ Config 파일 로딩 중...", 5)
            self._log(f"[*] Config 로드 시작: {conf_file}")

            with open(conf_file, 'r', encoding='utf-8', errors='replace') as f:
                lines = f.read().split('\n')
            self._log(f"    - 총 {len(lines):,} 라인 읽기 완료")

            # 1. 호스트네임 추출 / 1. Extract Hostname
            hostname = parse_hostname(lines, default_name="FortiGate")
            target_dir = os.path.join(out_dir, hostname)
            os.makedirs(target_dir, exist_ok=True)
            self.last_target_dir = target_dir
            self._log(f"[*] 호스트네임: {hostname}")
            self._log(f"[*] 출력 디렉토리: {target_dir}")

            # 2. VDOM 경계 분석 / 2. Analyze VDOM Boundaries
            self._set_status("⟳ VDOM 경계 분석 중...", 10)
            vdom_sections = find_vdom_boundaries(lines)
            num_vdoms = len(vdom_sections)
            self._log(f"[*] {num_vdoms}개 VDOM 감지: {', '.join(v[0] for v in vdom_sections)}")

            # 3. VDOM별 파싱 및 엑셀 개별 파일 생성 / 3. Parse per VDOM & Generate Individual Excel Files
            summary_list = []
            vdom_obj_counts = {}

            for idx, (vdom_name, vs, ve) in enumerate(vdom_sections, 1):
                pct = 10 + int((idx / num_vdoms) * 80)
                self._set_status(f"⟳ 처리 중 ({idx}/{num_vdoms}): {vdom_name}", pct)
                self._log(f"\n[{idx}/{num_vdoms}] VDOM '{vdom_name}' ({vs+1:,} ~ {ve+1:,} 라인) 분석 중...")

                addr_dict = parse_address_objects(lines, vs, ve)
                addrgrp_dict = parse_addrgrp_objects(lines, vs, ve)
                svc_dict = parse_service_objects(lines, vs, ve)
                svcgrp_dict = parse_service_groups(lines, vs, ve)
                ippool_dict = parse_ippool_objects(lines, vs, ve)

                resolver = ObjectResolver(addr_dict, addrgrp_dict, svc_dict, svcgrp_dict, ippool_dict)
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
                                                resolver, obj_counts, vdom_file_path)
                self._log(f"    -> [생성 완료] {clean_vdom_filename} (Policy: {counts[0]}, LocalIn: {counts[1]}, CNAT: {counts[2]}, VIP: {counts[3]}, DoS: {counts[4]})")

            # 4. 전체 요약 엑셀 생성 / 4. Generate Total Summary Excel
            self._set_status("⟳ 총괄 요약 파일 생성 중...", 95)
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
            self._log(f"\n[*] [총괄 요약 완료] _TOTAL_SUMMARY.xlsx")

            self._set_status("✔ 모든 엑셀 변환 완료", 100)
            self._log(f"\n[성공] 총 {num_vdoms}개 VDOM 엑셀 파일이 '{target_dir}' 폴더에 생성되었습니다.")

            self.root.after(0, self._on_success)

        except Exception as ex:
            self._set_status(f"✖ 오류 발생: {str(ex)}", 0)
            self._log(f"\n[ERROR] 변환 중 오류 발생:\n{str(ex)}")
            self.root.after(0, lambda: messagebox.showerror("변환 오류", f"변환 중 오류가 발생했습니다:\n{str(ex)}"))
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
        print("[ERROR] Tkinter 모듈을 불러올 수 없어 GUI를 실행할 수 없습니다.")
        print("Usage: python fortigate_policy_to_excel.py <config_file> [output_dir]")
        sys.exit(1)
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

    summary_list = []
    vdom_obj_counts = {}

    for vdom_name, vs, ve in vdom_sections:
        print(f"\n[*] Processing VDOM '{vdom_name}' ({vs+1}~{ve+1})...")

        addr_dict = parse_address_objects(lines, vs, ve)
        addrgrp_dict = parse_addrgrp_objects(lines, vs, ve)
        svc_dict = parse_service_objects(lines, vs, ve)
        svcgrp_dict = parse_service_groups(lines, vs, ve)
        ippool_dict = parse_ippool_objects(lines, vs, ve)

        resolver = ObjectResolver(addr_dict, addrgrp_dict, svc_dict, svcgrp_dict, ippool_dict)
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
                                 resolver, obj_counts, vdom_file_path)
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
    else:
        run_cli()


if __name__ == "__main__":
    main()
