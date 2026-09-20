# FortiGate Policy to Excel Exporter

포티게이트(FortiGate) 방화벽의 백업 설정 파일(`.conf`)을 분석하여, VDOM별 정책과 객체를 가독성 높은 엑셀(`.xlsx`) 문서로 자동 변환·분할 생성해 주는 도구입니다.

<img width="802" height="607" alt="image" src="https://github.com/user-attachments/assets/01b1390f-9840-4ad8-abbe-5a12f589af31" />


---

## 📌 목차
1. [주요 기능](#-주요-기능)
2. [시스템 요구 사양](#-시스템-요구-사양)
3. [파이썬 및 필수 모듈 설치](#-파이썬-및-필수-모듈-설치)
4. [프로그램 실행 방법](#-프로그램-실행-방법)
5. [엑셀 시트 구성 및 디자인 특징](#-엑셀-시트-구성-및-디자인-특징)
6. [폴더 및 파일 구조](#-폴더-및-파일-구조)
7. [자주 묻는 질문 (FAQ & Troubleshooting)](#-자주-묻는-질문-faq--troubleshooting)

---

## 🌟 주요 기능

### 1. 11대 핵심 정책, 라우팅, VPN 및 인터페이스 설정 완벽 추출 (개별 시트 분리 저장)
* **Firewall Policy** : 일반 방화벽 정책 (Action, Schedule, NAT, IP Pool, Inspection Mode, UTM Profile, Log 등 33개 상세 열 수록)
* **Local-in Policy** : 방화벽 장비 자체 접근 제어 정책
* **Central-NAT** : Central SNAT Map (Original IP/Port ↔ Translated IP/Port 매핑, NAT 기본 활성화 상태 표기)
* **DNAT (VIP)** : 가상 IP 및 포트포워딩 매핑 (External IP/Port ↔ Mapped IP/Port)
* **DoS Policy** : 서비스 거부 공격(DoS/DDoS) 방어 정책
* **Static Route** : 정적 라우팅 설정 (Destination 대역 CIDR 자동 변환, Gateway, Interface, Distance, Priority, 특수 플래그, 비활성화 음영 수록)
* **Policy Route** : 정책 기반 라우팅(PBR) (Incoming/Outgoing Interface, Gateway, Action, Protocol, Port Range, 주소 객체 다중 행 전개 및 셀 병합)
* **OSPF** : OSPF 동적 라우팅 설정 (Global Settings/Router ID, config network, config ospf-interface, config redistribute, Route-Map & Access-List 필터 세부 규칙 구조화 섹션 테이블)
* **IPsec VPN** : IPsec VPN Phase 1 및 Phase 2 터널 설정 (P1 vs P2 헤더 색상 구분, 1:N 터널 계층 연계, Proposal/암호화, P2 DH Group 및 기본값 '14 5' 자동 보정, 로컬/원격 서브넷 기본값 '0.0.0.0/0' 자동 보정, 병합 셀 중앙 정렬 수록)
* **Network Interface** : VDOM별 물리/VLAN/루프백/애그리게이션 인터페이스 설정 (`config system interface` 분석, Status UP/DOWN 상태, Name, Alias, Type, IP/Subnet, VLAN ID, Parent Interface, VRF, Mode, Admin Access, Speed, Description 등 14개 열)
* **External Resource** : 외부 위협 인텔리전스 피드(`config system external-resource`) 객체 (활성화 여부, 리소스 URL, 갱신 주기, IP 등)

### 2. 객체/그룹의 실제 IP·포트·코멘트 다중 행 전개 (Multi-row Flattening)
* 객체명이나 그룹명만 단순히 표기되는 기존 방식과 달리, 그룹에 포함된 **멤버 객체, 실제 IP, 서브넷 대역, 포트 번호**를 자동으로 추적하여 행 단위로 전개합니다.
* **IP 포맷 통합**:
  * 단일 호스트: `x.x.x.x/32`
  * C클래스 등 서브넷: `x.x.x.x/24`
  * IP 범위: `1.1.1.1-1.1.1.10`
* **객체별 코멘트 분리**: 각 객체 및 그룹의 코멘트를 추출하여 `Src Comment`, `Dst Comment`, `Svc Comment` 열에 개별 수록하고, 정책 자체 코멘트(`Policy Comment`)와 완벽히 분리했습니다.
* **Internet Service 및 External Resource 매핑**:
  * `Fortinet-DNS`, `Dropbox-Web` 등의 인터넷 서비스 객체를 서비스 포트가 아닌 출발지/목적지 객체(`Src/Dst OBJ Name`) 및 타입(`internet-service`)으로 정확히 분류합니다.
  * `AbuseIPDB` 등 외부 위협 피드 객체를 정확한 리소스 타입(`address`, `domain`, `malware` 등)으로 매핑합니다.
  * 목적지 VIP 객체는 `Dst Type`에 `static-nat` 등 실제 VIP 타입을 자동 매핑합니다.

### 3. 통합 보안 프로파일(Sec Profile) 및 코멘트 1:1 매핑
* 방화벽 정책에 적용된 다양한 보안 프로파일(SSL/SSH Inspection, IPS Sensor, Web Filter, Antivirus, DNS Filter, Application Control, File Filter 등)을 하나의 **`Sec Profile`** 열에 줄바꿈으로 통합 수록합니다.
* 각 프로파일 정의에 입력된 세부 코멘트를 추적하여 우측 **`Sec Profile Comment`** 열에 1:1로 대응되는 줄바꿈 텍스트로 함께 표기합니다.
* VDOM 및 정책 설정에 따른 검사 모드(**`Inspection Mode`**: `Flow-based` / `Proxy-based`)를 명확하게 표시합니다.

### 4. Src Interface 기준 오름차순 안정 정렬 (Stable Sort)
* 방화벽 정책 목록을 **`Src Interface` 기준으로 오름차순 정렬**하여 인터페이스별 정책 검토를 용이하게 합니다.
* 동일 인터페이스 내에서는 방화벽 정책의 상하 우선순위(First-Match Precedence)가 변경되지 않도록 **파이썬 Timsort 기반 안정 정렬(Stable Sort)을 적용하여 원래 순서를 100% 보존**합니다.

### 5. 세분화된 트래픽 로깅(Log Traffic) 및 UTM 상태 표기
* **Log Traffic**: `disable`, `all`, `utm` 기본 상태와 함께 `set logtraffic-start enable` 활성화 시 줄바꿈으로 **`session-start`**를 병기(예: `all\nsession-start`, `utm\nsession-start`)하여 세션 시작 로깅 여부를 명확히 확인할 수 있습니다.
* **UTM Status**: UTM 미설정 또는 비활성화 시 명시적으로 **`disable`**로 표기합니다.

### 6. 스마트 세로 셀 병합 및 세로 위쪽 맞춤 (Top Alignment)
* 한 정책에 여러 객체가 포함되어 다중 행으로 전개될 때, **정책 공통 속성(Seq, VDOM, Enable, ID, Name, Action, Schedule 등)** 및 **동일 그룹 멤버 영역**을 세로로 자동 병합합니다.
* 다중 행 병합 셀과 일반 데이터 셀의 텍스트 수직 정렬을 모두 **세로 상단(`top`)**으로 통일하여 가독성을 극대화했습니다.

### 7. 엔터프라이즈급 시각화 스타일링
* **헤더 및 글자색 구분**:
  * 출발지(Src) 영역: 파란색 계열 헤더 + 파란색 글씨
  * 목적지(Dst) 영역: 빨간색 계열 헤더 + 빨간색 글씨
  * 서비스(Service) 영역: 블루그레이 계열 헤더
* **줄무늬 음영 (Zebra Striping)**: 행 전체 열에 일관된 홀수/짝수 교차 배경색을 적용하여 시선 이동이 편안합니다.
* **비활성화 정책 음영**: 비활성화된 정책(`Enable == N`)은 진한 회색 배경으로 표시되어 활성 정책과 즉시 구별됩니다.
* **내용 없는 빈 시트 빨간색 탭 표시**: 설정된 정책이나 리소스가 없어 컬럼 헤더(1행) 외에 본문 데이터가 전혀 없는 시트는 엑셀 탭 색상을 **빨간색(Red)**으로 자동 지정하여 비어 있는 정책 항목을 한눈에 즉시 식별할 수 있습니다.

### 8. 호스트네임 디렉토리 생성 및 VDOM별 엑셀 파일 분할
* 대용량 설정 파일의 전체 정책을 단일 엑셀에 몰아넣을 때 발생하는 렉(버벅임)을 원천 차단합니다.
* 장비 호스트명(예: `JBNU_SVF-FW1`)으로 폴더를 자동 생성하고, 그 안에 **각 VDOM별 엑셀 파일(`<vdom_name>.xlsx`)을 독립적으로 생성**합니다.
* 전체 VDOM의 정책 수와 객체 수를 한눈에 볼 수 있는 총괄 요약 파일(**`_TOTAL_SUMMARY.xlsx`**)을 함께 제공합니다.

### 9. GUI 내장 및 동적 경로 자동 지정
* 직관적인 파일 탐색기 찾아보기(`Select File...`) 제공
* **동적 출력 경로 자동 갱신**: `.conf` 파일을 선택할 때마다 해당 파일이 위치한 폴더로 결과 저장 경로(`Output Directory`)가 자동으로 즉시 변경됩니다.
* 실시간 진행률 게이지(Progressbar) 및 구문 강조(Syntax Highlighting) 터미널 콘솔 로그
* 윈도우 고해상도(High-DPI / ClearType) 지원으로 QHD/4K 모니터에서도 번짐 없이 칼같이 선명한 텍스트 렌더링
* 변환 완료 후 번거로운 팝업 확인창 없이 하단 상태바 알림 및 결과 폴더 자동 열기 지원

---

## 💻 시스템 요구 사양

| 항목 | 권장 사양 |
|---|---|
| **운영체제** | Windows 10 / Windows 11 (64-bit) *(macOS / Linux는 CLI 모드 지원)* |
| **Python 버전** | **Python 3.8 이상** (Python 3.10 ~ 3.13 완벽 호환) |
| **디스플레이** | 1920×1080 (FHD) 이상 (QHD, 4K 배율 환경 완벽 지원) |
| **필수 라이브러리** | `openpyxl` |

---

## 📦 파이썬 및 필수 모듈 설치

### 1. Python 설치
1. [Python 공식 웹사이트](https://www.python.org/downloads/)에서 Python 3.10 이상 최신 버전을 다운로드합니다.
2. 설치 실행 창 첫 화면 맨 아래에 있는 **`[✔] Add python.exe to PATH`** 체크박스를 **반드시 체크**하고 설치를 진행합니다.

### 2. 필수 라이브러리(`openpyxl`) 설치
명령 프롬프트(CMD) 또는 파워쉘을 열고 다음 명령어를 입력합니다:

```bash
pip install openpyxl
```

---

## 🚀 프로그램 실행 방법

### 방법 1. 탐색기에서 더블 클릭 (가장 권장)
1. 프로그램 폴더 내의 **`start.bat`** 파일을 더블 클릭합니다.
2. GUI 창이 열립니다.
3. **[Select File...]** 버튼을 눌러 변환할 포티게이트 `.conf` 파일을 선택합니다.
4. **[Start Export to Excel]** 버튼을 클릭합니다.
5. 내보내기가 완료되면 결과 저장 폴더가 자동으로 열립니다.

### 방법 2. 파이썬 직접 실행 (GUI 모드)
터미널에서 인자 없이 실행하면 GUI 모드로 실행됩니다:

```bash
python fortigate_policy_to_excel.py
```

### 방법 3. 커맨드라인 실행 (CLI 모드 / 자동화 스크립트 연동)
GUI 창 없이 백그라운드나 배치 스크립트에서 명령줄 인자로 바로 실행할 수 있습니다:

```bash
# 기본 사용법: python fortigate_policy_to_excel.py <설정파일경로> [출력디렉토리]
python fortigate_policy_to_excel.py "C:\backup\my_firewall.conf"
```

---

## 📊 엑셀 시트 구성 및 디자인 특징

### 1. Firewall Policy 시트 컬럼 구성 (33개 열)

```
[정책 기본 정보]
  Col 1: Seq                 - 시퀀스 번호 (Src Interface 기준 오름차순 안정 정렬)
  Col 2: VDOM                - VDOM 명
  Col 3: Enable              - 활성화 여부 (Y / N)
  Col 4: ID                  - 정책 ID
  Col 5: Name                - 정책 이름
  Col 6: Action              - accept / deny

[출발지 영역 (Src) - 파란색 계열 헤더 & 글씨]
  Col 7: Src Interface       - 출발지 인터페이스
  Col 8: Src Group OBJ       - 출발지 그룹명
  Col 9: Src OBJ Name        - 세부 객체명 (주소 / 그룹 / Internet Service / External Resource)
  Col 10: Src Type           - ipmask / iprange / fqdn / internet-service 등
  Col 11: Src IP             - 실제 IP / 서브넷 대역 / 범위
  Col 12: Src Comment        - 출발지 객체 코멘트

[목적지 영역 (Dst) - 빨간색 계열 헤더 & 글씨]
  Col 13: Dst Interface      - 목적지 인터페이스
  Col 14: Dst Group OBJ      - 목적지 그룹명
  Col 15: Dst OBJ Name       - 세부 객체명 (주소 / 그룹 / Internet Service / External Resource)
  Col 16: Dst Type           - ipmask / iprange / fqdn / internet-service / static-nat(VIP) 등
  Col 17: Dst IP             - 실제 IP / 서브넷 대역 / 범위
  Col 18: Dst Comment        - 목적지 객체 코멘트

[서비스 영역 (Service) - 블루그레이 계열 헤더]
  Col 19: Svc Group OBJ      - 서비스 그룹명
  Col 20: Svc OBJ Name       - 서비스 객체명
  Col 21: Svc Port           - 포트 번호 (IP/ALL은 ALL 로 표기, Internet Service 적용 시 공란)
  Col 22: Svc Comment        - 서비스 객체 코멘트

[기타 및 보안 프로파일]
  Col 23: Schedule           - 스케줄
  Col 24: NAT                - NAT 사용 여부 (enable / disable)
  Col 25: IP Pool            - IP Pool 사용 여부
  Col 26: Pool Name          - Pool 이름
  Col 27: Pool IP            - Pool 실제 할당 IP 대역
  Col 28: Inspection Mode    - Flow-based / Proxy-based
  Col 29: UTM Status         - UTM 적용 여부 (enable / disable)
  Col 30: Sec Profile        - 적용된 모든 보안 프로파일 목록 (SSL/SSH, IPS, AV, WebFilter 등 줄바꿈 통합)
  Col 31: Sec Profile Comment- 각 보안 프로파일 설정의 코멘트 (1:1 매핑 줄바꿈)
  Col 32: Log Traffic        - 트래픽 로깅 (all / utm / disable 및 session-start 결합 표기)
  Col 33: Comment            - 정책 자체 코멘트
```

### 2. Local-in Policy 시트 컬럼 구성 (22개 열)

```
[정책 기본 정보]
  Col 1: Seq                 - 시퀀스 번호
  Col 2: VDOM                - VDOM 명
  Col 3: Enable              - 활성화 여부 (Y / N)
  Col 4: ID                  - 정책 ID
  Col 5: Interface           - 인바운드 수신 인터페이스

[출발지 영역 (Src)]
  Col 6: Src Group OBJ       - 출발지 그룹명
  Col 7: Src OBJ Name        - 출발지 객체명
  Col 8: Src Type            - 주소 객체 타입 (ipmask / fqdn 등)
  Col 9: Src IP              - 실제 출발지 IP 대역
  Col 10: Src Comment        - 출발지 객체 코멘트

[목적지 영역 (Dst)]
  Col 11: Dst Group OBJ      - 목적지 그룹명
  Col 12: Dst OBJ Name       - 목적지 객체명
  Col 13: Dst Type           - 주소 객체 타입
  Col 14: Dst IP             - 실제 목적지 IP 대역
  Col 15: Dst Comment        - 목적지 객체 코멘트

[액션 및 서비스]
  Col 16: Action             - accept / deny
  Col 17: Svc Group OBJ      - 서비스 그룹명
  Col 18: Svc OBJ Name       - 서비스 객체명
  Col 19: Svc Port           - 서비스 포트 번호
  Col 20: Svc Comment        - 서비스 객체 코멘트

[기타]
  Col 21: Schedule           - 스케줄
  Col 22: Comment            - 정책 코멘트
```

### 3. Central-NAT 시트 컬럼 구성 (21개 열)

```
[정책 기본 정보]
  Col 1: Seq                 - 시퀀스 번호
  Col 2: VDOM                - VDOM 명
  Col 3: Enable              - 활성화 여부 (Y / N)
  Col 4: ID                  - NAT 규칙 ID
  Col 5: Src Interface       - 출발지 인터페이스
  Col 6: Dst Interface       - 목적지 인터페이스

[원래 출발지 영역 (Orig Src)]
  Col 7: Orig Group OBJ      - 변환 전 원래 출발지 그룹명
  Col 8: Orig OBJ Name       - 변환 전 원래 출발지 객체명
  Col 9: Orig Type           - 객체 타입 (ipmask 등)
  Col 10: Orig IP            - 원래 출발지 IP 대역
  Col 11: Orig Comment       - 원래 출발지 코멘트

[목적지 영역 (Dst)]
  Col 12: Dst Group OBJ      - 목적지 그룹명
  Col 13: Dst OBJ Name       - 목적지 객체명
  Col 14: Dst Type           - 객체 타입
  Col 15: Dst IP             - 목적지 IP 대역
  Col 16: Dst Comment        - 목적지 코멘트

[NAT IP Pool 및 상태]
  Col 17: NAT IP Pool Name   - 매핑할 NAT IP Pool 이름
  Col 18: NAT Pool IP        - 실제 변환될 NAT Pool IP 대역
  Col 19: NAT Pool Type      - IP Pool 타입 (overload, one-to-one 등)
  Col 20: NAT                - NAT 활성화 여부 (enable / disable)
  Col 21: Comment            - NAT 규칙 코멘트
```

### 4. DNAT (VIP) 시트 컬럼 구성 (18개 열)

```
[VIP 기본 정보]
  Col 1: Seq                 - 시퀀스 번호
  Col 2: VDOM                - VDOM 명
  Col 3: Name                - VIP 객체명
  Col 4: Type                - VIP 타입 (static-nat, server-load-balance 등)

[외부 및 매핑 IP/인터페이스]
  Col 5: External IP         - 외부 공인 IP / 수신 IP (extip)
  Col 6: Mapped IP           - 내부 사설 매핑 IP (mappedip)
  Col 7: External Interface  - 연결 외부 인터페이스 (extintf)

[포트 포워딩 설정]
  Col 8: Port Forward        - 포트 포워딩 활성화 여부 (enable / disable)
  Col 9: Protocol            - 프로토콜 (tcp / udp / sctp 등)
  Col 10: External Port      - 외부 수신 포트 대역 (extport)
  Col 11: Mapped Port        - 내부 매핑 포트 대역 (mappedport)

[서버 로드밸런싱 설정 (SLB)]
  Col 12: Server Type        - 서버 타입 (http, https, ip 등)
  Col 13: LDB Method         - 부하분산 방식 (round-robin, weighted 등)
  Col 14: Monitor            - 헬스 체크 모니터 이름
  Col 15: Real Server IP     - 실제 리얼 서버 IP
  Col 16: Real Server Port   - 리얼 서버 서비스 포트
  Col 17: Real Server Weight - 서버 가중치 (Weight)

[기타]
  Col 18: Comment            - VIP 객체 코멘트
```

### 5. DoS Policy 시트 컬럼 구성 (26개 열)

```
[정책 기본 정보]
  Col 1: Seq                 - 시퀀스 번호
  Col 2: VDOM                - VDOM 명
  Col 3: Enable              - 활성화 여부 (Y / N)
  Col 4: ID                  - DoS 정책 ID
  Col 5: Interface           - 보호 인터페이스

[출발지 영역 (Src)]
  Col 6: Src Group OBJ       - 출발지 그룹명
  Col 7: Src OBJ Name        - 출발지 객체명
  Col 8: Src Type            - 주소 객체 타입
  Col 9: Src IP              - 실제 출발지 IP 대역
  Col 10: Src Comment        - 출발지 객체 코멘트

[목적지 영역 (Dst)]
  Col 11: Dst Group OBJ      - 목적지 그룹명
  Col 12: Dst OBJ Name       - 목적지 객체명
  Col 13: Dst Type           - 주소 객체 타입
  Col 14: Dst IP             - 실제 목적지 IP 대역
  Col 15: Dst Comment        - 목적지 객체 코멘트

[서비스 영역 (Service)]
  Col 16: Svc Group OBJ      - 서비스 그룹명
  Col 17: Svc OBJ Name       - 서비스 객체명
  Col 18: Svc Port           - 서비스 포트
  Col 19: Svc Comment        - 서비스 객체 코멘트

[비정상 트래픽 탐지 (Anomaly Detection)]
  Col 20: Anomaly Name       - 공격 유형 / 비정상 트래픽 명칭 (예: tcp_syn_flood 등)
  Col 21: Anomaly Status     - 아노말리 방어 활성화 상태 (enable / disable)
  Col 22: Anomaly Log        - 차단 로그 기록 여부 (enable / disable)
  Col 23: Anomaly Quarant    - 격리 설정 (attacker / disable)
  Col 24: Anomaly Action     - 조치 방식 (pass / block / disable)
  Col 25: Anomaly Threshold  - 임계값 (Threshold 패킷/초)

[정책 코멘트]
  Col 26: Comment            - DoS 정책 코멘트
```

### 6. ACL Policy 시트 컬럼 구성 (20개 열)

```
[정책 기본 정보]
  Col 1: Seq                 - 시퀀스 번호
  Col 2: VDOM                - VDOM 명
  Col 3: Enable              - 활성화 여부 (Y / N, status disable 시 N 및 진한 회색 음영)
  Col 4: ID                  - ACL 규칙 ID (edit 번호)
  Col 5: Interface           - 인바운드 수신 인터페이스

[출발지 영역 (Src) - 파란색 계열]
  Col 6: Src Group OBJ       - 출발지 그룹명 (그룹 객체 포함 시 자동 전개)
  Col 7: Src OBJ Name        - 출발지 세부 객체명
  Col 8: Src Type            - 주소 객체 타입 (ipmask / geography 등)
  Col 9: Src IP              - 실제 출발지 IP 대역
  Col 10: Src Comment        - 출발지 객체 코멘트

[목적지 영역 (Dst) - 빨간색 계열]
  Col 11: Dst Group OBJ      - 목적지 그룹명
  Col 12: Dst OBJ Name       - 목적지 세부 객체명
  Col 13: Dst Type           - 주소 객체 타입
  Col 14: Dst IP             - 실제 목적지 IP 대역
  Col 15: Dst Comment        - 목적지 객체 코멘트

[서비스 영역 (Service)]
  Col 16: Svc Group OBJ      - 서비스 그룹명
  Col 17: Svc OBJ Name       - 서비스 객체명
  Col 18: Svc Port           - 서비스 포트
  Col 19: Svc Comment        - 서비스 객체 코멘트

[정책 코멘트]
  Col 20: Comment            - ACL 정책 코멘트
※ firewall acl에 설정된 다중 주소 객체 및 그룹 객체를 완벽히 행 분리하여 풀어서 표기합니다.
```

### 7. Static Route 시트 컬럼 구성 (11개 열)

```
[라우팅 기본 정보]
  Col 1: Seq                 - 시퀀스 번호
  Col 2: VDOM                - VDOM 명
  Col 3: Enable              - 활성화 여부 (Y / N, status disable 시 N 및 진한 회색 음영 처리)
  Col 4: ID                  - Static Route ID (edit 번호)

[목적지 및 경로]
  Col 5: Destination         - 목적지 네트워크 대역 (IP/CIDR 자동 변환, 미지정 시 0.0.0.0/0)
  Col 6: Gateway             - 게이트웨이 IP 주소
  Col 7: Interface           - 아웃바운드 인터페이스 (device)

[메트릭 및 옵션]
  Col 8: Distance            - 관리 거리 (기본값 10)
  Col 9: Priority            - 우선순위 (Priority)
  Col 10: Options            - 특수 플래그 요약 (Blackhole, Dynamic-GW, BFD, Link-Mon-Exempt)
  Col 11: Comment            - 라우트 코멘트
```

### 8. Policy Route 시트 컬럼 구성 (21개 열)

```
[정책 라우팅 기본 정보]
  Col 1: Seq                 - 시퀀스 번호
  Col 2: VDOM                - VDOM 명
  Col 3: Enable              - 활성화 여부 (Y / N, disable 시 N 및 진한 회색 음영)
  Col 4: ID                  - 정책 라우팅 ID
  Col 5: Incoming Intf       - 인바운드 수신 인터페이스 (input-device)
  Col 6: Outgoing Intf       - 아웃바운드 송신 인터페이스 (output-device)
  Col 7: Gateway             - 게이트웨이 IP 주소
  Col 8: Action              - permit / deny
  Col 9: Protocol            - ALL / TCP(6) / UDP(17) / ICMP(1) 등
  Col 10: Port Range         - 포트 범위 (start-port ~ end-port, 미지정 시 ALL)

[출발지 영역 (Src)]
  Col 11: Src Group OBJ      - 출발지 그룹명
  Col 12: Src OBJ Name       - 세부 객체명 (주소 객체 / 그룹 / 직접 지정 서브넷)
  Col 13: Src Type           - 객체 타입 (ipmask / fqdn 등)
  Col 14: Src IP             - 실제 출발지 IP 대역 (객체 추적 자동 전개)
  Col 15: Src Comment        - 출발지 객체 코멘트

[목적지 영역 (Dst)]
  Col 16: Dst Group OBJ      - 목적지 그룹명
  Col 17: Dst OBJ Name       - 세부 객체명
  Col 18: Dst Type           - 객체 타입
  Col 19: Dst IP             - 실제 목적지 IP 대역 (객체 추적 자동 전개)
  Col 20: Dst Comment        - 목적지 객체 코멘트

[기타]
  Col 21: Comment            - 정책 라우팅 코멘트
```

### 9. OSPF 시트 구조 (5개 계층 섹션 테이블)

```
[1. OSPF Global Settings]
  - vDOM, Router ID, Total Networks, Total Interfaces, Total Areas
[2. OSPF Networks (config network)]
  - Seq, ID, Prefix (IP/CIDR), Area ID
[3. OSPF Interfaces (config ospf-interface)]
  - Seq, Name, Interface, Cost, Dead / Hello Interval, Network Type, Priority (기본값 1), Authentication (기본값 none)
[4. OSPF Redistribution (config redistribute)]
  - Seq, Protocol (Connected, Static, RIP, BGP, ISIS), Status, Route-Map, Metric, Metric Type
[5. Route-Map & Filter Details (config router route-map / access-list)]
  - Seq, Route-Map, Rule, Route-Map Action (PERMIT/DENY), Match Target, Filtered Prefix (규칙별 행 분리), Action (permit/deny), Exact Match (enable/disable), ACL Comment
※ OSPF 미설정 VDOM은 헤더만 표시되며 엑셀 탭 색상이 빨간색(Red)으로 자동 마킹됩니다.
※ 각 섹션 제목 배경색은 실제 표 열 너비(5열, 4열, 8열, 6열, 9열)에 정확히 맞춰져 가독성을 극대화합니다.
```

### 10. IPsec VPN 시트 컬럼 구성 (30개 열)

```
[공통 정보 - 기본 헤더 (#2F5496)]
  Col 1: Seq                 - 시퀀스 번호 (중앙 정렬 병합)
  Col 2: vDOM                - VDOM 명 (중앙 정렬 병합)

[Phase 1 터널 및 네트워크/고급 설정 - 네이비 블루 헤더 (#1F4E78)]
  Col 3: P1 Name             - Phase 1 터널 이름 (좌측 정렬 병합)
  Col 4: Interface           - 물리 바인딩 인터페이스 (중앙 정렬 병합)
  Col 5: Remote Gateway      - 상대방 공인 IP / (Dialup/Dynamic) (중앙 정렬 병합)
  Col 6: Local Gateway       - 로컬 바인딩 IP (중앙 정렬 병합)
  Col 7: IKE Version         - IKE 버전 (v1 / v2) (중앙 정렬 병합)
  Col 8: P1 Proposal         - 1단계 암호화/인증 알고리즘 (좌측 정렬 병합)
  Col 9: P1 DH Group         - Phase 1 Diffie-Hellman 그룹 (미지정 시 기본값 '14 5' 자동 기입)
  Col 10: NAT Traversal      - NAT 트래버설 (기본값 enable)
  Col 11: Keepalive Frequency- Keepalive 주기 (초 단위, 기본값 10)
  Col 12: Dead Peer Detection- DPD 모드 (disable / on-idle / on-demand, 기본값 on-demand)
  Col 13: DPD Retry Count    - DPD 재시도 횟수 (기본값 3)
  Col 14: DPD Retry Interval - DPD 재시도 간격 (초 단위 숫자만 표기, 기본값 20)
  Col 15: FEC Egress         - 순방향 오류 정정 송신 (기본값 disable)
  Col 16: FEC Ingress        - 순방향 오류 정정 수신 (기본값 disable)
  Col 17: Add Route          - 게이트웨이 경로 자동 추가 (add-gw-route, 기본값 enable)
  Col 18: Auto Discovery Sender   - 동적 터널 발신 탐색 (기본값 disable)
  Col 19: Auto Discovery Receiver - 동적 터널 수신 탐색 (기본값 disable)
  Col 20: Exchange Interface IP   - 인터페이스 IP 교환 (기본값 disable)
  Col 21: Device Creation    - 터널 가상 인터페이스 생성 (net-device, 기본값 disable)
  Col 22: P1 Comment         - Phase 1 코멘트 (좌측 정렬 병합)

[Phase 2 서브 터널 정보 - 다크 틸 그린 헤더 (#2A5C5A)]
  Col 23: P2 Name            - Phase 2 서브 터널 이름
  Col 24: P2 Proposal        - 2단계 암호화/인증 알고리즘
  Col 25: P2 DH Group        - Phase 2 Diffie-Hellman 그룹 (미지정 시 기본값 '14 5' 자동 기입)
  Col 26: Local Subnet / Src - 로컬 서브넷 대역 또는 주소 객체명 (미지정 시 기본값 '0.0.0.0/0' 자동 기입)
  Col 27: Remote Subnet / Dst- 원격 서브넷 대역 또는 주소 객체명 (미지정 시 기본값 '0.0.0.0/0' 자동 기입)
  Col 28: Auto Negotiate     - 자동 협상 활성화 여부 (enable / disable)
  Col 29: Keepalive          - Keepalive 활성화 여부 (enable / disable)
  Col 30: P2 Comment         - Phase 2 코멘트
※ 데이터 길이에 비해 헤더명이 긴 컬럼은 2줄 줄바꿈을 적용하여 열 너비를 최적화하고 가독성을 극대화했습니다.
```

### 11. Network Interface 시트 컬럼 구성 (16개 열)

```
[인터페이스 기본 정보]
  Col 1: Seq                 - 시퀀스 번호
  Col 2: vDOM                - VDOM 명
  Col 3: Status              - 인터페이스 활성화 상태 (UP / DOWN, DOWN 시 빨간색 음영)
  Col 4: Name                - 인터페이스 식별 이름 (굵은 글씨)
  Col 5: Alias               - 인터페이스 별칭(Alias)
  Col 6: Type                - 인터페이스 타입 (physical, vlan, loopback, aggregate, tunnel 등)

[네트워크 주소 및 계층 바인딩]
  Col 7: Primary IP          - 기본 할당 IP 주소 및 서브넷 마스크 (미지정 시 기본값 '0.0.0.0/0' 자동 기입, 파란색 글씨)
  Col 8: Secondary IP        - 2차 보조 IP 대역 (config secondaryip 파싱, 2개 이상 시 줄바꿈)
  Col 9: Remote IP (Tunnel)  - 터널 대향 IP 대역 (tunnel 인터페이스의 set remote-ip 파싱)
  Col 10: VLAN ID            - VLAN 태그 번호
  Col 11: Parent / Member Interface - 상위 부모 인터페이스(VLAN/터널) 또는 하위 멤버 목록(어그리게이트/리던던트) 통합 표기
  Col 12: VRF                - 할당된 VRF ID (미지정 시 기본값 '0' 자동 기입)
  Col 13: Addressing Mode    - 주소 할당 방식 (static, dhcp, pppoe 등)

[보안 및 상세 관리]
  Col 14: Administrative Access - 허용된 관리 접근 프로토콜 (ping, https, ssh, snmp, fgfm 등)
  Col 15: Speed / Duplex     - 링크 속도 및 듀플렉스 설정 (1000 / auto, 10000 / full 등으로 가독성 최적화)
  Col 16: Description        - 인터페이스 설명(설정 코멘트)
```

### 12. External Resource 시트 컬럼 구성 (9개 열)

```
  Col 1: Seq                 - 시퀀스 번호
  Col 2: VDOM                - VDOM 명
  Col 3: Enable              - 활성화 여부 (Y / N, set status disable 시 N)
  Col 4: Name                - 외부 리소스 객체명 (예: AbuseIPDB_Blacklist_Score-75 등)
  Col 5: Type                - 리소스 타입 (address / domain / malware 등)
  Col 6: Resource URL        - 외부 피드 다운로드 URL
  Col 7: Refresh Rate (min)  - 자동 갱신 주기 (분 단위)
  Col 8: Source IP           - 외부 접속 시 사용할 출발지 IP
  Col 9: Comment             - 리소스 코멘트
```

---

## 📁 폴더 및 파일 구조

```
fortigate_policy_to_excel/
│
├── fortigate_policy_to_excel.py   # [핵심] 변환 엔진 및 GUI 통합 단일 스크립트
├── README_ko.md                   # [문서] 한국어 사용자 매뉴얼 및 가이드
├── README_en.md                   # [문서] 영문 사용자 매뉴얼 및 가이드
├── Start.bat                      # [실행] Windows 간편 실행 배치 스크립트
│
└── <호스트네임>/                  # [결과물] 변환 완료 시 생성되는 결과 디렉토리
    ├── _TOTAL_SUMMARY.xlsx        # 전체 VDOM 정책·라우팅·VPN·인터페이스·객체 수 집계 총괄 요약
    ├── root.xlsx                  # root VDOM 엑셀
    └── ...                        # 각 VDOM별 독립 엑셀 파일들
```

---

## ❓ 자주 묻는 질문 (FAQ & Troubleshooting)

### Q1. "Python 실행 경로를 찾을 수 없습니다"라고 뜹니다.
* **원인**: Python이 PC에 설치되어 있지 않거나, 설치 시 "Add Python to PATH"가 누락된 경우입니다.
* **해결법**: Python 공식 홈페이지에서 Python 설치 프로그램을 다시 실행한 후 **`Modify`** 를 선택하고 **`Add Python to PATH`** 옵션을 체크하여 설치를 완료해 주세요.

### Q2. "openpyxl 필요: pip install openpyxl" 에러가 발생합니다.
* **원인**: 엑셀 제어 라이브러리가 미설치 상태입니다.
* **해결법**: 명령 프롬프트(CMD)를 열고 `pip install openpyxl` 명령을 입력하여 설치해 주세요.

### Q3. 화면 글씨가 흐릿하게 보이지는 않나요?
* 본 프로그램은 윈도우의 **Per-Monitor High-DPI(고해상도 네이티브 인식)** 및 **ClearType 렌더링**을 기본 내장하고 있어, QHD나 4K 모니터(125%, 150%, 175% 확대 환경)에서도 번짐 없이 선명한 고품질 폰트로 출력됩니다.

### Q4. 10만 라인이 넘는 대용량 config 파일도 변환 가능한가요?
* 네, 백그라운드 멀티스레딩(`threading.Thread`)으로 파싱과 엑셀 생성을 수행하므로, 12만 라인 이상의 대형 방화벽 설정 파일도 프로그램 멈춤(응답 없음) 없이 수초 내에 안정적으로 변환됩니다.

---
