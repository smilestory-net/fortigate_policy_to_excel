# FortiGate Policy to Excel Exporter v1.0

포티게이트(FortiGate) 방화벽의 백업 설정 파일(`.conf` 또는 `.txt`)을 분석하여, VDOM별 정책과 객체를 가독성 높은 엑셀(`.xlsx`) 문서로 자동 변환·분할 생성해 주는 도구입니다.


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

### 1. 5대 핵심 정책 완벽 추출 및 개별 시트 저장
* **Firewall Policy** : 일반 방화벽 정책 (Action, Schedule, NAT, IP Pool, UTM Profile, Log 등 수록)
* **Local-in Policy** : 방화벽 장비 자체 접근 제어 정책
* **Central-NAT** : Central SNAT Map (Original IP/Port ↔ Translated IP/Port 매핑)
* **DNAT (VIP)** : 가상 IP 및 포트포워딩 매핑 (External IP/Port ↔ Mapped IP/Port)
* **DoS Policy** : 서비스 거부 공격(DoS/DDoS) 방어 정책

### 2. 객체/그룹의 실제 IP·포트·코멘트 다중 행 전개 (Multi-row Flattening)
* 객체명이나 그룹명만 단순히 표기되는 기존 방식과 달리, 그룹에 포함된 **멤버 객체, 실제 IP, 서브넷 대역, 포트 번호**를 자동으로 추적하여 행 단위로 전개합니다.
* **IP 포맷 통합**:
  * 단일 호스트: `x.x.x.x/32`
  * C클래스 등 서브넷: `x.x.x.x/24`
  * IP 범위: `1.1.1.1-1.1.1.10`
* **객체별 코멘트 분리**: 각 객체 및 그룹의 코멘트를 추출하여 `Src Comment`, `Dst Comment`, `Svc Comment` 열에 개별 수록하고, 정책 자체 코멘트(`Comments`)와 완벽 분리했습니다.

### 3. 스마트 세로 셀 병합 (Vertical Cell Merging)
* 한 정책에 여러 객체가 포함되어 다중 행으로 전개될 때, **정책 공통 속성(Seq, VDOM, Enable, ID, Name, Action, Schedule 등)** 및 **동일 그룹 멤버 영역**을 세로로 자동 병합하여 가독성을 극대화합니다.

### 4. 엔터프라이즈급 시각화 스타일링
* **헤더 및 글자색 구분**:
  * 출발지(Src) 영역: 파란색 계열 헤더 + 파란색 글씨
  * 목적지(Dst) 영역: 빨간색 계열 헤더 + 빨간색 글씨
  * 서비스(Service) 영역: 블루그레이 계열 헤더
* **줄무늬 음영 (Zebra Striping)**: 행 전체 열에 일관된 홀수/짝수 교차 배경색을 적용하여 시선 이동이 편안합니다.
* **비활성화 정책 음영**: 비활성화된 정책(`Enable == N`)은 진한 회색 배경으로 표시되어 활성 정책과 즉시 구별됩니다.

### 5. 호스트네임 디렉토리 생성 및 VDOM별 엑셀 파일 분할
* 대용량 설정 파일의 전체 정책을 단일 엑셀에 몰아넣을 때 발생하는 렉(버벅임)을 원천 차단합니다.
* 장비 호스트명(예: `JBNU_SVF-FW1`)으로 폴더를 자동 생성하고, 그 안에 **각 VDOM별 엑셀 파일(`<vdom_name>.xlsx`)을 독립적으로 생성**합니다.
* 전체 VDOM의 정책 수와 객체 수를 한눈에 볼 수 있는 총괄 요약 파일(**`_TOTAL_SUMMARY.xlsx`**)을 함께 제공합니다.

### 6. GUI 내장
* 직관적인 파일 탐색기 찾아보기(`Browse...`) 제공
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

> **참고**: `tkinter`, `threading`, `ctypes`, `re` 등 GUI 및 시스템 연동 모듈은 Python 표준 라이브러리에 기본 내장되어 있으므로 추가 설치가 필요 없습니다.

---

## 🚀 프로그램 실행 방법

### 방법 1. 탐색기에서 더블 클릭 (가장 권장)
1. 프로그램 폴더 내의 **`실행하기.bat`** 파일을 더블 클릭합니다.
2. 다크 테마 GUI 창이 열립니다.
3. **[Browse...]** 버튼을 눌러 변환할 포티게이트 `.conf` (또는 `.txt`) 파일을 선택합니다.
4. 초록색 **`[▶ 엑셀 변환 실행 (Start Conversion)]`** 버튼을 클릭합니다.
5. 변환이 완료되면 결과 엑셀 폴더가 자동으로 열립니다.

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

### Firewall Policy 시트 컬럼 구성 (32개 열)

```
[정책 기본 정보]
  Col 1: Seq            - 시퀀스 번호
  Col 2: VDOM           - VDOM 명
  Col 3: Enable         - 활성화 여부 (Y / N)
  Col 4: ID             - 정책 ID
  Col 5: Name           - 정책 이름
  Col 6: Action         - accept / deny

[출발지 영역 (Src) - 파란색 계열 헤더 & 글씨]
  Col 7: Src Interface  - 출발지 인터페이스
  Col 8: Src Group OBJ  - 출발지 그룹명
  Col 9: Src OBJ Name   - 세부 객체명
  Col 10: Src Type      - ipmask / iprange / fqdn 등
  Col 11: Src IP        - 실제 IP / 서브넷 대역 / 범위
  Col 12: Src Comment   - 출발지 객체 코멘트

[목적지 영역 (Dst) - 빨간색 계열 헤더 & 글씨]
  Col 13: Dst Interface - 목적지 인터페이스
  Col 14: Dst Group OBJ - 목적지 그룹명
  Col 15: Dst OBJ Name  - 세부 객체명
  Col 16: Dst Type      - ipmask / iprange / fqdn 등
  Col 17: Dst IP        - 실제 IP / 서브넷 대역 / 범위
  Col 18: Dst Comment   - 목적지 객체 코멘트

[서비스 영역 (Service) - 블루그레이 계열 헤더]
  Col 19: Svc Group OBJ - 서비스 그룹명
  Col 20: Svc OBJ Name  - 서비스 객체명
  Col 21: Svc Port      - 포트 번호 (IP/ALL은 ALL 로 표기)
  Col 22: Svc Comment   - 서비스 객체 코멘트

[기타 및 보안 프로파일]
  Col 23: Schedule      - 스케줄
  Col 24: NAT           - NAT 사용 여부 (enable / disable)
  Col 25: IP Pool       - IP Pool 사용 여부
  Col 26: Pool Name     - Pool 이름
  Col 27: Pool IP       - Pool 실제 할당 IP 대역
  Col 28: UTM Status    - UTM 적용 여부
  Col 29: SSL/SSH Profile
  Col 30: IPS Sensor
  Col 31: Log Traffic   - all / utm / disable
  Col 32: Comments      - 정책 자체 코멘트
```

---

## 📁 폴더 및 파일 구조

```
fortigate_policy_to_excel/
│
├── fortigate_policy_to_excel.py   # [핵심] 변환 엔진 및 GUI 통합 단일 스크립트
├── README.md                      # [문서] 사용자 매뉴얼 및 가이드
│
└── <호스트네임>/                  # [결과물] 변환 완료 시 생성되는 결과 디렉토리
    ├── _TOTAL_SUMMARY.xlsx        # 전체 VDOM 정책 및 객체 수 집계 총괄 요약
    ├── root.xlsx                  # root VDOM 엑셀 (5대 정책 개별 시트 분리)
    ├── DMZ.xlsx                   # DMZ VDOM 엑셀
    ├── IDC.xlsx                   # IDC VDOM 엑셀
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
