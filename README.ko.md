# where-file

> "내 파일 어딨지?" — Windows PC에서 열린 모든 파일을 한눈에 확인하세요.

설치 필요 없음. 외부 의존성 없음. 그냥 실행하면 됩니다.

![Python](https://img.shields.io/badge/Python-3.6+-blue?logo=python&logoColor=white)
![Platform](https://img.shields.io/badge/Platform-Windows-0078D6?logo=windows&logoColor=white)
![Dependencies](https://img.shields.io/badge/Dependencies-Zero-brightgreen)
![License](https://img.shields.io/badge/License-MIT-yellow)

![where-file 웹 대시보드](image/image_1.png)

## 이게 뭔가요?

PowerPoint, Word, Excel, VS Code, 메모장 등 **모든 앱에서 열린 파일**을 한 곳에서 보여주는 Windows 유틸리티입니다.

창 20개 띄워놓고 그 파일 어디있는지 못 찾은 적 있죠? **where-file**이 해결해 드립니다.

### 주요 기능

- **열린 파일 한눈에 보기** (Office, 에디터, 미디어 플레이어 등)
- **파일 경로** 또는 **폴더 경로** 원클릭 복사
- **70개 이상의 파일 형식** 지원 (Office, PDF, 이미지, 코드, 동영상, 음악...)
- **외부 의존성 제로** — 순수 Python 표준 라이브러리만 사용
- **두 가지 모드**: 웹 대시보드 또는 시스템 트레이

## 바로 시작하기

![시작 가이드](image/image_3.png)

### exe 파일로 바로 사용 (Python 없어도 됨)

[Releases 페이지](https://github.com/kangyekwon/where-file/releases/latest)에서 다운로드 후 더블클릭하면 끝!

| 파일 | 설명 |
|------|------|
| `where-file-web.exe` | 웹 대시보드 - 브라우저에서 열린 파일 확인 |
| `where-file-tray.exe` | 시스템 트레이 - Ctrl+Shift+F로 파일 경로 복사 |

### Python으로 실행

```bash
# 클론
git clone https://github.com/kangyekwon/where-file.git
cd where-file

# 실행 (웹 대시보드)
python server.py
```

브라우저가 자동으로 `http://localhost:8765`에서 열립니다.

## 두 가지 사용법

### 1. 웹 대시보드 (`server.py`)

```bash
python server.py
```

- 다크 테마의 깔끔한 대시보드
- 파일 타입별 필터링 (Office, PDF, Image, Code 등)
- 5초마다 자동 새로고침
- 경로 복사 / 폴더 복사 버튼

### 2. 시스템 트레이 (`app.py`)

![시스템 트레이 모드](image/image_2.png)

```bash
python app.py
```

- 시스템 트레이에 조용히 상주
- **`Ctrl+Shift+F`** — 현재 활성 창의 파일 경로를 즉시 클립보드에 복사
- 트레이 아이콘 좌클릭으로 열린 파일 목록 확인
- 빠르게 파일 경로만 가져올 때 최적

## 보너스: 우클릭 컨텍스트 메뉴

Windows 탐색기 우클릭 메뉴에 "Copy Full Path" 추가:

```
# 더블클릭으로 설치
install_menu.reg

# 더블클릭으로 제거
uninstall_menu.reg
```

## 지원 파일 형식

| 카테고리 | 확장자 |
|----------|--------|
| Office | `.pptx` `.docx` `.xlsx` `.hwp` `.hwpx` `.csv` |
| 문서 | `.pdf` `.txt` `.rtf` `.md` `.log` |
| 이미지 | `.png` `.jpg` `.gif` `.svg` `.webp` `.psd` `.ai` |
| 코드 | `.py` `.js` `.ts` `.java` `.cpp` `.go` `.rs` `.rb` `.php` |
| 동영상 | `.mp4` `.avi` `.mkv` `.mov` |
| 음악 | `.mp3` `.wav` `.flac` `.aac` |
| 압축 | `.zip` `.rar` `.7z` `.tar` `.gz` |
| 디자인 | `.sketch` `.fig` `.xd` `.indd` `.dwg` |
| 기타 | 총 70개 이상 지원 |

## 작동 원리

1. **Office 앱** (PowerPoint, Word, Excel): PowerShell을 통해 COM 객체 직접 쿼리
2. **기타 앱**: 실행 중인 프로세스의 커맨드라인에서 파일 경로 추출
3. 3초 캐시로 빠른 응답 유지

## 요구 사항

- **Windows** (Windows API 사용)
- **Python 3.6+** (pip 설치 불필요) — exe 사용 시 Python도 불필요
- **PowerShell 5.0+** (Windows 10/11에 기본 설치)

## 알려진 제한사항

- Windows 전용 (Windows API와 PowerShell에 의존)
- 브라우저 탭에서 열린 파일 (예: Chrome/Edge에서 열린 PDF)은 감지 불가
- 일부 앱은 프로세스 커맨드라인에 파일 경로를 노출하지 않을 수 있음

## 라이선스

MIT
