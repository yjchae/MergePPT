# CLAUDE.md

이 파일은 Claude Code(claude.ai/code)가 이 저장소에서 작업할 때 참고하는 가이드입니다.

## 프로젝트 개요

**MergePPT(PPT 병합기)**는 여러 PowerPoint 파일(.ppt, .pptx)을 하나의 프레젠테이션으로 병합하는 한국어 데스크탑 GUI 앱입니다.

- **언어:** Python 3.11+
- **GUI:** PySide6 (Qt for Python)
- **핵심 라이브러리:** python-pptx, lxml
- **외부 의존성:** LibreOffice (`.ppt` → `.pptx` 변환에 필요)

## 주요 명령어

**앱 실행:**
```bash
python mergeppt.py
```

**의존성 설치:**
```bash
pip install PySide6 python-pptx lxml
```

**macOS 빌드 (DMG):**
```bash
bash build_mac.sh
# 결과물: dist/PPT병합기_mac.dmg
```

**Windows 빌드 (설치 파일):**
```batch
build_win.bat
# 결과물: dist/PPT병합기_Setup_v1.0.exe
```

자동화된 테스트는 없으며, 테스트용 파일은 `test/` 디렉토리에 있습니다.

## 코드 구조

앱 전체가 `mergeppt.py` 단일 파일(약 819줄)에 구현되어 있습니다.

### 클래스 구성 (파일 내 순서)

| 클래스 | 역할 |
|---|---|
| `MergeListWidget` | 병합 대기열 리스트 |
| `SearchResultsListWidget` | 파일 검색 결과 리스트 |
| `FileSearchWorker` | 백그라운드 파일 탐색 스레드 |
| `PPTItemDelegate` | 병합 목록 아이템 커스텀 렌더러 |
| `PPTMergerApp` | 메인 윈도우 |

---

#### `MergeListWidget(QListWidget)`
병합 대기열 리스트 위젯. 두 가지 드래그앤드롭을 지원합니다:
- **내부 이동** (`MoveAction`): 아이템 순서 변경
- **외부 파일 드롭** (`CopyAction`): OS 파인더/탐색기 또는 검색 패널에서 파일 추가

파일이 드롭되면 `files_dropped(list, int)` 시그널로 파일 경로와 삽입 위치를 전달합니다.

---

#### `SearchResultsListWidget(QListWidget)`
파일 검색 결과 패널. 드래그 전용(드롭 불가)이며, 아이템을 파일 URL MIME 데이터로 변환해 `MergeListWidget`으로 드래그할 수 있습니다.

---

#### `FileSearchWorker(QThread)`
백그라운드 디렉토리 탐색 스레드. .ppt/.pptx 파일을 비동기로 검색합니다.
- 350ms 디바운스 타이머 적용
- `threading.Event`로 취소 가능
- 한국어 파일명 처리를 위한 NFC 유니코드 정규화 적용

---

#### `PPTItemDelegate(QStyledItemDelegate)`
병합 목록 아이템의 커스텀 렌더러.
- 드래그 핸들(⠿), 순서 번호, 파일 타입 배지(PPTX/PPT), 삭제 버튼(✕) 표시
- PPTX는 파란색, PPT는 주황색 배지로 구분
- `editorEvent`에서 삭제 버튼 클릭 감지 처리

---

#### `PPTMergerApp(QWidget)`
메인 윈도우. 좌측(병합 대기열) / 우측(파일 검색 패널) 분할 레이아웃.

주요 메서드:

- **`merge_files()`** — 병합 전체 흐름 조율. `.ppt` 파일은 LibreOffice로 변환 후 `merge_pptx_files()` 호출
- **`convert_ppt_to_pptx()`** — LibreOffice 헤드리스 모드 실행(`soffice --headless --convert-to pptx`). LibreOffice 탐색 순서: 번들 앱 경로 → macOS 앱 번들 → Windows 레지스트리 → 시스템 PATH
- **`merge_pptx_files()`** — 실제 병합 처리. 새 프레젠테이션 생성 후 각 소스 파일에서 슬라이드를 복사하고 파일 사이에 검정 구분 슬라이드 삽입
- **`copy_slide()`** — OOXML 슬라이드 저수준 복사. 모든 관계(미디어, 하이퍼링크)를 복사하고 XML 내 r:id/r:embed/r:link 속성을 리매핑. cNvPr ID는 순차 재부여, paraId/textId는 32비트 랜덤 hex로 재부여
- **`remove_slide_background()`** — 변환된 .ppt 파일의 슬라이드에서 `<p:bg>` XML 요소를 제거해 첫 번째 파일의 테마가 전체에 적용되도록 함

## 핵심 기술 사항

- **슬라이드 병합**은 python-pptx의 고수준 API가 아닌 lxml을 통해 OOXML XML을 직접 조작합니다. 관계 ID 리매핑의 정확성을 위한 설계입니다.
- **LibreOffice**는 `.ppt` 지원에 필수입니다. 배포 빌드에는 번들로 포함되며, 개발 환경에서는 표준 설치 경로 또는 시스템 PATH에서 자동 탐지합니다.
- **출력 파일**은 `Merged_PPT_<6자리-uuid>.pptx`로 자동 명명되며, 기본 저장 위치는 바탕화면입니다. 저장 후 파일이 자동으로 열립니다.
- **다크 테마**는 `QApplication.setStyle("Fusion")`과 커스텀 다크 팔레트로 전역 적용됩니다.
- UI는 전체 한국어로 구성되어 있습니다.

## 빌드 & 배포

| 플랫폼 | 빌드 스크립트 | 결과물 |
|---|---|---|
| macOS | `build_mac.sh` | `dist/PPT병합기_mac.dmg` |
| Windows | `build_win.bat` | `dist/PPT병합기_Setup_v1.0.exe` |

- PyInstaller 설정: `PPTMerger.spec`
- Windows 인스톨러 설정: `installer.iss` (Inno Setup)
- GitHub Actions 자동 빌드: `v*` 태그 푸시 시 macOS/Windows 동시 빌드 후 GitHub Release 생성
