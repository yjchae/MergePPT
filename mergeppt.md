# mergeppt.py 로직 문서

## 파일 개요

| 항목 | 내용 |
|---|---|
| 파일명 | `mergeppt.py` |
| 역할 | PPT 병합기 앱 전체 (단일 파일 구조) |
| 언어 | Python 3.11+ |
| 프레임워크 | PySide6 (Qt for Python) |
| 주요 의존성 | `python-pptx`, `pptx.dml.color.RGBColor`, `lxml` |
| 실행 | `python mergeppt.py` |

---

## 전체 구조 (클래스 순서)

```
mergeppt.py
├── MergeListWidget          # 병합 대기열 리스트 위젯
├── SearchResultsListWidget  # 파일 검색 결과 리스트 위젯
├── FileSearchWorker         # 백그라운드 파일 탐색 스레드
├── PPTItemDelegate          # 병합 목록 아이템 커스텀 렌더러
└── PPTMergerApp             # 메인 윈도우 (앱 진입점)
```

---

## 클래스별 로직

### `MergeListWidget(QListWidget)`

병합 대기열을 표시하는 리스트 위젯.

- 내부 이동 (`Qt.MoveAction`): 같은 리스트 내 순서 변경
- 외부 파일 드롭 (`Qt.CopyAction`): OS 파인더/탐색기 또는 검색 패널에서 파일 추가
- 드롭 시 `.ppt` / `.pptx` 확장자 필터링 후 `files_dropped` 시그널 발행

**시그널**
```python
files_dropped = Signal(list)  # (file_paths, insert_row) 튜플
```

---

### `SearchResultsListWidget(QListWidget)`

파일 검색 결과를 표시하는 리스트 위젯.

- 드래그 전용 (`DragOnly`) — 드롭 불가
- 아이템을 `QUrl` MIME 데이터로 변환해 `MergeListWidget`으로 드래그 가능

---

### `FileSearchWorker(QThread)`

백그라운드 디렉토리 탐색 스레드.

- `start_search(root, keyword)`: 이전 탐색 취소 후 새 탐색 시작
- 숨김 폴더(`.`으로 시작) 건너뜀
- NFC 유니코드 정규화 → 한국어 파일명 정상 검색
- 350ms 디바운스 타이머 (PPTMergerApp에서 관리)
- `threading.Event`로 탐색 취소 가능
- 완료 시 `results_ready(list)` 시그널로 `(상대경로, 절대경로)` 리스트 전달

---

### `PPTItemDelegate(QStyledItemDelegate)`

병합 리스트 아이템 커스텀 렌더러.

| 요소 | 내용 |
|---|---|
| 드래그 핸들 | `⠿` (좌측) |
| 순서 번호 | 1-based 인덱스 |
| 파일 타입 배지 | PPTX → 파란색, PPT → 주황색 |
| 파일명 | 말줄임(`…`) 처리 |
| 삭제 버튼 | `✕` (우측, 호버 시 표시) |

- `editorEvent`에서 삭제 버튼 클릭 감지 → `delete_requested(row)` 시그널 발행

---

### `PPTMergerApp(QWidget)`

메인 윈도우. 모든 UI와 병합 로직 포함.

#### 인스턴스 변수 (상태 및 기본값)

| 변수 | 타입 | 기본값 | 설명 |
|---|---|---|---|
| `bg_color` | `QColor` | `#000000` (검정) | 병합 후 적용할 배경색 |
| `text_color` | `QColor` | `#FFFFFF` (흰색) | 글자색 일괄 적용 시 사용 |
| `slide_ratio` | `str` | `"16:9"` | 슬라이드 비율 |
| `text_valign` | `str` | `"top"` | 텍스트 세로 위치 (`top`/`center`/`bottom`) |
| `text_margin_pt` | `int` | `20` | 텍스트 여백 (pt 단위) |

#### UI 레이아웃

```
QVBoxLayout (root)
├── 헤더 행 (제목 + 힌트)
└── QSplitter (좌:우 = 640:420)
    ├── 좌측 (병합 패널)
    │   ├── MergeListWidget
    │   ├── 슬라이드 설정 행
    │   │   ├── 배경색 버튼 (bgColorBtn)
    │   │   ├── 글자색 버튼 (textColorBtn) + "일괄 적용" 체크박스 (textColorChk)
    │   │   ├── 슬라이드 비율: [16:9] [4:3]
    │   ├── 텍스트 위치 행
    │   │   ├── [상단] [가운데] [하단] 토글 버튼
    │   │   └── 여백 스핀박스 (0~300 pt)
    │   └── 하단 버튼 행
    │       ├── 전체 초기화 버튼
    │       └── 최종 파일로 합치기 버튼
    └── 우측 (파일 검색 패널)
        ├── 검색 폴더 입력 + 탐색 버튼
        ├── 파일명 검색 입력
        ├── 검색 결과 카운트 레이블
        └── SearchResultsListWidget
```

---

## 메서드 로직

### `merge_ppts()`

병합 전체 흐름 조율.

```
1. 리스트가 비어있으면 경고 후 종료
2. 저장 경로 선택 (기본: Desktop/Merged_PPT_<6자리uuid>.pptx)
3. 임시 디렉토리 생성
4. _convert_ppt_files() → .ppt 파일 .pptx 변환
5. 첫 번째 파일로 기본 Presentation 생성
6. 나머지 파일 순서대로:
   a. _add_divider_slide() → 선택한 배경색으로 구분 슬라이드 삽입
   b. 각 슬라이드를 _copy_slide()로 복사
7. 슬라이드 크기 설정
   - 16:9: 12192000 x 6858000 EMU
   - 4:3 : 9144000 x 6858000 EMU
8. _clean_slide_masters() → 마스터/레이아웃 이미지 제거
9. 모든 슬라이드 후처리:
   a. _remove_background_pictures() → 슬라이드 내 배경 이미지 제거
   b. _fit_text_shapes() → 가장 큰 텍스트 박스 위치·정렬
10. _apply_background_to_all_slides() → 배경색 일괄 적용
11. textColorChk 체크 시 _apply_text_color_to_all_slides() → 글자색 일괄 적용
12. 저장 후 파일 자동 오픈
13. 임시 디렉토리 삭제
```

---

### `_get_soffice()` (static)

LibreOffice 실행 파일 경로 탐색.

```
1. PyInstaller 번들 내부 경로 (배포 빌드)
2. macOS: /Applications/LibreOffice.app/...
3. Windows: 레지스트리 → 표준 설치 경로
4. 시스템 PATH (shutil.which)
```

---

### `_convert_ppt_files(file_paths, tmp_dir)`

`.ppt` 파일을 LibreOffice 헤드리스 모드로 일괄 변환.

```
soffice --headless --convert-to pptx --outdir <tmp_dir> <ppt_files...>
```

- 변환 후 `_strip_slide_backgrounds()` 호출 → 변환 파일의 `<p:bg>` 제거
- `.pptx` 파일은 그대로 경로 반환

---

### `_copy_slide(dest_prs, src_slide)`

소스 슬라이드를 목적 프레젠테이션에 복사 (OOXML 저수준 처리).

```
1. 빈 슬라이드 추가 (layout 6 또는 마지막 레이아웃)
2. slideLayout/notesSlide 제외한 모든 관계(rel) 복사
3. 구 rId → 새 rId 매핑 테이블 구성
4. src_slide의 spTree를 deepcopy하여 dest_slide에 교체
5. r:embed / r:id / r:link 속성값 매핑 치환
6. _reassign_ids() → ID 충돌 방지
```

> 배경 복사는 수행하지 않음. 병합 후 `_apply_background_to_all_slides()`에서 일괄 처리.

---

### `_add_divider_slide(prs, color)`

파일 사이 구분 슬라이드 삽입.

- 모든 도형 제거 후 선택한 배경색(`color`)으로 단색 `<p:bg>` 생성
- 항상 검정이 아닌 **사용자가 선택한 배경색** 사용

---

### `_clean_slide_masters(prs)` (static)

슬라이드 마스터 및 레이아웃에서 이미지 도형을 제거.

- 슬라이드 마스터 spTree + 모든 레이아웃 spTree 대상
- `blipFill`이 있고 텍스트(`a:t`)가 없는 요소 → 배경 장식으로 판단해 제거
- 마스터/레이아웃의 `<p:bg>` 배경 요소도 제거
- **근본 원인 차단**: 슬라이드 마스터에 이미지가 있으면 슬라이드 배경색을 설정해도 마스터 이미지가 그 위에 렌더링됨

---

### `_remove_background_pictures(slide)` (static)

슬라이드 개별 shape tree에서 배경 이미지 도형 제거.

- 대상 태그: `p:pic`, `p:sp`, `p:grpSp`
- 제거 조건: `blipFill`(이미지 데이터) 있음 **AND** `a:t`(텍스트) 없음
- 텍스트가 함께 있는 도형(이미지+캡션 등)은 보존

---

### `_fit_text_shapes(slide, slide_width, slide_height, valign, margin)` (static)

가장 큰 텍스트 박스(가사)만 위치·정렬 적용. 나머지(제목 등)는 건드리지 않음.

| 대상 | 처리 |
|---|---|
| 가장 큰 텍스트 박스 (width×height 최대) | 좌우 가운데 + valign 적용 |
| 나머지 텍스트 박스 | 위치·크기 일절 변경 없음 |

**valign별 top 계산**

| valign | top 값 |
|---|---|
| `top` | `margin` |
| `center` | `(slide_height - h) // 2` |
| `bottom` | `slide_height - h - margin` |

- 크기: 슬라이드 경계 초과 시에만 축소. 절대 원본보다 키우지 않음
- `margin` 단위: EMU (`text_margin_pt × 12700`)

---

### `_apply_background_to_all_slides(prs, color)`

모든 슬라이드에 단색 `<p:bg>` 삽입 (기존 배경 교체).

---

### `_apply_text_color_to_all_slides(prs, color)`

모든 슬라이드의 텍스트 런(run)에 글자색 일괄 적용.

- `textColorChk` 체크박스가 **체크된 경우에만** 호출됨
- 미체크 시 원본 색상 유지 (한글/영어 개별 색상 보존)

---

### `_reassign_ids(slide)`

슬라이드 내 XML 요소의 ID 재부여 → 병합 후 ID 충돌 방지.

| 대상 속성 | 처리 방식 |
|---|---|
| `cNvPr.id` | 1부터 순차 증가 |
| `paraId` | 32비트 랜덤 hex |
| `textId` | 32비트 랜덤 hex |

---

## UI 설정 항목 요약

| 항목 | 위젯 | 기본값 | 설명 |
|---|---|---|---|
| 배경색 | `QPushButton` (색상 미리보기) | `#000000` | 클릭 시 컬러 피커 |
| 글자색 | `QPushButton` (■ 견본) | `#FFFFFF` | 클릭 시 컬러 피커 |
| 글자색 일괄 적용 | `QCheckBox` | 미체크 | 체크 시에만 글자색 변경 |
| 슬라이드 비율 | 토글 버튼 `16:9` / `4:3` | `16:9` | 최종 슬라이드 크기 결정 |
| 텍스트 위치 | 토글 버튼 `상단` / `가운데` / `하단` | `상단` | 가장 큰 텍스트 박스 세로 정렬 |
| 여백 | `QSpinBox` (0~300 pt) | `20 pt` | 텍스트 박스와 슬라이드 경계 간격 |

---

## 슬라이드 크기 상수 (EMU)

| 비율 | width | height | 실제 크기 |
|---|---|---|---|
| 16:9 | 12,192,000 | 6,858,000 | 33.87cm × 19.05cm |
| 4:3 | 9,144,000 | 6,858,000 | 25.40cm × 19.05cm |

---

## OOXML 네임스페이스 참조

| 접두사 | URI |
|---|---|
| `p:` | `http://schemas.openxmlformats.org/presentationml/2006/main` |
| `a:` | `http://schemas.openxmlformats.org/drawingml/2006/main` |
| `r:` | `http://schemas.openxmlformats.org/officeDocument/2006/relationships` |

---

## 건너뛰는 관계 타입 (`_SKIP_RELTYPES`)

슬라이드 복사 시 목적 프레젠테이션이 자체 관리하므로 복사하지 않음.

```
.../relationships/slideLayout
.../relationships/notesSlide
```
