import sys, os, uuid, copy, random, subprocess, tempfile, threading
from pptx import Presentation
from lxml import etree
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout,
    QListWidget, QListWidgetItem, QPushButton, QMessageBox, QFileDialog,
    QStyledItemDelegate, QStyle, QLabel, QLineEdit, QSplitter, QFrame
)
from PySide6.QtCore import Qt, QSize, QRect, QRectF, QEvent, Signal, QUrl, QThread, QTimer
from PySide6.QtGui import (
    QPainter, QColor, QFont, QBrush, QPen, QFontMetrics, QPainterPath, QDrag
)
from PySide6.QtCore import QMimeData


# ── 병합 리스트 위젯 (내부 재정렬 + 외부 URL 드롭 지원) ──────────────────
class MergeListWidget(QListWidget):
    files_dropped = Signal(list)   # list of file path strings

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setDragDropMode(QListWidget.DragDrop)
        self.setDefaultDropAction(Qt.MoveAction)
        self.setMouseTracking(True)
        self.viewport().setMouseTracking(True)

    def dragEnterEvent(self, event):
        if event.source() is self:
            event.setDropAction(Qt.MoveAction)
            event.accept()
        elif event.mimeData().hasUrls():
            event.setDropAction(Qt.CopyAction)
            event.accept()
        else:
            event.ignore()

    def dragMoveEvent(self, event):
        if event.source() is self:
            event.setDropAction(Qt.MoveAction)
            event.accept()
        elif event.mimeData().hasUrls():
            event.setDropAction(Qt.CopyAction)
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):
        if event.source() is self:
            super().dropEvent(event)
        elif event.mimeData().hasUrls():
            drop_row = self._row_at(event.position().toPoint())
            paths = [
                url.toLocalFile() for url in event.mimeData().urls()
                if url.toLocalFile().lower().endswith(('.ppt', '.pptx'))
            ]
            if paths:
                self.files_dropped.emit((paths, drop_row))
            event.accept()
        else:
            event.ignore()

    def _row_at(self, pos):
        item = self.itemAt(pos)
        if item:
            return self.row(item)
        return self.count()


# ── 검색 결과 리스트 위젯 (드래그 시 URL 미임 데이터 제공) ─────────────────
class SearchResultsListWidget(QListWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setDragEnabled(True)
        self.setDragDropMode(QListWidget.DragOnly)
        self.setSelectionMode(QListWidget.ExtendedSelection)

    def startDrag(self, supported_actions):
        items = self.selectedItems()
        if not items:
            return
        urls = [
            QUrl.fromLocalFile(it.data(Qt.UserRole))
            for it in items
            if it.data(Qt.UserRole)
        ]
        if not urls:
            return
        mime = QMimeData()
        mime.setUrls(urls)
        drag = QDrag(self)
        drag.setMimeData(mime)
        drag.exec(Qt.CopyAction)


# ── 백그라운드 파일 검색 스레드 ────────────────────────────────────────────
class FileSearchWorker(QThread):
    # match_type: 'name' | 'content'
    partial_results = Signal(list)   # 1단계: 파일명 일치 결과 (빠르게)
    results_ready   = Signal(list)   # 2단계: 내용 검색까지 완료된 최종 결과
    progress_msg    = Signal(str)    # 진행 상태 메시지

    def __init__(self, parent=None):
        super().__init__(parent)
        self._root  = ''
        self._query = ''
        self._stop_event = threading.Event()

    def start_search(self, root, query):
        self._stop_event.set()
        self.wait(400)
        self._root  = root
        self._query = query
        self._stop_event.clear()
        self.start()

    def stop(self):
        self._stop_event.set()

    def run(self):
        q = self._query.lower().strip()

        # ── 1단계: 파일 시스템 탐색 + 파일명 매칭 ──
        name_results     = []   # (rel, full, 'name')
        content_pending  = []   # .pptx 파일 중 파일명 미일치 → 내용 검색 후보
        try:
            for dirpath, dirnames, filenames in os.walk(self._root):
                if self._stop_event.is_set():
                    return
                dirnames[:] = [d for d in dirnames if not d.startswith('.')]
                for fn in filenames:
                    if self._stop_event.is_set():
                        return
                    if fn.lower().endswith(('.ppt', '.pptx')):
                        full = os.path.join(dirpath, fn)
                        rel  = os.path.relpath(full, self._root)
                        if not q or q in fn.lower():
                            name_results.append((rel, full, 'name'))
                        elif fn.lower().endswith('.pptx'):
                            content_pending.append((rel, full))
        except Exception:
            pass

        if self._stop_event.is_set():
            return

        # 파일명 결과 먼저 바로 전달
        sorted_name = sorted(name_results, key=lambda x: x[0].lower())
        self.partial_results.emit(sorted_name)

        # ── 2단계: 내용 검색 (쿼리가 있을 때만) ──
        content_results = []
        if q and content_pending:
            total = len(content_pending)
            for i, (rel, full) in enumerate(content_pending, 1):
                if self._stop_event.is_set():
                    return
                self.progress_msg.emit(f"내용 검색 중... ({i}/{total})")
                if self._search_pptx_content(full, q, self._stop_event):
                    content_results.append((rel, full, 'content'))

        if self._stop_event.is_set():
            return

        all_results = sorted_name + sorted(content_results, key=lambda x: x[0].lower())
        self.results_ready.emit(all_results)

    @staticmethod
    def _search_pptx_content(path, query, stop_event):
        """슬라이드 텍스트에 query가 포함되어 있으면 True 반환."""
        try:
            from pptx import Presentation
            prs = Presentation(path)
            for slide in prs.slides:
                if stop_event.is_set():
                    return False
                for shape in slide.shapes:
                    if not shape.has_text_frame:
                        continue
                    for para in shape.text_frame.paragraphs:
                        text = ''.join(run.text for run in para.runs).lower()
                        if query in text:
                            return True
        except Exception:
            pass
        return False


# ── 아이템 델리게이트 (병합 리스트용) ─────────────────────────────────────
class PPTItemDelegate(QStyledItemDelegate):
    delete_requested = Signal(int)

    _H       = 52
    _M       = 3
    _HANDLE  = 24
    _NUM     = 28
    _BADGE   = 48
    _XBTN    = 32

    def paint(self, painter, option, index):
        painter.save()
        painter.setRenderHint(QPainter.Antialiasing)

        r = QRectF(option.rect.adjusted(self._M, self._M, -self._M, -self._M))
        selected = bool(option.state & QStyle.State_Selected)
        hovered  = bool(option.state & QStyle.State_MouseOver)

        if selected:
            bg, border = QColor("#2a3f5f"), QColor("#4a7abf")
        elif hovered:
            bg, border = QColor("#353550"), QColor("#4a4a66")
        else:
            bg, border = QColor("#2e2e42"), QColor("#3a3a52")

        path = QPainterPath()
        path.addRoundedRect(r, 8, 8)
        painter.fillPath(path, bg)
        painter.setPen(QPen(border, 1))
        painter.drawPath(path)

        file_path = index.data(Qt.UserRole) or ""
        name = os.path.basename(file_path)
        ext  = os.path.splitext(file_path)[1].lower()
        ri   = option.rect
        cy   = ri.center().y()

        # 햄버거 핸들
        f = QFont(); f.setPixelSize(18)
        painter.setFont(f)
        painter.setPen(QColor("#777" if hovered else "#555"))
        painter.drawText(
            QRect(ri.x() + 8, ri.y(), self._HANDLE, ri.height()),
            Qt.AlignCenter, "⠿"
        )

        # 번호
        f.setPixelSize(13)
        painter.setFont(f)
        painter.setPen(QColor("#888"))
        painter.drawText(
            QRect(ri.x() + 8 + self._HANDLE, ri.y(), self._NUM, ri.height()),
            Qt.AlignRight | Qt.AlignVCenter, f"{index.row() + 1}."
        )

        # X 버튼
        xr = self._xbtn_rect(ri)
        x_color = QColor("#ff6060") if hovered else QColor(180, 80, 80, 160)
        f.setPixelSize(13); f.setBold(True)
        painter.setFont(f)
        painter.setPen(x_color)
        painter.drawText(xr, Qt.AlignCenter, "✕")

        # 뱃지
        badge_right = xr.left() - 8
        badge_rect  = QRect(badge_right - self._BADGE, cy - 11, self._BADGE, 22)
        bp = QPainterPath()
        bp.addRoundedRect(QRectF(badge_rect), 4, 4)
        painter.fillPath(bp, QColor("#1a6bbf") if ext == '.pptx' else QColor("#c07000"))
        f.setPixelSize(11); f.setBold(True)
        painter.setFont(f)
        painter.setPen(QColor("white"))
        painter.drawText(badge_rect, Qt.AlignCenter, ext.lstrip('.').upper())

        # 파일명
        name_x = ri.x() + 8 + self._HANDLE + self._NUM + 8
        name_w = badge_rect.left() - name_x - 8
        f.setPixelSize(14); f.setBold(False)
        painter.setFont(f)
        painter.setPen(QColor("#e8e8f0"))
        elided = QFontMetrics(painter.font()).elidedText(name, Qt.ElideMiddle, name_w)
        painter.drawText(
            QRect(name_x, ri.y(), name_w, ri.height()),
            Qt.AlignLeft | Qt.AlignVCenter, elided
        )

        painter.restore()

    def _xbtn_rect(self, item_rect):
        r = item_rect
        return QRect(r.right() - self._XBTN - 4, r.y() + (r.height() - 26) // 2, 26, 26)

    def sizeHint(self, option, index):
        return QSize(0, self._H)

    def editorEvent(self, event, model, option, index):
        if event.type() == QEvent.Type.MouseButtonRelease:
            r = option.rect.adjusted(self._M, self._M, -self._M, -self._M)
            if self._xbtn_rect(r).contains(event.pos()):
                self.delete_requested.emit(index.row())
                return True
        return super().editorEvent(event, model, option, index)


# ── 메인 앱 ────────────────────────────────────────────────────────────────
class PPTMergerApp(QWidget):
    def __init__(self):
        super().__init__()
        self._search_timer  = QTimer(self)
        self._search_timer.setSingleShot(True)
        self._search_timer.setInterval(350)
        self._search_timer.timeout.connect(self._do_search)
        self._search_worker = FileSearchWorker(self)
        self._search_worker.partial_results.connect(self._on_partial_results)
        self._search_worker.results_ready.connect(self._on_search_results)
        self._search_worker.progress_msg.connect(
            lambda msg: self.resultCountLabel.setText(msg) if hasattr(self, 'resultCountLabel') else None
        )
        self.initUI()

    def initUI(self):
        self.setWindowTitle('PPT 병합기 ')
        self.resize(1080, 620)
        self.setAcceptDrops(True)
        self.setStyleSheet("background-color: #1e1e2e;")

        root = QVBoxLayout(self)
        root.setContentsMargins(12, 10, 12, 12)
        root.setSpacing(6)

        # 제목 + 힌트를 한 줄로
        header_row = QHBoxLayout()
        header_row.setSpacing(12)
        title = QLabel("PPT 병합기")
        title.setStyleSheet("color: #c0c0d8; font-size: 13px; font-weight: bold;")
        header_row.addWidget(title)
        hint = QLabel("드래그 앤 드롭  •  순서 드래그로 변경  •  우측 검색 결과 드래그하여 추가")
        hint.setStyleSheet("color: #555; font-size: 11px;")
        header_row.addWidget(hint, 1)
        root.addLayout(header_row)

        # ── 좌우 분할 스플리터 ──
        splitter = QSplitter(Qt.Horizontal)
        splitter.setHandleWidth(8)
        splitter.setStyleSheet("""
            QSplitter::handle { background: #2e2e42; border-radius: 4px; }
            QSplitter::handle:hover { background: #5555aa; }
        """)

        # ── 좌측: 병합 리스트 ──
        left_container = QWidget()
        left_container.setStyleSheet("background: transparent;")
        left_layout = QVBoxLayout(left_container)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(10)

        self.listWidget = MergeListWidget()
        self.listWidget.files_dropped.connect(self._on_files_dropped_to_list)
        self.listWidget.setStyleSheet("""
            QListWidget {
                background: #252535;
                border: 2px dashed #44445a;
                border-radius: 12px;
                padding: 4px;
                outline: none;
            }
            QListWidget::item { background: transparent; border: none; }
        """)
        self.delegate = PPTItemDelegate()
        self.delegate.delete_requested.connect(self._delete_row)
        self.listWidget.setItemDelegate(self.delegate)
        self.listWidget.model().rowsMoved.connect(
            lambda: self.listWidget.viewport().update()
        )
        left_layout.addWidget(self.listWidget)

        # 하단 버튼
        btn_row = QHBoxLayout()
        btn_row.setSpacing(12)

        self.clearButton = QPushButton("전체 초기화")
        self.clearButton.setFixedHeight(46)
        self.clearButton.setCursor(Qt.PointingHandCursor)
        self.clearButton.setStyleSheet("""
            QPushButton {
                color: #aaa; background: #2e2e42;
                border: 1px solid #44445a; border-radius: 10px; font-size: 14px;
            }
            QPushButton:hover { color: #fff; background: #3a3a55; border-color: #6666aa; }
            QPushButton:pressed { background: #2a2a45; }
        """)
        self.clearButton.clicked.connect(self.listWidget.clear)
        btn_row.addWidget(self.clearButton, 1)

        self.mergeButton = QPushButton("최종 파일로 합치기")
        self.mergeButton.setFixedHeight(46)
        self.mergeButton.setCursor(Qt.PointingHandCursor)
        self.mergeButton.setStyleSheet("""
            QPushButton {
                color: #fff;
                background: qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #0063cc,stop:1 #0096ff);
                border: none; border-radius: 10px;
                font-size: 16px; font-weight: bold;
            }
            QPushButton:hover {
                background: qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #0074ee,stop:1 #22aaff);
            }
            QPushButton:pressed {
                background: qlineargradient(x1:0,y1:0,x2:1,y2:0,stop:0 #0052aa,stop:1 #007acc);
            }
            QPushButton:disabled { background: #3a3a50; color: #666; }
        """)
        self.mergeButton.clicked.connect(self.merge_ppts)
        btn_row.addWidget(self.mergeButton, 3)

        left_layout.addLayout(btn_row)
        splitter.addWidget(left_container)

        # ── 우측: 파일 검색 패널 ──
        right_container = QWidget()
        right_container.setObjectName("searchPanel")
        right_container.setStyleSheet("""
            QWidget#searchPanel {
                background: #23233a;
                border: 1px solid #35355a;
                border-radius: 12px;
            }
        """)
        right_layout = QVBoxLayout(right_container)
        right_layout.setContentsMargins(14, 14, 14, 14)
        right_layout.setSpacing(8)

        # 검색 패널 제목
        search_title = QLabel("파일 검색")
        search_title.setStyleSheet(
            "color: #e0e0f0; font-size: 15px; font-weight: bold;"
            " background: transparent; border: none;"
        )
        right_layout.addWidget(search_title)

        sep = QFrame()
        sep.setFrameShape(QFrame.HLine)
        sep.setFixedHeight(1)
        sep.setStyleSheet("background: #35355a; border: none;")
        right_layout.addWidget(sep)

        # 폴더 경로
        lbl_folder = QLabel("검색 폴더")
        lbl_folder.setStyleSheet("color: #888; font-size: 11px; background: transparent; border: none;")
        right_layout.addWidget(lbl_folder)

        folder_row = QHBoxLayout()
        folder_row.setSpacing(6)
        self.folderInput = QLineEdit()
        self.folderInput.setPlaceholderText("최상위 폴더 경로...")
        self.folderInput.setStyleSheet(self._input_style())
        self.folderInput.returnPressed.connect(self._do_search)
        folder_row.addWidget(self.folderInput)

        browse_btn = QPushButton("…")
        browse_btn.setFixedSize(36, 32)
        browse_btn.setCursor(Qt.PointingHandCursor)
        browse_btn.setStyleSheet("""
            QPushButton {
                color: #bbb; background: #2e2e42;
                border: 1px solid #44445a; border-radius: 6px; font-size: 16px;
            }
            QPushButton:hover { color: #fff; background: #3a3a55; }
            QPushButton:pressed { background: #252540; }
        """)
        browse_btn.clicked.connect(self._browse_folder)
        folder_row.addWidget(browse_btn)
        right_layout.addLayout(folder_row)

        # 검색어 입력
        lbl_search = QLabel("파일명 검색")
        lbl_search.setStyleSheet("color: #888; font-size: 11px; background: transparent; border: none;")
        right_layout.addWidget(lbl_search)

        self.searchInput = QLineEdit()
        self.searchInput.setPlaceholderText("파일명을 입력하세요... (빈칸 = 전체)")
        self.searchInput.setStyleSheet(self._input_style())
        self.searchInput.textChanged.connect(self._on_search_changed)
        right_layout.addWidget(self.searchInput)

        # 결과 카운트
        self.resultCountLabel = QLabel("검색 결과: —")
        self.resultCountLabel.setStyleSheet(
            "color: #666; font-size: 11px; background: transparent; border: none;"
        )
        right_layout.addWidget(self.resultCountLabel)

        # 검색 결과 리스트
        self.searchList = SearchResultsListWidget()
        self.searchList.setStyleSheet("""
            QListWidget {
                background: #1e1e2e;
                border: 1px solid #35355a;
                border-radius: 8px;
                padding: 2px;
                outline: none;
            }
            QListWidget::item {
                color: #d0d0e8;
                padding: 5px 8px;
                border-bottom: 1px solid #252540;
                font-size: 12px;
                border-radius: 0px;
            }
            QListWidget::item:hover { background: #2a2a45; }
            QListWidget::item:selected { background: #2a3f5f; color: #ffffff; }
        """)
        self.searchList.itemDoubleClicked.connect(self._on_search_item_double_clicked)
        right_layout.addWidget(self.searchList, 1)

        drag_hint = QLabel("← 드래그하여 추가  |  더블클릭으로 추가  |  초록색 = 내용 일치")
        drag_hint.setAlignment(Qt.AlignCenter)
        drag_hint.setStyleSheet(
            "color: #4a4a6a; font-size: 10px; background: transparent; border: none;"
        )
        right_layout.addWidget(drag_hint)

        splitter.addWidget(right_container)
        splitter.setSizes([640, 420])
        splitter.setCollapsible(1, False)

        root.addWidget(splitter)

    # ── 스타일 헬퍼 ──────────────────────────────────────────────────────
    @staticmethod
    def _input_style():
        return """
            QLineEdit {
                color: #e0e0f0;
                background: #1e1e2e;
                border: 1px solid #44445a;
                border-radius: 6px;
                padding: 5px 8px;
                font-size: 12px;
            }
            QLineEdit:focus { border-color: #5555aa; }
        """

    # ── 폴더 탐색 ─────────────────────────────────────────────────────────
    def _browse_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "검색 폴더 선택")
        if folder:
            self.folderInput.setText(folder)
            self._do_search()

    # ── 검색 트리거 (디바운스) ────────────────────────────────────────────
    def _on_search_changed(self, _text):
        self._search_timer.stop()
        self._search_timer.start()

    def _do_search(self):
        root = self.folderInput.text().strip()
        if not root or not os.path.isdir(root):
            self.searchList.clear()
            self.resultCountLabel.setText("검색 결과: 폴더를 먼저 선택하세요")
            return
        self.resultCountLabel.setText("검색 중...")
        self._search_worker.start_search(root, self.searchInput.text().strip())

    def _populate_search_list(self, results):
        """results: list of (rel, full, match_type)"""
        self.searchList.clear()
        for rel, full, match_type in results:
            label = f"[내용]  {rel}" if match_type == 'content' else rel
            item = QListWidgetItem(label)
            item.setData(Qt.UserRole, full)
            item.setToolTip(f"{'내용 일치' if match_type == 'content' else '파일명 일치'}\n{full}")
            if match_type == 'content':
                item.setForeground(QColor("#7ddd88"))
            self.searchList.addItem(item)

    def _on_partial_results(self, results):
        """1단계: 파일명 일치 결과 즉시 표시 + 내용 검색 예고"""
        self._populate_search_list(results)
        n = len(results)
        q = self.searchInput.text().strip()
        if q:
            self.resultCountLabel.setText(
                f"파일명 일치: {n}개  |  내용 검색 중..."
            )
        else:
            self.resultCountLabel.setText(f"검색 결과: {n}개")

    def _on_search_results(self, results):
        """2단계: 내용 검색까지 완료된 최종 결과"""
        self._populate_search_list(results)
        n_name    = sum(1 for _, _, t in results if t == 'name')
        n_content = sum(1 for _, _, t in results if t == 'content')
        if not results:
            self.resultCountLabel.setText("검색 결과: 없음")
        elif n_content:
            self.resultCountLabel.setText(
                f"총 {len(results)}개  (파일명: {n_name}개  +  내용: {n_content}개)"
            )
        else:
            self.resultCountLabel.setText(f"검색 결과: {n_name}개")

    def _on_search_item_double_clicked(self, item):
        path = item.data(Qt.UserRole)
        if path:
            self.add_file(path)

    # ── 드래그앤 드롭 (윈도우 레벨 — 리스트 위젯 바깥 영역용) ──────────────
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    def dropEvent(self, event):
        for url in event.mimeData().urls():
            path = url.toLocalFile()
            if path.lower().endswith(('.ppt', '.pptx')):
                self.add_file(path)

    # ── 파일 추가 / 삭제 ──────────────────────────────────────────────────
    def add_file(self, file_path):
        item = QListWidgetItem()
        item.setData(Qt.UserRole, file_path)
        self.listWidget.addItem(item)

    def _add_files_at(self, paths, row):
        for i, path in enumerate(paths):
            item = QListWidgetItem()
            item.setData(Qt.UserRole, path)
            self.listWidget.insertItem(row + i, item)

    def _on_files_dropped_to_list(self, payload):
        paths, row = payload
        self._add_files_at(paths, row)

    def _delete_row(self, row):
        self.listWidget.takeItem(row)

    # ── 병합 ──────────────────────────────────────────────────────────────
    def merge_ppts(self):
        if self.listWidget.count() == 0:
            QMessageBox.warning(self, "파일 없음", "병합할 PPT 파일을 추가해주세요.")
            return

        file_paths = [
            self.listWidget.item(i).data(Qt.UserRole)
            for i in range(self.listWidget.count())
        ]

        default_name = f"Merged_PPT_{uuid.uuid4().hex[:6]}.pptx"
        default_path = os.path.join(os.path.expanduser("~/Desktop"), default_name)

        merged_file_path, _ = QFileDialog.getSaveFileName(
            self, "저장 위치 선택", default_path, "PowerPoint 파일 (*.pptx)"
        )
        if not merged_file_path:
            return

        self.mergeButton.setEnabled(False)
        self.mergeButton.setText('변환 및 병합 중...')

        tmp_dir = tempfile.mkdtemp()
        try:
            converted = self._convert_ppt_files(file_paths, tmp_dir)

            merged = Presentation(converted[0])
            for src_path in converted[1:]:
                self._add_black_slide(merged)
                src_prs = Presentation(src_path)
                for slide in src_prs.slides:
                    self._copy_slide(merged, slide)

            merged.save(merged_file_path)
            QMessageBox.information(self, "성공",
                f"병합 완료!\n'{merged_file_path}' 에 저장되었습니다.")
        except Exception as e:
            QMessageBox.critical(self, "병합 실패", f"오류가 발생했습니다:\n{str(e)}")
        finally:
            import shutil
            shutil.rmtree(tmp_dir, ignore_errors=True)
            self.mergeButton.setEnabled(True)
            self.mergeButton.setText('최종 파일로 합치기')

    # ── .ppt → .pptx 변환 ────────────────────────────────────────────────
    @staticmethod
    def _get_soffice():
        import shutil
        if getattr(sys, 'frozen', False):
            base = sys._MEIPASS
            if sys.platform == 'darwin':
                return os.path.join(base, 'LibreOffice.app', 'Contents', 'MacOS', 'soffice')
            else:
                return os.path.join(base, 'LibreOffice', 'program', 'soffice.exe')
        if sys.platform == 'darwin':
            p = '/Applications/LibreOffice.app/Contents/MacOS/soffice'
            return p if os.path.exists(p) else (shutil.which('soffice') or p)
        if sys.platform == 'win32':
            candidates = [
                r'C:\Program Files\LibreOffice\program\soffice.exe',
                r'C:\Program Files (x86)\LibreOffice\program\soffice.exe',
            ]
            try:
                import winreg
                for root in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
                    for sub in (r'SOFTWARE\LibreOffice\LibreOffice',
                                r'SOFTWARE\WOW6432Node\LibreOffice\LibreOffice'):
                        try:
                            key = winreg.OpenKey(root, sub)
                            idx = 0
                            while True:
                                try:
                                    ver = winreg.EnumKey(key, idx)
                                    vkey = winreg.OpenKey(key, ver)
                                    path, _ = winreg.QueryValueEx(vkey, 'Path')
                                    candidates.append(os.path.join(path, 'program', 'soffice.exe'))
                                    idx += 1
                                except OSError:
                                    break
                        except OSError:
                            pass
            except Exception:
                pass
            for p in candidates:
                if os.path.exists(p):
                    return p
            found = shutil.which('soffice') or shutil.which('soffice.exe')
            if found:
                return found
            return None
        found = shutil.which('soffice')
        return found or 'soffice'

    def _convert_ppt_files(self, file_paths, tmp_dir):
        ppt_files = [p for p in file_paths if p.lower().endswith('.ppt')]
        if ppt_files:
            soffice = self._get_soffice()
            if not soffice or not os.path.isfile(soffice):
                raise FileNotFoundError(
                    "LibreOffice를 찾을 수 없습니다.\n"
                    "LibreOffice를 설치한 뒤 다시 시도해주세요.\n"
                    "(https://www.libreoffice.org/download/libreoffice-fresh/)\n\n"
                    "설치 후에도 이 오류가 발생하면 LibreOffice 설치 경로의 'program' 폴더가\n"
                    "시스템 PATH에 등록되어 있는지 확인해주세요."
                )
            kwargs = {}
            if sys.platform == 'win32':
                kwargs['creationflags'] = subprocess.CREATE_NO_WINDOW
            subprocess.run(
                [soffice, '--headless', '--convert-to', 'pptx', '--outdir', tmp_dir] + ppt_files,
                check=True, capture_output=True, **kwargs
            )
        result = []
        for p in file_paths:
            if p.lower().endswith('.ppt'):
                base = os.path.splitext(os.path.basename(p))[0]
                result.append(os.path.join(tmp_dir, base + '.pptx'))
            else:
                result.append(p)
        return result

    # ── 슬라이드 복사 ──────────────────────────────────────────────────────
    _MEDIA_RELTYPES = {
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/media',
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio',
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/video',
    }

    def _copy_slide(self, dest_prs, src_slide):
        blank = dest_prs.slide_layouts[min(6, len(dest_prs.slide_layouts) - 1)]
        dest_slide = dest_prs.slides.add_slide(blank)

        rId_map = {}
        for rel_id, rel in src_slide.part.rels.items():
            if rel.is_external:
                try:
                    rId_map[rel_id] = dest_slide.part.relate_to(rel.target_ref, rel.reltype, is_external=True)
                except Exception:
                    pass
            elif rel.reltype in self._MEDIA_RELTYPES:
                try:
                    rId_map[rel_id] = dest_slide.part.relate_to(rel.target_part, rel.reltype)
                except Exception:
                    pass

        src_tree  = src_slide.shapes._spTree
        dest_tree = dest_slide.shapes._spTree
        for child in list(dest_tree):
            dest_tree.remove(child)
        for child in src_tree:
            dest_tree.append(copy.deepcopy(child))

        if rId_map:
            xml = etree.tostring(dest_tree, encoding='unicode')
            for old, new in rId_map.items():
                xml = xml.replace(f'r:embed="{old}"', f'r:embed="{new}"')
                xml = xml.replace(f'r:id="{old}"',    f'r:id="{new}"')
                xml = xml.replace(f'r:link="{old}"',  f'r:link="{new}"')
            new_tree = etree.fromstring(xml)
            dest_tree.getparent().replace(dest_tree, new_tree)

        self._reassign_ids(dest_slide)

        ns = 'http://schemas.openxmlformats.org/presentationml/2006/main'
        src_cSld  = src_slide._element.find(f'{{{ns}}}cSld')
        dest_cSld = dest_slide._element.find(f'{{{ns}}}cSld')
        if src_cSld is not None and dest_cSld is not None:
            src_bg = src_cSld.find(f'{{{ns}}}bg')
            if src_bg is not None:
                dest_bg = dest_cSld.find(f'{{{ns}}}bg')
                if dest_bg is not None:
                    dest_cSld.remove(dest_bg)
                dest_cSld.insert(0, copy.deepcopy(src_bg))

    def _add_black_slide(self, prs):
        slide = prs.slides.add_slide(prs.slide_layouts[min(6, len(prs.slide_layouts) - 1)])
        ns_p = 'http://schemas.openxmlformats.org/presentationml/2006/main'
        ns_a = 'http://schemas.openxmlformats.org/drawingml/2006/main'
        bg_xml = (
            f'<p:bg xmlns:p="{ns_p}" xmlns:a="{ns_a}">'
            '<p:bgPr><a:solidFill><a:srgbClr val="000000"/></a:solidFill>'
            '<a:effectLst/></p:bgPr></p:bg>'
        )
        cSld = slide._element.find(f'{{{ns_p}}}cSld')
        existing = cSld.find(f'{{{ns_p}}}bg')
        if existing is not None:
            cSld.remove(existing)
        cSld.insert(0, etree.fromstring(bg_xml))

    def _reassign_ids(self, slide):
        shape_id = 1
        for elem in slide._element.iter():
            local = elem.tag.split('}')[-1] if '}' in elem.tag else elem.tag
            if local == 'cNvPr':
                elem.set('id', str(shape_id))
                shape_id += 1
            if 'paraId' in elem.attrib:
                elem.set('paraId', format(random.randint(0x10000000, 0x7FFFFFFF), '08X'))
            if 'textId' in elem.attrib:
                elem.set('textId', format(random.randint(0x10000000, 0x7FFFFFFF), '08X'))


if __name__ == '__main__':
    app = QApplication(sys.argv)
    app.setStyleSheet("""
        QMessageBox {
            background-color: #2e2e42;
        }
        QMessageBox QLabel {
            color: #e8e8f0;
            font-size: 14px;
        }
        QMessageBox QPushButton {
            color: #e8e8f0;
            background-color: #3a3a55;
            border: 1px solid #55558a;
            border-radius: 6px;
            padding: 6px 18px;
            font-size: 13px;
            min-width: 72px;
        }
        QMessageBox QPushButton:hover {
            background-color: #4a4a70;
            border-color: #7777bb;
        }
        QMessageBox QPushButton:pressed {
            background-color: #252540;
        }
    """)
    ex = PPTMergerApp()
    ex.show()
    sys.exit(app.exec())
