import sys, os, uuid, copy, random, subprocess, tempfile
from pptx import Presentation
from lxml import etree
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout,
    QListWidget, QListWidgetItem, QPushButton, QMessageBox, QFileDialog,
    QStyledItemDelegate, QStyle
)
from PySide6.QtCore import Qt, QSize, QRect, QRectF, QEvent, Signal
from PySide6.QtGui import (
    QPainter, QColor, QFont, QBrush, QPen, QFontMetrics, QPainterPath
)


# ── 아이템 델리게이트 ─────────────────────────────────────────────
class PPTItemDelegate(QStyledItemDelegate):
    delete_requested = Signal(int)

    _H       = 52    # 아이템 높이
    _M       = 3     # 테두리 여백
    _HANDLE  = 24    # 햄버거 너비
    _NUM     = 28    # 번호 너비
    _BADGE   = 48    # 뱃지 너비
    _XBTN    = 32    # X 버튼 너비

    def paint(self, painter, option, index):
        painter.save()
        painter.setRenderHint(QPainter.Antialiasing)

        r = QRectF(option.rect.adjusted(self._M, self._M, -self._M, -self._M))
        selected = bool(option.state & QStyle.State_Selected)
        hovered  = bool(option.state & QStyle.State_MouseOver)

        # 배경
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
        ri   = option.rect  # int rect
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


# ── 메인 앱 ───────────────────────────────────────────────────────
class PPTMergerApp(QWidget):
    def __init__(self):
        super().__init__()
        self.initUI()

    def initUI(self):
        self.setWindowTitle('PPT 병합기')
        self.resize(680, 580)
        self.setAcceptDrops(True)
        self.setStyleSheet("background-color: #1e1e2e;")

        root = QVBoxLayout(self)
        root.setContentsMargins(20, 20, 20, 20)
        root.setSpacing(14)

        # 제목
        from PySide6.QtWidgets import QLabel
        title = QLabel("PPT 병합기")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("color: #e0e0f0; font-size: 20px; font-weight: bold;")
        root.addWidget(title)

        hint = QLabel("PPT / PPTX 파일을 아래로 드래그 앤 드롭  •  순서는 드래그로 변경")
        hint.setAlignment(Qt.AlignCenter)
        hint.setStyleSheet("color: #666; font-size: 12px;")
        root.addWidget(hint)

        # 리스트
        self.listWidget = QListWidget()
        self.listWidget.setDragDropMode(QListWidget.InternalMove)
        self.listWidget.setMouseTracking(True)
        self.listWidget.viewport().setMouseTracking(True)
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
        self.listWidget.model().rowsMoved.connect(lambda: self.listWidget.viewport().update())
        root.addWidget(self.listWidget)

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

        root.addLayout(btn_row)

        # 푸터
        footer = QLabel("v1.0  ·  Made by @ZionP")
        footer.setAlignment(Qt.AlignRight)
        footer.setStyleSheet("""
            color: #3a3a55;
            font-size: 11px;
            letter-spacing: 0.5px;
            padding-right: 2px;
        """)
        root.addWidget(footer)

    # ── 드래그앤 드롭 ──────────────────────────────────────────────
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

    # ── 파일 추가 / 삭제 ───────────────────────────────────────────
    def add_file(self, file_path):
        item = QListWidgetItem()
        item.setData(Qt.UserRole, file_path)
        self.listWidget.addItem(item)

    def _delete_row(self, row):
        self.listWidget.takeItem(row)

    # ── 병합 ───────────────────────────────────────────────────────
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

    # ── .ppt → .pptx 변환 ─────────────────────────────────────────
    @staticmethod
    def _get_soffice():
        if getattr(sys, 'frozen', False):
            base = sys._MEIPASS
            if sys.platform == 'darwin':
                return os.path.join(base, 'LibreOffice.app', 'Contents', 'MacOS', 'soffice')
            else:
                return os.path.join(base, 'LibreOffice', 'program', 'soffice.exe')
        if sys.platform == 'darwin':
            return '/Applications/LibreOffice.app/Contents/MacOS/soffice'
        for p in [r'C:\Program Files\LibreOffice\program\soffice.exe',
                  r'C:\Program Files (x86)\LibreOffice\program\soffice.exe']:
            if os.path.exists(p):
                return p
        return 'soffice'

    def _convert_ppt_files(self, file_paths, tmp_dir):
        ppt_files = [p for p in file_paths if p.lower().endswith('.ppt')]
        if ppt_files:
            soffice = self._get_soffice()
            if not os.path.exists(soffice):
                raise FileNotFoundError(
                    f"LibreOffice를 찾을 수 없습니다.\n경로: {soffice}\n"
                    "LibreOffice를 설치한 뒤 다시 시도해주세요."
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

    # ── 슬라이드 복사 ──────────────────────────────────────────────
    # python-pptx 가 자체 관리하는 관계 타입 — 직접 복사하면 충돌
    _SKIP_RELTYPES = {
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout',
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide',
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide',
    }

    @staticmethod
    def _remap_xml(xml: str, rId_map: dict) -> str:
        for old, new in rId_map.items():
            xml = xml.replace(f'r:embed="{old}"', f'r:embed="{new}"')
            xml = xml.replace(f'r:id="{old}"',    f'r:id="{new}"')
            xml = xml.replace(f'r:link="{old}"',  f'r:link="{new}"')
        return xml

    def _copy_slide(self, dest_prs, src_slide):
        blank = dest_prs.slide_layouts[min(6, len(dest_prs.slide_layouts) - 1)]
        dest_slide = dest_prs.slides.add_slide(blank)

        rId_map = {}
        for rel_id, rel in src_slide.part.rels.items():
            if rel.reltype in self._SKIP_RELTYPES:
                continue
            if rel.is_external:
                try:
                    rId_map[rel_id] = dest_slide.part.relate_to(
                        rel.target_ref, rel.reltype, is_external=True)
                except Exception:
                    pass
            else:
                try:
                    rId_map[rel_id] = dest_slide.part.relate_to(
                        rel.target_part, rel.reltype)
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
            xml = self._remap_xml(xml, rId_map)
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
                bg_copy = copy.deepcopy(src_bg)
                # 배경 이미지 rId 도 리매핑
                if rId_map:
                    bg_xml = etree.tostring(bg_copy, encoding='unicode')
                    bg_xml = self._remap_xml(bg_xml, rId_map)
                    bg_copy = etree.fromstring(bg_xml)
                dest_cSld.insert(0, bg_copy)

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
    ex = PPTMergerApp()
    ex.show()
    sys.exit(app.exec())
