"""漫画セリフ入れツール

セットアップ:
    pip install PyQt6

実行:
    python manga_serif_tool.py

使い方:
    - 画像フォルダをウィンドウにドラッグ&ドロップで読み込み
    - ホイール: ページ前後 / Ctrl+ホイール: 拡大縮小 / ホイールクリック+ドラッグ: パン
    - ツールバーから縦書き/横書きでセリフ追加、ダブルクリックで再編集
    - ページ移動・ウィンドウ閉じ時に <元フォルダ名>_serif/ へ JPG と edits.json を自動保存
"""

import sys
import os
import time
import traceback
import faulthandler
from pathlib import Path

_LOG_FILE = Path(__file__).resolve().parent / "manga_serif_tool.log"

# 起動時刻を即書き込み: ここに出ていなければ Python そのものが起動していない
try:
    with open(_LOG_FILE, "a", encoding="utf-8") as _f:
        _f.write(
            f"\n=== 起動 {time.strftime('%Y-%m-%d %H:%M:%S')} "
            f"pid={os.getpid()} exe={sys.executable} ===\n"
        )
    # ネイティブクラッシュ(SegFault等)もここに書き出す
    _fh_stream = open(_LOG_FILE, "a", encoding="utf-8")
    faulthandler.enable(file=_fh_stream, all_threads=True)
except Exception:
    pass


def _show_error_box(message: str) -> None:
    try:
        import ctypes
        ctypes.windll.user32.MessageBoxW(
            0, message[-3000:], "漫画セリフ入れツール - エラー", 0x10
        )
    except Exception:
        pass


def _excepthook(exc_type, exc_value, exc_tb):
    msg = "".join(traceback.format_exception(exc_type, exc_value, exc_tb))
    try:
        with open(_LOG_FILE, "a", encoding="utf-8") as f:
            f.write("\n--- 未捕捉エラー ---\n")
            f.write(msg)
    except Exception:
        pass
    _show_error_box(f"クラッシュしました。\n\n{msg}\n\n詳細ログ: {_LOG_FILE}")


sys.excepthook = _excepthook

# pythonw.exe で起動すると sys.stdout/stderr が None になり、
# そこへの書き込みでサイレントに落ちることがあるためログファイルへ向ける。
try:
    _log_stream = open(_LOG_FILE, "a", encoding="utf-8", buffering=1)
    sys.stdout = _log_stream
    sys.stderr = _log_stream
except Exception:
    pass

import json
from dataclasses import dataclass, asdict
from typing import List, Optional

try:
    from PyQt6.QtWidgets import (
        QApplication, QMainWindow, QGraphicsView, QGraphicsScene, QGraphicsItem,
        QGraphicsPixmapItem, QToolBar, QLabel, QSpinBox, QFontComboBox,
        QPushButton, QColorDialog, QStatusBar, QGraphicsObject, QInputDialog,
        QDialog, QVBoxLayout, QHBoxLayout, QListWidget, QListWidgetItem, QWidget,
        QSplitter, QPlainTextEdit, QTextEdit, QMenu,
    )
    from PyQt6.QtGui import (
        QPixmap, QImage, QPainter, QFont, QColor, QPen, QBrush, QAction,
        QFontMetrics, QKeySequence, QIcon, QTransform,
    )
    from PyQt6.QtCore import Qt, QPointF, QRectF, pyqtSignal, QSize
except Exception as _e:
    _excepthook(type(_e), _e, _e.__traceback__)
    sys.exit(1)

IMAGE_EXTS = {".jpg", ".jpeg", ".png", ".webp", ".bmp", ".gif", ".tiff"}

# 縦書き時に90°回転して描画する文字
_VERT_ROTATE = frozenset('ー―─━…‥〜～「」『』（）【】〔〕〈〉《》｛｝')
# 縦書き時に右上寄せにする句読点
_VERT_PUNCT  = frozenset('。、')


def _draw_char_vertical(painter: "QPainter", fm: "QFontMetrics",
                        ch: str, x: float, col_w: float, y_baseline: float) -> None:
    """縦書き1文字を適切な向き・位置で描画する。"""
    ch_w = fm.horizontalAdvance(ch)
    ch_h = fm.height()

    if ch in _VERT_ROTATE:
        # セル中心を軸に90°回転
        cx = x + col_w / 2
        cy = y_baseline - fm.ascent() + ch_h / 2
        painter.save()
        painter.translate(cx, cy)
        painter.rotate(90)
        # 回転後座標でセンタリング
        painter.drawText(QPointF(-ch_w / 2, (fm.ascent() - fm.descent()) / 2), ch)
        painter.restore()
    elif ch in _VERT_PUNCT:
        # 縦組みでは句読点を右上に配置する。
        # 日本語フォントの 、。 は全角 advance を持つが可視グリフは左下寄りに
        # 描かれるため、描画起点を右・上に約半セルずらしてインクを右上に来させる。
        x_ch = x + col_w / 2
        y_ch = y_baseline - ch_h / 2
        painter.drawText(QPointF(x_ch, y_ch), ch)
    else:
        painter.drawText(QPointF(x + (col_w - ch_w) / 2, y_baseline), ch)


class VerticalPreviewWidget(QWidget):
    """縦書きテキストを QPainter で描画するプレビューウィジェット"""

    def __init__(self):
        super().__init__()
        self._text = ""
        self._font_family = "Yu Gothic"
        self._font_size = 24
        self._char_spacing = 0
        self._line_spacing = 0
        self.setMinimumSize(80, 120)
        self.setStyleSheet("background: #2a2a2a; border: 1px solid #555;")

    def update_text(self, text: str, font_family: str, font_size: int,
                    char_spacing: int = 0, line_spacing: int = 0):
        self._text = text
        self._font_family = font_family
        self._font_size = font_size
        self._char_spacing = char_spacing
        self._line_spacing = line_spacing
        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.TextAntialiasing)
        painter.fillRect(self.rect(), QColor(42, 42, 42))

        text = self._text or " "
        pt = max(self._font_size * 2 // 3, 8)
        font = QFont(self._font_family, pt)
        painter.setFont(font)
        painter.setPen(QColor("#ffffff"))
        fm = QFontMetrics(font)
        char_sp = self._char_spacing
        line_sp = self._line_spacing

        columns = text.split("\n")
        col_widths = [
            max((fm.horizontalAdvance(c) for c in (col or " ")), default=fm.height())
            for col in columns
        ]
        col_heights = [
            fm.height() * len(col or " ") + char_sp * max(len(col or " ") - 1, 0)
            for col in columns
        ]
        total_w = sum(col_widths) + line_sp * max(len(columns) - 1, 0)
        max_h = max(col_heights) if col_heights else fm.height()

        # 中央揃え
        ox = (self.width() - total_w) // 2
        oy = (self.height() - max_h) // 2

        # 右から左へ列を配置
        x = ox + total_w
        for col_idx, col_text in enumerate(columns):
            col_w = col_widths[col_idx]
            x -= col_w
            y = oy + fm.ascent()
            for ch in (col_text or " "):
                _draw_char_vertical(painter, fm, ch, x, col_w, y)
                y += fm.height() + char_sp
            if col_idx < len(columns) - 1:
                x -= line_sp


class VerticalTextInputDialog(QDialog):
    """縦書きテキスト入力ダイアログ（プレビュー付き）"""

    def __init__(self, parent, initial_text: str = "", font_family: str = "Yu Gothic",
                 font_size: int = 24, char_spacing: int = 0, line_spacing: int = 0):
        super().__init__(parent)
        self.setWindowTitle("テキスト編集（縦書き）")
        self.resize(420, 300)
        self._font_family = font_family
        self._font_size = font_size
        self._char_spacing = char_spacing
        self._line_spacing = line_spacing

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(8, 8, 8, 8)
        main_layout.setSpacing(6)

        body = QHBoxLayout()
        body.setSpacing(8)

        left = QVBoxLayout()
        left.addWidget(QLabel("テキスト (Ctrl+Enter で改行):"))
        self.text_edit = QPlainTextEdit()
        self.text_edit.setPlainText(initial_text)
        self.text_edit.setFont(QFont(font_family, max(font_size * 2 // 3, 10)))
        self.text_edit.keyPressEvent = self._on_key_press
        self.text_edit.setMaximumHeight(180)
        left.addWidget(self.text_edit)

        right = QVBoxLayout()
        right.addWidget(QLabel("プレビュー:"))
        self.preview = VerticalPreviewWidget()
        right.addWidget(self.preview, 1)

        body.addLayout(left, 3)
        body.addLayout(right, 2)

        main_layout.addLayout(body, 1)

        btn_row = QHBoxLayout()
        btn_row.addStretch()
        ok_btn = QPushButton("OK")
        ok_btn.setDefault(True)
        ok_btn.clicked.connect(self.accept)
        cancel_btn = QPushButton("キャンセル")
        cancel_btn.clicked.connect(self.reject)
        btn_row.addWidget(ok_btn)
        btn_row.addWidget(cancel_btn)
        main_layout.addLayout(btn_row)

        self.text_edit.textChanged.connect(self._update_preview)
        self._update_preview()

    def _on_key_press(self, event):
        if (event.key() == Qt.Key.Key_Return
                and event.modifiers() == Qt.KeyboardModifier.ControlModifier):
            self.text_edit.insertPlainText("\n")
            event.accept()
        else:
            QPlainTextEdit.keyPressEvent(self.text_edit, event)

    def _update_preview(self):
        self.preview.update_text(
            self.text_edit.toPlainText() or " ",
            self._font_family, self._font_size,
            self._char_spacing, self._line_spacing,
        )

    def get_text(self) -> str:
        return self.text_edit.toPlainText()


@dataclass
class TextData:
    text: str = ""
    x: float = 0.0
    y: float = 0.0
    font_family: str = "Yu Gothic"
    font_size: int = 24
    color: str = "#000000"
    vertical: bool = True
    bold: bool = False
    char_spacing: int = 0   # 文字間隔 (extra px between chars / between rows)
    line_spacing: int = 0   # 行間隔   (extra px between columns / between lines)


class MangaTextItem(QGraphicsObject):
    """ドラッグ・編集可能な縦書き/横書きテキストアイテム。"""

    def __init__(self, data: TextData):
        super().__init__()
        self.data = data
        self.setFlag(QGraphicsItem.GraphicsItemFlag.ItemIsMovable, True)
        self.setFlag(QGraphicsItem.GraphicsItemFlag.ItemIsSelectable, True)
        self.setFlag(QGraphicsItem.GraphicsItemFlag.ItemSendsGeometryChanges, True)
        self.setPos(data.x, data.y)

    def _font(self) -> QFont:
        f = QFont(self.data.font_family, self.data.font_size)
        f.setBold(self.data.bold)
        return f

    def _col_sizes(self, fm: QFontMetrics):
        """列ごとの (width, height) リストを返す"""
        columns = self.data.text.split("\n") or [" "]
        cs = self.data.char_spacing
        ls = self.data.line_spacing
        col_widths = [
            max((fm.horizontalAdvance(c) for c in (col or " ")), default=fm.height())
            for col in columns
        ]
        col_heights = [
            fm.height() * len(col or " ") + cs * max(len(col or " ") - 1, 0)
            for col in columns
        ]
        total_w = sum(col_widths) + ls * max(len(columns) - 1, 0)
        return columns, col_widths, col_heights, total_w

    def boundingRect(self) -> QRectF:
        fm = QFontMetrics(self._font())
        if self.data.vertical:
            columns, col_widths, col_heights, total_w = self._col_sizes(fm)
            max_h = max(col_heights) if col_heights else fm.height()
            return QRectF(0, 0, total_w + 6, max_h + 6)
        # 横書き
        lines = self.data.text.split("\n") or [" "]
        w = max((fm.horizontalAdvance(line) for line in lines), default=fm.height())
        h = (fm.height() + self.data.line_spacing) * len(lines)
        return QRectF(0, 0, w + 6, h + 6)

    def paint(self, painter: QPainter, option, widget=None):
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.setRenderHint(QPainter.RenderHint.TextAntialiasing)
        painter.setFont(self._font())
        painter.setPen(QColor(self.data.color))

        fm = QFontMetrics(self._font())
        cs = self.data.char_spacing
        ls = self.data.line_spacing
        if self.data.vertical:
            columns, col_widths, _, total_w = self._col_sizes(fm)
            # 右から左へ列を配置
            x = total_w + 3
            for col_idx, col_text in enumerate(columns):
                col_w = col_widths[col_idx]
                x -= col_w
                y = fm.ascent() + 3
                for ch in (col_text or ""):
                    _draw_char_vertical(painter, fm, ch, x, col_w, y)
                    y += fm.height() + cs
                if col_idx < len(columns) - 1:
                    x -= ls
        else:
            y = fm.ascent() + 3
            for line in self.data.text.split("\n"):
                painter.drawText(QPointF(3, y), line)
                y += fm.height() + ls

        if self.isSelected():
            painter.setPen(QPen(QColor(0, 120, 215), 1, Qt.PenStyle.DashLine))
            painter.setBrush(Qt.BrushStyle.NoBrush)
            painter.drawRect(self.boundingRect())

    def itemChange(self, change, value):
        if change == QGraphicsItem.GraphicsItemChange.ItemPositionChange:
            self.data.x = value.x()
            self.data.y = value.y()
        return super().itemChange(change, value)

    def mouseDoubleClickEvent(self, event):
        view = self.scene().views()[0] if self.scene() and self.scene().views() else None
        if view:
            win = view.window()
            if hasattr(win, "edit_item_text"):
                win.edit_item_text(self)
        super().mouseDoubleClickEvent(event)

    def contextMenuEvent(self, event):
        view = self.scene().views()[0] if self.scene() and self.scene().views() else None
        win = view.window() if view else None
        menu = QMenu()
        menu.addAction("複製").triggered.connect(
            lambda: win.duplicate_item(self) if win else None
        )
        menu.addAction("削除").triggered.connect(
            lambda: win.delete_item(self) if win else None
        )
        menu.exec(event.screenPos())


class LayerPanel(QWidget):
    """セリフ一覧パネル"""

    item_visibility_changed = pyqtSignal(object, bool)
    item_selected = pyqtSignal(object)
    item_duplicate_requested = pyqtSignal(object)
    item_delete_requested = pyqtSignal(object)

    def __init__(self):
        super().__init__()
        layout = QVBoxLayout(self)
        layout.setContentsMargins(4, 4, 4, 4)

        header = QLabel("レイヤー")
        header.setStyleSheet("font-weight: bold; padding: 2px;")
        layout.addWidget(header)

        self.list_widget = QListWidget()
        self.list_widget.itemClicked.connect(self._on_item_clicked)
        self.list_widget.itemChanged.connect(self._on_list_item_changed)
        self.list_widget.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.list_widget.customContextMenuRequested.connect(self._show_context_menu)
        layout.addWidget(self.list_widget)

        self._item_map = {}        # MangaTextItem -> QListWidgetItem
        self._reverse_map = {}     # QListWidgetItem -> MangaTextItem
        self._ignore_changes = False

    def add_item(self, manga_text_item):
        self._ignore_changes = True
        text_preview = manga_text_item.data.text[:20].replace("\n", "¶")
        list_item = QListWidgetItem(text_preview)
        list_item.setCheckState(Qt.CheckState.Checked)
        list_item.setFlags(list_item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
        self.list_widget.addItem(list_item)
        self._item_map[manga_text_item] = list_item
        self._reverse_map[id(list_item)] = manga_text_item
        self._ignore_changes = False

    def remove_item(self, manga_text_item):
        if manga_text_item in self._item_map:
            list_item = self._item_map.pop(manga_text_item)
            self._reverse_map.pop(id(list_item), None)
            self.list_widget.takeItem(self.list_widget.row(list_item))

    def clear(self):
        self.list_widget.clear()
        self._item_map.clear()
        self._reverse_map.clear()

    def update_item_preview(self, manga_text_item):
        if manga_text_item in self._item_map:
            self._ignore_changes = True
            text_preview = manga_text_item.data.text[:20].replace("\n", "¶")
            self._item_map[manga_text_item].setText(text_preview)
            self._ignore_changes = False

    def set_selected(self, manga_text_item):
        if manga_text_item in self._item_map:
            self.list_widget.setCurrentItem(self._item_map[manga_text_item])

    def _on_item_clicked(self, list_item):
        item = self._reverse_map.get(id(list_item))
        if item:
            self.item_selected.emit(item)

    def _on_list_item_changed(self, list_item):
        if self._ignore_changes:
            return
        item = self._reverse_map.get(id(list_item))
        if item:
            visible = list_item.checkState() == Qt.CheckState.Checked
            self.item_visibility_changed.emit(item, visible)

    def _show_context_menu(self, pos):
        list_item = self.list_widget.itemAt(pos)
        if list_item is None:
            return
        item = self._reverse_map.get(id(list_item))
        if item is None:
            return
        menu = QMenu(self)
        menu.addAction("複製").triggered.connect(lambda: self.item_duplicate_requested.emit(item))
        menu.addAction("削除").triggered.connect(lambda: self.item_delete_requested.emit(item))
        menu.exec(self.list_widget.viewport().mapToGlobal(pos))


class MangaView(QGraphicsView):
    page_change_requested = pyqtSignal(int)
    folder_dropped = pyqtSignal(Path)

    def __init__(self, scene):
        super().__init__(scene)
        self.setRenderHint(QPainter.RenderHint.Antialiasing)
        self.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform)
        self.setTransformationAnchor(QGraphicsView.ViewportAnchor.AnchorUnderMouse)
        self.setResizeAnchor(QGraphicsView.ViewportAnchor.AnchorUnderMouse)
        self.setDragMode(QGraphicsView.DragMode.RubberBandDrag)
        self.setBackgroundBrush(QBrush(QColor(40, 40, 40)))
        self._panning = False
        self._pan_start = QPointF()
        self.setAcceptDrops(True)

    def wheelEvent(self, event):
        if event.modifiers() & Qt.KeyboardModifier.ControlModifier:
            factor = 1.15 if event.angleDelta().y() > 0 else 1 / 1.15
            self.scale(factor, factor)
        else:
            self.page_change_requested.emit(-1 if event.angleDelta().y() > 0 else 1)
        event.accept()

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.MiddleButton:
            self._panning = True
            self._pan_start = event.position()
            self.viewport().setCursor(Qt.CursorShape.ClosedHandCursor)
            event.accept()
            return
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if self._panning:
            pos = event.position()
            dx = pos.x() - self._pan_start.x()
            dy = pos.y() - self._pan_start.y()
            self._pan_start = pos
            hbar, vbar = self.horizontalScrollBar(), self.verticalScrollBar()
            hbar.setValue(hbar.value() - int(dx))
            vbar.setValue(vbar.value() - int(dy))
            event.accept()
            return
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.MouseButton.MiddleButton and self._panning:
            self._panning = False
            self.viewport().setCursor(Qt.CursorShape.ArrowCursor)
            event.accept()
            return
        super().mouseReleaseEvent(event)

    def fit_to_window(self):
        scene = self.scene()
        if scene and scene.sceneRect().width() > 0:
            self.resetTransform()
            self.fitInView(scene.sceneRect(), Qt.AspectRatioMode.KeepAspectRatio)

    def keyPressEvent(self, event):
        arrow_keys = {
            Qt.Key.Key_Left:  (-1,  0),
            Qt.Key.Key_Right: ( 1,  0),
            Qt.Key.Key_Up:    ( 0, -1),
            Qt.Key.Key_Down:  ( 0,  1),
        }
        if event.key() not in arrow_keys:
            super().keyPressEvent(event)
            return

        dx, dy = arrow_keys[event.key()]
        step = 10 if event.modifiers() & Qt.KeyboardModifier.ShiftModifier else 1

        selected = [it for it in self.scene().selectedItems()
                    if hasattr(it, "data")]  # MangaTextItem のみ
        if selected:
            for it in selected:
                it.moveBy(dx * step, dy * step)
            event.accept()
        else:
            # 選択なし → 左右はページ送り、上下は通常スクロール
            if event.key() == Qt.Key.Key_Left:
                self.page_change_requested.emit(-1)
                event.accept()
            elif event.key() == Qt.Key.Key_Right:
                self.page_change_requested.emit(1)
                event.accept()
            else:
                super().keyPressEvent(event)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        urls = event.mimeData().urls()
        if not urls:
            return
        path = Path(urls[0].toLocalFile())
        if path.is_dir():
            self.folder_dropped.emit(path)
        elif path.is_file() and path.suffix.lower() in IMAGE_EXTS:
            self.folder_dropped.emit(path.parent)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("漫画セリフ入れツール")
        self.resize(1400, 800)
        self.setAcceptDrops(True)

        self.image_paths: List[Path] = []
        self.current_index: int = -1
        self.folder_path: Optional[Path] = None
        self.serif_folder: Optional[Path] = None
        self.edits_data: dict = {}
        self.current_pixmap_item: Optional[QGraphicsPixmapItem] = None
        self._current_color = QColor("#000000")
        self._suspend_sync = False

        self.scene = QGraphicsScene(self)
        self.view = MangaView(self.scene)
        self.view.page_change_requested.connect(self.change_page)
        self.view.folder_dropped.connect(self.load_folder)

        # レイヤーパネル
        self.layer_panel = LayerPanel()
        self.layer_panel.item_visibility_changed.connect(self._on_layer_visibility_changed)
        self.layer_panel.item_selected.connect(self._on_layer_item_selected)
        self.layer_panel.item_duplicate_requested.connect(self.duplicate_item)
        self.layer_panel.item_delete_requested.connect(self.delete_item)

        # スプリッター：左に画像ビュー、右にレイヤーパネル
        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.addWidget(self.view)
        splitter.addWidget(self.layer_panel)
        splitter.setStretchFactor(0, 3)
        splitter.setStretchFactor(1, 1)
        splitter.setSizes([1050, 350])
        self.setCentralWidget(splitter)

        self.setStatusBar(QStatusBar())
        self.scene.selectionChanged.connect(self._on_selection_changed)

        self._build_toolbar()
        self.statusBar().showMessage(
            "画像フォルダをここにドラッグ&ドロップしてください  |  ホイール:ページ送り / Ctrl+ホイール:拡大縮小 / ホイールクリック:パン"
        )

    def _build_toolbar(self):
        tb = QToolBar("Tools")
        tb.setMovable(False)
        self.addToolBar(tb)

        a = QAction("横書き追加", self)
        a.triggered.connect(lambda: self.add_text(vertical=False))
        tb.addAction(a)

        a = QAction("縦書き追加", self)
        a.triggered.connect(lambda: self.add_text(vertical=True))
        tb.addAction(a)

        tb.addSeparator()
        self.layers_panel_action = QAction("レイヤーパネル ▼", self, checkable=True)
        self.layers_panel_action.setChecked(True)
        self.layers_panel_action.toggled.connect(self.layer_panel.setVisible)
        tb.addAction(self.layers_panel_action)

        tb.addSeparator()
        tb.addWidget(QLabel(" フォント: "))
        self.font_combo = QFontComboBox()
        self.font_combo.setCurrentFont(QFont("Yu Gothic"))
        self.font_combo.currentFontChanged.connect(self._on_font_changed)
        tb.addWidget(self.font_combo)

        tb.addWidget(QLabel(" サイズ: "))
        self.size_spin = QSpinBox()
        self.size_spin.setRange(6, 300)
        self.size_spin.setValue(24)
        self.size_spin.valueChanged.connect(self._on_size_changed)
        tb.addWidget(self.size_spin)

        tb.addSeparator()
        self.bold_action = QAction("太字", self)
        self.bold_action.setCheckable(True)
        self.bold_action.toggled.connect(self._on_bold_changed)
        tb.addAction(self.bold_action)

        self.color_btn = QPushButton("色")
        self.color_btn.clicked.connect(self._pick_color)
        self._update_color_btn()
        tb.addWidget(self.color_btn)

        tb.addWidget(QLabel(" 文字間: "))
        self.char_spacing_spin = QSpinBox()
        self.char_spacing_spin.setRange(-20, 200)
        self.char_spacing_spin.setValue(0)
        self.char_spacing_spin.setSuffix("px")
        self.char_spacing_spin.setFixedWidth(68)
        self.char_spacing_spin.valueChanged.connect(self._on_char_spacing_changed)
        tb.addWidget(self.char_spacing_spin)

        tb.addWidget(QLabel(" 行間: "))
        self.line_spacing_spin = QSpinBox()
        self.line_spacing_spin.setRange(-20, 200)
        self.line_spacing_spin.setValue(0)
        self.line_spacing_spin.setSuffix("px")
        self.line_spacing_spin.setFixedWidth(68)
        self.line_spacing_spin.valueChanged.connect(self._on_line_spacing_changed)
        tb.addWidget(self.line_spacing_spin)

        tb.addSeparator()
        a = QAction("縦↔横切替", self)
        a.triggered.connect(self._toggle_orientation)
        tb.addAction(a)

        a = QAction("削除", self)
        a.setShortcut(QKeySequence(Qt.Key.Key_Delete))
        a.triggered.connect(self._delete_selected)
        tb.addAction(a)

        tb.addSeparator()
        a = QAction("画面に合わせる", self)
        a.setShortcut(QKeySequence("Ctrl+0"))
        a.triggered.connect(self.view.fit_to_window)
        tb.addAction(a)

        a = QAction("◀ 前", self)
        a.triggered.connect(lambda: self.change_page(-1))
        tb.addAction(a)

        self.page_label = QLabel("  - / -  ")
        tb.addWidget(self.page_label)

        a = QAction("次 ▶", self)
        a.triggered.connect(lambda: self.change_page(1))
        tb.addAction(a)

        tb.addSeparator()
        a = QAction("💾 保存", self)
        a.setShortcut(QKeySequence("Ctrl+S"))
        a.setToolTip("現在のページを保存 (Ctrl+S)")
        a.triggered.connect(self._save_current_page_manual)
        tb.addAction(a)

    def _save_current_page_manual(self):
        if self.current_index < 0 or not self.image_paths:
            self.statusBar().showMessage("保存するページがありません")
            return
        self._save_current_page()
        path = self.image_paths[self.current_index]
        self.statusBar().showMessage(f"保存しました: {path.name}  →  {self.serif_folder}")

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dragMoveEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        urls = event.mimeData().urls()
        if not urls:
            return
        path = Path(urls[0].toLocalFile())
        if path.is_dir():
            self.load_folder(path)
        elif path.is_file() and path.suffix.lower() in IMAGE_EXTS:
            self.load_folder(path.parent)

    def load_folder(self, folder: Path):
        self._save_current_page()

        self.folder_path = folder
        self.serif_folder = folder.parent / f"{folder.name}_serif"
        self.serif_folder.mkdir(exist_ok=True)

        edits_file = self.serif_folder / "edits.json"
        if edits_file.exists():
            try:
                with open(edits_file, "r", encoding="utf-8") as f:
                    raw = json.load(f)
                self.edits_data = {
                    fn: [TextData(**td) for td in items] for fn, items in raw.items()
                }
            except Exception as e:
                self.statusBar().showMessage(f"edits.json 読み込みエラー: {e}")
                self.edits_data = {}
        else:
            self.edits_data = {}

        self.image_paths = sorted(
            [p for p in folder.iterdir() if p.is_file() and p.suffix.lower() in IMAGE_EXTS]
        )
        if not self.image_paths:
            self.statusBar().showMessage("画像が見つかりません")
            return

        self.current_index = 0
        self.show_page(self.current_index)
        self.statusBar().showMessage(f"{folder} ({len(self.image_paths)}枚)  →  保存先: {self.serif_folder}")

    def show_page(self, index: int):
        if not (0 <= index < len(self.image_paths)):
            return
        self.scene.clear()
        self.layer_panel.clear()
        self.current_pixmap_item = None

        path = self.image_paths[index]
        pixmap = QPixmap(str(path))
        if pixmap.isNull():
            self.statusBar().showMessage(f"画像読み込み失敗: {path.name}")
            return

        self.current_pixmap_item = self.scene.addPixmap(pixmap)
        self.current_pixmap_item.setZValue(-1)
        self.scene.setSceneRect(QRectF(pixmap.rect()))

        for td in self.edits_data.get(path.name, []):
            self._add_text_item_from_data(td)

        self.view.fit_to_window()
        self.page_label.setText(f"  {index + 1} / {len(self.image_paths)}  ")
        self.setWindowTitle(f"漫画セリフ入れツール - {path.name}")

    def change_page(self, delta: int):
        if not self.image_paths:
            return
        new_idx = self.current_index + delta
        if not (0 <= new_idx < len(self.image_paths)):
            return
        self._save_current_page()
        self.current_index = new_idx
        self.show_page(self.current_index)

    def _save_current_page(self):
        if self.current_index < 0 or not self.image_paths or self.serif_folder is None:
            return
        path = self.image_paths[self.current_index]

        items = [it for it in self.scene.items() if isinstance(it, MangaTextItem)]
        items.reverse()
        self.edits_data[path.name] = [it.data for it in items]

        try:
            with open(self.serif_folder / "edits.json", "w", encoding="utf-8") as f:
                json.dump(
                    {fn: [asdict(td) for td in tds] for fn, tds in self.edits_data.items()},
                    f, ensure_ascii=False, indent=2,
                )
        except Exception as e:
            self.statusBar().showMessage(f"JSON保存エラー: {e}")

        self._render_jpg(path)

    def _render_jpg(self, src_path: Path):
        if not self.current_pixmap_item or not self.serif_folder:
            return

        prev_selected = list(self.scene.selectedItems())
        self.scene.clearSelection()

        rect = self.scene.sceneRect()
        image = QImage(int(rect.width()), int(rect.height()), QImage.Format.Format_RGB32)
        image.fill(Qt.GlobalColor.white)
        painter = QPainter(image)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform)
        painter.setRenderHint(QPainter.RenderHint.TextAntialiasing)
        self.scene.render(painter, QRectF(image.rect()), rect)
        painter.end()

        out_path = self.serif_folder / (src_path.stem + ".jpg")
        image.save(str(out_path), "JPG", 95)

        for it in prev_selected:
            it.setSelected(True)

    def add_text(self, vertical: bool):
        if not self.current_pixmap_item:
            self.statusBar().showMessage("先に画像フォルダをドロップしてください")
            return
        center = self.view.mapToScene(self.view.viewport().rect().center())
        td = TextData(
            text="セリフ",
            x=center.x(),
            y=center.y(),
            font_family=self.font_combo.currentFont().family(),
            font_size=self.size_spin.value(),
            color=self._current_color.name(),
            vertical=vertical,
            bold=self.bold_action.isChecked(),
            char_spacing=self.char_spacing_spin.value(),
            line_spacing=self.line_spacing_spin.value(),
        )
        item = self._add_text_item_from_data(td)
        self.scene.clearSelection()
        item.setSelected(True)
        self.edit_item_text(item)

    def _add_text_item_from_data(self, td: TextData) -> MangaTextItem:
        item = MangaTextItem(td)
        self.scene.addItem(item)
        self.layer_panel.add_item(item)
        return item

    def edit_item_text(self, item: MangaTextItem):
        if item.data.vertical:
            # 縦書き用カスタムダイアログ
            dlg = VerticalTextInputDialog(
                self, item.data.text, item.data.font_family, item.data.font_size,
                item.data.char_spacing, item.data.line_spacing,
            )
            if dlg.exec() == QDialog.DialogCode.Accepted:
                item.prepareGeometryChange()
                item.data.text = dlg.get_text()
                item.update()
                self.layer_panel.update_item_preview(item)
        else:
            # 横書き用スタンダードダイアログ
            text, ok = QInputDialog.getMultiLineText(
                self, "テキスト編集", "セリフを入力（改行可）:", item.data.text
            )
            if ok:
                item.prepareGeometryChange()
                item.data.text = text
                item.update()
                self.layer_panel.update_item_preview(item)

    def _selected_items(self):
        return [it for it in self.scene.selectedItems() if isinstance(it, MangaTextItem)]

    def _on_selection_changed(self):
        sel = self._selected_items()
        if not sel:
            return
        it = sel[0]
        self._suspend_sync = True
        self.font_combo.setCurrentFont(QFont(it.data.font_family))
        self.size_spin.setValue(it.data.font_size)
        self.bold_action.setChecked(it.data.bold)
        self._current_color = QColor(it.data.color)
        self._update_color_btn()
        self.char_spacing_spin.setValue(it.data.char_spacing)
        self.line_spacing_spin.setValue(it.data.line_spacing)
        self.layer_panel.set_selected(it)
        self._suspend_sync = False

    def _on_font_changed(self, font: QFont):
        if self._suspend_sync:
            return
        for it in self._selected_items():
            it.prepareGeometryChange()
            it.data.font_family = font.family()
            it.update()

    def _on_size_changed(self, value: int):
        if self._suspend_sync:
            return
        for it in self._selected_items():
            it.prepareGeometryChange()
            it.data.font_size = value
            it.update()

    def _on_bold_changed(self, checked: bool):
        if self._suspend_sync:
            return
        for it in self._selected_items():
            it.prepareGeometryChange()
            it.data.bold = checked
            it.update()

    def _on_char_spacing_changed(self, value: int):
        if self._suspend_sync:
            return
        for it in self._selected_items():
            it.prepareGeometryChange()
            it.data.char_spacing = value
            it.update()

    def _on_line_spacing_changed(self, value: int):
        if self._suspend_sync:
            return
        for it in self._selected_items():
            it.prepareGeometryChange()
            it.data.line_spacing = value
            it.update()

    def _pick_color(self):
        c = QColorDialog.getColor(self._current_color, self, "色を選ぶ")
        if not c.isValid():
            return
        self._current_color = c
        self._update_color_btn()
        for it in self._selected_items():
            it.data.color = c.name()
            it.update()

    def _update_color_btn(self):
        c = self._current_color
        fg = "white" if c.lightness() < 128 else "black"
        self.color_btn.setStyleSheet(f"background-color: {c.name()}; color: {fg};")

    def _toggle_orientation(self):
        for it in self._selected_items():
            it.prepareGeometryChange()
            it.data.vertical = not it.data.vertical
            it.update()

    def _delete_selected(self):
        for it in self._selected_items():
            self.delete_item(it)

    def delete_item(self, item: "MangaTextItem"):
        self.scene.removeItem(item)
        self.layer_panel.remove_item(item)

    def duplicate_item(self, item: "MangaTextItem"):
        from dataclasses import replace as dc_replace
        new_data = dc_replace(item.data, x=item.data.x + 20, y=item.data.y + 20)
        new_item = self._add_text_item_from_data(new_data)
        self.scene.clearSelection()
        new_item.setSelected(True)

    def _on_layer_visibility_changed(self, item: MangaTextItem, visible: bool):
        item.setVisible(visible)

    def _on_layer_item_selected(self, item: MangaTextItem):
        self.scene.clearSelection()
        item.setSelected(True)

    def closeEvent(self, event):
        self._save_current_page()
        super().closeEvent(event)


def main():
    try:
        app = QApplication(sys.argv)
        win = MainWindow()
        win.show()
        sys.exit(app.exec())
    except SystemExit:
        raise
    except BaseException:
        _excepthook(*sys.exc_info())
        sys.exit(1)


if __name__ == "__main__":
    main()
