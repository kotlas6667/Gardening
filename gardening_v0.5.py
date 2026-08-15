from Config import NAZOV_APP, NAZOV_FIRMY, HEIGHT, WIDTH, X_POSITION, Y_POSITION, TABLES_2025_MORE, TABLES_2025_LESS
import sys
import os
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QLineEdit, QPushButton, QComboBox, QTableWidget, QTableWidgetItem,
    QCalendarWidget, QMessageBox, QFrame, QScrollArea, QSizePolicy,
    QCompleter, QAbstractItemView, QDateEdit, QSplitter, QFileDialog,
    QGraphicsDropShadowEffect, QHeaderView, QToolButton, QSpinBox, QMenu,
)
from PyQt6.QtCore import Qt, QDate, QSettings, QTimer
from PyQt6.QtGui import QFont, QIcon, QColor, QPalette, QTextCharFormat
import sqlite3
import openpyxl
import locale

# pyinstaller --onefile --windowed --icon=ikona.ico gardening_v0.5.py --add-data "Config.py;."

CALENDAR_STYLE = """
QCalendarWidget {
    background-color: #ffffff;
    border: 1px solid #d7e3d9;
    border-radius: 8px;
}
QCalendarWidget QWidget#qt_calendar_navigationbar {
    background-color: #2d6a4f;
    border: none;
    border-top-left-radius: 7px;
    border-top-right-radius: 7px;
    min-height: 32px;
}
QCalendarWidget QToolButton {
    color: #ffffff;
    background-color: transparent;
    font-family: "Segoe UI Semibold";
    font-size: 13px;
    border-radius: 6px;
    padding: 4px 8px;
    margin: 2px;
}
QCalendarWidget QToolButton:hover {
    background-color: rgba(255, 255, 255, 0.18);
}
QCalendarWidget QToolButton#qt_calendar_prevmonth,
QCalendarWidget QToolButton#qt_calendar_nextmonth {
    color: #ffffff;
    qproperty-iconSize: 16px 16px;
}
QCalendarWidget QSpinBox {
    background-color: #2d6a4f;
    color: #ffffff;
    border: none;
    font-family: "Segoe UI Semibold";
    selection-background-color: #40916c;
    selection-color: #ffffff;
}
QCalendarWidget QSpinBox::up-button,
QCalendarWidget QSpinBox::down-button {
    width: 0;
    height: 0;
    border: none;
}
QCalendarWidget QMenu {
    background-color: #ffffff;
    color: #1b4332;
    border: 1px solid #c9d8cc;
}
QCalendarWidget QMenu::item {
    background-color: #ffffff;
    color: #1b4332;
    padding: 6px 18px;
}
QCalendarWidget QMenu::item:selected {
    background-color: #d8eee1;
    color: #1b4332;
}
QCalendarWidget QAbstractItemView {
    background-color: #ffffff;
    color: #1f2d26;
    selection-background-color: #2d6a4f;
    selection-color: #ffffff;
    outline: none;
    font-family: "Segoe UI";
    font-size: 12px;
}
QCalendarWidget QAbstractItemView:enabled {
    background-color: #ffffff;
    color: #1f2d26;
    selection-background-color: #2d6a4f;
    selection-color: #ffffff;
}
QCalendarWidget QAbstractItemView:disabled {
    color: #a8b5ad;
}
QCalendarWidget QWidget {
    background-color: #ffffff;
    alternate-background-color: #f7faf8;
}
"""

APP_STYLE = """
QMainWindow {
    background-color: #eef4ef;
}
QScrollArea, QScrollArea > QWidget, QScrollArea > QWidget > QWidget {
    background-color: #ffffff;
    border: none;
}
QScrollArea#metricsScroll, QScrollArea#metricsScroll > QWidget,
QScrollArea#leftFormScroll, QScrollArea#leftFormScroll > QWidget {
    background-color: #ffffff;
}
QWidget#centralRoot {
    background: qlineargradient(
        x1:0, y1:0, x2:1, y2:1,
        stop:0 #e8f2ea, stop:0.45 #eef4ef, stop:1 #e6eef8
    );
}
QWidget#metricsWrap {
    background-color: #ffffff;
}
QFrame#headerBar {
    background: qlineargradient(
        x1:0, y1:0, x2:1, y2:0,
        stop:0 #1b4332, stop:1 #2d6a4f
    );
    border: none;
    border-radius: 14px;
}
QLabel#brandTitle {
    color: #ffffff;
    font-family: "Segoe UI Semibold";
    font-size: 22px;
    font-weight: 600;
    letter-spacing: 0.5px;
}
QLabel#brandSubtitle {
    color: rgba(255, 255, 255, 0.72);
    font-family: "Segoe UI";
    font-size: 11px;
}
QLabel#headerMeta {
    color: rgba(255, 255, 255, 0.88);
    font-family: "Segoe UI";
    font-size: 12px;
}
QFrame#toolbarCard, QFrame#formCard, QFrame#tableCard, QFrame#metricCard {
    background-color: #ffffff;
    border: 1px solid #d7e3d9;
    border-radius: 14px;
}
QLabel#sectionTitle {
    color: #1b4332;
    font-family: "Segoe UI Semibold";
    font-size: 14px;
    font-weight: 600;
}
QLabel#fieldLabel {
    color: #3d5a4a;
    font-family: "Segoe UI";
    font-size: 11px;
    font-weight: 600;
}
QLabel#metricCaption {
    color: #6b7f72;
    font-family: "Segoe UI";
    font-size: 11px;
}
QLabel#metricValue {
    color: #1b4332;
    font-family: "Segoe UI Semibold";
    font-size: 16px;
    font-weight: 600;
}
QLineEdit, QComboBox, QDateEdit {
    background-color: #f7faf8;
    border: 1px solid #c9d8cc;
    border-radius: 8px;
    padding: 7px 10px;
    color: #1f2d26;
    font-family: "Segoe UI";
    min-height: 18px;
}
QLineEdit:focus, QComboBox:focus, QDateEdit:focus {
    border: 1px solid #40916c;
    background-color: #ffffff;
}
QComboBox::drop-down {
    border: none;
    width: 24px;
}
QComboBox QAbstractItemView {
    background-color: #ffffff;
    color: #1f2d26;
    border: 1px solid #c9d8cc;
    selection-background-color: #d8eee1;
    selection-color: #1b4332;
    outline: 0;
}
QComboBox QAbstractItemView::item {
    min-height: 26px;
    padding: 4px 10px;
    color: #1f2d26;
    background-color: #ffffff;
}
QComboBox QAbstractItemView::item:hover,
QComboBox QAbstractItemView::item:selected {
    background-color: #d8eee1;
    color: #1b4332;
}
QDateEdit QAbstractItemView {
    background-color: #ffffff;
    color: #1f2d26;
    selection-background-color: #d8eee1;
    selection-color: #1b4332;
}
QMenu {
    background-color: #ffffff;
    color: #1f2d26;
    border: 1px solid #c9d8cc;
}
QMenu::item {
    background-color: #ffffff;
    color: #1f2d26;
    padding: 6px 18px;
}
QMenu::item:selected {
    background-color: #d8eee1;
    color: #1b4332;
}
QPushButton {
    font-family: "Segoe UI Semibold";
    font-size: 13px;
    border-radius: 9px;
    padding: 8px 16px;
    min-height: 20px;
}
QPushButton#btnPrimary {
    background-color: #2d6a4f;
    color: white;
    border: none;
}
QPushButton#btnPrimary:hover { background-color: #1b4332; }
QPushButton#btnPrimary:disabled { background-color: #9bb5a6; color: #e8f0eb; }
QPushButton#btnSecondary {
    background-color: #edf5ef;
    color: #1b4332;
    border: 1px solid #b7cfc0;
}
QPushButton#btnSecondary:hover { background-color: #dceee2; }
QPushButton#btnSecondary:disabled { color: #9aa89f; background-color: #f3f6f4; }
QPushButton#btnDanger {
    background-color: #fff1f0;
    color: #9b2226;
    border: 1px solid #f0c2c0;
}
QPushButton#btnDanger:hover { background-color: #ffe3e1; }
QPushButton#btnDanger:disabled { color: #c9a8a8; background-color: #faf5f5; }
QPushButton#btnGhost {
    background-color: rgba(255, 255, 255, 0.14);
    color: #ffffff;
    border: 1px solid rgba(255, 255, 255, 0.28);
}
QPushButton#btnGhost:hover { background-color: rgba(255, 255, 255, 0.24); }
QPushButton#btnTodaySmall {
    background-color: rgba(255, 255, 255, 0.14);
    color: #ffffff;
    border: 1px solid rgba(255, 255, 255, 0.28);
    font-size: 11px;
    padding: 3px 10px;
    min-height: 14px;
    max-height: 28px;
}
QPushButton#btnTodaySmall:hover { background-color: rgba(255, 255, 255, 0.24); }
QPushButton#btnCalendarToday {
    color: #ffffff;
    background-color: rgba(255, 255, 255, 0.14);
    border: 1px solid rgba(255, 255, 255, 0.28);
    border-radius: 6px;
    font-size: 11px;
    padding: 2px 8px;
    min-height: 20px;
    max-height: 24px;
}
QPushButton#btnCalendarToday:hover { background-color: rgba(255, 255, 255, 0.24); }
QTableWidget {
    background-color: #ffffff;
    alternate-background-color: #f5faf6;
    color: #1f2d26;
    gridline-color: #e2ebe4;
    border: none;
    border-radius: 8px;
    font-family: "Segoe UI";
    font-size: 13px;
    selection-background-color: #d8eee1;
    selection-color: #1b4332;
}
QHeaderView::section {
    background-color: #eef6f0;
    color: #1b4332;
    font-family: "Segoe UI Semibold";
    font-weight: 600;
    padding: 8px 6px;
    border: none;
    border-bottom: 2px solid #c9dccf;
    border-right: 1px solid #e2ebe4;
}
QScrollBar:vertical {
    background: #f0f5f1;
    width: 10px;
    margin: 2px;
}
QScrollBar::handle:vertical {
    background: #b7cfc0;
    border-radius: 5px;
    min-height: 30px;
}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0; }
QScrollBar:horizontal {
    background: #f0f5f1;
    height: 10px;
    margin: 2px;
}
QScrollBar::handle:horizontal {
    background: #b7cfc0;
    border-radius: 5px;
    min-width: 30px;
}
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal { width: 0; }
QMessageBox {
    background-color: #ffffff;
}
"""

DATE_UI = "dd.MM.yyyy"
DATE_DB = "yyyy-MM-dd"


def add_shadow(widget, blur=24, dx=0, dy=4, alpha=28):
    effect = QGraphicsDropShadowEffect(widget)
    effect.setBlurRadius(blur)
    effect.setOffset(dx, dy)
    effect.setColor(QColor(27, 67, 50, alpha))
    widget.setGraphicsEffect(effect)
    return widget


def parse_date(value):
    """Parsuje dátum z UI (dd.MM.yyyy) alebo DB (yyyy-MM-dd)."""
    if value is None:
        return QDate()
    text = str(value).strip()
    if not text:
        return QDate()
    for fmt in (DATE_UI, DATE_DB, "d.M.yyyy", "yyyy/MM/dd"):
        parsed = QDate.fromString(text, fmt)
        if parsed.isValid():
            return parsed
    return QDate()


def to_ui_date(value):
    parsed = parse_date(value)
    return parsed.toString(DATE_UI) if parsed.isValid() else (str(value).strip() if value else "")


def to_db_date(value):
    parsed = parse_date(value)
    return parsed.toString(DATE_DB) if parsed.isValid() else (str(value).strip() if value else "")


def fiscal_year_from_date(value):
    """Fiskálny rok z dátumu: apríl–marec (mesiac >= 4 → rok dátumu, inak rok - 1)."""
    parsed = parse_date(value)
    if not parsed.isValid():
        return None
    year = parsed.year()
    month = parsed.month()
    return str(year if month >= 4 else year - 1)


class FrmGarden(QMainWindow):
    def __init__(self):
        super().__init__()
        current_dir = os.path.dirname(os.path.abspath(__file__))
        icon_path = os.path.join(current_dir, "ikona.ico")
        self.app_icon = QIcon(icon_path)
        if not self.app_icon.isNull():
            self.setWindowIcon(self.app_icon)
            app = QApplication.instance()
            if app is not None:
                app.setWindowIcon(self.app_icon)

        try:
            locale.setlocale(locale.LC_NUMERIC, "en_US.UTF-8")
        except locale.Error:
            pass

        self.setWindowTitle(f"{NAZOV_APP} · v0.5")
        self.setStyleSheet(APP_STYLE)
        self.setMinimumSize(1024, 680)

        self.settings = QSettings(NAZOV_FIRMY, NAZOV_APP)
        self.table_initialized = False
        self.row_ids = []
        self.last_sorted_column = None
        self.sort_order = Qt.SortOrder.AscendingOrder
        self.editing_fiscal_year = None

        self.main_widget = QWidget()
        self.main_widget.setObjectName("centralRoot")
        self.main_widget.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)

        self.main_layout = QVBoxLayout(self.main_widget)
        self.main_layout.setContentsMargins(18, 16, 18, 18)
        self.main_layout.setSpacing(12)

        self.init_ui()
        self.setCentralWidget(self.main_widget)
        self.fit_window_to_screen()

        self.setup_database()
        self.load_initial_data()
        self.load_unique_clients()
        self.load_unique_staff()

        saved_font = self.settings.value("velkost_pisma")
        if saved_font:
            try:
                size = int(saved_font)
                idx = {10: 0, 12: 1, 14: 2}.get(size, 1)
                self.font_combobox.blockSignals(True)
                self.font_combobox.setCurrentIndex(idx)
                self.font_combobox.blockSignals(False)
                self.zmen_velkost_pisma(size, ulozit_nastavenie=False)
            except (TypeError, ValueError):
                self.zmen_velkost_pisma(12, ulozit_nastavenie=False)
        else:
            self.zmen_velkost_pisma(12, ulozit_nastavenie=False)

        QTimer.singleShot(0, self._finalize_layout)

    def fit_window_to_screen(self):
        """Prispôsobí okno dostupnej veľkosti monitora."""
        screen = self.screen() or QApplication.primaryScreen()
        if not screen:
            self.setGeometry(X_POSITION, Y_POSITION, WIDTH, HEIGHT)
            return

        geo = screen.availableGeometry()
        margin = 24
        target_w = min(max(int(geo.width() * 0.94), 1100), geo.width() - margin)
        target_h = min(max(int(geo.height() * 0.92), 720), geo.height() - margin)
        x = geo.x() + (geo.width() - target_w) // 2
        y = geo.y() + (geo.height() - target_h) // 2
        self.setGeometry(x, y, target_w, target_h)

    def _finalize_layout(self):
        """Po zobrazení doladí splitter a stĺpce tabuľky."""
        if hasattr(self, "main_splitter"):
            total = max(self.main_splitter.width(), 1)
            left = max(int(total * 0.30), 320)
            right = max(total - left, 500)
            self.main_splitter.setSizes([left, right])
        self._stretch_table_columns()

    def _stretch_table_columns(self):
        header = self.DGZoznam.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        header.setStretchLastSection(True)
        self.DGZoznam.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)

    def _style_calendar(self):
        """Biele pozadie kalendára + zelené prepínanie mesiacov (bez systémovej modrej)."""
        self.Calendar.setStyleSheet(CALENDAR_STYLE)

        palette = self.Calendar.palette()
        palette.setColor(QPalette.ColorRole.Base, QColor("#ffffff"))
        palette.setColor(QPalette.ColorRole.Window, QColor("#ffffff"))
        palette.setColor(QPalette.ColorRole.Text, QColor("#1f2d26"))
        palette.setColor(QPalette.ColorRole.Highlight, QColor("#2d6a4f"))
        palette.setColor(QPalette.ColorRole.HighlightedText, QColor("#ffffff"))
        self.Calendar.setPalette(palette)
        self.Calendar.setAutoFillBackground(True)

        nav = self.Calendar.findChild(QWidget, "qt_calendar_navigationbar")
        if nav is not None:
            nav.setStyleSheet("background-color: #2d6a4f; border: none;")
            nav_palette = nav.palette()
            nav_palette.setColor(QPalette.ColorRole.Window, QColor("#2d6a4f"))
            nav_palette.setColor(QPalette.ColorRole.Button, QColor("#2d6a4f"))
            nav.setPalette(nav_palette)
            nav.setAutoFillBackground(True)

        for btn in self.Calendar.findChildren(QToolButton):
            btn.setStyleSheet(
                "QToolButton { color: #ffffff; background: transparent; border: none; padding: 4px; }"
                "QToolButton:hover { background-color: rgba(255,255,255,0.18); border-radius: 6px; }"
            )

        for spin in self.Calendar.findChildren(QSpinBox):
            spin.setStyleSheet(
                "QSpinBox { background-color: #2d6a4f; color: #ffffff; border: none; }"
            )

        for menu in self.Calendar.findChildren(QMenu):
            menu.setStyleSheet(
                "QMenu { background-color: #ffffff; color: #1b4332; border: 1px solid #c9d8cc; }"
                "QMenu::item { background-color: #ffffff; color: #1b4332; padding: 6px 18px; }"
                "QMenu::item:selected { background-color: #d8eee1; color: #1b4332; }"
            )
            mp = menu.palette()
            mp.setColor(QPalette.ColorRole.Window, QColor("#ffffff"))
            mp.setColor(QPalette.ColorRole.Base, QColor("#ffffff"))
            mp.setColor(QPalette.ColorRole.Text, QColor("#1b4332"))
            mp.setColor(QPalette.ColorRole.ButtonText, QColor("#1b4332"))
            menu.setPalette(mp)

        weekend_fmt = QTextCharFormat()
        weekend_fmt.setForeground(QColor("#bc4749"))
        self.Calendar.setWeekdayTextFormat(Qt.DayOfWeek.Saturday, weekend_fmt)
        self.Calendar.setWeekdayTextFormat(Qt.DayOfWeek.Sunday, weekend_fmt)


        weekday_fmt = QTextCharFormat()
        weekday_fmt.setForeground(QColor("#1f2d26"))
        for day in (
            Qt.DayOfWeek.Monday, Qt.DayOfWeek.Tuesday, Qt.DayOfWeek.Wednesday,
            Qt.DayOfWeek.Thursday, Qt.DayOfWeek.Friday,
        ):
            self.Calendar.setWeekdayTextFormat(day, weekday_fmt)

        QTimer.singleShot(0, self._add_today_to_calendar_nav)

    def _add_today_to_calendar_nav(self):
        """Malé Today tlačidlo vpravo v zelenej lište kalendára."""
        nav = self.Calendar.findChild(QWidget, "qt_calendar_navigationbar")
        if nav is None:
            return

        if getattr(self, "_today_in_nav", False) and self.btnToday is not None:
            return

        self.btnToday = QPushButton("Today", nav)
        self.btnToday.setObjectName("btnCalendarToday")
        self.btnToday.setFixedHeight(24)
        self.btnToday.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btnToday.clicked.connect(self.btnToday_Click)

        layout = nav.layout()
        if layout is not None:
            insert_idx = max(0, layout.count() - 1)
            layout.insertWidget(insert_idx, self.btnToday)
            self._today_in_nav = True

    def showEvent(self, event):
        super().showEvent(event)
        QTimer.singleShot(0, self._finalize_layout)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self._stretch_table_columns()

    def _make_metric_card(self, caption, accent):
        card = QFrame()
        card.setObjectName("metricCard")
        card.setMinimumHeight(72)
        card.setMaximumHeight(88)
        layout = QVBoxLayout(card)
        layout.setContentsMargins(10, 8, 10, 8)
        layout.setSpacing(2)

        accent_bar = QFrame()
        accent_bar.setFixedHeight(3)
        accent_bar.setStyleSheet(f"background-color: {accent}; border: none; border-radius: 2px;")
        layout.addWidget(accent_bar)

        caption_lbl = QLabel(caption)
        caption_lbl.setObjectName("metricCaption")
        layout.addWidget(caption_lbl)

        value_lbl = QLabel("0.00 £")
        value_lbl.setObjectName("metricValue")
        value_lbl.setTextFormat(Qt.TextFormat.RichText)
        layout.addWidget(value_lbl)
        layout.addStretch()

        add_shadow(card, blur=12, dy=2, alpha=16)
        return card, value_lbl

    def _make_field(self, label_text, placeholder="", parent_layout=None):
        wrap = QVBoxLayout()
        wrap.setSpacing(4)
        lbl = QLabel(label_text)
        lbl.setObjectName("fieldLabel")
        edit = QLineEdit()
        edit.setPlaceholderText(placeholder)
        wrap.addWidget(lbl)
        wrap.addWidget(edit)
        if parent_layout is not None:
            parent_layout.addLayout(wrap)
        return edit

    def init_ui(self):
        # ---- Header ----
        header = QFrame()
        header.setObjectName("headerBar")
        header.setMinimumHeight(72)
        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(20, 12, 16, 12)
        header_layout.setSpacing(16)

        brand_col = QVBoxLayout()
        brand_col.setSpacing(0)
        title_label = QLabel(NAZOV_APP)
        title_label.setObjectName("brandTitle")
        subtitle = QLabel("Garden finance · modern workspace")
        subtitle.setObjectName("brandSubtitle")
        brand_col.addWidget(title_label)
        brand_col.addWidget(subtitle)
        header_layout.addLayout(brand_col)
        header_layout.addStretch()

        year_lbl = QLabel("Year")
        year_lbl.setObjectName("headerMeta")
        header_layout.addWidget(year_lbl)

        self.cbRok = QComboBox()
        self.cbRok.setMinimumWidth(90)
        self.cbRok.addItems([str(year) for year in range(2020, 2050)])
        self.cbRok.currentIndexChanged.connect(self.cbRok_SelectedIndexChanged)
        header_layout.addWidget(self.cbRok)

        month_lbl = QLabel("Month")
        month_lbl.setObjectName("headerMeta")
        header_layout.addWidget(month_lbl)

        self.cbMonths = QComboBox()
        self.cbMonths.setMinimumWidth(120)
        self.cbMonths.addItems([
            "All", "January", "February", "March", "April", "May", "June",
            "July", "August", "September", "October", "November", "December",
        ])
        self.cbMonths.currentIndexChanged.connect(self.cbMonths_SelectedIndexChanged)
        header_layout.addWidget(self.cbMonths)

        font_lbl = QLabel("Font")
        font_lbl.setObjectName("headerMeta")
        header_layout.addWidget(font_lbl)

        self.font_combobox = QComboBox()
        self.font_combobox.addItems(["Small", "Medium", "Large"])
        self.font_combobox.setCurrentIndex(1)
        self.font_combobox.currentIndexChanged.connect(self.on_font_size_change)
        header_layout.addWidget(self.font_combobox)

        self.btnExport = QPushButton("Export Excel")
        self.btnExport.setObjectName("btnGhost")
        self.btnExport.clicked.connect(self.export_to_excel)
        header_layout.addWidget(self.btnExport)

        add_shadow(header, blur=28, dy=6, alpha=40)
        self.main_layout.addWidget(header)

        # ---- Search toolbar ----
        toolbar = QFrame()
        toolbar.setObjectName("toolbarCard")
        search_layout = QHBoxLayout(toolbar)
        search_layout.setContentsMargins(14, 10, 14, 10)
        search_layout.setSpacing(10)

        search_title = QLabel("Advanced search")
        search_title.setObjectName("sectionTitle")
        search_layout.addWidget(search_title)

        self.cbColumns = QComboBox()
        self.cbColumns.addItems([
            "Date", "CLIENTS NAME", "Cash", "CHECK", "BANK TRANSFER", "Expenses",
            "EXPENSES COSTS", "CashForStaff", "CashForStaffName",
        ])
        self.cbColumns.setCurrentIndex(1)
        self.cbColumns.setMinimumWidth(140)
        search_layout.addWidget(self.cbColumns)

        self.txtSearch = QLineEdit()
        self.txtSearch.setPlaceholderText("Search text…")
        self.txtSearch.returnPressed.connect(self.advanced_search)
        search_layout.addWidget(self.txtSearch, 1)

        from_lbl = QLabel("From")
        from_lbl.setObjectName("fieldLabel")
        search_layout.addWidget(from_lbl)
        self.start_date_edit = QDateEdit()
        self.start_date_edit.setCalendarPopup(True)
        self.start_date_edit.setDisplayFormat(DATE_UI)
        self.start_date_edit.setDate(QDate.currentDate().addMonths(-12))
        self.start_date_edit.dateChanged.connect(self.reload_data_based_on_dates)
        search_layout.addWidget(self.start_date_edit)

        to_lbl = QLabel("To")
        to_lbl.setObjectName("fieldLabel")
        search_layout.addWidget(to_lbl)
        self.end_date_edit = QDateEdit()
        self.end_date_edit.setCalendarPopup(True)
        self.end_date_edit.setDisplayFormat(DATE_UI)
        self.end_date_edit.setDate(QDate.currentDate())
        self.end_date_edit.dateChanged.connect(self.reload_data_based_on_dates)
        search_layout.addWidget(self.end_date_edit)

        btn_search = QPushButton("Search")
        btn_search.setObjectName("btnSecondary")
        btn_search.clicked.connect(self.advanced_search)
        search_layout.addWidget(btn_search)

        add_shadow(toolbar, blur=18, dy=3, alpha=20)
        self.main_layout.addWidget(toolbar)

        # ---- Splitter: form | results ----
        self.main_splitter = QSplitter(Qt.Orientation.Horizontal)
        self.main_splitter.setChildrenCollapsible(False)
        self.main_splitter.setHandleWidth(8)

        # Left panel — scrollable, aby tlačidlá neboli odrezané
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(0, 0, 6, 0)
        left_layout.setSpacing(0)

        left_scroll = QScrollArea()
        left_scroll.setObjectName("leftFormScroll")
        left_scroll.setWidgetResizable(True)
        left_scroll.setFrameShape(QFrame.Shape.NoFrame)
        left_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        left_scroll.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        left_scroll.setStyleSheet(
            "QScrollArea, QScrollArea > QWidget, QScrollArea > QWidget > QWidget {"
            " background-color: #ffffff; border: none; }"
        )

        form_card = QFrame()
        form_card.setObjectName("formCard")
        form_layout = QVBoxLayout(form_card)
        form_layout.setContentsMargins(14, 12, 14, 14)
        form_layout.setSpacing(8)

        form_title = QLabel("Transaction details")
        form_title.setObjectName("sectionTitle")
        form_layout.addWidget(form_title)

        self.Calendar = QCalendarWidget()
        self.Calendar.setGridVisible(False)
        self.Calendar.setVerticalHeaderFormat(QCalendarWidget.VerticalHeaderFormat.NoVerticalHeader)
        self.Calendar.clicked.connect(self.Calendar_DateSelected)
        self.Calendar.setMinimumHeight(200)
        self.Calendar.setMaximumHeight(230)
        self._style_calendar()
        form_layout.addWidget(self.Calendar)

        self.txtSelectDate = self._make_field("Date", DATE_UI, form_layout)
        self.txtClients = self._make_field("Client", "Client name", form_layout)

        income_title = QLabel("Income")
        income_title.setObjectName("sectionTitle")
        form_layout.addWidget(income_title)

        income_row = QHBoxLayout()
        income_row.setSpacing(8)
        self.txtCash = self._make_field("Cash", "0.00", income_row)
        self.txtCheck = self._make_field("Check", "0.00", income_row)
        self.txtBank = self._make_field("Bank transfer", "0.00", income_row)
        form_layout.addLayout(income_row)

        expense_title = QLabel("Expenses")
        expense_title.setObjectName("sectionTitle")
        form_layout.addWidget(expense_title)

        exp_row1 = QHBoxLayout()
        exp_row1.setSpacing(8)
        self.txtExpenses = self._make_field("Description", "Expense description", exp_row1)
        self.txtExpensesCost = self._make_field("Cost", "0.00", exp_row1)
        form_layout.addLayout(exp_row1)

        exp_row2 = QHBoxLayout()
        exp_row2.setSpacing(8)
        self.txtCashForStaffName = self._make_field("Staff name", "Name", exp_row2)
        self.txtCashForStaff = self._make_field("Cash for staff", "0.00", exp_row2)
        form_layout.addLayout(exp_row2)

        buttons_layout = QHBoxLayout()
        buttons_layout.setSpacing(8)
        buttons_layout.setContentsMargins(0, 6, 0, 4)
        self.btnPridaj = QPushButton("Add")
        self.btnPridaj.setObjectName("btnPrimary")
        self.btnPridaj.setEnabled(True)
        self.btnPridaj.clicked.connect(self.btnPridaj_Click)

        self.btnEdit = QPushButton("Edit")
        self.btnEdit.setObjectName("btnSecondary")
        self.btnEdit.setEnabled(False)
        self.btnEdit.clicked.connect(self.btnEdit_Click)

        self.btnDelete = QPushButton("Delete")
        self.btnDelete.setObjectName("btnDanger")
        self.btnDelete.setEnabled(False)
        self.btnDelete.clicked.connect(self.btnDelete_Click)

        buttons_layout.addWidget(self.btnPridaj)
        buttons_layout.addWidget(self.btnEdit)
        buttons_layout.addWidget(self.btnDelete)
        form_layout.addLayout(buttons_layout)
        form_layout.addStretch(1)

        add_shadow(form_card, blur=16, dy=3, alpha=18)
        left_scroll.setWidget(form_card)
        left_layout.addWidget(left_scroll)
        self.main_splitter.addWidget(left_panel)

        # Right panel
        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(6, 0, 0, 0)
        right_layout.setSpacing(10)

        self.results_group = QFrame()
        self.results_group.setObjectName("tableCard")
        results_layout = QVBoxLayout(self.results_group)
        results_layout.setContentsMargins(12, 12, 12, 12)
        results_layout.setSpacing(10)
        self.results_layout = results_layout

        self.results_title = QLabel("Results")
        self.results_title.setObjectName("sectionTitle")
        results_layout.addWidget(self.results_title)

        metrics_row = QHBoxLayout()
        metrics_row.setSpacing(8)
        metrics_row.setContentsMargins(0, 0, 0, 0)

        cash_card, self.lblCashResult = self._make_metric_card("Cash", "#40916c")
        check_card, self.lblCheckResult = self._make_metric_card("Check", "#52b788")
        bank_card, self.lblBankResult = self._make_metric_card("Bank transfer", "#74c69d")
        income_card, self.lblTotalIncome = self._make_metric_card("Total income", "#1b4332")
        exp_card, self.lblExpensesResult = self._make_metric_card("Expenses", "#bc4749")
        staff_card, self.lblStaffCost = self._make_metric_card("Staff cost", "#e76f51")
        total_exp_card, self.lblTotalExpenses = self._make_metric_card("Total expenses", "#9b2226")
        profit_card, self.lblGrossProfit = self._make_metric_card("Gross profit", "#2a9d8f")
        margin_card, self.lblProfitMargin = self._make_metric_card("Profit margin", "#264653")

        self.lblGrossProfit.setToolTip("Celkový zisk pred odpočítaním nákladov.")
        self.lblTotalExpenses.setToolTip("Spočítanie Staff cost + Expenses.")
        self.lblProfitMargin.setToolTip("Percentuálny podiel zisku na celkových príjmoch.")

        for card in (
            cash_card, check_card, bank_card, income_card, exp_card,
            staff_card, total_exp_card, profit_card, margin_card,
        ):
            card.setMinimumWidth(112)
            card.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
            metrics_row.addWidget(card, 1)

        metrics_wrap = QWidget()
        metrics_wrap.setObjectName("metricsWrap")
        metrics_wrap.setLayout(metrics_row)
        metrics_wrap.setAutoFillBackground(True)
        metrics_palette = metrics_wrap.palette()
        metrics_palette.setColor(QPalette.ColorRole.Window, QColor("#ffffff"))
        metrics_palette.setColor(QPalette.ColorRole.Base, QColor("#ffffff"))
        metrics_wrap.setPalette(metrics_palette)

        metrics_scroll = QScrollArea()
        metrics_scroll.setObjectName("metricsScroll")
        metrics_scroll.setWidgetResizable(True)
        metrics_scroll.setFrameShape(QFrame.Shape.NoFrame)
        metrics_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        metrics_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        metrics_scroll.setMinimumHeight(96)
        metrics_scroll.setMaximumHeight(108)
        metrics_scroll.setStyleSheet(
            "QScrollArea, QScrollArea > QWidget, QScrollArea > QWidget > QWidget,"
            " QWidget#metricsWrap { background-color: #ffffff; border: none; }"
        )
        metrics_scroll.setWidget(metrics_wrap)
        results_layout.addWidget(metrics_scroll)

        self.DGZoznam = QTableWidget()
        self.DGZoznam.setObjectName("DGZoznam")
        self.DGZoznam.setAlternatingRowColors(True)
        self.DGZoznam.setColumnCount(9)
        self.DGZoznam.setHorizontalHeaderLabels(TABLES_2025_MORE)
        self.DGZoznam.horizontalHeader().setHighlightSections(False)
        self.DGZoznam.verticalHeader().setVisible(False)
        self.DGZoznam.setShowGrid(False)
        self.DGZoznam.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.DGZoznam.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.DGZoznam.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.DGZoznam.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.DGZoznam.setSortingEnabled(False)
        self.DGZoznam.cellClicked.connect(self.DGZoznam_CellClick)
        self.DGZoznam.cellDoubleClicked.connect(self.DGZoznam_CellContentDoubleClick)
        self.DGZoznam.horizontalHeader().setSortIndicatorShown(True)
        self.DGZoznam.horizontalHeader().sectionClicked.connect(self.handle_header_click)
        self._stretch_table_columns()
        results_layout.addWidget(self.DGZoznam, 1)

        add_shadow(self.results_group, blur=16, dy=3, alpha=18)
        right_layout.addWidget(self.results_group)
        self.main_splitter.addWidget(right_panel)

        self.main_splitter.setStretchFactor(0, 3)
        self.main_splitter.setStretchFactor(1, 7)
        self.main_splitter.setSizes([360, 900])
        self.main_layout.addWidget(self.main_splitter, 1)

        # Force light popup views (Year/Month/Font) — no black dropdowns
        for combo in (self.cbRok, self.cbMonths, self.font_combobox, self.cbColumns):
            view = combo.view()
            view.setStyleSheet(
                "background-color: #ffffff; color: #1f2d26;"
                " selection-background-color: #d8eee1; selection-color: #1b4332;"
            )
            view.setAutoFillBackground(True)
            vp = view.palette()
            vp.setColor(QPalette.ColorRole.Base, QColor("#ffffff"))
            vp.setColor(QPalette.ColorRole.Text, QColor("#1f2d26"))
            vp.setColor(QPalette.ColorRole.Highlight, QColor("#d8eee1"))
            vp.setColor(QPalette.ColorRole.HighlightedText, QColor("#1b4332"))
            view.setPalette(vp)

        self.table_initialized = True
        self.nastav_tucne_pismo_pre_vysledky_textboxe()

    def reload_data_based_on_dates(self):
        self.advanced_search()

    def export_to_excel(self):
        try:
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet.title = "DataExport_" + self.current_year

            headers = [self.DGZoznam.horizontalHeaderItem(i).text() for i in range(self.DGZoznam.columnCount())]
            sheet.append(headers)

            for row in range(self.DGZoznam.rowCount()):
                row_data = []
                for column in range(self.DGZoznam.columnCount()):
                    item = self.DGZoznam.item(row, column)
                    row_data.append(item.text() if item else "")
                sheet.append(row_data)

            default_filename = f"exported_data_{self.current_year}.xlsx"
            excel_file_path, _ = QFileDialog.getSaveFileName(
                self, "Exportovať do Excelu", default_filename, "Excel Files (*.xlsx);;All Files (*)"
            )

            if excel_file_path:
                workbook.save(excel_file_path)
                QMessageBox.information(self, "Export", f"Úspešne exportované do súboru {excel_file_path}")

        except Exception as ex:
            QMessageBox.critical(self, "Chyba", f"Chyba pri exportovaní: {str(ex)}")

    def on_font_size_change(self, index):
        if index == 0:
            self.zmen_velkost_pisma(10)
        elif index == 1:
            self.zmen_velkost_pisma(12)
        elif index == 2:
            self.zmen_velkost_pisma(14)

    def cbMonths_SelectedIndexChanged(self, index):
        try:
            selected_month = self.cbMonths.currentText()
            if selected_month == "All":
                self.NacitajDatabazu()
            else:
                month_index = self.cbMonths.currentIndex()
                if month_index > 0:
                    self.NacitajDatabazuPodlaMesiac(month_index, self.cbRok.currentText())
        except Exception as ex:
            QMessageBox.critical(self, "Error", f"Chyba pri zmene mesiaca: {str(ex)}")

    def NacitajDatabazuPodlaMesiac(self, month_number, year):
        try:
            self.txtSearch.clear()
            cursor = self.conn.cursor()
            cursor.execute(
                f"""
                SELECT * FROM TableGarden{year}
                WHERE strftime('%m', Date) = ?
                AND strftime('%Y', Date) = ?
                ORDER BY Date DESC, Id DESC
                """,
                (f"{month_number:02d}", year),
            )
            rows = cursor.fetchall()
            self._fill_table_rows(rows)
            self.StatistikaVypocet()
        except Exception as ex:
            QMessageBox.critical(self, "Error", str(ex))

    def advanced_search(self):
        try:
            if self.txtSearch.text() != "":
                self.MazanietxtPopridaniDoSql()
                column_name = self.cbColumns.currentText()
                search_text = self.txtSearch.text().strip()
                start_date = self.start_date_edit.date().toString(DATE_DB)
                end_date = self.end_date_edit.date().toString(DATE_DB)

                cursor = self.conn.cursor()
                query = f"SELECT * FROM TableGarden{self.current_year} WHERE Date BETWEEN ? AND ? "
                params = [start_date, end_date]

                if search_text:
                    query += f"AND [{column_name}] LIKE ? "
                    params.append(f"%{search_text}%")

                query += "ORDER BY Date DESC, Id DESC"

                cursor.execute(query, params)
                rows = cursor.fetchall()
                self._fill_table_rows(rows)
            else:
                self.NacitajDatabazu()

            self.StatistikaVypocet()
        except Exception as ex:
            QMessageBox.critical(self, "Chyba", f"Chyba počas vyhľadávania: {str(ex)}")

    def handle_header_click(self, column):
        if self.last_sorted_column == column:
            self.sort_order = (
                Qt.SortOrder.DescendingOrder
                if self.sort_order == Qt.SortOrder.AscendingOrder
                else Qt.SortOrder.AscendingOrder
            )
        else:
            self.sort_order = Qt.SortOrder.AscendingOrder

        for row in range(self.DGZoznam.rowCount()):
            item = self.DGZoznam.item(row, column)
            if not item:
                continue
            if column == 0:
                parsed = parse_date(item.text())
                if parsed.isValid():
                    item.setData(Qt.ItemDataRole.EditRole, parsed.toString(DATE_DB))
            else:
                try:
                    item.setData(Qt.ItemDataRole.EditRole, float(item.text()))
                except ValueError:
                    pass

        self.DGZoznam.sortItems(column, self.sort_order)
        self.last_sorted_column = column
        if column == 0:
            self._restore_date_column_format()
        self._sync_row_ids_after_sort()

    def nastav_tucne_pismo_pre_vysledky_textboxe(self):
        velkost = int(self.settings.value("velkost_pisma", 12))
        bold_font = QFont("Segoe UI Semibold", max(velkost, 11))
        bold_font.setBold(True)
        for label in (
            self.lblCashResult, self.lblCheckResult, self.lblBankResult,
            self.lblTotalIncome, self.lblExpensesResult, self.lblStaffCost,
            self.lblTotalExpenses, self.lblGrossProfit, self.lblProfitMargin,
        ):
            label.setFont(bold_font)

    def _font_stylesheet(self, velkost):
        """Dynamické font-size v stylesheet (pevné px v CSS inak prebíjajú setFont)."""
        base = {10: 11, 12: 13, 14: 16}.get(velkost, 13)
        title = base + 9
        section = base + 1
        caption = max(base - 2, 10)
        meta = max(base - 1, 11)
        metric = base + 3
        button = base
        table = base
        cal = max(base - 1, 11)

        return f"""
        QLabel#brandTitle {{ font-size: {title}px; }}
        QLabel#brandSubtitle {{ font-size: {caption}px; }}
        QLabel#headerMeta {{ font-size: {meta}px; }}
        QLabel#sectionTitle {{ font-size: {section}px; }}
        QLabel#fieldLabel {{ font-size: {caption}px; }}
        QLabel#metricCaption {{ font-size: {caption}px; }}
        QLabel#metricValue {{ font-size: {metric}px; }}
        QLineEdit, QComboBox, QDateEdit {{ font-size: {base}px; }}
        QComboBox QAbstractItemView {{ font-size: {base}px; }}
        QComboBox QAbstractItemView::item {{ font-size: {base}px; min-height: {base + 14}px; }}
        QPushButton {{ font-size: {button}px; padding: {max(6, base - 5)}px {max(12, base)}px; }}
        QTableWidget {{ font-size: {table}px; }}
        QHeaderView::section {{ font-size: {table}px; padding: {max(6, base - 5)}px 6px; }}
        QCalendarWidget QToolButton {{ font-size: {base}px; }}
        QCalendarWidget QAbstractItemView {{ font-size: {cal}px; }}
        QCalendarWidget QSpinBox {{ font-size: {base}px; }}
        """

    def zmen_velkost_pisma(self, velkost, ulozit_nastavenie=True):
        """Zmení veľkosť písma v celej aplikácii (stylesheet + QFont)."""
        self._current_font_size = velkost

        font = QFont("Segoe UI")
        font.setPointSize(velkost)
        app = QApplication.instance()
        if app is not None:
            app.setFont(font)
        self.setFont(font)

        # Stylesheet musí ísť neskôr ako APP_STYLE, aby prepísal pevné font-size
        self.setStyleSheet(APP_STYLE + "\n" + self._font_stylesheet(velkost))

        # Kalendár má vlastný stylesheet — doplniť aj fonty
        if hasattr(self, "Calendar"):
            self.Calendar.setStyleSheet(CALENDAR_STYLE + "\n" + self._font_stylesheet(velkost))

        for widget in self.findChildren(QWidget):
            wfont = QFont(widget.font())
            # Brand title väčší
            if isinstance(widget, QLabel) and widget.objectName() == "brandTitle":
                wfont.setFamily("Segoe UI Semibold")
                wfont.setPointSize(velkost + 8)
            elif isinstance(widget, QLabel) and widget.objectName() == "metricValue":
                wfont.setFamily("Segoe UI Semibold")
                wfont.setPointSize(velkost + 2)
            elif isinstance(widget, QLabel) and widget.objectName() == "sectionTitle":
                wfont.setFamily("Segoe UI Semibold")
                wfont.setPointSize(velkost + 1)
            else:
                wfont.setPointSize(velkost)
            widget.setFont(wfont)

        self.nastav_tucne_pismo_pre_vysledky_textboxe()

        if ulozit_nastavenie:
            self.settings.setValue("velkost_pisma", velkost)

    def setup_database(self):
        if getattr(sys, "frozen", False):
            app_dir = os.path.dirname(sys.executable)
        else:
            app_dir = os.path.dirname(os.path.abspath(__file__))

        self.FullPath = os.path.join(app_dir, "DbGarden.db")
        self.ConnectionString = f"Data Source={self.FullPath};"
        self.conn = sqlite3.connect(self.FullPath)

        current_date = QDate.currentDate()
        year = current_date.year()
        month = current_date.month()
        self.current_year = str(year if month >= 4 else year - 1)

        self.cbRok.setCurrentText(self.current_year)

        if not self.check_table_exists(self.current_year):
            self.create_table_if_not_exists(self.current_year)

        self.results_title.setText(f"Results for year {self.current_year}")

    def check_table_exists(self, year):
        cursor = self.conn.cursor()
        cursor.execute(
            f"SELECT name FROM sqlite_master WHERE type='table' AND name='TableGarden{year}';"
        )
        return cursor.fetchone() is not None

    def create_table_if_not_exists(self, year):
        cursor = self.conn.cursor()
        cursor.execute(
            f"""
            CREATE TABLE IF NOT EXISTS TableGarden{year} (
                Id INTEGER PRIMARY KEY AUTOINCREMENT,
                Date TEXT,
                [CLIENTS NAME] TEXT,
                Cash REAL,
                [CHECK] REAL,
                [BANK TRANSFER] REAL,
                Expenses TEXT,
                [EXPENSES COSTS] REAL,
                CashForStaff REAL,
                CashForStaffName TEXT
            )
            """
        )
        self.conn.commit()

    def load_initial_data(self):
        self.NacitajDatabazu()

    def btnDelete_Click(self):
        reply = QMessageBox.question(
            self,
            "Delete Record",
            f"Delete this record , nr. {self.Riadok} ?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if reply == QMessageBox.StandardButton.Yes:
            self.VymazatZaznam()
            self.NacitajDatabazu()

    def DGZoznam_CellClick(self, row, column):
        None

    def DGZoznam_CellContentDoubleClick(self, row, column):
        date_item = self.DGZoznam.item(row, 0)
        record_id = date_item.data(Qt.ItemDataRole.UserRole) if date_item else None
        self.Riadok = record_id if record_id is not None else self.row_ids[row]
        self.editing_fiscal_year = self.current_year
        if row != -1:
            self.MazanietxtPopridaniDoSql()
            self.txtSearch.clear()

            self.txtSelectDate.setText(to_ui_date(self.DGZoznam.item(row, 0).text()))
            self.txtClients.setText(self.DGZoznam.item(row, 1).text())
            self.txtCash.setText(self.DGZoznam.item(row, 2).text())
            self.txtCheck.setText(self.DGZoznam.item(row, 3).text())
            self.txtBank.setText(self.DGZoznam.item(row, 4).text())
            self.txtExpenses.setText(self.DGZoznam.item(row, 5).text())
            self.txtExpensesCost.setText(self.DGZoznam.item(row, 6).text())

            if self.current_year >= "2025":
                self.txtCashForStaff.setText(self.DGZoznam.item(row, 7).text())
                self.txtCashForStaffName.setText(self.DGZoznam.item(row, 8).text())

            self.btnEdit.setEnabled(True)
            self.btnDelete.setEnabled(True)

    def Calendar_DateSelected(self):
        selected_date = self.Calendar.selectedDate()
        self.txtSelectDate.setText(selected_date.toString(DATE_UI))

    def btnToday_Click(self):
        today = QDate.currentDate()
        self.Calendar.setSelectedDate(today)
        self.txtSelectDate.setText(today.toString(DATE_UI))

    def reinit_table(self):
        if hasattr(self, "DGZoznam") and self.DGZoznam is not None:
            self.DGZoznam.deleteLater()

        self.DGZoznam = QTableWidget()
        self.DGZoznam.setObjectName("DGZoznam")
        self.DGZoznam.setAlternatingRowColors(True)
        self.DGZoznam.setColumnCount(9)
        self.DGZoznam.setShowGrid(False)
        self.DGZoznam.verticalHeader().setVisible(False)

        if self.current_year >= "2025":
            self.DGZoznam.setHorizontalHeaderLabels(TABLES_2025_MORE)
        else:
            self.DGZoznam.setHorizontalHeaderLabels(TABLES_2025_LESS)

        self.DGZoznam.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.DGZoznam.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.DGZoznam.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.DGZoznam.cellClicked.connect(self.DGZoznam_CellClick)
        self.DGZoznam.cellDoubleClicked.connect(self.DGZoznam_CellContentDoubleClick)
        self.last_sorted_column = None
        self.sort_order = Qt.SortOrder.AscendingOrder
        self.DGZoznam.horizontalHeader().setSortIndicatorShown(True)
        self.DGZoznam.horizontalHeader().sectionClicked.connect(self.handle_header_click)
        self._stretch_table_columns()
        self.results_layout.addWidget(self.DGZoznam, 1)
        self.table_initialized = True

    def _numeric_field_has_value(self, field):
        text = field.text().strip()
        if not text:
            return False
        try:
            return float(text) != 0
        except ValueError:
            return True

    def _validate_record(self):
        """Validácia záznamu podľa typu (príjem / výdavok / staff cash)."""
        date_text = self.txtSelectDate.text().strip()
        date_ok = bool(date_text) and parse_date(date_text).isValid()

        has_cost = self._numeric_field_has_value(self.txtExpensesCost)
        has_staff_cash = self._numeric_field_has_value(self.txtCashForStaff)
        has_description = bool(self.txtExpenses.text().strip())
        has_staff_name = bool(self.txtCashForStaffName.text().strip())
        has_client = bool(self.txtClients.text().strip())

        expense_valid = has_cost and date_ok and has_description
        staff_valid = has_staff_cash and date_ok and has_staff_name
        client_valid = date_ok and has_client

        if expense_valid or staff_valid or client_valid:
            return True, ""

        if has_cost:
            missing = []
            if not date_ok:
                missing.append("Date")
            if not has_description:
                missing.append("Description")
            if not missing:
                missing.append("Cost")
            return False, f"For Expenses enter: {', '.join(missing)}"

        if has_staff_cash:
            missing = []
            if not date_ok:
                missing.append("Date")
            if not has_staff_name:
                missing.append("Staff name")
            if not missing:
                missing.append("Cash for staff")
            return False, f"For Cash for staff enter: {', '.join(missing)}"

        return False, "Write at least the Date and the Client"

    def btnPridaj_Click(self):
        valid, message = self._validate_record()
        if not valid:
            QMessageBox.information(self, "Error", message, QMessageBox.StandardButton.Ok)
            return

        fiscal_year = fiscal_year_from_date(self.txtSelectDate.text())
        if not fiscal_year:
            QMessageBox.information(self, "Error", "Enter a valid Date", QMessageBox.StandardButton.Ok)
            return

        reply = QMessageBox.question(
            self,
            "Add Record",
            f"Add this record for the fiscal year {fiscal_year}?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if reply == QMessageBox.StandardButton.Yes:
            self.ZalozitZaznam()
            self.NacitajDatabazu()

    def btnEdit_Click(self):
        valid, message = self._validate_record()
        if not valid:
            QMessageBox.information(self, "Error", message, QMessageBox.StandardButton.Ok)
            return

        reply = QMessageBox.question(
            self,
            "Edit Record",
            "Save changes?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if reply == QMessageBox.StandardButton.Yes:
            self.UpdateZaznam()
            self.NacitajDatabazu()
            self.MazanietxtPopridaniDoSql()

    def _fill_table_rows(self, rows):
        """Naplní tabuľku; dátum zobrazí vo formáte dd.MM.yyyy."""
        self.DGZoznam.setRowCount(len(rows))
        self.row_ids = []
        for row_idx, row in enumerate(rows):
            record_id = row[0]
            self.row_ids.append(record_id)
            for col_idx, value in enumerate(row[1:]):
                numeric_columns = [2, 3, 4, 6, 7]
                if col_idx == 0:
                    parsed = parse_date(value)
                    display_date = to_ui_date(value)
                    item = QTableWidgetItem(display_date)
                    if parsed.isValid():
                        item.setData(Qt.ItemDataRole.EditRole, parsed.toString(DATE_DB))
                    item.setData(Qt.ItemDataRole.UserRole, record_id)
                elif col_idx in numeric_columns:
                    item = QTableWidgetItem()
                    item.setData(Qt.ItemDataRole.EditRole, float(value) if value else 0.0)
                else:
                    item = QTableWidgetItem(str(value) if value is not None else "")
                self.DGZoznam.setItem(row_idx, col_idx, item)

        self._apply_default_date_sort()

    def _restore_date_column_format(self):
        """Po zoradení obnoví zobrazenie dátumu vo formáte dd.MM.yyyy."""
        for row in range(self.DGZoznam.rowCount()):
            item = self.DGZoznam.item(row, 0)
            if not item:
                continue
            parsed = parse_date(item.data(Qt.ItemDataRole.EditRole) or item.text())
            if parsed.isValid():
                item.setText(parsed.toString(DATE_UI))

    def _apply_default_date_sort(self):
        """Predvolené zoradenie podľa dátumu zostupne (dáta už zoradené v SQL)."""
        self.last_sorted_column = 0
        self.sort_order = Qt.SortOrder.DescendingOrder
        self.DGZoznam.horizontalHeader().setSortIndicator(0, self.sort_order)
        self._restore_date_column_format()

    def _sync_row_ids_after_sort(self):
        """Po zoradení tabuľky zosúladí row_ids s aktuálnym poradím riadkov."""
        self.row_ids = []
        for row in range(self.DGZoznam.rowCount()):
            item = self.DGZoznam.item(row, 0)
            record_id = item.data(Qt.ItemDataRole.UserRole) if item else None
            self.row_ids.append(record_id)

    def NacitajDatabazu(self):
        try:
            cursor = self.conn.cursor()
            cursor.execute(f"SELECT * FROM TableGarden{self.current_year} ORDER BY Date DESC, Id DESC")
            rows = cursor.fetchall()
            self._fill_table_rows(rows)
            self.StatistikaVypocet()
        except Exception as ex:
            QMessageBox.critical(self, "Error", str(ex))

    def load_unique_clients(self):
        try:
            cursor = self.conn.cursor()
            cursor.execute(f"SELECT DISTINCT [CLIENTS NAME] FROM TableGarden{self.current_year};")
            rows_current = cursor.fetchall()
            self.unique_clients = [row[0] for row in rows_current if row[0]]

            if not self.unique_clients:
                previous_year = str(int(self.current_year) - 1)
                cursor.execute(
                    f"SELECT name FROM sqlite_master WHERE type='table' AND name='TableGarden{previous_year}';"
                )
                if cursor.fetchone():
                    cursor.execute(
                        f"SELECT DISTINCT [CLIENTS NAME] FROM TableGarden{previous_year};"
                    )
                    rows_previous = cursor.fetchall()
                    self.unique_clients = [row[0] for row in rows_previous if row[0]]

            completer = QCompleter(self.unique_clients, self)
            completer.setCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
            self.txtClients.setCompleter(completer)
        except Exception as ex:
            QMessageBox.critical(self, "Error", f"Chyba pri načítaní klientov: {str(ex)}")

    def load_unique_staff(self):
        if self.cbRok.currentText() >= "2025":
            try:
                cursor = self.conn.cursor()
                cursor.execute(
                    f"SELECT DISTINCT [CashForStaffName] FROM TableGarden{self.current_year};"
                )
                rows_current = cursor.fetchall()
                self.unique_staff = [row[0] for row in rows_current if row[0]]

                completer = QCompleter(self.unique_staff, self)
                completer.setCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
                self.txtCashForStaffName.setCompleter(completer)
            except Exception as ex:
                QMessageBox.critical(self, "Error", f"Chyba pri načítaní klientov: {str(ex)}")

    def _record_field_values(self):
        return (
            to_db_date(self.txtSelectDate.text()),
            self.txtClients.text(),
            float(self.txtCash.text()) if self.txtCash.text() else 0,
            float(self.txtCheck.text()) if self.txtCheck.text() else 0,
            float(self.txtBank.text()) if self.txtBank.text() else 0,
            self.txtExpenses.text(),
            float(self.txtExpensesCost.text()) if self.txtExpensesCost.text() else 0,
            float(self.txtCashForStaff.text()) if self.txtCashForStaff.text() else 0,
            self.txtCashForStaffName.text(),
        )

    def ZalozitZaznam(self):
        try:
            fiscal_year = fiscal_year_from_date(self.txtSelectDate.text())
            if not fiscal_year:
                QMessageBox.critical(self, "Error", "Invalid date for fiscal year")
                return

            if not self.check_table_exists(fiscal_year):
                self.create_table_if_not_exists(fiscal_year)

            cursor = self.conn.cursor()
            sql = f"""
            INSERT INTO TableGarden{fiscal_year}
            (Date, [CLIENTS NAME], Cash, [CHECK], [BANK TRANSFER], Expenses, [EXPENSES COSTS], CashForStaff, CashForStaffName)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """
            cursor.execute(sql, self._record_field_values())
            self.conn.commit()
            self.MazanietxtPopridaniDoSql()
        except Exception as ex:
            QMessageBox.critical(self, "Error", str(ex))

    def UpdateZaznam(self):
        try:
            new_fiscal_year = fiscal_year_from_date(self.txtSelectDate.text())
            if not new_fiscal_year:
                QMessageBox.critical(self, "Error", "Invalid date for fiscal year")
                return

            old_fiscal_year = self.editing_fiscal_year or self.cbRok.currentText()
            values = self._record_field_values()
            cursor = self.conn.cursor()

            if new_fiscal_year == old_fiscal_year:
                sql = f"""
                UPDATE TableGarden{old_fiscal_year}
                SET Date=?, [CLIENTS NAME]=?, Cash=?, [CHECK]=?, [BANK TRANSFER]=?,
                Expenses=?, [EXPENSES COSTS]=?, CashForStaff=?, CashForStaffName=?
                WHERE Id=?
                """
                cursor.execute(sql, values + (self.Riadok,))
            else:
                if not self.check_table_exists(new_fiscal_year):
                    self.create_table_if_not_exists(new_fiscal_year)
                cursor.execute(f"DELETE FROM TableGarden{old_fiscal_year} WHERE Id=?", (self.Riadok,))
                sql = f"""
                INSERT INTO TableGarden{new_fiscal_year}
                (Date, [CLIENTS NAME], Cash, [CHECK], [BANK TRANSFER], Expenses, [EXPENSES COSTS], CashForStaff, CashForStaffName)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                """
                cursor.execute(sql, values)

            self.conn.commit()
        except Exception as ex:
            QMessageBox.critical(self, "Error", str(ex))

    def VymazatZaznam(self):
        try:
            fiscal_year = self.editing_fiscal_year or self.cbRok.currentText()
            cursor = self.conn.cursor()
            sql = f"DELETE FROM TableGarden{fiscal_year} WHERE Id=?"
            cursor.execute(sql, (self.Riadok,))
            self.conn.commit()
            self.MazanietxtPopridaniDoSql()
        except Exception as ex:
            QMessageBox.critical(self, "Error", str(ex))

    def MazanietxtPopridaniDoSql(self):
        self.editing_fiscal_year = None
        self.txtSelectDate.clear()
        self.txtClients.clear()
        self.txtCash.clear()
        self.txtCheck.clear()
        self.txtBank.clear()
        self.txtExpenses.clear()
        self.txtExpensesCost.clear()
        self.txtCashForStaff.clear()
        self.txtCashForStaffName.clear()
        self.btnDelete.setEnabled(False)
        self.btnEdit.setEnabled(False)
        self.btnPridaj.setEnabled(True)

    def StatistikaVypocet(self):
        try:
            total_cash = 0
            total_check = 0
            total_bank = 0
            total_expenses = 0
            staff_cost = 0

            for row in range(self.DGZoznam.rowCount()):
                cash_item = self.DGZoznam.item(row, 2)
                if cash_item and cash_item.text():
                    total_cash += float(cash_item.text())

                check_item = self.DGZoznam.item(row, 3)
                if check_item and check_item.text():
                    total_check += float(check_item.text())

                bank_item = self.DGZoznam.item(row, 4)
                if bank_item and bank_item.text():
                    total_bank += float(bank_item.text())

                cost_item = self.DGZoznam.item(row, 6)
                if cost_item and cost_item.text():
                    total_expenses += float(cost_item.text())

                staff_cost_item = self.DGZoznam.item(row, 7)
                if staff_cost_item and staff_cost_item.text():
                    staff_cost += float(staff_cost_item.text())

            total_income = total_cash + total_check + total_bank
            gross_profit = total_income - total_expenses - staff_cost
            profit_margin = (gross_profit / total_income * 100) if total_income > 0 else 0
            total_expenses_staff = total_expenses + staff_cost

            self.update_label_with_bold(self.lblCashResult, f"CASH: {self.format_number_with_spaces(total_cash)} £")
            self.update_label_with_bold(self.lblCheckResult, f"CHECK: {self.format_number_with_spaces(total_check)} £")
            self.update_label_with_bold(self.lblBankResult, f"BANK TRANSFER: {self.format_number_with_spaces(total_bank)} £")
            self.update_label_with_bold(self.lblTotalIncome, f"TOTAL INCOME: {self.format_number_with_spaces(total_income)} £")
            self.update_label_with_bold(self.lblExpensesResult, f"EXPENSES: {self.format_number_with_spaces(total_expenses)} £")
            self.update_label_with_bold(self.lblStaffCost, f"STAFF COST: {self.format_number_with_spaces(staff_cost)} £")
            self.update_label_with_bold(self.lblGrossProfit, f"GROSS PROFIT: {self.format_number_with_spaces(gross_profit)} £")
            self.update_label_with_bold(self.lblTotalExpenses, f"TOTAL EXPENSES: {self.format_number_with_spaces(total_expenses_staff)} £")
            self.update_label_with_bold(self.lblProfitMargin, f"PROFIT MARGIN: {self.format_number_with_spaces(profit_margin)}%")
        except Exception as ex:
            QMessageBox.critical(self, "Error", f"Chyba pri Statistike: {str(ex)}")

    def format_number_with_spaces(self, number):
        if number >= 1000 or number <= -1000:
            return f"{number:,.2f}".replace(",", " ")
        return f"{number:.2f}"

    def update_label_with_bold(self, label, text):
        parts = text.split(":")
        if len(parts) == 2:
            label.setText(f"<b>{parts[1].strip()}</b>")
        else:
            label.setText(text)

    def cbRok_SelectedIndexChanged(self, index):
        if hasattr(self, "DGZoznam") and self.table_initialized:
            try:
                self.MazanietxtPopridaniDoSql()
                self.cbMonths.setCurrentIndex(0)
                self.current_year = self.cbRok.currentText()

                if not self.check_table_exists(self.current_year):
                    self.create_table_if_not_exists(self.current_year)

                if self.current_year >= "2025":
                    self.DGZoznam.setHorizontalHeaderLabels(TABLES_2025_MORE)
                else:
                    self.DGZoznam.setHorizontalHeaderLabels(TABLES_2025_LESS)

                self.DGZoznam.clearContents()
                self.DGZoznam.setRowCount(0)
                self.txtSearch.clear()
                self.NacitajDatabazu()
                self.results_title.setText(f"Results for year {self.current_year}")
                self.load_unique_clients()
                self.load_unique_staff()
            except RuntimeError:
                self.reinit_table()
                self.NacitajDatabazu()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")

    # Rovnaká ikona v rohu okna aj na Windows taskbare
    icon_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ikona.ico")
    app_icon = QIcon(icon_path)
    if not app_icon.isNull():
        app.setWindowIcon(app_icon)

    # Svetlá paleta — žiadne systémové čierne popup/scroll pozadia
    light = QPalette()
    light.setColor(QPalette.ColorRole.Window, QColor("#eef4ef"))
    light.setColor(QPalette.ColorRole.WindowText, QColor("#1f2d26"))
    light.setColor(QPalette.ColorRole.Base, QColor("#ffffff"))
    light.setColor(QPalette.ColorRole.AlternateBase, QColor("#f5faf6"))
    light.setColor(QPalette.ColorRole.Text, QColor("#1f2d26"))
    light.setColor(QPalette.ColorRole.Button, QColor("#ffffff"))
    light.setColor(QPalette.ColorRole.ButtonText, QColor("#1f2d26"))
    light.setColor(QPalette.ColorRole.Highlight, QColor("#d8eee1"))
    light.setColor(QPalette.ColorRole.HighlightedText, QColor("#1b4332"))
    light.setColor(QPalette.ColorRole.ToolTipBase, QColor("#ffffff"))
    light.setColor(QPalette.ColorRole.ToolTipText, QColor("#1f2d26"))
    light.setColor(QPalette.ColorGroup.Disabled, QPalette.ColorRole.Text, QColor("#9aa89f"))
    app.setPalette(light)

    window = FrmGarden()
    window.show()
    sys.exit(app.exec())
