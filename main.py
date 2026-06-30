import sys
import pythoncom
from PySide6.QtWidgets import (
    QApplication, QWidget, QMainWindow, QVBoxLayout, QHBoxLayout, QSystemTrayIcon, QMenu,
    QPushButton, QLabel, QFrame, QStackedWidget, QLineEdit, QScrollArea, QGridLayout, QComboBox,
    QFileDialog, QDialog, QMessageBox, QSizePolicy, QTableWidget, QTableWidgetItem,
    QHeaderView, QAbstractItemView
)
from PySide6.QtGui import (
    QFontDatabase, QFont, QColor, QPalette, QIcon, QDesktopServices
)
from PySide6.QtCore import Qt, QObject, Signal, Slot, QThread, QTimer, QUrl, QEvent, QVariantAnimation, QEasingCurve

import ctypes
ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID('Mail AI')


from logic import *

try:
    pythoncom.CoInitialize()
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
except Exception as e:
    outlook = None


def list_outlook_accounts():
    """Top-level Outlook accounts (mailbox / store names) available via local Outlook."""
    if outlook is None:
        return []
    try:
        return [f.Name for f in outlook.Folders]
    except Exception:
        return []


def list_outlook_folders(account_name):
    """Direct subfolders of the given account — matches how extraction selects a folder."""
    if outlook is None or not account_name:
        return []
    try:
        for acct in outlook.Folders:
            if acct.Name.strip().upper() == account_name.strip().upper():
                return [sf.Name for sf in acct.Folders]
    except Exception:
        pass
    return []


def fill_combo(combo, options, current):
    """Populate an (editable) combo with options, preserving/selecting `current`."""
    combo.blockSignals(True)
    combo.clear()
    current = (current or "").strip()
    if current and current not in options:
        combo.addItem(current)
    combo.addItems(options)
    idx = combo.findText(current)
    if idx >= 0:
        combo.setCurrentIndex(idx)
    else:
        combo.setEditText(current)
    combo.blockSignals(False)


MONTH_NAMES = ["January", "February", "March", "April", "May", "June",
               "July", "August", "September", "October", "November", "December"]

csv_file = resource_path("WPIUpdated.csv")
if not os.path.exists(csv_file):
    QMessageBox.critical(None, "Missing Resource", f"Required file not found: WPI.csv\nPlease reinstall the application.")
    sys.exit(1)
csv_dict = merge_custom_zones(load_csv_into_dict(csv_file))
_unlocode = load_unlocode_dict()
for _k, _v in _unlocode.items():
    if _k not in csv_dict:
        csv_dict[_k] = _v
    elif len(set(csv_dict[_k])) > 1 and len(set(_v)) == 1:
        csv_dict[_k] = _v

class ExtractWorker(QObject):
    new_email = Signal(dict)
    done = Signal()

    def __init__(self, generator):
        super().__init__()
        self.generator = generator
        self.running = True
        self.api_error_key = None
        self.limit_reached = False

    def run(self):
        try:
            for email in self.generator:
                if not self.running:
                    break
                if email.get("type") == "api_error":
                    self.api_error_key = email["error_key"]
                    break
                if email.get("type") == "limit_reached":
                    self.limit_reached = True
                    break
                self.new_email.emit(email)
        except Exception as e:
            import traceback
            logger.error(f"Extraction worker crashed: {e}\n{traceback.format_exc()}")
            self.api_error_key = "proxy_error_generic"
        finally:
            self.done.emit()

    def stop(self):
        self.running = False


class UpdateChecker(QObject):
    """Background check against GitHub for a newer release."""
    update_available = Signal(str)

    def run(self):
        version = check_for_update()
        if version:
            self.update_available.emit(version)


class UpdateWorker(QObject):
    """Downloads + stages the update off the UI thread. It is applied on quit, not now."""
    done = Signal(bool, str)

    def run(self):
        try:
            download_update()
            self.done.emit(True, "")
        except Exception as e:
            self.done.emit(False, str(e))


_INTER_FAMILY = None

def get_font(language):
    """UI font: Inter for everything. Chinese uses Source Han Sans. Registers the bundled
    families once and caches Inter's family name."""
    global _INTER_FAMILY
    if _INTER_FAMILY is None:
        fid = QFontDatabase.addApplicationFont(resource_path("Inter/Inter-VariableFont.ttf"))
        fams = QFontDatabase.applicationFontFamilies(fid) if fid != -1 else []
        _INTER_FAMILY = fams[0] if fams else "Segoe UI"

    if language == "中文":
        font_id = QFontDatabase.addApplicationFont(resource_path("SourceHanSansSC-Regular.otf"))
        if font_id != -1:
            return QFont(QFontDatabase.applicationFontFamilies(font_id)[0], 10)
        return QApplication.font()

    return QFont(_INTER_FAMILY, 10)

class GridWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.theme = "dark"

    def set_theme(self, theme):
        self.theme = theme
        self.update()

    def paintEvent(self, event):
        from PySide6.QtGui import QPainter, QPen
        painter = QPainter(self)

        if self.theme == "dark":
            painter.fillRect(self.rect(), QColor("#0a0b0d"))
            pen = QPen(QColor("#141517"))
        else:
            painter.fillRect(self.rect(), QColor("#fafafa"))
            pen = QPen(QColor("#f0f0f1"))

        pen.setWidth(1)
        painter.setPen(pen)

        spacing = 40
        for x in range(0, self.width(), spacing):
            painter.drawLine(x, 0, x, self.height())
        for y in range(0, self.height(), spacing):
            painter.drawLine(0, y, self.width(), y)

class GridStack(QStackedWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.theme = "dark"

    def set_theme(self, theme):
        self.theme = theme
        self.update()

    def paintEvent(self, event):
        from PySide6.QtGui import QPainter, QPen
        painter = QPainter(self)

        if self.theme == "dark":
            painter.fillRect(self.rect(), QColor("#0a0b0d"))
            pen = QPen(QColor("#141517"))
        else:
            painter.fillRect(self.rect(), QColor("#fafafa"))
            pen = QPen(QColor("#f0f0f1"))

        pen.setWidth(1)
        painter.setPen(pen)

        spacing = 40
        for x in range(0, self.width(), spacing):
            painter.drawLine(x, 0, x, self.height())
        for y in range(0, self.height(), spacing):
            painter.drawLine(0, y, self.width(), y)

class SetupWizard(QWidget):
    finished = Signal()
    language_changed = Signal(str)

    def __init__(self, language="English", parent=None):
        super().__init__(parent)
        self.language = language
        config = load_config()

        self.stack = QStackedWidget()
        self.pages_list = []

        welcome = QWidget()
        wl = QVBoxLayout(welcome)
        wl.setAlignment(Qt.AlignCenter)
        wt = QLabel(t("setup_welcome_title", self.language))
        wt.setStyleSheet("font: bold 44px;")
        wt.setAlignment(Qt.AlignCenter)
        ws = QLabel(t("setup_welcome_subtitle", self.language))
        ws.setStyleSheet("font: normal 18px;")
        ws.setAlignment(Qt.AlignCenter)

        lang_row = QHBoxLayout()
        lang_row.setAlignment(Qt.AlignCenter)
        lang_label = QLabel(t("language", self.language))
        lang_label.setStyleSheet("font: 600 16px;")
        self.lang_combo = QComboBox()
        self.lang_combo.addItems(["English", "中文"])
        self.lang_combo.setFixedSize(150, 40)
        self.lang_combo.setFont(get_font(self.language))
        self.lang_combo.setCurrentText(self.language)
        self.lang_combo.currentTextChanged.connect(self._on_language_changed)
        lang_row.addWidget(lang_label)
        lang_row.addSpacing(5)
        lang_row.addWidget(self.lang_combo)

        wb = QPushButton(t("setup_get_started", self.language))
        wb.setFixedSize(250, 80)
        wb.setStyleSheet("font-weight: 600;")
        wb.clicked.connect(self.go_next)
        wl.addWidget(wt)
        wl.addSpacing(10)
        wl.addWidget(ws)
        wl.addSpacing(30)
        wl.addLayout(lang_row)
        wl.addSpacing(30)
        wl.addWidget(wb, alignment=Qt.AlignCenter)
        self.pages_list.append(welcome)

        email_page = QWidget()
        el = QVBoxLayout(email_page)
        el.setAlignment(Qt.AlignCenter)
        self.email_step = QLabel(f"{t('setup_step', self.language)} 1 / 3")
        self.email_step.setStyleSheet("font: 600 14px; color: #0891b2;")
        self.email_step.setAlignment(Qt.AlignCenter)
        et = QLabel(t("setup_email_title", self.language))
        et.setStyleSheet("font: bold 32px;")
        et.setAlignment(Qt.AlignCenter)
        ed = QLabel(t("setup_email_desc", self.language))
        ed.setStyleSheet("font: normal 17px;")
        ed.setAlignment(Qt.AlignCenter)
        self.email_input = QComboBox()
        self.email_input.setEditable(True)
        self.email_input.setFixedSize(500, 45)
        self.email_input.setFont(QFont(_INTER_FAMILY or "Inter", 12))
        fill_combo(self.email_input, list_outlook_accounts(), config.get("email_address", ""))
        self.email_input.currentTextChanged.connect(self._wizard_email_changed)
        el.addWidget(self.email_step)
        el.addSpacing(10)
        el.addWidget(et)
        el.addSpacing(8)
        el.addWidget(ed)
        el.addSpacing(25)
        el.addWidget(self.email_input, alignment=Qt.AlignCenter)
        self.pages_list.append(email_page)

        folder_page = QWidget()
        fl = QVBoxLayout(folder_page)
        fl.setAlignment(Qt.AlignCenter)
        self.folder_step = QLabel(f"{t('setup_step', self.language)} 2 / 3")
        self.folder_step.setStyleSheet("font: 600 14px; color: #0891b2;")
        self.folder_step.setAlignment(Qt.AlignCenter)
        ft = QLabel(t("setup_folder_title", self.language))
        ft.setStyleSheet("font: bold 32px;")
        ft.setAlignment(Qt.AlignCenter)
        fd = QLabel(t("setup_folder_desc", self.language))
        fd.setStyleSheet("font: normal 17px;")
        fd.setAlignment(Qt.AlignCenter)
        self.folder_input = QComboBox()
        self.folder_input.setEditable(True)
        self.folder_input.setFixedSize(500, 45)
        self.folder_input.setFont(QFont(_INTER_FAMILY or "Inter", 12))
        fill_combo(self.folder_input, list_outlook_folders(config.get("email_address", "")), config.get("folder", "") or "")
        self.folder_input.currentTextChanged.connect(self.update_nav)
        fl.addWidget(self.folder_step)
        fl.addSpacing(10)
        fl.addWidget(ft)
        fl.addSpacing(8)
        fl.addWidget(fd)
        fl.addSpacing(25)
        fl.addWidget(self.folder_input, alignment=Qt.AlignCenter)
        self.pages_list.append(folder_page)

        excel_page = QWidget()
        xl = QVBoxLayout(excel_page)
        xl.setAlignment(Qt.AlignCenter)
        self.excel_step = QLabel(f"{t('setup_step', self.language)} 3 / 3")
        self.excel_step.setStyleSheet("font: 600 14px; color: #0891b2;")
        self.excel_step.setAlignment(Qt.AlignCenter)
        xt = QLabel(t("setup_excel_title", self.language))
        xt.setStyleSheet("font: bold 32px;")
        xt.setAlignment(Qt.AlignCenter)
        xd = QLabel(t("setup_excel_desc", self.language))
        xd.setStyleSheet("font: normal 17px;")
        xd.setAlignment(Qt.AlignCenter)
        excel_row = QHBoxLayout()
        excel_row.setAlignment(Qt.AlignCenter)
        self.excel_input = QLineEdit()
        self.excel_input.setPlaceholderText("e.g. C:/Documents/extraction.xlsx")
        self.excel_input.setFixedSize(400, 45)
        self.excel_input.setStyleSheet("QLineEdit { font-size: 16px; }")
        self.excel_input.setMaxLength(254)
        self.excel_input.setText(config.get("excel", ""))
        self.excel_input.textChanged.connect(self.update_nav)
        browse_btn = QPushButton(t("setup_excel_browse", self.language))
        browse_btn.setFixedSize(100, 45)
        browse_btn.clicked.connect(self.browse_excel)
        browse_btn.setStyleSheet("font: normal 18px;")
        excel_row.addWidget(self.excel_input)
        excel_row.addSpacing(5)
        excel_row.addWidget(browse_btn)
        xl.addWidget(self.excel_step)
        xl.addSpacing(10)
        xl.addWidget(xt)
        xl.addSpacing(8)
        xl.addWidget(xd)
        xl.addSpacing(25)
        xl.addLayout(excel_row)
        self.pages_list.append(excel_page)

        finish = QWidget()
        fnl = QVBoxLayout(finish)
        fnl.setAlignment(Qt.AlignCenter)
        fnt = QLabel(t("setup_finish_title", self.language))
        fnt.setStyleSheet("font: bold 44px;")
        fnt.setAlignment(Qt.AlignCenter)
        fnd = QLabel(t("setup_finish_desc", self.language))
        fnd.setStyleSheet("font: normal 17px;")
        fnd.setAlignment(Qt.AlignCenter)
        fnd.setWordWrap(True)
        fnd.setMaximumWidth(700)
        fnb = QPushButton(t("setup_finish_btn", self.language))
        fnb.setFixedSize(300, 80)
        fnb.setStyleSheet("font-weight: 600;")
        fnb.clicked.connect(self.complete_setup)
        fnl.addWidget(fnt)
        fnl.addSpacing(15)
        fnl.addWidget(fnd)
        fnl.addSpacing(40)
        fnl.addWidget(fnb, alignment=Qt.AlignCenter)
        self.pages_list.append(finish)

        for page in self.pages_list:
            self.stack.addWidget(page)

        nav = QHBoxLayout()
        nav.setContentsMargins(40, 0, 40, 30)
        self.back_btn = QPushButton(t("setup_back", self.language))
        self.back_btn.setFixedSize(120, 50)
        self.back_btn.clicked.connect(self.go_back)
        self.back_btn.setStyleSheet("font: normal 17px;")

        self.dots = []
        dots_layout = QHBoxLayout()
        dots_layout.setAlignment(Qt.AlignCenter)
        dots_layout.setSpacing(10)
        for i in range(len(self.pages_list)):
            dot = QLabel()
            dot.setFixedSize(12, 12)
            dots_layout.addWidget(dot)
            self.dots.append(dot)

        self.next_btn = QPushButton(t("setup_next", self.language))
        self.next_btn.setFixedSize(120, 50)
        self.next_btn.clicked.connect(self.go_next)
        self.next_btn.setStyleSheet("font: normal 17px;")

        nav.addWidget(self.back_btn)
        nav.addStretch()
        nav.addLayout(dots_layout)
        nav.addStretch()
        nav.addWidget(self.next_btn)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.addWidget(self.stack, 1)
        layout.addLayout(nav)

        self.update_nav()

    def _on_language_changed(self, language):
        self.language = language
        config = load_config()
        config["language"] = language
        save_config(config)
        self.language_changed.emit(language)

    def update_nav(self):
        idx = self.stack.currentIndex()
        last = len(self.pages_list) - 1

        self.back_btn.setVisible(0 < idx < last)
        self.next_btn.setVisible(0 < idx < last)

        if idx == 1:
            self.next_btn.setEnabled(bool(self.email_input.currentText().strip()))
        elif idx == 2:
            self.next_btn.setEnabled(bool(self.folder_input.currentText().strip()))
        elif idx == 3:
            self.next_btn.setEnabled(True)

        for i, dot in enumerate(self.dots):
            if i == idx:
                dot.setStyleSheet("background-color: #22d3ee; border-radius: 6px;")
            else:
                dot.setStyleSheet("background-color: #1a3a5c; border-radius: 6px;")

    def go_next(self):
        idx = self.stack.currentIndex()
        if idx == 1:
            save_config(load_config() | {"email_address": self.email_input.currentText().strip()})
        elif idx == 2:
            save_config(load_config() | {"folder": self.folder_input.currentText().strip()})
        elif idx == 3:
            save_config(load_config() | {"excel": self.excel_input.text().strip()})

        if idx < len(self.pages_list) - 1:
            self.stack.setCurrentIndex(idx + 1)
            self.update_nav()

    def go_back(self):
        idx = self.stack.currentIndex()
        if idx > 0:
            self.stack.setCurrentIndex(idx - 1)
            self.update_nav()

    def _wizard_email_changed(self, text):
        # Refresh the folder dropdown to the chosen account's folders, then re-check nav.
        fill_combo(self.folder_input, list_outlook_folders(text), self.folder_input.currentText())
        self.update_nav()

    def browse_excel(self):
        path, _ = QFileDialog.getOpenFileName(self, "Select Excel File", "", "Excel Files (*.xlsx)")
        if path:
            self.excel_input.setText(path)

    def complete_setup(self):
        config = load_config()
        config["setup_complete"] = True
        save_config(config)
        self.finished.emit()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Mail AI")
        self.setMinimumSize(1000, 680)
        self.resize(1320, 860)
        self.setWindowIcon(QIcon(resource_path("icon.png")))


        config = load_config()
        self.email_address = config.get("email_address", "")
        self.folder = config.get("folder", "")
        self.excel = config.get("excel", "")
        self.language = config.get("language", "English")
        self.is_first_run = not config.get("setup_complete", False)

        self.col_widths = [160, 145, 150, 160, 140, 220, 200, 120]

        QApplication.setFont(get_font(self.language))

        self.setup_ui()

    def setup_ui(self):
        self.extracting_running = False
        self.listening_running = False
        config = load_config()
        self.emails_processed = config.get("emails_processed", 0)

        self.main_widget = GridWidget()
        main_layout = QHBoxLayout(self.main_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)
        self.setCentralWidget(self.main_widget)

        self.sidebar = QFrame()
        self.sidebar.setFixedWidth(210)
        self.sidebar.setStyleSheet("""
            QFrame {
                background-color: #0a1628;
                border: none;
                border-right: 1px solid #1a3a5c;
            }
        """)
        self.sidebar_layout = QVBoxLayout(self.sidebar)
        self.sidebar_layout.setContentsMargins(0, 0, 0, 0)
        self.sidebar_layout.setSpacing(0)

        self.logo = QLabel()
        self.logo.setAlignment(Qt.AlignCenter)
        self.logo.setFixedHeight(70)
        self.sidebar_layout.addWidget(self.logo)

        self.pages = GridStack()

        self.page_home = self.create_home_page()
        self.page_filtering = self.create_filtering_page()
        self.page_settings = self.create_settings_page()
        self.page_extract = None
        self.page_main = None
        self.page_listening = None

        self.pages.addWidget(self.page_home)
        self.pages.addWidget(self.page_filtering)
        self.pages.addWidget(self.page_settings)

        self.extract_sidebar_btn = QPushButton(t("extract", self.language))
        self.filtering_sidebar_btn = QPushButton(t("filtering", self.language))
        self.settings_sidebar_btn = QPushButton(t("settings", self.language))

        for btn in [self.extract_sidebar_btn, self.filtering_sidebar_btn, self.settings_sidebar_btn]:
            btn.setFixedHeight(48)
            btn.setStyleSheet("""
                QPushButton {
                    background-color: transparent;
                    color: #7ca4c0;
                    font-family: 'Inter';
                    font-size: 13px;
                    font-weight: 500;
                    border: none;
                    border-left: 2px solid transparent;
                    text-align: left;
                    padding-left: 18px;
                }
                QPushButton:hover {
                    background-color: #0d1f35;
                    color: #22d3ee;
                    border-left: 2px solid #0891b2;
                }
            """)
            self.sidebar_layout.addWidget(btn)

        self._active_sidebar_btn = None
        self.extract_sidebar_btn.clicked.connect(self.on_extract_sidebar_clicked)
        self.filtering_sidebar_btn.clicked.connect(
            lambda: (self._set_active_sidebar(self.filtering_sidebar_btn), self.switch_page(self.page_filtering)))
        self.settings_sidebar_btn.clicked.connect(
            lambda: (self._set_active_sidebar(self.settings_sidebar_btn), self.switch_page(self.page_settings)))

        self.sidebar_layout.addStretch()

        ver = QLabel(f"  mailai.uk                                v{APP_VERSION}")
        self.ver = ver
        ver.setFixedHeight(40)
        ver.setStyleSheet("""
            font-family: 'Inter';
            font-size: 11px;
            color: #5c5d66;
            background-color: transparent;
            border-top: 1px solid #26272b;
            padding-left: 16px;
        """)
        self.sidebar_layout.addWidget(ver)

        self.tray = QSystemTrayIcon(self)
        self.tray.setIcon(QIcon(resource_path("icon.png")))
        self.tray.setToolTip("Mail AI")
        self.tray.show()

        main_layout.addWidget(self.sidebar)
        main_layout.addSpacing(20)
        main_layout.addWidget(self.pages)

        if self.is_first_run:
            self.sidebar.hide()
            self.pages.hide()
            self.setup_wizard = SetupWizard(language=self.language, parent=self.main_widget)
            main_layout.addWidget(self.setup_wizard)
            self.setup_wizard.finished.connect(self.on_setup_complete)
            self.setup_wizard.language_changed.connect(self.on_setup_language_changed)

    def on_setup_language_changed(self, language):
        self.language = language
        QApplication.setFont(get_font(language))
        current_theme = load_config().get("theme", "light")
        self.apply_theme(current_theme)

        self.main_widget.layout().removeWidget(self.setup_wizard)
        self.setup_wizard.deleteLater()
        self.setup_wizard = SetupWizard(language=language, parent=self.main_widget)
        self.main_widget.layout().addWidget(self.setup_wizard)
        self.setup_wizard.finished.connect(self.on_setup_complete)
        self.setup_wizard.language_changed.connect(self.on_setup_language_changed)

    def on_setup_complete(self):
        self.setup_wizard.hide()
        self.main_widget.layout().removeWidget(self.setup_wizard)
        self.setup_wizard.deleteLater()
        self.setup_wizard = None

        config = load_config()
        self.email_address = config.get("email_address", "")
        self.folder = config.get("folder", "")
        self.excel = config.get("excel", "")

        old_home = self.page_home
        old_filtering = self.page_filtering
        self.page_home = self.create_home_page()
        self.page_filtering = self.create_filtering_page()
        self.pages.insertWidget(0, self.page_home)
        self.pages.insertWidget(1, self.page_filtering)
        self.pages.removeWidget(old_home)
        self.pages.removeWidget(old_filtering)
        old_home.deleteLater()
        old_filtering.deleteLater()

        self.sidebar.show()
        self.pages.show()
        self.pages.setCurrentWidget(self.page_home)
        self.is_first_run = False

    def create_home_page(self):
        content = QWidget()
        content_layout = QVBoxLayout(content)
        content_layout.setAlignment(Qt.AlignCenter)

        header = QLabel(t("welcome", self.language))
        header.setStyleSheet("font: bold 68px;")
        header.setAlignment(Qt.AlignCenter)


        caption = QLabel(t("extract_something", self.language))
        caption.setStyleSheet("font: normal 22px;")
        caption.setAlignment(Qt.AlignCenter)

        button_row = QHBoxLayout()
        button_row.setSpacing(30)

        btn_left = QPushButton(t("extract", self.language))
        btn_left.setProperty("variant", "primary")
        btn_left.clicked.connect(self.show_extract_page)

        btn_right = QPushButton(t("listen", self.language))
        btn_right.setProperty("variant", "primary")
        btn_right.clicked.connect(self.show_listening_page)

        btn_left.setFixedSize(250, 80)
        btn_right.setFixedSize(250, 80)

        button_row.addWidget(btn_left)
        button_row.addSpacing(20)
        button_row.addWidget(btn_right)
        button_row.setAlignment(Qt.AlignCenter)

        content_layout.addWidget(header)
        content_layout.addWidget(caption)
        content_layout.addSpacing(40)
        content_layout.addLayout(button_row)
        content_layout.addSpacing(40)
        return content

    def create_filtering_page(self):
        content = QWidget()
        content_layout = QVBoxLayout(content)
        content_layout.setAlignment(Qt.AlignTop | Qt.AlignLeft)

        header = QLabel(t("filtering_settings", self.language))
        header.setStyleSheet("font: bold 40px;")

        caption = QLabel(t("email_caption", self.language))
        caption.setStyleSheet("font: 600 17px;")
        self.email_combo = QComboBox()
        self.email_combo.setEditable(True)
        self.email_combo.setFixedSize(700, 40)
        self.email_combo.setFont(QFont(_INTER_FAMILY or "Inter", 12))
        fill_combo(self.email_combo, list_outlook_accounts(), getattr(self, "email_address", ""))
        self.email_combo.currentTextChanged.connect(self._settings_email_changed)

        caption2 = QLabel(t("folder_caption", self.language))
        caption2.setStyleSheet("font: 600 17px;")
        self.folder_combo = QComboBox()
        self.folder_combo.setEditable(True)
        self.folder_combo.setFixedSize(700, 40)
        self.folder_combo.setFont(QFont(_INTER_FAMILY or "Inter", 12))
        fill_combo(self.folder_combo, list_outlook_folders(getattr(self, "email_address", "")), getattr(self, "folder", ""))
        self.folder_combo.currentTextChanged.connect(self.folder_entered)

        caption3 = QLabel(t("excel_caption", self.language))
        caption3.setStyleSheet("font: 600 17px;")
        self.excel_settings_input = QLineEdit()
        self.excel_settings_input.setMaxLength(254)
        self.excel_settings_input.setPlaceholderText("e.g. extraction.xlsx")
        self.excel_settings_input.setFixedSize(592, 40)
        self.excel_settings_input.setStyleSheet("QLineEdit { font-size: 16px; }")
        self.excel_settings_input.setText(getattr(self, "excel", ""))
        self.excel_settings_input.textEdited.connect(self.excel_entered)
        excel_browse_btn = QPushButton(t("setup_excel_browse", self.language))
        excel_browse_btn.setFixedSize(100, 40)
        excel_browse_btn.setStyleSheet("font-weight: 600;")
        excel_browse_btn.clicked.connect(self.browse_excel_settings)
        excel_row = QHBoxLayout()
        excel_row.setAlignment(Qt.AlignLeft)
        excel_row.addWidget(self.excel_settings_input)
        excel_row.addSpacing(8)
        excel_row.addWidget(excel_browse_btn)

        caption4 = QLabel(t("clear_duplicates_caption", self.language))
        caption4.setStyleSheet("font: 600 17px;")
        self.refresh_btn = QPushButton(t("clear_duplicates_btn", self.language))
        self.refresh_btn.setStyleSheet("font-weight: 600;")
        self.refresh_btn.setFixedSize(250, 80)
        self.refresh_btn.clicked.connect(self.refresh_duplicates)

        content_layout.addWidget(header)
        content_layout.addSpacing(25)
        content_layout.addWidget(caption)
        content_layout.addWidget(self.email_combo)
        content_layout.addSpacing(25)
        content_layout.addWidget(caption2)
        content_layout.addWidget(self.folder_combo)
        content_layout.addSpacing(25)
        content_layout.addWidget(caption3)
        content_layout.addLayout(excel_row)
        content_layout.addSpacing(25)
        content_layout.addWidget(caption4)
        content_layout.addSpacing(5)
        content_layout.addWidget(self.refresh_btn)

        return content

    def create_settings_page(self):
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        scroll_area.setFrameShape(QFrame.NoFrame)

        content = GridWidget()
        content.set_theme(load_config().get("theme", "light"))
        self._settings_content = content
        content_layout = QVBoxLayout(content)
        content_layout.setAlignment(Qt.AlignTop | Qt.AlignLeft)

        header = QLabel(t("settings", self.language))
        header.setStyleSheet("font: bold 40px;")

        theme_label = QLabel(t("theme", self.language))
        theme_label.setStyleSheet("font: 600 17px;")

        if load_config().get("theme", "light") == "dark":
            self.theme_btn = QPushButton(t("switch_light", self.language))
        else:
            self.theme_btn = QPushButton(t("switch_dark", self.language))
            
        self.theme_btn.setFixedSize(250, 80)
        self.theme_btn.setStyleSheet("font: 600 17px;")
        self.theme_btn.clicked.connect(self.toggle_theme)

        language_label = QLabel(t("language", self.language))
        language_label.setStyleSheet("font: 600 17px;")

        self.language_combo = QComboBox()
        self.language_combo.addItems(["English", "中文"])
        self.language_combo.setFixedSize(250, 80)
        self.language_combo.setFont(get_font(self.language))
        self.language_combo.blockSignals(True)
        self.language_combo.setCurrentText(self.language)
        self.language_combo.blockSignals(False)
        self.language_combo.currentTextChanged.connect(self.language_changed)

        content_layout.addWidget(header)
        content_layout.addSpacing(25)
        content_layout.addWidget(theme_label)
        content_layout.addSpacing(5)
        content_layout.addWidget(self.theme_btn)
        content_layout.addSpacing(25)
        content_layout.addWidget(language_label)
        content_layout.addSpacing(5)
        content_layout.addWidget(self.language_combo)
        content_layout.addSpacing(40)

        zones_header = QLabel(t("custom_zones_header", self.language))
        zones_header.setStyleSheet("font: bold 26px;")
        content_layout.addWidget(zones_header)

        zones_desc = QLabel(t("custom_zones_desc", self.language))
        zones_desc.setStyleSheet("font: 16px;")
        zones_desc.setWordWrap(True)
        content_layout.addWidget(zones_desc)
        content_layout.addSpacing(10)

        input_row = QHBoxLayout()

        port_label = QLabel(t("port_name_label", self.language))
        port_label.setStyleSheet("font: 600 16px;")
        self.port_input = QLineEdit()
        self.port_input.setPlaceholderText("e.g. BAHIA BLANCA")
        self.port_input.setFixedSize(300, 40)
        self.port_input.setStyleSheet("QLineEdit { font-size: 16px; }")

        zone_label_input = QLabel(t("zone_label", self.language))
        zone_label_input.setStyleSheet("font: 600 16px;")
        self.zone_input = QLineEdit()
        self.zone_input.setPlaceholderText("e.g. ECSA")
        self.zone_input.setFixedSize(200, 40)
        self.zone_input.setStyleSheet("QLineEdit { font-size: 16px; }")

        self.add_zone_btn = QPushButton(t("add_zone_btn", self.language))
        self.add_zone_btn.setFixedSize(180, 40)
        self.add_zone_btn.setStyleSheet("font-weight: 600;")
        self.add_zone_btn.clicked.connect(self.add_custom_zone_clicked)

        input_row.addWidget(port_label)
        input_row.addWidget(self.port_input)
        input_row.addSpacing(10)
        input_row.addWidget(zone_label_input)
        input_row.addWidget(self.zone_input)
        input_row.addSpacing(10)
        input_row.addWidget(self.add_zone_btn)
        input_row.addStretch()
        content_layout.addLayout(input_row)

        self.zone_status_label = QLabel("")
        self.zone_status_label.setStyleSheet("font: 14px; color: #4CAF50;")
        content_layout.addWidget(self.zone_status_label)
        content_layout.addSpacing(15)

        zones_list_label = QLabel(t("custom_zones_list", self.language))
        zones_list_label.setStyleSheet("font: 600 17px;")
        content_layout.addWidget(zones_list_label)
        content_layout.addSpacing(5)

        self.zones_list_container = QVBoxLayout()
        content_layout.addLayout(self.zones_list_container)
        self.refresh_zones_list()

        content_layout.addSpacing(40)

        license_header = QLabel(t("pro_section_header", self.language))
        license_header.setStyleSheet("font: bold 26px;")
        content_layout.addWidget(license_header)
        content_layout.addSpacing(10)

        cfg = load_config()
        if cfg.get("is_pro", False):
            status_text = t("pro_active", self.language)
            status_color = "#22d3ee"
        elif trial_active():
            status_text = f"{t('trial_active', self.language)} — {trial_days_left()} {t('trial_days_left', self.language)}"
            status_color = "#7ca4c0"
        else:
            status_text = t("trial_expired", self.language)
            status_color = "#f87171"

        self.license_status_label = QLabel(status_text)
        self.license_status_label.setStyleSheet(f"font: 16px; color: {status_color};")
        content_layout.addWidget(self.license_status_label)
        content_layout.addSpacing(10)

        key_row = QHBoxLayout()
        license_label_w = QLabel(t("license_label", self.language))
        license_label_w.setStyleSheet("font: 600 16px;")

        self.license_input = QLineEdit()
        self.license_input.setPlaceholderText("MAILAI-YYYYMM-XXXXXXXXXX")
        self.license_input.setFixedSize(340, 40)
        self.license_input.setStyleSheet("QLineEdit { font-size: 15px; }")
        self.license_input.setText(cfg.get("license_key", "") if cfg.get("is_pro") else "")

        self.activate_btn = QPushButton(t("activate_btn", self.language))
        self.activate_btn.setFixedSize(140, 40)
        self.activate_btn.setProperty("variant", "primary")
        self.activate_btn.clicked.connect(self.activate_license)

        key_row.addWidget(license_label_w)
        key_row.addSpacing(8)
        key_row.addWidget(self.license_input)
        key_row.addSpacing(8)
        key_row.addWidget(self.activate_btn)
        key_row.addStretch()
        content_layout.addLayout(key_row)

        self.key_feedback_label = QLabel("")
        self.key_feedback_label.setStyleSheet("font: 14px;")
        content_layout.addWidget(self.key_feedback_label)
        content_layout.addSpacing(12)

        upgrade_btn = QPushButton(t("upgrade_btn", self.language))
        upgrade_btn.setFixedSize(340, 50)
        upgrade_btn.setProperty("variant", "primary")
        upgrade_btn.clicked.connect(lambda: QDesktopServices.openUrl(QUrl("https://ko-fi.com/mailaiuk/tiers")))
        content_layout.addWidget(upgrade_btn)
        content_layout.addSpacing(20)

        scroll_area.setWidget(content)
        return scroll_area

    def refresh_zones_list(self):
        while self.zones_list_container.count():
            item = self.zones_list_container.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
            elif item.layout():
                while item.layout().count():
                    child = item.layout().takeAt(0)
                    if child.widget():
                        child.widget().deleteLater()

        custom_zones = get_custom_zones_list()
        if not custom_zones:
            empty_label = QLabel(t("no_custom_zones", self.language))
            empty_label.setStyleSheet("font: 14px; color: gray;")
            self.zones_list_container.addWidget(empty_label)
            return

        for port, zones in custom_zones:
            row = QHBoxLayout()
            label = QLabel(f"{port}  ->  {zones}")
            label.setStyleSheet("font: 15px;")
            label.setMinimumWidth(500)

            remove_btn = QPushButton(t("remove_zone_btn", self.language))
            remove_btn.setFixedSize(120, 32)
            remove_btn.setStyleSheet("font-weight: 600;")
            remove_btn.clicked.connect(lambda checked, p=port: self.remove_custom_zone_clicked(p))

            row.addWidget(label)
            row.addWidget(remove_btn)
            row.addStretch()

            row_widget = QWidget()
            row_widget.setLayout(row)
            self.zones_list_container.addWidget(row_widget)

    def add_custom_zone_clicked(self):
        global csv_dict
        port = self.port_input.text().strip()
        zone = self.zone_input.text().strip()
        if not port or not zone:
            self.zone_status_label.setStyleSheet("font: 14px; color: #f44336;")
            self.zone_status_label.setText(t("zone_empty", self.language))
            return

        add_custom_zone(port, zone)
        csv_dict = merge_custom_zones(load_csv_into_dict(csv_file))
        self.port_input.clear()
        self.zone_input.clear()
        self.zone_status_label.setStyleSheet("font: 14px; color: #4CAF50;")
        self.zone_status_label.setText(t("zone_added", self.language))
        self.refresh_zones_list()

    def remove_custom_zone_clicked(self, port_name):
        global csv_dict
        remove_custom_zone(port_name)
        csv_dict = merge_custom_zones(load_csv_into_dict(csv_file))
        self.zone_status_label.setStyleSheet("font: 14px; color: #4CAF50;")
        self.zone_status_label.setText(t("zone_removed", self.language))
        self.refresh_zones_list()

    def create_extract_page(self):
        content = QWidget()
        content_layout = QVBoxLayout(content)
        content_layout.setAlignment(Qt.AlignTop | Qt.AlignLeft)

        header = QLabel(t("extract_page_header", self.language))
        header.setStyleSheet("font: bold 40px;")

        self.caption = QLabel("")
        self.caption2 = QLabel("")
        self.caption3 = QLabel("")
        self.caption4 = QLabel("")
        self.captione1 = QLabel("")
        self.captione2 = QLabel("")

        input_row = QHBoxLayout()
        input_row.setAlignment(Qt.AlignLeft)

        caption5 = QLabel(t("date_caption", self.language))
        caption5.setStyleSheet("font: 600 17px;")

        _cur_year = datetime.now().year

        self.input_day = self._dt_combo("Day", [(str(d), f"{d:02d}") for d in range(1, 32)], 110, editable=True)
        self.input_day.currentTextChanged.connect(self._date_combo_changed)
        self.input_month = self._dt_combo("Month", [(MONTH_NAMES[m - 1], f"{m:02d}") for m in range(1, 13)], 150, editable=True)
        self.input_month.currentTextChanged.connect(self._date_combo_changed)
        self.input_year = self._dt_combo("Year", [(str(y), str(y)) for y in range(_cur_year, _cur_year - 6, -1)], 110, editable=True)
        self.input_year.currentTextChanged.connect(self._date_combo_changed)

        input_row.addWidget(self.input_day)
        input_row.addWidget(self.input_month)
        input_row.addWidget(self.input_year)

        input_row2 = QHBoxLayout()
        input_row2.setAlignment(Qt.AlignLeft)

        caption6 = QLabel(t("time_caption", self.language))
        caption6.setWordWrap(True)
        caption6.setMaximumWidth(1000)
        caption6.setStyleSheet("font: 600 17px;")

        # 24-hour time (no AM/PM). Editable so a custom time can be typed; the minute
        # dropdown suggests 15-minute steps but any 0-59 value can be entered.
        self.input_hour = self._dt_combo("Hour", [(f"{h:02d}", f"{h:02d}") for h in range(0, 24)], 120, editable=True)
        self.input_hour.currentTextChanged.connect(self._time_combo_changed)
        self.input_minute = self._dt_combo("Min", [(f"{m:02d}", f"{m:02d}") for m in (0, 15, 30, 45)], 120, editable=True)
        self.input_minute.currentTextChanged.connect(self._time_combo_changed)
        time_colon = QLabel(":")
        time_colon.setStyleSheet("font: 700 18px;")

        input_row2.addWidget(self.input_hour)
        input_row2.addWidget(time_colon)
        input_row2.addWidget(self.input_minute)

        self.btn = QPushButton(t("start_extracting", self.language))
        self.btn.setProperty("variant", "primary")
        self.btn.setFixedSize(200, 80)
        self.btn.clicked.connect(self.handle_extract)
        self.btn.setEnabled(False)

        self.error = QLabel("")
        self.error.setStyleSheet("font: 600 16px; color: red;")

        content_layout.addWidget(header)
        content_layout.addSpacing(25)
        content_layout.addWidget(self.caption)
        content_layout.addWidget(self.caption2)
        content_layout.addSpacing(25)
        content_layout.addWidget(self.caption3)
        content_layout.addWidget(self.caption4)
        content_layout.addSpacing(25)
        content_layout.addWidget(self.captione1)
        content_layout.addWidget(self.captione2)
        content_layout.addSpacing(25)
        content_layout.addWidget(caption5)
        content_layout.addLayout(input_row)
        content_layout.addSpacing(25)
        content_layout.addWidget(caption6)
        content_layout.addLayout(input_row2)
        content_layout.addSpacing(40)
        content_layout.addWidget(self.btn)
        content_layout.addWidget(self.error)

        return content

    def _header_style(self):
        c = getattr(self, "_theme_colors", {})
        muted = c.get("muted", "#9b9ba3"); border = c.get("border", "#26272b")
        return (f"font-family:'Inter'; font-size:13px; font-weight:600; "
                f"letter-spacing:1px; color:{muted}; padding:10px 12px; "
                f"border-bottom:1px solid {border};")

    def _header_btn_qss(self, active):
        c = getattr(self, "_theme_colors", {})
        muted = c.get("muted", "#9b9ba3"); text = c.get("text", "#f4f4f5")
        border = c.get("border", "#26272b"); accent = c.get("accent", "#22d3ee")
        col = accent if active else muted
        return (f"QPushButton {{ background:transparent; border:none; border-radius:0; "
                f"border-bottom:1px solid {border}; color:{col}; font-family:'Inter'; "
                f"font-size:13px; font-weight:600; letter-spacing:1px; padding:10px 12px; "
                f"text-align:left; }}"
                f"QPushButton:hover {{ color:{text}; }}")

    def _cell_style(self, col, row_index):
        c = getattr(self, "_theme_colors", {})
        text = c.get("text", "#f4f4f5"); accent = c.get("accent", "#22d3ee")
        surface = c.get("surface", "#141517"); border = c.get("border", "#26272b")
        row_bg = surface if (row_index % 2 == 1) else "transparent"
        base = f"padding:10px 12px; background:{row_bg}; border-bottom:1px solid {border};"
        if col == 4:
            return f"font-family:'Inter'; font-size:15px; font-weight:600; color:{accent}; {base}"
        if col == 1:
            return f"font-family:'Inter'; font-size:15px; color:{text}; {base}"
        return f"font-size:16px; color:{text}; {base}"

    def _status_box_qss(self, accent_hex):
        c = getattr(self, "_theme_colors", {})
        surface = c.get("surface", "#141517"); border = c.get("border", "#26272b")
        return (f"background-color:{surface}; border:1px solid {border}; "
                f"border-left:3px solid {accent_hex}; border-radius:8px; margin-left:-10px;")

    def _style_extbox(self, color):
        """Style the main status banner and remember its colour, so a theme toggle can
        re-apply it with the new palette (its bg/border are baked from the theme)."""
        self._main_status_color = color
        try:
            self.extbox.setStyleSheet(self._status_box_qss(color))
        except RuntimeError:
            pass

    def create_main_page(self):
        content = QWidget()
        content_layout = QVBoxLayout(content)
        content_layout.setContentsMargins(24, 24, 24, 24)
        content_layout.setSpacing(8)

        self.extheader = QLabel(t("current_extraction", self.language))
        self.extheader.setStyleSheet("font: bold 40px;")

        self.extbox = QFrame()
        self.extbox.setWindowFlags(Qt.FramelessWindowHint)
        self.extbox.setAttribute(Qt.WA_TranslucentBackground)
        self.extbox.setFixedSize(600, 70)
        self._style_extbox("#34d399")

        self.status = QLabel(t("extraction_running", self.language))
        self.status.setStyleSheet("font: 600 17px; padding: 15px;")

        box_layout = QVBoxLayout(self.extbox)
        box_layout.addWidget(self.status)

        self.stop_btn = QPushButton(t("stop_extracting", self.language))
        self.stop_btn.setFixedSize(200, 80)
        self.stop_btn.clicked.connect(self.handle_stop)
        self.stop_btn.setEnabled(True)

        self.new_extract_btn = QPushButton(t("new_extraction", self.language))
        self.new_extract_btn.setFixedSize(250, 80)
        self.new_extract_btn.setStyleSheet("font-weight: 600;")
        self.new_extract_btn.clicked.connect(self.new_extraction)
        self.new_extract_btn.hide()

        self.caption5 = QLabel("")
        self.caption5.setStyleSheet("font: 600 17px;")

        self.table_data = []
        self._sort_col = None
        self._sort_asc = True
        self._seq_counter = 0
        self._populating = False
        self._visible_rows = []
        self._search_pos = -1
        self._search_last_q = ""

        self._data_offset = 2
        self._header_texts = [
            "MV", "DWT/Built",
            t("location", self.language),
            t("open_date", self.language),
            t("zone", self.language),
            t("sender", self.language),
            t("subject", self.language),
            t("date", self.language),
        ]
        self.table = QTableWidget(0, len(self._header_texts) + self._data_offset)
        self.table.setHorizontalHeaderLabels(["", ""] + [h.upper() for h in self._header_texts])
        self.table.verticalHeader().setVisible(False)
        self.table.setShowGrid(False)
        self.table.setWordWrap(True)
        self.table.setAlternatingRowColors(True)
        self.table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.table.setMinimumSize(560, 280)
        self.table.setSelectionBehavior(QAbstractItemView.SelectItems)
        self.table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.table.setEditTriggers(QAbstractItemView.EditKeyPressed)
        self.table.setStyleSheet(self._table_qss())
        self.table.setColumnWidth(0, 38)
        self.table.setColumnWidth(1, 30)
        for i, w in enumerate(self.col_widths):
            self.table.setColumnWidth(i + self._data_offset, w)
        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.Interactive)
        header.setSortIndicatorShown(False)
        header.setSectionsClickable(True)
        header.setDefaultAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        header.sectionClicked.connect(self._on_header_clicked)
        self.table.cellClicked.connect(self._on_cell_clicked)
        self.table.itemChanged.connect(self._on_item_changed)
        if not getattr(self, "_table_filter_installed", False):
            QApplication.instance().installEventFilter(self)
            self._table_filter_installed = True

        self.continue_listen_btn = QPushButton(t("continue_listen", self.language))
        self.continue_listen_btn.setFixedSize(250, 80)
        self.continue_listen_btn.setStyleSheet("font-weight: 600;")
        self.continue_listen_btn.clicked.connect(self.toggle_main_listening)
        self.continue_listen_btn.hide()

        self.open_excel_btn = QPushButton(t("open_excel_btn", self.language))
        self.open_excel_btn.setFixedSize(250, 80)
        self.open_excel_btn.setStyleSheet("font-weight: 600;")
        self.open_excel_btn.clicked.connect(self.open_excel_file)
        self.open_excel_btn.hide()

        tc = getattr(self, "_theme_colors", {})
        _accent = tc.get("accent", "#22d3ee"); _muted = tc.get("muted", "#9b9ba3")
        self.table_help = QLabel(
            f"<div style='line-height:140%; color:{_muted}; font-size:12px;'>"
            f"<b style='letter-spacing:1px;'>HOW TO USE</b><br>"
            f"<span style='color:{_accent};'>&#9679;</span> Received via live listening"
            f" &nbsp;&nbsp; <span style='color:{_accent};'>&#9733;</span> Click to star (pins to top)<br>"
            f"Tap any cell to edit &nbsp;·&nbsp; Click a column header to sort"
            f"</div>"
        )
        self.table_help.setTextFormat(Qt.RichText)
        self.table_help.setAlignment(Qt.AlignRight | Qt.AlignVCenter)

        btn_row = QHBoxLayout()
        btn_row.addWidget(self.stop_btn)
        btn_row.addWidget(self.new_extract_btn)
        btn_row.addWidget(self.continue_listen_btn)
        btn_row.addWidget(self.open_excel_btn)
        btn_row.addStretch()
        btn_row.addWidget(self.table_help, alignment=Qt.AlignRight | Qt.AlignVCenter)

        content_layout.addWidget(self.extheader, alignment=Qt.AlignLeft)
        content_layout.addSpacing(5)
        content_layout.addWidget(self.extbox, alignment=Qt.AlignLeft)
        content_layout.addSpacing(5)
        content_layout.addLayout(btn_row)

        self.table_search = QLineEdit()
        self.table_search.setPlaceholderText("Search the table…  (Enter = next match)")
        self.table_search.setClearButtonEnabled(True)
        self.table_search.setFixedSize(340, 36)
        self.table_search.setFont(QFont(_INTER_FAMILY or "Inter", 11))
        self.table_search.textChanged.connect(lambda: self._run_search(advance=False))
        self.table_search.returnPressed.connect(lambda: self._run_search(advance=True))
        search_row = QHBoxLayout()
        search_row.addWidget(self.caption5, alignment=Qt.AlignLeft | Qt.AlignVCenter)
        search_row.addStretch()
        search_row.addWidget(self.table_search, alignment=Qt.AlignRight | Qt.AlignVCenter)

        content_layout.addLayout(search_row)
        content_layout.addWidget(self.table, 1)

        return content

    def create_listening_page(self):
        content = QWidget()
        content_layout = QVBoxLayout(content)
        content_layout.setContentsMargins(24, 24, 24, 24)
        content_layout.setSpacing(8)

        self.lheader = QLabel(t("listening_header", self.language))
        self.lheader.setStyleSheet("font: bold 40px;")

        self.lbox = QFrame()
        self.lbox.setWindowFlags(Qt.FramelessWindowHint)
        self.lbox.setAttribute(Qt.WA_TranslucentBackground)
        self.lbox.setFixedSize(600, 70)
        self.lbox.setStyleSheet(self._status_box_qss("#f59e0b"))

        self.statusl = QLabel(t("listening_paused", self.language))
        self.statusl.setStyleSheet("font: 600 17px; padding: 15px;")

        lbox_layout = QVBoxLayout(self.lbox)
        lbox_layout.addWidget(self.statusl)

        self.listen_toggle_btn = QPushButton(t("resume_listen", self.language))
        self.listen_toggle_btn.clicked.connect(self.toggle_listening)
        self.listen_toggle_btn.setFixedSize(250, 80)

        self.lcount = QLabel("")
        self.lcount.setStyleSheet("font: 600 17px;")

        self.lscrollf = QScrollArea()
        self.lscrollf.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.lscrollf.setMinimumSize(560, 280)
        self.lscrollf.setWidgetResizable(True)
        self.lscrollf.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        self.lscrollf.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOn)

        self.lcontainer = QWidget()
        self.lcontainer.setMinimumWidth(1510)
        self.lrow = 1

        self.lgrid = QGridLayout(self.lcontainer)
        self.lgrid.setContentsMargins(10, 10, 10, 10)
        self.lgrid.setHorizontalSpacing(30)
        self.lgrid.setVerticalSpacing(15)
        self.lgrid.setRowStretch(0, 0)
        self.lgrid.setAlignment(Qt.AlignTop)

        headers = [
            "MV", "DWT/Built",
            t("location", self.language), t("open_date", self.language),
            t("zone", self.language), t("sender", self.language),
            t("subject", self.language), t("date", self.language),
        ]

        for i, text in enumerate(headers):
            h = QLabel(text.upper())
            h.setStyleSheet(self._header_style())
            h.setFixedWidth(self.col_widths[i])
            self.lgrid.addWidget(h, 0, i)

        self.lscrollf.setWidget(self.lcontainer)

        content_layout.addWidget(self.lheader, alignment=Qt.AlignLeft)
        content_layout.addSpacing(5)
        content_layout.addWidget(self.lbox, alignment=Qt.AlignLeft)
        content_layout.addSpacing(5)
        content_layout.addWidget(self.listen_toggle_btn, alignment=Qt.AlignLeft)
        content_layout.addWidget(self.lcount, alignment=Qt.AlignLeft)
        content_layout.addWidget(self.lscrollf, 1)
        return content


    def _set_active_sidebar(self, btn):
        self._active_sidebar_btn = btn
        base = getattr(self, "_sidebar_btn_qss", "")
        active = getattr(self, "_sidebar_btn_active_qss", "")
        for b in (self.extract_sidebar_btn, self.filtering_sidebar_btn, self.settings_sidebar_btn):
            b.setStyleSheet(active if b is btn else base)

    def on_extract_sidebar_clicked(self):
        self._set_active_sidebar(self.extract_sidebar_btn)
        if self.page_main is not None:
            self.switch_page(self.page_main)
        else:
            self.switch_page(self.page_home)

    def new_extraction(self):
        if getattr(self, "listen_worker", None):
            for sig, slot in ((self.listen_worker.new_email, self.add_listening_to_main_table),
                              (self.listen_worker.done, self.on_main_listen_done)):
                try:
                    sig.disconnect(slot)
                except (RuntimeError, TypeError):
                    pass
            if self.listening_running:
                self.listen_worker.stop()
        self.listening_running = False
        if self.page_main is not None:
            self.pages.removeWidget(self.page_main)
            self.page_main.deleteLater()
            self.page_main = None
        self.page_extract = None
        self.show_extract_page()

    def show_extract_page(self):
        if self.extracting_running:
            self.show_main_page()
            return

        if self.page_extract is None:
            self.page_extract = self.create_extract_page()
            self.pages.addWidget(self.page_extract)
        self.switch_page(self.page_extract)

        if getattr(self, "email_address", ""):
            self.caption.setText(t("email_extracting", self.language))
            self.caption.setStyleSheet("font: 600 17px;")
            self.caption2.setText(self.email_address)
            self.caption2.setStyleSheet("font: normal 17px;")
        else:
            self.caption.setText(t("no_email", self.language))
            self.caption.setStyleSheet("font: 600 17px;")
            self.caption2.setText("")

        if getattr(self, "folder", ""):
            self.caption3.setText(t("folder_extracting", self.language))
            self.caption3.setStyleSheet("font: 600 17px;")
         
            self.caption4.setText(self.folder)
            self.caption4.setStyleSheet("font: normal 17px;")
        else:
            self.caption3.setText(t("no_folder", self.language))
            self.caption3.setStyleSheet("font: 600 17px;")
            self.caption4.setText("")

        if getattr(self, "excel", ""):
            self.captione1.setText(t("excel_extracting", self.language))
            self.captione1.setStyleSheet("font: 600 17px;")
            self.captione2.setText(self.excel)
            self.captione2.setStyleSheet("font: normal 17px;")
        else:
            self.captione1.setText(t("excel_extracting", self.language))
            self.captione1.setStyleSheet("font: 600 17px;")
            self.captione2.setText(resolve_excel_path(""))
            self.captione2.setStyleSheet("font: normal 17px; color: grey;")

        if getattr(self, "email_address", "") and getattr(self, "folder", ""):
            self.btn.setEnabled(True)
        else:
            self.btn.setEnabled(False)

        self.date = None
        self.time = None
        

    def show_main_page(self):
        self.page_main = self.create_main_page()
        self.pages.addWidget(self.page_main)
        self.switch_page(self.page_main)
    
    def show_listening_page(self):
        if self.page_listening is None:
            self.page_listening = self.create_listening_page()
            self.pages.addWidget(self.page_listening)
        self.switch_page(self.page_listening)
        if not self.listening_running:
            self.toggle_listening()


    def switch_page(self, page):
        self.pages.setCurrentWidget(page)

    def email_entered(self, text):
        self.email_address = text
        save_config(load_config() | {"email_address": text})

    def folder_entered(self, text):
        self.folder = text
        save_config(load_config() | {"folder": text})

    def excel_entered(self, text):
        self.excel = text
        save_config(load_config() | {"excel": text})

    def _settings_email_changed(self, text):
        # Save the chosen account and refresh the folder dropdown to that account's folders.
        self.email_entered(text)
        fill_combo(self.folder_combo, list_outlook_folders(text), self.folder_combo.currentText())

    def browse_excel_settings(self):
        path, _ = QFileDialog.getOpenFileName(self, "Select Excel File", "", "Excel Files (*.xlsx)")
        if path:
            self.excel_settings_input.setText(path)
            self.excel_entered(path)
    
    def refresh_duplicates(self):
        delete_duplicates()
        self.refresh_btn.setText(t("cleared", self.language))
        QTimer.singleShot(2000, lambda: self.refresh_btn.setText(t("clear_duplicates_btn", self.language)))

    def _dt_combo(self, placeholder, values, width=90, editable=False):
        """Build a date/time dropdown. Non-editable combos get a placeholder item (data "");
        editable ones (the time fields) start empty with a placeholder and accept typed values."""
        cb = QComboBox()
        cb.setFixedSize(width, 40)
        cb.setFont(QFont(_INTER_FAMILY or "Inter", 12))
        for disp, val in values:
            cb.addItem(disp, val)
        cb.setCurrentIndex(-1)  # start unselected so the placeholder shows (not a list item)
        if editable:
            cb.setEditable(True)
            cb.setInsertPolicy(QComboBox.NoInsert)
            cb.lineEdit().setPlaceholderText(placeholder)
        else:
            cb.setPlaceholderText(placeholder)
        return cb

    @staticmethod
    def _norm_num(text, lo, hi):
        """Normalise typed/selected time text to a zero-padded 2-digit string in range, else ''."""
        text = (text or "").strip()
        if text.isdigit() and lo <= int(text) <= hi:
            return f"{int(text):02d}"
        return ""

    @staticmethod
    def _norm_year(text):
        text = (text or "").strip()
        return text if (text.isdigit() and len(text) == 4) else ""

    def _norm_month(self, text):
        text = (text or "").strip()
        if text.isdigit() and 1 <= int(text) <= 12:
            return f"{int(text):02d}"
        tl = text.lower()
        for i, name in enumerate(MONTH_NAMES, 1):
            if tl and name.lower().startswith(tl):
                return f"{i:02d}"
        return ""

    def _date_combo_changed(self, *_):
        # Editable: read the (possibly typed) text and normalise. Month accepts a name or number.
        self.day = self._norm_num(self.input_day.currentText(), 1, 31)
        self.month = self._norm_month(self.input_month.currentText())
        self.year = self._norm_year(self.input_year.currentText())
        self.date = f"{self.year}-{self.month}-{self.day}" if (self.day and self.month and self.year) else None

    def _time_combo_changed(self, *_):
        # Editable: read the (possibly typed) text and normalise. 24-hour, no AM/PM.
        self.hours = self._norm_num(self.input_hour.currentText(), 0, 23)
        self.minutes = self._norm_num(self.input_minute.currentText(), 0, 59)
        self.time = f"{self.hours}:{self.minutes}" if (self.hours and self.minutes) else None

    def handle_extract(self):
        self.btn.setEnabled(False)

        v, msg, dt = validate(self.date, self.time, self.email_address, self.folder, self.excel, outlook, self.language)
        if not v:
            self.error.setText(msg)
            self.btn.setEnabled(True)
            return

        if not access_allowed():
            self.show_upgrade_dialog()
            self.btn.setEnabled(True)
            return

        self.extracting_running = True
        self._current_excel = resolve_excel_path(self.excel)
        # Anchor for auto-listening: capture the moment extraction begins so listening later
        # picks up everything from here on, with no gap while the extraction runs.
        self._extraction_started = datetime.now()
        self.show_main_page()

        self.thread = QThread()
        self.worker = ExtractWorker(None)
        generator = night_extraction(dt, self.email_address, self.folder, self._current_excel, csv_dict, self.worker)
        self.worker.generator = generator
        self.worker.moveToThread(self.thread)

        self.worker.new_email.connect(self.add_email_to_table)
        self.thread.started.connect(self.worker.run)
        self.worker.done.connect(self.on_extraction_done)
        self.worker.done.connect(self.thread.quit)
        self.worker.done.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)

        self.thread.start()

    def handle_stop(self):
        if hasattr(self, "worker") and self.worker:
            self.worker.stop()
        self.stop_btn.setEnabled(False)

    def handle_listen(self):
        
        v, _, _ = validate(None, None, self.email_address, self.folder, self.excel, outlook, self.language)
        if not v:
            self.listen_toggle_btn.setEnabled(False)
            self.lheader.setText(t("listen_error", self.language))
            self.lheader.setStyleSheet("font: bold 25px; color: red;")
            return

        self.listening_running = True
        self._current_excel = resolve_excel_path(self.excel)

        self.listen_thread = QThread()
        self.listen_worker = ExtractWorker(None)
        generator = process_email(self.email_address, self.folder, self._current_excel, csv_dict, self.listen_worker)
        self.listen_worker.generator = generator
        self.listen_worker.moveToThread(self.listen_thread)

        self.listen_worker.new_email.connect(self.add_to_listening_table)
        self.listen_thread.started.connect(self.listen_worker.run)
        self.listen_worker.done.connect(self.on_listen_done)
        self.listen_worker.done.connect(self.listen_thread.quit)
        self.listen_worker.done.connect(self.listen_worker.deleteLater)
        self.listen_thread.finished.connect(self.listen_thread.deleteLater)
        self.listen_thread.finished.connect(lambda: setattr(self, 'listen_thread', None))

        self.listen_thread.start()
        self.listen_toggle_btn.setEnabled(True)
    
    
    @Slot(dict)
    
    def add_email_to_table(self, email_data):

        if email_data.get("type") == "excel_locked":
            self.status.setText("Waiting for Excel file to close...")
            self._style_extbox("#f59e0b")
            return
        if email_data.get("type") == "excel_unlocked":
            self.status.setText(t("extraction_running", self.language))
            self._style_extbox("#34d399")
            return

        try:
            def truncate(text, length=50):
                return text if len(text) <= length else text[:length] + "..."

            sender = email_data["sender"]
            subject = email_data["subject"]
            received_time = email_data["received_time"][:10]
            ves = email_data["ves"]
            vessel_data = email_data["vessel_data"]

            mv = vessel_data.get("MV", "")
            dwt = vessel_data.get("Deadweight", "") or ""
            built = vessel_data.get("Build Year", "") or ""
            dwt_built = f"{dwt}/{built}" if dwt and built else (dwt or built or "")
            location = vessel_data.get("Vessel Open Location", "")
            date = vessel_data.get("Vessel Open Date", "")
            zone = vessel_data.get("Zone", "")

            self.caption5.setText(f"{t('vessels_extracted', self.language)} {ves}")

            self._seq_counter += 1
            self.table_data.append({
                'zone': zone or "",
                'labels': [mv, dwt_built, location, date, zone or "", sender, truncate(subject), received_time],
                'listened': False,
                'starred': False,
                'seq': self._seq_counter,
            })
            self._render_main_table()
            self.table.scrollToBottom()
            self._flash_new_row(self._seq_counter)

            self.emails_processed += 1

        except Exception as e:
            print(f"Error adding email to table: {e}")

    def add_to_listening_table(self, email_data):

        if email_data.get("type") == "excel_locked":
            self.statusl.setText("Waiting for Excel file to close...")
            self.lbox.setStyleSheet(self._status_box_qss("#f59e0b"))
            return
        if email_data.get("type") == "excel_unlocked":
            self.statusl.setText(t("listening_running", self.language))
            self.lbox.setStyleSheet(self._status_box_qss("#34d399"))
            return

        try:
            def truncate(text, length=50):
                        return text if len(text) <= length else text[:length] + "..."

            sender = email_data["sender"]
            subject = email_data["subject"]
            received_time = email_data["received_time"][:10]
            ves = email_data["ves"]
            vessel_data = email_data["vessel_data"]

            mv = vessel_data.get("MV", "")
            dwt = vessel_data.get("Deadweight", "") or ""
            built = vessel_data.get("Build Year", "") or ""
            dwt_built = f"{dwt}/{built}" if dwt and built else (dwt or built or "")
            location = vessel_data.get("Vessel Open Location", "")
            date = vessel_data.get("Vessel Open Date", "")
            zone = vessel_data.get("Zone", "")

            self.lcount.setText(f"{t('vessels_extracted', self.language)} {ves}")

            labels = [
                QLabel(mv), QLabel(dwt_built), QLabel(location), QLabel(date),
                QLabel(zone), QLabel(sender), QLabel(truncate(subject)), QLabel(received_time)
            ]

            for i, label in enumerate(labels):
                label.setStyleSheet(self._cell_style(i, self.lrow))
                label.setWordWrap(True)
                label.setFixedWidth(self.col_widths[i])
                self.lgrid.addWidget(label, self.lrow, i)

            self.lrow += 1
            self.lscrollf.verticalScrollBar().setValue(self.lscrollf.verticalScrollBar().maximum())

            self.emails_processed += 1

        except Exception as e:
            print(f"Error adding email to table: {e}")

    def _truncate(self, text, length=50):
        return text if len(text) <= length else text[:length] + "..."

    def _row_from_email(self, email_data):
        """Build (zone, labels[]) in the table's column order from a worker email dict."""
        v = email_data["vessel_data"]
        mv = v.get("MV", "")
        dwt = v.get("Deadweight", "") or ""
        built = v.get("Build Year", "") or ""
        dwt_built = f"{dwt}/{built}" if dwt and built else (dwt or built or "")
        location = v.get("Vessel Open Location", "")
        date = v.get("Vessel Open Date", "")
        zone = v.get("Zone", "") or ""
        sender = email_data["sender"]
        subject = self._truncate(email_data["subject"])
        received = email_data["received_time"][:10]
        return zone, [mv, dwt_built, location, date, zone, sender, subject, received]

    def _table_qss(self):
        c = getattr(self, "_theme_colors", {})
        text = c.get("text", "#f4f4f5"); muted = c.get("muted", "#9b9ba3")
        border = c.get("border", "#26272b"); surface = c.get("surface", "#141517")
        surface2 = c.get("surface2", "#1a1b1e"); accent = c.get("accent", "#22d3ee")
        bg = c.get("bg", "#0a0b0d")
        return (
            f"QTableWidget {{ background:transparent; border:none; color:{text}; font-size:16px; "
            f"gridline-color:{border}; alternate-background-color:{surface}; outline:0; "
            f"selection-background-color:{surface2}; selection-color:{text}; }}"
            f"QTableWidget::item {{ padding:8px 10px; border-bottom:1px solid {border}; }}"
            f"QTableWidget::item:selected {{ background:{surface2}; color:{text}; }}"
            f"QTableWidget QLineEdit {{ background:{bg}; color:{text}; border:1px solid {accent}; "
            f"border-radius:3px; padding:7px 9px; selection-background-color:{accent}; "
            f"selection-color:{bg}; }}"
            f"QHeaderView {{ background:transparent; }}"
            f"QHeaderView::section {{ background:transparent; color:{muted}; border:none; "
            f"border-bottom:1px solid {border}; padding:10px 10px; font-family:'Inter'; "
            f"font-size:13px; font-weight:600; }}"
            f"QHeaderView::section:hover {{ color:{text}; }}"
            f"QTableCornerButton::section {{ background:transparent; border:none; }}"
        )

    def _dwt_num(self, row):
        """Leading integer of the DWT/Built value, for numeric sorting (9K before 10K)."""
        m = re.match(r'(\d+)', (row['labels'][1] or ''))
        return int(m.group(1)) if m else float('inf')

    def _ordered_rows(self):
        """Display order: starred rows first, then live-listened rows (newest first) pinned to
        the top, then everything else. The base block keeps extraction order (the order rows
        arrived) until the user clicks a column header to sort — no automatic zone sorting.
        Sorting a header clears the current listened flags (folding those rows in); a fresh
        listened arrival re-pins with its dot until the next sort."""
        rows = list(self.table_data)
        starred = [r for r in rows if r.get('starred')]
        listened = [r for r in rows if r.get('listened') and not r.get('starred')]
        base = [r for r in rows if not r.get('listened') and not r.get('starred')]
        starred.sort(key=lambda r: r.get('seq', 0))
        listened.sort(key=lambda r: r.get('seq', 0), reverse=True)
        if self._sort_col is not None:
            col = self._sort_col
            keyf = self._dwt_num if col == 1 else (lambda r: (r['labels'][col] or '').upper())
            base.sort(key=keyf, reverse=not self._sort_asc)
        else:
            base.sort(key=lambda r: r.get('seq', 0))   # extraction order; await header click
        return starred + listened + base

    def _render_main_table(self):
        """Rebuild the QTableWidget from table_data in the current display order.
        Col 0 = user star, col 1 = live-listened dot, col 2.. = the editable data columns."""
        if getattr(self, "table", None) is None:
            return
        c = getattr(self, "_theme_colors", {})
        accent = c.get("accent", "#22d3ee"); muted = c.get("muted", "#9b9ba3")
        off = self._data_offset
        self._populating = True
        try:
            ordered = self._ordered_rows()
            self._visible_rows = ordered
            self.table.setRowCount(len(ordered))
            for r, rowd in enumerate(ordered):
                starred = rowd.get('starred', False)
                listened = rowd.get('listened', False)

                star = QTableWidgetItem("★" if starred else "☆")
                star.setTextAlignment(Qt.AlignCenter)
                star.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
                star.setForeground(QColor(accent if starred else muted))
                star.setToolTip("Click to star — starred vessels pin to the top")
                self.table.setItem(r, 0, star)

                dot = QTableWidgetItem("●" if listened else "")
                dot.setTextAlignment(Qt.AlignCenter)
                dot.setFlags(Qt.ItemIsEnabled)
                dot.setForeground(QColor(accent))
                if listened:
                    dot.setToolTip("Received via live listening")
                self.table.setItem(r, 1, dot)

                for ci, val in enumerate(rowd['labels']):
                    item = QTableWidgetItem(val or "")
                    if ci == 4:  # Zone — accent colour
                        item.setForeground(QColor(accent))
                    self.table.setItem(r, ci + off, item)
            self.table.resizeRowsToContents()
        finally:
            self._populating = False

    def _on_header_clicked(self, section):
        off = self._data_offset
        if section < off:
            return
        col = section - off
        if self._sort_col == col:
            self._sort_asc = not self._sort_asc
        else:
            self._sort_col = col
            self._sort_asc = True
        for r in self.table_data:
            r['listened'] = False
        self._render_main_table()
        h = self.table.horizontalHeader()
        h.setSortIndicatorShown(True)
        h.setSortIndicator(section, Qt.AscendingOrder if self._sort_asc else Qt.DescendingOrder)

    def _on_cell_clicked(self, row, col):
        if row < 0 or row >= len(self._visible_rows):
            return
        if col == 0:
            rowd = self._visible_rows[row]
            rowd['starred'] = not rowd.get('starred', False)
            self._render_main_table()
        elif col >= self._data_offset:
            item = self.table.item(row, col)
            if item is not None:
                self.table.editItem(item)

    def _on_item_changed(self, item):
        if self._populating:
            return
        col = item.column()
        if col < self._data_offset:
            return
        r = item.row()
        if r < 0 or r >= len(self._visible_rows):
            return
        lbl_col = col - self._data_offset
        rowd = self._visible_rows[r]
        rowd['labels'][lbl_col] = item.text()
        if lbl_col == 4:
            rowd['zone'] = item.text()

    def _run_search(self, advance=False):
        """Find rows whose any data cell contains the query, and jump to one. Typing jumps to
        the first match; Enter cycles to the next (wrapping)."""
        if getattr(self, "table", None) is None:
            return
        q = self.table_search.text().strip().lower()
        if not q:
            self.table.clearSelection()
            self._search_pos = -1
            self._search_last_q = ""
            return
        off = self._data_offset
        matches = []   # (row, col) of the first matching cell per row
        for r in range(self.table.rowCount()):
            for c in range(off, self.table.columnCount()):
                it = self.table.item(r, c)
                if it and q in it.text().lower():
                    matches.append((r, c))
                    break
        if not matches:
            self.table.clearSelection()
            return
        if q != self._search_last_q or self._search_pos < 0:
            pos = 0                                   # new query -> first match
        elif advance:
            pos = (self._search_pos + 1) % len(matches)   # Enter -> next, wrap
        else:
            pos = min(self._search_pos, len(matches) - 1)
        self._search_last_q = q
        self._search_pos = pos
        row, col = matches[pos]
        self.table.setCurrentCell(row, col)           # highlights + makes it the current cell
        self.table.scrollToItem(self.table.item(row, col), QAbstractItemView.PositionAtCenter)

    def _flash_new_row(self, seq):
        """Subtle 'new arrival' cue: fade a freshly-added row's text in from transparent to its
        final colour. Targets the row by seq so reordering doesn't misfire."""
        if getattr(self, "table", None) is None:
            return
        row = next((i for i, r in enumerate(self._visible_rows) if r.get('seq') == seq), None)
        if row is None:
            return
        prev = getattr(self, "_row_anim", None)
        if prev is not None:
            prev.stop()
        c = getattr(self, "_theme_colors", {})
        accent = QColor(c.get("accent", "#22d3ee"))
        text = QColor(c.get("text", "#16181d"))
        muted = QColor(c.get("muted", "#5b616e"))
        starred = self._visible_rows[row].get('starred', False)
        zone_col = self._data_offset + 4
        cols = self.table.columnCount()

        def target(cc):
            if cc == 0:
                return accent if starred else muted
            if cc == 1 or cc == zone_col:
                return accent
            return text

        def tick(v):
            # Suppress itemChanged while we recolour (text content is unchanged).
            self._populating = True
            try:
                for cc in range(cols):
                    it = self.table.item(row, cc)
                    if it is not None:
                        col = QColor(target(cc)); col.setAlphaF(float(v))
                        it.setForeground(col)
            except RuntimeError:
                pass
            finally:
                self._populating = False

        tick(0.0)   # hide the row's text immediately so it fades in (no first-frame flash)
        anim = QVariantAnimation(self.table)
        anim.setDuration(500)
        anim.setStartValue(0.0)
        anim.setEndValue(1.0)
        anim.setEasingCurve(QEasingCurve.OutCubic)
        anim.valueChanged.connect(tick)
        anim.start()
        self._row_anim = anim

    def eventFilter(self, obj, event):
        """Deselect the current cell on any click that isn't on a valid cell — empty table
        space or anywhere off the table. App-wide so clicks on non-focusable widgets count."""
        if event.type() == QEvent.MouseButtonPress:
            tbl = getattr(self, "table", None)
            try:
                if tbl is not None and tbl.isVisible() and tbl.state() != QAbstractItemView.EditingState:
                    vp = tbl.viewport()
                    local = vp.mapFromGlobal(event.globalPosition().toPoint())
                    inside_cell = vp.rect().contains(local) and tbl.indexAt(local).isValid()
                    if not inside_cell:
                        tbl.clearSelection()
            except RuntimeError:
                self.table = None
        return super().eventFilter(obj, event)

    def _set_main_status(self, color, text):
        try:
            self.status.setText(text)
            self._style_extbox(color)
        except RuntimeError:
            pass

    def start_main_listening(self, since=None):
        """Continue listening within the main extraction view, feeding new vessels into
        the same (sorted) table rather than a separate window. `since` (a datetime) anchors
        which emails count as new — passed as the extraction start time when auto-started
        after an extraction, so nothing is missed in the gap; defaults to now() otherwise."""
        if not isinstance(since, datetime):
            since = None   # signal-connected calls (e.g. thread.finished) pass junk args
        v, _, _ = validate(None, None, self.email_address, self.folder, self.excel, outlook, self.language)
        if not v:
            self._set_main_status("#f87171", t("extraction_stopped", self.language))
            return
        self.listening_running = True
        self._current_excel = resolve_excel_path(self.excel)
        self.listen_thread = QThread()
        self.listen_worker = ExtractWorker(None)
        generator = process_email(self.email_address, self.folder, self._current_excel, csv_dict, self.listen_worker, since)
        self.listen_worker.generator = generator
        self.listen_worker.moveToThread(self.listen_thread)
        self.listen_worker.new_email.connect(self.add_listening_to_main_table)
        self.listen_thread.started.connect(self.listen_worker.run)
        self.listen_worker.done.connect(self.on_main_listen_done)
        self.listen_worker.done.connect(self.listen_thread.quit)
        self.listen_worker.done.connect(self.listen_worker.deleteLater)
        self.listen_thread.finished.connect(self.listen_thread.deleteLater)
        self.listen_thread.finished.connect(lambda: setattr(self, 'listen_thread', None))
        self.listen_thread.start()
        self._set_main_status("#22d3ee", t("listening_running", self.language))
        self.continue_listen_btn.setText(t("pause_listen", self.language))
        self.continue_listen_btn.show()

    def add_listening_to_main_table(self, email_data):
        if email_data.get("type") == "excel_locked":
            self._set_main_status("#f59e0b", t("listening_paused", self.language) + " — close Excel to continue")
            return
        if email_data.get("type") == "excel_unlocked":
            self._set_main_status("#22d3ee", t("listening_running", self.language))
            return
        try:
            zone, labels = self._row_from_email(email_data)
            self._seq_counter += 1
            self.table_data.append({
                'zone': zone, 'labels': labels,
                'listened': True, 'starred': False, 'seq': self._seq_counter,
            })
            self._render_main_table()
            self._flash_new_row(self._seq_counter)
            self.caption5.setText(f"{t('vessels_extracted', self.language)} {len(self.table_data)}")
        except Exception as e:
            print(f"Error adding listened email to table: {e}")

    def on_main_listen_done(self):
        self.listening_running = False
        if self.page_main is None:
            return
        err = getattr(self.listen_worker, "api_error_key", None)
        limit = getattr(self.listen_worker, "limit_reached", False)
        try:
            if limit:
                self.show_upgrade_dialog()
                self._set_main_status("#22d3ee", t("limit_reached_title", self.language))
            elif err:
                self._set_main_status("#f87171", t(err, self.language))
            else:
                self._set_main_status("#f59e0b", t("listening_paused", self.language))
            self.continue_listen_btn.setText(t("resume_listen", self.language))
        except RuntimeError:
            pass

    def toggle_main_listening(self):
        if self.listening_running:
            if getattr(self, "listen_worker", None):
                self.listen_worker.stop()
            self.listening_running = False
            self.continue_listen_btn.setText(t("resume_listen", self.language))
            self._set_main_status("#f59e0b", t("listening_paused", self.language))
        else:
            try:
                running = (hasattr(self, "listen_thread") and self.listen_thread
                           and self.listen_thread.isRunning())
            except RuntimeError:
                self.listen_thread = None
                running = False
            if running:
                try:
                    self.listen_thread.finished.disconnect(self.start_main_listening)
                except (RuntimeError, TypeError):
                    pass
                self.listen_thread.finished.connect(self.start_main_listening)
            else:
                self.start_main_listening()

    def on_extraction_done(self):
        error_key = getattr(self.worker, "api_error_key", None)
        limit_hit = getattr(self.worker, "limit_reached", False)

        if limit_hit:
            self.show_upgrade_dialog()
            self.extheader.setText(t("limit_reached_title", self.language))
            self.extheader.setStyleSheet("font: bold 25px; color: #22d3ee;")
            self.status.setText(t("extraction_stopped", self.language))
            self._style_extbox("#f87171")
        elif error_key:
            self.extheader.setText(t(error_key, self.language))
            self.extheader.setStyleSheet("font: bold 25px; color: red;")
            self.status.setText(t("extraction_stopped", self.language))
            self._style_extbox("#f87171")
        elif not self.table_data:
            self.extheader.setText(t("extraction_complete_none", self.language))
            self.extheader.setStyleSheet("font: bold 25px;")
            self.table.setRowCount(0)
            self.status.setText(t("extraction_stopped", self.language))
            self._style_extbox("#f87171")
        else:
            self.extheader.setText(t("extraction_complete", self.language))
            self.extheader.setStyleSheet("font: bold 25px;")
            self._render_main_table()
            # Anchor listening to when the extraction began, so emails that arrived while it
            # was running are caught (closes the snapshot→listen gap).
            self.start_main_listening(getattr(self, "_extraction_started", None))

        self.btn.setEnabled(True)
        self.extracting_running = False
        self.stop_btn.hide()
        self.new_extract_btn.show()
        self.open_excel_btn.show()

        self.tray.showMessage(
            "Extraction complete",
            f"{len(self.table_data)} vessels extracted",
            QIcon(resource_path("icon.png")),
            3000
        )

        for field in (self.input_day, self.input_month, self.input_year,
                      self.input_hour, self.input_minute):
            field.setCurrentIndex(-1)
            field.clearEditText()
        self.day = self.month = self.year = self.hours = self.minutes = ""
        self.date = None
        self.time = None

    def on_listen_done(self):
        error_key = getattr(self.listen_worker, "api_error_key", None)
        limit_hit = getattr(self.listen_worker, "limit_reached", False)
        if limit_hit:
            self.show_upgrade_dialog()
            self.lheader.setText(t("limit_reached_title", self.language))
            self.lheader.setStyleSheet("font: bold 40px; color: #22d3ee;")
            self.statusl.setText(t("extraction_stopped", self.language))
            self.lbox.setStyleSheet(self._status_box_qss("#22d3ee"))
            self.listening_running = False
            self.listen_toggle_btn.setText(t("resume_listen", self.language))
        elif error_key:
            self.lheader.setText(t(error_key, self.language))
            self.lheader.setStyleSheet("font: bold 40px; color: red;")
            self.statusl.setText(t(error_key, self.language))
            self.lbox.setStyleSheet(self._status_box_qss("#f87171"))
            self.listening_running = False
            self.listen_toggle_btn.setText(t("resume_listen", self.language))

    def toggle_listening(self):
        if self.listening_running:
            if hasattr(self, "listen_worker") and self.listen_worker:
                self.listen_worker.stop()

            self.listening_running = False
            self.listen_toggle_btn.setText(t("resume_listen", self.language))
            self.statusl.setText(t("listening_paused", self.language))
            self.lbox.setStyleSheet(self._status_box_qss("#f59e0b"))
        else:
            try:
                thread_running = (
                    hasattr(self, "listen_thread")
                    and self.listen_thread
                    and self.listen_thread.isRunning()
                )
            except RuntimeError:
                self.listen_thread = None
                thread_running = False
            if thread_running:
                try:
                    self.listen_thread.finished.disconnect(self._restart_listening)
                except (RuntimeError, TypeError):
                    pass
                self.listen_thread.finished.connect(self._restart_listening)
            else:
                self._restart_listening()

    def _restart_listening(self):
        if not access_allowed():
            self.show_upgrade_dialog()
            return
        self.handle_listen()
        self.listen_toggle_btn.setText(t("pause_listen", self.language))
        self.statusl.setText(t("listening_running", self.language))
        self.lbox.setStyleSheet(self._status_box_qss("#34d399"))

    def open_excel_file(self):
        path = getattr(self, '_current_excel', None) or resolve_excel_path(self.excel)
        QDesktopServices.openUrl(QUrl.fromLocalFile(path))

    def show_upgrade_dialog(self):
        dialog = QDialog(self)
        dialog.setWindowTitle(t("limit_reached_title", self.language))
        dialog.setFixedWidth(620)

        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(30, 30, 30, 30)
        layout.setSpacing(20)

        msg = QLabel(t("limit_reached_body", self.language))
        msg.setWordWrap(True)
        msg.setStyleSheet("font: 16px;")
        layout.addWidget(msg)

        btn_row = QHBoxLayout()
        upgrade_btn = QPushButton(t("upgrade_btn", self.language))
        upgrade_btn.setFixedSize(280, 60)
        upgrade_btn.setProperty("variant", "primary")
        upgrade_btn.clicked.connect(lambda: (QDesktopServices.openUrl(QUrl("https://ko-fi.com/mailaiuk/tiers")), dialog.accept()))

        close_btn = QPushButton(t("donation_close", self.language))
        close_btn.setFixedSize(150, 60)
        close_btn.clicked.connect(dialog.reject)

        btn_row.addWidget(upgrade_btn)
        btn_row.addSpacing(20)
        btn_row.addWidget(close_btn)
        layout.addLayout(btn_row)

        dialog.exec()

    def start_update_check(self):
        # Native-style updates: silently check + download in the background, then apply on the
        # next restart. No prompt, no forced relaunch.
        if not getattr(sys, "frozen", False):
            return
        if staged_update_ready():
            self._mark_update_ready()   # downloaded in a previous session, still waiting
            return
        self._upd_check_thread = QThread()
        self._upd_checker = UpdateChecker()
        self._upd_checker.moveToThread(self._upd_check_thread)
        self._upd_check_thread.started.connect(self._upd_checker.run)
        self._upd_checker.update_available.connect(self._begin_update_download)
        self._upd_checker.update_available.connect(self._upd_check_thread.quit)
        self._upd_check_thread.start()

    def _begin_update_download(self, version):
        # A newer version exists — download + stage it quietly off the UI thread.
        self._upd_thread = QThread()
        self._upd_worker = UpdateWorker()
        self._upd_worker.moveToThread(self._upd_thread)
        self._upd_thread.started.connect(self._upd_worker.run)
        self._upd_worker.done.connect(self._on_update_downloaded)
        self._upd_worker.done.connect(self._upd_thread.quit)
        self._upd_thread.start()

    def _on_update_downloaded(self, ok, err):
        if ok:
            self._mark_update_ready()
        else:
            logger.warning(f"Background update download failed: {err}")

    def _mark_update_ready(self):
        # Update is staged; it will be applied automatically when the app is restarted.
        self._update_ready = True
        try:
            self.ver.setText("  update ready · restart to apply")
            self.ver.setToolTip("A new version has been downloaded. "
                                "Restart Mail AI to apply it.")
        except Exception:
            pass
        try:
            self.tray.showMessage(
                "Mail AI — update ready",
                "A new version has been downloaded. Restart Mail AI to apply it.",
                QIcon(resource_path("icon.png")), 6000,
            )
        except Exception:
            pass

    def activate_license(self):
        key = self.license_input.text().strip()
        if validate_license_key(key):
            config = load_config()
            config["license_key"] = key
            config["is_pro"] = True
            save_config(config)
            self.key_feedback_label.setText(t("pro_active", self.language))
            self.key_feedback_label.setStyleSheet("font: 14px; color: #22d3ee;")
            self.license_status_label.setText(t("pro_active", self.language))
            self.license_status_label.setStyleSheet("font: 16px; color: #22d3ee;")
        else:
            self.key_feedback_label.setText(t("invalid_key", self.language))
            self.key_feedback_label.setStyleSheet("font: 14px; color: red;")

    def toggle_theme(self):
        config = load_config()
        current = config.get("theme", "light")

        if current == "dark":
            self.apply_theme("light")
            config["theme"] = "light"
            self.theme_btn.setText(t("switch_dark", self.language))
        else:
            self.apply_theme("dark")
            config["theme"] = "dark"
            self.theme_btn.setText(t("switch_light", self.language))

        save_config(config)

    def apply_theme(self, theme):
        if theme == "dark":
            c = {
                "bg": "#0a0b0d", "surface": "#141517", "surface2": "#1a1b1e", "border": "#26272b",
                "text": "#f4f4f5", "muted": "#9b9ba3", "dim": "#5c5d66",
                "accent": "#22d3ee", "accent_hover": "#67e8f9", "on_accent": "#0a0b0d",
                "sidebar_bg": "#0c0d0f",
            }
        else:
            c = {
                "bg": "#ffffff", "surface": "#f7f8f9", "surface2": "#eef0f2", "border": "#e5e7eb",
                "text": "#16181d", "muted": "#5b616e", "dim": "#9aa0aa",
                "accent": "#0891b2", "accent_hover": "#0e7490", "on_accent": "#ffffff",
                "sidebar_bg": "#f7f8f9",
            }
        self._theme_colors = c
        get_font(self.language)  # ensure Inter is registered before building QSS
        ui_font = "Source Han Sans SC" if self.language == "中文" else (_INTER_FAMILY or "Inter")
        # Custom combobox chevron (theme-matched). Forward slashes for QSS url() on Windows.
        _chev = resource_path("chevron-dark.svg" if theme == "dark" else "chevron-light.svg").replace("\\", "/")

        QApplication.instance().setStyleSheet(f"""
            QMainWindow {{ background-color: {c['bg']}; }}
            QWidget {{ background-color: transparent; color: {c['text']}; }}
            QFrame {{ background-color: transparent; }}
            QLabel {{ color: {c['text']}; background: transparent; }}

            QPushButton {{
                background-color: {c['surface']};
                color: {c['text']};
                border: 1px solid {c['border']};
                border-radius: 8px;
                padding: 8px 18px;
                font-family: '{ui_font}';
                font-size: 14px;
                font-weight: 600;
            }}
            QPushButton:hover {{ border-color: {c['accent']}; color: {c['accent']}; }}
            QPushButton:pressed {{ background-color: {c['surface2']}; }}
            QPushButton:disabled {{ color: {c['dim']}; border-color: {c['border']}; background-color: {c['bg']}; }}

            QLineEdit {{
                background-color: {c['surface']};
                color: {c['text']};
                border: 1px solid {c['border']};
                border-radius: 8px;
                padding: 8px 12px;
                selection-background-color: {c['accent']};
                selection-color: {c['on_accent']};
            }}
            QLineEdit:focus {{ border: 1px solid {c['accent']}; }}

            QComboBox {{
                background-color: {c['surface']};
                color: {c['text']};
                border: 1px solid {c['border']};
                border-radius: 8px;
                padding: 6px 12px;
                font-size: 13pt;
            }}
            QComboBox:hover {{ border-color: {c['accent']}; }}
            QComboBox::drop-down {{
                subcontrol-origin: padding;
                subcontrol-position: center right;
                width: 26px;
                border: none;
                background: transparent;
            }}
            QComboBox::down-arrow {{ image: url("{_chev}"); width: 14px; height: 14px; }}
            QComboBox QAbstractItemView {{
                background-color: {c['surface']};
                color: {c['text']};
                border: 1px solid {c['border']};
                selection-background-color: {c['surface2']};
                selection-color: {c['accent']};
                outline: none;
            }}

            QScrollArea {{ background-color: transparent; border: none; }}

            QScrollBar:vertical {{ background: transparent; width: 10px; margin: 2px; }}
            QScrollBar::handle:vertical {{ background: {c['border']}; border-radius: 5px; min-height: 30px; }}
            QScrollBar::handle:vertical:hover {{ background: {c['dim']}; }}
            QScrollBar:horizontal {{ background: transparent; height: 10px; margin: 2px; }}
            QScrollBar::handle:horizontal {{ background: {c['border']}; border-radius: 5px; min-width: 30px; }}
            QScrollBar::handle:horizontal:hover {{ background: {c['dim']}; }}
            QScrollBar::add-line, QScrollBar::sub-line {{ width: 0; height: 0; }}
            QScrollBar::add-page, QScrollBar::sub-page {{ background: transparent; }}

            QToolTip {{
                background-color: {c['surface2']};
                color: {c['text']};
                border: 1px solid {c['border']};
                padding: 4px 8px;
                border-radius: 6px;
            }}
        """)

        self.sidebar.setStyleSheet(f"""
            QFrame {{
                background-color: {c['sidebar_bg']};
                border: none;
                border-right: 1px solid {c['border']};
            }}
        """)

        self.logo.setText("Mail AI")
        self.logo.setStyleSheet(f"""
            font-family: '{ui_font}';
            font-size: 18px;
            font-weight: 800;
            color: {c['text']};
            background-color: {c['sidebar_bg']};
            border-bottom: 1px solid {c['border']};
        """)

        if hasattr(self, "ver") and self.ver:
            self.ver.setStyleSheet(f"""
                font-family: 'Inter';
                font-size: 11px;
                color: {c['dim']};
                background-color: transparent;
                border-top: 1px solid {c['border']};
                padding-left: 16px;
            """)

        sidebar_font = "Source Han Sans SC" if self.language == "中文" else ui_font
        self._sidebar_btn_qss = f"""
            QPushButton {{
                background-color: transparent;
                color: {c['muted']};
                font-family: '{sidebar_font}';
                font-size: 15px;
                font-weight: 500;
                border: none;
                border-left: 2px solid transparent;
                border-radius: 0px;
                text-align: left;
                padding-left: 18px;
            }}
            QPushButton:hover {{
                background-color: {c['surface']};
                color: {c['text']};
                border-left: 2px solid transparent;
            }}
        """
        self._sidebar_btn_active_qss = f"""
            QPushButton {{
                background-color: transparent;
                color: {c['text']};
                font-family: '{sidebar_font}';
                font-size: 15px;
                font-weight: 600;
                border: none;
                border-left: 2px solid {c['accent']};
                border-radius: 0px;
                text-align: left;
                padding-left: 18px;
            }}
            QPushButton:hover {{
                background-color: {c['surface']};
                color: {c['text']};
            }}
        """

        for btn in [self.extract_sidebar_btn, self.filtering_sidebar_btn, self.settings_sidebar_btn]:
            btn.setStyleSheet(self._sidebar_btn_qss)
        if hasattr(self, "_active_sidebar_btn") and self._active_sidebar_btn:
            self._active_sidebar_btn.setStyleSheet(self._sidebar_btn_active_qss)

        self.main_widget.set_theme(theme)
        self.pages.set_theme(theme)
        if hasattr(self, '_settings_content') and self._settings_content:
            self._settings_content.set_theme(theme)

        # Re-theme the results table: its stylesheet and per-cell colours/fonts are baked at
        # render time, so on a theme toggle we must re-apply the QSS and rebuild the rows,
        # otherwise some cells keep the old palette and become unreadable.
        if getattr(self, "table", None) is not None:
            try:
                self.table.setStyleSheet(self._table_qss())
                self._render_main_table()
            except RuntimeError:
                self.table = None  # page was torn down
        if getattr(self, "extbox", None) is not None:
            self._style_extbox(getattr(self, "_main_status_color", "#34d399"))

        QApplication.setFont(get_font(self.language))

        if hasattr(self, 'setup_wizard') and self.setup_wizard:
            self.setup_wizard.update_nav()

    def language_changed(self, language):
        self.language = language
        config = load_config()
        config["language"] = language
        save_config(config)
        self.retranslate()

    def retranslate(self):

        old_home = self.pages.widget(0)
        old_filtering = self.pages.widget(1)
        old_settings = self.pages.widget(2)

        self.page_home = self.create_home_page()
        self.page_filtering = self.create_filtering_page()
        self.page_settings = self.create_settings_page()

        self.pages.insertWidget(0, self.page_home)
        self.pages.insertWidget(1, self.page_filtering)
        self.pages.insertWidget(2, self.page_settings)

        self.pages.removeWidget(old_home)
        self.pages.removeWidget(old_filtering)
        self.pages.removeWidget(old_settings)

        old_home.deleteLater()
        old_filtering.deleteLater()
        old_settings.deleteLater()

        self.extract_sidebar_btn.setText(t("extract", self.language))
        self.filtering_sidebar_btn.setText(t("filtering", self.language))
        self.settings_sidebar_btn.setText(t("settings", self.language))

        self.pages.setCurrentWidget(self.page_settings)

        QApplication.setFont(get_font(self.language))
        current_theme = load_config().get("theme", "light")
        self.apply_theme(current_theme)


if __name__ == "__main__":
    load_existing_vessels()
    load_email_ids()
    config = load_config()
    refresh_access_state()
    cleanup_old_update()
    try:
        QApplication.setHighDpiScaleFactorRoundingPolicy(
            Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
    except Exception:
        pass
    app = QApplication(sys.argv)
    app.setWindowIcon(QIcon(resource_path("icon.png")))
    # Apply any staged update when the app quits, so the next launch runs the new version.
    app.aboutToQuit.connect(apply_staged_update)
    window = MainWindow()
    window.apply_theme(config.get("theme", "light"))
    window.show()
    # Dismiss the PyInstaller startup splash (frozen builds with --splash) now the UI is up.
    try:
        import pyi_splash
        pyi_splash.close()
    except Exception:
        pass
    window.start_update_check()
    sys.exit(app.exec())