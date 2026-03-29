import sys
import pythoncom
from PySide6.QtWidgets import (
    QApplication, QWidget, QMainWindow, QVBoxLayout, QHBoxLayout, QSystemTrayIcon, QMenu,
    QPushButton, QLabel, QFrame, QStackedWidget, QLineEdit, QScrollArea, QGridLayout, QComboBox,
    QFileDialog, QDialog, QMessageBox
)
from PySide6.QtGui import QFontDatabase, QFont, QColor, QPalette, QIcon, QDesktopServices
from PySide6.QtCore import Qt, QObject, Signal, Slot, QThread, QTimer, QUrl

import ctypes
ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID('Mail AI')


from logic import *

try:
    pythoncom.CoInitialize()
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
except Exception as e:
    outlook = None

csv_file = resource_path("WPIUpdated.csv")
if not os.path.exists(csv_file):
    QMessageBox.critical(None, "Missing Resource", f"Required file not found: WPI.csv\nPlease reinstall the application.")
    sys.exit(1)
csv_dict = merge_custom_zones(load_csv_into_dict(csv_file))

class ExtractWorker(QObject):
    new_email = Signal(dict)
    done = Signal()

    def __init__(self, generator):
        super().__init__()
        self.generator = generator
        self.running = True
        self.api_error_key = None

    def run(self):
        for email in self.generator:
            if not self.running:
                break
            if email.get("type") == "api_error":
                self.api_error_key = email["error_key"]
                break
            self.new_email.emit(email)
        self.done.emit()

    def stop(self):
        self.running = False

def get_font(language):

    QFontDatabase.addApplicationFont(resource_path("DM_Mono/DMMono-Regular.ttf"))
    QFontDatabase.addApplicationFont(resource_path("DM_Mono/DMMono-Medium.ttf"))

    if language == "中文":
        font_id = QFontDatabase.addApplicationFont(resource_path("SourceHanSansSC-Regular.otf"))
    else:
        font_id = QFontDatabase.addApplicationFont(resource_path("Syne/Syne-VariableFont_wght.ttf"))

    if font_id != -1:
        family = QFontDatabase.applicationFontFamilies(font_id)[0]
        return QFont(family, 10)
    return QApplication.font()

class GridWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.theme = "dark"

    def set_theme(self, theme):
        self.theme = theme
        self.update()  # triggers repaint

    def paintEvent(self, event):
        from PySide6.QtGui import QPainter, QPen
        painter = QPainter(self)

        if self.theme == "dark":
            painter.fillRect(self.rect(), QColor("#080f1a"))
            pen = QPen(QColor("#0b1a2e"))
        else:
            painter.fillRect(self.rect(), QColor("#f0f9ff"))
            pen = QPen(QColor("#bae6fd"))

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
        self.update()  # triggers repaint

    def paintEvent(self, event):
        from PySide6.QtGui import QPainter, QPen
        painter = QPainter(self)

        if self.theme == "dark":
            painter.fillRect(self.rect(), QColor("#080f1a"))
            pen = QPen(QColor("#0b1a2e"))
        else:
            painter.fillRect(self.rect(), QColor("#f0f9ff"))
            pen = QPen(QColor("#bae6fd"))

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

        # --- Page 0: Welcome ---
        welcome = QWidget()
        wl = QVBoxLayout(welcome)
        wl.setAlignment(Qt.AlignCenter)
        wt = QLabel(t("setup_welcome_title", self.language))
        wt.setStyleSheet("font: bold 55px;")
        wt.setAlignment(Qt.AlignCenter)
        ws = QLabel(t("setup_welcome_subtitle", self.language))
        ws.setStyleSheet("font: normal 22px;")
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

        # --- Page 1: Email ---
        email_page = QWidget()
        el = QVBoxLayout(email_page)
        el.setAlignment(Qt.AlignCenter)
        self.email_step = QLabel(f"{t('setup_step', self.language)} 1 / 3")
        self.email_step.setStyleSheet("font: 600 14px; color: #0891b2;")
        self.email_step.setAlignment(Qt.AlignCenter)
        et = QLabel(t("setup_email_title", self.language))
        et.setStyleSheet("font: bold 40px;")
        et.setAlignment(Qt.AlignCenter)
        ed = QLabel(t("setup_email_desc", self.language))
        ed.setStyleSheet("font: normal 18px;")
        ed.setAlignment(Qt.AlignCenter)
        self.email_input = QLineEdit()
        self.email_input.setPlaceholderText("e.g. johndoe@gmail.com")
        self.email_input.setFixedSize(500, 45)
        self.email_input.setStyleSheet("QLineEdit { font-size: 16px; }")
        self.email_input.setMaxLength(254)
        self.email_input.setText(config.get("email_address", ""))
        self.email_input.textChanged.connect(self.update_nav)
        el.addWidget(self.email_step)
        el.addSpacing(10)
        el.addWidget(et)
        el.addSpacing(8)
        el.addWidget(ed)
        el.addSpacing(25)
        el.addWidget(self.email_input, alignment=Qt.AlignCenter)
        self.pages_list.append(email_page)

        # --- Page 2: Folder ---
        folder_page = QWidget()
        fl = QVBoxLayout(folder_page)
        fl.setAlignment(Qt.AlignCenter)
        self.folder_step = QLabel(f"{t('setup_step', self.language)} 2 / 3")
        self.folder_step.setStyleSheet("font: 600 14px; color: #0891b2;")
        self.folder_step.setAlignment(Qt.AlignCenter)
        ft = QLabel(t("setup_folder_title", self.language))
        ft.setStyleSheet("font: bold 40px;")
        ft.setAlignment(Qt.AlignCenter)
        fd = QLabel(t("setup_folder_desc", self.language))
        fd.setStyleSheet("font: normal 18px;")
        fd.setAlignment(Qt.AlignCenter)
        self.folder_input = QLineEdit()
        self.folder_input.setPlaceholderText("e.g. Inbox, shipbroking")
        self.folder_input.setFixedSize(500, 45)
        self.folder_input.setStyleSheet("QLineEdit { font-size: 16px; }")
        self.folder_input.setMaxLength(254)
        self.folder_input.setText(config.get("folder"))
        self.folder_input.textChanged.connect(self.update_nav)
        fl.addWidget(self.folder_step)
        fl.addSpacing(10)
        fl.addWidget(ft)
        fl.addSpacing(8)
        fl.addWidget(fd)
        fl.addSpacing(25)
        fl.addWidget(self.folder_input, alignment=Qt.AlignCenter)
        self.pages_list.append(folder_page)

        # --- Page 3: Excel ---
        excel_page = QWidget()
        xl = QVBoxLayout(excel_page)
        xl.setAlignment(Qt.AlignCenter)
        self.excel_step = QLabel(f"{t('setup_step', self.language)} 3 / 3")
        self.excel_step.setStyleSheet("font: 600 14px; color: #0891b2;")
        self.excel_step.setAlignment(Qt.AlignCenter)
        xt = QLabel(t("setup_excel_title", self.language))
        xt.setStyleSheet("font: bold 40px;")
        xt.setAlignment(Qt.AlignCenter)
        xd = QLabel(t("setup_excel_desc", self.language))
        xd.setStyleSheet("font: normal 18px;")
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
        browse_btn.setStyleSheet("font: normal 16px;")
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

        # --- Page 4: Finish ---
        finish = QWidget()
        fnl = QVBoxLayout(finish)
        fnl.setAlignment(Qt.AlignCenter)
        fnt = QLabel(t("setup_finish_title", self.language))
        fnt.setStyleSheet("font: bold 55px;")
        fnt.setAlignment(Qt.AlignCenter)
        fnd = QLabel(t("setup_finish_desc", self.language))
        fnd.setStyleSheet("font: normal 20px;")
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

        # --- Navigation bar ---
        nav = QHBoxLayout()
        nav.setContentsMargins(40, 0, 40, 30)
        self.back_btn = QPushButton(t("setup_back", self.language))
        self.back_btn.setFixedSize(120, 50)
        self.back_btn.clicked.connect(self.go_back)
        self.back_btn.setStyleSheet("font: normal 20px;")

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
        self.next_btn.setStyleSheet("font: normal 20px;")

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

        # Disable next if current input field is empty
        if idx == 1:
            self.next_btn.setEnabled(bool(self.email_input.text().strip()))
        elif idx == 2:
            self.next_btn.setEnabled(bool(self.folder_input.text().strip()))
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
            save_config(load_config() | {"email_address": self.email_input.text().strip()})
        elif idx == 2:
            save_config(load_config() | {"folder": self.folder_input.text().strip()})
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
        self.setFixedSize(1280, 800)
        self.setWindowIcon(QIcon(resource_path("icon.png")))


        config = load_config()
        self.email_address = config.get("email_address", "")
        self.folder = config.get("folder", "")
        self.excel = config.get("excel", "")
        self.language = config.get("language", "English")
        self.is_first_run = not config.get("setup_complete", False)

        self.col_widths = [220, 200, 120, 160, 110, 150, 120, 140]

        QApplication.setFont(get_font(self.language))

        self.setup_ui()

    def setup_ui(self):
        self.extracting_running = False
        self.listening_running = False
        config = load_config()
        self.emails_processed = config.get("emails_processed", 0)
        self._last_donation_milestone = self.emails_processed // 5000

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

        self.logo = QLabel("  MAIL AI")
        self.logo.setFixedHeight(70)
        self.logo.setStyleSheet("""
            font: 800 16px;
            font-family: 'Syne';
            color: #f0f9ff;
            letter-spacing: 4px;
            background-color: #0a1628;
            border-bottom: 1px solid #1a3a5c;
            padding-left: 16px;
        """)
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
                    font-family: 'DM Mono';
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

        self.extract_sidebar_btn.clicked.connect(self.on_extract_sidebar_clicked)
        self.filtering_sidebar_btn.clicked.connect(lambda: self.switch_page(self.page_filtering))
        self.settings_sidebar_btn.clicked.connect(lambda: self.switch_page(self.page_settings))

        self.sidebar_layout.addStretch()

        ver = QLabel("  mailai.uk         v1.1")
        ver.setFixedHeight(40)
        ver.setStyleSheet("""
            font-family: 'DM Mono';
            font-size: 11px;
            color: #3a5a78;
            background-color: transparent;
            border-top: 1px solid #1a3a5c;
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
        current_theme = load_config().get("theme", "dark")
        self.apply_theme(current_theme)

        # Rebuild the wizard with the new language
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
        header.setStyleSheet("font: bold 75px;")
        header.setAlignment(Qt.AlignCenter)


        caption = QLabel(t("extract_something", self.language))
        caption.setStyleSheet("font: normal 30px;")
        caption.setAlignment(Qt.AlignCenter)

        button_row = QHBoxLayout()
        button_row.setSpacing(30)

        btn_left = QPushButton(t("extract", self.language))
        btn_left.setStyleSheet("font-weight: 600;")
        btn_left.clicked.connect(self.show_extract_page)

        btn_right = QPushButton(t("listen", self.language))
        btn_right.setStyleSheet("font-weight: 600;")
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
        header.setStyleSheet("font: bold 50px;")

        caption = QLabel(t("email_caption", self.language))
        caption.setStyleSheet("font: 600 20px;")
        input_box = QLineEdit()
        input_box.setMaxLength(254)
        input_box.setPlaceholderText("e.g. johndoe@gmail.com")
        input_box.setFixedSize(700, 40)
        input_box.setStyleSheet("QLineEdit { font-size: 16px; }")
        input_box.setText(getattr(self, "email_address", ""))
        input_box.textEdited.connect(self.email_entered)

        caption2 = QLabel(t("folder_caption", self.language))
        caption2.setStyleSheet("font: 600 20px;")
        input_box2 = QLineEdit()
        input_box2.setMaxLength(254)
        input_box2.setPlaceholderText("e.g. Inbox, Archive")
        input_box2.setFixedSize(700, 40)
        input_box2.setStyleSheet("QLineEdit { font-size: 16px; }")
        input_box2.setText(getattr(self, "folder", ""))
        input_box2.textEdited.connect(self.folder_entered)

        caption3 = QLabel(t("excel_caption", self.language))
        caption3.setStyleSheet("font: 600 20px;")
        input_box3 = QLineEdit()
        input_box3.setMaxLength(254)
        input_box3.setPlaceholderText("e.g. extraction.xlsx")
        input_box3.setFixedSize(700, 40)
        input_box3.setStyleSheet("QLineEdit { font-size: 16px; }")
        input_box3.setText(getattr(self, "excel", ""))
        input_box3.textEdited.connect(self.excel_entered)

        caption4 = QLabel(t("clear_duplicates_caption", self.language))
        caption4.setStyleSheet("font: 600 20px;")
        self.refresh_btn = QPushButton(t("clear_duplicates_btn", self.language))
        self.refresh_btn.setStyleSheet("font-weight: 600;")
        self.refresh_btn.setFixedSize(250, 80)
        self.refresh_btn.clicked.connect(self.refresh_duplicates)

        content_layout.addWidget(header)
        content_layout.addSpacing(25)
        content_layout.addWidget(caption)
        content_layout.addWidget(input_box)
        content_layout.addSpacing(25)
        content_layout.addWidget(caption2)
        content_layout.addWidget(input_box2)
        content_layout.addSpacing(25)
        content_layout.addWidget(caption3)
        content_layout.addWidget(input_box3)
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
        content.set_theme(load_config().get("theme", "dark"))
        self._settings_content = content
        content_layout = QVBoxLayout(content)
        content_layout.setAlignment(Qt.AlignTop | Qt.AlignLeft)

        header = QLabel(t("settings", self.language))
        header.setStyleSheet("font: bold 50px;")

        theme_label = QLabel(t("theme", self.language))
        theme_label.setStyleSheet("font: 600 20px;")

        if load_config().get("theme", "dark") == "dark":
            self.theme_btn = QPushButton(t("switch_light", self.language))
        else:
            self.theme_btn = QPushButton(t("switch_dark", self.language))
            
        self.theme_btn.setFixedSize(250, 80)
        self.theme_btn.setStyleSheet("font: 600 18px;")
        self.theme_btn.clicked.connect(self.toggle_theme)

        language_label = QLabel(t("language", self.language))
        language_label.setStyleSheet("font: 600 20px;")

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

        # Custom zone mappings section
        zones_header = QLabel(t("custom_zones_header", self.language))
        zones_header.setStyleSheet("font: bold 30px;")
        content_layout.addWidget(zones_header)

        zones_desc = QLabel(t("custom_zones_desc", self.language))
        zones_desc.setStyleSheet("font: 16px;")
        zones_desc.setWordWrap(True)
        content_layout.addWidget(zones_desc)
        content_layout.addSpacing(10)

        # Input row for adding new mappings
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

        # List of current custom mappings
        zones_list_label = QLabel(t("custom_zones_list", self.language))
        zones_list_label.setStyleSheet("font: 600 18px;")
        content_layout.addWidget(zones_list_label)
        content_layout.addSpacing(5)

        self.zones_list_container = QVBoxLayout()
        content_layout.addLayout(self.zones_list_container)
        self.refresh_zones_list()

        content_layout.addSpacing(40)
        donate_label = QLabel(t("donate_optional", self.language))
        donate_label.setStyleSheet("font: bold 30px;")
        content_layout.addWidget(donate_label)
        content_layout.addSpacing(10)
        donate_btn = QPushButton(t("donation_btn", self.language))
        donate_btn.setFixedSize(250, 80)
        donate_btn.setStyleSheet("font-weight: 600;")
        donate_btn.clicked.connect(self.show_donation_popup)
        content_layout.addWidget(donate_btn)
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
        header.setStyleSheet("font: bold 50px;")

        self.caption = QLabel("")
        self.caption2 = QLabel("")
        self.caption3 = QLabel("")
        self.caption4 = QLabel("")
        self.captione1 = QLabel("")
        self.captione2 = QLabel("")

        input_row = QHBoxLayout()
        input_row.setAlignment(Qt.AlignLeft)

        caption5 = QLabel(t("date_caption", self.language))
        caption5.setStyleSheet("font: 600 20px;")

        self.input_day = QLineEdit()
        self.input_day.setPlaceholderText("Day")
        self.input_day.setFixedSize(80, 40)
        self.input_day.setMaxLength(2)
        self.input_day.setStyleSheet("QLineEdit { font-size: 16px; }")
        self.input_day.textEdited.connect(lambda text: self.date_entered(text, "d"))

        self.input_month = QLineEdit()
        self.input_month.setPlaceholderText("Month")
        self.input_month.setFixedSize(80, 40)
        self.input_month.setMaxLength(2)
        self.input_month.setStyleSheet("QLineEdit { font-size: 16px; }")
        self.input_month.textEdited.connect(lambda text: self.date_entered(text, "m"))

        self.input_year = QLineEdit()
        self.input_year.setPlaceholderText("Year")
        self.input_year.setFixedSize(100, 40)
        self.input_year.setMaxLength(4)
        self.input_year.setStyleSheet("QLineEdit { font-size: 16px; }")
        self.input_year.textEdited.connect(lambda text: self.date_entered(text, "y"))

        input_row.addWidget(self.input_day)
        input_row.addWidget(self.input_month)
        input_row.addWidget(self.input_year)

        input_row2 = QHBoxLayout()
        input_row2.setAlignment(Qt.AlignLeft)

        caption6 = QLabel(t("time_caption", self.language))
        caption6.setWordWrap(True)
        caption6.setMaximumWidth(1000)
        caption6.setStyleSheet("font: 600 20px;")

        self.input_hour = QLineEdit()
        self.input_hour.setPlaceholderText("Hour")
        self.input_hour.setFixedSize(80, 40)
        self.input_hour.setMaxLength(2)
        self.input_hour.setStyleSheet("QLineEdit { font-size: 16px; }")
        self.input_hour.textEdited.connect(lambda text: self.time_entered(text, "h"))

        self.input_minute = QLineEdit()
        self.input_minute.setPlaceholderText("Minutes")
        self.input_minute.setFixedSize(80, 40)
        self.input_minute.setMaxLength(2)
        self.input_minute.setStyleSheet("QLineEdit { font-size: 16px; }")
        self.input_minute.textEdited.connect(lambda text: self.time_entered(text, "m"))

        self.input_ampm = QLineEdit()
        self.input_ampm.setPlaceholderText("am/pm")
        self.input_ampm.setFixedSize(100, 40)
        self.input_ampm.setMaxLength(4)
        self.input_ampm.setStyleSheet("QLineEdit { font-size: 16px; }")
        self.input_ampm.textEdited.connect(lambda text: self.time_entered(text, "ampm"))

        input_row2.addWidget(self.input_hour)
        input_row2.addWidget(self.input_minute)
        input_row2.addWidget(self.input_ampm)

        self.btn = QPushButton(t("start_extracting", self.language))
        self.btn.setStyleSheet("""
        QPushButton { font-weight: 600; }
        QToolTip {
            background-color: #333;
            color: white;
            padding: 6px;
            font-size: 15px;
        }
        """)
        self.btn.setFixedSize(200, 80)
        self.btn.clicked.connect(self.handle_extract)
        self.btn.setEnabled(False)
        self.btn.setToolTip(t("tooltip", self.language))

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

    def create_main_page(self):
        content = QWidget()
        content_layout = QVBoxLayout(content)
        content_layout.setAlignment(Qt.AlignTop | Qt.AlignLeft)

        self.extheader = QLabel(t("current_extraction", self.language))
        self.extheader.setStyleSheet("font: bold 50px;")

        self.extbox = QFrame()
        self.extbox.setWindowFlags(Qt.FramelessWindowHint)
        self.extbox.setAttribute(Qt.WA_TranslucentBackground)
        self.extbox.setFixedSize(600, 70)
        self.extbox.setStyleSheet("background-color: rgba(0, 255, 0, 100); margin-left: -10px;")

        self.status = QLabel(t("extraction_running", self.language))
        self.status.setStyleSheet("font: 600 20px; padding: 15px;")

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
        self.caption5.setStyleSheet("font: 600 20px;")

        self.scrollf = QScrollArea()
        self.scrollf.setFixedSize(1000, 500)
        self.scrollf.setWidgetResizable(True)
        self.scrollf.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        self.scrollf.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        self.container = QWidget()
        self.container.setMinimumWidth(1510)
        self.row = 1

        self.grid = QGridLayout(self.container)
        self.grid.setContentsMargins(10, 10, 10, 10)
        self.grid.setHorizontalSpacing(30)
        self.grid.setVerticalSpacing(15)
        self.grid.setRowStretch(0, 0)
        self.grid.setAlignment(Qt.AlignTop)

        headers = [
            t("sender", self.language),
            t("subject", self.language),
            t("date", self.language),
            "MV", "DWT/Built",
            t("location", self.language),
            t("open_date", self.language),
            t("zone", self.language),
        ]

        for i, text in enumerate(headers):
            h = QLabel(text)
            h.setStyleSheet("font-weight: 600; font-size: 20px;")
            h.setFixedWidth(self.col_widths[i])
            self.grid.addWidget(h, 0, i)

        self.scrollf.setWidget(self.container)

        self.continue_listen_btn = QPushButton(t("continue_listen", self.language))
        self.continue_listen_btn.setFixedSize(250, 80)
        self.continue_listen_btn.setStyleSheet("font-weight: 600;")
        self.continue_listen_btn.clicked.connect(self.show_listening_page)
        self.continue_listen_btn.hide()

        self.open_excel_btn = QPushButton(t("open_excel_btn", self.language))
        self.open_excel_btn.setFixedSize(250, 80)
        self.open_excel_btn.setStyleSheet("font-weight: 600;")
        self.open_excel_btn.clicked.connect(self.open_excel_file)
        self.open_excel_btn.hide()

        btn_row = QHBoxLayout()
        btn_row.setAlignment(Qt.AlignLeft)
        btn_row.addWidget(self.stop_btn)
        btn_row.addWidget(self.new_extract_btn)
        btn_row.addWidget(self.continue_listen_btn)
        btn_row.addWidget(self.open_excel_btn)

        content_layout.addWidget(self.extheader)
        content_layout.addSpacing(5)
        content_layout.addWidget(self.extbox)
        content_layout.addSpacing(5)
        content_layout.addLayout(btn_row)
        content_layout.addWidget(self.caption5)
        content_layout.addWidget(self.scrollf)

        return content

    def create_listening_page(self):
        content = QWidget()
        content_layout = QVBoxLayout(content)
        content_layout.setAlignment(Qt.AlignTop | Qt.AlignLeft)

        self.lheader = QLabel(t("listening_header", self.language))
        self.lheader.setStyleSheet("font: bold 50px;")

        self.lbox = QFrame()
        self.lbox.setWindowFlags(Qt.FramelessWindowHint)
        self.lbox.setAttribute(Qt.WA_TranslucentBackground)
        self.lbox.setFixedSize(600, 70)
        self.lbox.setStyleSheet("background-color: rgba(255, 165, 0, 100); margin-left: -10px;")

        self.statusl = QLabel(t("listening_paused", self.language))
        self.statusl.setStyleSheet("font: 600 20px; padding: 15px;")

        lbox_layout = QVBoxLayout(self.lbox)  # use lbox not extbox
        lbox_layout.addWidget(self.statusl)   # use statusl not status

        self.listen_toggle_btn = QPushButton(t("resume_listen", self.language))
        self.listen_toggle_btn.clicked.connect(self.toggle_listening)
        self.listen_toggle_btn.setFixedSize(250, 80)

        self.lcount = QLabel("")
        self.lcount.setStyleSheet("font: 600 20px;")

        self.lscrollf = QScrollArea()         # separate scroll area from extraction page
        self.lscrollf.setFixedSize(1000, 500)
        self.lscrollf.setWidgetResizable(True)
        self.lscrollf.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        self.lscrollf.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOn)

        self.lcontainer = QWidget()           # separate container
        self.lcontainer.setMinimumWidth(1510)
        self.lrow = 1                         # separate row counter

        self.lgrid = QGridLayout(self.lcontainer)  # separate grid
        self.lgrid.setContentsMargins(10, 10, 10, 10)
        self.lgrid.setHorizontalSpacing(30)
        self.lgrid.setVerticalSpacing(15)
        self.lgrid.setRowStretch(0, 0)
        self.lgrid.setAlignment(Qt.AlignTop)

        headers = [
            t("sender", self.language), t("subject", self.language),
            t("date", self.language), "MV", "DWT/Built",
            t("location", self.language), t("open_date", self.language),
            t("zone", self.language),
        ]

        for i, text in enumerate(headers):
            h = QLabel(text)
            h.setStyleSheet("font-weight: 600; font-size: 20px;")
            h.setFixedWidth(self.col_widths[i])
            self.lgrid.addWidget(h, 0, i)

        self.lscrollf.setWidget(self.lcontainer)

        content_layout.addWidget(self.lheader)    # use lheader
        content_layout.addSpacing(5)
        content_layout.addWidget(self.lbox)       # use lbox
        content_layout.addSpacing(5)
        content_layout.addWidget(self.listen_toggle_btn)
        content_layout.addWidget(self.lcount)
        content_layout.addWidget(self.lscrollf)   # use lscrollf

        return content


    def on_extract_sidebar_clicked(self):
        if self.page_main is not None:
            self.switch_page(self.page_main)
        else:
            self.switch_page(self.page_home)

    def new_extraction(self):
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
            self.caption.setStyleSheet("font: 600 20px;")
            self.caption2.setText(self.email_address)
            self.caption2.setStyleSheet("font: normal 20px;")
        else:
            self.caption.setText(t("no_email", self.language))
            self.caption.setStyleSheet("font: 600 20px;")
            self.caption2.setText("")

        if getattr(self, "folder", ""):
            self.caption3.setText(t("folder_extracting", self.language))
            self.caption3.setStyleSheet("font: 600 20px;")
         
            self.caption4.setText(self.folder)
            self.caption4.setStyleSheet("font: normal 20px;")
        else:
            self.caption3.setText(t("no_folder", self.language))
            self.caption3.setStyleSheet("font: 600 20px;")
            self.caption4.setText("")

        if getattr(self, "excel", ""):
            self.captione1.setText(t("excel_extracting", self.language))
            self.captione1.setStyleSheet("font: 600 20px;")
            self.captione2.setText(self.excel)
            self.captione2.setStyleSheet("font: normal 20px;")
        else:
            self.captione1.setText(t("excel_extracting", self.language))
            self.captione1.setStyleSheet("font: 600 20px;")
            self.captione2.setText(resolve_excel_path(""))
            self.captione2.setStyleSheet("font: normal 20px; color: grey;")

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
    
    def refresh_duplicates(self):
        delete_duplicates()
        self.refresh_btn.setText(t("cleared", self.language))
        QTimer.singleShot(2000, lambda: self.refresh_btn.setText(t("clear_duplicates_btn", self.language)))

    def date_entered(self, text, dmy):
        if not text.isdigit():
            return
        if dmy == "d":
            self.day = "0" + text if len(text) == 1 else text
        elif dmy == "m":
            self.month = "0" + text if len(text) == 1 else text
        else:
            self.year = text

        if getattr(self, "day", "") and getattr(self, "month", "") and getattr(self, "year", ""):
            self.date = self.year + "-" + self.month + "-" + self.day

    def time_entered(self, text, hm):
        if hm != "ampm" and not text.isdigit():
            return
        if hm == "h":
            self.hours = "0" + text if len(text) == 1 else text
        elif hm == "m":
            self.minutes = "0" + text if len(text) == 1 else text
        else:
            self.ampm = text.upper()

        if getattr(self, "hours", "") and getattr(self, "minutes", "") and getattr(self, "ampm", ""):
            self.time = self.hours + ":" + self.minutes + " " + self.ampm

    def handle_extract(self):
        self.btn.setEnabled(False)

        v, msg, dt = validate(self.date, self.time, self.email_address, self.folder, self.excel, outlook, self.language)
        if not v:
            self.error.setText(msg)
            self.btn.setEnabled(True)
            return

        self.extracting_running = True
        self._current_excel = resolve_excel_path(self.excel)
        self.show_main_page()

        generator = night_extraction(dt, self.email_address, self.folder, self._current_excel, csv_dict)

        self.thread = QThread()
        self.worker = ExtractWorker(generator)
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
            self.extbox.setStyleSheet("background-color: rgba(255, 165, 0, 100); margin-left: -10px;")
            return
        if email_data.get("type") == "excel_unlocked":
            self.status.setText(t("extraction_running", self.language))
            self.extbox.setStyleSheet("background-color: rgba(0, 255, 0, 100); margin-left: -10px;")
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

            labels = [
                QLabel(sender), QLabel(truncate(subject)), QLabel(received_time),
                QLabel(mv), QLabel(dwt_built), QLabel(location), QLabel(date), QLabel(zone)
            ]

            for i, label in enumerate(labels):
                label.setStyleSheet("font-size: 18px; padding-right: 15px;")
                label.setWordWrap(True)
                label.setFixedWidth(self.col_widths[i])
                self.grid.addWidget(label, self.row, i)

            self.row += 1
            self.scrollf.verticalScrollBar().setValue(self.scrollf.verticalScrollBar().maximum())

            self.emails_processed += 1
            if self.emails_processed // 5000 > self._last_donation_milestone:
                self._last_donation_milestone = self.emails_processed // 5000
                save_config(load_config() | {"emails_processed": self.emails_processed})
                self.show_donation_popup()
        
        except Exception as e:
            print(f"Error adding email to table: {e}")

    def add_to_listening_table(self, email_data):

        if email_data.get("type") == "excel_locked":
            self.statusl.setText("Waiting for Excel file to close...")
            self.lbox.setStyleSheet("background-color: rgba(255, 165, 0, 100); margin-left: -10px;")
            return
        if email_data.get("type") == "excel_unlocked":
            self.statusl.setText(t("listening_running", self.language))
            self.lbox.setStyleSheet("background-color: rgba(0, 255, 0, 100); margin-left: -10px;")
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
                QLabel(sender), QLabel(truncate(subject)), QLabel(received_time),
                QLabel(mv), QLabel(dwt_built), QLabel(location), QLabel(date), QLabel(zone)
            ]

            for i, label in enumerate(labels):
                label.setStyleSheet("font-size: 18px; padding-right: 15px;")
                label.setWordWrap(True)
                label.setFixedWidth(self.col_widths[i])
                self.lgrid.addWidget(label, self.lrow, i)

            self.lrow += 1
            self.lscrollf.verticalScrollBar().setValue(self.lscrollf.verticalScrollBar().maximum())

            self.emails_processed += 1
            if self.emails_processed // 5000 > self._last_donation_milestone:
                self._last_donation_milestone = self.emails_processed // 5000
                save_config(load_config() | {"emails_processed": self.emails_processed})
                self.show_donation_popup()
        
        except Exception as e:
            print(f"Error adding email to table: {e}")

    def on_extraction_done(self):
        error_key = getattr(self.worker, "api_error_key", None)

        if error_key:
            self.extheader.setText(t(error_key, self.language))
            self.extheader.setStyleSheet("font: bold 25px; color: red;")
            self.status.setText(t("extraction_stopped", self.language))
        elif self.row == 1:
            self.extheader.setText(t("extraction_complete_none", self.language))
            self.extheader.setStyleSheet("font: bold 25px;")
            no_results = QLabel(t("no_vessels", self.language))
            no_results.setStyleSheet("font: normal 20px; color: grey;")
            self.grid.addWidget(no_results, 1, 0, 1, 3)
            self.status.setText(t("extraction_stopped", self.language))
        else:
            self.extheader.setText(t("extraction_complete", self.language))
            self.extheader.setStyleSheet("font: bold 25px;")
            self.status.setText(t("extraction_stopped", self.language))
            self.continue_listen_btn.show()

        self.extbox.setStyleSheet("background-color: rgba(255, 0, 0, 100); margin-left: -10px;")
        self.btn.setEnabled(True)
        self.extracting_running = False
        self.stop_btn.hide()
        self.new_extract_btn.show()
        self.open_excel_btn.show()

        self.tray.showMessage(
            "Extraction complete",
            f"{self.row - 1} vessels extracted",
            QIcon(resource_path("icon.png")),
            3000
        )

        for field in (self.input_day, self.input_month, self.input_year,
                      self.input_hour, self.input_minute, self.input_ampm):
            field.clear()
        self.day = self.month = self.year = self.hours = self.minutes = self.ampm = ""
        self.date = None
        self.time = None

    def on_listen_done(self):
        error_key = getattr(self.listen_worker, "api_error_key", None)
        if error_key:
            self.lheader.setText(t(error_key, self.language))
            self.lheader.setStyleSheet("font: bold 50px; color: red;")
            self.statusl.setText(t(error_key, self.language))
            self.lbox.setStyleSheet("background-color: rgba(255, 0, 0, 100); margin-left: -10px;")
            self.listening_running = False
            self.listen_toggle_btn.setText(t("resume_listen", self.language))

    def toggle_listening(self):
        if self.listening_running:
            # pause
            if hasattr(self, "listen_worker") and self.listen_worker:
                self.listen_worker.stop()

            self.listening_running = False
            self.listen_toggle_btn.setText(t("resume_listen", self.language))
            self.statusl.setText(t("listening_paused", self.language))
            self.lbox.setStyleSheet("background-color: rgba(255, 165, 0, 100); margin-left: -10px;")
        else:
            # resume — wait for old thread to finish before creating a new one
            if hasattr(self, "listen_thread") and self.listen_thread and self.listen_thread.isRunning():
                self.listen_thread.wait(10000)
                if self.listen_thread and self.listen_thread.isRunning():
                    return
            self.handle_listen()
            self.listen_toggle_btn.setText(t("pause_listen", self.language))
            self.statusl.setText(t("listening_running", self.language))
            self.lbox.setStyleSheet("background-color: rgba(0, 255, 0, 100); margin-left: -10px;")

    def open_excel_file(self):
        path = getattr(self, '_current_excel', None) or resolve_excel_path(self.excel)
        QDesktopServices.openUrl(QUrl.fromLocalFile(path))

    def show_donation_popup(self):
        dialog = QDialog(self)
        dialog.setWindowTitle(t("donation_title", self.language))
        dialog.setFixedWidth(620)

        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(30, 30, 30, 30)
        layout.setSpacing(20)

        msg = QLabel(t("donation_message", self.language))
        msg.setWordWrap(True)
        msg.setStyleSheet("font: 16px;")
        layout.addWidget(msg)

        btn_row = QHBoxLayout()
        donate_btn = QPushButton(t("donation_btn", self.language))
        donate_btn.setFixedSize(200, 60)
        donate_btn.setStyleSheet("font-weight: 600;")
        donate_btn.clicked.connect(lambda: (QDesktopServices.openUrl(QUrl("https://ko-fi.com/jonathanfan")), dialog.accept()))

        close_btn = QPushButton(t("donation_close", self.language))
        close_btn.setFixedSize(200, 60)
        close_btn.clicked.connect(dialog.reject)

        btn_row.addWidget(donate_btn)
        btn_row.addSpacing(20)
        btn_row.addWidget(close_btn)
        layout.addLayout(btn_row)

        dialog.exec()

    def toggle_theme(self):
        config = load_config()
        current = config.get("theme", "dark")

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
            QApplication.instance().setStyleSheet("""
            QMainWindow { background-color: #080f1a; }
            QWidget { background-color: transparent; }
            QFrame { background-color: transparent; }
            QPushButton { 
                background-color: #0d1f35; 
                color: #22d3ee; 
                font: 600 18px;
            }
            QPushButton:hover {
                background-color: #0891b2;
                color: #080f1a;
            }
            QPushButton:disabled {
                background-color: #060d17;
                color: #2a4a5e;
            }
            QLineEdit { 
                background-color: #0d1f35; 
                color: #f0f9ff; 
                padding: 4px 8px;
            }
            QLineEdit:focus { border: 1px solid #0891b2; }
            QScrollArea { background-color: #080f1a;}
            QComboBox { 
                background-color: #0d1f35; 
                color: #22d3ee; 
                padding: 6px 12px;
                font-size: 14pt;
            }
            QLabel { color: #f0f9ff; }
        """)
        
            self.sidebar.setStyleSheet("""
                QFrame {
                    background-color: #0a1628;
                    border: none;
                    border-right: 1px solid #1a3a5c;
                }
            """)

            self.logo.setStyleSheet("""
                font: 800 16px;
                font-family: 'Syne';
                color: #f0f9ff;
                letter-spacing: 4px;
                background-color: #0a1628;
                border-bottom: 1px solid #1a3a5c;
                padding-left: 16px;
            """)

            sidebar_font = "Source Han Sans SC" if self.language == "中文" else "DM Mono"
            btn_qss = f"""
                QPushButton {{
                    background-color: transparent;
                    color: #7ca4c0;
                    font-family: '{sidebar_font}';
                    font-size: 13px;
                    font-weight: 500;
                    border: none;
                    border-left: 2px solid transparent;
                    text-align: left;
                    padding-left: 18px;
                }}
                QPushButton:hover {{
                    background-color: #0d1f35;
                    color: #22d3ee;
                    border-left: 2px solid #0891b2;
                }}
            """
        else:
            QApplication.instance().setStyleSheet("""
            QMainWindow { background-color: #f0f9ff; }
            QWidget { background-color: transparent; }
            QFrame { background-color: transparent; }
            QPushButton { 
                background-color: #e0f2fe; 
                color: #0891b2; 
                font: 600 18px;
            }
            QPushButton:hover {
                background-color: #0891b2;
                color: #f0f9ff;
            }
            QPushButton:disabled {
                background-color: #c5dce8;
                color: #7aaabb;
            }
            QLineEdit { 
                background-color: #ffffff; 
                color: #0a1628; 
                border: 1px solid #bae6fd;
                padding: 4px 8px;
            }
            QLineEdit:focus { border: 1px solid #0891b2; }
            QScrollArea { background-color: #f0f9ff; }
            QComboBox { 
                background-color: #e0f2fe;
                color: #0d1f35; 
                padding: 6px 12px;
                font-size: 14pt;
            }
            QLabel { color: #0a1628; }
            """)

            self.sidebar.setStyleSheet("""
                QFrame {
                    background-color: #e0f2fe;
                    border: none;
                    border-right: 1px solid #bae6fd;
                }
            """)

            self.logo.setStyleSheet("""
                font: 800 16px;
                font-family: 'Syne';
                color: #0d1f35;
                letter-spacing: 4px;
                background-color: #e0f2fe;
                border-bottom: 1px solid #0d1f35;
                padding-left: 16px;
            """)

            sidebar_font = "Source Han Sans SC" if self.language == "中文" else "DM Mono"
            btn_qss = f"""
                QPushButton {{
                    background-color: transparent;
                    color: #0891b2;
                    font-family: '{sidebar_font}';
                    font-size: 13px;
                    font-weight: 500;
                    border: none;
                    border-left: 2px solid transparent;
                    text-align: left;
                    padding-left: 18px;
                }}
                QPushButton:hover {{
                    background-color: #bae6fd;
                    color: #0a1628;
                    border-left: 2px solid #0891b2;
                }}
            """

        for btn in [self.extract_sidebar_btn, self.filtering_sidebar_btn, self.settings_sidebar_btn]:
            btn.setStyleSheet(btn_qss)

        self.main_widget.set_theme(theme)
        self.pages.set_theme(theme)
        if hasattr(self, '_settings_content') and self._settings_content:
            self._settings_content.set_theme(theme)
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

        # update sidebar buttons
        self.extract_sidebar_btn.setText(t("extract", self.language))
        self.filtering_sidebar_btn.setText(t("filtering", self.language))
        self.settings_sidebar_btn.setText(t("settings", self.language))

        self.pages.setCurrentWidget(self.page_settings)

        QApplication.setFont(get_font(self.language))
        current_theme = load_config().get("theme", "dark")
        self.apply_theme(current_theme)


if __name__ == "__main__":
    load_existing_vessels()
    load_email_ids()
    config = load_config()
    app = QApplication(sys.argv)
    app.setWindowIcon(QIcon(resource_path("icon.png")))
    window = MainWindow()
    window.apply_theme(config.get("theme", "dark"))
    window.show()
    sys.exit(app.exec())
