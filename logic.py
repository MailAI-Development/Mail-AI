import requests
import openpyxl
import os
import re
import sys
import win32com.client
import pythoncom
from datetime import datetime
import pytz
import time
import csv
import json
from math import log10, floor
import threading
import logging

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(os.path.join(os.path.expanduser("~"), "mailai.log"), encoding="utf-8"),
        logging.StreamHandler(),
    ]
)
logger = logging.getLogger(__name__)


class ProxyAuthError(Exception):
    pass

class ProxyAPIError(Exception):
    pass

def format_received_time(dt):
    s = dt.isoformat(timespec='minutes').replace('T', ' ')
    if len(s) > 16 and s[-6] in ('+', '-'):
        s = s[:-6] + ' ' + s[-6:]
    return s

def resource_path(relative_path):
    if getattr(sys, "frozen", False):
        base_path = getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

def data_path(relative_path):
    if getattr(sys, "frozen", False):
        base_path = os.path.dirname(sys.executable)
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

config_file = data_path("config.json")
duplicates_file = data_path("duplicates.json")
email_ids_file = data_path("email_ids.json")
custom_zones_file = data_path("custom_zones.json")
existing_vessels = {}
email_ids = set()
_email_ids_lock = threading.Lock()





TRANSLATIONS = {
    "English": {
        "welcome": "Welcome to the app",
        "extract_something": "Extract something",
        "extract": "Extract",
        "listen": "Listen for emails",
        "filtering": "Filtering",
        "settings": "Settings",
        "filtering_settings": "Filtering settings",
        "email_caption": "Email address you want to extract data from:",
        "folder_caption": "Folder you want to extract data from:",
        "excel_caption": "Path to Excel spreadsheet (where data will be extracted to):",
        "clear_duplicates_caption": "Delete all existing duplicates stored",
        "clear_duplicates_btn": "Clear duplicates",
        "cleared": "✓ Duplicates cleared",
        "theme": "Theme:",
        "switch_light": "Switch to Light Mode",
        "switch_dark": "Switch to Dark Mode",
        "language": "Language:",
        "extract_page_header": "Extract something",
        "date_caption": "Choose a date to extract emails from. If left empty, all emails received today will be extracted:",
        "time_caption": "Choose a time to extract emails from. If left empty, all emails received between midnight and now will be extracted:",
        "start_extracting": "Start extracting",
        "tooltip": "Enter your email/folder/Excel path in filtering settings first",
        "no_email": "No email address has been defined, go to filtering settings",
        "email_extracting": "Email address currently extracting from:",
        "no_folder": "No folder has been defined, go to filtering settings",
        "folder_extracting": "Folder currently extracting from:",
        "no_excel": "No excel path has been defined, go to filtering settings",
        "excel_extracting": "Excel spreadsheet path:",
        "current_extraction": "Current extraction",
        "extraction_running": "Extraction is running",
        "extraction_stopped": "Extraction complete.",
        "continue_listen": "Continue listening",
        "donation_title": "Support Mail AI",
        "donation_message": "Hi! I'm a 16-year-old independent student developer. I built Mail AI entirely on my own and am absorbing the AI API costs out of pocket. If this tool has been helpful to you, a small optional donation would mean the world to me — it helps keep this project alive. Thank you so much!",
        "donation_btn": "Donate",
        "donation_close": "Maybe later",
        "donate_optional": "Optional donation",
        "open_excel_btn": "Open spreadsheet",
        "excel_caption": "Path to Excel spreadsheet (leave empty to use default):",
        "extraction_complete_none": "Extraction complete. No results yielded.",
        "extraction_complete": "Extraction complete. Check Excel spreadsheet for results.",
        "stop_extracting": "Stop extracting",
        "new_extraction": "New extraction",
        "no_vessels": "No vessels found for the given date and time.",
        "vessels_extracted": "Vessels extracted:",
        "sender": "Sender",
        "subject": "Subject",
        "date": "Date",
        "location": "Location",
        "open_date": "Open Date",
        "build_year": "Built",
        "zone": "Zone",
        "listening_header": "Live listening",
        "listening_running": "Listening for emails...",
        "listening_paused": "Listening paused.",
        "pause_listen": "Pause listening",
        "resume_listen": "Resume listening",
        "listen_error": "Listening unable to run. Check filtering settings.",
        "outlook_not_running": "Check if Outlook is installed and running",
        "datetime_invalid": "Date/time is invalid! (check your format)",
        "excel_path_invalid": "Excel path is invalid! (check your format)",
        "folder_not_found": "Folder not found: ",
        "email_not_found": "Email address not found: ",
        "proxy_auth_error": "Service error. Please update the app or contact support.",
        "proxy_error_generic": "Extraction service error. Check your internet connection.",
        "setup_welcome_title": "Welcome to Mail AI",
        "setup_welcome_subtitle": "Let's get you set up in a few quick steps.",
        "setup_get_started": "Get started",
        "setup_email_title": "Your email address",
        "setup_email_desc": "Enter the Outlook email address you want to extract data from.",
        "setup_folder_title": "Outlook folder",
        "setup_folder_desc": "Enter the name of the Outlook folder to monitor for emails.",
        "setup_excel_title": "Excel spreadsheet",
        "setup_excel_desc": "Choose the Excel file where extracted data will be saved. Leave empty to use a default file in your Documents folder.",
        "setup_excel_browse": "Browse...",
        "setup_finish_title": "You're all set!",
        "setup_finish_desc": "Your configuration is saved. You can change these settings anytime from the Filtering page.",
        "setup_finish_btn": "Start using Mail AI",
        "setup_next": "Next",
        "setup_back": "Back",
        "setup_step": "Step",
        "custom_zones_header": "Custom Zone Mappings",
        "custom_zones_desc": "Add your own port-to-zone mappings to supplement the built-in WPI database:",
        "port_name_label": "Port Name:",
        "zone_label": "Zone:",
        "add_zone_btn": "Add Mapping",
        "zone_added": "Zone mapping added.",
        "zone_removed": "Zone mapping removed.",
        "zone_empty": "Please enter both a port name and zone.",
        "remove_zone_btn": "Remove",
        "no_custom_zones": "No custom zone mappings added yet.",
        "custom_zones_list": "Current custom mappings:",
    },
    "中文": {
        "welcome": "欢迎使用",
        "extract_something": "提取数据",
        "extract": "提取",
        "listen": "监听邮件",
        "filtering": "筛选",
        "settings": "设置",
        "filtering_settings": "筛选设置",
        "email_caption": "您想提取数据的邮箱地址：",
        "folder_caption": "您想提取数据的文件夹：",
        "excel_caption": "Excel表格路径（数据将被提取到此处）：",
        "clear_duplicates_caption": "删除所有已存储的重复项",
        "clear_duplicates_btn": "清除重复项",
        "cleared": "✓ 重复项已清除",
        "theme": "主题：",
        "switch_light": "切换到浅色模式",
        "switch_dark": "切换到深色模式",
        "language": "语言：",
        "extract_page_header": "提取数据",
        "date_caption": "选择提取邮件的日期。如果留空，将提取今天收到的所有邮件：",
        "time_caption": "选择提取邮件的时间。如果留空，将提取从午夜到现在收到的所有邮件：",
        "start_extracting": "开始提取",
        "tooltip": "请先在筛选设置中填写邮箱/文件夹/Excel路径",
        "no_email": "未定义邮箱地址，请前往筛选设置",
        "email_extracting": "当前提取的邮箱地址：",
        "no_folder": "未定义文件夹，请前往筛选设置",
        "folder_extracting": "当前提取的文件夹：",
        "no_excel": "未定义Excel路径，请前往筛选设置",
        "excel_extracting": "Excel表格路径：",
        "current_extraction": "当前提取",
        "extraction_running": "提取进行中",
        "extraction_stopped": "提取完成。",
        "continue_listen": "继续监听",
        "donation_title": "支持 Mail AI",
        "donation_message": "你好！我是一名16岁的独立学生开发者。Mail AI 完全由我独自开发，AI API 的费用由我自己承担。如果这个工具对你有帮助，任何一点小额捐赠对我来说都意义重大——这将帮助我继续维护这个项目。非常感谢！",
        "donation_btn": "捐赠",
        "donation_close": "也许以后",
        "donate_optional": "可选捐赠",
        "open_excel_btn": "打开表格",
        "excel_caption": "Excel表格路径（留空则使用默认位置）：",
        "extraction_complete_none": "提取完成，未找到结果。",
        "extraction_complete": "提取完成，请查看Excel表格。",
        "stop_extracting": "停止提取",
        "new_extraction": "新建提取",
        "no_vessels": "在指定日期和时间内未找到船只。",
        "vessels_extracted": "已提取船只：",
        "sender": "发件人",
        "subject": "主题",
        "date": "日期",
        "location": "位置",
        "open_date": "开放日期",
        "build_year": "建造年份",
        "zone": "区域",
        "listening_header": "实时监听",
        "listening_running": "正在监听邮件...",
        "listening_paused": "监听已暂停。",
        "pause_listen": "暂停监听",
        "resume_listen": "恢复监听",
        "listen_error": "无法开始监听，请检查筛选设置。",
        "outlook_not_running": "请检查Outlook是否已安装并运行",
        "datetime_invalid": "日期/时间无效！（请检查格式）",
        "excel_path_invalid": "Excel路径无效！（请检查格式）",
        "folder_not_found": "未找到文件夹：",
        "email_not_found": "未找到邮箱地址：",
        "proxy_auth_error": "服务错误，请更新应用或联系支持。",
        "proxy_error_generic": "提取服务错误，请检查您的网络连接。",
        "setup_welcome_title": "欢迎使用 Mail AI",
        "setup_welcome_subtitle": "让我们通过几个简单的步骤完成设置。",
        "setup_get_started": "开始设置",
        "setup_email_title": "您的邮箱地址",
        "setup_email_desc": "输入您想要提取数据的 Outlook 邮箱地址。",
        "setup_folder_title": "Outlook 文件夹",
        "setup_folder_desc": "输入要监控的 Outlook 文件夹名称。",
        "setup_excel_title": "Excel 表格",
        "setup_excel_desc": "选择用于保存提取数据的 Excel 文件。留空将在文档文件夹中使用默认文件。",
        "setup_excel_browse": "浏览...",
        "setup_finish_title": "设置完成！",
        "setup_finish_desc": "您的配置已保存。您可以随时在筛选页面中更改这些设置。",
        "setup_finish_btn": "开始使用 Mail AI",
        "setup_next": "下一步",
        "setup_back": "上一步",
        "setup_step": "步骤",
        "custom_zones_header": "自定义区域映射",
        "custom_zones_desc": "添加您自己的港口-区域映射以补充内置的WPI数据库：",
        "port_name_label": "港口名称：",
        "zone_label": "区域：",
        "add_zone_btn": "添加映射",
        "zone_added": "区域映射已添加。",
        "zone_removed": "区域映射已删除。",
        "zone_empty": "请输入港口名称和区域。",
        "remove_zone_btn": "删除",
        "no_custom_zones": "尚未添加自定义区域映射。",
        "custom_zones_list": "当前自定义映射：",
    }
}

def t(key, language="English"):
    return TRANSLATIONS.get(language, TRANSLATIONS["English"]).get(key, key)


keywords = [
    'MV', 'DWT', 'dwt', 'open', 'vessel position', 
    'bulk carrier', 'handy', 'supramax', 'ultramax', 'panamax', 'kamsarmax', 'ETA', 'ETD'
]


def load_csv_into_dict(csv_file):
    mapping = {}
    with open(csv_file, mode="r", encoding="latin-1") as f:
        reader = csv.reader(f)
        next(reader, None)  # skip header if present
        for row in reader:
            if len(row) >= 2:  # avoid malformed lines
                key = row[0].strip().upper()
                value = row[1].strip()
                if key not in mapping:
                    mapping[key] = []
                if value not in mapping[key]:  # avoid duplicates
                    mapping[key].append(value)
    return mapping

def load_custom_zones():
    if os.path.exists(custom_zones_file):
        try:
            with open(custom_zones_file, "r", encoding="utf-8") as f:
                content = f.read().strip()
                if not content:
                    return {}
                data = json.loads(content)
                if not isinstance(data, dict):
                    return {}
                return data
        except (json.JSONDecodeError, OSError):
            return {}
    return {}

def save_custom_zones(zones):
    with open(custom_zones_file, "w", encoding="utf-8") as f:
        json.dump(zones, f, indent=1)

def add_custom_zone(port_name, zone):
    zones = load_custom_zones()
    key = port_name.strip().upper()
    zone = zone.strip().upper()
    if key not in zones:
        zones[key] = []
    if zone not in zones[key]:
        zones[key].append(zone)
    save_custom_zones(zones)

def remove_custom_zone(port_name):
    zones = load_custom_zones()
    key = port_name.strip().upper()
    if key in zones:
        del zones[key]
        save_custom_zones(zones)

def get_custom_zones_list():
    zones = load_custom_zones()
    return [(port, ", ".join(zone_list)) for port, zone_list in sorted(zones.items())]

def merge_custom_zones(csv_mapping):
    custom = load_custom_zones()
    for key, zone_list in custom.items():
        if key not in csv_mapping:
            csv_mapping[key] = []
        for zone in zone_list:
            if zone not in csv_mapping[key]:
                csv_mapping[key].append(zone)
    return csv_mapping

def lookup_value(input_text, mapping):

    input_text = input_text.strip().upper()

    if input_text in mapping:
        return ", ".join(mapping[input_text])

    # Substring match
    substring_matches = [
        zone for key, zones in mapping.items() if input_text in key or key in input_text for zone in zones
    ]
    if substring_matches and len(set(substring_matches)) <= 3:
        zones = list(set(substring_matches))
        return ", ".join(zones)

    return "UNKNOWN"

def is_relevant_email(email_subject, email_body):
    for keyword in keywords:
        if re.search(r'\b' + re.escape(keyword) + r'\b', email_subject, re.IGNORECASE):
            return True
    # only check body if subject didn't match, to avoid scanning long emails
    first_100_lines = '\n'.join(email_body.splitlines()[:100])
    for keyword in keywords:
        if re.search(r'\b' + re.escape(keyword) + r'\b', first_100_lines, re.IGNORECASE):
            return True
    return False

def get_first_n_lines(email_body):
    noise_patterns = re.compile(
        r'(?i)^('
        # quoted replies
        r'(>\s*)'
        # legal / email footer boilerplate
        r'|.*confidential.*'
        r'|.*disclaimer.*'
        r'|.*unsubscribe.*'
        r'|.*this (e-?mail|message) (is |was )?(intended|sent|confidential).*'
        r'|.*best regards.*|.*kind regards.*|.*warm regards.*'
        r'|.*sincerely.*|.*yours faithfully.*'
        r'|\s*sent from .*'
        r'|\s*get outlook for .*'
        # separator lines
        r'|-{3,}.*forwarded.*-{3,}'
        r'|-{5,}$'
        r'|_{5,}$'
        r'|={5,}$'
        r'|\*{5,}$'
        # bunker specs
        r'|.*\bLSFO\b.*\bISO\b.*'
        r'|.*\bMGO\b.*\bISO\b.*'
        r'|.*\bLSMGO\b.*\bISO\b.*'
        r'|.*bimco.*'
        r'|.*marpol.*'
        r'|.*sulphur content.*'
        r'|.*bunker.*quality.*'
        r'|.*bunker.*spec.*'
        r'|.*mixing of bunkers.*'
        r'|.*charterers.*warrant.*'
        r'|.*iso\s*8217.*'
        # speed / consumption tables
        r'|.*\bspeed\b.*\bcons\b.*'
        r'|.*abt\s+\d+.*knot.*'
        r'|.*\bbeaufort\b.*'
        r'|.*douglas sea.*'
        r'|.*good weather condition.*'
        r'|.*extrapolation.*'
        r'|.*positive current.*'
        r'|.*clean bottom.*'
        r'|.*adverse current.*'
        r'|.*wave.*swell.*'
        r'|.*ballast.*deballast.*'
        r'|.*lsmgo for exchanging.*'
        r'|.*liberty of using.*'
        r'|.*maneuvering.*'
        r'|.*maintenance works.*'
        # vessel particulars / tech specs
        r'|.*\bimo no\b.*'
        r'|.*\bcall sign\b.*'
        r'|.*port of registry.*'
        r'|.*\bflag\b.*\bclass\b.*'
        r'|.*\bbuilder\b.*'
        r'|.*date of delivery.*'
        r'|.*\bloa\b.*\blbp\b.*'
        r'|.*\bloa\b.*\bbeam\b.*'
        r'|.*\bgross tonnage\b.*'
        r'|.*\bgrt\b.*\bnrt\b.*'
        r'|.*\bgt\b.*\bnt\b.*\d+.*'
        r'|.*panama tonnage.*'
        r'|.*main engine.*'
        r'|.*auxiliary engine.*'
        r'|.*hatch cover.*'
        r'|.*tank top strength.*'
        r'|.*\bco2 devices\b.*'
        r'|.*radio remote control.*'
        r'|.*\bport idle\b.*'
        r'|.*\bport working\b.*'
        # last cargo / port history
        r'|.*last (five|ten|five|3|5|10)\s*(cargo|port).*'
        r'|.*last \d+\s*(cargo|port).*'
        r'|.*bunker on delivery.*'
        r'|.*buker on delivery.*'
        r'|all (above|details).*'
        # loadline / draft tables (never in summaries)
        r'|.*\bSSW\b.*'
        r'|.*\b(WINTER|SUMMER|TROPICAL)\b.*\bDWT\b.*'
        r'|.*\bDWT\b.*\bTPC\b.*'
        r'|.*\bTPC\b.*\bDRAFT\b.*'
        r'|.*\bDRAFT\b.*\bTPC\b.*'
        # fuel consumption lines
        r'|.*\bVLSFO\b.*\bLSMGO\b.*'
        r'|.*\bLSFO\b.*\bLSMGO\b.*'
        r'|.*\bMDO\b.*\bLSFO\b.*'
        # speed lines using K (knots abbrev) alongside a fuel type — e.g. "ABT 13K ON LSFO ABT 25MT"
        r'|.*\b\d+\s*k\b.*\b(lsfo|vlsfo|lsmgo|mgo)\b.*'
        # vessel registry / class
        r'|.*\bimo\b.*\b\d{7}\b.*'
        r'|.*\bimo number\b.*'
        r'|.*\bclass\b.*\b(BV|DNV|LR|ABS|NK|GL|CCS|RINA|KR)\b.*'
        r'|.*\bbuilt\b.*\bflag\b.*'
        # hold / capacity lines
        r'|.*\bgrain\b.*\bbale\b.*\bcbm\b.*'
        r'|.*\ball details about\b.*'
        r')'
    )

    relevant = re.compile(
        r'(?i)('
        r'\bM[./]?V[./]?\b|\bDWT\b|\b\d+\s*K\b|\bOPEN\b'
        r'|\bO\s*/\s*A\b|\bO\.A\.?\b|\bOA\b'
        r'|\b(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)\b'
        r'|\b(JANUARY|FEBRUARY|MARCH|APRIL|MAY|JUNE|JULY|AUGUST|SEPTEMBER|OCTOBER|NOVEMBER|DECEMBER)\b'
        r'|\d{1,2}[.)]\s'
        r')'
    )

    # Clean lines, tracking original indices
    cleaned = []
    for line in email_body.splitlines():
        stripped = line.strip()
        if not stripped or noise_patterns.match(stripped):
            cleaned.append(None)
        else:
            cleaned.append(stripped)

    # Mark relevant line indices, then expand by 1 context line each side
    relevant_indices = set()
    for i, line in enumerate(cleaned):
        if line and relevant.search(line):
            relevant_indices.add(i)

    context_indices = set()
    for i in relevant_indices:
        for j in (i - 1, i, i + 1):
            if 0 <= j < len(cleaned) and cleaned[j]:
                context_indices.add(j)

    result = [cleaned[i] for i in sorted(context_indices)]

    mv_name_re = re.compile(
        r'\bM\.?V\.?\b\s*["\']?\s*([A-Z][A-Z\s\d\-]{2,40}?)(?=["\']?\s*(?:\(|\bDWT\b|\bOPEN\b|\b\d{5}\b|,|$))',
        re.IGNORECASE
    )
    has_summary_context = re.compile(
        r'(?i)(\bOPEN\b|\bO\s*/\s*A\b|\bARD\b'
        r'|\b(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)\b'
        r'|\b(JANUARY|FEBRUARY|MARCH|APRIL|MAY|JUNE|JULY|AUGUST|SEPTEMBER|OCTOBER|NOVEMBER|DECEMBER)\b)'
    )
    seen_vessels = set()
    deduped = []
    for line in result:
        m = mv_name_re.search(line)
        if m:
            name_key = ' '.join(m.group(1).strip().upper().split())
            if name_key in seen_vessels and not has_summary_context.search(line):
                continue  # repeated vessel name with no new open/date info
            seen_vessels.add(name_key)
        deduped.append(line)
    result = deduped

    if not result:
        return None

    return '\n'.join(result)

def extract_details_from_email(preprocessed_body, csv_dict):
    payload = {
        "messages": [
            {"role": "system", "content":
            "You are a data extraction tool. Output ONLY the requested fields in the exact format shown."
            "No explanations, no notes, no extra text. If a value is unknown, write None."
            },
            {"role": "user", "content":
            f"Extract the following details for each vessel mentioned in the shipbroking email:\n\n"
            f"1. MV (Motor Vessel): vessel name, sometimes prefixed with MV but may be numbered with no MV, never include DWT — e.g. MV OCEAN STAR\n"
            f"2. Deadweight (DWT): DWT in K to 2 significant figures — e.g. 70K, 58K. None if unknown\n"
            f"3. Build Year: 4-digit year the vessel was built — often written alongside DWT as DWT/YEAR e.g. 57K/2012 but may be elsewhere. None if unknown\n"
            f"4. Vessel Open Location: the port or region where the vessel becomes available. Port name only in capitals, strip country and prefixes like OPEN/AT/IN/EX. May be marked with O/A or phrased as 'Open <location>', 'EX SHIPYARD <location>', or 'delivery <location>'. If only a region code is given (e.g. CJK, ECI, WCI, NOPAC), return that. None if unknown\n"
            f"5. Vessel Open Date: the date the vessel becomes available. Date only, no year — e.g. 10 OCT, 20-22 NOV, EARLY OCT. May be phrased as open date, O/A, ARD, ETA, 'Expected time of delivery', or 'delivery date'. None if unknown\n"
            f"Format:\nMV: [MV name]\nDeadweight: [deadweight]\nBuild Year: [build year]\nVessel Open Location: [vessel open location]\nVessel Open Date: [vessel open date]\n"
            f"Repeat for each vessel mentioned in the email.\n"
            f"A vessel's details may span multiple lines.\n"
            f"Return None for any field that is missing.\n"
            f"Before moving on to the next vessel, separate each vessel with '---'.\n"
            f"Email:\n{preprocessed_body}\n\n"},
        ]
    }
    headers = {"X-App-Token": _APP_TOKEN}
    last_error = None
    for attempt in range(3):
        try:
            resp = requests.post(PROXY_URL, json=payload, headers=headers, timeout=60)
            if resp.status_code == 401:
                raise ProxyAuthError("Unauthorized")
            resp.raise_for_status()
            data = resp.json()
            break
        except ProxyAuthError:
            raise
        except requests.exceptions.RequestException as e:
            last_error = e
            logger.warning(f"API request failed (attempt {attempt + 1}/3): {e}")
            if attempt < 2:
                time.sleep(2 ** attempt)
    else:
        raise ProxyAPIError(str(last_error))

    details = data["content"].strip()

    vessels = details.split('---')
    extracted_vessels = []

    for vessel in vessels:
        mv = re.search(r'MV:\s*(.*)|MV\s*(.*)', vessel)
        deadweight = re.search(r'Deadweight:\s*(.*)', vessel)
        build_year = re.search(r'Build Year:\s*(.*)', vessel)
        location = re.search(r'Vessel Open Location:\s*(.*)', vessel)
        date_of_arrival = re.search(r'Vessel Open Date:\s*(.*)', vessel)

        def clean(value):
            if value is None:
                return None
            stripped = value.strip()
            return None if stripped.lower() == 'none' else stripped.upper()

        def clean_year(value):
            if value is None:
                return None
            stripped = value.strip()
            if stripped.lower() == 'none':
                return None
            match = re.match(r'(\d{4})', stripped)
            return match.group(1) if match else None

        vessel_data = {
            'MV': ensure_mv_prefix(clean(mv.group(1) if mv and mv.group(1) else (mv.group(2) if mv and mv.group(2) else None))),
            'Deadweight': normalize_dwt(clean(deadweight.group(1) if deadweight else None)),
            'Build Year': clean_year(build_year.group(1) if build_year else None),
            'Vessel Open Location': clean(location.group(1) if location else None),
            'Vessel Open Date': validate_date(clean(date_of_arrival.group(1) if date_of_arrival else None)),
        }

        if vessel_data['Vessel Open Location'] is None:
            vessel_data['Zone'] = None
        else:
            zone = lookup_value(vessel_data['Vessel Open Location'], csv_dict)
            vessel_data['Zone'] = zone

        extracted_vessels.append(vessel_data)
    return extracted_vessels



def resolve_excel_path(path):
    if not path or not path.strip():
        docs = os.path.join(os.path.expanduser("~"), "Documents")
        os.makedirs(docs, exist_ok=True)
        return os.path.join(docs, "Vessel_Data_Extraction.xlsx")
    path = os.path.normpath(path)
    if not path.lower().endswith(".xlsx"):
        path += ".xlsx"
    return path

def load_config():
    if os.path.exists(config_file):
        with open(config_file, "r", encoding="utf-8") as f:
            config = json.load(f)
        return config
    return {}

def save_config(config):
    with open(config_file, "w", encoding="utf-8") as f:
        json.dump(config, f, indent=1)

def get_config_value(key, prompt_text, change):
    config = load_config()

    if key in config and config[key] and change == False:
        logger.debug(f"Found saved {key}: {config[key]}")
        return config[key]

    value = input(prompt_text + ": ")
    config[key] = value
    save_config(config)
    return value
    

def load_email_ids():
    global email_ids
    with _email_ids_lock:
        if os.path.exists(email_ids_file):
            try:
                with open(email_ids_file, "r", encoding="utf-8") as f:
                    content = f.read().strip()
                    if not content:
                        return
                    email_ids = set(json.loads(content))
            except (json.JSONDecodeError, OSError):
                email_ids = set()


def save_email_ids():
    with _email_ids_lock:
        ids_list = list(email_ids)
        if len(ids_list) > 5000:
            ids_list = ids_list[-5000:]
            email_ids.clear()
            email_ids.update(ids_list)
    with open(email_ids_file, "w", encoding="utf-8") as f:
        json.dump(ids_list, f)



def validate(date, time_str, email_address, folder, excel_path, outlook, language="English"):
    if outlook is None:
        return False, t("outlook_not_running", language), None
    try:
        if not date:
            date = str(datetime.now(pytz.UTC).date())
        if not time_str:
            time_str = "12:00 AM"
        sp_datetime = datetime.strptime(f"{date} {time_str}", '%Y-%m-%d %I:%M %p')

        specific_datetime = sp_datetime.replace(tzinfo=pytz.UTC)
        if specific_datetime > datetime.now(pytz.UTC):
            return False, t("datetime_invalid", language), None

    except ValueError:
        return False, t("datetime_invalid", language), None

    if excel_path and not excel_path.endswith(".xlsx"):
        return False, t("excel_path_invalid", language), None

    for account in outlook.Folders:
        if account.Name.upper() == email_address.upper():
            for subfolder in account.Folders:
                if subfolder.Name.upper() == folder.upper():
                    folder = subfolder
                    return True, "", specific_datetime
            else:
                return False, t("folder_not_found", language) + folder, None
    else:
        return False, t("email_not_found", language) + email_address, None
    

def is_excel_open(file_path):
    if not os.path.exists(file_path):
        return False
    try:
        with open(file_path, 'a+b'):
            return False
    except IOError:
        return True


def append_data_excel(file_path, data, specific_datetime, listening):
    headers = ['Sender', 'Subject', 'Received Time', 'MV', 'DWT/Built', 'Vessel Open Location', 'Vessel Open Date', 'Zone']

    today = datetime.now().date()
    sheet_name = today.strftime("%d %b %Y")

    if not os.path.exists(file_path):
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = sheet_name
        sheet.append(headers)
        sheet.auto_filter.ref = "A1:H1"
        workbook.save(file_path)
        workbook = openpyxl.load_workbook(file_path)
        sheet = workbook[sheet_name]
    else:
        workbook = openpyxl.load_workbook(file_path)
        if sheet_name in workbook.sheetnames:
            sheet = workbook[sheet_name]
        else:
            sheet = workbook.create_sheet(title=sheet_name)
            sheet.append(headers)
            sheet.auto_filter.ref = "A1:H1"

    if listening and specific_datetime:
        sheet.append([f"Emails extracted from {specific_datetime} below:"])
        sheet.append(headers)
        sheet.auto_filter.ref = f"A{sheet.max_row}:H{sheet.max_row}"

    for entry in data:
        subject = entry.get('Subject', '') or ''
        if len(subject) > 50:
            subject = subject[:50] + '...'
        dwt = entry.get('Deadweight', '') or ''
        year = entry.get('Build Year', '') or ''
        dwt_built = f"{dwt}/{year}" if dwt and year else (dwt or year or '')
        row = [
            entry.get('Sender', ''),
            subject,
            entry.get('Received Time', ''),
            entry.get('MV', ''),
            dwt_built,
            entry.get('Vessel Open Location', ''),
            entry.get('Vessel Open Date', ''),
            entry.get('Zone', '')
        ]
        sheet.append(row)

    workbook.save(file_path)
    workbook.close()


def append_error_message(file_path, sender_email, email_subject):
    while is_excel_open(file_path):
        time.sleep(2)

    # Check if the Excel file exists
    if not os.path.exists(file_path):
        workbook = openpyxl.Workbook()
        sheet = workbook.active
    else:
        # Open the existing workbook
        workbook = openpyxl.load_workbook(file_path)
        sheet = workbook.active

    # Prepare the message to append
    error_message = f"Data not found: Email sender: {sender_email}, Email subject: {email_subject}"

    # Append the error message in a new row
    sheet.append([error_message])

    workbook.save(file_path)
    workbook.close()




def filter_data(vessels):
    return [v for v in vessels if v.get('MV') is not None]

def load_duplicates():
    if os.path.exists(duplicates_file):
        try:
            with open(duplicates_file, "r", encoding="utf-8") as f:
                content = f.read().strip()
                if not content:
                    return {}
                return json.loads(content)
        except json.JSONDecodeError:
            return {}
    return {}

def load_existing_vessels():
    global existing_vessels
    existing_vessels = load_duplicates()

def save_duplicates(duplicates):
    with open(duplicates_file, "w", encoding="utf-8") as f:
        json.dump(duplicates, f, indent=1)


def normalise_mv(name):
    if not name:
        return None
    name = re.sub(r'^(M\.?V\.?|M/V)\s*', '', str(name), flags=re.IGNORECASE).strip()
    name = re.sub(r'\s+\d+K.*$', '', name, flags=re.IGNORECASE).strip()
    return re.sub(r'\s+', ' ', name).upper()

def normalise_date(date):
    if not date:
        return None
    replacements = {
        'JANUARY': 'JAN', 'FEBRUARY': 'FEB', 'MARCH': 'MAR',
        'APRIL': 'APR', 'JUNE': 'JUN', 'JULY': 'JUL',
        'AUGUST': 'AUG', 'SEPTEMBER': 'SEP', 'OCTOBER': 'OCT',
        'NOVEMBER': 'NOV', 'DECEMBER': 'DEC'
    }
    for full, short in replacements.items():
        date = re.sub(rf'\b{full}\b', short, date, flags=re.IGNORECASE)
    date = re.sub(r'(\d+)(ST|ND|RD|TH)\b', r'\1', date, flags=re.IGNORECASE)
    date = re.sub(r'(\d+)\s+TO\s+(\d+)', r'\1-\2', date, flags=re.IGNORECASE)
    date = re.sub(r'\b(19|20)\d{2}\b', '', date)
    return date.strip().upper()

def ensure_mv_prefix(name):
    if not name:
        return None
    if not re.match(r'^MV\s+', name):
        return f"MV {name}"
    return name

def normalize_dwt(dwt_str):
    if not dwt_str:
        return None
    match = re.match(r'([\d,\.]+)\s*([KkMmTt]*)', dwt_str.strip())
    if not match:
        return None
    num = float(match.group(1).replace(',', ''))
    unit = match.group(2).upper()
    if 'MT' in unit or (num > 999 and 'K' in unit) or (num > 999 and not unit):
        num = num / 1000
    if num == 0:
        return None
    magnitude = floor(log10(abs(num)))
    rounded = round(num, -int(magnitude) + 1)
    return f"{int(rounded)}K"

def validate_date(date_str):
    if not date_str:
        return None
    date_str = normalise_date(date_str)
    months = r'(?:JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)'
    valid_patterns = [
        rf'^\d{{1,2}}\s+{months}$',                                                      # 10 OCT
        rf'^\d{{1,2}}-\d{{1,2}}\s+{months}$',                                            # 20-22 NOV
        rf'^\d{{1,2}}/\d{{1,2}}\s+{months}$',                                            # 10/11 OCT
        rf'^\d{{1,2}}\s+{months}\s*[-/]\s*\d{{1,2}}\s+{months}$',                       # 30 JUN - 2 JUL / 30 JUN / 2 JUL
        rf'^{months}\s+\d{{1,2}}$',                                                      # OCT 10
        rf'^(?:EARLY|MID|LATE|END)\s+{months}$',                                         # EARLY OCT
        rf'^(?:EARLY|MID|LATE|END)\s+{months}\s*/\s*(?:EARLY|MID|LATE|END)\s+{months}$', # LATE MAR / EARLY APR
    ]
    for pattern in valid_patterns:
        if re.match(pattern, date_str):
            return date_str
    return date_str

def is_valid_vessel(vessel):
    return (
        vessel.get('MV') not in (None, 'None') and
        vessel.get('Vessel Open Location') not in (None, 'None') and
        '->' not in str(vessel.get('Vessel Open Location', ''))
    )

def detect_duplicates(details):
    global existing_vessels
    vessels = []

    for vessel in details:
        if not is_valid_vessel(vessel):
            continue

        key = normalise_mv(vessel.get('MV'))
        if not key:
            continue

        date = normalise_date(vessel.get('Vessel Open Date'))

        if key not in existing_vessels:
            existing_vessels[key] = date
            vessels.append(vessel)
        else:
            if existing_vessels[key] != date:
                existing_vessels[key] = date
                vessels.append(vessel)

    if vessels:
        save_duplicates(existing_vessels)

    return vessels


def delete_duplicates():
    global existing_vessels
    existing_vessels = {}
    with open(duplicates_file, "w", encoding="utf-8") as f:
        json.dump({}, f)  # empty dict


def process_email(email_address,folder,excel_path,csv_dict,worker):

    global email_ids
    start_time = datetime.now()
    pythoncom.CoInitialize()
    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

        for account in outlook.Folders:
            if account.Name.upper() == email_address.upper():
                for subfolder in account.Folders:
                    if subfolder.Name.upper() == folder.upper():
                        folder = subfolder
                        break
                else:
                    return
                break
        else:
            return

        while True:
            if not worker.running:
                return

            messages = folder.Items
            messages.Sort("[ReceivedTime]", True)
            message = messages.GetFirst()

            while message:
                if not worker.running:
                    return

                with _email_ids_lock:
                    already_seen = message.EntryID in email_ids
                if not already_seen and hasattr(message, 'ReceivedTime'):
                    received_time = message.ReceivedTime
                    if received_time.replace(tzinfo=None) < start_time:
                        message = messages.GetNext()
                        continue

                    email_body = message.Body
                    email_subject = message.Subject
                    sender_email = message.SenderEmailAddress

                    if is_relevant_email(email_subject, email_body):
                        logger.info(f"Processing email from: {sender_email} with subject: {email_subject}")
                        preprocessed_body = get_first_n_lines(email_body)

                        if preprocessed_body:
                            try:
                                extracted_details = extract_details_from_email(preprocessed_body, csv_dict)
                            except ProxyAuthError:
                                yield {"type": "api_error", "error_key": "proxy_auth_error"}
                                return
                            except ProxyAPIError:
                                yield {"type": "api_error", "error_key": "proxy_error_generic"}
                                return
                            valid_vessels = filter_data(extracted_details)
                            vessels = detect_duplicates(valid_vessels)

                            if vessels:
                                for vessel in vessels:
                                    vessel['Sender'] = sender_email
                                    vessel['Subject'] = email_subject
                                    vessel['Received Time'] = format_received_time(received_time)
                                ves = len(vessels)
                                for vessel in vessels:
                                    yield {
                                        "sender": sender_email,
                                        "subject": email_subject,
                                        "received_time": format_received_time(received_time),
                                        "ves": ves,
                                        "vessel_data": vessel,
                                    }
                                if is_excel_open(excel_path):
                                    yield {"type": "excel_locked"}
                                    while is_excel_open(excel_path):
                                        if not worker.running:
                                            return
                                        time.sleep(2)
                                    yield {"type": "excel_unlocked"}
                                append_data_excel(excel_path, vessels, None, False)
                        else:
                            append_error_message(excel_path, sender_email, email_subject)

                    with _email_ids_lock:
                        email_ids.add(message.EntryID)
                    save_email_ids()

                message = messages.GetNext()

            for _ in range(10):
                if not worker.running:
                    return
                time.sleep(1)
    finally:
        pythoncom.CoUninitialize()




def night_extraction(specific_datetime,email_address,folder,excel_path,csv_dict):

    pythoncom.CoInitialize()
    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

        for account in outlook.Folders:
            if account.Name.upper() == email_address.upper():
                for subfolder in account.Folders:
                    if subfolder.Name.upper() == folder.upper():
                        folder = subfolder
                        break
                else:
                    return
                break
        else:
            return

        messages = folder.Items
        messages.Sort("[ReceivedTime]", True)
        message = messages.GetFirst()

        processed_emails = []
        ves = 0

        while message:
            if hasattr(message, 'ReceivedTime'):
                received_time = message.ReceivedTime

                if received_time.replace(tzinfo=None) <= specific_datetime.replace(tzinfo=None):
                    break
                if received_time.replace(tzinfo=None) > specific_datetime.replace(tzinfo=None):
                    email_body = message.Body
                    email_subject = message.Subject
                    sender_email = message.SenderEmailAddress

                    if is_relevant_email(email_subject, email_body):
                        logger.info(f"Processing email from: {sender_email} with subject: {email_subject}")
                        preprocessed_body = get_first_n_lines(email_body)
                        if preprocessed_body:
                            try:
                                extracted_details = extract_details_from_email(preprocessed_body, csv_dict)
                            except ProxyAuthError:
                                yield {"type": "api_error", "error_key": "proxy_auth_error"}
                                return
                            except ProxyAPIError:
                                yield {"type": "api_error", "error_key": "proxy_error_generic"}
                                return
                            valid_vessels = filter_data(extracted_details)
                            vessels = detect_duplicates(valid_vessels)

                            if vessels:
                                for vessel in vessels:
                                    vessel['Sender'] = sender_email
                                    vessel['Subject'] = email_subject
                                    vessel['Received Time'] = format_received_time(received_time)
                                processed_emails.extend(vessels)
                                ves += len(vessels)

                                for vessel in vessels:
                                    yield {
                                        "sender": sender_email,
                                        "subject": email_subject,
                                        "received_time": format_received_time(received_time),
                                        "ves": ves,
                                        "vessel_data": vessel,
                                    }
                        else:
                            append_error_message(excel_path, sender_email, email_subject)

            message = messages.GetNext()

        if processed_emails:
            if is_excel_open(excel_path):
                yield {"type": "excel_locked"}
                while is_excel_open(excel_path):
                    time.sleep(2)
                yield {"type": "excel_unlocked"}
            append_data_excel(excel_path, processed_emails, specific_datetime, True)
            return True
        else:
            return False
    finally:
        pythoncom.CoUninitialize()
