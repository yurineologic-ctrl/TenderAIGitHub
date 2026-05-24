import re
import time
from datetime import datetime
from pathlib import Path

import pandas as pd
from bs4 import BeautifulSoup
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

BASE_URL    = "https://www.asus.com"
CATALOG_URL = "https://www.asus.com/ua-ua/store/laptops/"
OUTPUT_PATH = Path.home() / "Desktop" / "NOUT_Asus.xlsx"
PAGE_TIMEOUT = 45_000   # 45 сек — ASUS CDN може вантажитись повільніше
LOAD_DETAIL_PAGES = True
MAX_ITEMS = 9999

COLUMNS = [
    "Seria", "Model", "Part Number",
    "CPU_Brand", "CPU_Model", "CPU_Cores", "CPU_Threads",
    "CPU_Base_Freq_GHz", "CPU_Max_Freq_GHz", "CPU_Cache_L3_MB",
    "RAM_Type", "RAM_Freq", "RAM_Size_GB", "RAM_Slots",
    "Nakopychuvach_SSD",
    "Display_Diagonal", "Display_Max_Resolution", "Display_Matrix_Type",
    "Display_Cover", "Display_Brightness_nits", "Display_Contrast",
    "Display_Response_Time", "Display_Refresh_Rate",
    "Video_Brand", "GPU_Type", "GPU_Model", "GPU_Memory_MB",
    "OS",
    "Battery_Capacity_Wh", "Camera_MP",
    "USB_TypeA", "USB_TypeC", "HDMI", "DisplayPort", "Ports",
    "Keyboard_Waterproof", "Keyboard_Ukrainian", "Keyboard_Pointing_Device",
    "Network_3G4G", "Bluetooth", "WiFi", "LAN_Mbps",
    "Certificates", "Warranty",
    "Merezha_WiFi", "Tsina_UAH", "URL",
]

SERIES_MAP = [
    "ROG", "TUF", "ZenBook", "VivoBook", "Vivobook", "ProArt",
    "ExpertBook", "Chromebook", "StudioBook", "OLED", "Flow",
    "Zenbook", "Expertbook",
]

# ── Compiled regex patterns (same as Lenovo parser) ──────────────────────────
RE_PART_NUMBER     = re.compile(r'\(([A-Z0-9_-]{6,})\)')
RE_PART_NUMBER_ALT = re.compile(r'\b[A-Z0-9]{8,}\b')
RE_DIAG            = re.compile(r'(\d+(?:[\.,]\d+)?)\s*(?:"|дюйм|дюйма|inch|in\b)', re.I)
RE_RESOLUTION      = re.compile(r'\d{3,4}[xх]\d{3,4}|fhd|wuxga|qhd|wqxga|uhd|4k', re.I)
RE_DISPLAY_TYPE    = re.compile(r'\b(ips|oled|tn|va|retina|nano\s*ips|amoled)\b', re.I)
RE_DISPLAY_COVER   = re.compile(r'(антиблік|антибл|матов|глян|antiglare|gloss|matte)', re.I)
RE_BRIGHTNESS      = re.compile(r'(?:яскравість|яскравость|brightness)[,:]?\s*(?:ніт|нит|nit)?\s*(\d+)|(\d+)\s*(?:ніт|нит|nit)', re.I)
RE_CONTRAST        = re.compile(r'контраст|contrast', re.I)
RE_RESPONSE_TIME   = re.compile(r'(\d+)\s*мс|(\d+)\s*ms', re.I)
RE_REFRESH_RATE    = re.compile(r'(?:частота\s*оновлення|refresh\s*rate)|(?:\d+\s*(?:гц|hz).*?(?:екран|дисплей|панель|wuxga|fhd|qhd|uhd|4k|oled|ips|tn))', re.I)
RE_CPU_FREQ        = re.compile(r'(\d+(?:[\.,]\d+)?)\s*ггц', re.I)
RE_CPU_FREQ_GHZ    = re.compile(r'(\d+(?:[\.,]\d+)?)\s*ghz', re.I)
RE_CPU_CORES       = re.compile(r'(\d+)\s*(?:ядер|ядра|cores?)', re.I)
RE_CPU_THREADS     = re.compile(r'(\d+)\s*(?:потоків|потоки|threads?)', re.I)
RE_L3_CACHE        = re.compile(r'\b(?:l3|кеш\s*l3|l3\s*кеш|cache\s*l3)\b', re.I)
RE_DDR             = re.compile(r'\b(?:LP)?DDR\d(?:X\b|\b)', re.I)
RE_RAM_FREQ        = re.compile(r'\b(\d{3,4})\s*(?:мгц|mhz)\b', re.I)
RE_RAM_SIZE        = re.compile(r'\b(\d+)\s*(?:гб|gb)\b', re.I)
RE_RAM_SLOT        = re.compile(r'(\d{1,2})\s*слот|[xх]\s*([12])(?=\s|[^0-9]|$)', re.I)
RE_STORAGE         = re.compile(r'\b(?:ssd|hdd|nvme|емкість|накоп)\b', re.I)
RE_GPU             = re.compile(r'\b(?:nvidia|geforce|rtx|gtx|radeon|intel graphics|arc|iris|gpu|вбудован|дискретн|відеокарт|video)\b', re.I)
RE_OS              = re.compile(r'\b(?:windows|linux|без ос|ubuntu|dos|freeDOS)\b', re.I)
RE_BATTERY         = re.compile(r'(?:ємність|енергетична\s*ємність|battery\s*capacity)[,:]?\s*(?:(\d+)\s*)?(?:вт\*год|вт·год|wh|watt.?hour)?\s*(\d+)?', re.I)
RE_BATTERY_WHR     = re.compile(r'(\d+(?:\.\d+)?)\s*(?:Вт·год|вт\*год|wh|вт\.год)', re.I)
RE_CAMERA          = re.compile(r'\b(?:камера|web-камера|веб-камера|hd|4k|ir\s*camera|webcam)|\d+\s*(?:мп|mp)\b', re.I)
RE_USB_A           = re.compile(r'\b(?:usb(?:\s*type-?a|\s*3\.?\d?|\s*2\.?\d?|\s*a)|usb-a|type-?a)\b', re.I)
RE_USB_C           = re.compile(r'\b(?:usb(?:\s*type-?c|\s*c)|usb-c|type-?c|thunderbolt)\b', re.I)
RE_HDMI            = re.compile(r'\bhdmi\b', re.I)
RE_DISPLAYPORT     = re.compile(r'\b(?:displayport|dp)\b', re.I)
RE_PORTS           = re.compile(r"\b(?:порт(?:и)?|роз['']єм(?:и)?|розєм(?:и)?|interface)\b", re.I)
RE_WATERPROOF      = re.compile(r'\b(?:вологозахист|водозахист|waterproof|splash\s*proof)\b', re.I)
RE_UKRAINIAN       = re.compile(r'\b(?:україн|укр)\b', re.I)
RE_TOUCHPAD        = re.compile(r'\b(?:тачпад|маніпулятор|трекпоінт|touchpad|trackpad|numberpad)\b', re.I)
RE_3G4G            = re.compile(r'\b(?:3g|4g|lte|nano\s*sim)\b', re.I)
RE_BLUETOOTH       = re.compile(r'\bbt\b|bluetooth', re.I)
RE_WIFI            = re.compile(r'\b(?:wi-fi|wifi|wireless|wlan)\b', re.I)
RE_LAN             = re.compile(r'\b(?:lan|rj-?45|ethernet)\b', re.I)
RE_WARRANTY        = re.compile(r'\b(?:гарант|warranty)\b', re.I)

RESOLUTION_MAP = {
    "fhd":   "1920x1080",
    "wuxga": "1920x1200",
    "qhd":   "2560x1440",
    "wqxga": "2560x1600",
    "uhd":   "3840x2160",
    "4k":    "3840x2160",
}


# ── Helper functions (same logic as Lenovo parser) ────────────────────────────

def detect_series(name):
    for kw in SERIES_MAP:
        if kw.lower() in name.lower():
            return kw
    return "-"


def extract_pn(name):
    name = name.strip()
    # ASUS part numbers like B1500CEAE-EJ3671W or 90NB0G71-M003K0
    m = re.search(r'\b([A-Z0-9]{2,4}[0-9]{3,4}[A-Z0-9_-]{4,})\b', name)
    if m:
        return m.group(1)
    candidates = RE_PART_NUMBER.findall(name)
    if candidates:
        return max(candidates, key=len)
    candidates = RE_PART_NUMBER_ALT.findall(name)
    return max(candidates, key=len) if candidates else "-"


def extract_cpu_details(cpu_model_str):
    if not cpu_model_str or cpu_model_str == "-":
        return {}
    details = {}
    s = cpu_model_str.lower()

    m = re.search(r'\((?:нов|кеш\s*)?(\d+)\s*(?:мб|mb)', s)
    if m:
        details["cache"] = m.group(1) + " МБ"

    freq_matches = re.findall(r'(\d+(?:[.,]\d+)?)\s*(?:ггц|ghz)', s)
    if freq_matches:
        freqs_numeric = []
        for f in freq_matches:
            try:
                freqs_numeric.append((float(f.replace(',', '.')), f.replace(',', '.')))
            except Exception:
                pass
        if len(freqs_numeric) >= 2:
            freqs_numeric.sort()
            details["base_freq"] = f"{freqs_numeric[0][1]} ГГц"
            details["max_freq"]  = f"{freqs_numeric[-1][1]} ГГц"
        elif len(freqs_numeric) == 1:
            if re.search(r'\b(?:до|up to|turbo|boost|max)\b', s):
                details["max_freq"] = freqs_numeric[0][1] + " ГГц"
            else:
                details["base_freq"] = freqs_numeric[0][1] + " ГГц"

    cores = re.search(r'(\d+)(?:\s*-)?core|(\d+)\s*ядер', s)
    if cores:
        details["cores"] = cores.group(1) or cores.group(2)

    threads = re.search(r'(\d+)(?:\s*-)?thread|(\d+)\s*потоків', s)
    if threads:
        details["threads"] = threads.group(1) or threads.group(2)

    return details


def parse_specs(items):
    keys = [
        "CPU_Brand", "CPU_Model", "CPU_Cores", "CPU_Threads",
        "CPU_Base_Freq_GHz", "CPU_Max_Freq_GHz", "CPU_Cache_L3_MB",
        "RAM_Type", "RAM_Freq", "RAM_Size_GB", "RAM_Slots",
        "Nakopychuvach_SSD", "Display_Diagonal", "Display_Max_Resolution",
        "Display_Matrix_Type", "Display_Cover", "Display_Brightness_nits",
        "Display_Contrast", "Display_Response_Time", "Display_Refresh_Rate",
        "Video_Brand", "GPU_Type", "GPU_Model", "GPU_Memory_MB",
        "OS", "Battery_Capacity_Wh", "Camera_MP", "USB_TypeA", "USB_TypeC",
        "HDMI", "DisplayPort", "Ports", "Keyboard_Waterproof",
        "Keyboard_Ukrainian", "Keyboard_Pointing_Device", "Network_3G4G",
        "Bluetooth", "WiFi", "LAN_Mbps", "Certificates", "Warranty",
        "Merezha_WiFi",
    ]
    r = {k: "-" for k in keys}

    def set_if_empty(key, value):
        if r[key] == "-" and value:
            r[key] = value.rstrip(";").strip() if isinstance(value, str) else str(value)

    def set_from_detail(key, value):
        if value and (r[key] in ("-", "Так", "Є")):
            r[key] = value.rstrip(";").strip() if isinstance(value, str) else str(value)

    full_text = " ".join([str(item).lower() for item in items if item])

    # First pass: DDR type
    for s in items:
        if not s:
            continue
        if RE_DDR.search(s):
            m = RE_DDR.search(s)
            if m:
                r["RAM_Type"] = m.group(0)
                break

    for s in items:
        if not s:
            continue
        sl = s.lower()

        # ── Display ──────────────────────────────────────────────────────────
        if RE_DIAG.search(sl):
            m = RE_DIAG.search(sl)
            if m and r["Display_Diagonal"] == "-":
                set_if_empty("Display_Diagonal", m.group(0))

        if RE_RESOLUTION.search(sl):
            m = RE_RESOLUTION.search(sl)
            if m:
                pixel_m = re.search(r'(\d{3,4})[xх](\d{3,4})', sl)
                name_m  = re.search(r'\b(fhd|wuxga|qhd|wqxga|uhd|4k)\b', sl, re.I)
                if pixel_m and name_m:
                    set_if_empty("Display_Max_Resolution",
                                 f"{name_m.group(0).upper()} ({pixel_m.group(0).upper()})")
                elif pixel_m:
                    set_if_empty("Display_Max_Resolution", pixel_m.group(0).upper())
                elif name_m:
                    name_key = name_m.group(0).lower()
                    pixels = RESOLUTION_MAP.get(name_key, "")
                    if pixels:
                        set_if_empty("Display_Max_Resolution",
                                     f"{name_m.group(0).upper()} ({pixels})")
                    else:
                        set_if_empty("Display_Max_Resolution", name_m.group(0).upper())
                else:
                    set_if_empty("Display_Max_Resolution", m.group(0).upper())

        if RE_DISPLAY_TYPE.search(sl):
            m = RE_DISPLAY_TYPE.search(sl)
            if m:
                set_if_empty("Display_Matrix_Type", m.group(0).upper())

        if RE_DISPLAY_COVER.search(sl):
            cover_m = RE_DISPLAY_COVER.search(sl)
            if cover_m:
                ct = cover_m.group(0).lower()
                if "глян" in ct or "gloss" in ct:
                    set_if_empty("Display_Cover", "Глянцеве покриття")
                elif "матов" in ct or "matte" in ct:
                    set_if_empty("Display_Cover", "Матове покриття")
                elif "антиблі" in ct or "antiglare" in ct:
                    set_if_empty("Display_Cover", "Антиблікове покриття")
                else:
                    set_if_empty("Display_Cover", ct)

        if RE_BRIGHTNESS.search(sl):
            m = RE_BRIGHTNESS.search(sl)
            if m:
                val = m.group(1) or m.group(2)
                if val:
                    set_if_empty("Display_Brightness_nits", f"{val} nits")

        if RE_CONTRAST.search(sl):
            cm = re.search(r'(?:контраст|contrast)[:\s]+(\d+(?:\.\d+)?(?:\s*:\s*\d+)?)', sl, re.I)
            if cm:
                set_if_empty("Display_Contrast", cm.group(1))

        if RE_RESPONSE_TIME.search(sl):
            m = RE_RESPONSE_TIME.search(sl)
            if m:
                val = m.group(1) or m.group(2)
                set_if_empty("Display_Response_Time", f"{val} мс")

        if RE_REFRESH_RATE.search(sl):
            freq_m = re.search(r'(\d+)\s*(?:hz|гц)\b', sl, re.I)
            if freq_m:
                set_if_empty("Display_Refresh_Rate", freq_m.group(0))
            else:
                freq_m = re.search(r'(?:частота\s*оновлення|refresh\s*rate)\s*[:\s]+(\d+)', sl, re.I)
                if freq_m:
                    set_if_empty("Display_Refresh_Rate", freq_m.group(1) + " Hz")

        # ── CPU ──────────────────────────────────────────────────────────────
        if r["CPU_Brand"] == "-":
            if "intel" in sl:
                set_if_empty("CPU_Brand", "Intel")
            elif "amd" in sl or "ryzen" in sl:
                set_if_empty("CPU_Brand", "AMD")
            elif "snapdragon" in sl:
                set_if_empty("CPU_Brand", "Snapdragon")
            elif "mediatek" in sl:
                set_if_empty("CPU_Brand", "MediaTek")

        cpu_related = any(k in sl for k in [
            "intel", "amd", "ryzen", "snapdragon", "celeron", "pentium",
            "core", "xeon", "i3", "i5", "i7", "i9", "процес", "cpu", "mediatek"
        ])
        if cpu_related and r["CPU_Model"] == "-":
            clean = re.sub(r'\s*\([^)]*(?:мб|ггц|кеш|ghz|mb|cache)[^)]*\)', '', s, flags=re.I).strip()
            clean = re.sub(r'\s*[®™]\s*', ' ', clean).strip()
            cl = clean.lower()
            if "intel" in cl:
                m = re.search(r'Core(?:\s+Ultra)?\s+(?:i[3-9][\w-]+|\w+\s+\w+)', clean, re.I)
                if not m:
                    m = re.search(r'(?:Celeron|Pentium|Xeon)\s+[\w-]+', clean, re.I)
                set_if_empty("CPU_Model", m.group(0).strip() if m else clean)
            elif "amd" in cl or "ryzen" in cl:
                m = re.search(r'Ryzen\s+\d+\s+\d{4,5}\w*', clean, re.I)
                if not m:
                    m = re.search(r'(?:Ryzen|Athlon)\s+\w+(?:\s+\w+)?', clean, re.I)
                set_if_empty("CPU_Model", m.group(0).strip() if m else clean)
            elif "snapdragon" in cl:
                m = re.search(r'Snapdragon\s+\w+(?:\s+[\w-]+)+', clean, re.I)
                if m:
                    set_if_empty("CPU_Model", m.group(0).strip())
            else:
                set_if_empty("CPU_Model", clean)

        if RE_CPU_CORES.search(sl):
            m = RE_CPU_CORES.search(sl)
            if m:
                set_if_empty("CPU_Cores", m.group(1))

        if RE_CPU_THREADS.search(sl):
            m = RE_CPU_THREADS.search(sl)
            if m:
                set_if_empty("CPU_Threads", m.group(1))

        if cpu_related:
            base_freq_m = re.search(r'(?:номін|базов)\w*\s*частота[\s,ггц]*(\d+(?:[.,]\d+)?)', sl, re.I)
            max_freq_m  = re.search(r'максимальна\s*частота[\s,ггц]*(\d+(?:[.,]\d+)?)', sl, re.I)
            if base_freq_m:
                set_if_empty("CPU_Base_Freq_GHz", base_freq_m.group(1) + " ГГц")
            if max_freq_m:
                set_if_empty("CPU_Max_Freq_GHz", max_freq_m.group(1) + " ГГц")

            freq_matches = RE_CPU_FREQ.findall(sl)
            if not freq_matches:
                freq_matches = RE_CPU_FREQ_GHZ.findall(sl)

            if len(freq_matches) >= 2:
                freqs = [float(f.replace(',', '.')) for f in freq_matches]
                low, high = min(freqs), max(freqs)
                set_if_empty("CPU_Base_Freq_GHz", f"{low} ГГц")
                set_if_empty("CPU_Max_Freq_GHz",  f"{high} ГГц")
            elif len(freq_matches) == 1:
                if "до" in sl or "turbo" in sl or "boost" in sl or "max" in sl or "up to" in sl:
                    set_if_empty("CPU_Max_Freq_GHz", freq_matches[0] + " ГГц")
                else:
                    set_if_empty("CPU_Base_Freq_GHz", freq_matches[0] + " ГГц")

        if RE_L3_CACHE.search(sl):
            m = re.search(r'l3[^0-9\n]{0,20}?(\d+)\s*(?:мб|mb)?', sl, re.I)
            if m:
                set_if_empty("CPU_Cache_L3_MB", m.group(1) + " МБ")

        # ── RAM ──────────────────────────────────────────────────────────────
        if RE_RAM_FREQ.search(sl):
            m = RE_RAM_FREQ.search(sl)
            if m:
                set_if_empty("RAM_Freq", m.group(0) + " МГц")

        if r["RAM_Size_GB"] == "-" and "ssd" not in sl and "hdd" not in sl and "nvme" not in sl:
            m = RE_RAM_SIZE.search(sl)
            if m:
                set_if_empty("RAM_Size_GB", m.group(0).upper())

        if "слот" in sl and any(k in sl for k in ["озу", "ddr", "пам", "гб", "ram"]):
            m = RE_RAM_SLOT.search(sl)
            if m:
                slot_num = m.group(1) or m.group(2)
                if slot_num:
                    set_if_empty("RAM_Slots", slot_num + " слотів")

        # ── Storage ──────────────────────────────────────────────────────────
        if RE_STORAGE.search(sl):
            set_if_empty("Nakopychuvach_SSD", s.strip())

        # ── GPU ──────────────────────────────────────────────────────────────
        if RE_GPU.search(sl):
            sl_lower = sl.lower()

            if r["GPU_Memory_MB"] == "-":
                m = re.search(r'\b(\d+)\s*(?:гб|gb)\b', sl, re.I)
                if m:
                    set_if_empty("GPU_Memory_MB", m.group(1) + " GB")
                elif "gpu memory" in sl_lower or "обсяг gpu" in sl_lower or "vram" in sl_lower:
                    m = re.search(r'\b(\d+)\b', sl)
                    if m:
                        set_if_empty("GPU_Memory_MB", m.group(1) + " GB")

            if r["Video_Brand"] == "-":
                if any(k in sl_lower for k in ["nvidia", "geforce", "rtx", "gtx"]):
                    set_if_empty("Video_Brand", "NVIDIA")
                elif any(k in sl_lower for k in ["amd", "radeon"]):
                    set_if_empty("Video_Brand", "AMD")
                elif any(k in sl_lower for k in ["intel", "iris", "arc"]):
                    set_if_empty("Video_Brand", "Intel")
                elif any(k in sl_lower for k in ["qualcomm", "adreno"]):
                    set_if_empty("Video_Brand", "Qualcomm")

            if r["GPU_Model"] == "-":
                m = re.search(r'\b(?:rtx|gtx|radeon\s+rx)\s*(\d{3,4})\b', sl, re.I)
                if m:
                    prefix = re.search(r'(rtx|gtx|radeon\s+rx)', sl, re.I)
                    set_if_empty("GPU_Model", f"{prefix.group(0).upper()} {m.group(1)}")
                else:
                    m = re.search(
                        r'\b(?:radeon\s+\d{3}m|arc\s+\w+|iris\s+\w+|intel\s+graphics|'
                        r'(?:intel\s+)?(?:uhd|hd|iris)\s+graphics|(?:qualcomm\s+)?adreno\s+gpu)\b',
                        sl, re.I)
                    if m:
                        model_text = m.group(0).strip()
                        parts = model_text.split()
                        model_text = ' '.join(
                            [p.upper() if p.upper() in ('UHD', 'HD', 'XE') else p.title()
                             for p in parts])
                        set_if_empty("GPU_Model", model_text)
                    else:
                        brand_m = re.search(
                            r'(?:nvidia|geforce|amd|radeon|intel|arc|qualcomm|adreno)', sl, re.I)
                        if brand_m:
                            gpu_text = sl[brand_m.start():].strip()
                            gpu_text = re.sub(r',.*', '', gpu_text).strip()
                            gpu_text = re.sub(r'\d+\s*(?:гб|gb|мб|mb)', '', gpu_text, flags=re.I).strip()
                            gpu_text = re.sub(r'\s*\([^)]*\)', '', gpu_text).strip()
                            if 3 < len(gpu_text) < 60:
                                set_if_empty("GPU_Model", gpu_text.title())

            if r["GPU_Type"] == "-":
                if r["Video_Brand"] == "NVIDIA" or any(k in sl_lower for k in ["rtx", "gtx", "radeon rx"]):
                    set_if_empty("GPU_Type", "дискретна")
                elif r["Video_Brand"] in ("Intel", "Qualcomm") or any(
                        k in sl_lower for k in ["iris", "arc", "uhd graphics", "adreno"]):
                    set_if_empty("GPU_Type", "інтегрована")
                elif r["Video_Brand"] == "AMD":
                    if any(k in sl_lower for k in ["radeon 780m", "radeon 680m", "radeon graphics"]):
                        set_if_empty("GPU_Type", "інтегрована")
                    else:
                        set_if_empty("GPU_Type", "дискретна")

        # ── OS ───────────────────────────────────────────────────────────────
        if RE_OS.search(sl):
            m = RE_OS.search(sl)
            if m:
                set_if_empty("OS", m.group(0).title())

        # ── Battery ──────────────────────────────────────────────────────────
        whr_m = RE_BATTERY_WHR.search(s)
        if whr_m:
            set_if_empty("Battery_Capacity_Wh", whr_m.group(1) + " Wh")
        elif RE_BATTERY.search(sl):
            m = RE_BATTERY.search(sl)
            if m:
                val = m.group(1) or m.group(2)
                if val:
                    set_if_empty("Battery_Capacity_Wh", val + " Wh")

        # ── Camera ───────────────────────────────────────────────────────────
        if RE_CAMERA.search(sl):
            cam_m = re.search(r'(\d+(?:\.\d+)?)\s*(?:мп|mp)', sl, re.I)
            if cam_m:
                set_if_empty("Camera_MP", cam_m.group(0))
            elif "fhd" in sl and ("camera" in sl or "камера" in sl):
                set_if_empty("Camera_MP", "FHD")
            elif "hd" in sl and ("camera" in sl or "камера" in sl):
                set_if_empty("Camera_MP", "HD")

        # ── Ports / Keyboard / Network — split on ‖ separator ────────────────
        if "‖" in s:
            usb_a_parts, usb_c_parts, dp_parts = [], [], []
            for entry in s.split("‖"):
                e  = entry.strip()
                el = e.lower()

                if re.match(r'процесор\s', el) or re.match(r'cpu\s', el):
                    raw = re.sub(r'^(?:процесор|cpu)\s*', '', e, flags=re.I).strip()
                    raw = re.sub(r'\s*\([^)]*(?:мб|ггц|кеш|ghz|mb|cache)[^)]*\)', '', raw, flags=re.I).strip()
                    raw = re.sub(r'\s*[®™]\s*', ' ', raw).strip()
                    if raw and r["CPU_Model"] == "-":
                        ms = re.search(r'Snapdragon\s+\w+(?:\s+[\w-]+)+', raw, re.I)
                        mi = re.search(r'Core(?:\s+Ultra)?\s+(?:i[3-9][\w-]+|\w+\s+\w+)', raw, re.I)
                        ma = re.search(r'Ryzen\s+\d+\s+\d{4,5}\w*', raw, re.I)
                        m  = ms or mi or ma
                        set_if_empty("CPU_Model", m.group(0).strip() if m else raw)

                elif re.match(r'(?:кеш\s*l3|l3\s*кеш)', el):
                    val = re.sub(r'^(?:кеш\s*l3|l3\s*кеш)\s*', '', e, flags=re.I).strip()
                    if val:
                        set_if_empty("CPU_Cache_L3_MB", val)

                elif re.match(r"тип\s*пам", el) or re.match(r"memory\s*type", el):
                    m = re.search(r'\b(?:LP)?DDR\d(?:X\b|\b)', e, re.I)
                    if m:
                        set_if_empty("RAM_Type", m.group(0).upper())

                elif re.match(r'usb\s*type-?a\s', el) or re.match(r'usb\s*3', el):
                    val = re.sub(r'^usb[\s\w-]*\s*', '', e, flags=re.I).strip()
                    if val:
                        usb_a_parts.append(val)

                elif re.match(r'usb[\s-]*type-?c\s', el) or re.match(r'thunderbolt', el):
                    val = re.sub(r'^(?:usb[\s-]*type-?c|thunderbolt)\s*', '', e, flags=re.I).strip()
                    if val:
                        usb_c_parts.append(val)

                elif re.match(r'hdmi[,\s]', el):
                    val = re.sub(r'^hdmi[,\s]*(?:шт\.?\s+)?', '', e, flags=re.I).strip()
                    if val:
                        set_if_empty("HDMI", val)

                elif re.match(r'displayport[,\s]', el) or re.match(r'mini\s*dp', el):
                    val = re.sub(r'^(?:displayport|mini\s*dp)[,\s]*(?:шт\.?\s+)?', '', e, flags=re.I).strip()
                    if val:
                        dp_parts.append(val)

                elif re.match(r'порти\s', el) or re.match(r'interface', el):
                    val = re.sub(r'^(?:порти|interface)\s*', '', e, flags=re.I).strip()
                    if val and val.lower() != "немає":
                        set_if_empty("Ports", val)

                elif re.match(r'клавіатура\s*\(вологозахист\)', el) or re.match(r'splash\s*proof', el):
                    val = re.sub(r'^(?:клавіатура\s*\(вологозахист\)|splash\s*proof)\s*', '', e, flags=re.I).strip()
                    if val:
                        set_if_empty("Keyboard_Waterproof", val)

                elif re.match(r'українська\s*мова\s', el):
                    val = re.sub(r'^українська\s*мова\s*', '', e, flags=re.I).strip()
                    if val:
                        set_if_empty("Keyboard_Ukrainian", val)

                elif re.match(r'маніпулятори\s', el) or re.match(r'touchpad', el):
                    val = re.sub(r'^(?:маніпулятори|touchpad)\s*', '', e, flags=re.I).strip()
                    if val and val.lower() != "немає":
                        set_if_empty("Keyboard_Pointing_Device", val)

                elif re.match(r'3g/4g\s', el) or re.match(r'lte', el):
                    val = re.sub(r'^(?:3g/4g|lte)\s*', '', e, flags=re.I).strip()
                    if val:
                        set_from_detail("Network_3G4G", val)

                elif re.match(r'bluetooth\s', el):
                    val = re.sub(r'^bluetooth\s*', '', e, flags=re.I).strip()
                    if val and val.lower() != "немає":
                        set_from_detail("Bluetooth", val)

                elif re.match(r'wi-fi\s', el) or re.match(r'wlan\s', el) or re.match(r'wireless\s', el):
                    val = re.sub(r'^(?:wi-fi|wlan|wireless)\s*', '', e, flags=re.I).strip()
                    if val and val.lower() != "немає":
                        set_from_detail("WiFi", val)
                        set_from_detail("Merezha_WiFi", val)

                elif re.match(r'lan\s*rj-?45', el) or re.match(r'ethernet', el):
                    val = re.sub(r'^(?:lan\s*rj-?45|ethernet)[,\s]*(?:мбіт/с\s*)?', '', e, flags=re.I).strip()
                    if val and val.lower() != "немає":
                        set_from_detail("LAN_Mbps", val)

                elif re.match(r'сертифікати\s', el) or re.match(r'certificate', el):
                    val = re.sub(r'^(?:сертифікати|certificate)\s*', '', e, flags=re.I).strip()
                    if val:
                        set_from_detail("Certificates", val)

                elif re.match(r'(?:термін\s*базово|warranty)', el):
                    m = re.search(r'(\d+\s*(?:рік|роки|років|місяц\w*|year|month))', e, re.I)
                    if m:
                        set_from_detail("Warranty", m.group(1).strip())

            if usb_a_parts:
                set_if_empty("USB_TypeA", " / ".join(usb_a_parts))
            if usb_c_parts:
                set_if_empty("USB_TypeC", " / ".join(usb_c_parts))
            if dp_parts:
                set_if_empty("DisplayPort", " / ".join(dp_parts))
        else:
            # Fallback: no ‖ separator
            if RE_USB_A.search(sl):
                set_if_empty("USB_TypeA", s.strip())
            if RE_USB_C.search(sl):
                set_if_empty("USB_TypeC", s.strip())
            if RE_HDMI.search(sl):
                set_if_empty("HDMI", s.strip())
            if RE_DISPLAYPORT.search(sl):
                set_if_empty("DisplayPort", s.strip())
            if RE_PORTS.search(sl):
                set_if_empty("Ports", s.strip())
            if RE_WATERPROOF.search(sl):
                set_if_empty("Keyboard_Waterproof", "Так")
            if RE_UKRAINIAN.search(sl):
                set_if_empty("Keyboard_Ukrainian", "Так")
            if RE_TOUCHPAD.search(sl):
                set_if_empty("Keyboard_Pointing_Device", "Так")
            if RE_3G4G.search(sl):
                set_if_empty("Network_3G4G", "Так")
            if RE_BLUETOOTH.search(sl):
                set_if_empty("Bluetooth", "Так")
            if RE_WIFI.search(sl):
                set_if_empty("WiFi", "Так")
                set_if_empty("Merezha_WiFi", "Так")
            if RE_LAN.search(sl):
                m = re.search(r'(\d{3,4})\s*(?:мбіт|mbps)', sl, re.I)
                if m:
                    set_if_empty("LAN_Mbps", m.group(0))
                else:
                    set_if_empty("LAN_Mbps", "Так")

        # ── Warranty ─────────────────────────────────────────────────────────
        if RE_WARRANTY.search(sl):
            m = re.search(r'(\d+)\s*(?:місяц|рік|year|month)', sl, re.I)
            if m:
                set_if_empty("Warranty", m.group(0))
            else:
                set_if_empty("Warranty", "Так")

    # Post-process CPU details from model string
    if r["CPU_Model"] != "-":
        cpu_details = extract_cpu_details(r["CPU_Model"])
        for k, v in [
            ("CPU_Cache_L3_MB", cpu_details.get("cache")),
            ("CPU_Base_Freq_GHz", cpu_details.get("base_freq")),
            ("CPU_Max_Freq_GHz", cpu_details.get("max_freq")),
            ("CPU_Cores", cpu_details.get("cores")),
            ("CPU_Threads", cpu_details.get("threads")),
        ]:
            if v:
                set_if_empty(k, v)

    # Search cores/threads in full text if still missing
    if r["CPU_Cores"] == "-" or r["CPU_Threads"] == "-":
        cores_m   = re.search(r'кількість ядер\s+(\d+)', full_text, re.I)
        threads_m = re.search(r'кількість потоків\s+(\d+)', full_text, re.I)
        if cores_m:
            set_if_empty("CPU_Cores", cores_m.group(1))
        if threads_m:
            set_if_empty("CPU_Threads", threads_m.group(1))
        if not cores_m:
            cores_m = re.search(r'(\d+)\s*(?:ядер|ядра|cores?)', full_text, re.I)
            if cores_m:
                set_if_empty("CPU_Cores", cores_m.group(1))
        if not threads_m:
            threads_m = re.search(r'(\d+)\s*(?:потоків|потоки|threads?)', full_text, re.I)
            if threads_m:
                set_if_empty("CPU_Threads", threads_m.group(1))

    # ASUS default assumptions
    if r["Bluetooth"] == "-" and "asus" in full_text:
        set_if_empty("Bluetooth", "Так")
    if r["WiFi"] == "-" and "asus" in full_text:
        set_if_empty("WiFi", "Так")
        set_if_empty("Merezha_WiFi", "Так")

    return r


# ── ASUS detail-page extraction ───────────────────────────────────────────────

def fetch_detail_page_text(pg, url):
    """Fetch spec text from an ASUS product detail page."""
    try:
        pg.goto(url, wait_until="domcontentloaded", timeout=PAGE_TIMEOUT)

        # Wait for specs section to appear (up to 10 s)
        deadline = time.time() + 10
        while time.time() < deadline:
            try:
                count = pg.evaluate(
                    "document.querySelectorAll('.asus-specificaton-table, .specs-list, "
                    "[class*=\"spec\"], [class*=\"Spec\"]').length"
                )
                if count > 0:
                    break
            except Exception:
                pass
            time.sleep(0.3)

        html = pg.content()
        soup = BeautifulSoup(html, "lxml")
        parts = []

        # Strategy 1: ASUS specification table (common layout)
        # Rows with th/td pairs in a specs table
        for table in soup.select("table.asus-specificaton-table, table[class*='spec'], table[class*='Spec']"):
            for row in table.select("tr"):
                th = row.select_one("th")
                td = row.select_one("td")
                if th and td:
                    label = th.get_text(strip=True)
                    value = td.get_text(separator=" ", strip=True)
                    if label and value:
                        parts.append(f"{label} {value}")

        # Strategy 2: definition-list style specs (dl/dt/dd)
        if not parts:
            for dl in soup.select("dl, .specs-list, [class*='SpecList'], [class*='spec-list']"):
                dts = dl.select("dt")
                dds = dl.select("dd")
                for dt, dd in zip(dts, dds):
                    label = dt.get_text(strip=True)
                    value = dd.get_text(separator=" ", strip=True)
                    if label and value:
                        parts.append(f"{label} {value}")

        # Strategy 3: div-based label/value pairs (ASUS React store)
        if not parts:
            for section in soup.select("[class*='spec'], [class*='Spec'], [class*='detail']"):
                labels = section.select("[class*='label'], [class*='Label'], [class*='name'], [class*='Name']")
                values = section.select("[class*='value'], [class*='Value'], [class*='desc'], [class*='content']")
                for lbl, val in zip(labels, values):
                    l_text = lbl.get_text(strip=True)
                    v_text = val.get_text(separator=" ", strip=True)
                    if l_text and v_text:
                        parts.append(f"{l_text} {v_text}")

        # Strategy 4: any table on the page with 2 columns
        if not parts:
            for table in soup.select("table"):
                for row in table.select("tr"):
                    cells = row.select("td, th")
                    if len(cells) >= 2:
                        label = cells[0].get_text(strip=True)
                        value = cells[1].get_text(separator=" ", strip=True)
                        if label and value and len(label) < 80:
                            parts.append(f"{label} {value}")

        # Strategy 5: scrape visible text from meta/og tags (fallback)
        if not parts:
            desc = soup.select_one('meta[name="description"]') or soup.select_one('meta[property="og:description"]')
            if desc:
                parts.append(desc.get("content", ""))

        return " ‖ ".join(parts) if parts else ""

    except Exception as e:
        print(f"    [!] detail page error: {e}")
        return ""


# ── Catalog card extraction ───────────────────────────────────────────────────

# ASUS store uses several possible card selectors — try them all
CARD_SELECTORS = [
    "div.ProductItem_ProductItem__",
    "div[class*='ProductItem']",
    "div[class*='product-item']",
    "div[class*='product_item']",
    "article[class*='product']",
    "li[class*='product']",
    "div[class*='ProductCard']",
    "div[class*='product-card']",
    "[data-product-id]",
]

LINK_SELECTORS = [
    "a[class*='ProductItem']",
    "a[class*='product']",
    "a[href*='/laptops/']",
    "a[href*='/notebook']",
    "a",
]

PRICE_SELECTORS = [
    "[class*='price']",
    "[class*='Price']",
    "[class*='cost']",
    "[data-price]",
]

NAME_SELECTORS = [
    "[class*='title']",
    "[class*='Title']",
    "[class*='name']",
    "[class*='Name']",
    "h2", "h3", "h4",
]


def _extract_cards_generic(soup):
    """Generic ASUS card extractor using heuristic selectors."""
    records_raw = []

    for card_sel in CARD_SELECTORS:
        try:
            cards = soup.select(card_sel)
        except Exception:
            cards = []
        if not cards:
            continue

        for card in cards:
            # Find product link
            link_el = None
            for lsel in LINK_SELECTORS:
                link_el = card.select_one(lsel)
                if link_el:
                    break
            if not link_el:
                continue

            href = link_el.get("href", "")
            if not href or href == "#":
                continue
            url = BASE_URL + href if href.startswith("/") else href
            if "asus.com" not in url:
                continue

            # Find name
            name = ""
            for nsel in NAME_SELECTORS:
                name_el = card.select_one(nsel)
                if name_el:
                    name = name_el.get_text(strip=True)
                    if name:
                        break
            if not name:
                name = link_el.get_text(strip=True)
            if not name:
                continue

            # Find price
            price = "-"
            for psel in PRICE_SELECTORS:
                price_el = card.select_one(psel)
                if price_el:
                    price_text = price_el.get_text()
                    price = re.sub(r"[^\d\s]", "", price_text).strip()
                    if price:
                        break

            records_raw.append((name, url, price, card))

        if records_raw:
            break

    return records_raw


def _extract_cards_js(pg):
    """Extract product data via JavaScript evaluation when HTML selectors fail."""
    try:
        data = pg.evaluate("""
            (() => {
                const results = [];
                // Try common ASUS product link patterns
                const links = Array.from(document.querySelectorAll('a[href]')).filter(a => {
                    const h = a.getAttribute('href') || '';
                    return (h.includes('/laptops/') || h.includes('/notebook')) &&
                           !h.includes('#') && a.textContent.trim().length > 5;
                });
                const seen = new Set();
                links.forEach(a => {
                    const href = a.getAttribute('href');
                    if (seen.has(href)) return;
                    seen.add(href);
                    // Look for price near the link
                    let price = '-';
                    const parent = a.closest('[class*="product"], [class*="Product"], li, article') || a.parentElement;
                    if (parent) {
                        const priceEl = parent.querySelector('[class*="price"], [class*="Price"]');
                        if (priceEl) price = priceEl.textContent.replace(/[^\\d\\s]/g, '').trim();
                    }
                    results.push({
                        name: a.textContent.trim(),
                        href: href,
                        price: price
                    });
                });
                return results;
            })()
        """)
        return data or []
    except Exception:
        return []


def parse_cards_from_page(pg, detail_pg=None):
    """Parse product cards from the current page loaded in `pg`."""
    html = pg.content()
    soup = BeautifulSoup(html, "lxml")
    records = []

    raw_cards = _extract_cards_generic(soup)

    # Fallback: JS extraction
    if not raw_cards:
        js_data = _extract_cards_js(pg)
        for item in js_data:
            name  = item.get("name", "").strip()
            href  = item.get("href", "")
            price = item.get("price", "-")
            if not name or not href:
                continue
            url = BASE_URL + href if href.startswith("/") else href
            raw_cards.append((name, url, price, None))

    for name, url, price, _card in raw_cards:
        specs_items = [name]

        # Collect any visible spec items from card HTML
        if _card is not None:
            for li in _card.select("li, [class*='spec'], [class*='Spec'], [class*='feature']"):
                text = li.get_text(strip=True)
                if text and text != name and len(text) < 300:
                    specs_items.append(text)

        if detail_pg is not None and LOAD_DETAIL_PAGES:
            detail_text = fetch_detail_page_text(detail_pg, url)
            if detail_text:
                specs_items.append(detail_text)

        specs = parse_specs(specs_items)
        records.append({
            "Seria":             detect_series(name),
            "Model":             name,
            "Part Number":       extract_pn(name),
            "CPU_Brand":         specs["CPU_Brand"],
            "CPU_Model":         specs["CPU_Model"],
            "CPU_Cores":         specs["CPU_Cores"],
            "CPU_Threads":       specs["CPU_Threads"],
            "CPU_Base_Freq_GHz": specs["CPU_Base_Freq_GHz"],
            "CPU_Max_Freq_GHz":  specs["CPU_Max_Freq_GHz"],
            "CPU_Cache_L3_MB":   specs["CPU_Cache_L3_MB"],
            "RAM_Type":          specs["RAM_Type"],
            "RAM_Freq":          specs["RAM_Freq"],
            "RAM_Size_GB":       specs["RAM_Size_GB"],
            "RAM_Slots":         specs["RAM_Slots"],
            "Nakopychuvach_SSD": specs["Nakopychuvach_SSD"],
            "Display_Diagonal":  specs["Display_Diagonal"],
            "Display_Max_Resolution": specs["Display_Max_Resolution"],
            "Display_Matrix_Type": specs["Display_Matrix_Type"],
            "Display_Cover":     specs["Display_Cover"],
            "Display_Brightness_nits": specs["Display_Brightness_nits"],
            "Display_Contrast":  specs["Display_Contrast"],
            "Display_Response_Time": specs["Display_Response_Time"],
            "Display_Refresh_Rate": specs["Display_Refresh_Rate"],
            "Video_Brand":       specs["Video_Brand"],
            "GPU_Type":          specs["GPU_Type"],
            "GPU_Model":         specs["GPU_Model"],
            "GPU_Memory_MB":     specs["GPU_Memory_MB"],
            "OS":                specs["OS"],
            "Battery_Capacity_Wh": specs["Battery_Capacity_Wh"],
            "Camera_MP":         specs["Camera_MP"],
            "USB_TypeA":         specs["USB_TypeA"],
            "USB_TypeC":         specs["USB_TypeC"],
            "HDMI":              specs["HDMI"],
            "DisplayPort":       specs["DisplayPort"],
            "Ports":             specs["Ports"],
            "Keyboard_Waterproof":     specs["Keyboard_Waterproof"],
            "Keyboard_Ukrainian":      specs["Keyboard_Ukrainian"],
            "Keyboard_Pointing_Device": specs["Keyboard_Pointing_Device"],
            "Network_3G4G":      specs["Network_3G4G"],
            "Bluetooth":         specs["Bluetooth"],
            "WiFi":              specs["WiFi"],
            "LAN_Mbps":          specs["LAN_Mbps"],
            "Certificates":      specs["Certificates"],
            "Warranty":          specs["Warranty"],
            "Merezha_WiFi":      specs["Merezha_WiFi"],
            "Tsina_UAH":         price,
            "URL":               url,
        })

    return records


# ── Page loading ──────────────────────────────────────────────────────────────

def _count_products(pg):
    """Return number of detectable product elements on page."""
    selectors = (
        "[class*='ProductItem'], [class*='product-item'], "
        "[class*='ProductCard'], [class*='product-card'], "
        "[data-product-id]"
    )
    try:
        return pg.evaluate(f"document.querySelectorAll(`{selectors}`).length")
    except Exception:
        return 0


def load_page_and_wait(pg, url):
    """Navigate to url, wait for product cards to appear. Returns True if found."""
    for attempt in range(2):
        try:
            pg.goto(url, wait_until="domcontentloaded", timeout=PAGE_TIMEOUT)
            break
        except Exception as err:
            if attempt >= 1:
                print(f"    Load error: {err}")
                return False
            time.sleep(1)

    # Dismiss cookie banner if present
    for sel in ["button[id*='accept'], button[class*='accept'], button[class*='cookie']"]:
        try:
            btn = pg.query_selector(sel)
            if btn:
                btn.click()
                time.sleep(0.3)
                break
        except Exception:
            pass

    # Wait for products to appear (up to 15 s)
    deadline = time.time() + 15
    last_count = 0
    while time.time() < deadline:
        count = _count_products(pg)
        if count > 0 and count == last_count:
            return True
        last_count = count
        time.sleep(0.4)

    # Last chance — also check for product links
    final_count = _count_products(pg)
    if final_count > 0:
        return True

    # Check product links as fallback
    try:
        link_count = pg.evaluate(
            "document.querySelectorAll('a[href*=\"/laptops/\"], a[href*=\"/notebook\"]').length"
        )
        if link_count > 3:
            print(f"    Found {link_count} product links (no card containers)")
            return True
    except Exception:
        pass

    # Save debug HTML
    html = pg.content()
    with open("debug_asus_page.html", "w", encoding="utf-8") as f:
        f.write(html)
    print("    No products found — debug HTML saved to debug_asus_page.html")
    return False


def detect_pagination(pg):
    """Detect total pages and pagination URL pattern for ASUS store."""
    # ASUS store typically uses ?page=N or /page/N or query params like ?p=N
    patterns_to_try = [
        CATALOG_URL + "?page={page}",
        CATALOG_URL + "?p={page}",
        CATALOG_URL + "page/{page}/",
    ]

    # Look for pagination elements to find total page count
    total_pages = 1
    try:
        # Try getting page count from pagination buttons
        nums = pg.evaluate("""
            (() => {
                const btns = Array.from(document.querySelectorAll(
                    '[class*="pagination"] a, [class*="pagination"] button, [class*="page-item"] a'
                ));
                const nums = btns.map(b => parseInt(b.textContent.trim(), 10)).filter(n => !isNaN(n) && n > 0);
                return nums.length ? Math.max(...nums) : 1;
            })()
        """)
        if nums and nums > 1:
            total_pages = nums
    except Exception:
        pass

    # If can't detect, also check "X results" text
    if total_pages == 1:
        try:
            count_text = pg.evaluate("""
                (() => {
                    const el = document.querySelector('[class*="result"], [class*="count"], [class*="total"]');
                    return el ? el.textContent : '';
                })()
            """)
            m = re.search(r'(\d+)', count_text or "")
            if m:
                items_total = int(m.group(1))
                # ASUS store usually shows ~24 items per page
                total_pages = max(1, (items_total + 23) // 24)
        except Exception:
            pass

    return total_pages, patterns_to_try[0]


# ── Excel formatting ──────────────────────────────────────────────────────────

def format_xlsx(path, total):
    wb = load_workbook(path)
    ws = wb.active
    ws.title = "ASUS Notebooks"

    UA = {
        "Seria": "Серія", "Model": "Модель", "Part Number": "Part Number",
        "CPU_Brand": "Бренд CPU", "CPU_Model": "Модель CPU", "CPU_Cores": "Кількість ядер",
        "CPU_Threads": "Кількість потоків", "CPU_Base_Freq_GHz": "Номінальна частота, ГГц",
        "CPU_Max_Freq_GHz": "Максимальна частота, ГГц", "CPU_Cache_L3_MB": "Кеш L3, МБ",
        "RAM_Type": "Тип ОЗП", "RAM_Freq": "Частота ОЗП",
        "RAM_Size_GB": "Обсяг ОЗП, ГБ", "RAM_Slots": "Кількість слотів",
        "Nakopychuvach_SSD": "M.2 SSD, ГБ",
        "Display_Diagonal": "Діагональ екрана", "Display_Max_Resolution": "Макс. роздільна здатність",
        "Display_Matrix_Type": "Тип матриці", "Display_Cover": "Покриття екрану",
        "Display_Brightness_nits": "Яскравість, ніт", "Display_Contrast": "Контраст",
        "Display_Response_Time": "Час реагування", "Display_Refresh_Rate": "Частота оновлення",
        "Video_Brand": "Видео_Бренд", "GPU_Type": "Тип_GPU", "GPU_Model": "Модель GPU",
        "GPU_Memory_MB": "Обсяг GPU, МБ", "OS": "ОС",
        "Battery_Capacity_Wh": "Енергетична ємність, Вт*год",
        "Camera_MP": "WEB-камера, Мп", "USB_TypeA": "USB Type-A", "USB_TypeC": "USB Type-C",
        "HDMI": "HDMI, шт.", "DisplayPort": "DisplayPort, шт.", "Ports": "Порти",
        "Keyboard_Waterproof": "Клавіатура (вологозахист)", "Keyboard_Ukrainian": "Українська мова",
        "Keyboard_Pointing_Device": "Маніпулятори", "Network_3G4G": "3G/4G",
        "Bluetooth": "Bluetooth", "WiFi": "Wi-Fi", "LAN_Mbps": "LAN RJ-45, Мбіт/с",
        "Certificates": "Сертифікати", "Warranty": "Термін базової гарантії від виробника",
        "Merezha_WiFi": "Мережа / Wi-Fi", "Tsina_UAH": "Ціна (UAH)", "URL": "URL",
    }
    for i, k in enumerate(COLUMNS, 1):
        ws.cell(row=1, column=i).value = UA.get(k, k)

    # ASUS brand colour: dark navy
    HFILL = PatternFill("solid", fgColor="00539B")
    AFILL = PatternFill("solid", fgColor="EEF4FF")
    t   = Side(style="thin", color="BBBBBB")
    brd = Border(left=t, right=t, top=t, bottom=t)

    for cell in ws[1]:
        cell.font      = Font(bold=True, color="FFFFFF", name="Arial", size=10)
        cell.fill      = HFILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border    = brd
    ws.row_dimensions[1].height = 32

    for r in range(2, ws.max_row + 1):
        for cell in ws[r]:
            cell.font      = Font(name="Arial", size=9)
            cell.fill      = AFILL if r % 2 == 0 else PatternFill()
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.border    = brd
        ws.row_dimensions[r].height = 45

    for i in range(1, len(COLUMNS) + 1):
        ws.column_dimensions[get_column_letter(i)].width = 20

    url_col = len(COLUMNS)
    for r in range(2, ws.max_row + 1):
        cell = ws.cell(row=r, column=url_col)
        if cell.value and str(cell.value).startswith("http"):
            cell.hyperlink = cell.value
            cell.font = Font(name="Arial", size=9, color="1155CC", underline="single")

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    info = wb.create_sheet("Info")
    info["A1"], info["B1"] = "Сформовано:",  datetime.now().strftime("%d.%m.%Y %H:%M")
    info["A2"], info["B2"] = "Джерело:",     CATALOG_URL
    info["A3"], info["B3"] = "Моделей:",     total
    for c in ["A1", "A2", "A3"]:
        info[c].font = Font(bold=True, name="Arial")
    wb.save(path)


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    print("=" * 60)
    print("  ASUS Notebooks Parser — asus.com/ua-ua/store/laptops/")
    print("=" * 60)

    all_records = []

    with sync_playwright() as pw:
        browser = pw.chromium.launch(
            headless=True,
            args=["--disable-blink-features=AutomationControlled", "--no-sandbox"]
        )
        ctx = browser.new_context(
            viewport={"width": 1440, "height": 900},
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/124.0.0.0 Safari/537.36"
            ),
            locale="uk-UA",
        )
        pg = ctx.new_page()

        # Page 1
        print("\nЗавантаження сторінки 1...")
        ok = load_page_and_wait(pg, CATALOG_URL)
        if not ok:
            print("ПОМИЛКА: товари не знайдено. Зупинка.")
            browser.close()
            return

        total_pages, url_pattern = detect_pagination(pg)
        print(f"Сторінок: {total_pages}  (приблизно {total_pages * 24} товарів)")
        print()

        # Parse page 1 (already loaded)
        recs = parse_cards_from_page(pg, pg)
        all_records.extend(recs)
        print(f"  [1/{total_pages}] зібрано {len(recs)} (всього: {len(all_records)})")

        # Pages 2..N
        for page_num in range(2, total_pages + 1):
            if len(all_records) >= MAX_ITEMS:
                print(f"\nДосягнуто ліміту {MAX_ITEMS} товарів.")
                break

            print(f"  [{page_num}/{total_pages}] ...", end=" ", flush=True)

            success  = False
            retry    = 0

            while not success and retry < 2:
                url = url_pattern.replace("{page}", str(page_num))
                ok  = load_page_and_wait(pg, url)
                if ok:
                    recs = parse_cards_from_page(pg, pg)
                    if recs:
                        all_records.extend(recs)
                        print(f"зібрано {len(recs)} (всього: {len(all_records)})")
                        success = True
                        break

                retry += 1
                if retry < 2:
                    print("\n  [!] Повтор...", end=" ", flush=True)
                    time.sleep(2)

            if not success:
                # Try scroll-based pagination: click "Завантажити ще" / "Load more"
                try:
                    more_btn = pg.query_selector(
                        "button[class*='load-more'], button[class*='LoadMore'], "
                        "button[class*='show-more'], a[class*='load-more']"
                    )
                    if more_btn:
                        more_btn.click()
                        time.sleep(2)
                        recs = parse_cards_from_page(pg, pg)
                        if recs:
                            all_records.extend(recs)
                            print(f"зібрано {len(recs)} (всього: {len(all_records)})")
                            success = True
                except Exception:
                    pass

            if not success:
                print("0 — пропуск")

        browser.close()

    # Save
    print()
    all_records = all_records[:MAX_ITEMS]
    print(f"Збереження Excel ({len(all_records)} моделей)...")
    df = pd.DataFrame(all_records, columns=COLUMNS)
    before = len(df)
    valid_mask  = df["Part Number"] != "-"
    valid_df    = df[valid_mask].drop_duplicates(subset=["Part Number"], keep="first")
    # Also deduplicate by URL for items without Part Number
    invalid_df  = df[~valid_mask].drop_duplicates(subset=["URL"], keep="first")
    df = pd.concat([valid_df, invalid_df], ignore_index=True)
    if len(df) < before:
        print(f"Дублікатів виключено: {before - len(df)}")

    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(OUTPUT_PATH, index=False, sheet_name="ASUS Notebooks")
    print("Форматування...")
    format_xlsx(OUTPUT_PATH, len(df))

    print()
    print("=" * 60)
    print(f"ГОТОВО!  {OUTPUT_PATH}")
    print(f"Моделей у таблиці: {len(df)}")
    print("=" * 60)


if __name__ == "__main__":
    main()
