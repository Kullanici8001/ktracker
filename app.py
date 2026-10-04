import csv
import io
import os
import re
import requests
import zipfile
import xml.etree.ElementTree as ET
from flask import Flask, render_template, redirect, url_for, request, jsonify, abort

app = Flask(__name__)

# Google Sheets API Key
GOOGLE_SHEETS_API_KEY = os.environ.get(
    "GOOGLE_SHEETS_API_KEY",
    "AIzaSyAwt90oJ0remLZoarvo-Iz2AFDKzfHUFAw"
)

# ─── Google Sheets helpers ────────────────────────────────────────────────────

def extract_sheet_id(raw: str) -> str | None:
    """URL veya bare ID kabul et."""
    raw = raw.strip()
    m = re.search(r"/(?:spreadsheets/d|sh)/([a-zA-Z0-9_-]+)", raw)
    if m:
        return m.group(1)
    if re.fullmatch(r"[a-zA-Z0-9_-]+", raw):
        return raw
    return None


# ─── Hücre İçi Resimler (Insert Image in Cell) Çıkarıcı ────────────────────────
_XLSX_IMAGES_CACHE: dict[str, dict[tuple[str, int, str], str]] = {}

def get_xlsx_cell_images(sheet_id: str) -> dict[tuple[str, int, str], str]:
    """
    Google Sheets REST API'sinin döndüremediği "Hücre içine eklenen" (in-cell) resimleri
    tablonun XLSX exportundan çekip yerel statik dizine kaydeder ve
    (sheet_name.lower(), row_1_indexed, col_letter) -> url eşlemesi döndürür.
    """
    if sheet_id in _XLSX_IMAGES_CACHE:
        return _XLSX_IMAGES_CACHE[sheet_id]

    cache_dir = os.path.join(app.static_folder or "static", "cache", sheet_id)
    os.makedirs(cache_dir, exist_ok=True)

    url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"
    cell_images: dict[tuple[str, int, str], str] = {}
    try:
        r = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=20)
        if r.status_code == 200:
            zf = zipfile.ZipFile(io.BytesIO(r.content))
            for name in zf.namelist():
                if name.startswith("xl/media/"):
                    img_name = os.path.basename(name)
                    out_path = os.path.join(cache_dir, img_name)
                    if not os.path.exists(out_path):
                        with open(out_path, "wb") as f:
                            f.write(zf.read(name))

            wb_tree = ET.fromstring(zf.read("xl/workbook.xml"))
            wb_rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
            rel_map = {rel.attrib["Id"]: rel.attrib["Target"] for rel in wb_rels}

            ns = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
            r_ns = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
            sheet_to_file = {}
            for s in wb_tree.findall(f".//{ns}sheet"):
                s_name = s.attrib["name"]
                r_id = s.attrib[f"{r_ns}id"]
                target = rel_map[r_id]
                if not target.startswith("xl/"):
                    target = "xl/" + target.lstrip("/")
                sheet_to_file[s_name.lower()] = target

            for s_name_lower, s_file in sheet_to_file.items():
                s_dir = os.path.dirname(s_file)
                s_base = os.path.basename(s_file)
                s_rels_file = f"{s_dir}/_rels/{s_base}.rels"
                if s_rels_file in zf.namelist():
                    s_rels_tree = ET.fromstring(zf.read(s_rels_file))
                    drawing_target = None
                    for rel in s_rels_tree:
                        if rel.attrib.get("Type", "").endswith("/drawing"):
                            drawing_target = rel.attrib.get("Target")
                            break

                    if drawing_target:
                        drawing_path = os.path.normpath(f"{s_dir}/{drawing_target}")
                        d_dir = os.path.dirname(drawing_path)
                        d_base = os.path.basename(drawing_path)
                        drawing_rels_path = f"{d_dir}/_rels/{d_base}.rels"

                        d_rels_map = {}
                        if drawing_rels_path in zf.namelist():
                            d_rels_tree = ET.fromstring(zf.read(drawing_rels_path))
                            for rel in d_rels_tree:
                                d_rels_map[rel.attrib["Id"]] = os.path.basename(rel.attrib["Target"])

                        if drawing_path in zf.namelist():
                            d_tree = ET.fromstring(zf.read(drawing_path))
                            for anchor in d_tree:
                                from_tag = anchor.find("{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}from")
                                if from_tag is not None:
                                    c_node = from_tag.find("{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}col")
                                    r_node = from_tag.find("{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}row")
                                    if c_node is not None and r_node is not None:
                                        c = int(c_node.text)
                                        r_idx = int(r_node.text)
                                        blip = anchor.find(".//{http://schemas.openxmlformats.org/drawingml/2006/main}blip")
                                        embed = blip.attrib.get(f"{r_ns}embed") if blip is not None else None
                                        img_name = d_rels_map.get(embed)
                                        if img_name:
                                            col_letter = chr(65 + c)
                                            img_path = f"/static/cache/{sheet_id}/{img_name}"
                                            cell_images[(s_name_lower, r_idx + 1, col_letter)] = img_path
                                            cell_images[(r_idx + 1, col_letter)] = img_path
    except Exception as e:
        print("XLSX hücre resmi çıkarma hatası:", e)

    _XLSX_IMAGES_CACHE[sheet_id] = cell_images
    return cell_images


def fetch_sheet_data(sheet_id: str, title: str) -> list[list[str]]:
    """Bir sekmeyi API ile JSON dizisi olarak indir ve hücre içi resimlerle zenginleştir."""
    if not title or title == "0":
        return []
    
    cell_images = {} # Kullanıcı isteği: resim işi şimdilik pas geçildi, sayfalar anında yüklenir
    title_lower = title.strip().lower()

    import urllib.parse
    title_enc = urllib.parse.quote(title)
    
    fields = "sheets.data.rowData.values(formattedValue,hyperlink)"
    url = f"https://sheets.googleapis.com/v4/spreadsheets/{sheet_id}?includeGridData=true&ranges={title_enc}!A:J&key={GOOGLE_SHEETS_API_KEY}&fields={fields}"
    
    try:
        r = requests.get(url, timeout=30)
        if r.status_code == 200:
            data = r.json()
            sheets = data.get("sheets", [])
            if not sheets:
                return []
            
            rows = []
            for row_data in sheets[0].get("data", []):
                for row_idx, row in enumerate(row_data.get("rowData", []), start=1):
                    new_row = []
                    for i, cell in enumerate(row.get("values", [])):
                        val = cell.get("formattedValue", "")
                        link = cell.get("hyperlink", "")
                        if link and i in (4, 7, 8, 9): # E, H, I, J columns can have hyperlinks replaced with their URL
                            val = link
                        col_letter = chr(65 + i)
                        img_url = cell_images.get((title_lower, row_idx, col_letter))
                        if img_url:
                            val = img_url
                        new_row.append(val)
                    while len(new_row) < 10:
                        col_letter = chr(65 + len(new_row))
                        img_url = cell_images.get((title_lower, row_idx, col_letter), "")
                        new_row.append(img_url)
                    rows.append(new_row)
            return rows
    except Exception as e:
        print("fetch_sheet_data hatası:", e)
    return []


def safe(row: list[str], idx: int) -> str:
    try:
        return row[idx].strip()
    except IndexError:
        return ""


# ─── URL Çözümleme (imgur.gg, ibb.co vb.) ────────────────────────────────────

_MEDIA_CACHE: dict[str, str] = {}

def resolve_media_url(url: str) -> str:
    """Sayfa linkini (örn: imgur.gg/f/..., ibb.co/...) doğrudan görsel/ses linkine çevirir."""
    if not url or url.lower() in ("-", "n/a", "none"):
        return ""
    if url.startswith("/static/"):
        return url
    if url in _MEDIA_CACHE:
        return _MEDIA_CACHE[url]

    resolved = url

    # 1. imgur.gg/f/<id>
    imgur_m = re.search(r"imgur\.gg/f/([a-zA-Z0-9_-]+)", url)
    if imgur_m:
        file_id = imgur_m.group(1)
        try:
            api_res = requests.get(f"https://imgur.gg/api/file/{file_id}", timeout=5)
            if api_res.status_code == 200:
                cdn = api_res.json().get("cdnUrl")
                if cdn:
                    resolved = cdn
        except Exception:
            pass

    # 2. pillows.su/f/<id> or pillowcase.su/f/<id>
    elif re.search(r"(?:pillows\.su|pillowcase\.su)/f/([a-zA-Z0-9_-]+)", url):
        p_m = re.search(r"(?:pillows\.su|pillowcase\.su)/f/([a-zA-Z0-9_-]+)", url)
        if p_m:
            resolved = f"https://api.pillows.su/api/download/{p_m.group(1)}"

    # 2. ibb.co/<id>
    elif "ibb.co/" in url and not re.search(r"\.(png|jpg|jpeg|webp|gif)$", url, re.I):
        try:
            res = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=5)
            if res.status_code == 200:
                og_m = re.search(r'property="og:image"\s+content="([^"]+)"', res.text)
                if og_m:
                    resolved = og_m.group(1)
        except Exception:
            pass

    _MEDIA_CACHE[url] = resolved
    return resolved


# ─── Tab Keşfi ve Eşleme (Google Sheets API ile) ──────────────────────────────

_TAB_CACHE: dict[str, dict[str, str]] = {}

def get_tab_map(sheet_id: str) -> dict[str, str]:
    """
    Google Sheets API ile tab isimlerini %100 doğrulukla bulur.
    """
    if sheet_id in _TAB_CACHE:
        return _TAB_CACHE[sheet_id]

    tab_map = {
        "unreleased": "0",
        "recent": "0",
        "best": "0",
        "art": "0",
        "tracklists": "0",
    }

    try:
        api_url = f"https://sheets.googleapis.com/v4/spreadsheets/{sheet_id}?key={GOOGLE_SHEETS_API_KEY}"
        r = requests.get(api_url, timeout=10)
        if r.status_code == 200:
            data = r.json()
            sheets = data.get("sheets", [])
            for s in sheets:
                p = s.get("properties", {})
                title = p.get("title", "")
                t_lower = title.strip().lower()
                
                if "unreleased" in t_lower or "songs" in t_lower:
                    if tab_map["unreleased"] == "0":
                        tab_map["unreleased"] = title
                elif "recent" in t_lower:
                    tab_map["recent"] = title
                elif "best of" in t_lower or "best" in t_lower:
                    tab_map["best"] = title
                elif t_lower == "art" or "art" in t_lower:
                    tab_map["art"] = title
                elif "tracklist" in t_lower:
                    tab_map["tracklists"] = title

            # Eğer unreleased bulunamadıysa ilk tabı unreleased yap
            if tab_map["unreleased"] == "0" and sheets:
                tab_map["unreleased"] = sheets[0].get("properties", {}).get("title", "0")

            _TAB_CACHE[sheet_id] = tab_map
            return tab_map
    except Exception as e:
        print("Tab API hatası:", e)

    _TAB_CACHE[sheet_id] = tab_map
    return tab_map


# ─── Otomatik Era Renkleri (Google Sheets API v4) ─────────────────────────────

_COLOR_CACHE: dict[str, dict[str, dict[str, str]]] = {}

def to_rgb_str(c_dict: dict | None, default: str = "rgb(24,24,24)") -> str:
    if not c_dict:
        return default
    r = int(c_dict.get("red", 0) * 255)
    g = int(c_dict.get("green", 0) * 255)
    b = int(c_dict.get("blue", 0) * 255)
    return f"rgb({r},{g},{b})"


def get_sheet_era_colors(sheet_id: str) -> dict[str, dict[str, str]]:
    """
    Google Sheets API v4 ile tablonun her era için tanımlanmış
    orijinal hücre arka plan (bg) ve yazı rengini (fg), ayrıca bannerdaki E kolonundan coverı çeker.
    """
    if sheet_id in _COLOR_CACHE:
        return _COLOR_CACHE[sheet_id]

    era_colors: dict[str, dict[str, str]] = {}
    xlsx_images = get_xlsx_cell_images(sheet_id)
    try:
        fields = "sheets.properties.title,sheets.data.rowData.values(formattedValue,hyperlink,userEnteredFormat.backgroundColor,userEnteredFormat.textFormat.foregroundColor,userEnteredValue.formulaValue)"
        api_url = (
            f"https://sheets.googleapis.com/v4/spreadsheets/{sheet_id}"
            f"?includeGridData=true&ranges=A:E&key={GOOGLE_SHEETS_API_KEY}&fields={fields}"
        )
        r = requests.get(api_url, timeout=20)
        if r.status_code == 200:
            data = r.json()
            for sheet in data.get("sheets", []):
                sheet_title = sheet.get("properties", {}).get("title", "").strip().lower()
                for r_data in sheet.get("data", []):
                    for row_idx, row in enumerate(r_data.get("rowData", []), start=1):
                        vals = row.get("values", [])
                        if len(vals) >= 2:
                            v0 = vals[0].get("formattedValue", "")
                            v1 = vals[1].get("formattedValue", "")
                            fmt = vals[0].get("userEnteredFormat", {}) or vals[1].get("userEnteredFormat", {})
                            bg = fmt.get("backgroundColor")
                            fg = fmt.get("textFormat", {}).get("foregroundColor")
                            
                            cover_url = ""
                            if len(vals) >= 5:
                                e_val = vals[4].get("userEnteredValue", {}).get("formulaValue", "")
                                if e_val and "http" in e_val:
                                    m = re.search(r'"(https?://[^"]+)"', e_val)
                                    if m:
                                        cover_url = m.group(1)
                                if not cover_url:
                                    e_link = vals[4].get("hyperlink", "")
                                    if e_link and "http" in e_link:
                                        cover_url = e_link
                                if not cover_url:
                                    e_fmt = vals[4].get("formattedValue", "")
                                    if e_fmt and "http" in e_fmt:
                                        cover_url = e_fmt
                            if not cover_url:
                                cover_url = xlsx_images.get((sheet_title, row_idx, "E"), "")

                            # Use same permissive keywords as is_stat_row to match banner rows
                            v0_lower = v0.lower()
                            keywords = ["full", "partial", "snippet", "session", "bounce", "tagged", "og file", "unavailable"]
                            has_keyword = any(k in v0_lower for k in keywords)
                            has_number = bool(re.search(r'\d+', v0_lower))

                            if (has_keyword and has_number and len(v0_lower) > 10):
                                era_name = v1.split("\n")[0].strip()
                                if era_name:
                                    if not cover_url:
                                        cover_url = xlsx_images.get((sheet_title, row_idx, "E"), "")
                                    era_colors[era_name.lower()] = {
                                        "bg": to_rgb_str(bg, "rgb(24,24,24)"),
                                        "fg": to_rgb_str(fg, "rgb(255,255,255)"),
                                        "banner_cover": cover_url
                                    }
    except Exception as e:
        print("Renk çekme hatası:", e)

    _COLOR_CACHE[sheet_id] = era_colors
    return era_colors



# ─── Kapak Resimleri: Yalnızca Era Banner (Col E) Kullanılır ─────────────────
# (Art sekmesinden kapak çekme kullanıcının isteği doğrultusunda tamamen kaldırıldı)


# ─── Tag Parsing ──────────────────────────────────────────────────────────────

TAG_COLORS = {
    "og file":       {"bg": "rgb(42,159,102)",   "fg": "rgb(255,255,255)"},
    "snippet":       {"bg": "rgb(153,0,0)",       "fg": "rgb(255,255,255)"},
    "low quality":   {"bg": "rgb(231,0,0)",       "fg": "rgb(255,255,255)"},
    "high quality":  {"bg": "rgb(39,78,19)",      "fg": "rgb(255,255,255)"},
    "cd quality":    {"bg": "rgb(76,175,80)",     "fg": "rgb(255,255,255)"},
    "partial":       {"bg": "rgb(191,144,0)",     "fg": "rgb(255,255,255)"},
    "beat only":     {"bg": "rgb(142,124,195)",   "fg": "rgb(255,255,255)"},
    "confirmed":     {"bg": "rgb(80,80,80)",      "fg": "rgb(200,200,200)"},
    "full":          {"bg": "rgb(7,55,99)",       "fg": "rgb(255,255,255)"},
    "lossless":      {"bg": "rgb(69,188,255)",    "fg": "rgb(0,0,0)"},
    "stem bounce":   {"bg": "rgb(189,123,194)",   "fg": "rgb(255,255,255)"},
    "recording":     {"bg": "rgb(0,0,0)",         "fg": "rgb(255,255,255)"},
    "throwaway":     {"bg": "rgb(100,60,20)",     "fg": "rgb(255,200,100)"},
    "production":    {"bg": "rgb(30,60,120)",     "fg": "rgb(150,200,255)"},
    "not available": {"bg": "rgb(60,60,60)",      "fg": "rgb(150,150,150)"},
    "unreleased":    {"bg": "rgb(50,50,50)",      "fg": "rgb(200,200,200)"},
    "released":      {"bg": "rgb(29,185,84)",     "fg": "rgb(255,255,255)"},
}

def make_tag(label: str) -> dict:
    key = label.strip().lower()
    colors = TAG_COLORS.get(key, {"bg": "rgb(60,60,60)", "fg": "rgb(200,200,200)"})
    return {"label": label.strip(), "bg": colors["bg"], "fg": colors["fg"]}


# ─── Track Name / Artists / Aliases Parser ───────────────────────────────────

def parse_name_field(raw: str) -> tuple[str, str, str]:
    """Col B → (title, artists, aliases)"""
    lines = raw.split("\n")
    first_line = lines[0].strip()
    extra_lines = [l.strip() for l in lines[1:] if l.strip()]

    m = re.search(r"(\((?:feat|prod|with|ft)\b.*)", first_line, re.IGNORECASE)
    if m:
        title = first_line[:m.start()].strip()
        rest = m.group(1).strip()
    else:
        title = first_line
        rest = ""

    combined = (rest + " " + " ".join(extra_lines)).strip()
    parens = re.findall(r"\([^)]*\)", combined)
    artists_list = []
    aliases_list = []

    for p in parens:
        if re.search(r"\((?:feat|prod|with|ft)\b", p, re.IGNORECASE):
            artists_list.append(p)
        else:
            cleaned = p.strip("()").strip()
            if cleaned:
                aliases_list.append(cleaned)

    artists = " ".join(artists_list)
    aliases = ", ".join(aliases_list)
    return title, artists, aliases


# ─── Satır Sınıflandırıcılar ──────────────────────────────────────────────────

def is_stat_row(row: list[str]) -> bool:
    col_a = safe(row, 0).lower()
    keywords = ["full", "partial", "snippet", "session", "bounce", "tagged", "og file", "unavailable"]
    has_keyword = any(k in col_a for k in keywords)
    has_number = bool(re.search(r'\d+', col_a))
    return has_keyword and has_number and len(col_a) > 10


def is_track_row(row: list[str]) -> bool:
    col_b = safe(row, 1)
    return bool(col_b) and not is_stat_row(row) and not is_subgroup_row(row)


def is_era_only_row(row: list[str]) -> bool:
    col_a = safe(row, 0)
    col_b = safe(row, 1)
    return bool(col_a) and not bool(col_b) and not is_stat_row(row)


def is_subgroup_row(row: list[str]) -> bool:
    col_a = safe(row, 0)
    col_b = safe(row, 1)
    has_track_info = any(bool(safe(row, i)) for i in range(3, len(row)))
    return not col_a and bool(col_b) and not has_track_info


# ─── Tracker CSV Parser ───────────────────────────────────────────────────────

def parse_tracker_csv(rows: list[list[str]]) -> list[dict]:
    eras: list[dict] = []
    current_era: dict | None = None
    last_era_name: str = ""

    if rows and rows[0] and safe(rows[0], 0).lower() in ("era", "album"):
        rows = rows[1:]

    for row in rows:
        if not any(c.strip() for c in row):
            continue

        col_a = safe(row, 0)
        col_b = safe(row, 1)
        col_c = safe(row, 2)
        col_d = safe(row, 3)
        col_e = safe(row, 4)
        col_f = safe(row, 5)

        # ── 1. Era Banner Satırı (Stat Satırı) ──
        if is_stat_row(row):
            era_name = col_b.split("\n")[0].strip() if col_b else col_a
            desc = col_f if col_f and "notes about" not in col_f.lower() else col_c
            cover = resolve_media_url(col_e) if ("http" in col_e or "/static/" in col_e) else ""

            current_era = {
                "era": era_name,
                "description": desc,
                "cover": cover,
                "bg_color": "rgb(24,24,24)",
                "text_color": "rgb(255,255,255)",
                "tracks": [],
            }
            eras.append(current_era)
            last_era_name = era_name
            continue

        # ── 2. Era-Only Satırı ──
        if is_era_only_row(row):
            if any(k in col_a.lower() for k in ("links", "update notes", "tracker guidelines")):
                continue

            desc = col_f if col_f and "notes about" not in col_f.lower() else col_c
            cover = resolve_media_url(col_e) if ("http" in col_e or "/static/" in col_e) else ""

            current_era = {
                "era": col_a,
                "description": desc,
                "cover": cover,
                "bg_color": "rgb(24,24,24)",
                "text_color": "rgb(255,255,255)",
                "tracks": [],
            }
            eras.append(current_era)
            last_era_name = col_a
            continue

        # ── 3. Sub-group label ──
        if is_subgroup_row(row):
            if current_era is not None:
                current_era["tracks"].append({
                    "name": col_b,
                    "artists": "",
                    "notes": col_c,
                    "tags": [],
                    "aliases": "",
                    "length": "",
                    "leak_date": "",
                    "link": "",
                    "is_subera": True,
                })
            continue

        # ── 4. Track Satırı ──
        if is_track_row(row):
            if col_a and col_a != last_era_name:
                current_era = {
                    "era": col_a,
                    "description": "",
                    "cover": "",
                    "bg_color": "rgb(24,24,24)",
                    "text_color": "rgb(255,255,255)",
                    "tracks": [],
                }
                eras.append(current_era)
                last_era_name = col_a

            if current_era is None:
                current_era = {
                    "era": col_a or "Misc",
                    "description": "",
                    "cover": "",
                    "bg_color": "rgb(24,24,24)",
                    "text_color": "rgb(255,255,255)",
                    "tracks": [],
                }
                eras.append(current_era)
                last_era_name = col_a

            # Dynamic link detection: scan columns from right to left
            track_link = ""
            for idx in range(len(row) - 1, 2, -1):
                val = safe(row, idx)
                if any(h in val.lower() for h in ("http://", "https://", "imgur.gg", "pillows.su", "pillowcase.su")):
                    track_link = val.split("\n")[0].strip()
                    break

            # Dynamic tags detection
            tags = []
            for idx in range(3, len(row)):
                val = safe(row, idx)
                if not val or val.lower() in ("", "n/a", "-"):
                    continue
                if any(h in val.lower() for h in ("http://", "https://", "imgur.gg", "pillows.su", "pillowcase.su")):
                    continue
                if re.match(r"^(\d{1,2}:\d{2}|~?\d+:\d+|\d{4}|(?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)\b)", val, re.I):
                    continue
                val_lower = val.lower()
                if val_lower in TAG_COLORS or any(k in val_lower for k in ("released", "unreleased", "throwaway", "snippet", "quality", "full", "demo", "og", "lossless", "recording", "partial", "confirmed")):
                    tags.append(make_tag(val))

            # Length detection
            length = ""
            for idx in [3, 4]:
                val = safe(row, idx)
                if re.match(r"^~?\d+:\d+", val):
                    length = val
                    break

            # Date detection
            leak_date = ""
            for idx in [4, 5, 6]:
                val = safe(row, idx)
                if re.search(r"\b(19\d\d|20\d\d|(?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec))\b", val, re.I):
                    leak_date = val
                    break

            title, artists, aliases = parse_name_field(col_b)

            current_era["tracks"].append({
                "name": title,
                "artists": artists,
                "notes": col_c,
                "tags": tags,
                "aliases": aliases,
                "length": length or col_d,
                "leak_date": leak_date or col_e,
                "link": track_link,
                "is_subera": False,
            })

    return eras


# ─── Art & Tracklists CSV Parsers ─────────────────────────────────────────────

def parse_art_csv(rows: list[list[str]]) -> list[dict]:
    eras_map: dict[str, dict] = {}
    if rows and rows[0] and safe(rows[0], 0).lower() in ("era", "album"):
        rows = rows[1:]

    for row in rows:
        if not any(c.strip() for c in row):
            continue
        era_name = safe(row, 0)
        name = safe(row, 1)
        notes = safe(row, 2)
        designer = safe(row, 3)
        art_type = safe(row, 4)
        project_type = safe(row, 6)
        use = safe(row, 7)
        link = safe(row, 8)

        if not era_name:
            continue

        img_url = ""
        for cell in [link, safe(row, 5)]:
            if "http" in cell:
                m = re.search(r"https?://[^\s\"]+", cell)
                if m:
                    img_url = resolve_media_url(m.group(0))
                    break

        if era_name not in eras_map:
            eras_map[era_name] = {"era": era_name, "artworks": []}

        eras_map[era_name]["artworks"].append({
            "title": name or "Artwork",
            "notes": notes,
            "designer": designer,
            "type": project_type or art_type,
            "use": use,
            "image": img_url,
            "link": link or img_url,
        })

    return list(eras_map.values())


def parse_tracklists_csv(rows: list[list[str]]) -> list[dict]:
    eras_map: dict[str, dict] = {}
    if rows and rows[0] and safe(rows[0], 0).lower() in ("era", "album"):
        rows = rows[1:]

    for row in rows:
        if not any(c.strip() for c in row):
            continue
        era_name = safe(row, 0)
        name = safe(row, 1)
        tracklist = safe(row, 2)
        date_made = safe(row, 4)
        quality = safe(row, 5)
        source = safe(row, 6)
        link = safe(row, 7)

        if not era_name or "completely unheard" in era_name.lower():
            continue

        img_url = ""
        if "http" in link:
            m = re.search(r"https?://[^\s\"]+", link)
            if m:
                img_url = resolve_media_url(m.group(0))

        if era_name not in eras_map:
            eras_map[era_name] = {"era": era_name, "tracklists": []}

        eras_map[era_name]["tracklists"].append({
            "name": name,
            "tracklist": tracklist,
            "date": date_made,
            "quality": quality,
            "source": source,
            "image": img_url,
            "link": link,
        })

    return list(eras_map.values())


# ─── Renk ve Veri Birleştirme ─────────────────────────────────────────────────

_FALLBACK_PALETTE = [
    {"bg": "rgb(166,27,0)",   "fg": "rgb(246,178,107)"},
    {"bg": "rgb(87,19,20)",   "fg": "rgb(238,162,1)"},
    {"bg": "rgb(41,10,6)",    "fg": "rgb(184,146,20)"},
    {"bg": "rgb(65,34,87)",   "fg": "rgb(234,80,171)"},
    {"bg": "rgb(195,205,204)","fg": "rgb(241,0,1)"},
    {"bg": "rgb(55,120,196)", "fg": "rgb(218,185,87)"},
    {"bg": "rgb(104,110,250)","fg": "rgb(255,255,255)"},
]

def apply_styling(sheet_id: str, eras: list[dict]):
    """
    Era bannerındaki (Col E) orijinal kapakları ve hücre renklerini uygular.
    Art sekmesinden kapak çekme tamamen kaldırıldı.
    """
    colors = get_sheet_era_colors(sheet_id)

    for i, era in enumerate(eras):
        k = era["era"].strip().lower()
        # 1. Renkler
        if k in colors:
            era["bg_color"] = colors[k]["bg"]
            era["text_color"] = colors[k]["fg"]
            
            banner_cover = colors[k].get("banner_cover")
            if banner_cover:
                era["cover"] = resolve_media_url(banner_cover)
        else:
            fb = _FALLBACK_PALETTE[i % len(_FALLBACK_PALETTE)]
            era["bg_color"] = fb["bg"]
            era["text_color"] = fb["fg"]


def get_tracker_data(sheet_id: str, tab_title: str) -> list[dict]:
    rows = fetch_sheet_data(sheet_id, tab_title)
    eras = parse_tracker_csv(rows)
    eras = [e for e in eras if e["tracks"]]
    apply_styling(sheet_id, eras)
    return eras


# ─── Routes ───────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/<sheet_id>")
def sheet_root(sheet_id):
    if not re.fullmatch(r"[a-zA-Z0-9_-]+", sheet_id):
        abort(404)
    return redirect(url_for("sheet_unreleased", sheet_id=sheet_id))


@app.route("/sh/<sheet_id>")
def sheet_sh_redirect(sheet_id):
    return redirect(url_for("sheet_unreleased", sheet_id=sheet_id))


@app.route("/<sheet_id>/unreleased")
def sheet_unreleased(sheet_id):
    try:
        tabs = get_tab_map(sheet_id)
        data = get_tracker_data(sheet_id, tabs["unreleased"])
    except Exception as e:
        return render_template("error.html", error=str(e), sheet_id=sheet_id)
    return render_template(
        "tracker.html",
        sheet_id=sheet_id,
        section="unreleased",
        eras=data,
        page_title="Unreleased",
        tabs=tabs,
    )


@app.route("/<sheet_id>/recent")
def sheet_recent(sheet_id):
    try:
        tabs = get_tab_map(sheet_id)
        data = get_tracker_data(sheet_id, tabs["recent"])
    except Exception as e:
        return render_template("error.html", error=str(e), sheet_id=sheet_id)
    return render_template(
        "recent.html",
        sheet_id=sheet_id,
        section="recent",
        eras=data,
        page_title="Recent",
        tabs=tabs,
    )


@app.route("/<sheet_id>/best")
def sheet_best(sheet_id):
    try:
        tabs = get_tab_map(sheet_id)
        data = get_tracker_data(sheet_id, tabs["best"])
    except Exception as e:
        return render_template("error.html", error=str(e), sheet_id=sheet_id)
    return render_template(
        "tracker.html",
        sheet_id=sheet_id,
        section="best",
        eras=data,
        page_title="Best Of",
        tabs=tabs,
    )


@app.route("/<sheet_id>/art")
def sheet_art(sheet_id):
    try:
        tabs = get_tab_map(sheet_id)
        rows = fetch_sheet_data(sheet_id, tabs["art"])
        data = parse_art_csv(rows)
        apply_styling(sheet_id, data)
    except Exception as e:
        return render_template("error.html", error=str(e), sheet_id=sheet_id)
    return render_template(
        "art.html",
        sheet_id=sheet_id,
        section="art",
        eras=data,
        page_title="Art",
        tabs=tabs,
    )


@app.route("/<sheet_id>/tracklists")
def sheet_tracklists(sheet_id):
    try:
        tabs = get_tab_map(sheet_id)
        rows = fetch_sheet_data(sheet_id, tabs["tracklists"])
        data = parse_tracklists_csv(rows)
        apply_styling(sheet_id, data)
    except Exception as e:
        return render_template("error.html", error=str(e), sheet_id=sheet_id)
    return render_template(
        "tracklists.html",
        sheet_id=sheet_id,
        section="tracklists",
        eras=data,
        page_title="Tracklists",
        tabs=tabs,
    )


@app.route("/<sheet_id>/tab/<path:tab_name>")
def sheet_custom_tab(sheet_id, tab_name):
    try:
        tabs = get_tab_map(sheet_id)
        rows = fetch_sheet_data(sheet_id, tab_name)
        t_lower = tab_name.lower()
        if "art" in t_lower:
            data = parse_art_csv(rows)
            apply_styling(sheet_id, data)
            return render_template("art.html", sheet_id=sheet_id, section=tab_name, eras=data, page_title=tab_name, tabs=tabs)
        elif "tracklist" in t_lower:
            data = parse_tracklists_csv(rows)
            apply_styling(sheet_id, data)
            return render_template("tracklists.html", sheet_id=sheet_id, section=tab_name, eras=data, page_title=tab_name, tabs=tabs)
        elif "recent" in t_lower:
            data = parse_tracker_csv(rows)
            apply_styling(sheet_id, data)
            return render_template("recent.html", sheet_id=sheet_id, section=tab_name, eras=data, page_title=tab_name, tabs=tabs)
        else:
            data = parse_tracker_csv(rows)
            apply_styling(sheet_id, data)
            return render_template("tracker.html", sheet_id=sheet_id, section=tab_name, eras=data, page_title=tab_name, tabs=tabs)
    except Exception as e:
        return render_template("error.html", error=str(e), sheet_id=sheet_id)


# ─── API ──────────────────────────────────────────────────────────────────────

@app.route("/go", methods=["POST"])
def go():
    raw = request.form.get("sheet_url", "").strip()
    sheet_id = extract_sheet_id(raw)
    if not sheet_id:
        return render_template("index.html", error="Invalid Google Sheets URL or ID.")
    return redirect(url_for("sheet_unreleased", sheet_id=sheet_id))


@app.route("/api/resolve")
def api_resolve():
    url = request.args.get("url", "").strip()
    return jsonify({"resolved": resolve_media_url(url)})


if __name__ == "__main__":
    app.run(debug=True, port=3131)
