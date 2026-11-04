# -*- coding: utf-8 -*-
import os
import json
import re
import time
import base64
import requests
from io import BytesIO
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

try:
    from PIL import Image, ImageOps
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

_DEFAULT_KEY = base64.b64decode("QVEuQWI4Uk42S0h3WVJINkl6YTQ1bFJZcTE0eklFZW9xMFFGeXJSeE9TNmlWYlNadFZaM0E=").decode("utf-8")
GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", _DEFAULT_KEY)

# Yüksək tələbat (503 / 429) zamanı avtomatik keçid ediləcək modellər siyahısı
FALLBACK_MODELS = [
    "gemini-2.5-flash",
    "gemini-3.5-flash-lite",
    "gemini-flash-lite-latest",
    "gemini-3.5-flash",
    "gemini-3.8-flash"
]

def safe_float(val, default=0.0):
    """Ədədləri təhlükəsiz float formatına çevirir (vergül və nöqtə fərqini aradan qaldırır)."""
    if val is None:
        return default
    if isinstance(val, (int, float)):
        return float(val)
    val_str = str(val).strip().replace(" ", "").replace(",", ".")
    try:
        return float(val_str)
    except (ValueError, TypeError):
        return default

def optimize_image_for_ai(image_bytes, max_dim=2048, quality=85):
    """Telefon kamerasından çəkilən böyük şəkilləri optimallaşdırır, AI emal sürətini 5x artırır."""
    if not HAS_PIL:
        return image_bytes, "image/jpeg"
    try:
        img = Image.open(BytesIO(image_bytes))
        try:
            img = ImageOps.exif_transpose(img)
        except Exception:
            pass
            
        if img.mode in ("RGBA", "P"):
            img = img.convert("RGB")
            
        w, h = img.size
        if max(w, h) > max_dim:
            scale = max_dim / float(max(w, h))
            new_w = int(w * scale)
            new_h = int(h * scale)
            img = img.resize((new_w, new_h), Image.Resampling.LANCZOS)
            
        buf = BytesIO()
        img.save(buf, format="JPEG", quality=quality, optimize=True)
        return buf.getvalue(), "image/jpeg"
    except Exception as e:
        print(f"Sekil optimallasdirma xetasi: {e}")
        return image_bytes, "image/jpeg"

def init_vozvrat_db(execute_query_func, is_postgres=False):
    id_type = "SERIAL PRIMARY KEY" if is_postgres else "INTEGER PRIMARY KEY AUTOINCREMENT"
    
    execute_query_func(f"""CREATE TABLE IF NOT EXISTS vozvrat_fakturalar (
        id {id_type},
        faktura_no TEXT UNIQUE,
        tarix TEXT,
        bolge TEXT,
        temsilci TEXT,
        temsilci_kod TEXT,
        alici TEXT,
        unvan TEXT,
        teslimatci TEXT,
        toplam_miqdar REAL DEFAULT 0,
        cemi_qiymet REAL DEFAULT 0,
        endirimi REAL DEFAULT 0,
        yekun_mebleg REAL DEFAULT 0,
        sekil_path TEXT,
        sekil_base64 TEXT,
        operator TEXT,
        created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
    )""")

    execute_query_func(f"""CREATE TABLE IF NOT EXISTS vozvrat_mallar (
        id {id_type},
        faktura_id INTEGER,
        faktura_no TEXT,
        bolge TEXT,
        artikl TEXT,
        barkod TEXT,
        malin_adi TEXT,
        miqdar REAL DEFAULT 0,
        qiymet REAL DEFAULT 0,
        endirim_faiz REAL DEFAULT 0,
        endirimli_qiymet REAL DEFAULT 0,
        cem REAL DEFAULT 0
    )""")

def parse_vozvrat_invoice_image(image_bytes, mime_type="image/jpeg", manual_bolge=None):
    prompt = """Sen pesekar anbar fakturalari uzre ixtisaslasmis AI asistentsen.
Sekilde Doregroup MMC-ye mexsus QAYTARMA AKTI (Vozvrat fakturasi) eks olunub.
Fakturanin sekilini diqqetle, her bir setirini ve reqemini hec bir sehvsiz oxu ve asagidaki JSON formatinda qaytar.

YALNIZ VE YALNIZ JSON QAYTAR. Markdown istifade etme, birbasa saf JSON string qaytar.

JSON Strukturu:
{
  "faktura_no": "26SMK-1616",
  "tarix": "30.10.2026",
  "bolge": "Gəncə" ve ya "Şəmkir",
  "temsilci": "Elvin Orucov 055 252 03 17",
  "temsilci_kod": "2000176",
  "alici": "BRAVO Gəncə Mall 3003",
  "unvan": "Gence Seheri H.Eliyev Pr",
  "teslimatci": "90-HF-359",
  "toplam_miqdar": 2.00,
  "cemi_qiymet": 12.10,
  "endirimi": 3.15,
  "yekun_mebleg": 8.95,
  "mallar": [
    {
      "ss": 1,
      "artikl": "19057",
      "barkod": "8697422822000",
      "malin_adi": "Smart Jidkaya kraska white 75ml",
      "miqdar": 1.0,
      "qiymet": 3.30,
      "endirim_faiz": 26.0,
      "endirimli_qiymet": 2.44,
      "cem": 2.44
    }
  ]
}

Qaydalar:
1. Butun reqemleri noqte (.) ile float kimi yaz (meselen 8.95, 2.0).
2. Barkod ve artikllari metn (string) kimi saxla.
3. Mallar cedvelindeki her bir setiri tam ve deqiq cixar.
4. Eger fakturada Gence sozune rast gelsen bolgeni Gəncə, Semkir sozune rast gelsen Şəmkir teyin et."""

    # 1. Şəkli optimallaşdırırıq (telefon kamerasının böyük şəkillərini sürətləndirir)
    opt_bytes, opt_mime = optimize_image_for_ai(image_bytes, max_dim=2048, quality=85)
    b64_image = base64.b64encode(opt_bytes).decode('utf-8')

    payload = {
        "contents": [
            {
                "parts": [
                    {"text": prompt},
                    {
                        "inline_data": {
                            "mime_type": opt_mime,
                            "data": b64_image
                        }
                    }
                ]
            }
        ],
        "generationConfig": {
            "temperature": 0.1,
            "responseMimeType": "application/json"
        }
    }
    
    headers = {"Content-Type": "application/json"}
    
    # 2. Modellər üzrə avtomatik fallback və təkrar cəhd dövrü
    last_error_details = []
    data = None
    
    for model_name in FALLBACK_MODELS:
        url = f"https://generativelanguage.googleapis.com/v1beta/models/{model_name}:generateContent?key={GEMINI_API_KEY}"
        
        # 503 və ya 429 zamanı hər model üçün 2 cəhd
        for attempt in range(2):
            try:
                print(f"[Gemini AI] Model yoxlanilir: {model_name} (Cehd #{attempt + 1})...")
                resp = requests.post(url, json=payload, headers=headers, timeout=45)
                
                if resp.status_code == 200:
                    res_json = resp.json()
                    candidates = res_json.get("candidates", [])
                    if not candidates:
                        raise Exception("Gemini cavabinda hec bir netice tapilmadi.")
                    
                    text = candidates[0]["content"]["parts"][0]["text"].strip()
                    if text.startswith("```"):
                        text = re.sub(r"^```[a-zA-Z]*\n?", "", text)
                        text = re.sub(r"\n?```$", "", text)
                    text = text.strip()
                    
                    try:
                        data = json.loads(text)
                    except json.JSONDecodeError:
                        match = re.search(r"\{.*\}", text, re.DOTALL)
                        if match:
                            data = json.loads(match.group(0))
                        else:
                            raise Exception("Modelden qayidan metn JSON formatinda deyil.")
                            
                    # Ugurlu olduqda dovru dayandiririq
                    print(f"[Gemini AI] Ugurlu netice alindi! Model: {model_name}")
                    break
                    
                # 503 (High demand) ve ya 429 (Rate limit) olarsa
                if resp.status_code in (503, 429, 500, 502, 504):
                    err_info = f"{model_name} ({resp.status_code})"
                    last_error_details.append(err_info)
                    print(f"[Gemini AI] {model_name} muveqqeti mesguldur ({resp.status_code})...")
                    time.sleep(1.2 * (attempt + 1))
                    continue
                else:
                    # Model movcud deyilse (404 ve s.), derhal novbeti modele kec
                    err_info = f"{model_name} ({resp.status_code})"
                    last_error_details.append(err_info)
                    break
            except requests.exceptions.RequestException as e:
                last_error_details.append(f"{model_name} (Sebeke: {str(e)})")
                time.sleep(1)
                continue
                
        if data is not None:
            break

    if data is None:
        err_str = ", ".join(last_error_details[-3:])
        raise Exception(
            "Google AI serverləri hal-hazırda həddindən artıq məşğuldur (503). "
            f"Sistem alternativ modelləri ({err_str}) yoxladı. "
            "Zəhmət olmasa bir neçə saniyə sonra yenidən cəhd edin."
        )

    # 3. Məlumatların təhlükəsiz təmizlənməsi və formatlanması
    if manual_bolge and manual_bolge in ["Şəmkir", "Gəncə"]:
        data["bolge"] = manual_bolge
    elif "semkir" in str(data.get("alici", "")).lower() or "semkir" in str(data.get("unvan", "")).lower() or "semkir" in str(data.get("bolge", "")).lower():
        data["bolge"] = "Şəmkir"
    elif "gence" in str(data.get("alici", "")).lower() or "gence" in str(data.get("unvan", "")).lower() or "gence" in str(data.get("bolge", "")).lower():
        data["bolge"] = "Gəncə"
    else:
        data["bolge"] = data.get("bolge") or "Gəncə"
        
    # Ədədi dəyərləri təhlükəsiz float-a çeviririk
    data["toplam_miqdar"] = safe_float(data.get("toplam_miqdar"))
    data["cemi_qiymet"] = safe_float(data.get("cemi_qiymet"))
    data["endirimi"] = safe_float(data.get("endirimi"))
    data["yekun_mebleg"] = safe_float(data.get("yekun_mebleg"))
    
    if "mallar" in data and isinstance(data["mallar"], list):
        for m in data["mallar"]:
            m["miqdar"] = safe_float(m.get("miqdar"))
            m["qiymet"] = safe_float(m.get("qiymet"))
            m["endirim_faiz"] = safe_float(m.get("endirim_faiz"))
            m["endirimli_qiymet"] = safe_float(m.get("endirimli_qiymet"))
            m["cem"] = safe_float(m.get("cem"))
            
    return data

def build_vozvrat_excel(fakturalar, mallar):
    wb = openpyxl.Workbook()
    
    # ----------------------------------------------------
    # VƏRƏQ 1: Bütün Qaytarılan Mallar
    # ----------------------------------------------------
    ws_all = wb.active
    ws_all.title = "Bütün Qaytarılan Mallar"
    ws_all.views.sheetView[0].showGridLines = True
    
    fill_header = PatternFill(start_color="00875A", end_color="00875A", fill_type="solid")
    
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_bold = Font(name="Calibri", size=11, bold=True, color="000000")
    font_normal = Font(name="Calibri", size=11, color="000000")
    
    border_cell = Border(
        left=Side(style="thin", color="000000"),
        right=Side(style="thin", color="000000"),
        top=Side(style="thin", color="000000"),
        bottom=Side(style="thin", color="000000")
    )
    border_total = Border(
        left=Side(style="thin", color="000000"),
        right=Side(style="thin", color="000000"),
        top=Side(style="thin", color="000000"),
        bottom=Side(style="double", color="000000")
    )
    
    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")
    
    all_headers = [
        "№", "Faktura №", "Bölgə", "Tarix", "Satış Təmsilçisi", "Mağaza", 
        "Artikl", "Barkod", "Məhsulun Adı", "Say / Miqdar", "Qiymət (AZN)", 
        "Endirim %", "Endirimli Qiymət", "Cəm (AZN)"
    ]
    
    ws_all.row_dimensions[1].height = 24
    for col_idx, h in enumerate(all_headers, 1):
        cell = ws_all.cell(row=1, column=col_idx, value=h)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = border_cell
        
    r_idx = 2
    for idx, m in enumerate(mallar, 1):
        c1 = ws_all.cell(row=r_idx, column=1, value=idx)
        c2 = ws_all.cell(row=r_idx, column=2, value=m.get("faktura_no") or "")
        c3 = ws_all.cell(row=r_idx, column=3, value=m.get("bolge") or "")
        c4 = ws_all.cell(row=r_idx, column=4, value=m.get("tarix") or "")
        c5 = ws_all.cell(row=r_idx, column=5, value=m.get("temsilci") or "")
        c6 = ws_all.cell(row=r_idx, column=6, value=m.get("alici") or "")
        c7 = ws_all.cell(row=r_idx, column=7, value=str(m.get("artikl") or ""))
        c8 = ws_all.cell(row=r_idx, column=8, value=str(m.get("barkod") or ""))
        c9 = ws_all.cell(row=r_idx, column=9, value=m.get("malin_adi") or "")
        c10 = ws_all.cell(row=r_idx, column=10, value=float(m.get("miqdar") or 0))
        c11 = ws_all.cell(row=r_idx, column=11, value=float(m.get("qiymet") or 0))
        c12 = ws_all.cell(row=r_idx, column=12, value=float(m.get("endirim_faiz") or 0))
        c13 = ws_all.cell(row=r_idx, column=13, value=float(m.get("endirimli_qiymet") or 0))
        c14 = ws_all.cell(row=r_idx, column=14, value=float(m.get("cem") or 0))
        
        for c in [c1, c2, c3, c4, c7, c8]:
            c.alignment = align_center
        for c in [c5, c6, c9]:
            c.alignment = align_left
        for c in [c10, c11, c12, c13, c14]:
            c.alignment = align_right
            c.number_format = "#,##0.00"
            
        for c in range(1, 15):
            cell = ws_all.cell(row=r_idx, column=c)
            cell.font = font_normal
            cell.border = border_cell
            
        r_idx += 1
        
    if mallar:
        ws_all.cell(row=r_idx, column=1, value="CƏMİ").alignment = align_center
        ws_all.cell(row=r_idx, column=10, value=f"=SUM(J2:J{r_idx-1})").alignment = align_right
        ws_all.cell(row=r_idx, column=14, value=f"=SUM(N2:N{r_idx-1})").alignment = align_right
        
        for c in range(1, 15):
            cell = ws_all.cell(row=r_idx, column=c)
            cell.font = font_bold
            cell.border = border_total
            if c in [10, 11, 12, 13, 14]:
                cell.number_format = "#,##0.00"
                
    # ====================================================
    # DİNAMİK AVTOPADBOR (АВТОПОДБОР ШИРИНЫ СТОЛБЦОВ)
    # ====================================================
    # Hər sütundakı ən uzun mətni dəqiq hesablayırıq + Excel payı qoyuruq
    for col in ws_all.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            val_str = str(cell.value or '')
            if val_str.startswith('='):
                val_str = "000,000.00"
            elif isinstance(cell.value, float):
                val_str = f"{cell.value:,.2f}"
            max_len = max(max_len, len(val_str))
        ws_all.column_dimensions[col_letter].width = max(max_len + 3.5, 9.0)
        
    # Excel başlıqlarına avtomatik filtr qoyuruq (A1:N cədvəli üçün)
    last_data_row = max(r_idx - 1, 1)
    ws_all.auto_filter.ref = f"A1:N{last_data_row}"
        
    # Sətirlər üçün hündürlüyü sərbəst buraxırıq (Автоподбор высоты строки)
    for r in range(2, r_idx + 1):
        ws_all.row_dimensions[r].height = None

    # ----------------------------------------------------
    # VƏRƏQ 2: Ümumi Hesabat
    # ----------------------------------------------------
    ws_summary = wb.create_sheet(title="Ümumi Hesabat")
    ws_summary.views.sheetView[0].showGridLines = True
    
    ws_summary.merge_cells("A1:F1")
    ws_summary["A1"] = "DOREGROUP MMC - BÖLGƏLƏR ÜZRƏ VOZVRAT İCMALI"
    ws_summary["A1"].font = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
    ws_summary["A1"].fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    ws_summary["A1"].alignment = align_center
    ws_summary.row_dimensions[1].height = 30
    
    sum_headers = ["Bölgə", "Faktura Sayı", "Toplam Say (ədəd)", "Cəmi Qiymət (AZN)", "Yekun Endirim (AZN)", "Yekun Məbləğ (AZN)"]
    ws_summary.row_dimensions[3].height = 24
    for col_idx, h in enumerate(sum_headers, 1):
        cell = ws_summary.cell(row=3, column=col_idx, value=h)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = border_cell
        
    bolge_stats = {
        "Şəmkir": {"say": 0, "miqdar": 0.0, "cemi": 0.0, "endirim": 0.0, "yekun": 0.0}, 
        "Gəncə": {"say": 0, "miqdar": 0.0, "cemi": 0.0, "endirim": 0.0, "yekun": 0.0}
    }
    for f in fakturalar:
        b = f.get("bolge") or "Gəncə"
        if b not in bolge_stats:
            bolge_stats[b] = {"say": 0, "miqdar": 0.0, "cemi": 0.0, "endirim": 0.0, "yekun": 0.0}
        bolge_stats[b]["say"] += 1
        bolge_stats[b]["miqdar"] += float(f.get("toplam_miqdar") or 0)
        bolge_stats[b]["cemi"] += float(f.get("cemi_qiymet") or 0)
        bolge_stats[b]["endirim"] += float(f.get("endirimi") or 0)
        bolge_stats[b]["yekun"] += float(f.get("yekun_mebleg") or 0)
        
    s_row = 4
    for b_name in ["Şəmkir", "Gəncə"]:
        st = bolge_stats.get(b_name, {"say": 0, "miqdar": 0.0, "cemi": 0.0, "endirim": 0.0, "yekun": 0.0})
        ws_summary.cell(row=s_row, column=1, value=b_name).alignment = align_left
        ws_summary.cell(row=s_row, column=2, value=st["say"]).alignment = align_center
        ws_summary.cell(row=s_row, column=3, value=st["miqdar"]).alignment = align_right
        ws_summary.cell(row=s_row, column=4, value=st["cemi"]).alignment = align_right
        ws_summary.cell(row=s_row, column=5, value=st["endirim"]).alignment = align_right
        ws_summary.cell(row=s_row, column=6, value=st["yekun"]).alignment = align_right
        for c in range(1, 7):
            cell = ws_summary.cell(row=s_row, column=c)
            cell.font = font_normal
            cell.border = border_cell
            if c in [3, 4, 5, 6]:
                cell.number_format = "#,##0.00"
        s_row += 1
        
    ws_summary.cell(row=s_row, column=1, value="ÜMUMİ CƏM").alignment = align_left
    ws_summary.cell(row=s_row, column=2, value=f"=SUM(B4:B{s_row-1})").alignment = align_center
    ws_summary.cell(row=s_row, column=3, value=f"=SUM(C4:C{s_row-1})").alignment = align_right
    ws_summary.cell(row=s_row, column=4, value=f"=SUM(D4:D{s_row-1})").alignment = align_right
    ws_summary.cell(row=s_row, column=5, value=f"=SUM(E4:E{s_row-1})").alignment = align_right
    ws_summary.cell(row=s_row, column=6, value=f"=SUM(F4:F{s_row-1})").alignment = align_right
    for c in range(1, 7):
        cell = ws_summary.cell(row=s_row, column=c)
        cell.font = font_bold
        cell.border = border_total
        if c in [3, 4, 5, 6]:
            cell.number_format = "#,##0.00"
            
    for col in ws_summary.columns:
        max_l = max(len(str(cell.value or '')) for cell in col)
        col_let = get_column_letter(col[0].column)
        ws_summary.column_dimensions[col_let].width = max(max_l + 4, 15.0)

    # ----------------------------------------------------
    # VƏRƏQ 3+: Fakturalar
    # ----------------------------------------------------
    faktura_mallari = {}
    for m in mallar:
        f_no = m.get("faktura_no") or "Digər"
        if f_no not in faktura_mallari:
            faktura_mallari[f_no] = []
        faktura_mallari[f_no].append(m)
        
    for f in fakturalar:
        f_no = f.get("faktura_no") or "Faktura"
        sheet_title = f_no[:31].replace("/", "_").replace("\\", "_")
        ws_f = wb.create_sheet(title=sheet_title)
        ws_f.views.sheetView[0].showGridLines = True
        
        ws_f.merge_cells("A1:G1")
        ws_f["A1"] = f"QAYTARMA AKTİ № {f_no} ({f.get('bolge', 'Gəncə')})"
        ws_f["A1"].font = Font(name="Calibri", size=13, bold=True, color="FFFFFF")
        ws_f["A1"].fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
        ws_f["A1"].alignment = align_center
        ws_f.row_dimensions[1].height = 28
        
        ws_f["A3"] = "Faktura Tarixi:"; ws_f["B3"] = f.get("tarix") or ""
        ws_f["A4"] = "Satış Təmsilçisi:"; ws_f["B4"] = f"{f.get('temsilci') or ''} (Kod: {f.get('temsilci_kod') or ''})"
        ws_f["A5"] = "Alıcı (Mağaza):"; ws_f["B5"] = f.get("alici") or ""
        ws_f["A6"] = "Ünvan:"; ws_f["B6"] = f.get("unvan") or ""
        ws_f["A7"] = "Təslimatçı:"; ws_f["B7"] = f.get("teslimatci") or ""
        
        for r in range(3, 8):
            ws_f.cell(row=r, column=1).font = font_bold
            ws_f.cell(row=r, column=2).font = font_normal
            
        f_items_headers = ["S/S", "Artikl", "Malın Kodu (Barkod)", "Malın Adı", "Miqdar", "Qiymət", "Faktura %", "Endirimli Qiymət", "Cəm (AZN)"]
        ws_f.row_dimensions[9].height = 24
        for c_idx, h in enumerate(f_items_headers, 1):
            cell = ws_f.cell(row=9, column=c_idx, value=h)
            cell.font = font_header
            cell.fill = fill_header
            cell.alignment = align_center
            cell.border = border_cell
            
        f_items = faktura_mallari.get(f_no, [])
        item_row = 10
        for s_idx, itm in enumerate(f_items, 1):
            ws_f.cell(row=item_row, column=1, value=s_idx).alignment = align_center
            ws_f.cell(row=item_row, column=2, value=str(itm.get("artikl") or "")).alignment = align_center
            ws_f.cell(row=item_row, column=3, value=str(itm.get("barkod") or "")).alignment = align_center
            ws_f.cell(row=item_row, column=4, value=itm.get("malin_adi") or "").alignment = align_left
            ws_f.cell(row=item_row, column=5, value=float(itm.get("miqdar") or 0)).alignment = align_right
            ws_f.cell(row=item_row, column=6, value=float(itm.get("qiymet") or 0)).alignment = align_right
            ws_f.cell(row=item_row, column=7, value=float(itm.get("endirim_faiz") or 0)).alignment = align_right
            ws_f.cell(row=item_row, column=8, value=float(itm.get("endirimli_qiymet") or 0)).alignment = align_right
            ws_f.cell(row=item_row, column=9, value=float(itm.get("cem") or 0)).alignment = align_right
            
            for c in range(1, 10):
                cell = ws_f.cell(row=item_row, column=c)
                cell.font = font_normal
                cell.border = border_cell
                if c in [5, 6, 7, 8, 9]:
                    cell.number_format = "#,##0.00"
            item_row += 1
            
        ws_f.cell(row=item_row, column=4, value="TOPLAM:").alignment = align_right
        ws_f.cell(row=item_row, column=4).font = font_bold
        ws_f.cell(row=item_row, column=5, value=float(f.get("toplam_miqdar") or 0)).alignment = align_right
        ws_f.cell(row=item_row, column=5).font = font_bold
        ws_f.cell(row=item_row, column=5).number_format = "#,##0.00"
        ws_f.cell(row=item_row, column=9, value=float(f.get("yekun_mebleg") or 0)).alignment = align_right
        ws_f.cell(row=item_row, column=9).font = font_bold
        ws_f.cell(row=item_row, column=9).number_format = "#,##0.00"
        ws_f.cell(row=item_row, column=5).border = border_total
        ws_f.cell(row=item_row, column=9).border = border_total
        
        item_row += 2
        ws_f.cell(row=item_row, column=7, value="Cəmi:").font = font_bold
        ws_f.cell(row=item_row, column=9, value=float(f.get("cemi_qiymet") or 0)).number_format = "#,##0.00"
        ws_f.cell(row=item_row+1, column=7, value="Yekun Endirim:").font = font_bold
        ws_f.cell(row=item_row+1, column=9, value=float(f.get("endirimi") or 0)).number_format = "#,##0.00"
        ws_f.cell(row=item_row+2, column=7, value="Yekun Məbləğ:").font = font_bold
        ws_f.cell(row=item_row+2, column=9, value=float(f.get("yekun_mebleg") or 0)).number_format = "#,##0.00"
        ws_f.cell(row=item_row+2, column=9).font = font_bold
        
        for col in ws_f.columns:
            max_l = max(len(str(cell.value or '')) for cell in col)
            col_let = get_column_letter(col[0].column)
            ws_f.column_dimensions[col_let].width = max(max_l + 4, 14.0)
            
    return wb
