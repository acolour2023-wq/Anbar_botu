import os
import sys
import time
import openpyxl
import psycopg2

if sys.stdout:
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

DATABASE_URL = os.environ.get(
    "DATABASE_URL",
    "postgresql://neondb_owner:npg_5ImHSDlnZJ6s@ep-blue-feather-as3nt6qo-pooler.c-4.eu-central-1.aws.neon.tech/neondb?sslmode=require"
)

def find_excel_file():
    curr_dir = os.path.dirname(os.path.abspath(__file__))
    candidates = [
        "Son-Anbar-qalıgı.xlsx",
        "Son-Anbar-qaligi.xlsx",
        "Son anbar qaligi.xlsx",
        "Son anbar qalıgı.xlsx",
        "Anbar_Qaligi.xlsx",
        "anbar_qaligi.xlsx"
    ]
    for c in candidates:
        p = os.path.join(curr_dir, c)
        if os.path.exists(p):
            return p
            
    try:
        for f in os.listdir(curr_dir):
            if f.startswith("~$") or f.endswith(".bak") or "netice" in f.lower() or "sayim" in f.lower():
                continue
            if f.lower().endswith(".xlsx"):
                f_norm = normalize_header(f)
                if "anbar" in f_norm or "qaliq" in f_norm:
                    return os.path.join(curr_dir, f)
    except Exception:
        pass
        
    return os.path.join(curr_dir, "Son anbar qaligi.xlsx")

EXCEL_FILE = find_excel_file()

def normalize_header(text):
    if not text:
        return ""
    text = str(text).strip()
    replacements = {
        'İ': 'i', 'I': 'i', 'İ': 'i', 'ı': 'i',
        'Ə': 'e', 'ə': 'e',
        'Ö': 'o', 'ö': 'o',
        'Ü': 'u', 'ü': 'u',
        'Ğ': 'g', 'ğ': 'g',
        'Ç': 'c', 'ç': 'c',
        'Ş': 's', 'ş': 's'
    }
    for k, v in replacements.items():
        text = text.replace(k, v)
    text = text.lower()
    for ch in ['_', '-', '.', '/', '\\', ':', ';', '(', ')', '[', ']']:
        text = text.replace(ch, ' ')
    return " ".join(text.split())

barkod_synonyms = [
    'barkod', 'barkkod', 'barcode', 'strixkod', 'strix-kod', 'strix_kod',
    'ean', 'shtrihkod', 'штрихкод', 'штрих-код', 'bar_code',
    'sh_kod', 'bar_kod', 'barkod_no', 'barkkod_no',
    'sto_kategori_kodu', 'sto_barkod', 'stok_barkod', 'barkodu'
]

kod_synonyms = [
    'kod', 'kodu', 'product_code', 'mal_kodu', 'mehsul_kodu',
    'kod_mehsul', 'kod_mal', 'код', 'артикул', 'artikul', 'item_code', 'item_id',
    'sto_kod', 'sto_kodu', 'stok_kodu'
]

brend_synonyms = [
    'brend', 'brand', 'marka', 'istehsalci', 'производитель',
    'фирма', 'firma', 'vendor', 'manufacturer', 'malin_markasi',
    'taminatci', 'supplier',
    'sto_uretici_kodu', 'sto_marka', 'sto_marka_kodu', 'uretici'
]

adi_synonyms = [
    'mehsulun adi', 'mehsulun_adi', 'malin adi', 'malin_adi',
    'adi', 'name', 'mehsul', 'description', 'nomenklatura',
    'наименование', 'название', 'товар', 'urun_adi', 'tam_adi',
    'mehsul_adi', 'item_name', 'aciglama',
    'sto_isim', 'sto_adi', 'stok_adi'
]

qaliq_synonyms = [
    'anbar qaligi', 'anbar_qaligi', 'anbar qaliq', 'anbar', 'shemkir anbar',
    'qaliq', 'stock', 'sistem qaligi',
    'sistem_qaligi', 'miqdar', 'sayi', 'say', 'qaliq_miqdari', 'остаток',
    'количество', 'stok_miktari', 'stok', 'sto_stok', 'balance', 'qty',
    'quantity', 'son_qaliq', 'tek_qaliq', 'mevcut'
]

qiymet_synonyms = [
    'mehsulun qiymeti', 'mehsulun_qiymeti', 'qiymet', 'qiymeti',
    'price', 'satis_qiymeti', 'satis qiymeti', 'satis_qiymet', 'satis qiymet',
    'satiş qiyməti', 'satiş qiymət', 'maya_qiymeti', 'maya qiymeti', 'цена',
    'стоимость', 'fiyat', 'sto_fiyat', 'cost', 'unit_price', 'qiymat', 'mebleg',
    'perakende', 'pərakəndə', 'retail', 'qiymet azn', 'qiymeti azn',
    'qiymet (azn)', 'qiymeti (azn)'
]

def find_col_index(headers_map, synonyms, exclude_indices=None):
    if exclude_indices is None:
        exclude_indices = set()
    for syn in synonyms:
        norm_syn = normalize_header(syn)
        if norm_syn in headers_map:
            col_idx = headers_map[norm_syn]
            if col_idx not in exclude_indices:
                return col_idx
    for syn in synonyms:
        norm_syn = normalize_header(syn)
        for h_key, col_idx in headers_map.items():
            if col_idx in exclude_indices:
                continue
            if any(k in norm_syn for k in ['kod', 'artikul', 'item_code']) and 'barkod' in h_key:
                continue
            if 'barkod' in norm_syn and any(k == h_key for k in ['kod', 'kodu']):
                continue
            if norm_syn in h_key or h_key in norm_syn:
                return col_idx
    return None

def parse_num(val):
    if val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip().replace(" ", "").replace("\xa0", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0

def sync_excel_to_db(filepath):
    if not os.path.exists(filepath):
        print(f"[XETA] Fayl tapılmadı: {filepath}")
        return False

    print(f"\n[OKUNUR] Fayl oxunur: {os.path.basename(filepath)}...")
    try:
        wb = openpyxl.load_workbook(filepath, data_only=True)
        sheet = wb.active
    except Exception as e:
        print(f"[XETA] Excel faylını açmaq mümkün olmadı (bəlkə başqa proqramda açıqdır?): {e}")
        return False

    headers = {}
    for col in range(1, sheet.max_column + 1):
        val = sheet.cell(row=1, column=col).value
        if val is not None:
            headers[normalize_header(val)] = col

    barkod_col = find_col_index(headers, barkod_synonyms)
    kod_col = find_col_index(headers, kod_synonyms, exclude_indices={barkod_col} if barkod_col else set())

    if not barkod_col and kod_col:
        barkod_col = kod_col
        kod_col = None

    brend_col = find_col_index(headers, brend_synonyms)
    adi_col = find_col_index(headers, adi_synonyms, exclude_indices={brend_col} if brend_col else set())
    if not brend_col:
        brend_col = adi_col

    anbar_qaligi_col = find_col_index(headers, qaliq_synonyms)
    qiymet_col = find_col_index(headers, qiymet_synonyms)

    if not anbar_qaligi_col or not barkod_col:
        print(f"[XETA] Vacib sütunlar tapılmadı! Barkod sütunu: {barkod_col}, Qalıq sütunu: {anbar_qaligi_col}")
        return False

    products_data = []
    seen = set()

    for r in range(2, sheet.max_row + 1):
        b_val = sheet.cell(row=r, column=barkod_col).value
        barcode = ""
        if b_val is not None:
            if isinstance(b_val, float):
                barcode = str(int(b_val)).strip()
            else:
                barcode = str(b_val).strip()

        kod_val = ""
        if kod_col:
            k_val = sheet.cell(row=r, column=kod_col).value
            if k_val is not None:
                if isinstance(k_val, float):
                    kod_val = str(int(k_val)).strip()
                else:
                    kod_val = str(k_val).strip()

        if not barcode:
            if kod_val:
                barcode = kod_val
            else:
                continue

        if barcode in seen:
            continue
        seen.add(barcode)

        brend_val = str(sheet.cell(row=r, column=brend_col).value or "Naməlum").strip() if brend_col else "Naməlum"
        adi_val = str(sheet.cell(row=r, column=adi_col).value or "").strip() if adi_col else ""
        qaliq = parse_num(sheet.cell(row=r, column=anbar_qaligi_col).value)
        qiymet = parse_num(sheet.cell(row=r, column=qiymet_col).value) if qiymet_col else 0.0

        products_data.append((barcode, kod_val, brend_val, adi_val, qaliq, qiymet))

    print(f"[MELUMAT] Excel-də {len(products_data)} unikal məhsul aşkar edildi. Bulud bazasına qoşulur...")

    try:
        from psycopg2.extras import execute_batch
        conn = psycopg2.connect(DATABASE_URL)
        cur = conn.cursor()

        # 1. Update existing products' stock, price, brand, name, code without touching counts!
        update_data = [
            (qaliq, qiymet, brend, adi, kod, barcode)
            for barcode, kod, brend, adi, qaliq, qiymet in products_data
        ]
        update_query = """
            UPDATE products 
            SET qaliq = %s, qiymet = %s,
                brend = COALESCE(NULLIF(%s, ''), brend),
                adi = COALESCE(NULLIF(%s, ''), adi),
                kod = COALESCE(NULLIF(%s, ''), kod)
            WHERE barcode = %s
        """
        execute_batch(cur, update_query, update_data, page_size=300)

        # 2. Insert any brand new products not yet present in database
        insert_data = [
            (barcode, kod, brend, adi, qaliq, qiymet)
            for barcode, kod, brend, adi, qaliq, qiymet in products_data
        ]
        insert_query = """
            INSERT INTO products (barcode, kod, brend, adi, qaliq, qiymet, yeni, operator, order_num)
            VALUES (%s, %s, %s, %s, %s, %s, NULL, '', 0)
            ON CONFLICT (barcode) DO NOTHING
        """
        execute_batch(cur, insert_query, insert_data, page_size=300)

        conn.commit()
        cur.close()
        conn.close()

        print(f"[UGURLU] ✅ Sinxronizasiya tamamlandı! {len(products_data)} məhsul Mikro ilə yeniləndi.")
        print(f"         - QEYD: Operatorların mövcud sayımları və jurnalları 100% toxunulmaz qorundu.\n")
        return True
    except Exception as e:
        print(f"[XETA] Verilənlər bazası ilə sinxronizasiyada xəta baş verdi: {e}")
        return False

def main():
    print("=" * 65)
    print("     MIKRO ANBAR QALIQLARI - AVTOMATIK SINXRONIZATOR")
    print("=" * 65)
    print(f"İzlənən fayl: {EXCEL_FILE}")
    print("Rejim: Hər 15 dəqiqədən bir və ya fayl dəyişdikdə avtomatik yenilənmə")
    print("Durdurmaq üçün: Ctrl + C sıxın\n")

    # Initial sync
    last_mtime = None
    if os.path.exists(EXCEL_FILE):
        last_mtime = os.path.getmtime(EXCEL_FILE)
        sync_excel_to_db(EXCEL_FILE)
    else:
        print(f"[XEBERDARLIQ] '{os.path.basename(EXCEL_FILE)}' hələlik tapılmadı. Fayl qovluğa düşən kimi sinxronizasiya başlayacaq...")

    try:
        while True:
            time.sleep(10)
            active_file = find_excel_file()
            if os.path.exists(active_file):
                try:
                    curr_mtime = os.path.getmtime(active_file)
                    if last_mtime is None or curr_mtime != last_mtime:
                        time.sleep(2)
                        print(f"\n[{time.strftime('%H:%M:%S')}] Dəyişiklik aşkar edildi ({os.path.basename(active_file)})! Yenilənmə aparılır...")
                        if sync_excel_to_db(active_file):
                            last_mtime = os.path.getmtime(active_file)
                except Exception as ex:
                    pass
    except KeyboardInterrupt:
        print("\nSinxronizator dayandırıldı.")

if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--once":
        sync_excel_to_db(EXCEL_FILE)
    else:
        main()
