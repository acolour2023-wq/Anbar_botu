# -*- coding: utf-8 -*-
import os, sys, glob
import xlrd, openpyxl
import psycopg2
from collections import defaultdict

print("=" * 80)
print("       DOREGROUP — ANBAR QALIĞI VƏ FAKTUR POLİSİ (FƏRQ AXTARIŞI)")
print("=" * 80)

DSN = os.environ.get("DATABASE_URL", "postgresql://neondb_owner:npg_5ImHSDlnZJ6s@ep-blue-feather-as3nt6qo-pooler.c-4.eu-central-1.aws.neon.tech/neondb?sslmode=require")

# 1. Neon DB-dən fakturaları oxuyuruq
vozvrat_by_key = defaultdict(float)
item_info = {}
item_invoices = defaultdict(list)
all_fakturalar_list = set()

try:
    conn = psycopg2.connect(DSN)
    cur = conn.cursor()
    cur.execute("""
        SELECT m.artikl, m.barkod, m.malin_adi, m.miqdar, m.faktura_no, f.anbar
        FROM vozvrat_mallar m
        JOIN vozvrat_fakturalar f ON m.faktura_id = f.id
        WHERE UPPER(f.anbar) LIKE '%YARARSIZ%' OR UPPER(f.anbar) LIKE '%XARAB%'
        ORDER BY m.id ASC
    """)
    rows = cur.fetchall()
    for artikl, barkod, name, qty, f_no, anbar in rows:
        art = str(artikl or '').strip()
        bar = str(barkod or '').strip()
        name = str(name or '').strip()
        q = float(qty or 0)
        all_fakturalar_list.add(f_no)
        
        # Əsas açar: Əgər 4-6 rəqəmli artikl varsa artikl, yoxsa barkod
        key = art if (art.isdigit() and len(art) in (4,5,6)) else (bar if bar else name)
        vozvrat_by_key[key] += q
        
        if key not in item_info:
            item_info[key] = {'artikl': art, 'barkod': bar, 'name': name}
        item_invoices[key].append((f_no, q))
        
    cur.close()
    conn.close()
    print(f"[+] Botun bazasından Yararsız Anbar üzrə {len(rows)} mal sətiri ({len(all_fakturalar_list)} faktura) oxundu.")
except Exception as e:
    print(f"[-] Bot bazasına qoşularkən xəta: {e}")
    sys.exit(1)

total_vozvrat_qty = sum(vozvrat_by_key.values())
print(f"[+] Botun bazasındakı YEKUN SAY: {total_vozvrat_qty:.0f} ədəd")

# 2. Desktop-dan XARAB_ANBAR faylını tapırıq
desktop = os.path.join(os.path.expanduser("~"), "Desktop")
found_files = glob.glob(os.path.join(desktop, "*xarab*.xls*")) + glob.glob(os.path.join(desktop, "*XARAB*.xls*"))
found_files = [f for f in found_files if not os.path.basename(f).startswith("~$")]

if not found_files:
    print("\n[-] DİQQƏT: Masaüstündə (Desktop) 'XARAB_ANBAR.xls' faylı tapılmadı!")
    print("    Zəhmət olmasa 1C/Mikro-dan çıxardığınız anbar qalığı faylını Masaüstünə atın.")
    input("\nÇıxmaq üçün Enter basın...")
    sys.exit(1)

target_file = found_files[0]
print(f"[+] Tapılan Anbar Faylı: {os.path.basename(target_file)}")

# 3. Anbar faylını oxuyuruq
xarab_by_key = defaultdict(float)

try:
    if target_file.endswith('.xls'):
        wb_x = xlrd.open_workbook(target_file)
        sh = wb_x.sheet_by_index(0)
        for r in range(1, sh.nrows):
            val0 = str(sh.cell_value(r, 0)).strip()
            if val0.endswith('.0'): val0 = val0[:-2]
            if not val0 or val0.lower() in ('cəmi', 'cemi', 'total'):
                continue
            bcode = str(sh.cell_value(r, 1)).strip()
            if bcode.endswith('.0'): bcode = bcode[:-2]
            name = str(sh.cell_value(r, 3)).strip()
            qty = float(sh.cell_value(r, 4) or 0)
            
            key = val0 if (val0.isdigit() and len(val0) in (4,5,6)) else (bcode if bcode else name)
            xarab_by_key[key] += qty
            if key not in item_info:
                item_info[key] = {'artikl': val0, 'barkod': bcode, 'name': name}
    else:
        wb_x = openpyxl.load_workbook(target_file, data_only=True)
        sh = wb_x.active
        for r in range(2, sh.max_row + 1):
            val0 = str(sh.cell(row=r, column=1).value or '').strip()
            if not val0 or val0.lower() in ('cəmi', 'cemi', 'total'):
                continue
            bcode = str(sh.cell(row=r, column=2).value or '').strip()
            name = str(sh.cell(row=r, column=4).value or '').strip()
            qty = float(sh.cell(row=r, column=5).value or 0)
            
            key = val0 if (val0.isdigit() and len(val0) in (4,5,6)) else (bcode if bcode else name)
            xarab_by_key[key] += qty
            if key not in item_info:
                item_info[key] = {'artikl': val0, 'barkod': bcode, 'name': name}
except Exception as e:
    print(f"[-] Anbar faylını oxuyarkən xəta: {e}")
    sys.exit(1)

total_xarab_qty = sum(xarab_by_key.values())
print(f"[+] Əsas Anbar Qalığındakı YEKUN SAY: {total_xarab_qty:.0f} ədəd")
print("-" * 80)

# 4. Müqayisə
all_keys = set(xarab_by_key.keys()).union(set(vozvrat_by_key.keys()))
discrepancies = []

for k in all_keys:
    xq = xarab_by_key.get(k, 0.0)
    vq = vozvrat_by_key.get(k, 0.0)
    diff = vq - xq
    if abs(diff) > 0.001:
        info = item_info.get(k, {})
        discrepancies.append({
            'key': k,
            'artikl': info.get('artikl', ''),
            'barkod': info.get('barkod', ''),
            'name': info.get('name', 'Bilinməyən'),
            'xarab_qty': xq,
            'vozvrat_qty': vq,
            'diff': diff,
            'invoices': item_invoices.get(k, [])
        })

print("\n" + "=" * 80)
print(f" NƏTİCƏ: Əsas Anbar = {total_xarab_qty:.0f} ədəd | Bot Hesabatı = {total_vozvrat_qty:.0f} ədəd")
print(f" FƏRQ:   {total_vozvrat_qty - total_xarab_qty:+.0f} ədəd")
print("=" * 80)

if not discrepancies:
    print("\n✅ TƏBRİKLƏR! 100% DƏQİQ UYĞUNLUQ!")
    print("   Bütün məhsullar və saylar hər iki siyahıda BİRƏ-BİR EYNİDİR!")
    print("   Heç bir artıq, əskik və ya kənar faktura yoxdur.\n")
else:
    print(f"\n⚠️  AŞKARLANAN FƏRQLƏR ({len(discrepancies)} məhsul):")
    print(f"{'Kod/Açar':<10} | {'Baza':>6} | {'Bot':>6} | {'Fərq':>6} | {'Məhsulun Adı':<30} | Fakturalar")
    print("-" * 85)
    discrepancies.sort(key=lambda d: abs(d['diff']), reverse=True)
    for d in discrepancies:
        inv_str = ", ".join([f"{inv}({q:.0f})" for inv, q in d['invoices']])
        print(f"{d['key']:<10} | {d['xarab_qty']:>6.0f} | {d['vozvrat_qty']:>6.0f} | {d['diff']:>+6.0f} | {d['name'][:28]:<30} | {inv_str}")
    print("-" * 85)

print("\n" + "=" * 80)
input("Yoxlama başa çatdı. Bağlamaq üçün Enter basın...")
