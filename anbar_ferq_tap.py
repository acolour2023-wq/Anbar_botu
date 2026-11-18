# -*- coding: utf-8 -*-
import os, sys, glob
import xlrd, openpyxl
import psycopg2
from collections import defaultdict

print("=" * 80)
print("       DOREGROUP — ANBAR QALIĞI VƏ VOZVRAT MÜQAYİSƏ ALƏTİ")
print("=" * 80)

DSN = os.environ.get("DATABASE_URL", "postgresql://neondb_owner:npg_5ImHSDlnZJ6s@ep-blue-feather-as3nt6qo-pooler.c-4.eu-central-1.aws.neon.tech/neondb?sslmode=require")

# 1. Desktop-dan XARAB_ANBAR faylını tapırıq
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

# 2. Anbar faylını oxuyuruq və Barkod <-> Kod xəritəsini qururuq
barcode_to_kod = {}
kod_to_barcode = {}
xarab_qty = defaultdict(float)
item_names = {}

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
            
            canon_key = val0 if val0 else bcode
            xarab_qty[canon_key] += qty
            if bcode and val0:
                barcode_to_kod[bcode] = val0
                kod_to_barcode[val0] = bcode
            if name:
                item_names[canon_key] = name
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
            
            canon_key = val0 if val0 else bcode
            xarab_qty[canon_key] += qty
            if bcode and val0:
                barcode_to_kod[bcode] = val0
                kod_to_barcode[val0] = bcode
            if name:
                item_names[canon_key] = name
except Exception as e:
    print(f"[-] Anbar faylını oxuyarkən xəta: {e}")
    sys.exit(1)

total_xarab = sum(xarab_qty.values())
print(f"[+] Əsas Anbar Qalığındakı YEKUN SAY: {total_xarab:.0f} ədəd ({len(xarab_qty)} adda məhsul)")

# 3. Neon DB-dən fakturaları oxuyuruq və eyni açarlarla birləşdiririk
vozvrat_qty = defaultdict(float)
vozvrat_invoices = defaultdict(list)
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
        q = float(qty or 0)
        all_fakturalar_list.add(f_no)
        
        # Açar təyini: Əvvəlcə artikl, yoxdursa barkod vasitəsilə kod, yoxsa barkod
        if art and art.isdigit() and len(art) in (4, 5, 6):
            canon_key = art
        elif bar and bar in barcode_to_kod:
            canon_key = barcode_to_kod[bar]
        elif bar:
            canon_key = bar
        else:
            canon_key = art or name
            
        vozvrat_qty[canon_key] += q
        vozvrat_invoices[canon_key].append((f_no, q))
        if canon_key not in item_names and name:
            item_names[canon_key] = name

    cur.close()
    conn.close()
    print(f"[+] Botun bazasından Yararsız Anbar üzrə {len(rows)} mal sətiri ({len(all_fakturalar_list)} faktura) oxundu.")
except Exception as e:
    print(f"[-] Bot bazasına qoşularkən xəta: {e}")
    sys.exit(1)

total_vozvrat = sum(vozvrat_qty.values())
print(f"[+] Botun bazasındakı YEKUN SAY: {total_vozvrat:.0f} ədəd ({len(vozvrat_qty)} adda məhsul)")
print("-" * 80)

# 4. Müqayisə
all_keys = set(xarab_qty.keys()).union(set(vozvrat_qty.keys()))
discrepancies = []

for k in all_keys:
    xq = xarab_qty.get(k, 0.0)
    vq = vozvrat_qty.get(k, 0.0)
    diff = vq - xq
    if abs(diff) > 0.001:
        discrepancies.append({
            'key': k,
            'name': item_names.get(k, 'Bilinməyən məhsul'),
            'xarab_qty': xq,
            'vozvrat_qty': vq,
            'diff': diff,
            'invoices': vozvrat_invoices.get(k, [])
        })

print("\n" + "=" * 80)
print(f" NƏTİCƏ: Əsas Anbar = {total_xarab:.0f} ədəd | Bot Hesabatı = {total_vozvrat:.0f} ədəd")
print(f" FƏRQ:   {total_vozvrat - total_xarab:+.0f} ədəd")
print("=" * 80)

if not discrepancies:
    print("\n✅ TƏBRİKLƏR! 100% DƏQİQ UYĞUNLUQ!")
    print("   Bütün məhsullar və saylar hər iki siyahıda BİRƏ-BİR EYNİDİR!")
    print("   Heç bir artıq, əskik və ya kənar faktura yoxdur.\n")
else:
    print(f"\n⚠️  AŞKARLANAN FƏRQLƏR ({len(discrepancies)} məhsul):")
    print(f"{'Kod':<8} | {'Əsas Anbar':>10} | {'Bot Faktura':>11} | {'Fərq':>6} | {'Məhsulun Adı':<32} | Fakturalar")
    print("-" * 90)
    discrepancies.sort(key=lambda d: abs(d['diff']), reverse=True)
    for d in discrepancies:
        inv_str = ", ".join([f"{inv}({q:.0f})" for inv, q in d['invoices']])
        print(f"{d['key']:<8} | {d['xarab_qty']:>10.0f} | {d['vozvrat_qty']:>11.0f} | {d['diff']:>+6.0f} | {d['name'][:30]:<32} | {inv_str}")
    print("-" * 90)

print("\n" + "=" * 80)
input("Yoxlama başa çatdı. Bağlamaq üçün Enter basın...")
