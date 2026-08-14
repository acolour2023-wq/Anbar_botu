========================================================================
                 ANBAR SAYIMI YARDIMÇISI (İNVENTARİZASİYA BOTU)
========================================================================

Bu proqram anbardakı məhsulların real sayını sürətli şəkildə daxil etmək,
sistem qalıqları ilə fərqini (say və qiymət olaraq) dərhal görmək və 
yekun hesabatı Excel faylına yazmaq üçün hazırlanmışdır.

------------------------------------------------------------------------
İŞƏ SALMAQ ÜÇÜN TƏLƏBLƏR:
------------------------------------------------------------------------
1. Cari qovluğa (anbar_botu qovluğuna) son anbar qalığı olan Excel (.xlsx) 
   faylını kopyalayın.
2. Excel faylınızın birinci (başlıq) sətirində aşağıdakı sütunların olduğundan
   əmin olun (adlar eynilə və ya oxşar ola bilər):
   * Brend (və ya Adı / Brand / Name)
   * Anbar Qaligi (və ya Qalıq / Stock)
   * Barkod (və ya Barcode)
   * Mehsulun qiymeti (və ya Qiyməti / Price)

   QEYD: Əgər faylınızda "Yeni Sayim", "Say ferqi" və "Qiymet ferqi" sütunları
   yoxdursa, proqram onları fayla avtomatik əlavə edəcəkdir.

3. "run.bat" faylına cüt klikləyin. O, lazım olan kitabxananı (openpyxl)
   avtomatik quraşdıracaq və proqramı başladacaq.

------------------------------------------------------------------------
PROQRAMDAN İSTİFADƏ QAYDASI:
------------------------------------------------------------------------
1. Proqram açılanda qovluqdakı Excel faylını aşkar edəcək. Əgər birdən çox
   fayl tapılsa, sizdən siyahıdan seçim etməyinizi istəyəcək.
2. Açılan ekranda "Barkod" sahəsinə klikləyin (və ya barkod oxuyucu aparatla skan edin).
3. Barkodu daxil edib Enter düyməsini sıxın:
   - Əgər məhsul bazada varsa, adı, sistemdəki qalığı və qiyməti görünəcək.
   - Kursor avtomatik olaraq "Real sayı daxil edin" sahəsinə keçəcək.
4. Faktiki (real) sayılan miqdarı yazıb Enter sıxın:
   - Proqram dərhal "SAY FƏRQİ" və "QİYMƏT FƏRQİ" bölmələrində fərqləri hesablayacaq.
   - Bu məhsul "Son Skan Edilənlər" tarixinə əlavə olunacaq.
   - Kursor yenidən növbəti məhsulu skan etmək üçün "Barkod" sahəsinə qayıdacaq
     (heç bir mausa toxunmadan sürətli skan edə bilərsiniz).
5. Əgər skan etdiyiniz barkod bazada yoxdursa, "Yeni Məhsul Əlavə Et" düyməsini 
   sıxaraq məhsulu cədvələ əlavə edə bilərsiniz.
6. Bütün sayım bitdikdən sonra "Yadda Saxla və Çıx" düyməsini sıxın.

⚠️ DİQQƏT: Yadda saxlamazdan əvvəl Excel faylının kompüterdə açıq olmadığından
(MS Excel proqramında açıq qalmadığından) əmin olun. Əks halda proqram yazmağa
icazə verməyəcək.

Uğurlar!
========================================================================
