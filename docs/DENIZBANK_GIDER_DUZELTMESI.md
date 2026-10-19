# DenizBank 2026Q2: gider dipnotundaki metin tutarların düzeltilmesi

## Neden ve kapsam

19 Ekim 2026 incelemesi: 30 Haziran 2026 konsolide olmayan DenizBank raporunun PDF sayfa 81 / basılı sayfa 75'indeki “Diğer faaliyet giderlerine ilişkin bilgiler” tablosu incelendi. Cari dönemde sekiz satırdaki tutar `--` olarak yayımlanıyor. Ayrıştırıcı bunu metin olarak taşıyordu. Ortak Excel yazıcısı yalnız tek çizgiyi tanıdığından çift çizgi Excel'de metin kalıyordu. Önceki referans Excel'de de aynı metinler bulunduğu için yalnız referans eşitliği yeterli kontrol değildi.

Düzeltme yalnız DenizBank'ın bu dipnotunda açık çizgi tutarlarını **sayısal 0** yapar. Gerçek boş değerler, boş TL/YP alanları, etiketler ve diğer dipnotlar değiştirilmez. Bilinmeyen tutar metni, eksik/boş gider tablosu veya yanlış kolon yapısı görülürse işlem durdurulur; bunlar sessizce sıfırlanmaz. Bu kontrol, ortak ayrıştırıcının tablo atlayan uyarı mekanizmasının dışında çalışır.

`CACHE_VERSION` 10'dan **11**'e çıktı. Eski v10 önbellek yerine PDF yeniden ayrıştırılır; sonraki çalışmada v11 kullanılabilir. Raporları veya önbellek klasörünü elle silmeniz gerekmez. Ortak motor, diğer bankaların kaynakları, EdgeDriver ve bağımlılıklar değişmedi.

## Diğer bilgisayarda zorunlu olan yalnız iki kaynak dosyası

GitHub'da aşağıdaki dosyaların **tam içeriklerini** aynı yollardaki mevcut dosyalara aktarın; yalnız birkaç satırı eklemeyin. İkisini birlikte güncelleyin.

1. [denizbank/peerpdf_scripts/denizbank_dipnot.py](../denizbank/peerpdf_scripts/denizbank_dipnot.py)
2. [denizbank/denizbank_dipnot_guncelle.py](../denizbank/denizbank_dipnot_guncelle.py)

Yeni belge, otomatik test ve iskelet listesi de depoda güncellendi; uygulamayı çalıştırmak için bunları şirket bilgisayarına taşımanız gerekmez. `requirements.txt`, `BASLAT.cmd`, panel dosyaları ve diğer banka dosyalarını değiştirmeniz gerekmez. Paketleri yeniden kurmayın; şirket içi adres tanımlamaya gerek yoktur.

## CMD ile uygulama sırası

1. Panelde işlem varsa bitmesini bekleyin. Panelin çalıştığı CMD'de `Ctrl+C` ile kapatın. Excel'i kapatın.
2. Proje klasörünü VS Code'da açın; aşağıdaki komutları **proje kökünde** çalıştırın. Zaten kökteyseniz `cd` gerekmez.
3. İki eski kaynak dosyasını ve çalışma Excel'ini proje dışında, yeni ve boş bir yedek klasöre kopyalayın. Referans Excel'i hiçbir aşamada girdi olarak kullanmayın.
4. GitHub dosya sayfasında **Raw** görünümünden tüm kodu kopyalayıp ilgili dosyaya yapıştırın. UTF-8 olarak kaydedin; `.py.txt` veya HTML olmasın.
5. Önce sözdizimi ve sürümü kontrol edin:

```cmd
.venv\Scripts\python.exe -m py_compile denizbank\peerpdf_scripts\denizbank_dipnot.py denizbank\denizbank_dipnot_guncelle.py
```

Başarıda bu komutun çıktı vermemesi normaldir. Ardından:

```cmd
.venv\Scripts\python.exe -c "from denizbank.denizbank_dipnot_guncelle import CACHE_VERSION; from denizbank.peerpdf_scripts.denizbank_dipnot import extract_all_tables; assert CACHE_VERSION == 11; print('OK: DenizBank gider duzeltmesi ve cache v11 hazir')"
```

Beklenen: `OK: DenizBank gider duzeltmesi ve cache v11 hazir`. Önceki kurulumda doğruladığınız Python komutu farklıysa aynı doğrulanmış yorumlayıcıyı kullanın; paketleri olmayan başka Python'a geçmeyin. İçe aktarma hatası, eksik dosya veya farklı sürüm varsa güncellemeye başlamayın.

## Kritik: Q2 zaten Excel'de varsa normal güncelleme onu onarmaz

Proje yalnız mevcut son dönemin **bir sonrasını** işler. Son dönem Q2 ise hedef Q3 olur. Kod düzeltmesini yaptıktan sonra Q2 içeren Excel'de normal güncellemeyi başlatıp Q2'nin düzelmesini beklemeyin. `--period` bu güvenlik kuralını atlayan bir “eski dönemi zorla yaz” seçeneği değildir.

2026Q2'yi yeniden sınamak için **Q2 öncesindeki kendi çalışma Excel yedeğinizin ayrı bir kopyasını** kullanın; bu kopyada DenizBank'ın son dönemi 31.03.2026 olmalı. Bu dosya farklı adla proje köküne `denizbank_Q1_TEST_GIRDI.xlsx` olarak konabilir. Doğru kabul edilen referans dosyası bu amaçla kullanılmamalı.

Q1 çalışma yedeğiniz yoksa burada durun. Başka bankaların son dönemlerini silmeyin, Q2 kolonlarını elle kaldırmayın ve ana dosyayı eski bir dosyayla değiştirmeyin. Önce mevcut dosya yapısına göre ayrı bir onarım planı gerekir.

### Güvenli Q2 testi (önceki indirilmiş kaynaklar mevcutsa)

Önce dosyaları kontrol edin:

```cmd
dir denizbank_Q1_TEST_GIRDI.xlsx
dir denizbank\data\history\2026Q2
```

2026Q2 klasöründe hem KAP Excel'i hem konsolide olmayan finansal PDF bulunmalı. Aşağıdaki çıktı dosyası adı yeni olmalı; önemli eski bir denemenin üzerine yazmayın.

```cmd
.venv\Scripts\python.exe denizbank\denizbank_otomatik_guncelle.py --workbook denizbank_Q1_TEST_GIRDI.xlsx --history-dir denizbank\data\history --period 30.06.2026 --skip-download --dry-run --dry-run-output denizbank_2026Q2_GIDER_DUZELTME_TEST.xlsx
```

Bu komut yeni kaynak indirmez, ana/çalışma girdisini değiştirmez, ayrı sonuç üretir. Kaynaklar yoksa `--skip-download` seçeneğini kaldırın; kurumca onaylı ve daha önce çalıştırdığınız Edge bağlantı ayarlarıyla yalnız hedef dönemi indirir. Ağ erişimi başarısızsa kaynak doğrulamasını atlamayın.

### Sonucun kontrolü

- İncelenen orijinal şablonda dipnot başlığı A497, satırlar 499–518 ve Q2 toplam kolonu CA'dır. Sizin dosyanızda **başlığı ve 30.06.2026 / TOPLAM kolonunu doğrulamadan bu koordinatları kullanmayın**.
- Eski metinler: `CA501`, `CA502`, `CA504`, `CA505`, `CA507`, `CA508`, `CA509`, `CA510`. Yeni sonuçta bunlar sayı türünde 0 olmalı.
- Dipnotun 20 toplam hücresinin tamamı sayısal olmalı; TL/YP için 40 boş hücre boş kalmalı. Etiketler ve satır sırası korunmalı.
- Excel'in mevcut sayı biçimi sıfırı ekranda `-` gösterebilir. Bu **metin `--` ile aynı değildir**: hücre seçildiğinde formül çubuğundaki gerçek değer 0 olmalıdır. Görünüşü değiştirmek için bütün tabloyu metne dönüştürmeyin.
- Diğer 12 tutar, toplam, geçmiş dönemler ve diğer bankalar aynı kalmalı. Alt kırılımları üst toplamlarla birlikte toplayarak çift sayım yapmayın.
- Deneme çıktısını kendi doğru referansınızla karşılaştırın; referanstaki `--` metinlerinin de eski kusuru taşıyabileceğini unutmayın. Yalnız eşitlik değil hücre türünü de kontrol edin.

Başarılı denemeden sonra paneli tekrar açabilirsiniz. Önceden çalışan `BASLAT.cmd` kullanılabilir; yeni Python kodunun yüklenmesi için eski panel süreci mutlaka kapanmış olmalı. Deneme çıktısını doğrulamadan ana Excel'in üzerine koymayın.

## Geliştirici doğrulaması ve sınırları

Yeni regresyonlar `tests/test_denizbank_expense.py` içinde: çizgi varyantları, boşların korunması, normal tutarlar, bilinmeyen metinde durma, diğer dipnotların korunması, ortak ayrıştırıcının değiştirilmemesi, Excel sayı dönüşümü ve v10→v11 önbellek yenilenmesi.

Gerçek 2026Q2 testi, daha önce KAP'tan indirilmiş yerel Excel/PDF kullanılarak Q1 çalışma dosyasının ayrı kopyasında `--dry-run` ile yapılır. Bu test KAP'a yeni canlı bağlantı/indirme testi değildir. Giriş, referans ve önceki sonuç dosyalarının SHA-256 değerleri karşılaştırılır; özgün dosyalar değiştirilmez. Yeni sonuç, önceki DenizBank Q2 çıktısıyla bütün sayfalarda değer, veri türü ve hücre stili bakımından karşılaştırılır.

Sonuç: **79/79 otomatik test başarılı**. Gerçek Q2 denemesinde yalnız yukarıdaki sekiz metin hücresi sayısal 0 oldu; 20/20 toplam tutarı sayısal ve 40/40 boş TL/YP alanı korundu. Tüm banka sayfaları karşılaştırıldı; başka değer/veri türü veya hücre stili farkı yok. Eski v10 önbellek yenilendi, gerçek v11 önbelleğin PDF yeniden ayrıştırılmadan kullanılabildiği ayrıca doğrulandı. Ayrıştırılmış 21 tablonun önceki sonuçla karşılaştırmasında da yalnız bu sekiz tutar farklıydı. [Test özeti](TEST_DURUMU.md).

Bu düzeltme tüm bankaların tüm dipnotlarındaki çizgilerin toplu dönüşümü değildir. Başka tablo veya dönemde farklı bir metin varsa ayrıca incelenmelidir.
