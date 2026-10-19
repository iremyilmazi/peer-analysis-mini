# Peer Analysis Mini

Mevcut banka Excel’ini KAP’taki **yalnız bir sonraki solo finansal dönem** ile güncelleyen yerel Python uygulaması. 13 banka, tarayıcıda açılan web paneli ve komut satırı kullanımı içerir.

Sürüm: **19 Ekim 2026**. Hedef/test ortamı: **Windows, Python 3.12**.

Kaynak depo: [iremyilmazi/peer-analysis-mini](https://github.com/iremyilmazi/peer-analysis-mini) (public).

Bu klasör tam kaynak projedir; önceki küçük ING düzeltme ZIP’i gibi bir yama değildir. Python ve kütüphaneler ayrıca kurulmalıdır. Bağımsız bir EXE, çevrimiçi hizmet veya boş Excel şablonu üreticisi değildir.

## İçindekiler

1. [Ne yapar?](#ne-yapar)
2. [Kurulum öncesi](#kurulum-öncesi)
3. [Windows kurulumu](#windows-kurulumu)
4. [Excel’i hazırlama](#exceli-hazırlama)
5. [Web arayüzünü açma](#web-arayüzünü-açma)
6. [İlk güvenli deneme](#ilk-güvenli-deneme)
7. [Gerçek güncelleme](#gerçek-güncelleme)
8. [Komut satırı](#komut-satırı)
9. [Dosyalar ve yedekler](#dosyalar-ve-yedekler)
10. [Sorun giderme](#sorun-giderme)
11. [Testler](#testler)
12. [GitHub’a hazırlama](#githuba-hazırlama)

## Ne yapar?

Her işlemde bir banka seçilir. Aynı Excel üzerinde bankalar sırayla çalıştırılır.

1. Excel’in seçili banka sayfasındaki son geçerli dönem başlığı okunur.
2. Hedef bunun **tam bir çeyrek sonrası** olarak belirlenir.
3. KAP’ta bu dönemin konsolide olmayan (solo) bildirimi aranır.
4. Yalnız hedef dönemin Excel ve PDF ekleri indirilir.
5. Finansal tablolar ve PDF dipnotları geçici kopyada işlenir.
6. Yalnız yeni dönemin TL / YP / TOPLAM kolonları mevcut şablona aktarılır.
7. Eski dönemler ve diğer banka hücreleri değer/biçim bakımından korunuyor mu kontrol edilir.
8. Güvenli denemede ayrı sonuç üretilir. Gerçek güncellemede kontrollerden sonra yedek alınır ve dosya atomik olarak değiştirilir.

Örnek: Excel’de son dönem `31.03.2026` ise yalnız `30.06.2026` aranır. Q2 yokken Q3’e atlanmaz. Ana Excel’de Q2 kaydedildikten sonra sonraki çalıştırma Q3’ü arar. Sadece indirilmiş bir Q2 dosyası veya deneme çıktısı, ana Excel’in dönemini ileri taşımaz.

**Geçmiş kaynak arşivi gerekmez ve ilk çalıştırmada 2020’den başlayan indirme yapılmaz.** Ancak Excel’in kendisi mevcut dönemleri ve doğru şablonu içermelidir. Geçerli dönem yoksa işlem durur.

Kapsam dışı: diğer yan Excel dosyaları, tarihçe oluşturma, tüm bankaları paralel güncelleme, otomatik zamanlama, bulut sunucusu ve Excel şablonunu sıfırdan kurma. Alt indirici/parser modüllerini doğrudan çalıştırmayın; bazılarında tarihçe işlevleri bulunur. Normal giriş noktaları `panel.py` ve `banka_otomatik_guncelle.py` dosyalarıdır.

## Kurulum öncesi

Gerekenler:

- Kurumunuzun izin verdiği Windows bilgisayar ve **Python 3.12** kurulumu. Kaynakta asgari sürüm kontrolü 3.11 olsa da doğrulanan kurulum 3.12’dir.
- Yazma izniniz olan yerel proje klasörü. Ağ sürücüsü/OneDrive eşzamanlaması ve dosya kilitleri yerine yerel çalışma önerilir.
- Kurumun onayladığı Python paket deposuna ve çalışma sırasında `https://www.kap.org.tr` adresine HTTPS erişimi.
- Halihazırda kullanılan, bu projenin banka sayfaları ve dönem düzeniyle uyumlu `butunlesik_banka_verisi.xlsx` dosyasının **çalışma kopyası**.
- Güncelleme sırasında Excel dosyasının Excel uygulamasında kapalı olması.

Microsoft Excel, Python işleminin çalışması için kurulu olmak zorunda değildir; sonuçları açıp kontrol etmek için uygun bir Excel görüntüleyici gerekir. Codex ve Node.js gerekmez. Git yalnız GitHub’a yükleme/güncelleme için gerekir; uygulamayı çalıştırmanın şartı değildir. Selenium `requirements.txt` ile kurulur; Edge bağlantısında ayrıca yerel Edge ve uyumlu `msedgedriver.exe` gerekir.

### Projeye yerel EdgeDriver eklemek

`msedgedriver.exe` dosyasını `panel.py` ile **aynı klasöre** koymanız yeterlidir. Varsayılan `auto` modunda proje bu dosyayı bulduğunda KAP sorgularını ve Excel/PDF indirmelerini yerel Edge üzerinden yapar; sürücüyü internetten indirmez. `drivers/msedgedriver.exe` ve `driver/msedgedriver.exe` konumları da desteklenir. Driver yoksa önceki HTTP yolu kullanılır; **Edge zorunlu olsun** istiyorsanız `kap_connection.json` içinde `mode` değerini `edge` yapın.

Detaylar, alternatif yollar, sürüm uyumu ve bağlantı testi: [EDGE_DRIVER_KURULUMU.md](docs/EDGE_DRIVER_KURULUMU.md). Driver `.exe` dosyası GitHub’a yüklenmez; hedef bilgisayarda ayrıca eklenir.

## Windows kurulumu

### 1. Klasörü yerleştirin

Bu klasörün tamamını örneğin aşağıdaki konuma koyun:

```text
C:\Projeler\peer-analysis-mini
```

Alt klasörleri ayırmayın. Komutları projenin kök klasöründe çalıştırın; `panel.py`, `requirements.txt` ve `kap_config.json` aynı seviyede olmalı.

PowerShell’de:

```powershell
Set-Location "C:\Projeler\peer-analysis-mini"
```

Buradaki yol örnektir; kendi gerçek klasörünüzü yazın. Yerel kurulum yolunuz proje koduna gömülmez.

### 2. Python’u doğrulayın

```powershell
py -3.12 --version
```

Çıktı `Python 3.12.x` olmalı. `py` bulunamıyorsa:

```powershell
python --version
```

Bu komut mağazayı açıyorsa veya Python bulunamıyorsa kurum BT’sinden onaylı kurulumu isteyin. Gerçek Python yolu biliniyorsa PowerShell’de tırnaklı program yolunu `&` ile çağırabilirsiniz:

```powershell
& "C:\Python312\python.exe" --version
```

Bu da yalnız yol örneğidir; bilgisayarınızda mevcut olmayan yolu aynen kullanmayın.

### 3. Projeye özel Python ortamı oluşturun

```powershell
py -3.12 -m venv .venv
```

`py` yerine onaylı `python` komutunu kullanıyorsanız:

```powershell
python -m venv .venv
```

İkisini birden çalıştırmanız gerekmez. Bu işlem proje içinde `.venv` oluşturur. Sistem Python’unu veya sistem paketlerini değiştirmez.

**Ortamı Activate.ps1 ile etkinleştirmek şart değildir.** Aşağıdaki bütün komutlar doğru Python’u açıkça çağırır; PowerShell execution policy’sini değiştirmeniz gerekmez.

### 4. Kütüphaneleri kurun

```powershell
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe -m pip check
```

`requirements.txt` test edilen doğrudan bağımlılıkların sürümlerini sabitler; geçişli bağımlılıklar için tüm platformları kapsayan bir lock dosyası değildir. İnternet/paket deposu erişimi gereklidir. Kurum pip ayarları ve onaylı proxy/paket deposu kullanılmalıdır. Sertifika kontrolünü kapatmayın; ağ engelini aşmaya çalışmayın.

BT bütün uygun wheel dosyalarını ve geçişli bağımlılıkları sağladıysa çevrimdışı kurulum mümkündür:

```powershell
.\.venv\Scripts\python.exe -m pip install --no-index --find-links .\kutuphaneler -r requirements.txt
```

`kutuphaneler` proje içinde oluşturulabilen yerel bir klasördür; bu dağıtımda wheel/Python dosyaları bulunmaz. Wheel’ler Windows/Python sürümü ve mimarinizle uyumlu olmalıdır.

### 5. Kaynak ve kütüphane kontrolünü çalıştırın

```powershell
.\.venv\Scripts\python.exe sistem_kontrolu.py
```

Kütüphaneler ve kaynaklar için `[OK]` ve sonunda `SISTEM HAZIR` beklenir. Henüz Excel kopyalamadıysanız Excel yok uyarısı normaldir; bu durumda güncelleme başlatılamaz.

## Excel’i hazırlama

1. Güncellenecek ana Excel’inizin yedeğini proje dışında alın.
2. **Çalışma kopyasını** proje köküne koyun.
3. Adını tam olarak `butunlesik_banka_verisi.xlsx` yapın.
4. Dosyayı Excel uygulamasında kapatın.
5. Doğru kabul ettiğiniz karşılaştırma/referans Excel’i başka bir yerde saklayın. Referansı çalışma girdisi olarak kullanmayın.

```text
peer-analysis-mini/
  butunlesik_banka_verisi.xlsx     ← siz eklersiniz; Git’e dahil edilmez
  panel.py
  banka_otomatik_guncelle.py
  kap_config.json
  requirements.txt
```

Bu proje boş bir `.xlsx` dosyasını banka şablonuna dönüştürmez. Sayfa adları ve mevcut satır kimlikleri önemlidir. Seçili bankanın sayfası ve mevcut dönem başlıkları bulunmalıdır. Birinci satırdaki dönemler `31.03.2026` gibi çeyrek sonu tarihleri; dönem alt kırılımları TL / YP / TOPLAM şeklindedir. Şablon satırlarını veya sayfa adlarını rastgele değiştirmeyin.

| Banka | Komut anahtarı | Excel sayfası |
|---|---|---|
| Akbank | `akbank` | `akbank` |
| DenizBank | `denizbank` | `denizbank` |
| Enpara | `enpara` | `enpara-bank` |
| Garanti BBVA | `garanti` | `garanti-bankasi` |
| Halkbank | `halkbank` | `halkbank` |
| HSBC | `hsbc` | `hsbc` |
| ING | `ing` | `ing-bank` |
| İş Bankası | `isbank` | `is-bankasi` |
| QNB | `qnb` | `qnb-bank` |
| TEB | `teb` | `turk-ekonomi-bankasi` |
| VakıfBank | `vakifbank` | `vakifbank` |
| Yapı Kredi | `yapikredi` | `yapi-kredi-bankasi` |
| Ziraat | `ziraat` | `ziraat-bankasi` |

## Web arayüzünü açma

```powershell
.\.venv\Scripts\python.exe panel.py
```

Alternatif olarak ortam kurulmuşsa `BASLAT.cmd` dosyasına çift tıklayın. Bu başlatıcı yalnız projenin `.venv` ortamını kullanır; farklı Python kurulumlarını tahmin ederek açmaz.

Panel varsayılan tarayıcıda açılır. Bağlantı adresi konsolda yazılır:

```text
http://127.0.0.1:OTOMATIK_SECILEN_PORT
```

Buradaki `OTOMATIK_SECILEN_PORT` yerine konsolda yazan gerçek sayı kullanılır. **Terminal/komut penceresini kullanım boyunca açık tutun.** Pencereyi kapatırsanız yerel sunucu durur.

- Banka listesinden istediğiniz bankayı seçin.
- Panel seçili bankanın Excel’deki son dönemini ve aranacak tek dönemi gösterir.
- Excel’i dışarıdan değiştirdiyseniz **Excel Durumunu Yenile** düğmesine basın.
- Excel/sayfa/dönem bulunmazsa başlatma düğmesi kapalı kalır.
- Bir işlem çalışırken banka/mod değiştirilemez ve ikinci işlem başlatılamaz.

Manuel adres kullanımı:

```powershell
.\.venv\Scripts\python.exe panel.py --no-browser --port 8766
```

Bu durumda `http://127.0.0.1:8766` adresini kendiniz açın. Port doluysa başka sayı kullanın veya `--port` vermeyin. `banka_paneli/index.html` dosyasına doğrudan çift tıklamayın; panelin Python sunucusu/API’si gereklidir. Arayüzde harici CDN veya bulut barındırma bağımlılığı yoktur.

## İlk güvenli deneme

1. Bankayı seçin ve dönemlerin doğru olduğundan emin olun.
2. **Güvenli deneme** varsayılan seçili kalsın.
3. **Güncellemeyi Başlat** düğmesine basın.
4. İndirme, Excel, PDF, birleştirme ve kontrol adımlarını logdan izleyin.
5. Başarı sonrası oluşan ayrı sonucu açıp kontrol edin.

Örnek çıktı:

```text
butunlesik_banka_verisi_ing_DRY_RUN.xlsx
```

Bu mod ana çalışma Excel’ini değiştirmez. Hedef kaynaklar ve kontrol raporları yerelde oluşur. Deneme çıktı adı aynı bankada tekrarlanırsa önceki deneme dosyası değiştirilebilir; önemli bir denemeyi farklı adla proje dışında saklayın veya CLI’de `--dry-run-output` kullanın.

Sonuç kontrolünde dönem, TL/YP/TOPLAM başlıkları, ana finansal tutarlar, dipnotlar, oran/kur tabloları, sıfır/boş ayrımı ve tutar birimlerini gözden geçirin. Bir dosyanın oluşmuş olması tek başına bütün ekonomik verilerin doğru olduğunun kanıtı değildir.

**Yeni dönem yok** hata değildir: beklenen sonraki dönem bulunmamıştır, ana Excel değişmez ve daha sonraki döneme atlanmaz.

## Gerçek güncelleme

Güvenli deneme çıktısını kontrol ettikten sonra aynı banka için **Gerçek güncelleme** modunu seçin. Başlatırken ek onay sorulur. Kontroller başarılıysa:

1. Güncelleme öncesi yedek kaydedilir.
2. Kontrolden geçen aday, çalışma Excel’inin yerine atomik olarak yazılır.
3. Panel son çıktı, yedek ve kontrol raporu yollarını gösterir.
4. Excel durumunu yenilediğinizde son dönem ilerlemiş görünür.

Aynı Excel’e başka program/kişi eşzamanlı yazmamalıdır. Girdi işlem sırasında değişmişse kayıt iptal edilir. **Çalışmayı Durdur** işlemi tamamlanmış bir kaydı geri almaz; durdurma sonrası logu ve dosyayı kontrol edin. İhtiyaç halinde yedekten çalışma kopyasını elle geri yükleyin.

Her banka ayrı seçilip sırayla çalıştırılır. Bu sürüm bütün bankaları aynı anda güncelleyen tek düğme içermez.

## Komut satırı

Panel olmadan aynı motor kullanılabilir. Ana motorun varsayılanı **gerçek kayıttır**; ilk denemede `--dry-run` unutulmamalıdır. Web paneli ise varsayılan olarak güvenli denemededir.

```powershell
# Banka listesini göster
.\.venv\Scripts\python.exe panel.py --list

# Kurulum ve seçili bankanın Excel durumunu kontrol et
.\.venv\Scripts\python.exe panel.py --check --bank ing

# Ana motorla güvenli deneme
.\.venv\Scripts\python.exe banka_otomatik_guncelle.py ing --dry-run --workers 1

# Ana motorla gerçek güncelleme: çalışma Excel'ine yazar
.\.venv\Scripts\python.exe banka_otomatik_guncelle.py ing --workers 1

# Panelin terminal girişinde varsayılan güvenli denemedir
.\.venv\Scripts\python.exe panel.py --bank ing

# Bu girişte gerçek kayıt açıkça seçilir
.\.venv\Scripts\python.exe panel.py --bank ing --live
```

Excel’i proje dışında tutmak isterseniz CLI’de açık yol verin; web panelinin girdisi proje kökündeki sabit dosyadır:

```powershell
.\.venv\Scripts\python.exe banka_otomatik_guncelle.py ing --workbook "C:\BankaCalisma\butunlesik_banka_verisi.xlsx" --dry-run --dry-run-output "C:\BankaCalisma\ing_kontrol.xlsx" --workers 1
```

Geliştirici/teşhis seçenekleri:

- `--period 30.06.2026`: yalnız Excel’in tam bir sonraki dönemine izin verir; geçmiş/uzak bir dönemi zorlamak için değildir. Tarih örnektir, kalıcı Q2 kilidi yoktur.
- `--skip-download`: yalnız beklenen sonraki dönemin daha önce indirilmiş ve doğrulanmış yerel Excel/PDF kaynağı varsa işler. Geçmiş arşiv gerekmez.
- `--history-dir`, `--cache-dir`, `--audit-report`, `--backup-root`: ayrı test dizinleri seçmek için.
- `--debug-dir`: ara aşama Excel kopyaları üretir; gerekmedikçe kullanmayın ve depoya yüklemeyin.
- `--force-download`: hedef kaynakları yeniden indirmek için; dönem atlama yetkisi vermez.

Tam seçenekler için `python banka_otomatik_guncelle.py --help` kullanın.

## Dosyalar ve yedekler

```text
panel.py                              Yerel HTTP paneli / terminal giriş noktası
banka_paneli/index.html + app.js       Arayüz kaynakları
banka_otomatik_guncelle.py             13 banka, tek dönem, denetim ve kayıt motoru
entegre_excel_olustur.py               KAP Excel tabloları okuyucusu
kap_config.json                       Banka yapılandırması; özel anahtar içermez
kap_connection.json                   auto / edge / http bağlantı ve yerel driver ayarları
kap_transport.py                      Paylaşılan HTTP / yerel Edge bağlantısı
sistem_kontrolu.py                    Kütüphane ve kaynak bütünlüğü kontrolleri
<banka klasörü>/                     Bankaya özel indirme/PDF adaptörleri
tests/                                Sentetik girdilerle otomatik testler
docs/                                 Mimari, test kapsamı, GitHub kontrol notları
```

Çalışırken oluşanlar:

```text
<banka klasörü>/data/history/YYYYQn/    Hedef dönemin Excel/PDF kaynakları
<banka klasörü>/data/pdf_cache/         Ayrıştırma önbelleği
outputs/otomatik_guncelleme_audit/      JSON işlem/kontrol raporları
outputs/otomatik_guncelleme_yedek/      Gerçek kayıt öncesi yedekler
```

Kaynak ve yedek dosyalarını sıradan cache sanarak silmeyin. `.gitignore` bunları GitHub’a dahil etmemek içindir; otomatik temizlik yapmaz. Kaynaklar, kullanıcı Excel’i ve kütüphaneler bu temiz dağıtımda bulunmaz. `teb_bankasi` ve bankalar arası bazı yardımcı kodlar ortak bağımlılıktır; yalnız tek bankayı kullanacaksanız bile rastgele klasör silmeyin.

## Sorun giderme

| Belirti | Kontrol / çözüm |
|---|---|
| `python` bulunamadı / mağaza açılıyor | Python 3.12’nin gerçek yolunu veya `py -3.12` kullanın. Kurum onaylı kurulum isteyin. |
| `No module named ...` | `.venv` Python’unu kullandığınızı doğrulayın; `-m pip install -r requirements.txt` komutunu aynı ortamda çalıştırın. |
| Python dosyası bulunamadı | Proje kökünde olduğunuzu ve alt klasörlerin tam taşındığını kontrol edin. |
| Excel/sayfa/dönem bulunamadı | Dosya adı, `.xlsx` uzantısı, seçili banka sayfası ve ilk satırdaki mevcut dönem başlıklarını kontrol edin. Boş Excel kullanmayın. |
| İzin/dosya kullanım hatası | Excel’i kapatın, lock/eşzamanlama sorununu kontrol edin; yerel yazılabilir dizin kullanın. |
| Panel açılmıyor | Terminal açık mı? Konsoldaki gerçek adresi kullanın; sabit port doluysa kaldırın. HTML dosyasını doğrudan açmayın. |
| Panel bağlantısı koptu | Python süreci kapandıysa yeniden başlatın; eski port sekmesini değil yeni adresi kullanın. |
| KAP erişim/sertifika/proxy hatası | Kurum BT’siyle onaylı HTTPS erişimini doğrulayın; TLS doğrulamasını kapatmayın. |
| `429` / geçici ağ hatası | Logdaki bekleme/yeniden deneme davranışını izleyin. Peş peşe çok sayıda işlem başlatmayın. |
| Yeni dönem yok | Excel’in tam sonraki döneminin yayımlanıp yayımlanmadığını kontrol edin. Q2 yokken Q3 beklemeyin. |
| PDF/Excel bozuk veya küçük | Kontrol raporunu okuyun; giriş/güvenlik sayfası inmiş olabilir. Kaynağı doğrulamadan gerçek kaydı zorlamayın. |
| ING cari sütun/eşleme hatası | Yeni başlık veya tablo düzeni olabilir; işlem güvenli şekilde durur. Sonuçları referanstan elle doldurup hatayı gizlemeyin. |
| `History/veri deseni degisikligi` | Hedef dışı hücre/biçim değişikliği bulundu; aday reddedildi. Audit ve gerekiyorsa ayrı debug çıktılarıyla inceleyin. |

Hata durumunda önce audit JSON içindeki `status`, `error` ve `steps` alanlarına bakın. Hata mesajı, KAP kaynak değişikliği ve gerçek kaynak dosyası olmadan tahmini finansal değer üretmeyin.

## Testler

Kurulumdan sonra internet indirmesi veya gerçek kullanıcı Excel’i gerektirmeyen testleri çalıştırın:

```powershell
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
```

Testler geçici/sentetik Excel’ler kullanır. Yerel HTTP panel uç noktaları, banka bağlamları, varsayılan güvenli mod, eşzamanlı işlem engeli, yeni dönem sınırı, yedek, eski hücrelerin korunması ve ING sütun seçimini kapsar. Canlı KAP indirmesi veya tam bankacılık doğruluğu ile aynı şey değildir.

Önceki 18–19 Ekim 2026 canlı doğrulamada, düzeltme sonrası 13 bankanın Q2 alanları kullanıcı referansındaki 12.236 dolu hücreyle eşleşti. ING 19 Ekim’de boş kaynak/önbellekle tekrarlandı; diğer 12 bankanın 18 Ekim başarılı sonuçları korundu. Bu temiz GitHub kopyasındaki yeni panel/taşınabilirlik düzenlemeleri için ayrıca çevrimdışı regresyon testleri yapılır. Her bankanın gerçek ağ üzerinden canlı testi bu hazırlık sırasında yeniden yapılmış sayılmaz.

Ayrıntılar: [mimari](docs/MIMARI_VE_KONTROLLER.md), [test kapsamı](docs/TEST_DURUMU.md).

## GitHub’a hazırlama

Public kaynak depo yukarıda bağlantılıdır. Üstteki büyük çalışma alanını değil, yalnız bu temiz proje klasörünü depo kökü olarak kullanın. Eski `outputs`, masaüstü Excel’leri ve eski `.git` klasörünü buraya taşımayın.

Mevcut depoyu GitHub Desktop’ta **File → Clone repository → URL** ile veya terminalden alabilirsiniz:

```powershell
git clone https://github.com/iremyilmazi/peer-analysis-mini.git
Set-Location peer-analysis-mini
```

İndirdikten sonra yukarıdaki Windows kurulum adımlarını uygulayın. Excel çalışma kopyanızı ayrıca ekleyin; kaynak depoda bulunmaz. Aşağıdaki `git init` / uzak depo oluşturma adımları yalnız ayrı, yeni bir depo hazırlamak içindir; klonlanan mevcut depoda tekrarlanmaz.

Git kuruluysa gelecekte bu klasör içinde:

```powershell
git init
git branch -M main
git add .
git status --short
git diff --cached --stat
git diff --cached --name-only
```

Commit’ten **önce** listede Excel/PDF, indirilmiş veri, log, `.venv`, erişim anahtarı, kurum dosyası veya kullanıcıya özel yol bulunmadığını kontrol edin. `.gitignore` yanlışlıkla eklemeyi azaltır; güvenlik taramasının veya elle incelemenin yerine geçmez. Önceden takip edilen dosyayı sonradan ignore etmek Git geçmişinden kaldırmaz.

Kaydedilecek dosyalar uygunsa:

```powershell
git commit -m "Initial mini application"
```

Ardından GitHub’da seçtiğiniz ad/görünürlükte boş bir depo oluşturup gerçek uzak depo URL’sini kullanın:

```powershell
git remote add origin REPO_URL
git push -u origin main
```

`REPO_URL` yer tutucudur; aynen çalıştırılmaz. Kimlik doğrulamayı Git’in/kurumunuzun onaylı yöntemiyle yapın; token’ı kaynak koda, URL’ye veya dokümana yazmayın. Bu projenin mevcut deposu kullanıcı isteğiyle public olarak oluşturulmuştur. `LICENSE` eklenmedi; lisans kararı ayrıca verilmelidir.

Kontrol listesi: [GITHUB_HAZIRLIK.md](docs/GITHUB_HAZIRLIK.md).
