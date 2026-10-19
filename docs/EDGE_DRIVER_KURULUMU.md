# Yerel EdgeDriver ile KAP bağlantısı

Bu özellik, KAP bağlantısını şirket bilgisayarındaki Microsoft Edge ve sizin sağladığınız yerel EdgeDriver üzerinden çalıştırır. Hem banka bildirim sorguları (JSON GET/POST) hem hedef Excel/PDF ekleri aynı Edge oturumundan alınır. Dosyaları indirdikten sonraki Excel/PDF işleme ve güvenlik kontrolleri değişmez; yine yalnız mevcut Excel'in tam bir sonraki dönemi işlenir.

## En kolay kurulum

1. Projenin güncel sürümünü alın.
2. Bilgisayardaki Microsoft Edge ile uyumlu, kurumunuzun onayladığı `msedgedriver.exe` dosyasını edinin. Uygulama bu dosyayı internetten indirmez.
3. ZIP içinden çıkarılmış gerçek `.exe` dosyasını `panel.py` ile aynı klasöre koyun:

```text
peer-analysis-mini/
  msedgedriver.exe              ← siz eklersiniz; GitHub'a yüklenmez
  panel.py
  kap_transport.py
  kap_connection.json
  requirements.txt
```

4. Yeni Selenium bağımlılığını aynı proje ortamına kurun:

```powershell
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe -m pip check
.\.venv\Scripts\python.exe sistem_kontrolu.py
```

5. Sistem kontrolünde `KAP baglantisi: edge` ve yerel sürücü yolu görülmelidir. Sistem kontrolü yalnız dosya/ayar/kütüphane varlığını kontrol eder; sürücüyü çalıştırmaz.
6. Excel'e dokunmadan ana sayfa bağlantısını deneyin:

```powershell
.\.venv\Scripts\python.exe kap_transport.py --check
```

Bu komut gerçek bir Edge oturumu açar, KAP ana sayfasını sınar ve normal bitişte kapatır. Finansal API erişimi, bütün ek indirmeleri veya tutarların doğruluğunu tek başına kanıtlamaz.

7. `BASLAT.cmd` veya `python panel.py` ile paneli açın. İlk işlemi **güvenli deneme** olarak çalıştırın. İşlem günlüğünde `KAP baglantisi: yerel EdgeDriver (...)` satırı görünür.

## Otomatik arama sırası

Varsayılan `mode: "auto"` iken aşağıdaki yerler sırayla aranır:

1. Proje kökü: `msedgedriver.exe`
2. Proje içi: `drivers/msedgedriver.exe`
3. Proje içi: `driver/msedgedriver.exe`

Dosya bulunursa Edge kullanılır. Hiçbiri yoksa önceki doğrudan HTTP bağlantısı korunur. Arama çalışma dizinine, masaüstüne, PATH'e veya eski bilgisayarın kullanıcı adına bağlı değildir; `kap_transport.py` dosyasının bulunduğu proje köküne göredir.

## Edge kullanımını zorunlu yapma

Şirket bilgisayarında Edge dışında bağlantı denenmesini istemiyorsanız `kap_connection.json` dosyasını şöyle ayarlayın:

```json
{
  "mode": "edge",
  "driver_path": "msedgedriver.exe",
  "headless": false,
  "edge_binary": null
}
```

Bu durumda sürücü yoksa, sürümler uyumsuzsa veya Edge başlatılamıyorsa işlem hata verir; sessizce HTTP'ye dönmez. `auto` modunda sürücü bulunup Edge seçildikten sonra oluşan hata da HTTP'ye geçişe yol açmaz.

Başka konum için `driver_path` değerini örneğin `drivers/msedgedriver.exe` yapın. Mutlak yol da kullanılabilir fakat bilgisayarlar arasında taşınabilirliği azaltır. Windows mutlak yollarında JSON içinde `C:/Araclar/msedgedriver.exe` veya çift ters eğik çizgi kullanın. Kişisel/kurumsal mutlak yollarınızı public depoya commit etmeyin.

Yalnız açıkça isterseniz `mode: "http"` doğrudan bağlantıyı seçer. `headless: false` Edge penceresini görünür açar; `true` penceresiz moddur. Kurum politikaları her iki modu da kısıtlayabilir. `edge_binary` normalde `null` kalır; standart dışı Edge kurulumu için kurumun verdiği `msedge.exe` yolu yazılabilir. Bu alan sürücü yolu değildir.

İsteğe bağlı ortam değişkenleri:

```powershell
$env:KAP_TRANSPORT = "edge"
$env:KAP_EDGE_DRIVER = "drivers/msedgedriver.exe"
.\.venv\Scripts\python.exe panel.py
```

Bu iki değişken JSON'daki mod/sürücü yolunu geçersiz kılar; panelin başlattığı banka alt süreçlerine de geçer. Göreli yollar yine proje köküne göredir. Açıkça belirtilmiş hatalı sürücü yolu, `auto` modunda da hata üretir. Değişkenleri aynı terminalde kaldırmak için `Remove-Item Env:KAP_TRANSPORT, Env:KAP_EDGE_DRIVER -ErrorAction SilentlyContinue` kullanın.

## Sürüm ve kurum politikaları

Edge sürümünü `edge://settings/help` sayfasından görün. Microsoft'a göre Edge ile EdgeDriver sürümlerinin ilk üç sayısal bölümü eşleşmelidir. Python tarafında Selenium 4 kullanılır; bu projede `requirements.txt` sürümü sabitler. [Microsoft resmi yönergesi](https://learn.microsoft.com/en-us/microsoft-edge/webdriver-chromium/).

Sürücü, Selenium'a açık bir yerel `Service(executable_path=...)` yolu ile verilir; Selenium Manager/PATH araması ve otomatik driver indirmesi kullanılmaz. [Selenium Edge Service API](https://www.selenium.dev/selenium/docs/api/py/selenium_webdriver_edge/selenium.webdriver.edge.service.html).

Normal indirme aşaması boyunca bir Edge oturumu tekrar kullanılır, aşama bitince veya hata oluşunca kapatılır. Kişisel Edge profiliniz/parolalarınız içe aktarılmaz. Yeni otomasyon oturumunda kurum proxy/SSO ayarlarının çalışması kurum politikasına bağlıdır; normal tarayıcının açılması otomasyonun da izinli olduğu anlamına gelmez.

Sertifika kontrolü, web güvenliği veya kurum politikaları kapatılmaz. Engel varsa kurum BT'sinden onaylı ayar desteği alın. Bu özellik ağ/erişim engellerini aşma yöntemi değildir. Tek bir HTTP yanıtı en fazla 64 MiB olabilir; daha büyük raporlar açık hata üretir.

| Hata | Kontrol |
|---|---|
| Yerel driver bulunamadı | ZIP değil çıkarılmış `.exe` var mı? Dosya adı/konumu ve `mode`/`driver_path` doğru mu? |
| EdgeDriver başlatılamadı / session not created | Edge ve driver sürümleri, mimari ve kurumun uygulama çalıştırma politikalarını kontrol edin. Ayrıntılı logdaki asıl hata nedeni yardımcı olabilir. |
| Selenium bulunamadı | `requirements.txt` kurulumunu paneli çalıştıran aynı `.venv` Python'u ile yapın. |
| KAP erişim/sertifika/proxy hatası | Kurum BT'si ile otomasyon oturumunun HTTPS/KAP erişimini doğrulayın; TLS'i kapatmayın. |
| Sistem kontrolü `http` gösteriyor | Driver otomatik arama konumlarında yoktur veya mod açıkça `http` seçilmiştir. |

`.gitignore` `.exe` ve driver dosyalarını depodan dışlar. Yerel sürücüyü, sürücü ZIP'ini, ortamınızı veya kurum ayarlarını public GitHub'a yüklemeyin.
