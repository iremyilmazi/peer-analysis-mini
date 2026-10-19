# CMD ile kurulum ve ilk güvenli başlatma

Bu rehber, dosyaları başka bilgisayara elle kopyalanmış `peer_analysis_mini` projesi içindir. PowerShell, Git, GitHub Desktop veya yönetici yetkisi gerektiren bir işlem içermez. Kurumunuz Python/driver çalıştırmayı veya paket kurulumunu kısıtlıyorsa BT onayı alın; güvenlik politikalarını değiştirmeyin.

**Her komutu sırayla çalıştırın, bitmesini bekleyin. Bir adım hata verirse sonraki adıma geçmeyin. Bütün rehberi tek seferde CMD'ye yapıştırmayın.** Kod bloklarındaki kopyalama düğmesini kullanabilirsiniz. Aynı CMD penceresini kullanım boyunca açık tutun.

Kurum içi paket deposunun gerçek adresi bu public belgeye özellikle konulmadı. Adresi kendi notunuzdan/BT'den alıp aşağıdaki soruya bir kez yapıştıracaksınız. Adres, kimlik bilgisi ve kurum ayarlarını GitHub'a eklemeyin.

**Paketleri IT'nin yeni adresiyle tek tek kurduysanız ve `.venv` hazırsa:** kurulum komutlarını tekrar çalıştırmadan doğrudan [paket sürümü ve bağımlılık kontrolüne](#paket-surum-kontrolu) geçin. Yalnız `.whl` dosyasını indirmek, paketi Python'a kurmak değildir. Kontroller doğru sonuç vermeden Edge/KAP testine veya güncellemeye başlamayın.

## 1. Doğru klasörde CMD açın

Masaüstündeki `peer_analysis_mini` klasörünü Dosya Gezgini'nde açın. Adres çubuğuna `cmd` yazıp Enter'a basın. Aşağıdaki komutlar bu klasörde çalıştırılmalıdır.

Önce Python sürümü:

```cmd
python --version
```

Proje Python 3.12 ile test edildi; Python 3.11'den eskiyse burada durun. `python` bulunmuyorsa:

```cmd
py --version
```

Yalnız `py` çalışıyorsa 3. adımdaki ortam oluşturma komutunda da `python` yerine `py` kullanın. Hiçbiri çalışmıyorsa kurumun onaylı Python kurulumu gerekir.

Dosyaları kontrol edin:

```cmd
dir requirements.txt
dir panel.py
dir banka_otomatik_guncelle.py
dir kap_config.json
dir kap_connection.json
dir kap_transport.py
dir sistem_kontrolu.py
dir msedgedriver.exe
```

Dosyalar gerçekten bulunmalı. Yalnız boş iskelet oluşturmak yeterli değildir: gereken kaynak kodları kopyalanmış olmalıdır. `.py.txt`, `.json.txt` gibi yanlış uzantılar olmamalı. Driver ZIP içinde değil, çıkarılmış gerçek `msedgedriver.exe` olmalı ve `panel.py` ile aynı klasörde bulunmalıdır.

`.git`, `.gitignore`, `.gitattributes`, `BOS_PROJE_OLUSTUR.py`, `tests` ve `docs` çalışma için zorunlu değildir. Bu rehberi GitHub'dan okumak yeterlidir; hedef bilgisayarda ayrıca saklamanız gerekmez. Banka klasörlerini/ortak Python kaynaklarını silmeyin.

## 2. Edge ve driver sürümlerini kontrol edin

Microsoft Edge'de `edge://settings/help` adresini açıp sürümü görün. CMD'de:

```cmd
msedgedriver.exe --version
```

Edge ile EdgeDriver sürümlerinin ilk üç sayısal bölümü eşleşmeli. Uyumlu, kurumca onaylanmış sürücüyü siz sağlarsınız; proje otomatik driver indirmez. [Microsoft resmi sürüm uyumu yönergesi](https://learn.microsoft.com/en-us/microsoft-edge/webdriver/).

## 3. Projeye özel Python ortamını oluşturun

```cmd
python -m venv .venv
```

Yalnız `py` komutu çalışıyorsa yukarıdaki komut yerine bunu kullanın; ikisini birden çalıştırmanız gerekmez:

```cmd
py -m venv .venv
```

Sonra doğrulayın:

```cmd
.venv\Scripts\python.exe --version
```

`.venv` zaten bu bilgisayarda doğru Python ile oluşturulduysa yeniden oluşturmak gerekmez. Başka bilgisayarın `.venv` klasörünü taşımayın. Ortamı `activate` ile etkinleştirmek şart değildir; aşağıdaki komutlar doğru Python'u açıkça çağırır. Bu kurulumda `--user` kullanmayın.

## 4. Kütüphaneleri kurumun onaylı deposundan kurun

Önce paket listesinin dolu olduğunu görün:

```cmd
type requirements.txt
```

Listede `beautifulsoup4`, `lxml`, `openpyxl`, `pandas`, `pdfplumber`, `pypdf`, `selenium` ve sabit sürümleri olmalıdır.

Kurum adresini bir kez tanımlamak için aşağıdaki **tek komutu** kopyalayın. Soru geldiğinde kendi notunuzdaki/BT'nin onayladığı güncel paket deposu adresini yapıştırıp Enter'a basın. Adres, paket listesinin `/simple` uç noktası olmalı; kullanıcı adı/parola/token içermemeli.

```cmd
set /p "KURUM_PYPI_URL=BT onayli paket deposu adresini yapistirin: "
```

Bu değişken yalnız açık CMD penceresi için geçerlidir; Windows'un kalıcı ayarını değiştirmez. **Aşağıdaki A veya B yolundan yalnız uygun olanı seçin.** Dış PyPI'ye geçerek kurum politikasını aşmaya çalışmayın.

### A. Kurum deposu HTTPS kullanıyorsa

```cmd
.venv\Scripts\python.exe -m pip install -r requirements.txt --index-url "%KURUM_PYPI_URL%" --disable-pip-version-check
```

Sertifika hatası olursa BT'nin onaylı CA/sertifika yapılandırmasını kullanın; hatayı bastırmak için güven kontrolünü kapatmayın.

### B. Eski notunuzdaki gibi yalnız BT onaylı HTTP deposu kullanılıyorsa

HTTP şifreli ve kimliği doğrulanmış HTTPS değildir. `--trusted-host` güvenli HTTPS doğrulaması olmayan sunucuyu kabul eder. **Yalnız BT bu HTTP adresini ve bu kullanım biçimini onaylıyorsa** bu yolu kullanın; mümkünse BT'nin güncel HTTPS adresini tercih edin. [pip trusted-host açıklaması](https://pip.pypa.io/en/stable/cli/pip/#cmdoption-trusted-host).

Şu tek komutu çalıştırın; soru geldiğinde sadece depo sunucusunun IP/host adını girin. `http://`, port veya `/repository/...` yolu eklemeyin:

```cmd
set /p "KURUM_PYPI_HOST=BT onayli HTTP deposunun yalniz IP veya host adini girin: "
```

Ardından kurulumu başlatın:

```cmd
.venv\Scripts\python.exe -m pip install -r requirements.txt --index-url "%KURUM_PYPI_URL%" --trusted-host "%KURUM_PYPI_HOST%" --disable-pip-version-check
```

Paket deposuna bağlanılamazsa veya `No matching distribution found` çıkarsa burada durun. Önceki adres yerine IT'nin ilettiği güncel onaylı adresi kullanın. `ConnectTimeoutError` sunucuya bağlantı kurulamadığını gösterir; tek başına paket sürümünün depoda olmadığını kanıtlamaz. Kurum deposunda istenen sürüm/Python mimarinize uygun paket bulunmayabilir; sürüm numaralarını rastgele silmeyin. Hata mesajını inceleyin/BT ile görüşün. Gereksinim dosyasından toplu kurulum ve depo seçimi [pip tarafından desteklenir](https://pip.pypa.io/en/stable/cli/pip_install/).

<a id="paket-surum-kontrolu"></a>

### C. Kurulum sonrası paket sürümü ve bağımlılık kontrolü

Bu bölüm, ister toplu ister tek tek kurulum yapmış olun aynıdır. Komutlar paket indirmez/kurmaz ve ana Excel'e yazmaz. Kontrolü uygulamanın kullanacağı `.venv` Python'u üzerinden yapıyoruz; ana Python'da veya başka projede kurulu olması tek başına yeterli değildir.

#### C1. Kullanılan Python ortamını doğrulayın

```cmd
.venv\Scripts\python.exe -c "import sys; print(sys.executable)"
```

Çıktı, bu projenin `peer_analysis_mini\.venv\Scripts\python.exe` yolunu göstermeli. Dosya bulunamadı hatasında önce 3. adımı tamamlayın; ortamı silmeyin veya başka projeden `.venv` kopyalamayın.

#### C2. Yedi paketin kurulu sürümlerini gösterin

Aşağıdaki komutu **tek satır olarak** kopyalayın:

```cmd
.venv\Scripts\python.exe -m pip list --format=freeze --disable-pip-version-check | findstr /i /b "beautifulsoup4== lxml== openpyxl== pandas== pdfplumber== pypdf== selenium=="
```

Bu sürümün `requirements.txt` dosyasına göre beklenen değerler:

| Paket | Gerekli sürüm |
|---|---|
| beautifulsoup4 | 4.14.3 |
| lxml | 6.1.1 |
| openpyxl | 3.1.5 |
| pandas | 3.0.1 |
| pdfplumber | 0.11.9 |
| pypdf | 6.10.0 |
| selenium | 4.51.0 |

Yedi paketin tamamı görünmeli ve sürümleri eşleşmeli. Komut yalnız bu yedi adı süzer; diğer yardımcı bağımlılıkları göstermemesi normaldir. Esas sürüm kaynağı projenin güncel `requirements.txt` dosyasıdır.

Bir paket görünmüyor veya sürümü farklıysa **burada durun**. Listeyi ve sürümleri paylaşın; rastgele `--upgrade`, kaldırma veya sürüm silme uygulamayın. Paketleri ana Python'a kurmuş olabileceğinizi kontrol etmek için yalnız teşhis amaçlı:

```cmd
python -m pip list --format=freeze --disable-pip-version-check | findstr /i /b "beautifulsoup4== lxml== openpyxl== pandas== pdfplumber== pypdf== selenium=="
```

Bu ikinci listede paketlerin görünmesi `.venv` kontrolünün yerine geçmez; panel hâlâ `.venv` ortamını kullanacaktır. Doğru ortama kurulum gerekiyorsa 4A/4B'deki `.venv\Scripts\python.exe -m pip ...` komutlarını IT'nin güncel onaylı deposuyla değerlendirin.

#### C3. Yardımcı bağımlılıkların uyumunu kontrol edin

```cmd
.venv\Scripts\python.exe -m pip check
```

Beklenen sonuç: `No broken requirements found.` Bu mesaj, kurulu paketlerin bildirdiği bağımlılıklarda sorun bulunmadığını söyler; projenin yedi paketinin eksiksiz ve istenen sürümde olduğunu tek başına kanıtlamaz. **C2 ve C3 birlikte başarılı olmalıdır.** Sonra 5. adımda yerel sürücüyü seçip 6. adımda sistem kontrolünü çalıştırın.

## 5. Yerel EdgeDriver kullanımını zorunlu seçin

Aynı CMD penceresinde:

```cmd
set "KAP_TRANSPORT=edge"
set "KAP_EDGE_DRIVER=msedgedriver.exe"
```

Bu seçim yalnız bu CMD ve buradan başlatılan süreçlerde geçerlidir. Sürücü yoksa/bozuksa HTTP'ye sessiz geçiş yapılmaz. Sürücü farklı bir konumdaysa doğru göreli dosya yolunu kullanın; varsayılan örnek proje kökündeki sürücüdür.

## 6. Sistem kontrolünü çalıştırın

```cmd
.venv\Scripts\python.exe -B sistem_kontrolu.py
```

Beklenen ana satırlar:

```text
[OK] KAP baglantisi: edge
[OK] Yerel EdgeDriver: ...
SISTEM HAZIR
```

Henüz çalışma Excel'i yoksa Excel bulunamadı uyarısı normaldir; güncelleme başlamadan önce 8. adım uygulanmalıdır. Bu kontrol kaynak/kütüphane/ayar varlığını ve Python sözdizimini denetler; boş bırakılmış her kaynak dosyasını veya bütün iş kurallarını doğruladığı anlamına gelmez.

## 7. Excel'e dokunmadan KAP bağlantısını sınayın

```cmd
.venv\Scripts\python.exe kap_transport.py --check
```

Gerçek yerel EdgeDriver ile ayrı bir Edge oturumu açılır. Başarı halinde `KAP ana sayfa baglanti kontrolu basarili` mesajı beklenir; normal bitişte oturum kapanır. Bu komut Excel değiştirmez. Ana sayfa erişimi başarılı olması finansal API/Excel/PDF indirmelerinin veya tutarların doğruluğunu tek başına kanıtlamaz.

Sürücü/sürüm/proxy/sertifika/politika hatası varsa panelden güncellemeye geçmeyin. Kurum BT'si ile otomasyon oturumunun izinlerini doğrulayın; TLS/web güvenliğini veya kurum politikalarını kapatmayın.

## 8. Çalışma Excel'ini hazırlayın

Güncellenecek Excel'in yedeğini proje dışında alın. **Ayrı çalışma kopyasını** proje köküne şu adla koyun:

```text
butunlesik_banka_verisi.xlsx
```

Doğru kabul edilen referans Excel'i girdi olarak kullanmayın; başka yerde saklayın. Şablon banka sayfalarını ve mevcut dönemleri içermeli; boş/sıfır bayt `.xlsx` kullanmayın. İşlem sırasında çalışma Excel'ini Excel uygulamasında kapalı tutun.

```cmd
dir butunlesik_banka_verisi.xlsx
```

## 9. Web arayüzünü başlatın

```cmd
.venv\Scripts\python.exe panel.py
```

Tarayıcı otomatik açılır; gerçek `http://127.0.0.1:PORT` adresi CMD'de yazılır. **CMD penceresini kapatmayın.** GitHub sayfası uygulamanın kendisi değildir; panel bu bilgisayarda yerel olarak çalışır. HTML dosyasına doğrudan çift tıklamayın.

Otomatik tarayıcı açılmazsa CMD'deki adresi tarayıcıya yapıştırın. Gerekirse sunucuyu Ctrl+C ile durdurup şu alternatifle başlatın:

```cmd
.venv\Scripts\python.exe panel.py --no-browser --port 8766
```

Bu alternatifin adresi `http://127.0.0.1:8766` olur. Port dolu hatasında aynı komutu tekrar tekrar başlatmayın; `--port 8766` kısmını kaldırıp otomatik port kullanın. Yerel bağlantı kurumca engelleniyorsa BT ile görüşün.

## 10. İlk işlemi güvenli deneme olarak yapın

1. İlk deneme için tek banka seçin.
2. Paneldeki son dönem ve aranacak hedef dönemi kontrol edin. Örneğin son dönem Q1 ise hedef Q2; son dönem Q2 ise hedef Q3'tür. Yalnız tam bir sonraki dönem aranır.
3. **Güvenli deneme** seçili kalsın; gerçek güncellemeye geçmeyin.
4. Güncellemeyi başlatın, logu ve sonucu kontrol edin.
5. Deneme çıktısını Excel'de açıp doğru referansla karşılaştırın. Gerekli kontrollerden sonra gerçek güncellemeyi ayrıca değerlendirin.

Örnek DenizBank deneme çıktısı:

```text
butunlesik_banka_verisi_denizbank_DRY_RUN.xlsx
```

Kaynak dosyalar `denizbank/data/history/YYYYQn/` altında, audit raporları `outputs/otomatik_guncelleme_audit/` altında oluşur. Klasörler gerektiğinde oluşturulur. Güvenli denemede ana Excel değişmez; kaynak indirmeleri ve ayrı sonuç dosyası oluşur. Aynı banka denemesi tekrarlandığında önceki deneme çıktısı değişebilir; önemli sonucu farklı adla saklayın.

## Sonraki açılışlar

Kurulum başarıyla tamamlandıktan sonra paketleri her açışta yeniden kurmayın. Proje klasöründe yeni CMD açıp aşağıdakileri çalıştırmanız yeterlidir:

```cmd
set "KAP_TRANSPORT=edge"
set "KAP_EDGE_DRIVER=msedgedriver.exe"
.venv\Scripts\python.exe panel.py
```

Python kaynakları/bağımlılıklar güncellenmişse gereken kurulum ve kontrolleri tekrar uygulayın. Hata paylaşırken parola/token, kurum içi adres ve kişisel dosya bilgilerini ekran görüntüsünde gizleyin.
