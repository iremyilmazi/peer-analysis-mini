# Test kapsamı ve sınırları

## Önceki canlı finansal doğrulama

18–19 Ekim 2026’da kullanıcının çalışma Excel’inin ayrı kopyalarında 13 banka için solo 2026Q2 KAP Excel/PDF akışı test edildi. ING’nin paket bağımlılığı ve cari gelir sütunu sorunu düzeltildikten sonra, birleşik sonuçtaki Q2 alanları kullanıcı referansındaki 12.236 dolu hücreyle eşleşti. ING 19 Ekim’de boş kaynak/önbellekle yeniden indirildi; diğer 12 bankanın 18 Ekim başarılı sonuçları korundu.

Gerçek kullanıcı Excel’i, referans, kaynak PDF/Excel dosyaları, kişisel yollar ve ayrıntılı özel audit raporları bu depoya dahil edilmez. Yukarıdaki bilgi önceki yerel doğrulamanın özetidir; veri içermeyen dağıtımdan tek komutla o referans testi tekrarlanamaz.

## Bu temiz projedeki otomatik testler

`python -m unittest discover -s tests -v` geçici sentetik Excel’lerle, sahte indirme/ayrıştırma adımları ve gerçek yerel HTTP istekleriyle çalışır. Testler gerçek KAP dosyası indirmez ve kullanıcının Excel’ine dokunmaz.

Kapsam: 13 banka kayıt/giriş noktası ve runtime bağımlılığı; tek sonraki dönem ve yıl geçişi; yeni dönem yok akışı; yanlış hedefin reddi; deneme çıktısı güvenliği; yedek ve atomik kayıt akışı; dışarıdan girdi değişikliği; eski değer/biçim koruması; ING başlık tabanlı sütun seçimi; sıfır/tire/eksik ayrımı; panel banka bağlamı ve JSON API; varsayılan güvenli mod; paralel başlatma engeli; Host/Origin ve çıktı yolu sınırları.

Web paneli 13 banka seçimine uyarlandı. Saf JavaScript sözdizimi ve backend/API davranışı ayrıca kontrol edilir. Tarayıcıda insan kullanımı, şirket bilgisayarı politikaları, gerçek Excel uygulamasının hesaplama davranışı ve bütün bankaların bu yeni klasörden gerçek ağ üzerinden tekrar çalışması bu otomatik testlerin kapsamı dışındadır.

## 19 Ekim 2026 hazırlık sonucu

- Python 3.12.14 ortamında **31 testin 31’i başarılı** (`unittest`, 10,438 saniye).
- `sistem_kontrolu.py`: altı doğrudan kütüphane, ana dosyalar, 13 bankanın gerekli kaynakları ve Python sözdizimi kontrolleri başarılı.
- `node --check banka_paneli/app.js`: JavaScript sözdizimi başarılı. Node yalnız hazırlık kontrolünde kullanıldı; uygulamanın kurulumu için gerekmez.
- Projede kullanıcı Excel’i bulunmadığından sistem kontrolü beklenen Excel yok uyarısını verdi. Bu, veri içermeyen kaynak dağıtımı için normaldir; gerçek güncelleme için çalışma kopyası ayrıca eklenmelidir.

Test ortamında mevcut Python kütüphaneleri kullanıldı; boş bir şirket bilgisayarında sıfırdan paket kurulumu denenmedi. Yeni bir bağımlılık veya finansal kural değişikliğinde önce testleri çalıştırın; kritik değişiklikte ayrı çalışma kopyası ve güvenilir referansla canlı doğrulama yapın.

## Yerel EdgeDriver güncellemesi — 19 Ekim 2026

Yeni bağlantı katmanı ile **55 testin 55'i başarılı** (Python 3.12.14, 9,146 saniye). Önceki 31 teste ek olarak yerel sürücü arama, proje köküne göre yol çözme, zorunlu Edge modu, hatada HTTP'ye geçmeme, açık Service yolu, gerçek Selenium DriverFinder'ın Manager'ı atlaması, ortam değişkeninin sürücüyü değiştirememesi, JSON POST, binary PDF aktarımı, aynı oturumun tekrar kullanımı, HTTP 429, adres sınırları ve oturum kapanışı sınandı.

Selenium 4.51.0 ve bağımlılıkları proje dışındaki test dizinine ayrıca kuruldu. Sistem kontrolü yedi kütüphaneyi, 13 bankanın kaynaklarını ve Python sözdizimini doğruladı. Tarayıcıya gönderilen JavaScript ayrıca Node ile ağsız ortamda POST/binary/başlık/boyut sınırı/hata senaryolarında sınandı.

**Gerçek EdgeDriver dosyası bu dağıtımda yoktur.** Bu hazırlıkta gerçek Edge oturumu/KAP finansal indirmesi ve hedef şirket bilgisayarının proxy/SSO/politika uyumu denenmedi. Otomatik testlerde tarayıcı yanıtları sahtedir; DriverFinder testi gerçek Selenium yol seçim kodunu çalıştırır, `.exe` çalıştırmaz. Hedefte `kap_transport.py --check` ve ardından ayrı Excel kopyasıyla güvenli deneme yapılmalıdır. Referans Excel'ler değiştirilmedi ve depoya eklenmedi.

## Tek dosyalık boş iskelet — 19 Ekim 2026

`BOS_PROJE_OLUSTUR.py` için 14 test eklendi: manifestteki 96 dosyanın eksiksiz ve sıfır bayt oluşturulması, mevcut dosyanın içeriği/zamanının korunması, ikinci çalıştırmada üzerine yazılmaması, tek indirilen dosya ile bağımsız çalışma, varsayılan hedef, listeleme modunda yazma olmaması, dosya/klasör çakışmalarında ön kontrol, hedef dışına çıkma ve link/junction reddi, dosya listesinin mevcut kaynaklarla eşleşmesi.

Testler geçici klasörlerde çalışır; gerçek projeyi veya kullanıcı Excel'lerini boşaltmaz. Yardımcının ürettiği dosyalar yalnız ad/klasör iskeletidir; çalışır uygulama veya gerçek Excel/driver değildir. Genel test komutu yeni testleri de otomatik keşfeder. Son tam çalıştırmada **69 testin 69'u başarılı** (Python 3.12.14, 10,887 saniye).

## DenizBank gider tutar türü düzeltmesi — 19 Ekim 2026

Yeni 10 regresyonla **79 testin 79'u başarılı** (Python 3.12.14, 13,277 saniye). [Düzeltme ve hedef bilgisayarda uygulama yönergesi](DENIZBANK_GIDER_DUZELTMESI.md).

Gerçek 2026Q2 yerel KAP Excel/PDF kaynakları Q1 çalışma dosyasının ayrı kopyasında uçtan uca işlendi: Excel okuma, 21 PDF tablosu, dönem birleştirme ve koruma denetimi başarılı; audit 1.337.433 hücre kontrolü bildirdi. Bu sefer canlı ağ indirmesi yapılmadı; şirket bilgisayarında çalıştırılmadı.

Önceki Q2 sonucuyla tüm banka sayfalarında değer, veri türü ve hücre stili karşılaştırıldı: yalnız DenizBank gider dipnotundaki **sekiz `--` metni sayısal 0 oldu**; stil farkı yok. İlgili dipnotta 20/20 toplam tutarı sayısal, 40/40 TL/YP boş alanı korunmuş. Diğer tutarlar, geçmiş dönemler ve diğer bankalar önceki sonuçla aynı. Kaynak PDF'deki cari dönem satırları ayrıca incelendi. Özgün girdi, referans ve önceki sonuç dosyalarının SHA-256 değerleri değişmedi.

Eski v10 önbellekle başlatılan gerçek test PDF'yi yeniden ayrıştırıp v11 yazdı. Eski referansta aynı `--` metinlerinin bulunması, önceki 12.236 hücre eşleşmesinin tek başına sayısal tür doğruluğunu kanıtlamadığını gösterdi; yeni testler bu açığı kapatır. Bu sonuç diğer bankaların tüm tablolarını yeniden kaynak PDF'leriyle doğruladığımız anlamına gelmez.

## Yerel kaynak varsa KAP'a bağlanmama — 20 Ekim 2026

Yedi yeni regresyonla **86 testin 86'sı başarılı** (Python 3.12.14, 23,559 saniye). Yeni testlerde 13 banka ayrı ayrı çalıştırılarak geçerli yerel hedef kaynakları varken indiricinin ve `kap_connection` girişinin hiç çağrılmadığı, yalnız beklenen çeyreğin işlendiği ve audit'te indirilen dönem listesinin boş kaldığı doğrulandı. Deneme girdisi ve kaynak dosyalarının SHA-256 değerleri korundu.

Ek kapsam: iki kaynaktan birinin veya ikisinin eksik olması; sıfır bayt, yanlış imza ve giriş sayfası kaynakları; başka çeyreğin hedef yerine kullanılmaması; açık dönem seçimi; eski zorla indirme argümanının geçerli yerel kaynakları geçersiz kılmaması; sonraki PDF işleme hatasında tekrar KAP'a dönülmemesi ve ana girdinin değişmemesi.

Bu testler sentetik Excel girdileri ve imza/boyut kontrolüne uygun sentetik kaynaklarla işlem yönlendirmesini sınar. Excel/PDF ayrıştırma adımları bu yeni senaryolarda taklit edilir; gerçek KAP finansal içeriğinin doğruluğu, canlı ağ bağlantısı ve şirket bilgisayarı bu çalıştırmada yeniden sınanmadı. Mevcut finansal kurallar ve kaynak doğrulamaları değiştirilmedi. Hedef bilgisayara yalnız güncel `banka_otomatik_guncelle.py` aktarılmalıdır; [CMD rehberinde](CMD_ILE_KURULUM_VE_BASLATMA.md) uygulama adımları ve beklenen log vardır.
