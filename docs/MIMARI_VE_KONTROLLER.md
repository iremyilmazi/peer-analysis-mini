# Mimari ve işlem kontrolleri

## Giriş noktaları

`panel.py` tarayıcı panelini veya terminal girişini sunar. Seçilen banka ayrı Python alt sürecinde kendi `*_otomatik_guncelle.py` girişinden ortak motora bağlanır. Bu izolasyon banka PDF modüllerinin paylaştığı parser ayarlarının başka bankaya taşınmasını önler. İşlemler aynı Excel’de seri yürütülür.

Ana API’ler: `GET /api/banks`, `GET /api/context?bank=ing`, `GET /api/status?cursor=0`, `POST /api/run`, `POST /api/stop`, `POST /api/open`. Arayüz `127.0.0.1` üzerinde çalışır; JSON yazma istekleri ve yerel Host/Origin kontrolü kullanılır. Dosya açma isteği yalnız son çıktılar listesindeki yolları kabul eder.

Bu, internete yayınlanacak çok kullanıcılı bir web servisi değildir. Kimlik doğrulama, TLS, uzak kullanıcı yetkilendirmesi ve çoklu Excel kilitleme altyapısı içermez. Host’u `0.0.0.0` yapıp dışarıya açmayın. Kurumsal uç nokta politikaları ayrıca geçerlidir.

## Dönem seçimi

Ana Excel tek dönem otoritesidir. Son başlık çeyrek kimliğine çevrilir ve yalnız tam bir sonraki dönem seçilir. İndirilen klasörün daha ileri bir dönem içermesi hedefi değiştirmez. Eksik ara dönem atlanmaz. Boş dönem listesi tarihçe indirmeye dönüşmez; hata üretir.

`--period` bu sınırı gevşetmez. `--skip-download` sadece beklenen dönemin doğrulanmış yerel kaynağını kullanır. İndirici, metadata listesinde başka dönemler görse bile yalnız hedef dönem dosyalarını indirir. Kaynak işleme aşamasına da sadece hedef dönemin dizini verilir.

## Excel ve PDF

`entegre_excel_olustur.py` KAP’ın `.xls` adıyla sunduğu HTML/XML temelli tabloları okuyabilir. Bankaya özel PDF adaptörleri dipnot tablolarını çıkarır. Bu modüllerde kamuya açık eski rapor düzenlerine ilişkin özel kurallar/uyumluluk mantığı bulunabilir; bunlar ilk çalıştırmada eski rapor indirilmesini gerektirmez.

ING gelir sütunu, cari kümülatif dönem başlığından seçilir. Üç aylık ve önceki dönem sütunları seçilmez. Belirsiz başlık, yanlış dönem, boş finansal blok ve eksik hedef gelir satırı eşlemesi kayıt öncesi hata üretir. Açık yayımlanmış tire/sıfır ile gerçek eksik girdi ayrılır.

Tutar ölçeği, oran ve kur tabloları için mevcut banka/dönem kuralları korunur. Gelecekte KAP’ın birim veya rapor düzenini değiştirmesi ayrıca doğrulanmalıdır; bir dönemin test başarısı sonraki dönem doğruluğunu garanti etmez.

## Birleştirme ve kayıt

Geçici yeniden üretimdeki satırlar mevcut şablonun bölüm/etiket/tekrar kimliklerine eşlenir. Şablonda karşılığı olmayan kaynak satırları loglanabilir; yeni satır eklenmez. Yeni dönem kolonlarında önceden biçim varsa korunur; yoksa önceki dönemin biçimi devralınır.

Denetim sayfa sırası, birleşik hücreler, görünüm, satır/sütun boyutları ve hedef dışı değer/etkin biçim özelliklerini karşılaştırır. Yeni dönemde değişiklik oluşmadıysa kayıt reddedilir. İşlem başlangıcından sonra girdi hash’i değişmişse ana dosyanın üzerine yazılmaz. Güvenli deneme çıktısı ana Excel ile aynı yol olamaz.

Gerçek kayıtta önce yedek alınır; aday dosya aynı dizindeki geçici dosya üzerinden atomik olarak değiştirilir. Atomik kayıt, farklı programların eşzamanlı yazmasına karşı tam bir dosya-kilit protokolü veya otomatik geri alma sistemi değildir. Excel kapalı tutulmalıdır.

Denetim XML/ZIP’in byte düzeyinde aynılığını garanti etmez. Excel kütüphanesi eşdeğer XML öğe sıralamasını ve yardımcı alanları yeniden yazabilir. Ana koruma veri ve etkin biçim düzeyindedir. Mevcut motorun kapsamadığı özel Excel nesneleri için şablonun ayrıca test edilmesi gerekir.

## Kaynak düzeni ve geliştirme

Ortak motoru veya satır eşleme/ölçek kurallarını değiştirirken yalnız bir hücreyi referanstan kopyalayarak testi geçirmeyin. Kaynak, birim, dönem ve satır anlamı doğrulanmalıdır. Banka modüllerini rastgele silmeyin; TEB ve Garanti altında ortak parser bağımlılıkları bulunur.

Önerilen geliştirme sırası: çevrimdışı testler → ayrı Excel kopyasıyla güvenli deneme → referans ve koruma karşılaştırması → gerektiğinde izole canlı tekrar testi. Gerçek kullanıcı verisini test fixture’ı veya GitHub eki yapmayın.
