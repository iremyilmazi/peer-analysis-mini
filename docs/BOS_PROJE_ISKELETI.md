# Tek dosyayla boş proje iskeleti oluşturma

`BOS_PROJE_OLUSTUR.py`, başka hiçbir proje dosyasına ihtiyaç duymayan, Python standart kütüphanesiyle çalışan bir yardımcıdır. Dosya/klasör adlarının sabit listesini kendi içinde taşır. İnternete bağlanmaz; Git, Selenium veya pip kurulumu gerektirmez. Bilgisayarda Python 3.11 veya daha yeni bir sürüm bulunmalıdır.

**Bu bir kurucu veya kod indirici değildir.** Oluşturulan dosyaların içerikleri tamamen boştur (0 bayt). Bunlarla uygulama çalışmaz. Dolu kaynak kodlarını/ayarları daha sonra ayrıca yerleştirmeniz gerekir. Oluşan boş `README.md` ve boş `BOS_PROJE_OLUSTUR.py` dosyaları da buna dahildir; çalıştırdığınız dolu yardımcı dosyayı ayrı tutun.

## Kullanım

1. GitHub'daki [BOS_PROJE_OLUSTUR.py](../BOS_PROJE_OLUSTUR.py) sayfasından **Download raw file** ile yalnız bu dosyayı indirin. HTML sayfası olarak kaydetmeyin; dosya adı `.py` ile bitmeli.
2. Dosyanın bulunduğu klasörde PowerShell açın.
3. İstediğiniz hedef klasörü belirterek çalıştırın:

```powershell
python .\BOS_PROJE_OLUSTUR.py --hedef "C:\Projeler\peer-analysis-mini"
```

Buradaki yol örnektir; kendi hedefinizi yazın. `python` yerine kurumunuzun onaylı `py -3.12` komutu veya Python'un gerçek yolu da kullanılabilir.

Hedef belirtmezseniz yardımcı dosyanın yanında `peer-analysis-mini-bos` adlı bir alt klasör oluşturulur:

```powershell
python .\BOS_PROJE_OLUSTUR.py
```

Önce yalnız dosya adlarını görmek için:

```powershell
python .\BOS_PROJE_OLUSTUR.py --listele
```

Bu seçenek hiçbir klasör/dosya oluşturmaz.

## Oluşan yapı

19 Ekim 2026 güncel sürümünde listede **97 dosya** vardır: 13 bankanın kaynak dosyaları, ortak motor, panelin HTML/JS dosyaları, yapılandırma dosyaları, gereksinimler, belgeler, testler ve bu yardımcı dosyanın kendi adı. CMD rehberi eklenmeden önceki sürümde 96 dosya vardı; önceden oluşturduğunuz iskeleti baştan kurmanız gerekmez. Eklenen rehber çalışma için zorunlu değildir ve GitHub'dan okunabilir. Alt klasörler dosya yollarından otomatik oluşturulur. Listede dosya adı olması, içerik de getirildiği anlamına gelmez.

`.git`, `.venv`, kütüphaneler, önbellek, indirilmiş raporlar, özel Excel'ler ve EdgeDriver bu yapıya dahil değildir. Özellikle **boş/sahte `.xlsx` veya `.exe` oluşturulmaz**. Gerçek çalışma Excel'inizi ve uyumlu gerçek `msedgedriver.exe` dosyasını ayrıca ekleyin. Çalışma sırasında gereken `data`/`outputs` dizinleri dolu uygulama tarafından gerektiğinde oluşturulur; geçmiş arşivi taşımanız gerekmez.

## Mevcut dosyaların güvenliği

- Olmayan dosya adları 0 bayt olarak oluşturulur.
- Mevcut dosyalar dolu veya boş fark etmeksizin **atlanır**; içerikleri, adları ve değiştirilme tarihleri korunur.
- Dosya adı yerinde klasör veya klasör adı yerinde dosya varsa işlem, ilk yazımdan önce hata verir.
- Link/junction üzerinden dosya oluşturulması ve hedef dışına çıkan manifest yolları reddedilir.
- Yardımcı hiçbir şeyi silmez, indirmez veya çalıştırmaz. Sürücüyü/kurum politikasını değiştirmez.
- Oluşturma sırasında izin/disk hatası veya kesinti yaşanırsa kısmi bir iskelet kalabilir; otomatik geri alma/silme yapılmaz. Yeniden çalıştırmak mevcutları atlayıp eksikleri tamamlar.

Tüm dosyaların boş olmasını istiyorsanız **yeni, boş bir hedef klasör** kullanın. Daha önceden dolu dosyalar bulunan bir klasörü boşaltmak bu aracın işi değildir.

Manifest, bu sürümün anlık görüntüsüdür; sonraki sürümlerde yeni dosya eklenirse güncel yardımcıyı yeniden indirin. Geliştirirken `FILES` listesini yeni kaynak yollarıyla güncelleyin; `tests/test_scaffold.py` liste ile kaynak ağacının eşleşmesini kontrol eder.
