# GitHub yükleme öncesi kontrol listesi

- [ ] Yalnız `peer-analysis-mini` klasörü depo kökü olacak; üst çalışma alanı yüklenmeyecek.
- [x] Kullanıcı isteğiyle `iremyilmazi/peer-analysis-mini` adlı public depo oluşturuldu. Lisans ayrıca kararlaştırılacak.
- [ ] `README.md` kurulum adımları kullanılacak; Python/venv dosyaları depoya alınmayacak.
- [ ] `python -m unittest discover -s tests -v` başarılı.
- [ ] `python sistem_kontrolu.py` kaynak ve kütüphaneleri doğruluyor. Excel’in dağıtımda bulunmaması beklenir.
- [ ] `git diff --cached --name-only` ve `git diff --cached --stat` elle incelendi.
- [ ] Excel, PDF, CSV, indirilen kaynak, yedek, cache ve gerçek audit/log dosyası listede yok.
- [ ] Kullanıcı adları, masaüstü yolları, e-posta adresleri ve kurum içi bilgi kaynaklara eklenmedi.
- [ ] Parola, token, özel anahtar, `.env`, proxy kimlik bilgisi ve pip yapılandırması yok.
- [ ] Gerçek test referansları veya ekran görüntüleri paylaşılmıyor.
- [ ] KAP kimlikleri kamuya açık banka eşleme bilgileridir; bunlar kullanıcı erişim anahtarı değildir.
- [ ] Daha önce takip edilmiş/commit edilmiş bir dosyanın `.gitignore` ile geçmişten silinmediği biliniyor.
- [ ] Uzak depo URL’si doğru ve kimlik bilgisi içermiyor.

Public kaynak depo: [iremyilmazi/peer-analysis-mini](https://github.com/iremyilmazi/peer-analysis-mini). Yükleme kapsamı yalnız temiz kaynaklar, arayüz, sentetik testler ve belgelerdir. Gerçek kullanıcı Excel’i, kaynak raporlar, Python ortamı ve önceki çalışma geçmişi dahil edilmez. README’de GitHub Desktop ile klonlama ve kurulum adımları bulunur. Kontrol listesi sonraki yüklemelerde de tekrar uygulanmalıdır.
