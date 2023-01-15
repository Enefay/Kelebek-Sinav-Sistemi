# Kelebek-Sinav-Sistemi
kelebek sınav sistemi c#

[Ornek_Ogrenci_Listesi.xls](https://github.com/Enefay/Kelebek-Sinav-Sistemi/files/10419752/Ornek_Ogrenci_Listesi.xls)
Sisteme yüklenecek örncek excel dosyayı bu sekilde olmalıdır.


![site](https://user-images.githubusercontent.com/102833949/212531210-1ac1f3bb-7d76-49a5-a4f2-aaf5e537a513.png)
Sitenin genel görünümü bu şekilde. 

Yanlış formatta veya düzende bir excel dosyası yüklenirse hata mesajı ekrana geliyor. Eğer doğruysa sisteme yüklenen öğrenci listesi sol tarafta yer alıyor. Ardından kullanıcının sınıf seçimi yapması gerekiyor. Seçilen sınıf kontenjan sayısı öğrenci sayısından fazla ise "Rastgele Kaydet ve Devam Et" butonu aktif oluyor. Bu butonu tıklandığında ise öğrenciler rastgele sınıflara dağıtılıyor ve bu kayıtlar veritabanına kaydediliyor. Ardından seçilen sınıfların adında excel dosyaları oluşturularak hangi öğrenci hangi sınıftaysa o excel dosyasına kaydediliyor. 
Arama butonunda ise geçerli bir numara girilmediğinde hata döndürüyor. Eğer aranan öğrenci numarasıyla bir öğrenci var ise o öğrencinin giriş kartını "Resimler" klasörüne kaydediyor.

Ekran Görüntüleri

![ogrencitablosu](https://user-images.githubusercontent.com/102833949/212531528-182b787f-19de-45ca-977b-1c0c0e13e424.png)

Sisteme yüklenen dosyada bir sorun yok ise öğrenciler veritabanına kaydediliyor. Eğer o öğrenci daha önce veritabanında yer alıyor ise aynı öğrenci eklenmiyor.


![yerlestirmetablosu](https://user-images.githubusercontent.com/102833949/212531578-cb23534e-b3bb-4b6c-9d5f-8ebe2b6d2566.png)
"Rastgele Kaydet ve Devam et" butonuna tıkladıktan sonra öğrenciler seçilen sınıflara rastgele dağıtılıyor ve veritabanına kaydediliyor. Eğer o öğrenci daha önce bir sınıfa atandıysa seçilen sınıflara göre tekrar güncelleniyor.


![dosyalar](https://user-images.githubusercontent.com/102833949/212531632-d7956cb4-c3c0-4a42-a83e-c6abf75f4911.png)
Veritabanına kaydedildikten sonr oluşan excel dosyaları


![excelderay](https://user-images.githubusercontent.com/102833949/212531698-4d2a0606-4c46-4a4c-af12-a36cb44e4958.png)



Öğrenci Arama

![resdmlrdosya](https://user-images.githubusercontent.com/102833949/212531724-6b219aa1-f521-475e-b0e9-18ae1d65757f.png)

Öğrenci no ile arama başarılı ise o öğrencinin giriş kartı "Resimler" Klasörüne kaydediliyor.

![123456](https://user-images.githubusercontent.com/102833949/212531764-57103078-deb2-42ec-b534-478cc664c78f.jpg)



