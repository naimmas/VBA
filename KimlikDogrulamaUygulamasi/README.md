# Kimlik Doğrulama

Excel VBA ile kimlik bilgilerini doğrulama aracıdır.

## Kullanım

İlk olarak VBA Code ekranından Tools -> References'a girip 
- Microsoft WinHTTP Services
- Microsoft XML, v6.0

Referans dosyalarını projemize eklememiz gerekmektedir, ardından Excel penceresine girip formül çubuğuna aşağıdaki fonkisyonu yazabiliriz

```bash
=kimlik(metin olarak T.C. KİMLİK NO. ,metin olarak AD ,metin olarak SOYAD, Excel tarih formatı olarak DOGUM TARİHİ, metin olarak UYRUK)
```
