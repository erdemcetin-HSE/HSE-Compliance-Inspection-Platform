# HSE Platform MVP Plani (Direktor Odakli)

## 1) Hedef
Kolay kullanilabilir, profesyonel ve yonetici seviyesinde karar aldiran bir HSE platformu:
- Haftalik ve aylik KPI raporlama
- Operasyonel veri girisi (inspection, PPE, PTW, observation, risk, egitim)
- CEO ve senior management icin ozet dashboard

## 2) Roller
- Platform Owner / Super Admin
- HSE Direktor
- HSE Uzmani
- Saha Sefi
- Taseron Sorumlusu
- CEO / Senior Management (read-only agirlikli)

## 3) Ana Moduller

### A) KPI ve Raporlama (Haftalik / Aylik)
Temel KPI kartlari:
- Toplam Inspection
- Acik / Kapali PTW
- Acik / Kapali Observation
- Yuksek Risk kayit sayisi
- PPE kullanim ve stok trendi
- Egitim tamamlama orani

Rapor ozellikleri:
- Tarih filtresi (hafta, ay, ozel aralik)
- Bolge, taseron, departman filtreleri
- PDF/Excel export
- Drill-down (karttan detay listeye inis)

### B) Inspection Veri Girisi
Inspection form alanlari:
- Personel ad
- Personel soyad
- Taseron firma
- Bolge
- Isin niteligi
- Inspection tarihi/saati
- Bulgular ve aciklama
- Risk seviyesi (Dusuk/Orta/Yuksek/Kritik)
- Duzeltici faaliyet ve termin tarihi
- Durum (Acik/Kapali)

Ekler:
- Dosya yukleme (rapor, tutanak vb.)
- Sertifika dokumani
- Fotograf yukleme

### C) PPE Modulu
Veri alanlari:
- PPE turu
- Giren miktar
- Kullanilan miktar
- Depoda kalan miktar (otomatik hesap + manuel duzeltme logu)
- Lokasyon / depo
- Son guncelleme tarihi

Gosterimler:
- Stok seviyesi bar chart
- Kullanilan vs kalan trend line
- Kritik stok alarmi (esik altina dusunce)

### D) PTW (Permit To Work)
Alanlar:
- PTW no
- Is tipi
- Baslangic/bitis tarihi
- Sorumlu kisi
- Bolge
- Durum: Acik / Kapali
- Notlar ve ekler

### E) Observation Modulu
Alanlar:
- Observation tipi (guvenli/guvensiz)
- Konum
- Aciklama
- Sorumlu
- Durum: Acik / Kapali
- Kapanis tarihi
- Kanit dosyalari/fotograf

### F) Risk Management
- Risk kaydi olusturma
- Risk seviyesi: Dusuk, Orta, Yuksek, Kritik
- Olasilik x Etki matrix skoru
- Aksiyon atama ve takip
- Gecikmis aksiyon listesi

### G) Egitim Bilgileri
Calisan egitim kayitlari:
- Calisan ad soyad
- Egitim adi
- Egitim tarihi
- Gecerlilik tarihi
- Sertifika no
- Sertifika dosyasi
- Durum: Gecerli / Yaklasan Son Tarih / Suresi Dolmus

## 4) CEO ve Senior Management Dashboard
Gormek istedigi ozetler:
- Kurumsal HSE skoru (kompozit KPI)
- Son 30 gun olay/inspection trendi
- En riskli bolgeler / taseronlar
- Acik kritik aksiyonlar
- PTW ve observation closure oranlari
- Egitim uyum orani

Dashboard prensipleri:
- 1 ekranda yonetsel ozet
- Kirmizi-sari-yesil durum renkleri
- Kart -> detay sayfa gecisi

## 5) Veri Modeli (MVP tablolar)
- employees
- contractors
- regions
- inspections
- inspection_attachments
- ppe_items
- ppe_transactions
- ptw_records
- observations
- risks
- trainings
- training_certificates
- kpi_snapshots_weekly
- kpi_snapshots_monthly

## 6) Is Kurallari
- Tum kritik modullerde Acik/Kapali durum zorunlu
- Dosya yuklemelerde format ve boyut kontrolu
- Kritik risklerde otomatik bildirim
- Gecerliligi bitecek egitimlerde hatirlatici
- KPI hesaplari gecelik job ile snapshotlanir

## 7) MVP Yol Haritasi (6 Hafta)
1. Hafta 1: Inspection + ek dosya/fotograf + temel listeleme
2. Hafta 2: PPE stok/harcama + grafikler
3. Hafta 3: PTW + Observation (acik/kapali takip)
4. Hafta 4: Risk matrix + aksiyon takibi
5. Hafta 5: Egitim kayitlari + sertifika yonetimi
6. Hafta 6: CEO dashboard + haftalik/aylik rapor export

## 8) Basari Kriterleri
- HSE ekipleri 2 dakikadan kisa surede kayit acabilmeli
- CEO dashboard acilisinda 5 saniye altinda ozet gelmeli
- Haftalik/aylik rapor tek tikla indirilebilmeli
- Kritik risk ve dusuk PPE stok alarmlari gecikmesiz gorulebilmeli

## 9) Sonraki Faz (Opsiyonel)
- Mobil saha uygulamasi
- QR ile PPE zimmet
- AI destekli risk tahmin modeli
- Otomatik mevzuat uyum kontrolu
