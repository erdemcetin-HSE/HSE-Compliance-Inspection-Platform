# HSE Uygulama Backlogu (Baslanabilir)

## Epic 1 - Inspection Management
### User Story 1.1
HSE uzmani olarak inspection kaydi acabilmek istiyorum, boylece saha bulgularini standart girebileyim.

Kabul Kriterleri:
- Personel ad, soyad, taseron firma, bolge, isin niteligi zorunlu alanlar.
- Risk seviyesi secimi zorunlu (Dusuk/Orta/Yuksek/Kritik).
- Durum varsayilan Acik.

### User Story 1.2
Inspection kaydina dosya, sertifika ve fotograf ekleyebilmek istiyorum.

Kabul Kriterleri:
- En az 1 fotograf yuklenebilmeli.
- Desteklenen formatlar: pdf, jpg, png.
- Dosya boyut limiti asildiginda hata mesaji gosterilmeli.

## Epic 2 - PPE Tracking
### User Story 2.1
Depo sorumlusu olarak PPE giren/kullanilan/kalan verilerini girmek istiyorum.

Kabul Kriterleri:
- Giren ve kullanilan girisi sonrasi kalan stok otomatik hesaplanir.
- Manuel duzeltme yapilirsa audit log tutulur.

### User Story 2.2
HSE direktor olarak PPE trend grafikleri gormek istiyorum.

Kabul Kriterleri:
- Son 12 hafta kullanim trend cizgisi gosterilir.
- Kritik stok esik altinda kirmizi uyarilir.

## Epic 3 - PTW
### User Story 3.1
Saha sefleri PTW kaydi acip acik/kapali takip edebilmeli.

Kabul Kriterleri:
- PTW no benzersiz olmalı.
- Acik kayitlar listesi ayrica filtrelenebilmeli.

## Epic 4 - Observation
### User Story 4.1
Observation kaydi acilip acik/kapali yonetilebilmeli.

Kabul Kriterleri:
- Observation tipi secilmeli.
- Kapanis tarihini kapali durumunda zorunlu yap.

## Epic 5 - Risk Management
### User Story 5.1
Risk seviyesi ve matrix skoru hesaplanmali.

Kabul Kriterleri:
- Olasilik x Etki ile skor otomatik hesaplanir.
- Yuksek/Kritik riskler dashboardda vurgulanir.

## Epic 6 - Training Records
### User Story 6.1
Calisan egitim kayitlari ve sertifikalari girilebilmeli.

Kabul Kriterleri:
- Gecerlilik tarihi zorunlu.
- Suresi dolmus sertifikalar ayri listelenir.

## Epic 7 - Executive Dashboard
### User Story 7.1
CEO ve senior management ozet KPI dashboard gormeli.

Kabul Kriterleri:
- Haftalik/aylik KPI kartlari tek ekranda.
- PTW/Observation closure oranlari gosterilmeli.
- En riskli bolgeler ve taseronlar listelenmeli.

## Epic 8 - Reporting
### User Story 8.1
Raporlar PDF/Excel alinabilmeli.

Kabul Kriterleri:
- Haftalik rapor tek tik export.
- Aylik rapor tek tik export.
- Filtrelenmis veri exportta korunur.

## Teknik Gorevler
- API: /inspections CRUD + upload endpoints
- API: /ppe-transactions + stock summary endpoint
- API: /ptw CRUD
- API: /observations CRUD
- API: /risks CRUD + scoring utility
- API: /trainings CRUD + expiry query
- API: /dashboard/executive-kpis
- Web: modul bazli formlar + tablolar + grafikler
- Web: role-based menu ve route guard
- Jobs: haftalik/aylik KPI snapshot scheduler
