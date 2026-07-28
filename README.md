# penzioneri – evidencija isporuke ogrjeva

Alati za praćenje isporuke drva penzionerima (Udruženje penzionera Bosanska Krupa):
čišćenje postojećih `.xls` spiskova, zbirovi po ulicama i mjestima i karta u Google Sheetsu.

**Cijepani ogrjev i ogrjev „u dugom“ vode se odvojeno.**

| Folder | Sadržaj |
|---|---|
| `podaci/` | izvorni `.xls` spiskovi |
| `alati/` | Python skripte za čišćenje (`pripremi.py`, `normalizacija.py`, `geokodiraj.py`) |
| `izlaz/` | `PENZIONERI_DRVA.xlsx` + CSV listovi za uvoz u Google Sheets |
| `apps-script/` | `Kod.gs` i `Karta.html` – meni i karta unutar Google Sheetsa |

Javna stranica (uvijek aktuelno stanje iz tabele): `apps-script/Stranica.html` + `docs/` za GitHub Pages.

Uputstvo korak po korak: **[UPUTSTVO.md](UPUTSTVO.md)**
