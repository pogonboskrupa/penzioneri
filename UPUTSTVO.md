# Evidencija isporuke ogrjeva – Udruženje penzionera Bosanska Krupa

Sve radi u **Google Sheetsu**: unos isporuke, zbir po ulicama i mjestima, te
**karta** na kojoj se vidi gdje je još ostalo za otpremiti i kome je već isporučeno.

**Dvije vrste ogrjeva vode se odvojeno:**

| Vrsta | Izvorni fajl | Listovi |
|---|---|---|
| **CIJEPANO** | `DRVA_PENZ_2025.xls` | `PODACI_CIJEPANO`, `ULICE_CIJEPANO` |
| **U DUGOM** | `PENZ_BOS_KRUPA_U_DUGOM.xls` | `PODACI_U_DUGOM`, `ULICE_U_DUGOM` |

Zajednički su samo `MJESTA` (zbirni pregled po vrstama) i `ULICE_GEO`
(koordinate ulica za kartu).

---

## 1. Šta je pripremljeno

Fajl **`izlaz/PENZIONERI_DRVA.xlsx`** je očišćena verzija oba spiska:

- **spojeni dupli nazivi mjesta** – `BOS.KRUPA`, `Bos.Krupa`, `Bos Krupa`,
  `BOSANSKA KRUPA` → **Bosanska Krupa** (isto za Otoku, Jezerski, Pištaline…);
- **spojeni dupli nazivi ulica**, ali **samo unutar istog mjesta** (kolona E),
  jer ista ulica postoji u više mjesta – npr. *Unska* u Bosanskoj Krupi i
  *Unska* u Bosanskoj Otoci ostaju **dvije različite stavke**;
- skinuti kućni brojevi (`27. Juli br 7` → `27 Juli`) i razriješene skraćenice
  (`ALEJA ZL. LJ` → `Aleja Zlatnih Ljiljana`, `101 MUSLIM` → `101 Muslimanska`);
- **sve količine svedene na m³**: u spisku cijepanog KOL je bio u prostornim
  metrima (5 prm = 3,5 m³; 10 prm = 7 m³), u spisku „u dugom“ je već bio m³;
- izračunato **Isporučeno / Preostalo / Status** za svakog korisnika
  (isporuka je prepoznata po upisanom datumu, otpremnici ili količini).

Trenutno stanje:

| Vrsta | Korisnika | Ulica | Odobreno | Isporučeno | Preostalo |
|---|---|---|---|---|---|
| CIJEPANO | 358 | 71 | 2030,0 m³ | 1113,5 m³ | **980,0 m³** |
| U DUGOM | 45 | 28 | 341,0 m³ | 189,0 m³ | **152,0 m³** |

---

## 2. Postavljanje (jednom, ~10 minuta)

1. **Napravi tabelu**: Google Drive → *Novo → Google tabele → Uvezi* →
   učitaj `izlaz/PENZIONERI_DRVA.xlsx` (opcija *Zamijeni tabelu*).
   Dobiješ listove `PODACI_CIJEPANO`, `PODACI_U_DUGOM`, `ULICE_CIJEPANO`,
   `ULICE_U_DUGOM`, `MJESTA`, `ULICE_GEO`.
2. **Dodaj skriptu**: u tabeli → *Proširenja → Apps Script*.
   - u fajl `Code.gs` zalijepi sadržaj `apps-script/Kod.gs`;
   - *+ → HTML* → nazovi fajl **`Karta`** → zalijepi `apps-script/Karta.html`;
   - sačuvaj (💾) i zatvori.
3. **Osvježi tabelu** (F5) – pojavi se meni **🪵 Drva**.
4. Prvi put klikni bilo koju stavku menija i **odobri dozvole**
   (Google pita „Autoriziraj“ → tvoj nalog → *Napredno → Nastavi*).

---

## 3. Svakodnevni rad

### Upis isporuke
Na listu `PODACI_CIJEPANO` ili `PODACI_U_DUGOM` klikni na red penzionera →
**🪵 Drva → Upiši isporuku za označeni red** → upiši m³ i broj otpremnice.
Skripta sama upiše datum, novo stanje, status i osvježi sve zbirove.

*(Može i ručno: samo upiši broj u kolonu „Isporučeno m3“, pa pokreni
„Osvježi sažetke“.)*

### Pregled po ulicama
**🪵 Drva → 1. Osvježi sažetke** – ponovo izračuna:
- `ULICE_CIJEPANO` i `ULICE_U_DUGOM`: po ulici → broj korisnika, odobreno,
  isporučeno, **preostalo m³**, koliko je korisnika gotovo, koliko čeka, i
  status ulice (🟢 ZAVRŠENO / 🟡 U TOKU / 🔴 NIJE POČETO, obojeno);
- `MJESTA`: zbir po mjestu i vrsti ogrjeva + postotak realizacije;
- `ULICE_GEO`: jedan red po ulici s preostalim količinama obje vrste
  (koordinate se čuvaju, ne brišu se pri osvježavanju).

**Plan rute za dan**: sortiraj `ULICE_GEO` po „Preostalo ukupno m3“ i
kreni od vrha – ili to isto pogledaj na karti.

### Karta
**🪵 Drva → 2. Geokodiraj ulice** pa **3. Otvori kartu**.

Na karti:
- **zeleno** = ta ulica je u potpunosti isporučena, **crveno** = ima još
  preostalo (djelimično ili ništa); veličina kruga = koliko m³ je ostalo;
- gore biraš **Cijepano / U dugom / Obje** i „Prikaži samo ulice gdje je ostalo“;
- prelaskom miša preko kruga vidi se ulica, koliko je ostalo/isporučeno i
  **imena onih koji čekaju** (s kategorijom), a klik otvara oblačić s
  **punim spiskom imena i prezimena** i količinama;
- desno je isti spisak detaljnije: **kome je isporučeno** (✓ s datumom i
  otpremnicom) i **ko još čeka** (koliko m³, telefon), plus dugme
  **Navigacija (Google Maps)** za vozača kamiona.

**Tačnost lokacije** – geokodiranje sad radi u tri koraka da se izbjegnu
pogrešni pogoci (npr. da ulica iz Bosanske Otoke ispadne u centru Bosanske
Krupe) i da se snađe i kad je naziv ulice upisan približno:
1. prvo se traži **tačna ulica**; rezultat se prihvata samo ako Google vrati
   pun (ne približan) pogodak unutar područja opštine – takva ulica je
   označena punim bijelim obrubom kruga;
2. ako tačna ulica nije nađena (npr. neko je upisao „Bihacka” bez č/ć/ž/š,
   ili skraćeno kao „Dž Čaušević”), traži se **najsličniji naziv ulice u
   istom mjestu koja je već tačno geokodirana** (npr. postojeća „Bihaćka”)
   i preuzima se njena lokacija – krug ima isprekidan obrub i piše
   „(približno, po sličnoj ulici …)”; ako nijedna ulica u tom mjestu nije
   dovoljno slična, ide se dalje;
3. ako ni to ne uspije, koristi se **centar naselja** (kolona Mjesto) –
   isto isprekidan obrub, piše samo „(približno)”;
4. ako ni naselje nije prepoznato, ulica ide u **spisak „Nesvrstano”** u
   lijevom/bočnom panelu na karti (bez pina) – nju postavljaš ručno:
   klikni **Postavi pin** pored ulice u spisku, pa klikni na kartu tačno
   gdje treba (npr. gdje živi kupac) – pin se odmah snimi u `ULICE_GEO` s
   oznakom „ručno” i geokodiranje je više neće dirati.

Ponovno pokretanje **Geokodiraj ulice** ne dira ulice već označene „ulica”
ili „ručno” – samo pokušava poboljšati one označene „slično”, „naselje” ili
„nesvrstano”. Na koju je ulicu nešto namapirano vidi se u koloni
**„Podudaranje sa”** u `ULICE_GEO`.

### Kategorija korisnika (penzioner / RVI / sindikat)
U listovima `PODACI_*` postoji kolona **Kategorija**. U izvornim `.xls`
spiskovima te podjele nema, pa su svi upisani kao `PENZIONER` – promijeni gdje
treba. **🪵 Drva → Postavi izbor kategorija** ubaci padajući izbor
(PENZIONER / RVI / SINDIKAT / OSTALO) da se ne kuca ručno.

Kategorija se vidi uz svako ime na karti i u spisku, a na javnoj stranici
postoji i filter po kategoriji (tada se i kubici zbrajaju samo za nju).

### Kontrola duplikata
**🪵 Drva → Provjeri moguće duplikate ulica** ispisuje slične nazive u istom
mjestu (npr. `Voloder` / `Volodor`). Ispravi naziv u `PODACI_*` i osvježi sažetke.

---

## 4. Javna stranica povezana s tabelom (web app + GitHub Pages)

Stranica **ne drži svoju kopiju podataka** – svaki put čita tabelu, pa se svaka
izmjena u Google Sheetsu odmah vidi na stranici (i sama se osvježava svakih 5 minuta).

### 4.1 Objavi stranicu iz tabele
1. U Apps Scriptu dodaj i treći fajl: **+ → HTML → `Stranica`** →
   zalijepi `apps-script/Stranica.html` (uz već postojeće `Kod.gs` i `Karta`).
2. Gore desno **Deploy → New deployment → tip: Web app**:
   - *Execute as*: **Me** (tvoj nalog čita tabelu),
   - *Who has access*: **Anyone with the link** (ili *Anyone within…* / samo ti).
3. Kopiraj dobiveni link (`…/exec`) – to je već gotova stranica, radi i na telefonu.

Na stranici su: kartice sa zbirovima (odvojeno cijepano / u dugom / ukupno),
karta s bojama po statusu, spisak ulica sortiran po preostalim kubicima,
pretraga po ulici, mjestu ili prezimenu, i detalji kome je isporučeno.

> **Telefoni:** ako link dijeliš šire, koristi `…/exec?telefoni=0` – tada se
> brojevi telefona ne prikazuju. Imena penzionera i količine su lični podaci,
> pa link ne objavljuj javno bez potrebe (najbolje *Anyone with the link*
> poslan samo vozaču i članovima uprave).

### 4.2 (Opcionalno) GitHub Pages adresa
Ako želiš lijepu stalnu adresu tipa `https://pogonboskrupa.github.io/penzioneri/`:
1. U `docs/index.html` upiši svoj `…/exec` link u red `var ADRESA_APLIKACIJE = '';`.
2. Na GitHubu: *Settings → Pages → Source: Deploy from a branch →
   grana `main`, folder `/docs`*.

Stranica na Pagesu je samo okvir koji prikazuje web aplikaciju – podaci i dalje
dolaze direktno iz Google Sheetsa, ništa se ne kopira u repozitorij.

---

## 5. Ponovna priprema iz .xls fajlova (opcionalno)

Ako stignu novi/ispravljeni `.xls` spiskovi, zamijeni ih u `podaci/` i pokreni:

```bash
python3 alati/pripremi.py      # čišćenje + spajanje duplikata + zbirovi
python3 alati/geokodiraj.py    # (opcionalno) popuni Lat/Lng preko OpenStreetMapa
```

`geokodiraj.py` traži slobodan pristup internetu i nije pokrenut pri ovoj
pripremi (mreža je bila zatvorena), pa je kolona `Lat`/`Lng` prazna –
**koordinate se dobiju iz same tabele**, stavkom menija
„2. Geokodiraj ulice“, koja koristi Googleov geokoder.

Pravila spajanja mjesta i ulica su u `alati/normalizacija.py`
(`MJESTA`, `ALIJASI_ULICA`, `SKRACENICE`) – tu se dodaje svaki novi izuzetak.

Napomena: `alati/pripremi.py` piše novi `izlaz/PENZIONERI_DRVA.xlsx` iz izvornih
`.xls` fajlova, pa **ne pokreći ga preko isporuka koje su upisane u Google
Sheetsu** – u tom slučaju prvo preuzmi aktuelnu tabelu iz Sheetsa.
