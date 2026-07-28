# -*- coding: utf-8 -*-
"""Čisti oba .xls spiska i priprema tabele za Google Sheets.

VAŽNO: dvije vrste ogrjeva se vode ODVOJENO
    CIJEPANO  - podaci/DRVA_PENZ_2025.xls        (KOL u prostornim metrima)
    U DUGOM   - podaci/PENZ_BOS_KRUPA_U_DUGOM.xls (KOL već u m3)

Pokretanje:
    python3 alati/pripremi.py

Rezultat (folder izlaz/):
    PENZIONERI_DRVA.xlsx  - listovi PODACI_CIJEPANO, PODACI_U_DUGOM,
                            ULICE_CIJEPANO, ULICE_U_DUGOM, MJESTA, ULICE_GEO
    *.csv                 - isti listovi za uvoz u Google Sheets
"""
import os
import sys
from collections import Counter, defaultdict

import pandas as pd

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from normalizacija import (  # noqa: E402
    broj,
    kolicina_u_m3,
    naziv_ulice_za_prikaz,
    normalizuj_mjesto,
    normalizuj_ulicu,
    spoji_slicne,
    uljepsaj_prikaz,
)

KORIJEN = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

VRSTE = [
    # (vrsta, fajl, KOL je već u m3?)
    ("CIJEPANO", "podaci/DRVA_PENZ_2025.xls", False),
    ("U DUGOM", "podaci/PENZ_BOS_KRUPA_U_DUGOM.xls", True),
]

KOLONE = ["rb", "maticni", "prezime", "ime", "mjesto", "ulica", "telefon",
          "kol", "datum", "otpremnica", "kolicina", "oznaka"]

# Kategorija korisnika; u izvornim spiskovima nije zapisana, pa svi kreću kao
# PENZIONER, a u Google Sheetsu se po potrebi mijenja u RVI ili SINDIKAT.
PODRAZUMIJEVANA_KATEGORIJA = "PENZIONER"

REDOSLIJED = ["Vrsta", "Redni broj", "Matični broj", "Prezime", "Ime", "Kategorija",
              "Mjesto",
              "Ulica", "Telefon", "Odobreno m3", "Isporučeno m3", "Preostalo m3",
              "Status", "Datum isporuke", "Otpremnica", "Napomena",
              "Adresa ključ", "Original mjesto", "Original ulica"]


def ucitaj(vrsta, putanja, u_m3):
    sirovo = pd.read_excel(os.path.join(KORIJEN, putanja),
                           sheet_name="Naknadni spisak", header=None)
    d = sirovo.iloc[2:, :12].copy()
    d.columns = KOLONE

    redovi = []
    for _, r in d.iterrows():
        prezime = str(r.prezime).strip() if pd.notna(r.prezime) else ""
        ime = str(r.ime).strip() if pd.notna(r.ime) else ""
        if not prezime and not ime:
            continue
        odobreno = kolicina_u_m3(r.kol, u_m3) or 0.0
        isporuceno = broj(r.kolicina)
        datum = str(r.datum).strip() if pd.notna(r.datum) else ""
        otpremnica = str(r.otpremnica).strip() if pd.notna(r.otpremnica) else ""
        if isporuceno is None and (datum or otpremnica):
            isporuceno = odobreno  # otpremnica postoji, količina nije upisana
        isporuceno = isporuceno or 0.0
        preostalo = max(round(odobreno - isporuceno, 2), 0.0)
        redovi.append({
            "Vrsta": vrsta,
            "Redni broj": str(r.rb).strip().replace(".0", "") if pd.notna(r.rb) else "",
            "Matični broj": str(r.maticni).strip().replace(".0", "")
            if pd.notna(r.maticni) else "",
            "Prezime": prezime,
            "Ime": ime,
            "Kategorija": PODRAZUMIJEVANA_KATEGORIJA,
            "Mjesto": normalizuj_mjesto(r.mjesto),
            "_ulica_kljuc": normalizuj_ulicu(r.ulica) or "(BEZ ULICE)",
            "_ulica_original": str(r.ulica).strip() if pd.notna(r.ulica) else "",
            "Telefon": str(r.telefon).strip() if pd.notna(r.telefon) else "",
            "Odobreno m3": odobreno,
            "Isporučeno m3": isporuceno,
            "Preostalo m3": preostalo,
            "Status": ("ISPORUČENO" if preostalo <= 0.001 and isporuceno > 0
                       else "DJELIMIČNO" if isporuceno > 0 else "ZA ISPORUKU"),
            "Datum isporuke": datum,
            "Otpremnica": otpremnica,
            "Napomena": str(r.oznaka).strip() if pd.notna(r.oznaka) else "",
            "Original mjesto": str(r.mjesto).strip() if pd.notna(r.mjesto) else "",
            "Original ulica": str(r.ulica).strip() if pd.notna(r.ulica) else "",
        })
    return pd.DataFrame(redovi)


def ujednaci_ulice(podaci):
    """Spaja različite zapise iste ulice, ali samo unutar istog mjesta."""
    po_mjestu = defaultdict(Counter)
    for m, k in zip(podaci["Mjesto"], podaci["_ulica_kljuc"]):
        po_mjestu[m][k] += 1
    mapa = spoji_slicne(po_mjestu)

    # za svaki glavni ključ biramo najljepši originalni zapis (s dijakritikom)
    kandidati = defaultdict(Counter)
    for m, k, o in zip(podaci["Mjesto"], podaci["_ulica_kljuc"],
                       podaci["_ulica_original"]):
        glavni = mapa[(m, k)]
        prikaz = uljepsaj_prikaz(naziv_ulice_za_prikaz(o)) if o else "(bez ulice)"
        if prikaz:
            kandidati[(m, glavni)][prikaz] += 1

    def najbolji(par):
        c = kandidati.get(par)
        if not c:
            return "(bez ulice)"
        # prednost zapisu s dijakritikom, pa češćem, pa dužem
        return sorted(c, key=lambda p: (-any(z in p for z in "čćžšđČĆŽŠĐ"),
                                        -c[p], -len(p), p))[0]

    glavni = [mapa[(m, k)] for m, k in zip(podaci["Mjesto"], podaci["_ulica_kljuc"])]
    podaci["Ulica"] = [najbolji((m, g)) for m, g in zip(podaci["Mjesto"], glavni)]
    podaci["Adresa ključ"] = [f"{m}|{g}" for m, g in zip(podaci["Mjesto"], glavni)]
    return podaci


def sazetak_ulica(podaci):
    u = (podaci.groupby(["Mjesto", "Ulica", "Adresa ključ"], as_index=False)
         .agg(**{"Broj korisnika": ("Prezime", "count"),
                 "Odobreno m3": ("Odobreno m3", "sum"),
                 "Isporučeno m3": ("Isporučeno m3", "sum"),
                 "Preostalo m3": ("Preostalo m3", "sum")}))
    gotovi = (podaci[podaci["Status"] == "ISPORUČENO"]
              .groupby("Adresa ključ").size())
    u["Isporučeno korisnika"] = [int(gotovi.get(k, 0)) for k in u["Adresa ključ"]]
    u["Za isporuku korisnika"] = u["Broj korisnika"] - u["Isporučeno korisnika"]
    u["Status ulice"] = ["ZAVRŠENO" if p <= 0.001 else
                         ("U TOKU" if i > 0.001 else "NIJE POČETO")
                         for p, i in zip(u["Preostalo m3"], u["Isporučeno m3"])]
    for k in ("Odobreno m3", "Isporučeno m3", "Preostalo m3"):
        u[k] = u[k].round(2)
    return u.sort_values(["Preostalo m3", "Mjesto", "Ulica"],
                         ascending=[False, True, True]).reset_index(drop=True)


BOJA_ISPORUCENO = "FFD9EAD3"    # svijetlo zelena
BOJA_NEISPORUCENO = "FFF4CCCC"  # svijetlo crvena


def oboji_listove(writer, listovi):
    """Cijeli red korisnika/ulice u zelenoj ili crvenoj nijansi."""
    from openpyxl.styles import PatternFill

    zelena = PatternFill("solid", fgColor=BOJA_ISPORUCENO)
    crvena = PatternFill("solid", fgColor=BOJA_NEISPORUCENO)
    for naziv, df in listovi.items():
        if "Preostalo m3" not in df.columns:
            continue
        list_ = writer.sheets[naziv]
        for i, preostalo in enumerate(df["Preostalo m3"], start=2):
            ispuna = zelena if float(preostalo) <= 0.001 else crvena
            for kolona in range(1, len(df.columns) + 1):
                list_.cell(row=i, column=kolona).fill = ispuna


def main():
    dijelovi = {v: ucitaj(v, p, m) for v, p, m in VRSTE}
    sve = pd.concat(dijelovi.values(), ignore_index=True)
    sve = ujednaci_ulice(sve)
    sve = sve.sort_values(["Vrsta", "Mjesto", "Ulica", "Prezime", "Ime"],
                          kind="stable").reset_index(drop=True)

    po_vrsti = {v: sve[sve["Vrsta"] == v][REDOSLIJED].reset_index(drop=True)
                for v in dijelovi}
    ulice = {v: sazetak_ulica(d) for v, d in po_vrsti.items()}

    mjesta = (sve.groupby(["Vrsta", "Mjesto"], as_index=False)
              .agg(**{"Broj korisnika": ("Prezime", "count"),
                      "Odobreno m3": ("Odobreno m3", "sum"),
                      "Isporučeno m3": ("Isporučeno m3", "sum"),
                      "Preostalo m3": ("Preostalo m3", "sum")})
              .sort_values(["Vrsta", "Preostalo m3"], ascending=[True, False]))
    for k in ("Odobreno m3", "Isporučeno m3", "Preostalo m3"):
        mjesta[k] = mjesta[k].round(2)

    # jedan red po ulici (obje vrste zajedno) - osnova za kartu
    geo = (sve.groupby(["Mjesto", "Ulica", "Adresa ključ"], as_index=False)
           .size().drop(columns="size"))
    for v in dijelovi:
        s = ulice[v].set_index("Adresa ključ")
        geo[f"Preostalo {v} m3"] = [float(s["Preostalo m3"].get(k, 0.0))
                                    for k in geo["Adresa ključ"]]
        geo[f"Isporučeno {v} m3"] = [float(s["Isporučeno m3"].get(k, 0.0))
                                     for k in geo["Adresa ključ"]]
    geo["Preostalo ukupno m3"] = (geo["Preostalo CIJEPANO m3"] +
                                  geo["Preostalo U DUGOM m3"]).round(2)
    geo["Adresa za kartu"] = [
        (f"{u}, {m}, Bosna i Hercegovina" if u != "(bez ulice)"
         else f"{m}, Bosna i Hercegovina")
        for u, m in zip(geo["Ulica"], geo["Mjesto"])]
    geo["Lat"] = ""
    geo["Lng"] = ""
    geo = geo.sort_values("Preostalo ukupno m3", ascending=False).reset_index(drop=True)

    izlaz = os.path.join(KORIJEN, "izlaz")
    os.makedirs(izlaz, exist_ok=True)
    listovi = {
        "PODACI_CIJEPANO": po_vrsti["CIJEPANO"],
        "PODACI_U_DUGOM": po_vrsti["U DUGOM"],
        "ULICE_CIJEPANO": ulice["CIJEPANO"],
        "ULICE_U_DUGOM": ulice["U DUGOM"],
        "MJESTA": mjesta,
        "ULICE_GEO": geo,
    }
    with pd.ExcelWriter(os.path.join(izlaz, "PENZIONERI_DRVA.xlsx"),
                        engine="openpyxl") as w:
        for naziv, df in listovi.items():
            df.to_excel(w, sheet_name=naziv, index=False)
        oboji_listove(w, listovi)
    for naziv, df in listovi.items():
        df.to_csv(os.path.join(izlaz, f"{naziv}.csv"), index=False)

    for v, d in po_vrsti.items():
        print(f"[{v}] korisnika {len(d)}, ulica {len(ulice[v])}, "
              f"odobreno {d['Odobreno m3'].sum():.2f} m3, "
              f"isporučeno {d['Isporučeno m3'].sum():.2f} m3, "
              f"preostalo {d['Preostalo m3'].sum():.2f} m3")
    print(mjesta.to_string(index=False))


if __name__ == "__main__":
    main()
