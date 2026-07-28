# -*- coding: utf-8 -*-
"""Pravi podaci/SINDIKAT.csv iz teksta ekstraktovanog Word dokumenta.

Original je .doc ('_PD_US__SINDIKAT.doc') koji Google/LibreOffice u ovom
okruženju nije mogao otvoriti; tekst je izvučen alatom `catdoc` i ručno
provjeren red po red (tabela u Wordu je razbijena na nepravilne redove pa
automatski parser ne bi bio pouzdan). Spisak ima 3 cjeline (Direkcija,
Direkcija-nastavak, Pogon gospodarenja) - četvrta "DVOMET" tabela na kraju
dokumenta je isključena jer ponavlja iste ljude (provjereno po imenu i
otpremnici). Jedan red (r.b. 20 u trećoj cjelini) je u dokumentu oštećen
(binarni "šum" umjesto imena) i preskočen - vidljivo je da je odobreno
7,00 m3 i isporučeno, ali ime nije čitljivo.

Kolone su svedene na isti oblik kao ostali izvori: nema matičnog broja ni
adrese (dokument ih ne sadrži), pa Mjesto ostaje "Bosanska Krupa" osim kod
jednog upisa gdje je mjesto eksplicitno navedeno (Polje, Cazin).
"""
import csv
import os

KORIJEN = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# (prezime i ime, telefon, mjesto, odobreno_m3, datum, otpremnica, isporuceno_m3, napomena)
REDOVI = [
    # --- Direkcija, dispozicija 12/26-R ---
    ("ALAGIĆ ISMET", "", "", 10.50, "04.05.2026", "046646/046647", 10.50, ""),
    ("BEŠIĆ SANEL", "", "", 10.50, "", "", 0, ""),
    ("BRATIĆ MIREL", "", "", 7.00, "03.06.2026", "004109", 7.00, ""),
    ("BURZIĆ DENIS", "", "", 3.50, "", "", 0, ""),
    ("DŽAFIĆ ALMA", "", "", 7.00, "", "", 0, ""),
    ("GROŠIĆ JASMIN", "", "", 7.00, "", "", 0, ""),
    ("GROŠIĆ MEHO", "", "", 3.50, "", "", 0, ""),
    ("HASANAGIĆ FARUK", "", "", 7.00, "", "", 0, ""),
    ("KAHRIĆ EDIS", "", "", 10.50, "23.07.2026", "055681", 10.50, ""),
    ("KASIĆ LEJLA", "", "", 10.50, "", "", 0, ""),
    ("DERVIŠEVIĆ ALMIR", "", "", 7.00, "", "", 0, ""),
    ("MAHMIĆ EMIR", "", "", 7.00, "", "", 0, ""),
    ("MEHIĆ ELVIRA", "", "", 10.50, "", "", 0, ""),
    ("MESIĆ AZEMINA", "", "", 3.50, "", "", 0, ""),
    ("MUSTAFIĆ MEHMED", "", "", 10.50, "29.04.2026", "046627/046628", 10.50, ""),
    ("SULJIĆ SAFETA", "", "", 3.50, "", "", 0, ""),
    ("MAHMIĆ ELVIR", "", "", 7.00, "", "", 0, ""),
    ("REKIĆ NEZIR", "", "", 7.00, "", "", 0, ""),
    ("REKIĆ OSMAN", "", "", 7.00, "", "", 0, ""),
    ("ŠARIĆ ELVIRA", "", "", 7.00, "05.05.2026", "004002", 7.00, ""),
    ("HALKIĆ MEDINA", "", "", 7.00, "18.06.2026", "051540", 7.00, ""),
    ("KABILJAGIĆ AMINA", "", "", 7.00, "", "", 0, ""),
    ("SUBAŠIĆ EMSAD", "", "", 3.50, "23.07.2026", "055679", 3.50, ""),
    ("MUŠINOVIĆ BEKIR", "", "", 7.00, "", "", 0, ""),
    ("REKIĆ SAIF", "", "", 7.00, "", "", 0, ""),
    ("KABILJAGIĆ MERIMA", "", "", 3.50, "05.05.2026", "004003", 3.50, ""),
    ("TATAREVIĆ HALID", "", "", 7.00, "", "", 0, ""),
    # --- Direkcija, dispozicija 14/26-R i 31/26-R (nastavak) ---
    ("HASANOVIĆ AMIR", "", "", 7.00, "", "", 0, "dispozicija 14/26-R"),
    ("SARAČEVIĆ KENAN", "", "", 3.50, "", "", 0, "dispozicija 14/26-R"),
    ("EMRIĆ ŠEFIK", "", "", 10.50, "", "", 0, "dispozicija 14/26-R"),
    ("HADŽIĆ BEHRUDIN", "061/600-615", "Polje, Cazin", 10.50, "", "", 0,
     "dispozicija 31/26-R"),
    # --- Pogon gospodarenja, dispozicija 03/26-R ---
    # r.b. 20 u originalu je oštećen (nečitko ime), preskočen
    ("GROŠIĆ EDIN", "", "", 10.50, "06.05.2026 / 22.06.2026",
     "004009/003897", 7.00, ""),
    ("HADŽIPAŠIĆ IRFAN", "", "", 7.00, "05.06.2026 / 21.07.2026",
     "004088/004673", 7.00, ""),
    ("HADŽIĆ JASMIN", "", "", 7.00, "29.05.2026", "003936", 7.00, ""),
    ("HADŽIPAŠIĆ IBRAHIM", "", "", 3.50, "05.05.2026", "004006", 3.50, ""),
    ("HARBAŠ MEHMEDALIJA", "", "", 10.50, "15.06.2026", "004159", 7.00, ""),
    ("HARČEVIĆ IBRAHIM", "", "", 7.00, "29.05.2026", "003935", 7.00, ""),
    ("HODŽIĆ AVDO", "", "", 7.00, "", "", 0, ""),
    ("ISMAILOVSKI AMRA", "", "", 10.50, "11.06.2026", "004090", 10.50, ""),
    ("JOGIĆ MUJO", "", "", 3.50, "01.06.2026", "004078", 3.50, ""),
    ("KOVAČEVIĆ ASIM", "", "", 7.00, "07.07.2026", "004562", 7.00, ""),
    ("MAHMIĆ HUSEIN", "", "", 7.00, "", "", 0, ""),
    ("MAHMUTOVIĆ MIRZA", "", "", 10.50, "23.07.2026", "004738", 10.50, ""),
    ("MUŠELJIĆ ESAD", "", "", 7.00, "23.04.2026", "043741", 7.00, ""),
    ("NUHANOVIĆ AMIRA", "", "", 7.00, "06.05.2026 / 24.07.2026",
     "004011/004748", 7.00, ""),
    ("PAŠIĆ ALADIN", "", "", 3.50, "10.07.2026", "004627", 3.50, ""),
    ("PAŠIĆ FEDIS", "", "", 7.00, "", "", 0, ""),
    ("PORIĆ EDIN", "", "", 7.00, "", "", 0, ""),
    ("PORIĆ JASMIN", "", "", 3.50, "", "", 0, ""),
    ("SAMARDŽIĆ HAZIM", "", "", 7.00, "23.07.2026", "004740", 3.50, ""),
    ("SALKIĆ ADNAN", "", "", 7.00, "30.04.2026", "046636", 7.00, "višemetrica"),
    ("ŠARIĆ VELAD", "", "", 7.00, "29.05.2026", "003933", 7.00, ""),
    ("TULIĆ AMIR", "", "", 7.00, "06.05.2026", "004007/003932", 7.00, ""),
    ("VELAGIĆ JASMIN", "", "", 7.00, "", "", 0, ""),
    ("ŠERTOVIĆ ELVIRA", "", "", 3.50, "", "", 0, ""),
    ("SAMARDŽIĆ ISMET", "", "", 7.00, "23.07.2026", "055680", 7.00, ""),
    ("SARAĐAN RAMIZ", "", "", 7.00, "18.06.2026", "051539", 7.00, ""),
    ("KOVAČEVIĆ NURIJA", "", "", 7.00, "07.07.2026", "004561", 7.00, ""),
    ("KOVAČEVIĆ SEAD", "", "", 7.00, "07.07.2026", "004563", 7.00, ""),
    ("BUŽIMKIĆ SABAHUDIN", "", "", 7.00, "06.07.2026", "004475", 7.00, ""),
    ("BEĆIREVIĆ OMER", "", "", 7.00, "29.05.2026", "003934", 7.00, ""),
    ("ĆEHIĆ AZRA", "", "", 7.00, "24.04.2026", "043750", 7.00, "dispozicija 68/26-R"),
]


def main():
    put = os.path.join(KORIJEN, "podaci", "SINDIKAT.csv")
    with open(put, "w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(["rb", "prezime_ime", "telefon", "mjesto", "odobreno_m3",
                    "datum", "otpremnica", "isporuceno_m3", "napomena"])
        for i, r in enumerate(REDOVI, start=1):
            ime, tel, mjesto, odobreno, datum, otpr, isp, nap = r
            w.writerow([i, ime, tel, mjesto or "Bosanska Krupa", odobreno,
                        datum, otpr, isp, nap])
    print(f"Upisano {len(REDOVI)} redova u {put}")
    print(f"Odobreno ukupno: {sum(r[3] for r in REDOVI):.2f} m3, "
          f"isporučeno: {sum(r[6] for r in REDOVI):.2f} m3")


if __name__ == "__main__":
    main()
