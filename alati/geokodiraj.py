# -*- coding: utf-8 -*-
"""Popunjava Lat/Lng u izlaz/ULICE_GEO.csv preko OpenStreetMap (Nominatim).

Pokretanje:
    python3 alati/geokodiraj.py

Nominatim dozvoljava 1 upit u sekundi, zato skripta traje par minuta.
Ono što ne nađe ostaje prazno - te ulice se dopune ručno u Google Sheetsu
(ili preko menija "Geokodiraj ulice", koji koristi Google geokoder).
"""
import csv
import json
import os
import sys
import time
import urllib.parse
import urllib.request

KORIJEN = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PUT = os.path.join(KORIJEN, "izlaz", "ULICE_GEO.csv")
API = "https://nominatim.openstreetmap.org/search?"
ZAGLAVLJE = {"User-Agent": "udruzenje-penzionera-bos-krupa/1.0 (evidencija ogrjeva)"}

# Granice šireg područja Bosanske Krupe - odbacujemo pogotke izvan njih.
OKVIR = (44.60, 45.15, 15.80, 16.60)  # lat_min, lat_max, lng_min, lng_max


def upit(tekst):
    url = API + urllib.parse.urlencode({
        "q": tekst, "format": "json", "limit": 1, "countrycodes": "ba"})
    zahtjev = urllib.request.Request(url, headers=ZAGLAVLJE)
    with urllib.request.urlopen(zahtjev, timeout=30) as o:
        return json.loads(o.read().decode())


def nadji(ulica, mjesto):
    varijante = []
    if ulica and ulica != "(bez ulice)":
        varijante.append(f"{ulica}, {mjesto}, Bosna i Hercegovina")
        varijante.append(f"{ulica}, {mjesto}")
    varijante.append(f"{mjesto}, Bosna i Hercegovina")
    for v in varijante:
        try:
            r = upit(v)
        except Exception as e:
            print("  greška:", e, file=sys.stderr)
            r = []
        time.sleep(1.1)
        if r:
            lat, lng = float(r[0]["lat"]), float(r[0]["lon"])
            if OKVIR[0] <= lat <= OKVIR[1] and OKVIR[2] <= lng <= OKVIR[3]:
                return lat, lng, ("ulica" if v is varijante[0] else "približno")
    return None, None, "nije nađeno"


def main():
    with open(PUT, encoding="utf-8") as f:
        redovi = list(csv.DictReader(f))
    polja = list(redovi[0].keys())
    if "Tačnost" not in polja:
        polja.append("Tačnost")

    for r in redovi:
        if r.get("Lat"):
            continue
        lat, lng, tacnost = nadji(r["Ulica"], r["Mjesto"])
        r["Lat"] = "" if lat is None else round(lat, 6)
        r["Lng"] = "" if lng is None else round(lng, 6)
        r["Tačnost"] = tacnost
        print(f'{r["Mjesto"]:16s} {r["Ulica"]:28s} {tacnost:12s} {r["Lat"]},{r["Lng"]}')

    with open(PUT, "w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=polja)
        w.writeheader()
        w.writerows(redovi)
    nadjeno = sum(1 for r in redovi if r["Lat"])
    print(f"\nKoordinate ima {nadjeno}/{len(redovi)} ulica.")


if __name__ == "__main__":
    main()
