# -*- coding: utf-8 -*-
"""Zajednička pravila za čišćenje i spajanje duplih naziva mjesta i ulica."""
import re
import unicodedata

# --- osnovni alati -----------------------------------------------------------

DIA = {
    "Č": "C", "Ć": "C", "Ž": "Z", "Š": "S", "Đ": "D",
    "č": "c", "ć": "c", "ž": "z", "š": "s", "đ": "d",
}


def bez_dijakritike(s: str) -> str:
    s = "".join(DIA.get(z, z) for z in s)
    return unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode()


def kljuc(s: str) -> str:
    """Ključ za poređenje: bez dijakritike, bez interpunkcije, VELIKA slova."""
    s = bez_dijakritike(str(s)).upper()
    s = re.sub(r"[^A-Z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


# --- mjesta (kolona E) -------------------------------------------------------
# Ključ (normalizovan) -> zvanični naziv koji se koristi svuda.
MJESTA = {
    "BOS KRUPA": "Bosanska Krupa",
    "BOSANSKA KRUPA": "Bosanska Krupa",
    "B KRUPA": "Bosanska Krupa",
    "KRUPA": "Bosanska Krupa",
    "BOS OTOKA": "Bosanska Otoka",
    "BOSANSKA OTOKA": "Bosanska Otoka",
    "OTOKA": "Bosanska Otoka",
    "JEZERSKI": "Jezerski",
    "PISTALINE": "Pištaline",
    "BUZIM": "Bužim",
    "MAHMIC SELO": "Mahmić Selo",
    "VOLODER": "Voloder",
    "ARAPUSA": "Arapuša",
    "VRANJSKA": "Vranjska",
    "PERNA": "Perna",
    "BUZIMSKA": "Bužim",
}


def normalizuj_mjesto(v) -> str:
    k = kljuc(v)
    if not k:
        return ""
    return MJESTA.get(k, k.title())


# --- ulice (kolona F) --------------------------------------------------------
# Skraćenice koje se pojavljuju u tabelama -> puni oblik.
SKRACENICE = {
    "MUSLIM": "MUSLIMANSKA",
    "MUSLIMANSKE": "MUSLIMANSKA",
    "SLAVNE": "SLAVNA",
    "VITESKE": "VITESKA",
    "BRIG": "BRIGADA",
    "BRIGADE": "BRIGADA",
    "BR IGADA": "BRIGADA",
    "ZL": "ZLATNIH",
    "LJ": "LJILJANA",
    "ZLATNIHLJ": "ZLATNIH LJILJANA",
    "GL": "GLAVICA",
    "UL": "",
    "ULICA": "",
    "BB": "",
    "KOR": "KORPUSA",
    "KORPUS": "KORPUSA",
    "HER": "HEROJA",
    "HEROJI": "HEROJA",
    "HEROJIH": "HEROJA",
    "OTOCKIH": "OTOCKIH",
    "N": "NASELJE",
    "NAS": "NASELJE",
    "MZ": "",
    "JULI": "JULA",
    "JULY": "JULA",
}

# Kućni broj na kraju naziva: "27 JULI br 7", "GAZIJSKA 12", "OMLADINSKA bb"
KUCNI_BROJ = re.compile(r"\s+(BR\.?\s*\d+[A-Z]?|BB|\d+[A-Z]?)$")


def normalizuj_ulicu(v) -> str:
    """Naziv ulice sveden na jedinstven oblik (bez kućnog broja i skraćenica)."""
    k = kljuc(v)
    if not k:
        return ""
    # skini kućni broj s kraja, ali NE i broj koji je dio imena ("101 MUSLIMANSKA")
    prije = None
    while prije != k:
        prije = k
        kandidat = KUCNI_BROJ.sub("", k).strip()
        # ako bi ostalo prazno ili samo broj, ne diraj
        if kandidat and not kandidat.isdigit():
            k = kandidat
    rijeci = []
    for r in k.split():
        z = SKRACENICE.get(r, r)
        if z:
            rijeci.extend(z.split())
    return " ".join(rijeci).strip()


def naziv_ulice_za_prikaz(sirovi: str) -> str:
    """Lijep prikaz iz originalnog zapisa: čuva č/ć/ž/š/đ, skida kućni broj."""
    s = re.sub(r"\s+", " ", str(sirovi)).strip(" .,")
    s = re.sub(r"\s+(br\.?\s*\d+[a-zA-Z]?|bb|\d+[a-zA-Z]?)$", "", s, flags=re.I)
    s = s.replace(".", ". ").replace(". ", ". ")
    rijeci = []
    for r in re.split(r"\s+", s):
        r = r.strip(",")
        if not r:
            continue
        rijeci.append(r if r[0].isdigit() else r[:1].upper() + r[1:].lower())
    return " ".join(rijeci).strip(" .") or naziv_ulice_za_prikaz_iz_kljuca(sirovi)


def naziv_ulice_za_prikaz_iz_kljuca(sirovi: str) -> str:
    n = normalizuj_ulicu(sirovi)
    return " ".join(w if w.isdigit() else w.capitalize() for w in n.split())


# --- spajanje sličnih naziva ulica unutar istog mjesta -----------------------

# Ručno potvrđeni parovi koje automatika ne prepozna (ključ -> ključ).
ALIJASI_ULICA = {
    "OTOCKOG BAT": "OTOC BATALJON",
    "OTOCKOG BATALJONA": "OTOC BATALJON",
    "OTOCKI HEROJA": "OTOCKIH HEROJA",
    "ALEJA ZL LJILJANA": "ALEJA ZLATNIH LJILJANA",
    "ALEJA ZLATNIH LJ": "ALEJA ZLATNIH LJILJANA",
    "V KORPUSA": "5 KORPUSA",
    "PETOG KORPUSA": "5 KORPUSA",
    "PATR LIGE": "PATRIOTSKE LIGE",
}

# Kozmetika za prikaz (skraćenice u originalnom zapisu -> puni naziv).
DISPLAY_ZAMJENE = [
    (r"\bGl\.?\b", "Glavica"),
    (r"\bZl\.?\s*Lj\.?\b", "Zlatnih Ljiljana"),
    (r"\bMuslim\.?\b", "Muslimanska"),
    (r"\bBat\.?\b", "Bataljona"),
    (r"\bBrig\.?\b", "Brigade"),
    (r"\bHer\.?\b", "Heroja"),
]


def uljepsaj_prikaz(s: str) -> str:
    for uzorak, zamjena in DISPLAY_ZAMJENE:
        s = re.sub(uzorak, zamjena, s)
    return re.sub(r"\s+", " ", s).strip()


def spoji_slicne(kljucevi_po_mjestu, prag=0.82):
    """{mjesto: {kljuc: broj pojavljivanja}} -> {(mjesto, kljuc): glavni_kljuc}.

    Spaja samo nazive u ISTOM mjestu (ista ulica postoji u više mjesta),
    koji se poklapaju preko `prag` i dijele prva tri znaka
    ("VOLODER"/"VOLODOR" da, "503 SLAVNA"/"511 SLAVNA" ne).
    """
    from difflib import SequenceMatcher

    mapa = {}
    for mjesto, brojaci in kljucevi_po_mjestu.items():
        # najčešći (pa najduži) nazivi postaju nosioci grupe
        poredak = sorted(brojaci, key=lambda k: (-brojaci[k], -len(k), k))
        nosioci = []
        for original in poredak:
            k = ALIJASI_ULICA.get(original, original)
            glavni = None
            for n in nosioci:
                prefiks = (k.startswith(n + " ") or n.startswith(k + " "))
                slicno = (k[:3] == n[:3]
                          and SequenceMatcher(None, k, n).ratio() >= prag)
                if prefiks or slicno:
                    glavni = n
                    break
            if glavni is None:
                nosioci.append(k)
                glavni = k
            mapa[(mjesto, original)] = glavni
    return mapa


# --- količine ----------------------------------------------------------------

PRM_U_M3 = 0.7  # 5 prm = 3,5 m3 ; 10 prm = 7 m3 (odnos iz postojećih tabela)


def broj(v):
    if v is None:
        return None
    s = str(v).strip().replace(",", ".")
    m = re.search(r"\d+(?:\.\d+)?", s)  # uzmi prvi broj ("14 m3" -> 14, ne 143)
    if not m:
        return None
    try:
        return float(m.group(0))
    except ValueError:
        return None


def kolicina_u_m3(sirovi, izvor_je_m3: bool):
    """KOL kolona: u starom spisku je u prostornim metrima, u novom već u m3."""
    n = broj(sirovi)
    if n is None:
        return None
    if izvor_je_m3 or "M3" in kljuc(sirovi):
        return round(n, 2)
    return round(n * PRM_U_M3, 2)
