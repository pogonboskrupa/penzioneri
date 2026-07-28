/**
 * Udruženje penzionera Bosanska Krupa - praćenje isporuke ogrjeva.
 *
 * Dvije vrste ogrjeva vode se ODVOJENO:
 *   CIJEPANO  -> listovi PODACI_CIJEPANO i ULICE_CIJEPANO
 *   U DUGOM   -> listovi PODACI_U_DUGOM  i ULICE_U_DUGOM
 * Zajednički su samo MJESTA (zbirni pregled) i ULICE_GEO (koordinate za kartu).
 */

var VRSTE = [
  { naziv: 'CIJEPANO', podaci: 'PODACI_CIJEPANO', ulice: 'ULICE_CIJEPANO' },
  { naziv: 'U DUGOM', podaci: 'PODACI_U_DUGOM', ulice: 'ULICE_U_DUGOM' }
];
var LIST_MJESTA = 'MJESTA';
var LIST_GEO = 'ULICE_GEO';

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🪵 Drva')
    .addItem('1. Osvježi sažetke (ulice i mjesta)', 'osvjeziSazetke')
    .addItem('2. Geokodiraj ulice (za kartu)', 'geokodirajUlice')
    .addItem('3. Otvori kartu', 'otvoriKartu')
    .addSeparator()
    .addItem('Upiši isporuku za označeni red', 'upisiIsporuku')
    .addItem('Provjeri moguće duplikate ulica', 'provjeriDuplikate')
    .addItem('Postavi izbor kategorija (penzioner/RVI/sindikat)', 'postaviKategorije')
    .addItem('Postavi lozinku za pomjeranje pina na javnoj stranici', 'postaviLozinkuUredjivanja')
    .addToUi();
}

/* ------------------------------------------------------------------ pomoćno */

function list_(naziv) {
  var s = SpreadsheetApp.getActive().getSheetByName(naziv);
  if (!s) throw new Error('Nedostaje list "' + naziv + '".');
  return s;
}

function citaj_(naziv) {
  var v = list_(naziv).getDataRange().getValues();
  var zaglavlje = v.shift();
  var indeks = {};
  zaglavlje.forEach(function (h, i) { indeks[String(h).trim()] = i; });
  return { zaglavlje: zaglavlje, redovi: v, i: indeks };
}

function broj_(v) {
  if (v === '' || v === null || v === undefined) return 0;
  if (typeof v === 'number') return v;
  var m = String(v).replace(',', '.').match(/\d+(\.\d+)?/);
  return m ? parseFloat(m[0]) : 0;
}

function kljucAdrese_(mjesto, ulica) {
  return String(mjesto).trim() + '|' + String(ulica).trim().toUpperCase();
}

function upisi_(naziv, zaglavlje, redovi) {
  var s = SpreadsheetApp.getActive().getSheetByName(naziv);
  if (!s) s = SpreadsheetApp.getActive().insertSheet(naziv);
  s.clear();
  s.getRange(1, 1, 1, zaglavlje.length).setValues([zaglavlje])
    .setFontWeight('bold').setBackground('#e8eaed');
  if (redovi.length) {
    s.getRange(2, 1, redovi.length, zaglavlje.length).setValues(redovi);
  }
  s.setFrozenRows(1);
  s.autoResizeColumns(1, zaglavlje.length);
  return s;
}

/* --------------------------------------------------------- 1. sažeci ulica */

function osvjeziSazetke() {
  var geoPostojeci = ucitajKoordinate_();
  var poVrsti = {};
  var sveUlice = {};

  VRSTE.forEach(function (vrsta) {
    var t = citaj_(vrsta.podaci);
    var grupe = {};
    t.redovi.forEach(function (r) {
      var mjesto = String(r[t.i['Mjesto']]).trim();
      var ulica = String(r[t.i['Ulica']]).trim();
      if (!mjesto && !ulica) return;
      var k = kljucAdrese_(mjesto, ulica);
      if (!grupe[k]) {
        grupe[k] = { mjesto: mjesto, ulica: ulica, korisnika: 0, odobreno: 0,
                     isporuceno: 0, preostalo: 0, gotovih: 0 };
      }
      var g = grupe[k];
      var odobreno = broj_(r[t.i['Odobreno m3']]);
      var isporuceno = broj_(r[t.i['Isporučeno m3']]);
      var preostalo = Math.max(Math.round((odobreno - isporuceno) * 100) / 100, 0);
      g.korisnika++;
      g.odobreno += odobreno;
      g.isporuceno += isporuceno;
      g.preostalo += preostalo;
      if (preostalo <= 0.001 && isporuceno > 0) g.gotovih++;
      if (!sveUlice[k]) sveUlice[k] = { mjesto: mjesto, ulica: ulica };
    });
    poVrsti[vrsta.naziv] = grupe;

    var redovi = Object.keys(grupe).map(function (k) {
      var g = grupe[k];
      return [g.mjesto, g.ulica, k, g.korisnika,
        okrugli_(g.odobreno), okrugli_(g.isporuceno), okrugli_(g.preostalo),
        g.gotovih, g.korisnika - g.gotovih, statusUlice_(g)];
    }).sort(function (a, b) { return b[6] - a[6]; });

    var s = upisi_(vrsta.ulice,
      ['Mjesto', 'Ulica', 'Adresa ključ', 'Broj korisnika', 'Odobreno m3',
        'Isporučeno m3', 'Preostalo m3', 'Isporučeno korisnika',
        'Za isporuku korisnika', 'Status ulice'], redovi);
    obojiStatus_(s, 10, redovi.length);
    obojiPodatke_(vrsta.podaci);
  });

  osvjeziMjesta_(poVrsti);
  osvjeziGeo_(poVrsti, sveUlice, geoPostojeci);
  SpreadsheetApp.getActive().toast('Sažeci osvježeni.', 'Drva', 5);
}

function okrugli_(x) { return Math.round(x * 100) / 100; }

function statusUlice_(g) {
  if (g.preostalo <= 0.001) return 'ZAVRŠENO';
  return g.isporuceno > 0.001 ? 'U TOKU' : 'NIJE POČETO';
}

// Dvije boje kroz cijelu tabelu, kartu i javnu stranicu.
var BOJA_ISPORUCENO = '#d9ead3';    // svijetlo zelena - sve isporučeno
var BOJA_NEISPORUCENO = '#f4cccc';  // svijetlo crvena - ima još preostalo

function obojiStatus_(list, kolona, brojRedova) {
  if (!brojRedova) return;
  var opseg = list.getRange(2, kolona, brojRedova, 1);
  opseg.setBackgrounds(opseg.getValues().map(function (r) {
    return [r[0] === 'ZAVRŠENO' ? BOJA_ISPORUCENO : BOJA_NEISPORUCENO];
  }));
}

/** Boji cijeli red svakog korisnika prema tome je li isporuka završena. */
function obojiPodatke_(nazivLista) {
  var list = list_(nazivLista);
  var t = citaj_(nazivLista);
  if (!t.redovi.length) return;
  var brojKolona = t.zaglavlje.length;
  var boje = t.redovi.map(function (r) {
    var odobreno = broj_(r[t.i['Odobreno m3']]);
    var isporuceno = broj_(r[t.i['Isporučeno m3']]);
    var preostalo = Math.max(okrugli_(odobreno - isporuceno), 0);
    var boja = (preostalo <= 0.001 && isporuceno > 0)
      ? BOJA_ISPORUCENO : BOJA_NEISPORUCENO;
    var red = [];
    for (var i = 0; i < brojKolona; i++) red.push(boja);
    return red;
  });
  list.getRange(2, 1, boje.length, brojKolona).setBackgrounds(boje);
}

function osvjeziMjesta_(poVrsti) {
  var redovi = [];
  VRSTE.forEach(function (vrsta) {
    var zbir = {};
    var grupe = poVrsti[vrsta.naziv];
    Object.keys(grupe).forEach(function (k) {
      var g = grupe[k];
      if (!zbir[g.mjesto]) {
        zbir[g.mjesto] = { korisnika: 0, odobreno: 0, isporuceno: 0,
                           preostalo: 0, ulica: 0 };
      }
      var z = zbir[g.mjesto];
      z.korisnika += g.korisnika;
      z.odobreno += g.odobreno;
      z.isporuceno += g.isporuceno;
      z.preostalo += g.preostalo;
      z.ulica++;
    });
    Object.keys(zbir).sort().forEach(function (m) {
      var z = zbir[m];
      redovi.push([vrsta.naziv, m, z.ulica, z.korisnika, okrugli_(z.odobreno),
        okrugli_(z.isporuceno), okrugli_(z.preostalo),
        z.odobreno ? Math.round(z.isporuceno / z.odobreno * 100) + '%' : '0%']);
    });
  });
  upisi_(LIST_MJESTA, ['Vrsta', 'Mjesto', 'Broj ulica', 'Broj korisnika',
    'Odobreno m3', 'Isporučeno m3', 'Preostalo m3', 'Realizacija'], redovi);
}

function ucitajKoordinate_() {
  var mapa = {};
  var s = SpreadsheetApp.getActive().getSheetByName(LIST_GEO);
  if (!s) return mapa;
  var t = citaj_(LIST_GEO);
  t.redovi.forEach(function (r) {
    var k = String(r[t.i['Adresa ključ']]).trim();
    if (k && r[t.i['Lat']] !== '' && r[t.i['Lng']] !== '') {
      mapa[k] = {
        lat: broj_(r[t.i['Lat']]), lng: broj_(r[t.i['Lng']]),
        tacnost: t.i['Tačnost'] !== undefined ? String(r[t.i['Tačnost']] || '') : '',
        podudaranje: t.i['Podudaranje sa'] !== undefined
          ? String(r[t.i['Podudaranje sa']] || '') : ''
      };
    }
  });
  return mapa;
}

function osvjeziGeo_(poVrsti, sveUlice, koordinate) {
  var redovi = Object.keys(sveUlice).map(function (k) {
    var u = sveUlice[k];
    var c = poVrsti['CIJEPANO'][k] || { preostalo: 0, isporuceno: 0, korisnika: 0 };
    var d = poVrsti['U DUGOM'][k] || { preostalo: 0, isporuceno: 0, korisnika: 0 };
    var xy = koordinate[k] || { lat: '', lng: '', tacnost: '', podudaranje: '' };
    var adresa = u.ulica && u.ulica !== '(bez ulice)'
      ? u.ulica + ', ' + u.mjesto + ', Bosna i Hercegovina'
      : u.mjesto + ', Bosna i Hercegovina';
    return [u.mjesto, u.ulica, k,
      okrugli_(c.preostalo), okrugli_(c.isporuceno),
      okrugli_(d.preostalo), okrugli_(d.isporuceno),
      okrugli_(c.preostalo + d.preostalo), adresa, xy.lat, xy.lng, xy.tacnost,
      xy.podudaranje];
  }).sort(function (a, b) { return b[7] - a[7]; });

  var s = upisi_(LIST_GEO, ['Mjesto', 'Ulica', 'Adresa ključ',
    'Preostalo CIJEPANO m3', 'Isporučeno CIJEPANO m3',
    'Preostalo U DUGOM m3', 'Isporučeno U DUGOM m3',
    'Preostalo ukupno m3', 'Adresa za kartu', 'Lat', 'Lng', 'Tačnost',
    'Podudaranje sa'], redovi);
  if (redovi.length) {
    var opseg = s.getRange(2, 12, redovi.length, 1);
    opseg.setBackgrounds(opseg.getValues().map(function (r) {
      return [r[0] === 'ulica' ? '#d9ead3' : r[0] === 'naselje' ? '#fff2cc'
        : r[0] === 'ručno' ? '#cfe2f3' : r[0] === 'slično' ? '#d0e0fb'
        : r[0] === 'nesvrstano' ? '#f4cccc' : null];
    }));
  }
}

/* ---------------------------------------------------------- 2. geokodiranje */

// Granice šireg područja opštine Bosanska Krupa - odbacujemo pogotke izvan
// njih (npr. kad Google ne nađe ulicu pa vrati centar neke druge opštine).
var GEO_OKVIR = { latMin: 44.60, latMax: 45.15, lngMin: 15.80, lngMax: 16.60 };

function uOkviru_(loc) {
  return loc.lat >= GEO_OKVIR.latMin && loc.lat <= GEO_OKVIR.latMax &&
    loc.lng >= GEO_OKVIR.lngMin && loc.lng <= GEO_OKVIR.lngMax;
}

/**
 * Pokušava geokodirati tačnu ulicu. Ako Google ne nađe baš tu ulicu (ili
 * vrati samo približan pogodak - partial_match), NE prihvata taj rezultat
 * jer u praksi zna vratiti centar sasvim drugog mjesta.
 */
function geokodirajUlicu_(geokoder, adresaUlice) {
  if (!adresaUlice) return null;
  try {
    var o = geokoder.geocode(adresaUlice);
    if (o.status === 'OK' && o.results.length && !o.results[0].partial_match) {
      var loc = o.results[0].geometry.location;
      if (uOkviru_(loc)) return { lat: loc.lat, lng: loc.lng };
    }
  } catch (e) { /* nema rezultata */ }
  return null;
}

function geokodirajMjesto_(geokoder, adresaMjesta) {
  try {
    var o = geokoder.geocode(adresaMjesta);
    if (o.status === 'OK' && o.results.length) {
      var loc = o.results[0].geometry.location;
      if (uOkviru_(loc)) return { lat: loc.lat, lng: loc.lng };
    }
  } catch (e) { /* nema rezultata */ }
  return null;
}

var PRAG_SLICNOSTI_ULICE = 0.5;

function dodajUzorak_(poMjestu, mjesto, ulica, lat, lng) {
  mjesto = String(mjesto).trim();
  if (!poMjestu[mjesto]) poMjestu[mjesto] = [];
  poMjestu[mjesto].push({ ulica: String(ulica).trim(), lat: lat, lng: lng });
}

/**
 * Traži najsličniji naziv ulice koja je VEĆ tačno geokodirana u istom
 * mjestu - npr. upisano "Bihacka" (bez dijakritike ili skraćeno) nađe
 * već geokodiranu "Bihaćka" u istom gradu i preuzme njenu lokaciju.
 * To je pouzdanije nego centar cijelog naselja.
 */
function nadjiSlicnuUlicu_(poMjestu, mjesto, ulica) {
  var uzorci = poMjestu[String(mjesto).trim()] || [];
  var a = pojednostavi_(ulica), najbolji = null, ocjena = -1;
  // isti naziv bez dijakritike (npr. "Bihacka" / "Bihaćka") = najbolji mogući pogodak
  uzorci.forEach(function (u) {
    var o = slicnost_(a, pojednostavi_(u.ulica));
    if (o > ocjena) { ocjena = o; najbolji = u; }
  });
  return ocjena >= PRAG_SLICNOSTI_ULICE ? najbolji : null;
}

function geokodirajUlice() {
  var s = list_(LIST_GEO);
  var t = citaj_(LIST_GEO);
  var geokoder = Maps.newGeocoder().setRegion('ba');
  var uliceN = 0, slicnoN = 0, naseljeN = 0, nesvrstanoN = 0;
  var poMjestu = {};         // mjesto -> ulice već tačno geokodirane (uzorak za "slično")
  var zaDrugiPokusaj = [];   // redovi kojima tačna ulica nije nađena

  // uzorci iz ranijih pokretanja (kolone Tačnost = "ulica")
  t.redovi.forEach(function (r) {
    var tac = t.i['Tačnost'] !== undefined ? String(r[t.i['Tačnost']] || '') : '';
    if (tac === 'ulica' && r[t.i['Lat']] !== '' && r[t.i['Lng']] !== '') {
      dodajUzorak_(poMjestu, r[t.i['Mjesto']], r[t.i['Ulica']],
        broj_(r[t.i['Lat']]), broj_(r[t.i['Lng']]));
    }
  });

  // 1. prolaz - tačna ulica
  t.redovi.forEach(function (r, idx) {
    var tacnostSad = t.i['Tačnost'] !== undefined ? String(r[t.i['Tačnost']] || '') : '';
    if (tacnostSad === 'ručno' || tacnostSad === 'ulica') return;  // ne diraj

    var ulica = String(r[t.i['Ulica']] || '').trim();
    var mjesto = String(r[t.i['Mjesto']] || '').trim();
    var adresaUlice = ulica && ulica !== '(bez ulice)'
      ? ulica + ', ' + mjesto + ', Bosna i Hercegovina' : '';

    var rez = geokodirajUlicu_(geokoder, adresaUlice);
    if (adresaUlice) Utilities.sleep(200);
    if (rez) {
      upisiGeoRed_(s, t, idx, rez.lat, rez.lng, 'ulica', '');
      dodajUzorak_(poMjestu, mjesto, ulica, rez.lat, rez.lng);
      uliceN++;
    } else {
      zaDrugiPokusaj.push({ idx: idx, ulica: ulica, mjesto: mjesto });
    }
  });

  // 2. prolaz - slična ulica u istom mjestu, već tačno geokodirana
  var zaTreciPokusaj = [];
  zaDrugiPokusaj.forEach(function (stavka) {
    var slicna = stavka.ulica && stavka.ulica !== '(bez ulice)'
      ? nadjiSlicnuUlicu_(poMjestu, stavka.mjesto, stavka.ulica) : null;
    if (slicna) {
      upisiGeoRed_(s, t, stavka.idx, slicna.lat, slicna.lng, 'slično', slicna.ulica);
      slicnoN++;
    } else {
      zaTreciPokusaj.push(stavka);
    }
  });

  // 3. prolaz - centar naselja (Mjesto)
  zaTreciPokusaj.forEach(function (stavka) {
    var rez = geokodirajMjesto_(geokoder, stavka.mjesto + ', Bosna i Hercegovina');
    Utilities.sleep(200);
    if (rez) {
      upisiGeoRed_(s, t, stavka.idx, rez.lat, rez.lng, 'naselje', '');
      naseljeN++;
    } else {
      upisiGeoRed_(s, t, stavka.idx, '', '', 'nesvrstano', '');
      nesvrstanoN++;
    }
  });

  SpreadsheetApp.getUi().alert(
    'Tačno po ulici: ' + uliceN + '\n' +
    'Po sličnoj već geokodiranoj ulici u istom mjestu: ' + slicnoN + '\n' +
    'Približno, po naselju (Mjesto): ' + naseljeN + '\n' +
    'Nesvrstano (bez lokacije): ' + nesvrstanoN +
    (nesvrstanoN
      ? '\n\nNesvrstane ulice postavi ručno: 🪵 Drva → Otvori kartu → ' +
        'lijevi spisak "Nesvrstano" → Postavi pin → klik na kartu.'
      : ''));
}

function upisiGeoRed_(list, t, idx, lat, lng, tacnost, podudaranje) {
  var red = idx + 2;
  list.getRange(red, t.i['Lat'] + 1).setValue(lat);
  list.getRange(red, t.i['Lng'] + 1).setValue(lng);
  if (t.i['Tačnost'] !== undefined) list.getRange(red, t.i['Tačnost'] + 1).setValue(tacnost);
  if (t.i['Podudaranje sa'] !== undefined) {
    list.getRange(red, t.i['Podudaranje sa'] + 1).setValue(podudaranje);
  }
}

/**
 * Ručno postavljanje pina na kartu (klikom) za ulicu koja nema tačnu ili
 * ima pogrešnu lokaciju. Pretpostavlja tačnost "ručno" - geokodiranje je
 * više nikad automatski ne dira.
 */
function spremiRucnuKoordinatu(kljucAdrese, lat, lng) {
  var s = list_(LIST_GEO);
  var t = citaj_(LIST_GEO);
  for (var i = 0; i < t.redovi.length; i++) {
    if (String(t.redovi[i][t.i['Adresa ključ']]).trim() === kljucAdrese) {
      var red = i + 2;
      s.getRange(red, t.i['Lat'] + 1).setValue(lat);
      s.getRange(red, t.i['Lng'] + 1).setValue(lng);
      if (t.i['Tačnost'] !== undefined) {
        s.getRange(red, t.i['Tačnost'] + 1).setValue('ručno');
      }
      return true;
    }
  }
  throw new Error('Adresa nije nađena u ULICE_GEO: ' + kljucAdrese);
}

var SVOJSTVO_LOZINKE = 'LOZINKA_UREDJIVANJA';

/** Postavlja/mijenja lozinku kojom javna stranica smije pomjerati pinove. */
function postaviLozinkuUredjivanja() {
  var ui = SpreadsheetApp.getUi();
  var svojstva = PropertiesService.getScriptProperties();
  var trenutna = svojstva.getProperty(SVOJSTVO_LOZINKE);
  var odgovor = ui.prompt(
    trenutna ? 'Lozinka je već postavljena' : 'Postavi lozinku',
    'Upiši novu lozinku za pomjeranje pina na javnoj stranici ' +
    '(ostavi prazno da isključiš uređivanje na javnoj stranici).',
    ui.ButtonSet.OK_CANCEL);
  if (odgovor.getSelectedButton() !== ui.Button.OK) return;
  var lozinka = odgovor.getResponseText().trim();
  if (lozinka) {
    svojstva.setProperty(SVOJSTVO_LOZINKE, lozinka);
    ui.alert('Lozinka je spremljena. Vozač/administrator je unosi samo jednom - ' +
      'stranica je pamti dok ne zatvori karticu preglednika.');
  } else {
    svojstva.deleteProperty(SVOJSTVO_LOZINKE);
    ui.alert('Uređivanje karte na javnoj stranici je isključeno (nema lozinke).');
  }
}

/**
 * Isto što i spremiRucnuKoordinatu, ali za javnu web stranicu (doGet) gdje
 * je "Execute as: Me" - bez ove provjere bi svako s linkom mogao pomjerati
 * pinove. Radi samo ako je lozinka prethodno postavljena u meniju.
 */
function javniSpremiRucnuKoordinatu(kljucAdrese, lat, lng, lozinka) {
  var tacna = PropertiesService.getScriptProperties().getProperty(SVOJSTVO_LOZINKE);
  if (!tacna) {
    throw new Error('Uređivanje karte nije uključeno. U Sheetsu: 🪵 Drva → ' +
      'Postavi lozinku za pomjeranje pina na javnoj stranici.');
  }
  if (String(lozinka || '') !== tacna) {
    throw new Error('Pogrešna lozinka.');
  }
  return spremiRucnuKoordinatu(kljucAdrese, lat, lng);
}

/* ------------------------------------------------------------------ 3. karta */

function otvoriKartu() {
  var html = HtmlService.createHtmlOutputFromFile('Karta')
    .setWidth(1200).setHeight(800);
  SpreadsheetApp.getUi().showModalDialog(html, 'Karta isporuke ogrjeva');
}

/** Podaci za kartu: jedna tačka po ulici + spisak korisnika. */
function podaciZaKartu() {
  var geo = citaj_(LIST_GEO);
  var tacke = {};
  geo.redovi.forEach(function (r) {
    var lat = r[geo.i['Lat']], lng = r[geo.i['Lng']];
    var imaXY = lat !== '' && lng !== '';
    tacke[String(r[geo.i['Adresa ključ']]).trim()] = {
      mjesto: r[geo.i['Mjesto']],
      ulica: r[geo.i['Ulica']],
      lat: imaXY ? broj_(lat) : null,   // bez koordinata: samo u spisku, ne na karti
      lng: imaXY ? broj_(lng) : null,
      tacnost: geo.i['Tačnost'] !== undefined ? String(r[geo.i['Tačnost']] || '') : '',
      podudaranje: geo.i['Podudaranje sa'] !== undefined
        ? String(r[geo.i['Podudaranje sa']] || '') : '',
      vrste: {}
    };
  });

  VRSTE.forEach(function (vrsta) {
    var t = citaj_(vrsta.podaci);
    t.redovi.forEach(function (r) {
      var k = kljucAdrese_(r[t.i['Mjesto']], r[t.i['Ulica']]);
      var tacka = tacke[k];
      if (!tacka) return;
      if (!tacka.vrste[vrsta.naziv]) {
        tacka.vrste[vrsta.naziv] = { preostalo: 0, isporuceno: 0, korisnici: [] };
      }
      var v = tacka.vrste[vrsta.naziv];
      var odobreno = broj_(r[t.i['Odobreno m3']]);
      var isporuceno = broj_(r[t.i['Isporučeno m3']]);
      var preostalo = Math.max(okrugli_(odobreno - isporuceno), 0);
      v.preostalo += preostalo;
      v.isporuceno += isporuceno;
      v.korisnici.push({
        ime: String(r[t.i['Prezime']]) + ' ' + String(r[t.i['Ime']]),
        kategorija: t.i['Kategorija'] === undefined ? 'PENZIONER'
          : (String(r[t.i['Kategorija']] || 'PENZIONER').trim().toUpperCase()),
        telefon: String(r[t.i['Telefon']] || ''),
        odobreno: odobreno,
        isporuceno: isporuceno,
        preostalo: preostalo,
        datum: String(r[t.i['Datum isporuke']] || ''),
        otpremnica: String(r[t.i['Otpremnica']] || '')
      });
    });
  });

  var bezKoordinata = [];
  geo.redovi.forEach(function (r) {
    if (r[geo.i['Lat']] === '' || r[geo.i['Lng']] === '') {
      bezKoordinata.push({
        kljuc: String(r[geo.i['Adresa ključ']]).trim(),
        mjesto: String(r[geo.i['Mjesto']]),
        ulica: String(r[geo.i['Ulica']])
      });
    }
  });

  return {
    tacke: Object.keys(tacke).map(function (k) {
      var t = tacke[k];
      t.kljuc = k;
      return t;
    }),
    bezKoordinata: bezKoordinata
  };
}

/* ------------------------------------------- javna stranica (web aplikacija) */

/**
 * Otvara se na URL-u koji dobiješ preko "Deploy -> New deployment -> Web app".
 * Podaci se čitaju iz tabele u trenutku otvaranja stranice, pa je prikaz
 * uvijek onakav kakva je tabela.
 *
 * Parametri u URL-u:
 *   ?telefoni=0   sakriva brojeve telefona (za javno dijeljenje linka)
 */
function doGet(e) {
  var param = (e && e.parameter) || {};
  var t = HtmlService.createTemplateFromFile('Stranica');
  t.prikaziTelefone = param.telefoni !== '0';
  return t.evaluate()
    .setTitle('Isporuka ogrjeva - Udruženje penzionera Bosanska Krupa')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Sve što stranici treba: zbirovi, ulice s koordinatama i spisak korisnika. */
function javniPodaci(prikaziTelefone) {
  var karta = podaciZaKartu();
  var zbir = {};
  VRSTE.forEach(function (vrsta) {
    var t = citaj_(vrsta.podaci);
    var z = { korisnika: 0, odobreno: 0, isporuceno: 0, preostalo: 0, gotovih: 0 };
    t.redovi.forEach(function (r) {
      var odobreno = broj_(r[t.i['Odobreno m3']]);
      var isporuceno = broj_(r[t.i['Isporučeno m3']]);
      if (!odobreno && !isporuceno) return;
      var preostalo = Math.max(okrugli_(odobreno - isporuceno), 0);
      z.korisnika++;
      z.odobreno += odobreno;
      z.isporuceno += isporuceno;
      z.preostalo += preostalo;
      if (preostalo <= 0.001 && isporuceno > 0) z.gotovih++;
    });
    ['odobreno', 'isporuceno', 'preostalo'].forEach(function (k) {
      z[k] = okrugli_(z[k]);
    });
    zbir[vrsta.naziv] = z;
  });

  if (prikaziTelefone === false) {
    karta.tacke.forEach(function (t) {
      Object.keys(t.vrste).forEach(function (v) {
        t.vrste[v].korisnici.forEach(function (k) { k.telefon = ''; });
      });
    });
  }

  return {
    zbir: zbir,
    tacke: karta.tacke,
    bezKoordinata: karta.bezKoordinata,
    osvjezeno: Utilities.formatDate(new Date(), 'Europe/Sarajevo',
      'dd.MM.yyyy. HH:mm')
  };
}

/* ------------------------------------------------- upis isporuke iz tabele */

function upisiIsporuku() {
  var ui = SpreadsheetApp.getUi();
  var list = SpreadsheetApp.getActiveSheet();
  var naziv = list.getName();
  if (naziv !== 'PODACI_CIJEPANO' && naziv !== 'PODACI_U_DUGOM') {
    ui.alert('Označite red na listu PODACI_CIJEPANO ili PODACI_U_DUGOM.');
    return;
  }
  var red = list.getActiveRange().getRow();
  if (red < 2) { ui.alert('Označite red s podacima o penzioneru.'); return; }

  var t = citaj_(naziv);
  var podaci = list.getRange(red, 1, 1, t.zaglavlje.length).getValues()[0];
  var ime = podaci[t.i['Prezime']] + ' ' + podaci[t.i['Ime']];
  var odobreno = broj_(podaci[t.i['Odobreno m3']]);
  var vecIsporuceno = broj_(podaci[t.i['Isporučeno m3']]);
  var preostalo = Math.max(okrugli_(odobreno - vecIsporuceno), 0);

  var odgovor = ui.prompt('Isporuka: ' + ime,
    'Koliko m3 je sada isporučeno? (preostalo ' + preostalo + ' m3)',
    ui.ButtonSet.OK_CANCEL);
  if (odgovor.getSelectedButton() !== ui.Button.OK) return;
  var kolicina = broj_(odgovor.getResponseText());
  if (!kolicina) { ui.alert('Nije upisana količina.'); return; }

  var otpremnica = ui.prompt('Broj otpremnice', 'Upišite broj otpremnice (može ostati prazno).',
    ui.ButtonSet.OK_CANCEL);
  if (otpremnica.getSelectedButton() !== ui.Button.OK) return;

  var novoIsporuceno = okrugli_(vecIsporuceno + kolicina);
  var novoPreostalo = Math.max(okrugli_(odobreno - novoIsporuceno), 0);
  list.getRange(red, t.i['Isporučeno m3'] + 1).setValue(novoIsporuceno);
  list.getRange(red, t.i['Preostalo m3'] + 1).setValue(novoPreostalo);
  list.getRange(red, t.i['Status'] + 1).setValue(
    novoPreostalo <= 0.001 ? 'ISPORUČENO' : 'DJELIMIČNO');
  list.getRange(red, t.i['Datum isporuke'] + 1)
    .setValue(Utilities.formatDate(new Date(), 'Europe/Sarajevo', 'dd.MM.yyyy'));
  if (otpremnica.getResponseText()) {
    list.getRange(red, t.i['Otpremnica'] + 1).setValue(otpremnica.getResponseText());
  }
  osvjeziSazetke();
}

/* ----------------------------------------------------------- kategorije */

var KATEGORIJE = ['PENZIONER', 'RVI', 'SINDIKAT', 'OSTALO'];

/** Padajući izbor u koloni "Kategorija" na oba lista s podacima. */
function postaviKategorije() {
  var pravilo = SpreadsheetApp.newDataValidation()
    .requireValueInList(KATEGORIJE, true).setAllowInvalid(false).build();
  var dodano = 0;
  VRSTE.forEach(function (vrsta) {
    var list = list_(vrsta.podaci);
    var t = citaj_(vrsta.podaci);
    var kolona = t.i['Kategorija'];
    if (kolona === undefined) {                 // stara tabela bez kolone
      kolona = t.zaglavlje.length;
      list.insertColumnAfter(list.getLastColumn());
      list.getRange(1, kolona + 1).setValue('Kategorija').setFontWeight('bold')
        .setBackground('#e8eaed');
    }
    var brojRedova = Math.max(list.getLastRow() - 1, 1);
    var opseg = list.getRange(2, kolona + 1, brojRedova, 1);
    opseg.setDataValidation(pravilo);
    opseg.getValues().forEach(function (r, i) {   // prazna polja -> PENZIONER
      if (String(r[0]).trim() === '') {
        list.getRange(i + 2, kolona + 1).setValue('PENZIONER');
        dodano++;
      }
    });
  });
  SpreadsheetApp.getUi().alert('Kategorije su spremne (' + KATEGORIJE.join(', ') +
    ').\nPopunjeno praznih polja: ' + dodano +
    '\n\nPromijeni kategoriju gdje treba, pa pokreni "Osvježi sažetke".');
}

/* -------------------------------------------- kontrola preostalih duplikata */

function provjeriDuplikate() {
  var t = citaj_(LIST_GEO);
  var poMjestu = {};
  t.redovi.forEach(function (r) {
    var m = String(r[t.i['Mjesto']]).trim();
    (poMjestu[m] = poMjestu[m] || []).push(String(r[t.i['Ulica']]).trim());
  });
  var sumnjivi = [];
  Object.keys(poMjestu).forEach(function (m) {
    var u = poMjestu[m];
    for (var i = 0; i < u.length; i++) {
      for (var j = i + 1; j < u.length; j++) {
        if (slicnost_(pojednostavi_(u[i]), pojednostavi_(u[j])) >= 0.8) {
          sumnjivi.push(m + ': "' + u[i] + '"  ~  "' + u[j] + '"');
        }
      }
    }
  });
  SpreadsheetApp.getUi().alert(sumnjivi.length
    ? 'Moguće isti nazivi ulica:\n\n' + sumnjivi.join('\n')
      + '\n\nIspravite naziv u listovima PODACI_* pa pokrenite "Osvježi sažetke".'
    : 'Nema sumnjivih duplikata.');
}

function pojednostavi_(s) {
  return String(s).toUpperCase()
    .replace(/[ČĆ]/g, 'C').replace(/Ž/g, 'Z').replace(/Š/g, 'S').replace(/Đ/g, 'D')
    .replace(/[^A-Z0-9]+/g, ' ').trim();
}

/** Jednostavna mjera sličnosti (dijeljeni bigrami). */
function slicnost_(a, b) {
  if (a === b) return 1;
  if (a.length < 2 || b.length < 2) return 0;
  var par = function (s) {
    var m = {};
    for (var i = 0; i < s.length - 1; i++) {
      var p = s.substr(i, 2);
      m[p] = (m[p] || 0) + 1;
    }
    return m;
  };
  var A = par(a), B = par(b), zajedno = 0, ukupnoA = 0, ukupnoB = 0;
  Object.keys(A).forEach(function (p) {
    ukupnoA += A[p];
    if (B[p]) zajedno += Math.min(A[p], B[p]);
  });
  Object.keys(B).forEach(function (p) { ukupnoB += B[p]; });
  return (2 * zajedno) / (ukupnoA + ukupnoB);
}
