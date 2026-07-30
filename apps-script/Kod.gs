/**
 * Udruženje penzionera Bosanska Krupa - praćenje isporuke ogrjeva.
 *
 * Svaka grupa korisnika se vodi ODVOJENO - svoj odobreni fond kubika,
 * svoj list PODACI_* i ULICE_* (vidi VRSTE ispod). Zajednički su samo
 * MJESTA (zbirni pregled po grupi) i ULICE_GEO (koordinate za kartu).
 */

var VRSTE = [
  { naziv: 'CIJEPANO', podaci: 'PODACI_CIJEPANO', ulice: 'ULICE_CIJEPANO' },
  { naziv: 'U DUGOM', podaci: 'PODACI_U_DUGOM', ulice: 'ULICE_U_DUGOM' },
  { naziv: 'RVI', podaci: 'PODACI_RVI', ulice: 'ULICE_RVI' },
  { naziv: 'PORODICE ŠEHIDA', podaci: 'PODACI_PORODICE_SEHIDA', ulice: 'ULICE_PORODICE_SEHIDA' },
  { naziv: 'SINDIKAT', podaci: 'PODACI_SINDIKAT', ulice: 'ULICE_SINDIKAT' }
];
var LIST_MJESTA = 'MJESTA';
var LIST_GEO = 'ULICE_GEO';
var LIST_ISPORUKE = 'ISPORUKE';
var LIST_PREGLED = 'PREGLED';
var LIST_REKAP = 'REKAP';

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🪵 Drva')
    .addItem('1. Osvježi sažetke (ulice i mjesta)', 'osvjeziSazetke')
    .addItem('2. Geokodiraj ulice (za kartu)', 'geokodirajUlice')
    .addItem('3. Otvori kartu', 'otvoriKartu')
    .addSeparator()
    .addItem('Upiši isporuku za označeni red', 'upisiIsporuku')
    .addItem('Otvori zbirni pregled', 'otvoriPregled')
    .addItem('Otvori rekap (grupe + mjeseci)', 'otvoriRekap')
    .addSeparator()
    .addSubMenu(ui.createMenu('Provjere')
      .addItem('Mogući duplikati ulica', 'provjeriDuplikate')
      .addItem('Dupli matični brojevi u istom spisku', 'provjeriDupleKorisnike'))
    .addSubMenu(ui.createMenu('Podešavanja')
      .addItem('Izbor kategorija (penzioner/RVI/sindikat)', 'postaviKategorije')
      .addItem('Ispravi "Redni broj" u pravi broj (tekst → broj)', 'ispraviRedniBroj')
      .addItem('Zaključaj računate kolone', 'zakljucajRacunateKolone')
      .addItem('Otključaj računate kolone', 'otkljucajRacunateKolone')
      .addItem('Uključi automatsko jutarnje osvježavanje', 'ukljuciDnevnoOsvjezavanje')
      .addItem('Isključi automatsko osvježavanje', 'iskljuciDnevnoOsvjezavanje'))
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

/**
 * Automatski se pokreće kad neko upiše/promijeni broj u koloni
 * "Isporučeno m3" na bilo kojem listu PODACI_* - operater ne mora otvarati
 * dijalog "Upiši isporuku": čim upiše kubike, sam se izračuna Preostalo,
 * Status i boja reda. Datum isporuke se NE upisuje automatski - to je
 * namjerno ostavljeno za ručni unos (piše se kratko, npr. "28.07." bez
 * godine i vremena).
 * (Ne mijenja ULICE_GEO/MJESTA/PREGLED - za to i dalje treba
 * "Osvježi sažetke", ili uključi automatsko jutarnje osvježavanje.)
 */
function onEdit(e) {
  try {
    var opseg = e.range;
    var list = opseg.getSheet();
    var jePodaci = false;
    VRSTE.forEach(function (v) { if (v.podaci === list.getName()) jePodaci = true; });
    if (!jePodaci || opseg.getRow() < 2) return;

    var zadnjaKolona = list.getLastColumn();
    var zaglavlje = list.getRange(1, 1, 1, zadnjaKolona).getValues()[0];
    var idx = {};
    zaglavlje.forEach(function (h, i) { idx[String(h).trim()] = i; });
    var kolIsporuceno = idx['Isporučeno m3'];
    if (kolIsporuceno === undefined) return;
    // reaguje samo kad je izmijenjena baš kolona Isporučeno m3 (uključivo pri
    // ljepljenju više redova); izmjene u Preostalo/Status koje ovaj kod sam
    // upiše ne pokreću ponovo ovu granu jer su u drugoj koloni
    if (opseg.getColumn() > kolIsporuceno + 1 ||
        opseg.getColumn() + opseg.getNumColumns() - 1 < kolIsporuceno + 1) return;

    var kolOdobreno = idx['Odobreno m3'], kolPreostalo = idx['Preostalo m3'],
      kolStatus = idx['Status'];
    if ([kolOdobreno, kolPreostalo, kolStatus].some(function (k) { return k === undefined; })) return;

    for (var r = 0; r < opseg.getNumRows(); r++) {
      var red = opseg.getRow() + r;
      var isporuceno = broj_(list.getRange(red, kolIsporuceno + 1).getValue());
      var odobreno = broj_(list.getRange(red, kolOdobreno + 1).getValue());
      var preostalo = Math.max(Math.round((odobreno - isporuceno) * 100) / 100, 0);
      var status = (preostalo <= 0.001 && isporuceno > 0) ? 'ISPORUČENO'
        : (isporuceno > 0.001 ? 'DJELIMIČNO' : 'ZA ISPORUKU');

      list.getRange(red, kolPreostalo + 1).setValue(preostalo);
      list.getRange(red, kolStatus + 1).setValue(status);
      var boja = (preostalo <= 0.001 && isporuceno > 0) ? BOJA_ISPORUCENO : BOJA_NEISPORUCENO;
      list.getRange(red, 1, 1, zadnjaKolona).setBackground(boja);
    }
  } catch (err) {
    // ne prekidati uređivanje korisniku zbog greške ovdje
  }
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
  if (SpreadsheetApp.getActive().getSheetByName(LIST_PREGLED)) osvjeziPregled_();
  if (SpreadsheetApp.getActive().getSheetByName(LIST_REKAP)) osvjeziRekap_();
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
  var zaglavlje = ['Mjesto', 'Ulica', 'Adresa ključ'];
  VRSTE.forEach(function (vrsta) {
    zaglavlje.push('Preostalo ' + vrsta.naziv + ' m3', 'Isporučeno ' + vrsta.naziv + ' m3');
  });
  zaglavlje.push('Preostalo ukupno m3', 'Adresa za kartu', 'Lat', 'Lng', 'Tačnost',
    'Podudaranje sa');
  var kolonaTacnosti = zaglavlje.indexOf('Tačnost') + 1;

  var redovi = Object.keys(sveUlice).map(function (k) {
    var u = sveUlice[k];
    var xy = koordinate[k] || { lat: '', lng: '', tacnost: '', podudaranje: '' };
    var adresa = u.ulica && u.ulica !== '(bez ulice)'
      ? u.ulica + ', ' + u.mjesto + ', Bosna i Hercegovina'
      : u.mjesto + ', Bosna i Hercegovina';
    var red = [u.mjesto, u.ulica, k];
    var ukupnoPreostalo = 0;
    VRSTE.forEach(function (vrsta) {
      var z = poVrsti[vrsta.naziv][k] || { preostalo: 0, isporuceno: 0, korisnika: 0 };
      red.push(okrugli_(z.preostalo), okrugli_(z.isporuceno));
      ukupnoPreostalo += z.preostalo;
    });
    red.push(okrugli_(ukupnoPreostalo), adresa, xy.lat, xy.lng, xy.tacnost, xy.podudaranje);
    return red;
  }).sort(function (a, b) {
    return b[zaglavlje.indexOf('Preostalo ukupno m3')] - a[zaglavlje.indexOf('Preostalo ukupno m3')];
  });

  var s = upisi_(LIST_GEO, zaglavlje, redovi);
  if (redovi.length) {
    var opseg = s.getRange(2, kolonaTacnosti, redovi.length, 1);
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


/* ------------------------------------------------------------------ 3. karta */

/** Nazivi svih grupa (VRSTE) - Karta.html gradi izbor dinamički iz ovoga. */
function spisakVrsta() {
  return VRSTE.map(function (v) { return v.naziv; });
}

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
    vrste: VRSTE.map(function (v) { return v.naziv; }),
    zbir: zbir,
    tacke: karta.tacke,
    bezKoordinata: karta.bezKoordinata,
    osvjezeno: Utilities.formatDate(new Date(), 'Europe/Sarajevo',
      'dd.MM.yyyy. HH:mm')
  };
}

/* ------------------------------------------------- upis isporuke iz tabele */

function nazivListaZaVrstu_(nazivLista) {
  var vrsta = null;
  VRSTE.forEach(function (v) { if (v.podaci === nazivLista) vrsta = v.naziv; });
  return vrsta;
}

function upisiIsporuku() {
  var ui = SpreadsheetApp.getUi();
  var list = SpreadsheetApp.getActiveSheet();
  var naziv = list.getName();
  var vrstaNaziv = nazivListaZaVrstu_(naziv);
  if (!vrstaNaziv) {
    ui.alert('Označite red na jednom od listova: ' +
      VRSTE.map(function (v) { return v.podaci; }).join(', ') + '.');
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

  var datumUnos = ui.prompt('Datum isporuke',
    'Upišite datum (kratko, npr. "28.07." - bez godine i vremena). Može ostati prazno.',
    ui.ButtonSet.OK_CANCEL);
  if (datumUnos.getSelectedButton() !== ui.Button.OK) return;

  var novoIsporuceno = okrugli_(vecIsporuceno + kolicina);
  var novoPreostalo = Math.max(okrugli_(odobreno - novoIsporuceno), 0);
  list.getRange(red, t.i['Isporučeno m3'] + 1).setValue(novoIsporuceno);
  list.getRange(red, t.i['Preostalo m3'] + 1).setValue(novoPreostalo);
  list.getRange(red, t.i['Status'] + 1).setValue(
    novoPreostalo <= 0.001 ? 'ISPORUČENO' : 'DJELIMIČNO');
  if (datumUnos.getResponseText().trim()) {
    list.getRange(red, t.i['Datum isporuke'] + 1).setValue(datumUnos.getResponseText().trim());
  }
  if (otpremnica.getResponseText()) {
    list.getRange(red, t.i['Otpremnica'] + 1).setValue(otpremnica.getResponseText());
  }

  zapisiIsporuku_({
    vrsta: vrstaNaziv,
    maticni: podaci[t.i['Matični broj']],
    prezime: podaci[t.i['Prezime']],
    ime: podaci[t.i['Ime']],
    mjesto: podaci[t.i['Mjesto']],
    ulica: podaci[t.i['Ulica']],
    kolicina: kolicina,
    otpremnica: otpremnica.getResponseText(),
    ukupnoNakon: novoIsporuceno,
    preostaloNakon: novoPreostalo
  });
  osvjeziSazetke();
}

/* --------------------------------------------------- historija isporuka */

var ZAGLAVLJE_ISPORUKE = ['Datum', 'Vrijeme', 'Vrsta', 'Matični broj', 'Prezime',
  'Ime', 'Mjesto', 'Ulica', 'Isporučeno sada m3', 'Otpremnica',
  'Ukupno isporučeno m3', 'Preostalo m3', 'Upisao'];

function listIsporuka_() {
  var ss = SpreadsheetApp.getActive();
  var s = ss.getSheetByName(LIST_ISPORUKE);
  if (!s) {
    s = ss.insertSheet(LIST_ISPORUKE);
    s.getRange(1, 1, 1, ZAGLAVLJE_ISPORUKE.length).setValues([ZAGLAVLJE_ISPORUKE])
      .setFontWeight('bold').setBackground('#e8eaed');
    s.setFrozenRows(1);
  }
  return s;
}

/**
 * Svaka isporuka se dodaje kao novi red - tako se vidi i historija
 * djelimičnih isporuka (npr. 5 m3 danas, 5 m3 za mjesec), što se u
 * kolonama "Datum isporuke"/"Otpremnica" gubi jer čuvaju samo zadnju.
 */
function zapisiIsporuku_(p) {
  var s = listIsporuka_();
  var sada = new Date();
  var korisnik = '';
  try { korisnik = Session.getActiveUser().getEmail() || ''; } catch (e) { /* nema pristupa */ }
  s.appendRow([
    Utilities.formatDate(sada, 'Europe/Sarajevo', 'dd.MM.yyyy.'),
    Utilities.formatDate(sada, 'Europe/Sarajevo', 'HH:mm'),
    p.vrsta, p.maticni, p.prezime, p.ime, p.mjesto, p.ulica,
    p.kolicina, p.otpremnica || '', p.ukupnoNakon, p.preostaloNakon, korisnik
  ]);
}

/* ------------------------------------------- zaključavanje računatih kolona */

// Ove kolone računa skripta; ručna izmjena bi razišla zbirove s podacima.
var RACUNATE_KOLONE = ['Preostalo m3', 'Status'];
var OPIS_ZASTITE = 'Računa skripta - ne mijenjati ručno';

function zakljucajRacunateKolone() {
  var zakljucano = [];
  VRSTE.forEach(function (vrsta) {
    var list = list_(vrsta.podaci);
    var t = citaj_(vrsta.podaci);
    var brojRedova = Math.max(list.getLastRow() - 1, 1);
    RACUNATE_KOLONE.forEach(function (naziv) {
      if (t.i[naziv] === undefined) return;
      var opseg = list.getRange(2, t.i[naziv] + 1, brojRedova, 1);
      var zastita = opseg.protect().setDescription(
        OPIS_ZASTITE + ' (' + vrsta.podaci + ' / ' + naziv + ')');
      // vlasnik ostaje urednik da skripta i dalje može pisati
      zastita.removeEditors(zastita.getEditors().map(function (u) { return u.getEmail(); }));
      if (zastita.canDomainEdit && zastita.canDomainEdit()) zastita.setDomainEdit(false);
      zakljucano.push(vrsta.podaci + ' → ' + naziv);
    });
  });
  SpreadsheetApp.getUi().alert('Zaključano (samo vlasnik može mijenjati):\n\n' +
    zakljucano.join('\n') +
    '\n\nOve kolone se same računaju pri "Osvježi sažetke" i upisu isporuke.');
}

function otkljucajRacunateKolone() {
  var skinuto = 0;
  SpreadsheetApp.getActive().getSheets().forEach(function (list) {
    list.getProtections(SpreadsheetApp.ProtectionType.RANGE).forEach(function (z) {
      if (String(z.getDescription()).indexOf(OPIS_ZASTITE) === 0) {
        z.remove();
        skinuto++;
      }
    });
  });
  SpreadsheetApp.getUi().alert('Skinuto zaštićenih opsega: ' + skinuto + '.');
}

/* ------------------------------------------------ provjera duplih korisnika */

/** Isti matični broj dvaput u ISTOM spisku je greška unosa. */
function provjeriDupleKorisnike() {
  var nalazi = [];
  VRSTE.forEach(function (vrsta) {
    var t = citaj_(vrsta.podaci);
    var vidjeni = {};
    t.redovi.forEach(function (r, idx) {
      var mb = String(r[t.i['Matični broj']] || '').trim();
      if (!mb) return;
      var opis = String(r[t.i['Prezime']]) + ' ' + String(r[t.i['Ime']]) +
        ' (red ' + (idx + 2) + ')';
      if (vidjeni[mb]) {
        nalazi.push(vrsta.podaci + ' – MB ' + mb + ': ' + vidjeni[mb] + '  ↔  ' + opis);
      } else {
        vidjeni[mb] = opis;
      }
    });
  });
  SpreadsheetApp.getUi().alert(nalazi.length
    ? 'Isti matični broj dvaput u istom spisku (' + nalazi.length + '):\n\n' +
      nalazi.slice(0, 30).join('\n') +
      (nalazi.length > 30 ? '\n…' : '') +
      '\n\nProvjerite je li riječ o grešci unosa ili o dvije stavke istog korisnika.'
    : 'Nema duplih matičnih brojeva unutar istog spiska.\n\n' +
      '(Isti korisnik u oba spiska - cijepano i u dugom - je uredu i ne prijavljuje se.)');
}

/* --------------------------------------------------------- zbirni pregled */

function otvoriPregled() {
  osvjeziPregled_();
  SpreadsheetApp.getActive().setActiveSheet(list_(LIST_PREGLED));
}

/** Jedan list s krupnim brojkama i grafikonom - za upravu. */
function osvjeziPregled_() {
  var ss = SpreadsheetApp.getActive();
  var s = ss.getSheetByName(LIST_PREGLED);
  if (!s) {
    s = ss.insertSheet(LIST_PREGLED, 0);
  } else {
    s.getCharts().forEach(function (g) { s.removeChart(g); });
    s.clear();
  }

  var zbir = {}, ukupno = { korisnika: 0, odobreno: 0, isporuceno: 0, preostalo: 0, gotovih: 0 };
  VRSTE.forEach(function (vrsta) {
    var t = citaj_(vrsta.podaci);
    var z = { korisnika: 0, odobreno: 0, isporuceno: 0, preostalo: 0, gotovih: 0 };
    t.redovi.forEach(function (r) {
      var odobreno = broj_(r[t.i['Odobreno m3']]);
      var isporuceno = broj_(r[t.i['Isporučeno m3']]);
      if (!odobreno && !isporuceno) return;
      var preostalo = Math.max(okrugli_(odobreno - isporuceno), 0);
      z.korisnika++; z.odobreno += odobreno; z.isporuceno += isporuceno; z.preostalo += preostalo;
      if (preostalo <= 0.001 && isporuceno > 0) z.gotovih++;
    });
    zbir[vrsta.naziv] = z;
    ['korisnika', 'odobreno', 'isporuceno', 'preostalo', 'gotovih'].forEach(function (k) {
      ukupno[k] += z[k];
    });
  });

  var ulicaNijePoceto = 0, ulicaUkupno = 0;
  VRSTE.forEach(function (vrsta) {
    var u = SpreadsheetApp.getActive().getSheetByName(vrsta.ulice);
    if (!u) return;
    var t = citaj_(vrsta.ulice);
    t.redovi.forEach(function (r) {
      ulicaUkupno++;
      if (String(r[t.i['Status ulice']]) === 'NIJE POČETO') ulicaNijePoceto++;
    });
  });

  var postotak = ukupno.odobreno ? ukupno.isporuceno / ukupno.odobreno : 0;
  s.getRange('A1').setValue('PREGLED ISPORUKE OGRJEVA')
    .setFontSize(16).setFontWeight('bold');
  s.getRange('A2').setValue('Osvježeno: ' +
    Utilities.formatDate(new Date(), 'Europe/Sarajevo', 'dd.MM.yyyy. HH:mm'))
    .setFontColor('#5f6368');

  var kartice = [
    ['PREOSTALO UKUPNO', okrugli_(ukupno.preostalo) + ' m³'],
    ['ISPORUČENO', okrugli_(ukupno.isporuceno) + ' m³  (' + Math.round(postotak * 100) + '%)'],
    ['ODOBRENO', okrugli_(ukupno.odobreno) + ' m³'],
    ['KORISNIKA ČEKA', (ukupno.korisnika - ukupno.gotovih) + ' od ' + ukupno.korisnika],
    ['ULICA NIJE POČETO', ulicaNijePoceto + ' od ' + ulicaUkupno]
  ];
  kartice.forEach(function (k, i) {
    var red = 4 + i;
    s.getRange(red, 1).setValue(k[0]).setFontColor('#5f6368').setFontSize(10);
    s.getRange(red, 2).setValue(k[1]).setFontWeight('bold').setFontSize(14);
  });

  var pocetak = 4 + kartice.length + 1;
  s.getRange(pocetak, 1, 1, 5).setValues([['Vrsta', 'Korisnika', 'Odobreno m³',
    'Isporučeno m³', 'Preostalo m³']]).setFontWeight('bold').setBackground('#e8eaed');
  VRSTE.forEach(function (vrsta, i) {
    var z = zbir[vrsta.naziv];
    s.getRange(pocetak + 1 + i, 1, 1, 5).setValues([[vrsta.naziv, z.korisnika,
      okrugli_(z.odobreno), okrugli_(z.isporuceno), okrugli_(z.preostalo)]]);
  });
  var redUkupno = pocetak + 1 + VRSTE.length;
  s.getRange(redUkupno, 1, 1, 5).setValues([['UKUPNO', ukupno.korisnika,
    okrugli_(ukupno.odobreno), okrugli_(ukupno.isporuceno), okrugli_(ukupno.preostalo)]])
    .setFontWeight('bold');

  // zbir po mjestu (preko svih grupa) - MJESTA ima po jedan red za svaku
  // grupu, pa se prvo sabere po mjestu u pomoćnu tabelu za grafikon
  if (SpreadsheetApp.getActive().getSheetByName(LIST_MJESTA)) {
    var tm = citaj_(LIST_MJESTA);
    var poMjestu = {}, redoslijedMjesta = [];
    tm.redovi.forEach(function (r) {
      var m = String(r[tm.i['Mjesto']]);
      if (!poMjestu[m]) { poMjestu[m] = { isporuceno: 0, preostalo: 0 }; redoslijedMjesta.push(m); }
      poMjestu[m].isporuceno += broj_(r[tm.i['Isporučeno m3']]);
      poMjestu[m].preostalo += broj_(r[tm.i['Preostalo m3']]);
    });
    var pomocnaKolona = 11;  // K - van vidokruga tabele iznad
    s.getRange(1, pomocnaKolona, 1, 3).setValues([['Mjesto', 'Isporučeno m³', 'Preostalo m³']]);
    redoslijedMjesta.sort(function (a, b) { return poMjestu[b].preostalo - poMjestu[a].preostalo; });
    redoslijedMjesta.forEach(function (m, i) {
      s.getRange(2 + i, pomocnaKolona, 1, 3).setValues(
        [[m, okrugli_(poMjestu[m].isporuceno), okrugli_(poMjestu[m].preostalo)]]);
    });
    if (redoslijedMjesta.length) {
      var grafikon = s.newChart().setChartType(Charts.ChartType.COLUMN)
        .addRange(s.getRange(1, pomocnaKolona, redoslijedMjesta.length + 1, 3))
        .setPosition(4, 7, 0, 0)
        .setOption('title', 'Isporučeno i preostalo po mjestu (m³, sve grupe zajedno)')
        .setOption('width', 620).setOption('height', 340)
        .setOption('colors', ['#188038', '#d93025'])
        .setOption('isStacked', true)
        .build();
      s.insertChart(grafikon);
    }
  }
  s.setColumnWidth(1, 190);
  s.setColumnWidth(2, 190);
  return s;
}

/* ------------------------------------------------------------ rekap */

/**
 * Pretvara vrijednost kolone "Datum isporuke" (Date objekat ili tekst) u
 * ključ mjeseca "yyyy-MM". Datum se upisuje ručno i obično bez godine
 * (npr. "28.07."), pa se u tom slučaju uzima tekuća godina - ako se to
 * pokaže netačno oko prijelaza godine, upiši godinu eksplicitno
 * ("28.07.2026") i biće tačno prepoznato. Vraća null ako datum uopšte
 * nije prepoznat - takav red se ne broji u mjesečnom pregledu, ali
 * ostaje uračunat u ukupnim zbirovima.
 */
function mjesecKljuc_(v) {
  if (!v) return null;
  var d = null;
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v)) {
    d = v;
  } else {
    var s = String(v).trim();
    var sGod = s.match(/(\d{1,2})[.\/](\d{1,2})[.\/](\d{4})/);
    if (sGod) {
      d = new Date(parseInt(sGod[3], 10), parseInt(sGod[2], 10) - 1, parseInt(sGod[1], 10));
    } else {
      var bezGod = s.match(/^(\d{1,2})[.\/](\d{1,2})\.?$/);
      if (bezGod) {
        d = new Date(new Date().getFullYear(), parseInt(bezGod[2], 10) - 1,
          parseInt(bezGod[1], 10));
      }
    }
  }
  return d ? Utilities.formatDate(d, 'Europe/Sarajevo', 'yyyy-MM') : null;
}

var NAZIVI_MJESECI_ = ['januar', 'februar', 'mart', 'april', 'maj', 'juni',
  'juli', 'avgust', 'septembar', 'oktobar', 'novembar', 'decembar'];

function opisMjeseca_(kljuc) {
  var dijelovi = kljuc.split('-');
  return NAZIVI_MJESECI_[parseInt(dijelovi[1], 10) - 1] + ' ' + dijelovi[0];
}

function otvoriRekap() {
  osvjeziRekap_();
  SpreadsheetApp.getActive().setActiveSheet(list_(LIST_REKAP));
}

// Svako udruženje (osnovna grupa) ima ugovor za OBA tipa drva - cijepano i
// u dugom. Listovi CIJEPANO/U DUGOM već jesu ta podjela za penzionere; za
// RVI/PORODICE ŠEHIDA/SINDIKAT tip drva po osobi prepoznaje se iz kolone
// "Napomena" (npr. oznaka "DUG" u izvornom spisku Porodica šehida). Gdje
// izvor ne razlikuje tip (RVI, Sindikat - trenutno nema te oznake u
// dobijenim spiskovima), sve ide pod CIJEPANO dok se ne dobije precizniji
// podatak - lako se ispravi upisom "DUG" u kolonu Napomena za te redove.
var GRUPA_ZA_VRSTU_ = {
  'CIJEPANO': 'PENZIONERI', 'U DUGOM': 'PENZIONERI',
  'RVI': 'RVI', 'PORODICE ŠEHIDA': 'PORODICE ŠEHIDA', 'SINDIKAT': 'SINDIKAT'
};
var REDOSLIJED_GRUPA_ = ['PENZIONERI', 'RVI', 'PORODICE ŠEHIDA', 'SINDIKAT'];
var TIPOVI_DRVA_ = ['CIJEPANO', 'U DUGOM'];

function tipDrvaZaRed_(vrsta, napomena) {
  if (vrsta.naziv === 'CIJEPANO' || vrsta.naziv === 'U DUGOM') return vrsta.naziv;
  return String(napomena || '').toUpperCase().indexOf('DUG') >= 0 ? 'U DUGOM' : 'CIJEPANO';
}

function prazanZbir_() {
  return { korisnika: 0, odobreno: 0, isporuceno: 0, preostalo: 0, gotovih: 0 };
}

function dodajUZbir_(z, odobreno, isporuceno, preostalo) {
  z.korisnika++; z.odobreno += odobreno; z.isporuceno += isporuceno; z.preostalo += preostalo;
  if (preostalo <= 0.001 && isporuceno > 0) z.gotovih++;
}

/**
 * List REKAP: isporučeno/neisporučeno po GRUPI (Penzioneri, RVI, Porodice
 * šehida, Sindikat) razdvojeno na CIJEPANO i U DUGOM, i otprema po
 * mjesecima za iste kombinacije. Mjesečni podaci dolaze iz kolone
 * "Datum isporuke" u PODACI_* (datum zadnje/jedine isporuke za tog
 * korisnika) - za korisnike sa više djelimičnih isporuka to je mjesec
 * posljednje; tačniju historiju po datumu svake pojedinačne isporuke ima
 * list ISPORUKE.
 */
function osvjeziRekap_() {
  var ss = SpreadsheetApp.getActive();
  var s = ss.getSheetByName(LIST_REKAP);
  if (!s) {
    s = ss.insertSheet(LIST_REKAP, 0);
  } else {
    s.getCharts().forEach(function (g) { s.removeChart(g); });
    s.clear();
  }

  // --- zbir po (grupa, tip drva) ---
  var poGrupiTip = {};  // { 'PENZIONERI': { 'CIJEPANO': {...}, 'U DUGOM': {...} }, ... }
  REDOSLIJED_GRUPA_.forEach(function (g) {
    poGrupiTip[g] = {};
    TIPOVI_DRVA_.forEach(function (t) { poGrupiTip[g][t] = prazanZbir_(); });
  });
  var mjeseci = {};  // { 'yyyy-MM': { 'GRUPA|TIP': isporuceno_m3 } }

  VRSTE.forEach(function (vrsta) {
    var t = citaj_(vrsta.podaci);
    var grupa = GRUPA_ZA_VRSTU_[vrsta.naziv];
    t.redovi.forEach(function (r) {
      var odobreno = broj_(r[t.i['Odobreno m3']]);
      var isporuceno = broj_(r[t.i['Isporučeno m3']]);
      if (!odobreno && !isporuceno) return;
      var preostalo = Math.max(okrugli_(odobreno - isporuceno), 0);
      var napomena = t.i['Napomena'] !== undefined ? r[t.i['Napomena']] : '';
      var tip = tipDrvaZaRed_(vrsta, napomena);
      dodajUZbir_(poGrupiTip[grupa][tip], odobreno, isporuceno, preostalo);

      if (isporuceno > 0.001 && t.i['Datum isporuke'] !== undefined) {
        var kljuc = mjesecKljuc_(r[t.i['Datum isporuke']]);
        if (kljuc) {
          var kljucGT = grupa + '|' + tip;
          if (!mjeseci[kljuc]) mjeseci[kljuc] = {};
          mjeseci[kljuc][kljucGT] = okrugli_((mjeseci[kljuc][kljucGT] || 0) + isporuceno);
        }
      }
    });
  });

  s.getRange('A1').setValue('REKAP – ISPORUKA OGRJEVA PO GRUPAMA I TIPU DRVA')
    .setFontSize(16).setFontWeight('bold');
  s.getRange('A2').setValue('Osvježeno: ' +
    Utilities.formatDate(new Date(), 'Europe/Sarajevo', 'dd.MM.yyyy. HH:mm'))
    .setFontColor('#5f6368');

  // --- tabela 1: isporučeno / neisporučeno po grupi i tipu drva ---
  s.getRange('A4').setValue('Isporučena i neisporučena drvna masa po grupama (cijepano / u dugom)')
    .setFontWeight('bold').setFontSize(13);
  var zaglavlje1 = ['Grupa', 'Tip drva', 'Korisnika', 'Odobreno m³', 'Isporučeno m³',
    'Neisporučeno m³', 'Realizacija'];
  s.getRange(5, 1, 1, zaglavlje1.length).setValues([zaglavlje1])
    .setFontWeight('bold').setBackground('#e8eaed');

  var red = 6, ukupnoSve = prazanZbir_();
  REDOSLIJED_GRUPA_.forEach(function (grupa) {
    var ukupnoGrupa = prazanZbir_();
    TIPOVI_DRVA_.forEach(function (tip) {
      var z = poGrupiTip[grupa][tip];
      var postotak = z.odobreno ? Math.round(z.isporuceno / z.odobreno * 100) : 0;
      s.getRange(red, 1, 1, zaglavlje1.length).setValues([[nazivGrupe_(grupa),
        nazivTipa_(tip), z.korisnika, okrugli_(z.odobreno), okrugli_(z.isporuceno),
        okrugli_(z.preostalo), postotak + '%']]);
      red++;
      ['korisnika', 'odobreno', 'isporuceno', 'preostalo'].forEach(function (k) {
        ukupnoGrupa[k] += z[k];
      });
    });
    var postotakGrupa = ukupnoGrupa.odobreno
      ? Math.round(ukupnoGrupa.isporuceno / ukupnoGrupa.odobreno * 100) : 0;
    s.getRange(red, 1, 1, zaglavlje1.length).setValues([[nazivGrupe_(grupa) + ' - ukupno', '',
      ukupnoGrupa.korisnika, okrugli_(ukupnoGrupa.odobreno), okrugli_(ukupnoGrupa.isporuceno),
      okrugli_(ukupnoGrupa.preostalo), postotakGrupa + '%']])
      .setFontWeight('bold').setBackground('#f1f3f4');
    red++;
    ['korisnika', 'odobreno', 'isporuceno', 'preostalo'].forEach(function (k) {
      ukupnoSve[k] += ukupnoGrupa[k];
    });
  });
  var postotakSve = ukupnoSve.odobreno ? Math.round(ukupnoSve.isporuceno / ukupnoSve.odobreno * 100) : 0;
  s.getRange(red, 1, 1, zaglavlje1.length).setValues([['SVEUKUPNO', '', ukupnoSve.korisnika,
    okrugli_(ukupnoSve.odobreno), okrugli_(ukupnoSve.isporuceno), okrugli_(ukupnoSve.preostalo),
    postotakSve + '%']]).setFontWeight('bold').setBackground('#dadce0');
  var redZadnji1 = red;

  var pomKolona1 = 10;  // J - pomoćna tabela za graf, van vidokruga glavne tabele
  var pomZaglavlje1 = ['Grupa i tip', 'Isporučeno m³', 'Neisporučeno m³'];
  s.getRange(5, pomKolona1, 1, 3).setValues([pomZaglavlje1]);
  var pomRedovi1 = [];
  REDOSLIJED_GRUPA_.forEach(function (grupa) {
    TIPOVI_DRVA_.forEach(function (tip) {
      var z = poGrupiTip[grupa][tip];
      pomRedovi1.push([nazivGrupe_(grupa) + ' - ' + nazivTipa_(tip),
        okrugli_(z.isporuceno), okrugli_(z.preostalo)]);
    });
  });
  s.getRange(6, pomKolona1, pomRedovi1.length, 3).setValues(pomRedovi1);
  var grafikonIsporuke = s.newChart().setChartType(Charts.ChartType.COLUMN)
    .addRange(s.getRange(5, pomKolona1, pomRedovi1.length + 1, 3))
    .setPosition(6, pomKolona1 + 4, 0, 0)
    .setOption('title', 'Isporučeno i neisporučeno po grupi i tipu drva (m³)')
    .setOption('width', 620).setOption('height', 340)
    .setOption('colors', ['#188038', '#d93025'])
    .setOption('isStacked', true)
    .build();
  s.insertChart(grafikonIsporuke);

  // --- tabela 2: otprema po mjesecima, po grupi i tipu drva ---
  var redNaslov2 = redZadnji1 + 3;
  s.getRange(redNaslov2, 1).setValue('Otprema po mjesecima, po grupi i tipu drva (m³ isporučeno)')
    .setFontWeight('bold').setFontSize(13);
  var redZaglavlje2 = redNaslov2 + 1;
  var koloneGT = [];
  REDOSLIJED_GRUPA_.forEach(function (grupa) {
    TIPOVI_DRVA_.forEach(function (tip) { koloneGT.push(grupa + '|' + tip); });
  });
  var zaglavlje2 = ['Mjesec'].concat(koloneGT.map(function (k) {
    var dio = k.split('|');
    return nazivGrupe_(dio[0]) + ' - ' + nazivTipa_(dio[1]);
  }), ['UKUPNO']);
  s.getRange(redZaglavlje2, 1, 1, zaglavlje2.length).setValues([zaglavlje2])
    .setFontWeight('bold').setBackground('#e8eaed');

  var kljucevi = Object.keys(mjeseci).sort();
  if (kljucevi.length) {
    var redoviMjeseci = kljucevi.map(function (k) {
      var redM = [opisMjeseca_(k)];
      var ukupnoMjesec = 0;
      koloneGT.forEach(function (kgt) {
        var iznos = okrugli_(mjeseci[k][kgt] || 0);
        redM.push(iznos);
        ukupnoMjesec += iznos;
      });
      redM.push(okrugli_(ukupnoMjesec));
      return redM;
    });
    s.getRange(redZaglavlje2 + 1, 1, redoviMjeseci.length, zaglavlje2.length)
      .setValues(redoviMjeseci);

    var grafikon2 = s.newChart().setChartType(Charts.ChartType.COLUMN)
      .addRange(s.getRange(redZaglavlje2, 1, redoviMjeseci.length + 1, koloneGT.length + 1))
      .setPosition(redZaglavlje2, 9, 0, 0)
      .setOption('title', 'Otprema po mjesecima, po grupi i tipu drva (m³)')
      .setOption('width', 760).setOption('height', 360)
      .setOption('isStacked', true)
      .build();
    s.insertChart(grafikon2);
  } else {
    s.getRange(redZaglavlje2 + 1, 1).setValue(
      '(još nema isporuka s prepoznatim datumom - upiši datum uz isporuku)');
  }

  s.setColumnWidth(1, 190);
  for (var c = 2; c <= Math.max(zaglavlje1.length, zaglavlje2.length); c++) s.setColumnWidth(c, 130);
}

function nazivGrupe_(g) {
  return g.charAt(0) + g.slice(1).toLowerCase();
}

function nazivTipa_(t) {
  return t === 'CIJEPANO' ? 'Cijepano' : 'U dugom';
}

/* --------------------------------------------- automatsko osvježavanje */

var FUNKCIJA_OKIDACA = 'dnevnoOsvjezavanje';

/** Pokreće se okidačem svako jutro; bez dijaloga jer nema korisnika. */
function dnevnoOsvjezavanje() {
  osvjeziSazetke();
  osvjeziPregled_();
  osvjeziRekap_();
}

function ukljuciDnevnoOsvjezavanje() {
  iskljuciOkidace_();
  ScriptApp.newTrigger(FUNKCIJA_OKIDACA).timeBased().atHour(6).everyDays(1).create();
  SpreadsheetApp.getUi().alert(
    'Uključeno: svako jutro između 6 i 7 sati tabela sama osvježi sažetke i pregled.');
}

function iskljuciDnevnoOsvjezavanje() {
  var broj = iskljuciOkidace_();
  SpreadsheetApp.getUi().alert(broj
    ? 'Automatsko osvježavanje je isključeno.'
    : 'Automatsko osvježavanje nije ni bilo uključeno.');
}

function iskljuciOkidace_() {
  var broj = 0;
  ScriptApp.getProjectTriggers().forEach(function (o) {
    if (o.getHandlerFunction() === FUNKCIJA_OKIDACA) { ScriptApp.deleteTrigger(o); broj++; }
  });
  return broj;
}

/**
 * Ako je "Redni broj" upisan kao tekst (čest slučaj poslije uvoza iz
 * .xls/.csv), Google Sheets ga sortira i filtrira slovno (1, 10, 11, 2...)
 * umjesto brojčano. Ovo pretvara postojeće vrijednosti u pravi broj, u
 * svim listovima PODACI_*, bez diranja bilo koje druge kolone ili podatka.
 */
function ispraviRedniBroj() {
  var ispravljeno = 0, preskoceno = 0;
  VRSTE.forEach(function (vrsta) {
    var list = list_(vrsta.podaci);
    var t = citaj_(vrsta.podaci);
    var kolona = t.i['Redni broj'];
    if (kolona === undefined || !t.redovi.length) return;
    var opseg = list.getRange(2, kolona + 1, t.redovi.length, 1);
    var nove = opseg.getValues().map(function (r) {
      var v = r[0];
      if (typeof v === 'number') { preskoceno++; return [v]; }
      var broj = parseInt(String(v).trim().replace(/\.$/, ''), 10);
      if (isNaN(broj)) { preskoceno++; return [v]; }
      ispravljeno++;
      return [broj];
    });
    opseg.setValues(nove);
    opseg.setNumberFormat('0');       // spriječi da Sheets nazad prikaže kao tekst
  });
  SpreadsheetApp.getUi().alert('Ispravljeno u broj: ' + ispravljeno +
    '.\nVeć bilo u redu ili nije prepoznato: ' + preskoceno + '.');
}

/* ----------------------------------------------------------- kategorije */

var KATEGORIJE = ['PENZIONER', 'RVI', 'PORODICA ŠEHIDA', 'SINDIKAT', 'OSTALO'];

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
