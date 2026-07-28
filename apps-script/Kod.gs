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

function obojiStatus_(list, kolona, brojRedova) {
  if (!brojRedova) return;
  var opseg = list.getRange(2, kolona, brojRedova, 1);
  opseg.setBackgrounds(opseg.getValues().map(function (r) {
    return [r[0] === 'ZAVRŠENO' ? '#d9ead3'
      : r[0] === 'U TOKU' ? '#fff2cc' : '#f4cccc'];
  }));
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
      mapa[k] = { lat: broj_(r[t.i['Lat']]), lng: broj_(r[t.i['Lng']]) };
    }
  });
  return mapa;
}

function osvjeziGeo_(poVrsti, sveUlice, koordinate) {
  var redovi = Object.keys(sveUlice).map(function (k) {
    var u = sveUlice[k];
    var c = poVrsti['CIJEPANO'][k] || { preostalo: 0, isporuceno: 0, korisnika: 0 };
    var d = poVrsti['U DUGOM'][k] || { preostalo: 0, isporuceno: 0, korisnika: 0 };
    var xy = koordinate[k] || { lat: '', lng: '' };
    var adresa = u.ulica && u.ulica !== '(bez ulice)'
      ? u.ulica + ', ' + u.mjesto + ', Bosna i Hercegovina'
      : u.mjesto + ', Bosna i Hercegovina';
    return [u.mjesto, u.ulica, k,
      okrugli_(c.preostalo), okrugli_(c.isporuceno),
      okrugli_(d.preostalo), okrugli_(d.isporuceno),
      okrugli_(c.preostalo + d.preostalo), adresa, xy.lat, xy.lng];
  }).sort(function (a, b) { return b[7] - a[7]; });

  upisi_(LIST_GEO, ['Mjesto', 'Ulica', 'Adresa ključ',
    'Preostalo CIJEPANO m3', 'Isporučeno CIJEPANO m3',
    'Preostalo U DUGOM m3', 'Isporučeno U DUGOM m3',
    'Preostalo ukupno m3', 'Adresa za kartu', 'Lat', 'Lng'], redovi);
}

/* ---------------------------------------------------------- 2. geokodiranje */

function geokodirajUlice() {
  var s = list_(LIST_GEO);
  var t = citaj_(LIST_GEO);
  var geokoder = Maps.newGeocoder().setRegion('ba');
  var uspjeh = 0, neuspjeh = [];

  t.redovi.forEach(function (r, idx) {
    if (r[t.i['Lat']] !== '' && r[t.i['Lng']] !== '') return;  // već ima
    var adresa = String(r[t.i['Adresa za kartu']]).trim();
    if (!adresa) return;
    var odgovor;
    try {
      odgovor = geokoder.geocode(adresa);
    } catch (e) {
      neuspjeh.push(adresa + ' (' + e.message + ')');
      return;
    }
    if (odgovor.status === 'OK' && odgovor.results.length) {
      var loc = odgovor.results[0].geometry.location;
      s.getRange(idx + 2, t.i['Lat'] + 1).setValue(loc.lat);
      s.getRange(idx + 2, t.i['Lng'] + 1).setValue(loc.lng);
      uspjeh++;
    } else {
      neuspjeh.push(adresa);
    }
    Utilities.sleep(200);  // da se ne pređe dnevna kvota prebrzo
  });

  var poruka = 'Geokodirano: ' + uspjeh + '.';
  if (neuspjeh.length) {
    poruka += '\n\nNije pronađeno (' + neuspjeh.length + ') - upišite Lat/Lng ručno:\n'
      + neuspjeh.slice(0, 25).join('\n');
  }
  SpreadsheetApp.getUi().alert(poruka);
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
      bezKoordinata.push(String(r[geo.i['Mjesto']]) + ' - ' + String(r[geo.i['Ulica']]));
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
