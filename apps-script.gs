/**
 * Knihobot – knižní kvíz: registrační backend (Google Apps Script)
 *
 * Čo to robí:
 *  - Čte a vynucuje kapacitu (5 týmů / 25 osob celkem) priamo z Google Sheetu.
 *  - Každé odoslane = jeden riadok. Každá otázka = jeden stĺpec.
 *  - Komunikuje cez JSONP (GET + parameter ?callback=...), takže frontend
 *    z GitHub Pages nenaráža na CORS a vie prečítať výsledok (ok / plno).
 *
 * ── AKO NASADIŤ ──────────────────────────────────────────────────────────────
 *  1. Otvor cieľový Google Sheet
 *     (https://docs.google.com/spreadsheets/d/14ZFlyXrGc3X4focG65UPyo_PiCnQD0BEpcEqhImVqpk/edit)
 *  2. Menu Rozšírenia → Apps Script.
 *  3. Zmaž ukážkový kód a vlož celý tento súbor. Ulož (ikona diskety).
 *  4. Nasadiť (Deploy) → Nové nasadene → typ „Webová aplikácia".
 *       - Spustiť ako: Ja (tvoj účet)
 *       - Kto má prístup: Ktokoľvek (Anyone)
 *  5. Autorizuj (pri „Google toto neoveril" → Rozšírené → Prejsť na projekt).
 *  6. Skopíruj „URL webovej aplikácie" (končí na /exec) a vlož ju do index.html
 *     do konstanty ENDPOINT_URL (nebo mi ji pošli).
 *  ── Když později změníš kód: Spravovat nasazení → upravit → Nová verze (jinak
 *     beží stará verzia ďalej).
 * ─────────────────────────────────────────────────────────────────────────────
 */

var SHEET_PRIHLASKY = 'Prihlasky';
var SHEET_NAHRADNICI = 'Náhradníci';

var MAX_TEAMS = 5;      // max počet týmů
var MAX_PEOPLE = 25;    // max počet osob celkem (5 týmů × 5)

var HEADERS_PRIHLASKY = ['Čas', 'Typ', 'Název týmu', 'Počet členů', 'Jméno/přezdívka', 'E-mail', 'Souhlas', 'Počet osob'];
var HEADERS_NAHRADNICI = ['Čas', 'E-mail', 'Kontext'];

// Jediný vstupný bod – všetko cez GET/JSONP.
function doGet(e) {
  var params = (e && e.parameter) ? e.parameter : {};
  var action = params.action || 'status';
  var out;

  if (action === 'status') {
    out = getCounts();
  } else if (action === 'register') {
    out = handleRegister(params);
  } else if (action === 'nahradnik') {
    out = handleNahradnik(params);
  } else {
    out = { result: 'error', message: 'neznámá akce' };
  }

  return respond(out, params.callback);
}

// Vráti JSONP (callback(...)) alebo čistý JSON, ak callback ne je zadaný.
function respond(obj, callback) {
  var json = JSON.stringify(obj);
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function sheet_(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); }
  if (sh.getLastRow() === 0) { sh.appendRow(headers); }
  return sh;
}

// Spočíta aktuálny stav zo Sheetu.
function getCounts() {
  var sh = sheet_(SHEET_PRIHLASKY, HEADERS_PRIHLASKY);
  var last = sh.getLastRow();
  var teamsCount = 0, peopleCount = 0;
  if (last > 1) {
    // sloupec B = Typ (index 2), sloupec H = Počet osob (index 8)
    var values = sh.getRange(2, 1, last - 1, HEADERS_PRIHLASKY.length).getValues();
    for (var i = 0; i < values.length; i++) {
      var typ = String(values[i][1] || '').toLowerCase();
      var osoby = Number(values[i][7]) || 0;
      if (typ === 'tým' || typ === 'tim') { teamsCount++; }
      peopleCount += osoby;
    }
  }
  var spotsLeft = Math.max(0, MAX_PEOPLE - peopleCount);
  return {
    result: 'ok',
    teamsCount: teamsCount,
    peopleCount: peopleCount,
    individualSpotsLeft: spotsLeft,
    teamFull: (teamsCount >= MAX_TEAMS) || (peopleCount >= MAX_PEOPLE),
    peopleFull: (peopleCount >= MAX_PEOPLE)
  };
}

// Zápis přihlášky – autoritatívna kontrola kapacity pod zámkom.
function handleRegister(p) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(15000); } catch (err) { return { result: 'error', message: 'busy' }; }
  try {
    var type = String(p.type || '').toLowerCase();
    var isTeam = (type === 'tim' || type === 'tým');
    var members = isTeam ? Math.max(2, Math.min(5, parseInt(p.members, 10) || 0)) : 1;

    var c = getCounts();

    if (isTeam) {
      if (c.teamsCount >= MAX_TEAMS) return { result: 'full', reason: 'teams', individualSpotsLeft: c.individualSpotsLeft };
      if (c.peopleCount + members > MAX_PEOPLE) return { result: 'full', reason: 'people' };
    } else {
      if (c.peopleCount >= MAX_PEOPLE) return { result: 'full', reason: 'people' };
    }

    var sh = sheet_(SHEET_PRIHLASKY, HEADERS_PRIHLASKY);
    sh.appendRow([
      new Date(),                       // Čas
      isTeam ? 'tým' : 'jednotlivec',   // Typ
      isTeam ? (p.teamName || '') : '', // Názov týmu
      isTeam ? members : '',            // Počet členů
      p.name || '',                     // Meno/prezývka
      p.email || '',                    // E-mail
      (String(p.consent) === 'ano') ? 'ano' : 'ne',  // Souhlas (dobrovolný)
      members                           // Počet osob (pro součet kapacity)
    ]);
    return { result: 'ok' };
  } catch (err) {
    return { result: 'error', message: String(err) };
  } finally {
    lock.releaseLock();
  }
}

// Zápis e-mailu náhradníka.
function handleNahradnik(p) {
  try {
    var email = String(p.email || '').trim();
    if (!email) return { result: 'error', message: 'chýba e-mail' };
    var sh = sheet_(SHEET_NAHRADNICI, HEADERS_NAHRADNICI);
    sh.appendRow([new Date(), email, p.context || '']);
    return { result: 'ok' };
  } catch (err) {
    return { result: 'error', message: String(err) };
  }
}
