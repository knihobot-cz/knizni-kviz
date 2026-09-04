/**
 * KNIHOBOT – KNIŽNÍ KVÍZ
 * Registrační backend – Google Apps Script
 *
 * Kapacita:
 *   5 týmů
 *   25 lidí celkem
 *
 * Každá registrace vytvoří právě jeden řádek.
 * Pokud je kapacita plná, registrace se neztratí –
 * uloží se do čekací listiny.
 *
 * Komunikace s GitHub Pages probíhá přes JSONP.
 */

// ═══════════════════════════════════════════════════════════════════════════════
// NASTAVENÍ
// ═══════════════════════════════════════════════════════════════════════════════

var SPREADSHEET_ID = '14ZFlyXrGc3X4focG65UPyo_PiCnQD0BEpcEqhImVqpk';

var SHEET_PRIHLASKY = 'Prihlasky';
var SHEET_NAHRADNICI = 'Náhradníci';

var MAX_TEAMS = 5;
var MAX_PEOPLE = 25;


// Hlavičky hlavní tabulky
var HEADERS_PRIHLASKY = [
  'Čas',
  'Typ',
  'Název týmu',
  'Počet členů',
  'Jméno/přezdívka',
  'E-mail',
  'Souhlas',
  'Počet osob',
  'Stav'
];


// Hlavičky čekací listiny
var HEADERS_NAHRADNICI = [
  'Čas',
  'Typ',
  'Název týmu',
  'Počet členů',
  'Jméno/přezdívka',
  'E-mail',
  'Souhlas',
  'Počet osob',
  'Důvod',
  'Stav'
];


// ═══════════════════════════════════════════════════════════════════════════════
// HLAVNÍ VSTUP
// ═══════════════════════════════════════════════════════════════════════════════

function doGet(e) {

  var params = (e && e.parameter) ? e.parameter : {};
  var action = params.action || 'status';

  var out;

  try {

    if (action === 'status') {
      out = getCounts();

    } else if (action === 'register') {
      out = handleRegister(params);

    } else if (action === 'nahradnik') {
      out = handleNahradnik(params);

    } else {

      out = {
        result: 'error',
        message: 'Neznámá akce.'
      };

    }

  } catch (err) {

    out = {
      result: 'error',
      message: String(err)
    };

  }

  return respond(out, params.callback);
}


// ═══════════════════════════════════════════════════════════════════════════════
// JSON / JSONP
// ═══════════════════════════════════════════════════════════════════════════════

function respond(obj, callback) {

  var json = JSON.stringify(obj);

  if (callback) {

    return ContentService
      .createTextOutput(callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);

  }

  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}


// ═══════════════════════════════════════════════════════════════════════════════
// GOOGLE SHEET
// ═══════════════════════════════════════════════════════════════════════════════

function getSpreadsheet_() {

  return SpreadsheetApp.openById(SPREADSHEET_ID);

}


function getSheet_(name, headers) {

  var ss = getSpreadsheet_();

  var sh = ss.getSheetByName(name);

  if (!sh) {
    sh = ss.insertSheet(name);
  }

  // Pokud je tabulka úplně prázdná, vytvoříme hlavičku.
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  return sh;
}


// ═══════════════════════════════════════════════════════════════════════════════
// NORMALIZACE
// ═══════════════════════════════════════════════════════════════════════════════

function normalizeEmail_(email) {

  return String(email || '')
    .trim()
    .toLowerCase();

}


function isTeam_(type) {

  type = String(type || '').toLowerCase().trim();

  return (
    type === 'tým' ||
    type === 'tim' ||
    type === 'team'
  );

}


function getMembers_(p, isTeam) {

  if (!isTeam) {
    return 1;
  }

  var members = parseInt(p.members, 10);

  if (isNaN(members)) {
    members = 0;
  }

  // Tým musí mít 2–5 lidí.
  return Math.max(2, Math.min(5, members));

}


// ═══════════════════════════════════════════════════════════════════════════════
// STAV KAPACITY
// ═══════════════════════════════════════════════════════════════════════════════

function getCounts() {

  var sh = getSheet_(
    SHEET_PRIHLASKY,
    HEADERS_PRIHLASKY
  );

  var lastRow = sh.getLastRow();

  var teamsCount = 0;
  var peopleCount = 0;

  if (lastRow > 1) {

    var values = sh
      .getRange(2, 1, lastRow - 1, HEADERS_PRIHLASKY.length)
      .getValues();

    for (var i = 0; i < values.length; i++) {

      var type = String(values[i][1] || '')
        .toLowerCase()
        .trim();

      var members = Number(values[i][7]) || 0;

      var status = String(values[i][8] || '')
        .toLowerCase()
        .trim();

      // Do kapacity počítáme pouze skutečně potvrzené registrace.
      if (
        status === 'potvrzeno' ||
        status === 'ok' ||
        status === ''
      ) {

        if (
          type === 'tým' ||
          type === 'tim' ||
          type === 'team'
        ) {
          teamsCount++;
        }

        peopleCount += members;
      }
    }
  }

  var peopleLeft = Math.max(
    0,
    MAX_PEOPLE - peopleCount
  );

  var teamsLeft = Math.max(
    0,
    MAX_TEAMS - teamsCount
  );

  return {

    result: 'ok',

    teamsCount: teamsCount,

    teamsLeft: teamsLeft,

    peopleCount: peopleCount,

    peopleLeft: peopleLeft,

    individualSpotsLeft: peopleLeft,

    teamFull:
      teamsCount >= MAX_TEAMS ||
      peopleCount >= MAX_PEOPLE,

    peopleFull:
      peopleCount >= MAX_PEOPLE,

    fullyBooked:
      teamsCount >= MAX_TEAMS ||
      peopleCount >= MAX_PEOPLE

  };

}


// ═══════════════════════════════════════════════════════════════════════════════
// REGISTRACE
// ═══════════════════════════════════════════════════════════════════════════════

function handleRegister(p) {

  var lock = LockService.getScriptLock();

  try {

    // Zabrání tomu, aby dva lidé registrovaní ve stejný okamžik
    // oba prošli kontrolou kapacity.
    lock.waitLock(15000);

  } catch (err) {

    return {
      result: 'error',
      message: 'Registrace se právě zpracovává. Zkus to prosím znovu.'
    };

  }


  try {

    var isTeam = isTeam_(p.type);

    var members = getMembers_(p, isTeam);

    var email = normalizeEmail_(p.email);

    var name = String(p.name || '').trim();

    var teamName = String(p.teamName || '').trim();

    var consent =
      String(p.consent || '').toLowerCase() === 'ano'
        ? 'ano'
        : 'ne';


    // ─────────────────────────────────────────────────────────────────────────
    // ZÁKLADNÍ VALIDACE
    // ─────────────────────────────────────────────────────────────────────────

    if (!email) {

      return {
        result: 'error',
        message: 'Chybí e-mailová adresa.'
      };

    }


    if (!name) {

      return {
        result: 'error',
        message: 'Chybí jméno nebo přezdívka.'
      };

    }


    if (consent !== 'ano') {

      return {
        result: 'error',
        message: 'Pro odeslání registrace je potřeba souhlasit se zpracováním osobních údajů.'
      };

    }


    if (isTeam && !teamName) {

      return {
        result: 'error',
        message: 'Chybí název týmu.'
      };

    }


    // ─────────────────────────────────────────────────────────────────────────
    // DUPLICITNÍ REGISTRACE
    // ─────────────────────────────────────────────────────────────────────────

    var existing = findRegistrationByEmail_(email);

    if (existing) {

      return {
        result: 'already',
        status: existing.status,
        message:
          existing.status === 'čekací listina'
            ? 'Tato e-mailová adresa už je na čekací listině.'
            : 'Tato e-mailová adresa už je zaregistrovaná.'
      };

    }


    // ─────────────────────────────────────────────────────────────────────────
    // KONTROLA KAPACITY
    // ─────────────────────────────────────────────────────────────────────────

    var counts = getCounts();

    var hasCapacity = true;
    var fullReason = '';


    if (isTeam) {

      if (counts.teamsCount >= MAX_TEAMS) {

        hasCapacity = false;
        fullReason = 'naplněná kapacita týmů';

      } else if (
        counts.peopleCount + members > MAX_PEOPLE
      ) {

        hasCapacity = false;
        fullReason = 'naplněná kapacita účastníků';

      }

    } else {

      if (counts.peopleCount >= MAX_PEOPLE) {

        hasCapacity = false;
        fullReason = 'naplněná kapacita účastníků';

      }

    }


    // ─────────────────────────────────────────────────────────────────────────
    // MÁME MÍSTO → POTVRZENÁ REGISTRACE
    // ─────────────────────────────────────────────────────────────────────────

    if (hasCapacity) {

      var sh = getSheet_(
        SHEET_PRIHLASKY,
        HEADERS_PRIHLASKY
      );

      sh.appendRow([

        new Date(),

        isTeam
          ? 'tým'
          : 'jednotlivec',

        isTeam
          ? teamName
          : '',

        isTeam
          ? members
          : '',

        name,

        email,

        consent,

        members,

        'potvrzeno'

      ]);


      return {

        result: 'ok',

        status: 'potvrzeno',

        message: 'Registrace byla úspěšně potvrzena.',

        teamsCount: counts.teamsCount + (isTeam ? 1 : 0),

        peopleCount: counts.peopleCount + members

      };

    }


    // ─────────────────────────────────────────────────────────────────────────
    // NEMÁME MÍSTO → ČEKACÍ LISTINA
    // ─────────────────────────────────────────────────────────────────────────

    var waitSheet = getSheet_(
      SHEET_NAHRADNICI,
      HEADERS_NAHRADNICI
    );

    waitSheet.appendRow([

      new Date(),

      isTeam
        ? 'tým'
        : 'jednotlivec',

      isTeam
        ? teamName
        : '',

      isTeam
        ? members
        : '',

      name,

      email,

      consent,

      members,

      fullReason,

      'čekací listina'

    ]);


    return {

      result: 'waitlist',

      status: 'čekací listina',

      message:
        'Kapacita je momentálně naplněná. Tvoje registrace byla zaznamenána na čekací listině.',

      reason: fullReason

    };


  } catch (err) {

    return {

      result: 'error',

      message: String(err)

    };

  } finally {

    lock.releaseLock();

  }

}


// ═══════════════════════════════════════════════════════════════════════════════
// HLEDÁNÍ DUPLICITY
// ═══════════════════════════════════════════════════════════════════════════════

function findRegistrationByEmail_(email) {

  email = normalizeEmail_(email);

  var ss = getSpreadsheet_();


  // Nejprve hlavní registrace
  var sh = ss.getSheetByName(SHEET_PRIHLASKY);

  if (sh && sh.getLastRow() > 1) {

    var values = sh.getRange(
      2,
      1,
      sh.getLastRow() - 1,
      Math.max(9, sh.getLastColumn())
    ).getValues();

    for (var i = 0; i < values.length; i++) {

      var rowEmail = normalizeEmail_(values[i][5]);

      if (rowEmail === email) {

        return {
          status: String(values[i][8] || 'potvrzeno')
            .toLowerCase()
            .trim()
        };

      }
    }
  }


  // Potom čekací listina
  var wait = ss.getSheetByName(SHEET_NAHRADNICI);

  if (wait && wait.getLastRow() > 1) {

    var waitValues = wait.getRange(
      2,
      1,
      wait.getLastRow() - 1,
      Math.max(10, wait.getLastColumn())
    ).getValues();

    for (var j = 0; j < waitValues.length; j++) {

      var waitEmail = normalizeEmail_(waitValues[j][5]);

      if (waitEmail === email) {

        return {
          status: 'čekací listina'
        };

      }
    }
  }


  return null;

}


// ═══════════════════════════════════════════════════════════════════════════════
// RUČNÍ PŘIDÁNÍ NÁHRADNÍKA
// ═══════════════════════════════════════════════════════════════════════════════

function handleNahradnik(p) {

  try {

    var email = normalizeEmail_(p.email);

    if (!email) {

      return {
        result: 'error',
        message: 'Chybí e-mail.'
      };

    }


    var sh = getSheet_(
      SHEET_NAHRADNICI,
      HEADERS_NAHRADNICI
    );


    sh.appendRow([

      new Date(),

      String(p.type || ''),

      String(p.teamName || ''),

      String(p.members || ''),

      String(p.name || ''),

      email,

      String(p.consent || ''),

      String(p.members || ''),

      String(p.context || ''),

      'čekací listina'

    ]);


    return {

      result: 'ok',

      status: 'čekací listina'

    };


  } catch (err) {

    return {

      result: 'error',

      message: String(err)

    };

  }

}
