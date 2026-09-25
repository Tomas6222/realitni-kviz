/**
 * REMAX Delta – realitní kvíz
 * Google Apps Script Web App endpoint.
 *
 * Nasazení:
 * 1) Otevřete Google Sheet "Leady – realitní kvíz" -> Rozšíření -> Apps Script.
 * 2) Vložte tento soubor jako Code.gs.
 * 3) Project Settings -> Script properties:
 *    SMARTEMAILING_API_KEY = <API klíč>
 *    SMARTEMAILING_USERNAME = ypw
 *    SMARTEMAILING_LIST_ID = 368
 *    SMARTEMAILING_CF_LEAD_ID = <ID vlastního textového pole "Realitní kvíz – Lead">
 *    Volitelně:
 *    SMARTEMAILING_CF_SCORE_ID = <ID pole skóre>
 *    SMARTEMAILING_CF_PROFILE_ID = <ID pole profil>
 *    SMARTEMAILING_CF_REGION_ID = <ID pole lokalita>
 *    SMARTEMAILING_CF_TIMING_ID = <ID pole termín změny>
 * 4) Deploy -> New deployment -> Web app
 *    Execute as: Me
 *    Who has access: Anyone
 * 5) /exec URL vložte do konstanty ENDPOINT v index.html.
 */

const SHEET_NAME = 'Leady';

function doGet() {
  return json_({ ok: true, service: 'remax-delta-realitni-kviz' });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    validate_(data);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error('List "' + SHEET_NAME + '" nebyl nalezen.');

    const now = new Date();
    const profile = String(data.profile || '');
    const answers = JSON.stringify(data.answerDetails || data.answers || []);

    const row = [
      now,                                  // A Přijato
      data.createdAt || '',                 // B Vyplněno
      data.name || '',                      // C Jméno a příjmení
      data.email || '',                     // D E-mail
      data.phone || '',                     // E Telefon
      data.region || '',                    // F Lokalita
      data.timing || '',                    // G Termín změny
      num_(data.score && data.score.overall),      // H Celkové skóre
      num_(data.score && data.score.sales),        // I Obchodní tah
      num_(data.score && data.score.resilience),   // J Odolnost
      num_(data.score && data.score.discipline),   // K Disciplína
      num_(data.score && data.score.people),       // L Práce s lidmi
      profile,                              // M Profil
      data.leadClass || '',                 // N Lead A/B/C
      answers,                              // O Odpovědi
      data.consent ? 'ANO' : 'NE',          // P Souhlas
      data.consentAt || '',                 // Q Souhlas udělen
      data.source || 'realitni-kviz',       // R Zdroj
      data.utm_source || '',                // S
      data.utm_medium || '',                // T
      data.utm_campaign || '',              // U
      data.utm_content || '',               // V
      data.utm_term || '',                  // W
      data.page || '',                      // X Stránka
      data.referrer || '',                  // Y Referrer
      'ČEKÁ',                               // Z SmartEmailing
      ''                                    // AA SmartEmailing odpověď
    ];

    sheet.appendRow(row);
    const rowNumber = sheet.getLastRow();

    let seStatus = 'PŘESKOČENO';
    let seResponse = 'SmartEmailing není nakonfigurován.';

    try {
      const se = sendToSmartEmailing_(data, profile);
      seStatus = se.ok ? 'OK' : 'CHYBA';
      seResponse = se.response;
    } catch (err) {
      seStatus = 'CHYBA';
      seResponse = String(err && err.message ? err.message : err);
    }

    sheet.getRange(rowNumber, 26, 1, 2).setValues([[seStatus, seResponse]]);

    return json_({
      ok: true,
      row: rowNumber,
      smartEmailing: seStatus
    });
  } catch (err) {
    return json_({
      ok: false,
      error: String(err && err.message ? err.message : err)
    });
  } finally {
    lock.releaseLock();
  }
}

function sendToSmartEmailing_(data, profile) {
  const props = PropertiesService.getScriptProperties();
  const username = props.getProperty('SMARTEMAILING_USERNAME') || 'ypw';
  const apiKey = props.getProperty('SMARTEMAILING_API_KEY');
  const listId = Number(props.getProperty('SMARTEMAILING_LIST_ID') || '368');

  if (!apiKey) {
    return { ok: false, response: 'Chybí SMARTEMAILING_API_KEY v Script Properties.' };
  }

  const person = splitName_(data.name || '');
  const note = [
    'REALITNI_KVIZ',
    'LEAD=' + (data.leadClass || ''),
    'SCORE=' + num_(data.score && data.score.overall),
    'PROFILE=' + profile,
    'REGION=' + (data.region || ''),
    'TIMING=' + (data.timing || ''),
    'SOURCE=' + (data.source || 'realitni-kviz')
  ].join(' | ');

  const customFields = [];
  addCustomField_(customFields, props.getProperty('SMARTEMAILING_CF_LEAD_ID'), data.leadClass || '');
  addCustomField_(customFields, props.getProperty('SMARTEMAILING_CF_SCORE_ID'), num_(data.score && data.score.overall));
  addCustomField_(customFields, props.getProperty('SMARTEMAILING_CF_PROFILE_ID'), profile);
  addCustomField_(customFields, props.getProperty('SMARTEMAILING_CF_REGION_ID'), data.region || '');
  addCustomField_(customFields, props.getProperty('SMARTEMAILING_CF_TIMING_ID'), data.timing || '');

  const payload = {
    settings: {
      update: true,
      add_namedays: true,
      add_genders: true,
      add_salutions: true,
      preserve_unsubscribed: true,
      skip_invalid_emails: false
    },
    data: [{
      emailaddress: String(data.email || '').trim(),
      name: person.first,
      surname: person.last,
      cellphone: String(data.phone || '').trim(),
      language: 'cs_CZ',
      notes: note,
      customfields: customFields,
      contactlists: [{
        id: listId,
        status: 'confirmed'
      }]
    }]
  };

  const response = UrlFetchApp.fetch('https://app.smartemailing.cz/api/v3/import', {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    headers: {
      Authorization: 'Basic ' + Utilities.base64Encode(username + ':' + apiKey)
    },
    muteHttpExceptions: true
  });

  const code = response.getResponseCode();
  const text = response.getContentText();
  return {
    ok: code >= 200 && code < 300,
    response: 'HTTP ' + code + ': ' + text.slice(0, 1200)
  };
}

function addCustomField_(target, id, value) {
  const fieldId = Number(id);
  if (!fieldId || value === '' || value === null || typeof value === 'undefined') return;
  target.push({ id: fieldId, value: String(value) });
}

function validate_(data) {
  if (!data || typeof data !== 'object') throw new Error('Neplatná data.');
  if (!String(data.name || '').trim()) throw new Error('Chybí jméno.');
  if (!String(data.email || '').trim()) throw new Error('Chybí e-mail.');
  if (!String(data.phone || '').trim()) throw new Error('Chybí telefon.');
  if (!String(data.region || '').trim()) throw new Error('Chybí lokalita.');
  if (!String(data.timing || '').trim()) throw new Error('Chybí termín změny.');
  if (!data.consent) throw new Error('Chybí souhlas.');
}

function splitName_(fullName) {
  const parts = String(fullName || '').trim().replace(/\s+/g, ' ').split(' ').filter(Boolean);
  if (!parts.length) return { first: '', last: '' };
  if (parts.length === 1) return { first: parts[0], last: '' };
  return { first: parts.slice(0, -1).join(' '), last: parts[parts.length - 1] };
}

function num_(value) {
  const n = Number(value);
  return Number.isFinite(n) ? n : '';
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
