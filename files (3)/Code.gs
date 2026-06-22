// ============================================================
//  FREE FIRE TOURNAMENT — Google Apps Script Backend
//  Squad-based: 12 squads max, 4 players each
//  Paste into Extensions > Apps Script, then deploy as Web App
// ============================================================

const SHEET_NAME = 'Squads';
const MAX_SQUADS = 12;

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    let data;
  if (e.postData && e.postData.contents) {
    try {
      data = JSON.parse(e.postData.contents);
    } catch (parseErr) {
      return jsonResponse({ ok: false, error: 'Invalid JSON payload' });
    }
  } else if (e.parameter && e.parameter.team && e.parameter.players) {
    try {
      data = {
        team: e.parameter.team,
        players: JSON.parse(e.parameter.players)
      };
    } catch (parseErr) {
      return jsonResponse({ ok: false, error: 'Invalid form payload' });
    }
  } else {
    return jsonResponse({ ok: false, error: 'Missing required data' });
  }
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    let sheet   = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      const headers = [
        'Squad No', 'Team Name',
        'P1 Name', 'P1 IGN', 'P1 UID', 'P1 Phone (Captain)',
        'P2 Name', 'P2 IGN', 'P2 UID',
        'P3 Name', 'P3 IGN', 'P3 UID',
        'P4 Name', 'P4 IGN', 'P4 UID',
        'Registered At'
      ];
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f5a623').setFontColor('#000000');
      sheet.setFrozenRows(1);
    }

    const existing = sheet.getDataRange().getValues();
    const count    = existing.length - 1; // subtract header row

    if (count >= MAX_SQUADS) {
      return jsonResponse({ ok: false, error: 'All 12 squad slots are full!' });
    }

    // Check for duplicate team name
    const teamNames = existing.slice(1).map(r => String(r[1]).toLowerCase());
    if (teamNames.includes(data.team.toLowerCase())) {
      return jsonResponse({ ok: false, error: 'A squad with this team name is already registered.' });
    }

    // Check for duplicate UIDs across all existing entries
    const existingUIDs = [];
    existing.slice(1).forEach(r => {
      [4, 8, 11, 14].forEach(col => { if (r[col]) existingUIDs.push(String(r[col])); });
    });

    const players = data.players;
    for (let i = 0; i < players.length; i++) {
      if (existingUIDs.includes(String(players[i].uid))) {
        return jsonResponse({ ok: false, error: `Player ${i+1} UID (${players[i].uid}) is already registered in another squad.` });
      }
    }

    const now = Utilities.formatDate(
      new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss'
    );

    const newRow = [
      count + 1,
      data.team,
      players[0].name, players[0].ign, players[0].uid, players[0].phone,
      players[1].name, players[1].ign, players[1].uid,
      players[2].name, players[2].ign, players[2].uid,
      players[3].name, players[3].ign, players[3].uid,
      now
    ];

    sheet.appendRow(newRow);

    return jsonResponse({ ok: true, count: count + 1 });

  } catch (err) {
    return jsonResponse({ ok: false, error: err.message });
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    return jsonResponse({ ok: true, count: 0, squads: [] });
  }

  const rows  = sheet.getDataRange().getValues();
  const data  = rows.slice(1); // skip header
  const count = data.length;

  if (e && e.parameter && typeof e.parameter.delete !== 'undefined') {
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const index = parseInt(e.parameter.delete, 10);
      if (!isNaN(index) && index >= 0) {
        const rowToDelete = index + 2; // +1 for 1-based, +1 for header row
        const maxRow = sheet.getLastRow();
        if (rowToDelete <= maxRow) {
          sheet.deleteRow(rowToDelete);
          return jsonResponse({ ok: true });
        } else {
          return jsonResponse({ ok: false, error: 'Squad not found' });
        }
      }
      return jsonResponse({ ok: false, error: 'Invalid index' });
    } catch (err) {
      return jsonResponse({ ok: false, error: err.message });
    } finally {
      lock.releaseLock();
    }
  }

  if (e && e.parameter && e.parameter.full === '1') {
    const squads = data.map(r => ({
      squadNo : r[0],
      team    : r[1],
      players : [
        { name: r[2],  ign: r[3],  uid: r[4],  phone: r[5] },
        { name: r[6],  ign: r[7],  uid: r[8],  phone: ''   },
        { name: r[9],  ign: r[10], uid: r[11], phone: ''   },
        { name: r[12], ign: r[13], uid: r[14], phone: ''   },
      ],
      time: r[15]
    }));
    return jsonResponse({ ok: true, count, squads });
  }

  return jsonResponse({ ok: true, count });
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
