// FIX_02_NameMapAlertOnUnmatched.gs
// -----------------------------------------------------------------
// PURPOSE: Actively flags any player name in Participants sheet
//          that does NOT appear in the Scores sheet.
//          Writes a WARNING to Name Map sheet and sends email alert.
//
// PROBLEM: If a participant picks a player whose name does not
//          exactly match the ESPN import (typo, nickname, extra
//          space), the engine silently gives them the MC penalty.
//          Nobody knows until it is too late.
//
// HOW TO USE: Add one line at the end of importESPNScores() or
//             calculateLeaderboard():
//
//               checkNameMapAlerts();
// -----------------------------------------------------------------

function checkNameMapAlerts() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var config = getConfigFix02_(ss);

  var scoresSheet = ss.getSheetByName('Scores');
  var participantsSheet = ss.getSheetByName('Participants');
  var nameMapSheet = ss.getSheetByName('Name Map');

  if (!scoresSheet || !participantsSheet || !nameMapSheet) {
    Logger.log('FIX_02: Missing required sheet (Scores, Participants, or Name Map). Skipping.');
    return;
  }

  // Build set of all player names in Scores sheet (column A)
  var scoreData = scoresSheet.getDataRange().getValues();
  var scoredPlayers = {};
  for (var i = 1; i < scoreData.length; i++) {
    var playerName = (scoreData[i][0] || '').trim();
    if (playerName) scoredPlayers[playerName.toLowerCase()] = playerName;
  }

  // Check all picks in Participants sheet
  // Columns: Name(0), B1Pick1(1), B1Pick2(2), B2Pick1(3), B2Pick2(4),
  //          B3Pick1(5), B3Pick2(6), B4Pick1(7), B4Pick2(8)
  var partData = participantsSheet.getDataRange().getValues();
  var unmatched = {};

  for (var r = 1; r < partData.length; r++) {
    var participantName = (partData[r][0] || '').trim();
    if (!participantName) continue;
    for (var c = 1; c <= 8; c++) {
      var pick = (partData[r][c] || '').trim();
      if (!pick) continue;
      if (!scoredPlayers[pick.toLowerCase()]) {
        if (!unmatched[pick]) unmatched[pick] = [];
        unmatched[pick].push(participantName);
      }
    }
  }

  var unmatchedNames = Object.keys(unmatched);

  if (unmatchedNames.length === 0) {
    Logger.log('FIX_02: All picks matched to Scores sheet. No alerts needed.');
    return;
  }

  // Write warnings to Name Map sheet
  var nameMapData = nameMapSheet.getDataRange().getValues();
  var nameMapIndex = {};
  for (var nm = 1; nm < nameMapData.length; nm++) {
    nameMapIndex[(nameMapData[nm][0] || '').trim().toLowerCase()] = nm;
  }

  var alertLines = [];
  unmatchedNames.forEach(function(playerName) {
    var whoPickedThem = unmatched[playerName].join(', ');
    alertLines.push('  WARNING: "' + playerName + '" - picked by: ' + whoPickedThem);

    var existingRow = nameMapIndex[playerName.toLowerCase()];
    if (existingRow !== undefined) {
      nameMapSheet.getRange(existingRow + 1, 2).setValue('WARNING - NOT IN SCORES');
      nameMapSheet.getRange(existingRow + 1, 3).setValue('UNMATCHED - picked by: ' + whoPickedThem);
      nameMapSheet.getRange(existingRow + 1, 1, 1, 3).setBackground('#f4cccc');
    } else {
      var lastRow = nameMapSheet.getLastRow() + 1;
      nameMapSheet.getRange(lastRow, 1).setValue(playerName);
      nameMapSheet.getRange(lastRow, 2).setValue('WARNING - NOT IN SCORES');
      nameMapSheet.getRange(lastRow, 3).setValue('UNMATCHED - picked by: ' + whoPickedThem);
      nameMapSheet.getRange(lastRow, 1, 1, 3).setBackground('#f4cccc');
    }
  });

  // Send email alert to admin
  var adminEmail = config.adminEmail || Session.getActiveUser().getEmail();
  if (adminEmail) {
    var subject = 'ALERT: ' + unmatchedNames.length + ' unmatched player name(s) in sweep';
    var body = 'The following player names are in participant picks but NOT in the Scores sheet.\n'
      + 'These picks are receiving the MC penalty (+' + config.mcPenalty + ' strokes).\n'
      + 'Check spelling and update the Name Map sheet.\n\n'
      + alertLines.join('\n')
      + '\n\nTournament: ' + config.tournamentName
      + '\nChecked at: ' + new Date().toLocaleString();
    try {
      GmailApp.sendEmail(adminEmail, subject, body);
      Logger.log('FIX_02: Alert email sent to ' + adminEmail);
    } catch (e) {
      Logger.log('FIX_02: Could not send email: ' + e.message);
    }
  }

  Logger.log('FIX_02: ' + unmatchedNames.length + ' unmatched names: ' + unmatchedNames.join(', '));
}

function getConfigFix02_(ss) {
  var sheet = ss.getSheetByName('Config');
  if (!sheet) return { tournamentName: '', mcPenalty: 5, adminEmail: '' };
  var data = sheet.getDataRange().getValues();
  var c = {};
  data.forEach(function(row) { if (row[0]) c[row[0].toString().trim()] = (row[1] || '').toString().trim(); });
  return {
    tournamentName: c['Tournament Name'] || '',
    mcPenalty: parseInt(c['MC Penalty (strokes added)'] || c['MC Penalty'] || '5', 10),
    adminEmail: c['Pro Shop Email'] || ''
  };
}

// TEST FUNCTION - safe to run, only reads live data and logs results
// Run: Extensions > Apps Script > Run > testNameMapAlerts
function testNameMapAlerts() {
  Logger.log('FIX_02 TEST: Running name map check...');
  checkNameMapAlerts();
  Logger.log('FIX_02 TEST: Done. Check Name Map sheet and this log.');
}
