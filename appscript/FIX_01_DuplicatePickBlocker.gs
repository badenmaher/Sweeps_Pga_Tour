// FIX_01_DuplicatePickBlocker.gs
// -----------------------------------------------------------------
// PURPOSE: Prevents a participant from selecting the same golfer
//          more than once across any of the 8 pick slots.
//
// HOW TO USE: Call validateNoDuplicatePicks(picks) from your entry
//             submission handler BEFORE writing to Participants sheet.
//             If it returns {ok:false}, reject the submission.
//
//   var result = validateNoDuplicatePicks({
//     b1p1: formData.b1pick1, b1p2: formData.b1pick2,
//     b2p1: formData.b2pick1, b2p2: formData.b2pick2,
//     b3p1: formData.b3pick1, b3p2: formData.b3pick2,
//     b4p1: formData.b4pick1, b4p2: formData.b4pick2
//   });
//   if (!result.ok) return ContentService
//     .createTextOutput(JSON.stringify({error: result.error}))
//     .setMimeType(ContentService.MimeType.JSON);
// -----------------------------------------------------------------

function validateNoDuplicatePicks(picks) {
  var slots = [
    { key: 'b1p1', label: 'Bracket 1, Pick 1' },
    { key: 'b1p2', label: 'Bracket 1, Pick 2' },
    { key: 'b2p1', label: 'Bracket 2, Pick 1' },
    { key: 'b2p2', label: 'Bracket 2, Pick 2' },
    { key: 'b3p1', label: 'Bracket 3, Pick 1' },
    { key: 'b3p2', label: 'Bracket 3, Pick 2' },
    { key: 'b4p1', label: 'Bracket 4, Pick 1' },
    { key: 'b4p2', label: 'Bracket 4, Pick 2' }
  ];

  var seen = {};
  var duplicates = [];

  for (var i = 0; i < slots.length; i++) {
    var slot = slots[i];
    var name = (picks[slot.key] || '').trim();
    if (!name) continue;
    var key = name.toLowerCase();
    if (seen[key]) {
      duplicates.push(name + ' (in ' + seen[key] + ' AND ' + slot.label + ')');
    } else {
      seen[key] = slot.label;
    }
  }

  if (duplicates.length > 0) {
    return {
      ok: false,
      error: 'Duplicate player: ' + duplicates.join('; ') +
             '. Each golfer can only appear once across all 8 pick slots.'
    };
  }

  return { ok: true, error: null };
}


// TEST FUNCTION
// Run: Extensions > Apps Script > Run > testDuplicatePickBlocker
function testDuplicatePickBlocker() {
  var results = [];

  // Test 1: Clean entry - should PASS
  var clean = validateNoDuplicatePicks({
    b1p1: 'Robert MacIntyre', b1p2: 'Russell Henley',
    b2p1: 'Corey Conners',    b2p2: 'Sam Burns',
    b3p1: 'Nick Taylor',      b3p2: 'Ryan Fox',
    b4p1: 'Denny McCarthy',   b4p2: 'Adam Schenk'
  });
  results.push('Test 1 (clean): ' + (clean.ok ? 'PASS' : 'FAIL: ' + clean.error));

  // Test 2: Same-bracket duplicate - should BLOCK
  var dupSame = validateNoDuplicatePicks({
    b1p1: 'J.J. Spaun',     b1p2: 'Ben Griffin',
    b2p1: 'Shane Lowry',    b2p2: 'Jake Knapp',
    b3p1: 'Jordan Smith',   b3p2: 'Jordan Smith',
    b4p1: 'Denny McCarthy', b4p2: 'Denny McCarthy'
  });
  results.push('Test 2 (same-bracket dupe): ' + (!dupSame.ok ? 'BLOCKED: ' + dupSame.error : 'FAIL'));

  // Test 3: Cross-bracket duplicate - should BLOCK
  var dupCross = validateNoDuplicatePicks({
    b1p1: 'Scottie Scheffler', b1p2: 'Cameron Young',
    b2p1: 'Cameron Young',     b2p2: 'Ryan Gerard',
    b3p1: 'Michael Thorbjornsen', b3p2: 'Ryo Hisatsune',
    b4p1: 'Billy Horschel',   b4p2: 'Tony Finau'
  });
  results.push('Test 3 (cross-bracket dupe): ' + (!dupCross.ok ? 'BLOCKED: ' + dupCross.error : 'FAIL'));

  // Test 4: Partial empty picks - should PASS
  var partial = validateNoDuplicatePicks({
    b1p1: 'Scottie Scheffler', b1p2: '',
    b2p1: '', b2p2: '', b3p1: '', b3p2: '', b4p1: '', b4p2: ''
  });
  results.push('Test 4 (partial/empty): ' + (partial.ok ? 'PASS' : 'FAIL: ' + partial.error));

  Logger.log(results.join('\n'));
  return results;
}
