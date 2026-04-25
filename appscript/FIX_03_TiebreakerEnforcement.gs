// FIX_03_TiebreakerEnforcement.gs
// -----------------------------------------------------------------
// PURPOSE: Two-part fix for the missing tiebreaker gap.
//
// PART A - Entry validation: TieBreaker field must be filled in
//          before an entry is accepted. Must be one of the 8 picks.
//
// PART B - Leaderboard sort: When two participants share identical
//          total score, use TieBreaker pick's final score as
//          secondary sort (lower = wins). Name alphabetical is
//          the final fallback.
//
// HOW TO USE:
//   PART A - In your entry submission handler:
//     var tbCheck = validateTiebreakerPresent(tiebreakerValue, picks);
//     if (!tbCheck.ok) return error to user;
//
//   PART B - In calculateLeaderboard(), replace your .sort() with:
//     var sorted = sortLeaderboardWithTiebreaker(participantRows);
// -----------------------------------------------------------------


// PART A: Entry-time tiebreaker validation

function validateTiebreakerPresent(tiebreakerValue, picks) {
  var tb = (tiebreakerValue || '').trim();

  if (!tb) {
    return {
      ok: false,
      error: 'Tiebreaker is required. Please select one of your 8 picks as your tiebreaker player. ' +
             'In the event of a tied score, the participant whose tiebreaker player finishes highest wins.'
    };
  }

  if (picks) {
    var allPicks = [
      picks.b1p1, picks.b1p2, picks.b2p1, picks.b2p2,
      picks.b3p1, picks.b3p2, picks.b4p1, picks.b4p2
    ].map(function(p) { return (p || '').trim().toLowerCase(); }).filter(Boolean);

    if (allPicks.length > 0 && allPicks.indexOf(tb.toLowerCase()) === -1) {
      return {
        ok: false,
        error: 'Tiebreaker must be one of your 8 picks. "' + tb + '" is not in your selections.'
      };
    }
  }

  return { ok: true, error: null };
}


// PART B: Leaderboard tiebreaker sort

// Low-level sort - call if you already have scoresMap built
// @param rows - array of { name, totalScore, tieBreaker, ... }
// @param scoresMap - { 'player name lowercase': scoreNumber }
// @param mcPenalty - number e.g. 5
function applyTiebreakerSort(rows, scoresMap, mcPenalty) {
  var penalty = mcPenalty || 5;

  rows.forEach(function(row) {
    var tb = (row.tieBreaker || '').trim();
    if (tb) {
      var tbScore = scoresMap[tb.toLowerCase()];
      row.tiebreakerScore = (tbScore !== undefined) ? Number(tbScore) : penalty;
      row.tiebreakerUsed = tb;
    } else {
      row.tiebreakerScore = 999;
      row.tiebreakerUsed = null;
    }
  });

  rows.sort(function(a, b) {
    if (a.totalScore !== b.totalScore) return a.totalScore - b.totalScore;
    if (a.tiebreakerScore !== b.tiebreakerScore) return a.tiebreakerScore - b.tiebreakerScore;
    return (a.name || '').localeCompare(b.name || '');
  });

  for (var i = 0; i < rows.length; i++) {
    if (i === 0) {
      rows[i].rank = 1;
    } else {
      var prev = rows[i - 1];
      var curr = rows[i];
      if (curr.totalScore === prev.totalScore && curr.tiebreakerScore === prev.tiebreakerScore) {
        curr.rank = prev.rank;
      } else {
        curr.rank = i + 1;
      }
    }
  }

  return rows;
}


// Convenience wrapper - reads Scores sheet and builds scoresMap
// Replace plain .sort() in calculateLeaderboard() with this:
//   var sorted = sortLeaderboardWithTiebreaker(participantRows);
function sortLeaderboardWithTiebreaker(participantRows) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var scoresSheet = ss.getSheetByName('Scores');
  var config = getConfigFix03_(ss);

  var scoresMap = {};
  if (scoresSheet) {
    var scoreData = scoresSheet.getDataRange().getValues();
    for (var i = 1; i < scoreData.length; i++) {
      var name = (scoreData[i][0] || '').trim();
      var score = parseFloat(scoreData[i][2]); // ToPar column index 2
      if (name && !isNaN(score)) {
        scoresMap[name.toLowerCase()] = score;
      }
    }
  }

  return applyTiebreakerSort(participantRows, scoresMap, config.mcPenalty);
}

function getConfigFix03_(ss) {
  var sheet = ss.getSheetByName('Config');
  if (!sheet) return { mcPenalty: 5 };
  var data = sheet.getDataRange().getValues();
  var mc = 5;
  data.forEach(function(row) {
    var key = (row[0] || '').toString().trim();
    if (key === 'MC Penalty (strokes added)' || key === 'MC Penalty') {
      mc = parseInt(row[1] || '5', 10);
    }
  });
  return { mcPenalty: mc };
}


// TEST FUNCTION
// Run: Extensions > Apps Script > Run > testTiebreakerSort
function testTiebreakerSort() {
  var fakeScores = {
    'matt fitzpatrick': -18,
    'si woo kim': -16,
    'bud cauley': -12,
    'ryan fox': -9,
    'denny mccarthy': -2
  };

  var fakeRows = [
    { name: 'Alice',   totalScore: -48, tieBreaker: 'Matt Fitzpatrick' },
    { name: 'Bob',     totalScore: -48, tieBreaker: 'Si Woo Kim' },
    { name: 'Charlie', totalScore: -48, tieBreaker: '' },
    { name: 'Dave',    totalScore: -42, tieBreaker: 'Ryan Fox' },
    { name: 'Eve',     totalScore: -42, tieBreaker: 'Denny McCarthy' },
    { name: 'Frank',   totalScore: -30, tieBreaker: 'Bud Cauley' }
  ];

  var sorted = applyTiebreakerSort(fakeRows, fakeScores, 5);

  Logger.log('FIX_03 TIEBREAKER SORT TEST:');
  sorted.forEach(function(r) {
    Logger.log('  Rank ' + r.rank + ': ' + r.name +
      ' | Total: ' + r.totalScore +
      ' | TB: ' + (r.tiebreakerUsed || 'NONE') +
      ' | TB score: ' + r.tiebreakerScore);
  });

  var pass = sorted[0].name === 'Alice' && sorted[1].name === 'Bob' &&
             sorted[2].name === 'Charlie' && sorted[3].name === 'Dave' &&
             sorted[4].name === 'Eve' && sorted[5].name === 'Frank';

  Logger.log(pass ? 'ALL ASSERTIONS PASS' : 'ASSERTIONS FAILED - check order above');
  return sorted;
}
