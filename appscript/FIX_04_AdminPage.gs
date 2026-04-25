// FIX_04_AdminPage.gs
// -----------------------------------------------------------------
// PURPOSE: Serves the secure pub admin web page.
//
// WHAT IT DOES:
//   1. Validates admin token from URL (?adminToken=TOKEN)
//   2. Shows: Add New Participant / Amend Participant / View All
//   3. Enforces lock - no adds/edits after Wednesday 5pm
//   4. Calls FIX_01 (duplicate check) and FIX_03 (tiebreaker) on submit
//   5. Writes every action to AdminActions audit log sheet
//
// HOW TO DEPLOY:
//   Deploy > New Deployment > Web App
//   Execute as: Me  |  Who has access: Anyone
//   Admin link format: [deploy_url]?adminToken=[TOKEN]&mode=admin
//
// ADD TO YOUR EXISTING doGet():
//   function doGet(e) {
//     if ((e.parameter || {}).mode === 'admin') return doGetAdmin(e);
//     // ... rest of your existing doGet ...
//   }
//
// REQUIRES: FIX_01_DuplicatePickBlocker.gs and
//           FIX_03_TiebreakerEnforcement.gs in the same project.
// -----------------------------------------------------------------


function doGetAdmin(e) {
  var params = e ? (e.parameter || {}) : {};
  var token = params.adminToken || '';

  if (!token) {
    return HtmlService.createHtmlOutput(buildAdminErrorPage_('No admin token provided.'));
  }

  var tokenData = validateAdminToken_(token);
  if (!tokenData.valid) {
    return HtmlService.createHtmlOutput(buildAdminErrorPage_('Invalid or expired token. Contact the sweep organiser.'));
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var config = getConfigAdmin_(ss);
  var isLocked = isTournamentLocked_(config);
  var participants = getGroupParticipants_(ss, tokenData.groupCode);

  var html = buildAdminPageHtml_({
    groupName: tokenData.groupName,
    contactName: tokenData.contactName,
    groupCode: tokenData.groupCode,
    token: token,
    participants: participants,
    isLocked: isLocked,
    tournamentName: config.tournamentName,
    lockDateTime: config.lockDateTime,
    entryFee: config.entryFee
  });

  return HtmlService.createHtmlOutput(html).setTitle('Admin - ' + tokenData.groupName);
}


// Called by google.script.run from admin page
function handleAdminAction(payload) {
  var token = payload.adminToken || '';
  var tokenData = validateAdminToken_(token);
  if (!tokenData.valid) return { ok: false, error: 'Invalid token. Action rejected.' };

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var config = getConfigAdmin_(ss);
  var action = payload.action || '';
  var isLocked = isTournamentLocked_(config);

  var picks = {
    b1p1: payload.b1p1, b1p2: payload.b1p2,
    b2p1: payload.b2p1, b2p2: payload.b2p2,
    b3p1: payload.b3p1, b3p2: payload.b3p2,
    b4p1: payload.b4p1, b4p2: payload.b4p2
  };

  if (action === 'add') {
    if (isLocked) return { ok: false, error: 'Entries are locked. Cannot add after the deadline.' };

    var dupCheck = validateNoDuplicatePicks(picks);
    if (!dupCheck.ok) return { ok: false, error: dupCheck.error };

    var tbCheck = validateTiebreakerPresent(payload.tieBreaker, picks);
    if (!tbCheck.ok) return { ok: false, error: tbCheck.error };

    var existing = getParticipantByName_(ss, payload.participantName);
    if (existing) return { ok: false, error: '"' + payload.participantName + '" already exists. Use Amend to edit.' };

    var partSheet = ss.getSheetByName('Participants');
    partSheet.appendRow([
      payload.participantName.trim(),
      payload.b1p1, payload.b1p2,
      payload.b2p1, payload.b2p2,
      payload.b3p1, payload.b3p2,
      payload.b4p1, payload.b4p2,
      payload.tieBreaker,
      new Date().toISOString()
    ]);

    var gmSheet = ss.getSheetByName('GroupMembers');
    if (gmSheet) gmSheet.appendRow([tokenData.groupCode, payload.participantName.trim()]);

    logAdminAction_(ss, {
      timestamp: new Date(), groupCode: tokenData.groupCode,
      groupName: tokenData.groupName, action: 'ADD',
      participantName: payload.participantName, detail: 'Added via admin page'
    });

    return { ok: true, message: payload.participantName + ' has been added successfully.' };
  }

  if (action === 'amend') {
    if (isLocked) return { ok: false, error: 'Entries are locked. Amendments not permitted after the deadline.' };

    var dupCheck2 = validateNoDuplicatePicks(picks);
    if (!dupCheck2.ok) return { ok: false, error: dupCheck2.error };

    var tbCheck2 = validateTiebreakerPresent(payload.tieBreaker, picks);
    if (!tbCheck2.ok) return { ok: false, error: tbCheck2.error };

    var updated = updateParticipantPicks_(ss, payload.participantName, payload);
    if (!updated) return { ok: false, error: 'Could not find "' + payload.participantName + '".' };

    logAdminAction_(ss, {
      timestamp: new Date(), groupCode: tokenData.groupCode,
      groupName: tokenData.groupName, action: 'AMEND',
      participantName: payload.participantName, detail: 'Picks updated via admin page'
    });

    return { ok: true, message: payload.participantName + '\'s picks have been updated.' };
  }

  return { ok: false, error: 'Unknown action: ' + action };
}


// Returns array of 4 arrays of player names, one per bracket
// Called by google.script.run from client-side JS
function getBracketPlayers() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var brackets = ['Bracket1', 'Bracket2', 'Bracket3', 'Bracket4'];
  return brackets.map(function(bName) {
    var sheet = ss.getSheetByName(bName);
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    var players = [];
    for (var i = 1; i < data.length; i++) {
      if (data[i][0]) players.push(data[i][0].toString().trim());
    }
    return players;
  });
}


// HELPERS

function validateAdminToken_(token) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tokensSheet = ss.getSheetByName('Tokens');
  if (!tokensSheet) return { valid: false };

  var data = tokensSheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var storedToken = (data[i][0] || '').toString().trim();
    if (storedToken !== token) continue;

    var groupCode = (data[i][2] || '').toString().trim();
    var contactName = (data[i][1] || '').toString().trim();
    var groupName = groupCode;

    var grpSheet = ss.getSheetByName('Groups');
    if (grpSheet) {
      var grpData = grpSheet.getDataRange().getValues();
      for (var g = 1; g < grpData.length; g++) {
        if ((grpData[g][1] || '').toString().trim() === groupCode) {
          groupName = (grpData[g][0] || '').toString().trim();
          break;
        }
      }
    }
    return { valid: true, groupCode: groupCode, groupName: groupName, contactName: contactName };
  }
  return { valid: false };
}

function isTournamentLocked_(config) {
  if (!config.lockDateTime) return false;
  try { return new Date() > new Date(config.lockDateTime); } catch (e) { return false; }
}

function getGroupParticipants_(ss, groupCode) {
  var gmSheet = ss.getSheetByName('GroupMembers');
  var partSheet = ss.getSheetByName('Participants');
  if (!gmSheet || !partSheet) return [];

  var gmData = gmSheet.getDataRange().getValues();
  var memberNames = {};
  for (var i = 1; i < gmData.length; i++) {
    if ((gmData[i][0] || '').toString().trim() === groupCode) {
      memberNames[(gmData[i][1] || '').toString().trim().toLowerCase()] = true;
    }
  }

  var partData = partSheet.getDataRange().getValues();
  var result = [];
  for (var r = 1; r < partData.length; r++) {
    var pName = (partData[r][0] || '').toString().trim();
    if (memberNames[pName.toLowerCase()]) {
      result.push({ name: pName, hasPicks: !!(partData[r][1] || partData[r][2]) });
    }
  }
  return result;
}

function getParticipantByName_(ss, name) {
  var partSheet = ss.getSheetByName('Participants');
  if (!partSheet) return null;
  var data = partSheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if ((data[i][0] || '').toString().trim().toLowerCase() === (name || '').trim().toLowerCase()) {
      return { row: i + 1, data: data[i] };
    }
  }
  return null;
}

function updateParticipantPicks_(ss, name, payload) {
  var existing = getParticipantByName_(ss, name);
  if (!existing) return false;
  var partSheet = ss.getSheetByName('Participants');
  partSheet.getRange(existing.row, 2, 1, 9).setValues([[
    payload.b1p1, payload.b1p2,
    payload.b2p1, payload.b2p2,
    payload.b3p1, payload.b3p2,
    payload.b4p1, payload.b4p2,
    payload.tieBreaker
  ]]);
  return true;
}

function logAdminAction_(ss, entry) {
  var logSheet = ss.getSheetByName('AdminActions');
  if (!logSheet) {
    logSheet = ss.insertSheet('AdminActions');
    logSheet.appendRow(['Timestamp','GroupCode','GroupName','Action','ParticipantName','Detail']);
  }
  logSheet.appendRow([
    entry.timestamp || new Date(),
    entry.groupCode || '',
    entry.groupName || '',
    entry.action || '',
    entry.participantName || '',
    entry.detail || ''
  ]);
}

function getConfigAdmin_(ss) {
  var sheet = ss.getSheetByName('Config');
  if (!sheet) return {};
  var data = sheet.getDataRange().getValues();
  var c = {};
  data.forEach(function(row) { if (row[0]) c[row[0].toString().trim()] = (row[1] || '').toString().trim(); });
  return {
    tournamentName: c['Tournament Name'] || '',
    entryFee: c['Entry Fee'] || '10',
    lockDateTime: c['Lock DateTime'] || '',
    mcPenalty: parseInt(c['MC Penalty (strokes added)'] || c['MC Penalty'] || '5', 10)
  };
}

function buildAdminErrorPage_(msg) {
  return '<html><body style="font-family:sans-serif;padding:40px;max-width:500px;margin:auto">'
    + '<h2>Access Denied</h2><p>' + msg + '</p>'
    + '<p style="color:#888;font-size:0.9em">SnipeGolf Admin</p></body></html>';
}

function buildAdminPageHtml_(data) {
  var lockBanner = data.isLocked
    ? '<div style="background:#f4cccc;border:1px solid #c00;padding:12px 16px;border-radius:8px;margin-bottom:20px">LOCKED: Entries closed (past Wednesday 5pm). View only.</div>'
    : '';

  var participantRows = (data.participants || []).map(function(p) {
    return '<tr><td>' + p.name + '</td>'
      + '<td>' + (p.hasPicks ? 'Picks submitted' : 'No picks yet') + '</td>'
      + '<td>' + (data.isLocked ? '' : '<button onclick="loadAmend(\'' + p.name.replace(/\'/g, '') + '\')">Edit</button>') + '</td></tr>';
  }).join('');

  return '<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8">'
    + '<meta name="viewport" content="width=device-width,initial-scale=1">'
    + '<title>Admin - ' + data.groupName + '</title>'
    + '<style>'
    + 'body{font-family:Arial,sans-serif;background:#f7f3ea;color:#1b1b16;margin:0}'
    + '.hdr{background:#163d28;color:#fff;padding:18px 22px}'
    + '.hdr h1{margin:0;font-size:1.25rem}.hdr p{margin:4px 0 0;opacity:.75;font-size:.88rem}'
    + '.main{max-width:660px;margin:0 auto;padding:18px}'
    + '.card{background:#fff;border:1px solid #ddd;border-radius:10px;padding:16px;margin-bottom:16px}'
    + 'h2{color:#1f5a3a;margin:0 0 12px;font-size:.98rem}'
    + 'label{display:block;font-weight:bold;font-size:.83rem;margin-bottom:3px}'
    + 'select,input[type=text]{width:100%;padding:8px 10px;border:1px solid #ccc;border-radius:6px;font-size:.93rem;margin-bottom:9px;box-sizing:border-box}'
    + '.btn{padding:10px 20px;border:none;border-radius:7px;font-weight:bold;font-size:.93rem;cursor:pointer}'
    + '.btn-primary{background:#1f5a3a;color:#fff;width:100%;margin-top:4px}'
    + '.btn-secondary{background:#eee;color:#333;margin-top:8px}'
    + '.act-row{display:flex;gap:10px;flex-wrap:wrap}'
    + '.act-btn{flex:1;min-width:150px;background:#fff;border:2px solid #1f5a3a;color:#1f5a3a;padding:13px;border-radius:9px;font-weight:bold;text-align:center;cursor:pointer}'
    + '.act-btn:hover{background:#1f5a3a;color:#fff}'
    + '.msg{padding:9px 13px;border-radius:7px;margin-bottom:10px;font-weight:bold}'
    + '.msg-ok{background:#d9ead3;color:#2a5a0d}.msg-err{background:#f4cccc;color:#900}'
    + 'table{width:100%;border-collapse:collapse;font-size:.88rem}'
    + 'th{text-align:left;padding:7px 8px;background:#f0ece0;font-size:.76rem;text-transform:uppercase}'
    + 'td{padding:8px;border-bottom:1px solid #eee}'
    + '.bg{background:#f7f3ea;border-radius:7px;padding:10px;margin-bottom:8px}'
    + '.bg h3{margin:0 0 7px;font-size:.8rem;text-transform:uppercase;color:#777}'
    + '.pr{display:grid;grid-template-columns:1fr 1fr;gap:8px}'
    + '#form-section{display:none}#list-section{display:none}'
    + '</style></head><body>'
    + '<div class="hdr"><h1>Golf Sweep - ' + data.groupName + ' Admin</h1>'
    + '<p>Welcome ' + data.contactName + ' | Tournament: ' + data.tournamentName + '</p></div>'
    + '<div class="main">' + lockBanner
    + '<div class="card"><h2>What would you like to do?</h2><div class="act-row">'
    + (data.isLocked ? '' : '<button class="act-btn" onclick="showForm(\'add\')">+ Add New Participant</button>')
    + (data.isLocked ? '' : '<button class="act-btn" onclick="showSection(\'list\')">Edit a Participant</button>')
    + '<button class="act-btn" onclick="showSection(\'list\')">' + (data.isLocked ? 'View' : 'View') + ' All Participants</button>'
    + '</div></div>'
    + '<div id="msg-area"></div>'
    + '<div id="list-section" class="card"><h2>Participants - ' + data.groupName + '</h2>'
    + (participantRows ? '<table><thead><tr><th>Name</th><th>Status</th><th></th></tr></thead><tbody>' + participantRows + '</tbody></table>' : '<p>No participants yet.</p>')
    + '</div>'
    + '<div id="form-section" class="card"><h2 id="form-title">Participant</h2>'
    + '<div id="msg-form"></div>'
    + '<label>Participant Name</label><input type="text" id="f-name" placeholder="e.g. John Murphy" />'
    + '<div class="bg"><h3>Bracket 1</h3><div class="pr">'
    + '<div><label>Pick 1</label><select id="f-b1p1"><option value="">-- Select --</option></select></div>'
    + '<div><label>Pick 2</label><select id="f-b1p2"><option value="">-- Select --</option></select></div>'
    + '</div></div>'
    + '<div class="bg"><h3>Bracket 2</h3><div class="pr">'
    + '<div><label>Pick 1</label><select id="f-b2p1"><option value="">-- Select --</option></select></div>'
    + '<div><label>Pick 2</label><select id="f-b2p2"><option value="">-- Select --</option></select></div>'
    + '</div></div>'
    + '<div class="bg"><h3>Bracket 3</h3><div class="pr">'
    + '<div><label>Pick 1</label><select id="f-b3p1"><option value="">-- Select --</option></select></div>'
    + '<div><label>Pick 2</label><select id="f-b3p2"><option value="">-- Select --</option></select></div>'
    + '</div></div>'
    + '<div class="bg"><h3>Bracket 4</h3><div class="pr">'
    + '<div><label>Pick 1</label><select id="f-b4p1"><option value="">-- Select --</option></select></div>'
    + '<div><label>Pick 2</label><select id="f-b4p2"><option value="">-- Select --</option></select></div>'
    + '</div></div>'
    + '<div class="bg"><h3>Tiebreaker (one of your 8 picks)</h3>'
    + '<select id="f-tb"><option value="">-- Select tiebreaker --</option></select></div>'
    + '<button class="btn btn-primary" onclick="submitForm()">Submit</button>'
    + '<br><button class="btn btn-secondary" onclick="resetView()">Back</button>'
    + '</div></div>'
    + '<script>'
    + 'var currentAction="",adminToken="' + data.token + '";'
    + 'function showSection(s){document.getElementById("list-section").style.display=s==="list"?"block":"none";document.getElementById("form-section").style.display=s==="form"?"block":"none";}'
    + 'function showForm(a){currentAction=a;document.getElementById("form-title").textContent=a==="add"?"Add New Participant":"Amend Participant";document.getElementById("f-name").readOnly=(a==="amend");showSection("form");loadBracketPlayers();}'
    + 'function loadAmend(n){document.getElementById("f-name").value=n;showForm("amend");}'
    + 'function resetView(){showSection("none");document.getElementById("msg-area").innerHTML="";}'
    + 'function gp(){return{b1p1:document.getElementById("f-b1p1").value,b1p2:document.getElementById("f-b1p2").value,b2p1:document.getElementById("f-b2p1").value,b2p2:document.getElementById("f-b2p2").value,b3p1:document.getElementById("f-b3p1").value,b3p2:document.getElementById("f-b3p2").value,b4p1:document.getElementById("f-b4p1").value,b4p2:document.getElementById("f-b4p2").value};}'
    + 'function updateTB(){var p=gp(),tb=document.getElementById("f-tb"),cur=tb.value,vals=Object.values(p).filter(Boolean);tb.innerHTML="<option value=\'\'>-- Select tiebreaker --</option>";vals.forEach(function(v){var o=document.createElement("option");o.value=v;o.textContent=v;tb.appendChild(o);});if(cur&&vals.indexOf(cur)>-1)tb.value=cur;}'
    + 'function loadBracketPlayers(){google.script.run.withSuccessHandler(function(pl){var bs=["b1","b2","b3","b4"];bs.forEach(function(b,bi){["p1","p2"].forEach(function(p){var s=document.getElementById("f-"+b+p);if(!s)return;s.innerHTML="<option value=\'\'>-- Select --</option>";(pl[bi]||[]).forEach(function(n){var o=document.createElement("option");o.value=n;o.textContent=n;s.appendChild(o);});s.addEventListener("change",updateTB);});});}).getBracketPlayers();}'
    + 'function submitForm(){var name=document.getElementById("f-name").value.trim();if(!name){showMsg("form","err","Please enter a participant name.");return;}var picks=gp(),tb=document.getElementById("f-tb").value,payload=Object.assign({},picks,{action:currentAction,participantName:name,tieBreaker:tb,adminToken:adminToken});document.querySelector(".btn-primary").disabled=true;document.querySelector(".btn-primary").textContent="Saving...";google.script.run.withSuccessHandler(function(r){document.querySelector(".btn-primary").disabled=false;document.querySelector(".btn-primary").textContent="Submit";if(r.ok){showMsg("area","ok",r.message);resetView();}else{showMsg("form","err",r.error);}}).withFailureHandler(function(e){document.querySelector(".btn-primary").disabled=false;document.querySelector(".btn-primary").textContent="Submit";showMsg("form","err","Error: "+e.message);}).handleAdminAction(payload);}'
    + 'function showMsg(area,type,text){var el=document.getElementById("msg-"+area);if(el)el.innerHTML="<div class=\'msg msg-"+type+"\'>"+(type==="ok"?"OK: ":"Error: ")+text+"</div>";}'
    + '<\/script></body></html>';
}
