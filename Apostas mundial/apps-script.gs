// ============================================================
//  UNCLE PREDICTOR 2026 — Google Apps Script Backend
//  Cola este código no Google Apps Script e faz deploy como Web App
// ============================================================

// ID do teu Google Sheet (da URL: https://docs.google.com/spreadsheets/d/ESTE_ID/edit)
const SHEET_ID = 'COLE_AQUI_O_ID_DO_GOOGLE_SHEET';

// Password do administrador (muda antes de fazer deploy)
const ADMIN_PASSWORD_DEFAULT = 'admin2026';

// ============================================================
//  ENTRY POINTS
// ============================================================

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;
    let result;

    if      (action === 'login')        result = handleLogin(body);
    else if (action === 'adminLogin')   result = handleAdminLogin(body);
    else if (action === 'getData')      result = handleGetData(body);
    else if (action === 'savePreds')    result = handleSavePreds(body);
    else if (action === 'savePrebets')  result = handleSavePrebets(body);
    else if (action === 'saveResults')  result = handleSaveResults(body);
    else if (action === 'getUsers')     result = handleGetUsers(body);
    else if (action === 'addUser')      result = handleAddUser(body);
    else if (action === 'removeUser')   result = handleRemoveUser(body);
    else if (action === 'saveConfig')     result = handleSaveConfig(body);
    else if (action === 'createLeague')  result = handleCreateLeague(body);
    else if (action === 'joinLeague')    result = handleJoinLeague(body);
    else if (action === 'leaveLeague')   result = handleLeaveLeague(body);
    else if (action === 'getMyLeagues')  result = handleGetMyLeagues(body);
    else if (action === 'getLeagues')    result = handleGetLeagues(body);
    else if (action === 'removeLeagueMember') result = handleRemoveLeagueMember(body);
    else if (action === 'deleteLeague')  result = handleDeleteLeague(body);
    else if (action === 'fetchResults')  result = handleFetchResults(body);
    else result = {success: false, error: 'Unknown action'};

    return jsonResponse(result);
  } catch(err) {
    return jsonResponse({success: false, error: err.toString()});
  }
}

function doGet(e) {
  return jsonResponse({status: 'Uncle Predictor 2026 API online'});
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
//  SHEETS HELPER
// ============================================================

function getSheet(name) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    initSheet(sheet, name);
  }
  return sheet;
}

function initSheet(sheet, name) {
  if (name === 'Participants')  sheet.appendRow(['code','name','created_at']);
  if (name === 'Predictions')   sheet.appendRow(['code','game_id','t1','t2','joker','saved_at']);
  if (name === 'Prebets')       sheet.appendRow(['code','key','value','saved_at']);
  if (name === 'Results')       sheet.appendRow(['game_id','t1','t2','updated_at']);
  if (name === 'Config')        sheet.appendRow(['key','value']);
  if (name === 'Ligas')         sheet.appendRow(['id','name','admin_code','invite_code','created_at']);
  if (name === 'Liga_Members')  sheet.appendRow(['league_id','user_code','joined_at']);
}

function sheetToObjects(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}

// ============================================================
//  CONFIG
// ============================================================

function getConfig(key, defaultVal) {
  const sheet = getSheet('Config');
  const rows = sheetToObjects(sheet);
  const row = rows.find(r => r.key === key);
  return row ? row.value : defaultVal;
}

function setConfig(key, value) {
  const sheet = getSheet('Config');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      sheet.getRange(i+1, 2).setValue(value);
      return;
    }
  }
  sheet.appendRow([key, value]);
}

// ============================================================
//  AUTH
// ============================================================

function handleLogin(body) {
  const {name, code} = body;
  if (!name || !code) return {success: false};
  const sheet = getSheet('Participants');
  const rows = sheetToObjects(sheet);
  const user = rows.find(r => r.code === code.toUpperCase());
  if (user) return {success: true, name: user.name};
  return {success: false};
}

function handleAdminLogin(body) {
  const pw = getConfig('admin_password', ADMIN_PASSWORD_DEFAULT);
  return {success: body.password === pw};
}

function requireAdmin(body) {
  const pw = getConfig('admin_password', ADMIN_PASSWORD_DEFAULT);
  if (body.password !== pw) throw new Error('Unauthorized');
}

// ============================================================
//  GET DATA (predictions + results + leaderboard)
// ============================================================

function handleGetData(body) {
  const {code} = body;

  // User predictions
  const predSheet = getSheet('Predictions');
  const predRows = sheetToObjects(predSheet);
  const userPreds = predRows.filter(r => r.code === code);
  const preds = {};
  userPreds.forEach(r => {
    preds[r.game_id] = {
      t1: r.t1 !== '' ? Number(r.t1) : undefined,
      t2: r.t2 !== '' ? Number(r.t2) : undefined,
      joker: r.joker === true || r.joker === 'TRUE',
    };
  });

  // User prebets
  const pbSheet = getSheet('Prebets');
  const pbRows = sheetToObjects(pbSheet);
  const userPb = pbRows.filter(r => r.code === code);
  const prebets = {};
  userPb.forEach(r => { prebets[r.key] = r.value; });

  // Results
  const resSheet = getSheet('Results');
  const resRows = sheetToObjects(resSheet);
  const results = {};
  resRows.forEach(r => {
    if (r.t1 !== '') {
      results[r.game_id] = {t1: Number(r.t1), t2: Number(r.t2)};
    }
  });

  // Leaderboard
  const lb = buildLeaderboard(results);

  // Config
  const deadlineH = Number(getConfig('deadline_hours', 2));

  return {preds, prebets, results, lb, deadlineH};
}

// ============================================================
//  LEADERBOARD CALCULATION
// ============================================================

function buildLeaderboard(results) {
  const partSheet = getSheet('Participants');
  const participants = sheetToObjects(partSheet);

  const predSheet = getSheet('Predictions');
  const allPreds = sheetToObjects(predSheet);

  const pbSheet = getSheet('Prebets');
  const allPrebets = sheetToObjects(pbSheet);

  return participants.map(p => {
    const userPreds = allPreds.filter(r => r.code === p.code);
    const userPrebets = allPrebets.filter(r => r.code === p.code);

    let wins = 0, draws = 0, jokers = 0, groupPts = 0;

    userPreds.forEach(pred => {
      const res = results[pred.game_id];
      if (!res || pred.t1 === '') return;
      // Determine if this is a knockout game by game_id prefix (KO games use phases: R32, R16, QF, SF, F3, FIN)
      const isKO = /^(R32|R16|QF|SF|F3|FIN)/.test(pred.game_id);
      const jokerVal = pred.joker === true || pred.joker === 'TRUE';
      const predObj = {t1: Number(pred.t1), t2: Number(pred.t2), joker: jokerVal};
      const pts = calcPoints(predObj, res, isKO);
      groupPts += pts;
      // Determine exact vs winner for tiebreaks
      const pw = predObj.t1 > predObj.t2 ? 1 : predObj.t1 < predObj.t2 ? -1 : 0;
      const rw = res.t1 > res.t2 ? 1 : res.t1 < res.t2 ? -1 : 0;
      if (predObj.t1 === res.t1 && predObj.t2 === res.t2) {
        wins++;
        if (jokerVal && pts > 0) jokers++;
      } else if (pw === rw) {
        draws++;
      }
    });

    // Pre-bets points (calculated separately when you implement prebets result checking)
    const prebets_pts = calcPrebetPoints(userPrebets, results);

    return {
      name: p.name,
      code: p.code,
      total: groupPts + prebets_pts,
      wins, draws, jokers, prebets_pts
    };
  });
}

function calcPoints(pred, res, isKO) {
  // isKO: knockout game (extra time result, no penalties)
  // Group stage: exact=3, winner=1. Knockout: exact=4, winner=2.
  // JOKER doubles the points obtained (0→0, 1→2, 2→4, 3→6, 4→8)
  if (pred.t1 === undefined || pred.t2 === undefined) return 0;
  if (res.t1 === undefined || res.t2 === undefined) return 0;
  const pw = pred.t1 > pred.t2 ? 1 : pred.t1 < pred.t2 ? -1 : 0;
  const rw = res.t1 > res.t2 ? 1 : res.t1 < res.t2 ? -1 : 0;
  const exactBase  = isKO ? 4 : 3;
  const winnerBase = isKO ? 2 : 1;
  let pts = 0;
  if (pred.t1 === res.t1 && pred.t2 === res.t2) pts = exactBase;
  else if (pw === rw) pts = winnerBase;
  return pred.joker ? pts * 2 : pts;
}

function calcPrebetPoints(userPrebets, results) {
  // Group stage prebets: checked once all group games are done
  // Special bets: checked at end of tournament
  // Returns points earned so far
  let pts = 0;
  // TODO: implement after tournament — requires knowing group standings
  // For now returns 0 until implemented
  return pts;
}

// ============================================================
//  SAVE PREDICTIONS
// ============================================================

function handleSavePreds(body) {
  const {code, preds} = body;
  if (!code || !preds) return {success: false};

  const sheet = getSheet('Predictions');
  const data = sheet.getDataRange().getValues();
  const now = new Date().toISOString();

  // Remove existing predictions for this user
  const toDelete = [];
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0] === code) toDelete.push(i + 1);
  }
  // Delete from bottom to top to preserve row indices
  toDelete.forEach(row => sheet.deleteRow(row));

  // Insert new predictions
  const rows = [];
  Object.entries(preds).forEach(([gameId, pred]) => {
    if (pred.t1 !== undefined && pred.t2 !== undefined) {
      rows.push([code, gameId, pred.t1, pred.t2, pred.joker || false, now]);
    }
  });
  if (rows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 6).setValues(rows);
  }

  return {success: true};
}

// ============================================================
//  SAVE PREBETS
// ============================================================

function handleSavePrebets(body) {
  const {code, prebets} = body;
  if (!code || !prebets) return {success: false};

  const sheet = getSheet('Prebets');
  const data = sheet.getDataRange().getValues();
  const now = new Date().toISOString();

  // Remove existing
  const toDelete = [];
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0] === code) toDelete.push(i + 1);
  }
  toDelete.forEach(row => sheet.deleteRow(row));

  // Insert new
  const rows = Object.entries(prebets)
    .filter(([k,v]) => v)
    .map(([k,v]) => [code, k, v, now]);
  if (rows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 4).setValues(rows);
  }

  return {success: true};
}

// ============================================================
//  SAVE RESULTS (admin)
// ============================================================

function handleSaveResults(body) {
  requireAdmin(body);
  const {results} = body;

  const sheet = getSheet('Results');
  // Clear existing results (keep header)
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);

  const now = new Date().toISOString();
  const rows = [];
  Object.entries(results).forEach(([gameId, res]) => {
    if (res && res.t1 !== undefined && res.t2 !== undefined) {
      rows.push([gameId, res.t1, res.t2, now]);
    }
  });

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 4).setValues(rows);
  }

  return {success: true};
}

// ============================================================
//  USER MANAGEMENT (admin)
// ============================================================

function handleGetUsers(body) {
  requireAdmin(body);
  const sheet = getSheet('Participants');
  const rows = sheetToObjects(sheet);
  return {users: rows.map(r => ({name: r.name, code: r.code}))};
}

function handleAddUser(body) {
  requireAdmin(body);
  const {name, code} = body;
  if (!name || !code) return {success: false, error: 'Missing fields'};

  const sheet = getSheet('Participants');
  const rows = sheetToObjects(sheet);
  if (rows.find(r => r.code === code.toUpperCase())) {
    return {success: false, error: 'Code already exists'};
  }

  sheet.appendRow([code.toUpperCase(), name, new Date().toISOString()]);
  return {success: true};
}

function handleRemoveUser(body) {
  requireAdmin(body);
  const {code} = body;

  // Remove from Participants
  removeRowsByCode('Participants', code);
  removeRowsByCode('Predictions', code);
  removeRowsByCode('Prebets', code);

  return {success: true};
}

function removeRowsByCode(sheetName, code) {
  const sheet = getSheet(sheetName);
  const data = sheet.getDataRange().getValues();
  const toDelete = [];
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0] === code) toDelete.push(i + 1);
  }
  toDelete.forEach(row => sheet.deleteRow(row));
}

// ============================================================
//  CONFIG (admin)
// ============================================================

function handleSaveConfig(body) {
  requireAdmin(body);
  if (body.deadlineH) setConfig('deadline_hours', body.deadlineH);
  if (body.newPassword) setConfig('admin_password', body.newPassword);
  if (body.apiKey) setConfig('football_api_key', body.apiKey);
  return {success: true};
}

// ============================================================
//  LIGAS
// ============================================================

function generateInviteCode(name) {
  const base = name.toUpperCase().replace(/[^A-Z]/g, '').substring(0, 3) || 'LIG';
  const num = Math.floor(Math.random() * 9000) + 1000;
  return base + num;
}

function getLeagueMembers(leagueId) {
  const sheet = getSheet('Liga_Members');
  const rows = sheetToObjects(sheet);
  return rows.filter(r => r.league_id === leagueId).map(r => r.user_code);
}

function handleGetMyLeagues(body) {
  const {code} = body;
  const ligasSheet = getSheet('Ligas');
  const membersSheet = getSheet('Liga_Members');
  const allLeagues = sheetToObjects(ligasSheet);
  const allMembers = sheetToObjects(membersSheet);
  const myLeagueIds = allMembers.filter(m => m.user_code === code.toUpperCase()).map(m => m.league_id);
  const myLeagues = allLeagues
    .filter(l => myLeagueIds.includes(l.id))
    .map(l => ({
      id: l.id,
      name: l.name,
      invite: l.invite_code,
      admin: l.admin_code,
      members: allMembers.filter(m => m.league_id === l.id).map(m => m.user_code)
    }));
  return {leagues: myLeagues};
}

function handleGetLeagues(body) {
  requireAdmin(body);
  const ligasSheet = getSheet('Ligas');
  const membersSheet = getSheet('Liga_Members');
  const allLeagues = sheetToObjects(ligasSheet);
  const allMembers = sheetToObjects(membersSheet);
  const leagues = allLeagues.map(l => ({
    id: l.id,
    name: l.name,
    invite: l.invite_code,
    admin: l.admin_code,
    members: allMembers.filter(m => m.league_id === l.id).map(m => m.user_code)
  }));
  return {leagues};
}

function handleCreateLeague(body) {
  const {name, code} = body;
  if (!name || !code) return {success: false, error: 'Missing fields'};

  const ligasSheet = getSheet('Ligas');
  const membersSheet = getSheet('Liga_Members');
  const now = new Date().toISOString();

  // Generate unique ID and invite code
  const allLeagues = sheetToObjects(ligasSheet);
  const id = 'L' + (allLeagues.length + 1).toString().padStart(4, '0');
  let invite = generateInviteCode(name);
  // Ensure invite is unique
  while (allLeagues.find(l => l.invite_code === invite)) {
    invite = generateInviteCode(name);
  }

  const adminCode = code.toUpperCase();
  ligasSheet.appendRow([id, name, adminCode, invite, now]);
  membersSheet.appendRow([id, adminCode, now]);

  return {success: true, league: {id, name, invite, admin: adminCode, members: [adminCode]}};
}

function handleJoinLeague(body) {
  const {code, invite_code} = body;
  if (!code || !invite_code) return {success: false, error: 'Missing fields'};

  const ligasSheet = getSheet('Ligas');
  const membersSheet = getSheet('Liga_Members');
  const leagues = sheetToObjects(ligasSheet);
  const league = leagues.find(l => l.invite_code === invite_code.toUpperCase());

  if (!league) return {success: false, error: 'Código de convite inválido'};

  const userCode = code.toUpperCase();
  const allMembers = sheetToObjects(membersSheet);
  const alreadyMember = allMembers.find(m => m.league_id === league.id && m.user_code === userCode);
  if (!alreadyMember) {
    membersSheet.appendRow([league.id, userCode, new Date().toISOString()]);
  }

  const members = allMembers.filter(m => m.league_id === league.id).map(m => m.user_code);
  if (!members.includes(userCode)) members.push(userCode);

  return {success: true, league: {id: league.id, name: league.name, invite: league.invite_code, admin: league.admin_code, members}};
}

function handleLeaveLeague(body) {
  const {code, league_id} = body;
  const sheet = getSheet('Liga_Members');
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0] === league_id && data[i][1] === code.toUpperCase()) {
      sheet.deleteRow(i + 1);
    }
  }
  return {success: true};
}

function handleRemoveLeagueMember(body) {
  requireAdmin(body);
  const {league_id, member_code} = body;
  const sheet = getSheet('Liga_Members');
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][0] === league_id && data[i][1] === member_code) {
      sheet.deleteRow(i + 1);
    }
  }
  return {success: true};
}

function handleDeleteLeague(body) {
  requireAdmin(body);
  const {league_id} = body;

  const ligasSheet = getSheet('Ligas');
  const lData = ligasSheet.getDataRange().getValues();
  for (let i = lData.length - 1; i >= 1; i--) {
    if (lData[i][0] === league_id) ligasSheet.deleteRow(i + 1);
  }

  const membersSheet = getSheet('Liga_Members');
  const mData = membersSheet.getDataRange().getValues();
  for (let i = mData.length - 1; i >= 1; i--) {
    if (mData[i][0] === league_id) membersSheet.deleteRow(i + 1);
  }

  return {success: true};
}

// ============================================================
//  AUTO-FETCH RESULTADOS — football-data.org API
// ============================================================

// Mapeamento de nomes em inglês (API) → nomes em português (nosso sistema)
const TEAM_NAME_EN_PT = {
  'Mexico': 'México',
  'South Africa': 'África do Sul',
  'Korea Republic': 'Coreia do Sul',
  'South Korea': 'Coreia do Sul',
  'Czech Republic': 'Chéquia',
  'Czechia': 'Chéquia',
  'Canada': 'Canadá',
  'Bosnia and Herzegovina': 'Bósnia e Herzegovina',
  'Qatar': 'Qatar',
  'Switzerland': 'Suíça',
  'Brazil': 'Brasil',
  'Morocco': 'Marrocos',
  'Haiti': 'Haiti',
  'Scotland': 'Escócia',
  'United States': 'EUA',
  'USA': 'EUA',
  'Paraguay': 'Paraguai',
  'Australia': 'Austrália',
  'Turkey': 'Turquia',
  'Türkiye': 'Turquia',
  'Turkiye': 'Turquia',
  'Germany': 'Alemanha',
  'Curaçao': 'Curaçao',
  'Curacao': 'Curaçao',
  "Côte d'Ivoire": 'Costa do Marfim',
  'Ivory Coast': 'Costa do Marfim',
  'Ecuador': 'Equador',
  'Netherlands': 'Holanda',
  'Holland': 'Holanda',
  'Japan': 'Japão',
  'Sweden': 'Suécia',
  'Tunisia': 'Tunísia',
  'Spain': 'Espanha',
  'Cape Verde': 'Cabo Verde',
  'Saudi Arabia': 'Arábia Saudita',
  'Uruguay': 'Uruguai',
  'Belgium': 'Bélgica',
  'Egypt': 'Egipto',
  'Iran': 'Irão',
  'New Zealand': 'Nova Zelândia',
  'France': 'França',
  'Senegal': 'Senegal',
  'Iraq': 'Iraque',
  'Norway': 'Noruega',
  'Argentina': 'Argentina',
  'Algeria': 'Argélia',
  'Austria': 'Áustria',
  'Jordan': 'Jordânia',
  'Portugal': 'Portugal',
  'DR Congo': 'R.D. Congo',
  'Democratic Republic of Congo': 'R.D. Congo',
  'Congo DR': 'R.D. Congo',
  'Uzbekistan': 'Uzbequistão',
  'Colombia': 'Colômbia',
  'England': 'Inglaterra',
  'Croatia': 'Croácia',
  'Ghana': 'Gana',
  'Panama': 'Panamá',
};

// Lista de jogos da fase de grupos para matching por equipas
// Formato: {id, t1, t2}
function getGroupGamesList() {
  return [
    {id:'A1',t1:'México',t2:'África do Sul'},
    {id:'A2',t1:'Coreia do Sul',t2:'Chéquia'},
    {id:'B1',t1:'Canadá',t2:'Bósnia e Herzegovina'},
    {id:'D1',t1:'EUA',t2:'Paraguai'},
    {id:'B2',t1:'Qatar',t2:'Suíça'},
    {id:'C1',t1:'Brasil',t2:'Marrocos'},
    {id:'C2',t1:'Haiti',t2:'Escócia'},
    {id:'D2',t1:'Austrália',t2:'Turquia'},
    {id:'E1',t1:'Alemanha',t2:'Curaçao'},
    {id:'F1',t1:'Holanda',t2:'Japão'},
    {id:'E2',t1:'Costa do Marfim',t2:'Equador'},
    {id:'F2',t1:'Suécia',t2:'Tunísia'},
    {id:'H1',t1:'Espanha',t2:'Cabo Verde'},
    {id:'G1',t1:'Bélgica',t2:'Egipto'},
    {id:'H2',t1:'Arábia Saudita',t2:'Uruguai'},
    {id:'G2',t1:'Irão',t2:'Nova Zelândia'},
    {id:'I1',t1:'França',t2:'Senegal'},
    {id:'I2',t1:'Iraque',t2:'Noruega'},
    {id:'J1',t1:'Argentina',t2:'Argélia'},
    {id:'J2',t1:'Áustria',t2:'Jordânia'},
    {id:'K1',t1:'Portugal',t2:'R.D. Congo'},
    {id:'L1',t1:'Inglaterra',t2:'Croácia'},
    {id:'L2',t1:'Gana',t2:'Panamá'},
    {id:'K2',t1:'Uzbequistão',t2:'Colômbia'},
    {id:'A3',t1:'Chéquia',t2:'África do Sul'},
    {id:'B3',t1:'Suíça',t2:'Bósnia e Herzegovina'},
    {id:'B4',t1:'Canadá',t2:'Qatar'},
    {id:'A4',t1:'México',t2:'Coreia do Sul'},
    {id:'D3',t1:'EUA',t2:'Austrália'},
    {id:'C3',t1:'Escócia',t2:'Marrocos'},
    {id:'C4',t1:'Brasil',t2:'Haiti'},
    {id:'D4',t1:'Turquia',t2:'Paraguai'},
    {id:'F3',t1:'Holanda',t2:'Suécia'},
    {id:'E3',t1:'Alemanha',t2:'Costa do Marfim'},
    {id:'E4',t1:'Equador',t2:'Curaçao'},
    {id:'F4',t1:'Tunísia',t2:'Japão'},
    {id:'H3',t1:'Espanha',t2:'Arábia Saudita'},
    {id:'G3',t1:'Bélgica',t2:'Irão'},
    {id:'H4',t1:'Uruguai',t2:'Cabo Verde'},
    {id:'G4',t1:'Nova Zelândia',t2:'Egipto'},
    {id:'J3',t1:'Argentina',t2:'Áustria'},
    {id:'I3',t1:'França',t2:'Iraque'},
    {id:'I4',t1:'Noruega',t2:'Senegal'},
    {id:'J4',t1:'Jordânia',t2:'Argélia'},
    {id:'K3',t1:'Portugal',t2:'Uzbequistão'},
    {id:'L3',t1:'Inglaterra',t2:'Gana'},
    {id:'L4',t1:'Panamá',t2:'Croácia'},
    {id:'K4',t1:'Colômbia',t2:'R.D. Congo'},
    {id:'B5',t1:'Suíça',t2:'Canadá'},
    {id:'B6',t1:'Bósnia e Herzegovina',t2:'Qatar'},
    {id:'C5',t1:'Escócia',t2:'Brasil'},
    {id:'C6',t1:'Marrocos',t2:'Haiti'},
    {id:'A5',t1:'Chéquia',t2:'México'},
    {id:'A6',t1:'África do Sul',t2:'Coreia do Sul'},
    {id:'E5',t1:'Equador',t2:'Alemanha'},
    {id:'E6',t1:'Curaçao',t2:'Costa do Marfim'},
    {id:'F5',t1:'Japão',t2:'Suécia'},
    {id:'F6',t1:'Tunísia',t2:'Holanda'},
    {id:'D5',t1:'Turquia',t2:'EUA'},
    {id:'D6',t1:'Paraguai',t2:'Austrália'},
    {id:'I5',t1:'Noruega',t2:'França'},
    {id:'I6',t1:'Senegal',t2:'Iraque'},
    {id:'H5',t1:'Cabo Verde',t2:'Arábia Saudita'},
    {id:'H6',t1:'Uruguai',t2:'Espanha'},
    {id:'G5',t1:'Egipto',t2:'Irão'},
    {id:'G6',t1:'Nova Zelândia',t2:'Bélgica'},
    {id:'L5',t1:'Panamá',t2:'Inglaterra'},
    {id:'L6',t1:'Croácia',t2:'Gana'},
    {id:'K5',t1:'Colômbia',t2:'Portugal'},
    {id:'K6',t1:'R.D. Congo',t2:'Uzbequistão'},
    {id:'J5',t1:'Argélia',t2:'Áustria'},
    {id:'J6',t1:'Jordânia',t2:'Argentina'},
  ];
}

function findGameIdByTeams(apiHome, apiAway) {
  const t1 = TEAM_NAME_EN_PT[apiHome] || apiHome;
  const t2 = TEAM_NAME_EN_PT[apiAway] || apiAway;
  const games = getGroupGamesList();
  const match = games.find(g =>
    (g.t1 === t1 && g.t2 === t2) || (g.t1 === t2 && g.t2 === t1)
  );
  return match ? {id: match.id, swapped: match.t1 === t2} : null;
}

function handleFetchResults(body) {
  requireAdmin(body);
  const apiKey = getConfig('football_api_key', '');
  if (!apiKey) {
    return {success: false, error: 'API key não configurada. Vai a Admin → Config e adiciona a tua chave da football-data.org.'};
  }

  try {
    // Fetch finished WC 2026 matches from football-data.org
    const url = 'https://api.football-data.org/v4/competitions/WC/matches?status=FINISHED';
    const response = UrlFetchApp.fetch(url, {
      headers: {'X-Auth-Token': apiKey},
      muteHttpExceptions: true
    });

    const code = response.getResponseCode();
    if (code === 403) return {success: false, error: 'API key inválida ou sem permissão para o Mundial.'};
    if (code === 429) return {success: false, error: 'Limite de pedidos excedido. Tenta mais tarde.'};
    if (code !== 200) return {success: false, error: 'Erro na API (HTTP ' + code + ').'};

    const data = JSON.parse(response.getContentText());
    const matches = data.matches || [];

    if (!matches.length) {
      return {success: true, updated: 0, message: 'Nenhum jogo terminado encontrado na API.'};
    }

    const sheet = getSheet('Results');
    // Load existing results into a map for fast lookup
    const existingData = sheet.getDataRange().getValues();
    const existingMap = {};
    for (let i = 1; i < existingData.length; i++) {
      existingMap[existingData[i][0]] = i + 1; // game_id → row number
    }

    const now = new Date().toISOString();
    let updated = 0;
    const skipped = [];

    matches.forEach(match => {
      const score = match.score;
      if (!score || !score.fullTime) return;
      // Use fullTime score (includes extra time if played; penalties NOT included)
      const t1 = score.fullTime.home;
      const t2 = score.fullTime.away;
      if (t1 === null || t2 === null) return;

      const homeTeam = match.homeTeam.shortName || match.homeTeam.name;
      const awayTeam = match.awayTeam.shortName || match.awayTeam.name;
      const result = findGameIdByTeams(homeTeam, awayTeam);

      if (!result) {
        skipped.push(homeTeam + ' vs ' + awayTeam);
        return;
      }

      const gameId = result.id;
      const finalT1 = result.swapped ? t2 : t1;
      const finalT2 = result.swapped ? t1 : t2;

      if (existingMap[gameId]) {
        // Update existing row
        sheet.getRange(existingMap[gameId], 2, 1, 3).setValues([[finalT1, finalT2, now]]);
      } else {
        // Append new row
        sheet.appendRow([gameId, finalT1, finalT2, now]);
      }
      updated++;
    });

    // Save fetch timestamp
    setConfig('last_auto_fetch', now);

    const msg = skipped.length
      ? skipped.length + ' jogo(s) não identificados: ' + skipped.slice(0, 3).join(', ')
      : '';

    return {success: true, updated, message: msg};

  } catch (err) {
    return {success: false, error: 'Erro ao aceder à API: ' + err.toString()};
  }
}

// ── Trigger automático (opcional — configurar manualmente no Apps Script) ──
// Para ativar: no editor do Apps Script, vai a Triggers → Add Trigger
// Função: autoFetchTrigger | Tipo: Time-driven | Interval: Every 5 minutes
function autoFetchTrigger() {
  const apiKey = getConfig('football_api_key', '');
  if (!apiKey) return;
  try {
    handleFetchResults({password: getConfig('admin_password', ADMIN_PASSWORD_DEFAULT)});
  } catch(e) {
    Logger.log('autoFetchTrigger error: ' + e.toString());
  }
}
