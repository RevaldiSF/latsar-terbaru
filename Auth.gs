var SESSION_DURATION = 30 * 60 * 1000;

function login(username, password) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('master_user');
  var data = sheet.getDataRange().getValues();
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (String(row[1]).trim() === String(username).trim() && 
        String(row[2]).trim() === String(password).trim()) {
      var token = Utilities.getUuid();
      var expiry = new Date().getTime() + SESSION_DURATION;
      var cache = CacheService.getScriptCache();
      
      var role = String(row[4]).trim();
      var terminal = String(row[5]).trim();
      
      // Semua role mulai dari dashboard
      var homePage = 'dashboard';
      
      var sessionData = {
        user_id: String(row[0]),
        username: String(row[1]).trim(),
        nama: String(row[3]).trim(),
        role: role,
        terminal: terminal,
        expiry: expiry,
        homePage: homePage
      };
      
      cache.put(token, JSON.stringify(sessionData), 1800);
      
      return { 
        success: true, 
        token: token, 
        role: role, 
        terminal: terminal, 
        nama: String(row[3]).trim(),
        homePage: homePage
      };
    }
  }
  return { success: false, message: 'Username atau password salah!' };
}

function checkSession(token) {
  if (!token) return null;
  var cache = CacheService.getScriptCache();
  var data = cache.get(token);
  if (!data) return null;
  var session = JSON.parse(data);
  var now = new Date().getTime();
  if (now > session.expiry) {
    cache.remove(token);
    return null;
  }
  session.expiry = now + SESSION_DURATION;
  cache.put(token, JSON.stringify(session), 1800);
  return session;
}

function logout(token) {
  if (token) {
    var cache = CacheService.getScriptCache();
    cache.remove(token);
  }
  return { success: true };
}

function validateSession(token) {
  return checkSession(token);
}

function loginTamu() {
  var token = 'tamu_' + Utilities.getUuid();
  var expiry = new Date().getTime() + SESSION_DURATION;
  var cache = CacheService.getScriptCache();
  cache.put(token, JSON.stringify({
    user_id: 'TAMU',
    username: 'tamu',
    nama: 'Tamu',
    role: 'tamu',
    terminal: 'Semua',
    expiry: expiry,
    homePage: 'dashboard'
  }), 1800);
  return { success: true, token: token, role: 'tamu', terminal: 'Semua', nama: 'Tamu', homePage: 'dashboard' };
}
