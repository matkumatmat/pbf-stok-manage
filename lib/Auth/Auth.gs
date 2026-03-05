// ============================================================
// AUTH MODULE — Session Management & Login
// ============================================================
// Note: SH_AUTH is defined in Kode.js global config

var AUTH_USERNAME = 0;  // Column A
var AUTH_PASSWD   = 1;  // Column B
var AUTH_ROLE     = 2;  // Column C (AS)

// Session key prefix
var SESSION_PREFIX = 'auth_session_';

/**
 * Generate a random session token
 */
function generateSessionToken() {
  var chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  var token = '';
  for (var i = 0; i < 32; i++) {
    token += chars.charAt(Math.floor(Math.random() * chars.length));
  }
  return token + '_' + new Date().getTime();
}

/**
 * Authenticate user credentials
 * Returns: { success: true, username: 'X', role: 'Y', token: 'Z' } OR { success: false, error: 'msg' }
 */
function authenticateUser(username, password) {
  try {
    if (!username || !password) {
      return { success: false, error: 'Username dan password harus diisi' };
    }

    var ss = SpreadsheetApp.openById(SS_ID);
    var authSheet = ss.getSheetByName(SH_AUTH);

    if (!authSheet) {
      Logger.log('ERROR: Auth_ sheet not found');
      return { success: false, error: 'Konfigurasi sistem tidak valid' };
    }

    var data = authSheet.getDataRange().getValues();

    // Skip header row (row 0), iterate from row 1
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var dbUser = str(row[AUTH_USERNAME]).toUpperCase();
      var dbPass = str(row[AUTH_PASSWD]);
      var dbRole = str(row[AUTH_ROLE]);

      if (!dbUser) continue;

      if (dbUser === username.toUpperCase() && dbPass === password) {
        // Match found - create session
        var token = generateSessionToken();
        var userProps = PropertiesService.getUserProperties();

        // Store session data
        userProps.setProperty(SESSION_PREFIX + 'token', token);
        userProps.setProperty(SESSION_PREFIX + 'username', dbUser);
        userProps.setProperty(SESSION_PREFIX + 'role', dbRole);

        return {
          success: true,
          username: dbUser,
          role: dbRole,
          token: token
        };
      }
    }

    return { success: false, error: 'Username atau password salah' };
  } catch (e) {
    Logger.log('ERROR in authenticateUser: ' + e.message);
    return { success: false, error: 'Terjadi kesalahan sistem' };
  }
}

/**
 * Check if current user has valid session
 * Returns: { authenticated: true, username: 'X', role: 'Y' } OR { authenticated: false }
 */
function checkSession() {
  try {
    var userProps = PropertiesService.getUserProperties();
    var token = userProps.getProperty(SESSION_PREFIX + 'token');

    if (!token) {
      return { authenticated: false };
    }

    var username = userProps.getProperty(SESSION_PREFIX + 'username');
    var role = userProps.getProperty(SESSION_PREFIX + 'role');

    return {
      authenticated: true,
      username: username || '',
      role: role || ''
    };
  } catch (e) {
    Logger.log('ERROR in checkSession: ' + e.message);
    return { authenticated: false };
  }
}

/**
 * Logout - clear session
 */
function logout() {
  try {
    var userProps = PropertiesService.getUserProperties();
    userProps.deleteProperty(SESSION_PREFIX + 'token');
    userProps.deleteProperty(SESSION_PREFIX + 'username');
    userProps.deleteProperty(SESSION_PREFIX + 'role');

    return { success: true };
  } catch (e) {
    Logger.log('ERROR in logout: ' + e.message);
    return { success: false, error: 'Gagal logout' };
  }
}
