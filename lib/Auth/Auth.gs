// ============================================================
// AUTH MODULE — Client-side Authentication (localStorage-based)
// ============================================================
// Note: SH_AUTH is defined in Kode.js global config

var AUTH_USERNAME = 0;  // Column A
var AUTH_PASSWD   = 1;  // Column B
var AUTH_ROLE     = 2;  // Column C (AS)

/**
 * Authenticate user credentials (client-side session via localStorage)
 * Returns: { success: true, username: 'X', role: 'Y' } OR { success: false, error: 'msg' }
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
        // Match found - return user info (client will store in localStorage)
        return {
          success: true,
          username: dbUser,
          role: dbRole
        };
      }
    }

    return { success: false, error: 'Username atau password salah' };
  } catch (e) {
    Logger.log('ERROR in authenticateUser: ' + e.message);
    return { success: false, error: 'Terjadi kesalahan sistem' };
  }
}
