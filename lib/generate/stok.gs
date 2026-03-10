// ============================================================
// STOCK PIVOT EXPORT MODULE
// ============================================================
// This file is automatically included in GAS project scope
// Functions here are callable from Kode.js router

/**
 * Export all stock aggregated by product code + batch
 * Returns array of items with all required fields for pivot export
 */
function getStockPivotExport() {
  try {
    var ss = SpreadsheetApp.openById(SS_ID);
    var shProduk = ss.getSheetByName(SH_PRODUK);
    var produkData = shProduk.getDataRange().getValues();

    var shMaster = ss.getSheetByName(SH_MASTER);
    var masterData = shMaster.getDataRange().getValues();

    // Group by kode + batch
    var pivotMap = {}; // key: kode|batch

    for (var i = 3; i < masterData.length; i++) {
      var row = masterData[i];
      var kode = str(row[MD_KODE]);
      var batch = str(row[MD_BATCH]);
      if (!kode || !batch) continue;

      var key = kode + '|' + batch;
      if (!pivotMap[key]) {
        pivotMap[key] = {
          kategori:    str(row[MD_KATEGORI]),
          kodeBarang:  kode,
          namaBarang:  str(row[MD_NAMA_ERP]),
          namaDagang:  str(row[MD_NAMA_DGG]),
          batch:       batch,
          mfg:         fmtDate(row[MD_MFG]),
          expire:      fmtDate(row[MD_EXPIRE]),
          expireSort:  row[MD_EXPIRE] instanceof Date ? row[MD_EXPIRE].getTime() : 0,
          rsl:         num(row[MD_RSL]),
          penerimaan:  0,
          distribusi:  0
        };
      }
      pivotMap[key].penerimaan += num(row[MD_PEN]);
      pivotMap[key].distribusi += num(row[MD_DIST]);
    }

    // Convert to array and calculate stok
    var results = [];
    for (var k in pivotMap) {
      var item = pivotMap[k];
      item.stok = item.penerimaan - item.distribusi;
      results.push(item);
    }

    // Sort by kategori, then kode, then expire date
    results.sort(function(a, b) {
      var cmpKat = a.kategori.localeCompare(b.kategori);
      if (cmpKat !== 0) return cmpKat;
      var cmpKode = a.kodeBarang.localeCompare(b.kodeBarang);
      if (cmpKode !== 0) return cmpKode;
      return a.expireSort - b.expireSort;
    });

    // Remove expireSort (internal field)
    results.forEach(function(r) { delete r.expireSort; });

    return { success: true, items: results };
  } catch(e) {
    return { success: false, error: e.message };
  }
}
