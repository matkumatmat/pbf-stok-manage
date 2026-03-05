// ============================================================
// STOK PBF — Google Apps Script Backend (v3 - FIXED INDICES)
// Deploy: Web App → Execute as me → Anyone with link
// ============================================================

var SS_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
var SH_PRODUK = 'PRODUK';
var SH_MASTER = 'MASTER DISTRIBUSI';

// ── PRODUK column indices (col B "NAMA BARANG" sudah di-drop)
// A=0  B=1       C=2          D=3       E=4       F=5    G=6  H=7
// KODE KODE_NEW  NAMA_DAGANG  NAMA_ERP  KATEGORI  JENIS  HJP  HET
var PR_KODE_LAMA   = 0;
var PR_KODE_BARU   = 1;
var PR_NAMA_DAGANG = 2;
var PR_NAMA_ERP    = 3;
var PR_KATEGORI    = 4;
var PR_JENIS       = 5;
var PR_HJP         = 6;
var PR_HET         = 7;

// ── MASTER DISTRIBUSI column indices (row 3 = header, row 4+ = data)
// SEKTOR column sudah DIHAPUS, MFG_DATE sudah DITAMBAH
//
// A=0         B=1       C=2     D=3        E=4            F=5
// TGL_KIRIM   NO_DOK    NO_PO   NO_SO/MOV  NAMA_KONSUMEN  KOTA/CABANG
//
// G=6        H=7          I=8             J=9          K=10
// KATEGORI   KODE BARANG  NAMA BARANG(ERP) NAMA DAGANG  BATCH
//
// L=11       M=12         N=13    O=14        P=15        Q=16        R=17        S=18
// MFG_DATE   EXPIRE_DATE  STATUS  RSL_(BULAN) PENERIMAAN  DISTRIBUSI  KETERANGAN  CATATAN

var MD_TGL       = 0;
var MD_NO_DOK    = 1;
var MD_NO_PO     = 2;
var MD_SO        = 3;
var MD_KONSUMEN  = 4;
var MD_KOTA      = 5;
var MD_KATEGORI  = 6;
var MD_KODE      = 7;
var MD_NAMA_ERP  = 8;
var MD_NAMA_DGG  = 9;
var MD_BATCH     = 10;
var MD_MFG       = 11;
var MD_EXPIRE    = 12;
var MD_STATUS    = 13;
var MD_RSL       = 14;
var MD_PEN       = 15;
var MD_DIST      = 16;
var MD_KET       = 17;

// ═══════════════════════════════════════════════════════════
// WEB APP ENTRY
// ═══════════════════════════════════════════════════════════

function doGet() {
  return HtmlService
    .createHtmlOutputFromFile('index')
    .setTitle('Stok PBF')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ═══════════════════════════════════════════════════════════
// ROUTER
// ═══════════════════════════════════════════════════════════

function handleRequest(action, params) {
  switch (action) {
    case 'searchProduct':      return searchProduct(params.query);
    case 'getProductList':     return getProductList();
    case 'getProductDetail':   return getProductDetail(params.kodeBarang);
    case 'getPengiriman':      return getPengiriman(params.startDate, params.endDate, params.konsumen, params.kota);
    case 'getCycleCount':      return getCycleCount(params.tanggal);
    case 'getDistributorList': return getDistributorList();
    case 'searchBatch':        return searchBatch(params.query);
    case 'getBatchDetail':     return getBatchDetail(params.kodeBarang, params.batch);
    // Embalage
    case 'getEmbalageList':      return getEmbalageList();
    case 'getEmbalagePenerimaan': return getEmbalagePenerimaan();
    case 'addEmbalagePenerimaan': return addEmbalagePenerimaan(params);
    case 'getEmbalageDashboard':  return getEmbalageDashboard();
    case 'getEmbalagePengeluaran': return getEmbalagePengeluaran(params.page, params.pageSize);
    case 'getDistribusiHistory':  return getDistribusiHistory(params.page, params.pageSize);
    case 'getDistribusiDetail':   return getDistribusiDetail(params.tanggal, params.tujuan);
    default: return { error: 'Unknown action: ' + action };
  }
}

// ═══════════════════════════════════════════════════════════
// TAB 1 — SEARCH PRODUK (by nama/kode OR by batch)
// ═══════════════════════════════════════════════════════════

function searchProduct(query) {
  var q = (query || '').toString().toLowerCase().trim();
  if (!q) return { results: [] };

  var ss = SpreadsheetApp.openById(SS_ID);

  var shProduk = ss.getSheetByName(SH_PRODUK);
  var produkData = shProduk.getDataRange().getValues();
  var results = [];
  var seenKodes = {};

  for (var i = 1; i < produkData.length; i++) {
    var row = produkData[i];
    if (!row[PR_KODE_BARU]) continue;

    var haystack = [str(row[PR_KODE_LAMA]), str(row[PR_KODE_BARU]),
                    str(row[PR_NAMA_DAGANG]), str(row[PR_NAMA_ERP])]
                   .join(' ').toLowerCase();

    if (haystack.indexOf(q) !== -1) {
      var kode = str(row[PR_KODE_BARU]);
      if (!seenKodes[kode]) {
        seenKodes[kode] = true;
        results.push({
          kodeLama:   str(row[PR_KODE_LAMA]),
          kodeBaru:   kode,
          namaDagang: str(row[PR_NAMA_DAGANG]),
          namaERP:    str(row[PR_NAMA_ERP]),
          kategori:   str(row[PR_KATEGORI]),
          jenis:      str(row[PR_JENIS]),
          hjp:        num(row[PR_HJP]),
          het:        num(row[PR_HET])
        });
      }
    }
  }

  // Search MASTER DISTRIBUSI by batch number
  var shMaster = ss.getSheetByName(SH_MASTER);
  var masterData = shMaster.getDataRange().getValues();

  for (var i = 3; i < masterData.length; i++) {
    var row = masterData[i];
    var batch = str(row[MD_BATCH]).toLowerCase();
    var kode = str(row[MD_KODE]);
    if (!kode || !batch) continue;
    if (seenKodes[kode]) continue;

    if (batch.indexOf(q) !== -1) {
      seenKodes[kode] = true;
      var master = lookupProduk(produkData, kode);
      results.push({
        kodeLama:     master ? master.kodeLama : '',
        kodeBaru:     kode,
        namaDagang:   master ? master.namaDagang : str(row[MD_NAMA_DGG]),
        namaERP:      master ? master.namaERP : str(row[MD_NAMA_ERP]),
        kategori:     master ? master.kategori : '',
        jenis:        master ? master.jenis : '',
        hjp:          master ? master.hjp : 0,
        het:          master ? master.het : 0,
        matchedBatch: str(row[MD_BATCH])
      });
    }
  }

  return { results: results };
}

/**
 * Return ALL products from PRODUK sheet (for initial load)
 */
function getProductList() {
  var sh = SpreadsheetApp.openById(SS_ID).getSheetByName(SH_PRODUK);
  var data = sh.getDataRange().getValues();
  var results = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[PR_KODE_BARU]) continue;
    results.push({
      kodeLama:   str(row[PR_KODE_LAMA]),
      kodeBaru:   str(row[PR_KODE_BARU]),
      namaDagang: str(row[PR_NAMA_DAGANG]),
      namaERP:    str(row[PR_NAMA_ERP]),
      kategori:   str(row[PR_KATEGORI]),
      jenis:      str(row[PR_JENIS]),
      hjp:        num(row[PR_HJP]),
      het:        num(row[PR_HET])
    });
  }

  return { results: results };
}

function lookupProduk(produkData, kodeBarang) {
  for (var i = 1; i < produkData.length; i++) {
    if (str(produkData[i][PR_KODE_BARU]) === kodeBarang) {
      return {
        kodeLama:   str(produkData[i][PR_KODE_LAMA]),
        kodeBaru:   kodeBarang,
        namaDagang: str(produkData[i][PR_NAMA_DAGANG]),
        namaERP:    str(produkData[i][PR_NAMA_ERP]),
        kategori:   str(produkData[i][PR_KATEGORI]),
        jenis:      str(produkData[i][PR_JENIS]),
        hjp:        num(produkData[i][PR_HJP]),
        het:        num(produkData[i][PR_HET])
      };
    }
  }
  return null;
}

// ═══════════════════════════════════════════════════════════
// TAB 1 — PRODUCT DETAIL
// ═══════════════════════════════════════════════════════════

function getProductDetail(kodeBarang) {
  if (!kodeBarang) return { error: 'kodeBarang required' };

  var ss = SpreadsheetApp.openById(SS_ID);
  var shProduk = ss.getSheetByName(SH_PRODUK);
  var produkData = shProduk.getDataRange().getValues();
  var master = lookupProduk(produkData, kodeBarang);

  var shMaster = ss.getSheetByName(SH_MASTER);
  var data = shMaster.getDataRange().getValues();

  var totalPen = 0, totalDist = 0;
  var batches = {};
  var konsumenMap = {};

  for (var i = 3; i < data.length; i++) {
    var row = data[i];
    if (str(row[MD_KODE]) !== kodeBarang) continue;

    var batch    = str(row[MD_BATCH]);
    var expire   = row[MD_EXPIRE];
    var pen      = num(row[MD_PEN]);
    var dist     = num(row[MD_DIST]);
    var status   = str(row[MD_STATUS]);
    var konsumen = str(row[MD_KONSUMEN]);

    totalPen  += pen;
    totalDist += dist;

    if (batch) {
      if (!batches[batch]) {
        batches[batch] = {
          batch:      batch,
          mfg:        fmtDate(row[MD_MFG]),
          expire:     fmtDate(expire),
          expireSort: expire instanceof Date ? expire.getTime() : 0,
          penerimaan: 0,
          distribusi: 0,
          status:     status
        };
      }
      batches[batch].penerimaan += pen;
      batches[batch].distribusi += dist;
      if (status) batches[batch].status = status;
    }

    if (konsumen) {
      if (!konsumenMap[konsumen]) {
        konsumenMap[konsumen] = { konsumen: konsumen, masuk: 0, keluar: 0 };
      }
      konsumenMap[konsumen].masuk  += pen;
      konsumenMap[konsumen].keluar += dist;
    }
  }

  var batchList = [];
  for (var b in batches) {
    var item = batches[b];
    item.stok = item.penerimaan - item.distribusi;
    batchList.push(item);
  }
  batchList.sort(function(a, b) { return a.expireSort - b.expireSort; });
  batchList.forEach(function(x) { delete x.expireSort; });

  var riwayat = [];
  for (var k in konsumenMap) {
    var entry = konsumenMap[k];
    entry.total = entry.masuk - entry.keluar;
    riwayat.push(entry);
  }
  riwayat.sort(function(a, b) {
    return (b.masuk + b.keluar) - (a.masuk + a.keluar);
  });

  return {
    kodeBarang:      kodeBarang,
    kodeLama:        master ? master.kodeLama : '',
    namaDagang:      master ? master.namaDagang : '',
    namaERP:         master ? master.namaERP : '',
    kategori:        master ? master.kategori : '',
    jenis:           master ? master.jenis : '',
    hjp:             master ? master.hjp : 0,
    het:             master ? master.het : 0,
    totalPenerimaan: totalPen,
    totalDistribusi: totalDist,
    totalStok:       totalPen - totalDist,
    batches:         batchList,
    riwayat:         riwayat
  };
}

// ═══════════════════════════════════════════════════════════
// DISTRIBUTOR LIST (cascading dropdowns)
// ═══════════════════════════════════════════════════════════

function getDistributorList() {
  var sh = SpreadsheetApp.openById(SS_ID).getSheetByName('DISTRIBUTOR');
  var data = sh.getDataRange().getValues();
  var map = {};

  for (var i = 2; i < data.length; i++) {
    var konsumen = (data[i][1] || '').toString().trim();
    var kota     = (data[i][2] || '').toString().trim();
    if (!konsumen || konsumen === 'NAMA KONSUMEN') continue;

    if (!map[konsumen]) map[konsumen] = [];
    if (kota && map[konsumen].indexOf(kota) === -1) {
      map[konsumen].push(kota);
    }
  }

  for (var k in map) { map[k].sort(); }

  return { distributor: map };
}

// ═══════════════════════════════════════════════════════════
// TAB 2 — PENGIRIMAN
// ═══════════════════════════════════════════════════════════

function getPengiriman(startDate, endDate, konsumenFilter, kotaFilter) {
  var sh = SpreadsheetApp.openById(SS_ID).getSheetByName(SH_MASTER);
  var data = sh.getDataRange().getValues();

  var start = startDate ? new Date(startDate) : null;
  var end   = endDate   ? new Date(endDate)   : null;
  if (start) start.setHours(0, 0, 0, 0);
  if (end)   end.setHours(23, 59, 59, 999);

  var fKonsumen = (konsumenFilter || '').trim();
  var fKota     = (kotaFilter || '').trim();
  var results = [];

  for (var i = 3; i < data.length; i++) {
    var row = data[i];
    var dist = num(row[MD_DIST]);
    if (dist <= 0) continue;

    var tgl = row[MD_TGL];
    if (!(tgl instanceof Date)) continue;
    if (start && tgl < start) continue;
    if (end   && tgl > end)   continue;

    if (fKonsumen && str(row[MD_KONSUMEN]) !== fKonsumen) continue;
    if (fKota     && str(row[MD_KOTA]) !== fKota)         continue;

    results.push({
      tanggal:    fmtDate(tgl),
      noDokumen:  str(row[MD_NO_DOK]),
      konsumen:   str(row[MD_KONSUMEN]),
      kota:       str(row[MD_KOTA]),
      kodeBarang: str(row[MD_KODE]),
      namaDagang: str(row[MD_NAMA_DGG]),
      batch:      str(row[MD_BATCH]),
      mfg:        fmtDate(row[MD_MFG]),
      expire:     fmtDate(row[MD_EXPIRE]),
      qty:        dist,
      keterangan: str(row[MD_KET])
    });
  }

  var totalQty = 0;
  results.forEach(function(r) { totalQty += r.qty; });

  return { results: results, count: results.length, totalQty: totalQty };
}

// ═══════════════════════════════════════════════════════════
// TAB 2 — CYCLE COUNT
// ═══════════════════════════════════════════════════════════

function getCycleCount(tanggal) {
  if (!tanggal) return { error: 'tanggal required' };

  var sh = SpreadsheetApp.openById(SS_ID).getSheetByName(SH_MASTER);
  var data = sh.getDataRange().getValues();

  var targetDate = new Date(tanggal);
  targetDate.setHours(0, 0, 0, 0);
  var targetEnd = new Date(tanggal);
  targetEnd.setHours(23, 59, 59, 999);

  // Pass 1: find kode|batch combos active ON the target date
  var activeKeys = {};

  for (var i = 3; i < data.length; i++) {
    var row = data[i];
    var kode = str(row[MD_KODE]);
    if (!kode) continue;

    var tgl = row[MD_TGL];
    if (!(tgl instanceof Date)) continue;

    var batch = str(row[MD_BATCH]);
    var key = kode + '|' + batch;

    if (tgl >= targetDate && tgl <= targetEnd) {
      if (!activeKeys[key]) {
        activeKeys[key] = {
          kodeBarang:  kode,
          namaDagang:  str(row[MD_NAMA_DGG]),
          batch:       batch,
          mfg:         fmtDate(row[MD_MFG]),
          expire:      fmtDate(row[MD_EXPIRE]),
          expireSort:  row[MD_EXPIRE] instanceof Date ? row[MD_EXPIRE].getTime() : 0,
          penHariIni:  0,
          distHariIni: 0,
          penKum:      0,
          distKum:     0
        };
      }
      activeKeys[key].penHariIni  += num(row[MD_PEN]);
      activeKeys[key].distHariIni += num(row[MD_DIST]);
    }
  }

  // Pass 2: cumulative up to (<= targetEnd)
  for (var i = 3; i < data.length; i++) {
    var row = data[i];
    var kode = str(row[MD_KODE]);
    if (!kode) continue;

    var tgl = row[MD_TGL];
    if (!(tgl instanceof Date)) continue;
    if (tgl > targetEnd) continue;

    var batch = str(row[MD_BATCH]);
    var key = kode + '|' + batch;

    if (activeKeys[key]) {
      activeKeys[key].penKum  += num(row[MD_PEN]);
      activeKeys[key].distKum += num(row[MD_DIST]);
    }
  }

  var results = [];
  var sumPen = 0, sumDist = 0, sumStok = 0;

  for (var k in activeKeys) {
    var g = activeKeys[k];
    g.stokAkhir = g.penKum - g.distKum;
    sumPen  += g.penHariIni;
    sumDist += g.distHariIni;
    sumStok += g.stokAkhir;
    results.push(g);
  }

  results.sort(function(a, b) {
    var cmp = a.namaDagang.localeCompare(b.namaDagang);
    return cmp !== 0 ? cmp : a.expireSort - b.expireSort;
  });

  results.forEach(function(x) {
    delete x.expireSort;
    delete x.penKum;
    delete x.distKum;
  });

  return {
    results: results,
    count:   results.length,
    tanggal: fmtDate(targetDate),
    summary: { totalPenerimaan: sumPen, totalDistribusi: sumDist, totalStok: sumStok }
  };
}

// ═══════════════════════════════════════════════════════════
// TAB 2 — BATCH RECORD (search → cards → detail with running stok)
// ═══════════════════════════════════════════════════════════

/**
 * Search by product name/code/batch → return unique batch cards
 */
function searchBatch(query) {
  var q = (query || '').toString().toLowerCase().trim();
  if (!q) return { results: [] };

  var ss = SpreadsheetApp.openById(SS_ID);
  var shMaster = ss.getSheetByName(SH_MASTER);
  var data = shMaster.getDataRange().getValues();

  var shProduk = ss.getSheetByName(SH_PRODUK);
  var produkData = shProduk.getDataRange().getValues();

  var batchMap = {}; // key: kode|batch

  for (var i = 3; i < data.length; i++) {
    var row = data[i];
    var kode      = str(row[MD_KODE]);
    var namaDagang = str(row[MD_NAMA_DGG]);
    var batch     = str(row[MD_BATCH]);
    if (!kode || !batch) continue;

    var haystack = [kode, namaDagang, str(row[MD_NAMA_ERP]), batch].join(' ').toLowerCase();
    if (haystack.indexOf(q) === -1) continue;

    var key = kode + '|' + batch;
    if (!batchMap[key]) {
      batchMap[key] = {
        kodeBarang: kode,
        namaDagang: namaDagang,
        batch:      batch,
        mfg:        fmtDate(row[MD_MFG]),
        expire:     fmtDate(row[MD_EXPIRE]),
        penerimaan: 0,
        distribusi: 0
      };
    }
    batchMap[key].penerimaan += num(row[MD_PEN]);
    batchMap[key].distribusi += num(row[MD_DIST]);
  }

  var results = [];
  for (var k in batchMap) {
    var b = batchMap[k];
    b.stok = b.penerimaan - b.distribusi;
    results.push(b);
  }

  // Sort by namaDagang then batch
  results.sort(function(a, b) {
    var cmp = a.namaDagang.localeCompare(b.namaDagang);
    return cmp !== 0 ? cmp : a.batch.localeCompare(b.batch);
  });

  return { results: results };
}

/**
 * Get all transactions for a specific kode+batch, in row order, with running stok
 */
function getBatchDetail(kodeBarang, batch) {
  if (!kodeBarang || !batch) return { error: 'kodeBarang and batch required' };

  var ss = SpreadsheetApp.openById(SS_ID);
  var shMaster = ss.getSheetByName(SH_MASTER);
  var data = shMaster.getDataRange().getValues();

  var shProduk = ss.getSheetByName(SH_PRODUK);
  var produkData = shProduk.getDataRange().getValues();
  var master = lookupProduk(produkData, kodeBarang);

  // Pass 1: group by tanggal+konsumen+kota, preserve sheet order
  var groups = [];
  var groupMap = {};

  for (var i = 3; i < data.length; i++) {
    var row = data[i];
    if (str(row[MD_KODE]) !== kodeBarang) continue;
    if (str(row[MD_BATCH]) !== batch) continue;

    var tgl  = fmtDate(row[MD_TGL]);
    var kons = str(row[MD_KONSUMEN]);
    var kota = str(row[MD_KOTA]);
    var key  = tgl + '|' + kons + '|' + kota;

    if (!groupMap[key]) {
      groupMap[key] = { tanggal: tgl, konsumen: kons, kota: kota, masuk: 0, keluar: 0 };
      groups.push(groupMap[key]);
    }
    groupMap[key].masuk  += num(row[MD_PEN]);
    groupMap[key].keluar += num(row[MD_DIST]);
  }

  // Pass 2: calculate running stok
  var runningStok = 0;
  var rows = [];
  for (var g = 0; g < groups.length; g++) {
    var grp = groups[g];
    runningStok += grp.masuk - grp.keluar;
    rows.push({
      tanggal:  grp.tanggal,
      konsumen: grp.konsumen,
      kota:     grp.kota,
      masuk:    grp.masuk,
      keluar:   grp.keluar,
      stok:     runningStok
    });
  }

  // Get batch info from first matching row
  var batchInfo = { mfg: '', expire: '', status: '' };
  for (var i = 3; i < data.length; i++) {
    var row = data[i];
    if (str(row[MD_KODE]) === kodeBarang && str(row[MD_BATCH]) === batch) {
      batchInfo.mfg    = fmtDate(row[MD_MFG]);
      batchInfo.expire = fmtDate(row[MD_EXPIRE]);
      batchInfo.status = str(row[MD_STATUS]);
      break;
    }
  }

  return {
    kodeBarang: kodeBarang,
    namaDagang: master ? master.namaDagang : '',
    namaERP:    master ? master.namaERP : '',
    batch:      batch,
    mfg:        batchInfo.mfg,
    expire:     batchInfo.expire,
    status:     batchInfo.status,
    stokAkhir:  runningStok,
    rows:       rows
  };
}

/**
 * Export batch record to a new Google Sheet (A4 formatted), return URL
 */
// ═══════════════════════════════════════════════════════════
// EMBALAGE — Item List (from EMBALAGE sheet)
// ═══════════════════════════════════════════════════════════

function getEmbalageList() {
  var sh = SpreadsheetApp.openById(SS_ID).getSheetByName('EMBALAGE');
  var data = sh.getDataRange().getValues();
  var items = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var nama = str(row[2]);
    if (!nama) continue;
    items.push({
      no:          num(row[0]),
      nama:        nama,
      kode:        str(row[3]),
      aliasERP:    str(row[4]),
      kategori:    str(row[5]),
      satuan:      str(row[6]),
      keterangan:  str(row[7]),
      safetyStock: row[8] === '-' || row[8] == null ? 0 : num(row[8]),
      reorderPoint: row[9] == null ? 0 : num(row[9])
    });
  }
  return { items: items };
}

// ═══════════════════════════════════════════════════════════
// EMBALAGE — Penerimaan (left side of MASTER EMBALAGE)
// Read: row 5 = header, row 6+ = data
// Cols B-I (idx 1-8): No, NomorBON, Tanggal, KodeBarang, NamaBarang, Jumlah, Satuan, Keterangan
// ═══════════════════════════════════════════════════════════

function getEmbalagePenerimaan() {
  var sh = SpreadsheetApp.openById(SS_ID).getSheetByName('MASTER EMBALAGE');
  var data = sh.getDataRange().getValues();
  var results = [];

  for (var i = 5; i < data.length; i++) {
    var row = data[i];
    var tgl = row[3]; // col D = Tanggal
    if (!tgl) continue;
    results.push({
      no:        num(row[1]),
      nomorBon:  str(row[2]),
      tanggal:   fmtDate(tgl),
      kode:      str(row[4]),
      nama:      str(row[5]),
      jumlah:    num(row[6]),
      satuan:    str(row[7]),
      keterangan: str(row[8])
    });
  }

  return { results: results, count: results.length };
}

function addEmbalagePenerimaan(params) {
  var sh = SpreadsheetApp.openById(SS_ID).getSheetByName('MASTER EMBALAGE');
  var data = sh.getDataRange().getValues();

  // Find next empty row on the left side (check col B = idx 1)
  var nextRow = -1;
  var lastNo = 0;
  for (var i = 5; i < data.length; i++) {
    if (data[i][3]) { // col D has date = row is filled
      lastNo = num(data[i][1]);
    } else {
      nextRow = i + 1; // sheet rows are 1-indexed
      break;
    }
  }
  if (nextRow === -1) nextRow = data.length + 1;

  // Write data: cols B-I (2-9)
  var newNo = lastNo + 1;
  sh.getRange(nextRow, 2).setValue(newNo);                    // B: No
  sh.getRange(nextRow, 3).setValue(params.nomorBon || '');     // C: Nomor BON
  sh.getRange(nextRow, 4).setValue(new Date(params.tanggal));  // D: Tanggal
  sh.getRange(nextRow, 5).setValue(params.kode);               // E: Kode Barang
  sh.getRange(nextRow, 6).setValue(params.nama);               // F: Nama Barang
  sh.getRange(nextRow, 7).setValue(Number(params.jumlah));     // G: Jumlah
  sh.getRange(nextRow, 8).setValue(params.satuan);             // H: Satuan
  sh.getRange(nextRow, 9).setValue(params.keterangan || '');   // I: Keterangan

  // Update D3:D4 (UPDATED AT / UPDATED BY)
  sh.getRange('D3').setValue(new Date());
  sh.getRange('D4').setValue('Web App');

  return { success: true, row: nextRow, no: newNo };
}

// ═══════════════════════════════════════════════════════════
// EMBALAGE — Dashboard (stock = penerimaan - pengeluaran)
// ═══════════════════════════════════════════════════════════

function getEmbalageDashboard() {
  var ss = SpreadsheetApp.openById(SS_ID);

  // 1) Get embalage master list with safety stock
  var shEmb = ss.getSheetByName('EMBALAGE');
  var embData = shEmb.getDataRange().getValues();
  var items = {};
  for (var i = 1; i < embData.length; i++) {
    var nama = str(embData[i][2]);
    var kode = str(embData[i][3]);
    if (!nama) continue;
    var key = kode || nama;
    items[key] = {
      nama: nama,
      kode: kode,
      satuan: str(embData[i][6]),
      safetyStock: embData[i][8] === '-' || embData[i][8] == null ? 0 : num(embData[i][8]),
      reorderPoint: embData[i][9] == null ? 0 : num(embData[i][9]),
      penerimaan: 0,
      pengeluaran: 0,
      stok: 0
    };
  }

  // 2) Read MASTER EMBALAGE - left side (penerimaan) + right side (pengeluaran)
  var shMaster = ss.getSheetByName('MASTER EMBALAGE');
  var data = shMaster.getDataRange().getValues();

  for (var i = 5; i < data.length; i++) {
    // Left side: penerimaan (cols D,E,F,G = idx 3,4,5,6)
    if (data[i][3]) { // has date
      var kode = str(data[i][4]);
      if (items[kode]) items[kode].penerimaan += num(data[i][6]);
    }
    // Right side: pengeluaran (cols M,N,O,P = idx 12,13,14,15)
    if (data[i][12]) { // has date
      var kode2 = str(data[i][13]);
      if (items[kode2]) items[kode2].pengeluaran += num(data[i][15]);
    }
  }

  // Calculate stok
  var results = [];
  for (var k in items) {
    var item = items[k];
    item.stok = item.penerimaan - item.pengeluaran;
    results.push(item);
  }

  results.sort(function(a, b) { return a.nama.localeCompare(b.nama); });

  return { items: results };
}

// ═══════════════════════════════════════════════════════════
// EMBALAGE — Pengeluaran (right side of MASTER EMBALAGE)
// Cols K-R (idx 10-17): No, NoPacking, Tanggal, KodeBarang, NamaBarang, Jumlah, Satuan, Keterangan
// ═══════════════════════════════════════════════════════════

function getEmbalagePengeluaran(page, pageSize) {
  var sh = SpreadsheetApp.openById(SS_ID).getSheetByName('MASTER EMBALAGE');
  var data = sh.getDataRange().getValues();
  var all = [];

  for (var i = 5; i < data.length; i++) {
    var row = data[i];
    var tgl = row[12]; // col M = Tanggal pengeluaran
    if (!tgl) continue;
    all.push({
      no:         num(row[10]),
      noPacking:  str(row[11]),
      tanggal:    fmtDate(tgl),
      kode:       str(row[13]),
      nama:       str(row[14]),
      jumlah:     num(row[15]),
      satuan:     str(row[16]),
      keterangan: str(row[17])
    });
  }

  // Reverse so newest first
  all.reverse();

  var pg = num(page) || 1;
  var ps = num(pageSize) || 20;
  var total = all.length;
  var totalPages = Math.ceil(total / ps);
  var startIdx = (pg - 1) * ps;
  var paged = all.slice(startIdx, startIdx + ps);

  return { results: paged, total: total, page: pg, totalPages: totalPages };
}

// ═══════════════════════════════════════════════════════════
// EMBALAGE — Distribusi History (from sheet DISTRIBUSI)
// Group by TGL_KIRIM + TUJUAN_KIRIM, paginated
// ═══════════════════════════════════════════════════════════

var DS_TGL     = 3;  // D: TGL KIRIM
var DS_NO_PO   = 1;  // B: NOMOR PO
var DS_NO_DOK  = 2;  // C: NOMOR DOKUMEN
var DS_PRODUK  = 6;  // G: NAMA PRODUK
var DS_TUJUAN  = 7;  // H: TUJUAN KIRIM
var DS_ALOKASI = 8;  // I: ALOKASI
var DS_BATCH   = 9;  // J: NO BATCH
var DS_EXPIRE  = 10; // K: EXPIRE DATE
var DS_PEN     = 11; // L: PENERIMAAN
var DS_DIST    = 12; // M: DISTRIBUSI
var DS_BERAT   = 18; // S: BERAT PRODUK
var DS_FORWARDER = 19; // T: FORWARDER

function getDistribusiHistory(page, pageSize) {
  var sh = SpreadsheetApp.openById(SS_ID).getSheetByName('DISTRIBUSI');
  if (!sh) return { groups: [], total: 0, page: 1, totalPages: 0 };

  var data = sh.getDataRange().getValues();
  var pg = num(page) || 1;
  var ps = num(pageSize) || 20;

  // Group by tgl + tujuan (preserve order of first occurrence)
  var groups = [];
  var groupMap = {};

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var tgl = row[DS_TGL];
    if (!(tgl instanceof Date)) continue;

    var tglStr  = fmtDate(tgl);
    var tujuan  = str(row[DS_TUJUAN]);
    var key     = tglStr + '|' + tujuan;

    if (!groupMap[key]) {
      groupMap[key] = {
        tanggal:   tglStr,
        tujuan:    tujuan,
        noDokumen: str(row[DS_NO_DOK]),
        noPO:      str(row[DS_NO_PO]),
        berat:     num(row[DS_BERAT]),
        forwarder: str(row[DS_FORWARDER]),
        itemCount: 0,
        totalQty:  0
      };
      groups.push(groupMap[key]);
    }
    groupMap[key].itemCount++;
    groupMap[key].totalQty += num(row[DS_DIST]);
    // berat & forwarder from first row only (already captured)
    if (!groupMap[key].forwarder && str(row[DS_FORWARDER])) {
      groupMap[key].forwarder = str(row[DS_FORWARDER]);
    }
  }

  // Sort by date descending
  groups.sort(function(a, b) { return b.tanggal.localeCompare(a.tanggal); });

  var total = groups.length;
  var totalPages = Math.ceil(total / ps);
  var startIdx = (pg - 1) * ps;
  var paged = groups.slice(startIdx, startIdx + ps);

  return { groups: paged, total: total, page: pg, totalPages: totalPages };
}

function getDistribusiDetail(tanggal, tujuan) {
  var sh = SpreadsheetApp.openById(SS_ID).getSheetByName('DISTRIBUSI');
  if (!sh) return { rows: [] };

  var data = sh.getDataRange().getValues();
  var results = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var tgl = row[DS_TGL];
    if (!(tgl instanceof Date)) continue;
    if (fmtDate(tgl) !== tanggal) continue;
    if (str(row[DS_TUJUAN]) !== tujuan) continue;

    results.push({
      no:         num(row[0]),
      noPO:       str(row[DS_NO_PO]),
      noDokumen:  str(row[DS_NO_DOK]),
      produk:     str(row[DS_PRODUK]),
      batch:      str(row[DS_BATCH]),
      expire:     fmtDate(row[DS_EXPIRE]),
      penerimaan: num(row[DS_PEN]),
      distribusi: num(row[DS_DIST]),
      keterangan: str(row[15]),
      berat:      num(row[DS_BERAT]),
      forwarder:  str(row[DS_FORWARDER])
    });
  }

  return { tanggal: tanggal, tujuan: tujuan, rows: results };
}

// ═══════════════════════════════════════════════════════════
// UTILITIES
// ═══════════════════════════════════════════════════════════

function str(v) {
  if (v == null) return '';
  if (v instanceof Date) return fmtDate(v);
  return v.toString().trim();
}

function num(v) {
  if (v == null || v === '') return 0;
  var n = Number(v);
  return isNaN(n) ? 0 : n;
}

function fmtDate(v) {
  if (!(v instanceof Date)) return '';
  var y = v.getFullYear();
  var m = ('0' + (v.getMonth() + 1)).slice(-2);
  var d = ('0' + v.getDate()).slice(-2);
  return y + '-' + m + '-' + d;
}