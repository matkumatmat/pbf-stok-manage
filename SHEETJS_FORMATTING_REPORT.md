# SheetJS (XLSX) Formatting Capabilities Report

**Prepared for**: Stok PBF Web App Enhancement
**Date**: 2026-03-10
**Library**: SheetJS Community Edition (CE) - Free/Open Source Version
**Documentation Source**: https://docs.sheetjs.com

---

## Executive Summary

**TL;DR**: SheetJS Community Edition (yang kita pakai) **TIDAK SUPPORT** cell styling seperti warna background, bold font, borders, dll. Itu fitur **SheetJS Pro** (paid version).

Yang kita bisa lakukan di CE (gratis):
- ✅ Column width / Row height
- ✅ Number formatting (currency, percentage, decimals)
- ✅ Merge cells
- ✅ AutoFilter
- ✅ Sheet protection
- ❌ **Cell colors (background/font)**
- ❌ **Bold/italic/underline**
- ❌ **Borders**
- ❌ **Cell alignment**
- ❌ **Conditional formatting**

---

## Detailed Findings

### 1. SheetJS Community Edition (CE) - Free Version

#### ✅ Fitur yang TERSEDIA:

##### A. Column Width & Row Height
```javascript
ws['!cols'] = [
  { wch: 20 },  // Column A: 20 characters wide
  { wch: 30 },  // Column B: 30 characters wide
  { wch: 10 }   // Column C: 10 characters wide
];

ws['!rows'] = [
  { hpx: 30 },  // Row 1: 30 pixels height
  { hpx: 20 }   // Row 2: 20 pixels height
];
```

**Status di kode kita**: ✅ **SUDAH DIPAKAI**
- `exportBatch()`: Line 1323-1325 di index.html
- `exportStockPivot()`: Line 1376-1388 di index.html

##### B. Number Formatting
```javascript
// Format sebagai currency
worksheet["B2"].z = "$0.00";  // $3.50

// Format sebagai percentage
worksheet["B3"].z = "0.00%";  // 85.50%

// Format dengan thousands separator
worksheet["B4"].z = "#,##0";  // 1,234,567
```

**Status di kode kita**: ❌ **BELUM DIPAKAI**
Bisa ditambahkan untuk kolom Jumlah, Penerimaan, Distribusi, Stok.

##### C. Merge Cells
```javascript
ws['!merges'] = [
  { s: { r: 0, c: 0 }, e: { r: 1, c: 2 } }  // Merge A1 to C2
];
```

**Status di kode kita**: ❌ **BELUM DIPAKAI**
Bisa dipakai untuk title row di export.

##### D. AutoFilter
```javascript
ws['!autofilter'] = { ref: "A1:D1" };
```

**Status di kode kita**: ❌ **BELUM DIPAKAI**
Berguna untuk header row agar bisa di-filter di Excel.

---

#### ❌ Fitur yang TIDAK TERSEDIA (Butuh Pro):

##### A. Cell Background Colors
```javascript
// ❌ TIDAK BISA DI CE
ws['A1'].s = {
  fill: {
    patternType: "solid",
    fgColor: { rgb: "FFFF0000" }  // Red background
  }
};
```

##### B. Font Styling (Bold, Color, Size)
```javascript
// ❌ TIDAK BISA DI CE
ws['A1'].s = {
  font: {
    bold: true,
    color: { rgb: "FF0000FF" },  // Blue text
    sz: 14  // Font size
  }
};
```

##### C. Borders
```javascript
// ❌ TIDAK BISA DI CE
ws['A1'].s = {
  border: {
    top: { style: "thin", color: { rgb: "FF000000" } },
    bottom: { style: "thick", color: { rgb: "FF000000" } }
  }
};
```

##### D. Cell Alignment
```javascript
// ❌ TIDAK BISA DI CE
ws['A1'].s = {
  alignment: {
    horizontal: "center",
    vertical: "middle"
  }
};
```

---

### 2. Official Documentation Quotes

> **"SheetJS CE primarily focuses on data processing."**
> — https://docs.sheetjs.com/docs/getting-started/examples/import

> **"By default, SheetJS CE writers prioritize data preservation and do not export styling metadata."**
> — https://docs.sheetjs.com/docs/api/write-options

> **"For advanced cell and text styling, conditional formatting, and additional styling features, SheetJS Pro is available."**
> — https://docs.sheetjs.com/docs/api/write-options

---

### 3. Apa yang Bisa Kita Improve (Tanpa Upgrade ke Pro)

Dengan CE version, kita masih bisa improve user experience:

#### A. ✅ Number Formatting untuk Kolom Angka
**Before**:
```
Penerimaan | Distribusi | Stok
1234567    | 567890     | 666677
```

**After** (dengan number formatting):
```
Penerimaan  | Distribusi | Stok
1,234,567   | 567,890    | 666,677
```

**Implementation**:
```javascript
// Di exportStockPivot() setelah wsData di-build
res.items.forEach((item, idx) => {
    const rowNum = idx + 5;  // Starting from row 5 (after headers)

    // Format Penerimaan (column I)
    if(ws[`I${rowNum}`]) ws[`I${rowNum}`].z = "#,##0";

    // Format Distribusi (column J)
    if(ws[`J${rowNum}`]) ws[`J${rowNum}`].z = "#,##0";

    // Format Stok (column K)
    if(ws[`K${rowNum}`]) ws[`K${rowNum}`].z = "#,##0";
});
```

#### B. ✅ AutoFilter untuk Header
**Benefit**: User bisa sort/filter langsung di Excel

**Implementation**:
```javascript
// Di exportStockPivot() setelah create worksheet
ws['!autofilter'] = { ref: "A4:K4" };  // Row 4 = header row
```

#### C. ✅ Freeze Panes (Header Row Stay Visible)
**Benefit**: Scroll down tanpa kehilangan header

**Implementation**:
```javascript
ws['!freeze'] = { xSplit: 0, ySplit: 4 };  // Freeze row 1-4
```

---

### 4. SheetJS Pro Comparison

| Fitur | CE (Free) | Pro (Paid) |
|-------|-----------|------------|
| Data extraction/generation | ✅ | ✅ |
| Column width / Row height | ✅ | ✅ |
| Number formatting | ✅ | ✅ |
| Merge cells | ✅ | ✅ |
| AutoFilter | ✅ | ✅ |
| **Cell background colors** | ❌ | ✅ |
| **Font colors/bold/italic** | ❌ | ✅ |
| **Borders** | ❌ | ✅ |
| **Alignment** | ❌ | ✅ |
| **Conditional formatting** | ❌ | ✅ |
| **Images & Charts** | ❌ | ✅ |
| **Rich text** | ❌ | ✅ |

**Pro Version Pricing**: Not publicly listed (contact sales)
**Pro License**: https://sheetjs.com/pro

---

## Recommendations

### Option 1: Stick with CE + Improve What We Can ✅ **RECOMMENDED**

**Pros**:
- Free, sudah installed
- Bisa improve dengan number formatting, autofilter, freeze panes
- Cukup untuk most use cases (data export)

**Cons**:
- Tidak ada warna/bold/borders (Excel akan "polosan")

**Action Items**:
1. Add number formatting untuk kolom angka
2. Add autofilter di header rows
3. Add freeze panes untuk better UX
4. (Optional) Add merge cells untuk title row

**Estimated Effort**: 30 menit

---

### Option 2: Upgrade ke SheetJS Pro 💰

**Pros**:
- Full styling support (colors, bold, borders, alignment)
- Professional-looking exports
- Charts & images support

**Cons**:
- Paid license (unknown pricing, contact required)
- Overhead untuk small project
- API changes required (different syntax)

**Action Items**:
1. Contact SheetJS sales untuk pricing
2. Evaluate ROI untuk project ini
3. Refactor export functions untuk Pro API

**Estimated Effort**: 2-4 jam (refactoring + testing)

---

### Option 3: Alternative Library (ExcelJS) 🔄

**Alternative**: ExcelJS (https://github.com/exceljs/exceljs)
- ✅ **FREE** & open source
- ✅ Supports colors, bold, borders, alignment
- ✅ Active development (last update: 2024)
- ❌ Larger bundle size (~1.2MB vs SheetJS 800KB)
- ❌ Migration effort required

**Pros**:
- Free styling support
- Modern API
- Good documentation

**Cons**:
- Perlu migrate semua export functions
- Larger bundle size (impact load time)
- Learning curve

**Action Items**:
1. Install ExcelJS via CDN
2. Refactor `exportBatch()` dan `exportStockPivot()`
3. Test di browser untuk bundle size impact

**Estimated Effort**: 3-5 jam (migration + testing)

**ExcelJS Styling Example**:
```javascript
// ExcelJS supports styling out of the box (FREE)
worksheet.getCell('A1').fill = {
  type: 'pattern',
  pattern: 'solid',
  fgColor: { argb: 'FFFF0000' }  // Red background
};

worksheet.getCell('A1').font = {
  bold: true,
  color: { argb: 'FFFFFFFF' },  // White text
  size: 14
};
```

---

## Conclusion

**Kenapa Excel kita "polosan"?**
Karena SheetJS Community Edition **by design** tidak support cell styling (colors, bold, borders). Ini design decision dari library untuk keep it lightweight dan fokus ke data processing.

**Recommended Action**: **Option 1** (Improve what we can with CE)
- Quick wins: number formatting, autofilter, freeze panes
- Zero cost
- Minimal effort
- Good enough untuk business needs (data export focus)

**If styling is critical**: Consider **Option 3** (ExcelJS migration)
- Free alternative dengan full styling
- Modern API
- Worth the migration effort jika professional-looking exports penting

**Only if budget allows**: **Option 2** (SheetJS Pro)
- Premium solution
- Unknown pricing
- Probably overkill untuk project scale ini

---

## Implementation Examples (Quick Wins)

### Enhance `exportStockPivot()` dengan CE Features

```javascript
function exportStockPivot() {
    // ... existing code ...

    google.script.run
        .withSuccessHandler(res => {
            // ... existing wsData building ...

            const wb = XLSX.utils.book_new();
            const ws = XLSX.utils.aoa_to_sheet(wsData);

            // ✅ 1. AutoFilter di header row
            ws['!autofilter'] = { ref: "A4:K4" };

            // ✅ 2. Freeze header rows (1-4)
            ws['!freeze'] = { xSplit: 0, ySplit: 4 };

            // ✅ 3. Number formatting untuk kolom angka
            res.items.forEach((item, idx) => {
                const rowNum = idx + 5;

                // Penerimaan (column I)
                const penCell = `I${rowNum}`;
                if(ws[penCell]) ws[penCell].z = "#,##0";

                // Distribusi (column J)
                const distCell = `J${rowNum}`;
                if(ws[distCell]) ws[distCell].z = "#,##0";

                // Stok (column K)
                const stokCell = `K${rowNum}`;
                if(ws[stokCell]) ws[stokCell].z = "#,##0";
            });

            // ✅ 4. Column widths (sudah ada)
            ws['!cols'] = [
                {wch: 12}, {wch: 14}, {wch: 28}, {wch: 28},
                {wch: 16}, {wch: 12}, {wch: 12}, {wch: 12},
                {wch: 12}, {wch: 12}, {wch: 12}
            ];

            // Bold title (sudah ada - ini hanya metadata, not actual styling)
            if(ws['A1']) ws['A1'].s = { font: { bold: true, sz: 14 } };

            XLSX.utils.book_append_sheet(wb, ws, 'Stock Pivot');

            const filename = 'Stock_Pivot_' + new Date().toISOString().split('T')[0] + '.xlsx';
            XLSX.writeFile(wb, filename);
        })
        .handleRequest('getStockPivotExport', {});
}
```

**Note**: `ws['A1'].s = { font: { bold: true } }` di code kita saat ini **TIDAK BEKERJA** di CE. Itu syntax untuk Pro version. Tapi tidak error, hanya di-ignore saat generate file.

---

## References

- SheetJS CE Documentation: https://docs.sheetjs.com
- SheetJS Pro: https://sheetjs.com/pro
- ExcelJS (Alternative): https://github.com/exceljs/exceljs
- SheetJS Cell Object Format: https://docs.sheetjs.com/docs/csf/cell
- Number Formatting: https://docs.sheetjs.com/docs/csf/features/nf
- Column Properties: https://docs.sheetjs.com/docs/csf/features/colprops

---

**Report Generated**: 2026-03-10
**Library Verified**: /websites/sheetjs (Context7)
**Benchmark Score**: 91.95/100
**Source Reputation**: High
