// function onEdit(e) {
//   if (!e) return; 

//   const sheet = e.source.getActiveSheet();
//   const sheetName = sheet.getName();
//   const range = e.range;
//   const editedCol = range.getColumn();
//   const editedRow = range.getRow();
  
//   // Hitung seberapa besar area yang di-copas/diedit
//   const numCols = range.getNumColumns(); 
//   const numRows = range.getNumRows(); 
  
//   const sheetConfigs = {
//     "MASTER DISTRIBUSI": { distCol: 5, cabangCol: 6 }, // Kolom E (5) dan F (6)
//     "MASTER KONSINYASI": { distCol: 3, cabangCol: 4 }  // Kolom C (3) dan D (4)
//   };

//   if (sheetConfigs[sheetName] && editedRow > 1) {
//     const config = sheetConfigs[sheetName];
    
//     // PENGAMAN 1: Kalau copas lebih dari 1 kolom (Konsumen & Cabang masuk bareng), batalkan script.
//     // Ini biar data Cabang hasil copas Fitrah nggak ke-reset jadi kosong.
//     if (numCols > 1) return; 

//     // Pastikan yang diedit / di-copas ada di kolom Konsumen/Distributor
//     if (editedCol === config.distCol) {
      
//       const values = range.getValues(); // Ambil semua data yang baru diinput (bisa 1 sel, bisa banyak sel ke bawah)
//       const sourceSheet = e.source.getSheetByName("DISTRIBUTOR");
      
//       if (!sourceSheet) return;
      
//       const lastRow = sourceSheet.getLastRow();
//       if (lastRow < 3) return; 
      
//       const sourceData = sourceSheet.getRange(3, 2, lastRow - 2, 2).getValues();

//       // PENGAMAN 2: Looping sebanyak baris yang diedit (bisa tangani bulk-paste 1 kolom)
//       for (let i = 0; i < numRows; i++) {
//         const currentRow = editedRow + i;
//         const distributorValue = values[i][0];
//         const cabangCell = sheet.getRange(currentRow, config.cabangCol);

//         // Reset cell cabang di baris ini
//         cabangCell.clearDataValidations();
//         cabangCell.setValue("");

//         // Kalau sel Konsumen di baris ini nggak kosong, buatkan dropdown Cabangnya
//         if (distributorValue !== "") {
//           const validCabang = sourceData
//             .filter(row => String(row[0]).trim() === String(distributorValue).trim())
//             .map(row => String(row[1]).trim());

//           if (validCabang.length > 0) {
//             const uniqueCabang = [...new Set(validCabang)];
//             const rule = SpreadsheetApp.newDataValidation()
//               .requireValueInList(uniqueCabang, true)
//               .build();
//             cabangCell.setDataValidation(rule);
//           }
//         }
//       }
//     }
//   }
// }