// Tambahkan variabel global ini di bagian atas file
var SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1seHa652xv4LbQe1iftnU_Ed7tThpf6kdDJFwKpe2Whc/edit"; // Ganti dengan URL spreadsheet Anda

// Fungsi untuk mendapatkan spreadsheet dengan aman
function getSpreadsheet() {
  return SpreadsheetApp.openByUrl(SPREADSHEET_URL);
}


function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Data Management System')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Fungsi untuk mendapatkan semua data
function getAllData() {
  const ss = getSpreadsheet();
  const dataSheet = ss.getSheetByName('Data');
  
  // Jika tidak ada data, kembalikan array kosong
  if (dataSheet.getLastRow() <= 1) {
    return [];
  }
  
  const dataRange = dataSheet.getRange(2, 1, dataSheet.getLastRow() - 1, 8).getValues();
  const data = [];
  
  dataRange.forEach(function(row) {
    // Konversi tanggal menjadi format YYYY-MM-DD untuk input date HTML
    let formattedDate = '';
    if (row[1] instanceof Date) {
      const date = row[1];
      const year = date.getFullYear();
      const month = String(date.getMonth() + 1).padStart(2, '0');
      const day = String(date.getDate()).padStart(2, '0');
      formattedDate = `${year}-${month}-${day}`;
    }
    
    data.push({
      id: row[0],
      tanggal: formattedDate,
      bulan: row[2],
      bidang: row[3],
      subKegiatan: row[4],
      sumberDana: row[5],
      uraian: row[6],
      total: row[7]
    });
  });
  
  return data;
}

// Fungsi untuk mendapatkan data referensi (dropdowns)
function getReferenceData() {
  const ss = getSpreadsheet();
  
  // Ambil data bulan
  const bulanSheet = ss.getSheetByName('bulan');
  const bulanData = bulanSheet.getRange(1, 1, bulanSheet.getLastRow(), 1).getValues().flat();
  
  // Ambil data bidang
  const bidangSheet = ss.getSheetByName('bid');
  const bidangData = bidangSheet.getRange(1, 1, bidangSheet.getLastRow(), 1).getValues().flat();
  
  // Ambil data sub kegiatan
  const skSheet = ss.getSheetByName('SK');
  const skData = skSheet.getRange(1, 1, skSheet.getLastRow(), 1).getValues().flat();
  
  // Ambil data sumber dana
  const sdSheet = ss.getSheetByName('sd');
  const sdData = sdSheet.getRange(1, 1, sdSheet.getLastRow(), 1).getValues().flat();
  
  return {
    bulan: bulanData,
    bidang: bidangData,
    subKegiatan: skData,
    sumberDana: sdData
  };
}

// Fungsi untuk menyimpan data baru
function saveData(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = ss.getSheetByName('Data');
  
  // Konversi string tanggal ke objek Date
  const tanggalParts = data.tanggal.split('-');
  const tanggalObj = new Date(tanggalParts[0], tanggalParts[1] - 1, tanggalParts[2]);
  
  // Generate ID unik (timestamp)
  const id = new Date().getTime().toString();
  
  // Tambahkan data baru
  dataSheet.appendRow([
    id,
    tanggalObj,
    data.bulan,
    data.bidang,
    data.subKegiatan,
    data.sumberDana,
    data.uraian,
    parseFloat(data.total)
  ]);
  
  return { success: true, id: id };
}

// Fungsi untuk mengupdate data yang ada
function updateData(data) {
  try {
    const ss = getSpreadsheet();
    const dataSheet = ss.getSheetByName('Data');
    
    // Log data yang diterima untuk debugging
    Logger.log('Data untuk update: ' + JSON.stringify(data));
    
    // Konversi string tanggal ke objek Date
    const tanggalParts = data.tanggal.split('-');
    const tanggalObj = new Date(tanggalParts[0], tanggalParts[1] - 1, tanggalParts[2]);
    
    // Cari baris dengan ID yang sesuai
    const dataRange = dataSheet.getRange(2, 1, dataSheet.getLastRow() - 1, 1).getValues();
    let rowIndex = -1;
    
    // Log untuk debugging
    Logger.log('Mencari ID: ' + data.id);
    Logger.log('Data Range: ' + JSON.stringify(dataRange));
    
    for (let i = 0; i < dataRange.length; i++) {
      const currentId = String(dataRange[i][0]); // Konversi ke string untuk perbandingan yang konsisten
      Logger.log('Checking ID: ' + currentId + ' vs ' + String(data.id));
      
      if (currentId === String(data.id)) {
        rowIndex = i + 2; // +2 karena index array dimulai dari 0 dan baris header
        break;
      }
    }
    
    if (rowIndex === -1) {
      Logger.log('Data tidak ditemukan dengan ID: ' + data.id);
      return { success: false, message: 'Data tidak ditemukan' };
    }
    
    Logger.log('Updating data at row: ' + rowIndex);
    
    // Update data
    dataSheet.getRange(rowIndex, 2).setValue(tanggalObj);
    dataSheet.getRange(rowIndex, 3).setValue(data.bulan);
    dataSheet.getRange(rowIndex, 4).setValue(data.bidang);
    dataSheet.getRange(rowIndex, 5).setValue(data.subKegiatan);
    dataSheet.getRange(rowIndex, 6).setValue(data.sumberDana);
    dataSheet.getRange(rowIndex, 7).setValue(data.uraian);
    dataSheet.getRange(rowIndex, 8).setValue(parseFloat(data.total));
    
    return { success: true, message: 'Data berhasil diupdate' };
  } catch (error) {
    Logger.log('Error saat update data: ' + error.toString());
    return { success: false, message: 'Error update data: ' + error.toString() };
  }
}

// Fungsi untuk menghapus data
function deleteData(id) {
  try {
    // Gunakan fungsi getSpreadsheet() yang sudah dibuat sebelumnya
    const ss = getSpreadsheet();
    const dataSheet = ss.getSheetByName('Data');
    
    if (!dataSheet) {
      return { success: false, message: 'Sheet "Data" tidak ditemukan' };
    }
    
    // Pastikan lastRow adalah jumlah baris yang sebenarnya
    const lastRow = dataSheet.getLastRow();
    if (lastRow <= 1) {
      return { success: false, message: 'Tidak ada data untuk dihapus' };
    }
    
    // Ambil semua ID dari kolom pertama
    const dataRange = dataSheet.getRange(2, 1, lastRow - 1, 1).getValues();
    let rowIndex = -1;
    
    // Log untuk debugging
    Logger.log('Mencari ID: ' + id);
    Logger.log('Data Range: ' + JSON.stringify(dataRange));
    
    // Cari baris dengan ID yang sesuai
    for (let i = 0; i < dataRange.length; i++) {
      const currentId = String(dataRange[i][0]); // Konversi ke string untuk perbandingan yang konsisten
      Logger.log('Checking ID: ' + currentId);
      
      if (currentId === String(id)) {
        rowIndex = i + 2; // +2 karena index array dimulai dari 0 dan baris header
        break;
      }
    }
    
    if (rowIndex === -1) {
      return { success: false, message: 'Data tidak ditemukan (ID: ' + id + ')' };
    }
    
    // Hapus baris
    dataSheet.deleteRow(rowIndex);
    
    return { success: true, message: 'Data berhasil dihapus' };
  } catch (error) {
    Logger.log('Error saat menghapus data: ' + error.toString());
    return { success: false, message: 'Error menghapus data: ' + error.toString() };
  }
}