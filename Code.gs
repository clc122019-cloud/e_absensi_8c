const CONFIG = {
  SHEET_SISWA: "SISWA",
  SHEET_ABSENSI: "ABSENSI",
  SHEET_USER: "USER",
  SHEET_CONFIG: "KONFIGURASI",
  SHEET_REKAP_HARIAN: "REKAP_HARIAN",
  SHEET_REKAP_BULANAN: "REKAP_BULANAN",
  SHEET_LOG_NOTIFIKASI: "LOG_NOTIFIKASI"
};

function doGet(e) {
  try {
    const action = e && e.parameter ? e.parameter.action : "";

    if (action === "ambilSiswa") {
      return output(ambilSiswa());
    }

    if (action === "ambilAbsensi") {
      return output(ambilAbsensi({
        tanggal: e.parameter.tanggal || formatTanggal(new Date())
      }));
    }

    return output({
      success: true,
      message: "e-ABSENSI 8C API aktif",
      timestamp: new Date().toISOString()
    });
  } catch (error) {
    return output({success:false, message:error.message});
  }
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents || "{}");
    const action = data.action;

    if (action === "simpanAbsensi") return output(simpanAbsensi(data));
    if (action === "ambilAbsensi") return output(ambilAbsensi(data));
    if (action === "ambilSiswa") return output(ambilSiswa());
    if (action === "rekapHarian") return output(rekapHarian(data));
    if (action === "rekapBulanan") return output(rekapBulanan(data));

    return output({success:false,message:"Action tidak dikenali: "+action});
  } catch (error) {
    return output({success:false,message:error.message});
  }
}

function ambilSiswa() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_SISWA);
  if (!sheet) throw new Error('Sheet "SISWA" tidak ditemukan.');

  const values = sheet.getDataRange().getDisplayValues();
  if (values.length <= 1) return {success:true,data:[]};

  const data = values.slice(1)
    .filter(row => row[0] || row[1] || row[2])
    .map(row => ({
      id: row[0],
      nis: row[1],
      nama: row[2],
      kelas: row[3] || "VIII-C",
      wa: row[4],
      status: row[5] || "AKTIF"
    }));

  return {success:true,jumlah:data.length,data:data};
}

function simpanAbsensi(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_ABSENSI);
  if (!sheet) throw new Error('Sheet "ABSENSI" tidak ditemukan.');

  const siswa = data.siswa || [];
  if (!siswa.length) throw new Error("Data siswa kosong.");

  const tanggal = data.tanggal || formatTanggal(new Date());
  const inputOleh = data.inputOleh || "Wali Kelas VIII-C";
  const waktuInput = new Date();

  const rows = siswa.map(item => [
    generateId("ABS"), tanggal, item.nis || "", item.nama || "",
    item.kelas || "VIII-C", item.status || "H", item.keterangan || "",
    inputOleh, waktuInput
  ]);

  sheet.getRange(sheet.getLastRow()+1,1,rows.length,rows[0].length).setValues(rows);

  return {success:true,message:"Absensi berhasil disimpan.",jumlah:rows.length,tanggal:tanggal};
}

function ambilAbsensi(data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_ABSENSI);
  if (!sheet) throw new Error('Sheet "ABSENSI" tidak ditemukan.');

  const tanggal = data.tanggal || formatTanggal(new Date());
  const values = sheet.getDataRange().getDisplayValues();

  const hasil = values.slice(1).filter(row => row[1] === tanggal).map(row => ({
    id:row[0], tanggal:row[1], nis:row[2], nama:row[3], kelas:row[4],
    status:row[5], keterangan:row[6], inputOleh:row[7], waktuInput:row[8]
  }));

  return {success:true,data:hasil};
}

function rekapHarian(data) {
  const absensi = ambilAbsensi({tanggal:data.tanggal || formatTanggal(new Date())}).data;
  const statistik={total:absensi.length,hadir:0,sakit:0,izin:0,alpa:0};
  absensi.forEach(x=>{if(x.status==="H")statistik.hadir++;if(x.status==="S")statistik.sakit++;if(x.status==="I")statistik.izin++;if(x.status==="A")statistik.alpa++;});
  statistik.persentase=statistik.total?((statistik.hadir/statistik.total)*100).toFixed(2):"0.00";
  return {success:true,statistik};
}

function rekapBulanan(data) {
  const bulan=Number(data.bulan)||(new Date().getMonth()+1);
  const tahun=Number(data.tahun)||new Date().getFullYear();
  const sheet=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_ABSENSI);
  if(!sheet)throw new Error('Sheet "ABSENSI" tidak ditemukan.');
  const values=sheet.getDataRange().getValues();
  const statistik={hadir:0,sakit:0,izin:0,alpa:0};
  values.slice(1).forEach(row=>{
    const d=new Date(row[1]);
    if(d.getMonth()+1===bulan && d.getFullYear()===tahun){
      if(row[5]==="H")statistik.hadir++;
      if(row[5]==="S")statistik.sakit++;
      if(row[5]==="I")statistik.izin++;
      if(row[5]==="A")statistik.alpa++;
    }
  });
  const total=statistik.hadir+statistik.sakit+statistik.izin+statistik.alpa;
  statistik.total=total;
  statistik.persentase=total?((statistik.hadir/total)*100).toFixed(2):"0.00";
  return {success:true,bulan,tahun,statistik};
}

function generateId(prefix){
  return prefix+"-"+Utilities.formatDate(new Date(),Session.getScriptTimeZone(),"yyyyMMddHHmmss")+"-"+Math.floor(Math.random()*1000);
}
function formatTanggal(date){
  if(!(date instanceof Date))date=new Date(date);
  return Utilities.formatDate(date,Session.getScriptTimeZone(),"yyyy-MM-dd");
}
function output(data){
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
