function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('TestOnline')
      .setTitle('Tes Online - PT Multi Indomandiri')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getInitData() {
  var ss = SpreadsheetApp.openById("1hlKN5Daz4PqS-60Oso8Rl1CwijaW72saNRMtT9OvP-U");
  var sheetPengaturan = ss.getSheetByName("PENGATURAN");
  var posisiString = sheetPengaturan.getRange("B2").getValue().toString();
  var pilihanPosisi = posisiString.split(",").map(function(item) { return item.trim(); });
  return { posisi: pilihanPosisi };
}

function verifikasiDanAmbilSoal(tokenInput, nama, email, posisi, deviceId) {
  var ss = SpreadsheetApp.openById("1hlKN5Daz4PqS-60Oso8Rl1CwijaW72saNRMtT9OvP-U");

  try {
    // === 1. CEK WAKTU ===
    var sheetPengaturan = ss.getSheetByName("PENGATURAN");
    var waktuMulai   = sheetPengaturan.getRange("B3").getValue();
    var waktuSelesai = sheetPengaturan.getRange("B4").getValue();
    var waktuSekarang = new Date();

    if (waktuMulai && waktuMulai.toString().trim() !== "") {
      var tMulai = new Date(waktuMulai);
      if (!isNaN(tMulai.getTime()) && waktuSekarang.getTime() < tMulai.getTime()) {
        return { error: "⏳ UJIAN BELUM DIMULAI!\n" + sheetPengaturan.getRange("B3").getDisplayValue() };
      }
    }
    if (waktuSelesai && waktuSelesai.toString().trim() !== "") {
      var tSelesai = new Date(waktuSelesai);
      if (!isNaN(tSelesai.getTime()) && waktuSekarang.getTime() > tSelesai.getTime()) {
        return { error: "⛔ AKSES DITUTUP!\n" + sheetPengaturan.getRange("B4").getDisplayValue() };
      }
    }

    // === 2. CEK TOKEN ===
    var sheetToken = ss.getSheetByName("DATA_TOKEN");
    if (!sheetToken) return { error: "Sistem Error: Sheet DATA_TOKEN tidak ditemukan." };

    var dataToken = sheetToken.getDataRange().getValues();
    var barisDitemukan = -1;
    for (var i = 1; i < dataToken.length; i++) {
      if (dataToken[i][0].toString().trim() === tokenInput.trim()) {
        barisDitemukan = i;
        break;
      }
    }
    if (barisDitemukan === -1) return { error: "❌ TOKEN TIDAK TERDAFTAR!" };

    var statusToken = dataToken[barisDitemukan][1].toString().trim().toUpperCase();

    if (statusToken === "SELESAI") {
      return { error: "✅ Ujian Anda sudah tercatat dan selesai.\nTidak dapat mengakses ujian kembali." };
    }
    if (statusToken === "SUDAH DIPAKAI") {
      return { error: "❌ TOKEN HANGUS! Token ini sudah digunakan." };
    }

    var isRestore = (statusToken === "SEDANG DIKERJAKAN");

    // === PROTEKSI JOKI (CEK DEVICE ID) ===
    // Asumsi Device ID disimpan di Kolom E (index 4)
    var savedDeviceId = dataToken[barisDitemukan][4] ? dataToken[barisDitemukan][4].toString().trim() : "";
    
    if (isRestore) {
      // Jika status sedang dikerjakan, cocokkan ID Perangkatnya
      if (savedDeviceId !== "" && deviceId !== savedDeviceId) {
        return { error: "🚨 AKSES DITOLAK!\nToken ini sedang dikerjakan di perangkat/browser lain. Anda tidak bisa menggunakan lebih dari satu perangkat." };
      }
    }

    // === 3. SET STATUS TOKEN → SEDANG DIKERJAKAN ===
    var waktuAksesDetail = Utilities.formatDate(waktuSekarang, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
    sheetToken.getRange(barisDitemukan + 1, 2).setValue("SEDANG DIKERJAKAN");
    sheetToken.getRange(barisDitemukan + 1, 3).setValue(nama + " (" + email + ")");
    
    if (!isRestore) {
      sheetToken.getRange(barisDitemukan + 1, 4).setValue(waktuAksesDetail);
      sheetToken.getRange(barisDitemukan + 1, 5).setValue(deviceId); // Simpan Device ID ke Kolom E
    }

    // === 4. SEDOT SOAL ===
    var sheetVerbal   = ss.getSheetByName("SOAL_VERBAL");
    var sheetHitungan = ss.getSheetByName("SOAL_HITUNGAN");

    function prosesSheet(sheet, kodePrefix) {
      if (!sheet) return [];
      var data = sheet.getDataRange().getValues();
      if (data.length <= 1) return [];
      var paketTersedia = [];
      for (var i = 1; i < data.length; i++) {
        var namaPaket = data[i][0].toString().trim();
        if (namaPaket !== "" && paketTersedia.indexOf(namaPaket) === -1) paketTersedia.push(namaPaket);
      }
      if (paketTersedia.length === 0) return [];
      var paketTerpilih = paketTersedia[Math.floor(Math.random() * paketTersedia.length)];
      var soalTerpilih = [];
      for (var i = 1; i < data.length; i++) {
        if (data[i][0].toString().trim() === paketTerpilih) {
          var opsi = [];
          for (var col = 2; col <= 6; col++) {
            var teksOpsi = data[i][col].toString().trim();
            if (teksOpsi !== "" && teksOpsi !== "-") opsi.push(teksOpsi);
          }
          for (var j = opsi.length - 1; j > 0; j--) {
            var k = Math.floor(Math.random() * (j + 1));
            var temp = opsi[j]; opsi[j] = opsi[k]; opsi[k] = temp;
          }
          soalTerpilih.push({ id: kodePrefix + "_" + (i + 1), soal: data[i][1].toString(), opsi: opsi });
        }
      }
      for (var j = soalTerpilih.length - 1; j > 0; j--) {
        var k = Math.floor(Math.random() * (j + 1));
        var temp = soalTerpilih[j]; soalTerpilih[j] = soalTerpilih[k]; soalTerpilih[k] = temp;
      }
      return soalTerpilih;
    }

    return {
      success: true,
      isRestore: isRestore,
      verbal: prosesSheet(sheetVerbal, "V"),
      hitungan: prosesSheet(sheetHitungan, "H")
    };

  } catch(e) {
    return { error: "💥 Terjadi kesalahan sistem: " + e.message };
  }
}

function simpanHasilUjian(dataKandidat) {
  var ss = SpreadsheetApp.openById("1hlKN5Daz4PqS-60Oso8Rl1CwijaW72saNRMtT9OvP-U");
  var sheetHasil = ss.getSheetByName("HASIL_UJIAN");
  if (!sheetHasil) {
    sheetHasil = ss.insertSheet("HASIL_UJIAN");
    sheetHasil.appendRow(["Waktu Submit", "Nama", "Email", "Posisi", "Skor Total", "Verbal (%)", "Hitungan (%)", "Detail Verbal", "Detail Hitungan", "Status Pelanggaran"]);
    sheetHasil.getRange("A1:J1").setFontWeight("bold").setBackground("#0056b3").setFontColor("white");
    sheetHasil.setFrozenRows(1);
  }

  var sheetVerbal   = ss.getSheetByName("SOAL_VERBAL");
  var sheetHitungan = ss.getSheetByName("SOAL_HITUNGAN");
  var dataSheetVerbal   = sheetVerbal   ? sheetVerbal.getDataRange().getValues()   : [];
  var dataSheetHitungan = sheetHitungan ? sheetHitungan.getDataRange().getValues() : [];

  function koreksiDetail(dataJawabanUser, dataSheet) {
    var benar = 0; var detailLog = [];
    for (var key in dataJawabanUser) {
      if (dataJawabanUser.hasOwnProperty(key)) {
        var barisExcel    = parseInt(key.split("_")[1]) - 1;
        var kunciJawaban  = dataSheet[barisExcel][7].toString().trim().toLowerCase();
        var jawabanUser   = dataJawabanUser[key].toString().trim().toLowerCase();
        if (jawabanUser === kunciJawaban) {
          benar++;
          detailLog.push("Brs" + (barisExcel+1) + ":BENAR");
        } else {
          detailLog.push("Brs" + (barisExcel+1) + ":SALAH(" + (jawabanUser === "-" ? "Kosong" : jawabanUser) + ")");
        }
      }
    }
    return { skor: benar, log: detailLog.join(" | ") };
  }

  var hasilVerbal   = koreksiDetail(dataKandidat.verbal,   dataSheetVerbal);
  var hasilHitungan = koreksiDetail(dataKandidat.hitungan, dataSheetHitungan);
  var totalSkor     = Math.round((hasilVerbal.skor + hasilHitungan.skor) / 2);
  var statusPelanggaran = dataKandidat.pelanggaran >= 3
    ? "DISKUALIFIKASI (" + dataKandidat.pelanggaran + "x)"
    : dataKandidat.pelanggaran + " Peringatan";
  var waktuSubmitDetail = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");

  sheetHasil.appendRow([
    waktuSubmitDetail, dataKandidat.nama, dataKandidat.email, dataKandidat.posisi,
    totalSkor, hasilVerbal.skor, hasilHitungan.skor,
    hasilVerbal.log, hasilHitungan.log, statusPelanggaran
  ]);

  // === UPDATE STATUS TOKEN → SELESAI ===
  var sheetToken = ss.getSheetByName("DATA_TOKEN");
  if (sheetToken) {
    var dataToken = sheetToken.getDataRange().getValues();
    for (var i = 1; i < dataToken.length; i++) {
      if (dataToken[i][0].toString().trim() === dataKandidat.token) {
        sheetToken.getRange(i + 1, 2).setValue("SELESAI");
        break;
      }
    }
  }

  return "Sukses";
}
function cekStatusToken(tokenInput) {
  var ss = SpreadsheetApp.openById("1hlKN5Daz4PqS-60Oso8Rl1CwijaW72saNRMtT9OvP-U");
  var sheetToken = ss.getSheetByName("DATA_TOKEN");
  if (!sheetToken) return { status: "ERROR" };

  var dataToken = sheetToken.getDataRange().getValues();
  for (var i = 1; i < dataToken.length; i++) {
    if (dataToken[i][0].toString().trim() === tokenInput.trim()) {
      return { status: dataToken[i][1].toString().trim().toUpperCase() };
    }
  }
  return { status: "TIDAK DITEMUKAN" };
}
