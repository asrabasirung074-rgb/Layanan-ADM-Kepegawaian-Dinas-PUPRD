/**
 * BACKEND APPS SCRIPT — Formulir Pengajuan Kepegawaian + Panel Verifikasi
 * -------------------------------------------------------------------
 * CARA PASANG (sama seperti sebelumnya):
 * 1. Buka Google Sheet yang sudah dipakai (yang sudah berisi data lama Anda).
 * 2. Menu Extensions > Apps Script.
 * 3. Hapus isi lama, tempel SELURUH kode ini (menggantikan Code.gs lama).
 * 4. Wajib isi ADMIN_TOKEN di bawah dengan kode rahasia Anda sendiri —
 *    ini seperti "password" panel verifikasi. Jangan dibagikan ke pegawai.
 * 5. Deploy > Manage deployments > pilih deployment aktif > ikon pensil
 *    > Version: "New version" > Deploy. URL Web App tetap sama, tidak
 *    perlu diganti di file formulir.html maupun admin.html.
 *
 * APA YANG BARU:
 * - doGet(?action=list&token=...) -> mengembalikan semua baris pengajuan
 *   sebagai JSON, hanya jika token cocok dengan ADMIN_TOKEN.
 * - doPost dengan {action:'updateStatus', id, status, token, catatan,
 *   verifiedBy} -> mengubah status satu pengajuan berdasarkan Kode
 *   Referensi, mencatat siapa & kapan diverifikasi, dan (opsional)
 *   mengirim email pemberitahuan ke pegawai yang mengajukan.
 * - Sheet otomatis mendapat 3 kolom baru: "Catatan Verifikasi",
 *   "Diverifikasi Oleh", "Waktu Verifikasi" (aman untuk sheet lama,
 *   kolom akan ditambahkan otomatis tanpa menghapus data yang ada).
 *
 * CATATAN KEAMANAN:
 * ADMIN_TOKEN adalah satu-satunya proteksi panel verifikasi. Pilih
 * string yang panjang & acak, dan hanya beri tahu staf kepegawaian
 * yang berwenang. Siapa pun yang tahu token ini bisa mengubah status
 * dan melihat seluruh data pengajuan.
 */

const SHEET_NAME = 'Pengajuan';
const FOLDER_NAME = 'Berkas Pengajuan Kepegawaian';
const NOTIFY_EMAIL = 'asrabasirung074@gmail.com'; // <-- email tujuan notifikasi pengajuan baru
const SECRET_TOKEN = ''; // opsional: proteksi tambahan untuk form pegawai (lihat form HTML)
const ADMIN_TOKEN = 'GANTI-DENGAN-KODE-RAHASIA-ANDA'; // WAJIB DIGANTI — password panel verifikasi

const HEADERS = [
  'Kode Referensi', 'Waktu Kirim', 'Nama', 'NIP 18', 'NIP 16', 'No HP', 'Email',
  'Unit Kerja', 'Jabatan', 'Golongan Saat Ini', 'TMT Pangkat Terakhir', 'Jenis Pengajuan',
  'Pangkat Diusulkan', 'TMT Diusulkan', 'Jenjang Ijazah', 'Tanggal Ijazah',
  'TMT KGB Terakhir', 'Tanggal SK KGB', 'Status', 'Link Berkas', 'Folder Drive',
  'Catatan Verifikasi', 'Diverifikasi Oleh', 'Waktu Verifikasi',
];

function doGet(e) {
  const action = e && e.parameter && e.parameter.action;

  if (action === 'list') {
    if (!checkAdminToken(e.parameter.token)) {
      return jsonOutput({ ok: false, error: 'unauthorized' });
    }
    return jsonOutput({ ok: true, rows: listAllSubmissions() });
  }

  // Buka URL Web App langsung di browser untuk memastikan deployment aktif.
  return jsonOutput({ ok: true, message: 'Apps Script formulir kepegawaian aktif. Gunakan metode POST untuk mengirim data.' });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000); // tunggu maks 15 detik kalau ada request lain berjalan
  } catch (lockErr) {
    return jsonOutput({ ok: false, error: 'server-sibuk-coba-lagi' });
  }

  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonOutput({ ok: false, error: 'tidak-ada-data' });
    }

    const data = JSON.parse(e.postData.contents);

    if (data.action === 'updateStatus') {
      return handleUpdateStatus(data);
    }

    return handleSubmitPengajuan(data);
  } catch (err) {
    return jsonOutput({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/* ---------------------------------------------------------------------
 * PENGAJUAN BARU (form pegawai) — logika sama seperti sebelumnya
 * ------------------------------------------------------------------- */
function handleSubmitPengajuan(data) {
  if (SECRET_TOKEN && data.token !== SECRET_TOKEN) {
    return jsonOutput({ ok: false, error: 'unauthorized' });
  }

  const sheet = getOrCreateSheet();
  const subFolder = getOrCreateFolder(FOLDER_NAME).createFolder((data.id || 'TANPA-ID') + ' - ' + (data.nama || 'tanpa-nama'));
  const fileLinks = [];

  if (data.files) {
    Object.keys(data.files).forEach(function (key) {
      try {
        const f = data.files[key];
        if (!f || !f.base64) return;
        const parts = f.base64.split(',');
        const raw = parts.length > 1 ? parts[1] : parts[0];
        const blob = Utilities.newBlob(Utilities.base64Decode(raw), f.type || 'application/octet-stream', f.name || key);
        const savedFile = subFolder.createFile(blob);
        fileLinks.push(key + ': ' + savedFile.getUrl());
      } catch (fileErr) {
        // Satu file gagal tidak boleh menggagalkan seluruh pengajuan.
        fileLinks.push(key + ': GAGAL DISIMPAN (' + String(fileErr) + ')');
      }
    });
  }

  sheet.appendRow([
    data.id || '',
    data.submittedAt || new Date().toISOString(),
    data.nama || '',
    data.nip18 || '',
    data.nip16 || '',
    data.hp || '',
    data.email || '',
    data.unit || '',
    data.jabatan || '',
    data.golongan || '',
    data.tmtPangkat || '',
    data.jenis || '',
    data.pangkatDiusulkan || '',
    data.tmtDiusulkan || '',
    data.jenjangIjazah || '',
    data.tglIjazah || '',
    data.tmtKgb || '',
    data.tglSkKgb || '',
    data.status || 'Diterima - Menunggu Verifikasi',
    fileLinks.join('\n'),
    subFolder.getUrl(),
    '', // Catatan Verifikasi
    '', // Diverifikasi Oleh
    '', // Waktu Verifikasi
  ]);

  // Data sudah tersimpan di titik ini. Kegagalan kirim email TIDAK BOLEH
  // membuat respons ke form jadi ok:false, karena datanya sudah aman.
  if (NOTIFY_EMAIL) {
    try {
      MailApp.sendEmail({
        to: NOTIFY_EMAIL,
        subject: 'Pengajuan Baru: ' + (data.jenis || 'Kepegawaian') + ' — ' + (data.nama || ''),
        body:
          'Ada pengajuan baru masuk, menunggu verifikasi.\n\n' +
          'Kode Referensi : ' + data.id + '\n' +
          'Nama           : ' + data.nama + '\n' +
          'NIP            : ' + data.nip18 + '\n' +
          'Unit Kerja     : ' + data.unit + '\n' +
          'Jenis Pengajuan: ' + data.jenis + '\n' +
          'Waktu Kirim    : ' + data.submittedAt + '\n\n' +
          'Folder berkas  : ' + subFolder.getUrl() + '\n' +
          'Buka panel verifikasi untuk memproses pengajuan ini.',
      });
    } catch (mailErr) {
      Logger.log('Gagal kirim email notifikasi: ' + String(mailErr));
    }
  }

  return jsonOutput({ ok: true, id: data.id });
}

/* ---------------------------------------------------------------------
 * PANEL VERIFIKASI (admin) — baru
 * ------------------------------------------------------------------- */
function checkAdminToken(token) {
  return !!ADMIN_TOKEN && ADMIN_TOKEN !== 'GANTI-DENGAN-KODE-RAHASIA-ANDA' && token === ADMIN_TOKEN;
}

function listAllSubmissions() {
  const sheet = getOrCreateSheet();
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0];
  const rows = values.slice(1).map(function (row) {
    const obj = {};
    headers.forEach(function (h, i) { obj[h] = row[i] instanceof Date ? row[i].toISOString() : row[i]; });
    return obj;
  });
  return rows.reverse(); // terbaru tampil lebih dulu
}

function handleUpdateStatus(data) {
  if (!checkAdminToken(data.token)) {
    return jsonOutput({ ok: false, error: 'unauthorized' });
  }
  if (!data.id || !data.status) {
    return jsonOutput({ ok: false, error: 'data-tidak-lengkap' });
  }

  const sheet = getOrCreateSheet();
  const values = sheet.getDataRange().getValues();
  const headers = values[0];
  const idCol = headers.indexOf('Kode Referensi');
  const statusCol = headers.indexOf('Status');
  const catatanCol = headers.indexOf('Catatan Verifikasi');
  const verifByCol = headers.indexOf('Diverifikasi Oleh');
  const verifAtCol = headers.indexOf('Waktu Verifikasi');
  const emailCol = headers.indexOf('Email');
  const jenisCol = headers.indexOf('Jenis Pengajuan');

  for (let r = 1; r < values.length; r++) {
    if (String(values[r][idCol]) === String(data.id)) {
      const rowNum = r + 1;
      sheet.getRange(rowNum, statusCol + 1).setValue(data.status);
      if (catatanCol > -1) sheet.getRange(rowNum, catatanCol + 1).setValue(data.catatan || '');
      if (verifByCol > -1) sheet.getRange(rowNum, verifByCol + 1).setValue(data.verifiedBy || '');
      if (verifAtCol > -1) sheet.getRange(rowNum, verifAtCol + 1).setValue(new Date().toISOString());

      const email = emailCol > -1 ? values[r][emailCol] : '';
      if (email) {
        try {
          MailApp.sendEmail({
            to: email,
            subject: 'Status Pengajuan Diperbarui — ' + data.id,
            body:
              'Status pengajuan Anda telah diperbarui.\n\n' +
              'Kode Referensi : ' + data.id + '\n' +
              'Jenis Pengajuan: ' + (jenisCol > -1 ? values[r][jenisCol] : '') + '\n' +
              'Status Baru    : ' + data.status + '\n' +
              (data.catatan ? '\nCatatan dari Kepegawaian:\n' + data.catatan + '\n' : '') +
              '\nHubungi Bidang Kepegawaian jika ada pertanyaan.',
          });
        } catch (mailErr) {
          Logger.log('Gagal kirim email status: ' + String(mailErr));
        }
      }

      return jsonOutput({ ok: true });
    }
  }

  return jsonOutput({ ok: false, error: 'pengajuan-tidak-ditemukan' });
}

/* ---------------------------------------------------------------------
 * UTIL
 * ------------------------------------------------------------------- */
function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  } else {
    // Tambahkan kolom baru otomatis kalau sheet lama belum punya
    // (mis. sheet dibuat sebelum fitur verifikasi ada) — data lama aman.
    const lastCol = Math.max(sheet.getLastColumn(), 1);
    const existingHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    HEADERS.forEach(function (h) {
      if (existingHeaders.indexOf(h) === -1) {
        sheet.getRange(1, sheet.getLastColumn() + 1).setValue(h);
      }
    });
  }
  return sheet;
}

function getOrCreateFolder(name) {
  const folders = DriveApp.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(name);
}

function jsonOutput(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** Uji cepat dari editor Apps Script (opsional) */
function testDoPost() {
  const fakeEvent = {
    postData: {
      contents: JSON.stringify({
        id: 'TEST-0001',
        submittedAt: new Date().toISOString(),
        nama: 'Contoh Pegawai',
        nip18: '199001012020121001',
        nip16: '1234567890123456',
        hp: '081234567890',
        email: 'contoh@instansi.go.id',
        unit: 'Bidang Bina Marga',
        jabatan: 'Pengadministrasi Umum',
        golongan: 'III/b',
        tmtPangkat: '2022-01-01',
        jenis: 'Kenaikan Gaji Berkala - KGB',
        tmtKgb: '2024-01-01',
        files: {},
      }),
    },
  };
  Logger.log(doPost(fakeEvent).getContent());
}

/** Uji cepat endpoint verifikasi (opsional) — ganti ADMIN_TOKEN dulu sebelum jalankan */
function testUpdateStatus() {
  const fakeEvent = {
    postData: {
      contents: JSON.stringify({
        action: 'updateStatus',
        token: ADMIN_TOKEN,
        id: 'TEST-0001',
        status: 'Diverifikasi - Berkas Lengkap',
        catatan: 'Berkas lengkap dan sesuai.',
        verifiedBy: 'Staf Kepegawaian',
      }),
    },
  };
  Logger.log(doPost(fakeEvent).getContent());
}