/**
 * BACKEND APPS SCRIPT — Formulir Pengajuan Kepegawaian
 * -----------------------------------------------------
 * CARA PASANG:
 * 1. Buka Google Sheet baru (kosong).
 * 2. Menu Extensions > Apps Script.
 * 3. Hapus isi default, tempel seluruh kode ini.
 * 4. Ganti NOTIFY_EMAIL di bawah dengan email kepegawaian yang dituju.
 * 5. Klik Deploy > New deployment.
 *    - Type: Web app
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 6. Klik Deploy, izinkan akses saat diminta (Authorize access).
 * 7. Salin "Web app URL" yang muncul.
 * 8. Tempel URL tersebut ke variabel APPS_SCRIPT_URL di file HTML formulir.
 *
 * CATATAN KEAMANAN:
 * "Who has access: Anyone" berarti siapa pun yang tahu URL ini bisa
 * mengirim data ke Sheet Anda. Untuk formulir internal kantor ini umumnya
 * cukup karena URL tidak disebarluaskan, tapi jangan publikasikan URL-nya.
 * Untuk keamanan lebih, bisa ditambahkan token rahasia sederhana
 * (lihat catatan SECRET_TOKEN di bawah).
 */

const SHEET_NAME = 'Pengajuan';
const FOLDER_NAME = 'Berkas Pengajuan Kepegawaian';
const NOTIFY_EMAIL = 'asrabasirung074@gmail.com'; // <-- GANTI dengan email tujuan
const SECRET_TOKEN = ''; // opsional: isi string rahasia sama di sisi HTML jika ingin proteksi tambahan

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    if (SECRET_TOKEN && data.token !== SECRET_TOKEN) {
      return jsonOutput({ ok: false, error: 'unauthorized' });
    }

    const sheet = getOrCreateSheet();
    const subFolder = getOrCreateFolder(FOLDER_NAME).createFolder(data.id + ' - ' + (data.nama || 'tanpa-nama'));
    const fileLinks = [];

    if (data.files) {
      Object.keys(data.files).forEach(function (key) {
        const f = data.files[key];
        if (!f || !f.base64) return;
        const parts = f.base64.split(',');
        const raw = parts.length > 1 ? parts[1] : parts[0];
        const blob = Utilities.newBlob(Utilities.base64Decode(raw), f.type || 'application/octet-stream', f.name || key);
        const savedFile = subFolder.createFile(blob);
        fileLinks.push(key + ': ' + savedFile.getUrl());
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
    ]);

    if (NOTIFY_EMAIL) {
      MailApp.sendEmail({
        to: NOTIFY_EMAIL,
        subject: 'Pengajuan Baru: ' + (data.jenis || 'Kepegawaian') + ' — ' + (data.nama || ''),
        body:
          'Ada pengajuan baru masuk.\n\n' +
          'Kode Referensi : ' + data.id + '\n' +
          'Nama           : ' + data.nama + '\n' +
          'NIP            : ' + data.nip18 + '\n' +
          'Unit Kerja     : ' + data.unit + '\n' +
          'Jenis Pengajuan: ' + data.jenis + '\n' +
          'Waktu Kirim    : ' + data.submittedAt + '\n\n' +
          'Folder berkas  : ' + subFolder.getUrl() + '\n' +
          'Lihat detail lengkap di Google Sheet.',
      });
    }

    return jsonOutput({ ok: true, id: data.id });
  } catch (err) {
    return jsonOutput({ ok: false, error: String(err) });
  }
}

function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow([
      'Kode Referensi', 'Waktu Kirim', 'Nama', 'NIP 18', 'NIP 16', 'No HP', 'Email',
      'Unit Kerja', 'Jabatan', 'Golongan Saat Ini', 'TMT Pangkat Terakhir', 'Jenis Pengajuan',
      'Pangkat Diusulkan', 'TMT Diusulkan', 'Jenjang Ijazah', 'Tanggal Ijazah',
      'TMT KGB Terakhir', 'Tanggal SK KGB', 'Status', 'Link Berkas', 'Folder Drive',
    ]);
    sheet.setFrozenRows(1);
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
