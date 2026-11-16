// ============================================================================
// KONFIGURASI UTAMA (Sudah diisi dengan kedua tautan Google Form Anda)
// ============================================================================
const DEFAULT_CONFIG = {
  // Tempel URL Web App Google Apps Script (Skrip B) agar dropdown membaca DB ponpes & db lomba
  appsScriptWebAppUrl: '',
  googleFormUrls: {
    'daftar-pondok':
      'https://docs.google.com/forms/d/e/1FAIpQLSfGD4jcFpDz4PfbdNT8k8t2cpnuX8RsOGKp7uxVkarjKs_I1A/viewform?usp=header',
    'lomba-mabahis':
      'https://docs.google.com/forms/d/e/1FAIpQLSfPrhwdGJtnKnZPisXkrNCWhVHsnchSVnc2wjQpfNWmyUsvnA/viewform?usp=header',
  },
};

const STORAGE_CONFIG_KEY = 'mabahis_google_config_v2';
const STORAGE_HISTORY_KEY = 'mabahis_submission_history_v2';
const STORAGE_LOCAL_PONPES_KEY = 'mabahis_local_ponpes_v1';

// ============================================================================
// DEFINISI 2 FORMULIR SESUAI KOLOM ASLI GOOGLE FORM & DATABASE_MABAHIS
// ============================================================================
const FORMS = [
  {
    id: 'daftar-pondok',
    title: 'Pendaftaran Pondok Pesantren (Daftar Pondok)',
    category: 'Pondok Pesantren',
    sheetTabName: 'DB ponpes',
    description:
      'Pendataan resmi pondok pesantren peserta kegiatan MABAHIS.',
    schedule: 'Pendaftaran Terbuka',
    location: 'Pondok Pesantren',
    codePrefix: 'PDK',
    image: '/assets/aistudio/IMG_3891_64c4589c.JPG',
    fields: [
      {
        id: 'nama_lengkap_pondok',
        label: 'Nama Lengkap Pondok',
        entryId: 'entry.1617669335',
        type: 'text',
        required: true,
        placeholder: 'Contoh: Pondok Pesantren Al-Hikmah',
      },
      {
        id: 'alamat_pondok',
        label: 'Alamat Pondok',
        entryId: 'entry.876434650',
        type: 'textarea',
        required: true,
        placeholder: 'Tuliskan alamat jalan / dusun / RT RW lengkap',
      },
      {
        id: 'desa',
        label: 'Desa',
        entryId: 'entry.1600717297',
        type: 'text',
        required: true,
        placeholder: 'Nama desa / kelurahan',
      },
      {
        id: 'kecamatan',
        label: 'Kecamatan',
        entryId: 'entry.827729337',
        type: 'text',
        required: true,
        placeholder: 'Nama kecamatan',
      },
      {
        id: 'nama_pengasuh',
        label: 'Nama Pengasuh',
        entryId: 'entry.1536102953',
        type: 'text',
        required: true,
        placeholder: 'Nama lengkap kiai / pengasuh pondok',
      },
    ],
    emailOpening:
      'Data pendataan Pondok Pesantren Anda telah kami terima dan dicatat ke dalam sheet DB ponpes pada Google Spreadsheet Database_Mabahis.',
  },
  {
    id: 'lomba-mabahis',
    title: 'Pendaftaran Lomba MABAHIS',
    category: 'Lomba MABAHIS',
    sheetTabName: 'Pendaftaran Lomba MABAHIS',
    description:
      'Formulir pendaftaran peserta dan delegasi cabang Lomba MABAHIS.',
    schedule: 'Pendaftaran Lomba MABAHIS',
    location: 'Peserta Delegasi Pondok Pesantren',
    codePrefix: 'MBH',
    image: '/assets/aistudio/IMG_3891_64c4589c.JPG',
    fields: [
      {
        id: 'nama_peserta',
        label: 'Nama Peserta',
        entryId: 'entry.513119518',
        type: 'text',
        required: true,
        placeholder: 'Contoh: Ilham',
      },
      {
        id: 'nama_pondok',
        label: 'Nama Pondok',
        entryId: 'entry.146855471',
        type: 'select',
        dynamicSource: 'ponpes',
        required: true,
      },
      {
        id: 'lomba',
        label: 'Lomba',
        entryId: 'entry.840692791',
        type: 'select',
        dynamicSource: 'lomba',
        required: true,
      },
      {
        id: 'kategori',
        label: 'Kategori',
        entryId: 'entry.1177265094',
        type: 'select',
        required: true,
        options: ['Putra', 'Putri', 'Campuran'],
      },
      {
        id: 'no_wa_pj',
        label: 'NO WA Penanggung Jawab',
        entryId: 'entry.10985689',
        type: 'tel',
        required: true,
        placeholder: 'Contoh: 081234567890',
      },
    ],
    emailOpening:
      'Pendaftaran peserta Lomba MABAHIS telah tercatat di Database_Mabahis. Simpan kode registrasi berikut untuk verifikasi peserta dan koordinasi penanggung jawab.',
  },
];

// ============================================================================
// TEMPLATE GOOGLE APPS SCRIPT (Code.gs)
// ============================================================================
const APPS_SCRIPT_FORM_TRIGGER = `/**
 * SKRIP A: Trigger Google Form + Sinkronisasi Dropdown dari "DB ponpes" & "db lomba"
 * 1) Pasang Trigger "On form submit" pada fungsi kirimEmailDariGoogleForm(e)
 * 2) Opsional: Pasang Trigger "On change" atau jalankan sinkronkanDropdownGoogleForm()
 *    agar opsi dropdown di Google Form Lomba selalu update dari DB ponpes & db lomba.
 */
function kirimEmailDariGoogleForm(e) {
  var sheet = e.range.getSheet();
  var namaTab = sheet.getName();
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var values = e.values;

  var data = {};
  for (var i = 0; i < headers.length; i++) {
    data[String(headers[i]).trim()] = values[i] || '-';
  }

  var emailTujuan =
    data['Email Address'] ||
    data['Alamat Email'] ||
    data['Email'] ||
    '';

  if (!emailTujuan || String(emailTujuan).indexOf('@') === -1) {
    Logger.log('Email pengirim belum ada pada tab: ' + namaTab);
    return;
  }

  var isLomba = Boolean(data['Nama Peserta'] || data['Lomba'] || data['Kategori']);
  var namaPengirim = isLomba
    ? (data['Nama Peserta'] || 'Peserta')
    : (data['Nama Pengasuh'] || data['Nama Lengkap Pondok'] || 'Pengirim');

  var judulProgram = isLomba
    ? 'Pendaftaran Lomba MABAHIS'
    : 'Pendaftaran Pondok Pesantren (Daftar Pondok)';

  var awalanKode = isLomba ? 'MBH' : 'PDK';
  var kodeRegistrasi =
    awalanKode + '-' +
    Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyMMdd') + '-' +
    Math.floor(1000 + Math.random() * 9000);

  var subjek = 'Konfirmasi [' + kodeRegistrasi + '] - ' + judulProgram;

  var barisTabel = '';
  for (var j = 0; j < headers.length; j++) {
    barisTabel +=
      '<tr>' +
      '<td style="padding:8px 12px;border-bottom:1px solid #dce6e0;color:#5f766b;font-size:13px;">' + escapeHtml(headers[j]) + '</td>' +
      '<td style="padding:8px 12px;border-bottom:1px solid #dce6e0;color:#0b1f17;font-size:13px;font-weight:500;">' + escapeHtml(values[j] || '-') + '</td>' +
      '</tr>';
  }

  var htmlBody =
    '<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;border:1px solid #dce6e0;border-radius:12px;overflow:hidden;">' +
      '<div style="background:#064e3b;color:#ffffff;padding:24px;">' +
        '<div style="font-size:12px;color:#a7f3d0;margin-bottom:4px;">Notifikasi Otomatis · Database_Mabahis</div>' +
        '<h2 style="margin:0;font-size:20px;">' + escapeHtml(judulProgram) + '</h2>' +
      '</div>' +
      '<div style="padding:24px;">' +
        '<p style="font-size:15px;color:#0b1f17;">Halo <strong>' + escapeHtml(namaPengirim) + '</strong>,</p>' +
        '<p style="font-size:14px;color:#3f574c;line-height:1.6;">Terima kasih telah mengirimkan pendaftaran. Data Anda telah tersimpan otomatis dengan kode referensi:</p>' +
        '<div style="background:#ecfdf5;border:1px solid #a7f3d0;padding:14px 18px;border-radius:10px;font-family:monospace;font-size:18px;font-weight:700;color:#047857;margin:16px 0;">' + kodeRegistrasi + '</div>' +
        '<table style="width:100%;border-collapse:collapse;border:1px solid #dce6e0;">' + barisTabel + '</table>' +
      '</div>' +
    '</div>';

  MailApp.sendEmail({
    to: emailTujuan,
    subject: subjek,
    htmlBody: htmlBody,
  });
}

// Opsional: Jalankan fungsi ini untuk menyinkronkan dropdown "Nama Pondok" & "Lomba" pada Google Form asli
function sinkronkanDropdownGoogleForm() {
  var GOOGLE_FORM_LOMBA_ID = 'ISI_DENGAN_ID_EDIT_GOOGLE_FORM_LOMBA';
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var daftarPonpes = ambilKolomDariSheet_(ss, ['DB ponpes', 'DB Ponpes', 'db ponpes'], ['Nama Lengkap Pondok', 'Nama Pondok', 'Pondok']);
  var daftarLomba = ambilKolomDariSheet_(ss, ['db lomba', 'DB lomba', 'DB Lomba'], ['Nama Lomba', 'Lomba', 'Cabang Lomba']);

  var form = FormApp.openById(GOOGLE_FORM_LOMBA_ID);
  var items = form.getItems();
  items.forEach(function(item) {
    if (item.getTitle() === 'Nama Pondok' && daftarPonpes.length > 0) {
      item.asListItem().setChoiceValues(daftarPonpes);
    }
    if (item.getTitle() === 'Lomba' && daftarLomba.length > 0) {
      item.asMultipleChoiceItem().setChoiceValues(daftarLomba);
    }
  });
}

function ambilKolomDariSheet_(ss, sheetCandidates, headerCandidates) {
  var sheets = ss.getSheets();
  var targetSheet = null;
  for (var c = 0; c < sheetCandidates.length && !targetSheet; c++) {
    for (var i = 0; i < sheets.length; i++) {
      if (sheets[i].getName().toLowerCase().trim() === sheetCandidates[c].toLowerCase()) {
        targetSheet = sheets[i];
        break;
      }
    }
  }
  if (!targetSheet || targetSheet.getLastRow() < 2) return [];
  var data = targetSheet.getDataRange().getValues();
  var headers = data[0].map(function(h) { return String(h).trim().toLowerCase(); });
  var colIdx = -1;
  for (var h = 0; h < headerCandidates.length; h++) {
    var idx = headers.indexOf(headerCandidates[h].toLowerCase());
    if (idx !== -1) { colIdx = idx; break; }
  }
  if (colIdx === -1) colIdx = headers[0] === 'timestamp' ? 1 : 0;
  var seen = {};
  var out = [];
  for (var r = 1; r < data.length; r++) {
    var val = String(data[r][colIdx] || '').trim();
    if (val && !seen[val.toLowerCase()]) {
      seen[val.toLowerCase()] = true;
      out.push(val);
    }
  }
  return out;
}

function escapeHtml(text) {
  return String(text || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}`;

const APPS_SCRIPT_WEBAPP_DOPOST = `/**
 * SKRIP B: Web App (doGet + doPost) untuk 1 Spreadsheet (Database_Mabahis)
 * - doGet(): Mengambil daftar dropdown dari sheet "DB ponpes" dan "db lomba"
 * - doPost(): Menyimpan Daftar Pondok ke "DB ponpes", Lomba ke "Pendaftaran Lomba MABAHIS",
 *             dan langsung mengirim email konfirmasi otomatis ke pengirim.
 * Deploy sebagai Web App: Execute as "Me", Who has access "Anyone"
 */

function findSheetIgnoreCase_(ss, candidates) {
  var sheets = ss.getSheets();
  for (var c = 0; c < candidates.length; c++) {
    var target = candidates[c].toLowerCase();
    for (var i = 0; i < sheets.length; i++) {
      if (sheets[i].getName().toLowerCase().trim() === target) {
        return sheets[i];
      }
    }
  }
  return null;
}

function extractColumnValues_(ss, sheetCandidates, headerCandidates) {
  var sheet = findSheetIgnoreCase_(ss, sheetCandidates);
  if (!sheet || sheet.getLastRow() < 2) return [];

  var data = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h) { return String(h).trim().toLowerCase(); });

  // Cari kolom yang cocok dengan nama header yang diharapkan
  var colIdx = -1;
  for (var h = 0; h < headerCandidates.length; h++) {
    var idx = headers.indexOf(headerCandidates[h].toLowerCase());
    if (idx !== -1) {
      colIdx = idx;
      break;
    }
  }

  // Fallback jika header beda nama: pilih kolom pertama yang bukan Timestamp/No
  if (colIdx === -1) {
    for (var k = 0; k < headers.length; k++) {
      if (headers[k] !== 'timestamp' && headers[k] !== 'no' && headers[k] !== 'waktu kirim') {
        colIdx = k;
        break;
      }
    }
    if (colIdx === -1) colIdx = 0;
  }

  var seen = {};
  var results = [];
  for (var r = 1; r < data.length; r++) {
    var val = String(data[r][colIdx] || '').trim();
    if (val && !seen[val.toLowerCase()]) {
      seen[val.toLowerCase()] = true;
      results.push(val);
    }
  }
  return results;
}

function doGet() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ponpes = extractColumnValues_(
      ss,
      ['DB ponpes', 'DB Ponpes', 'Daftar Pondok'],
      ['Nama Lengkap Pondok', 'Nama Pondok', 'Pondok']
    );
    var lomba = extractColumnValues_(
      ss,
      ['db lomba', 'DB lomba', 'DB Lomba', 'Lomba'],
      ['Nama Lomba', 'Lomba', 'Cabang Lomba']
    );

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, ponpes: ponpes, lomba: lomba }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message, ponpes: [], lomba: [] }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var answers = payload.answers || {};
    var isLomba = payload.formId === 'lomba-mabahis';

    var sheet = isLomba
      ? (findSheetIgnoreCase_(ss, ['Pendaftaran Lomba MABAHIS']) || ss.insertSheet('Pendaftaran Lomba MABAHIS'))
      : (findSheetIgnoreCase_(ss, ['DB ponpes', 'DB Ponpes']) || ss.insertSheet('DB ponpes'));

    if (sheet.getLastRow() === 0) {
      if (isLomba) {
        sheet.appendRow([
          'Timestamp',
          'Nama Peserta',
          'Nama Pondok',
          'Lomba',
          'Kategori',
          'NO WA Penanggung Jawab',
          'Email Pengirim',
          'Kode Registrasi'
        ]);
      } else {
        sheet.appendRow([
          'Timestamp',
          'Nama Lengkap Pondok',
          'Alamat Pondok',
          'Desa',
          'Kecamatan',
          'Nama Pengasuh',
          'Email Pengirim',
          'Kode Registrasi'
        ]);
      }
    }

    var kode = payload.registrationCode || (
      (isLomba ? 'MBH-' : 'PDK-') +
      Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyMMdd') + '-' +
      Math.floor(1000 + Math.random() * 9000)
    );

    if (isLomba) {
      sheet.appendRow([
        new Date(),
        answers['Nama Peserta'] || '-',
        answers['Nama Pondok'] || '-',
        answers['Lomba'] || '-',
        answers['Kategori'] || '-',
        answers['NO WA Penanggung Jawab'] || '-',
        payload.email || '-',
        kode
      ]);
    } else {
      sheet.appendRow([
        new Date(),
        answers['Nama Lengkap Pondok'] || '-',
        answers['Alamat Pondok'] || '-',
        answers['Desa'] || '-',
        answers['Kecamatan'] || '-',
        answers['Nama Pengasuh'] || '-',
        payload.email || '-',
        kode
      ]);
    }

    // Kirim email konfirmasi otomatis ke pengirim
    if (payload.email && String(payload.email).indexOf('@') !== -1) {
      var subjek = 'Konfirmasi [' + kode + '] - ' + (payload.formTitle || sheet.getName());
      var sapaan = isLomba
        ? (answers['Nama Peserta'] || 'Peserta')
        : (answers['Nama Pengasuh'] || answers['Nama Lengkap Pondok'] || 'Pengirim');

      var rowsHtml = Object.keys(answers).map(function(k) {
        return '<tr><td style="padding:8px 12px;border-bottom:1px solid #dce6e0;color:#5f766b;font-size:13px;">' + escapeHtml(k) + '</td><td style="padding:8px 12px;border-bottom:1px solid #dce6e0;color:#0b1f17;font-size:13px;font-weight:500;">' + escapeHtml(answers[k]) + '</td></tr>';
      }).join('');

      var htmlBody =
        '<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;border:1px solid #dce6e0;border-radius:12px;overflow:hidden;">' +
          '<div style="background:#064e3b;color:#ffffff;padding:24px;">' +
            '<div style="font-size:12px;color:#a7f3d0;margin-bottom:4px;">MABAHIS · Konfirmasi Otomatis</div>' +
            '<h2 style="margin:0;font-size:20px;">' + escapeHtml(payload.formTitle || sheet.getName()) + '</h2>' +
          '</div>' +
          '<div style="padding:24px;">' +
            '<p style="font-size:15px;color:#0b1f17;">Halo <strong>' + escapeHtml(sapaan) + '</strong>,</p>' +
            '<p style="font-size:14px;color:#3f574c;line-height:1.6;">' + escapeHtml(payload.emailOpening || 'Data pendaftaran Anda telah tersimpan ke Database_Mabahis.') + '</p>' +
            '<div style="background:#ecfdf5;border:1px solid #a7f3d0;padding:14px 18px;border-radius:10px;font-family:monospace;font-size:18px;font-weight:700;color:#047857;margin:16px 0;">' + escapeHtml(kode) + '</div>' +
            '<table style="width:100%;border-collapse:collapse;border:1px solid #dce6e0;">' + rowsHtml + '</table>' +
          '</div>' +
        '</div>';

      MailApp.sendEmail({
        to: payload.email,
        subject: subjek,
        htmlBody: htmlBody,
      });
    }

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, registrationCode: kode, sheetTab: sheet.getName(), emailSent: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function escapeHtml(text) {
  return String(text || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}`;

// ============================================================================
// STATE & HELPERS
// ============================================================================
let state = {
  activeView: 'portal',
  selectedCategory: 'Semua',
  searchQuery: '',
  selectedFormId: null,
  formMode: 'html', // 'html' | 'google-form'
  activeScriptTab: 'webapp-dopost',
  config: loadConfig(),
  history: loadHistory(),
  localPonpes: loadLocalPonpes(),
  dropdownData: {
    ponpes: [],
    lomba: [],
    loading: false,
    loaded: false,
    error: '',
    source: 'local', // 'spreadsheet' | 'local'
  },
};

function loadConfig() {
  try {
    const raw = localStorage.getItem(STORAGE_CONFIG_KEY);
    if (!raw) return structuredClone(DEFAULT_CONFIG);
    const parsed = JSON.parse(raw);
    return {
      appsScriptWebAppUrl: parsed.appsScriptWebAppUrl || DEFAULT_CONFIG.appsScriptWebAppUrl,
      googleFormUrls: {
        ...DEFAULT_CONFIG.googleFormUrls,
        ...(parsed.googleFormUrls || {}),
      },
    };
  } catch {
    return structuredClone(DEFAULT_CONFIG);
  }
}

function saveConfig() {
  try {
    localStorage.setItem(STORAGE_CONFIG_KEY, JSON.stringify(state.config));
  } catch {
    // Ignore storage error
  }
}

function loadHistory() {
  try {
    const raw = localStorage.getItem(STORAGE_HISTORY_KEY);
    const parsed = raw ? JSON.parse(raw) : [];
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function saveHistory() {
  try {
    localStorage.setItem(STORAGE_HISTORY_KEY, JSON.stringify(state.history.slice(0, 30)));
  } catch {
    // Ignore storage error
  }
}

function loadLocalPonpes() {
  try {
    const raw = localStorage.getItem(STORAGE_LOCAL_PONPES_KEY);
    const parsed = raw ? JSON.parse(raw) : [];
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function addLocalPonpes(name) {
  const clean = String(name || '').trim();
  if (!clean) return;
  if (!state.localPonpes.some((p) => p.toLowerCase() === clean.toLowerCase())) {
    state.localPonpes.push(clean);
    try {
      localStorage.setItem(STORAGE_LOCAL_PONPES_KEY, JSON.stringify(state.localPonpes));
    } catch {
      // Ignore storage error
    }
  }
}

function escapeHtml(str) {
  return String(str || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function generateRegistrationCode(prefix = 'REG') {
  const now = new Date();
  const yy = String(now.getFullYear()).slice(-2);
  const mm = String(now.getMonth() + 1).padStart(2, '0');
  const dd = String(now.getDate()).padStart(2, '0');
  const rand = Math.random().toString(36).substring(2, 6).toUpperCase();
  return `${prefix}-${yy}${mm}${dd}-${rand}`;
}

function normalizeGoogleFormEmbedUrl(url) {
  const trimmed = String(url || '').trim();
  if (!trimmed) return '';
  try {
    const u = new URL(trimmed);
    if (!u.hostname.includes('docs.google.com')) return '';
    u.searchParams.set('embedded', 'true');
    return u.toString();
  } catch {
    return '';
  }
}

function showToast(message) {
  const toast = document.getElementById('toast');
  if (!toast) return;
  toast.textContent = message;
  toast.classList.remove('is-hidden');
  clearTimeout(showToast._timer);
  showToast._timer = setTimeout(() => {
    toast.classList.add('is-hidden');
  }, 3200);
}

async function loadDropdownOptions(force = false) {
  const scriptUrl = state.config.appsScriptWebAppUrl.trim();
  if (!scriptUrl) {
    state.dropdownData = {
      ponpes: [...state.localPonpes],
      lomba: [],
      loading: false,
      loaded: true,
      error: '',
      source: 'local',
    };
    return;
  }

  if (state.dropdownData.loaded && !force && state.dropdownData.source === 'spreadsheet') {
    return;
  }

  state.dropdownData.loading = true;
  state.dropdownData.error = '';

  const form = FORMS.find((f) => f.id === state.selectedFormId);
  if (form && state.activeView === 'form') {
    renderFormWorkspace(form);
  }

  try {
    const res = await fetch(scriptUrl, { method: 'GET' });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    const data = await res.json();
    if (!data || data.ok === false) {
      throw new Error(data?.error || 'Respons Apps Script tidak valid.');
    }

    const ponpesList = Array.isArray(data.ponpes) ? data.ponpes : [];
    const lombaList = Array.isArray(data.lomba) ? data.lomba : [];

    state.dropdownData = {
      ponpes: ponpesList,
      lomba: lombaList,
      loading: false,
      loaded: true,
      error: '',
      source: 'spreadsheet',
    };
  } catch (err) {
    state.dropdownData = {
      ponpes: [...state.localPonpes],
      lomba: [],
      loading: false,
      loaded: true,
      error:
        'Gagal membaca sheet DB ponpes & db lomba dari URL Apps Script. Pastikan Skrip B terbaru sudah di-deploy ulang.',
      source: 'local',
    };
  }

  const activeForm = FORMS.find((f) => f.id === state.selectedFormId);
  if (activeForm && state.activeView === 'form') {
    renderFormWorkspace(activeForm);
  }
}

function buildConfirmationEmailHtml(form, payload) {
  const rows = Object.entries(payload.answers || {})
    .map(
      ([k, v]) =>
        `<tr>
          <td style="padding:9px 12px;border-bottom:1px solid #dce6e0;color:#5f766b;font-size:13px;width:42%;">${escapeHtml(k)}</td>
          <td style="padding:9px 12px;border-bottom:1px solid #dce6e0;color:#0b1f17;font-size:13px;font-weight:500;">${escapeHtml(v || '-')}</td>
        </tr>`
    )
    .join('');

  return `<!doctype html>
<html lang="id">
<head><meta charset="utf-8" /></head>
<body style="margin:0;padding:24px;background:#f7faf8;font-family:Arial,sans-serif;color:#0b1f17;">
  <div style="max-width:580px;margin:0 auto;background:#ffffff;border:1px solid #dce6e0;border-radius:12px;overflow:hidden;">
    <div style="background:#064e3b;color:#ffffff;padding:24px 28px;">
      <div style="font-size:12px;color:#a7f3d0;margin-bottom:4px;">MABAHIS · Konfirmasi Pendaftaran Otomatis</div>
      <h1 style="margin:0;font-size:19px;font-weight:600;">${escapeHtml(form.title)}</h1>
    </div>
    <div style="padding:24px 28px;">
      <p style="margin:0 0 12px;font-size:15px;color:#0b1f17;">Halo <strong>${escapeHtml(payload.recipientName)}</strong>,</p>
      <p style="margin:0 0 18px;font-size:14px;color:#3f574c;line-height:1.6;">${escapeHtml(form.emailOpening)}</p>
      <div style="background:#ecfdf5;border:1px solid #a7f3d0;border-radius:10px;padding:14px 18px;margin-bottom:20px;">
        <div style="font-size:12px;color:#065f46;margin-bottom:4px;">Kode Registrasi Resmi</div>
        <div style="font-family:monospace;font-size:18px;font-weight:700;color:#047857;">${escapeHtml(payload.registrationCode)}</div>
      </div>
      <table style="width:100%;border-collapse:collapse;border:1px solid #dce6e0;">
        ${rows}
      </table>
    </div>
  </div>
</body>
</html>`;
}

// ============================================================================
// NAVIGATION & RENDERING
// ============================================================================
function setView(viewName) {
  state.activeView = viewName;

  ['portal', 'form', 'guide', 'integrations', 'history'].forEach((name) => {
    const el = document.getElementById(`view-${name}`);
    if (el) el.classList.toggle('is-hidden', name !== viewName);
  });

  document.querySelectorAll('[data-nav]').forEach((btn) => {
    const target = btn.getAttribute('data-nav');
    const isActive = target === viewName || (viewName === 'form' && target === 'portal');
    btn.classList.toggle('is-active', isActive);
  });

  if (viewName === 'guide') renderAppsScriptBlock();
  if (viewName === 'integrations') renderIntegrationsView();
  if (viewName === 'history') renderHistoryView();

  window.scrollTo({ top: 0, behavior: 'smooth' });
}

function renderCatalog() {
  const categories = ['Semua', ...new Set(FORMS.map((f) => f.category))];
  const filterBar = document.getElementById('category-filters');
  if (filterBar) {
    filterBar.innerHTML = categories
      .map(
        (cat) => `
        <button
          type="button"
          class="filter-tab ${state.selectedCategory === cat ? 'is-active' : ''}"
          data-category="${escapeHtml(cat)}"
        >${escapeHtml(cat)}</button>`
      )
      .join('');

    filterBar.querySelectorAll('.filter-tab').forEach((btn) => {
      btn.addEventListener('click', () => {
        state.selectedCategory = btn.getAttribute('data-category') || 'Semua';
        renderCatalog();
      });
    });
  }

  const q = state.searchQuery.trim().toLowerCase();
  const filtered = FORMS.filter((f) => {
    const catOk = state.selectedCategory === 'Semua' || f.category === state.selectedCategory;
    const queryOk =
      !q ||
      f.title.toLowerCase().includes(q) ||
      f.description.toLowerCase().includes(q) ||
      f.category.toLowerCase().includes(q);
    return catOk && queryOk;
  });

  const grid = document.getElementById('catalog-grid');
  if (!grid) return;

  if (filtered.length === 0) {
    grid.innerHTML = `
      <div class="empty-state" style="grid-column: 1 / -1;">
        <h2>Formulir Tidak Ditemukan</h2>
        <p>Tidak ada formulir yang cocok dengan kata kunci pencarian Anda.</p>
        <button type="button" class="btn btn-dark btn-sm" id="btn-reset-catalog">Reset Pencarian</button>
      </div>`;
    document.getElementById('btn-reset-catalog')?.addEventListener('click', () => {
      state.selectedCategory = 'Semua';
      state.searchQuery = '';
      const searchInput = document.getElementById('catalog-search');
      if (searchInput) searchInput.value = '';
      renderCatalog();
    });
    return;
  }

  grid.innerHTML = filtered
    .map((form, index) => {
      const numLabel = String(index + 1).padStart(2, '0');
      return `
      <article class="form-card">
        <div>
          <div class="form-card-media">
            <img
              src="${escapeHtml(form.image)}"
              alt="${escapeHtml(form.title)}"
              referrerpolicy="no-referrer"
              onerror="this.style.display='none';this.nextElementSibling.classList.remove('is-hidden');"
            />
            <div class="img-fallback is-hidden">${escapeHtml(form.title)}</div>
          </div>
          <div class="form-card-body">
            <div class="meta-line">
              <span class="meta-index">${numLabel}</span>
              <span aria-hidden="true">·</span>
              <span class="meta-category">${escapeHtml(form.category)}</span>
            </div>
            <h3 class="form-card-title">${escapeHtml(form.title)}</h3>
            <p class="form-card-desc">${escapeHtml(form.description)}</p>
          </div>
        </div>
        <div class="form-card-footer">
          <button type="button" class="btn btn-dark btn-card" data-open-form="${escapeHtml(form.id)}">
            <span>Buka Formulir</span>
            <svg class="btn-arrow" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true">
              <path d="M5 12h14"></path>
              <path d="m12 5 7 7-7 7"></path>
            </svg>
          </button>
        </div>
      </article>`;
    })
    .join('');

  grid.querySelectorAll('[data-open-form]').forEach((btn) => {
    btn.addEventListener('click', () => {
      const id = btn.getAttribute('data-open-form');
      if (id) openFormWorkspace(id);
    });
  });
}

function openFormWorkspace(formId) {
  const form = FORMS.find((f) => f.id === formId);
  if (!form) return;
  state.selectedFormId = formId;
  state.formMode = 'html';
  renderFormWorkspace(form);
  setView('form');

  if (formId === 'lomba-mabahis') {
    void loadDropdownOptions(false);
  }
}

function renderDynamicSelectField(field) {
  const sourceKey = field.dynamicSource; // 'ponpes' | 'lomba'
  const options = sourceKey === 'ponpes' ? state.dropdownData.ponpes : state.dropdownData.lomba;
  const isLoading = state.dropdownData.loading;

  let placeholderText = 'Pilih salah satu...';
  if (isLoading) {
    placeholderText =
      sourceKey === 'ponpes' ? 'Memuat daftar pondok...' : 'Memuat daftar lomba...';
  } else if (options.length === 0) {
    placeholderText =
      sourceKey === 'ponpes'
        ? 'Daftar pondok belum tersedia'
        : 'Daftar lomba belum tersedia';
  }

  const optionsHtml = [
    `<option value="">${escapeHtml(placeholderText)}</option>`,
    ...options.map((opt) => `<option value="${escapeHtml(opt)}">${escapeHtml(opt)}</option>`),
  ].join('');

  const noteHtml = field.note
    ? `<div class="text-muted" style="margin-top:4px;font-size:12px;">${escapeHtml(field.note)}</div>`
    : '';

  return `
  <div class="field-group">
    <label class="field-label" for="fld-${escapeHtml(field.id)}">
      ${escapeHtml(field.label)} ${field.required ? '<span class="req">*</span>' : ''}
    </label>
    <select class="field-select" id="fld-${escapeHtml(field.id)}" data-field-label="${escapeHtml(field.label)}" ${field.required ? 'required' : ''}>
      ${optionsHtml}
    </select>
    ${noteHtml}
  </div>`;
}

function renderFormWorkspace(form, submittedReceipt = null) {
  const container = document.getElementById('form-workspace-content');
  if (!container) return;

  if (submittedReceipt) {
    const isLive = submittedReceipt.deliveryMode === 'apps-script';
    container.innerHTML = `
      <div class="receipt-card">
        <div class="receipt-head">
          <div class="receipt-status-tag">
            ${isLive ? 'Pendaftaran Terkirim' : 'Pendaftaran Tercatat'}
          </div>
          <h2 class="section-title">${escapeHtml(form.title)}</h2>
          <p class="section-subtitle">
            Terima kasih, <strong>${escapeHtml(submittedReceipt.recipientName)}</strong>. Data pendaftaran Anda telah diproses.
          </p>
        </div>

        <div class="receipt-meta-grid">
          <div>
            <div class="text-muted mb-4">Kode Registrasi</div>
            <div class="code-badge">${escapeHtml(submittedReceipt.registrationCode)}</div>
          </div>
          <div>
            <div class="text-muted mb-4">Email Konfirmasi</div>
            <div style="font-weight: 600;">${escapeHtml(submittedReceipt.email)}</div>
          </div>
        </div>

        <div class="receipt-actions">
          <button type="button" class="btn btn-primary" id="btn-preview-receipt-email">
            Lihat Bukti Email
          </button>
          <button type="button" class="btn btn-ghost" id="btn-reset-form">
            Isi Formulir Lagi
          </button>
        </div>
      </div>`;

    document.getElementById('btn-preview-receipt-email')?.addEventListener('click', () => {
      openEmailModal(
        submittedReceipt.email,
        `Konfirmasi [${submittedReceipt.registrationCode}] - ${form.title}`,
        submittedReceipt.htmlPreview,
        isLive ? 'Terkirim ke Email' : 'Pratinjau Bukti Email'
      );
    });
    document.getElementById('btn-reset-form')?.addEventListener('click', () => {
      renderFormWorkspace(form, null);
    });
    return;
  }

  const embedUrl = normalizeGoogleFormEmbedUrl(state.config.googleFormUrls[form.id]);

  const fieldsHtml = form.fields
    .map((field) => {
      if (field.type === 'select' && field.dynamicSource) {
        return renderDynamicSelectField(field);
      }

      const noteHtml = field.note
        ? `<div class="text-muted" style="margin-top:4px;font-size:12px;">${escapeHtml(field.note)}</div>`
        : '';

      if (field.type === 'select' && field.options) {
        return `
        <div class="field-group">
          <label class="field-label" for="fld-${escapeHtml(field.id)}">
            ${escapeHtml(field.label)} ${field.required ? '<span class="req">*</span>' : ''}
          </label>
          <select class="field-select" id="fld-${escapeHtml(field.id)}" data-field-label="${escapeHtml(field.label)}" ${field.required ? 'required' : ''}>
            ${field.options.map((opt) => `<option value="${escapeHtml(opt)}">${escapeHtml(opt)}</option>`).join('')}
          </select>
          ${noteHtml}
        </div>`;
      }

      if (field.type === 'textarea') {
        return `
        <div class="field-group">
          <label class="field-label" for="fld-${escapeHtml(field.id)}">
            ${escapeHtml(field.label)} ${field.required ? '<span class="req">*</span>' : ''}
          </label>
          <textarea class="field-textarea" id="fld-${escapeHtml(field.id)}" data-field-label="${escapeHtml(field.label)}" placeholder="${escapeHtml(field.placeholder || '')}" ${field.required ? 'required' : ''}></textarea>
          ${noteHtml}
        </div>`;
      }

      return `
      <div class="field-group">
        <label class="field-label" for="fld-${escapeHtml(field.id)}">
          ${escapeHtml(field.label)} ${field.required ? '<span class="req">*</span>' : ''}
        </label>
        <input class="field-input" type="${escapeHtml(field.type || 'text')}" id="fld-${escapeHtml(field.id)}" data-field-label="${escapeHtml(field.label)}" placeholder="${escapeHtml(field.placeholder || '')}" ${field.required ? 'required' : ''} />
        ${noteHtml}
      </div>`;
    })
    .join('');

  const dropdownToolbarHtml =
    form.id === 'lomba-mabahis' && (state.dropdownData.loading || state.dropdownData.error)
      ? `
      <div style="display:flex;flex-wrap:wrap;align-items:center;justify-content:space-between;gap:10px;padding:10px 14px;background:#f7faf8;border:1px solid #dce6e0;border-radius:10px;margin-bottom:18px;">
        <div style="font-size:12px;color:#3f574c;">
          ${
            state.dropdownData.loading
              ? 'Memuat daftar pondok dan lomba...'
              : `<span style="color:#b91c1c;">Gagal memuat daftar pondok &amp; lomba.</span>`
          }
        </div>
        <button type="button" class="btn btn-secondary btn-sm" id="btn-refresh-dropdowns" ${state.dropdownData.loading ? 'disabled' : ''}>
          Muat Ulang
        </button>
      </div>`
      : '';

  const rightPanelContent =
    state.formMode === 'google-form'
      ? embedUrl
        ? `<iframe class="gform-iframe" src="${escapeHtml(embedUrl)}" title="Google Form ${escapeHtml(form.title)}">Memuat Google Form...</iframe>`
        : `
          <div class="gform-placeholder">
            <h3>Tautan Google Form Belum Dipasang</h3>
            <p class="text-muted mb-4">
              Tempelkan URL Google Form Anda di menu Hubungkan Google.
            </p>
          </div>`
      : `
        ${dropdownToolbarHtml}
        <div id="form-error-box" class="alert alert-error is-hidden" role="alert"></div>
        <form id="registration-html-form">
          ${fieldsHtml}

          <div class="field-group" style="padding-top:14px;border-top:1px solid #dce6e0;">
            <label class="field-label" for="reg-email">
              Alamat Email <span class="req">*</span>
            </label>
            <input class="field-input" id="reg-email" type="email" required maxlength="160" placeholder="nama@gmail.com" />
            <div class="text-muted" style="margin-top:4px;font-size:12px;">
              Bukti pendaftaran dikirim otomatis ke email ini.
            </div>
          </div>

          <div class="form-submit-bar">
            <button type="submit" class="btn btn-primary btn-lg" id="btn-submit-form">
              <span>Kirim Pendaftaran</span>
              <svg class="btn-arrow" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" aria-hidden="true">
                <path d="M5 12h14"></path>
                <path d="m12 5 7 7-7 7"></path>
              </svg>
            </button>
          </div>
        </form>`;

  container.innerHTML = `
    <div class="workspace-grid">
      <aside class="workspace-info">
        <div class="form-card-media">
          <img
            src="${escapeHtml(form.image)}"
            alt="${escapeHtml(form.title)}"
            referrerpolicy="no-referrer"
            onerror="this.style.display='none';this.nextElementSibling.classList.remove('is-hidden');"
          />
          <div class="img-fallback is-hidden">${escapeHtml(form.title)}</div>
        </div>
        <div class="form-card-body">
          <div class="meta-line">
            <span class="meta-category">${escapeHtml(form.category)}</span>
          </div>
          <h1 class="form-card-title">${escapeHtml(form.title)}</h1>
          <p class="form-card-desc">${escapeHtml(form.description)}</p>
        </div>
      </aside>

      <div class="workspace-form-card">
        <div class="mode-switcher" role="tablist">
          <button type="button" class="seg-btn ${state.formMode === 'html' ? 'is-active' : ''}" data-form-mode="html">
            Formulir Website
          </button>
          <button type="button" class="seg-btn ${state.formMode === 'google-form' ? 'is-active' : ''}" data-form-mode="google-form">
            Google Form
          </button>
        </div>

        ${rightPanelContent}
      </div>
    </div>`;

  container.querySelectorAll('[data-form-mode]').forEach((btn) => {
    btn.addEventListener('click', () => {
      state.formMode = btn.getAttribute('data-form-mode') || 'html';
      renderFormWorkspace(form);
    });
  });

  document.getElementById('btn-refresh-dropdowns')?.addEventListener('click', () => {
    if (!state.config.appsScriptWebAppUrl.trim()) {
      showToast('Pasang URL Web App Apps Script di menu Hubungkan Google terlebih dahulu.');
      return;
    }
    void loadDropdownOptions(true);
  });

  const htmlForm = document.getElementById('registration-html-form');
  if (htmlForm) {
    htmlForm.addEventListener('submit', async (e) => {
      e.preventDefault();
      const errBox = document.getElementById('form-error-box');
      const submitBtn = document.getElementById('btn-submit-form');
      if (errBox) errBox.classList.add('is-hidden');

      const email = document.getElementById('reg-email').value.trim();
      if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
        if (errBox) {
          errBox.textContent = 'Masukkan alamat email yang valid untuk menerima notifikasi otomatis.';
          errBox.classList.remove('is-hidden');
        }
        return;
      }

      const answers = {};
      let missingRequired = '';
      htmlForm.querySelectorAll('[data-field-label]').forEach((el) => {
        const label = el.getAttribute('data-field-label');
        if (label) {
          const val = el.value.trim();
          answers[label] = val;
          if (el.hasAttribute('required') && !val && !missingRequired) {
            missingRequired = label;
          }
        }
      });

      if (missingRequired) {
        if (errBox) {
          errBox.textContent = `Kolom "${missingRequired}" wajib diisi atau dipilih terlebih dahulu.`;
          errBox.classList.remove('is-hidden');
        }
        return;
      }

      const recipientName =
        answers['Nama Peserta'] ||
        answers['Nama Pengasuh'] ||
        answers['Nama Lengkap Pondok'] ||
        'Pengirim';

      const registrationCode = generateRegistrationCode(form.codePrefix);
      const payload = {
        registrationCode,
        formId: form.id,
        formTitle: form.title,
        email,
        answers,
        emailOpening: form.emailOpening,
      };

      if (submitBtn) {
        submitBtn.disabled = true;
        const lbl = submitBtn.querySelector('span');
        if (lbl) lbl.textContent = 'Mengirim Pendaftaran...';
      }

      let deliveryMode = 'local-preview';
      const scriptUrl = state.config.appsScriptWebAppUrl.trim();

      if (scriptUrl) {
        try {
          const response = await fetch(scriptUrl, {
            method: 'POST',
            headers: { 'Content-Type': 'text/plain;charset=utf-8' },
            body: JSON.stringify(payload),
          });
          if (!response.ok) {
            throw new Error(`Status HTTP ${response.status}`);
          }
          deliveryMode = 'apps-script';
          if (form.id === 'daftar-pondok') {
            state.dropdownData.loaded = false;
          }
        } catch {
          if (errBox) {
            errBox.textContent =
              'Gagal terhubung ke URL Google Apps Script. Pastikan URL benar dan akses deployment diatur ke "Anyone".';
            errBox.classList.remove('is-hidden');
          }
          if (submitBtn) {
            submitBtn.disabled = false;
            const lbl = submitBtn.querySelector('span');
            if (lbl) lbl.textContent = 'Kirim Pendaftaran';
          }
          return;
        }
      }

      if (form.id === 'daftar-pondok' && answers['Nama Lengkap Pondok']) {
        addLocalPonpes(answers['Nama Lengkap Pondok']);
      }

      const htmlPreview = buildConfirmationEmailHtml(form, {
        ...payload,
        recipientName,
      });

      const receipt = {
        id: `rec_${Date.now()}`,
        registrationCode,
        formId: form.id,
        formTitle: form.title,
        recipientName,
        email,
        deliveryMode,
        htmlPreview,
        submittedAt: new Date().toISOString(),
      };

      state.history.unshift(receipt);
      saveHistory();
      renderFormWorkspace(form, receipt);
    });
  }
}

function renderAppsScriptBlock() {
  const codeEl = document.getElementById('apps-script-code');
  const labelEl = document.getElementById('active-script-filename');
  const isFormTrigger = state.activeScriptTab === 'form-trigger';

  if (codeEl) {
    codeEl.textContent = isFormTrigger
      ? APPS_SCRIPT_FORM_TRIGGER
      : APPS_SCRIPT_WEBAPP_DOPOST;
  }
  if (labelEl) {
    labelEl.textContent = isFormTrigger
      ? 'Code.gs (Skrip A · Trigger Google Form + Sync Dropdown)'
      : 'Code.gs (Skrip B · Web App doGet DB ponpes & db lomba + doPost)';
  }

  document.querySelectorAll('#script-tabs .seg-btn').forEach((btn) => {
    btn.classList.toggle(
      'is-active',
      btn.getAttribute('data-script') === state.activeScriptTab
    );
  });
}

function renderIntegrationsView() {
  const scriptInput = document.getElementById('input-apps-script-url');
  const statusEl = document.getElementById('script-url-status');
  if (scriptInput) scriptInput.value = state.config.appsScriptWebAppUrl;
  if (statusEl) {
    statusEl.textContent = state.config.appsScriptWebAppUrl
      ? `Terhubung ke: ${state.config.appsScriptWebAppUrl}`
      : 'Belum ada URL Web App yang diisi. Pasang URL dari Deploy Web App Skrip B agar dropdown Nama Pondok & Lomba tertarik otomatis dari Spreadsheet.';
  }

  const listEl = document.getElementById('google-form-links-list');
  if (listEl) {
    listEl.innerHTML = FORMS.map(
      (f) => `
      <div class="gform-link-item">
        <div>
          <div style="font-size:13px;font-weight:600;color:#0b1f17;">${escapeHtml(f.title)}</div>
          <div class="text-muted">${escapeHtml(f.category)}</div>
        </div>
        <input
          type="url"
          data-gform-id="${escapeHtml(f.id)}"
          value="${escapeHtml(state.config.googleFormUrls[f.id] || '')}"
          placeholder="https://docs.google.com/forms/d/e/.../viewform"
        />
      </div>`
    ).join('');
  }

  const snippetEl = document.getElementById('exported-config-snippet');
  if (snippetEl) {
    snippetEl.textContent = `const DEFAULT_CONFIG = ${JSON.stringify(state.config, null, 2)};`;
  }
}

function renderHistoryView() {
  const listEl = document.getElementById('history-list');
  if (!listEl) return;

  if (state.history.length === 0) {
    listEl.innerHTML = `
      <div class="empty-state">
        <h2>Belum Ada Riwayat Kiriman</h2>
        <p>Kirimkan salah satu formulir pendaftaran terlebih dahulu untuk melihat bukti kode registrasi dan salinan email otomatisnya.</p>
        <button type="button" class="btn btn-primary" id="btn-empty-go-portal">Lihat Daftar Formulir</button>
      </div>`;
    document.getElementById('btn-empty-go-portal')?.addEventListener('click', () => {
      setView('portal');
    });
    return;
  }

  listEl.innerHTML = state.history
    .map((item) => {
      const dateStr = new Date(item.submittedAt).toLocaleString('id-ID', {
        dateStyle: 'medium',
        timeStyle: 'short',
      });
      const modeLabel =
        item.deliveryMode === 'apps-script'
          ? 'Terkirim ke Database_Mabahis & Email'
          : 'Simulasi Uji Lokal';

      return `
      <div class="history-item">
        <div>
          <div class="meta-line">
            <span class="meta-category" style="font-family:var(--font-mono);">${escapeHtml(item.registrationCode)}</span>
            <span aria-hidden="true">·</span>
            <span>${escapeHtml(dateStr)}</span>
            <span aria-hidden="true">·</span>
            <span>${escapeHtml(modeLabel)}</span>
          </div>
          <h3 style="font-size:17px;font-weight:700;margin-bottom:4px;">${escapeHtml(item.formTitle)}</h3>
          <div class="text-muted">${escapeHtml(item.recipientName || item.fullName)} · ${escapeHtml(item.email)}</div>
        </div>
        <div class="history-actions">
          <button type="button" class="btn btn-secondary btn-sm" data-preview-history="${escapeHtml(item.id)}">
            Lihat Email Otomatis
          </button>
        </div>
      </div>`;
    })
    .join('');

  listEl.querySelectorAll('[data-preview-history]').forEach((btn) => {
    btn.addEventListener('click', () => {
      const id = btn.getAttribute('data-preview-history');
      const item = state.history.find((h) => h.id === id);
      if (!item) return;
      openEmailModal(
        item.email,
        `Konfirmasi [${item.registrationCode}] - ${item.formTitle}`,
        item.htmlPreview,
        item.deliveryMode === 'apps-script'
          ? 'Dikirim via Google Apps Script'
          : 'Pratinjau Format Email'
      );
    });
  });
}

function openEmailModal(recipient, subject, html, modeText) {
  const modal = document.getElementById('email-modal');
  const metaEl = document.getElementById('modal-email-meta');
  const subjEl = document.getElementById('modal-email-subject');
  const iframeEl = document.getElementById('modal-email-iframe');

  if (metaEl) metaEl.textContent = `Kepada: ${recipient} · ${modeText}`;
  if (subjEl) subjEl.textContent = subject;
  if (iframeEl) iframeEl.srcdoc = html;
  if (modal) modal.classList.remove('is-hidden');
}

function closeEmailModal() {
  document.getElementById('email-modal')?.classList.add('is-hidden');
}

// ============================================================================
// INITIALIZATION
// ============================================================================
function init() {
  renderCatalog();

  document.querySelectorAll('[data-nav]').forEach((el) => {
    el.addEventListener('click', () => {
      const target = el.getAttribute('data-nav');
      if (target) setView(target);
    });
  });

  document.getElementById('btn-hero-cta')?.addEventListener('click', () => {
    setView('portal');
    setTimeout(() => {
      document.getElementById('katalog-formulir')?.scrollIntoView({ behavior: 'smooth' });
    }, 40);
  });

  document.getElementById('btn-scroll-catalog')?.addEventListener('click', () => {
    document.getElementById('katalog-formulir')?.scrollIntoView({ behavior: 'smooth' });
  });

  document.getElementById('btn-back-portal')?.addEventListener('click', () => {
    setView('portal');
  });

  document.getElementById('catalog-search')?.addEventListener('input', (e) => {
    state.searchQuery = e.target.value || '';
    renderCatalog();
  });

  const heroImg = document.getElementById('hero-img');
  heroImg?.addEventListener('error', () => {
    heroImg.classList.add('is-hidden');
    document.getElementById('hero-img-fallback')?.classList.remove('is-hidden');
  });

  document.querySelectorAll('#script-tabs .seg-btn').forEach((btn) => {
    btn.addEventListener('click', () => {
      state.activeScriptTab = btn.getAttribute('data-script') || 'webapp-dopost';
      renderAppsScriptBlock();
    });
  });

  document.getElementById('btn-copy-script')?.addEventListener('click', async () => {
    const code =
      state.activeScriptTab === 'form-trigger'
        ? APPS_SCRIPT_FORM_TRIGGER
        : APPS_SCRIPT_WEBAPP_DOPOST;
    try {
      await navigator.clipboard.writeText(code);
      showToast('Kode Google Apps Script berhasil disalin.');
    } catch {
      showToast('Gagal menyalin otomatis. Silakan salin manual dari kotak kode.');
    }
  });

  document.getElementById('btn-download-script')?.addEventListener('click', () => {
    const code =
      state.activeScriptTab === 'form-trigger'
        ? APPS_SCRIPT_FORM_TRIGGER
        : APPS_SCRIPT_WEBAPP_DOPOST;
    const blob = new Blob([code], { type: 'text/plain;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'Code.gs';
    a.click();
    URL.revokeObjectURL(url);
    showToast('File Code.gs berhasil diunduh.');
  });

  document.getElementById('btn-save-script-url')?.addEventListener('click', () => {
    const input = document.getElementById('input-apps-script-url');
    const url = input ? input.value.trim() : '';
    state.config.appsScriptWebAppUrl = url;
    state.dropdownData.loaded = false;
    saveConfig();
    renderIntegrationsView();
    renderCatalog();
    showToast(url ? 'URL Apps Script disimpan. Dropdown siap memuat data Spreadsheet.' : 'URL Apps Script dikosongkan.');
  });

  document.getElementById('btn-save-gform-urls')?.addEventListener('click', () => {
    document.querySelectorAll('[data-gform-id]').forEach((input) => {
      const id = input.getAttribute('data-gform-id');
      if (id) state.config.googleFormUrls[id] = input.value.trim();
    });
    saveConfig();
    renderIntegrationsView();
    renderCatalog();
    showToast('Semua tautan Google Form berhasil disimpan.');
  });

  document.getElementById('btn-copy-config')?.addEventListener('click', async () => {
    const snippet = `const DEFAULT_CONFIG = ${JSON.stringify(state.config, null, 2)};`;
    try {
      await navigator.clipboard.writeText(snippet);
      showToast('Konfigurasi JS berhasil disalin.');
    } catch {
      showToast('Gagal menyalin konfigurasi.');
    }
  });

  document.getElementById('btn-clear-history')?.addEventListener('click', () => {
    state.history = [];
    saveHistory();
    renderHistoryView();
    showToast('Riwayat lokal telah dibersihkan.');
  });

  document.getElementById('btn-close-modal')?.addEventListener('click', closeEmailModal);
  document.getElementById('btn-close-modal-footer')?.addEventListener('click', closeEmailModal);
  document.getElementById('email-modal')?.addEventListener('click', (e) => {
    if (e.target.id === 'email-modal') closeEmailModal();
  });
}

init();
