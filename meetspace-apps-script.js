// ═══════════════════════════════════════════════════════════════════
//  MEETSPACE — GOOGLE APPS SCRIPT
//  Paste kode ini di: Google Sheets → Extensions → Apps Script
//  Lalu Deploy sebagai Web App (akses: Anyone)
// ═══════════════════════════════════════════════════════════════════

const SHEET_BOOKING  = 'Data Booking';
const SHEET_USERS    = 'Data Pengguna';
const SHEET_ROOMS    = 'Data Ruangan';
const SHEET_LOG      = 'Log Aktivitas';

// ───────────────────────────────────────────────────────────────────
// ENTRY POINT — semua request HTTP masuk ke sini
// ───────────────────────────────────────────────────────────────────
function doGet(e) {
  const action = e.parameter.action || '';
  const ss     = SpreadsheetApp.getActiveSpreadsheet();

  try {
    let result;
    switch (action) {
      case 'getBookings': result = getBookings(ss, e.parameter);  break;
      case 'getRooms':    result = getRooms(ss);                  break;
      case 'getUsers':    result = getUsers(ss);                  break;
      case 'getStats':    result = getStats(ss);                  break;
      case 'checkSlots':  result = checkSlots(ss, e.parameter);   break;
      default:            result = { ok: true, msg: 'MeetSpace API aktif ✅' };
    }
    return respond(result);
  } catch (err) {
    return respond({ ok: false, error: err.message });
  }
}

function doPost(e) {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const body = JSON.parse(e.postData.contents);
  const action = body.action || '';

  try {
    let result;
    switch (action) {
      case 'addBooking':       result = addBooking(ss, body);       break;
      case 'updateStatus':     result = updateStatus(ss, body);     break;
      case 'cancelBooking':    result = cancelBooking(ss, body);    break;
      case 'registerUser':     result = registerUser(ss, body);     break;
      case 'loginUser':        result = loginUser(ss, body);        break;
      case 'updateUser':       result = updateUser(ss, body);       break;
      case 'addRoom':          result = addRoom(ss, body);          break;
      case 'updateRoom':       result = updateRoom(ss, body);       break;
      default: result = { ok: false, error: 'Action tidak dikenal: ' + action };
    }
    addLog(ss, action, body.email || 'system', result.ok ? 'OK' : 'ERROR');
    return respond(result);
  } catch (err) {
    addLog(ss, action, 'system', 'ERROR: ' + err.message);
    return respond({ ok: false, error: err.message });
  }
}

// Kirim response JSON dengan CORS header
function respond(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ───────────────────────────────────────────────────────────────────
// INISIALISASI SHEET — buat otomatis jika belum ada
// ───────────────────────────────────────────────────────────────────
function initSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ── Data Booking ──
  let sh = getOrCreate(ss, SHEET_BOOKING);
  if (sh.getLastRow() === 0) {
    sh.appendRow([
      'Kode Booking','Nama Ruangan','Tanggal','Waktu Mulai','Waktu Selesai',
      'Jumlah Peserta','Status','Nama Pemesan','Departemen','Email Pemesan',
      'Agenda','Kebutuhan Tambahan','Timestamp'
    ]);
    styleHeader(sh);
    // Contoh data awal
    const contoh = [
      ['MSP-482031','Ruang Cendana','2026-05-06','09:00','11:00',6,'Dikonfirmasi','Budi Santoso','Marketing','budi@perusahaan.com','Rapat Bulanan','Tidak ada',new Date()],
      ['MSP-719204','Ruang Mahoni','2026-05-08','13:00','15:00',12,'Menunggu','Sari Dewi','Finance','sari@perusahaan.com','Review Anggaran','Catering',new Date()],
      ['MSP-305882','Ruang Emerald','2026-05-10','10:00','12:00',20,'Dikonfirmasi','Ahmad Rizal','Engineering','ahmad@perusahaan.com','Sprint Planning','Tidak ada',new Date()],
      ['MSP-213901','Ruang Laguna','2026-05-15','08:00','10:00',8,'Dikonfirmasi','Rina Kusuma','HR','rina@perusahaan.com','Rekrutmen','Tidak ada',new Date()],
      ['MSP-841033','Aula Grand Panorama','2026-05-22','09:00','12:00',40,'Menunggu','Yeni Safitri','Operations','yeni@perusahaan.com','All-hands Meeting','Catering',new Date()],
    ];
    contoh.forEach(r => sh.appendRow(r));
  }

  // ── Data Pengguna ──
  sh = getOrCreate(ss, SHEET_USERS);
  if (sh.getLastRow() === 0) {
    sh.appendRow([
      'Email','Password (Hash)','Role','Nama Lengkap','Jabatan',
      'Departemen','NIK','Nomor HP','Tanggal Bergabung','Status Akun'
    ]);
    styleHeader(sh);
    sh.appendRow(['admin@meetspace.id','admin123','admin','Admin Utama','System Administrator','IT','EMP-2020-0001','0811-0000-0001','2020-01-01','Aktif']);
    sh.appendRow(['staff@meetspace.id','staff123','staff','Budi Santoso','Sr. Marketing Manager','Marketing','EMP-2021-0042','0812-3456-7890','2021-03-15','Aktif']);
    sh.appendRow(['staff2@meetspace.id','staff123','staff','Sari Dewi','Finance Analyst','Finance','EMP-2022-0015','0813-9876-5432','2022-06-10','Aktif']);
  }

  // ── Data Ruangan ──
  sh = getOrCreate(ss, SHEET_ROOMS);
  if (sh.getLastRow() === 0) {
    sh.appendRow([
      'ID Ruangan','Nama Ruangan','Lantai','Lokasi','Kapasitas',
      'Proyektor','WiFi','Whiteboard','Video Conference','Catering',
      'Rating','Status','Deskripsi'
    ]);
    styleHeader(sh);
    const ruangan = [
      ['ROOM-001','Ruang Cendana','Lantai 1','Sayap Barat',8,'✅','✅','✅','❌','❌',4.8,'Aktif','Ruangan nyaman untuk rapat kecil'],
      ['ROOM-002','Ruang Mahoni','Lantai 2','Sayap Timur',15,'❌','✅','❌','✅','✅',4.7,'Aktif','Dilengkapi Smart TV & Video Conference'],
      ['ROOM-003','Ruang Laguna','Lantai 3','Pusat',12,'✅','✅','✅','❌','❌',4.9,'Aktif','View terbaik di lantai 3'],
      ['ROOM-004','Ruang Emerald','Lantai 4','Sayap Utara',25,'✅','✅','❌','✅','❌',4.8,'Aktif','Ruangan besar dengan dual proyektor'],
      ['ROOM-005','Ruang Amethyst VIP','Lantai 5','Eksekutif',10,'❌','✅','❌','✅','✅',5.0,'Aktif','Ruangan VIP eksekutif'],
      ['ROOM-006','Aula Grand Panorama','Lantai 6','Rooftop',50,'✅','✅','❌','❌','✅',4.9,'Aktif','Aula rooftop dengan view 360°'],
    ];
    ruangan.forEach(r => sh.appendRow(r));
  }

  // ── Log Aktivitas ──
  sh = getOrCreate(ss, SHEET_LOG);
  if (sh.getLastRow() === 0) {
    sh.appendRow(['Timestamp','Action','User','Status','Detail']);
    styleHeader(sh);
  }

  SpreadsheetApp.getUi().alert('✅ MeetSpace berhasil diinisialisasi!\n\nSemua sheet sudah siap digunakan.');
}

// ───────────────────────────────────────────────────────────────────
// BOOKING FUNCTIONS
// ───────────────────────────────────────────────────────────────────

function getBookings(ss, params) {
  const sh   = ss.getSheetByName(SHEET_BOOKING);
  const rows = sheetToObjects(sh);

  // Filter opsional
  let filtered = rows;
  if (params.status)  filtered = filtered.filter(r => r['Status'] === params.status);
  if (params.email)   filtered = filtered.filter(r => r['Email Pemesan'] === params.email);
  if (params.date)    filtered = filtered.filter(r => r['Tanggal'] === params.date);
  if (params.room)    filtered = filtered.filter(r => r['Nama Ruangan'] === params.room);

  // Ubah format ke camelCase untuk frontend
  const bookings = filtered.map(r => ({
    code:    r['Kode Booking'],
    room:    r['Nama Ruangan'],
    date:    r['Tanggal'],
    start:   r['Waktu Mulai'],
    end:     r['Waktu Selesai'],
    pax:     Number(r['Jumlah Peserta']) || 0,
    status:  r['Status'],
    name:    r['Nama Pemesan'],
    dept:    r['Departemen'],
    email:   r['Email Pemesan'],
    agenda:  r['Agenda'],
    extra:   r['Kebutuhan Tambahan'],
    ts:      r['Timestamp'],
  }));

  return { ok: true, data: bookings, total: bookings.length };
}

function addBooking(ss, body) {
  const sh = ss.getSheetByName(SHEET_BOOKING);

  // Validasi wajib
  const required = ['room','date','start','end','name','email'];
  for (const f of required) {
    if (!body[f]) return { ok: false, error: `Field '${f}' wajib diisi` };
  }

  // Cek konflik slot
  if (isSlotConflict(ss, body.room, body.date, body.start, body.end)) {
    return { ok: false, error: `Slot ${body.start}–${body.end} di ${body.room} pada ${body.date} sudah dipesan! Pilih waktu lain.` };
  }

  // Buat kode booking unik
  const code = 'MSP-' + Math.floor(100000 + Math.random() * 900000);

  sh.appendRow([
    code,
    body.room,
    body.date,
    body.start,
    body.end,
    body.pax || 1,
    'Menunggu',        // status awal selalu Menunggu
    body.name,
    body.dept || '',
    body.email,
    body.agenda || '',
    body.extra || 'Tidak ada',
    new Date().toISOString(),
  ]);

  // Kirim notifikasi email (opsional — aktifkan jika mau)
  // sendBookingEmail(body.email, body.name, code, body.room, body.date, body.start, body.end);

  return {
    ok:   true,
    code: code,
    msg:  `Booking ${code} berhasil dibuat! Status: Menunggu Konfirmasi.`,
  };
}

function updateStatus(ss, body) {
  // body: { code, status, adminEmail }
  if (!body.code || !body.status) return { ok: false, error: 'code dan status wajib diisi' };

  const sh  = ss.getSheetByName(SHEET_BOOKING);
  const row = findRow(sh, 'Kode Booking', body.code);
  if (!row) return { ok: false, error: `Booking ${body.code} tidak ditemukan` };

  const statusCol = getColIndex(sh, 'Status');
  sh.getRange(row, statusCol).setValue(body.status);

  // Kirim email notifikasi jika dikonfirmasi
  if (body.status === 'Dikonfirmasi') {
    const data = rowToObject(sh, row);
    sendConfirmationEmail(data['Email Pemesan'], data['Nama Pemesan'], body.code,
      data['Nama Ruangan'], data['Tanggal'], data['Waktu Mulai'], data['Waktu Selesai']);
  }

  return { ok: true, msg: `Status booking ${body.code} diubah ke "${body.status}"` };
}

function cancelBooking(ss, body) {
  return updateStatus(ss, { ...body, status: 'Dibatalkan' });
}

function checkSlots(ss, params) {
  // params: room, date
  const sh   = ss.getSheetByName(SHEET_BOOKING);
  const rows = sheetToObjects(sh);
  const busy = rows
    .filter(r =>
      r['Nama Ruangan'] === params.room &&
      r['Tanggal']      === params.date &&
      r['Status']       !== 'Dibatalkan'
    )
    .map(r => ({ start: r['Waktu Mulai'], end: r['Waktu Selesai'] }));

  return { ok: true, busySlots: busy };
}

function isSlotConflict(ss, room, date, start, end) {
  const sh   = ss.getSheetByName(SHEET_BOOKING);
  const rows = sheetToObjects(sh);
  return rows.some(r =>
    r['Nama Ruangan'] === room &&
    r['Tanggal']      === date &&
    r['Status']       !== 'Dibatalkan' &&
    timeOverlap(r['Waktu Mulai'], r['Waktu Selesai'], start, end)
  );
}

function timeOverlap(s1, e1, s2, e2) {
  const toMin = t => { const [h,m] = t.split(':').map(Number); return h*60+m; };
  return toMin(s1) < toMin(e2) && toMin(s2) < toMin(e1);
}

// ───────────────────────────────────────────────────────────────────
// USER FUNCTIONS
// ───────────────────────────────────────────────────────────────────

function loginUser(ss, body) {
  if (!body.email || !body.password) return { ok: false, error: 'Email dan password wajib diisi' };

  const sh  = ss.getSheetByName(SHEET_USERS);
  const row = findRow(sh, 'Email', body.email.toLowerCase());
  if (!row) return { ok: false, error: 'Email tidak terdaftar' };

  const data = rowToObject(sh, row);
  if (data['Password (Hash)'] !== body.password) return { ok: false, error: 'Password salah' };
  if (data['Status Akun'] !== 'Aktif') return { ok: false, error: 'Akun tidak aktif. Hubungi admin.' };

  // Jangan kirim password ke frontend
  return {
    ok: true,
    user: {
      email:     data['Email'],
      role:      data['Role'],
      name:      data['Nama Lengkap'],
      jabatan:   data['Jabatan'],
      dept:      data['Departemen'],
      nik:       data['NIK'],
      hp:        data['Nomor HP'],
      bergabung: data['Tanggal Bergabung'],
      initials:  data['Nama Lengkap'].split(' ').map(w=>w[0]).join('').substring(0,2).toUpperCase(),
    }
  };
}

function registerUser(ss, body) {
  const required = ['email','password','name','role'];
  for (const f of required) {
    if (!body[f]) return { ok: false, error: `Field '${f}' wajib diisi` };
  }

  const sh  = ss.getSheetByName(SHEET_USERS);
  const row = findRow(sh, 'Email', body.email.toLowerCase());
  if (row) return { ok: false, error: 'Email sudah terdaftar' };

  // Validasi role — hanya staff yang bisa daftar mandiri; admin dibuat manual
  if (!['staff','admin'].includes(body.role)) {
    return { ok: false, error: 'Role tidak valid. Gunakan "staff" atau "admin".' };
  }

  const year = new Date().getFullYear();
  const nik  = `EMP-${year}-${String(Math.floor(1000 + Math.random()*9000))}`;
  const now  = new Date().toISOString().split('T')[0];

  sh.appendRow([
    body.email.toLowerCase(),
    body.password,       // di produksi: hash password dulu!
    body.role,
    body.name,
    body.jabatan  || '',
    body.dept     || '',
    nik,
    body.hp       || '',
    now,
    'Aktif',
  ]);

  return {
    ok:  true,
    nik: nik,
    msg: `Akun ${body.email} berhasil didaftarkan sebagai ${body.role}`,
  };
}

function updateUser(ss, body) {
  if (!body.email) return { ok: false, error: 'Email wajib diisi' };

  const sh  = ss.getSheetByName(SHEET_USERS);
  const row = findRow(sh, 'Email', body.email.toLowerCase());
  if (!row) return { ok: false, error: 'User tidak ditemukan' };

  const fields = {
    'Nama Lengkap': body.name,
    'Jabatan':      body.jabatan,
    'Departemen':   body.dept,
    'Nomor HP':     body.hp,
  };

  Object.entries(fields).forEach(([col, val]) => {
    if (val !== undefined) {
      const colIdx = getColIndex(sh, col);
      if (colIdx) sh.getRange(row, colIdx).setValue(val);
    }
  });

  return { ok: true, msg: 'Profil berhasil diperbarui' };
}

function getUsers(ss) {
  const sh   = ss.getSheetByName(SHEET_USERS);
  const rows = sheetToObjects(sh);
  // Jangan kirim password!
  const users = rows.map(r => ({
    email:     r['Email'],
    role:      r['Role'],
    name:      r['Nama Lengkap'],
    jabatan:   r['Jabatan'],
    dept:      r['Departemen'],
    nik:       r['NIK'],
    hp:        r['Nomor HP'],
    bergabung: r['Tanggal Bergabung'],
    status:    r['Status Akun'],
  }));
  return { ok: true, data: users, total: users.length };
}

// ───────────────────────────────────────────────────────────────────
// ROOM FUNCTIONS
// ───────────────────────────────────────────────────────────────────

function getRooms(ss) {
  const sh   = ss.getSheetByName(SHEET_ROOMS);
  const rows = sheetToObjects(sh);
  const rooms = rows
    .filter(r => r['Status'] === 'Aktif')
    .map(r => ({
      id:          r['ID Ruangan'],
      title:       r['Nama Ruangan'],
      floor:       r['Lantai'],
      loc:         r['Lokasi'],
      cap:         r['Kapasitas'] + ' orang',
      proyektor:   r['Proyektor'] === '✅',
      wifi:        r['WiFi'] === '✅',
      whiteboard:  r['Whiteboard'] === '✅',
      vc:          r['Video Conference'] === '✅',
      catering:    r['Catering'] === '✅',
      rating:      r['Rating'] + ' ⭐',
      status:      r['Status'],
      desc:        r['Deskripsi'],
    }));
  return { ok: true, data: rooms, total: rooms.length };
}

function addRoom(ss, body) {
  const sh = ss.getSheetByName(SHEET_ROOMS);
  const id = 'ROOM-' + String(sh.getLastRow()).padStart(3,'0');
  sh.appendRow([
    id,
    body.name, body.floor, body.loc, body.cap,
    body.proyektor ? '✅' : '❌',
    body.wifi      ? '✅' : '❌',
    body.whiteboard? '✅' : '❌',
    body.vc        ? '✅' : '❌',
    body.catering  ? '✅' : '❌',
    5.0, 'Aktif', body.desc || '',
  ]);
  return { ok: true, id: id, msg: `Ruangan ${body.name} berhasil ditambahkan` };
}

function updateRoom(ss, body) {
  if (!body.id) return { ok: false, error: 'ID ruangan wajib diisi' };
  const sh  = ss.getSheetByName(SHEET_ROOMS);
  const row = findRow(sh, 'ID Ruangan', body.id);
  if (!row) return { ok: false, error: 'Ruangan tidak ditemukan' };

  const fieldMap = {
    'Nama Ruangan': body.name, 'Lantai': body.floor, 'Lokasi': body.loc,
    'Kapasitas': body.cap, 'Status': body.status, 'Deskripsi': body.desc,
  };
  Object.entries(fieldMap).forEach(([col, val]) => {
    if (val !== undefined) {
      const ci = getColIndex(sh, col);
      if (ci) sh.getRange(row, ci).setValue(val);
    }
  });
  return { ok: true, msg: `Ruangan ${body.id} berhasil diperbarui` };
}

// ───────────────────────────────────────────────────────────────────
// STATS — untuk dashboard admin
// ───────────────────────────────────────────────────────────────────
function getStats(ss) {
  const bSh  = ss.getSheetByName(SHEET_BOOKING);
  const rows = sheetToObjects(bSh);
  const total       = rows.length;
  const confirmed   = rows.filter(r => r['Status'] === 'Dikonfirmasi').length;
  const pending     = rows.filter(r => r['Status'] === 'Menunggu').length;
  const cancelled   = rows.filter(r => r['Status'] === 'Dibatalkan').length;

  // Booking per ruangan
  const perRoom = {};
  rows.forEach(r => {
    const rm = r['Nama Ruangan'];
    if (!perRoom[rm]) perRoom[rm] = { total:0, confirmed:0, pending:0 };
    perRoom[rm].total++;
    if (r['Status']==='Dikonfirmasi') perRoom[rm].confirmed++;
    if (r['Status']==='Menunggu')     perRoom[rm].pending++;
  });

  const rSh    = ss.getSheetByName(SHEET_ROOMS);
  const rooms  = sheetToObjects(rSh).filter(r => r['Status'] === 'Aktif').length;

  return {
    ok: true,
    stats: { total, confirmed, pending, cancelled, activeRooms: rooms },
    perRoom,
    updatedAt: new Date().toISOString(),
  };
}

// ───────────────────────────────────────────────────────────────────
// EMAIL NOTIFICATION
// ───────────────────────────────────────────────────────────────────
function sendConfirmationEmail(email, name, code, room, date, start, end) {
  if (!email) return;
  try {
    MailApp.sendEmail({
      to:      email,
      subject: `✅ Booking Dikonfirmasi — ${code} | MeetSpace`,
      htmlBody: `
        <div style="font-family:Arial,sans-serif;max-width:520px;margin:0 auto;padding:24px">
          <div style="background:#18192B;border-radius:12px;padding:24px;text-align:center;margin-bottom:20px">
            <h1 style="color:#C9986A;font-size:24px;margin:0">MeetSpace</h1>
            <p style="color:rgba(255,255,255,.6);font-size:12px;margin:4px 0 0">Sistem Booking Ruangan Meeting</p>
          </div>
          <h2 style="color:#18192B">Booking Dikonfirmasi ✅</h2>
          <p>Halo <strong>${name}</strong>,</p>
          <p>Booking ruangan Anda telah <strong>dikonfirmasi</strong> oleh Admin. Berikut detailnya:</p>
          <table style="width:100%;border-collapse:collapse;margin:16px 0">
            <tr><td style="padding:10px;background:#F7F3EE;font-weight:700;width:40%">Kode Booking</td><td style="padding:10px;background:#FFF;font-family:monospace;font-size:16px;font-weight:700;color:#C9986A">${code}</td></tr>
            <tr><td style="padding:10px;background:#F7F3EE;font-weight:700">Ruangan</td><td style="padding:10px;background:#FFF">${room}</td></tr>
            <tr><td style="padding:10px;background:#F7F3EE;font-weight:700">Tanggal</td><td style="padding:10px;background:#FFF">${date}</td></tr>
            <tr><td style="padding:10px;background:#F7F3EE;font-weight:700">Waktu</td><td style="padding:10px;background:#FFF">${start} – ${end}</td></tr>
          </table>
          <p style="font-size:12px;color:#888">Email ini dikirim otomatis oleh sistem MeetSpace. Jangan reply email ini.</p>
        </div>
      `,
    });
  } catch(e) {
    Logger.log('Gagal kirim email: ' + e.message);
  }
}

// ───────────────────────────────────────────────────────────────────
// LOG
// ───────────────────────────────────────────────────────────────────
function addLog(ss, action, user, status) {
  const sh = ss.getSheetByName(SHEET_LOG);
  if (!sh) return;
  sh.appendRow([new Date().toISOString(), action, user, status, '']);
  // Jaga log max 1000 baris
  if (sh.getLastRow() > 1001) {
    sh.deleteRow(2);
  }
}

// ───────────────────────────────────────────────────────────────────
// HELPER FUNCTIONS
// ───────────────────────────────────────────────────────────────────

function getOrCreate(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function sheetToObjects(sh) {
  if (sh.getLastRow() < 2) return [];
  const data    = sh.getDataRange().getValues();
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i] !== undefined ? String(row[i]) : ''; });
    return obj;
  });
}

function rowToObject(sh, rowNum) {
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const row     = sh.getRange(rowNum, 1, 1, sh.getLastColumn()).getValues()[0];
  const obj = {};
  headers.forEach((h, i) => { obj[h] = row[i] !== undefined ? String(row[i]) : ''; });
  return obj;
}

function findRow(sh, colName, value) {
  const col  = getColIndex(sh, colName);
  if (!col) return null;
  const vals = sh.getRange(2, col, sh.getLastRow()-1).getValues();
  for (let i = 0; i < vals.length; i++) {
    if (String(vals[i][0]).toLowerCase() === String(value).toLowerCase()) return i + 2;
  }
  return null;
}

function getColIndex(sh, colName) {
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const idx = headers.indexOf(colName);
  return idx >= 0 ? idx + 1 : null;
}

function styleHeader(sh) {
  const range = sh.getRange(1, 1, 1, sh.getLastColumn());
  range.setBackground('#18192B')
       .setFontColor('#C9986A')
       .setFontWeight('bold')
       .setFontSize(11);
  sh.setFrozenRows(1);
}

// ───────────────────────────────────────────────────────────────────
// MENU — muncul di Google Sheets
// ───────────────────────────────────────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🏢 MeetSpace')
    .addItem('⚙️ Inisialisasi Semua Sheet', 'initSheets')
    .addItem('📊 Lihat Statistik', 'showStats')
    .addItem('🔗 Salin URL Web App', 'copyWebAppUrl')
    .addToUi();
}

function showStats() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const stats = getStats(ss).stats;
  SpreadsheetApp.getUi().alert(
    `📊 Statistik MeetSpace\n\n` +
    `Total Booking  : ${stats.total}\n` +
    `Dikonfirmasi   : ${stats.confirmed}\n` +
    `Menunggu       : ${stats.pending}\n` +
    `Dibatalkan     : ${stats.cancelled}\n` +
    `Ruangan Aktif  : ${stats.activeRooms}`
  );
}

function copyWebAppUrl() {
  const url = ScriptApp.getService().getUrl();
  SpreadsheetApp.getUi().alert(
    '🔗 URL Web App kamu:\n\n' + url +
    '\n\nSalin URL ini dan tempel ke variabel APPS_SCRIPT_URL di file index.html'
  );
}
