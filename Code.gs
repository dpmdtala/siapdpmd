/**
 * SIAP-DPMD - Google Apps Script backend
 * Database: Google Sheets; optional image storage: Google Drive.
 * Database remains the existing Google Sheet configured in Script Properties.
 * Do not run setupDatabase() during routine deployment or testing.
 */
const APP = {
  TZ: 'Asia/Makassar',
  SHEETS: {
    Users: ['id','nama','nik','hp','email','passwordHash','bidang','jabatan','role','status','ttl','tmt','avatar','rekening','kawin','alamat','jenisKelamin','pendidikanTerakhir','createdAt','updatedAt'],
    Attendance: ['id','userId','tanggal','jam','status','sesi','catatan','latitude','longitude','distanceMeters','lateBySeconds','lateByText','selfieFileId','signatureFileId','createdAt'],
    Reports: ['id','userId','nama','bidang','tanggal','jamMulai','jamSelesai','kegiatan','output','fotoFileId','status','verifiedBy','verifiedAt','createdAt'],
    Archives: ['id','userId','jenis','nama','driveFileId','tanggal','status','createdAt'],
    AuditLog: ['id','actorId','action','entity','entityId','details','createdAt']
  },
  SESSION_TTL_SECONDS: 21600,
  MAX_UPLOAD_BYTES: 5 * 1024 * 1024,
  MAX_ARCHIVE_BYTES: 1 * 1024 * 1024
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('SIAP-DPMD | DPMD Kabupaten Tanah Laut')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Jalankan satu kali dari editor Apps Script untuk membuat spreadsheet dan sheet. */
function setupDatabase() {
  const props = PropertiesService.getScriptProperties();
  let spreadsheetId = props.getProperty('SIAP_DPMD_SPREADSHEET_ID');
  let ss;
  if (spreadsheetId) {
    try { ss = SpreadsheetApp.openById(spreadsheetId); } catch (e) { ss = null; }
  }
  if (!ss) {
    throw new Error('Database SIAP-DPMD tidak ditemukan atau SIAP_DPMD_SPREADSHEET_ID belum benar. Tidak ada database baru yang dibuat. Periksa Script Properties dan ID spreadsheet yang sudah digunakan.');
  }

  Object.keys(APP.SHEETS).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    const headers = APP.SHEETS[name];
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers]);
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#dbeafe');
      sh.autoResizeColumns(1, headers.length);
    } else {
      const existing = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
      const missing = headers.filter(h => existing.indexOf(h) < 0);
      if (missing.length) {
        const startCol = sh.getLastColumn() + 1;
        sh.getRange(1, startCol, 1, missing.length).setValues([missing]);
        sh.getRange(1, startCol, 1, missing.length).setFontWeight('bold').setBackground('#dbeafe');
        sh.autoResizeColumns(startCol, missing.length);
      }
    }
  });

  // Remove unused default Sheet1 only if empty.
  const defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet && defaultSheet.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(defaultSheet);

  // No fixed/default credentials. If Users is empty, require explicit one-time
  // credentials in Script Properties before creating bootstrap accounts.
  const users = readRows_('Users');
  if (users.length === 0) {
    const adminPassword = props.getProperty('SIAP_DPMD_BOOTSTRAP_ADMIN_PASSWORD') || '';
    const verifierPassword = props.getProperty('SIAP_DPMD_BOOTSTRAP_VERIFIKATOR_PASSWORD') || '';
    if (adminPassword.length < 14 || verifierPassword.length < 14) {
      throw new Error('Sheet Users masih kosong. Untuk bootstrap, atur SIAP_DPMD_BOOTSTRAP_ADMIN_PASSWORD dan SIAP_DPMD_BOOTSTRAP_VERIFIKATOR_PASSWORD di Script Properties (masing-masing minimal 14 karakter), lalu jalankan setupDatabase secara sadar. Tidak ada akun bawaan yang dibuat.');
    }
    const now = new Date().toISOString();
    const initial = [
      makeUser_({id: 'admin-' + Utilities.getUuid(), nama: 'Administrator SIAP-DPMD', nik: 'ADMIN-DPMD', hp: '', email: 'admin@dpmd.tanahlautkab.go.id', password: adminPassword, bidang: 'Sekretariat', jabatan: 'Super Admin System', role: 'SuperAdmin', status: 'Aktif'}, now),
      makeUser_({id: 'verif-' + Utilities.getUuid(), nama: 'Verifikator Sekretariat', nik: 'VERIF-SEKRETARIAT', hp: '', email: 'verifikator@dpmd.tanahlautkab.go.id', password: verifierPassword, bidang: 'Sekretariat', jabatan: 'Verifikator', role: 'Verifikator', status: 'Aktif'}, now)
    ];
    appendObjects_('Users', initial);
    props.deleteProperty('SIAP_DPMD_BOOTSTRAP_ADMIN_PASSWORD');
    props.deleteProperty('SIAP_DPMD_BOOTSTRAP_VERIFIKATOR_PASSWORD');
  }
  return { ok: true, spreadsheetId: ss.getId(), spreadsheetUrl: ss.getUrl(), sheets: Object.keys(APP.SHEETS), message: 'Pemeriksaan struktur selesai. Tidak ada kata sandi bawaan di kode.' };
}

/**
 * Pemulihan Super Admin hanya menggunakan kredensial sekali pakai dari Script Properties.
 * Jangan jalankan kecuali sedang melakukan pemulihan yang disengaja.
 * Properti yang diperlukan: SIAP_DPMD_RECOVERY_SUPERADMIN_PASSWORD (min. 14 karakter).
 * Properti kata sandi dihapus setelah berhasil; tidak ada kata sandi tetap di source code.
 */
function createOrRepairSuperAdminAccount() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) throw new Error('Proses lain sedang berjalan. Coba lagi nanti.');
  try {
    const props = PropertiesService.getScriptProperties();
    const password = props.getProperty('SIAP_DPMD_RECOVERY_SUPERADMIN_PASSWORD') || '';
    if (password.length < 14) {
      throw new Error('Pemulihan dibatalkan. Atur SIAP_DPMD_RECOVERY_SUPERADMIN_PASSWORD di Script Properties dengan kata sandi acak minimal 14 karakter. Fungsi ini tidak menetapkan kata sandi bawaan.');
    }
    const username = String(props.getProperty('SIAP_DPMD_RECOVERY_SUPERADMIN_USERNAME') || 'superadmin').trim().toLowerCase();
    const email = String(props.getProperty('SIAP_DPMD_RECOVERY_SUPERADMIN_EMAIL') || 'superadmin@dpmd.tanahlautkab.go.id').trim().toLowerCase();
    if (!/^[a-z0-9._-]{3,64}$/.test(username)) throw new Error('Username pemulihan tidak valid.');
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('Email pemulihan tidak valid.');

    // Read only the already-configured database. Never call setupDatabase() here.
    const users = readRows_('Users');
    const matches = users.filter(u => String(u.nik || '').toLowerCase() === username || String(u.email || '').toLowerCase() === email);
    if (matches.some(u => u.role !== 'SuperAdmin')) {
      throw new Error('Username/email pemulihan dipakai akun non-SuperAdmin. Tidak ada perubahan dilakukan; periksa sheet Users.');
    }
    if (matches.length > 1) throw new Error('Ditemukan lebih dari satu akun yang cocok. Tidak ada perubahan dilakukan; periksa sheet Users.');

    const now = new Date().toISOString();
    if (matches.length === 1) {
      const target = matches[0];
      updateById_('Users', target.id, {
        nama: 'Super Admin SIAP-DPMD', nik: username, email,
        passwordHash: hashPassword_(password), role: 'SuperAdmin', status: 'Aktif',
        jabatan: 'Super Admin System', bidang: target.bidang || 'Sekretariat', updatedAt: now
      });
      audit_(target.id, 'ADMIN_ACCOUNT_RECOVERED', 'Users', target.id, 'Akun Super Admin dipulihkan menggunakan kredensial sekali pakai');
    } else {
      const row = makeUser_({
        id:'admin-' + Utilities.getUuid(), nama:'Super Admin SIAP-DPMD', nik:username, hp:'', email,
        password, bidang:'Sekretariat', jabatan:'Super Admin System', role:'SuperAdmin', status:'Aktif'
      }, now);
      appendObjects_('Users', [row]);
      audit_(row.id, 'ADMIN_ACCOUNT_RECOVERED', 'Users', row.id, 'Akun Super Admin dibuat menggunakan kredensial sekali pakai');
    }
    props.deleteProperty('SIAP_DPMD_RECOVERY_SUPERADMIN_PASSWORD');
    props.deleteProperty('SIAP_DPMD_RECOVERY_SUPERADMIN_USERNAME');
    props.deleteProperty('SIAP_DPMD_RECOVERY_SUPERADMIN_EMAIL');
    return {ok:true, created:matches.length === 0, username, email, message:'Pemulihan akun berhasil. Properti kredensial sekali pakai telah dihapus. Segera masuk dan pastikan kata sandi tersimpan aman.'};
  } finally {
    lock.releaseLock();
  }
}

function api(action, payload, token) {
  payload = payload || {};
  if (action === 'healthCheck') return {ok:true, app:'SIAP-DPMD', time:new Date().toISOString()};
  if (action === 'login') return login_(payload);
  if (action === 'register') return register_(payload);
  if (action === 'logout') { if (token) CacheService.getScriptCache().remove('sess:' + token); return {ok:true}; }

  const user = requireSession_(token);
  switch (action) {
    case 'getAppData': return getAppData_(user, payload);
    case 'getDataPage': return getDataPage_(user, payload);
    case 'reviewRegistration': requireRole_(user, ['SuperAdmin']); return reviewRegistration_(user, payload);
    case 'getReportPhoto': return getReportPhoto_(user, payload.reportId);
    case 'updateProfile': return updateProfile_(user, payload);
    case 'changePassword': return changePassword_(user, payload);
    case 'adminCreateUser': requireRole_(user, ['SuperAdmin']); return adminCreateUser_(payload, user);
    case 'deleteUser': requireRole_(user, ['SuperAdmin']); return deleteUser_(payload.userId, user);
    case 'deleteAttendance': requireRole_(user, ['SuperAdmin']); return deleteRecord_('Attendance', payload.recordId, user);
    case 'deleteReport': requireRole_(user, ['SuperAdmin']); return deleteRecord_('Reports', payload.recordId, user);
    case 'deleteArchive': requireRole_(user, ['SuperAdmin']); return deleteRecord_('Archives', payload.recordId, user);
    case 'downloadArchive': return downloadArchive_(user, payload.archiveId);
    case 'saveAttendance': requireRole_(user, ['PJLP']); return saveAttendance_(user, payload);
    case 'saveReport': requireRole_(user, ['PJLP']); return saveReport_(user, payload);
    case 'saveArchive': requireRole_(user, ['PJLP']); return saveArchive_(user, payload);
    case 'approveReport': requireRole_(user, ['SuperAdmin','Verifikator']); return approveReport_(user, payload.reportId);
    case 'resetPassword': requireRole_(user, ['SuperAdmin']); return resetPassword_(user, payload);
    default: throw new Error('Aksi tidak dikenal: ' + action);
  }
}

function login_(p) {
  const identifier = String(p.identifier || '').trim().toLowerCase();
  const password = String(p.password || '');
  if (!identifier || !password) throw new Error('Email/NIK dan kata sandi wajib diisi.');
  const cache = CacheService.getScriptCache();
  const throttleKey = 'loginfail:' + sha256Hex_(identifier);
  const failCount = Number(cache.get(throttleKey) || 0);
  if (failCount >= 8) throw new Error('Terlalu banyak percobaan masuk. Coba kembali setelah 15 menit.');
  const users = readRows_('Users');
  const user = users.find(u => String(u.email || '').toLowerCase() === identifier || String(u.nik || '').toLowerCase() === identifier);
  if (!user || user.status !== 'Aktif' || !verifyPassword_(password, user.passwordHash)) {
    cache.put(throttleKey, String(failCount + 1), 900);
    throw new Error('Email/NIK atau kata sandi salah, atau akun tidak aktif.');
  }
  // Migrate old unsalted SHA-256 hashes transparently after successful authentication.
  if (!String(user.passwordHash || '').startsWith('pbkdf2-sha256$')) {
    updateById_('Users', user.id, {passwordHash:hashPassword_(password), updatedAt:new Date().toISOString()});
  }
  cache.remove(throttleKey);
  const token = Utilities.getUuid() + Utilities.getUuid();
  CacheService.getScriptCache().put('sess:' + token, JSON.stringify({id:user.id, role:user.role, bidang:user.bidang, createdAt:Date.now()}), APP.SESSION_TTL_SECONDS);
  audit_(user.id, 'LOGIN', 'Users', user.id, 'Login berhasil');
  return {ok:true, token, user:safeUser_(user, false)};
}

function register_(p) {
  ensureSheetColumns_('Users', ['pendidikanTerakhir']);
  const nama = required_(p.nama, 'Nama');
  const nik = required_(p.nik, 'NIK');
  const email = required_(p.email, 'Email').toLowerCase();
  const password = String(p.password || '');
  if (password.length < 10) throw new Error('Kata sandi minimal 10 karakter.');
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('Format email tidak valid.');
  const users = readRows_('Users');
  if (users.some(u => String(u.email).toLowerCase() === email || String(u.nik) === nik)) throw new Error('Email atau NIK sudah terdaftar.');
  const obj = makeUser_({id:'usr-' + Utilities.getUuid(), nama, nik, hp:String(p.hp||''), email, password, bidang:required_(p.bidang,'Bidang'), jabatan:required_(p.jabatan,'Jabatan'), role:'PJLP', status:'Menunggu Persetujuan', pendidikanTerakhir:String(p.pendidikanTerakhir || '')}, new Date().toISOString());
  appendObjects_('Users', [obj]);
  audit_(obj.id, 'REGISTER', 'Users', obj.id, 'Pendaftaran akun PJLP menunggu persetujuan');
  return {ok:true, message:'Pendaftaran diterima dan menunggu persetujuan Super Admin.'};
}

function getAppData_(user, options) {
  // Filter the user list BEFORE sanitizing it. Only the signed-in user's avatar
  // is read from Drive; profile images are not needed for every table row.
  let rawUsers = readRows_('Users');
  if (user.role === 'PJLP') rawUsers = rawUsers.filter(u => String(u.id) === String(user.id));
  else if (user.role === 'Verifikator') rawUsers = rawUsers.filter(u => String(u.bidang || '') === String(user.bidang || ''));
  let users = rawUsers.map(u => safeUser_(u, String(u.id) === String(user.id)));

  const pageSize = Math.min(200, Math.max(10, Number(options && options.pageSize) || 100));
  const periodDays = Math.min(365, Math.max(1, Number(options && options.periodDays) || 90));
  const attendancePage = scopedPage_('Attendance', user, rawUsers, 0, pageSize, periodDays);
  const reportsPage = scopedPage_('Reports', user, rawUsers, 0, pageSize, periodDays);
  const archivesPage = scopedPage_('Archives', user, rawUsers, 0, pageSize, periodDays);
  return {ok:true, users, attendance:attendancePage.items, reports:reportsPage.items, archives:archivesPage.items,
    pagination:{attendance:attendancePage.meta, reports:reportsPage.meta, archives:archivesPage.meta}};
}

function reviewRegistration_(actor, p) {
  const id = required_(p.userId, 'ID pendaftar');
  const decision = String(p.decision || '');
  if (!['approve','reject'].includes(decision)) throw new Error('Keputusan tidak valid.');
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const target = findById_('Users', id);
    if (!target || target.role !== 'PJLP' || target.status !== 'Menunggu Persetujuan') throw new Error('Pendaftaran tidak ditemukan atau telah diproses.');
    const status = decision === 'approve' ? 'Aktif' : 'Ditolak';
    updateById_('Users', id, {status, updatedAt:new Date().toISOString()});
    audit_(actor.id, decision === 'approve' ? 'APPROVE_REGISTRATION' : 'REJECT_REGISTRATION', 'Users', id, 'Status: ' + status);
    return {ok:true, status};
  } finally { lock.releaseLock(); }
}

// Membaca rentang dari bagian akhir sheet, bukan seluruh riwayat. Cursor adalah
// nomor baris terakhir yang belum dibaca; pagination berbasis baris stabil saat append.
function scopedPage_(name, user, visibleUsers, cursor, pageSize, periodDays) {
  const sh = sheet_(name), last = sh.getLastRow(), width = sh.getLastColumn();
  const headers = sh.getRange(1,1,1,width).getValues()[0];
  const allowedIds = new Set(visibleUsers.map(u => String(u.id)));
  const cutoff = Date.now() - periodDays * 86400000;
  let end = cursor > 1 ? Math.min(cursor, last) : last;
  const items = [];
  if (user.role === 'Verifikator' && name === 'Archives') return {items:[],meta:{nextCursor:null,hasMore:false,pageSize,periodDays}};
  while (end >= 2 && items.length < pageSize) {
    const start = Math.max(2, end - 199), count = end - start + 1;
    const rows = sh.getRange(start,1,count,width).getValues();
    for (let i = rows.length - 1; i >= 0 && items.length < pageSize; i--) {
      const row = rows[i], o = {};
      headers.forEach((key,j) => { const v=row[j]; o[key]=v instanceof Date ? v.toISOString() : v; });
      if (!o.id) continue;
      if (user.role !== 'SuperAdmin' && !allowedIds.has(String(o.userId))) continue;
      const d = o.tanggal || o.createdAt;
      if (d) { const ms = new Date(d).getTime(); if (Number.isFinite(ms) && ms < cutoff) continue; }
      if (name === 'Attendance') {o.latitude=numberOrBlank_(o.latitude);o.longitude=numberOrBlank_(o.longitude);}
      if (name === 'Reports') {o.foto=o.fotoFileId ? 'available' : '';delete o.fotoFileId;}
      if (name === 'Archives') delete o.driveFileId;
      items.push(o);
    }
    // On a full page, retain a cursor pointing before the last included row.
    if (items.length >= pageSize) {
      const lastItem = items[items.length-1];
      const idx = rows.findIndex(r => String(r[headers.indexOf('id')]) === String(lastItem.id));
      end = start + idx - 1;
      break;
    }
    end = start - 1;
  }
  return {items,meta:{nextCursor:end >= 2 ? end : null,hasMore:end >= 2,pageSize,periodDays}};
}
function getDataPage_(user, p) {
  const names = {attendance:'Attendance',reports:'Reports',archives:'Archives'};
  const type = String(p.type || ''), name=names[type];
  if (!name) throw new Error('Jenis data tidak valid.');
  const cursor = Number(p.cursor);
  if (!Number.isInteger(cursor) || cursor < 2) throw new Error('Cursor tidak valid.');
  let visibleUsers = readRows_('Users');
  if (user.role === 'PJLP') visibleUsers=visibleUsers.filter(u => String(u.id) === String(user.id));
  else if (user.role === 'Verifikator') visibleUsers=visibleUsers.filter(u => u.bidang === user.bidang);
  const page=scopedPage_(name,user,visibleUsers,cursor,Math.min(200,Math.max(10,Number(p.pageSize)||100)),Math.min(365,Math.max(1,Number(p.periodDays)||90)));
  return {ok:true,type,items:page.items,pagination:page.meta};
}

function getReportPhoto_(user, reportId) {
  const row = findById_('Reports', String(reportId || ''));
  if (!row || !row.fotoFileId) throw new Error('Foto laporan tidak ditemukan.');
  if (user.role === 'PJLP' && String(row.userId) !== String(user.id)) throw new Error('Anda tidak berwenang melihat foto laporan ini.');
  if (user.role === 'Verifikator') {
    const owner = findById_('Users', String(row.userId));
    if (!owner || owner.bidang !== user.bidang) throw new Error('Anda tidak berwenang melihat foto laporan bidang lain.');
  }
  if (!['PJLP','Verifikator','SuperAdmin'].includes(user.role)) throw new Error('Akses ditolak.');
  const foto = driveImageDataUrl_(row.fotoFileId);
  if (!foto) throw new Error('Foto tidak dapat dibaca dari Drive. Periksa file dan izin aplikasi.');
  return {ok:true, foto};
}

function saveAttendance_(user, p) {
  const status = required_(p.status, 'Status');
  const sesi = required_(p.sesi, 'Sesi presensi');
  const catatan = required_(p.catatan, 'Catatan');
  if (!['Hadir','Dinas Luar','Izin','Sakit'].includes(status)) throw new Error('Status presensi tidak valid.');
  if (!['Masuk','Pulang'].includes(sesi)) throw new Error('Sesi presensi tidak valid.');

  const now = new Date();
  const policy = getAttendancePolicy_(now, sesi);
  if (!policy.allowed) throw new Error(policy.message);

  const latitude = Number(p.latitude), longitude = Number(p.longitude);
  const accuracy = Number(p.accuracy);
  if (!Number.isFinite(latitude) || latitude < -90 || latitude > 90 ||
      !Number.isFinite(longitude) || longitude < -180 || longitude > 180) {
    throw new Error('Koordinat GPS tidak valid. Tekan Cek Lokasi dan izinkan akses lokasi.');
  }
  if (!Number.isFinite(accuracy) || accuracy <= 0 || accuracy > 30) {
    throw new Error('Akurasi GPS belum memadai (maksimal toleransi akurasi 30 meter). Aktifkan lokasi presisi dan coba lagi.');
  }

  // Calculate distance on the server; never trust a browser-supplied distance.
  const office = {lat:-3.7994516, lon:114.7790374};
  const distanceMeters = haversineMeters_(latitude, longitude, office.lat, office.lon);
  // Untuk status Hadir, server mewajibkan lokasi dalam radius 30 meter.
  if (status === 'Hadir' && distanceMeters > 30) {
    throw new Error('Presensi Hadir ditolak karena di luar radius kantor.');
  }

  const today = Utilities.formatDate(now, APP.TZ, 'yyyy-MM-dd');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const existing = readRows_('Attendance').filter(a =>
      String(a.userId) === String(user.id) &&
      String(a.tanggal) === today &&
      String(a.sesi) === sesi
    );
    if (existing.length) {
      const sesiLain = sesi === 'Masuk' ? 'Pulang' : 'Masuk';
      const otherDone = readRows_('Attendance').some(a => String(a.userId) === String(user.id) && String(a.tanggal) === today && String(a.sesi) === sesiLain);
      const detail = otherDone ? ' Presensi Masuk dan Pulang Anda hari ini sudah tercatat lengkap.' : (sesi === 'Masuk' ? ' Presensi Pulang dapat dilakukan sesuai jadwal.' : ' Pastikan presensi Masuk sudah tercatat.');
      throw new Error('Anda sudah melakukan presensi ' + sesi.toLowerCase() + ' pada ' + today + '. Presensi ganda tidak diperbolehkan.' + detail);
    }

    const selfieFileId = saveDataUrlToDrive_(p.selfie, 'Selfie_' + user.id + '_' + Date.now() + '.jpg', null);
    const signatureFileId = saveDataUrlToDrive_(p.signature, 'TandaTangan_' + user.id + '_' + Date.now() + '.png', 'image/png');
    const lateBySeconds = sesi === 'Masuk' && status === 'Hadir' ? Math.max(0, policy.lateBySeconds) : 0;
    const row = {
      id:'abs-' + Utilities.getUuid(), userId:user.id,
      tanggal:today,
      jam:Utilities.formatDate(now, APP.TZ, 'HH:mm:ss'),
      status, sesi, catatan,
      latitude, longitude, distanceMeters:Math.round(distanceMeters * 100) / 100,
      lateBySeconds,
      lateByText:lateBySeconds > 0 ? formatDuration_(lateBySeconds) : '',
      selfieFileId, signatureFileId, createdAt:now.toISOString()
    };
    appendObjects_('Attendance', [row]);
    audit_(user.id, 'CREATE', 'Attendance', row.id, 'Presensi ' + status + ' / ' + sesi + (lateBySeconds ? ' / Terlambat ' + formatDuration_(lateBySeconds) : ''));
    return {ok:true, id:row.id, lateBySeconds, lateByText:row.lateByText, distanceMeters:row.distanceMeters,
      message:lateBySeconds > 0 ? 'Presensi tercatat. Terlambat ' + row.lateByText : 'Presensi tercatat tepat waktu.'};
  } finally {
    lock.releaseLock();
  }
}

function getAttendancePolicy_(date, sesi) {
  const day = Number(Utilities.formatDate(date, APP.TZ, 'u')); // 1=Mon ... 7=Sun
  const hh = Number(Utilities.formatDate(date, APP.TZ, 'H'));
  const mm = Number(Utilities.formatDate(date, APP.TZ, 'm'));
  const ss = Number(Utilities.formatDate(date, APP.TZ, 's'));
  const nowSeconds = hh * 3600 + mm * 60 + ss;
  const start = 8 * 3600;
  if (day >= 6) return {allowed:false, message:'Presensi hanya tersedia pada hari kerja.'};
  const end = day === 5 ? 11 * 3600 + 30 * 60 : 16 * 3600 + 30 * 60;
  if (sesi === 'Masuk') {
    if (nowSeconds < start - 1800) return {allowed:false, message:'Presensi masuk dibuka pukul 07:30 WITA.'};
    if (nowSeconds > end) return {allowed:false, message:'Batas presensi masuk hari ini telah lewat.'};
    return {allowed:true, lateBySeconds:Math.max(0, nowSeconds - start)};
  }
  if (sesi === 'Pulang') {
    if (nowSeconds < end) return {allowed:false, message:'Presensi pulang belum dibuka.'};
    if (nowSeconds > end + 7200) return {allowed:false, message:'Batas presensi pulang telah berakhir.'};
    return {allowed:true, lateBySeconds:0};
  }
  return {allowed:false, message:'Sesi presensi tidak valid.'};
}
function formatDuration_(seconds) {
  seconds = Math.max(0, Math.floor(Number(seconds) || 0));
  const h = Math.floor(seconds / 3600), m = Math.floor((seconds % 3600) / 60), s = seconds % 60;
  return String(h).padStart(2,'0') + ':' + String(m).padStart(2,'0') + ':' + String(s).padStart(2,'0') + ' (jam:menit:detik)';
}
function haversineMeters_(lat1, lon1, lat2, lon2) {
  const rad = x => x * Math.PI / 180;
  const dLat = rad(lat2-lat1), dLon = rad(lon2-lon1);
  const a = Math.sin(dLat/2)**2 + Math.cos(rad(lat1))*Math.cos(rad(lat2))*Math.sin(dLon/2)**2;
  return 6371000 * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a));
}

function saveReport_(user, p) {
  const kegiatan = required_(p.kegiatan, 'Uraian kegiatan');
  const output = required_(p.output, 'Output pekerjaan');
  const tanggalInput = required_(p.tanggal, 'Tanggal');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(tanggalInput)) throw new Error('Tanggal laporan harus berformat YYYY-MM-DD.');
  const parsedDate = new Date(tanggalInput + 'T12:00:00+08:00');
  if (isNaN(parsedDate.getTime()) || Utilities.formatDate(parsedDate, APP.TZ, 'yyyy-MM-dd') !== tanggalInput) throw new Error('Tanggal laporan tidak valid.');
  const tanggal = tanggalInput;
  const jamMulai = required_(p.jamMulai, 'Jam mulai kegiatan');
  const jamSelesai = required_(p.jamSelesai, 'Jam selesai kegiatan');
  if (!/^([01]\d|2[0-3]):[0-5]\d$/.test(jamMulai) || !/^([01]\d|2[0-3]):[0-5]\d$/.test(jamSelesai)) {
    throw new Error('Jam kegiatan harus diisi dengan format HH:MM.');
  }
  if (jamSelesai <= jamMulai) throw new Error('Jam selesai harus setelah jam mulai.');
  // Ensure new columns exist even if the user has not rerun setupDatabase().
  ensureSheetColumns_('Reports', ['jamMulai','jamSelesai']);
  const fotoFileId = saveDataUrlToDrive_(p.foto, 'Dokumentasi_' + user.id + '_' + Date.now(), null);
  const row = {id:'lap-' + Utilities.getUuid(), userId:user.id, nama:user.nama, bidang:user.bidang, tanggal, jamMulai, jamSelesai, kegiatan, output, fotoFileId, status:'Pending', verifiedBy:'', verifiedAt:'', createdAt:new Date().toISOString()};
  appendObjects_('Reports', [row]);
  audit_(user.id, 'CREATE', 'Reports', row.id, 'Laporan harian dibuat (' + jamMulai + '-' + jamSelesai + ')');
  return {ok:true, id:row.id};
}

function ensureSheetColumns_(sheetName, requiredHeaders) {
  const sh = sheet_(sheetName);
  const lastCol = Math.max(1, sh.getLastColumn());
  const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  const missing = requiredHeaders.filter(h => headers.indexOf(h) < 0);
  if (missing.length) {
    sh.getRange(1, lastCol + 1, 1, missing.length).setValues([missing]);
    sh.getRange(1, lastCol + 1, 1, missing.length).setFontWeight('bold').setBackground('#dbeafe');
  }
}

function saveArchive_(user, p) {
  const jenis = required_(p.jenis, 'Jenis arsip');
  const nama = required_(p.nama, 'Nama berkas');
  const dataUrl = required_(p.fileData, 'File arsip');
  const mime = dataUrl.match(/^data:([^;,]+);base64,/);
  if (!mime || mime[1] !== 'application/pdf') throw new Error('Berkas kepegawaian wajib berformat PDF.');
  const encoded = dataUrl.substring(dataUrl.indexOf(',') + 1);
  let pdfBytes;
  try { pdfBytes = Utilities.base64Decode(encoded); } catch (e) { throw new Error('File PDF tidak valid.'); }
  if (!pdfBytes.length || pdfBytes.length > APP.MAX_ARCHIVE_BYTES) throw new Error('Ukuran PDF maksimal 1 MB per file.');
  if (Utilities.newBlob(pdfBytes).getDataAsString().substring(0, 5) !== '%PDF-') throw new Error('Isi file bukan PDF yang valid.');
  const driveFileId = saveAnyDataUrlToDrive_(dataUrl, sanitizeFilename_(nama), 'application/pdf', APP.MAX_ARCHIVE_BYTES);
  const row = {id:'ars-' + Utilities.getUuid(), userId:user.id, jenis, nama:sanitizeFilename_(nama), driveFileId,
    tanggal:Utilities.formatDate(new Date(), APP.TZ, 'yyyy-MM-dd'), status:'Tersimpan (privat)', createdAt:new Date().toISOString()};
  appendObjects_('Archives', [row]);
  audit_(user.id, 'CREATE', 'Archives', row.id, 'File arsip diunggah ke Drive privat');
  return {ok:true, id:row.id};
}

function downloadArchive_(user, archiveId) {
  const row = findById_('Archives', archiveId);
  if (!row) throw new Error('Arsip tidak ditemukan.');
  if (user.role !== 'SuperAdmin' && String(row.userId) !== String(user.id)) {
    throw new Error('Akses berkas ditolak.');
  }
  if (!row.driveFileId) throw new Error('File arsip tidak tersedia.');
  const file = DriveApp.getFileById(String(row.driveFileId));
  const blob = file.getBlob();
  const bytes = blob.getBytes();
  if (bytes.length > APP.MAX_UPLOAD_BYTES) throw new Error('File terlalu besar untuk diunduh melalui aplikasi.');
  return {ok:true, name:row.nama, mimeType:blob.getContentType(), base64:Utilities.base64Encode(bytes)};
}

function deleteRecord_(sheetName, recordId, actor) {
  const row = findById_(sheetName, recordId);
  if (!row) throw new Error('Data tidak ditemukan.');
  [row.driveFileId, row.fotoFileId, row.selfieFileId, row.signatureFileId].filter(Boolean).forEach(fileId => {
    try { DriveApp.getFileById(String(fileId)).setTrashed(true); } catch (e) {}
  });
  const sh = sheet_(sheetName), values = sh.getDataRange().getValues();
  const idCol = values[0].indexOf('id');
  for (let r=1; r<values.length; r++) {
    if (String(values[r][idCol]) === String(recordId)) {
      sh.deleteRow(r+1);
      audit_(actor.id, 'DELETE', sheetName, recordId, 'Data dihapus oleh Super Admin');
      return {ok:true};
    }
  }
  throw new Error('Data tidak ditemukan.');
}
function sanitizeFilename_(name) {
  const clean = String(name || 'dokumen').replace(/[\\/:*?"<>|\x00-\x1F]/g, '_').trim();
  return clean.slice(0, 120) || 'dokumen';
}
function saveAnyDataUrlToDrive_(dataUrl, filename, expectedMime, maxBytes) {
  const match = String(dataUrl || '').match(/^data:([^;,]+);base64,([A-Za-z0-9+/=]+)$/);
  if (!match || match[1] !== expectedMime) throw new Error('Data file tidak valid.');
  const bytes = Utilities.base64Decode(match[2]);
  const limit = Number(maxBytes || APP.MAX_UPLOAD_BYTES);
  if (!bytes.length || bytes.length > limit) throw new Error(limit === APP.MAX_ARCHIVE_BYTES ? 'Ukuran PDF maksimal 1 MB per file.' : 'Ukuran file maksimal 5 MB.');
  const file = DriveApp.createFile(Utilities.newBlob(bytes, expectedMime, filename));
  file.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
  return file.getId();
}

function approveReport_(user, reportId) {
  const row = findById_('Reports', reportId);
  if (!row) throw new Error('Laporan tidak ditemukan.');
  if (user.role === 'Verifikator') {
    const owner = findById_('Users', row.userId);
    if (!owner || owner.bidang !== user.bidang) throw new Error('Anda tidak berwenang memverifikasi laporan bidang lain.');
  }
  updateById_('Reports', reportId, {status:'Disetujui', verifiedBy:user.id, verifiedAt:new Date().toISOString()});
  audit_(user.id, 'APPROVE', 'Reports', reportId, 'Laporan disetujui');
  return {ok:true};
}

function updateProfile_(user, p) {
  ensureSheetColumns_('Users', ['pendidikanTerakhir']);
  const targetId = String(p.userId || user.id);
  if (user.role !== 'SuperAdmin' && targetId !== user.id) throw new Error('Tidak boleh mengubah profil pengguna lain.');
  if (!findById_('Users', targetId)) throw new Error('Pengguna tidak ditemukan.');
  const patch = {};
  ['ttl','tmt','hp','rekening','kawin','alamat','pendidikanTerakhir'].forEach(k => { if (p[k] !== undefined) patch[k] = String(p[k]); });
  if (p.nik !== undefined) {
    const nik = required_(p.nik, 'NIK');
    const duplicate = readRows_('Users').some(u => String(u.id) !== targetId && String(u.nik || '').trim() === nik);
    if (duplicate) throw new Error('NIK sudah digunakan oleh akun lain.');
    patch.nik = nik;
  }
  if (p.jenisKelamin !== undefined) {
    const jk = String(p.jenisKelamin || '');
    if (jk && !['Laki-laki','Perempuan'].includes(jk)) throw new Error('Jenis kelamin tidak valid.');
    patch.jenisKelamin = jk;
  }
  if (user.role === 'SuperAdmin' && p.nama) patch.nama = String(p.nama);
  if (p.avatar && String(p.avatar).startsWith('data:image/')) {
    const avatarFileId = saveDataUrlToDrive_(p.avatar, 'FotoProfil_' + targetId + '_' + Date.now(), null);
    patch.avatar = 'drive:' + avatarFileId;
  }
  patch.updatedAt = new Date().toISOString();
  updateById_('Users', targetId, patch);
  audit_(user.id, 'UPDATE', 'Users', targetId, 'Profil diperbarui');
  return {ok:true};
}

function adminCreateUser_(p, actor) {
  ensureSheetColumns_('Users', ['pendidikanTerakhir']);
  const password = String(p.password || '');
  if (password.length < 10) throw new Error('Kata sandi minimal 10 karakter.');
  const email = required_(p.email,'Email').toLowerCase();
  const nik = required_(p.nik,'NIK');
  const users = readRows_('Users');
  if (users.some(u => String(u.email).toLowerCase() === email || String(u.nik) === nik)) throw new Error('Email atau NIK sudah digunakan.');
  const allowedRoles = ['PJLP','Verifikator'];
  const requestedRole = String(p.role || 'PJLP');
  if (!allowedRoles.includes(requestedRole)) throw new Error('Peran pengguna tidak valid.');
  const row = makeUser_({id:'usr-' + Utilities.getUuid(), nama:required_(p.nama,'Nama'), nik, hp:String(p.hp||''), email, password, bidang:required_(p.bidang,'Bidang'), jabatan:required_(p.jabatan,'Jabatan'), role:requestedRole, status:'Aktif', jenisKelamin:String(p.jenisKelamin || ''), pendidikanTerakhir:String(p.pendidikanTerakhir || '')}, new Date().toISOString());
  appendObjects_('Users', [row]);
  audit_(actor.id, 'CREATE', 'Users', row.id, 'Akun baru: ' + row.email);
  return {ok:true, id:row.id};
}

function deleteUser_(id, actor) {
  if (!id || id === actor.id) throw new Error('Akun sendiri tidak dapat dihapus.');
  const target = findById_('Users', id);
  if (target && target.role === 'SuperAdmin' && readRows_('Users').filter(u => u.role === 'SuperAdmin' && u.status === 'Aktif').length <= 1) throw new Error('Super Admin aktif terakhir tidak boleh dihapus.');
  const sh = sheet_('Users'); const values = sh.getDataRange().getValues();
  const idCol = APP.SHEETS.Users.indexOf('id');
  for (let r = values.length - 1; r >= 1; r--) if (String(values[r][idCol]) === String(id)) { sh.deleteRow(r+1); audit_(actor.id,'DELETE','Users',id,'Akun dihapus'); return {ok:true}; }
  throw new Error('Pengguna tidak ditemukan.');
}

function changePassword_(user, p) {
  const currentPassword = String(p.currentPassword || '');
  const newPassword = String(p.newPassword || '');
  const confirmPassword = String(p.confirmPassword || '');
  if (!currentPassword || !newPassword || !confirmPassword) throw new Error('Semua kolom kata sandi wajib diisi.');
  if (newPassword.length < 10) throw new Error('Kata sandi baru minimal 10 karakter.');
  if (newPassword !== confirmPassword) throw new Error('Konfirmasi kata sandi baru tidak cocok.');
  if (currentPassword === newPassword) throw new Error('Kata sandi baru harus berbeda dari kata sandi saat ini.');
  const fresh = findById_('Users', user.id);
  if (!fresh || fresh.status !== 'Aktif') throw new Error('Akun tidak ditemukan atau tidak aktif.');
  if (!verifyPassword_(currentPassword, fresh.passwordHash)) throw new Error('Kata sandi saat ini salah.');
  updateById_('Users', user.id, {passwordHash:hashPassword_(newPassword), updatedAt:new Date().toISOString()});
  audit_(user.id, 'CHANGE_PASSWORD', 'Users', user.id, 'Pengguna mengganti kata sandi sendiri');
  return {ok:true, message:'Kata sandi berhasil diganti.'};
}

function resetPassword_(user, p) {
  const targetId = String(p.userId || '').trim();
  if (!targetId || !findById_('Users', targetId)) throw new Error('Akun tujuan tidak ditemukan.');
  const pass = String(p.newPassword || '');
  if (pass.length < 10) throw new Error('Kata sandi baru minimal 10 karakter.');
  updateById_('Users', targetId, {passwordHash:hashPassword_(pass), updatedAt:new Date().toISOString()});
  audit_(user.id, 'RESET_PASSWORD', 'Users', targetId, 'Kata sandi direset oleh Super Admin');
  return {ok:true};
}

function requireSession_(token) {
  if (!token) throw new Error('Sesi tidak tersedia. Silakan masuk kembali.');
  const raw = CacheService.getScriptCache().get('sess:' + token);
  if (!raw) throw new Error('Sesi kedaluwarsa. Silakan masuk kembali.');
  const sess = JSON.parse(raw);
  const user = findById_('Users', sess.id);
  if (!user || user.status !== 'Aktif') throw new Error('Akun tidak aktif.');
  return safeUser_(user);
}
function requireRole_(user, roles) { if (roles.indexOf(user.role) < 0) throw new Error('Akses ditolak.'); }
function makeUser_(p, now) {
  return {id:p.id, nama:p.nama, nik:p.nik, hp:p.hp||'', email:p.email, passwordHash:hashPassword_(p.password), bidang:p.bidang, jabatan:p.jabatan, role:p.role, status:p.status||'Aktif', ttl:p.ttl||'', tmt:p.tmt||'', avatar:p.avatar||'', rekening:p.rekening||'', kawin:p.kawin||'', alamat:p.alamat||'', jenisKelamin:p.jenisKelamin||'', pendidikanTerakhir:p.pendidikanTerakhir||'', createdAt:now, updatedAt:now};
}
function safeUser_(u, includeAvatarData) {
  const x = Object.assign({}, u); delete x.passwordHash;
  // Never read Drive for an avatar on every authenticated API call.
  // The dashboard payload requests image bytes only for the current user.
  if (typeof x.avatar === 'string' && x.avatar.indexOf('drive:') === 0) {
    x.avatar = includeAvatarData ? driveImageDataUrl_(x.avatar.slice(6)) : '';
  }
  return x;
}
// Password storage: PBKDF2-HMAC-SHA256 with unique random salt.
// Iteration count is deliberately configurable in one place; test login latency on the live GAS runtime.
const PASSWORD_HASH_ITERATIONS = 10000;
function hashPassword_(password) {
  const saltText = Utilities.getUuid() + Utilities.getUuid();
  const salt = Utilities.newBlob(saltText).getBytes();
  const derived = pbkdf2Sha256_(String(password), salt, PASSWORD_HASH_ITERATIONS, 32);
  return 'pbkdf2-sha256$' + PASSWORD_HASH_ITERATIONS + '$' + Utilities.base64Encode(salt) + '$' + Utilities.base64Encode(derived);
}
function verifyPassword_(password, storedHash) {
  const stored = String(storedHash || '');
  // Backward compatibility for existing accounts; a successful login upgrades this hash.
  if (/^[a-f0-9]{64}$/i.test(stored)) return constantTimeEqual_(stored.toLowerCase(), sha256Hex_(String(password)));
  const parts = stored.split('$');
  if (parts.length !== 4 || parts[0] !== 'pbkdf2-sha256') return false;
  const iterations = Number(parts[1]);
  if (!Number.isInteger(iterations) || iterations < 10000 || iterations > 200000) return false;
  let salt, expected;
  try { salt = Utilities.base64Decode(parts[2]); expected = Utilities.base64Decode(parts[3]); } catch (e) { return false; }
  if (!salt.length || expected.length !== 32) return false;
  const actual = Utilities.base64Encode(pbkdf2Sha256_(String(password), salt, iterations, 32));
  return constantTimeEqual_(actual, parts[3]);
}
function pbkdf2Sha256_(password, salt, iterations, keyLength) {
  const passwordBytes = Utilities.newBlob(password).getBytes();
  const output = [];
  const blockCount = Math.ceil(keyLength / 32);
  for (let block = 1; block <= blockCount; block++) {
    const blockIndex = [(block >>> 24) & 255, (block >>> 16) & 255, (block >>> 8) & 255, block & 255];
    const firstInput = salt.concat(blockIndex);
    let u = Utilities.computeHmacSha256Signature(firstInput, passwordBytes);
    const accumulator = u.map(b => (b + 256) % 256);
    for (let i = 1; i < iterations; i++) {
      u = Utilities.computeHmacSha256Signature(u, passwordBytes);
      for (let j = 0; j < accumulator.length; j++) accumulator[j] ^= ((u[j] + 256) % 256);
    }
    output.push.apply(output, accumulator);
  }
  return output.slice(0, keyLength).map(b => b > 127 ? b - 256 : b);
}
function sha256Hex_(value) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(value), Utilities.Charset.UTF_8);
  return bytes.map(b => ('0' + ((b + 256) % 256).toString(16)).slice(-2)).join('');
}
function constantTimeEqual_(a, b) {
  a = String(a); b = String(b);
  let diff = a.length ^ b.length;
  const max = Math.max(a.length, b.length);
  for (let i = 0; i < max; i++) diff |= (a.charCodeAt(i) || 0) ^ (b.charCodeAt(i) || 0);
  return diff === 0;
}
function saveDataUrlToDrive_(dataUrl, filename, mimeType) {
  if (!dataUrl || typeof dataUrl !== 'string') throw new Error('Data gambar wajib diunggah.');
  const match = dataUrl.match(/^data:(image\/(?:png|jpeg|jpg|webp));base64,([A-Za-z0-9+/=]+)$/);
  if (!match) throw new Error('Format gambar tidak valid. Gunakan PNG, JPG/JPEG, atau WEBP.');
  const actualType = match[1] === 'image/jpg' ? 'image/jpeg' : match[1];
  if (mimeType && mimeType !== actualType && !(mimeType === 'image/jpeg' && match[1] === 'image/jpg')) {
    throw new Error('Jenis file gambar tidak cocok dengan data yang diunggah.');
  }
  const bytes = Utilities.base64Decode(match[2]);
  if (!bytes.length || bytes.length > APP.MAX_UPLOAD_BYTES) throw new Error('Ukuran gambar maksimal 5 MB.');
  const blob = Utilities.newBlob(bytes, actualType, filename);
  const file = DriveApp.createFile(blob);
  file.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
  return file.getId();
}

function driveImageDataUrl_(id) {
  try {
    const file = DriveApp.getFileById(String(id));
    const blob = file.getBlob();
    const type = blob.getContentType() || 'image/jpeg';
    if (!/^image\/(png|jpeg|jpg|webp)$/i.test(type)) return '';
    const bytes = blob.getBytes();
    if (!bytes.length || bytes.length > APP.MAX_UPLOAD_BYTES) return '';
    const normalizedType = type === 'image/jpg' ? 'image/jpeg' : type;
    return 'data:' + normalizedType + ';base64,' + Utilities.base64Encode(bytes);
  } catch (e) { return ''; }
}
function driveUrl_(id) { return 'https://drive.google.com/uc?export=view&id=' + encodeURIComponent(id); }
function required_(v, label) { const s = String(v == null ? '' : v).trim(); if (!s) throw new Error(label + ' wajib diisi.'); return s; }
function numberOrBlank_(v) { return v === '' || v == null ? '' : Number(v); }

function sheet_(name) {
  const id = PropertiesService.getScriptProperties().getProperty('SIAP_DPMD_SPREADSHEET_ID');
  if (!id) throw new Error('Database belum disiapkan. Jalankan setupDatabase() dari editor Apps Script.');
  const ss = SpreadsheetApp.openById(id);
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error('Sheet database tidak ditemukan: ' + name);
  return sh;
}
function readRows_(name) {
  const sh = sheet_(name); const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0];
  return values.slice(1).filter(r => r.some(v => v !== '')).map(r => {
    const o = {};
    headers.forEach((h,i) => { let v = r[i]; if (v instanceof Date) v = v.toISOString(); o[h] = v; });
    return o;
  });
}
function appendObjects_(name, objects) {
  if (!objects || !objects.length) return;
  const sh = sheet_(name);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  const rows = objects.map(o => headers.map(h => o[h] === undefined ? '' : o[h]));
  sh.getRange(sh.getLastRow()+1, 1, rows.length, headers.length).setValues(rows);
}
function findById_(name, id) { return readRows_(name).find(o => String(o.id) === String(id)); }
function updateById_(name, id, patch) {
  const sh = sheet_(name); const values = sh.getDataRange().getValues();
  const headers = values[0]; const idCol = headers.indexOf('id');
  for (let r=1; r<values.length; r++) {
    if (String(values[r][idCol]) === String(id)) {
      Object.keys(patch).forEach(k => { const c=headers.indexOf(k); if(c>=0) sh.getRange(r+1,c+1).setValue(patch[k]); });
      return true;
    }
  }
  throw new Error('Data tidak ditemukan untuk diperbarui.');
}
function audit_(actorId, action, entity, entityId, details) {
  try { appendObjects_('AuditLog', [{id:'log-' + Utilities.getUuid(), actorId, action, entity, entityId, details, createdAt:new Date().toISOString()}]); } catch (e) { console.error(e); }
}
