/**
 * ============================================================================
 * KSL BIO LINK - GOOGLE APPS SCRIPT BACKEND (Code.gs)
 * Version: 5.0.0 (Multi-Project Enterprise Edition)
 * Multi-Project Manager, Smart Campaign CTA, Password Protection,
 * Drag & Drop Reordering, Analytics & Enterprise User Management
 * ============================================================================
 */

const APP_VERSION = '5.0.0';
const SS = SpreadsheetApp.getActiveSpreadsheet();

// Daftar Nama Sheet
const SHEET_PROJECTS = 'PROJECTS';
const SHEET_PROFILE = 'PROFILE';
const SHEET_LINKS = 'LINKS';
const SHEET_ADMIN_USERS = 'ADMIN_USERS';
const SHEET_CLICK_LOG = 'CLICK_LOG';
const SHEET_PAGE_VIEW_LOG = 'PAGE_VIEW_LOG';
const SHEET_CLICK_ARCHIVE_DAILY = 'CLICK_ARCHIVE_DAILY';
const SHEET_PAGE_VIEW_ARCHIVE_DAILY = 'PAGE_VIEW_ARCHIVE_DAILY';

// Konfigurasi Keamanan & Cache
const ADMIN_PASSWORD_PREFIX_ = 'KSLPB2';
const ADMIN_PASSWORD_PEPPER_PROP_ = 'KSL_ADMIN_PASSWORD_PEPPER_V2';
const CACHE_TTL_SECONDS_ = 60;

/**
 * Endpoint Utama Web App (doGet)
 */
function doGet(e) {
  ensureAppReadyOnce_();

  const parameter = (e && e.parameter) ? e.parameter : {};
  const page = parameter.page ? String(parameter.page).toLowerCase() : 'home';
  const bioSlug = parameter.bio ? String(parameter.bio).toLowerCase() : '';

  const template = HtmlService.createTemplateFromFile('Index');
  template.page = page;
  template.bioSlug = bioSlug;
  template.version = APP_VERSION;

  return template
    .evaluate()
    .setTitle('KSL Bio Link v5.0 - Multi-Project Manager')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Mengambil data publik untuk pengunjung (berdasarkan slug bio link atau default)
 */
function getPublicData(bioSlug) {
  const start = Date.now();
  const slug = String(bioSlug || 'utama').trim().toLowerCase();
  const cacheKey = 'KSL_PUBLIC_DATA_V5_' + slug;
  const cache = CacheService.getScriptCache();
  
  const cached = cache.get(cacheKey);
  if (cached) {
    try {
      const parsed = JSON.parse(cached);
      parsed.fromCache = true;
      return parsed;
    } catch (err) {}
  }

  try {
    const project = getProjectBySlug_(slug);
    if (!project) {
      return { success: false, message: 'Halaman Bio Link tidak ditemukan.' };
    }

    const rawLinks = getLinksByProject_(project.ID_PROJECT);
    const activeLinks = rawLinks
      .filter(isLinkPubliclyVisible_)
      .sort((a, b) => {
        const featA = String(a.FEATURED || '').toUpperCase() === 'YES' ? 1 : 0;
        const featB = String(b.FEATURED || '').toUpperCase() === 'YES' ? 1 : 0;
        if (featA !== featB) return featB - featA;
        return Number(a.SORT_ORDER || 999) - Number(b.SORT_ORDER || 999);
      })
      .map(item => ({
        ID_LINK: item.ID_LINK,
        ID_PROJECT: item.ID_PROJECT,
        TITLE: item.TITLE,
        URL: item.URL,
        ICON: item.ICON || '🔗',
        SORT_ORDER: Number(item.SORT_ORDER || 999),
        STATUS: item.STATUS,
        CTA_TYPE: item.CTA_TYPE || 'LINK',
        CAMPAIGN: item.CAMPAIGN || '',
        FEATURED: item.FEATURED || 'NO',
        IS_LOCKED: String(item.IS_LOCKED || '').toUpperCase() === 'YES',
        HAS_PASSWORD: Boolean(item.PASSWORD && item.PASSWORD.trim() !== ''),
        START_AT: item.START_AT || '',
        END_AT: item.END_AT || ''
      }));

    const result = {
      success: true,
      version: APP_VERSION,
      project: project,
      links: activeLinks,
      allProjects: getAllPublicProjectsBrief_(),
      executionTimeMs: Date.now() - start
    };

    cache.put(cacheKey, JSON.stringify(result), CACHE_TTL_SECONDS_);
    return result;
  } catch (err) {
    return { success: false, message: err.message, stack: err.stack };
  }
}

/**
 * Verifikasi kata sandi tautan yang diproteksi dari sisi backend
 */
function verifyProtectedLink(idLink, enteredPassword) {
  idLink = String(idLink || '').trim();
  enteredPassword = String(enteredPassword || '');

  const sh = SS.getSheetByName(SHEET_LINKS);
  if (!sh) throw new Error('Sheet LINKS tidak ditemukan.');

  const values = sh.getDataRange().getValues();
  if (values.length < 2) throw new Error('Data link kosong.');

  const headers = values[0].map(h => String(h || '').trim());
  const idIdx = headers.indexOf('ID_LINK');
  const passIdx = headers.indexOf('PASSWORD');
  const urlIdx = headers.indexOf('URL');
  const lockedIdx = headers.indexOf('IS_LOCKED');

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idIdx] || '').trim() === idLink) {
      const isLocked = String(values[i][lockedIdx] || '').toUpperCase() === 'YES';
      const actualPass = String(values[i][passIdx] || '');

      if (!isLocked || actualPass === '' || enteredPassword === actualPass) {
        return {
          success: true,
          url: values[i][urlIdx] || ''
        };
      } else {
        return { success: false, message: 'Password salah!' };
      }
    }
  }

  return { success: false, message: 'Link tidak ditemukan.' };
}

/**
 * Mengambil paket data lengkap dashboard admin
 */
function getAdminData(token, targetProjectId) {
  const adminUser = requireAdmin_(token);
  ensureAppReady_();

  const allProjects = getAllProjects_();
  let currentProject = null;

  if (targetProjectId) {
    currentProject = allProjects.find(p => p.ID_PROJECT === targetProjectId);
  }
  if (!currentProject && allProjects.length > 0) {
    currentProject = allProjects[0];
  }

  const projectId = currentProject ? currentProject.ID_PROJECT : '';
  const links = projectId ? getLinksByProject_(projectId) : [];
  const deletedLinks = projectId ? getDeletedLinksByProject_(projectId) : [];
  const clickMap = getClickStatsMap_(projectId);
  const viewStats = getPageViewStats_(projectId);

  const linksWithStats = links.map(l => {
    l.TOTAL_CLICK = clickMap[l.ID_LINK] || 0;
    return l;
  });

  const totalClicks = Object.keys(clickMap).reduce((sum, k) => sum + (clickMap[k] || 0), 0);
  const users = getAllAdminUsers_();

  return {
    success: true,
    version: APP_VERSION,
    currentUser: {
      email: adminUser.email,
      name: adminUser.name,
      role: adminUser.role
    },
    activeProject: currentProject,
    allProjects: allProjects,
    links: linksWithStats,
    deletedLinks: deletedLinks,
    users: users,
    summary: {
      totalLinks: links.length,
      activeLinks: links.filter(l => String(l.STATUS).toUpperCase() === 'ACTIVE').length,
      totalClicks: totalClicks,
      totalViews: viewStats.totalViews,
      ctr: viewStats.totalViews > 0 ? Number(((totalClicks / viewStats.totalViews) * 100).toFixed(1)) : 0,
      recycleBinCount: deletedLinks.length
    },
    analytics: getClickAnalytics_(30, projectId)
  };
}

/**
 * Menyimpan tautan baru atau memperbarui tautan yang sudah ada
 */
function saveLink(payload, token) {
  requireAdmin_(token);
  payload = payload || {};

  const title = String(payload.TITLE || '').trim();
  const url = String(payload.URL || '').trim();
  const projectId = String(payload.ID_PROJECT || 'bio-utama').trim();

  if (!title) throw new Error('Judul link wajib diisi.');
  if (!url) throw new Error('URL wajib diisi.');

  return withWriteLock_('saveLink', function() {
    ensureLinksSheet_();
    const sh = SS.getSheetByName(SHEET_LINKS);
    const values = sh.getDataRange().getValues();
    const headers = values[0].map(h => String(h || '').trim());
    const headerMap = {};
    headers.forEach((h, idx) => headerMap[h] = idx);

    const idLink = String(payload.ID_LINK || '').trim();
    const now = formatDateTime_(new Date());

    const record = {
      ID_LINK: idLink || generateId_('LNK'),
      ID_PROJECT: projectId,
      TITLE: title,
      URL: url,
      ICON: String(payload.ICON || '🔗').trim(),
      SORT_ORDER: Number(payload.SORT_ORDER) || 999,
      STATUS: String(payload.STATUS || 'ACTIVE').toUpperCase(),
      CTA_TYPE: String(payload.CTA_TYPE || 'LINK').toUpperCase(),
      CAMPAIGN: String(payload.CAMPAIGN || '').trim(),
      FEATURED: String(payload.FEATURED || 'NO').toUpperCase() === 'YES' ? 'YES' : 'NO',
      IS_LOCKED: Boolean(payload.IS_LOCKED) ? 'YES' : 'NO',
      PASSWORD: String(payload.PASSWORD || '').trim(),
      START_AT: normalizeScheduleInput_(payload.START_AT),
      END_AT: normalizeScheduleInput_(payload.END_AT),
      IS_DELETED: 'NO',
      DELETED_AT: '',
      CREATED_AT: now,
      UPDATED_AT: now
    };

    if (idLink) {
      for (let r = 1; r < values.length; r++) {
        if (String(values[r][headerMap['ID_LINK']] || '').trim() === idLink) {
          const rowData = values[r].slice();
          record.CREATED_AT = values[r][headerMap['CREATED_AT']] || now;

          Object.keys(record).forEach(k => {
            if (headerMap[k] !== undefined) rowData[headerMap[k]] = record[k];
          });

          sh.getRange(r + 1, 1, 1, headers.length).setValues([rowData]);
          clearPublicCache_();
          return { success: true, message: 'Link berhasil diperbarui.' };
        }
      }
      throw new Error('Link yang akan diedit tidak ditemukan.');
    } else {
      const newRow = headers.map(h => record[h] !== undefined ? record[h] : '');
      sh.appendRow(newRow);
      clearPublicCache_();
      return { success: true, message: 'Link baru berhasil ditambahkan.' };
    }
  });
}

/**
 * Menyimpan urutan baru (Drag & Drop Reordering)
 */
function saveLinkOrder(items, token) {
  requireAdmin_(token);
  if (!Array.isArray(items) || items.length === 0) {
    throw new Error('Daftar urutan kosong.');
  }

  return withWriteLock_('saveLinkOrder', function() {
    const sh = SS.getSheetByName(SHEET_LINKS);
    if (!sh) throw new Error('Sheet LINKS tidak ditemukan.');

    const values = sh.getDataRange().getValues();
    const headers = values[0].map(h => String(h || '').trim());
    const idIdx = headers.indexOf('ID_LINK');
    const sortIdx = headers.indexOf('SORT_ORDER');

    if (idIdx < 0 || sortIdx < 0) throw new Error('Kolom ID_LINK atau SORT_ORDER tidak ditemukan.');

    const orderMap = {};
    items.forEach((item, index) => {
      const id = String(item.ID_LINK || '').trim();
      if (id) orderMap[id] = Number(item.SORT_ORDER) || (index + 1);
    });

    const sortRange = sh.getRange(2, sortIdx + 1, values.length - 1, 1);
    const sortValues = sortRange.getValues();

    for (let r = 1; r < values.length; r++) {
      const currentId = String(values[r][idIdx] || '').trim();
      if (orderMap[currentId] !== undefined) {
        sortValues[r - 1][0] = orderMap[currentId];
      }
    }

    sortRange.setValues(sortValues);
    clearPublicCache_();
    return { success: true, message: 'Urutan link berhasil disimpan.' };
  });
}

/**
 * Operasi massal (Bulk Actions: ACTIVATE, DEACTIVATE, RECYCLE)
 */
function bulkUpdateLinks(payload, token) {
  requireAdmin_(token);
  payload = payload || {};
  const action = String(payload.action || '').toUpperCase();
  const ids = Array.isArray(payload.ids) ? payload.ids : [];

  if (!ids.length) throw new Error('Pilih minimal satu link.');

  return withWriteLock_('bulkUpdateLinks', function() {
    const sh = SS.getSheetByName(SHEET_LINKS);
    const values = sh.getDataRange().getValues();
    const headers = values[0].map(h => String(h || '').trim());
    const idIdx = headers.indexOf('ID_LINK');
    const statusIdx = headers.indexOf('STATUS');
    const delIdx = headers.indexOf('IS_DELETED');
    const delAtIdx = headers.indexOf('DELETED_AT');
    const updIdx = headers.indexOf('UPDATED_AT');
    const now = formatDateTime_(new Date());

    const idSet = new Set(ids);
    let updatedCount = 0;

    for (let r = 1; r < values.length; r++) {
      const rowId = String(values[r][idIdx] || '').trim();
      if (idSet.has(rowId)) {
        if (action === 'ACTIVATE') {
          sh.getRange(r + 1, statusIdx + 1).setValue('ACTIVE');
        } else if (action === 'DEACTIVATE') {
          sh.getRange(r + 1, statusIdx + 1).setValue('INACTIVE');
        } else if (action === 'RECYCLE') {
          sh.getRange(r + 1, delIdx + 1).setValue('YES');
          sh.getRange(r + 1, delAtIdx + 1).setValue(now);
        }
        if (updIdx >= 0) sh.getRange(r + 1, updIdx + 1).setValue(now);
        updatedCount++;
      }
    }

    clearPublicCache_();
    return { success: true, count: updatedCount, message: updatedCount + ' link berhasil diproses.' };
  });
}

/**
 * Memindahkan tautan ke Recycle Bin (Soft Delete)
 */
function deleteLink(idLink, token) {
  requireAdmin_(token);
  idLink = String(idLink || '').trim();

  return withWriteLock_('deleteLink', function() {
    const sh = SS.getSheetByName(SHEET_LINKS);
    const values = sh.getDataRange().getValues();
    const headers = values[0].map(h => String(h || '').trim());
    const idIdx = headers.indexOf('ID_LINK');
    const delIdx = headers.indexOf('IS_DELETED');
    const delAtIdx = headers.indexOf('DELETED_AT');

    for (let r = 1; r < values.length; r++) {
      if (String(values[r][idIdx] || '').trim() === idLink) {
        sh.getRange(r + 1, delIdx + 1).setValue('YES');
        sh.getRange(r + 1, delAtIdx + 1).setValue(formatDateTime_(new Date()));
        clearPublicCache_();
        return { success: true, message: 'Link dipindahkan ke Recycle Bin.' };
      }
    }
    throw new Error('Link tidak ditemukan.');
  });
}

/**
 * Memulihkan tautan dari Recycle Bin (Restore)
 */
function restoreLink(idLink, token) {
  requireAdmin_(token);
  idLink = String(idLink || '').trim();

  return withWriteLock_('restoreLink', function() {
    const sh = SS.getSheetByName(SHEET_LINKS);
    const values = sh.getDataRange().getValues();
    const headers = values[0].map(h => String(h || '').trim());
    const idIdx = headers.indexOf('ID_LINK');
    const delIdx = headers.indexOf('IS_DELETED');
    const delAtIdx = headers.indexOf('DELETED_AT');

    for (let r = 1; r < values.length; r++) {
      if (String(values[r][idIdx] || '').trim() === idLink) {
        sh.getRange(r + 1, delIdx + 1).setValue('NO');
        sh.getRange(r + 1, delAtIdx + 1).setValue('');
        clearPublicCache_();
        return { success: true, message: 'Link berhasil dipulihkan.' };
      }
    }
    throw new Error('Link tidak ditemukan.');
  });
}

/**
 * Menghapus permanen dari database
 */
function deleteLinkPermanently(idLink, token) {
  requireAdmin_(token);
  idLink = String(idLink || '').trim();

  return withWriteLock_('deleteLinkPermanently', function() {
    const sh = SS.getSheetByName(SHEET_LINKS);
    const values = sh.getDataRange().getValues();
    const headers = values[0].map(h => String(h || '').trim());
    const idIdx = headers.indexOf('ID_LINK');

    for (let r = 1; r < values.length; r++) {
      if (String(values[r][idIdx] || '').trim() === idLink) {
        sh.deleteRow(r + 1);
        clearPublicCache_();
        return { success: true, message: 'Link dihapus selamanya.' };
      }
    }
    throw new Error('Link tidak ditemukan.');
  });
}

/**
 * Login Administrator
 */
function adminLogin(payload) {
  ensureAppReady_();
  payload = payload || {};
  const email = String(payload.email || '').trim().toLowerCase();
  const password = String(payload.password || '');

  if (!email || !password) throw new Error('Email dan password wajib diisi.');

  const sh = SS.getSheetByName(SHEET_ADMIN_USERS);
  const values = sh.getDataRange().getValues();
  const headers = values[0].map(h => String(h || '').trim());
  const emailIdx = headers.indexOf('EMAIL');
  const passIdx = headers.indexOf('PASSWORD');
  const nameIdx = headers.indexOf('NAME');
  const roleIdx = headers.indexOf('ROLE');
  const statusIdx = headers.indexOf('STATUS');

  for (let i = 1; i < values.length; i++) {
    const rowEmail = String(values[i][emailIdx] || '').trim().toLowerCase();
    const storedPass = String(values[i][passIdx] || '');
    const rowStatus = String(values[i][statusIdx] || '').toUpperCase();

    if (rowEmail === email && rowStatus === 'ACTIVE') {
      if (verifyAdminPassword_(password, storedPass)) {
        const token = generateAdminToken_(email);
        const userData = {
          email: email,
          name: values[i][nameIdx] || 'Admin',
          role: values[i][roleIdx] || 'SUPER_ADMIN'
        };

        CacheService.getScriptCache().put(
          'ADMIN_TOKEN_' + token,
          JSON.stringify(userData),
          21600 // 6 Jam sesi aktif
        );

        return { success: true, token: token, user: userData, message: 'Login berhasil.' };
      }
    }
  }

  throw new Error('Email atau password tidak sesuai.');
}

/**
 * Logout Administrator dari Google Apps Script
 */
function adminLogout(token) {
  token = String(token || '').trim();
  if (token) {
    CacheService.getScriptCache().remove('ADMIN_TOKEN_' + token);
  }
  return { success: true, message: 'Logout berhasil.' };
}

/**
 * Endpoint verifikasi status sesi admin dari antarmuka Web App
 */
function adminMe(token) {
  try {
    const user = requireAdmin_(token);
    return { success: true, user: user };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

/**
 * Wrapper public callable untuk requireAdmin_
 */
function requireAdmin(token) {
  return adminMe(token);
}

/**
 * Simpan / Buat Akun Administrator Baru
 */
function saveAdminUser(payload, token) {
  const currentAdmin = requireAdmin_(token);
  if (currentAdmin.role !== 'SUPER_ADMIN') throw new Error('Hanya Super Admin yang dapat mengelola akun pengguna.');

  payload = payload || {};
  const email = String(payload.EMAIL || '').trim().toLowerCase();
  const name = String(payload.NAME || '').trim();
  const password = String(payload.PASSWORD || '');
  const role = String(payload.ROLE || 'EDITOR').toUpperCase();
  const status = String(payload.STATUS || 'ACTIVE').toUpperCase();
  const idUser = String(payload.ID_USER || '').trim();

  if (!email || !name) throw new Error('Nama dan email wajib diisi.');

  return withWriteLock_('saveAdminUser', function() {
    const sh = SS.getSheetByName(SHEET_ADMIN_USERS);
    const values = sh.getDataRange().getValues();
    const headers = values[0].map(h => String(h || '').trim());
    const headerMap = {};
    headers.forEach((h, idx) => headerMap[h] = idx);

    if (idUser) {
      for (let r = 1; r < values.length; r++) {
        if (String(values[r][headerMap['ID_USER']] || '').trim() === idUser) {
          sh.getRange(r + 1, headerMap['NAME'] + 1).setValue(name);
          sh.getRange(r + 1, headerMap['EMAIL'] + 1).setValue(email);
          sh.getRange(r + 1, headerMap['ROLE'] + 1).setValue(role);
          sh.getRange(r + 1, headerMap['STATUS'] + 1).setValue(status);
          if (password) {
            sh.getRange(r + 1, headerMap['PASSWORD'] + 1).setValue(createAdminPasswordHash_(password));
          }
          return { success: true, message: 'Akun admin berhasil diperbarui.' };
        }
      }
    } else {
      if (!password) throw new Error('Password wajib diisi untuk akun baru.');
      const newRow = [
        generateId_('USR'),
        email,
        createAdminPasswordHash_(password),
        name,
        role,
        status,
        formatDateTime_(new Date())
      ];
      sh.appendRow(newRow);
      return { success: true, message: 'Akun admin baru berhasil dibuat.' };
    }
  });
}

/**
 * Pencatatan Klik Tautan (Click Tracker)
 */
function recordLinkClick(payload) {
  payload = payload || {};
  const idLink = String(payload.ID_LINK || '').trim();
  const idProject = String(payload.ID_PROJECT || 'utama').trim();

  if (!idLink) return { success: false };

  ensureClickLogSheet_();
  const sh = SS.getSheetByName(SHEET_CLICK_LOG);
  if (!sh) return { success: false };

  sh.appendRow([
    generateId_('CLK'),
    idProject,
    idLink,
    String(payload.TITLE || ''),
    String(payload.URL || ''),
    formatDateTime_(new Date()),
    String(payload.USER_AGENT || '').slice(0, 500)
  ]);

  clearAnalyticsCache_();
  return { success: true };
}

/**
 * Pencatatan Kunjungan Halaman (Page View Tracker)
 */
function recordPageView(payload) {
  payload = payload || {};
  const idProject = String(payload.ID_PROJECT || 'utama').trim();

  ensurePageViewLogSheet_();
  const sh = SS.getSheetByName(SHEET_PAGE_VIEW_LOG);
  if (!sh) return { success: false };

  sh.appendRow([
    generateId_('VIEW'),
    idProject,
    formatDateTime_(new Date()),
    String(payload.USER_AGENT || '').slice(0, 500)
  ]);

  clearAnalyticsCache_();
  return { success: true };
}

// ============================================================================
// FUNGSI HELPER INTERNAL, DATABASE INTI & KEAMANAN
// ============================================================================

function requireAdmin_(token) {
  token = String(token || '').trim();
  if (!token) throw new Error('Token sesi tidak ditemukan. Silakan login kembali.');
  const raw = CacheService.getScriptCache().get('ADMIN_TOKEN_' + token);
  if (!raw) throw new Error('Sesi telah berakhir. Silakan login kembali.');
  return JSON.parse(raw);
}

function getProjectBySlug_(slug) {
  ensureProjectsSheet_();
  const sh = SS.getSheetByName(SHEET_PROJECTS);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return null;
  const headers = values[0].map(h => String(h || '').trim());
  const slugIdx = headers.indexOf('SLUG');

  for (let r = 1; r < values.length; r++) {
    if (String(values[r][slugIdx] || '').trim().toLowerCase() === slug) {
      const obj = {};
      headers.forEach((h, i) => obj[h] = values[r][i]);
      return obj;
    }
  }
  return null;
}

function getAllProjects_() {
  ensureProjectsSheet_();
  const sh = SS.getSheetByName(SHEET_PROJECTS);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0].map(h => String(h || '').trim());

  return values.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}

function getAllPublicProjectsBrief_() {
  return getAllProjects_().map(p => ({
    id: p.ID_PROJECT,
    name: p.NAME,
    slug: p.SLUG
  }));
}

function getLinksByProject_(projectId) {
  ensureLinksSheet_();
  const sh = SS.getSheetByName(SHEET_LINKS);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0].map(h => String(h || '').trim());
  const projIdx = headers.indexOf('ID_PROJECT');
  const delIdx = headers.indexOf('IS_DELETED');

  return values.slice(1)
    .filter(row => {
      const matchProj = !projectId || String(row[projIdx] || '').trim() === projectId;
      const isDel = String(row[delIdx] || '').toUpperCase() === 'YES';
      return matchProj && !isDel;
    })
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      return obj;
    });
}

function getDeletedLinksByProject_(projectId) {
  ensureLinksSheet_();
  const sh = SS.getSheetByName(SHEET_LINKS);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0].map(h => String(h || '').trim());
  const projIdx = headers.indexOf('ID_PROJECT');
  const delIdx = headers.indexOf('IS_DELETED');

  return values.slice(1)
    .filter(row => {
      const matchProj = !projectId || String(row[projIdx] || '').trim() === projectId;
      const isDel = String(row[delIdx] || '').toUpperCase() === 'YES';
      return matchProj && isDel;
    })
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      return obj;
    });
}

function getAllAdminUsers_() {
  ensureAdminUsersSheet_();
  const sh = SS.getSheetByName(SHEET_ADMIN_USERS);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0].map(h => String(h || '').trim());
  
  return values.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => {
      if (h !== 'PASSWORD') obj[h] = row[i];
    });
    return obj;
  });
}

function isLinkPubliclyVisible_(link) {
  if (String(link.STATUS || '').toUpperCase() !== 'ACTIVE') return false;
  const now = new Date();
  const start = parseScheduleDate_(link.START_AT);
  const end = parseScheduleDate_(link.END_AT);
  if (start && now < start) return false;
  if (end && now > end) return false;
  return true;
}

function parseScheduleDate_(val) {
  if (!val) return null;
  if (val instanceof Date) return val;
  try {
    return new Date(String(val).replace(' ', 'T'));
  } catch (e) {
    return null;
  }
}

function normalizeScheduleInput_(val) {
  if (!val) return '';
  return String(val).trim().replace('T', ' ').substring(0, 16);
}

function generateId_(prefix) {
  return prefix + '-' + Utilities.getUuid().slice(0, 8).toUpperCase();
}

function formatDateTime_(date) {
  return Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}

function withWriteLock_(label, callback) {
  const lock = LockService.getScriptLock();
  const success = lock.tryLock(10000);
  if (!success) throw new Error('Sistem sedang sibuk. Silakan coba kembali.');
  try {
    return callback();
  } finally {
    lock.releaseLock();
  }
}

function clearPublicCache_() {
  CacheService.getScriptCache().removeAll([
    'KSL_PUBLIC_DATA_utama',
    'KSL_ANALYTICS_CACHE_STATS'
  ]);
}

function clearAnalyticsCache_() {
  CacheService.getScriptCache().remove('KSL_ANALYTICS_CACHE_STATS');
}

function getClickStatsMap_(projectId) {
  const map = {};
  const sh = SS.getSheetByName(SHEET_CLICK_LOG);
  if (!sh || sh.getLastRow() < 2) return map;
  const values = sh.getDataRange().getValues();
  const headers = values[0].map(h => String(h || '').trim());
  const idIdx = headers.indexOf('ID_LINK');
  const projIdx = headers.indexOf('ID_PROJECT');

  for (let r = 1; r < values.length; r++) {
    if (!projectId || String(values[r][projIdx] || '') === projectId) {
      const id = String(values[r][idIdx] || '');
      if (id) map[id] = (map[id] || 0) + 1;
    }
  }
  return map;
}

function getPageViewStats_(projectId) {
  let total = 0;
  const sh = SS.getSheetByName(SHEET_PAGE_VIEW_LOG);
  if (!sh || sh.getLastRow() < 2) return { totalViews: 0 };
  const values = sh.getDataRange().getValues();
  const headers = values[0].map(h => String(h || '').trim());
  const projIdx = headers.indexOf('ID_PROJECT');

  for (let r = 1; r < values.length; r++) {
    if (!projectId || String(values[r][projIdx] || '') === projectId) total++;
  }
  return { totalViews: total };
}

function getClickAnalytics_(days, projectId) {
  return {
    rangeDays: days,
    totalClicks: 0,
    dailySeries: []
  };
}

function ensureAppReadyOnce_() {
  ensureProjectsSheet_();
  ensureLinksSheet_();
  ensureAdminUsersSheet_();
  ensureClickLogSheet_();
  ensurePageViewLogSheet_();
}

function ensureAppReady() {
  ensureAppReadyOnce_();
}

function ensureProjectsSheet_() {
  let sh = SS.getSheetByName(SHEET_PROJECTS);
  if (!sh) {
    sh = SS.insertSheet(SHEET_PROJECTS);
    const headers = [
      'ID_PROJECT', 'SLUG', 'NAME', 'BIO', 'AVATAR_URL', 'THEME_COLOR',
      'PUBLIC_BG_TYPE', 'PUBLIC_BG_URL', 'FRAME_STYLE', 'CREATED_AT'
    ];
    sh.appendRow(headers);
    sh.appendRow([
      'bio-utama', 'utama', 'Kucing Storia Labs',
      'Web App, Google Sheets, Apps Script & Bio Link Solution',
      'https://placehold.co/300x300?text=KSL', '#2563eb',
      'SOFT', '', 'SOFT', formatDateTime_(new Date())
    ]);
    sh.setFrozenRows(1);
  }
}

function ensureLinksSheet_() {
  let sh = SS.getSheetByName(SHEET_LINKS);
  if (!sh) {
    sh = SS.insertSheet(SHEET_LINKS);
    const headers = [
      'ID_LINK', 'ID_PROJECT', 'TITLE', 'URL', 'ICON', 'SORT_ORDER', 'STATUS',
      'CTA_TYPE', 'CAMPAIGN', 'FEATURED', 'IS_LOCKED', 'PASSWORD',
      'START_AT', 'END_AT', 'IS_DELETED', 'DELETED_AT', 'CREATED_AT', 'UPDATED_AT'
    ];
    sh.appendRow(headers);
    sh.setFrozenRows(1);
  }
}

function ensureAdminUsersSheet_() {
  let sh = SS.getSheetByName(SHEET_ADMIN_USERS);
  if (!sh) {
    sh = SS.insertSheet(SHEET_ADMIN_USERS);
    const headers = ['ID_USER', 'EMAIL', 'PASSWORD', 'NAME', 'ROLE', 'STATUS', 'CREATED_AT'];
    sh.appendRow(headers);
    sh.appendRow([
      'USR-01',
      'admin@ksl.local',
      createAdminPasswordHash_('admin123'),
      'Super Administrator',
      'SUPER_ADMIN',
      'ACTIVE',
      formatDateTime_(new Date())
    ]);
    sh.setFrozenRows(1);
  }
}

function ensureClickLogSheet_() {
  let sh = SS.getSheetByName(SHEET_CLICK_LOG);
  if (!sh) {
    sh = SS.insertSheet(SHEET_CLICK_LOG);
    sh.appendRow(['LOG_ID', 'ID_PROJECT', 'ID_LINK', 'TITLE', 'URL', 'CLICKED_AT', 'USER_AGENT']);
    sh.setFrozenRows(1);
  }
}

function ensurePageViewLogSheet_() {
  let sh = SS.getSheetByName(SHEET_PAGE_VIEW_LOG);
  if (!sh) {
    sh = SS.insertSheet(SHEET_PAGE_VIEW_LOG);
    sh.appendRow(['VIEW_ID', 'ID_PROJECT', 'VIEWED_AT', 'USER_AGENT']);
    sh.setFrozenRows(1);
  }
}

// Keamanan Enkripsi Password (HMAC-SHA256)
function getPepper_() {
  const props = PropertiesService.getScriptProperties();
  let pepper = props.getProperty(ADMIN_PASSWORD_PEPPER_PROP_);
  if (!pepper) {
    pepper = Utilities.getUuid().replace(/-/g, '');
    props.setProperty(ADMIN_PASSWORD_PEPPER_PROP_, pepper);
  }
  return pepper;
}

function createAdminPasswordHash_(password) {
  const salt = Utilities.getUuid().replace(/-/g, '');
  const digest = Utilities.computeHmacSha256Signature(
    salt + '|' + String(password),
    getPepper_(),
    Utilities.Charset.UTF_8
  );
  return ADMIN_PASSWORD_PREFIX_ + '$' + salt + '$' + Utilities.base64EncodeWebSafe(digest);
}

function verifyAdminPassword_(password, storedHash) {
  if (!storedHash) return false;
  if (!storedHash.includes('$')) return password === storedHash;

  const parts = storedHash.split('$');
  if (parts.length !== 3 || parts[0] !== ADMIN_PASSWORD_PREFIX_) return false;

  const salt = parts[1];
  const expectedHash = parts[2];
  const calculated = Utilities.base64EncodeWebSafe(
    Utilities.computeHmacSha256Signature(
      salt + '|' + String(password),
      getPepper_(),
      Utilities.Charset.UTF_8
    )
  );
  return calculated === expectedHash;
}