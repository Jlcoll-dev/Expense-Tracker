'use strict';

/* ══════════════════════════════════════════════════════
   DRIVE.JS — Google Drive sync para Finanzas PWA
   
   Flujo:
   1. Usuario toca "Conectar" → popup OAuth Google
   2. App obtiene access_token (scope: drive.file)
   3. Al guardar un movimiento → sube finanzas-data.json a Drive
   4. Al abrir la app → baja Drive y hace merge por ID
   
   IMPORTANTE: reemplazá GOOGLE_CLIENT_ID con el tuyo.
   Ver README.md sección "Configurar Google OAuth"
   ══════════════════════════════════════════════════════ */

const GOOGLE_CLIENT_ID = 'TU_CLIENT_ID_AQUI.apps.googleusercontent.com';
const DRIVE_SCOPE      = 'https://www.googleapis.com/auth/drive.file';
const DRIVE_FILE_NAME  = 'finanzas-data.json';
const DRIVE_FILE_KEY   = 'finanzas_drive_file_id';
const DRIVE_TOKEN_KEY  = 'finanzas_drive_token';

let driveFileId   = localStorage.getItem(DRIVE_FILE_KEY) || null;
let accessToken   = null;
let tokenClient   = null;
let syncPending   = false;

/* ── UI helpers ── */
function showToast(msg, type = 'info') {
  const el = document.getElementById('sync-toast');
  el.textContent = msg;
  el.className = `sync-toast show ${type}`;
  clearTimeout(el._t);
  el._t = setTimeout(() => el.classList.remove('show'), 3000);
}

function setDriveStatus(connected, text = '', syncing = false) {
  const banner   = document.getElementById('drive-banner');
  const topRow   = document.getElementById('drive-banner-content');
  const statusRow = document.getElementById('drive-status-row');
  const dot      = document.getElementById('drive-dot');
  const statusTxt = document.getElementById('drive-status-text');

  if (!banner) return;

  if (connected) {
    topRow.style.display    = 'none';
    statusRow.style.display = 'flex';
    dot.className = 'drive-dot ' + (syncing ? 'syncing' : 'ok');
    statusTxt.textContent = text || 'Sincronizado';
  } else {
    topRow.style.display    = 'flex';
    statusRow.style.display = 'none';
  }
}

/* ── Init Google Identity Services ── */
function initDrive() {
  // Wait for GSI script to load
  if (typeof google === 'undefined') {
    setTimeout(initDrive, 500);
    return;
  }

  // Restore token from session (not localStorage — tokens expire)
  const saved = sessionStorage.getItem(DRIVE_TOKEN_KEY);
  if (saved) {
    accessToken = saved;
    setDriveStatus(true, 'Sincronizado');
    // Pull from Drive on app load to catch any changes from other devices
    pullFromDrive();
  }

  tokenClient = google.accounts.oauth2.initTokenClient({
    client_id: GOOGLE_CLIENT_ID,
    scope: DRIVE_SCOPE,
    callback: onTokenReceived,
  });

  // Bind buttons
  document.getElementById('drive-connect-btn')?.addEventListener('click', () => {
    tokenClient.requestAccessToken({ prompt: 'consent' });
  });
  document.getElementById('drive-sync-btn')?.addEventListener('click', () => {
    pushToDrive();
  });
  document.getElementById('drive-disconnect-btn')?.addEventListener('click', () => {
    disconnectDrive();
  });
}

function onTokenReceived(response) {
  if (response.error) {
    showToast('No se pudo conectar con Google', 'error');
    return;
  }
  accessToken = response.access_token;
  sessionStorage.setItem(DRIVE_TOKEN_KEY, accessToken);
  setDriveStatus(true, 'Conectando…', true);
  showToast('Conectado a Google Drive', 'success');
  // First sync: pull → merge → push
  pullFromDrive().then(() => pushToDrive());
}

function disconnectDrive() {
  if (accessToken) {
    google.accounts.oauth2.revoke(accessToken, () => {});
  }
  accessToken = null;
  driveFileId = null;
  sessionStorage.removeItem(DRIVE_TOKEN_KEY);
  localStorage.removeItem(DRIVE_FILE_KEY);
  setDriveStatus(false);
  showToast('Desconectado de Google Drive');
}

/* ══════════════════════════════════════
   PUSH — local → Drive
   ══════════════════════════════════════ */
async function pushToDrive() {
  if (!accessToken) return;
  setDriveStatus(true, 'Guardando en Drive…', true);

  const payload = {
    version: 2,
    updatedAt: new Date().toISOString(),
    txs: window._appGetTxs ? window._appGetTxs() : [],
    nextId: window._appGetNextId ? window._appGetNextId() : 1,
  };
  const body = JSON.stringify(payload);

  try {
    if (driveFileId) {
      await updateDriveFile(driveFileId, body);
    } else {
      driveFileId = await createDriveFile(body);
      localStorage.setItem(DRIVE_FILE_KEY, driveFileId);
    }
    setDriveStatus(true, 'Sincronizado · ' + timeNow());
  } catch (e) {
    if (e.status === 401) {
      // Token expired → request new one silently
      tokenClient.requestAccessToken({ prompt: '' });
    } else {
      setDriveStatus(true, 'Error al sincronizar');
      showToast('Error al guardar en Drive', 'error');
    }
  }
}

async function createDriveFile(content) {
  // Multipart upload: metadata + content
  const boundary = 'finanzas_boundary_x9k';
  const meta = JSON.stringify({ name: DRIVE_FILE_NAME, mimeType: 'application/json' });
  const multipart = [
    `--${boundary}`,
    'Content-Type: application/json; charset=UTF-8',
    '',
    meta,
    `--${boundary}`,
    'Content-Type: application/json',
    '',
    content,
    `--${boundary}--`,
  ].join('\r\n');

  const res = await apiFetch(
    'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart',
    {
      method: 'POST',
      headers: { 'Content-Type': `multipart/related; boundary=${boundary}` },
      body: multipart,
    }
  );
  const data = await res.json();
  return data.id;
}

async function updateDriveFile(fileId, content) {
  await apiFetch(
    `https://www.googleapis.com/upload/drive/v3/files/${fileId}?uploadType=media`,
    {
      method: 'PATCH',
      headers: { 'Content-Type': 'application/json' },
      body: content,
    }
  );
}

/* ══════════════════════════════════════
   PULL — Drive → local (merge)
   ══════════════════════════════════════ */
async function pullFromDrive() {
  if (!accessToken) return;
  setDriveStatus(true, 'Bajando de Drive…', true);

  try {
    // Find the file if we don't have its ID
    if (!driveFileId) {
      const searchRes = await apiFetch(
        `https://www.googleapis.com/drive/v3/files?q=name='${DRIVE_FILE_NAME}'+and+trashed=false&fields=files(id,modifiedTime)&orderBy=modifiedTime+desc`
      );
      const searchData = await searchRes.json();
      if (searchData.files && searchData.files.length > 0) {
        driveFileId = searchData.files[0].id;
        localStorage.setItem(DRIVE_FILE_KEY, driveFileId);
      }
    }

    if (!driveFileId) {
      setDriveStatus(true, 'Sin archivo en Drive aún');
      return;
    }

    const fileRes = await apiFetch(
      `https://www.googleapis.com/drive/v3/files/${driveFileId}?alt=media`
    );
    const remote = await fileRes.json();

    if (remote && remote.txs) {
      mergeWithLocal(remote);
      setDriveStatus(true, 'Sincronizado · ' + timeNow());
    }
  } catch (e) {
    setDriveStatus(true, 'Sin conexión — datos locales');
  }
}

/* ══════════════════════════════════════
   MERGE — une local + Drive por ID
   El movimiento con ID mayor gana en
   caso de conflicto de monto/concepto.
   Los de Drive que no están en local
   se agregan, y viceversa.
   ══════════════════════════════════════ */
function mergeWithLocal(remote) {
  if (!window._appGetTxs || !window._appSetTxs) return;

  const local    = window._appGetTxs();
  const localIds = new Set(local.map(t => t.id));
  const remoteIds = new Set(remote.txs.map(t => t.id));

  // Movimientos que están en Drive pero no en local → agregar
  const toAdd = remote.txs.filter(t => !localIds.has(t.id));

  // Movimientos que están en local pero no en Drive → quedan (ya están)
  // (no borramos nada en el merge)

  if (toAdd.length > 0) {
    const merged = [...local, ...toAdd].sort((a, b) => a.id - b.id);
    const newNextId = Math.max(
      window._appGetNextId(),
      remote.nextId || 1,
      ...merged.map(t => t.id + 1)
    );
    window._appSetTxs(merged, newNextId);
    showToast(`+${toAdd.length} movimiento${toAdd.length > 1 ? 's' : ''} sincronizado${toAdd.length > 1 ? 's' : ''} de Drive`, 'success');

    // Re-render current screen
    if (window._appRefresh) window._appRefresh();
  } else {
    showToast('Todo al día', 'info');
  }
}

/* ══════════════════════════════════════
   HTTP helper with auth header
   ══════════════════════════════════════ */
async function apiFetch(url, options = {}) {
  const res = await fetch(url, {
    ...options,
    headers: {
      Authorization: `Bearer ${accessToken}`,
      ...(options.headers || {}),
    },
  });
  if (!res.ok) {
    const err = new Error(`Drive API error ${res.status}`);
    err.status = res.status;
    throw err;
  }
  return res;
}

function timeNow() {
  return new Date().toLocaleTimeString('es-AR', { hour: '2-digit', minute: '2-digit' });
}

/* ══════════════════════════════════════
   PUBLIC API — llamado desde app.js
   ══════════════════════════════════════ */
window._driveGetToken = () => accessToken;

window.driveSync = {
  push: pushToDrive,
  pull: pullFromDrive,
  isConnected: () => !!accessToken,
};

/* Boot */
document.addEventListener('DOMContentLoaded', initDrive);
