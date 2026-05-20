/**
 * Local File System Access API helpers
 * Uses the browser's native File System Access API (Chrome/Edge) to read and write
 * files in a user-selected local folder — works perfectly with OneDrive / SharePoint
 * sync folders without any Azure App registration.
 *
 * The selected folder handle is stored in IndexedDB so the user only has to pick it once
 * per browser session (permission re-granted automatically on subsequent visits).
 */

const DB_NAME  = "tsm_localfs";
const DB_STORE = "handles";
const DB_KEY   = "rootHandle";

// ─── IndexedDB: persist the directory handle ─────────────────────────────────
function openIDB() {
  return new Promise(function(resolve, reject) {
    var req = indexedDB.open(DB_NAME, 1);
    req.onupgradeneeded = function(e) { e.target.result.createObjectStore(DB_STORE); };
    req.onsuccess = function(e) { resolve(e.target.result); };
    req.onerror   = function(e) { reject(e.target.error); };
  });
}

export async function saveHandle(handle) {
  var db = await openIDB();
  return new Promise(function(resolve, reject) {
    var tx = db.transaction(DB_STORE, "readwrite");
    tx.objectStore(DB_STORE).put(handle, DB_KEY);
    tx.oncomplete = resolve;
    tx.onerror    = function(e) { reject(e.target.error); };
  });
}

export async function loadHandle() {
  var db = await openIDB();
  return new Promise(function(resolve, reject) {
    var tx  = db.transaction(DB_STORE, "readonly");
    var req = tx.objectStore(DB_STORE).get(DB_KEY);
    req.onsuccess = function(e) { resolve(e.target.result || null); };
    req.onerror   = function(e) { reject(e.target.error); };
  });
}

export async function clearHandle() {
  var db = await openIDB();
  return new Promise(function(resolve, reject) {
    var tx = db.transaction(DB_STORE, "readwrite");
    tx.objectStore(DB_STORE).delete(DB_KEY);
    tx.oncomplete = resolve;
    tx.onerror    = function(e) { reject(e.target.error); };
  });
}

// ─── Capability check ────────────────────────────────────────────────────────
export function isSupported() {
  return typeof window !== "undefined" && typeof window.showDirectoryPicker === "function";
}

// ─── Pick a folder (opens OS dialog) ─────────────────────────────────────────
export async function pickFolder() {
  var handle = await window.showDirectoryPicker({ mode: "readwrite" });
  await saveHandle(handle);
  return handle;
}

// ─── Restore saved handle + verify permission ────────────────────────────────
export async function restoreFolder() {
  var handle = await loadHandle();
  if (!handle) return null;
  // Check if permission is still granted (doesn't prompt the user)
  var perm = await handle.queryPermission({ mode: "readwrite" });
  if (perm === "granted") return handle;
  // Try to re-request permission (requires a user gesture, so we just return null
  // and let the UI prompt the user to click "Reconnect")
  return null;
}

// Re-request permission (call from a click handler — requires user gesture)
export async function requestPermission(handle) {
  var perm = await handle.requestPermission({ mode: "readwrite" });
  return perm === "granted";
}

// ─── List files in a directory ───────────────────────────────────────────────
export async function listFiles(dirHandle, filterFn) {
  var items = [];
  for await (var [name, handle] of dirHandle.entries()) {
    if (handle.kind === "file") {
      if (!filterFn || filterFn(name)) items.push({ name, handle, kind: "file" });
    } else {
      items.push({ name, handle, kind: "directory" });
    }
  }
  return items.sort(function(a, b) {
    // Directories first, then alphabetical
    if (a.kind !== b.kind) return a.kind === "directory" ? -1 : 1;
    return a.name.localeCompare(b.name);
  });
}

// ─── Read a file as ArrayBuffer ───────────────────────────────────────────────
export async function readFile(fileHandle) {
  var file = await fileHandle.getFile();
  return file.arrayBuffer();
}

// ─── Write / overwrite a file ─────────────────────────────────────────────────
export async function writeFile(dirHandle, filename, content) {
  var fh = await dirHandle.getFileHandle(filename, { create: true });
  var writable = await fh.createWritable();
  await writable.write(content);
  await writable.close();
}

// ─── Database file operations ────────────────────────────────────────────────
export const DB_FILENAME = "TimesheetManager_DB.json";

export async function loadDatabase(dirHandle) {
  try {
    var fh = await dirHandle.getFileHandle(DB_FILENAME);
    var file = await fh.getFile();
    var text = await file.text();
    return JSON.parse(text);
  } catch (e) {
    if (e.name === "NotFoundError") return null; // file doesn't exist yet
    throw e;
  }
}

export async function saveDatabase(dirHandle, data) {
  var json = JSON.stringify(data, null, 2);
  await writeFile(dirHandle, DB_FILENAME, json);
}

// ─── Navigate into a sub-folder ──────────────────────────────────────────────
export async function getSubDir(dirHandle, name, create = false) {
  return dirHandle.getDirectoryHandle(name, { create });
}
