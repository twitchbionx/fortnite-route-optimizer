// ═══════════════════════════════════════════════════════════════════
//  AUTO-UPDATER MODULE
//  Checks GitHub Releases for new versions, downloads in background,
//  notifies the renderer, and installs on quit.
// ═══════════════════════════════════════════════════════════════════

const { autoUpdater } = require("electron-updater");
const { ipcMain } = require("electron");

let mainWindow = null;

// Disable auto-download — we want to show the user first
autoUpdater.autoDownload = false;
autoUpdater.autoInstallOnAppQuit = true;

// Logging
autoUpdater.logger = require("electron").app
  ? { info: console.log, warn: console.warn, error: console.error }
  : console;

// ── State ────────────────────────────────────────────────────────────
let updateState = {
  status: "idle", // idle | checking | available | downloading | downloaded | error | up-to-date
  currentVersion: null,
  newVersion: null,
  releaseNotes: "",
  downloadProgress: 0,
  error: null,
};

function sendToRenderer(event, data) {
  if (mainWindow && !mainWindow.isDestroyed()) {
    mainWindow.webContents.send(event, data);
  }
}

// ── Auto-Updater Events ──────────────────────────────────────────────

autoUpdater.on("checking-for-update", () => {
  updateState.status = "checking";
  updateState.error = null;
  sendToRenderer("update-status", updateState);
});

autoUpdater.on("update-available", (info) => {
  updateState.status = "available";
  updateState.newVersion = info.version;
  updateState.releaseNotes = info.releaseNotes || "";
  sendToRenderer("update-status", updateState);
});

autoUpdater.on("update-not-available", () => {
  updateState.status = "up-to-date";
  sendToRenderer("update-status", updateState);
});

autoUpdater.on("download-progress", (progress) => {
  updateState.status = "downloading";
  updateState.downloadProgress = Math.round(progress.percent);
  sendToRenderer("update-status", updateState);
});

autoUpdater.on("update-downloaded", (info) => {
  updateState.status = "downloaded";
  updateState.newVersion = info.version;
  sendToRenderer("update-status", updateState);
});

autoUpdater.on("error", (err) => {
  updateState.status = "error";
  updateState.error = err.message || "Update check failed";
  sendToRenderer("update-status", updateState);
});

// ── IPC Handlers ─────────────────────────────────────────────────────

// Check for updates
ipcMain.handle("updater-check", async () => {
  try {
    await autoUpdater.checkForUpdates();
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
});

// Download the update
ipcMain.handle("updater-download", async () => {
  try {
    await autoUpdater.downloadUpdate();
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
});

// Install and restart
ipcMain.handle("updater-install", () => {
  autoUpdater.quitAndInstall(false, true);
});

// Get current state
ipcMain.handle("updater-status", () => {
  return updateState;
});

// Dismiss update notification
ipcMain.handle("updater-dismiss", () => {
  updateState.status = "idle";
  sendToRenderer("update-status", updateState);
  return { success: true };
});

// ── Init ─────────────────────────────────────────────────────────────
function initUpdater(win) {
  mainWindow = win;
  updateState.currentVersion = require("./package.json").version;

  // Check for updates 5 seconds after launch (don't slow down startup)
  setTimeout(() => {
    autoUpdater.checkForUpdates().catch(() => {});
  }, 5000);

  // Then check every 30 minutes
  setInterval(() => {
    autoUpdater.checkForUpdates().catch(() => {});
  }, 30 * 60 * 1000);
}

module.exports = { initUpdater };
