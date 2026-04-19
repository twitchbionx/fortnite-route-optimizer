const { contextBridge, ipcRenderer } = require("electron");

contextBridge.exposeInMainWorld("api", {
  // ── Network / Ping ──
  pingServer: (serverId) => ipcRenderer.invoke("ping-server", serverId),
  pingAll: () => ipcRenderer.invoke("ping-all"),
  traceroute: (serverId) => ipcRenderer.invoke("traceroute", serverId),
  applyOptimization: (optId) => ipcRenderer.invoke("apply-optimization", optId),
  getNetworkInfo: () => ipcRenderer.invoke("get-network-info"),
  // ── Window ──
  windowMinimize: () => ipcRenderer.invoke("window-minimize"),
  windowMaximize: () => ipcRenderer.invoke("window-maximize"),
  windowClose: () => ipcRenderer.invoke("window-close"),
  // ── Tunnel ──
  tunnelStatus: () => ipcRenderer.invoke("tunnel-status"),
  tunnelSaveConfig: (config) => ipcRenderer.invoke("tunnel-save-config", config),
  tunnelLoadConfig: () => ipcRenderer.invoke("tunnel-load-config"),
  tunnelConnect: () => ipcRenderer.invoke("tunnel-connect"),
  tunnelDisconnect: () => ipcRenderer.invoke("tunnel-disconnect"),
  tunnelToggle: () => ipcRenderer.invoke("tunnel-toggle"),
  tunnelStats: () => ipcRenderer.invoke("tunnel-stats"),
  // ── Benchmark ──
  benchmarkRun: (serverId) => ipcRenderer.invoke("benchmark-run", serverId),
  // ── Updater ──
  updaterCheck: () => ipcRenderer.invoke("updater-check"),
  updaterDownload: () => ipcRenderer.invoke("updater-download"),
  updaterInstall: () => ipcRenderer.invoke("updater-install"),
  updaterStatus: () => ipcRenderer.invoke("updater-status"),
  updaterDismiss: () => ipcRenderer.invoke("updater-dismiss"),
  onUpdateStatus: (callback) => ipcRenderer.on("update-status", (_e, data) => callback(data)),
  // ── Debloater ──
  debloatGetLists: () => ipcRenderer.invoke("debloat-get-lists"),
  debloatScanApps: () => ipcRenderer.invoke("debloat-scan-apps"),
  debloatRemoveApp: (appId) => ipcRenderer.invoke("debloat-remove-app", appId),
  debloatRemoveAllApps: () => ipcRenderer.invoke("debloat-remove-all-apps"),
  debloatScanServices: () => ipcRenderer.invoke("debloat-scan-services"),
  debloatDisableService: (svcId) => ipcRenderer.invoke("debloat-disable-service", svcId),
  debloatDisableAllServices: () => ipcRenderer.invoke("debloat-disable-all-services"),
  debloatApplyTelemetry: (id) => ipcRenderer.invoke("debloat-apply-telemetry", id),
  debloatApplyAllTelemetry: () => ipcRenderer.invoke("debloat-apply-all-telemetry"),
  debloatApplyPerf: (id) => ipcRenderer.invoke("debloat-apply-perf", id),
  debloatApplyAllPerf: () => ipcRenderer.invoke("debloat-apply-all-perf"),
  debloatFullNuke: () => ipcRenderer.invoke("debloat-full-nuke"),
  // ── Gaming Tweaks ──
  tweaksGetCategories: () => ipcRenderer.invoke("tweaks-get-categories"),
  tweaksApply: (tweakId) => ipcRenderer.invoke("tweaks-apply", tweakId),
  tweaksApplyCategory: (catId) => ipcRenderer.invoke("tweaks-apply-category", catId),
  tweaksApplyAll: () => ipcRenderer.invoke("tweaks-apply-all"),
  tweaksGetApplied: () => ipcRenderer.invoke("tweaks-get-applied"),
  // ── Hardware / Overclock ──
  hwDetect: () => ipcRenderer.invoke("hw-detect"),
  hwStats: () => ipcRenderer.invoke("hw-stats"),
  hwGpuPresets: () => ipcRenderer.invoke("hw-gpu-presets"),
  hwGpuApplyPreset: (id) => ipcRenderer.invoke("hw-gpu-apply-preset", id),
  hwGpuReset: () => ipcRenderer.invoke("hw-gpu-reset"),
  hwCpuPresets: () => ipcRenderer.invoke("hw-cpu-presets"),
  hwCpuApplyPreset: (id) => ipcRenderer.invoke("hw-cpu-apply-preset", id),
  hwRamTweaks: () => ipcRenderer.invoke("hw-ram-tweaks"),
  hwRamApply: (id) => ipcRenderer.invoke("hw-ram-apply", id),
  hwStressTest: (dur) => ipcRenderer.invoke("hw-stress-test", dur),
  // ── BIOS ──
  biosInfo: () => ipcRenderer.invoke("bios-info"),
  biosHidden: () => ipcRenderer.invoke("bios-hidden"),
  biosGuide: () => ipcRenderer.invoke("bios-guide"),
  biosReboot: () => ipcRenderer.invoke("bios-reboot"),
  biosApplyBcd: (setting, value) => ipcRenderer.invoke("bios-apply-bcd", setting, value),
  biosDisableHyperV: () => ipcRenderer.invoke("bios-disable-hyperv"),
});
