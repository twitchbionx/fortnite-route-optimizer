const { app, BrowserWindow, ipcMain } = require("electron");
const path = require("path");
const { exec } = require("child_process");
const net = require("net");
const dns = require("dns");
const os = require("os");
const TunnelManager = require("./tunnel.js");
const { initUpdater } = require("./updater.js");
const Debloater = require("./debloater.js");
const GamingTweaks = require("./tweaks.js");
const HardwareMonitor = require("./overclock.js");
const BIOSHelper = require("./bios.js");
const DependencyManager = require("./installer.js");

// ── Module instances ──────────────────────────────────────────────────
const tunnel = new TunnelManager();
const debloater = new Debloater();
const gamingTweaks = new GamingTweaks();
const hardware = new HardwareMonitor();
const biosHelper = new BIOSHelper();
const installer = new DependencyManager();

// ── Fortnite Server Endpoints ─────────────────────────────────────────
// These are real AWS region endpoints Epic uses for Fortnite matchmaking
const FORTNITE_SERVERS = {
  nae:  { host: "dynamodb.us-east-1.amazonaws.com",       name: "NA-East",      city: "Virginia"  },
  naw:  { host: "dynamodb.us-west-2.amazonaws.com",       name: "NA-West",      city: "Oregon"    },
  nac:  { host: "dynamodb.us-east-2.amazonaws.com",       name: "NA-Central",   city: "Ohio"      },
  dal:  { host: "tx-us-ping.vultr.com",                    name: "NA-South",     city: "Dallas"    },
  eu1:  { host: "dynamodb.eu-west-2.amazonaws.com",       name: "EU-West",      city: "London"    },
  eu2:  { host: "dynamodb.eu-central-1.amazonaws.com",    name: "EU-Central",   city: "Frankfurt" },
  oce:  { host: "dynamodb.ap-southeast-2.amazonaws.com",  name: "Oceania",      city: "Sydney"    },
  br:   { host: "dynamodb.sa-east-1.amazonaws.com",       name: "Brazil",       city: "São Paulo" },
  asia: { host: "dynamodb.ap-northeast-1.amazonaws.com",  name: "Asia",         city: "Tokyo"     },
  me:   { host: "dynamodb.me-south-1.amazonaws.com",      name: "Middle East",  city: "Bahrain"   },
  ind:  { host: "dynamodb.ap-south-1.amazonaws.com",      name: "India",        city: "Mumbai"    },
};

// ── Real TCP Ping ─────────────────────────────────────────────────────
// Measures actual round-trip time by opening a TCP connection to port 443
function tcpPing(host, port = 443, timeout = 5000) {
  return new Promise((resolve) => {
    const start = process.hrtime.bigint();
    const socket = new net.Socket();

    socket.setTimeout(timeout);

    socket.on("connect", () => {
      const end = process.hrtime.bigint();
      const ms = Number(end - start) / 1_000_000;
      socket.destroy();
      resolve({ success: true, ping: Math.round(ms * 10) / 10 });
    });

    socket.on("timeout", () => {
      socket.destroy();
      resolve({ success: false, ping: -1 });
    });

    socket.on("error", () => {
      socket.destroy();
      resolve({ success: false, ping: -1 });
    });

    socket.connect(port, host);
  });
}

// Run multiple pings and return stats
async function multiPing(host, count = 5) {
  const results = [];
  for (let i = 0; i < count; i++) {
    const result = await tcpPing(host);
    if (result.success) results.push(result.ping);
    // Small delay between pings
    await new Promise((r) => setTimeout(r, 100));
  }

  if (results.length === 0) {
    return { ping: -1, jitter: 0, loss: 100, min: 0, max: 0, samples: results };
  }

  const avg = results.reduce((a, b) => a + b, 0) / results.length;
  const min = Math.min(...results);
  const max = Math.max(...results);
  // Jitter = average deviation from mean
  const jitter = results.reduce((sum, p) => sum + Math.abs(p - avg), 0) / results.length;
  const loss = ((count - results.length) / count) * 100;

  return {
    ping: Math.round(avg),
    jitter: Math.round(jitter * 10) / 10,
    loss: Math.round(loss * 10) / 10,
    min: Math.round(min),
    max: Math.round(max),
    samples: results.map((r) => Math.round(r)),
  };
}

// ── Traceroute (Windows) ──────────────────────────────────────────────
function runTraceroute(host) {
  return new Promise((resolve) => {
    const isWin = os.platform() === "win32";
    const cmd = isWin ? `tracert -d -h 20 -w 2000 ${host}` : `traceroute -n -m 20 -w 2 ${host}`;

    exec(cmd, { timeout: 30000 }, (err, stdout) => {
      if (err && !stdout) {
        resolve([]);
        return;
      }

      const lines = stdout.split("\n").filter((l) => l.trim());
      const hops = [];

      for (const line of lines) {
        // Windows: "  1    <1 ms    <1 ms    <1 ms  192.168.1.1"
        // Linux:   " 1  192.168.1.1  0.543 ms  0.345 ms  0.298 ms"
        let match;
        if (isWin) {
          match = line.match(/^\s*(\d+)\s+(.+?)\s+([\d.]+(?:\.\d+)*)\s*$/);
          if (!match) {
            // Try alternate Windows format
            match = line.match(/^\s*(\d+)\s+(?:(<?\d+)\s*ms\s+(?:<?\d+)\s*ms\s+(?:<?\d+)\s*ms)\s+([\d.]+)/);
          }
          if (match) {
            const hopNum = parseInt(match[1]);
            const ip = match[3] || match[match.length - 1];
            // Extract first ms value
            const msMatch = line.match(/(\d+)\s*ms/);
            const ping = msMatch ? parseInt(msMatch[1]) : 0;
            hops.push({ hop: hopNum, host: ip.trim(), ping: Math.max(1, ping), loss: 0 });
          }
        } else {
          match = line.match(/^\s*(\d+)\s+([\d.]+)\s+([\d.]+)\s*ms/);
          if (match) {
            hops.push({ hop: parseInt(match[1]), host: match[2], ping: Math.round(parseFloat(match[3])), loss: 0 });
          }
        }
      }
      resolve(hops);
    });
  });
}

// ── DNS Resolution helper ─────────────────────────────────────────────
function resolveHost(hostname) {
  return new Promise((resolve) => {
    dns.resolve4(hostname, (err, addresses) => {
      if (err) resolve(hostname);
      else resolve(addresses[0]);
    });
  });
}

// ── Network Optimization Commands (Windows) ───────────────────────────
function runOptimization(id) {
  const isWin = os.platform() === "win32";
  if (!isWin) return Promise.resolve({ success: false, msg: "Windows only" });

  const commands = {
    dns: 'ipconfig /flushdns && netsh interface ip set dns "Ethernet" static 1.1.1.1 && netsh interface ip add dns "Ethernet" 1.0.0.1 index=2',
    nagle: 'powershell -Command "Get-NetAdapter | ForEach-Object { Set-NetTCPSetting -SettingName InternetCustom -AutoTuningLevelLocal Normal }"',
    throttle: 'reg add "HKLM\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Multimedia\\SystemProfile" /v NetworkThrottlingIndex /t REG_DWORD /d 4294967295 /f',
    qos: 'reg add "HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\Psched" /v NonBestEffortLimit /t REG_DWORD /d 0 /f',
    mtu: 'netsh interface ipv4 set subinterface "Ethernet" mtu=1472 store=persistent',
    tcp: "netsh int tcp set global autotuninglevel=experimental",
    rss: "netsh int tcp set global rss=enabled",
    ecn: "netsh int tcp set global ecncapability=enabled",
  };

  const cmd = commands[id];
  if (!cmd) return Promise.resolve({ success: false, msg: "Unknown optimization" });

  return new Promise((resolve) => {
    exec(cmd, { timeout: 10000 }, (err, stdout, stderr) => {
      resolve({
        success: !err,
        msg: err ? stderr || err.message : stdout || "Applied successfully",
      });
    });
  });
}

// ── Electron Window ───────────────────────────────────────────────────
let mainWindow;

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1400,
    height: 900,
    minWidth: 1100,
    minHeight: 700,
    title: "Fortnite Route Optimizer",
    backgroundColor: "#06060e",
    autoHideMenuBar: true,
    frame: false,
    titleBarStyle: "hidden",
    webPreferences: {
      preload: path.join(__dirname, "preload.js"),
      contextIsolation: true,
      nodeIntegration: false,
    },
  });

  mainWindow.loadFile("index.html");

  // Init auto-updater after window is ready
  mainWindow.webContents.on("did-finish-load", () => {
    initUpdater(mainWindow);
  });
}

app.whenReady().then(async () => {
  createWindow();

  // Run boot check — detect crashes and restore last known good settings if needed
  try {
    const bootResult = await hardware.bootCheck();
    if (bootResult.restored) {
      console.log("[BootCheck] Crash detected — restored last known good settings.", bootResult.crashes);
    }
  } catch (e) {
    console.error("[BootCheck] Error during boot check:", e.message);
  }
});

app.on("window-all-closed", () => {
  if (process.platform !== "darwin") app.quit();
});

app.on("activate", () => {
  if (BrowserWindow.getAllWindows().length === 0) createWindow();
});

// ── IPC Handlers ──────────────────────────────────────────────────────

// Ping a specific server
ipcMain.handle("ping-server", async (_event, serverId) => {
  const server = FORTNITE_SERVERS[serverId];
  if (!server) return { error: "Unknown server" };

  const ip = await resolveHost(server.host);
  const stats = await multiPing(ip, 5);
  return { ...stats, ip, serverId };
});

// Ping all servers
ipcMain.handle("ping-all", async () => {
  const results = {};
  for (const [id, server] of Object.entries(FORTNITE_SERVERS)) {
    const ip = await resolveHost(server.host);
    const stats = await multiPing(ip, 4);
    results[id] = { ...stats, ip };
  }
  return results;
});

// Traceroute
ipcMain.handle("traceroute", async (_event, serverId) => {
  const server = FORTNITE_SERVERS[serverId];
  if (!server) return [];
  const ip = await resolveHost(server.host);
  return await runTraceroute(ip);
});

// Apply optimization
ipcMain.handle("apply-optimization", async (_event, optId) => {
  return await runOptimization(optId);
});

// Window controls
ipcMain.handle("window-minimize", () => mainWindow.minimize());
ipcMain.handle("window-maximize", () => {
  if (mainWindow.isMaximized()) mainWindow.unmaximize();
  else mainWindow.maximize();
});
ipcMain.handle("window-close", () => mainWindow.close());

// Get system network info
ipcMain.handle("get-network-info", () => {
  const interfaces = os.networkInterfaces();
  const info = [];
  for (const [name, addrs] of Object.entries(interfaces)) {
    for (const addr of addrs) {
      if (addr.family === "IPv4" && !addr.internal) {
        info.push({ name, address: addr.address, mac: addr.mac });
      }
    }
  }
  return info;
});

// ── Tunnel IPC Handlers ───────────────────────────────────────────────

// Get tunnel status
ipcMain.handle("tunnel-status", () => {
  return tunnel.getStatus();
});

// Save tunnel config
ipcMain.handle("tunnel-save-config", (_event, configText) => {
  return tunnel.saveConfig(configText);
});

// Load saved config
ipcMain.handle("tunnel-load-config", () => {
  return tunnel.loadConfig();
});

// Connect tunnel
ipcMain.handle("tunnel-connect", async () => {
  return await tunnel.connect();
});

// Disconnect tunnel
ipcMain.handle("tunnel-disconnect", async () => {
  return await tunnel.disconnect();
});

// Toggle tunnel
ipcMain.handle("tunnel-toggle", async () => {
  return await tunnel.toggle();
});

// Get tunnel stats
ipcMain.handle("tunnel-stats", async () => {
  return await tunnel.getStats();
});

// Retry WireGuard detection
ipcMain.handle("tunnel-retry-detect", () => {
  return tunnel.retryDetect();
});

// ── Benchmark IPC Handler ─────────────────────────────────────────
// Runs 10 pings to a server and returns detailed stats
// Used for before/after tunnel comparison
ipcMain.handle("benchmark-run", async (_event, serverId) => {
  const server = FORTNITE_SERVERS[serverId];
  if (!server) return { error: "Unknown server" };

  const ip = await resolveHost(server.host);
  const stats = await multiPing(ip, 10);

  // Also collect individual samples with timestamps for charting
  const samples = [];
  for (let i = 0; i < 10; i++) {
    const result = await tcpPing(ip);
    samples.push({
      index: i,
      ping: result.success ? Math.round(result.ping * 10) / 10 : -1,
      success: result.success,
    });
    await new Promise((r) => setTimeout(r, 150));
  }

  return {
    ...stats,
    ip,
    serverId,
    detailedSamples: samples,
    timestamp: Date.now(),
  };
});

// ── Debloater IPC Handlers ────────────────────────────────────────
ipcMain.handle("debloat-get-lists", () => debloater.getLists());
ipcMain.handle("debloat-scan-apps", async () => await debloater.scanBloatware());
ipcMain.handle("debloat-remove-app", async (_e, appId) => await debloater.removeBloatware(appId));
ipcMain.handle("debloat-remove-all-apps", async () => await debloater.removeAllBloatware());
ipcMain.handle("debloat-scan-services", async () => await debloater.scanServices());
ipcMain.handle("debloat-disable-service", async (_e, svcId) => await debloater.disableService(svcId));
ipcMain.handle("debloat-disable-all-services", async () => await debloater.disableAllServices());
ipcMain.handle("debloat-apply-telemetry", async (_e, id) => await debloater.applyTelemetryTweak(id));
ipcMain.handle("debloat-apply-all-telemetry", async () => await debloater.applyAllTelemetry());
ipcMain.handle("debloat-apply-perf", async (_e, id) => await debloater.applyPerfTweak(id));
ipcMain.handle("debloat-apply-all-perf", async () => await debloater.applyAllPerf());
ipcMain.handle("debloat-full-nuke", async () => await debloater.fullDebloat());
ipcMain.handle("debloat-scan-all-apps", async () => await debloater.scanAllApps());

// ── Gaming Tweaks IPC Handlers ───────────────────────────────────
ipcMain.handle("tweaks-get-categories", () => gamingTweaks.getCategories());
ipcMain.handle("tweaks-apply", async (_e, tweakId) => await gamingTweaks.applyTweak(tweakId));
ipcMain.handle("tweaks-apply-category", async (_e, catId) => await gamingTweaks.applyCategory(catId));
ipcMain.handle("tweaks-apply-all", async () => await gamingTweaks.applyAll());
ipcMain.handle("tweaks-get-applied", () => gamingTweaks.getApplied());

// ── Hardware / Overclock IPC Handlers ────────────────────────────
ipcMain.handle("hw-detect", async () => await hardware.detectHardware());
ipcMain.handle("hw-stats", async () => await hardware.getStats());
ipcMain.handle("hw-gpu-presets", () => hardware.getGPUPresets());
ipcMain.handle("hw-gpu-apply-preset", async (_e, id) => await hardware.applyGPUPreset(id));
ipcMain.handle("hw-gpu-reset", async () => await hardware.gpuResetClocks());
ipcMain.handle("hw-cpu-presets", () => hardware.getCPUPresets());
ipcMain.handle("hw-cpu-apply-preset", async (_e, id) => await hardware.applyCPUPreset(id));
ipcMain.handle("hw-ram-tweaks", () => hardware.getRAMTweaks());
ipcMain.handle("hw-ram-apply", async (_e, id) => await hardware.applyRAMTweak(id));
ipcMain.handle("hw-stress-test", async (_e, dur) => await hardware.runStressTest(dur || 10));
ipcMain.handle("hw-auto-tune", async () => await hardware.autoTuneGPU());
ipcMain.handle("hw-boot-check", async () => await hardware.bootCheck());
ipcMain.handle("hw-crash-log", async () => await hardware.getCrashLog());
ipcMain.handle("hw-restore-last-good", async () => await hardware.restoreLastKnownGood());

// ── SceWin / AI Overclock IPC Handlers ──────────────────────────
ipcMain.handle("hw-scewin-status", () => hardware.getScewinStatus());
ipcMain.handle("hw-scewin-export", async () => await hardware.scewin.exportCurrentSettings());
ipcMain.handle("hw-scewin-backup", async () => await hardware.scewin.backupSettings());
ipcMain.handle("hw-scewin-restore", async () => await hardware.scewin.restoreBackup());
ipcMain.handle("hw-ai-auto-oc", async (_, opts) => await hardware.aiAutoOC(opts));
ipcMain.handle("hw-ai-progress", () => hardware.getAIProgress());
ipcMain.handle("hw-ai-stop", () => hardware.stopAIOC());

// ── BIOS IPC Handlers ────────────────────────────────────────────
ipcMain.handle("bios-info", async () => await biosHelper.getBIOSInfo());
ipcMain.handle("bios-hidden", async () => await biosHelper.readHiddenSettings());
ipcMain.handle("bios-guide", async () => await biosHelper.getGuide());
ipcMain.handle("bios-reboot", async () => await biosHelper.rebootToBIOS());
ipcMain.handle("bios-apply-bcd", async (_e, setting, value) => await biosHelper.applyBCDSetting(setting, value));
ipcMain.handle("bios-disable-hyperv", async () => await biosHelper.disableHyperV());

// ── Installer / Dependency IPC Handlers ─────────────────────────
ipcMain.handle("deps-check-all", async () => await installer.checkAll());
ipcMain.handle("deps-install", async (_e, depId) => await installer.installDep(depId));
ipcMain.handle("deps-install-all", async () => await installer.installAll());

// Load config on startup
tunnel.loadConfig();
