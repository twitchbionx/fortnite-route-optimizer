// ═══════════════════════════════════════════════════════════════════
//  AUTO-INSTALLER MODULE
//  Detects missing dependencies and downloads/installs them.
//  Dependencies: WireGuard, SceWin (AMISCE), NVIDIA drivers
// ═══════════════════════════════════════════════════════════════════

const { exec } = require("child_process");
const fs = require("fs");
const path = require("path");
const os = require("os");
const https = require("https");
const http = require("http");

const TOOLS_DIR = path.join(os.homedir(), ".fn-optimizer", "tools");
const DOWNLOADS_DIR = path.join(os.homedir(), ".fn-optimizer", "downloads");

function ensureDir(dir) {
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
}

function runCmd(cmd, timeout = 60000) {
  return new Promise((resolve) => {
    exec(cmd, { timeout }, (err, stdout, stderr) => {
      resolve({ success: !err, output: stdout?.trim() || "", error: stderr?.trim() || err?.message || "" });
    });
  });
}

function runPS(cmd, timeout = 60000) {
  return new Promise((resolve) => {
    exec(cmd, { timeout, shell: "powershell.exe" }, (err, stdout, stderr) => {
      resolve({ success: !err, output: stdout?.trim() || "", error: stderr?.trim() || err?.message || "" });
    });
  });
}

// ── Download a file with progress tracking (follows redirects) ──
function downloadFile(url, destPath, onProgress) {
  return new Promise((resolve) => {
    // Clean up any locked/stale file from a previous attempt
    try { fs.unlinkSync(destPath); } catch(e) {}
    // If still locked, use a unique temp name
    let actualPath = destPath;
    try {
      fs.writeFileSync(destPath, "");
      fs.unlinkSync(destPath);
    } catch(e) {
      actualPath = destPath.replace(/(\.\w+)$/, `-${Date.now()}$1`);
    }

    const follow = (targetUrl, redirects = 0) => {
      if (redirects > 10) return resolve({ success: false, error: "Too many redirects" });
      const mod = targetUrl.startsWith("https") ? https : http;
      mod.get(targetUrl, { headers: { "User-Agent": "FN-Optimizer" } }, (res) => {
        // Handle redirects
        if ([301, 302, 303, 307, 308].includes(res.statusCode) && res.headers.location) {
          let next = res.headers.location;
          if (next.startsWith("/")) {
            const u = new URL(targetUrl);
            next = u.protocol + "//" + u.host + next;
          }
          return follow(next, redirects + 1);
        }
        if (res.statusCode !== 200) {
          return resolve({ success: false, error: `HTTP ${res.statusCode}` });
        }

        const totalBytes = parseInt(res.headers["content-length"] || "0", 10);
        let downloadedBytes = 0;
        const file = fs.createWriteStream(actualPath);

        res.on("data", (chunk) => {
          downloadedBytes += chunk.length;
          if (onProgress && totalBytes > 0) {
            const pct = Math.round((downloadedBytes / totalBytes) * 100);
            const mbDown = (downloadedBytes / 1024 / 1024).toFixed(1);
            const mbTotal = (totalBytes / 1024 / 1024).toFixed(1);
            onProgress({ percent: pct, downloadedMB: mbDown, totalMB: mbTotal });
          } else if (onProgress) {
            const mbDown = (downloadedBytes / 1024 / 1024).toFixed(1);
            onProgress({ percent: -1, downloadedMB: mbDown, totalMB: "?" });
          }
        });

        res.pipe(file);
        file.on("finish", () => {
          file.close(() => {
            // Rename to intended path if we used a temp name
            if (actualPath !== destPath) {
              try { fs.renameSync(actualPath, destPath); } catch(e) {
                // If rename fails, just use the temp path
                return resolve({ success: true, path: actualPath });
              }
            }
            resolve({ success: true, path: destPath });
          });
        });
        file.on("error", (e) => {
          fs.unlink(actualPath, () => {});
          resolve({ success: false, error: e.message });
        });
      }).on("error", (e) => {
        resolve({ success: false, error: e.message });
      });
    };
    follow(url);
  });
}

// ── Run an installer elevated ──
function runInstallerElevated(installerPath, args = "", silent = true) {
  return new Promise((resolve) => {
    const silentArgs = silent ? "/S /SILENT /VERYSILENT /SUPPRESSMSGBOXES /NORESTART /NOCANCEL" : "";
    const fullArgs = `${args} ${silentArgs}`.trim();
    const psCmd = `Start-Process -FilePath '${installerPath.replace(/'/g, "''")}' -ArgumentList '${fullArgs}' -Verb RunAs -Wait -WindowStyle Hidden`;
    exec(psCmd, { timeout: 300000, shell: "powershell.exe" }, (err, stdout, stderr) => {
      resolve({ success: !err, output: stdout?.trim() || "", error: stderr?.trim() || err?.message || "" });
    });
  });
}

// ═══════════════════════════════════════════════════════════════════
//  DEPENDENCY DEFINITIONS
// ═══════════════════════════════════════════════════════════════════

class DependencyManager {
  constructor() {
    this.isWin = os.platform() === "win32";
    ensureDir(TOOLS_DIR);
    ensureDir(DOWNLOADS_DIR);
  }

  // ── Check all dependencies ──────────────────────────────────────
  async checkAll() {
    if (!this.isWin) return { error: "Windows only" };

    const [wireguard, nvidiaSmi, scewin] = await Promise.all([
      this.checkWireGuard(),
      this.checkNvidiaSmi(),
      this.checkSceWin(),
    ]);

    return {
      wireguard,
      nvidiaSmi,
      scewin,
      allInstalled: wireguard.installed && nvidiaSmi.installed,
      missingCount: [wireguard, nvidiaSmi, scewin].filter(d => !d.installed).length,
    };
  }

  // ── WireGuard ───────────────────────────────────────────────────
  async checkWireGuard() {
    const paths = [
      "C:\\Program Files\\WireGuard\\wireguard.exe",
      "C:\\Program Files (x86)\\WireGuard\\wireguard.exe",
      path.join(os.homedir(), "AppData\\Local\\Programs\\WireGuard\\wireguard.exe"),
      path.join(process.env.LOCALAPPDATA || "", "WireGuard\\wireguard.exe"),
      path.join(process.env.LOCALAPPDATA || "", "Programs\\WireGuard\\wireguard.exe"),
    ];

    for (const p of paths) {
      try { if (fs.existsSync(p)) return { installed: true, path: p, name: "WireGuard", id: "wireguard" }; } catch(e) {}
    }

    // Check PATH
    try {
      const r = await runCmd("where wireguard.exe 2>nul", 5000);
      if (r.success && r.output) return { installed: true, path: r.output.split("\n")[0].trim(), name: "WireGuard", id: "wireguard" };
    } catch(e) {}

    return {
      installed: false,
      name: "WireGuard",
      id: "wireguard",
      desc: "VPN tunnel for routing game traffic through your VPS",
      downloadUrl: "https://download.wireguard.com/windows-client/wireguard-installer.exe",
      size: "~8 MB",
      required: false,
      feature: "Tunnel",
    };
  }

  async installWireGuard(onProgress) {
    if (onProgress) onProgress({ status: "downloading", message: "Downloading WireGuard installer...", percent: 0 });
    const destPath = path.join(DOWNLOADS_DIR, "wireguard-installer.exe");
    const dl = await downloadFile("https://download.wireguard.com/windows-client/wireguard-installer.exe", destPath, (p) => {
      if (onProgress) onProgress({ status: "downloading", message: `Downloading WireGuard... ${p.downloadedMB}/${p.totalMB} MB`, percent: p.percent });
    });
    if (!dl.success) return { success: false, error: "Download failed: " + dl.error };
    const installerPath = dl.path || destPath;

    if (onProgress) onProgress({ status: "installing", message: "Installing WireGuard (admin required)..." });
    const result = await runInstallerElevated(installerPath, "", true);

    // Verify install
    const check = await this.checkWireGuard();
    return { success: check.installed, error: check.installed ? null : "Install may have failed — check manually" };
  }

  // ── NVIDIA SMI ─────────────────────────────────────────────────
  async checkNvidiaSmi() {
    const paths = [
      "C:\\Program Files\\NVIDIA Corporation\\NVSMI\\nvidia-smi.exe",
      "C:\\Windows\\System32\\nvidia-smi.exe",
    ];

    for (const p of paths) {
      try { if (fs.existsSync(p)) return { installed: true, path: p, name: "NVIDIA SMI", id: "nvidia-smi" }; } catch(e) {}
    }

    // Check PATH
    try {
      const r = await runCmd("where nvidia-smi.exe 2>nul", 5000);
      if (r.success && r.output) return { installed: true, path: r.output.split("\n")[0].trim(), name: "NVIDIA SMI", id: "nvidia-smi" };
    } catch(e) {}

    // Check if any NVIDIA GPU exists
    const gpuCheck = await runPS("Get-CimInstance Win32_VideoController | Where-Object { $_.Name -like '*NVIDIA*' } | Select-Object -First 1 Name | ConvertTo-Json", 10000);
    const hasNvidia = gpuCheck.success && gpuCheck.output && gpuCheck.output.includes("NVIDIA");

    return {
      installed: false,
      name: "NVIDIA SMI",
      id: "nvidia-smi",
      desc: hasNvidia ? "GPU monitoring tool — comes with NVIDIA drivers. Update your drivers to get it." : "GPU monitoring (NVIDIA GPU not detected)",
      downloadUrl: hasNvidia ? "https://www.nvidia.com/Download/index.aspx" : null,
      size: "~600 MB (full driver)",
      required: false,
      feature: "GPU Overclock & Monitoring",
      hasNvidiaGpu: hasNvidia,
      installNote: hasNvidia ? "nvidia-smi comes bundled with NVIDIA drivers. Download the latest GeForce driver from nvidia.com." : "No NVIDIA GPU detected — AMD/Intel GPUs use different tools.",
    };
  }

  async installNvidiaSmi(onProgress) {
    // nvidia-smi can't be installed standalone — it comes with the driver
    // We can try to update the driver using winget or guide the user
    if (onProgress) onProgress({ status: "checking", message: "Checking for NVIDIA driver updates..." });

    // Try winget first
    const wingetCheck = await runCmd("winget list NVIDIA.GeForceExperience 2>nul", 15000);
    if (wingetCheck.success) {
      if (onProgress) onProgress({ status: "installing", message: "Updating NVIDIA drivers via winget..." });
      const update = await runCmd("winget upgrade NVIDIA.GeForceExperience --silent --accept-package-agreements --accept-source-agreements 2>nul", 300000);
      if (update.success) {
        // Check if smi appeared
        const check = await this.checkNvidiaSmi();
        if (check.installed) return { success: true };
      }
    }

    // Try installing GeForce Experience via winget
    if (onProgress) onProgress({ status: "installing", message: "Installing NVIDIA GeForce Experience..." });
    const install = await runCmd("winget install NVIDIA.GeForceExperience --silent --accept-package-agreements --accept-source-agreements 2>nul", 300000);

    const check = await this.checkNvidiaSmi();
    return {
      success: check.installed,
      error: check.installed ? null : "Could not auto-install nvidia-smi. Download the latest NVIDIA drivers from nvidia.com and install manually.",
    };
  }

  // ── SceWin (AMISCE) ─────────────────────────────────────────────
  async checkSceWin() {
    const scewinDir = path.join(os.homedir(), ".fn-optimizer", "scewin");
    const paths = [
      path.join(scewinDir, "SCEWNX64.exe"),
      path.join(scewinDir, "scewin_64.exe"),
      path.join(__dirname, "tools", "SCEWNX64.exe"),
      "C:\\SCEWNX64.exe",
    ];

    for (const p of paths) {
      try { if (fs.existsSync(p)) return { installed: true, path: p, name: "SceWin (AMISCE)", id: "scewin" }; } catch(e) {}
    }

    return {
      installed: false,
      name: "SceWin (AMISCE)",
      id: "scewin",
      desc: "BIOS settings editor — enables CPU multiplier & voltage control from Windows",
      downloadUrl: "https://github.com/ab3lkaizen/SCEHUB/releases/download/1.2.0/DL_SCEWIN.exe",
      size: "~9 MB",
      required: false,
      feature: "CPU/RAM Overclock",
      installNote: "Downloads SCEHUB installer which extracts SceWin binaries automatically.",
    };
  }

  async installSceWin(onProgress) {
    const scewinDir = path.join(os.homedir(), ".fn-optimizer", "scewin");
    ensureDir(scewinDir);

    if (onProgress) onProgress({ status: "downloading", message: "Downloading SceWin (SCEHUB)...", percent: 0 });
    const destPath = path.join(DOWNLOADS_DIR, "DL_SCEWIN.exe");
    const dl = await downloadFile("https://github.com/ab3lkaizen/SCEHUB/releases/download/1.2.0/DL_SCEWIN.exe", destPath, (p) => {
      if (onProgress) onProgress({ status: "downloading", message: `Downloading SceWin... ${p.downloadedMB}/${p.totalMB} MB`, percent: p.percent });
    });
    if (!dl.success) return { success: false, error: "Download failed: " + dl.error };
    const installerPath = dl.path || destPath;

    if (onProgress) onProgress({ status: "installing", message: "Launching SceWin installer — follow the prompts if a window appears..." });

    // DL_SCEWIN.exe is a custom downloader/extractor from SCEHUB, NOT an NSIS installer.
    // It does not support /S or /SILENT flags.
    // Method 1: Run it non-silently so user can interact, with working dir set to scewinDir
    const psCmd1 = `Start-Process -FilePath '${installerPath.replace(/'/g, "''")}' -WorkingDirectory '${scewinDir.replace(/'/g, "''")}' -Verb RunAs -Wait`;
    await runPS(psCmd1, 300000);

    // Wait for files to appear
    await new Promise(r => setTimeout(r, 3000));

    if (onProgress) onProgress({ status: "searching", message: "Searching for SceWin binaries..." });

    // Search everywhere for the binary
    const found = await this._findSceWinBinary(scewinDir);
    if (found) return { success: true, path: found };

    // Method 2: Maybe it's a self-extractor — try 7z/tar to extract it
    try {
      const extract7z = await runCmd(`"C:\\Program Files\\7-Zip\\7z.exe" x "${installerPath}" -o"${scewinDir}" -y 2>nul`, 30000);
      if (extract7z.success) {
        await new Promise(r => setTimeout(r, 1000));
        const found2 = await this._findSceWinBinary(scewinDir);
        if (found2) return { success: true, path: found2 };
      }
    } catch(e) {}

    // Method 3: Try PowerShell Expand-Archive in case it's a renamed zip
    try {
      const zipCopy = installerPath.replace(/\.exe$/i, ".zip");
      fs.copyFileSync(installerPath, zipCopy);
      await runPS(`Expand-Archive -Path '${zipCopy.replace(/'/g, "''")}' -DestinationPath '${scewinDir.replace(/'/g, "''")}' -Force`, 30000);
      await new Promise(r => setTimeout(r, 1000));
      const found3 = await this._findSceWinBinary(scewinDir);
      if (found3) return { success: true, path: found3 };
    } catch(e) {}

    // Final check
    const check = await this.checkSceWin();
    return {
      success: check.installed,
      error: check.installed ? null : "SceWin binary not found after install. The installer was downloaded to: " + installerPath + "\nTry running it manually and placing SCEWNX64.exe in: " + scewinDir,
    };
  }

  // Helper: search common locations for SceWin binary and copy to scewinDir
  async _findSceWinBinary(scewinDir) {
    // Check scewinDir and subdirs
    const localPaths = [
      path.join(scewinDir, "SCEWNX64.exe"),
      path.join(scewinDir, "scewin_64.exe"),
      path.join(scewinDir, "SCEWIN_64.exe"),
      path.join(scewinDir, "SCEWIN", "SCEWNX64.exe"),
      path.join(scewinDir, "SCEWIN", "scewin_64.exe"),
      path.join(scewinDir, "AMISCE", "SCEWNX64.exe"),
    ];

    for (const p of localPaths) {
      try { if (fs.existsSync(p)) return p; } catch(e) {}
    }

    // Recursive search in scewinDir
    try {
      const findResult = await runCmd(`dir /s /b "${scewinDir}\\*SCEWN*" "${scewinDir}\\*scewin*" 2>nul`, 10000);
      if (findResult.success && findResult.output) {
        const found = findResult.output.split("\n").find(l => l.trim().toLowerCase().endsWith(".exe"));
        if (found && found.trim()) return found.trim();
      }
    } catch(e) {}

    // Check common system locations and copy to scewinDir if found
    const systemPaths = [
      path.join(os.homedir(), "SCEWIN", "SCEWNX64.exe"),
      path.join(os.homedir(), "Desktop", "SCEWIN", "SCEWNX64.exe"),
      path.join(os.homedir(), "Downloads", "SCEWNX64.exe"),
      path.join(os.homedir(), "Downloads", "SCEWIN", "SCEWNX64.exe"),
      "C:\\SCEWIN\\SCEWNX64.exe",
      "C:\\SCEWNX64.exe",
      path.join(DOWNLOADS_DIR, "SCEWIN", "SCEWNX64.exe"),
      path.join(DOWNLOADS_DIR, "SCEWNX64.exe"),
    ];

    for (const p of systemPaths) {
      try {
        if (fs.existsSync(p)) {
          const dest = path.join(scewinDir, path.basename(p));
          fs.copyFileSync(p, dest);
          return dest;
        }
      } catch(e) {}
    }

    // Global search for SCEWNX64.exe on C drive
    try {
      const globalSearch = await runCmd(`where /R C:\\ SCEWNX64.exe 2>nul`, 15000);
      if (globalSearch.success && globalSearch.output) {
        const found = globalSearch.output.split("\n")[0].trim();
        if (found && fs.existsSync(found)) {
          const dest = path.join(scewinDir, path.basename(found));
          try { fs.copyFileSync(found, dest); } catch(e) {}
          return dest;
        }
      }
    } catch(e) {}

    return null;
  }

  // ── Install all missing dependencies ────────────────────────────
  async installAll(onProgress) {
    const status = await this.checkAll();
    const results = {};

    if (!status.wireguard.installed) {
      if (onProgress) onProgress({ dep: "wireguard", status: "installing" });
      results.wireguard = await this.installWireGuard((p) => {
        if (onProgress) onProgress({ dep: "wireguard", ...p });
      });
    } else {
      results.wireguard = { success: true, skipped: true };
    }

    if (!status.nvidiaSmi.installed && status.nvidiaSmi.hasNvidiaGpu) {
      if (onProgress) onProgress({ dep: "nvidia-smi", status: "installing" });
      results.nvidiaSmi = await this.installNvidiaSmi((p) => {
        if (onProgress) onProgress({ dep: "nvidia-smi", ...p });
      });
    } else {
      results.nvidiaSmi = { success: status.nvidiaSmi.installed, skipped: true };
    }

    if (!status.scewin.installed) {
      if (onProgress) onProgress({ dep: "scewin", status: "installing" });
      results.scewin = await this.installSceWin((p) => {
        if (onProgress) onProgress({ dep: "scewin", ...p });
      });
    } else {
      results.scewin = { success: true, skipped: true };
    }

    return results;
  }

  // ── Install a specific dependency ───────────────────────────────
  async installDep(depId, onProgress) {
    switch (depId) {
      case "wireguard": return await this.installWireGuard(onProgress);
      case "nvidia-smi": return await this.installNvidiaSmi(onProgress);
      case "scewin": return await this.installSceWin(onProgress);
      default: return { success: false, error: "Unknown dependency: " + depId };
    }
  }
}

module.exports = DependencyManager;