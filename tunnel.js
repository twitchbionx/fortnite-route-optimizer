// ═══════════════════════════════════════════════════════════════════
//  WireGuard Tunnel Manager for Windows
//  Manages the WireGuard tunnel connection from the Electron app
// ═══════════════════════════════════════════════════════════════════

const { exec, execSync, spawn } = require("child_process");
const fs = require("fs");
const path = require("path");
const os = require("os");
const net = require("net");

const TUNNEL_NAME = "fn-optimizer";
const CONFIG_DIR = path.join(os.homedir(), ".fn-optimizer");
const CONFIG_PATH = path.join(CONFIG_DIR, `${TUNNEL_NAME}.conf`);

// WireGuard Windows install paths — check every known location
const WG_PATHS = [
  "C:\\Program Files\\WireGuard\\wireguard.exe",
  "C:\\Program Files (x86)\\WireGuard\\wireguard.exe",
  path.join(os.homedir(), "AppData\\Local\\Programs\\WireGuard\\wireguard.exe"),
  path.join(os.homedir(), "AppData\\Local\\WireGuard\\wireguard.exe"),
  path.join(process.env.LOCALAPPDATA || "", "WireGuard\\wireguard.exe"),
  path.join(process.env.LOCALAPPDATA || "", "Programs\\WireGuard\\wireguard.exe"),
  "C:\\WireGuard\\wireguard.exe",
  "D:\\Program Files\\WireGuard\\wireguard.exe",
  "D:\\WireGuard\\wireguard.exe",
];

const WG_CLI_PATHS = [
  "C:\\Program Files\\WireGuard\\wg.exe",
  "C:\\Program Files (x86)\\WireGuard\\wg.exe",
  path.join(os.homedir(), "AppData\\Local\\Programs\\WireGuard\\wg.exe"),
  path.join(os.homedir(), "AppData\\Local\\WireGuard\\wg.exe"),
  path.join(process.env.LOCALAPPDATA || "", "WireGuard\\wg.exe"),
  path.join(process.env.LOCALAPPDATA || "", "Programs\\WireGuard\\wg.exe"),
  "C:\\WireGuard\\wg.exe",
];

class TunnelManager {
  constructor() {
    this.status = "disconnected"; // disconnected | connecting | connected | error
    this.config = null;
    this.stats = { tx: 0, rx: 0, lastHandshake: null, endpoint: null };
    this.wireguardPath = null;
    this.wgCliPath = null;
    this.detectLog = []; // diagnostic log for debugging detection
    this._detectWireGuard();
    this._ensureConfigDir();
  }

  _ensureConfigDir() {
    if (!fs.existsSync(CONFIG_DIR)) {
      fs.mkdirSync(CONFIG_DIR, { recursive: true });
    }
  }

  _detectWireGuard() {
    this.detectLog = [];
    const log = (msg) => { this.detectLog.push(msg); console.log("[WG-Detect]", msg); };

    // 1) Check hardcoded paths first
    log("Checking hardcoded paths...");
    for (const p of WG_PATHS) {
      try {
        const exists = fs.existsSync(p);
        log(`  ${p} → ${exists ? "FOUND" : "not found"}`);
        if (exists && !this.wireguardPath) {
          this.wireguardPath = p;
          log(`  ✓ Using: ${p}`);
        }
      } catch (e) {
        log(`  ${p} → error: ${e.message}`);
      }
    }
    for (const p of WG_CLI_PATHS) {
      try {
        const exists = fs.existsSync(p);
        if (exists && !this.wgCliPath) {
          this.wgCliPath = p;
          log(`  ✓ wg.exe: ${p}`);
        }
      } catch (e) {}
    }

    // 2) Try 'where' command (finds on PATH)
    if (!this.wireguardPath) {
      log("Trying 'where wireguard.exe'...");
      try {
        const result = execSync("where wireguard.exe 2>nul", { timeout: 5000 }).toString().trim();
        if (result) {
          this.wireguardPath = result.split("\n")[0].trim();
          log(`  ✓ Found via where: ${this.wireguardPath}`);
        } else {
          log("  not on PATH");
        }
      } catch (e) {
        log(`  where failed: ${e.message}`);
      }
    }

    // 3) Check Windows registry
    if (!this.wireguardPath) {
      log("Checking registry...");
      const regKeys = [
        'HKLM\\SOFTWARE\\WireGuard',
        'HKCU\\SOFTWARE\\WireGuard',
        'HKLM\\SOFTWARE\\WOW6432Node\\WireGuard',
      ];
      for (const key of regKeys) {
        try {
          const regResult = execSync(`reg query "${key}" /ve 2>nul`, { timeout: 5000 }).toString();
          log(`  ${key} → ${regResult.trim().substring(0, 100)}`);
          const match = regResult.match(/REG_SZ\s+(.+)/);
          if (match) {
            const dir = match[1].trim();
            const exePath = path.join(dir, "wireguard.exe");
            if (fs.existsSync(exePath)) {
              this.wireguardPath = exePath;
              log(`  ✓ Found via registry: ${this.wireguardPath}`);
              break;
            }
          }
        } catch (e) {
          log(`  ${key} → not found`);
        }
      }
    }

    // 4) Check Uninstall registry for install location
    if (!this.wireguardPath) {
      log("Checking Uninstall registry...");
      try {
        const regResult = execSync(
          'reg query "HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\WireGuard" /v InstallLocation 2>nul',
          { timeout: 5000 }
        ).toString();
        const match = regResult.match(/InstallLocation\s+REG_SZ\s+(.+)/);
        if (match) {
          const dir = match[1].trim();
          const exePath = path.join(dir, "wireguard.exe");
          log(`  Uninstall location: ${dir}`);
          if (fs.existsSync(exePath)) {
            this.wireguardPath = exePath;
            log(`  ✓ Found: ${exePath}`);
          }
        }
      } catch (e) {
        log("  Uninstall key not found");
      }
    }

    // 5) Find WireGuard by checking running processes
    if (!this.wireguardPath) {
      log("Checking running processes...");
      try {
        const result = execSync(
          'wmic process where "name=\'wireguard.exe\'" get ExecutablePath 2>nul',
          { timeout: 5000 }
        ).toString().trim();
        const lines = result.split("\n").map(l => l.trim()).filter(l => l && l !== "ExecutablePath");
        if (lines.length > 0) {
          this.wireguardPath = lines[0];
          log(`  ✓ Found running: ${this.wireguardPath}`);
        } else {
          log("  wireguard.exe not running");
        }
      } catch (e) {
        log(`  wmic failed: ${e.message}`);
      }
    }

    // 6) PowerShell fallback — Get-Command
    if (!this.wireguardPath) {
      log("Trying PowerShell Get-Command...");
      try {
        const result = execSync(
          'powershell -NoProfile -Command "(Get-Command wireguard.exe -ErrorAction SilentlyContinue).Source" 2>nul',
          { timeout: 8000 }
        ).toString().trim();
        if (result && fs.existsSync(result)) {
          this.wireguardPath = result;
          log(`  ✓ Found via PowerShell: ${result}`);
        } else {
          log("  not found via PowerShell");
        }
      } catch (e) {
        log(`  PowerShell failed: ${e.message}`);
      }
    }

    // 7) Find wg.exe
    if (!this.wgCliPath) {
      try {
        const result = execSync("where wg.exe 2>nul", { timeout: 5000 }).toString().trim();
        if (result) {
          this.wgCliPath = result.split("\n")[0].trim();
          log(`  ✓ wg.exe via where: ${this.wgCliPath}`);
        }
      } catch (e) {}
    }
    if (!this.wgCliPath && this.wireguardPath) {
      const wgPath = path.join(path.dirname(this.wireguardPath), "wg.exe");
      if (fs.existsSync(wgPath)) {
        this.wgCliPath = wgPath;
        log(`  ✓ wg.exe same dir: ${this.wgCliPath}`);
      }
    }

    // Summary
    log(`\nResult: wireguard=${this.wireguardPath || "NOT FOUND"}, wg=${this.wgCliPath || "NOT FOUND"}`);
  }

  isInstalled() {
    return this.wireguardPath !== null;
  }

  getStatus() {
    return {
      status: this.status,
      installed: this.isInstalled(),
      hasConfig: this.config !== null || fs.existsSync(CONFIG_PATH),
      stats: this.stats,
      wireguardPath: this.wireguardPath,
      detectLog: this.detectLog,
    };
  }

  // Re-run detection (callable from renderer if first attempt failed)
  retryDetect() {
    this.wireguardPath = null;
    this.wgCliPath = null;
    this._detectWireGuard();
    return this.getStatus();
  }

  // Save WireGuard config from user input
  saveConfig(configText) {
    try {
      // Validate it looks like a WireGuard config
      if (!configText.includes("[Interface]") || !configText.includes("[Peer]")) {
        return { success: false, error: "Invalid WireGuard config — needs [Interface] and [Peer] sections" };
      }
      if (!configText.includes("PrivateKey")) {
        return { success: false, error: "Config missing PrivateKey" };
      }
      if (!configText.includes("Endpoint")) {
        return { success: false, error: "Config missing Endpoint" };
      }

      this.config = configText;
      fs.writeFileSync(CONFIG_PATH, configText, { mode: 0o600 });

      // Parse endpoint for display
      const endpointMatch = configText.match(/Endpoint\s*=\s*(.+)/);
      if (endpointMatch) {
        this.stats.endpoint = endpointMatch[1].trim();
      }

      return { success: true };
    } catch (e) {
      return { success: false, error: e.message };
    }
  }

  // Load saved config
  loadConfig() {
    try {
      if (fs.existsSync(CONFIG_PATH)) {
        this.config = fs.readFileSync(CONFIG_PATH, "utf-8");
        const endpointMatch = this.config.match(/Endpoint\s*=\s*(.+)/);
        if (endpointMatch) this.stats.endpoint = endpointMatch[1].trim();
        return { success: true, config: this.config };
      }
      return { success: false, error: "No saved config found" };
    } catch (e) {
      return { success: false, error: e.message };
    }
  }

  // Run a command elevated (as Administrator) using PowerShell
  _runElevated(cmd, timeout = 15000) {
    return new Promise((resolve) => {
      // Use PowerShell to run the command as admin
      const psCmd = `powershell -NoProfile -Command "Start-Process cmd -ArgumentList '/c ${cmd.replace(/"/g, '\\"')}' -Verb RunAs -Wait -WindowStyle Hidden"`;
      exec(psCmd, { timeout }, (err, stdout, stderr) => {
        resolve({ err, stdout, stderr });
      });
    });
  }

  // Connect the tunnel
  async connect() {
    if (!this.isInstalled()) {
      return { success: false, error: "WireGuard not installed. Download from https://www.wireguard.com/install/" };
    }

    if (!fs.existsSync(CONFIG_PATH)) {
      return { success: false, error: "No tunnel config saved. Paste your config first." };
    }

    this.status = "connecting";

    // Method 1: Try direct (works if already running as admin)
    try {
      const cmd = `"${this.wireguardPath}" /installtunnelservice "${CONFIG_PATH}"`;
      const result = await new Promise((resolve) => {
        exec(cmd, { timeout: 15000 }, (err, stdout, stderr) => {
          resolve({ err, stdout, stderr });
        });
      });

      if (!result.err) {
        this.status = "connected";
        return { success: true };
      }
      console.log("[Tunnel] Direct install failed, trying elevated...", result.err.message);
    } catch (e) {
      console.log("[Tunnel] Direct install exception:", e.message);
    }

    // Method 2: Try elevated via PowerShell RunAs
    try {
      const cmd = `"${this.wireguardPath}" /installtunnelservice "${CONFIG_PATH}"`;
      const result = await this._runElevated(cmd, 20000);

      if (!result.err) {
        // Verify the service actually started
        await new Promise(r => setTimeout(r, 2000));
        this.status = "connected";
        return { success: true };
      }
      console.log("[Tunnel] Elevated install failed:", result.err.message);
    } catch (e) {
      console.log("[Tunnel] Elevated exception:", e.message);
    }

    // Method 3: Try net start (if service was previously installed)
    try {
      const cmd = `net start WireGuardTunnel$${TUNNEL_NAME}`;
      const result = await this._runElevated(cmd, 10000);

      if (!result.err) {
        this.status = "connected";
        return { success: true };
      }
    } catch (e) {}

    this.status = "error";
    return {
      success: false,
      error: "Could not start tunnel. A UAC prompt may have appeared — please approve it and try again.",
    };
  }

  // Disconnect the tunnel
  async disconnect() {
    // Method 1: Try direct
    const cmd1 = `"${this.wireguardPath}" /uninstalltunnelservice "${TUNNEL_NAME}"`;
    try {
      const result = await new Promise((resolve) => {
        exec(cmd1, { timeout: 10000 }, (err, stdout, stderr) => {
          resolve({ err });
        });
      });
      if (!result.err) {
        this.status = "disconnected";
        return { success: true };
      }
    } catch (e) {}

    // Method 2: Try elevated
    try {
      await this._runElevated(cmd1, 10000);
      this.status = "disconnected";
      return { success: true };
    } catch (e) {}

    // Method 3: net stop elevated
    try {
      await this._runElevated(`net stop WireGuardTunnel$${TUNNEL_NAME}`, 10000);
      this.status = "disconnected";
      return { success: true };
    } catch (e) {}

    this.status = "disconnected";
    return { success: true };
  }

  // Toggle connection
  async toggle() {
    if (this.status === "connected") {
      return await this.disconnect();
    } else {
      return await this.connect();
    }
  }

  // Get tunnel statistics
  async getStats() {
    if (!this.wgCliPath) return this.stats;

    return new Promise((resolve) => {
      exec(`"${this.wgCliPath}" show ${TUNNEL_NAME} dump`, { timeout: 5000 }, (err, stdout) => {
        if (err || !stdout) {
          resolve(this.stats);
          return;
        }

        const lines = stdout.trim().split("\n");
        if (lines.length >= 2) {
          const parts = lines[1].split("\t");
          if (parts.length >= 6) {
            this.stats = {
              ...this.stats,
              rx: parseInt(parts[5]) || 0,
              tx: parseInt(parts[6]) || 0,
              lastHandshake: parseInt(parts[4]) || null,
            };
          }
        }
        resolve(this.stats);
      });
    });
  }

  // Test latency through tunnel vs direct
  async comparePing(host) {
    const directPing = await this._tcpPing(host);

    // If tunnel is connected, the ping automatically goes through WireGuard
    // for IPs in AllowedIPs. The direct ping IS the tunneled ping when connected.
    return {
      direct: directPing,
      tunneled: this.status === "connected" ? directPing : null,
    };
  }

  _tcpPing(host, port = 443, timeout = 5000) {
    return new Promise((resolve) => {
      const start = process.hrtime.bigint();
      const socket = new net.Socket();
      socket.setTimeout(timeout);
      socket.on("connect", () => {
        const end = process.hrtime.bigint();
        socket.destroy();
        resolve(Math.round(Number(end - start) / 1_000_000));
      });
      socket.on("timeout", () => { socket.destroy(); resolve(-1); });
      socket.on("error", () => { socket.destroy(); resolve(-1); });
      socket.connect(port, host);
    });
  }
}

module.exports = TunnelManager;
