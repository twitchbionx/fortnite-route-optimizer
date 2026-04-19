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

// WireGuard Windows install paths
const WG_PATHS = [
  "C:\\Program Files\\WireGuard\\wireguard.exe",
  "C:\\Program Files (x86)\\WireGuard\\wireguard.exe",
  path.join(os.homedir(), "AppData\\Local\\Programs\\WireGuard\\wireguard.exe"),
];

const WG_CLI_PATHS = [
  "C:\\Program Files\\WireGuard\\wg.exe",
  "C:\\Program Files (x86)\\WireGuard\\wg.exe",
];

class TunnelManager {
  constructor() {
    this.status = "disconnected"; // disconnected | connecting | connected | error
    this.config = null;
    this.stats = { tx: 0, rx: 0, lastHandshake: null, endpoint: null };
    this.wireguardPath = null;
    this.wgCliPath = null;
    this._detectWireGuard();
    this._ensureConfigDir();
  }

  _ensureConfigDir() {
    if (!fs.existsSync(CONFIG_DIR)) {
      fs.mkdirSync(CONFIG_DIR, { recursive: true });
    }
  }

  _detectWireGuard() {
    // Check hardcoded paths first
    for (const p of WG_PATHS) {
      if (fs.existsSync(p)) {
        this.wireguardPath = p;
        break;
      }
    }
    for (const p of WG_CLI_PATHS) {
      if (fs.existsSync(p)) {
        this.wgCliPath = p;
        break;
      }
    }

    // If not found, try 'where' command to find it on PATH or registry
    if (!this.wireguardPath) {
      try {
        const result = execSync("where wireguard.exe 2>nul", { timeout: 5000 }).toString().trim();
        if (result) this.wireguardPath = result.split("\n")[0].trim();
      } catch (e) {}
    }
    if (!this.wireguardPath) {
      // Check registry for install path
      try {
        const regResult = execSync('reg query "HKLM\\SOFTWARE\\WireGuard" /ve 2>nul', { timeout: 5000 }).toString();
        const match = regResult.match(/REG_SZ\s+(.+)/);
        if (match) {
          const dir = match[1].trim();
          const exePath = path.join(dir, "wireguard.exe");
          if (fs.existsSync(exePath)) this.wireguardPath = exePath;
        }
      } catch (e) {}
    }
    if (!this.wgCliPath) {
      try {
        const result = execSync("where wg.exe 2>nul", { timeout: 5000 }).toString().trim();
        if (result) this.wgCliPath = result.split("\n")[0].trim();
      } catch (e) {}
    }
    if (!this.wgCliPath && this.wireguardPath) {
      // wg.exe is usually in the same folder as wireguard.exe
      const wgPath = path.join(path.dirname(this.wireguardPath), "wg.exe");
      if (fs.existsSync(wgPath)) this.wgCliPath = wgPath;
    }
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
    };
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

  // Connect the tunnel
  async connect() {
    if (!this.isInstalled()) {
      return { success: false, error: "WireGuard not installed. Download from https://www.wireguard.com/install/" };
    }

    if (!fs.existsSync(CONFIG_PATH)) {
      return { success: false, error: "No tunnel config saved. Paste your config first." };
    }

    this.status = "connecting";

    return new Promise((resolve) => {
      // Use wireguard.exe /installtunnelservice to install and start
      const cmd = `"${this.wireguardPath}" /installtunnelservice "${CONFIG_PATH}"`;

      exec(cmd, { timeout: 15000 }, (err, stdout, stderr) => {
        if (err) {
          // Try alternative: wg-quick approach via wg.exe
          this._tryAlternativeConnect().then(resolve);
          return;
        }
        this.status = "connected";
        resolve({ success: true });
      });
    });
  }

  async _tryAlternativeConnect() {
    return new Promise((resolve) => {
      // Alternative: use netsh + route commands to set up tunnel manually
      // Or try the WireGuard service command
      const cmd = `net start WireGuardTunnel$${TUNNEL_NAME}`;
      exec(cmd, { timeout: 10000 }, (err) => {
        if (err) {
          this.status = "error";
          resolve({
            success: false,
            error: "Could not start tunnel. Make sure WireGuard is installed and you're running as Administrator.",
          });
        } else {
          this.status = "connected";
          resolve({ success: true });
        }
      });
    });
  }

  // Disconnect the tunnel
  async disconnect() {
    return new Promise((resolve) => {
      const cmd = `"${this.wireguardPath}" /uninstalltunnelservice "${TUNNEL_NAME}"`;

      exec(cmd, { timeout: 10000 }, (err) => {
        if (err) {
          // Try stopping the service directly
          exec(`net stop WireGuardTunnel$${TUNNEL_NAME}`, { timeout: 5000 }, () => {
            this.status = "disconnected";
            resolve({ success: true });
          });
          return;
        }
        this.status = "disconnected";
        resolve({ success: true });
      });
    });
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
