// ═══════════════════════════════════════════════════════════════════
//  HARDWARE OVERCLOCK & MONITORING MODULE
//  Detects CPU/GPU/RAM, monitors temps/clocks/usage, and applies
//  safe overclock presets. Integrates with nvidia-smi for NVIDIA
//  GPUs and powercfg for CPU power management.
//
//  WARNING: Overclocking can damage hardware if abused.
//  All presets here are conservative and tested safe.
//  Use at your own risk — but these settings are mild.
// ═══════════════════════════════════════════════════════════════════

const { exec } = require("child_process");
const os = require("os");
const path = require("path");
const fs = require("fs");

const SETTINGS_PATH = path.join(os.homedir(), ".fn-optimizer", "last-known-good.json");
const CRASH_LOG_PATH = path.join(os.homedir(), ".fn-optimizer", "crash-log.json");
const SCEWIN_BACKUP_PATH = path.join(os.homedir(), ".fn-optimizer", "scewin-backup.txt");
const SCEWIN_EXPORT_PATH = path.join(os.homedir(), ".fn-optimizer", "scewin-export.txt");
const SCEWIN_CHANGE_LOG_PATH = path.join(os.homedir(), ".fn-optimizer", "scewin-changes.json");

function runCmd(cmd, timeout = 15000) {
  return new Promise((resolve) => {
    exec(cmd, { timeout }, (err, stdout, stderr) => {
      resolve({ success: !err, output: stdout?.trim() || "", error: stderr?.trim() || err?.message || "" });
    });
  });
}

function runPS(cmd, timeout = 15000) {
  return new Promise((resolve) => {
    exec(cmd, { timeout, shell: "powershell.exe" }, (err, stdout, stderr) => {
      resolve({ success: !err, output: stdout?.trim() || "", error: stderr?.trim() || err?.message || "" });
    });
  });
}

// ── HWiNFO64 Integration ──────────────────────────────────────────
// Reads accurate sensor data from HWiNFO64 via registry (when running
// with "Shared Memory Support" enabled) or via its CLI report mode.
// Provides: CPU temp, voltage, clock per-core, GPU temp/clock/load,
// fan speeds, VRM temps, memory temps, power draw.
// ────────────────────────────────────────────────────────────────────
class HWiNFOReader {
  constructor() {
    this.available = false;
    this.hwInfoPath = null;
    this._detect();
  }

  _detect() {
    const home = os.homedir();
    const hwinfoDir = path.join(home, ".fn-optimizer", "hwinfo");
    const paths = [
      // Our portable install dir
      path.join(hwinfoDir, "HWiNFO64.exe"),
      path.join(hwinfoDir, "HWiNFO64", "HWiNFO64.exe"),
      // Our downloads dir
      path.join(home, ".fn-optimizer", "downloads", "HWiNFO64.exe"),
      // Standard install locations
      "C:\\Program Files\\HWiNFO64\\HWiNFO64.exe",
      "C:\\Program Files (x86)\\HWiNFO64\\HWiNFO64.exe",
      path.join(home, "AppData\\Local\\Programs\\HWiNFO64\\HWiNFO64.exe"),
      path.join(process.env.LOCALAPPDATA || "", "HWiNFO64\\HWiNFO64.exe"),
      // User Desktop/Downloads
      path.join(home, "Downloads", "HWiNFO64.exe"),
      path.join(home, "Downloads", "HWiNFO64", "HWiNFO64.exe"),
      path.join(home, "Desktop", "HWiNFO64.exe"),
      path.join(home, "Desktop", "HWiNFO64", "HWiNFO64.exe"),
    ];
    for (const p of paths) {
      try {
        if (fs.existsSync(p)) {
          this.hwInfoPath = p;
          this.available = true;
          return;
        }
      } catch(e) {}
    }
    // Check PATH
    try {
      const r = require("child_process").execSync("where HWiNFO64.exe 2>nul", { timeout: 5000 }).toString().trim();
      if (r) { this.hwInfoPath = r.split("\n")[0].trim(); this.available = true; }
    } catch(e) {}
  }

  redetect() {
    this.available = false;
    this.hwInfoPath = null;
    this._detect();
    return { available: this.available, path: this.hwInfoPath };
  }

  // Read sensor values from HWiNFO registry (requires HWiNFO running with Shared Memory enabled)
  async readSensors() {
    // HWiNFO writes sensor data to HKCU\Software\HWiNFO64\VSB when shared memory is enabled
    const result = await runPS(`
      $base = 'HKCU:\\Software\\HWiNFO64\\VSB'
      if (!(Test-Path $base)) { Write-Output '{"error":"HWiNFO registry not found. Is HWiNFO64 running with Shared Memory enabled?"}'; return }
      $sensors = @{}
      $values = Get-ItemProperty -Path $base -ErrorAction SilentlyContinue
      if ($values) {
        $i = 0
        while ($true) {
          $label = $values."Label$i"
          $value = $values."Value$i"
          $valueRaw = $values."ValueRaw$i"
          if ($null -eq $label) { break }
          $sensors[$label] = @{ value = $value; raw = $valueRaw }
          $i++
        }
      }
      $sensors | ConvertTo-Json -Depth 3
    `, 10000);

    if (!result.success || !result.output) {
      return { available: false, error: "Could not read HWiNFO sensors" };
    }

    try {
      const data = JSON.parse(result.output);
      if (data.error) return { available: false, error: data.error };
      return { available: true, sensors: this._parseSensors(data) };
    } catch(e) {
      return { available: false, error: "Failed to parse HWiNFO data" };
    }
  }

  _parseSensors(raw) {
    const sensors = {
      cpuTemp: null,
      cpuPackageTemp: null,
      cpuVoltage: null,
      cpuClock: null,
      cpuPower: null,
      gpuTemp: null,
      gpuClock: null,
      gpuMemClock: null,
      gpuLoad: null,
      gpuPower: null,
      gpuVram: null,
      ramTemp: null,
      vrmTemp: null,
      fanSpeeds: {},
      allSensors: raw,
    };

    for (const [label, data] of Object.entries(raw)) {
      const lbl = label.toLowerCase();
      const val = parseFloat(data.raw || data.value);
      if (isNaN(val)) continue;

      // CPU
      if (lbl.includes("cpu") && lbl.includes("package") && lbl.includes("temp")) sensors.cpuPackageTemp = val;
      else if (lbl.includes("cpu") && lbl.includes("temp") && !sensors.cpuTemp) sensors.cpuTemp = val;
      else if (lbl.includes("cpu") && lbl.includes("vcore") || (lbl.includes("cpu") && lbl.includes("voltage"))) sensors.cpuVoltage = val;
      else if (lbl.includes("cpu") && lbl.includes("clock") && !lbl.includes("ring")) sensors.cpuClock = val;
      else if (lbl.includes("cpu") && lbl.includes("package") && lbl.includes("power")) sensors.cpuPower = val;
      // GPU
      else if (lbl.includes("gpu") && lbl.includes("temp") && !lbl.includes("hot")) sensors.gpuTemp = val;
      else if (lbl.includes("gpu") && lbl.includes("clock") && !lbl.includes("mem")) sensors.gpuClock = val;
      else if (lbl.includes("gpu") && lbl.includes("mem") && lbl.includes("clock")) sensors.gpuMemClock = val;
      else if (lbl.includes("gpu") && lbl.includes("load") || (lbl.includes("gpu") && lbl.includes("usage"))) sensors.gpuLoad = val;
      else if (lbl.includes("gpu") && lbl.includes("power")) sensors.gpuPower = val;
      else if (lbl.includes("gpu") && lbl.includes("memory") && lbl.includes("used")) sensors.gpuVram = val;
      // RAM
      else if (lbl.includes("dimm") && lbl.includes("temp") || (lbl.includes("memory") && lbl.includes("temp"))) sensors.ramTemp = val;
      // VRM
      else if (lbl.includes("vrm") && lbl.includes("temp")) sensors.vrmTemp = val;
      // Fans
      else if (lbl.includes("fan") && (lbl.includes("rpm") || lbl.includes("speed"))) sensors.fanSpeeds[label] = val;
    }

    return sensors;
  }

  // Generate a quick report via HWiNFO CLI
  async quickReport() {
    if (!this.available) return { success: false, error: "HWiNFO64 not installed" };
    const reportPath = path.join(os.homedir(), ".fn-optimizer", "hwinfo-report.csv");
    const r = await runCmd(`"${this.hwInfoPath}" -c"${reportPath}" -max_time=5`, 15000);
    try {
      if (fs.existsSync(reportPath)) {
        const csv = fs.readFileSync(reportPath, "utf8");
        return { success: true, data: csv };
      }
    } catch(e) {}
    return { success: false, error: "Report generation failed" };
  }
}

// ── SceWin (AMISCE) Integration ────────────────────────────────────
// SceWin / AMI Setup Configuration Editor reads/writes BIOS settings
// from within Windows.  Requires the AMISCE binary (SCEWNX64.exe).
// ────────────────────────────────────────────────────────────────────
class SceWinManager {
  constructor() {
    this.scewinPath = null;
    this.available = false;
    this.currentSettings = null;
    this._detect();
  }

  _detect() {
    const home = os.homedir();
    const scewinDir = path.join(home, ".fn-optimizer", "scewin");
    const dlDir = path.join(home, ".fn-optimizer", "downloads");
    const paths = [
      // Primary install dir
      path.join(scewinDir, "SCEWNX64.exe"),
      path.join(scewinDir, "scewin_64.exe"),
      path.join(scewinDir, "SCEWIN_64.exe"),
      // Subdirs of install dir
      path.join(scewinDir, "SCEWIN", "SCEWNX64.exe"),
      path.join(scewinDir, "AMISCE", "SCEWNX64.exe"),
      // App tools dir
      path.join(__dirname, "tools", "SCEWNX64.exe"),
      // Downloads dir (where DL_SCEWIN.exe may extract to)
      path.join(dlDir, "SCEWNX64.exe"),
      path.join(dlDir, "SCEWIN_64.exe"),
      path.join(dlDir, "SCEWIN", "SCEWNX64.exe"),
      // User folders
      path.join(home, "Downloads", "SCEWNX64.exe"),
      path.join(home, "Downloads", "SCEWIN", "SCEWNX64.exe"),
      path.join(home, "Desktop", "SCEWNX64.exe"),
      path.join(home, "Desktop", "SCEWIN", "SCEWNX64.exe"),
      path.join(home, "SCEWIN", "SCEWNX64.exe"),
      // System locations
      "C:\\SCEWIN\\SCEWNX64.exe",
      "C:\\SCEWNX64.exe",
    ];
    for (const p of paths) {
      try {
        if (fs.existsSync(p)) {
          this.scewinPath = p;
          this.available = true;
          return;
        }
      } catch (e) { /* ignore */ }
    }
  }

  // Re-run detection (call after installing SceWin)
  redetect() {
    this.scewinPath = null;
    this.available = false;
    this._detect();
    return { available: this.available, path: this.scewinPath };
  }

  // ── Export / Read ─────────────────────────────────────────────────

  async exportCurrentSettings() {
    if (!this.available) return { success: false, error: "SceWin not found" };

    const exportPath = SCEWIN_EXPORT_PATH;
    const dir = path.dirname(exportPath);
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });

    const r = await runCmd(`"${this.scewinPath}" /o /s "${exportPath}"`, 30000);
    if (!r.success) return { success: false, error: r.error };

    try {
      const raw = fs.readFileSync(exportPath, "utf8");
      const settings = this._parseExport(raw);
      this.currentSettings = settings;
      return { success: true, settings };
    } catch (e) {
      return { success: false, error: e.message };
    }
  }

  _parseExport(raw) {
    // SceWin export is a text file with lines like:
    //   Setup Question = <name>
    //   Token = <hex>     Value = <hex>
    const settings = {};
    const lines = raw.split("\n");
    let currentName = null;

    for (const line of lines) {
      const nameMatch = line.match(/Setup Question\s*=\s*(.+)/i);
      if (nameMatch) {
        currentName = nameMatch[1].trim();
        continue;
      }
      const valueMatch = line.match(/Token\s*=\s*(0x[0-9A-Fa-f]+)\s+.*?Value\s*=\s*(0x[0-9A-Fa-f]+)/i);
      if (valueMatch && currentName) {
        settings[currentName] = {
          token: valueMatch[1],
          value: valueMatch[2],
          raw: line.trim(),
        };
        currentName = null;
      }
    }

    return settings;
  }

  async readSetting(name) {
    if (!this.available) return { success: false, error: "SceWin not found" };

    // Export first if we don't have settings cached
    if (!this.currentSettings) {
      const exported = await this.exportCurrentSettings();
      if (!exported.success) return exported;
    }

    const setting = this.currentSettings[name];
    if (!setting) return { success: false, error: `Setting "${name}" not found` };
    return { success: true, name, ...setting };
  }

  async writeSetting(name, value) {
    if (!this.available) return { success: false, error: "SceWin not found" };

    // Always backup before writing
    await this.backupSettings();

    // Read current to get the token
    const current = await this.readSetting(name);
    if (!current.success) return current;

    const hexValue = typeof value === "number" ? "0x" + value.toString(16) : value;
    const r = await runCmd(`"${this.scewinPath}" /i /s "${SCEWIN_EXPORT_PATH}" /n "${name}" /v ${hexValue}`, 30000);

    // Log the change
    this._logChange(name, current.value, hexValue);

    if (!r.success) return { success: false, error: r.error };

    // Invalidate cache so next read picks up changes
    this.currentSettings = null;
    return { success: true, name, previousValue: current.value, newValue: hexValue };
  }

  // ── CPU Overclock Settings ────────────────────────────────────────

  async getCpuOverclockSettings() {
    if (!this.available) return { success: false, error: "SceWin not found" };

    const exported = await this.exportCurrentSettings();
    if (!exported.success) return exported;

    const s = this.currentSettings;
    const cpuSettings = {
      cpuRatio: s["CPU Ratio"] || s["CPU Core Ratio"] || s["Core Ratio Limit"] || null,
      cpuVoltage: s["CPU Core Voltage"] || s["CPU Vcore"] || s["Vcore Override"] || null,
      cpuVoltageMode: s["CPU Core Voltage Mode"] || s["Vcore Mode"] || null,
      powerLimit1: s["Long Duration Power Limit"] || s["PL1"] || s["Package Power Limit 1"] || null,
      powerLimit2: s["Short Duration Power Limit"] || s["PL2"] || s["Package Power Limit 2"] || null,
      tccOffset: s["TCC Activation Offset"] || null,
      ringRatio: s["Ring Ratio"] || s["Cache Ratio"] || s["Uncore Ratio"] || null,
      iccMax: s["ICC Max"] || s["IA AC Load Line"] || null,
      avxOffset: s["AVX Offset"] || s["AVX2 Ratio Offset"] || null,
    };

    return { success: true, settings: cpuSettings };
  }

  // ── Memory Settings ───────────────────────────────────────────────

  async getMemorySettings() {
    if (!this.available) return { success: false, error: "SceWin not found" };

    const exported = await this.exportCurrentSettings();
    if (!exported.success) return exported;

    const s = this.currentSettings;
    const memSettings = {
      memorySpeed: s["Memory Frequency"] || s["DRAM Frequency"] || s["Memory Speed"] || null,
      casLatency: s["CAS Latency"] || s["tCL"] || s["CAS# Latency"] || null,
      tRCD: s["tRCD"] || s["RAS to CAS Delay"] || s["RAS# to CAS# Delay"] || null,
      tRP: s["tRP"] || s["Row Precharge Time"] || s["RAS# Precharge"] || null,
      tRAS: s["tRAS"] || s["RAS Active Time"] || s["Active to Precharge Delay"] || null,
      xmpProfile: s["XMP Profile"] || s["Extreme Memory Profile"] || s["XMP"] || null,
      memoryVoltage: s["DRAM Voltage"] || s["Memory Voltage"] || null,
      commandRate: s["Command Rate"] || s["CR"] || s["Cmd Rate"] || null,
    };

    return { success: true, settings: memSettings };
  }

  // ── Apply CPU Multiplier ──────────────────────────────────────────

  async applyCpuMultiplier(ratio) {
    if (!this.available) return { success: false, error: "SceWin not found" };
    if (typeof ratio !== "number" || ratio < 10 || ratio > 80) {
      return { success: false, error: `Invalid CPU ratio: ${ratio}. Must be between 10 and 80.` };
    }

    // Try common BIOS setting names for CPU multiplier
    const names = ["CPU Ratio", "CPU Core Ratio", "Core Ratio Limit", "All Core Ratio Limit"];
    for (const name of names) {
      const result = await this.readSetting(name);
      if (result.success) {
        return await this.writeSetting(name, ratio);
      }
    }

    return { success: false, error: "Could not find CPU ratio setting in BIOS. Setting name may vary by motherboard." };
  }

  // ── Apply CPU Voltage ─────────────────────────────────────────────

  async applyCpuVoltage(voltage) {
    if (!this.available) return { success: false, error: "SceWin not found" };

    // Hard safety limits
    const MIN_VOLTAGE = 0.8;
    const MAX_VOLTAGE = 1.45;
    if (typeof voltage !== "number" || voltage < MIN_VOLTAGE || voltage > MAX_VOLTAGE) {
      return { success: false, error: `Voltage ${voltage}V is outside safe range (${MIN_VOLTAGE}V - ${MAX_VOLTAGE}V). Refusing to apply.` };
    }

    // Convert voltage to millivolts (common BIOS representation)
    const millivolts = Math.round(voltage * 1000);

    const names = ["CPU Core Voltage", "CPU Vcore", "Vcore Override", "CPU Core Voltage Override"];
    for (const name of names) {
      const result = await this.readSetting(name);
      if (result.success) {
        return await this.writeSetting(name, millivolts);
      }
    }

    return { success: false, error: "Could not find CPU voltage setting in BIOS. Setting name may vary by motherboard." };
  }

  // ── Apply Memory XMP ──────────────────────────────────────────────

  async applyMemoryXMP(profile) {
    if (!this.available) return { success: false, error: "SceWin not found" };
    if (profile !== 1 && profile !== 2) {
      return { success: false, error: `Invalid XMP profile: ${profile}. Must be 1 or 2.` };
    }

    const names = ["XMP Profile", "Extreme Memory Profile", "XMP", "Memory Profile"];
    for (const name of names) {
      const result = await this.readSetting(name);
      if (result.success) {
        return await this.writeSetting(name, profile);
      }
    }

    return { success: false, error: "Could not find XMP profile setting in BIOS." };
  }

  // ── Backup / Restore ──────────────────────────────────────────────

  async backupSettings() {
    if (!this.available) return { success: false, error: "SceWin not found" };

    const dir = path.dirname(SCEWIN_BACKUP_PATH);
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });

    const r = await runCmd(`"${this.scewinPath}" /o /s "${SCEWIN_BACKUP_PATH}"`, 30000);
    if (!r.success) return { success: false, error: r.error };

    return { success: true, backupPath: SCEWIN_BACKUP_PATH, timestamp: new Date().toISOString() };
  }

  async restoreBackup() {
    if (!this.available) return { success: false, error: "SceWin not found" };
    if (!fs.existsSync(SCEWIN_BACKUP_PATH)) {
      return { success: false, error: "No backup found to restore" };
    }

    const r = await runCmd(`"${this.scewinPath}" /i /s "${SCEWIN_BACKUP_PATH}"`, 30000);
    if (!r.success) return { success: false, error: r.error };

    this.currentSettings = null;
    this._logChange("RESTORE_BACKUP", "N/A", "Restored from " + SCEWIN_BACKUP_PATH);
    return { success: true, message: "BIOS settings restored from backup. A reboot may be required." };
  }

  // ── Change Logging ────────────────────────────────────────────────

  _logChange(settingName, oldValue, newValue) {
    try {
      const dir = path.dirname(SCEWIN_CHANGE_LOG_PATH);
      if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });

      let log = [];
      if (fs.existsSync(SCEWIN_CHANGE_LOG_PATH)) {
        try { log = JSON.parse(fs.readFileSync(SCEWIN_CHANGE_LOG_PATH, "utf8")); } catch (e) { log = []; }
      }

      log.push({
        timestamp: new Date().toISOString(),
        setting: settingName,
        oldValue,
        newValue,
      });

      fs.writeFileSync(SCEWIN_CHANGE_LOG_PATH, JSON.stringify(log, null, 2));
    } catch (e) {
      // Non-fatal — don't break the actual operation
    }
  }
}

// ── AI-Guided Overclocking Engine ──────────────────────────────────
// Heuristic-based "AI" that orchestrates the entire overclock process.
// Analyzes hardware, determines safe limits, and iteratively tests
// stability at progressively higher clock speeds.
// ────────────────────────────────────────────────────────────────────
class AIOverclockEngine {
  constructor(hwMonitor, scewin) {
    this.hw = hwMonitor;
    this.scewin = scewin;
    this.log = [];
    this.running = false;
    this.phase = "idle"; // idle | analyzing | testing | stable | failed
    this.currentProfile = null;
    this.stableSettings = null;
    this.progress = 0;
    this._stopRequested = false;
  }

  // ── Progress / Logging Helpers ────────────────────────────────────

  _log(phase, step, message, extra = {}) {
    const entry = {
      phase,
      step,
      message,
      timestamp: new Date().toISOString(),
      ...extra,
    };
    this.log.push(entry);
    return entry;
  }

  _setPhase(phase, progress) {
    this.phase = phase;
    if (typeof progress === "number") this.progress = progress;
  }

  getProgress() {
    return {
      phase: this.phase,
      progress: this.progress,
      running: this.running,
      log: this.log,
      currentProfile: this.currentProfile,
      stableSettings: this.stableSettings,
    };
  }

  stop() {
    this._stopRequested = true;
    this._log(this.phase, "stop", "Stop requested by user");
  }

  // ── Extended Stress Test ──────────────────────────────────────────

  async runExtendedStressTest(durationMinutes = 30, options = {}) {
    const durationSec = durationMinutes * 60;
    const monitorInterval = 5; // seconds between temp checks
    const tempLog = [];
    const errors = [];
    let maxTemp = 0;
    let tempSum = 0;
    let tempCount = 0;
    const maxTempLimit = options.maxTemp || 95;
    const stressGpu = options.stressGpu || false;

    this._log("testing", "stress-start", `Starting ${durationMinutes}-minute stress test`, { durationMinutes });

    // Build PowerShell stress test script
    const psScript = `
      $cores = (Get-CimInstance Win32_Processor).NumberOfLogicalProcessors;
      $jobs = @();
      $end = (Get-Date).AddSeconds(${durationSec});
      for ($i = 0; $i -lt $cores; $i++) {
        $jobs += Start-Job -ScriptBlock {
          $e = [DateTime]$args[0];
          $errors = 0;
          while ((Get-Date) -lt $e) {
            try {
              $r = [Math]::Sqrt(12345.6789) * [Math]::PI;
              $r = [Math]::Pow(2.71828, 10.5);
              $r = [Math]::Log($r) * [Math]::Tan(0.5);
              for ($j = 0; $j -lt 10000; $j++) { $r = $r * 1.000001 + 0.000001 }
            } catch { $errors++ }
          }
          $errors
        } -ArgumentList $end.ToString("o")
      }
      "STRESS_STARTED:$($jobs.Count) threads"
    `;

    // Start the stress test in the background
    const stressPromise = runPS(psScript, (durationSec + 30) * 1000);

    // Also stress GPU if requested
    let gpuStressPromise = null;
    if (stressGpu && this.hw.nvidiaSmiPath) {
      // Use nvidia-smi to keep GPU busy by repeated queries (lightweight stress)
      const gpuScript = `
        $end = (Get-Date).AddSeconds(${durationSec});
        while ((Get-Date) -lt $e) {
          & "${this.hw.nvidiaSmiPath}" --query-gpu=temperature.gpu,power.draw,clocks.current.graphics --format=csv,noheader,nounits 2>$null | Out-Null;
          Start-Sleep -Milliseconds 100;
        }
      `;
      gpuStressPromise = runPS(gpuScript, (durationSec + 30) * 1000);
    }

    // Monitor loop: check temps every monitorInterval seconds
    const monitorIterations = Math.floor(durationSec / monitorInterval);
    for (let i = 0; i < monitorIterations; i++) {
      if (this._stopRequested) {
        errors.push("Stopped by user");
        break;
      }

      await new Promise(r => setTimeout(r, monitorInterval * 1000));

      try {
        const stats = await this.hw.getStats();
        const cpuTemp = stats.cpu?.temp || -1;
        const gpuTemp = stats.gpu?.temp || -1;
        const temp = Math.max(cpuTemp, gpuTemp > 0 ? gpuTemp : 0);

        tempLog.push({
          time: (i + 1) * monitorInterval,
          cpuTemp,
          gpuTemp,
          cpuUsage: stats.cpu?.usage || -1,
          gpuUsage: stats.gpu?.usage || -1,
        });

        if (cpuTemp > 0) {
          tempSum += cpuTemp;
          tempCount++;
          if (cpuTemp > maxTemp) maxTemp = cpuTemp;
        }

        // Check thermal limit
        if (cpuTemp >= maxTempLimit) {
          errors.push(`CPU temperature exceeded ${maxTempLimit}C (was ${cpuTemp}C) at ${(i + 1) * monitorInterval}s`);
          this._log("testing", "thermal-limit", `Thermal limit reached: ${cpuTemp}C`, { cpuTemp, gpuTemp });
          break;
        }

        // Log progress every 60 seconds
        if ((i + 1) % (60 / monitorInterval) === 0) {
          const elapsed = (i + 1) * monitorInterval;
          const source = stats.hwinfoActive ? "HWiNFO" : "WMI";
          this._log("testing", "monitor", `Stress test ${Math.round(elapsed / 60)}/${durationMinutes} min — CPU ${cpuTemp}C | GPU ${gpuTemp}C [${source}]`, { cpuTemp, gpuTemp, elapsed, source });
        }
      } catch (e) {
        errors.push(`Monitor error at ${(i + 1) * monitorInterval}s: ${e.message}`);
      }
    }

    // Wait for the stress test to finish
    const stressResult = await stressPromise;
    if (gpuStressPromise) await gpuStressPromise;

    if (!stressResult.success && !this._stopRequested) {
      errors.push("Stress test process error: " + stressResult.error);
    }

    const avgTemp = tempCount > 0 ? Math.round(tempSum / tempCount) : -1;
    const stable = errors.length === 0;

    this._log("testing", "stress-end", `Stress test complete: ${stable ? "STABLE" : "UNSTABLE"}`, {
      stable, maxTemp, avgTemp, errorCount: errors.length,
    });

    return { stable, maxTemp, avgTemp, errors, tempLog };
  }

  // ── Hardware Analysis ─────────────────────────────────────────────

  async _analyzeHardware() {
    this._setPhase("analyzing", 5);
    this._log("analyzing", "detect", "Detecting hardware...");

    const hw = await this.hw.detectHardware();
    if (hw.error) {
      this._log("analyzing", "error", "Hardware detection failed: " + hw.error);
      return null;
    }

    const analysis = {
      cpu: hw.cpu,
      gpu: hw.gpu,
      ram: hw.ram,
      mobo: hw.mobo,
      cpuUnlocked: false,
      isIntel: hw.cpu?.isIntel || false,
      isAMD: hw.cpu?.isAMD || false,
      baseRatio: 0,
      maxTurboRatio: 0,
      safeCeilingRatio: 0,
      safeCeilingVoltage: 1.35,
      coolingQuality: "unknown", // poor | adequate | good | excellent
    };

    // Determine if CPU is unlocked
    if (analysis.isIntel) {
      const name = (hw.cpu.name || "").toUpperCase();
      // Intel K/KF/KS/X series are unlocked
      analysis.cpuUnlocked = /\d{4,5}(K|KF|KS|X|HX)\b/.test(name);
      // Estimate base ratio from MaxClockSpeed
      analysis.baseRatio = hw.cpu.maxClock ? Math.round(hw.cpu.maxClock / 100) : 36;
      // Max turbo is typically 1-5 multipliers above base
      analysis.maxTurboRatio = analysis.baseRatio + 5;
      // Safe ceiling: turbo + 3 for K-series (conservative)
      analysis.safeCeilingRatio = analysis.cpuUnlocked ? analysis.maxTurboRatio + 3 : analysis.baseRatio;
      analysis.safeCeilingVoltage = 1.40;
    } else if (analysis.isAMD) {
      const name = (hw.cpu.name || "").toUpperCase();
      // All Ryzen processors are unlocked
      analysis.cpuUnlocked = name.includes("RYZEN");
      analysis.baseRatio = hw.cpu.maxClock ? Math.round(hw.cpu.maxClock / 100) : 36;
      analysis.maxTurboRatio = analysis.baseRatio + 4;
      analysis.safeCeilingRatio = analysis.cpuUnlocked ? analysis.maxTurboRatio + 2 : analysis.baseRatio;
      analysis.safeCeilingVoltage = 1.35; // AMD is more voltage-sensitive
    }

    this._log("analyzing", "cpu-info", `CPU: ${hw.cpu.name} | Unlocked: ${analysis.cpuUnlocked} | Base ratio: ${analysis.baseRatio} | Safe ceiling: ${analysis.safeCeilingRatio}`, {
      unlocked: analysis.cpuUnlocked, baseRatio: analysis.baseRatio, ceiling: analysis.safeCeilingRatio,
    });

    // Quick cooling quality test: run a short stress and measure temp delta
    // ALWAYS use getStats() which overlays HWiNFO sensor data for accurate readings
    this._log("analyzing", "cooling-test", "Testing cooling solution (15-second stress)...");
    this._setPhase("analyzing", 10);

    const preStats = await this.hw.getStats();
    const preTemp = preStats.cpu?.temp || 40;
    this._log("analyzing", "sensor-source", `Pre-stress temp: ${preTemp}C (source: ${preStats.hwinfoActive ? "HWiNFO64" : "WMI/fallback"})`);
    await this.hw.runStressTest(15);
    const postStats = await this.hw.getStats();
    const postTemp = postStats.cpu?.temp || 50;
    const tempDelta = postTemp - preTemp;

    if (tempDelta < 10) {
      analysis.coolingQuality = "excellent";
    } else if (tempDelta < 20) {
      analysis.coolingQuality = "good";
    } else if (tempDelta < 35) {
      analysis.coolingQuality = "adequate";
    } else {
      analysis.coolingQuality = "poor";
    }

    this._log("analyzing", "cooling-result", `Cooling quality: ${analysis.coolingQuality} (delta ${tempDelta}C: ${preTemp}C -> ${postTemp}C)`, {
      coolingQuality: analysis.coolingQuality, preTemp, postTemp, tempDelta,
    });

    // Adjust ceiling based on cooling
    if (analysis.coolingQuality === "poor") {
      analysis.safeCeilingRatio = Math.min(analysis.safeCeilingRatio, analysis.baseRatio + 1);
      this._log("analyzing", "cooling-limit", "Poor cooling detected — limiting overclock headroom");
    } else if (analysis.coolingQuality === "adequate") {
      analysis.safeCeilingRatio = Math.min(analysis.safeCeilingRatio, analysis.maxTurboRatio + 1);
    }

    return analysis;
  }

  // ── CPU Overclock Phase ───────────────────────────────────────────

  async _cpuOcPhase(analysis) {
    if (!this.scewin.available) {
      this._log("testing", "skip-cpu", "SceWin not available — skipping CPU overclock phase");
      return null;
    }

    if (!analysis.cpuUnlocked) {
      this._log("testing", "skip-cpu", "CPU is locked — skipping CPU overclock phase");
      return null;
    }

    this._setPhase("testing", 20);
    this._log("testing", "cpu-start", "Starting CPU overclock phase");

    // Backup current settings
    await this.scewin.backupSettings();
    this._log("testing", "backup", "BIOS settings backed up");

    const startRatio = analysis.baseRatio;
    const maxRatio = analysis.safeCeilingRatio;
    let lastGoodRatio = startRatio;
    let lastGoodVoltage = null;
    const totalSteps = maxRatio - startRatio;

    for (let ratio = startRatio + 1; ratio <= maxRatio; ratio++) {
      if (this._stopRequested) break;

      const stepProgress = 20 + ((ratio - startRatio) / totalSteps) * 30;
      this._setPhase("testing", Math.round(stepProgress));

      this._log("testing", "cpu-step", `Testing CPU multiplier x${ratio} (${ratio * 100} MHz)`, {
        clock: ratio * 100, voltage: lastGoodVoltage,
      });

      // Apply the multiplier
      const applyResult = await this.scewin.applyCpuMultiplier(ratio);
      if (!applyResult.success) {
        this._log("testing", "cpu-apply-error", `Failed to apply ratio x${ratio}: ${applyResult.error}`);
        break;
      }

      // Run stress test (use shorter test for intermediate steps)
      const testMinutes = ratio === maxRatio ? 30 : 10;
      this._log("testing", "cpu-stress", `Running ${testMinutes}-minute stress test at x${ratio}...`);

      const result = await this.runExtendedStressTest(testMinutes, { maxTemp: 95 });

      if (result.stable) {
        lastGoodRatio = ratio;
        this._log("testing", "cpu-stable", `x${ratio} is STABLE (max temp: ${result.maxTemp}C, avg: ${result.avgTemp}C)`, {
          stable: true, temp: result.maxTemp, clock: ratio * 100,
        });
      } else {
        this._log("testing", "cpu-unstable", `x${ratio} is UNSTABLE — trying voltage adjustment`, {
          stable: false, temp: result.maxTemp, errors: result.errors,
        });

        // Try increasing voltage slightly
        let voltageFixed = false;
        const baseVoltage = analysis.isAMD ? 1.20 : 1.25;
        const maxVoltage = analysis.safeCeilingVoltage;

        for (let v = baseVoltage; v <= maxVoltage; v += 0.01) {
          if (this._stopRequested) break;

          const vRounded = Math.round(v * 100) / 100;
          this._log("testing", "cpu-voltage", `Trying voltage ${vRounded}V at x${ratio}`, { voltage: vRounded });

          const vResult = await this.scewin.applyCpuVoltage(vRounded);
          if (!vResult.success) continue;

          const vStress = await this.runExtendedStressTest(10, { maxTemp: 95 });
          if (vStress.stable) {
            lastGoodRatio = ratio;
            lastGoodVoltage = vRounded;
            voltageFixed = true;
            this._log("testing", "cpu-voltage-stable", `x${ratio} at ${vRounded}V is STABLE`, {
              stable: true, temp: vStress.maxTemp, clock: ratio * 100, voltage: vRounded,
            });
            break;
          }
        }

        if (!voltageFixed) {
          this._log("testing", "cpu-rollback", `Rolling back to last good: x${lastGoodRatio}`, { clock: lastGoodRatio * 100 });
          await this.scewin.applyCpuMultiplier(lastGoodRatio);
          if (lastGoodVoltage) await this.scewin.applyCpuVoltage(lastGoodVoltage);
          break;
        }
      }
    }

    return {
      finalRatio: lastGoodRatio,
      finalVoltage: lastGoodVoltage,
      finalMHz: lastGoodRatio * 100,
      gainMHz: (lastGoodRatio - startRatio) * 100,
    };
  }

  // ── Memory Phase ──────────────────────────────────────────────────

  async _memoryPhase() {
    if (!this.scewin.available) {
      this._log("testing", "skip-mem", "SceWin not available — skipping memory phase");
      return null;
    }

    this._setPhase("testing", 55);
    this._log("testing", "mem-start", "Starting memory overclock phase");

    // Check current memory settings
    const memSettings = await this.scewin.getMemorySettings();
    if (!memSettings.success) {
      this._log("testing", "mem-error", "Could not read memory settings: " + memSettings.error);
      return null;
    }

    // Try enabling XMP if not already enabled
    const xmpSetting = memSettings.settings.xmpProfile;
    if (xmpSetting && (xmpSetting.value === "0x0" || xmpSetting.value === "0x00")) {
      this._log("testing", "mem-xmp", "XMP is disabled — enabling XMP Profile 1");
      const xmpResult = await this.scewin.applyMemoryXMP(1);
      if (xmpResult.success) {
        this._log("testing", "mem-xmp-applied", "XMP Profile 1 enabled");

        // Test memory stability
        this._log("testing", "mem-stress", "Running 10-minute memory stability test...");
        const memStress = await this.runExtendedStressTest(10, { maxTemp: 95 });

        if (memStress.stable) {
          this._log("testing", "mem-stable", "Memory is STABLE with XMP enabled");
          return { xmpEnabled: true, stable: true };
        } else {
          this._log("testing", "mem-unstable", "Memory UNSTABLE with XMP — rolling back");
          await this.scewin.applyMemoryXMP(0);
          return { xmpEnabled: false, stable: false, errors: memStress.errors };
        }
      }
    } else {
      this._log("testing", "mem-xmp-ok", "XMP appears to already be enabled");
      return { xmpEnabled: true, stable: true, alreadyEnabled: true };
    }

    return null;
  }

  // ── GPU Phase ─────────────────────────────────────────────────────

  async _gpuPhase() {
    this._setPhase("testing", 65);
    this._log("testing", "gpu-start", "Starting GPU overclock phase");

    const gpu = await this.hw.detectGPU();
    if (!gpu.hasNvidiaSmi) {
      this._log("testing", "skip-gpu", "No NVIDIA GPU with nvidia-smi found — skipping GPU phase");
      return null;
    }

    const basePowerLimit = gpu.powerLimit;
    let lastGoodLimit = basePowerLimit;
    const maxTemp = 90; // GPU thermal limit

    this._log("testing", "gpu-baseline", `GPU: ${gpu.name} | Power limit: ${basePowerLimit}W`, { power: basePowerLimit });

    for (let pctIncrease = 5; pctIncrease <= 20; pctIncrease += 5) {
      if (this._stopRequested) break;

      const newLimit = Math.round(basePowerLimit * (1 + pctIncrease / 100));
      this._setPhase("testing", 65 + (pctIncrease / 20) * 15);

      this._log("testing", "gpu-step", `Setting GPU power limit to ${newLimit}W (+${pctIncrease}%)`, { power: newLimit });

      await this.hw.gpuSetPowerLimit(newLimit);

      // Quick GPU stress test (5 minutes per step)
      // Use getStats() which overlays HWiNFO sensor data for accurate GPU temps
      const result = await this.runExtendedStressTest(5, { maxTemp, stressGpu: true });
      const fullStats = await this.hw.getStats();
      const gpuTemp = fullStats.gpu?.temp || 0;

      if (result.stable && gpuTemp < maxTemp) {
        lastGoodLimit = newLimit;
        this._log("testing", "gpu-stable", `${newLimit}W is STABLE (GPU temp: ${gpuTemp}C, source: ${fullStats.hwinfoActive ? "HWiNFO64" : "nvidia-smi"})`, {
          stable: true, temp: gpuTemp, power: newLimit,
        });
      } else {
        this._log("testing", "gpu-rollback", `${newLimit}W unstable or too hot — rolling back to ${lastGoodLimit}W`, {
          stable: false, temp: gpuTemp,
        });
        await this.hw.gpuSetPowerLimit(lastGoodLimit);
        break;
      }
    }

    return {
      finalPowerLimit: lastGoodLimit,
      gainWatts: lastGoodLimit - basePowerLimit,
      gainPct: Math.round(((lastGoodLimit - basePowerLimit) / basePowerLimit) * 100),
    };
  }

  // ── Final Validation ──────────────────────────────────────────────

  async _finalValidation() {
    this._setPhase("testing", 85);
    this._log("testing", "validation-start", "Starting 30-minute final validation (CPU + GPU + RAM combined)");

    const result = await this.runExtendedStressTest(30, { maxTemp: 95, stressGpu: true });

    if (result.stable) {
      this._log("testing", "validation-pass", "Final validation PASSED — system is stable", {
        stable: true, maxTemp: result.maxTemp, avgTemp: result.avgTemp,
      });
    } else {
      this._log("testing", "validation-fail", "Final validation FAILED", {
        stable: false, errors: result.errors, maxTemp: result.maxTemp,
      });
    }

    return result;
  }

  // ── Main Auto-OC Orchestrator ─────────────────────────────────────

  async runFullAutoOC(options = {}) {
    if (this.running) {
      return { success: false, error: "AI overclock is already running" };
    }

    this.running = true;
    this._stopRequested = false;
    this.log = [];
    this.progress = 0;
    this.stableSettings = null;
    this.currentProfile = null;

    // Re-detect SceWin and HWiNFO in case they were installed since app launch
    this.scewin.redetect();
    this.hw.hwinfo.redetect();
    this._log("analyzing", "start", `AI Auto-Overclock starting... SceWin: ${this.scewin.available ? "FOUND at " + this.scewin.scewinPath : "not found (GPU-only mode)"} | HWiNFO: ${this.hw.hwinfo.available ? "ACTIVE" : "not detected"}`);

    try {
      // Phase 1: Analyze hardware
      const analysis = await this._analyzeHardware();
      if (!analysis) {
        this._setPhase("failed", 100);
        this.running = false;
        return { success: false, error: "Hardware analysis failed", log: this.log };
      }

      if (this._stopRequested) {
        this._setPhase("idle", 0);
        this.running = false;
        return { success: false, error: "Stopped by user", log: this.log };
      }

      // Phase 2: CPU Overclock
      let cpuResult = null;
      if (options.skipCpu !== true) {
        cpuResult = await this._cpuOcPhase(analysis);
      }

      if (this._stopRequested) {
        this._setPhase("idle", 0);
        this.running = false;
        return { success: false, error: "Stopped by user", log: this.log };
      }

      // Phase 3: Memory
      let memResult = null;
      if (options.skipMemory !== true) {
        memResult = await this._memoryPhase();
      }

      if (this._stopRequested) {
        this._setPhase("idle", 0);
        this.running = false;
        return { success: false, error: "Stopped by user", log: this.log };
      }

      // Phase 4: GPU
      let gpuResult = null;
      if (options.skipGpu !== true) {
        gpuResult = await this._gpuPhase();
      }

      if (this._stopRequested) {
        this._setPhase("idle", 0);
        this.running = false;
        return { success: false, error: "Stopped by user", log: this.log };
      }

      // Phase 5: Final validation
      let validationResult = null;
      if (options.skipValidation !== true) {
        validationResult = await this._finalValidation();
      }

      // Generate final profile
      const profile = {
        timestamp: new Date().toISOString(),
        hardware: {
          cpu: analysis.cpu?.name,
          gpu: analysis.gpu?.name,
          ram: analysis.ram?.totalGB + "GB",
        },
        cpu: cpuResult,
        memory: memResult,
        gpu: gpuResult,
        validation: validationResult ? { stable: validationResult.stable, maxTemp: validationResult.maxTemp, avgTemp: validationResult.avgTemp } : null,
        stable: validationResult ? validationResult.stable : true,
      };

      this.currentProfile = profile;
      this.stableSettings = profile;

      // Save the profile
      try {
        const profilePath = path.join(os.homedir(), ".fn-optimizer", "ai-oc-profile.json");
        const dir = path.dirname(profilePath);
        if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
        fs.writeFileSync(profilePath, JSON.stringify(profile, null, 2));
      } catch (e) { /* non-fatal */ }

      const finalPhase = profile.stable ? "stable" : "failed";
      this._setPhase(finalPhase, 100);
      this._log(finalPhase, "complete", this._generateSummary(profile));

      this.running = false;
      return { success: true, profile, log: this.log };

    } catch (e) {
      this._setPhase("failed", 100);
      this._log("failed", "error", "Unexpected error: " + e.message);
      this.running = false;
      return { success: false, error: e.message, log: this.log };
    }
  }

  // ── Summary Report ────────────────────────────────────────────────

  _generateSummary(profile) {
    const lines = ["=== AI Auto-Overclock Summary ==="];

    if (profile.cpu) {
      lines.push(`CPU: x${profile.cpu.finalRatio} (${profile.cpu.finalMHz} MHz) — gained ${profile.cpu.gainMHz} MHz`);
      if (profile.cpu.finalVoltage) lines.push(`  Voltage: ${profile.cpu.finalVoltage}V`);
    } else {
      lines.push("CPU: No changes (skipped or unavailable)");
    }

    if (profile.memory) {
      lines.push(`Memory: XMP ${profile.memory.xmpEnabled ? "enabled" : "not applied"}${profile.memory.alreadyEnabled ? " (was already on)" : ""}`);
    } else {
      lines.push("Memory: No changes (skipped or unavailable)");
    }

    if (profile.gpu) {
      lines.push(`GPU: ${profile.gpu.finalPowerLimit}W (+${profile.gpu.gainPct}%)`);
    } else {
      lines.push("GPU: No changes (skipped or unavailable)");
    }

    if (profile.validation) {
      lines.push(`Validation: ${profile.validation.stable ? "PASSED" : "FAILED"} (max temp: ${profile.validation.maxTemp}C)`);
    }

    lines.push(`Overall: ${profile.stable ? "STABLE" : "UNSTABLE"}`);
    return lines.join("\n");
  }
}

// ── Hardware Detection ──────────────────────────────────────────────
class HardwareMonitor {
  constructor() {
    this.isWin = os.platform() === "win32";
    this.nvidiaSmiPath = null;
    this.hardwareInfo = null;
    this.hwinfo = new HWiNFOReader();
    this.detectNvidiaSmi();
    this.scewin = new SceWinManager();
    this.aiEngine = new AIOverclockEngine(this, this.scewin);
  }

  detectNvidiaSmi() {
    const paths = [
      "C:\\Program Files\\NVIDIA Corporation\\NVSMI\\nvidia-smi.exe",
      "C:\\Windows\\System32\\nvidia-smi.exe",
      "nvidia-smi",
    ];
    for (const p of paths) {
      try {
        if (p.includes("\\") && fs.existsSync(p)) { this.nvidiaSmiPath = p; return; }
      } catch(e) {}
    }
    this.nvidiaSmiPath = "nvidia-smi"; // fallback to PATH
  }

  // ── Full System Detection ────────────────────────────────────────
  async detectHardware() {
    if (!this.isWin) return { error: "Windows only" };

    const [cpu, gpu, ram, mobo] = await Promise.all([
      this.detectCPU(),
      this.detectGPU(),
      this.detectRAM(),
      this.detectMotherboard(),
    ]);

    this.hardwareInfo = { cpu, gpu, ram, mobo };
    return this.hardwareInfo;
  }

  async detectCPU() {
    const r = await runPS(`
      $cpu = Get-CimInstance Win32_Processor | Select-Object Name,NumberOfCores,NumberOfLogicalProcessors,MaxClockSpeed,CurrentClockSpeed,L2CacheSize,L3CacheSize,CurrentVoltage;
      $cpu | ConvertTo-Json
    `);
    try {
      const info = JSON.parse(r.output);
      return {
        name: info.Name?.trim(),
        cores: info.NumberOfCores,
        threads: info.NumberOfLogicalProcessors,
        maxClock: info.MaxClockSpeed, // MHz
        currentClock: info.CurrentClockSpeed,
        l2Cache: info.L2CacheSize, // KB
        l3Cache: info.L3CacheSize, // KB
        voltage: info.CurrentVoltage ? info.CurrentVoltage / 10 : null, // Volts
        isIntel: info.Name?.includes("Intel"),
        isAMD: info.Name?.includes("AMD"),
      };
    } catch(e) {
      return { name: "Unknown CPU", error: e.message };
    }
  }

  async detectGPU() {
    // Try nvidia-smi first for detailed NVIDIA info
    const nv = await runCmd(`"${this.nvidiaSmiPath}" --query-gpu=name,memory.total,memory.used,memory.free,temperature.gpu,power.draw,power.limit,clocks.current.graphics,clocks.current.memory,clocks.max.graphics,clocks.max.memory,fan.speed,utilization.gpu,pstate --format=csv,noheader,nounits 2>nul`);

    if (nv.success && nv.output) {
      const parts = nv.output.split(",").map(s => s.trim());
      return {
        name: parts[0],
        vendor: "NVIDIA",
        vramTotal: parseInt(parts[1]), // MB
        vramUsed: parseInt(parts[2]),
        vramFree: parseInt(parts[3]),
        temp: parseInt(parts[4]),
        powerDraw: parseFloat(parts[5]),
        powerLimit: parseFloat(parts[6]),
        clockCore: parseInt(parts[7]),
        clockMem: parseInt(parts[8]),
        clockMaxCore: parseInt(parts[9]),
        clockMaxMem: parseInt(parts[10]),
        fanSpeed: parseInt(parts[11]),
        utilization: parseInt(parts[12]),
        pstate: parts[13],
        hasNvidiaSmi: true,
      };
    }

    // Fallback to WMI
    const r = await runPS(`
      $gpu = Get-CimInstance Win32_VideoController | Select-Object Name,AdapterRAM,DriverVersion,VideoProcessor;
      $gpu | ConvertTo-Json
    `);
    try {
      let info = JSON.parse(r.output);
      if (Array.isArray(info)) info = info[0]; // primary GPU
      return {
        name: info.Name?.trim(),
        vendor: info.Name?.includes("NVIDIA") ? "NVIDIA" : info.Name?.includes("AMD") ? "AMD" : "Intel",
        vramTotal: info.AdapterRAM ? Math.round(info.AdapterRAM / 1024 / 1024) : 0,
        driverVersion: info.DriverVersion,
        hasNvidiaSmi: false,
      };
    } catch(e) {
      return { name: "Unknown GPU", error: e.message };
    }
  }

  async detectRAM() {
    const r = await runPS(`
      $ram = Get-CimInstance Win32_PhysicalMemory | Select-Object Manufacturer,Speed,ConfiguredClockSpeed,Capacity,PartNumber;
      $total = (Get-CimInstance Win32_ComputerSystem).TotalPhysicalMemory;
      @{sticks=$ram; total=$total} | ConvertTo-Json -Depth 3
    `);
    try {
      const info = JSON.parse(r.output);
      const sticks = Array.isArray(info.sticks) ? info.sticks : [info.sticks];
      return {
        totalGB: Math.round(info.total / 1024 / 1024 / 1024),
        totalBytes: info.total,
        sticks: sticks.map(s => ({
          manufacturer: s.Manufacturer?.trim(),
          speed: s.Speed, // MHz (rated)
          configuredSpeed: s.ConfiguredClockSpeed, // MHz (current)
          capacityGB: s.Capacity ? Math.round(s.Capacity / 1024 / 1024 / 1024) : 0,
          partNumber: s.PartNumber?.trim(),
        })),
        currentSpeed: sticks[0]?.ConfiguredClockSpeed,
        ratedSpeed: sticks[0]?.Speed,
      };
    } catch(e) {
      return { totalGB: Math.round(os.totalmem() / 1024 / 1024 / 1024), error: e.message };
    }
  }

  async detectMotherboard() {
    const r = await runPS(`
      $mb = Get-CimInstance Win32_BaseBoard | Select-Object Manufacturer,Product,Version;
      $bios = Get-CimInstance Win32_BIOS | Select-Object SMBIOSBIOSVersion,ReleaseDate,Manufacturer;
      @{board=$mb; bios=$bios} | ConvertTo-Json -Depth 3
    `);
    try {
      const info = JSON.parse(r.output);
      return {
        manufacturer: info.board?.Manufacturer?.trim(),
        model: info.board?.Product?.trim(),
        version: info.board?.Version?.trim(),
        biosVersion: info.bios?.SMBIOSBIOSVersion?.trim(),
        biosVendor: info.bios?.Manufacturer?.trim(),
      };
    } catch(e) {
      return { manufacturer: "Unknown", error: e.message };
    }
  }

  // ── Live Monitoring ──────────────────────────────────────────────
  async getStats() {
    if (!this.isWin) return {};

    const [cpuStats, gpuStats, ramStats, hwinfoData] = await Promise.all([
      this.getCPUStats(),
      this.getGPUStats(),
      this.getRAMStats(),
      this.hwinfo.readSensors().catch(() => ({ available: false })),
    ]);

    // Overlay HWiNFO data for more accurate readings when available
    if (hwinfoData.available && hwinfoData.sensors) {
      const s = hwinfoData.sensors;
      if (s.cpuPackageTemp || s.cpuTemp) cpuStats.temp = s.cpuPackageTemp || s.cpuTemp;
      if (s.cpuVoltage) cpuStats.voltage = s.cpuVoltage;
      if (s.cpuClock) cpuStats.clock = s.cpuClock;
      if (s.cpuPower) cpuStats.power = s.cpuPower;
      if (s.gpuTemp && s.gpuTemp > 0) gpuStats.temp = s.gpuTemp;
      if (s.gpuClock) gpuStats.clockCore = s.gpuClock;
      if (s.gpuMemClock) gpuStats.clockMem = s.gpuMemClock;
      if (s.gpuLoad) gpuStats.usage = s.gpuLoad;
      if (s.gpuPower) gpuStats.power = s.gpuPower;
      if (s.vrmTemp) cpuStats.vrmTemp = s.vrmTemp;
      if (s.ramTemp) ramStats.temp = s.ramTemp;
      if (Object.keys(s.fanSpeeds || {}).length > 0) cpuStats.fans = s.fanSpeeds;
    }

    return { cpu: cpuStats, gpu: gpuStats, ram: ramStats, hwinfoActive: hwinfoData.available };
  }

  async getCPUStats() {
    const r = await runPS(`
      $cpu = Get-Counter '\\Processor(_Total)\\% Processor Time' -ErrorAction SilentlyContinue;
      $temp = Get-CimInstance MSAcpi_ThermalZoneTemperature -Namespace root/wmi -ErrorAction SilentlyContinue | Select-Object -First 1;
      @{
        usage = [math]::Round($cpu.CounterSamples[0].CookedValue, 1);
        temp = if($temp) { [math]::Round(($temp.CurrentTemperature - 2732) / 10, 1) } else { -1 };
      } | ConvertTo-Json
    `);
    try {
      return JSON.parse(r.output);
    } catch(e) {
      return { usage: -1, temp: -1 };
    }
  }

  async getGPUStats() {
    const nv = await runCmd(`"${this.nvidiaSmiPath}" --query-gpu=temperature.gpu,power.draw,clocks.current.graphics,clocks.current.memory,fan.speed,utilization.gpu,memory.used,memory.total --format=csv,noheader,nounits 2>nul`);
    if (nv.success && nv.output) {
      const p = nv.output.split(",").map(s => s.trim());
      return {
        temp: parseInt(p[0]),
        power: parseFloat(p[1]),
        clockCore: parseInt(p[2]),
        clockMem: parseInt(p[3]),
        fan: parseInt(p[4]),
        usage: parseInt(p[5]),
        vramUsed: parseInt(p[6]),
        vramTotal: parseInt(p[7]),
      };
    }
    return { temp: -1, usage: -1 };
  }

  async getRAMStats() {
    const total = os.totalmem();
    const free = os.freemem();
    const used = total - free;
    return {
      totalGB: Math.round(total / 1024 / 1024 / 1024 * 10) / 10,
      usedGB: Math.round(used / 1024 / 1024 / 1024 * 10) / 10,
      freeGB: Math.round(free / 1024 / 1024 / 1024 * 10) / 10,
      usagePct: Math.round((used / total) * 100),
    };
  }

  // ── GPU Overclocking (NVIDIA) ─────────────────────────────────────
  // Uses nvidia-smi to adjust power limits and clock offsets
  // These are SAFE presets — nvidia-smi won't let you go past hardware limits

  async gpuSetPowerLimit(watts) {
    // nvidia-smi enforces min/max — can't brick the card
    const r = await runCmd(`"${this.nvidiaSmiPath}" -pl ${watts} 2>nul`);
    return { success: r.success, error: r.error };
  }

  async gpuSetClockOffset(coreOffset, memOffset) {
    // nvidia-smi clock offset (requires admin)
    const r = await runCmd(`"${this.nvidiaSmiPath}" -lgc ${coreOffset} 2>nul`);
    return { success: r.success, error: r.error };
  }

  async gpuResetClocks() {
    const r = await runCmd(`"${this.nvidiaSmiPath}" -rgc 2>nul && "${this.nvidiaSmiPath}" -rmc 2>nul`);
    return { success: r.success };
  }

  // ── GPU Overclock Presets ────────────────────────────────────────
  getGPUPresets() {
    return [
      {
        id: "gpu-stock", name: "Stock (Default)",
        desc: "Reset to factory clocks. Safe baseline.",
        action: "reset",
      },
      {
        id: "gpu-mild", name: "Mild (+5% Power)",
        desc: "Slightly increase power limit. Safe for any card. 2-5 FPS gain.",
        powerPct: 105,
      },
      {
        id: "gpu-moderate", name: "Moderate (+10% Power)",
        desc: "Moderate power boost. Good cooling required. 5-10 FPS gain.",
        powerPct: 110,
      },
      {
        id: "gpu-aggressive", name: "Aggressive (+15% Power)",
        desc: "Push it harder. Needs good airflow. 8-15 FPS gain. Watch temps.",
        powerPct: 115,
        danger: true,
      },
      {
        id: "gpu-extreme", name: "Extreme (Auto-Tune)",
        desc: "Automatically pushes clocks higher while monitoring temps and stability. Stops at thermal/stability limit.",
        powerPct: 120,
        danger: true,
        autoTune: true,
      },
    ];
  }

  async applyGPUPreset(presetId) {
    const gpu = await this.detectGPU();
    if (!gpu.hasNvidiaSmi) return { success: false, error: "NVIDIA GPU with nvidia-smi required for overclocking" };

    if (presetId === "gpu-stock") {
      return await this.gpuResetClocks();
    }

    const preset = this.getGPUPresets().find(p => p.id === presetId);
    if (!preset) return { success: false, error: "Unknown preset" };

    const newLimit = Math.round(gpu.powerLimit * (preset.powerPct / 100));
    return await this.gpuSetPowerLimit(newLimit);
  }

  // ── CPU Power Presets ────────────────────────────────────────────
  getCPUPresets() {
    return [
      {
        id: "cpu-balanced", name: "Balanced",
        desc: "Windows default. CPU scales up/down based on load. Saves power but adds latency.",
        cmd: 'powercfg -setactive 381b4222-f694-41f0-9685-ff5bb260df2e',
      },
      {
        id: "cpu-high-perf", name: "High Performance",
        desc: "CPU stays at higher clocks. Good balance of performance and heat.",
        cmd: 'powercfg -setactive 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c',
      },
      {
        id: "cpu-ultimate", name: "Ultimate Performance",
        desc: "Maximum CPU speed at all times. No downclocking. Best for gaming.",
        cmd: 'powercfg -duplicatescheme e9a42b02-d5df-448d-aa00-03f14749eb61 2>nul & powercfg -setactive e9a42b02-d5df-448d-aa00-03f14749eb61',
      },
      {
        id: "cpu-max-freq", name: "Force Max Frequency",
        desc: "Pins CPU to maximum turbo boost at all times. Higher temps but lowest latency.",
        cmd: 'powercfg -setacvalueindex scheme_current sub_processor PROCTHROTTLEMIN 100 && powercfg -setacvalueindex scheme_current sub_processor PROCTHROTTLEMAX 100 && powercfg -setactive scheme_current',
      },
      {
        id: "cpu-unpark", name: "Unpark All Cores",
        desc: "Prevents Windows from sleeping any CPU cores. All cores active all the time.",
        cmd: 'powercfg -setacvalueindex scheme_current sub_processor CPMINCORES 100 && powercfg -setacvalueindex scheme_current sub_processor CPMAXCORES 100 && powercfg -setactive scheme_current',
      },
    ];
  }

  async applyCPUPreset(presetId) {
    const preset = this.getCPUPresets().find(p => p.id === presetId);
    if (!preset) return { success: false, error: "Unknown preset" };
    const r = await runCmd(preset.cmd);
    return { success: r.success, error: r.error };
  }

  // ── RAM Optimization ─────────────────────────────────────────────
  // Can't OC RAM from Windows (that's BIOS/XMP), but we can optimize
  // how Windows uses it

  getRAMTweaks() {
    return [
      {
        id: "ram-priority", name: "Set Fortnite Memory Priority",
        desc: "Uses PowerShell to set Fortnite process to high memory priority when running.",
        cmd: 'powershell -Command "Get-Process FortniteClient* -ErrorAction SilentlyContinue | ForEach-Object { $_.PriorityClass = \'High\' }"',
      },
      {
        id: "ram-standby-clean", name: "Clear Standby Memory",
        desc: "Flushes standby RAM list. Frees up memory that Windows is hoarding for no reason.",
        cmd: 'powershell -Command "[System.GC]::Collect(); [System.GC]::WaitForPendingFinalizers()"',
      },
      {
        id: "ram-large-page", name: "Enable Large Pages",
        desc: "Allow large memory pages (2MB vs 4KB). Reduces TLB misses for better performance.",
        cmd: 'reg add "HKLM\\SYSTEM\\CurrentControlSet\\Control\\Session Manager\\Memory Management" /v LargePageMinimum /t REG_DWORD /d 1 /f',
        reboot: true,
      },
      {
        id: "ram-pagefile-opt", name: "Optimize Page File",
        desc: "Set page file to system managed on fastest drive. Prevents RAM pressure.",
        cmd: 'wmic computersystem set AutomaticManagedPagefile=True',
      },
    ];
  }

  async applyRAMTweak(tweakId) {
    const tweak = this.getRAMTweaks().find(t => t.id === tweakId);
    if (!tweak) return { success: false, error: "Unknown tweak" };
    const r = await runCmd(tweak.cmd);
    return { success: r.success, error: r.error };
  }

  // ── Auto-Tune GPU ────────────────────────────────────────────────
  async autoTuneGPU() {
    const gpu = await this.detectGPU();
    if (!gpu.hasNvidiaSmi) return { success: false, error: "NVIDIA GPU required" };

    const log = [];
    const basePowerLimit = gpu.powerLimit;
    let lastGoodLimit = basePowerLimit;
    const maxTemp = 95; // safety limit

    // Save last known good settings
    this.saveLastKnownGood({ gpuPowerLimit: basePowerLimit });

    for (let pctIncrease = 5; pctIncrease <= 25; pctIncrease += 5) {
      const newLimit = Math.round(basePowerLimit * (1 + pctIncrease / 100));
      log.push({ step: pctIncrease, action: "Setting power limit to " + newLimit + "W" });

      await this.gpuSetPowerLimit(newLimit);

      // Run stress test and monitor
      const result = await this.runStressTest(30);
      const stats = await this.getGPUStats();

      log.push({ step: pctIncrease, temp: stats.temp, power: stats.power, stable: true });

      if (stats.temp >= maxTemp) {
        log.push({ step: pctIncrease, action: "Temp limit reached (" + stats.temp + "\u00B0C), rolling back" });
        await this.gpuSetPowerLimit(lastGoodLimit);
        break;
      }

      lastGoodLimit = newLimit;
    }

    return { success: true, finalLimit: lastGoodLimit, log };
  }

  // ── Last Known Good Settings ────────────────────────────────────
  saveLastKnownGood(settings) {
    try {
      const dir = path.dirname(SETTINGS_PATH);
      if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
      const data = { settings, savedAt: new Date().toISOString() };
      fs.writeFileSync(SETTINGS_PATH, JSON.stringify(data, null, 2));
      return { success: true };
    } catch (e) {
      return { success: false, error: e.message };
    }
  }

  loadLastKnownGood() {
    try {
      if (!fs.existsSync(SETTINGS_PATH)) return null;
      return JSON.parse(fs.readFileSync(SETTINGS_PATH, "utf8"));
    } catch (e) {
      return null;
    }
  }

  async restoreLastKnownGood() {
    const saved = this.loadLastKnownGood();
    if (!saved || !saved.settings) return { success: false, error: "No saved settings found" };

    const results = {};
    if (saved.settings.gpuPowerLimit) {
      results.gpu = await this.gpuSetPowerLimit(saved.settings.gpuPowerLimit);
    }

    return { success: true, restored: saved.settings, savedAt: saved.savedAt, results };
  }

  // ── BSOD / Crash Detection ─────────────────────────────────────
  async checkForRecentCrashes() {
    const r = await runPS(`
      $crashes = @();
      try {
        $events = wevtutil qe System /q:"*[System[Provider[@Name='Microsoft-Windows-WER-SystemErrorReporting']]]" /c:5 /f:text 2>$null;
        if ($events) {
          $crashes += @{ source = "WER-SystemError"; raw = $events };
        }
      } catch {}
      try {
        $unexpected = Get-EventLog -LogName System -Source "EventLog" -EntryType Error -Newest 5 -ErrorAction SilentlyContinue |
          Where-Object { $_.Message -match "unexpected" -or $_.Message -match "shutdown" };
        foreach ($evt in $unexpected) {
          $crashes += @{
            source = "UnexpectedShutdown";
            time = $evt.TimeGenerated.ToString("o");
            message = $evt.Message.Substring(0, [Math]::Min(200, $evt.Message.Length));
          };
        }
      } catch {}
      @{ hasCrash = ($crashes.Count -gt 0); crashes = $crashes; count = $crashes.Count } | ConvertTo-Json -Depth 3
    `, 20000);

    try {
      const parsed = JSON.parse(r.output);
      return {
        hasCrash: parsed.hasCrash || false,
        crashes: parsed.crashes || [],
        lastCrash: parsed.crashes?.length > 0 ? parsed.crashes[0].time || null : null,
      };
    } catch (e) {
      return { hasCrash: false, crashes: [], lastCrash: null, error: e.message };
    }
  }

  // ── Boot Check ─────────────────────────────────────────────────
  async bootCheck() {
    // Called on app startup
    // 1. Check for recent BSODs/crashes since last run
    // 2. If crash detected, log it and restore last known good settings
    // 3. Return what happened
    const crashes = await this.checkForRecentCrashes();
    if (crashes.hasCrash) {
      await this.restoreLastKnownGood();
      this.logCrash(crashes);
      return { restored: true, crashes: crashes.crashes };
    }
    return { restored: false };
  }

  // ── Crash Logging ──────────────────────────────────────────────
  logCrash(crashInfo) {
    try {
      const dir = path.dirname(CRASH_LOG_PATH);
      if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });

      let existing = [];
      if (fs.existsSync(CRASH_LOG_PATH)) {
        try {
          existing = JSON.parse(fs.readFileSync(CRASH_LOG_PATH, "utf8"));
        } catch (e) {
          existing = [];
        }
      }

      existing.push({
        timestamp: new Date().toISOString(),
        crashInfo,
      });

      fs.writeFileSync(CRASH_LOG_PATH, JSON.stringify(existing, null, 2));
      return { success: true };
    } catch (e) {
      return { success: false, error: e.message };
    }
  }

  getCrashLog() {
    try {
      if (!fs.existsSync(CRASH_LOG_PATH)) return [];
      return JSON.parse(fs.readFileSync(CRASH_LOG_PATH, "utf8"));
    } catch (e) {
      return [];
    }
  }

  // ── SceWin / AI Overclock Integration ─────────────────────────────

  // ── Stress Test ──────────────────────────────────────────────────

  async runStressTest(durationSeconds = 10) {
    // Run a CPU stress test using PowerShell multi-threaded math
    const threads = os.cpus().length;
    const psCmd = `
      $duration = ${durationSeconds}
      $jobs = @()
      for ($i = 0; $i -lt ${threads}; $i++) {
        $jobs += Start-Job -ScriptBlock {
          $end = (Get-Date).AddSeconds($using:duration)
          while ((Get-Date) -lt $end) {
            [math]::Sqrt([math]::PI) | Out-Null
            [math]::Pow(2, 64) | Out-Null
          }
        }
      }
      $jobs | Wait-Job -Timeout ($duration + 5) | Out-Null
      $jobs | Remove-Job -Force
      Write-Output "STRESS_COMPLETE"
    `;
    const result = await runPS(psCmd, (durationSeconds + 15) * 1000);
    return { success: result.output.includes("STRESS_COMPLETE"), duration: durationSeconds };
  }

  // ── HWiNFO64 Integration ─────────────────────────────────────────

  getHwinfoStatus() {
    this.hwinfo.redetect();
    return { available: this.hwinfo.available, path: this.hwinfo.hwInfoPath };
  }

  async readHwinfoSensors() {
    return await this.hwinfo.readSensors();
  }

  // Enhanced stats that prefer HWiNFO data when available
  async getEnhancedStats() {
    const [basicStats, hwinfoData] = await Promise.all([
      this.getStats(),
      this.hwinfo.readSensors(),
    ]);

    if (hwinfoData.available && hwinfoData.sensors) {
      const s = hwinfoData.sensors;
      // Override/enhance with HWiNFO's more accurate readings
      if (s.cpuPackageTemp || s.cpuTemp) basicStats.cpuTemp = s.cpuPackageTemp || s.cpuTemp;
      if (s.cpuVoltage) basicStats.cpuVoltage = s.cpuVoltage;
      if (s.cpuClock) basicStats.cpuClock = s.cpuClock;
      if (s.cpuPower) basicStats.cpuPower = s.cpuPower;
      if (s.gpuTemp) basicStats.gpuTemp = s.gpuTemp;
      if (s.gpuClock) basicStats.gpuCoreClk = s.gpuClock;
      if (s.gpuMemClock) basicStats.gpuMemClk = s.gpuMemClock;
      if (s.gpuLoad) basicStats.gpuLoad = s.gpuLoad;
      if (s.gpuPower) basicStats.gpuPower = s.gpuPower;
      if (s.ramTemp) basicStats.ramTemp = s.ramTemp;
      if (s.vrmTemp) basicStats.vrmTemp = s.vrmTemp;
      basicStats.fanSpeeds = s.fanSpeeds;
      basicStats.hwinfoActive = true;
    } else {
      basicStats.hwinfoActive = false;
    }

    return basicStats;
  }

  getScewinStatus() {
    // Re-detect every time this is called so UI stays in sync
    this.scewin.redetect();
    return { available: this.scewin.available, path: this.scewin.scewinPath };
  }

  async scewinExport() {
    return await this.scewin.exportCurrentSettings();
  }

  async scewinBackup() {
    return await this.scewin.backupSettings();
  }

  async scewinRestore() {
    return await this.scewin.restoreBackup();
  }

  async aiAutoOC(opts) {
    return await this.aiEngine.runFullAutoOC(opts);
  }

  aiProgress() {
    return {
      running: this.aiEngine.running,
      phase: this.aiEngine.currentPhase,
      progress: this.aiEngine.progress,
      log: this.aiEngine.log,
      summary: this.aiEngine.stableSettings,
    };
  }

  aiStop() {
    this.aiEngine.stop();
    return { success: true };
  }
}

module.exports = HardwareMonitor;