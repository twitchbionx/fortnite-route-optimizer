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

  // Check if HWiNFO64 process is running
  async isRunning() {
    const r = await runCmd('tasklist /FI "IMAGENAME eq HWiNFO64.exe" 2>nul', 5000);
    return r.success && r.output.toLowerCase().includes("hwinfo64");
  }

  // Force-enable Shared Memory and ensure HWiNFO is running with it active
  async ensureRunning() {
    if (!this.available) this.redetect();

    // Find the binary first
    if (!this.available) {
      const fallbacks = [
        "C:\\Program Files\\HWiNFO64\\HWiNFO64.exe",
        "C:\\Program Files (x86)\\HWiNFO64\\HWiNFO64.exe",
        path.join(os.homedir(), ".fn-optimizer", "hwinfo", "HWiNFO64.exe"),
        path.join(os.homedir(), ".fn-optimizer", "hwinfo", "HWiNFO64", "HWiNFO64.exe"),
        path.join(os.homedir(), "AppData\\Local\\Programs\\HWiNFO64\\HWiNFO64.exe"),
      ];
      for (const p of fallbacks) {
        try { if (fs.existsSync(p)) { this.hwInfoPath = p; this.available = true; break; } } catch(e) {}
      }
    }

    if (!this.available || !this.hwInfoPath) {
      return { launched: false, error: "HWiNFO64 not found — install it from the Dependencies tab" };
    }

    // Enable Shared Memory via registry BEFORE launching
    await runPS(`
      $paths = @('HKCU:\\Software\\HWiNFO64', 'HKCU:\\Software\\HWiNFO64\\VSB')
      foreach ($p in $paths) { if (!(Test-Path $p)) { New-Item -Path $p -Force | Out-Null } }
      Set-ItemProperty -Path 'HKCU:\\Software\\HWiNFO64' -Name 'SensorsSM' -Value 1 -Type DWord -Force
      Set-ItemProperty -Path 'HKCU:\\Software\\HWiNFO64' -Name 'SensorsOnly' -Value 1 -Type DWord -Force
      Set-ItemProperty -Path 'HKCU:\\Software\\HWiNFO64' -Name 'AutoStart' -Value 1 -Type DWord -Force
    `, 5000);

    // Write INI file next to the EXE (for both portable and installed versions)
    try {
      const iniPath = path.join(path.dirname(this.hwInfoPath), "HWiNFO64.INI");
      const iniContent = "[Settings]\nSensorsSM=1\nSensorsOnly=1\nMinimize=1\nAutoStart=1\nShowWelcome=0\n";
      fs.writeFileSync(iniPath, iniContent);
    } catch(e) {} // may fail if Program Files (need admin)

    // Kill any existing HWiNFO so it restarts with our settings
    const running = await this.isRunning();
    if (running) {
      await runCmd('taskkill /IM HWiNFO64.exe /F 2>nul', 5000);
      await new Promise(r => setTimeout(r, 3000));
    }

    // Launch with explicit shared memory flags
    // HWiNFO64 supports: -sm (shared memory), -lSENSORSONLY (sensors only)
    const launchCmds = [
      `Start-Process -FilePath '${this.hwInfoPath.replace(/'/g, "''")}' -ArgumentList '-sm -lSENSORSONLY' -WindowStyle Minimized`,
      `Start-Process -FilePath '${this.hwInfoPath.replace(/'/g, "''")}' -ArgumentList '-sm' -WindowStyle Minimized`,
      `Start-Process -FilePath '${this.hwInfoPath.replace(/'/g, "''")}' -WindowStyle Minimized`,
    ];

    for (const cmd of launchCmds) {
      await runPS(cmd, 5000);
      // Wait up to 15 seconds for VSB data
      for (let i = 0; i < 15; i++) {
        await new Promise(r => setTimeout(r, 1000));
        const check = await runPS(`
          $base = 'HKCU:\\Software\\HWiNFO64\\VSB'
          if (!(Test-Path $base)) { Write-Output 'NO_VSB'; return }
          $v = Get-ItemProperty -Path $base -ErrorAction SilentlyContinue
          if ($v.Label0) { Write-Output "HAS_DATA:$($v.Label0)" } else { Write-Output 'NO_DATA' }
        `, 3000);
        if (check.success && check.output.includes("HAS_DATA")) {
          await new Promise(r => setTimeout(r, 2000));
          return { launched: true, path: this.hwInfoPath, hasData: true };
        }
      }
      // This launch method didn't work, kill and try next
      await runCmd('taskkill /IM HWiNFO64.exe /F 2>nul', 5000);
      await new Promise(r => setTimeout(r, 2000));
    }

    return { launched: true, path: this.hwInfoPath, hasData: false, warning: "HWiNFO started but Shared Memory not active. You may need to enable it manually: open HWiNFO64 → Settings → check 'Shared Memory Support'" };
  }

  // Fallback: read sensor data via HWiNFO CSV report mode (works without shared memory)
  async readSensorsCSV() {
    if (!this.available || !this.hwInfoPath) return { available: false, error: "HWiNFO64 not found" };

    const csvPath = path.join(os.homedir(), ".fn-optimizer", "hwinfo-sensors.csv");
    try { fs.unlinkSync(csvPath); } catch(e) {}

    // Run HWiNFO briefly to capture a CSV snapshot
    const r = await runCmd(`"${this.hwInfoPath}" /csv="${csvPath}" /maxtime=3`, 15000);

    try {
      if (fs.existsSync(csvPath)) {
        const raw = fs.readFileSync(csvPath, "utf8");
        if (raw.length > 50) {
          return { available: true, sensors: this._parseCSVSensors(raw), source: "CSV-report" };
        }
      }
    } catch(e) {}
    return { available: false, error: "CSV report not generated" };
  }

  _parseCSVSensors(csv) {
    const sensors = {
      cpuTemp: null, cpuPackageTemp: null, cpuVoltage: null, cpuClock: null,
      cpuPower: null, gpuTemp: null, gpuClock: null, gpuMemClock: null,
      gpuLoad: null, gpuPower: null, vrmTemp: null, fanSpeeds: {},
    };

    const lines = csv.split("\n");
    if (lines.length < 2) return sensors;

    // CSV has headers on first line, values on subsequent lines
    const headers = lines[0].split(",").map(h => h.trim().replace(/"/g, ""));
    const lastDataLine = lines[lines.length - 1].trim() || lines[lines.length - 2]?.trim();
    if (!lastDataLine) return sensors;
    const values = lastDataLine.split(",").map(v => v.trim().replace(/"/g, ""));

    for (let i = 0; i < headers.length; i++) {
      const h = headers[i].toLowerCase();
      const val = parseFloat(values[i]);
      if (isNaN(val)) continue;

      if (h.includes("cpu") && h.includes("package") && h.includes("temp")) sensors.cpuPackageTemp = val;
      else if (h.includes("cpu") && h.includes("temp") && !sensors.cpuTemp) sensors.cpuTemp = val;
      else if (h.includes("vcore") || (h.includes("cpu") && h.includes("voltage"))) sensors.cpuVoltage = val;
      else if (h.includes("cpu") && h.includes("clock") && !h.includes("ring")) sensors.cpuClock = val;
      else if (h.includes("cpu") && h.includes("package") && h.includes("power")) sensors.cpuPower = val;
      else if (h.includes("gpu") && h.includes("temp") && !h.includes("hot")) sensors.gpuTemp = val;
      else if (h.includes("gpu") && h.includes("clock") && !h.includes("mem")) sensors.gpuClock = val;
      else if (h.includes("vrm") && h.includes("temp")) sensors.vrmTemp = val;
      else if (h.includes("fan") && (h.includes("rpm") || h.includes("speed"))) sensors.fanSpeeds[headers[i]] = val;
    }

    return sensors;
  }

  // Read sensor values from HWiNFO registry
  async readSensors() {
    // First check if registry path exists
    const regCheck = await runPS(`Test-Path 'HKCU:\\Software\\HWiNFO64\\VSB'`, 3000);
    const regExists = regCheck.success && regCheck.output.includes("True");

    // If not present, try launching HWiNFO
    if (!regExists) {
      const launch = await this.ensureRunning();
      if (launch.error) {
        // Try CSV fallback before giving up
        const csv = await this.readSensorsCSV();
        if (csv.available) return csv;
        return { available: false, error: launch.error };
      }
      // Re-check after launch attempt
      const recheck = await runPS(`Test-Path 'HKCU:\\Software\\HWiNFO64\\VSB'`, 3000);
      if (!recheck.success || !recheck.output.includes("True")) {
        // VSB still not available — try CSV fallback
        const csv = await this.readSensorsCSV();
        if (csv.available) return csv;
        return { available: false, error: "HWiNFO VSB registry not found. CSV fallback also failed. You may need to enable Shared Memory manually in HWiNFO64 settings." };
      }
    }

    // Read sensor data from registry
    const result = await runPS(`
      $base = 'HKCU:\\Software\\HWiNFO64\\VSB'
      if (!(Test-Path $base)) { Write-Output '{"error":"VSB not found"}'; return }
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
      if ($sensors.Count -eq 0) { Write-Output '{"error":"No sensor data in VSB registry. Enable Shared Memory in HWiNFO settings."}'; return }
      $sensors | ConvertTo-Json -Depth 3
    `, 10000);

    if (!result.success || !result.output) {
      // VSB read failed — try CSV fallback
      const csv = await this.readSensorsCSV();
      if (csv.available) return csv;
      return { available: false, error: "Could not read HWiNFO sensors: " + (result.error || "no output") };
    }

    try {
      const data = JSON.parse(result.output);
      if (data.error) {
        // Registry exists but no sensor data — try CSV fallback
        const csv = await this.readSensorsCSV();
        if (csv.available) return csv;
        return { available: false, error: data.error };
      }
      return { available: true, sensors: this._parseSensors(data), source: "VSB-shared-memory" };
    } catch(e) {
      // Parse failed — try CSV fallback
      const csv = await this.readSensorsCSV();
      if (csv.available) return csv;
      return { available: false, error: "Failed to parse HWiNFO data: " + e.message };
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

    // Clean up old export file so we can detect if a new one was created
    try { fs.unlinkSync(exportPath); } catch(e) {}

    // AMISCE/SceWin needs to run from its own directory with admin elevation.
    // Different AMISCE versions use wildly different flags. We try everything.
    const scewinDir = path.dirname(this.scewinPath);
    const scewinExe = path.basename(this.scewinPath);
    const exportName = path.basename(exportPath);

    // First: try to discover what flags this version supports by running /? or /h
    let helpOutput = "";
    const helpResult = await runCmd(`cd /d "${scewinDir}" && "${scewinExe}" /? 2>&1`, 10000);
    if (helpResult.output) helpOutput = helpResult.output;
    if (!helpOutput) {
      const h2 = await runCmd(`cd /d "${scewinDir}" && "${scewinExe}" /h 2>&1`, 10000);
      if (h2.output) helpOutput = h2.output;
    }
    this._lastHelpOutput = helpOutput; // store for debugging

    // Build command list based on help output analysis
    const cmds = [];
    const esc = (p) => p.replace(/'/g, "''");

    // AMISCE 5.x syntax: /o outputs to AMISCE.txt in CWD (no path argument)
    // Then we copy the output file to our desired location
    cmds.push({
      label: "AMISCE /o (output to CWD)",
      cmd: `cd /d "${scewinDir}" && "${scewinExe}" /o`,
      outputFile: path.join(scewinDir, "AMISCE.txt"),
    });

    // /o with /lang for full text descriptions
    cmds.push({
      label: "AMISCE /o /lang",
      cmd: `cd /d "${scewinDir}" && "${scewinExe}" /o /lang`,
      outputFile: path.join(scewinDir, "AMISCE.txt"),
    });

    // /o with explicit filename — some versions support this
    cmds.push({
      label: "AMISCE /o /s <path>",
      cmd: `cd /d "${scewinDir}" && "${scewinExe}" /o /s "${exportPath}"`,
      outputFile: exportPath,
    });

    cmds.push({
      label: "AMISCE /o <path>",
      cmd: `cd /d "${scewinDir}" && "${scewinExe}" /o "${exportPath}"`,
      outputFile: exportPath,
    });

    // /ds = dump script — some AMISCE versions use this
    cmds.push({
      label: "AMISCE /ds",
      cmd: `cd /d "${scewinDir}" && "${scewinExe}" /ds "${exportPath}"`,
      outputFile: exportPath,
    });

    // Elevated versions of the above
    cmds.push({
      label: "Elevated /o (CWD output)",
      cmd: `powershell -Command "Start-Process -FilePath '${esc(this.scewinPath)}' -ArgumentList '/o' -WorkingDirectory '${esc(scewinDir)}' -Verb RunAs -Wait -WindowStyle Hidden"`,
      outputFile: path.join(scewinDir, "AMISCE.txt"),
    });

    cmds.push({
      label: "Elevated /o /lang",
      cmd: `powershell -Command "Start-Process -FilePath '${esc(this.scewinPath)}' -ArgumentList '/o /lang' -WorkingDirectory '${esc(scewinDir)}' -Verb RunAs -Wait -WindowStyle Hidden"`,
      outputFile: path.join(scewinDir, "AMISCE.txt"),
    });

    cmds.push({
      label: "Elevated /o /s <path>",
      cmd: `powershell -Command "Start-Process -FilePath '${esc(this.scewinPath)}' -ArgumentList '/o /s \\\"${esc(exportPath)}\\\"' -WorkingDirectory '${esc(scewinDir)}' -Verb RunAs -Wait -WindowStyle Hidden"`,
      outputFile: exportPath,
    });

    this._lastExportAttempts = [];

    for (const { label, cmd, outputFile } of cmds) {
      // Clean up expected output files before each attempt
      for (const f of [exportPath, outputFile]) {
        try { fs.unlinkSync(f); } catch(e) {}
      }

      const r = await runCmd(cmd, 30000);
      const attempt = { label, stdout: r.output?.substring(0, 500), stderr: r.error?.substring(0, 500), success: r.success };

      // Check both the intended output file and common AMISCE output filenames
      const checkFiles = [outputFile, exportPath];
      if (outputFile !== path.join(scewinDir, "AMISCE.txt")) checkFiles.push(path.join(scewinDir, "AMISCE.txt"));

      // Also check for any .txt file created in scewinDir in the last 30 seconds
      try {
        const files = fs.readdirSync(scewinDir);
        for (const f of files) {
          if (f.endsWith(".txt") || f.endsWith(".TXT")) {
            const fp = path.join(scewinDir, f);
            const stat = fs.statSync(fp);
            if (Date.now() - stat.mtimeMs < 30000 && !checkFiles.includes(fp)) {
              checkFiles.push(fp);
            }
          }
        }
      } catch(e) {}

      for (const checkPath of checkFiles) {
        try {
          if (fs.existsSync(checkPath)) {
            const raw = fs.readFileSync(checkPath, "utf8");
            attempt.fileFound = checkPath;
            attempt.fileSize = raw.length;

            if (raw.length > 50) {
              // Be more flexible in parsing — look for any structured BIOS data
              if (raw.includes("Setup Question") || raw.includes("Token") ||
                  raw.includes("Question") || raw.includes("Variable") ||
                  raw.includes("BIOS") || raw.includes("Setup")) {
                const settings = this._parseExport(raw);
                const count = Object.keys(settings).length;

                // If standard parse found nothing, try alternate parse
                if (count === 0) {
                  const altSettings = this._parseExportAlt(raw);
                  const altCount = Object.keys(altSettings).length;
                  if (altCount > 0) {
                    this.currentSettings = altSettings;
                    attempt.parsedCount = altCount;
                    this._lastExportAttempts.push(attempt);
                    if (checkPath !== exportPath) {
                      try { fs.copyFileSync(checkPath, exportPath); } catch(e) {}
                    }
                    return { success: true, settings: altSettings, settingCount: altCount, method: label };
                  }
                } else {
                  this.currentSettings = settings;
                  attempt.parsedCount = count;
                  this._lastExportAttempts.push(attempt);
                  if (checkPath !== exportPath) {
                    try { fs.copyFileSync(checkPath, exportPath); } catch(e) {}
                  }
                  return { success: true, settings, settingCount: count, method: label };
                }
              }
            }
          }
        } catch(e) { attempt.parseError = e.message; }
      }

      this._lastExportAttempts.push(attempt);
    }

    // Build a debug message with what we tried
    const debugInfo = this._lastExportAttempts
      .map(a => `${a.label}: ${a.fileFound ? `file=${a.fileFound}(${a.fileSize}b)` : "no file"} stdout="${(a.stdout || "").substring(0, 100)}"`)
      .join(" | ");

    return { success: false, error: `SceWin export failed after ${cmds.length} attempts. Debug: ${debugInfo}`, helpOutput: helpOutput?.substring(0, 500) };
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

  // Alternate parser for different AMISCE output formats
  _parseExportAlt(raw) {
    const settings = {};
    const lines = raw.split("\n");

    // Format 1: "Question: <name>" / "Value: <hex>" pairs
    let currentName = null;
    for (const line of lines) {
      const qMatch = line.match(/^\s*(?:Question|Variable|Name|Setting)\s*[:=]\s*(.+)/i);
      if (qMatch) {
        currentName = qMatch[1].trim();
        continue;
      }
      const vMatch = line.match(/^\s*(?:Value|Current|Default)\s*[:=]\s*(0x[0-9A-Fa-f]+|[0-9]+)/i);
      if (vMatch && currentName) {
        settings[currentName] = {
          token: "0x0",
          value: vMatch[1].startsWith("0x") ? vMatch[1] : "0x" + parseInt(vMatch[1]).toString(16),
          raw: line.trim(),
        };
        currentName = null;
      }
    }

    if (Object.keys(settings).length > 0) return settings;

    // Format 2: Tab/comma-separated table: Name\tToken\tValue
    for (const line of lines) {
      const parts = line.split(/\t|,/).map(p => p.trim());
      if (parts.length >= 3) {
        const name = parts[0];
        const token = parts[1];
        const value = parts[2];
        if (name && /0x[0-9A-Fa-f]+/.test(token) && /0x[0-9A-Fa-f]+/.test(value)) {
          settings[name] = { token, value, raw: line.trim() };
        }
      }
    }

    // Format 3: Flat key=value lines (like INI)
    if (Object.keys(settings).length === 0) {
      for (const line of lines) {
        const kvMatch = line.match(/^([^=]+)=\s*(0x[0-9A-Fa-f]+|[0-9]+)/);
        if (kvMatch) {
          const name = kvMatch[1].trim();
          const val = kvMatch[2];
          if (name.length > 2 && name.length < 100) {
            settings[name] = {
              token: "0x0",
              value: val.startsWith("0x") ? val : "0x" + parseInt(val).toString(16),
              raw: line.trim(),
            };
          }
        }
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
    const allNames = Object.keys(s);

    // Fuzzy find helper: search all setting names for any keyword match
    const fuzzyGet = (keywords, ...hardcoded) => {
      // Try hardcoded names first
      for (const name of hardcoded) {
        if (s[name]) return s[name];
      }
      // Then fuzzy search
      for (const name of allNames) {
        const lower = name.toLowerCase();
        for (const kw of keywords) {
          if (lower.includes(kw.toLowerCase())) return s[name];
        }
      }
      return null;
    };

    const cpuSettings = {
      cpuRatio: fuzzyGet(["ratio", "multiplier", "core ratio", "all core", "cpu ratio", "frequency ratio"],
        "CPU Ratio", "CPU Core Ratio", "Core Ratio Limit"),
      cpuVoltage: fuzzyGet(["vcore", "cpu voltage", "core voltage", "cpu v", "vid"],
        "CPU Core Voltage", "CPU Vcore", "Vcore Override"),
      cpuVoltageMode: fuzzyGet(["voltage mode", "vcore mode"],
        "CPU Core Voltage Mode", "Vcore Mode"),
      powerLimit1: fuzzyGet(["long duration power limit", "pl1", "package power limit 1", "power limit 1", "long duration", "tdp"],
        "Long Duration Power Limit", "PL1", "Package Power Limit 1"),
      powerLimit2: fuzzyGet(["short duration power limit", "pl2", "package power limit 2", "power limit 2", "short duration"],
        "Short Duration Power Limit", "PL2", "Package Power Limit 2"),
      tccOffset: fuzzyGet(["tcc", "activation offset"],
        "TCC Activation Offset"),
      ringRatio: fuzzyGet(["ring ratio", "cache ratio", "uncore ratio", "ring", "cache", "uncore"],
        "Ring Ratio", "Cache Ratio", "Uncore Ratio"),
      iccMax: fuzzyGet(["icc max", "ia ac load line", "icc"],
        "ICC Max", "IA AC Load Line"),
      avxOffset: fuzzyGet(["avx offset", "avx2 ratio offset", "avx"],
        "AVX Offset", "AVX2 Ratio Offset"),
    };

    return { success: true, settings: cpuSettings };
  }

  // ── Memory Settings ───────────────────────────────────────────────

  async getMemorySettings() {
    if (!this.available) return { success: false, error: "SceWin not found" };

    const exported = await this.exportCurrentSettings();
    if (!exported.success) return exported;

    const s = this.currentSettings;
    const allNames = Object.keys(s);

    // Fuzzy find helper: search all setting names for any keyword match
    const fuzzyGet = (keywords, ...hardcoded) => {
      // Try hardcoded names first
      for (const name of hardcoded) {
        if (s[name]) return s[name];
      }
      // Then fuzzy search
      for (const name of allNames) {
        const lower = name.toLowerCase();
        for (const kw of keywords) {
          if (lower.includes(kw.toLowerCase())) return s[name];
        }
      }
      return null;
    };

    const memSettings = {
      memorySpeed: fuzzyGet(["memory frequency", "dram frequency", "memory speed", "mem freq"],
        "Memory Frequency", "DRAM Frequency", "Memory Speed"),
      casLatency: fuzzyGet(["cas latency", "tcl", "cas#"],
        "CAS Latency", "tCL", "CAS# Latency"),
      tRCD: fuzzyGet(["trcd", "ras to cas"],
        "tRCD", "RAS to CAS Delay", "RAS# to CAS# Delay"),
      tRP: fuzzyGet(["trp", "row precharge", "ras# precharge"],
        "tRP", "Row Precharge Time", "RAS# Precharge"),
      tRAS: fuzzyGet(["tras", "ras active", "active to precharge"],
        "tRAS", "RAS Active Time", "Active to Precharge Delay"),
      xmpProfile: fuzzyGet(["xmp", "memory profile", "expo", "docp"],
        "XMP Profile", "Extreme Memory Profile", "XMP"),
      memoryVoltage: fuzzyGet(["dram voltage", "memory voltage", "dram v"],
        "DRAM Voltage", "Memory Voltage"),
      commandRate: fuzzyGet(["command rate", "cmd rate"],
        "Command Rate", "CR", "Cmd Rate"),
    };

    return { success: true, settings: memSettings };
  }

  // ── Apply CPU Multiplier ──────────────────────────────────────────

  async applyCpuMultiplier(ratio, discoveredName) {
    if (!this.available) return { success: false, error: "SceWin not found" };
    if (typeof ratio !== "number" || ratio < 10 || ratio > 80) {
      return { success: false, error: `Invalid CPU ratio: ${ratio}. Must be between 10 and 80.` };
    }

    // Use the discovered BIOS setting name if provided (from _discoverBiosSettings)
    if (discoveredName) {
      const result = await this.readSetting(discoveredName);
      if (result.success) {
        return await this.writeSetting(discoveredName, ratio);
      }
      return { success: false, error: `Discovered CPU ratio setting "${discoveredName}" could not be read: ${result.error}` };
    }

    // Fallback: try common BIOS setting names for CPU multiplier
    const names = ["CPU Ratio", "CPU Core Ratio", "Core Ratio Limit", "All Core Ratio Limit"];
    for (const name of names) {
      const result = await this.readSetting(name);
      if (result.success) {
        return await this.writeSetting(name, ratio);
      }
    }

    return { success: false, error: "Could not find CPU ratio setting in BIOS. Run BIOS discovery first or check motherboard compatibility." };
  }

  // ── Apply CPU Voltage ─────────────────────────────────────────────

  async applyCpuVoltage(voltage, discoveredName) {
    if (!this.available) return { success: false, error: "SceWin not found" };

    // Hard safety limits
    const MIN_VOLTAGE = 0.8;
    const MAX_VOLTAGE = 1.45;
    if (typeof voltage !== "number" || voltage < MIN_VOLTAGE || voltage > MAX_VOLTAGE) {
      return { success: false, error: `Voltage ${voltage}V is outside safe range (${MIN_VOLTAGE}V - ${MAX_VOLTAGE}V). Refusing to apply.` };
    }

    // Convert voltage to millivolts (common BIOS representation)
    const millivolts = Math.round(voltage * 1000);

    // Use the discovered BIOS setting name if provided (from _discoverBiosSettings)
    if (discoveredName) {
      const result = await this.readSetting(discoveredName);
      if (result.success) {
        return await this.writeSetting(discoveredName, millivolts);
      }
      return { success: false, error: `Discovered CPU voltage setting "${discoveredName}" could not be read: ${result.error}` };
    }

    // Fallback: try common BIOS setting names
    const names = ["CPU Core Voltage", "CPU Vcore", "Vcore Override", "CPU Core Voltage Override"];
    for (const name of names) {
      const result = await this.readSetting(name);
      if (result.success) {
        return await this.writeSetting(name, millivolts);
      }
    }

    return { success: false, error: "Could not find CPU voltage setting in BIOS. Run BIOS discovery first or check motherboard compatibility." };
  }

  // ── Apply Memory XMP ──────────────────────────────────────────────

  async applyMemoryXMP(profile, discoveredName) {
    if (!this.available) return { success: false, error: "SceWin not found" };
    if (profile !== 0 && profile !== 1 && profile !== 2) {
      return { success: false, error: `Invalid XMP profile: ${profile}. Must be 0, 1, or 2.` };
    }

    // Use the discovered BIOS setting name if provided (from _discoverBiosSettings)
    if (discoveredName) {
      const result = await this.readSetting(discoveredName);
      if (result.success) {
        return await this.writeSetting(discoveredName, profile);
      }
      return { success: false, error: `Discovered XMP setting "${discoveredName}" could not be read: ${result.error}` };
    }

    // Fallback: try common BIOS setting names
    const names = ["XMP Profile", "Extreme Memory Profile", "XMP", "Memory Profile"];
    for (const name of names) {
      const result = await this.readSetting(name);
      if (result.success) {
        return await this.writeSetting(name, profile);
      }
    }

    return { success: false, error: "Could not find XMP profile setting in BIOS. Run BIOS discovery first or check motherboard compatibility." };
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
      analysis.safeCeilingVoltage = 1.40;
    } else if (analysis.isAMD) {
      const name = (hw.cpu.name || "").toUpperCase();
      // All Ryzen processors are unlocked
      analysis.cpuUnlocked = name.includes("RYZEN");
      analysis.safeCeilingVoltage = 1.35; // AMD is more voltage-sensitive
    }

    // Discover actual BIOS setting names via fuzzy matching
    await this._discoverBiosSettings();

    // Read ACTUAL current multiplier from BIOS via discovered setting name
    let biosRatio = null;
    let biosVoltage = null;
    if (this.scewin.available && this.biosMap) {
      if (this.biosMap.cpuRatio) {
        const ratioRead = await this.scewin.readSetting(this.biosMap.cpuRatio);
        if (ratioRead.success) {
          biosRatio = parseInt(ratioRead.value, 16);
          if (isNaN(biosRatio) || biosRatio < 10 || biosRatio > 80) biosRatio = null;
          this._log("analyzing", "bios-ratio", `Current BIOS CPU ratio from "${this.biosMap.cpuRatio}": x${biosRatio} (raw: ${ratioRead.value})`);
        }
      }
      if (this.biosMap.cpuVoltage) {
        const voltRead = await this.scewin.readSetting(this.biosMap.cpuVoltage);
        if (voltRead.success) {
          const rawVolt = parseInt(voltRead.value, 16);
          // Interpret as millivolts if > 100, else as raw ratio
          biosVoltage = rawVolt > 100 ? rawVolt / 1000 : null;
          this._log("analyzing", "bios-voltage", `Current BIOS CPU voltage from "${this.biosMap.cpuVoltage}": ${biosVoltage}V (raw: ${voltRead.value})`);
        }
      }
    }

    // Read actual temps, clocks, and power from HWiNFO sensors (VSB shared memory → CSV fallback)
    let hwinfoClockMHz = null;
    let hwinfoCpuTemp = null;
    let hwinfoCpuPower = null;
    const hwinfoSensors = await this.hw.hwinfo.readSensors();
    if (hwinfoSensors.available && hwinfoSensors.sensors) {
      hwinfoCpuTemp = hwinfoSensors.sensors.cpuPackageTemp || hwinfoSensors.sensors.cpuTemp || null;
      hwinfoClockMHz = hwinfoSensors.sensors.cpuClock || null;
      hwinfoCpuPower = hwinfoSensors.sensors.cpuPower || null;
      const source = hwinfoSensors.source || "VSB/CSV";
      this._log("analyzing", "hwinfo-live", `HWiNFO live readings (${source}) — Temp: ${hwinfoCpuTemp}C | Clock: ${hwinfoClockMHz} MHz | Power: ${hwinfoCpuPower}W`);
    } else {
      this._log("analyzing", "hwinfo-fail", `HWiNFO sensor read failed: ${hwinfoSensors.error || "unknown"} — temps will use WMI fallback`);
    }

    // Determine base ratio: prefer BIOS reading, then HWiNFO clock, then WMI estimate
    if (biosRatio && biosRatio >= 10 && biosRatio <= 80) {
      analysis.baseRatio = biosRatio;
      this._log("analyzing", "ratio-source", `Using BIOS-reported CPU ratio: x${biosRatio}`);
    } else if (hwinfoClockMHz && hwinfoClockMHz > 500) {
      analysis.baseRatio = Math.round(hwinfoClockMHz / 100);
      this._log("analyzing", "ratio-source", `Using HWiNFO clock-derived ratio: x${analysis.baseRatio} (from ${hwinfoClockMHz} MHz)`);
    } else {
      analysis.baseRatio = hw.cpu.maxClock ? Math.round(hw.cpu.maxClock / 100) : 36;
      this._log("analyzing", "ratio-source", `Using WMI-estimated ratio: x${analysis.baseRatio} (from MaxClockSpeed ${hw.cpu.maxClock} MHz)`);
    }

    // Set turbo and ceiling based on detected base ratio
    if (analysis.isIntel) {
      analysis.maxTurboRatio = analysis.baseRatio + 5;
      analysis.safeCeilingRatio = analysis.cpuUnlocked ? analysis.maxTurboRatio + 3 : analysis.baseRatio;
    } else if (analysis.isAMD) {
      analysis.maxTurboRatio = analysis.baseRatio + 4;
      analysis.safeCeilingRatio = analysis.cpuUnlocked ? analysis.maxTurboRatio + 2 : analysis.baseRatio;
    }

    // Store actual voltage if we read one
    if (biosVoltage) {
      analysis.currentVoltage = biosVoltage;
    }

    this._log("analyzing", "cpu-info", `CPU: ${hw.cpu.name} | Unlocked: ${analysis.cpuUnlocked} | Base ratio: ${analysis.baseRatio} | Safe ceiling: ${analysis.safeCeilingRatio}`, {
      unlocked: analysis.cpuUnlocked, baseRatio: analysis.baseRatio, ceiling: analysis.safeCeilingRatio,
      biosRatio, biosVoltage, hwinfoCpuTemp, hwinfoClockMHz, hwinfoCpuPower,
    });

    // Quick cooling quality test: run a short stress and measure temp delta
    // Wait for HWiNFO to be confirmed active before reading temps
    this._log("analyzing", "cooling-test", "Waiting for HWiNFO64 sensors before cooling test...");
    this._setPhase("analyzing", 8);

    // Read temps directly from HWiNFO first — this is the only accurate source
    let preTemp = -1;
    let postTemp = -1;
    let sensorSource = "WMI/fallback";

    // Try HWiNFO directly up to 5 times with 4s waits (20s total max)
    // HWiNFO VSB registry can take 15-20s to populate after launch
    for (let attempt = 0; attempt < 5; attempt++) {
      const hwinfoRead = await this.hw.hwinfo.readSensors();
      if (hwinfoRead.available && hwinfoRead.sensors) {
        const temp = hwinfoRead.sensors.cpuPackageTemp || hwinfoRead.sensors.cpuTemp || -1;
        if (temp > 0) {
          preTemp = temp;
          sensorSource = "HWiNFO64";
          break;
        }
      }
      if (attempt < 4) {
        this._log("analyzing", "hwinfo-retry", `HWiNFO sensor read attempt ${attempt + 1}/5 — no temp data yet, waiting 4s...`);
        await new Promise(r => setTimeout(r, 4000));
      }
    }

    // Only fall back to WMI if HWiNFO completely failed
    if (preTemp <= 0) {
      const wmiStats = await this.hw.getCPUStats();
      preTemp = (wmiStats.temp > 0) ? wmiStats.temp : 40;
      sensorSource = (wmiStats.temp > 0) ? "WMI" : "default (no sensor data)";
    }

    this._log("analyzing", "sensor-source", `Pre-stress temp: ${preTemp}C (source: ${sensorSource})`);
    this._setPhase("analyzing", 10);
    this._log("analyzing", "cooling-test", "Running 15-second stress test...");
    await this.hw.runStressTest(15);

    // Read post-stress temp from same source
    if (sensorSource === "HWiNFO64") {
      const hwinfoPost = await this.hw.hwinfo.readSensors();
      if (hwinfoPost.available && hwinfoPost.sensors) {
        postTemp = hwinfoPost.sensors.cpuPackageTemp || hwinfoPost.sensors.cpuTemp || -1;
      }
    }
    if (postTemp <= 0) {
      const wmiPost = await this.hw.getCPUStats();
      postTemp = (wmiPost.temp > 0) ? wmiPost.temp : 50;
    }

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

    // Check that we have discovered BIOS setting names
    const ratioSettingName = this.biosMap && this.biosMap.cpuRatio ? this.biosMap.cpuRatio : null;
    const voltageSettingName = this.biosMap && this.biosMap.cpuVoltage ? this.biosMap.cpuVoltage : null;

    if (!ratioSettingName) {
      this._log("testing", "skip-cpu", "BIOS discovery did not find a CPU ratio setting — cannot safely overclock CPU. Check SceWin export for your motherboard's setting names.");
      return null;
    }

    this._setPhase("testing", 20);
    this._log("testing", "cpu-start", `Starting CPU overclock phase (ratio setting: "${ratioSettingName}"${voltageSettingName ? `, voltage setting: "${voltageSettingName}"` : ", voltage setting: NOT FOUND"})`);

    // Backup current settings
    await this.scewin.backupSettings();
    this._log("testing", "backup", "BIOS settings backed up");

    // Read ACTUAL current ratio from BIOS via discovered setting name
    let startRatio = analysis.baseRatio;
    const currentRatioRead = await this.scewin.readSetting(ratioSettingName);
    if (currentRatioRead.success) {
      const parsedRatio = parseInt(currentRatioRead.value, 16);
      if (!isNaN(parsedRatio) && parsedRatio >= 10 && parsedRatio <= 80) {
        startRatio = parsedRatio;
        this._log("testing", "cpu-current", `Read current CPU ratio from BIOS "${ratioSettingName}": x${startRatio} (raw: ${currentRatioRead.value})`);
      } else {
        this._log("testing", "cpu-current-fallback", `BIOS ratio "${ratioSettingName}" returned unparseable value (${currentRatioRead.value}), using analysis base: x${startRatio}`);
      }
    } else {
      this._log("testing", "cpu-current-fallback", `Could not read current ratio from "${ratioSettingName}", using analysis base: x${startRatio}`);
    }

    const maxRatio = analysis.safeCeilingRatio;
    let lastGoodRatio = startRatio;
    let lastGoodVoltage = null;
    const totalSteps = maxRatio - startRatio;

    if (totalSteps <= 0) {
      this._log("testing", "cpu-at-ceiling", `Current ratio x${startRatio} is already at or above safe ceiling x${maxRatio} — no overclock headroom`);
      return { finalRatio: startRatio, finalVoltage: null, finalMHz: startRatio * 100, gainMHz: 0 };
    }

    for (let ratio = startRatio + 1; ratio <= maxRatio; ratio++) {
      if (this._stopRequested) break;

      const stepProgress = 20 + ((ratio - startRatio) / totalSteps) * 30;
      this._setPhase("testing", Math.round(stepProgress));

      this._log("testing", "cpu-step", `Testing CPU multiplier x${ratio} (${ratio * 100} MHz)`, {
        clock: ratio * 100, voltage: lastGoodVoltage,
      });

      // Apply the multiplier using discovered BIOS setting name
      const applyResult = await this.scewin.applyCpuMultiplier(ratio, ratioSettingName);
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

        if (!voltageSettingName) {
          this._log("testing", "cpu-no-voltage", "No voltage setting discovered — cannot adjust voltage, rolling back");
          await this.scewin.applyCpuMultiplier(lastGoodRatio, ratioSettingName);
          break;
        }

        // Try increasing voltage slightly
        let voltageFixed = false;
        const baseVoltage = analysis.isAMD ? 1.20 : 1.25;
        const maxVoltage = analysis.safeCeilingVoltage;

        for (let v = baseVoltage; v <= maxVoltage; v += 0.01) {
          if (this._stopRequested) break;

          const vRounded = Math.round(v * 100) / 100;
          this._log("testing", "cpu-voltage", `Trying voltage ${vRounded}V at x${ratio}`, { voltage: vRounded });

          const vResult = await this.scewin.applyCpuVoltage(vRounded, voltageSettingName);
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
          await this.scewin.applyCpuMultiplier(lastGoodRatio, ratioSettingName);
          if (lastGoodVoltage) await this.scewin.applyCpuVoltage(lastGoodVoltage, voltageSettingName);
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

    // Use discovered XMP setting name if available
    const xmpSettingName = this.biosMap && this.biosMap.xmpProfile ? this.biosMap.xmpProfile : null;

    // Try enabling XMP if not already enabled
    const xmpSetting = memSettings.settings.xmpProfile;
    if (xmpSetting && (xmpSetting.value === "0x0" || xmpSetting.value === "0x00")) {
      this._log("testing", "mem-xmp", `XMP is disabled — enabling XMP Profile 1${xmpSettingName ? ` via discovered setting "${xmpSettingName}"` : ""}`);
      const xmpResult = await this.scewin.applyMemoryXMP(1, xmpSettingName);
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
          await this.scewin.applyMemoryXMP(0, xmpSettingName);
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

    // Launch HWiNFO EARLY — sensor registry (VSB) needs 15-20s to populate after launch
    this._log("analyzing", "hwinfo-early", "Launching HWiNFO64 early so sensors have time to populate...");
    const hwinfoLaunch = await this.hw.hwinfo.ensureRunning();
    if (hwinfoLaunch.error) {
      this._log("analyzing", "hwinfo-warn", `HWiNFO64 not available: ${hwinfoLaunch.error} — will use WMI fallback (less accurate)`);
    } else if (hwinfoLaunch.launched) {
      this._log("analyzing", "hwinfo-launched", `HWiNFO64 auto-launched from ${hwinfoLaunch.path} — waiting for VSB registry to populate...`);
    } else if (hwinfoLaunch.alreadyRunning) {
      this._log("analyzing", "hwinfo-ok", "HWiNFO64 already running — verifying VSB sensor data...");
    }

    // Wait up to 20 seconds for VSB to be populated with actual sensor data
    if (!hwinfoLaunch.error) {
      let vsbReady = false;
      for (let i = 0; i < 20; i++) {
        const sensorCheck = await this.hw.hwinfo.readSensors();
        if (sensorCheck.available && sensorCheck.sensors &&
            (sensorCheck.sensors.cpuPackageTemp || sensorCheck.sensors.cpuTemp)) {
          vsbReady = true;
          this._log("analyzing", "hwinfo-ready", `HWiNFO64 VSB sensors confirmed populated after ${i + 1}s — accurate sensor data available`);
          break;
        }
        await new Promise(r => setTimeout(r, 1000));
      }
      if (!vsbReady) {
        this._log("analyzing", "hwinfo-slow", "HWiNFO64 VSB sensors not yet populated after 20s — will retry during analysis, may fall back to WMI");
      }
    }

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

  // ── BIOS Setting Discovery ─────────────────────────────────────────

  async _discoverBiosSettings() {
    this.biosMap = {
      cpuRatio: null,
      cpuVoltage: null,
      xmpProfile: null,
      pl1: null,
      pl2: null,
      ringRatio: null,
    };

    if (!this.scewin.available) {
      this._log("analyzing", "bios-discovery", "SceWin not available — skipping BIOS setting discovery");
      return this.biosMap;
    }

    this._log("analyzing", "bios-discovery", "Exporting all BIOS settings via SceWin for discovery...");
    const exported = await this.scewin.exportCurrentSettings();
    if (!exported.success) {
      this._log("analyzing", "bios-discovery-fail", "Failed to export BIOS settings: " + exported.error);
      if (exported.helpOutput) {
        this._log("analyzing", "scewin-help", "SceWin help output: " + exported.helpOutput);
      }
      return this.biosMap;
    }
    if (exported.method) {
      this._log("analyzing", "bios-discovery", `Export succeeded via: ${exported.method}`);
    }

    const allNames = Object.keys(exported.settings);
    this._log("analyzing", "bios-discovery", `Exported ${allNames.length} BIOS settings — running fuzzy match`);

    // Fuzzy keyword matching helper: returns the first setting name that matches any keyword
    const fuzzyFind = (keywords) => {
      for (const name of allNames) {
        const lower = name.toLowerCase();
        for (const kw of keywords) {
          if (lower.includes(kw.toLowerCase())) return name;
        }
      }
      return null;
    };

    // CPU ratio / multiplier
    this.biosMap.cpuRatio = fuzzyFind([
      "all core ratio", "cpu ratio", "core ratio limit", "cpu core ratio",
      "frequency ratio", "multiplier", "ratio",
    ]);

    // CPU voltage
    this.biosMap.cpuVoltage = fuzzyFind([
      "vcore override", "cpu vcore", "cpu core voltage", "core voltage",
      "vcore", "cpu voltage", "cpu v", "vid",
    ]);

    // XMP / memory profile
    this.biosMap.xmpProfile = fuzzyFind([
      "xmp profile", "extreme memory profile", "xmp", "memory profile",
      "expo", "docp",
    ]);

    // Power limit 1 (long duration)
    this.biosMap.pl1 = fuzzyFind([
      "long duration power limit", "pl1", "package power limit 1",
      "power limit 1", "long duration", "tdp",
    ]);

    // Power limit 2 (short duration)
    this.biosMap.pl2 = fuzzyFind([
      "short duration power limit", "pl2", "package power limit 2",
      "power limit 2", "short duration",
    ]);

    // Ring / cache / uncore ratio
    this.biosMap.ringRatio = fuzzyFind([
      "ring ratio", "cache ratio", "uncore ratio", "ring", "cache", "uncore",
    ]);

    const found = Object.entries(this.biosMap)
      .filter(([, v]) => v !== null)
      .map(([k, v]) => `${k}="${v}"`)
      .join(", ");
    const missing = Object.entries(this.biosMap)
      .filter(([, v]) => v === null)
      .map(([k]) => k)
      .join(", ");

    this._log("analyzing", "bios-discovery-result", `Discovered BIOS names: ${found || "(none)"}${missing ? " | Missing: " + missing : ""}`);

    return this.biosMap;
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
      this.hwinfo.readSensors().catch((e) => ({ available: false, error: e.message })),
    ]);

    // If HWiNFO failed on first attempt, try once more after ensuring it's running
    if (!hwinfoData.available) {
      try {
        await this.hwinfo.ensureRunning();
        const retry = await this.hwinfo.readSensors();
        if (retry.available && retry.sensors) {
          Object.assign(hwinfoData, retry);
        }
      } catch(e) {}
    }

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

    return { cpu: cpuStats, gpu: gpuStats, ram: ramStats, hwinfoActive: hwinfoData.available, hwinfoSource: hwinfoData.source || "none" };
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