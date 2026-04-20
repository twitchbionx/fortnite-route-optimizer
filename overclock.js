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
const AI_OC_STATE_PATH = path.join(os.homedir(), ".fn-optimizer", "ai-oc-state.json");
const AI_OC_AUTOSTART_TASK = "FNOptimizerAutoOC";

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
  // NOTE: HWiNFO64 Free does NOT support CLI flags (-sm, /csv, etc) — Pro only.
  // We configure shared memory via registry + INI, then launch normally.
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

    // Enable Shared Memory + Sensors-Only mode via registry BEFORE launching
    await runPS(`
      $paths = @('HKCU:\\Software\\HWiNFO64', 'HKCU:\\Software\\HWiNFO64\\VSB')
      foreach ($p in $paths) { if (!(Test-Path $p)) { New-Item -Path $p -Force | Out-Null } }
      Set-ItemProperty -Path 'HKCU:\\Software\\HWiNFO64' -Name 'SensorsSM' -Value 1 -Type DWord -Force
      Set-ItemProperty -Path 'HKCU:\\Software\\HWiNFO64' -Name 'SensorsOnly' -Value 1 -Type DWord -Force
      Set-ItemProperty -Path 'HKCU:\\Software\\HWiNFO64' -Name 'AutoStart' -Value 1 -Type DWord -Force
      Set-ItemProperty -Path 'HKCU:\\Software\\HWiNFO64' -Name 'MinimizeMainWindow' -Value 1 -Type DWord -Force
      Set-ItemProperty -Path 'HKCU:\\Software\\HWiNFO64' -Name 'MinimizeSensors' -Value 1 -Type DWord -Force
      Set-ItemProperty -Path 'HKCU:\\Software\\HWiNFO64' -Name 'ShowWelcome' -Value 0 -Type DWord -Force
    `, 5000);

    // Write INI file next to the EXE (works for portable versions)
    // Also write to %APPDATA%/HWiNFO64 for installed versions
    const iniContent = "[Settings]\nSensorsSM=1\nSensorsOnly=1\nMinimize=1\nAutoStart=1\nShowWelcome=0\nMinimizeMainWindow=1\nMinimizeSensors=1\n";
    const iniLocations = [
      path.join(path.dirname(this.hwInfoPath), "HWiNFO64.INI"),
      path.join(process.env.APPDATA || "", "HWiNFO64", "HWiNFO64.INI"),
    ];
    for (const iniPath of iniLocations) {
      try {
        const iniDir = path.dirname(iniPath);
        if (!fs.existsSync(iniDir)) fs.mkdirSync(iniDir, { recursive: true });
        fs.writeFileSync(iniPath, iniContent);
      } catch(e) {} // may fail in Program Files without admin
    }

    // Kill any existing HWiNFO so it restarts with our new settings
    const running = await this.isRunning();
    if (running) {
      await runCmd('taskkill /IM HWiNFO64.exe /F 2>nul', 5000);
      await new Promise(r => setTimeout(r, 3000));
    }

    // Launch HWiNFO normally — NO CLI flags (Free version doesn't support them)
    // The registry + INI settings above will make it start in sensors-only mode
    // with shared memory enabled automatically
    await runPS(`Start-Process -FilePath '${this.hwInfoPath.replace(/'/g, "''")}' -WindowStyle Minimized`, 5000);

    // Wait up to 30 seconds for VSB data to appear
    // HWiNFO Free shows a welcome/nag screen that the user may need to click through
    for (let i = 0; i < 30; i++) {
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

    // Check if HWiNFO is at least running (user may need to click through welcome screen)
    const stillRunning = await this.isRunning();
    if (stillRunning) {
      return { launched: true, path: this.hwInfoPath, hasData: false, warning: "HWiNFO is running but Shared Memory data not detected yet. Please check: (1) Click through any welcome/nag screen in HWiNFO, (2) Go to Settings → check 'Shared Memory Support', (3) Make sure sensors are running (not just the summary)" };
    }

    return { launched: true, path: this.hwInfoPath, hasData: false, warning: "HWiNFO was launched but may have closed. Open HWiNFO64 manually → Settings → enable 'Shared Memory Support' → run Sensors Only" };
  }

  // Enhanced WMI + nvidia-smi fallback for when HWiNFO VSB isn't available
  // This gets real temps/clocks from the OS and GPU driver directly
  async readSensorsDirect() {
    const sensors = {
      cpuTemp: null, cpuPackageTemp: null, cpuVoltage: null, cpuClock: null,
      cpuPower: null, gpuTemp: null, gpuClock: null, gpuMemClock: null,
      gpuLoad: null, gpuPower: null, vrmTemp: null, fanSpeeds: {},
    };

    // CPU temp: try multiple WMI sources
    const tempResult = await runPS(`
      $temp = $null
      # Method 1: MSAcpi_ThermalZoneTemperature (most common)
      try {
        $t = Get-CimInstance MSAcpi_ThermalZoneTemperature -Namespace root/wmi -ErrorAction Stop | Select-Object -First 1
        if ($t -and $t.CurrentTemperature -gt 2000) { $temp = [math]::Round(($t.CurrentTemperature - 2732) / 10, 1) }
      } catch {}
      # Method 2: Win32_TemperatureProbe
      if (!$temp -or $temp -le 0) {
        try {
          $t2 = Get-CimInstance Win32_TemperatureProbe -ErrorAction Stop | Select-Object -First 1
          if ($t2 -and $t2.CurrentReading -gt 0) { $temp = $t2.CurrentReading }
        } catch {}
      }
      # CPU clock from Win32_Processor
      $clock = $null
      try {
        $p = Get-CimInstance Win32_Processor -ErrorAction Stop | Select-Object -First 1
        $clock = $p.CurrentClockSpeed
      } catch {}
      @{ temp = $temp; clock = $clock } | ConvertTo-Json
    `, 10000);

    try {
      const data = JSON.parse(tempResult.output);
      if (data.temp && data.temp > 0) sensors.cpuTemp = data.temp;
      if (data.clock && data.clock > 0) sensors.cpuClock = data.clock;
    } catch(e) {}

    // GPU: nvidia-smi gives us excellent data for NVIDIA cards
    const nvPath = "C:\\Windows\\System32\\nvidia-smi.exe";
    const nv = await runCmd(`"${nvPath}" --query-gpu=temperature.gpu,power.draw,clocks.current.graphics,clocks.current.memory,fan.speed,utilization.gpu --format=csv,noheader,nounits 2>nul`, 10000);
    if (nv.success && nv.output) {
      const p = nv.output.split(",").map(s => parseFloat(s.trim()));
      if (p[0] > 0) sensors.gpuTemp = p[0];
      if (p[1] > 0) sensors.gpuPower = p[1];
      if (p[2] > 0) sensors.gpuClock = p[2];
      if (p[3] > 0) sensors.gpuMemClock = p[3];
      if (p[5] > 0) sensors.gpuLoad = p[5];
    }

    const hasSomething = sensors.cpuTemp || sensors.gpuTemp || sensors.cpuClock;
    return { available: hasSomething, sensors, source: "WMI-direct" };
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
        // HWiNFO not found — fall back to direct WMI + nvidia-smi
        return await this.readSensorsDirect();
      }
      // Re-check after launch attempt
      const recheck = await runPS(`Test-Path 'HKCU:\\Software\\HWiNFO64\\VSB'`, 3000);
      if (!recheck.success || !recheck.output.includes("True")) {
        // VSB still not available — fall back to direct WMI + nvidia-smi
        return await this.readSensorsDirect();
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
      // VSB read failed — fall back to direct WMI + nvidia-smi
      return await this.readSensorsDirect();
    }

    try {
      const data = JSON.parse(result.output);
      if (data.error) {
        // Registry exists but no sensor data — fall back to direct
        return await this.readSensorsDirect();
      }
      return { available: true, sensors: this._parseSensors(data), source: "VSB-shared-memory" };
    } catch(e) {
      // Parse failed — fall back to direct
      return await this.readSensorsDirect();
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

    // ═══════════════════════════════════════════════════════════════════
    // HWiNFO VSB SENSOR PARSING — UNIT-AWARE
    //
    // The VSB registry stores Label/Value/ValueRaw for each sensor.
    // The Value string contains the unit: "50.0 °C", "1.250 V", "4800 MHz", etc.
    // We use the unit from Value to CATEGORIZE the sensor type first,
    // then match by label name. This avoids the old bug where "CPU Package"
    // (no "temp" in label) was missed, and "CPU DTS" or thermal zone
    // sensors with wrong values were matched instead.
    //
    // Priority for CPU temperature (highest to lowest):
    //   1. "CPU Package" — Intel package temp (most accurate for overall CPU)
    //   2. "CPU (Tctl/Tdie)" — AMD equivalent
    //   3. "Core Max" / "CPU Core Max" — highest core
    //   4. "Core Average" — average across all cores
    //   5. Any label with "cpu" + unit is °C
    // ═══════════════════════════════════════════════════════════════════

    // Classify each sensor by unit from Value string
    const tempSensors = [];   // { label, lbl, val }
    const voltSensors = [];
    const clockSensors = [];
    const powerSensors = [];
    const loadSensors = [];
    const fanSensors = [];
    const otherSensors = [];

    for (const [label, data] of Object.entries(raw)) {
      const val = parseFloat(data.raw != null ? data.raw : data.value);
      if (isNaN(val)) continue;

      const valueStr = (data.value || "").toString();
      const lbl = label.toLowerCase();
      const entry = { label, lbl, val };

      // Determine sensor type from Value string unit
      if (valueStr.includes("°C") || valueStr.includes("\u00B0C") || valueStr.includes("deg C")) {
        tempSensors.push(entry);
      } else if (/\bV\b/.test(valueStr) || valueStr.includes("Volt")) {
        voltSensors.push(entry);
      } else if (valueStr.includes("MHz") || valueStr.includes("GHz")) {
        clockSensors.push(entry);
      } else if (valueStr.includes(" W") || valueStr.includes("Watt")) {
        powerSensors.push(entry);
      } else if (valueStr.includes("%")) {
        loadSensors.push(entry);
      } else if (valueStr.includes("RPM")) {
        fanSensors.push(entry);
      } else {
        // Fallback: guess from label keywords
        if (lbl.includes("temp") || lbl.includes("thermal")) tempSensors.push(entry);
        else if (lbl.includes("volt") || lbl.includes("vcore")) voltSensors.push(entry);
        else if (lbl.includes("clock") || lbl.includes("freq") || lbl.includes("mhz")) clockSensors.push(entry);
        else if (lbl.includes("power") || lbl.includes("watt") || lbl.includes("tdp")) powerSensors.push(entry);
        else if (lbl.includes("load") || lbl.includes("usage") || lbl.includes("util")) loadSensors.push(entry);
        else if (lbl.includes("fan") || lbl.includes("rpm")) fanSensors.push(entry);
        else otherSensors.push(entry);
      }
    }

    // ── CPU TEMPERATURE (from confirmed temperature sensors only) ──
    // Priority: CPU Package > Tctl/Tdie > Core Max > Core Average > any CPU temp
    const cpuTempPriority = [
      (s) => s.lbl.includes("cpu") && s.lbl.includes("package") && !s.lbl.includes("power"),
      (s) => s.lbl.includes("tctl") || s.lbl.includes("tdie") || s.lbl.includes("tctl/tdie"),
      (s) => s.lbl.includes("cpu") && (s.lbl.includes("core max") || s.lbl.includes("ccd")),
      (s) => s.lbl.includes("core") && s.lbl.includes("average"),
      (s) => s.lbl.includes("core") && s.lbl.includes("max"),
      (s) => s.lbl.includes("cpu") && !s.lbl.includes("vrm") && !s.lbl.includes("vr "),
      (s) => s.lbl.includes("core") && s.val > 15 && s.val < 110,
    ];
    for (const matcher of cpuTempPriority) {
      const found = tempSensors.find(matcher);
      if (found && found.val > 10 && found.val < 115) {
        sensors.cpuPackageTemp = found.val;
        break;
      }
    }
    // Also grab any generic CPU temp that's different from package
    if (!sensors.cpuPackageTemp) {
      const anyCpu = tempSensors.find(s => s.lbl.includes("cpu") && s.val > 10 && s.val < 115);
      if (anyCpu) sensors.cpuPackageTemp = anyCpu.val;
    }
    // Set cpuTemp as alias for backward compatibility
    sensors.cpuTemp = sensors.cpuPackageTemp;

    // ── GPU TEMPERATURE ──
    const gpuTemp = tempSensors.find(s =>
      (s.lbl.includes("gpu") || s.lbl.includes("geforce") || s.lbl.includes("radeon")) &&
      !s.lbl.includes("hot spot") && !s.lbl.includes("hotspot") &&
      !s.lbl.includes("memory") && !s.lbl.includes("junction")
    ) || tempSensors.find(s => s.lbl.includes("gpu"));
    if (gpuTemp && gpuTemp.val > 0) sensors.gpuTemp = gpuTemp.val;

    // ── VRM TEMPERATURE ──
    const vrmTemp = tempSensors.find(s => s.lbl.includes("vrm") || s.lbl.includes("vr mos") || s.lbl.includes("mos temp"));
    if (vrmTemp && vrmTemp.val > 0) sensors.vrmTemp = vrmTemp.val;

    // ── RAM TEMPERATURE ──
    const ramTemp = tempSensors.find(s => s.lbl.includes("dimm") || (s.lbl.includes("memory") && !s.lbl.includes("gpu")));
    if (ramTemp && ramTemp.val > 0) sensors.ramTemp = ramTemp.val;

    // ── CPU VOLTAGE ──
    const cpuVolt = voltSensors.find(s => s.lbl.includes("vcore") || s.lbl.includes("cpu core"))
      || voltSensors.find(s => s.lbl.includes("cpu") && s.lbl.includes("volt"))
      || voltSensors.find(s => s.lbl.includes("cpu") && s.lbl.includes("vid"))
      || voltSensors.find(s => s.lbl.includes("core") && s.val > 0.5 && s.val < 2.0);
    if (cpuVolt && cpuVolt.val > 0) sensors.cpuVoltage = cpuVolt.val;

    // ── CPU CLOCK ──
    const cpuClock = clockSensors.find(s => s.lbl.includes("cpu") && s.lbl.includes("core") && !s.lbl.includes("ring") && !s.lbl.includes("uncore"))
      || clockSensors.find(s => s.lbl.includes("core") && s.lbl.includes("clock") && !s.lbl.includes("ring"))
      || clockSensors.find(s => s.lbl.includes("cpu") && !s.lbl.includes("ring") && !s.lbl.includes("uncore") && !s.lbl.includes("bus") && !s.lbl.includes("bclk"))
      || clockSensors.find(s => s.lbl.includes("p-core") || s.lbl.includes("core 0"));
    if (cpuClock && cpuClock.val > 100) sensors.cpuClock = cpuClock.val;

    // ── CPU POWER ──
    const cpuPower = powerSensors.find(s => s.lbl.includes("cpu") && s.lbl.includes("package"))
      || powerSensors.find(s => s.lbl.includes("cpu package"))
      || powerSensors.find(s => s.lbl.includes("cpu") && s.lbl.includes("power"))
      || powerSensors.find(s => s.lbl.includes("package power"));
    if (cpuPower && cpuPower.val > 0) sensors.cpuPower = cpuPower.val;

    // ── GPU CLOCK ──
    const gpuClock = clockSensors.find(s => s.lbl.includes("gpu") && s.lbl.includes("core"))
      || clockSensors.find(s => s.lbl.includes("gpu") && s.lbl.includes("clock") && !s.lbl.includes("mem"))
      || clockSensors.find(s => s.lbl.includes("gpu") && !s.lbl.includes("mem") && !s.lbl.includes("video"));
    if (gpuClock && gpuClock.val > 0) sensors.gpuClock = gpuClock.val;

    // ── GPU MEM CLOCK ──
    const gpuMemClk = clockSensors.find(s => s.lbl.includes("gpu") && (s.lbl.includes("mem") || s.lbl.includes("memory")));
    if (gpuMemClk && gpuMemClk.val > 0) sensors.gpuMemClock = gpuMemClk.val;

    // ── GPU LOAD ──
    const gpuLoad = loadSensors.find(s => s.lbl.includes("gpu") && (s.lbl.includes("core") || s.lbl.includes("load") || s.lbl.includes("util")))
      || loadSensors.find(s => s.lbl.includes("gpu") && !s.lbl.includes("mem") && !s.lbl.includes("video"));
    if (gpuLoad) sensors.gpuLoad = gpuLoad.val;

    // ── GPU POWER ──
    const gpuPow = powerSensors.find(s => s.lbl.includes("gpu") || s.lbl.includes("geforce") || s.lbl.includes("radeon"));
    if (gpuPow && gpuPow.val > 0) sensors.gpuPower = gpuPow.val;

    // ── GPU VRAM ──
    const gpuVram = otherSensors.concat(loadSensors).find(s => s.lbl.includes("gpu") && s.lbl.includes("memory") && s.lbl.includes("used"));
    if (gpuVram && gpuVram.val > 0) sensors.gpuVram = gpuVram.val;

    // ── FAN SPEEDS ──
    for (const f of fanSensors) {
      if (f.val > 0) sensors.fanSpeeds[f.label] = f.val;
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

    const scewinDir = path.dirname(this.scewinPath);
    const logFile = path.join(scewinDir, "fn_export_log.txt");

    // ═══════════════════════════════════════════════════════════════════
    // Strategy v3.9.0: PRACTICAL APPROACH
    //
    // What we learned from v3.8.1–3.8.9:
    //  - SCEWIN launched from Electron (any method) gives EXIT:16
    //  - The user's own Export.bat works when run manually as admin
    //  - nvram.txt (1.4MB) already exists from a previous manual export
    //  - The BIOS settings rarely change, so re-exporting every time is wasteful
    //
    // New approach (in order):
    //  1. Check for existing export files (nvram.txt, fn_export.txt, etc.)
    //     If any exist and are >100KB, parse them directly — no re-export needed
    //  2. If no existing files, try running the user's Export.bat elevated
    //  3. If that fails, try direct SCEWIN elevation
    //  4. If everything fails, show clear instructions for manual export
    // ═══════════════════════════════════════════════════════════════════

    // Helper: try to parse a file and return settings if valid
    const tryParseFile = (filePath) => {
      try {
        if (!fs.existsSync(filePath)) return null;
        const raw = fs.readFileSync(filePath, "utf8");
        if (raw.length < 100) return null;

        // Try AMISCE script format parser
        const settings = this._parseExport(raw);
        let count = Object.keys(settings).length;
        if (count > 5) return { settings, count, path: filePath, method: "existing-file" };

        // Try alternate parser
        const altSettings = this._parseExportAlt(raw);
        count = Object.keys(altSettings).length;
        if (count > 5) return { settings: altSettings, count, path: filePath, method: "existing-file-alt" };

        return null;
      } catch(e) { return null; }
    };

    // ── Step 1: Check existing export files ──
    // Don't delete anything! Use what's already there.
    const existingFiles = [
      path.join(scewinDir, "nvram.txt"),        // SCEWIN default output name
      path.join(scewinDir, "fn_export.txt"),     // our custom name
      path.join(scewinDir, "AMISCE.txt"),        // AMISCE default name
      exportPath,                                // our target path
    ];

    // Also scan scewin dir for any large .txt files that could be exports
    try {
      const files = fs.readdirSync(scewinDir);
      for (const f of files) {
        if (f.endsWith(".txt") || f.endsWith(".TXT")) {
          const fp = path.join(scewinDir, f);
          if (!existingFiles.includes(fp)) {
            try {
              const stat = fs.statSync(fp);
              if (stat.size > 100000) existingFiles.push(fp); // >100KB = likely an export
            } catch(e) {}
          }
        }
      }
    } catch(e) {}

    for (const filePath of existingFiles) {
      const result = tryParseFile(filePath);
      if (result) {
        this.currentSettings = result.settings;
        if (filePath !== exportPath) {
          try { fs.copyFileSync(filePath, exportPath); } catch(e) {}
        }
        return {
          success: true,
          settings: result.settings,
          settingCount: result.count,
          method: result.method,
          source: filePath,
        };
      }
    }

    // ── Step 2: Try running user's Export.bat if it exists ──
    const userExportBat = path.join(scewinDir, "Export.bat");
    if (fs.existsSync(userExportBat)) {
      const psScript = [
        "try {",
        "  $psi = New-Object System.Diagnostics.ProcessStartInfo",
        `  $psi.FileName = '${userExportBat.replace(/'/g, "''")}'`,
        `  $psi.WorkingDirectory = '${scewinDir.replace(/'/g, "''")}'`,
        "  $psi.Verb = 'RunAs'",
        "  $psi.UseShellExecute = $true",
        "  $psi.WindowStyle = 'Normal'",
        "  $proc = [System.Diagnostics.Process]::Start($psi)",
        "  $proc.WaitForExit(60000)",
        "  if (!$proc.HasExited) { try { $proc.Kill() } catch {} }",
        "  Write-Output \"BAT_EXIT:$($proc.ExitCode)\"",
        "} catch { Write-Output \"BAT_ERROR:$($_.Exception.Message)\" }",
      ].join("\n");

      await runPS(psScript, 75000);
      await new Promise(r => setTimeout(r, 3000));

      // Check for new files
      for (const filePath of existingFiles) {
        const result = tryParseFile(filePath);
        if (result) {
          this.currentSettings = result.settings;
          if (filePath !== exportPath) {
            try { fs.copyFileSync(filePath, exportPath); } catch(e) {}
          }
          return {
            success: true,
            settings: result.settings,
            settingCount: result.count,
            method: "export-bat",
            source: filePath,
          };
        }
      }
    }

    // ── Step 3: Try direct SCEWIN elevation ──
    const attempts = [
      `/o /s nvram.txt`,
      `/o /s fn_export.txt`,
      `/o`,
    ];

    for (const args of attempts) {
      const psScript = [
        "try {",
        "  $psi = New-Object System.Diagnostics.ProcessStartInfo",
        `  $psi.FileName = '${this.scewinPath.replace(/'/g, "''")}'`,
        `  $psi.Arguments = '${args}'`,
        `  $psi.WorkingDirectory = '${scewinDir.replace(/'/g, "''")}'`,
        "  $psi.Verb = 'RunAs'",
        "  $psi.UseShellExecute = $true",
        "  $psi.WindowStyle = 'Normal'",
        "  $proc = [System.Diagnostics.Process]::Start($psi)",
        "  $proc.WaitForExit(30000)",
        "  if (!$proc.HasExited) { try { $proc.Kill() } catch {} }",
        "} catch {}",
      ].join("\n");

      await runPS(psScript, 45000);
      await new Promise(r => setTimeout(r, 3000));

      for (const filePath of existingFiles) {
        const result = tryParseFile(filePath);
        if (result) {
          this.currentSettings = result.settings;
          if (filePath !== exportPath) {
            try { fs.copyFileSync(filePath, exportPath); } catch(e) {}
          }
          return {
            success: true,
            settings: result.settings,
            settingCount: result.count,
            method: "direct-elevation",
            source: filePath,
          };
        }
      }
    }

    // ── Step 4: Everything failed — give clear instructions ──
    return {
      success: false,
      error: "Could not export BIOS settings. Please run Export.bat manually as Administrator from: " + scewinDir + " — then restart the AI OC. The app will automatically find and parse the export file.",
    };
  }

  _parseExport(raw) {
    // ═══════════════════════════════════════════════════════════════════
    // AMISCE Script Format Parser (SCEWIN /o /s output)
    //
    // The actual format from SCEWIN_64 /o /s is a script file:
    //
    //   Setup Question = BCLK Output Source
    //   Help String = <description text>
    //   Token =14 // Do NOT change this line
    //   Offset =CBS
    //   Width =01
    //   BIOS Default =[00]CPU BCLK
    //   Options =*[00]CPU BCLK   // asterisk marks current value
    //            [02]Buffer
    //
    // Key fields per setting:
    //   - "Setup Question = <name>" — the setting name
    //   - "Token =<decimal>" — the token number (used for writes)
    //   - "Options =*[xx]<label>" — current value (marked with *)
    //   - "BIOS Default =[xx]<label>" — factory default
    //   - "Offset =<hex|string>" — NVRAM offset
    //   - "Width =<number>" — data width in bytes
    //
    // For numeric settings (ratios, voltages, etc.) the format is:
    //   Setup Question = CPU Base Clock 100.00MHz
    //   Token =14
    //   Offset =01
    //   Width =02
    //   BIOS Default =[2710]
    //   Value =2710    // current value in hex (no options list)
    // ═══════════════════════════════════════════════════════════════════
    const settings = {};
    const lines = raw.split("\n");
    let current = null; // current setting being parsed

    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];
      const trimmed = line.trim();

      // "Setup Question = <name>" starts a new setting block
      const nameMatch = trimmed.match(/^Setup Question\s*=\s*(.+)/i);
      if (nameMatch) {
        // Save previous setting if it had enough info
        if (current && current.name && (current.value !== null || current.token !== null)) {
          settings[current.name] = {
            token: current.token || "0",
            value: current.value !== null ? current.value : current.biosDefault || "0x0",
            currentLabel: current.currentLabel || "",
            biosDefault: current.biosDefault || "",
            defaultLabel: current.defaultLabel || "",
            offset: current.offset || "",
            width: current.width || "",
            options: current.options || [],
            helpString: current.helpString || "",
          };
        }
        current = {
          name: nameMatch[1].trim(),
          token: null,
          value: null,
          currentLabel: "",
          biosDefault: "",
          defaultLabel: "",
          offset: "",
          width: "",
          options: [],
          helpString: "",
        };
        continue;
      }

      if (!current) continue;

      // "Token =<number>" — token ID (decimal or hex)
      const tokenMatch = trimmed.match(/^Token\s*=\s*([0-9A-Fa-fx]+)/i);
      if (tokenMatch) {
        const t = tokenMatch[1];
        current.token = t.startsWith("0x") ? t : "0x" + parseInt(t, 10).toString(16);
        continue;
      }

      // "Value =<hex>" — direct numeric value (for numeric settings)
      const valueMatch = trimmed.match(/^Value\s*=\s*([0-9A-Fa-f]+)/i);
      if (valueMatch && !trimmed.startsWith("//")) {
        current.value = "0x" + valueMatch[1];
        continue;
      }

      // "BIOS Default =[xx]<label>" or "BIOS Default =[xx]"
      const defaultMatch = trimmed.match(/^BIOS Default\s*=\s*\[([0-9A-Fa-f]+)\]\s*(.*)/i);
      if (defaultMatch) {
        current.biosDefault = "0x" + defaultMatch[1];
        current.defaultLabel = defaultMatch[2] ? defaultMatch[2].trim() : "";
        continue;
      }

      // "Options =*[xx]<label>" or "Options =[xx]<label>" — option list
      // The asterisk (*) marks the currently selected option
      const optionsMatch = trimmed.match(/^Options\s*=\s*(\*?)\[([0-9A-Fa-f]+)\]\s*(.*)/i);
      if (optionsMatch) {
        const isSelected = optionsMatch[1] === "*";
        const val = "0x" + optionsMatch[2];
        const label = (optionsMatch[3] || "").split("//")[0].trim();
        current.options.push({ value: val, label, selected: isSelected });
        if (isSelected) {
          current.value = val;
          current.currentLabel = label;
        }
        continue;
      }

      // Continuation option lines: "  *[xx]<label>" or "  [xx]<label>"
      const contOptionMatch = trimmed.match(/^(\*?)\[([0-9A-Fa-f]+)\]\s*(.*)/);
      if (contOptionMatch) {
        const isSelected = contOptionMatch[1] === "*";
        const val = "0x" + contOptionMatch[2];
        const label = (contOptionMatch[3] || "").split("//")[0].trim();
        current.options.push({ value: val, label, selected: isSelected });
        if (isSelected) {
          current.value = val;
          current.currentLabel = label;
        }
        continue;
      }

      // "Offset =<hex|string>"
      const offsetMatch = trimmed.match(/^Offset\s*=\s*(.+)/i);
      if (offsetMatch) {
        current.offset = offsetMatch[1].trim();
        continue;
      }

      // "Width =<number>"
      const widthMatch = trimmed.match(/^Width\s*=\s*(.+)/i);
      if (widthMatch) {
        current.width = widthMatch[1].trim();
        continue;
      }

      // "Help String = <text>"
      const helpMatch = trimmed.match(/^Help String\s*=\s*(.*)/i);
      if (helpMatch) {
        current.helpString = helpMatch[1].trim();
        continue;
      }
    }

    // Don't forget the last setting
    if (current && current.name && (current.value !== null || current.token !== null)) {
      settings[current.name] = {
        token: current.token || "0",
        value: current.value !== null ? current.value : current.biosDefault || "0x0",
        currentLabel: current.currentLabel || "",
        biosDefault: current.biosDefault || "",
        defaultLabel: current.defaultLabel || "",
        offset: current.offset || "",
        width: current.width || "",
        options: current.options || [],
        helpString: current.helpString || "",
      };
    }

    return settings;
  }

  // Alternate parser for non-script AMISCE formats (simple Token/Value pairs)
  _parseExportAlt(raw) {
    const settings = {};
    const lines = raw.split("\n");

    // Format 1: "Setup Question = <name>" with "Token = 0xHEX  Value = 0xHEX" on one line
    let currentName = null;
    for (const line of lines) {
      const qMatch = line.match(/^\s*(?:Setup Question|Question|Variable|Name|Setting)\s*[:=]\s*(.+)/i);
      if (qMatch) {
        currentName = qMatch[1].trim();
        continue;
      }
      const vMatch = line.match(/Token\s*=\s*(0x[0-9A-Fa-f]+)\s+.*?Value\s*=\s*(0x[0-9A-Fa-f]+)/i);
      if (vMatch && currentName) {
        settings[currentName] = { token: vMatch[1], value: vMatch[2], raw: line.trim() };
        currentName = null;
        continue;
      }
      // Also try simple "Value = <hex>" on its own line
      const simpleVal = line.match(/^\s*(?:Value|Current|Default)\s*[:=]\s*(0x[0-9A-Fa-f]+|[0-9]+)/i);
      if (simpleVal && currentName) {
        const v = simpleVal[1];
        settings[currentName] = {
          token: "0x0",
          value: v.startsWith("0x") ? v : "0x" + parseInt(v).toString(16),
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

  // Helper: run any SCEWIN command elevated via scheduled task (SYSTEM)
  // The AMISCE driver only loads when exe runs as SYSTEM or is direct UAC target.
  // Scheduled task approach gives us SYSTEM + stdin piping via batch.
  async _runScewinElevated(args, timeout = 45000) {
    const scewinDir = path.dirname(this.scewinPath);
    const logFile = path.join(scewinDir, "fn_scewin_cmd_log.txt");
    try { fs.unlinkSync(logFile); } catch(e) {}

    const taskName = "FNOptSceWinCmd";
    const batchPath = path.join(scewinDir, "fn_task_cmd.bat");

    // Build batch that pipes stdin and logs output
    const batchContent = [
      "@echo off",
      `cd /d "${scewinDir}"`,
      `echo [CMD_START] >> "${logFile}"`,
      `echo [ARGS] ${args} >> "${logFile}"`,
      `echo. | "${this.scewinPath}" ${args} >> "${logFile}" 2>&1`,
      `echo EXIT_CODE:%ERRORLEVEL% >> "${logFile}"`,
      `echo [CMD_DONE] >> "${logFile}"`,
    ].join("\r\n");
    fs.writeFileSync(batchPath, batchContent, "utf8");

    // Create and run scheduled task as SYSTEM
    const createPS = [
      "try { schtasks /delete /tn '" + taskName + "' /f 2>$null } catch {}",
      "schtasks /create /tn '" + taskName + "' /tr '\"" + batchPath.replace(/'/g, "''") + "\"' /sc once /st 00:00 /ru SYSTEM /rl HIGHEST /f",
      "Start-Sleep -Milliseconds 500",
      "schtasks /run /tn '" + taskName + "'",
    ].join("; ");

    const encodedCreate = Buffer.from(createPS, "utf16le").toString("base64");
    const elevatePS = [
      "try {",
      "  $psi = New-Object System.Diagnostics.ProcessStartInfo",
      "  $psi.FileName = 'powershell.exe'",
      `  $psi.Arguments = '-NoProfile -ExecutionPolicy Bypass -EncodedCommand ${encodedCreate}'`,
      "  $psi.Verb = 'RunAs'",
      "  $psi.UseShellExecute = $true",
      "  $psi.WindowStyle = 'Hidden'",
      "  $proc = [System.Diagnostics.Process]::Start($psi)",
      "  $proc.WaitForExit(20000)",
      "  Write-Output 'TASK_LAUNCHED'",
      "} catch { Write-Output \"ERROR:$($_.Exception.Message)\" }",
    ].join("\n");

    const result = await runPS(elevatePS, 30000);

    // Wait for task completion
    let log = "";
    const deadline = Date.now() + timeout;
    while (Date.now() < deadline) {
      await new Promise(r => setTimeout(r, 2000));
      try { log = fs.readFileSync(logFile, "utf8"); } catch(e) {}
      if (log.includes("[CMD_DONE]")) break;
    }

    // Clean up
    try { await runPS("try { schtasks /delete /tn '" + taskName + "' /f 2>$null } catch {}", 10000); } catch(e) {}
    try { fs.unlinkSync(batchPath); } catch(e) {}

    if (!log) log = result.output || result.error || "no output";
    const exitMatch = log.match(/EXIT_CODE:(\d+)/);
    const exitCode = exitMatch ? parseInt(exitMatch[1]) : -1;

    return { success: exitCode === 0, exitCode, log, psOutput: result.output, psError: result.error };
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
    const r = await this._runScewinElevated(`/i /s "${SCEWIN_EXPORT_PATH}" /n "${name}" /v ${hexValue}`, 30000);

    // Log the change
    this._logChange(name, current.value, hexValue);

    if (!r.success) return { success: false, error: r.log || `Exit code: ${r.exitCode}` };

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

    // Fuzzy find helper with exclude support (same as _discoverBiosSettings)
    const fuzzyGet = (keywords, excludes = [], ...hardcoded) => {
      // Try hardcoded names first
      for (const name of hardcoded) {
        if (s[name]) return s[name];
      }
      // Then fuzzy search with excludes
      for (const kw of keywords) {
        const kwLower = kw.toLowerCase();
        for (const name of allNames) {
          const lower = name.toLowerCase();
          if (!lower.includes(kwLower)) continue;
          let excluded = false;
          for (const ex of excludes) {
            if (lower.includes(ex.toLowerCase())) { excluded = true; break; }
          }
          if (!excluded && s[name]) return s[name];
        }
      }
      return null;
    };

    const cpuSettings = {
      cpuRatio: fuzzyGet([
        "all-core ratio limit", "all core ratio limit", "performance core ratio",
        "adjust cpu ratio", "per core ratio limit", "cpu clock ratio",
        "all core ratio", "cpu core ratio", "core ratio limit",
        "p-core ratio", "pcore ratio", "processor core ratio",
        "cpu ratio", "core ratio", "frequency ratio",
        "ratio limit", "multiplier", "cpu multi", "oc ratio", "max ratio",
      ], ["igp", "gt ratio", "ring", "uncore", "memory", "dram", "cache", "e-core", "ecore"],
        "CPU Ratio", "CPU Core Ratio", "Core Ratio Limit", "All Core Ratio Limit"),
      cpuVoltage: fuzzyGet([
        "cpu core/cache voltage", "cpu core voltage override",
        "vcore override voltage", "cpu vcore override", "dynamic vcore",
        "core voltage override", "cpu core voltage", "cpu cache voltage",
        "vcore override", "vcore voltage", "ia voltage offset",
        "cpu vcore", "core voltage", "override voltage", "adaptive voltage",
        "vcore", "cpu voltage",
      ], ["loadline", "calibration", "llc", "load line", "mode", "svid", "ratio", "ring", "sa ", "dram", "gt ", "pll", "io "],
        "CPU Core Voltage", "CPU Vcore", "Vcore Override", "CPU Core/Cache Voltage"),
      cpuVoltageMode: fuzzyGet([
        "cpu core/cache voltage mode", "svid behavior",
        "cpu core voltage mode", "cpu vcore mode",
        "voltage mode", "vcore mode", "svid support",
      ], ["loadline", "calibration", "llc", "ratio", "dram"],
        "CPU Core Voltage Mode", "Vcore Mode", "SVID Behavior"),
      powerLimit1: fuzzyGet([
        "long duration package power limit", "long duration power limit",
        "package power limit1", "package power limit 1",
        "power limit 1 value", "cpu power limit value",
        "processor base power", "base power limit",
        "power limit 1", "long duration", "pl1 (w)", "pl1(w)", "pl1",
        "tdp limit", "tdp power",
      ], [], "Long Duration Power Limit", "PL1", "Package Power Limit 1"),
      powerLimit2: fuzzyGet([
        "short duration package power limit", "short duration power limit",
        "package power limit2", "package power limit 2",
        "power limit 2 value", "maximum turbo power",
        "max turbo power", "power limit 2", "short duration",
        "pl2 (w)", "pl2(w)", "pl2", "turbo power limit",
      ], [], "Short Duration Power Limit", "PL2", "Package Power Limit 2"),
      tccOffset: fuzzyGet([
        "tcc activation offset", "cpu tcc offset", "tcc offset", "tcc",
      ], [], "TCC Activation Offset"),
      ringRatio: fuzzyGet([
        "min cpu cache ratio", "max cpu cache ratio", "cpu cache ratio",
        "adjust ring ratio", "ring multiplier", "cpu uncore ratio",
        "ring ratio", "cache ratio", "uncore ratio", "uncore frequency",
      ], ["igp", "cpu core", "cpu ratio"], "Ring Ratio", "Cache Ratio", "Uncore Ratio"),
      iccMax: fuzzyGet([
        "iccmax unlimited", "icc max override", "ia ac/dc loadline",
        "ia ac loadline", "cpu vr current limit", "current limit",
        "ia ac load line", "icc max", "iccmax", "icc",
      ], [], "ICC Max", "IA AC Load Line"),
      avxOffset: fuzzyGet([
        "avx instruction core ratio", "avx2 ratio offset",
        "avx ratio offset", "avx frequency trim",
        "avx 512 offset", "avx-512 offset", "avx512 ratio offset",
        "avx offset", "avx2 offset", "avx negative offset",
      ], [], "AVX Offset", "AVX2 Ratio Offset"),
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

    // Fuzzy find helper with exclude support
    const fuzzyGet = (keywords, excludes = [], ...hardcoded) => {
      for (const name of hardcoded) {
        if (s[name]) return s[name];
      }
      for (const kw of keywords) {
        const kwLower = kw.toLowerCase();
        for (const name of allNames) {
          const lower = name.toLowerCase();
          if (!lower.includes(kwLower)) continue;
          let excluded = false;
          for (const ex of excludes) {
            if (lower.includes(ex.toLowerCase())) { excluded = true; break; }
          }
          if (!excluded && s[name]) return s[name];
        }
      }
      return null;
    };

    const memSettings = {
      memorySpeed: fuzzyGet([
        "memory frequency", "dram frequency", "memory speed", "mem freq",
        "system memory multiplier", "memory multiplier", "dram speed",
        "memory clock", "dram clock", "memory ratio",
      ], [], "Memory Frequency", "DRAM Frequency", "Memory Speed", "System Memory Multiplier"),
      casLatency: fuzzyGet([
        "cas latency", "cas# latency", "tcl", "cl value",
        "dram cas# latency", "cas",
      ], [], "CAS Latency", "tCL", "CAS# Latency", "DRAM CAS# Latency"),
      tRCD: fuzzyGet([
        "trcd", "ras to cas", "ras# to cas#", "ras to cas delay",
        "dram ras# to cas# delay", "row address to column address",
      ], [], "tRCD", "RAS to CAS Delay", "RAS# to CAS# Delay", "DRAM RAS# to CAS# Delay"),
      tRP: fuzzyGet([
        "trp", "row precharge", "ras# precharge", "ras precharge",
        "dram ras# pre time",
      ], [], "tRP", "Row Precharge Time", "RAS# Precharge", "DRAM RAS# PRE Time"),
      tRAS: fuzzyGet([
        "tras", "ras active", "active to precharge", "ras# act time",
        "dram ras# act time",
      ], [], "tRAS", "RAS Active Time", "Active to Precharge Delay", "DRAM RAS# ACT Time"),
      xmpProfile: fuzzyGet([
        "ai overclock tuner", "a-xmp", "extreme memory profile",
        "load xmp setting", "memory profile", "xmp profile",
        "system memory multiplier", "d.o.c.p", "docp",
        "expo profile", "expo", "xmp", "dram profile",
      ], [], "XMP Profile", "Extreme Memory Profile", "XMP", "Ai Overclock Tuner", "A-XMP"),
      memoryVoltage: fuzzyGet([
        "dram voltage", "memory voltage", "dram v", "dimm voltage",
        "dram core voltage", "ddr voltage", "memory vddq",
        "vddq voltage", "sa voltage", "system agent voltage",
      ], [], "DRAM Voltage", "Memory Voltage", "DRAM Core Voltage"),
      commandRate: fuzzyGet([
        "command rate", "cmd rate", "command rate mode",
        "dram command rate", "cr mode",
      ], [], "Command Rate", "CR", "Cmd Rate", "DRAM Command Rate"),
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

    const r = await this._runScewinElevated(`/o /s "${SCEWIN_BACKUP_PATH}"`, 30000);
    if (!r.success) return { success: false, error: r.log || `Exit code: ${r.exitCode}` };

    return { success: true, backupPath: SCEWIN_BACKUP_PATH, timestamp: new Date().toISOString() };
  }

  async restoreBackup() {
    if (!this.available) return { success: false, error: "SceWin not found" };
    if (!fs.existsSync(SCEWIN_BACKUP_PATH)) {
      return { success: false, error: "No backup found to restore" };
    }

    const r = await this._runScewinElevated(`/i /s "${SCEWIN_BACKUP_PATH}"`, 30000);
    if (!r.success) return { success: false, error: r.log || `Exit code: ${r.exitCode}` };

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

  // ── Reboot-and-Resume State Management ────────────────────────────
  // SceWin writes BIOS NVRAM but changes only take effect after reboot.
  // We save the current OC state to disk, register the app to auto-start
  // after reboot, and on next launch detect the saved state and resume.
  // ────────────────────────────────────────────────────────────────────

  _saveOcState(state) {
    try {
      const dir = path.dirname(AI_OC_STATE_PATH);
      if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
      const data = {
        ...state,
        savedAt: new Date().toISOString(),
        log: this.log, // preserve log across reboots
      };
      fs.writeFileSync(AI_OC_STATE_PATH, JSON.stringify(data, null, 2));
      return true;
    } catch (e) { return false; }
  }

  _loadOcState() {
    try {
      if (!fs.existsSync(AI_OC_STATE_PATH)) return null;
      const data = JSON.parse(fs.readFileSync(AI_OC_STATE_PATH, "utf8"));
      // State expires after 2 hours (in case user never rebooted)
      const savedAt = new Date(data.savedAt).getTime();
      if (Date.now() - savedAt > 2 * 60 * 60 * 1000) {
        this._clearOcState();
        return null;
      }
      return data;
    } catch (e) { return null; }
  }

  _clearOcState() {
    try { fs.unlinkSync(AI_OC_STATE_PATH); } catch (e) {}
    // Remove the auto-start scheduled task
    this._unregisterAutoStart().catch(() => {});
  }

  async _registerAutoStart() {
    // Register a Windows scheduled task to launch our app at next logon
    // This ensures we resume after reboot
    const appPath = process.execPath; // Electron exe path
    const psScript = [
      `try { schtasks /delete /tn '${AI_OC_AUTOSTART_TASK}' /f 2>$null } catch {}`,
      `schtasks /create /tn '${AI_OC_AUTOSTART_TASK}' /tr '"${appPath}" --resume-oc' /sc onlogon /rl HIGHEST /f`,
    ].join("; ");
    const result = await runPS(psScript, 15000);
    return result.success;
  }

  async _unregisterAutoStart() {
    await runPS(`try { schtasks /delete /tn '${AI_OC_AUTOSTART_TASK}' /f 2>$null } catch {}`, 10000);
  }

  async _rebootAndResume(state) {
    this._log("testing", "reboot-save", `Saving OC state before reboot: phase=${state.resumePhase}, ratio=x${state.currentRatio || "?"}, voltage=${state.currentVoltage || "auto"}`);
    this._saveOcState(state);
    await this._registerAutoStart();

    this._log("testing", "reboot-trigger", "Initiating system reboot in 10 seconds... BIOS changes will take effect after restart.");

    // Give the user 10 seconds to see the message
    await new Promise(r => setTimeout(r, 10000));

    // Trigger Windows reboot
    const rebootResult = await runPS(
      "Start-Process -FilePath 'shutdown.exe' -ArgumentList '/r /t 5 /c \"FN Optimizer: Rebooting to apply BIOS overclock settings...\"' -Verb RunAs",
      15000
    );

    if (!rebootResult.success) {
      this._log("testing", "reboot-fail", "Could not initiate reboot automatically. Please reboot manually and relaunch the app to continue the AI OC process.");
      return false;
    }
    return true;
  }

  hasResumeState() {
    return this._loadOcState() !== null;
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

      // DIAGNOSTIC: Dump ALL sensor labels with BOTH raw and value strings
      // The Value string contains the unit (e.g. "50.0 °C") which is critical
      // for correct sensor type identification
      if (hwinfoSensors.sensors.allSensors) {
        const allLabels = Object.entries(hwinfoSensors.sensors.allSensors);
        this._log("analyzing", "hwinfo-sensor-count", `Total HWiNFO sensors: ${allLabels.length}`);

        // Log ALL temperature-related sensors specifically (show Value AND Raw)
        const tempRelated = allLabels.filter(([lbl, data]) => {
          const l = lbl.toLowerCase();
          const v = (data.value || "").toString();
          return l.includes("temp") || l.includes("cpu") || l.includes("core") ||
                 l.includes("package") || l.includes("tctl") || l.includes("tdie") ||
                 v.includes("°C") || v.includes("deg");
        });
        this._log("analyzing", "hwinfo-temp-sensors", `Temperature-related sensors (${tempRelated.length}):`);
        for (const [lbl, data] of tempRelated) {
          this._log("analyzing", "hwinfo-temp", `  "${lbl}" → Value="${data.value}" | Raw="${data.raw}"`);
        }

        // Log other interesting sensors in chunks
        const other = allLabels.filter(([lbl]) => {
          const l = lbl.toLowerCase();
          return l.includes("gpu") || l.includes("clock") || l.includes("volt") ||
                 l.includes("power") || l.includes("hot") || l.includes("fan");
        });
        for (let i = 0; i < other.length; i += 5) {
          const chunk = other.slice(i, i + 5)
            .map(([lbl, data]) => `"${lbl}"=Val:"${data.value}"|Raw:"${data.raw}"`)
            .join(" | ");
          this._log("analyzing", "hwinfo-sensors", `Other [${i}]: ${chunk}`);
        }
      }
    } else {
      this._log("analyzing", "hwinfo-fail", `HWiNFO sensor read failed: ${hwinfoSensors.error || "unknown"} — temps will use WMI fallback`);
    }

    // ═══════════════════════════════════════════════════════════════════
    // RATIO DETECTION: For modern Intel, we need ACTUAL boost clock, not base
    //
    // i9-14900K: Base = 3.2GHz (x32), P-Core Turbo = 5.6GHz (x56), Turbo Max 3.0 = 6.0GHz (x60)
    // The "base ratio" from WMI CurrentClockSpeed is often the BASE, not boost.
    // HWiNFO shows ACTUAL current clock which could be anywhere from idle to boost.
    //
    // For OC, we need to know:
    //  - Stock max turbo ratio (from CPU spec, not current measurement)
    //  - Current actual clock under load (from HWiNFO during stress)
    //  - What BIOS has set (from SceWin export)
    //
    // Intel 14th gen (Raptor Lake Refresh) max boost clocks:
    //   i9-14900K/KF: P-core 5.6GHz (x56), E-core 4.4GHz (x44)
    //   i9-14900KS: P-core 6.0GHz (x60)
    //   i7-14700K/KF: P-core 5.5GHz (x55)
    //   i5-14600K/KF: P-core 5.3GHz (x53)
    // Intel 13th gen:
    //   i9-13900K/KF/KS: P-core 5.4-5.8GHz
    //   i7-13700K/KF: P-core 5.3-5.4GHz
    // ═══════════════════════════════════════════════════════════════════

    // Detect stock max turbo from CPU model name
    const cpuName = (hw.cpu.name || "").toUpperCase();
    let stockMaxTurbo = 0;

    // Intel 14th gen
    if (cpuName.includes("14900KS")) stockMaxTurbo = 60;
    else if (cpuName.includes("14900K")) stockMaxTurbo = 56;
    else if (cpuName.includes("14700K")) stockMaxTurbo = 55;
    else if (cpuName.includes("14600K")) stockMaxTurbo = 53;
    // Intel 13th gen
    else if (cpuName.includes("13900KS")) stockMaxTurbo = 58;
    else if (cpuName.includes("13900K")) stockMaxTurbo = 54;
    else if (cpuName.includes("13700K")) stockMaxTurbo = 54;
    else if (cpuName.includes("13600K")) stockMaxTurbo = 53;
    // Intel 12th gen
    else if (cpuName.includes("12900K")) stockMaxTurbo = 52;
    else if (cpuName.includes("12700K")) stockMaxTurbo = 50;
    else if (cpuName.includes("12600K")) stockMaxTurbo = 49;
    // AMD Ryzen 7000/9000
    else if (cpuName.includes("9950X")) stockMaxTurbo = 57;
    else if (cpuName.includes("9900X")) stockMaxTurbo = 56;
    else if (cpuName.includes("9700X")) stockMaxTurbo = 53;
    else if (cpuName.includes("9600X")) stockMaxTurbo = 53;
    else if (cpuName.includes("7950X")) stockMaxTurbo = 57;
    else if (cpuName.includes("7900X")) stockMaxTurbo = 56;
    else if (cpuName.includes("7800X3D")) stockMaxTurbo = 50;
    else if (cpuName.includes("7700X")) stockMaxTurbo = 54;
    else if (cpuName.includes("7600X")) stockMaxTurbo = 53;

    // Determine base ratio from WMI (base clock, NOT boost)
    const wmiBaseClock = hw.cpu.maxClock || 3200; // WMI reports base clock
    const wmiBaseRatio = Math.round(wmiBaseClock / 100);

    // Use stock turbo if we identified the CPU, otherwise estimate
    if (stockMaxTurbo > 0) {
      analysis.baseRatio = stockMaxTurbo; // Start from stock turbo, not base
      analysis.stockTurboRatio = stockMaxTurbo;
      analysis.wmiBaseRatio = wmiBaseRatio;
      this._log("analyzing", "ratio-source", `Stock max turbo for ${hw.cpu.name}: x${stockMaxTurbo} (${stockMaxTurbo * 100}MHz) | Base: x${wmiBaseRatio} (${wmiBaseClock}MHz)`);
    } else if (biosRatio && biosRatio >= 10 && biosRatio <= 80) {
      analysis.baseRatio = biosRatio;
      analysis.stockTurboRatio = biosRatio;
      this._log("analyzing", "ratio-source", `Using BIOS-reported CPU ratio: x${biosRatio}`);
    } else if (hwinfoClockMHz && hwinfoClockMHz > 500) {
      analysis.baseRatio = Math.round(hwinfoClockMHz / 100);
      analysis.stockTurboRatio = analysis.baseRatio;
      this._log("analyzing", "ratio-source", `Using HWiNFO clock-derived ratio: x${analysis.baseRatio} (from ${hwinfoClockMHz} MHz)`);
    } else {
      analysis.baseRatio = wmiBaseRatio;
      analysis.stockTurboRatio = wmiBaseRatio;
      this._log("analyzing", "ratio-source", `Using WMI-estimated ratio: x${analysis.baseRatio} (from ${wmiBaseClock} MHz)`);
    }

    // Set ceiling: for unlocked Intel K-series, try 1-3 ratios above stock turbo
    if (analysis.isIntel) {
      analysis.maxTurboRatio = analysis.baseRatio;
      analysis.safeCeilingRatio = analysis.cpuUnlocked ? analysis.baseRatio + 3 : analysis.baseRatio;
    } else if (analysis.isAMD) {
      analysis.maxTurboRatio = analysis.baseRatio;
      analysis.safeCeilingRatio = analysis.cpuUnlocked ? analysis.baseRatio + 2 : analysis.baseRatio;
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
  // Strategy: Start from stock max turbo (e.g. x56 for i9-14900K).
  // Step 1: Find MAX STABLE RATIO by going up 1 multiplier at a time.
  //         Each step: write BIOS via SceWin → reboot → stress test → check stability.
  // Step 2: At the max stable ratio, find the LOWEST stable voltage by
  //         stepping down from auto/default in 10mV increments (undervolt).
  //         This reduces heat and power while keeping the higher clock.
  //
  // BIOS changes via SceWin require a system reboot to take effect.
  // The reboot-and-resume system saves state to disk, registers an
  // auto-start task, and resumes automatically after each reboot.
  // ────────────────────────────────────────────────────────────────────

  async _cpuOcPhase(analysis, resumeState = null) {
    if (!this.scewin.available) {
      this._log("testing", "skip-cpu", "SceWin not available — skipping CPU overclock phase");
      return null;
    }

    if (!analysis.cpuUnlocked) {
      this._log("testing", "skip-cpu", "CPU is locked — skipping CPU overclock phase");
      return null;
    }

    const ratioSettingName = this.biosMap && this.biosMap.cpuRatio ? this.biosMap.cpuRatio : null;
    const voltageSettingName = this.biosMap && this.biosMap.cpuVoltage ? this.biosMap.cpuVoltage : null;
    const voltageModeSettingName = this.biosMap && this.biosMap.cpuVoltageMode ? this.biosMap.cpuVoltageMode : null;

    if (!ratioSettingName) {
      this._log("testing", "skip-cpu", "BIOS discovery did not find a CPU ratio setting — cannot safely overclock CPU. Check SceWin export for your motherboard's setting names.");
      return null;
    }

    this._setPhase("testing", 20);
    this._log("testing", "cpu-start", `CPU OC Phase — ratio: "${ratioSettingName}" | voltage: "${voltageSettingName || "NOT FOUND"}" | voltage mode: "${voltageModeSettingName || "NOT FOUND"}"`);

    // Backup current BIOS settings before making any changes
    await this.scewin.backupSettings();
    this._log("testing", "backup", "BIOS settings backed up before OC");

    // Determine starting point and ceiling
    const stockTurbo = analysis.stockTurboRatio || analysis.baseRatio;
    const maxCeiling = analysis.safeCeilingRatio;

    // If resuming from reboot, pick up where we left off
    let startRatio = stockTurbo;
    let lastGoodRatio = stockTurbo;
    let lastGoodVoltage = null;
    let subPhase = "ratio-hunt"; // ratio-hunt | voltage-optimize | done
    let needsReboot = false; // track if we just wrote BIOS and need reboot

    if (resumeState && resumeState.resumePhase === "cpu") {
      startRatio = resumeState.testingRatio || stockTurbo;
      lastGoodRatio = resumeState.lastGoodRatio || stockTurbo;
      lastGoodVoltage = resumeState.lastGoodVoltage || null;
      subPhase = resumeState.subPhase || "ratio-hunt";
      needsReboot = false; // we just rebooted, so BIOS changes are active

      this._log("testing", "cpu-resume", `Resumed after reboot — testing x${startRatio}, last good x${lastGoodRatio}, sub-phase: ${subPhase}`);

      // After reboot, the BIOS changes are now active — run stress test immediately
      if (subPhase === "ratio-hunt") {
        this._log("testing", "cpu-post-reboot-stress", `BIOS changes active after reboot — stress testing x${startRatio} (${startRatio * 100} MHz) for 10 minutes...`);
        const postRebootStress = await this.runExtendedStressTest(10, { maxTemp: 95 });

        if (postRebootStress.stable) {
          lastGoodRatio = startRatio;
          this._log("testing", "cpu-stable", `x${startRatio} is STABLE after reboot (max temp: ${postRebootStress.maxTemp}C, avg: ${postRebootStress.avgTemp}C)`, {
            stable: true, temp: postRebootStress.maxTemp, clock: startRatio * 100,
          });
          startRatio++; // move to next ratio for continued testing
        } else {
          this._log("testing", "cpu-unstable", `x${startRatio} is UNSTABLE after reboot — rolling back to x${lastGoodRatio}`, {
            stable: false, temp: postRebootStress.maxTemp, errors: postRebootStress.errors,
          });
          // Write the last good ratio back to BIOS
          if (lastGoodRatio !== startRatio) {
            await this.scewin.applyCpuMultiplier(lastGoodRatio, ratioSettingName);
            // Need another reboot for the rollback, but skip stress test — we already know it's stable
            needsReboot = true;
          }
          subPhase = voltageSettingName ? "voltage-optimize" : "done";
        }
      } else if (subPhase === "voltage-optimize") {
        // We rebooted after a voltage change — stress test it
        const vTest = resumeState.testingVoltage;
        this._log("testing", "cpu-voltage-post-reboot", `Stress testing voltage ${vTest}V at x${lastGoodRatio} after reboot...`);
        const vStress = await this.runExtendedStressTest(10, { maxTemp: 95 });

        if (vStress.stable) {
          lastGoodVoltage = vTest;
          this._log("testing", "cpu-voltage-stable", `${vTest}V at x${lastGoodRatio} is STABLE — trying lower...`, {
            stable: true, temp: vStress.maxTemp, voltage: vTest, clock: lastGoodRatio * 100,
          });
        } else {
          this._log("testing", "cpu-voltage-unstable", `${vTest}V at x${lastGoodRatio} is UNSTABLE — previous voltage was the minimum`, {
            stable: false, temp: vStress.maxTemp, voltage: vTest,
          });
          // Roll back to last good voltage
          if (lastGoodVoltage && lastGoodVoltage !== vTest) {
            await this.scewin.applyCpuVoltage(lastGoodVoltage, voltageSettingName);
            needsReboot = true;
          }
          subPhase = "done";
        }
      }
    }

    // ═══════════════════════════════════════════════════════════════════
    // PHASE 1: RATIO HUNT — Find the highest stable CPU multiplier
    // Start from stock turbo and go up. Each step requires a reboot
    // because SceWin NVRAM writes only take effect after restart.
    // ═══════════════════════════════════════════════════════════════════
    if (subPhase === "ratio-hunt") {
      const totalSteps = maxCeiling - stockTurbo;
      this._log("testing", "ratio-hunt-start", `Ratio hunt: stock turbo x${stockTurbo} → ceiling x${maxCeiling} (${totalSteps} steps)`);

      for (let ratio = startRatio; ratio <= maxCeiling; ratio++) {
        if (this._stopRequested) break;

        const stepProgress = 20 + ((ratio - stockTurbo) / Math.max(totalSteps, 1)) * 25;
        this._setPhase("testing", Math.round(stepProgress));

        this._log("testing", "cpu-step", `Setting CPU multiplier x${ratio} (${ratio * 100} MHz) via SceWin...`, {
          clock: ratio * 100,
        });

        // Write the new ratio to BIOS NVRAM
        const applyResult = await this.scewin.applyCpuMultiplier(ratio, ratioSettingName);
        if (!applyResult.success) {
          this._log("testing", "cpu-apply-error", `Failed to write ratio x${ratio} to BIOS: ${applyResult.error}`);
          break;
        }

        this._log("testing", "cpu-written", `Ratio x${ratio} written to BIOS NVRAM — system reboot required for changes to take effect`);

        // Save state and trigger reboot
        const saved = await this._rebootAndResume({
          resumePhase: "cpu",
          subPhase: "ratio-hunt",
          testingRatio: ratio,
          lastGoodRatio,
          lastGoodVoltage,
          analysis: {
            stockTurboRatio: analysis.stockTurboRatio,
            safeCeilingRatio: analysis.safeCeilingRatio,
            safeCeilingVoltage: analysis.safeCeilingVoltage,
            cpuUnlocked: analysis.cpuUnlocked,
            isIntel: analysis.isIntel,
            isAMD: analysis.isAMD,
            coolingQuality: analysis.coolingQuality,
            cpu: analysis.cpu,
            gpu: analysis.gpu,
            ram: analysis.ram,
          },
          biosMap: this.biosMap,
        });

        if (saved) {
          // Reboot initiated — execution stops here. Will resume after reboot.
          return { pending: true, message: "Rebooting to apply BIOS changes..." };
        } else {
          // Reboot failed — fall back to testing without reboot (less accurate)
          this._log("testing", "reboot-fallback", "Reboot failed — testing at current clocks (BIOS changes may not be active yet)");
          const stress = await this.runExtendedStressTest(10, { maxTemp: 95 });
          if (stress.stable) {
            lastGoodRatio = ratio;
            this._log("testing", "cpu-stable", `x${ratio} stress test passed (note: BIOS change may not be active without reboot)`, {
              stable: true, temp: stress.maxTemp, clock: ratio * 100,
            });
          } else {
            this._log("testing", "cpu-unstable", `x${ratio} stress test failed — stopping ratio hunt`, {
              stable: false, temp: stress.maxTemp,
            });
            // Roll back to last good
            await this.scewin.applyCpuMultiplier(lastGoodRatio, ratioSettingName);
            break;
          }
        }
      }

      // Move to voltage optimization if we have a voltage setting
      subPhase = voltageSettingName ? "voltage-optimize" : "done";
    }

    // ═══════════════════════════════════════════════════════════════════
    // PHASE 2: VOLTAGE OPTIMIZATION — Find the lowest stable voltage
    // for the max stable ratio found above. This reduces heat/power.
    //
    // Strategy: Start from a safe default voltage and step DOWN in
    // 10mV increments until instability is found. The last stable
    // voltage is the optimal one (lowest safe voltage for the clock).
    //
    // For Intel K-series: Start around 1.30V, step down to ~1.10V
    // For AMD Ryzen: Start around 1.25V, step down to ~1.05V
    // ═══════════════════════════════════════════════════════════════════
    if (subPhase === "voltage-optimize" && voltageSettingName && !this._stopRequested) {
      this._setPhase("testing", 50);

      // Determine voltage range based on CPU type
      const voltStart = analysis.isAMD ? 1.25 : 1.30; // safe starting voltage
      const voltFloor = analysis.isAMD ? 1.05 : 1.10;  // absolute minimum to try
      const voltStep = 0.010; // 10mV steps

      // If resuming from a voltage test, pick up from where we stopped
      let currentVoltage = voltStart;
      if (resumeState && resumeState.subPhase === "voltage-optimize" && resumeState.testingVoltage) {
        // We already tested testingVoltage and it was stable (since we continued)
        currentVoltage = resumeState.testingVoltage - voltStep;
      }

      this._log("testing", "voltage-start", `Voltage optimization at x${lastGoodRatio}: starting at ${voltStart}V, stepping down by ${voltStep * 1000}mV to find minimum stable voltage`);

      // First, set the voltage mode to "Override" / "Manual" if we can
      if (voltageModeSettingName) {
        this._log("testing", "voltage-mode", `Setting voltage mode to manual via "${voltageModeSettingName}"...`);
        // Most BIOS: 0x0 = Auto, 0x1 = Manual/Override, 0x2 = Offset
        await this.scewin.writeSetting(voltageModeSettingName, 1);
      }

      if (!lastGoodVoltage) lastGoodVoltage = voltStart;

      for (let v = currentVoltage; v >= voltFloor; v -= voltStep) {
        if (this._stopRequested) break;

        const vRounded = Math.round(v * 1000) / 1000;
        const stepPct = 50 + ((voltStart - vRounded) / (voltStart - voltFloor)) * 10;
        this._setPhase("testing", Math.round(stepPct));

        this._log("testing", "cpu-voltage-step", `Writing ${vRounded}V to BIOS for x${lastGoodRatio}...`);

        const vResult = await this.scewin.applyCpuVoltage(vRounded, voltageSettingName);
        if (!vResult.success) {
          this._log("testing", "cpu-voltage-write-fail", `Failed to write voltage ${vRounded}V: ${vResult.error}`);
          continue;
        }

        // Save state and reboot to apply voltage change
        const saved = await this._rebootAndResume({
          resumePhase: "cpu",
          subPhase: "voltage-optimize",
          testingVoltage: vRounded,
          lastGoodRatio,
          lastGoodVoltage,
          analysis: {
            stockTurboRatio: analysis.stockTurboRatio,
            safeCeilingRatio: analysis.safeCeilingRatio,
            safeCeilingVoltage: analysis.safeCeilingVoltage,
            cpuUnlocked: analysis.cpuUnlocked,
            isIntel: analysis.isIntel,
            isAMD: analysis.isAMD,
            coolingQuality: analysis.coolingQuality,
            cpu: analysis.cpu,
            gpu: analysis.gpu,
            ram: analysis.ram,
          },
          biosMap: this.biosMap,
        });

        if (saved) {
          return { pending: true, message: `Rebooting to test ${vRounded}V...` };
        } else {
          // Reboot failed — test without reboot (inaccurate)
          this._log("testing", "reboot-fallback", "Reboot failed — testing voltage without reboot");
          const vStress = await this.runExtendedStressTest(10, { maxTemp: 95 });
          if (vStress.stable) {
            lastGoodVoltage = vRounded;
            this._log("testing", "cpu-voltage-stable", `${vRounded}V STABLE (reboot-less test)`, {
              stable: true, temp: vStress.maxTemp, voltage: vRounded,
            });
          } else {
            this._log("testing", "cpu-voltage-floor", `${vRounded}V UNSTABLE — minimum voltage is ${lastGoodVoltage}V`, {
              stable: false, voltage: vRounded, minVoltage: lastGoodVoltage,
            });
            // Roll back to last good voltage
            await this.scewin.applyCpuVoltage(lastGoodVoltage, voltageSettingName);
            break;
          }
        }
      }
    }

    // If we need a final reboot to ensure last-good settings are active
    if (needsReboot && !this._stopRequested) {
      this._log("testing", "final-reboot", `Applying final settings: x${lastGoodRatio} at ${lastGoodVoltage || "auto"}V — one last reboot...`);
      await this.scewin.applyCpuMultiplier(lastGoodRatio, ratioSettingName);
      if (lastGoodVoltage && voltageSettingName) {
        await this.scewin.applyCpuVoltage(lastGoodVoltage, voltageSettingName);
      }
      // Don't reboot again for this — let the final validation handle it
    }

    const gainMHz = (lastGoodRatio - stockTurbo) * 100;
    this._log("testing", "cpu-complete", `CPU OC complete: x${lastGoodRatio} (${lastGoodRatio * 100} MHz) at ${lastGoodVoltage || "auto"}V — ${gainMHz >= 0 ? "+" : ""}${gainMHz} MHz vs stock turbo`, {
      finalRatio: lastGoodRatio, finalVoltage: lastGoodVoltage, gainMHz,
    });

    return {
      finalRatio: lastGoodRatio,
      finalVoltage: lastGoodVoltage,
      finalMHz: lastGoodRatio * 100,
      gainMHz,
      stockTurboMHz: stockTurbo * 100,
    };
  }

  // ── Memory Phase ──────────────────────────────────────────────────
  // Strategy:
  //  1. Check if XMP/EXPO/DOCP is enabled. If not, enable it.
  //  2. Verify RAM is running at its rated speed (e.g. DDR5-6000 = 6000 MHz).
  //  3. If BIOS timing settings are available, try tightening primary timings
  //     (tCL, tRCD, tRP, tRAS) one step at a time with stability testing.
  //  4. Memory changes via SceWin also require reboot to take effect.
  // ────────────────────────────────────────────────────────────────────

  async _memoryPhase(resumeState = null) {
    if (!this.scewin.available) {
      this._log("testing", "skip-mem", "SceWin not available — skipping memory phase");
      return null;
    }

    this._setPhase("testing", 55);
    this._log("testing", "mem-start", "Starting memory optimization phase");

    // Check current memory settings from BIOS
    const memSettings = await this.scewin.getMemorySettings();
    if (!memSettings.success) {
      this._log("testing", "mem-error", "Could not read memory settings: " + memSettings.error);
      return null;
    }

    // Read actual RAM info from Windows
    const ramInfo = await this.hw.detectRAM();
    const ratedSpeed = ramInfo.ratedSpeed || 0;
    const currentSpeed = ramInfo.currentSpeed || 0;
    const stickCount = ramInfo.sticks ? ramInfo.sticks.length : 0;

    this._log("testing", "mem-info", `RAM: ${ramInfo.totalGB}GB (${stickCount} sticks) | Rated: ${ratedSpeed} MHz | Current: ${currentSpeed} MHz | Part: ${ramInfo.sticks?.[0]?.partNumber || "unknown"}`);

    const xmpSettingName = this.biosMap && this.biosMap.xmpProfile ? this.biosMap.xmpProfile : null;
    const result = {
      xmpEnabled: false,
      xmpAlreadyEnabled: false,
      ratedSpeed,
      currentSpeed,
      timingsTightened: false,
      timingsChanged: [],
      stable: true,
    };

    // ── Step 1: Ensure XMP / EXPO / DOCP is enabled ──
    const xmpSetting = memSettings.settings.xmpProfile;
    if (xmpSetting) {
      const xmpValue = xmpSetting.value || "0x0";
      const isDisabled = xmpValue === "0x0" || xmpValue === "0x00" || (xmpSetting.currentLabel || "").toLowerCase().includes("disabled") || (xmpSetting.currentLabel || "").toLowerCase().includes("auto");

      if (isDisabled) {
        this._log("testing", "mem-xmp-off", `XMP/EXPO is currently disabled (value: ${xmpValue}, label: "${xmpSetting.currentLabel || ""}")${xmpSettingName ? ` — enabling via "${xmpSettingName}"` : ""}`);

        // Find profile 1 option value from the setting's options
        let enableValue = 1;
        if (xmpSetting.options && xmpSetting.options.length > 1) {
          // Profile 1 is usually the second option (first is Disabled/Auto)
          const profile1 = xmpSetting.options.find(o =>
            (o.label || "").toLowerCase().includes("profile 1") ||
            (o.label || "").toLowerCase().includes("xmp 1") ||
            (o.label || "").toLowerCase().includes("expo 1")
          ) || xmpSetting.options[1]; // fallback to second option
          if (profile1) {
            enableValue = parseInt(profile1.value, 16);
            this._log("testing", "mem-xmp-profile", `Using XMP option: "${profile1.label}" (value: ${profile1.value})`);
          }
        }

        const xmpResult = await this.scewin.applyMemoryXMP(enableValue, xmpSettingName);
        if (xmpResult.success) {
          this._log("testing", "mem-xmp-written", "XMP/EXPO profile written to BIOS NVRAM — requires reboot to take effect");
          result.xmpEnabled = true;

          // Note: XMP change requires reboot. If we're also doing CPU OC with reboots,
          // the XMP change will be picked up on the next reboot. Otherwise we'd need
          // to trigger one here. For now, log it and continue — the final validation
          // or next CPU OC reboot will activate it.
          this._log("testing", "mem-xmp-note", "XMP will be active after next system reboot. If CPU OC didn't trigger a reboot, memory may still be at JEDEC speeds.");
        } else {
          this._log("testing", "mem-xmp-fail", `Failed to enable XMP: ${xmpResult.error}`);
        }
      } else {
        this._log("testing", "mem-xmp-ok", `XMP/EXPO is already enabled (value: ${xmpValue}, label: "${xmpSetting.currentLabel || ""}")`);
        result.xmpEnabled = true;
        result.xmpAlreadyEnabled = true;
      }
    } else {
      this._log("testing", "mem-no-xmp", "No XMP/EXPO/DOCP setting found in BIOS export — cannot check or enable memory profiles");
    }

    // ── Step 2: Verify memory speed matches rated ──
    if (ratedSpeed > 0 && currentSpeed > 0) {
      const speedRatio = currentSpeed / ratedSpeed;
      if (speedRatio < 0.95) {
        this._log("testing", "mem-speed-mismatch", `WARNING: RAM running at ${currentSpeed} MHz but rated for ${ratedSpeed} MHz (${Math.round(speedRatio * 100)}%). XMP may not be active or may need a reboot.`, {
          currentSpeed, ratedSpeed, ratio: speedRatio,
        });
      } else {
        this._log("testing", "mem-speed-ok", `RAM speed verified: ${currentSpeed} MHz (rated ${ratedSpeed} MHz)`);
      }
    }

    // ── Step 3: Try tightening primary timings ──
    // Only attempt if we have the BIOS settings AND the BIOS change log shows
    // we're comfortable with SceWin writes working on this system
    const ms = memSettings.settings;
    const hasTimingSettings = ms.casLatency || ms.tRCD || ms.tRP || ms.tRAS;

    if (hasTimingSettings) {
      this._log("testing", "mem-timings", "Memory timing settings found in BIOS — analyzing current timings...");

      // Read current timing values
      const timings = {};
      const timingNames = [
        { key: "casLatency", name: "CAS Latency (tCL)", min: 14, safeStep: -1 },
        { key: "tRCD", name: "tRCD", min: 14, safeStep: -1 },
        { key: "tRP", name: "tRP", min: 14, safeStep: -1 },
        { key: "tRAS", name: "tRAS", min: 28, safeStep: -2 },
      ];

      for (const t of timingNames) {
        const setting = ms[t.key];
        if (setting && setting.value) {
          const val = parseInt(setting.value, 16);
          if (!isNaN(val) && val > 0 && val < 200) {
            timings[t.key] = { current: val, min: t.min, step: t.safeStep, settingObj: setting };
            this._log("testing", "mem-timing-read", `${t.name}: current = ${val} (raw: ${setting.value})`);
          }
        }
      }

      if (Object.keys(timings).length > 0) {
        this._log("testing", "mem-timing-note", `Found ${Object.keys(timings).length} timing settings. Tightening timings improves latency but requires careful stability testing.`);
        this._log("testing", "mem-timing-note2", "Memory timing changes require reboot to take effect. Each timing change will be logged but applied together with the next system reboot.");

        // For each timing, try tightening by one step
        for (const t of timingNames) {
          if (this._stopRequested) break;
          if (!timings[t.key]) continue;

          const currentVal = timings[t.key].current;
          const newVal = currentVal + t.safeStep; // safeStep is negative = tighter
          const minVal = timings[t.key].min;

          if (newVal < minVal) {
            this._log("testing", "mem-timing-floor", `${t.name}: ${currentVal} already at or near minimum (floor: ${minVal}) — skipping`);
            continue;
          }

          // Find the BIOS setting name for this timing
          const settingName = Object.keys(this.scewin.currentSettings || {}).find(name => {
            const lower = name.toLowerCase();
            if (t.key === "casLatency") return lower.includes("cas") && (lower.includes("latency") || lower.includes("tcl"));
            if (t.key === "tRCD") return lower.includes("trcd") || (lower.includes("ras") && lower.includes("cas"));
            if (t.key === "tRP") return lower.includes("trp") || lower.includes("precharge");
            if (t.key === "tRAS") return lower.includes("tras") || lower.includes("active");
            return false;
          });

          if (settingName) {
            this._log("testing", "mem-timing-set", `${t.name}: ${currentVal} → ${newVal} (tighter by ${Math.abs(t.safeStep)}) via "${settingName}"`);
            const writeResult = await this.scewin.writeSetting(settingName, newVal);
            if (writeResult.success) {
              result.timingsChanged.push({ name: t.name, from: currentVal, to: newVal, setting: settingName });
              result.timingsTightened = true;
            } else {
              this._log("testing", "mem-timing-fail", `Failed to write ${t.name}: ${writeResult.error}`);
            }
          } else {
            this._log("testing", "mem-timing-no-setting", `Could not find BIOS setting name for ${t.name}`);
          }
        }

        if (result.timingsTightened) {
          this._log("testing", "mem-timing-written", `Wrote ${result.timingsChanged.length} timing changes to BIOS NVRAM. Changes take effect after reboot. Will verify stability during final validation.`);
        }
      }
    } else {
      this._log("testing", "mem-no-timings", "No individual timing settings (tCL/tRCD/tRP/tRAS) found in BIOS export — timing optimization skipped. XMP profile controls timings instead.");
    }

    this._log("testing", "mem-complete", `Memory phase complete: XMP ${result.xmpEnabled ? (result.xmpAlreadyEnabled ? "already on" : "ENABLED") : "not available"} | Timings tightened: ${result.timingsTightened ? result.timingsChanged.map(c => `${c.name}: ${c.from}→${c.to}`).join(", ") : "no"}`);

    return result;
  }

  // ── GPU Phase ─────────────────────────────────────────────────────
  // GPU OC doesn't require reboot — nvidia-smi applies changes instantly.
  // We increase the power limit in 5% steps and verify each change actually
  // took effect by querying nvidia-smi after applying.
  // ────────────────────────────────────────────────────────────────────

  async _gpuPhase() {
    this._setPhase("testing", 65);
    this._log("testing", "gpu-start", "Starting GPU overclock phase");

    const gpu = await this.hw.detectGPU();
    if (!gpu.hasNvidiaSmi) {
      this._log("testing", "skip-gpu", "No NVIDIA GPU with nvidia-smi found — skipping GPU phase");
      return null;
    }

    const basePowerLimit = gpu.powerLimit;
    const baseClockCore = gpu.clockCore || 0;
    const baseClockMem = gpu.clockMem || 0;
    let lastGoodLimit = basePowerLimit;
    const maxTemp = 90;

    this._log("testing", "gpu-baseline", `GPU: ${gpu.name} | Power: ${basePowerLimit}W | Core: ${baseClockCore} MHz | Mem: ${baseClockMem} MHz | Temp: ${gpu.temp}C`, {
      power: basePowerLimit, clockCore: baseClockCore, clockMem: baseClockMem, temp: gpu.temp,
    });

    for (let pctIncrease = 5; pctIncrease <= 20; pctIncrease += 5) {
      if (this._stopRequested) break;

      const newLimit = Math.round(basePowerLimit * (1 + pctIncrease / 100));
      this._setPhase("testing", 65 + (pctIncrease / 20) * 15);

      this._log("testing", "gpu-step", `Setting GPU power limit to ${newLimit}W (+${pctIncrease}%)`, { power: newLimit });

      const setResult = await this.hw.gpuSetPowerLimit(newLimit);

      // ── VERIFICATION: Read back the power limit to confirm it was applied ──
      await new Promise(r => setTimeout(r, 2000)); // give nvidia-smi a moment
      const verifyGpu = await this.hw.detectGPU();
      const actualPL = verifyGpu.powerLimit || 0;
      const plDiff = Math.abs(actualPL - newLimit);

      if (plDiff > 5) {
        this._log("testing", "gpu-verify-fail", `VERIFICATION FAILED: Requested ${newLimit}W but nvidia-smi reports ${actualPL}W (diff: ${plDiff}W). nvidia-smi may not have admin privileges or the GPU may not support this limit.`, {
          requested: newLimit, actual: actualPL, diff: plDiff,
        });
        // If verification failed on the first step, the whole GPU OC won't work
        if (pctIncrease === 5) {
          this._log("testing", "gpu-abort", "GPU power limit change not taking effect — aborting GPU phase. Try running the app as Administrator.");
          return { finalPowerLimit: basePowerLimit, gainWatts: 0, gainPct: 0, verified: false, error: "Power limit changes not applied — admin required" };
        }
        break;
      }

      this._log("testing", "gpu-verify-ok", `VERIFIED: Power limit confirmed at ${actualPL}W (requested ${newLimit}W)`, {
        requested: newLimit, actual: actualPL,
      });

      // Stress test with GPU load
      this._log("testing", "gpu-stress", `Running 5-minute GPU stress test at ${newLimit}W...`);
      const result = await this.runExtendedStressTest(5, { maxTemp, stressGpu: true });

      // Read GPU stats after stress (most accurate under/just after load)
      const postStats = await this.hw.getStats();
      const gpuTemp = postStats.gpu?.temp || 0;
      const gpuClockNow = postStats.gpu?.clockCore || 0;
      const gpuPowerNow = postStats.gpu?.power || 0;

      this._log("testing", "gpu-post-stress", `Post-stress: ${gpuTemp}C | Core: ${gpuClockNow} MHz | Power draw: ${gpuPowerNow}W (source: ${postStats.hwinfoActive ? "HWiNFO" : "nvidia-smi"})`, {
        temp: gpuTemp, clock: gpuClockNow, power: gpuPowerNow,
      });

      if (result.stable && gpuTemp < maxTemp) {
        lastGoodLimit = newLimit;
        this._log("testing", "gpu-stable", `${newLimit}W is STABLE — GPU boosted to ${gpuClockNow} MHz at ${gpuTemp}C`, {
          stable: true, temp: gpuTemp, power: newLimit, clock: gpuClockNow,
        });
      } else {
        this._log("testing", "gpu-rollback", `${newLimit}W unstable or too hot (${gpuTemp}C) — rolling back to ${lastGoodLimit}W`, {
          stable: false, temp: gpuTemp,
        });
        await this.hw.gpuSetPowerLimit(lastGoodLimit);

        // Verify rollback too
        await new Promise(r => setTimeout(r, 1000));
        const rollbackGpu = await this.hw.detectGPU();
        this._log("testing", "gpu-rollback-verify", `Rolled back: power limit now ${rollbackGpu.powerLimit}W`);
        break;
      }
    }

    // ── FINAL GPU VERIFICATION ──
    const finalGpu = await this.hw.detectGPU();
    const finalPL = finalGpu.powerLimit || basePowerLimit;
    const gainW = finalPL - basePowerLimit;
    const gainPct = Math.round((gainW / basePowerLimit) * 100);

    this._log("testing", "gpu-complete", `GPU OC complete: ${finalPL}W (${gainPct >= 0 ? "+" : ""}${gainPct}%) | Verified power limit: ${finalPL}W | Core clock: ${finalGpu.clockCore || "?"}MHz`, {
      finalPowerLimit: finalPL, gainWatts: gainW, gainPct, verified: true,
    });

    return {
      finalPowerLimit: finalPL,
      gainWatts: gainW,
      gainPct,
      verified: true,
      baseClockCore,
      finalClockCore: finalGpu.clockCore || 0,
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

    // ── CHECK FOR RESUME STATE (after reboot) ──
    const resumeState = this._loadOcState();
    if (resumeState) {
      this._log("analyzing", "resume-detected", `Resuming AI OC after reboot (saved at ${resumeState.savedAt})`);
      // Restore the log from before the reboot
      if (resumeState.log && Array.isArray(resumeState.log)) {
        this.log = [...resumeState.log];
        this._log("analyzing", "log-restored", `Restored ${resumeState.log.length} log entries from before reboot`);
      }
      // Restore BIOS discovery map
      if (resumeState.biosMap) {
        this.biosMap = resumeState.biosMap;
        this._log("analyzing", "biosmap-restored", "BIOS setting names restored from saved state");
      }
      // Clear the saved state now that we've loaded it
      this._clearOcState();
    }

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

    this._log("analyzing", "start", `AI Auto-Overclock ${resumeState ? "RESUMING" : "starting"}... SceWin: ${this.scewin.available ? "FOUND at " + this.scewin.scewinPath : "not found (GPU-only mode)"} | HWiNFO: ${this.hw.hwinfo.available ? "ACTIVE" : "not detected"}`);

    try {
      // Phase 1: Analyze hardware (skip full analysis if resuming — use saved data)
      let analysis = null;
      if (resumeState && resumeState.analysis) {
        analysis = resumeState.analysis;
        this._log("analyzing", "analysis-restored", `Using saved hardware analysis: ${analysis.cpu?.name || "?"} | Unlocked: ${analysis.cpuUnlocked} | Stock turbo: x${analysis.stockTurboRatio} | Ceiling: x${analysis.safeCeilingRatio}`);
        // Still run BIOS discovery if we don't have it cached
        if (!this.biosMap) {
          await this._discoverBiosSettings();
        }
      } else {
        analysis = await this._analyzeHardware();
        if (!analysis) {
          this._setPhase("failed", 100);
          this.running = false;
          return { success: false, error: "Hardware analysis failed", log: this.log };
        }
      }

      if (this._stopRequested) {
        this._setPhase("idle", 0);
        this.running = false;
        return { success: false, error: "Stopped by user", log: this.log };
      }

      // Phase 2: CPU Overclock (pass resume state if resuming CPU phase)
      let cpuResult = null;
      if (options.skipCpu !== true) {
        const cpuResumeState = resumeState && resumeState.resumePhase === "cpu" ? resumeState : null;
        cpuResult = await this._cpuOcPhase(analysis, cpuResumeState);

        // If CPU phase returned pending (waiting for reboot), stop here
        if (cpuResult && cpuResult.pending) {
          this.running = false;
          return { success: true, pending: true, message: cpuResult.message, log: this.log };
        }
      }

      if (this._stopRequested) {
        this._setPhase("idle", 0);
        this.running = false;
        return { success: false, error: "Stopped by user", log: this.log };
      }

      // Phase 3: Memory (pass resume state if resuming memory phase)
      let memResult = null;
      if (options.skipMemory !== true) {
        const memResumeState = resumeState && resumeState.resumePhase === "memory" ? resumeState : null;
        memResult = await this._memoryPhase(memResumeState);
      }

      if (this._stopRequested) {
        this._setPhase("idle", 0);
        this.running = false;
        return { success: false, error: "Stopped by user", log: this.log };
      }

      // Phase 4: GPU (no reboot needed — applies instantly)
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

      // Clean up any leftover auto-start tasks
      this._clearOcState();

      const finalPhase = profile.stable ? "stable" : "failed";
      this._setPhase(finalPhase, 100);
      this._log(finalPhase, "complete", this._generateSummary(profile));

      this.running = false;
      return { success: true, profile, log: this.log };

    } catch (e) {
      this._setPhase("failed", 100);
      this._log("failed", "error", "Unexpected error: " + e.message);
      this._clearOcState(); // clean up on error
      this.running = false;
      return { success: false, error: e.message, log: this.log };
    }
  }

  // ── BIOS Setting Discovery ─────────────────────────────────────────

  async _discoverBiosSettings() {
    this.biosMap = {
      cpuRatio: null,
      cpuVoltage: null,
      cpuVoltageMode: null,
      xmpProfile: null,
      pl1: null,
      pl2: null,
      ringRatio: null,
      iccMax: null,
      avxOffset: null,
      tccOffset: null,
      cStates: null,
      speedStep: null,
      turboBoost: null,
      thermalVelocityBoost: null,
    };

    if (!this.scewin.available) {
      this._log("analyzing", "bios-discovery", "SceWin not available — skipping BIOS setting discovery");
      return this.biosMap;
    }

    this._log("analyzing", "bios-discovery", "Exporting all BIOS settings via SceWin for discovery...");
    const exported = await this.scewin.exportCurrentSettings();
    if (!exported.success) {
      this._log("analyzing", "bios-discovery-fail", "Failed to export BIOS settings: " + exported.error);
      if (exported.log) {
        this._log("analyzing", "scewin-log", "Batch log: " + exported.log.replace(/\r?\n/g, " | ").substring(0, 500));
      }
      return this.biosMap;
    }
    if (exported.method) {
      this._log("analyzing", "bios-discovery", `Export succeeded via: ${exported.method}`);
    }

    // Log raw file preview if available (diagnostic — helps identify wrong file or bad parse)
    if (exported.rawPreview) {
      this._log("analyzing", "bios-raw-preview", `Raw export first 300 chars: ${exported.rawPreview.replace(/\r?\n/g, "\\n")}`);
    }
    // Also try to read the export file directly for diagnostics
    try {
      const rawFile = fs.readFileSync(SCEWIN_EXPORT_PATH, "utf8");
      this._log("analyzing", "bios-file-size", `Export file: ${rawFile.length} bytes`);
      this._log("analyzing", "bios-file-head", `First 500 chars: ${rawFile.substring(0, 500).replace(/\r?\n/g, "\\n")}`);
      // Count lines that look like actual settings
      const lines = rawFile.split("\n");
      this._log("analyzing", "bios-file-lines", `Total lines: ${lines.length}`);
      const setupLines = lines.filter(l => /setup question|token.*value/i.test(l));
      this._log("analyzing", "bios-file-setup", `Lines matching 'Setup Question' or 'Token/Value': ${setupLines.length}`);
      if (setupLines.length > 0) {
        this._log("analyzing", "bios-file-sample", `Sample setup lines: ${setupLines.slice(0, 5).join(" | ")}`);
      }
    } catch(e) {
      this._log("analyzing", "bios-file-read-error", e.message);
    }

    const allNames = Object.keys(exported.settings);
    this._log("analyzing", "bios-discovery", `Exported ${allNames.length} BIOS settings — running fuzzy match`);

    // Log ALL setting names so we can see exactly what the BIOS exports
    // Group them in chunks of ~10 for readability
    for (let i = 0; i < allNames.length; i += 10) {
      const chunk = allNames.slice(i, i + 10).join(" | ");
      this._log("analyzing", "bios-settings-dump", `Settings [${i}-${Math.min(i+9, allNames.length-1)}]: ${chunk}`);
    }

    // ═══════════════════════════════════════════════════════════════════
    // COMPREHENSIVE FUZZY MATCHING
    // Every motherboard manufacturer uses different names for the same
    // BIOS settings. This covers: ASUS, MSI, Gigabyte, ASRock, EVGA,
    // Biostar, SuperMicro, Intel reference BIOS, and AMI Aptio defaults.
    //
    // The fuzzyFind helper returns the first setting name that contains
    // any of the keywords (case-insensitive). Keywords are ordered from
    // most specific to least specific to avoid false positives.
    // ═══════════════════════════════════════════════════════════════════

    // fuzzyFind with EXCLUDE support: returns the first setting name that
    // contains any keyword but does NOT contain any exclude term.
    // This prevents false positives like "IGP Ratio" matching for CPU ratio
    // or "Loadline Calibration" matching for CPU voltage.
    const fuzzyFind = (keywords, excludes = []) => {
      for (const kw of keywords) {
        const kwLower = kw.toLowerCase();
        for (const name of allNames) {
          const lower = name.toLowerCase();
          if (!lower.includes(kwLower)) continue;
          // Check exclusions
          let excluded = false;
          for (const ex of excludes) {
            if (lower.includes(ex.toLowerCase())) { excluded = true; break; }
          }
          if (!excluded) return name;
        }
      }
      return null;
    };

    // CPU ratio / multiplier
    // ASUS: "CPU Core Ratio", "All-Core Ratio Limit", "Performance Core Ratio"
    // MSI: "Adjust CPU Ratio", "CPU Ratio", "Ratio Limit", "Per Core Ratio"
    // Gigabyte: "CPU Clock Ratio", "CPU Ratio", "Host Clock Ratio"
    // ASRock: "CPU Ratio", "All Core", "Multi"
    // Intel ref: "Processor Core Ratio", "Active Processor Cores"
    // EVGA: "CPU Core Ratio", "All Core Multiplier"
    //
    // EXCLUDE: "IGP" (Integrated Graphics), "GT" (Intel iGPU), "Ring" (cache),
    //          "Uncore", "Memory", "DRAM", "SA" (System Agent), "FCLK"
    this.biosMap.cpuRatio = fuzzyFind([
      "all-core ratio limit", "all core ratio limit", "performance core ratio",
      "adjust cpu ratio", "per core ratio limit",
      "cpu clock ratio", "host clock ratio",
      "all core ratio", "cpu core ratio", "core ratio limit",
      "p-core ratio", "pcore ratio",
      "processor core ratio", "active core ratio",
      "cpu ratio", "core ratio", "frequency ratio",
      "ratio limit", "multiplier", "cpu multi",
      "oc ratio", "overclock ratio",
      "flex ratio", "non-turbo ratio",
      "max ratio",
    ], ["igp", "gt ratio", "ring", "uncore", "memory", "dram", "sa ", "fclk", "cache", "e-core", "ecore"]);

    // CPU voltage / Vcore
    // ASUS: "CPU Core/Cache Voltage", "CPU Core Voltage Override", "CPU SVID Support"
    // MSI: "CPU Core Voltage", "CPU Vcore", "Vcore Override Voltage"
    // Gigabyte: "CPU Vcore", "Dynamic Vcore(DVID)", "CPU Vcore Override"
    // ASRock: "CPU Core/Cache Voltage", "Vcore Override Voltage", "CPU Voltage"
    // Intel ref: "Core Voltage Override", "VR Configuration"
    // EVGA: "CPU Vcore Override", "Vcore", "VCore Voltage"
    //
    // EXCLUDE: "Loadline", "Calibration", "LLC", "Load Line", "Mode", "SVID",
    //          "Frequency", "Ratio", "Ring", "SA", "IO", "DRAM", "GT", "PLL"
    this.biosMap.cpuVoltage = fuzzyFind([
      "cpu core/cache voltage", "cpu core voltage override",
      "vcore override voltage", "cpu vcore override",
      "dynamic vcore", "dvid",
      "core voltage override", "vr voltage override",
      "cpu core voltage", "cpu cache voltage",
      "vcore override", "vcore voltage",
      "ia voltage offset", "ia voltage override",
      "cpu vcore", "core voltage",
      "override voltage", "adaptive voltage",
      "voltage offset", "voltage override",
      "vcore", "cpu voltage",
    ], ["loadline", "calibration", "llc", "load line", "mode", "svid", "frequency",
        "ratio", "ring", "sa ", "dram", "gt ", "pll", "uncore", "memory", "io "]);

    // CPU voltage mode (adaptive vs manual vs offset)
    // ASUS: "CPU Core/Cache Voltage Mode", "SVID Behavior"
    // MSI: "CPU Core Voltage Mode", "Voltage Mode"
    // Gigabyte: "CPU Vcore Mode"
    // ASRock: "Voltage Mode", "CPU Voltage Mode"
    this.biosMap.cpuVoltageMode = fuzzyFind([
      "cpu core/cache voltage mode", "svid behavior",
      "cpu core voltage mode", "cpu vcore mode",
      "voltage mode", "vcore mode",
      "svid support", "svid control",
    ], ["loadline", "calibration", "llc", "ratio", "dram"]);

    // XMP / EXPO / DOCP memory profile
    // ASUS: "Ai Overclock Tuner", "XMP", "DOCP"
    // MSI: "A-XMP", "Extreme Memory Profile", "XMP Profile"
    // Gigabyte: "Extreme Memory Profile(X.M.P.)", "XMP", "System Memory Multiplier"
    // ASRock: "DRAM Configuration > XMP", "Load XMP Setting"
    // AMD: "EXPO", "DOCP", "D.O.C.P."
    this.biosMap.xmpProfile = fuzzyFind([
      "ai overclock tuner", "a-xmp",
      "extreme memory profile", "load xmp setting",
      "memory profile", "xmp profile", "xmp setting",
      "system memory multiplier",
      "d.o.c.p", "docp",
      "expo profile", "expo",
      "xmp", "dram profile",
    ]);

    // Power Limit 1 (long duration / PBP / TDP)
    // ASUS: "Long Duration Package Power Limit", "PL1", "Package Power Limit"
    // MSI: "Long Duration Power Limit (W)", "PL1 (W)", "CPU Power Limit Value"
    // Gigabyte: "Package Power Limit1 - TDP (Watts)", "PL1"
    // ASRock: "Long Duration Power Limit", "PBP", "Base Power Limit"
    // Intel ref: "Package Power Limit 1", "Power Limit 1 Value"
    this.biosMap.pl1 = fuzzyFind([
      "long duration package power limit", "long duration power limit",
      "package power limit1", "package power limit 1",
      "power limit 1 value", "cpu power limit value",
      "processor base power", "base power limit",
      "power limit 1", "long duration",
      "pbp power", "pl1 power",
      "pl1 (w)", "pl1(w)", "pl1",
      "tdp limit", "tdp power",
    ]);

    // Power Limit 2 (short duration / max turbo power)
    // ASUS: "Short Duration Package Power Limit", "PL2"
    // MSI: "Short Duration Power Limit (W)", "PL2 (W)"
    // Gigabyte: "Package Power Limit2 (Watts)", "PL2"
    // ASRock: "Short Duration Power Limit", "Maximum Turbo Power"
    // Intel ref: "Package Power Limit 2", "Power Limit 2 Value"
    this.biosMap.pl2 = fuzzyFind([
      "short duration package power limit", "short duration power limit",
      "package power limit2", "package power limit 2",
      "power limit 2 value", "maximum turbo power",
      "max turbo power", "power limit 2", "short duration",
      "pl2 (w)", "pl2(w)", "pl2",
      "turbo power limit",
    ]);

    // Ring / Cache / Uncore ratio
    // ASUS: "Min/Max CPU Cache Ratio", "Ring Ratio", "Uncore Ratio"
    // MSI: "Ring Ratio", "CPU Cache Ratio", "Adjust Ring Ratio"
    // Gigabyte: "Uncore Ratio", "Ring Multiplier", "CPU Uncore Ratio"
    // ASRock: "Cache Ratio", "Ring Ratio", "Min Ring Ratio"
    // Intel ref: "Uncore Frequency"
    this.biosMap.ringRatio = fuzzyFind([
      "min cpu cache ratio", "max cpu cache ratio", "cpu cache ratio",
      "adjust ring ratio", "ring multiplier",
      "cpu uncore ratio", "uncore frequency",
      "ring ratio", "cache ratio", "uncore ratio",
      "ring down bin", "min ring ratio", "max ring ratio",
      "ring", "cache", "uncore",
    ]);

    // ICC Max (current limit)
    // ASUS: "IccMax", "IA AC Load Line", "IA DC Load Line"
    // MSI: "IA AC/DC Loadline", "ICC Max Override"
    // Gigabyte: "IA AC Loadline", "CPU VR Current Limit"
    // ASRock: "IccMax", "Long Duration Maintained"
    // Intel ref: "ICC Max Unlimited", "Current Limit"
    this.biosMap.iccMax = fuzzyFind([
      "iccmax unlimited", "icc max override",
      "ia ac/dc loadline", "ia ac loadline", "ia dc loadline",
      "cpu vr current limit", "current limit",
      "ia ac load line", "ia dc load line",
      "icc max", "iccmax", "icc",
    ]);

    // AVX Offset (reduces ratio during AVX workloads)
    // ASUS: "AVX2 Ratio Offset", "AVX Instruction Core Ratio Negative Offset"
    // MSI: "AVX Ratio Offset", "AVX2 Offset"
    // Gigabyte: "AVX Offset", "AVX Frequency Trim"
    // ASRock: "AVX2 Ratio Offset", "AVX 512 Offset"
    this.biosMap.avxOffset = fuzzyFind([
      "avx instruction core ratio", "avx2 ratio offset",
      "avx ratio offset", "avx frequency trim",
      "avx 512 offset", "avx-512 offset",
      "avx512 ratio offset",
      "avx offset", "avx2 offset",
      "avx negative offset",
    ]);

    // TCC Activation Offset (thermal throttle temperature offset)
    // ASUS: "TCC Activation Offset"
    // MSI: "TCC Activation Offset", "CPU TCC Offset"
    // Gigabyte: "TCC Offset"
    // ASRock: "TCC Activation Offset"
    this.biosMap.tccOffset = fuzzyFind([
      "tcc activation offset", "cpu tcc offset",
      "tcc offset", "tcc",
    ]);

    // C-States (CPU power saving states)
    // ASUS: "CPU - C6 Report", "C-State", "Intel C-State"
    // MSI: "Intel C-State", "C1E Support", "Package C State"
    // Gigabyte: "C-States Support", "Enhanced Halt State (C1E)"
    // ASRock: "CPU C States Support", "C State Support"
    this.biosMap.cStates = fuzzyFind([
      "cpu c states support", "cpu c-state", "intel c-state",
      "c-states support", "c state support",
      "package c state", "c6 report", "c1e support",
      "c-state", "c state",
    ]);

    // SpeedStep / EIST
    // ASUS: "Intel SpeedStep Technology", "EIST"
    // MSI: "Intel SpeedStep", "EIST"
    // Gigabyte: "Enhanced Intel SpeedStep", "EIST"
    // ASRock: "Intel SpeedStep Technology", "Speed Step"
    this.biosMap.speedStep = fuzzyFind([
      "intel speedstep technology", "enhanced intel speedstep",
      "intel speedstep", "intel speed step",
      "speedstep", "speed step", "eist",
    ]);

    // Turbo Boost
    // ASUS: "Intel Turbo Boost Technology", "Turbo Mode"
    // MSI: "Intel Turbo Boost", "Turbo Boost"
    // Gigabyte: "Intel(R) Turbo Boost Technology", "Turbo Boost"
    // ASRock: "Intel Turbo Boost Technology", "Turbo Mode"
    this.biosMap.turboBoost = fuzzyFind([
      "intel turbo boost technology", "intel(r) turbo boost",
      "intel turbo boost", "turbo boost technology",
      "turbo boost", "turbo mode",
      "multi core enhancement", "mce",
    ]);

    // Thermal Velocity Boost (Intel 14th gen feature)
    // ASUS: "Intel Thermal Velocity Boost", "TVB"
    // MSI: "Thermal Velocity Boost"
    // Gigabyte: "TVB", "Thermal Velocity Boost"
    this.biosMap.thermalVelocityBoost = fuzzyFind([
      "intel thermal velocity boost", "thermal velocity boost",
      "tvb voltage optimizations", "tvb",
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
      if (profile.cpu.pending) {
        lines.push("CPU: In progress (waiting for reboot)");
      } else {
        const stockNote = profile.cpu.stockTurboMHz ? ` (stock turbo: ${profile.cpu.stockTurboMHz} MHz)` : "";
        lines.push(`CPU: x${profile.cpu.finalRatio} (${profile.cpu.finalMHz} MHz) — ${profile.cpu.gainMHz >= 0 ? "+" : ""}${profile.cpu.gainMHz} MHz${stockNote}`);
        if (profile.cpu.finalVoltage) lines.push(`  Voltage: ${profile.cpu.finalVoltage}V (optimized — lowest stable)`);
      }
    } else {
      lines.push("CPU: No changes (skipped or unavailable)");
    }

    if (profile.memory) {
      const xmpStatus = profile.memory.xmpEnabled
        ? (profile.memory.xmpAlreadyEnabled ? "already enabled" : "ENABLED")
        : "not available";
      lines.push(`Memory: XMP/EXPO ${xmpStatus}`);
      if (profile.memory.ratedSpeed) lines.push(`  Speed: ${profile.memory.currentSpeed || "?"} MHz (rated ${profile.memory.ratedSpeed} MHz)`);
      if (profile.memory.timingsTightened && profile.memory.timingsChanged.length > 0) {
        const changes = profile.memory.timingsChanged.map(c => `${c.name}: ${c.from}→${c.to}`).join(", ");
        lines.push(`  Timings tightened: ${changes}`);
      }
    } else {
      lines.push("Memory: No changes (skipped or unavailable)");
    }

    if (profile.gpu) {
      if (profile.gpu.verified === false) {
        lines.push(`GPU: Changes could not be verified — ${profile.gpu.error || "unknown"}`);
      } else {
        lines.push(`GPU: ${profile.gpu.finalPowerLimit}W (${profile.gpu.gainPct >= 0 ? "+" : ""}${profile.gpu.gainPct}%) — VERIFIED`);
        if (profile.gpu.baseClockCore && profile.gpu.finalClockCore) {
          lines.push(`  Core clock: ${profile.gpu.baseClockCore} → ${profile.gpu.finalClockCore} MHz`);
        }
      }
    } else {
      lines.push("GPU: No changes (skipped or unavailable)");
    }

    if (profile.validation) {
      lines.push(`Validation: ${profile.validation.stable ? "PASSED" : "FAILED"} (max temp: ${profile.validation.maxTemp}C, avg: ${profile.validation.avgTemp}C)`);
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

  aiHasResumeState() {
    return this.aiEngine.hasResumeState();
  }
}

module.exports = HardwareMonitor;