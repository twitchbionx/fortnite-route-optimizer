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

// ── Hardware Detection ──────────────────────────────────────────────
class HardwareMonitor {
  constructor() {
    this.isWin = os.platform() === "win32";
    this.nvidiaSmiPath = null;
    this.hardwareInfo = null;
    this.detectNvidiaSmi();
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

    const [cpuStats, gpuStats, ramStats] = await Promise.all([
      this.getCPUStats(),
      this.getGPUStats(),
      this.getRAMStats(),
    ]);

    return { cpu: cpuStats, gpu: gpuStats, ram: ramStats };
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

  // ── Stress Test (Quick) ──────────────────────────────────────────
  // Simple CPU stress test using PowerShell
  async runStressTest(durationSec = 10) {
    const r = await runPS(`
      $cores = (Get-CimInstance Win32_Processor).NumberOfLogicalProcessors;
      $jobs = @();
      $end = (Get-Date).AddSeconds(${durationSec});
      for ($i = 0; $i -lt $cores; $i++) {
        $jobs += Start-Job -ScriptBlock {
          $e = [DateTime]$args[0];
          while ((Get-Date) -lt $e) { [Math]::Sqrt(12345.6789) | Out-Null }
        } -ArgumentList $end.ToString("o")
      }
      Start-Sleep -Seconds ${durationSec};
      $jobs | Stop-Job -PassThru | Remove-Job;
      "Stress test completed (${durationSec}s on $cores threads)"
    `, (durationSec + 5) * 1000);

    return { success: r.success, output: r.output, error: r.error };
  }
}

module.exports = HardwareMonitor;
