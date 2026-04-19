// ═══════════════════════════════════════════════════════════════════
//  BIOS SETTINGS HELPER MODULE
//  Reads current BIOS/firmware settings via WMI, exposes hidden
//  settings that most BIOS menus don't show, and provides guided
//  optimization recommendations based on your hardware.
//
//  NOTE: Most BIOS changes REQUIRE a reboot into BIOS/UEFI.
//  This module reads what's set and tells you what to change.
//  It can also reboot straight into BIOS for you.
// ═══════════════════════════════════════════════════════════════════

const { exec } = require("child_process");
const os = require("os");

function runCmd(cmd, timeout = 15000) {
  return new Promise((resolve) => {
    exec(cmd, { timeout }, (err, stdout, stderr) => {
      resolve({ success: !err, output: stdout?.trim() || "", error: stderr?.trim() || err?.message || "" });
    });
  });
}

function runPS(cmd, timeout = 20000) {
  return new Promise((resolve) => {
    exec(cmd, { timeout, shell: "powershell.exe" }, (err, stdout, stderr) => {
      resolve({ success: !err, output: stdout?.trim() || "", error: stderr?.trim() || err?.message || "" });
    });
  });
}

// ── BIOS Settings Guide ─────────────────────────────────────────────
// Settings you should change in BIOS for competitive gaming
const BIOS_GUIDE = [
  {
    id: "xmp", category: "Memory",
    name: "Enable XMP / EXPO Profile",
    desc: "Your RAM ships rated at a speed (e.g., 3600MHz) but runs at 2133MHz by default. XMP tells the motherboard to use the rated speed. This is the single biggest free performance boost you can get.",
    where: "BIOS > AI Tweaker / OC / DRAM > XMP / D.O.C.P / EXPO",
    impact: "HIGH",
    safe: true,
  },
  {
    id: "resizable-bar", category: "GPU",
    name: "Enable Resizable BAR / Smart Access Memory",
    desc: "Lets the CPU access the full GPU VRAM at once instead of through a 256MB window. 3-10% FPS improvement in many games including Fortnite.",
    where: "BIOS > Advanced > PCI > Above 4G Decoding = Enabled, Resizable BAR = Enabled",
    impact: "HIGH",
    safe: true,
  },
  {
    id: "pbo", category: "CPU (AMD)",
    name: "Enable PBO (Precision Boost Overdrive)",
    desc: "AMD's auto-overclock feature. Lets the CPU boost higher and longer than stock limits. Free performance with zero risk — AMD's own feature.",
    where: "BIOS > AMD Overclocking > PBO > Enable / Advanced",
    impact: "HIGH",
    safe: true,
    amdOnly: true,
  },
  {
    id: "curve-optimizer", category: "CPU (AMD)",
    name: "Curve Optimizer (AMD 5000/7000 Series)",
    desc: "Undervolts each core individually for lower temps AND higher boost clocks. Start with All Cores Negative 15-20. Test stability.",
    where: "BIOS > AMD Overclocking > PBO > Curve Optimizer > All Cores > Negative > 15-20",
    impact: "HIGH",
    safe: false,
    amdOnly: true,
  },
  {
    id: "intel-turbo", category: "CPU (Intel)",
    name: "Enable All-Core Turbo / Turbo Boost Max",
    desc: "Ensure all turbo boost features are enabled. Some boards limit boost by default.",
    where: "BIOS > CPU Config > Intel Turbo Boost > Enabled / Max Turbo Ratio = All Cores",
    impact: "MEDIUM",
    safe: true,
    intelOnly: true,
  },
  {
    id: "c-states", category: "CPU",
    name: "Disable C-States (Power Saving)",
    desc: "C-States put CPU cores to sleep to save power. Waking them adds latency. Disable for gaming.",
    where: "BIOS > CPU Config > C-States > Disabled / C1E = Disabled",
    impact: "MEDIUM",
    safe: true,
  },
  {
    id: "speedstep", category: "CPU",
    name: "Disable SpeedStep / Cool'n'Quiet",
    desc: "Prevents CPU from downclocking during light loads. Keeps CPU at max speed always.",
    where: "BIOS > CPU Config > SpeedStep (Intel) / Cool'n'Quiet (AMD) > Disabled",
    impact: "LOW",
    safe: true,
  },
  {
    id: "hpet-bios", category: "Timing",
    name: "Disable HPET in BIOS",
    desc: "High Precision Event Timer can cause micro-stuttering. Disable in BIOS AND Windows (we already disabled Windows side).",
    where: "BIOS > Advanced > PCH Config / ACPI > HPET > Disabled",
    impact: "MEDIUM",
    safe: true,
  },
  {
    id: "fast-boot", category: "Boot",
    name: "Enable Fast Boot / Ultra Fast Boot",
    desc: "Skips POST hardware checks for faster boot. Only downside: harder to enter BIOS on restart.",
    where: "BIOS > Boot > Fast Boot > Enabled / Ultra Fast",
    impact: "LOW",
    safe: true,
  },
  {
    id: "pcie-gen", category: "GPU",
    name: "Force PCIe Gen 4/5",
    desc: "Some boards auto-detect PCIe gen wrong. Force it to match your GPU for max bandwidth.",
    where: "BIOS > Advanced > PCIe Config > PCIe Speed > Gen 4 or Gen 5",
    impact: "LOW",
    safe: true,
  },
  {
    id: "virtualization", category: "CPU",
    name: "Disable Virtualization (VT-x / SVM)",
    desc: "If you don't use VMs or WSL2, disable CPU virtualization. Frees up resources and reduces overhead.",
    where: "BIOS > CPU Config > Intel VT-x / AMD SVM > Disabled",
    impact: "LOW",
    safe: true,
  },
  {
    id: "spread-spectrum", category: "Advanced",
    name: "Disable Spread Spectrum",
    desc: "Spread spectrum reduces EMI by slightly varying clock speeds. Disabling gives more stable clocks.",
    where: "BIOS > Advanced > Spread Spectrum > Disabled (for CPU, BCLK, PCIe)",
    impact: "LOW",
    safe: true,
  },
  {
    id: "erp", category: "Power",
    name: "Disable ErP Ready",
    desc: "ErP cuts power to USB ports on shutdown. Disable so your mouse/keyboard wake the PC instantly.",
    where: "BIOS > Advanced > APM > ErP Ready > Disabled",
    impact: "LOW",
    safe: true,
  },
];

// ── Module Class ─────────────────────────────────────────────────────
class BIOSHelper {
  constructor() {
    this.isWin = os.platform() === "win32";
  }

  // Read current BIOS info
  async getBIOSInfo() {
    if (!this.isWin) return { error: "Windows only" };

    const r = await runPS(`
      $bios = Get-CimInstance Win32_BIOS | Select-Object SMBIOSBIOSVersion,Manufacturer,ReleaseDate,SerialNumber;
      $board = Get-CimInstance Win32_BaseBoard | Select-Object Manufacturer,Product,Version;
      $boot = bcdedit /enum "{current}" 2>&1 | Out-String;
      @{bios=$bios; board=$board; boot=$boot} | ConvertTo-Json -Depth 3
    `);

    try {
      const info = JSON.parse(r.output);
      return {
        biosVersion: info.bios?.SMBIOSBIOSVersion,
        biosVendor: info.bios?.Manufacturer?.trim(),
        boardMfg: info.board?.Manufacturer?.trim(),
        boardModel: info.board?.Product?.trim(),
        bootConfig: info.boot,
      };
    } catch(e) {
      return { error: e.message };
    }
  }

  // Read hidden/advanced settings via WMI (varies by manufacturer)
  async readHiddenSettings() {
    if (!this.isWin) return { settings: [], error: "Windows only" };

    const settings = [];

    // Try to read boot configuration
    const boot = await runCmd("bcdedit /enum");
    if (boot.success) {
      const lines = boot.output.split("\n");
      for (const line of lines) {
        if (line.includes("useplatformclock") || line.includes("useplatformtick") ||
            line.includes("disabledynamictick") || line.includes("hypervisorlaunchtype") ||
            line.includes("nx") || line.includes("debug")) {
          const parts = line.trim().split(/\s+/);
          settings.push({
            source: "BCD",
            name: parts[0],
            value: parts.slice(1).join(" "),
            editable: true,
          });
        }
      }
    }

    // Virtualization status
    const virt = await runPS("(Get-CimInstance Win32_Processor).VirtualizationFirmwareEnabled");
    settings.push({
      source: "CPU",
      name: "Virtualization (VT-x/SVM)",
      value: virt.output === "True" ? "Enabled" : "Disabled",
      editable: false,
      recommendation: virt.output === "True" ? "Disable in BIOS if you don't use VMs" : "Already optimal",
    });

    // Secure Boot
    const sb = await runPS("Confirm-SecureBootUEFI 2>$null");
    settings.push({
      source: "UEFI",
      name: "Secure Boot",
      value: sb.output === "True" ? "Enabled" : "Disabled",
      editable: false,
    });

    // Hyper-V
    const hv = await runPS("(Get-WindowsOptionalFeature -FeatureName Microsoft-Hyper-V -Online -ErrorAction SilentlyContinue).State");
    settings.push({
      source: "Windows",
      name: "Hyper-V",
      value: hv.output || "Not Available",
      editable: true,
      recommendation: hv.output === "Enabled" ? "Disable for less VM overhead. Run: bcdedit /set hypervisorlaunchtype off" : "Already optimal",
    });

    // Power Plan
    const pp = await runCmd("powercfg -getactivescheme");
    if (pp.success) {
      settings.push({
        source: "Power",
        name: "Active Power Plan",
        value: pp.output.replace(/^.*:\s*/, ""),
        editable: true,
      });
    }

    // HPET status
    const hpet = await runCmd("bcdedit /enum | findstr useplatformclock");
    settings.push({
      source: "Timer",
      name: "HPET (Platform Clock)",
      value: hpet.success && hpet.output.includes("Yes") ? "Enabled (BAD)" : "Disabled (GOOD)",
      editable: true,
    });

    // Timer resolution
    const timer = await runCmd("bcdedit /enum | findstr useplatformtick");
    settings.push({
      source: "Timer",
      name: "Platform Tick",
      value: timer.success && timer.output.includes("Yes") ? "Enabled (GOOD)" : "Disabled",
      editable: true,
    });

    return { settings };
  }

  // Get the BIOS optimization guide filtered by hardware
  async getGuide() {
    // Detect CPU type
    const cpuInfo = await runPS("(Get-CimInstance Win32_Processor).Name");
    const isAMD = cpuInfo.output?.includes("AMD");
    const isIntel = cpuInfo.output?.includes("Intel");

    return BIOS_GUIDE.filter(item => {
      if (item.amdOnly && !isAMD) return false;
      if (item.intelOnly && !isIntel) return false;
      return true;
    }).map(item => ({
      ...item,
      cpuType: isAMD ? "AMD" : isIntel ? "Intel" : "Unknown",
    }));
  }

  // Reboot straight into BIOS/UEFI firmware settings
  async rebootToBIOS() {
    if (!this.isWin) return { success: false, error: "Windows only" };
    // This command reboots and goes directly to BIOS/UEFI setup
    const r = await runCmd("shutdown /r /fw /t 3");
    return { success: r.success, error: r.error, msg: "Rebooting to BIOS in 3 seconds..." };
  }

  // Apply a BCD (boot config) setting
  async applyBCDSetting(setting, value) {
    const allowed = ["useplatformclock", "useplatformtick", "disabledynamictick", "hypervisorlaunchtype"];
    if (!allowed.includes(setting)) return { success: false, error: "Setting not allowed" };

    const r = await runCmd(`bcdedit /set ${setting} ${value}`);
    return { success: r.success, error: r.error, reboot: true };
  }

  // Disable Hyper-V for less overhead
  async disableHyperV() {
    const r = await runCmd("bcdedit /set hypervisorlaunchtype off");
    return { success: r.success, error: r.error, reboot: true };
  }
}

module.exports = BIOSHelper;
