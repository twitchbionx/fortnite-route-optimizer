// ═══════════════════════════════════════════════════════════════════
//  GAMING REGISTRY TWEAKS MODULE
//  Deep Windows registry and system optimizations specifically
//  tuned for competitive FPS gaming. These go way beyond the
//  basic network tweaks — we're touching GPU scheduling, input
//  latency, timer resolution, MMCSS priorities, and more.
// ═══════════════════════════════════════════════════════════════════

const { exec } = require("child_process");
const os = require("os");

function runCmd(cmd, timeout = 15000) {
  return new Promise((resolve) => {
    exec(cmd, { timeout }, (err, stdout, stderr) => {
      resolve({
        success: !err,
        output: stdout?.trim() || "",
        error: stderr?.trim() || err?.message || "",
      });
    });
  });
}

function runPS(cmd, timeout = 15000) {
  return new Promise((resolve) => {
    exec(cmd, { timeout, shell: "powershell.exe" }, (err, stdout, stderr) => {
      resolve({
        success: !err,
        output: stdout?.trim() || "",
        error: stderr?.trim() || err?.message || "",
      });
    });
  });
}

// ── GPU & Display Tweaks ────────────────────────────────────────────
const GPU_TWEAKS = [
  {
    id: "hw-gpu-sched", name: "Hardware GPU Scheduling",
    desc: "Lets the GPU manage its own VRAM scheduling instead of Windows. Reduces input lag by 1-3ms in most cases.",
    check: 'reg query "HKLM\\SYSTEM\\CurrentControlSet\\Control\\GraphicsDrivers" /v HwSchMode 2>nul',
    checkValue: "0x2",
    cmd: 'reg add "HKLM\\SYSTEM\\CurrentControlSet\\Control\\GraphicsDrivers" /v HwSchMode /t REG_DWORD /d 2 /f',
    reboot: true,
  },
  {
    id: "fsoptimize", name: "Disable Fullscreen Optimizations (Global)",
    desc: "Windows fullscreen 'optimizations' add a composition layer that increases input lag. Kill it globally.",
    check: 'reg query "HKCU\\System\\GameConfigStore" /v GameDVR_FSEBehaviorMode 2>nul',
    checkValue: "0x2",
    cmd: 'reg add "HKCU\\System\\GameConfigStore" /v GameDVR_FSEBehaviorMode /t REG_DWORD /d 2 /f && reg add "HKCU\\System\\GameConfigStore" /v GameDVR_HonorUserFSEBehaviorMode /t REG_DWORD /d 1 /f && reg add "HKCU\\System\\GameConfigStore" /v GameDVR_FSEBehavior /t REG_DWORD /d 2 /f',
  },
  {
    id: "preemption", name: "GPU Preemption Granularity",
    desc: "Sets GPU preemption to DMA packet level. Reduces frame scheduling overhead for smoother frametimes.",
    check: 'reg query "HKLM\\SYSTEM\\CurrentControlSet\\Control\\GraphicsDrivers\\Scheduler" /v EnablePreemption 2>nul',
    checkValue: "0x0",
    cmd: 'reg add "HKLM\\SYSTEM\\CurrentControlSet\\Control\\GraphicsDrivers\\Scheduler" /v EnablePreemption /t REG_DWORD /d 0 /f',
    reboot: true,
  },
  {
    id: "mpo-disable", name: "Disable Multi-Plane Overlay (MPO)",
    desc: "MPO causes stuttering and black screen issues on many GPUs. NVIDIA themselves recommend disabling it.",
    check: 'reg query "HKLM\\SOFTWARE\\Microsoft\\Windows\\Dwm" /v OverlayTestMode 2>nul',
    checkValue: "0x5",
    cmd: 'reg add "HKLM\\SOFTWARE\\Microsoft\\Windows\\Dwm" /v OverlayTestMode /t REG_DWORD /d 5 /f',
  },
  {
    id: "flip-model", name: "Force DirectX Flip Model",
    desc: "Flip model presentation reduces latency by ~1 frame vs bitblt. Forces modern swap chain.",
    check: 'reg query "HKCU\\System\\GameConfigStore" /v GameDVR_DXGIHonorFSEWindowsCompatible 2>nul',
    checkValue: "0x1",
    cmd: 'reg add "HKCU\\System\\GameConfigStore" /v GameDVR_DXGIHonorFSEWindowsCompatible /t REG_DWORD /d 1 /f',
  },
];

// ── Input & Mouse Tweaks ────────────────────────────────────────────
const INPUT_TWEAKS = [
  {
    id: "mouse-accel", name: "Disable Mouse Acceleration",
    desc: "Raw mouse input. No acceleration curve, no smoothing. Essential for aim consistency.",
    check: 'reg query "HKCU\\Control Panel\\Mouse" /v MouseSpeed 2>nul',
    checkValue: "0",
    cmd: 'reg add "HKCU\\Control Panel\\Mouse" /v MouseSpeed /t REG_SZ /d 0 /f && reg add "HKCU\\Control Panel\\Mouse" /v MouseThreshold1 /t REG_SZ /d 0 /f && reg add "HKCU\\Control Panel\\Mouse" /v MouseThreshold2 /t REG_SZ /d 0 /f',
  },
  {
    id: "enhance-pointer", name: "Disable Enhance Pointer Precision",
    desc: "Same as mouse accel but through a different registry path. Belt and suspenders.",
    check: 'reg query "HKCU\\Control Panel\\Mouse" /v MouseSensitivity 2>nul',
    checkValue: "10",
    cmd: 'reg add "HKCU\\Control Panel\\Mouse" /v MouseSensitivity /t REG_SZ /d 10 /f',
  },
  {
    id: "usb-polling", name: "Optimize USB Polling",
    desc: "Removes USB power management that can add latency to mouse/keyboard input.",
    check: 'reg query "HKLM\\SYSTEM\\CurrentControlSet\\Control\\Power\\PowerSettings\\2a737441-1930-4402-8d77-b2bebba308a3\\48e6b7a6-50f5-4782-a5d4-53bb8f07e226" /v Attributes 2>nul',
    checkValue: "0x0",
    cmd: 'reg add "HKLM\\SYSTEM\\CurrentControlSet\\Control\\Power\\PowerSettings\\2a737441-1930-4402-8d77-b2bebba308a3\\48e6b7a6-50f5-4782-a5d4-53bb8f07e226" /v Attributes /t REG_DWORD /d 0 /f',
  },
  {
    id: "keyboard-speed", name: "Maximize Keyboard Response",
    desc: "Sets keyboard repeat delay to minimum and repeat rate to maximum.",
    check: 'reg query "HKCU\\Control Panel\\Keyboard" /v KeyboardDelay 2>nul',
    checkValue: "0",
    cmd: 'reg add "HKCU\\Control Panel\\Keyboard" /v KeyboardDelay /t REG_SZ /d 0 /f && reg add "HKCU\\Control Panel\\Keyboard" /v KeyboardSpeed /t REG_SZ /d 31 /f',
  },
];

// ── CPU & Scheduler Tweaks ──────────────────────────────────────────
const CPU_TWEAKS = [
  {
    id: "mmcss-gaming", name: "MMCSS Gaming Priority",
    desc: "Sets Multimedia Class Scheduler to prioritize gaming threads. Guarantees CPU time for your game.",
    check: 'reg query "HKLM\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Multimedia\\SystemProfile\\Tasks\\Games" /v "GPU Priority" 2>nul',
    checkValue: "0x8",
    cmd: 'reg add "HKLM\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Multimedia\\SystemProfile\\Tasks\\Games" /v "GPU Priority" /t REG_DWORD /d 8 /f && reg add "HKLM\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Multimedia\\SystemProfile\\Tasks\\Games" /v Priority /t REG_DWORD /d 6 /f && reg add "HKLM\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Multimedia\\SystemProfile\\Tasks\\Games" /v "Scheduling Category" /t REG_SZ /d High /f && reg add "HKLM\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Multimedia\\SystemProfile\\Tasks\\Games" /v "SFIO Priority" /t REG_SZ /d High /f',
  },
  {
    id: "mmcss-priority", name: "MMCSS System Priority Boost",
    desc: "Increase MMCSS system-wide scheduling priority so game threads get serviced first.",
    check: 'reg query "HKLM\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Multimedia\\SystemProfile" /v SystemResponsiveness 2>nul',
    checkValue: "0x0",
    cmd: 'reg add "HKLM\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Multimedia\\SystemProfile" /v SystemResponsiveness /t REG_DWORD /d 0 /f',
  },
  {
    id: "timer-res", name: "Force High Timer Resolution",
    desc: "Sets Windows timer to 0.5ms resolution. Default is 15.6ms — that's 15ms of wasted precision on every frame.",
    check: 'reg query "HKLM\\SYSTEM\\CurrentControlSet\\Control\\Session Manager\\kernel" /v GlobalTimerResolutionRequests 2>nul',
    checkValue: "0x1",
    cmd: 'reg add "HKLM\\SYSTEM\\CurrentControlSet\\Control\\Session Manager\\kernel" /v GlobalTimerResolutionRequests /t REG_DWORD /d 1 /f',
    reboot: true,
  },
  {
    id: "hpet-disable", name: "Disable HPET",
    desc: "High Precision Event Timer can cause micro-stuttering on some systems. BIOS + Windows side.",
    check: 'bcdedit /enum | findstr /i "useplatformtick" 2>nul',
    checkValue: "Yes",
    cmd: 'bcdedit /deletevalue useplatformclock 2>nul & bcdedit /set useplatformtick yes & bcdedit /set disabledynamictick yes',
    reboot: true,
  },
  {
    id: "core-parking", name: "Disable CPU Core Parking",
    desc: "Prevents Windows from sleeping CPU cores. All cores stay active for maximum thread throughput.",
    check: 'powercfg -query scheme_current sub_processor CPMINCORES 2>nul',
    checkValue: "0x00000064",
    cmd: 'powercfg -setacvalueindex scheme_current sub_processor CPMINCORES 100 && powercfg -setactive scheme_current',
  },
  {
    id: "cpu-priority", name: "CPU Priority Separation",
    desc: "Tells Windows to give foreground apps (your game) more CPU quantum time.",
    check: 'reg query "HKLM\\SYSTEM\\CurrentControlSet\\Control\\PriorityControl" /v Win32PrioritySeparation 2>nul',
    checkValue: "0x26",
    cmd: 'reg add "HKLM\\SYSTEM\\CurrentControlSet\\Control\\PriorityControl" /v Win32PrioritySeparation /t REG_DWORD /d 38 /f',
  },
  {
    id: "spectre-off", name: "Disable Spectre/Meltdown Mitigations",
    desc: "CPU vulnerability patches cost 5-15% performance. If this is a dedicated gaming PC, disable them.",
    check: 'reg query "HKLM\\SYSTEM\\CurrentControlSet\\Control\\Session Manager\\Memory Management" /v FeatureSettingsOverride 2>nul',
    checkValue: "0x3",
    cmd: 'reg add "HKLM\\SYSTEM\\CurrentControlSet\\Control\\Session Manager\\Memory Management" /v FeatureSettingsOverride /t REG_DWORD /d 3 /f && reg add "HKLM\\SYSTEM\\CurrentControlSet\\Control\\Session Manager\\Memory Management" /v FeatureSettingsOverrideMask /t REG_DWORD /d 3 /f',
    reboot: true,
    danger: true,
  },
];

// ── Memory & Disk Tweaks ────────────────────────────────────────────
const MEMORY_TWEAKS = [
  {
    id: "large-pages", name: "Enable Large System Pages",
    desc: "Large memory pages (2MB instead of 4KB) reduce TLB misses. Measurable FPS improvement in some games.",
    check: 'reg query "HKLM\\SYSTEM\\CurrentControlSet\\Control\\Session Manager\\Memory Management" /v LargePageMinimum 2>nul',
    checkValue: "0x1",
    cmd: 'reg add "HKLM\\SYSTEM\\CurrentControlSet\\Control\\Session Manager\\Memory Management" /v LargePageMinimum /t REG_DWORD /d 1 /f',
    reboot: true,
  },
  {
    id: "disable-paging", name: "Optimize Paging Executive",
    desc: "Keep kernel and drivers in RAM instead of paging to disk. Reduces random stutter.",
    check: 'reg query "HKLM\\SYSTEM\\CurrentControlSet\\Control\\Session Manager\\Memory Management" /v DisablePagingExecutive 2>nul',
    checkValue: "0x1",
    cmd: 'reg add "HKLM\\SYSTEM\\CurrentControlSet\\Control\\Session Manager\\Memory Management" /v DisablePagingExecutive /t REG_DWORD /d 1 /f',
  },
  {
    id: "ndu-disable", name: "Disable Network Data Usage Monitor",
    desc: "NDU service leaks memory over time. Known Windows bug that causes gradual performance degradation.",
    check: 'reg query "HKLM\\SYSTEM\\CurrentControlSet\\Services\\Ndu" /v Start 2>nul',
    checkValue: "0x4",
    cmd: 'reg add "HKLM\\SYSTEM\\CurrentControlSet\\Services\\Ndu" /v Start /t REG_DWORD /d 4 /f',
    reboot: true,
  },
  {
    id: "prefetch-off", name: "Disable Prefetch/Superfetch",
    desc: "Prefetch preloads apps but fights with games for I/O bandwidth. Disable for SSDs.",
    check: 'reg query "HKLM\\SYSTEM\\CurrentControlSet\\Control\\Session Manager\\Memory Management\\PrefetchParameters" /v EnablePrefetcher 2>nul',
    checkValue: "0x0",
    cmd: 'reg add "HKLM\\SYSTEM\\CurrentControlSet\\Control\\Session Manager\\Memory Management\\PrefetchParameters" /v EnablePrefetcher /t REG_DWORD /d 0 /f && reg add "HKLM\\SYSTEM\\CurrentControlSet\\Control\\Session Manager\\Memory Management\\PrefetchParameters" /v EnableSuperfetch /t REG_DWORD /d 0 /f',
  },
  {
    id: "last-access", name: "Disable Last Access Timestamp",
    desc: "NTFS updates a timestamp every time a file is read. Disable for less disk overhead.",
    check: 'fsutil behavior query DisableLastAccess 2>nul',
    checkValue: "1",
    cmd: 'fsutil behavior set DisableLastAccess 1',
  },
  {
    id: "trim-enable", name: "Enable SSD TRIM",
    desc: "Ensures SSD TRIM is enabled for optimal SSD performance and longevity.",
    check: 'fsutil behavior query DisableDeleteNotify 2>nul',
    checkValue: "0",
    cmd: 'fsutil behavior set DisableDeleteNotify 0',
  },
];

// ── All Categories ──────────────────────────────────────────────────
const ALL_CATEGORIES = {
  gpu: { name: "GPU & Display", icon: "\u{1F3AE}", tweaks: GPU_TWEAKS },
  input: { name: "Input & Mouse", icon: "\u{1F5B1}\uFE0F", tweaks: INPUT_TWEAKS },
  cpu: { name: "CPU & Scheduler", icon: "\u{2699}\uFE0F", tweaks: CPU_TWEAKS },
  memory: { name: "Memory & Disk", icon: "\u{1F4BE}", tweaks: MEMORY_TWEAKS },
};

// ── Module Class ─────────────────────────────────────────────────────
class GamingTweaks {
  constructor() {
    this.isWin = os.platform() === "win32";
    this.applied = {}; // tweakId -> status
  }

  // Get all tweak categories and their tweaks
  getCategories() {
    return ALL_CATEGORIES;
  }

  // Apply a single tweak then verify it actually took effect
  async applyTweak(tweakId) {
    if (!this.isWin) return { success: false, error: "Windows only" };

    let tweak = null;
    for (const cat of Object.values(ALL_CATEGORIES)) {
      tweak = cat.tweaks.find(t => t.id === tweakId);
      if (tweak) break;
    }
    if (!tweak) return { success: false, error: "Unknown tweak" };

    const result = await runCmd(tweak.cmd);

    // Verify the tweak actually applied by reading back the value
    let verified = false;
    if (result.success && tweak.check) {
      const v = await runCmd(tweak.check);
      verified = v.success && v.output.includes(tweak.checkValue);
    } else if (result.success) {
      verified = true; // No check available, trust the exit code
    }

    this.applied[tweakId] = verified ? "verified" : (result.success ? "applied-unverified" : "failed");

    return {
      success: result.success,
      verified,
      name: tweak.name,
      reboot: tweak.reboot || false,
      danger: tweak.danger || false,
      error: result.error,
    };
  }

  // Apply all tweaks in a category
  async applyCategory(categoryId) {
    const cat = ALL_CATEGORIES[categoryId];
    if (!cat) return { success: false, error: "Unknown category" };

    const results = [];
    let needsReboot = false;
    for (const tweak of cat.tweaks) {
      if (tweak.danger) continue; // Skip dangerous tweaks in bulk apply
      const r = await this.applyTweak(tweak.id);
      results.push({ id: tweak.id, ...r });
      if (r.reboot) needsReboot = true;
    }
    return { results, needsReboot };
  }

  // Apply ALL safe tweaks across all categories
  async applyAll() {
    const allResults = {};
    let needsReboot = false;

    for (const [catId, cat] of Object.entries(ALL_CATEGORIES)) {
      allResults[catId] = [];
      for (const tweak of cat.tweaks) {
        if (tweak.danger) continue;
        const r = await this.applyTweak(tweak.id);
        allResults[catId].push({ id: tweak.id, ...r });
        if (r.reboot) needsReboot = true;
      }
    }

    return { results: allResults, needsReboot };
  }

  // Check current state of a tweak (if it has a check command)
  async checkTweak(tweakId) {
    let tweak = null;
    for (const cat of Object.values(ALL_CATEGORIES)) {
      tweak = cat.tweaks.find(t => t.id === tweakId);
      if (tweak) break;
    }
    if (!tweak || !tweak.check) return { checked: false };

    const result = await runCmd(tweak.check);
    const isApplied = result.output.includes(tweak.checkValue);
    return { checked: true, applied: isApplied };
  }

  getApplied() {
    return this.applied;
  }

  // Verify ALL tweaks — reads back every registry/system value to confirm current state
  async verifyAll() {
    if (!this.isWin) return { error: "Windows only" };

    const results = {};
    let totalChecked = 0, totalApplied = 0, totalNotApplied = 0, totalNoCheck = 0;

    for (const [catId, cat] of Object.entries(ALL_CATEGORIES)) {
      results[catId] = { name: cat.name, tweaks: [] };
      for (const tweak of cat.tweaks) {
        if (!tweak.check) {
          results[catId].tweaks.push({ id: tweak.id, name: tweak.name, status: "no-check" });
          totalNoCheck++;
          continue;
        }
        totalChecked++;
        const v = await runCmd(tweak.check);
        const applied = v.success && v.output.includes(tweak.checkValue);
        if (applied) totalApplied++;
        else totalNotApplied++;
        results[catId].tweaks.push({
          id: tweak.id,
          name: tweak.name,
          status: applied ? "applied" : "not-applied",
          currentValue: v.success ? v.output.split("\n").pop().trim() : "unknown",
          expectedValue: tweak.checkValue,
        });
      }
    }

    return {
      categories: results,
      summary: { totalChecked, totalApplied, totalNotApplied, totalNoCheck },
    };
  }
}

module.exports = GamingTweaks;
