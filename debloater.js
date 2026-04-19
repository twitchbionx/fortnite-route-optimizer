// ═══════════════════════════════════════════════════════════════════
//  WINDOWS DEBLOATER MODULE
//  Strips Windows down to a lean gaming machine.
//  Removes bloatware, kills telemetry, disables useless services,
//  sets power plan to Ultra, and cleans temp junk.
// ═══════════════════════════════════════════════════════════════════

const { exec } = require("child_process");
const os = require("os");

function run(cmd, timeout = 30000) {
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

function runCmd(cmd, timeout = 30000) {
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

// ── Bloatware Apps ───────────────────────────────────────────────────
// These are the Windows Store apps that have zero gaming value
const BLOATWARE_APPS = [
  { id: "cortana", pkg: "Microsoft.549981C3F5F10", name: "Cortana", desc: "Voice assistant nobody asked for. Uses CPU and RAM in background." },
  { id: "bing-news", pkg: "Microsoft.BingNews", name: "Bing News", desc: "News widget that burns CPU cycles refreshing stories you'll never read." },
  { id: "bing-weather", pkg: "Microsoft.BingWeather", name: "Bing Weather", desc: "Look out the window instead. Frees up background resources." },
  { id: "get-help", pkg: "Microsoft.GetHelp", name: "Get Help", desc: "Microsoft's help app. You have Google for that." },
  { id: "getstarted", pkg: "Microsoft.Getstarted", name: "Tips", desc: "Windows tips that pop up and steal focus mid-game." },
  { id: "maps", pkg: "Microsoft.WindowsMaps", name: "Windows Maps", desc: "Nobody uses this. Just uses disk space." },
  { id: "messaging", pkg: "Microsoft.Messaging", name: "Messaging", desc: "Dead messaging app. Resources wasted." },
  { id: "mixedreality", pkg: "Microsoft.MixedReality.Portal", name: "Mixed Reality Portal", desc: "VR portal service running for no reason." },
  { id: "office-hub", pkg: "Microsoft.MicrosoftOfficeHub", name: "Office Hub", desc: "Office ad launcher. Not needed if you have Office installed." },
  { id: "onenote", pkg: "Microsoft.Office.OneNote", name: "OneNote", desc: "Note-taking app with background sync eating bandwidth." },
  { id: "paint3d", pkg: "Microsoft.MSPaint", name: "Paint 3D", desc: "3D paint app nobody uses. Free up the disk space." },
  { id: "people", pkg: "Microsoft.People", name: "People", desc: "Contact app with background sync. Pointless for gaming." },
  { id: "skype", pkg: "Microsoft.SkypeApp", name: "Skype", desc: "Legacy chat app running in background. Use Discord instead." },
  { id: "solitaire", pkg: "Microsoft.MicrosoftSolitaireCollection", name: "Solitaire", desc: "Card games with ads. Takes up space." },
  { id: "sticky-notes", pkg: "Microsoft.MicrosoftStickyNotes", name: "Sticky Notes", desc: "Background sync for sticky notes you don't use." },
  { id: "feedback", pkg: "Microsoft.WindowsFeedbackHub", name: "Feedback Hub", desc: "Sends telemetry to Microsoft. Zero gaming value." },
  { id: "your-phone", pkg: "Microsoft.YourPhone", name: "Phone Link", desc: "Phone mirroring service burning CPU in background." },
  { id: "xbox-gamebar", pkg: "Microsoft.XboxGamingOverlay", name: "Xbox Game Bar", desc: "Overlay that causes FPS drops and input lag. Use built-in Fortnite stats." },
  { id: "xbox-app", pkg: "Microsoft.GamingApp", name: "Xbox App", desc: "Xbox services running in background even if you don't use Xbox." },
  { id: "xbox-identity", pkg: "Microsoft.XboxIdentityProvider", name: "Xbox Identity", desc: "Xbox auth service. Only needed if you play Xbox games on PC." },
  { id: "xbox-speech", pkg: "Microsoft.XboxSpeechToTextOverlay", name: "Xbox Speech", desc: "Speech-to-text overlay. Uses CPU for nothing." },
  { id: "zune-music", pkg: "Microsoft.ZuneMusic", name: "Groove Music", desc: "Dead music app. Use Spotify instead." },
  { id: "zune-video", pkg: "Microsoft.ZuneVideo", name: "Movies & TV", desc: "Video player nobody uses. VLC exists." },
  { id: "clipchamp", pkg: "Clipchamp.Clipchamp", name: "Clipchamp", desc: "Video editor running services in background." },
  { id: "todos", pkg: "Microsoft.Todos", name: "Microsoft To Do", desc: "To-do list with background sync." },
  { id: "powerautomate", pkg: "Microsoft.PowerAutomateDesktop", name: "Power Automate", desc: "Automation tool you don't need. Background service." },
  { id: "widgets", pkg: "Microsoft.WidgetsPlatformRuntime", name: "Widgets Runtime", desc: "Powers the widgets panel that eats RAM." },
  { id: "family", pkg: "MicrosoftCorporationII.MicrosoftFamily", name: "Microsoft Family", desc: "Parental controls service running for no reason." },
  { id: "teams", pkg: "MicrosoftTeams", name: "Microsoft Teams", desc: "Auto-starts and hogs RAM. Close it if you don't need it." },
  { id: "calculator", pkg: "Microsoft.WindowsCalculator", name: "Calculator", desc: "Built-in calculator. Low impact but not needed for gaming." },
  { id: "camera", pkg: "Microsoft.WindowsCamera", name: "Camera", desc: "Camera app with background access. Not needed for gaming." },
  { id: "alarms", pkg: "Microsoft.WindowsAlarms", name: "Alarms & Clock", desc: "Clock/alarms app running in background. Use your phone instead." },
  { id: "commapps", pkg: "microsoft.windowscommunicationsapps", name: "Mail & Calendar", desc: "Built-in Mail and Calendar apps with background sync eating resources." },
  { id: "photos", pkg: "Microsoft.Windows.Photos", name: "Photos", desc: "Photo viewer with background indexing. Use a lightweight viewer instead." },
  { id: "soundrecorder", pkg: "Microsoft.WindowsSoundRecorder", name: "Sound Recorder", desc: "Voice recorder app. Zero gaming value." },
  { id: "screensketch", pkg: "Microsoft.ScreenSketch", name: "Snipping Tool", desc: "Screenshot tool running in background. Use Print Screen or ShareX instead." },
  { id: "xboxapp", pkg: "Microsoft.XboxApp", name: "Xbox Console Companion", desc: "Legacy Xbox companion app. Background services waste resources." },
];

// ── Services to Disable ─────────────────────────────────────────────
// Services that waste resources and have zero gaming benefit
const USELESS_SERVICES = [
  { id: "sysmain", svc: "SysMain", name: "SysMain (Superfetch)", desc: "Pre-loads apps into RAM. Causes stutters when it conflicts with games. Disable it — you have enough RAM." },
  { id: "search", svc: "WSearch", name: "Windows Search Indexer", desc: "Constantly indexes files in background. Major source of random disk spikes during gameplay." },
  { id: "diagtrack", svc: "DiagTrack", name: "Connected User Experiences", desc: "Microsoft telemetry service. Sends data home constantly. Pure waste." },
  { id: "dmwappush", svc: "dmwappushservice", name: "WAP Push Service", desc: "Telemetry push service. Another data collector with zero value." },
  { id: "mapbroker", svc: "MapsBroker", name: "Downloaded Maps Manager", desc: "Manages offline maps nobody uses. Background disk writes." },
  { id: "retaildemo", svc: "RetailDemo", name: "Retail Demo Service", desc: "Store demo mode. Why is this even running." },
  { id: "wisvc", svc: "wisvc", name: "Windows Insider Service", desc: "Insider preview updates. Disable unless you're testing preview builds." },
  { id: "xbl-auth", svc: "XblAuthManager", name: "Xbox Live Auth", desc: "Xbox Live authentication. Only needed for Xbox/Game Pass games." },
  { id: "xbl-save", svc: "XblGameSave", name: "Xbox Live Game Save", desc: "Cloud saves for Xbox games. Fortnite has its own cloud saves." },
  { id: "xboxnet", svc: "XboxNetApiSvc", name: "Xbox Live Networking", desc: "Xbox network service. Fortnite doesn't use Xbox networking." },
  { id: "fax", svc: "Fax", name: "Fax Service", desc: "A fax machine. In 2024. On your gaming PC." },
  { id: "print-spooler", svc: "Spooler", name: "Print Spooler", desc: "Manages print jobs. Disable if you don't print from this PC." },
  { id: "remote-registry", svc: "RemoteRegistry", name: "Remote Registry", desc: "Allows remote registry editing. Security risk + resource waste." },
  { id: "wmp-share", svc: "WMPNetworkSvc", name: "WMP Network Sharing", desc: "Windows Media Player network sharing. Dead technology." },
  { id: "phone-svc", svc: "PhoneSvc", name: "Phone Service", desc: "Telephony service. Your PC isn't a phone." },
  { id: "tablet-input", svc: "TabletInputService", name: "Touch Keyboard", desc: "Touch/tablet input service. Disable if you use mouse/keyboard." },
];

// ── Telemetry Settings ──────────────────────────────────────────────
const TELEMETRY_TWEAKS = [
  { id: "tele-basic", name: "Set Telemetry to Minimum", desc: "Reduce Windows diagnostic data to minimum required level.",
    check: 'reg query "HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\DataCollection" /v AllowTelemetry 2>nul', checkValue: "0x0",
    cmd: 'reg add "HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\DataCollection" /v AllowTelemetry /t REG_DWORD /d 0 /f' },
  { id: "tele-feedback", name: "Disable Feedback Notifications", desc: "Stop Windows from asking for feedback during gameplay.",
    check: 'reg query "HKCU\\SOFTWARE\\Microsoft\\Siuf\\Rules" /v NumberOfSIUFInPeriod 2>nul', checkValue: "0x0",
    cmd: 'reg add "HKCU\\SOFTWARE\\Microsoft\\Siuf\\Rules" /v NumberOfSIUFInPeriod /t REG_DWORD /d 0 /f' },
  { id: "tele-advertising", name: "Disable Advertising ID", desc: "Stop ad tracking. Zero gaming benefit, pure privacy invasion.",
    check: 'reg query "HKCU\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\AdvertisingInfo" /v Enabled 2>nul', checkValue: "0x0",
    cmd: 'reg add "HKCU\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\AdvertisingInfo" /v Enabled /t REG_DWORD /d 0 /f' },
  { id: "tele-activity", name: "Disable Activity History", desc: "Stop Windows from tracking your activity and sending it to Microsoft.",
    check: 'reg query "HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\System" /v EnableActivityFeed 2>nul', checkValue: "0x0",
    cmd: 'reg add "HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\System" /v EnableActivityFeed /t REG_DWORD /d 0 /f && reg add "HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\System" /v PublishUserActivities /t REG_DWORD /d 0 /f' },
  { id: "tele-location", name: "Disable Location Tracking", desc: "Games don't need GPS. Stop background location polling.",
    check: 'reg query "HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\LocationAndSensors" /v DisableLocation 2>nul', checkValue: "0x0",
    cmd: 'reg add "HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\LocationAndSensors" /v DisableLocation /t REG_DWORD /d 0 /f' },
  { id: "tele-wifi-sense", name: "Disable WiFi Sense", desc: "Stop auto-sharing WiFi passwords with contacts.",
    check: 'reg query "HKLM\\SOFTWARE\\Microsoft\\WcmSvc\\wifinetworkmanager\\config" /v AutoConnectAllowedOEM 2>nul', checkValue: "0x0",
    cmd: 'reg add "HKLM\\SOFTWARE\\Microsoft\\WcmSvc\\wifinetworkmanager\\config" /v AutoConnectAllowedOEM /t REG_DWORD /d 0 /f' },
  { id: "tele-error-report", name: "Disable Error Reporting", desc: "Stop Windows Error Reporting from uploading crash data in background.",
    check: 'reg query "HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\Windows Error Reporting" /v Disabled 2>nul', checkValue: "0x1",
    cmd: 'reg add "HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\Windows Error Reporting" /v Disabled /t REG_DWORD /d 1 /f' },
];

// ── Visual Effects & Power ──────────────────────────────────────────
const PERFORMANCE_TWEAKS = [
  { id: "power-ultimate", name: "Ultimate Performance Power Plan", desc: "Forces CPU to max speed at all times. No throttling, no power saving.",
    check: 'powercfg -getactivescheme 2>nul', checkValue: "e9a42b02-d5df-448d-aa00-03f14749eb61",
    cmd: 'powercfg -duplicatescheme e9a42b02-d5df-448d-aa00-03f14749eb61 2>nul & powercfg -setactive e9a42b02-d5df-448d-aa00-03f14749eb61 2>nul || powercfg -setactive 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c' },
  { id: "visual-perf", name: "Disable Visual Effects", desc: "Kill transparency, animations, shadows — pure performance mode.",
    check: 'reg query "HKCU\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\VisualEffects" /v VisualFXSetting 2>nul', checkValue: "0x2",
    cmd: 'reg add "HKCU\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\VisualEffects" /v VisualFXSetting /t REG_DWORD /d 2 /f && reg add "HKCU\\Control Panel\\Desktop" /v UserPreferencesMask /t REG_BINARY /d 9012038010000000 /f' },
  { id: "anim-disable", name: "Disable Window Animations", desc: "No more slide/fade animations stealing CPU time.",
    check: 'reg query "HKCU\\Control Panel\\Desktop\\WindowMetrics" /v MinAnimate 2>nul', checkValue: "0",
    cmd: 'reg add "HKCU\\Control Panel\\Desktop\\WindowMetrics" /v MinAnimate /t REG_SZ /d 0 /f' },
  { id: "transparency", name: "Disable Transparency", desc: "Transparency effects use GPU resources that should go to Fortnite.",
    check: 'reg query "HKCU\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize" /v EnableTransparency 2>nul', checkValue: "0x0",
    cmd: 'reg add "HKCU\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize" /v EnableTransparency /t REG_DWORD /d 0 /f' },
  { id: "game-mode", name: "Enable Game Mode", desc: "Tells Windows to prioritize your game. Reduces background interruptions.",
    check: 'reg query "HKCU\\SOFTWARE\\Microsoft\\GameBar" /v AutoGameModeEnabled 2>nul', checkValue: "0x1",
    cmd: 'reg add "HKCU\\SOFTWARE\\Microsoft\\GameBar" /v AutoGameModeEnabled /t REG_DWORD /d 1 /f' },
  { id: "game-dvr", name: "Disable Game DVR", desc: "Game DVR records gameplay in background — massive FPS hit. Kill it.",
    check: 'reg query "HKCU\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\GameDVR" /v AppCaptureEnabled 2>nul', checkValue: "0x0",
    cmd: 'reg add "HKCU\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\GameDVR" /v AppCaptureEnabled /t REG_DWORD /d 0 /f && reg add "HKLM\\SOFTWARE\\Policies\\Microsoft\\Windows\\GameDVR" /v AllowGameDVR /t REG_DWORD /d 0 /f' },
  { id: "notif-disable", name: "Disable Notifications", desc: "No more toast notifications stealing focus from your game.",
    check: 'reg query "HKCU\\SOFTWARE\\Policies\\Microsoft\\Windows\\Explorer" /v DisableNotificationCenter 2>nul', checkValue: "0x1",
    cmd: 'reg add "HKCU\\SOFTWARE\\Policies\\Microsoft\\Windows\\Explorer" /v DisableNotificationCenter /t REG_DWORD /d 1 /f' },
  { id: "bg-apps", name: "Disable Background Apps", desc: "Stop UWP apps from running in background eating resources.",
    check: 'reg query "HKCU\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\BackgroundAccessApplications" /v GlobalUserDisabled 2>nul', checkValue: "0x1",
    cmd: 'reg add "HKCU\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\BackgroundAccessApplications" /v GlobalUserDisabled /t REG_DWORD /d 1 /f' },
  { id: "startup-delay", name: "Remove Startup Delay", desc: "Windows adds a 10 second delay to startup apps. Remove it.",
    check: 'reg query "HKCU\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Serialize" /v StartupDelayInMSec 2>nul', checkValue: "0x0",
    cmd: 'reg add "HKCU\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Serialize" /v StartupDelayInMSec /t REG_DWORD /d 0 /f' },
  { id: "temp-clean", name: "Clean Temp Files", desc: "Purge temp folders. Frees disk space and reduces background I/O.",
    cmd: 'del /q/f/s %TEMP%\\* 2>nul & del /q/f/s C:\\Windows\\Temp\\* 2>nul & del /q/f/s C:\\Windows\\Prefetch\\* 2>nul' },
];

// ── Module Class ─────────────────────────────────────────────────────
class Debloater {
  constructor() {
    this.isWin = os.platform() === "win32";
  }

  // Get list of installed bloatware
  async scanBloatware() {
    if (!this.isWin) return { apps: [], error: "Windows only" };

    const result = await run("Get-AppxPackage | Select-Object Name | ConvertTo-Json");
    if (!result.success) return { apps: BLOATWARE_APPS.map(a => ({ ...a, installed: false })), error: result.error };

    let installed = [];
    try {
      const parsed = JSON.parse(result.output);
      installed = (Array.isArray(parsed) ? parsed : [parsed]).map(p => p.Name);
    } catch(e) {
      installed = [];
    }

    return {
      apps: BLOATWARE_APPS.map(a => ({
        ...a,
        installed: installed.some(name => name && a.pkg && name.toLowerCase().includes(a.pkg.toLowerCase())),
      })),
    };
  }

  // Remove a specific bloatware app then verify it's gone
  async removeBloatware(appId) {
    if (!this.isWin) return { success: false, error: "Windows only" };
    const app = BLOATWARE_APPS.find(a => a.id === appId);
    if (!app) return { success: false, error: "Unknown app" };

    const result = await run(`Get-AppxPackage *${app.pkg}* | Remove-AppxPackage -ErrorAction SilentlyContinue`);
    // Also prevent reinstall
    await run(`Get-AppxProvisionedPackage -Online | Where-Object {$_.PackageName -like "*${app.pkg}*"} | Remove-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue`);

    // Verify removal — check if the package is still installed
    const verify = await run(`Get-AppxPackage *${app.pkg}* | Select-Object Name | ConvertTo-Json`);
    const stillInstalled = verify.success && verify.output && verify.output.includes(app.pkg);

    return { success: !stillInstalled, verified: !stillInstalled, app: app.name };
  }

  // Remove ALL bloatware at once
  async removeAllBloatware() {
    const results = [];
    for (const app of BLOATWARE_APPS) {
      const r = await this.removeBloatware(app.id);
      results.push({ ...r, id: app.id });
    }
    return results;
  }

  // Get service statuses
  async scanServices() {
    if (!this.isWin) return { services: USELESS_SERVICES.map(s => ({ ...s, status: "unknown" })) };

    const result = await run("Get-Service | Select-Object Name,Status,StartType | ConvertTo-Json");
    let svcList = [];
    try {
      svcList = JSON.parse(result.output);
      if (!Array.isArray(svcList)) svcList = [svcList];
    } catch(e) {}

    return {
      services: USELESS_SERVICES.map(s => {
        const found = svcList.find(sv => sv.Name === s.svc);
        return {
          ...s,
          status: found ? (found.Status === 4 ? "running" : found.Status === 1 ? "stopped" : String(found.Status)) : "not-found",
          startType: found ? found.StartType : "unknown",
          exists: !!found,
        };
      }),
    };
  }

  // Disable a service then verify it's actually disabled
  async disableService(svcId) {
    if (!this.isWin) return { success: false, error: "Windows only" };
    const svc = USELESS_SERVICES.find(s => s.id === svcId);
    if (!svc) return { success: false, error: "Unknown service" };

    const r1 = await runCmd(`sc stop "${svc.svc}" 2>nul`, 10000);
    const r2 = await runCmd(`sc config "${svc.svc}" start=disabled`, 10000);

    // Verify: query the service start type
    const verify = await runCmd(`sc qc "${svc.svc}" 2>nul`, 5000);
    const verified = verify.success && verify.output.includes("DISABLED");

    return { success: r2.success, verified, service: svc.name, error: r2.error };
  }

  // Disable ALL useless services
  async disableAllServices() {
    const results = [];
    for (const svc of USELESS_SERVICES) {
      const r = await this.disableService(svc.id);
      results.push({ ...r, id: svc.id });
    }
    return results;
  }

  // Apply a telemetry tweak then verify
  async applyTelemetryTweak(tweakId) {
    if (!this.isWin) return { success: false, error: "Windows only" };
    const tweak = TELEMETRY_TWEAKS.find(t => t.id === tweakId);
    if (!tweak) return { success: false, error: "Unknown tweak" };

    const result = await runCmd(tweak.cmd, 10000);
    let verified = false;
    if (result.success && tweak.check) {
      const v = await runCmd(tweak.check);
      verified = v.success && v.output.includes(tweak.checkValue);
    } else if (result.success) {
      verified = true;
    }
    return { success: result.success, verified, name: tweak.name, error: result.error };
  }

  // Apply ALL telemetry tweaks
  async applyAllTelemetry() {
    const results = [];
    for (const tweak of TELEMETRY_TWEAKS) {
      const r = await this.applyTelemetryTweak(tweak.id);
      results.push({ ...r, id: tweak.id });
    }
    return results;
  }

  // Apply a performance tweak then verify
  async applyPerfTweak(tweakId) {
    if (!this.isWin) return { success: false, error: "Windows only" };
    const tweak = PERFORMANCE_TWEAKS.find(t => t.id === tweakId);
    if (!tweak) return { success: false, error: "Unknown tweak" };

    const result = await runCmd(tweak.cmd, 15000);
    let verified = false;
    if (result.success && tweak.check) {
      const v = await runCmd(tweak.check);
      verified = v.success && v.output.includes(tweak.checkValue);
    } else if (result.success) {
      verified = true;
    }
    return { success: result.success, verified, name: tweak.name, error: result.error };
  }

  // Apply ALL performance tweaks
  async applyAllPerf() {
    const results = [];
    for (const tweak of PERFORMANCE_TWEAKS) {
      const r = await this.applyPerfTweak(tweak.id);
      results.push({ ...r, id: tweak.id });
    }
    return results;
  }

  // NUKE IT ALL — full debloat in one shot
  async fullDebloat() {
    const results = {
      bloatware: await this.removeAllBloatware(),
      services: await this.disableAllServices(),
      telemetry: await this.applyAllTelemetry(),
      performance: await this.applyAllPerf(),
    };
    return results;
  }

  // Discover ALL installed UWP/Store apps (not just predefined ones)
  async scanAllApps() {
    if (!this.isWin) return { apps: [], error: "Windows only" };

    const result = await run("Get-AppxPackage | Where-Object {$_.IsFramework -eq $false} | Select-Object Name,PackageFullName | ConvertTo-Json");
    if (!result.success) return { apps: [], error: result.error };

    let allApps = [];
    try {
      const parsed = JSON.parse(result.output);
      allApps = Array.isArray(parsed) ? parsed : [parsed];
    } catch(e) {
      return { apps: [], error: "Failed to parse app list" };
    }

    // Build a lookup of known bloatware by package name (lowercase)
    const knownMap = {};
    for (const app of BLOATWARE_APPS) {
      knownMap[app.pkg.toLowerCase()] = app;
    }

    return {
      apps: allApps.map(a => {
        const name = a.Name || "";
        const known = knownMap[name.toLowerCase()];
        return {
          id: known ? known.id : name.toLowerCase().replace(/[^a-z0-9]/g, "-"),
          pkg: name,
          fullName: a.PackageFullName || "",
          name: known ? known.name : name.replace(/Microsoft\.|Windows\./g, "").replace(/\./g, " "),
          desc: known ? known.desc : "Installed UWP app.",
          installed: true,
          knownBloatware: !!known,
        };
      }),
    };
  }

  // Get lists for the UI
  getLists() {
    return {
      bloatware: BLOATWARE_APPS,
      services: USELESS_SERVICES,
      telemetry: TELEMETRY_TWEAKS,
      performance: PERFORMANCE_TWEAKS,
    };
  }

  // Verify everything — scan bloatware, services, telemetry, and performance tweaks
  async verifyAll() {
    if (!this.isWin) return { error: "Windows only" };

    // Check which bloatware is still installed
    const appScan = await this.scanBloatware();
    const appsStillInstalled = appScan.apps ? appScan.apps.filter(a => a.installed) : [];

    // Check which services are still running/enabled
    const svcScan = await this.scanServices();
    const svcsStillEnabled = svcScan.services ? svcScan.services.filter(s => s.exists && s.startType !== 4 && s.startType !== "Disabled") : [];

    // Check telemetry tweaks
    const telemetryResults = [];
    for (const tweak of TELEMETRY_TWEAKS) {
      if (!tweak.check) { telemetryResults.push({ id: tweak.id, name: tweak.name, status: "no-check" }); continue; }
      const v = await runCmd(tweak.check);
      const applied = v.success && v.output.includes(tweak.checkValue);
      telemetryResults.push({ id: tweak.id, name: tweak.name, status: applied ? "applied" : "not-applied" });
    }

    // Check performance tweaks
    const perfResults = [];
    for (const tweak of PERFORMANCE_TWEAKS) {
      if (!tweak.check) { perfResults.push({ id: tweak.id, name: tweak.name, status: "no-check" }); continue; }
      const v = await runCmd(tweak.check);
      const applied = v.success && v.output.includes(tweak.checkValue);
      perfResults.push({ id: tweak.id, name: tweak.name, status: applied ? "applied" : "not-applied" });
    }

    const telApplied = telemetryResults.filter(t => t.status === "applied").length;
    const perfApplied = perfResults.filter(t => t.status === "applied").length;

    return {
      bloatware: { total: BLOATWARE_APPS.length, stillInstalled: appsStillInstalled.length, apps: appsStillInstalled },
      services: { total: USELESS_SERVICES.length, stillEnabled: svcsStillEnabled.length, services: svcsStillEnabled },
      telemetry: { total: TELEMETRY_TWEAKS.length, applied: telApplied, tweaks: telemetryResults },
      performance: { total: PERFORMANCE_TWEAKS.length, applied: perfApplied, tweaks: perfResults },
      summary: {
        bloatwareRemoved: BLOATWARE_APPS.length - appsStillInstalled.length,
        servicesDisabled: USELESS_SERVICES.length - svcsStillEnabled.length,
        telemetryApplied: telApplied,
        perfApplied: perfApplied,
      },
    };
  }
}

module.exports = Debloater;
