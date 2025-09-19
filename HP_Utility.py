#!/usr/bin/env python3
# TCS GPU Driver Toolkit v1
# - Inventories HP + GPU + HWIDs
# - Installs known-good driver from USB library by HWID
# - Blocks Windows Update driver replacements
# - Can export current display driver into the library
# Run as Administrator.

import ctypes, subprocess, sys, os, re, json, datetime, shutil, argparse
from pathlib import Path

APP_NAME = "TCS GPU Driver Toolkit"
ROOT = Path(__file__).resolve().parent           # ...\TCS
USB_ROOT = ROOT                                   # toolkit root is the USB root/TCS
DRIVER_LIB = USB_ROOT / "Drivers" / "Intel"
LOG_DIR = USB_ROOT / "Logs"
LOG_DIR.mkdir(parents=True, exist_ok=True)
LOG_FILE = LOG_DIR / f"log-{datetime.datetime.now().strftime('%Y%m%d-%H%M%S')}.txt"

def log(msg:str):
    stamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{stamp}] {msg}"
    print(line)
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass

def require_admin():
    if not ctypes.windll.shell32.IsUserAnAdmin():
        # Relaunch as admin
        params = " ".join([f'"{a}"' for a in sys.argv])
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, f'"{__file__}" {params}', None, 1)
        sys.exit(0)

def run_ps(script:str) -> str:
    """Run a PowerShell command and return stdout text (trimmed)."""
    cmd = ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script]
    cp = subprocess.run(cmd, capture_output=True, text=True)
    if cp.stderr and cp.returncode != 0:
        log(f"PowerShell error: {cp.stderr.strip()}")
    return (cp.stdout or "").strip()

def run_cmd(args:list) -> subprocess.CompletedProcess:
    cp = subprocess.run(args, capture_output=True, text=True)
    if cp.stdout:
        log(cp.stdout.strip())
    if cp.stderr and cp.returncode != 0:
        log(cp.stderr.strip())
    return cp

def get_system_info():
    ps = r"""
$cs  = Get-CimInstance -ClassName Win32_ComputerSystem
$bios= Get-CimInstance -ClassName Win32_BIOS
$prod= Get-CimInstance -ClassName Win32_ComputerSystemProduct
$gpu = Get-CimInstance -ClassName Win32_VideoController | Select-Object -First 1 Name, DriverVersion, DriverDate
$disp= Get-PnpDevice -Class Display | Where-Object {$_.Status -eq 'OK'} | Select-Object -First 1
$hw  = Get-PnpDeviceProperty -InstanceId $disp.InstanceId -KeyName 'DEVPKEY_Device_HardwareIds'
$ids = @($hw.Data)

# also fetch INF name for current display driver
$pnpsigned = Get-WmiObject Win32_PnPSignedDriver | Where-Object {$_.ClassName -eq 'Display'} | Select-Object -First 1 InfName, DriverVersion

[PSCustomObject]@{
  Manufacturer = $cs.Manufacturer
  Model        = $cs.Model
  Product      = $prod.Name
  BIOSVersion  = $bios.SMBIOSBIOSVersion
  Serial       = $bios.SerialNumber
  GPUName      = $gpu.Name
  GPUDriver    = $gpu.DriverVersion
  GPUDriverDate= $gpu.DriverDate
  HardwareIds  = $ids
  CurrentInf   = $pnpsigned.InfName
  CurrentDrvVer= $pnpsigned.DriverVersion
} | ConvertTo-Json -Compress
"""
    out = run_ps(ps)
    return json.loads(out) if out else {}

def parse_ids(ids):
    # Typical: "PCI\VEN_8086&DEV_8A56&SUBSYS_86AB103C&REV_0C"
    ven, dev, subsys = None, None, None
    for s in ids:
        m = re.search(r"VEN_([0-9A-F]{4})", s, re.I)
        if m: ven = m.group(1).upper()
        m = re.search(r"DEV_([0-9A-F]{4})", s, re.I)
        if m: dev = m.group(1).upper()
        m = re.search(r"SUBSYS_([0-9A-F]{8})", s, re.I)
        if m: subsys = m.group(1).upper()
        if ven and dev and subsys:
            break
    return ven, dev, subsys

def newest_driver_folder(base:Path) -> Path|None:
    # pick newest by last modified time
    candidates = [p for p in base.iterdir() if p.is_dir()]
    if not candidates:
        return None
    candidates.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    return candidates[0]

def find_inf(folder:Path) -> Path|None:
    for p in folder.rglob("*.inf"):
        # prefer Intel display INF names (typical: iigd*.inf, igdlh*.inf)
        if re.match(r"iigd.*\.inf$", p.name, re.I) or re.match(r"igd.*\.inf$", p.name, re.I):
            return p
    # fallback to any inf
    for p in folder.rglob("*.inf"):
        return p
    return None

def install_driver_from_folder(folder:Path) -> bool:
    inf = find_inf(folder)
    if not inf:
        log(f"No .inf found under {folder}")
        return False
    log(f"Installing driver from {inf}")
    # Add + install using pnputil
    cp = run_cmd(["pnputil", "/add-driver", str(inf), "/install"])
    return cp.returncode == 0

def block_windows_driver_updates():
    log("Blocking Windows Update driver delivery (policy & device settings)...")
    # Policy: ExcludeWUDriversInQualityUpdate = 1
    run_ps(r"""
New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate" -Force | Out-Null
New-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate" -Name "ExcludeWUDriversInQualityUpdate" -PropertyType DWord -Value 1 -Force | Out-Null
""")
    # Device Installation Settings UI equivalent: SearchOrderConfig = 0
    run_ps(r"""
New-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\DriverSearching" -Name "SearchOrderConfig" -PropertyType DWord -Value 0 -Force | Out-Null
""")
    # Stop Intel DSA (if present)
    run_ps(r"""
$svc = Get-Service | Where-Object {$_.Name -like 'Intel*Driver*Support*Assistant*' -or $_.DisplayName -like '*Intel*Driver*Support*Assistant*'}
if ($svc) { Stop-Service $svc -Force -ErrorAction SilentlyContinue; Set-Service $svc -StartupType Disabled }
""")
    return True

def export_current_display_driver(dest_root:Path, inf_name:str|None):
    dest_root.mkdir(parents=True, exist_ok=True)
    if not inf_name:
        log("No current INF name detected; exporting ALL drivers may be large. Skipping.")
        return False
    # Export just the current display driver oemXX.inf
    log(f"Exporting current display driver package ({inf_name}) to {dest_root} ...")
    cp = run_cmd(["pnputil", "/export-driver", inf_name, str(dest_root)])
    return cp.returncode == 0

def main():
    parser = argparse.ArgumentParser(description="TCS GPU Driver Toolkit")
    parser.add_argument("--dry-run", action="store_true",
                        help="Run in read-only mode (no install, no registry changes).")
    args = parser.parse_args()

    require_admin()
    log(f"{APP_NAME} starting…")
    if args.dry_run:
        log("=== DRY-RUN MODE ENABLED: No changes will be made ===")

    info = get_system_info()
    if not info:
        log("Failed to gather system info.")
        input("\nPress Enter to exit…")
        sys.exit(1)

    ven, dev, subsys = parse_ids(info.get("HardwareIds"))
    log(f"Detected GPU: {info.get('GPUName')} | VEN={ven} DEV={dev} SUBSYS={subsys or 'N/A'}")

    ven_dev_dir = DRIVER_LIB / f"{ven}_{dev}"
    chosen = None
    if subsys and (ven_dev_dir / f"SUBSYS_{subsys}").exists():
        cands = list_candidate_versions(ven_dev_dir / f"SUBSYS_{subsys}")
        if cands:
            chosen = choose_version_interactively(cands) if len(cands) > 1 else newest(cands)

    if not chosen and ven_dev_dir.exists():
        cands = list_candidate_versions(ven_dev_dir)
        if cands:
            chosen = choose_version_interactively(cands) if len(cands) > 1 else newest(cands)

    if chosen:
        log(f"Matched driver folder: {chosen}")
        if args.dry_run:
            log("DRY-RUN: Would install driver from this folder.")
        else:
            ok = install_driver_from_folder(chosen)
            if not ok:
                log("Install failed. Try another version folder or verify .inf supports this hardware ID.")
                input("\nPress Enter to exit…")
                sys.exit(3)
    else:
        log(f"No matching driver folder found under {ven_dev_dir}")
        if args.dry_run:
            log("DRY-RUN: Would prompt user to export driver if this was a good state.")

    if args.dry_run:
        log("DRY-RUN: Skipping Windows Update blocking and driver export.")
    else:
        block_windows_driver_updates()
        drv_ver = (info.get("CurrentDrvVer") or info.get("GPUDriver") or "unknown").replace(" ", "_")
        export_dir = ven_dev_dir / (f"SUBSYS_{subsys}" if subsys else "") / f"{drv_ver}-exported"
        export_current_display_driver(export_dir, info.get("CurrentInf"))

    log("Done.")
    input("\nPress Enter to exit…")

if __name__ == "__main__":
    main()
