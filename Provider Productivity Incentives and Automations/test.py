import subprocess
import os
import sys
from shutil import which
from datetime import datetime
from pathlib import Path


# =========================
# Project path helpers
# =========================
def project_root() -> Path:
    """Directory that contains this Python file."""
    return Path(__file__).resolve().parent

def scripts_dir() -> Path:
    """
    Local 'Example R Scripts' folder inside the repo.
    NOTE: The folder name below includes a double space after 'and' to match your tree.
    """
    return project_root() / "Provider Productivity and  Incentives Automations" / "Provider Productivity Incentives and Automations" / "Example R Scripts"

def script_path(filename: str) -> Path:
    """Full path to an R script in the local scripts folder, with validation."""
    p = scripts_dir() / filename
    if not p.exists():
        raise FileNotFoundError(
            f"R script not found:\n  {p}\n"
            f"Scripts dir resolved to:\n  {scripts_dir()}\n"
            "Adjust the folder/filename strings if your layout differs."
        )
    return p


# =========================
# Rscript detection
# =========================
def _rscript_from_env():
    """Try R_HOME and PATH for Rscript."""
    r_home = os.environ.get("R_HOME")
    if r_home:
        for rel in (os.path.join("bin", "Rscript.exe"), os.path.join("bin", "Rscript")):
            candidate = os.path.join(r_home, rel)
            if os.path.isfile(candidate):
                return candidate
    candidate = which("Rscript.exe") or which("Rscript")
    if candidate:
        return candidate
    return None

def _rscript_from_registry():
    """Windows-only: look in R-core registry keys for InstallPath."""
    try:
        import winreg  # type: ignore
    except ImportError:
        return None
    keys = [
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\R-core\R"),
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\R-core\R64"),
        (winreg.HKEY_CURRENT_USER,  r"SOFTWARE\R-core\R"),
        (winreg.HKEY_CURRENT_USER,  r"SOFTWARE\R-core\R64"),
    ]
    for hive, subkey in keys:
        try:
            with winreg.OpenKey(hive, subkey) as k:
                r_home, _ = winreg.QueryValueEx(k, "InstallPath")
                for rel in (os.path.join("bin", "Rscript.exe"), os.path.join("bin", "Rscript")):
                    candidate = os.path.join(r_home, rel)
                    if os.path.isfile(candidate):
                        return candidate
        except FileNotFoundError:
            continue
        except OSError:
            continue
    return None

def detect_rscript():
    """
    Detect an Rscript executable for the current user.
    Order: R_HOME -> PATH -> Windows Registry -> common fallbacks.
    """
    cand = _rscript_from_env()
    if cand and os.path.isfile(cand):
        return cand

    cand = _rscript_from_registry()
    if cand and os.path.isfile(cand):
        return cand

    candidates = [
        r"C:/Program Files/R/R-4.4.1/bin/Rscript.exe",
        r"C:/Program Files/R/R-4.4.0/bin/Rscript.exe",
        r"C:/Program Files/R/R-4.3.2/bin/Rscript.exe",
        r"C:/Users/%USERNAME%/AppData/Local/Programs/R/R-4.4.1/bin/Rscript.exe",
        r"C:/Users/%USERNAME%/AppData/Local/Programs/R/R-4.4.0/bin/Rscript.exe",
        r"C:/Users/%USERNAME%/AppData/Local/Programs/R/R-3.6.3/bin/Rscript.exe",
    ]
    for p in candidates:
        p = os.path.expandvars(p)
        if os.path.isfile(p):
            return p

    raise RuntimeError(
        "Could not locate Rscript. Set R_HOME, add Rscript to PATH, or install R "
        "(you can verify by running `Rscript --version`)."
    )


# =========================
# Run helpers
# =========================
def run_rscript(script_file: str, *args: str) -> int:
    """Run an R script from the local scripts folder and stream results."""
    r_exe = detect_rscript()
    spath = script_path(script_file)
    print(f"🕵️ Detected Rscript: {r_exe}")
    print(f"▶️ Running: {spath.name} {' '.join(args) if args else ''}".strip())

    result = subprocess.run(
        [r_exe, str(spath), *args],
        capture_output=True,
        universal_newlines=True
    )

    print("📤 STDOUT:\n", result.stdout or "[no output]")
    print("❌ STDERR:\n", result.stderr or "[no error]")
    print(f"🔚 Exit code: {result.returncode}")
    return result.returncode


# =========================
# Prompts / Commands
# =========================
def prompt_pay_period_date():
    while True:
        date_input = input("📅 Enter PAY PERIOD START (YYYYMMDD): ").strip()
        try:
            return datetime.strptime(date_input, "%Y%m%d").strftime("%Y-%m-%d")
        except ValueError:
            print("❌ Invalid format. Please use YYYYMMDD.")

# Map friendly names to filenames (edit here if your filenames differ)
SCRIPT_FILES = {
    "INCENTIVE": "Incentive_Calc.R",
    "FOUR_WEEK": "Run_4Week.R",
    "ISO_WEEK": "Run_IsoWeek.R",
    "ISO_WEEK_BY_PROVIDER": "Run_ISO_Week_By_Provider.R",
}

def R_ScriptRunIncentive():
    pay_period = prompt_pay_period_date()
    return run_rscript(SCRIPT_FILES["INCENTIVE"], pay_period)

def R_Script_4Week():
    return run_rscript(SCRIPT_FILES["FOUR_WEEK"])

def RScript_ISoWeek():
    return run_rscript(SCRIPT_FILES["ISO_WEEK"])

def RSCRIPT_ISoweek_By_Provider():
    return run_rscript(SCRIPT_FILES["ISO_WEEK_BY_PROVIDER"])


# =========================
# Main menu
# =========================
if __name__ == "__main__":
    print("🧠 Which R script would you like to run?")
    print("1 = Run Incentive Calculation")
    print("2 = Run 4 Week Interval-Only Report")
    print("3 = Run ISo Week-Only Report")
    print("4 = Run ISo Week By Provider")

    choice = input("Enter 1, 2, 3, or 4: ").strip()

    if choice == "1":
        R_ScriptRunIncentive()
    elif choice == "2":
        R_Script_4Week()
    elif choice == "3":
        RScript_ISoWeek()
    elif choice == "4":
        RSCRIPT_ISoweek_By_Provider()
    else:
        print("❌ Invalid selection. Exiting.")
