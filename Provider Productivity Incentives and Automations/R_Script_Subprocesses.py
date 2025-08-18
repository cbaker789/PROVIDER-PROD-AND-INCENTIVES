import subprocess
import os
import sys
from shutil import which
from datetime import datetime

def _rscript_from_env():
    """Try R_HOME and PATH for Rscript."""
    r_home = os.environ.get("R_HOME")
    if r_home:
        # Windows layout
        for rel in (os.path.join("bin", "Rscript.exe"), os.path.join("bin", "Rscript")):
            candidate = os.path.join(r_home, rel)
            if os.path.isfile(candidate):
                return candidate
    # PATH lookup
    candidate = which("Rscript.exe") or which("Rscript")
    if candidate:
        return candidate
    return None

def _rscript_from_registry():
    """Windows-only: look in R-core registry keys for InstallPath."""
    try:
        import winreg
    except ImportError:
        return None
    keys = [
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\\R-core\\R"),
        (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\\R-core\\R64"),
        (winreg.HKEY_CURRENT_USER,  r"SOFTWARE\\R-core\\R"),
        (winreg.HKEY_CURRENT_USER,  r"SOFTWARE\\R-core\\R64"),
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
    """Detect an Rscript executable for the current user.

    Order: R_HOME -> PATH -> Windows Registry -> common fallbacks.
    Returns absolute path or raises RuntimeError with guidance.
    """
    # 1) Env + PATH
    cand = _rscript_from_env()
    if cand and os.path.isfile(cand):
        return cand

    # 2) Windows Registry (if available)
    cand = _rscript_from_registry()
    if cand and os.path.isfile(cand):
        return cand

    # 3) Common fallbacks (adjust as needed)
    candidates = [
        r"C:/Program Files/R/R-4.4.1/bin/Rscript.exe",
        r"C:/Program Files/R/R-4.4.0/bin/Rscript.exe",
        r"C:/Program Files/R/R-4.3.2/bin/Rscript.exe",
        r"C:/Users/%USERNAME%/AppData/Local/Programs/R/R-4.4.1/bin/Rscript.exe",
        r"C:/Users/%USERNAME%/AppData/Local/Programs/R/R-4.4.0/bin/Rscript.exe",
        r"C:/Users/%USERNAME%/AppData/Local/Programs/R/R-4.3.2/bin/Rscript.exe",
    ]
    for p in candidates:
        p = os.path.expandvars(p)
        if os.path.isfile(p):
            return p

    raise RuntimeError(
        "Could not locate Rscript. Set R_HOME, add Rscript to PATH, or install R (you can verify by running `Rscript --version`)."
    )

# --- Ask for Pay Period Start Date ---
def prompt_pay_period_date():
    while True:
        date_input = input("📅 Enter PAY PERIOD START (YYYYMMDD): ").strip()
        try:
            return datetime.strptime(date_input, "%Y%m%d").strftime("%Y-%m-%d")
        except ValueError:
            print("❌ Invalid format. Please use YYYYMMDD.")

# --- Run Incentive Calculation R script with argument ---
def R_ScriptRunIncentive():
    pay_period = prompt_pay_period_date()
    r_exe = detect_rscript()
    print(f"🕵️ Detected Rscript: {r_exe}")
    r_script_path = r"\\SBNC-file1\users\calvin.baker_SBNC\Documents\R Scripts\Prov Prod\Auto_R_Filtering\Incentive_Calc.R"

    print("▶️ Running Incentive R script with date:", pay_period)
    result = subprocess.run(
        [r_exe, r_script_path, pay_period],
        capture_output=True,
        universal_newlines=True
    )

    print("📤 STDOUT:\n", result.stdout or "[no output]")
    print("❌ STDERR:\n", result.stderr or "[no error]")
    print(f"🔚 Exit code: {result.returncode}")

# --- Run Summary-Only R script without arguments ---
def R_Script_4Week():
    r_exe = detect_rscript()
    print(f"🕵️ Detected Rscript: {r_exe}")
    r_script_path = r"\\SBNC-file1\users\calvin.baker_SBNC\Documents\R Scripts\Prov Prod\Auto_R_Filtering\Run_4Week.R"

    print("▶️ Running Summary-Only R script...")
    result = subprocess.run(
        [r_exe, r_script_path],
        capture_output=True,
        universal_newlines=True
    )

    print("📤 STDOUT:\n", result.stdout or "[no output]")
    print("❌ STDERR:\n", result.stderr or "[no error]")
    print(f"🔚 Exit code: {result.returncode}")

# --- Run Summary-Only R script without arguments ---
def RScript_ISoWeek():
    r_exe = detect_rscript()
    print(f"🕵️ Detected Rscript: {r_exe}")
    r_script_path = r"\\SBNC-file1\users\calvin.baker_SBNC\Documents\R Scripts\Prov Prod\Auto_R_Filtering\Run_IsoWeek.R"

    print("▶️ Running Summary-Only R script...")
    result = subprocess.run(
        [r_exe, r_script_path],
        capture_output=True,
        universal_newlines=True
    )

    print("📤 STDOUT:\n", result.stdout or "[no output]")
    print("❌ STDERR:\n", result.stderr or "[no error]")
    print(f"🔚 Exit code: {result.returncode}")


# --- Run Summary-Only R script without arguments ---
def RSCRIPT_ISoweek_By_Provider():
    r_exe = detect_rscript()
    print(f"🕵️ Detected Rscript: {r_exe}")
    r_script_path = r"\\SBNC-file1\users\calvin.baker_SBNC\Documents\R Scripts\Prov Prod\Auto_R_Filtering\ISO_Week_Split_By_Provider.R"

    print("▶️ Running Summary-Only R script...")
    result = subprocess.run(
        [r_exe, r_script_path],
        capture_output=True,
        universal_newlines=True
    )

    print("📤 STDOUT:\n", result.stdout or "[no output]")
    print("❌ STDERR:\n", result.stderr or "[no error]")
    print(f"🔚 Exit code: {result.returncode}")






# --- Main menu ---
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
    elif choice == '3':
        RScript_ISoWeek()
    elif choice == '4':
        RSCRIPT_ISoweek_By_Provider()

    else:
        print("❌ Invalid selection. Exiting.")