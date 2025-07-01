import subprocess
from datetime import datetime

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
    r_exe = r"C:/users/calvin.baker_SBNC/AppData/Local/Programs/R/R-4.3.2/bin/Rscript.exe"
    r_script_path = r"\\SBNC-file1\users\calvin.baker_SBNC\Documents\R Scripts\Prov Prod\Auto_R_Filtering\Incentive_Calc.R"

    print("▶️ Running Incentive R script with date:", pay_period)
    result = subprocess.run(
        [r_exe, r_script_path, pay_period],
        capture_output=True,
        text=True
    )

    print("📤 STDOUT:\n", result.stdout or "[no output]")
    print("❌ STDERR:\n", result.stderr or "[no error]")
    print(f"🔚 Exit code: {result.returncode}")

# --- Run Summary-Only R script without arguments ---
def R_Script_4Week():
    r_exe = r"C:/Users/calvin.baker_SBNC/AppData/Local/Programs/R/R-4.3.2/bin/Rscript.exe"
    r_script_path = r"\\SBNC-file1\users\calvin.baker_SBNC\Documents\R Scripts\Prov Prod\Auto_R_Filtering\Run_4Week.R"

    print("▶️ Running Summary-Only R script...")
    result = subprocess.run(
        [r_exe, r_script_path],
        capture_output=True,
        text=True
    )

    print("📤 STDOUT:\n", result.stdout or "[no output]")
    print("❌ STDERR:\n", result.stderr or "[no error]")
    print(f"🔚 Exit code: {result.returncode}")

# --- Run Summary-Only R script without arguments ---
def RScript_ISoWeek():
    r_exe = r"C:/Users/calvin.baker_SBNC/AppData/Local/Programs/R/R-4.3.2/bin/Rscript.exe"
    r_script_path = r"\\SBNC-file1\users\calvin.baker_SBNC\Documents\R Scripts\Prov Prod\Auto_R_Filtering\Run_IsoWeek.R"

    print("▶️ Running Summary-Only R script...")
    result = subprocess.run(
        [r_exe, r_script_path],
        capture_output=True,
        text=True
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


    choice = input("Enter 1,2, or 3: ").strip()

    if choice == "1":
        R_ScriptRunIncentive()
    elif choice == "2":
        R_Script_4Week()
    elif choice == '3':
        RScript_ISoWeek()

    else:
        print("❌ Invalid selection. Exiting.")