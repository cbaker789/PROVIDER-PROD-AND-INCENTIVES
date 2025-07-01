from Core_SQL_Connection_and_Query import run_main_template_query
from User_Prompt_Functions import R_ScriptRunIncentive, R_Script_4Week, RScript_ISoWeek

if __name__ == "__main__":
    print("Which report would you like to run?")
    print("1 = Provider Productivity File Pull")
    print("2 = Run R Filtering Sequence\n"
          "BEFORE RUNNING THIS OPTION ENSURE THE FOLLOWING are saved to the 'C:/Reports/Provider Prod Pulls Folder':\n"
          "1.) File with 'Kept' in the name is saved\n"
          "2.) File with 'Specialty' is saved\n"
          "3.) File from Option 1 is saved successfully")

    choice = input("Enter 1 or 2: ").strip()

    if choice == "1":
        run_main_template_query()

    elif choice == "2":
        while True:
            print("\n📊 Which R script do you want to run?")
            print("1 = Run Incentive Calculation (Requires Pay Period Date)")
            print("2 = Run 4 Week Interval Workbook Only")
            print("3 = Run ISO Week Workbook Only")
            print("X = Exit back to main menu")

            inner_choice = input("Enter 1, 2, 3, or X: ").strip().lower()

            if inner_choice == "1":
                R_ScriptRunIncentive()
            elif inner_choice == "2":
                R_Script_4Week()
            elif inner_choice == "3":
                RScript_ISoWeek()
            elif inner_choice == "x":
                break
            else:
                print("❌ Invalid selection. Try again.")

    else:
        print("❌ Invalid selection. Exiting.")
