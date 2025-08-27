
# --- Libraries ---
suppressMessages(library(tidyverse))
suppressMessages(library(openxlsx))
suppressMessages(library(lubridate))
suppressMessages(library(readxl))
suppressMessages(library(stringr))

cat("???? R script started...\n")
args <- commandArgs(trailingOnly = TRUE)
cat("???? Args received:", toString(args), "\n")

# --- Capture and validate the argument ---
if (length(args) == 0 || is.na(ymd(args[1]))) {
  stop("??? PAY PERIOD START date was not provided or is invalid. Please provide YYYY-MM-DD format.")
}
pay_period_start <- ymd(args[1])
cat("??? PAY PERIOD START =", as.character(pay_period_start), "\n")

# --- Directories ---
input_dir <- "C:/Reports/Provider Prod Data Pulls"
output_dir <- file.path(input_dir, "Filtered")
if (!dir.exists(output_dir)) dir.create(output_dir, recursive = TRUE)

# --- File Detection ---
files <- list.files(input_dir, pattern = "\\.xlsx$", full.names = TRUE)
prod_file <- files[str_detect(basename(files), regex("Prov Prod Data", ignore_case = TRUE))][1]
appt_file <- files[str_detect(basename(files), regex("Kept|Appt|appointments|NG", ignore_case = TRUE)) &
                     !str_detect(basename(files), regex("PROD|Productivity", ignore_case = TRUE))][1]
specialty_file <- files[str_detect(basename(files), regex("Provider_Productivity_Weeks|special|specialties", ignore_case = TRUE))][1]

if (is.na(prod_file) || !file.exists(prod_file)) stop("??? Productivity file not found.")
if (is.na(appt_file) || !file.exists(appt_file)) stop("??? Appointment file not found.")
if (is.na(specialty_file) || !file.exists(specialty_file)) stop("??? Specialty file not found.")

cat("???? Loading Excel files...\n")
Prod_Raw <- read_excel(prod_file)
cat("??? Prod file loaded.\n")
Appt_Raw <- read_excel(appt_file)
cat("??? Appt file loaded.\n")
Specialty_Data <- read_excel(specialty_file) %>%
  select(`Provider Specialty`, `Provider`, `Productivity Target?`)
cat("??? Specialty file loaded.\n")

# --- Clean and Transform ---
Prod_Cleaned <- Prod_Raw %>% mutate(ISoweek_start_date = isoweek(ymd(week_end_date)))
Charting <- Prod_Cleaned %>% filter(category == 'Charting Time')
Exempt <- Prod_Cleaned %>% filter(`Prevent Appointments?` == 'Y' & !category %in% c('Charting Time', 'Administrative Time'))
NonExempt <- Prod_Cleaned %>% filter(`Prevent Appointments?` == 'N')
Bound <- bind_rows(NonExempt, Charting)

NonExempt_Summary <- Bound %>%
  group_by(ISoweek_start_date, Provider) %>%
  summarise(
    `Total Non Exemption Time (Mins)` = round(sum(duration), 2),
    `Total Non-Exempt Hours On Schedule` = round(sum(duration) / 60, 2),
    week_start_date = first(week_start_date),
    week_end_date = first(week_end_date),
    .groups = "drop"
  )

Exempt_Summary <- Exempt %>%
  group_by(ISoweek_start_date, Provider) %>%
  summarise(
    `Total Exemption Time` = sum(duration),
    `Total Exempt Hours on Schedule` = round(sum(duration) / 60, 2),
    .groups = "drop"
  )

Appt_Summary <- Appt_Raw %>%
  rename(Provider = `Res Name`) %>%
  mutate(ISoweek_start_date = isoweek(ymd(`Appt Dt`))) %>%
  group_by(ISoweek_start_date, Provider) %>%
  summarise(`Total Number of Kept Appointments` = n(), .groups = "drop")

Final_Binding <- NonExempt_Summary %>%
  full_join(Appt_Summary, by = c('Provider', 'ISoweek_start_date')) %>%
  full_join(Exempt_Summary, by = c('Provider', 'ISoweek_start_date')) %>%
  mutate(
    `Total Productivity` = round(`Total Number of Kept Appointments` / `Total Non-Exempt Hours On Schedule`, 4)
  ) %>%
  distinct(Provider, ISoweek_start_date, .keep_all = TRUE)

ProviderSummary <- Final_Binding %>%
  mutate(Two_Week_Group = floor(as.numeric(difftime(ymd(week_start_date), pay_period_start, units = "days")) / 14) + 1) %>%
  group_by(Provider, Two_Week_Group) %>%
  summarise(
    `Total Kept Appointments` = sum(`Total Number of Kept Appointments`, na.rm = TRUE),
    `Total Non-Exempt Hours On Schedule` = sum(`Total Non-Exempt Hours On Schedule`, na.rm = TRUE),
    `Total Exempt Hours on Schedule` = sum(`Total Exempt Hours on Schedule`, na.rm = TRUE),
    `Average Productivity` = round(mean(`Total Number of Kept Appointments` / `Total Non-Exempt Hours On Schedule`, na.rm = TRUE), 2),
    .groups = "drop"
  ) %>%
  full_join(Specialty_Data, by = 'Provider')

FourWeekMapping <- Final_Binding %>%
  mutate(Two_Week_Group = floor(as.numeric(difftime(ymd(week_start_date), pay_period_start, units = "days")) / 14) + 1) %>%
  group_by(Two_Week_Group) %>%
  summarise(
    Min_Date = min(ymd(week_start_date), na.rm = TRUE),
    Max_Date = max(ymd(week_end_date), na.rm = TRUE),
    .groups = "drop"
  ) %>%
  mutate(Two_Week_Label = paste0("Period ", Two_Week_Group, ": ", format(Min_Date, "%Y-%m-%d"), " to ", format(Max_Date + 1, "%Y-%m-%d")))

ProviderSummaryLabeled <- ProviderSummary %>%
  left_join(FourWeekMapping, by = "Two_Week_Group") %>%
  arrange(Two_Week_Group)

Incentive_Payment_Calculation <- ProviderSummaryLabeled %>%
  filter(Two_Week_Group != 0) %>%
  mutate(
    `Number of Non Exempt Hours Per Two Week Period` = `Total Non-Exempt Hours On Schedule`,
    `Encounters Needed To Hit Goal at Prod Target` = `Number of Non Exempt Hours Per Two Week Period` * `Productivity Target?`,
    `Total Encounters Rendered Vs Goal` = round(`Total Kept Appointments` - `Encounters Needed To Hit Goal at Prod Target`, 0),
    `Incentive Payment` = if_else(`Total Encounters Rendered Vs Goal` > 0, `Total Encounters Rendered Vs Goal` * 10, 0)
  ) %>%
  select(Provider, Two_Week_Label, `Total Non-Exempt Hours On Schedule`, `Productivity Target?`,
         `Encounters Needed To Hit Goal at Prod Target`, `Total Kept Appointments`, `Total Encounters Rendered Vs Goal`, `Incentive Payment`) %>%
  distinct(Provider, Two_Week_Label, .keep_all = TRUE)

Incentive_Payment_Pivot <- Incentive_Payment_Calculation %>%
  select(Provider, Two_Week_Label, `Incentive Payment`) %>%
  pivot_wider(names_from = Two_Week_Label, values_from = `Incentive Payment`)

# --- 2.0 Productivity Target ---
ProviderSummaryLabeled_prod_2 <- ProviderSummaryLabeled %>%
  filter(Two_Week_Group != 0) %>%
  mutate(`Productivity Target?` = as.numeric(recode(as.character(`Productivity Target?`), '2.2' = '2.0')))

Incentive_Payment_Calculation_at_2 <- ProviderSummaryLabeled_prod_2 %>%
  mutate(
    `Number of Non Exempt Hours Per Two Week Period` = `Total Non-Exempt Hours On Schedule`,
    `Encounters Needed To Hit Goal at Prod Target` = `Number of Non Exempt Hours Per Two Week Period` * `Productivity Target?`,
    `Total Encounters Rendered Vs Goal` = round(`Total Kept Appointments` - `Encounters Needed To Hit Goal at Prod Target`, 0),
    `Incentive Payment` = if_else(`Total Encounters Rendered Vs Goal` > 0, `Total Encounters Rendered Vs Goal` * 10, 0)
  ) %>%
  select(Provider, Two_Week_Label, `Total Non-Exempt Hours On Schedule`, `Productivity Target?`,
         `Encounters Needed To Hit Goal at Prod Target`, `Total Kept Appointments`, `Total Encounters Rendered Vs Goal`, `Incentive Payment`) %>%
  distinct(Provider, Two_Week_Label, .keep_all = TRUE)

Incentive_Payment_Pivot_at_2 <- Incentive_Payment_Calculation_at_2 %>%
  select(Provider, Two_Week_Label, `Incentive Payment`) %>%
  pivot_wider(names_from = Two_Week_Label, values_from = `Incentive Payment`)

# --- Save Workbooks ---
x2.2_WB <- createWorkbook()
x2.0_WB <- createWorkbook()

addWorksheet(x2.2_WB, 'Incent Payment Calc')
writeDataTable(x2.2_WB, 'Incent Payment Calc', Incentive_Payment_Calculation)
addWorksheet(x2.2_WB, 'Pivot')
writeDataTable(x2.2_WB, 'Pivot', Incentive_Payment_Pivot)

addWorksheet(x2.0_WB, 'Incent Payment Calc')
writeDataTable(x2.0_WB, 'Incent Payment Calc', Incentive_Payment_Calculation_at_2)
addWorksheet(x2.0_WB, 'Pivot')
writeDataTable(x2.0_WB, 'Pivot', Incentive_Payment_Pivot_at_2)

saveWorkbook(x2.2_WB, file = file.path(output_dir, 'Prod Target 2.2 Incentive WB.xlsx'), overwrite = TRUE)
saveWorkbook(x2.0_WB, file = file.path(output_dir, 'Prod Target 2.0 Incentive WB.xlsx'), overwrite = TRUE)

cat("??? Incentive workbooks exported successfully.\n")