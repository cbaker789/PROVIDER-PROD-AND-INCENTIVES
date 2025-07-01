# --- Libraries ---
library(tidyverse)
library(openxlsx)
library(lubridate)
library(readxl)
library(stringr)

# --- Directories ---
input_dir <- "C:/Reports/Provider Prod Data Pulls"
output_dir <- file.path(input_dir, "Filtered/By 4 Week")
if (!dir.exists(output_dir)) dir.create(output_dir, recursive = TRUE)

# --- File Detection ---
files <- list.files(input_dir, pattern = "\\.xlsx$", full.names = TRUE)
prod_file <- files[str_detect(basename(files), regex("Prov Prod Data", ignore_case = TRUE))][1]
appt_file <- files[str_detect(basename(files), regex("Kept_Appointments|NG Kept|Appt|appointments", ignore_case = TRUE)) &
                     !str_detect(basename(files), regex("PROD|Productivity", ignore_case = TRUE))][1]
specialty_file <- files[str_detect(basename(files), regex("Provider_Productivity_Weeks|special", ignore_case = TRUE))][1]

cat("??? Matched files:\n")
cat("   - Productivity:", prod_file, "\n")
cat("   - Appointments:", appt_file, "\n")
cat("   - Specialty:", specialty_file, "\n")

# --- Load Data ---
Specialty_Data <- read_excel(specialty_file) %>%
  select(Provider, `Productivity Target?`)

Prod_Cleaned <- read_excel(prod_file) %>%
  mutate(ISoweek_start_date = isoweek(ymd(week_end_date)))

Appt_Cleaned <- read_excel(appt_file) %>%
  rename(Provider = `Res Name`) %>%
  mutate(ISoweek_start_date = isoweek(ymd(`Appt Dt`))) %>%
  group_by(ISoweek_start_date, Provider) %>%
  summarise(`Total Number of Kept Appointments` = n(), .groups = "drop")

# --- Time Categories ---
Charting_Time_Calc <- Prod_Cleaned %>% filter(category == "Charting Time")
Excemption_Calc <- Prod_Cleaned %>%
  filter(`Prevent Appointments?` == "Y" & !category %in% c("Charting Time", "Administrative Time"))
Non_Excemption_Calc <- Prod_Cleaned %>% filter(`Prevent Appointments?` == "N")
Bound_Charting_Time_Non_Excemption_Time <- bind_rows(Non_Excemption_Calc, Charting_Time_Calc)

Non_Exempt_Time_Calc_By_Provider <- Bound_Charting_Time_Non_Excemption_Time %>%
  group_by(ISoweek_start_date, Provider) %>%
  summarise(
    `Total Non Exemption Time (Mins)` = round(sum(duration), 2),
    `Total Non-Exempt Hours On Schedule` = round(sum(duration) / 60, 2),
    week_start_date = first(week_start_date),
    week_end_date = first(week_end_date),
    .groups = "drop"
  )

Exempt_Time_Calc_By_Provider <- Excemption_Calc %>%
  group_by(ISoweek_start_date, Provider) %>%
  summarise(
    `Total Exemption Time` = sum(duration),
    `Total Exempt Hours on Schedule` = round(`Total Exemption Time` / 60, 2),
    .groups = "drop"
  )

# --- Final Binding ---
Final_Binding <- Non_Exempt_Time_Calc_By_Provider %>%
  full_join(Appt_Cleaned, by = c("Provider", "ISoweek_start_date")) %>%
  full_join(Exempt_Time_Calc_By_Provider, by = c("Provider", "ISoweek_start_date")) %>%
  mutate(
    `Total Productivity` = ifelse(
      is.na(`Total Number of Kept Appointments`) |
        is.na(`Total Non-Exempt Hours On Schedule`) |
        `Total Non-Exempt Hours On Schedule` == 0,
      NA,
      round(`Total Number of Kept Appointments` / `Total Non-Exempt Hours On Schedule`, 4)
    )
  ) %>%
  mutate(Four_Week_Group = ((ISoweek_start_date - 1) %/% 4) + 1) %>%
  distinct(Provider, ISoweek_start_date, .keep_all = TRUE)

# --- Complete Provider × 4-Week Grid ---
all_four_week_groups <- sort(unique(Final_Binding$Four_Week_Group))
all_providers <- sort(unique(Final_Binding$Provider))
provider_week_grid <- expand.grid(
  Provider = all_providers,
  Four_Week_Group = all_four_week_groups,
  stringsAsFactors = FALSE
)

Final_Binding_Completed <- provider_week_grid %>%
  left_join(Final_Binding, by = c("Provider", "Four_Week_Group"))

# --- Summary ---
ProviderSummary <- Final_Binding_Completed %>%
  group_by(Provider, Four_Week_Group) %>%
  summarise(
    `Total Kept Appointments` = if (all(is.na(`Total Number of Kept Appointments`))) NA else sum(`Total Number of Kept Appointments`, na.rm = TRUE),
    `Total Non-Exempt Hours On Schedule` = if (all(is.na(`Total Non-Exempt Hours On Schedule`))) NA else sum(`Total Non-Exempt Hours On Schedule`, na.rm = TRUE),
    `Total Exempt Hours on Schedule` = if (all(is.na(`Total Exempt Hours on Schedule`))) NA else sum(`Total Exempt Hours on Schedule`, na.rm = TRUE),
    `Average Productivity` = if (
      all(is.na(`Total Number of Kept Appointments`)) |
      all(is.na(`Total Non-Exempt Hours On Schedule`)) |
      sum(`Total Non-Exempt Hours On Schedule`, na.rm = TRUE) == 0
    ) NA else round(
      sum(`Total Number of Kept Appointments`, na.rm = TRUE) /
        sum(`Total Non-Exempt Hours On Schedule`, na.rm = TRUE), 2
    ),
    .groups = "drop"
  ) %>%
  left_join(Specialty_Data, by = "Provider")

# --- Week Mapping ---
FourWeekMapping <- Final_Binding %>%
  group_by(Four_Week_Group) %>%
  summarise(
    Min_Date = min(ymd(week_start_date), na.rm = TRUE),
    Max_Date = max(ymd(week_end_date), na.rm = TRUE),
    .groups = "drop"
  ) %>%
  mutate(
    Four_Week_Label = paste0(
      "Weeks ", (Four_Week_Group - 1) * 4 + 1, "-", Four_Week_Group * 4,
      "; ", format(Min_Date, "%Y-%m-%d"), " to ", format(Max_Date, "%Y-%m-%d")
    )
  )

ProviderSummaryLabeled <- ProviderSummary %>%
  left_join(FourWeekMapping, by = "Four_Week_Group") %>%
  arrange(Four_Week_Group)

# --- Pivot Table Generator ---
pivot_table <- function(data, value_col) {
  data %>%
    select(Provider, Four_Week_Label, all_of(value_col)) %>%
    filter(!is.na(Four_Week_Label)) %>%
    pivot_wider(
      names_from = Four_Week_Label,
      values_from = all_of(value_col)
    ) %>%
    select(Provider, everything())
}

Productivity_Summary <- ProviderSummaryLabeled %>%
  pivot_table("Average Productivity") %>%
  left_join(Specialty_Data, by = "Provider") %>%
  select(Provider, `Productivity Target?`, everything())

Kept_Appt_Summary <- pivot_table(ProviderSummaryLabeled, "Total Kept Appointments")
Non_Exempt_Summary <- pivot_table(ProviderSummaryLabeled, "Total Non-Exempt Hours On Schedule")
Exempt_Summary <- pivot_table(ProviderSummaryLabeled, "Total Exempt Hours on Schedule")

# --- Save Workbook ---
output_file <- file.path(output_dir, paste0("PROVIDER_4WeekGROUPING_FINAL_", format(Sys.Date(), "%Y-%m-%d"), ".xlsx"))
wb <- createWorkbook()

add_styled_sheet <- function(wb, sheet_name, data) {
  addWorksheet(wb, sheet_name)
  writeDataTable(wb, sheet = sheet_name, x = data)
  freezePane(wb, sheet = sheet_name, firstRow = TRUE)
  setColWidths(wb, sheet = sheet_name, cols = 1:ncol(data), widths = "auto")
}

add_styled_sheet(wb, "Productivity Summary", Productivity_Summary)
add_styled_sheet(wb, "Kept Appointments", Kept_Appt_Summary)
add_styled_sheet(wb, "Non-Exempt Summary", Non_Exempt_Summary)
add_styled_sheet(wb, "Exempt Summary", Exempt_Summary)
add_styled_sheet(wb, "Raw Summary", ProviderSummaryLabeled)

saveWorkbook(wb, file = output_file, overwrite = TRUE)
cat("??? Workbook saved to:", output_file, "\n")
# --- Libraries ---
library(tidyverse)
library(openxlsx)
library(lubridate)
library(readxl)
library(stringr)

# --- Directories ---
input_dir <- "C:/Reports/Provider Prod Data Pulls"
output_dir <- file.path(input_dir, "Filtered/By 4 Week")
if (!dir.exists(output_dir)) dir.create(output_dir, recursive = TRUE)

# --- File Detection ---
files <- list.files(input_dir, pattern = "\\.xlsx$", full.names = TRUE)
prod_file <- files[str_detect(basename(files), regex("Prov Prod Data", ignore_case = TRUE))][1]
appt_file <- files[str_detect(basename(files), regex("Kept_Appointments|NG Kept|Appt|appointments", ignore_case = TRUE)) &
                     !str_detect(basename(files), regex("PROD|Productivity", ignore_case = TRUE))][1]
specialty_file <- files[str_detect(basename(files), regex("Provider_Productivity_Weeks|special", ignore_case = TRUE))][1]

cat("??? Matched files:\n")
cat("   - Productivity:", prod_file, "\n")
cat("   - Appointments:", appt_file, "\n")
cat("   - Specialty:", specialty_file, "\n")

# --- Load Data ---
Specialty_Data <- read_excel(specialty_file) %>%
  select(Provider, `Productivity Target?`)

Prod_Cleaned <- read_excel(prod_file) %>%
  mutate(ISoweek_start_date = isoweek(ymd(week_end_date)))

Appt_Cleaned <- read_excel(appt_file) %>%
  rename(Provider = `Res Name`) %>%
  mutate(ISoweek_start_date = isoweek(ymd(`Appt Dt`))) %>%
  group_by(ISoweek_start_date, Provider) %>%
  summarise(`Total Number of Kept Appointments` = n(), .groups = "drop")

# --- Time Categories ---
Charting_Time_Calc <- Prod_Cleaned %>% filter(category == "Charting Time")
Excemption_Calc <- Prod_Cleaned %>%
  filter(`Prevent Appointments?` == "Y" & !category %in% c("Charting Time", "Administrative Time"))
Non_Excemption_Calc <- Prod_Cleaned %>% filter(`Prevent Appointments?` == "N")
Bound_Charting_Time_Non_Excemption_Time <- bind_rows(Non_Excemption_Calc, Charting_Time_Calc)

Non_Exempt_Time_Calc_By_Provider <- Bound_Charting_Time_Non_Excemption_Time %>%
  group_by(ISoweek_start_date, Provider) %>%
  summarise(
    `Total Non Exemption Time (Mins)` = round(sum(duration), 2),
    `Total Non-Exempt Hours On Schedule` = round(sum(duration) / 60, 2),
    week_start_date = first(week_start_date),
    week_end_date = first(week_end_date),
    .groups = "drop"
  )

Exempt_Time_Calc_By_Provider <- Excemption_Calc %>%
  group_by(ISoweek_start_date, Provider) %>%
  summarise(
    `Total Exemption Time` = sum(duration),
    `Total Exempt Hours on Schedule` = round(`Total Exemption Time` / 60, 2),
    .groups = "drop"
  )

# --- Final Binding ---
Final_Binding <- Non_Exempt_Time_Calc_By_Provider %>%
  full_join(Appt_Cleaned, by = c("Provider", "ISoweek_start_date")) %>%
  full_join(Exempt_Time_Calc_By_Provider, by = c("Provider", "ISoweek_start_date")) %>%
  mutate(
    `Total Productivity` = ifelse(
      is.na(`Total Number of Kept Appointments`) |
        is.na(`Total Non-Exempt Hours On Schedule`) |
        `Total Non-Exempt Hours On Schedule` == 0,
      NA,
      round(`Total Number of Kept Appointments` / `Total Non-Exempt Hours On Schedule`, 4)
    )
  ) %>%
  mutate(Four_Week_Group = ((ISoweek_start_date - 1) %/% 4) + 1) %>%
  distinct(Provider, ISoweek_start_date, .keep_all = TRUE)

# --- Complete Provider × 4-Week Grid ---
all_four_week_groups <- sort(unique(Final_Binding$Four_Week_Group))
all_providers <- sort(unique(Final_Binding$Provider))
provider_week_grid <- expand.grid(
  Provider = all_providers,
  Four_Week_Group = all_four_week_groups,
  stringsAsFactors = FALSE
)

Final_Binding_Completed <- provider_week_grid %>%
  left_join(Final_Binding, by = c("Provider", "Four_Week_Group"))

# --- Summary ---
ProviderSummary <- Final_Binding_Completed %>%
  group_by(Provider, Four_Week_Group) %>%
  summarise(
    `Total Kept Appointments` = if (all(is.na(`Total Number of Kept Appointments`))) NA else sum(`Total Number of Kept Appointments`, na.rm = TRUE),
    `Total Non-Exempt Hours On Schedule` = if (all(is.na(`Total Non-Exempt Hours On Schedule`))) NA else sum(`Total Non-Exempt Hours On Schedule`, na.rm = TRUE),
    `Total Exempt Hours on Schedule` = if (all(is.na(`Total Exempt Hours on Schedule`))) NA else sum(`Total Exempt Hours on Schedule`, na.rm = TRUE),
    `Average Productivity` = if (
      all(is.na(`Total Number of Kept Appointments`)) |
      all(is.na(`Total Non-Exempt Hours On Schedule`)) |
      sum(`Total Non-Exempt Hours On Schedule`, na.rm = TRUE) == 0
    ) NA else round(
      sum(`Total Number of Kept Appointments`, na.rm = TRUE) /
        sum(`Total Non-Exempt Hours On Schedule`, na.rm = TRUE), 2
    ),
    .groups = "drop"
  ) %>%
  left_join(Specialty_Data, by = "Provider")

# --- Week Mapping ---
FourWeekMapping <- Final_Binding %>%
  group_by(Four_Week_Group) %>%
  summarise(
    Min_Date = min(ymd(week_start_date), na.rm = TRUE),
    Max_Date = max(ymd(week_end_date), na.rm = TRUE),
    .groups = "drop"
  ) %>%
  mutate(
    Four_Week_Label = paste0(
      "Weeks ", (Four_Week_Group - 1) * 4 + 1, "-", Four_Week_Group * 4,
      "; ", format(Min_Date, "%Y-%m-%d"), " to ", format(Max_Date, "%Y-%m-%d")
    )
  )

ProviderSummaryLabeled <- ProviderSummary %>%
  left_join(FourWeekMapping, by = "Four_Week_Group") %>%
  arrange(Four_Week_Group)

# --- Pivot Table Generator ---
pivot_table <- function(data, value_col) {
  data %>%
    select(Provider, Four_Week_Label, all_of(value_col)) %>%
    filter(!is.na(Four_Week_Label)) %>%
    pivot_wider(
      names_from = Four_Week_Label,
      values_from = all_of(value_col)
    ) %>%
    select(Provider, everything())
}

Productivity_Summary <- ProviderSummaryLabeled %>%
  pivot_table("Average Productivity") %>%
  left_join(Specialty_Data, by = "Provider") %>%
  select(Provider, `Productivity Target?`, everything())

Kept_Appt_Summary <- pivot_table(ProviderSummaryLabeled, "Total Kept Appointments")
Non_Exempt_Summary <- pivot_table(ProviderSummaryLabeled, "Total Non-Exempt Hours On Schedule")
Exempt_Summary <- pivot_table(ProviderSummaryLabeled, "Total Exempt Hours on Schedule")

# --- Save Workbook ---
output_file <- file.path(output_dir, paste0("PROVIDER_4WeekGROUPING_FINAL_", format(Sys.Date(), "%Y-%m-%d"), ".xlsx"))
wb <- createWorkbook()

add_styled_sheet <- function(wb, sheet_name, data) {
  addWorksheet(wb, sheet_name)
  writeDataTable(wb, sheet = sheet_name, x = data)
  freezePane(wb, sheet = sheet_name, firstRow = TRUE)
  setColWidths(wb, sheet = sheet_name, cols = 1:ncol(data), widths = "auto")
}

add_styled_sheet(wb, "Productivity Summary", Productivity_Summary)
add_styled_sheet(wb, "Kept Appointments", Kept_Appt_Summary)
add_styled_sheet(wb, "Non-Exempt Summary", Non_Exempt_Summary)
add_styled_sheet(wb, "Exempt Summary", Exempt_Summary)
add_styled_sheet(wb, "Raw Summary", ProviderSummaryLabeled)

saveWorkbook(wb, file = output_file, overwrite = TRUE)
cat("??? Workbook saved to:", output_file, "\n")
