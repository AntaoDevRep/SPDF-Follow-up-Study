################################################################################
# SPDF follow-up study: public repository main analysis script
#
# Purpose:
#   Generate Tables 1–4 and Figures 1–4 source-data/validation outputs for:
#   "Asymmetric plantar temperature downshifts are associated with atrial
#   fibrillation and thromboembolic events: an observational analysis of the
#   SmartPreventDiabeticFeet study".
#
# Public repository notes:
#   Raw patient-level clinical data are not included in the repository because
#   they contain sensitive health information and are subject to institutional
#   and ethical restrictions. To rerun the full pipeline locally, place the
#   restricted input files under data_private/follow_up/ and keep data_private/
#   excluded from version control. Public helper functions should be placed in R/.
#
# Recommended repository structure:
#   R/                         cleaned helper functions used below
#   scripts/                   this main analysis script
#   data_private/follow_up/    restricted input data; do not commit
#   outputs/main_analysis/     generated outputs; do not commit by default
#   templates/ppt/             optional PowerPoint template resources
################################################################################

################################################################################
# 0. Setup
################################################################################

## 0.0 Global switches and helper utilities -------------------------------------

export.internal.datasets <- FALSE       # Keep FALSE for public repository use.
export.model.datasets <- FALSE          # Keep FALSE to avoid exporting patient-level model datasets by default.
auto.install.missing.packages <- FALSE  # Prefer explicit dependency installation or renv::restore().

## Suggested .gitignore entries for this repository: data_private/, outputs/, *.pptx, internal_*do_not_upload*.xlsx.

dir_create <- function(path) if (!dir.exists(path)) dir.create(path, recursive = TRUE, showWarnings = TRUE)  # Helper: create a folder only when it does not already exist.

require_pkg <- function(pkg) {  # Helper: require a package and optionally install it if explicitly enabled.
  if (!requireNamespace(pkg, quietly = TRUE)) {  # Run a conditional validation or processing step.
    if (auto.install.missing.packages) install.packages(pkg)  # Run a conditional validation or processing step.
    else stop("Package '", pkg, "' is required. Please install it before running this script.")
  }
  invisible(TRUE)
}

check_objects <- function(objects, context) {  # Helper: stop early when required upstream objects are missing.
  missing <- objects[!sapply(objects, exists)]  # Identify required objects that have not been created yet.
  if (length(missing) > 0) stop("Missing objects before ", context, ": ", paste(missing, collapse = ", "))  # Run a conditional validation or processing step.
  invisible(TRUE)
}

check_vars <- function(data, vars, data_name, context = NULL) {  # Helper: stop early when required columns are missing.
  missing <- setdiff(vars, names(data))  # Identify required variables that are absent from the input data.
  if (length(missing) > 0) stop(data_name, " is missing variables", if (!is.null(context)) paste0(" for ", context) else "", ": ", paste(missing, collapse = ", "))  # Run a conditional validation or processing step.
  invisible(TRUE)
}

p_label <- function(p, prefix = "p", digits = 4) {  # Helper: format p values consistently for figures and tables.
  ifelse(is.na(p), NA_character_, ifelse(p < 0.0001, paste0(prefix, " < 0.0001"), paste0(prefix, " = ", sprintf(paste0("%.", digits, "f"), p))))
}

## 0.1 Define project paths -----------------------------------------------------

workspace.root <- normalizePath(getwd(), winslash = "/", mustWork = FALSE)  # Repository root; run this script from the project root.

data.followup.folder <- file.path(workspace.root, "data_private", "follow_up")  # Restricted input data folder; keep outside public version control.
functional.scripts.folder <- file.path(workspace.root, "R")  # Public repository helper functions folder.

output.folder <- file.path(workspace.root, "outputs", "main_analysis")  # Generated analysis outputs; do not commit unless explicitly approved.
table.output.folder <- file.path(output.folder, "tables")  # Folder for Table 1-4 output workbooks.
figure.output.folder <- file.path(output.folder, "figure_source_data")  # Folder for figure source-data workbooks and validation PPTs.
invisible(lapply(c(output.folder, table.output.folder, figure.output.folder), dir_create))

ppt.template.path <- file.path(workspace.root, "templates", "ppt")  # Optional PowerPoint template folder used by save.fig.to.pptx().
save.as.image.in.ppt <- FALSE  # FALSE keeps editable vector graphics in validation PPTs when supported.

output.ppt.path <- file.path(figure.output.folder, paste0("SPDF_Figure_2_validation_", Sys.Date(), ".pptx"))  # Build a platform-independent file path.
fig3.output.ppt.path <- file.path(figure.output.folder, paste0("SPDF_Figure_3_validation_", Sys.Date(), ".pptx"))  # Build a platform-independent file path.
fig4.output.ppt.path <- file.path(figure.output.folder, paste0("SPDF_Figure_4_validation_", Sys.Date(), ".pptx"))  # Build a platform-independent file path.

## 0.2 Input data paths ---------------------------------------------------------

alarms.data.save.path <- file.path(data.followup.folder, "SPDF_Sensors_Alarms.xlsx")  # Build a platform-independent file path.
study.days.data.save.path <- file.path(data.followup.folder, "SPDF_CRF.xlsx")  # Build a platform-independent file path.
patient.charac.save.path <- file.path(data.followup.folder, "SPDF_ITT patients characteristics.xlsx")  # Build a platform-independent file path.
temp.data.save.path <- file.path(data.followup.folder, "SPDF_temp_sensor_recordings.csv")  # Build a platform-independent file path.
cardiovas.events.save.path <- file.path(data.followup.folder, "Follow-up report atm 2024-05-30.xlsx")  # Build a platform-independent file path.
full.ABI.path <- file.path(data.followup.folder, "SPDF_ABI_Values.xlsx")  # Build a platform-independent file path.

clustering.results.path <- file.path(data.followup.folder, "Follow-up clustering results.xlsx")  # Build a platform-independent file path.
infrared.results.path <- file.path(data.followup.folder, "Follow-up infrared analysis results.xlsx")  # Build a platform-independent file path.

fig3_temp_data.path.candidates <- c( file.path(data.followup.folder, "Follow-up temp for clustering.csv"), file.path(data.followup.folder, "Follow-up temp for clustering.CSV") )  # Define a compact constant vector used below.
fig3_temp_data.path <- fig3_temp_data.path.candidates[file.exists(fig3_temp_data.path.candidates)][1]  # Store an intermediate object used in the following analysis step.
if (length(fig3_temp_data.path) == 0 || is.na(fig3_temp_data.path)) fig3_temp_data.path <- NA_character_  # Store an intermediate object used in the following analysis step.

fig3_temp_features.path <- file.path(data.followup.folder, "Follow-up temp features no labels.csv")  # Build a platform-independent file path.

required.input.paths <- c( alarms.data.save.path, study.days.data.save.path, patient.charac.save.path, temp.data.save.path, cardiovas.events.save.path, full.ABI.path, clustering.results.path, infrared.results.path, fig3_temp_features.path)  # Define a compact constant vector used below.

missing.input.paths <- required.input.paths[!file.exists(required.input.paths)]  # Store an intermediate object used in the following analysis step.
if (length(missing.input.paths) > 0) stop("The following required input files are missing:\n", paste(missing.input.paths, collapse = "\n"))  # Run a conditional validation or processing step.

if (is.na(fig3_temp_data.path) || !file.exists(fig3_temp_data.path)) {  # Run a conditional validation or processing step.
  stop("Figure 3 cleaned temperature data file not found. Checked:\n", paste(fig3_temp_data.path.candidates, collapse = "\n"))
}

message("Input paths initialized successfully.")
message("Core input folder: ", data.followup.folder)
message("Figure 3 cleaned temperature file: ", fig3_temp_data.path)

## 0.3 Load packages and public helper functions -------------------------------

required.packages <- c("openxlsx", "ggplot2", "lubridate", "survival", "patchwork", "ggrepel", "factoextra", "cluster")  # Packages needed by this script and helper functions.
invisible(lapply(required.packages, require_pkg))  # Stop early if required packages are missing.
suppressPackageStartupMessages(invisible(lapply(required.packages, library, character.only = TRUE)))  # Attach required packages quietly.

helper.scripts <- c("functions_intergroup_tests.R", "functions_global.R", "themes.R", "functions_survival.R")  # Cleaned helper scripts expected in R/.
helper.paths <- file.path(functional.scripts.folder, helper.scripts)  # Full paths to public helper scripts.
missing.helpers <- helper.paths[!file.exists(helper.paths)]  # Check that repository helper scripts are present.
if (length(missing.helpers) > 0) stop("Missing public helper scripts. Please place cleaned helper functions in R/:\n", paste(missing.helpers, collapse = "\n"))  # Stop before analysis if helper files are absent.
invisible(lapply(helper.paths, source))  # Load project helper functions after package setup.

required.helper.functions <- c("build.df.summary", "save.fig.to.pptx", "draw.survial.plot")  # Functions supplied by helper scripts and used below.
missing.helper.functions <- required.helper.functions[!sapply(required.helper.functions, exists, mode = "function")]  # Confirm required functions were loaded.
if (length(missing.helper.functions) > 0) stop("Missing required helper functions after sourcing R/: ", paste(missing.helper.functions, collapse = ", "))  # Fail fast for incomplete repository helpers.
required.theme.objects <- c("m.normal.theme", "m.top.legend.theme")  # Theme objects supplied by R/themes.R.
check_objects(required.theme.objects, "theme setup")  # Confirm required ggplot themes are available.
if (exists("check.wd", mode = "function")) check.wd()  # Optional helper retained for local diagnostics when available.

################################################################################
# 1. Load input data
################################################################################

## 1.1 Patient baseline characteristics ----------------------------------------

patients.df <- read.xlsx( file = patient.charac.save.path, sheetIndex = 2, header = TRUE )  # Read an Excel input sheet.

patients.df$Group <- factor( patients.df$Group, levels = c("Control", "Intervention") )  # Standardize factor levels for consistent tables, models, or plots.

message("Patient baseline data loaded:")
print(table(patients.df$Group))

## 1.2 Sensor alarm / PTD classification data ----------------------------------

alarms.data.df <- read.xlsx( file = alarms.data.save.path, sheetIndex = 1, header = TRUE )  # Read an Excel input sheet.

message("Alarm data loaded.")
print(table(alarms.data.df$AlarmTypeEN, useNA = "ifany"))

## 1.3 SPDF study-day data ------------------------------------------------------

study.days.df <- read.xlsx( file = study.days.data.save.path, sheetName = "Study Days", header = TRUE )  # Read an Excel input sheet.

withdraw.patients.n <- nrow(study.days.df[study.days.df$StudyEndType == 7, ])  # Count rows for reporting.

study.days.df <- study.days.df[study.days.df$StudyEndType != 7, ]  # Store an intermediate object used in the following analysis step.

message( "Removed ", withdraw.patients.n, " patients who withdrew before intervention start." )

## Manual correction retained from the original analysis script.
study.days.df[study.days.df$PatientId == "P005", "StudyEndDate"] <- "2018-10-28"  # Update selected rows/columns according to the analysis definition.
study.days.df[study.days.df$PatientId == "P005", "StudyDays"] <-  # Update selected rows/columns according to the analysis definition.
  1 +
  study.days.df[study.days.df$PatientId == "P005", "StudyEndDate"] -
  study.days.df[study.days.df$PatientId == "P005", "StudyStartDate"]

study.days.df$StudyEndType <- factor(  # Standardize factor levels for consistent tables, models, or plots.
  study.days.df$StudyEndType,
  levels = seq(1, 5, 1),
  labels = c("Reached 24-month follow-up", "Primary endpoints (DFU)", "Lost to follow-up", "Died", "Discontinued participation" )
)

message("Study-day data loaded:")
print(table(study.days.df$Group, useNA = "ifany"))
print(table(study.days.df$StudyEndType, useNA = "ifany"))

## 1.4 Temperature sensor recordings -------------------------------------------

temp.data <- read.csv( temp.data.save.path, header = TRUE, sep = "," )  # Read a CSV input file.

message("Temperature sensor recordings loaded.")
message("Number of patients with temperature data: ", length(unique(temp.data$PatientId)))

## 1.5 Cardiovascular and thromboembolic follow-up events -----------------------

CVDs.df <- read.xlsx( file = cardiovas.events.save.path, sheetIndex = 1, header = TRUE )  # Read an Excel input sheet.

names(CVDs.df)[1] <- "PatientId"  # Update selected rows/columns according to the analysis definition.
CVDs.df$PatientId <- sapply(CVDs.df$PatientId, substr, 1, 4)  # Store an intermediate object used in the following analysis step.

selected.cvd.vars <- c("PatientId", "VHF", "pAVK", "pAVK.Interventionsdatum", "LAE..Thrombose..Embolien.", "LAE.TVT.Embolie.Datum", "Apoplexia.cerebri", "Apoplex.Datum" )  # raw follow-up-report columns imported for CVD event recoding.

CVDs.df <- subset(CVDs.df, select = selected.cvd.vars)  # Subset rows or columns for this analysis step.

names(CVDs.df) <- c("PatientId", "AF", "PAD", "PAD.Date", "LAE", "LAE.Date", "CA", "CA.Date" )  # Define a compact constant vector used below.

message("Follow-up cardiovascular/thromboembolic event data loaded.")

################################################################################
# 2. Data cleaning and full follow-up cohort construction
################################################################################

## 2.1 Clean cardiovascular and thromboembolic event variables ------------------

## Coding used in the follow-up report:
## AF: atrial fibrillation
## PAD: peripheral arterial disease
## LAE: pulmonary embolism / thrombosis / embolic event field
## CA: cerebral apoplexy / stroke
## 996: died; 997: withdrew; 998: lost to follow-up

CVDs.df.v2 <- merge( patients.df[, 1:2], CVDs.df[1:270, ], by = "PatientId", all = TRUE )  # Merge source tables using their shared patient identifier.

## Exclude one TVT event from LAE, as in the original analysis.
CVDs.df.v2[CVDs.df.v2$LAE == 2, "LAE"] <- 0  # Update selected rows/columns according to the analysis definition.

## Merge PAD subtypes into a binary PAD outcome.
CVDs.df.v2[
  CVDs.df.v2$PAD == 2 | CVDs.df.v2$PAD == 3,
  "PAD"
] <- 1  # Store an intermediate object used in the following analysis step.

message("Cleaned CVD event variables:")
print(table(CVDs.df.v2$AF, useNA = "ifany"))
print(table(CVDs.df.v2$PAD, useNA = "ifany"))
print(table(CVDs.df.v2$LAE, useNA = "ifany"))
print(table(CVDs.df.v2$CA, useNA = "ifany"))

## 2.2 Build initial full follow-up dataset -------------------------------------

full.df <- merge( patients.df, study.days.df[, -2], by = "PatientId", all = FALSE )  # Merge source tables using their shared patient identifier.

full.df$FollowupEndDate <- as.Date("2024-05-27")  # Convert values to Date format.
full.df$FollowedUpDays <- as.numeric(full.df$FollowupEndDate - full.df$StudyStartDate)  # Convert values to numeric format.

full.df$FollowedUpMonths <- time_length( interval(full.df$StudyStartDate, full.df$FollowupEndDate), "month" )  # Store an intermediate object used in the following analysis step.

full.df$FollowedUpMonths <- round(full.df$FollowedUpMonths, 1)  # Round numeric values for reporting.

full.df <- merge( full.df, CVDs.df.v2[, -2], by = "PatientId", all = FALSE )  # Merge source tables using their shared patient identifier.

message("Initial merged full dataset:")
print(table(full.df$Group, useNA = "ifany"))

## 2.3 Apply exclusion criteria -------------------------------------------------

## Patients with <=30 days of temperature recordings in the intervention group.
pat.measured.days <- data.frame()  # Create a labelled source-data table for validation/export.

for (pat.n in unique(temp.data$PatientId)) {  # Iterate over patients or model outputs.
  pat.df <- subset(temp.data, PatientId == pat.n)  # Subset rows or columns for this analysis step.
  measured.days <- length(unique(pat.df$Timestamp))  # Count records or patients for reporting.

  pat.measured.days <- rbind( pat.measured.days, data.frame( PatientId = pat.n, MeasuredDays = measured.days ) )  # Store an intermediate object used in the following analysis step.
}

patients.excluded.temp <- pat.measured.days[  # Store an intermediate object used in the following analysis step.
  pat.measured.days$MeasuredDays <= 30,
  "PatientId"
]

message( "Patients excluded due to <=30 days of temperature measurements: ", length(patients.excluded.temp) )
print(patients.excluded.temp)

## Patients with <=30 days of SPDF observation.
patients.excluded.obs <- study.days.df[  # Store an intermediate object used in the following analysis step.
  study.days.df$StudyDays <= 30,
  "PatientId"
]

message( "Patients excluded due to <=30 days of observation: ", length(patients.excluded.obs) )
print(patients.excluded.obs)

## Patients who died during the SPDF period.
patients.excluded.died <- study.days.df[  # Store an intermediate object used in the following analysis step.
  study.days.df$StudyEndType == "Died",
  "PatientId"
]

message( "Patients excluded because they died before follow-up ascertainment: ", length(patients.excluded.died) )
print(patients.excluded.died)

## Patients without valid telephone follow-up report.
NANReason <- c(996, 997, 998)  # follow-up report codes indicating non-valid follow-up status.

patients.excluded.followup <- c()  # Define a compact constant vector used below.

for (n.row in seq_len(nrow(CVDs.df.v2))) {  # Iterate over patients or model outputs.
  pat.id <- CVDs.df.v2$PatientId[n.row]  # Store an intermediate object used in the following analysis step.

  if (CVDs.df.v2$AF[n.row] %in% NANReason) {  # Run a conditional validation or processing step.
    patients.excluded.followup <- c(patients.excluded.followup, pat.id)  # Define a compact constant vector used below.
  }
}

message( "Patients excluded due to missing telephone follow-up: ", length(patients.excluded.followup) )
print(patients.excluded.followup)

patients.excluded.all <- unique( c( patients.excluded.temp, patients.excluded.obs, patients.excluded.died, patients.excluded.followup ) )  # Keep unique values for counts or exclusions.

full.df <- subset(full.df, !PatientId %in% patients.excluded.all)  # Subset rows or columns for this analysis step.

message("Final full follow-up cohort after exclusions:")
print(table(full.df$Group, useNA = "ifany"))
message("Total n = ", nrow(full.df))

## 2.4 Format baseline and outcome variables -----------------------------------

full.df$Sex <- factor( full.df$Sex, levels = c("Male", "Female") )  # Standardize factor levels for consistent tables, models, or plots.

full.df[full.df$DMType == "Type 3", "DMType"] <- "Type 2"  # Update selected rows/columns according to the analysis definition.

full.df$DMType <- factor( full.df$DMType, levels = c("Type 1", "Type 2") )  # Standardize factor levels for consistent tables, models, or plots.

full.df$WalkingDistance <- factor( full.df$WalkingDistance, levels = c("<200", "200─500", "500─1000", ">1000") )  # Standardize factor levels for consistent tables, models, or plots.

full.df$QoLClassification <- factor( full.df$QoLClassification, levels = c("Normal (14-25)", "Limited (0-13)") )  # Standardize factor levels for consistent tables, models, or plots.

full.df$ScrNSS <- factor( full.df$ScrNSS, levels = c("Minimal (0-2)", "Mild (3-4)", "Moderate (5-6)", "Severe (7-10)" ) )  # Standardize factor levels for consistent tables, models, or plots.

full.df$ScrNDS <- factor( full.df$ScrNDS, levels = c("Minimal (0-2)", "Mild (3-5)", "Moderate (6-8)", "Severe (9-10)" ) )  # Standardize factor levels for consistent tables, models, or plots.

full.df$VibrationRight <- factor( full.df$VibrationRight, levels = c("Moderate (3-4)", "Severe (0-2)") )  # Standardize factor levels for consistent tables, models, or plots.

full.df$VibrationLeft <- factor( full.df$VibrationLeft, levels = c("Moderate (3-4)", "Severe (0-2)") )  # Standardize factor levels for consistent tables, models, or plots.

full.df$MonofilamentRight <- factor( full.df$MonofilamentRight, levels = c("Absent", "Reduced") )  # Standardize factor levels for consistent tables, models, or plots.

full.df$MonofilamentLeft <- factor( full.df$MonofilamentLeft, levels = c("Absent", "Reduced") )  # Standardize factor levels for consistent tables, models, or plots.

full.df$PulsRight <- factor( full.df$PulsRight, levels = c("strong", "weak", "absent") )  # Standardize factor levels for consistent tables, models, or plots.

full.df$PulsLeft <- factor( full.df$PulsLeft, levels = c("strong", "weak", "absent") )  # Standardize factor levels for consistent tables, models, or plots.

full.df$ABILeft <- factor( full.df$ABILeft, levels = c(">0,9", "≤0,75-0,5", "≤0,9-0,75", "≥1,3") )  # Standardize factor levels for consistent tables, models, or plots.

full.df$ABIRight <- factor( full.df$ABIRight, levels = c(">0,9", "≤0,75-0,5", "≤0,9-0,75", "≥1,3") )  # Standardize factor levels for consistent tables, models, or plots.

full.df$RiskLevel <- factor( full.df$RiskLevel, levels = c(2, 3) )  # Standardize factor levels for consistent tables, models, or plots.

full.df$PreviousDFUClass <- factor( full.df$PreviousDFUClass, levels = c("None", "A0", "A1", "A2", "A3", "Unknown") )  # Standardize factor levels for consistent tables, models, or plots.

full.df$AF <- factor( full.df$AF, levels = c(0, 1), labels = c("No", "Yes") )  # Standardize factor levels for consistent tables, models, or plots.

full.df$PAD <- factor( full.df$PAD, levels = c(0, 1), labels = c("No", "Yes") )  # Standardize factor levels for consistent tables, models, or plots.

full.df$LAE <- factor( full.df$LAE, levels = c(0, 1), labels = c("No", "Yes") )  # Standardize factor levels for consistent tables, models, or plots.

full.df$CA <- factor( full.df$CA, levels = c(0, 1), labels = c("No", "Yes") )  # Standardize factor levels for consistent tables, models, or plots.

## 2.5 Define composite outcome -------------------------------------------------

full.df$AF.CA.LAE.PAD <- "No"  # Set or standardize analysis group labels.

full.df[
  full.df$AF != "No" |
    full.df$CA != "No" |
    full.df$PAD != "No" |
    full.df$LAE != "No",
  "AF.CA.LAE.PAD"
] <- "Yes"  # Store an intermediate object used in the following analysis step.

full.df$AF.CA.LAE.PAD <- factor( full.df$AF.CA.LAE.PAD, levels = c("No", "Yes") )  # Standardize factor levels for consistent tables, models, or plots.

message("Composite AF/thromboembolic outcome:")
print(table(full.df$AF.CA.LAE.PAD, useNA = "ifany"))

################################################################################
# 3. Tables
################################################################################

################################################################################
# 3.1 Table 1: Baseline characteristics by randomized group
################################################################################

## Variables excluded from Table 1 generation.
ignore.vars.table1 <- c("PatientId", "Race", "ShoeSize", "FeetInspectionFrequency", "QoL", "NDS", "NSS", "PAD.Date", "LAE.Date", "CA.Date", "StudyStartDate", "StudyEndDate", "StudyEndType", "StudyEnd", "PreviousDFUClass")  # variables excluded from Table 1 summary generation.

table1.df <- build.df.summary(  # Generate Table 1 baseline summary by randomized group.
  target.df = full.df,
  group.label.name = "Group",
  skip.column.names = ignore.vars.table1,
  effective.digits = 1,
  debug.activated = FALSE,
  only.mean.sd = FALSE,
  group.test.enabled = TRUE
)

message("Table 1 generated.")
print(table1.df[, 1:min(6, ncol(table1.df))])

## Export Table 1 only.
write.xlsx( table1.df, file = file.path(table.output.folder, "Table_1_baseline_characteristics.xlsx"), sheetName = "Table 1", append = FALSE, row.names = FALSE )  # Export analysis output to an Excel workbook.

## Optional internal patient-level export; disabled by default for public use.
if (export.internal.datasets) {  # Run a conditional validation or processing step.
  write.xlsx(full.df, file = file.path(output.folder, "internal_full_followup_dataset_do_not_upload.xlsx"), sheetName = "Full cohort", append = FALSE, row.names = FALSE)  # Export analysis output to an Excel workbook.
}

################################################################################
# 3.1.1 Follow-up duration checks for manuscript text
################################################################################

followup.days.q <- quantile(full.df$FollowedUpDays, probs = c(0.25, 0.75), na.rm = TRUE)  # Calculate interquartile range values.
followup.months.q <- quantile(full.df$FollowedUpMonths, probs = c(0.25, 0.75), na.rm = TRUE)  # Calculate interquartile range values.

followup.summary <- data.frame(  # Create a labelled source-data table for validation/export.
  variable = c("FollowedUpDays_median", "FollowedUpDays_Q1_Q3", "FollowedUpDays_range", "FollowedUpMonths_median", "FollowedUpMonths_Q1_Q3", "FollowedUpMonths_range" ),
  value = c(
    median(full.df$FollowedUpDays, na.rm = TRUE),
    paste(followup.days.q, collapse = " - "),
    paste(range(full.df$FollowedUpDays, na.rm = TRUE), collapse = " - "),
    median(full.df$FollowedUpMonths, na.rm = TRUE),
    paste(followup.months.q, collapse = " - "),
    paste(range(full.df$FollowedUpMonths, na.rm = TRUE), collapse = " - ")
  )
)

write.xlsx( followup.summary, file = file.path(table.output.folder, "Followup_duration_summary.xlsx"), sheetName = "Follow-up summary", append = FALSE, row.names = FALSE )  # Export analysis output to an Excel workbook.

message("Analysis completed up to Table 1.")

################################################################################
# 3.2 Table 2: AF and thromboembolic events by PTD, cluster and PIRI subgroups
################################################################################

## This section builds the intervention-group analysis cohort and generates
## Table 2 summaries by:
##   1) PTD / post-hoc temperature group
##   2) temperature variability cluster
##   3) PIRI symmetry/asymmetry group
##
## Required objects from previous sections:
##   full.df
##   alarms.data.df
##   clustering.results.path
##   infrared.results.path
##   output.folder
##   table.output.folder

################################################################################
# 3.2.1 Check required objects and input files
################################################################################

required.objects.table2 <- c("full.df", "alarms.data.df", "clustering.results.path", "infrared.results.path", "output.folder", "table.output.folder" )  # Define a compact constant vector used below.

missing.objects.table2 <- required.objects.table2[  # Store an intermediate object used in the following analysis step.
  !sapply(required.objects.table2, exists)
]

if (length(missing.objects.table2) > 0) {  # Run a conditional validation or processing step.
  stop(
    "The following required objects are missing before Table 2: ",
    paste(missing.objects.table2, collapse = ", "),
    ". Please run the setup, data loading, cohort construction and Table 1 sections first."
  )
}

if (!file.exists(clustering.results.path)) {  # Run a conditional validation or processing step.
  stop( "Clustering results file not found. Please check clustering.results.path: ", clustering.results.path )
}

if (!file.exists(infrared.results.path)) {  # Run a conditional validation or processing step.
  stop( "Infrared/PIRI results file not found. Please check infrared.results.path: ", infrared.results.path )
}

################################################################################
# 3.2.2 Build intervention-group analysis cohort
################################################################################

int.full.df <- subset(full.df, Group == "Intervention")  # Subset rows or columns for this analysis step.

message("Intervention-group analysis cohort:")
message("n = ", nrow(int.full.df))
print(table(int.full.df$Group, useNA = "ifany"))

################################################################################
# 3.2.3 Add PTD / post-hoc temperature group
################################################################################

## In the original analysis, patients with AlarmTypeEN == "Ischemia" were
## classified as having lowered plantar temperature / PTDs.

if (!"AlarmTypeEN" %in% names(alarms.data.df)) {  # Run a conditional validation or processing step.
  stop("Column 'AlarmTypeEN' not found in alarms.data.df.")
}

if (!"PatientId" %in% names(alarms.data.df)) {  # Run a conditional validation or processing step.
  stop("Column 'PatientId' not found in alarms.data.df.")
}

message("Alarm types:")
print(table(alarms.data.df$AlarmTypeEN, useNA = "ifany"))

ischemie.pats <- alarms.data.df[  # Store an intermediate object used in the following analysis step.
  alarms.data.df$AlarmTypeEN == "Ischemia",
  "PatientId"
]

int.full.df$PostHocGroup <- "Normal Temp."  # Set or standardize analysis group labels.

int.full.df[
  int.full.df$PatientId %in% ischemie.pats,
  "PostHocGroup"
] <- "Lowered Temp."  # Store an intermediate object used in the following analysis step.

int.full.df$PostHocGroup <- factor( int.full.df$PostHocGroup, levels = c("Normal Temp.", "Lowered Temp.") )  # Standardize factor levels for consistent tables, models, or plots.

message("PTD / post-hoc temperature groups:")
print(table(int.full.df$PostHocGroup, useNA = "ifany"))

################################################################################
# 3.2.4 Add temperature variability cluster group
################################################################################

clustering.df <- read.xlsx( file = clustering.results.path, sheetIndex = 1 )  # Read an Excel input sheet.

required.cluster.vars <- c("PatientId", "Cluster")  # Define a compact constant vector used below.
missing.cluster.vars <- setdiff(required.cluster.vars, names(clustering.df))  # Store an intermediate object used in the following analysis step.

if (length(missing.cluster.vars) > 0) {  # Run a conditional validation or processing step.
  stop( "The clustering results file is missing required columns: ", paste(missing.cluster.vars, collapse = ", ") )
}

clustering.subdf <- clustering.df[, required.cluster.vars]  # Store an intermediate object used in the following analysis step.

int.full.df <- merge( int.full.df, clustering.subdf, by = "PatientId", all.x = TRUE )  # Merge source tables using their shared patient identifier.

## Standardize cluster labels.
if (is.numeric(int.full.df$Cluster) || all(na.omit(int.full.df$Cluster) %in% c(1, 2))) {  # Run a conditional validation or processing step.
  int.full.df$Cluster <- factor( int.full.df$Cluster, levels = c(1, 2), labels = c("C1 (n=71)", "C2 (n=47)") )  # Standardize factor levels for consistent tables, models, or plots.
} else {
  int.full.df$Cluster <- as.character(int.full.df$Cluster)  # Convert factor-like values to character strings before relabelling.

  int.full.df$Cluster[int.full.df$Cluster %in% c("1", "Cluster 1", "C1")] <- "C1 (n=71)"  # Update selected rows/columns according to the analysis definition.
  int.full.df$Cluster[int.full.df$Cluster %in% c("2", "Cluster 2", "C2")] <- "C2 (n=47)"  # Update selected rows/columns according to the analysis definition.

  int.full.df$Cluster <- factor( int.full.df$Cluster, levels = c("C1 (n=71)", "C2 (n=47)") )  # Standardize factor levels for consistent tables, models, or plots.
}

message("Temperature variability clusters:")
print(table(int.full.df$Cluster, useNA = "ifany"))

################################################################################
# 3.2.5 Add PIRI / infrared image group
################################################################################

infrared.df <- read.xlsx( file = infrared.results.path, sheetIndex = 1 )  # Read an Excel input sheet.

required.infrared.vars <- c("PatientId", "InfraredImage")  # Define a compact constant vector used below.
missing.infrared.vars <- setdiff(required.infrared.vars, names(infrared.df))  # Store an intermediate object used in the following analysis step.

if (length(missing.infrared.vars) > 0) {  # Run a conditional validation or processing step.
  stop( "The infrared/PIRI results file is missing required columns: ", paste(missing.infrared.vars, collapse = ", ") )
}

infrared.subdf <- infrared.df[, required.infrared.vars]  # Store an intermediate object used in the following analysis step.

int.full.df <- merge( int.full.df, infrared.subdf, by = "PatientId", all.x = TRUE )  # Merge source tables using their shared patient identifier.

## Standardize PIRI labels.
int.full.df$PIRI.Group <- as.character(int.full.df$InfraredImage)  # Convert factor-like values to character strings before relabelling.

int.full.df$PIRI.Group[
  int.full.df$PIRI.Group %in% c("symm.", "symm", "Symmetric", "symmetric")
] <- "Symmetric"  # Store an intermediate object used in the following analysis step.

int.full.df$PIRI.Group[
  int.full.df$PIRI.Group %in% c("asym.", "asym", "Asymmetric", "asymmetric")
] <- "Asymmetric"  # Store an intermediate object used in the following analysis step.

int.full.df$PIRI.Group <- factor( int.full.df$PIRI.Group, levels = c("Symmetric", "Asymmetric") )  # Standardize factor levels for consistent tables, models, or plots.

message("PIRI / infrared image groups:")
print(table(int.full.df$PIRI.Group, useNA = "ifany"))

################################################################################
# 3.2.6 Sanity checks for subgroup sizes
################################################################################

## These checks are intended to catch accidental mismatches in input files.
## If expected values change after data updates, revise the expected counts below.

expected.posthoc <- c("Normal Temp." = 52, "Lowered Temp." = 66)  # Define a compact constant vector used below.
expected.cluster <- c("C1 (n=71)" = 71, "C2 (n=47)" = 47)  # Define a compact constant vector used below.
expected.piri <- c("Symmetric" = 81, "Asymmetric" = 37)  # Define a compact constant vector used below.

observed.posthoc <- table(int.full.df$PostHocGroup)  # Tabulate counts for quality control or plotting.
observed.cluster <- table(int.full.df$Cluster)  # Tabulate counts for quality control or plotting.
observed.piri <- table(int.full.df$PIRI.Group)  # Tabulate counts for quality control or plotting.

message("Sanity check: expected vs observed subgroup sizes")
print(observed.posthoc)
print(observed.cluster)
print(observed.piri)

if (!all(observed.posthoc[names(expected.posthoc)] == expected.posthoc)) {  # Run a conditional validation or processing step.
  warning("PTD/PostHocGroup counts differ from expected manuscript counts.")
}

if (!all(observed.cluster[names(expected.cluster)] == expected.cluster)) {  # Run a conditional validation or processing step.
  warning("Cluster counts differ from expected manuscript counts.")
}

if (!all(observed.piri[names(expected.piri)] == expected.piri)) {  # Run a conditional validation or processing step.
  warning("PIRI counts differ from expected manuscript counts.")
}

################################################################################
# 3.2.7 Generate concise Table 2 subgroup summaries
################################################################################

## Table 2 focuses on follow-up duration and AF/thromboembolic outcomes
## in the intervention-group analysis cohort.

table2.select.vars.posthoc <- c("PostHocGroup", "FollowedUpMonths", "AF", "LAE", "CA", "PAD", "AF.CA.LAE.PAD" )  # Table 2 variables summarized by PTD/post-hoc temperature group.

table2.select.vars.cluster <- c("Cluster", "FollowedUpMonths", "AF", "LAE", "CA", "PAD", "AF.CA.LAE.PAD" )  # Table 2 variables summarized by temperature-variability cluster.

table2.select.vars.piri <- c("PIRI.Group", "FollowedUpMonths", "AF", "LAE", "CA", "PAD", "AF.CA.LAE.PAD" )  # Table 2 variables summarized by PIRI group.

## Ensure all required variables exist before generating summaries.
required.table2.vars <- unique(c( table2.select.vars.posthoc, table2.select.vars.cluster, table2.select.vars.piri ))  # Keep unique values for counts or exclusions.

missing.table2.vars <- setdiff(required.table2.vars, names(int.full.df))  # Store an intermediate object used in the following analysis step.

if (length(missing.table2.vars) > 0) {  # Run a conditional validation or processing step.
  stop( "The following variables required for Table 2 are missing in int.full.df: ", paste(missing.table2.vars, collapse = ", ") )
}

## PTD / post-hoc temperature group.
table2_posthoc <- build.df.summary(  # Store an intermediate object used in the following analysis step.
  target.df = subset(int.full.df, select = table2.select.vars.posthoc),
  group.label.name = "PostHocGroup",
  skip.column.names = c(),
  effective.digits = 1,
  debug.activated = FALSE,
  only.mean.sd = FALSE,
  group.test.enabled = TRUE
)

table2_posthoc <- table2_posthoc[, c(1:3, 6)]  # Store an intermediate object used in the following analysis step.

## Temperature variability cluster group.
table2_cluster <- build.df.summary(  # Store an intermediate object used in the following analysis step.
  target.df = subset(int.full.df, select = table2.select.vars.cluster),
  group.label.name = "Cluster",
  skip.column.names = c(),
  effective.digits = 1,
  debug.activated = FALSE,
  only.mean.sd = FALSE,
  group.test.enabled = TRUE
)

table2_cluster <- table2_cluster[, c(1:3, 6)]  # Store an intermediate object used in the following analysis step.

## PIRI symmetry/asymmetry group.
table2_piri <- build.df.summary(  # Store an intermediate object used in the following analysis step.
  target.df = subset(int.full.df, select = table2.select.vars.piri),
  group.label.name = "PIRI.Group",
  skip.column.names = c(),
  effective.digits = 1,
  debug.activated = FALSE,
  only.mean.sd = FALSE,
  group.test.enabled = TRUE
)

table2_piri <- table2_piri[, c(1:3, 6)]  # Store an intermediate object used in the following analysis step.

################################################################################
# 3.2.8 Check consistency before integrating Table 2
################################################################################

if (!identical(as.character(table2_posthoc[, 1]), as.character(table2_cluster[, 1]))) {  # Run a conditional validation or processing step.
  warning("Variable order differs between PTD and cluster Table 2 summaries.")
  print(data.frame( PTD = table2_posthoc[, 1], Cluster = table2_cluster[, 1] ))
}

if (!identical(as.character(table2_posthoc[, 1]), as.character(table2_piri[, 1]))) {  # Run a conditional validation or processing step.
  warning("Variable order differs between PTD and PIRI Table 2 summaries.")
  print(data.frame( PTD = table2_posthoc[, 1], PIRI = table2_piri[, 1] ))
}

################################################################################
# 3.2.9 Integrate Table 2
################################################################################

table2_integrated <- cbind( table2_posthoc, table2_cluster[, c(2:4)], table2_piri[, c(2:4)] )  # Store an intermediate object used in the following analysis step.

colnames(table2_integrated) <- c("Variable", "Normal Temp.", "Lowered Temp.", "P value: PTD group", "Cluster 1", "Cluster 2", "P value: cluster", "PIRI symmetric", "PIRI asymmetric", "P value: PIRI")  # Define a compact constant vector used below.

message("Integrated Table 2:")
print(table2_integrated)

################################################################################
# 3.2.10 Optional full subgroup reports
################################################################################

## Keep this off by default to avoid unnecessarily large outputs.
## Set TRUE only when you need to inspect all variables by subgroup.

export.full.table2.reports <- FALSE  # Store an intermediate object used in the following analysis step.

if (export.full.table2.reports) {  # Run a conditional validation or processing step.

  ignore.vars.posthoc <- c("PatientId", "Group", "Cluster", "InfraredImage", "PIRI.Group", "Race", "ShoeSize", "FeetInspectionFrequency", "QoL", "NDS", "NSS", "PAD.Date", "LAE.Date", "CA.Date", "StudyStartDate", "StudyEndDate", "StudyEndType", "StudyEnd")  # Define a compact constant vector used below.

  ignore.vars.cluster <- c("PatientId", "Group", "PostHocGroup", "InfraredImage", "PIRI.Group", "Race", "ShoeSize", "FeetInspectionFrequency", "QoL", "NDS", "NSS", "PAD.Date", "LAE.Date", "CA.Date", "StudyStartDate", "StudyEndDate", "StudyEndType", "StudyEnd")

  ignore.vars.piri <- c("PatientId", "Group", "PostHocGroup", "Cluster", "InfraredImage", "Race", "ShoeSize", "FeetInspectionFrequency", "QoL", "NDS", "NSS", "PAD.Date", "LAE.Date", "CA.Date", "StudyStartDate", "StudyEndDate", "StudyEndType", "StudyEnd")  # Define a compact constant vector used below.

  table2_posthoc_full <- build.df.summary(  # Store an intermediate object used in the following analysis step.
    target.df = int.full.df,
    group.label.name = "PostHocGroup",
    skip.column.names = ignore.vars.posthoc,
    effective.digits = 1,
    debug.activated = FALSE,
    only.mean.sd = FALSE,
    group.test.enabled = TRUE
  )

  table2_cluster_full <- build.df.summary(  # Store an intermediate object used in the following analysis step.
    target.df = int.full.df,
    group.label.name = "Cluster",
    skip.column.names = ignore.vars.cluster,
    effective.digits = 1,
    debug.activated = FALSE,
    only.mean.sd = FALSE,
    group.test.enabled = TRUE
  )

  table2_piri_full <- build.df.summary(  # Store an intermediate object used in the following analysis step.
    target.df = int.full.df,
    group.label.name = "PIRI.Group",
    skip.column.names = ignore.vars.piri,
    effective.digits = 1,
    debug.activated = FALSE,
    only.mean.sd = FALSE,
    group.test.enabled = TRUE
  )
}

################################################################################
# 3.2.11 Export Table 2 outputs
################################################################################

table2.output.path <- file.path( table.output.folder, "Table_2_subgroup_events.xlsx" )  # Build a platform-independent file path.

if (export.full.table2.reports) {  # Run a conditional validation or processing step.
  table2.export.list <- list(  # Store an intermediate object used in the following analysis step.
    "Table 2 integrated" = table2_integrated,
    "PTD group" = table2_posthoc,
    "Cluster" = table2_cluster,
    "PIRI" = table2_piri,
    "PTD full" = table2_posthoc_full,
    "Cluster full" = table2_cluster_full,
    "PIRI full" = table2_piri_full
  )
} else {
  table2.export.list <- list( "Table 2 integrated" = table2_integrated, "PTD group" = table2_posthoc, "Cluster" = table2_cluster, "PIRI" = table2_piri )  # Store an intermediate object used in the following analysis step.
}

openxlsx::write.xlsx( x = table2.export.list, file = table2.output.path, overwrite = TRUE, rowNames = FALSE )  # Export analysis output to an Excel workbook.

message("Table 2 exported to:")
message(table2.output.path)

################################################################################
# 3.2.12 Optional internal patient-level intervention dataset
################################################################################

## Disabled by default for public repository use.
if (export.internal.datasets) {  # Run a conditional validation or processing step.
  openxlsx::write.xlsx(int.full.df, file = file.path(output.folder, "internal_intervention_dataset_do_not_upload.xlsx"), sheetName = "Intervention cohort", overwrite = TRUE, rowNames = FALSE)  # Export analysis output to an Excel workbook.
}

message("Analysis completed up to Table 2.")

################################################################################
# 3.3 Table 3: Logistic regression analysis for composite AF/thromboembolic outcome
################################################################################

## This section generates Table 3:
## logistic regression models for the composite outcome of AF, CA, LAE, or PAD
## in the intervention-group analysis cohort.
##
## Required objects from previous sections:
##   int.full.df
##   full.ABI.path
##   clustering.results.path
##   table.output.folder

################################################################################
# 3.3.1 Check required objects and input files
################################################################################

required.objects.table3 <- c("int.full.df", "full.ABI.path", "clustering.results.path", "table.output.folder" )  # Define a compact constant vector used below.

missing.objects.table3 <- required.objects.table3[  # Store an intermediate object used in the following analysis step.
  !sapply(required.objects.table3, exists)
]

if (length(missing.objects.table3) > 0) {  # Run a conditional validation or processing step.
  stop(
    "The following required objects are missing before Table 3: ",
    paste(missing.objects.table3, collapse = ", "),
    ". Please run the setup, cohort construction, Table 1 and Table 2 sections first."
  )
}

if (!file.exists(full.ABI.path)) {  # Run a conditional validation or processing step.
  stop("ABI data file not found. Please check full.ABI.path: ", full.ABI.path)
}

if (!file.exists(clustering.results.path)) {  # Run a conditional validation or processing step.
  stop("Clustering results file not found. Please check clustering.results.path: ", clustering.results.path)
}

################################################################################
# 3.3.2 Load ABI data and prepare model dataset
################################################################################

full.ABI.df <- read.xlsx( file = full.ABI.path, sheetIndex = 1 )  # Read an Excel input sheet.

required.abi.vars <- c("PatientId", "Group", "ScrABIRightValue", "ScrABILeftValue" )  # minimum ABI columns required for regression models.

missing.abi.vars <- setdiff(required.abi.vars, names(full.ABI.df))  # Store an intermediate object used in the following analysis step.

if (length(missing.abi.vars) > 0) {  # Run a conditional validation or processing step.
  stop( "The ABI file is missing required columns: ", paste(missing.abi.vars, collapse = ", ") )
}

## Keep only intervention-group ABI data.
abi.model.df <- full.ABI.df[  # Store an intermediate object used in the following analysis step.
  full.ABI.df$Group == "Intervention",
  c("PatientId", "ScrABIRightValue", "ScrABILeftValue")
]

## Use Cluster already merged into int.full.df during Table 2.
## This avoids re-merging cluster data and keeps Table 2 and Table 3 consistent.
required.int.vars.table3 <- c("PatientId", "Sex", "Age", "BMI", "DMDuration", "Cluster", "AF.CA.LAE.PAD" )  # intervention-cohort variables required for logistic regression.

missing.int.vars.table3 <- setdiff(required.int.vars.table3, names(int.full.df))  # Store an intermediate object used in the following analysis step.

if (length(missing.int.vars.table3) > 0) {  # Run a conditional validation or processing step.
  stop( "int.full.df is missing required columns for Table 3: ", paste(missing.int.vars.table3, collapse = ", ") )
}

log.df <- merge( int.full.df[, required.int.vars.table3], abi.model.df, by = "PatientId", all.x = FALSE, all.y = FALSE )  # Merge source tables using their shared patient identifier.

## Format variables for logistic regression.
log.df$Sex <- factor( log.df$Sex, levels = c("Male", "Female") )  # Standardize factor levels for consistent tables, models, or plots.

log.df$Cluster <- factor( log.df$Cluster, levels = c("C1 (n=71)", "C2 (n=47)") )  # Standardize factor levels for consistent tables, models, or plots.

log.df$AF.CA.LAE.PAD <- factor( log.df$AF.CA.LAE.PAD, levels = c("No", "Yes") )  # Standardize factor levels for consistent tables, models, or plots.

## Numeric coercion for covariates if needed.
log.df$Age <- as.numeric(log.df$Age)  # Convert values to numeric format.
log.df$BMI <- as.numeric(log.df$BMI)  # Convert values to numeric format.
log.df$DMDuration <- as.numeric(log.df$DMDuration)  # Convert values to numeric format.
log.df$ScrABIRightValue <- as.numeric(log.df$ScrABIRightValue)  # Convert values to numeric format.
log.df$ScrABILeftValue <- as.numeric(log.df$ScrABILeftValue)  # Convert values to numeric format.

message("Table 3 logistic regression dataset:")
message("n before complete-case filtering = ", nrow(log.df))
print(table(log.df$Cluster, useNA = "ifany"))
print(table(log.df$AF.CA.LAE.PAD, useNA = "ifany"))

################################################################################
# 3.3.3 Fit logistic regression models
################################################################################

## Model 1: unadjusted
log.model1 <- glm( AF.CA.LAE.PAD ~ Cluster, family = binomial(link = "logit"), data = log.df )  # Fit a logistic regression model.

## Model 2: adjusted for demographic and clinical covariates
log.model2 <- glm(  # Fit a logistic regression model.
  AF.CA.LAE.PAD ~ Age + Sex + BMI + DMDuration + Cluster,
  family = binomial(link = "logit"),
  data = log.df
)

## Model 3: ABI model without cluster
log.model3 <- glm(  # Fit a logistic regression model.
  AF.CA.LAE.PAD ~ Age + Sex + BMI + DMDuration + ScrABIRightValue + ScrABILeftValue,
  family = binomial(link = "logit"),
  data = log.df
)

## Model 4: fully adjusted model
log.model4 <- glm(  # Fit a logistic regression model.
  AF.CA.LAE.PAD ~ Age + Sex + BMI + DMDuration + ScrABIRightValue + ScrABILeftValue + Cluster,
  family = binomial(link = "logit"),
  data = log.df
)

################################################################################
# 3.3.4 Helper function to extract OR, 95% CI and p values
################################################################################

extract_logistic_results <- function(model, model_name) {  # Store an intermediate object used in the following analysis step.

  model_summary <- summary(model)  # Store an intermediate object used in the following analysis step.
  coef_table <- as.data.frame(model_summary$coefficients)  # Convert summary output to a data frame for export.

  coef_table$term <- rownames(coef_table)  # Store an intermediate object used in the following analysis step.
  rownames(coef_table) <- NULL  # Store an intermediate object used in the following analysis step.

  names(coef_table) <- c("estimate_log_odds", "standard_error", "z_value", "p_value", "term" )  # Define a compact constant vector used below.

  ## Profile likelihood confidence intervals are used by confint().
  ## If confint() fails, fall back to Wald confidence intervals.
  ci <- tryCatch(  # Store an intermediate object used in the following analysis step.
    {
      suppressMessages(confint(model))
    },
    error = function(e) {
      est <- coef(model)  # Extract model coefficient estimates.
      se <- sqrt(diag(vcov(model)))  # Store an intermediate object used in the following analysis step.
      cbind(
        est - 1.96 * se,
        est + 1.96 * se
      )
    }
  )

  ci <- as.data.frame(ci)  # Convert summary output to a data frame for export.
  ci$term <- rownames(ci)  # Store an intermediate object used in the following analysis step.
  rownames(ci) <- NULL  # Store an intermediate object used in the following analysis step.
  names(ci)[1:2] <- c("ci_lower_log_odds", "ci_upper_log_odds")  # Define a compact constant vector used below.

  out <- merge( coef_table, ci, by = "term", all.x = TRUE, sort = FALSE )  # Merge source tables using their shared patient identifier.

  out$model <- model_name  # Store an intermediate object used in the following analysis step.
  out$n_observations <- stats::nobs(model)  # Store an intermediate object used in the following analysis step.
  out$OR <- exp(out$estimate_log_odds)  # Store an intermediate object used in the following analysis step.
  out$CI_lower <- exp(out$ci_lower_log_odds)  # Store an intermediate object used in the following analysis step.
  out$CI_upper <- exp(out$ci_upper_log_odds)  # Store an intermediate object used in the following analysis step.

  out$p_value_formatted <- ifelse( is.na(out$p_value), NA, ifelse(out$p_value < 0.0001, "<0.0001", sprintf("%.4f", out$p_value)) )  # Store an intermediate object used in the following analysis step.

  out$OR_95CI <- paste0( sprintf("%.2f", out$OR), " (", sprintf("%.2f", out$CI_lower), "–", sprintf("%.2f", out$CI_upper), ")" )  # Create a readable label for plotting or reporting.

  out <- out[  # Store an intermediate object used in the following analysis step.
    ,
    c("model", "n_observations", "term", "OR", "CI_lower", "CI_upper", "OR_95CI", "p_value", "p_value_formatted", "estimate_log_odds", "standard_error", "z_value" )
  ]

  out
}

################################################################################
# 3.3.5 Extract and assemble Table 3
################################################################################

table3_model1 <- extract_logistic_results(log.model1, "Model 1")  # Store an intermediate object used in the following analysis step.
table3_model2 <- extract_logistic_results(log.model2, "Model 2")  # Store an intermediate object used in the following analysis step.
table3_model3 <- extract_logistic_results(log.model3, "Model 3")  # Store an intermediate object used in the following analysis step.
table3_model4 <- extract_logistic_results(log.model4, "Model 4")  # Store an intermediate object used in the following analysis step.

table3_all_terms <- rbind( table3_model1, table3_model2, table3_model3, table3_model4 )  # Store an intermediate object used in the following analysis step.

## Main manuscript Table 3 can focus on the cluster term and ABI terms,
## while the full model output is retained in a separate sheet.
table3_main <- table3_all_terms[  # Store an intermediate object used in the following analysis step.
  table3_all_terms$term %in% c("ClusterC2 (n=47)", "ScrABIRightValue", "ScrABILeftValue" ),
]

message("Table 3 main results:")
print(table3_main)

################################################################################
# 3.3.6 Add model descriptions
################################################################################

table3_model_descriptions <- data.frame(  # Create a labelled source-data table for validation/export.
  model = c("Model 1", "Model 2", "Model 3", "Model 4"),
  outcome = "Composite AF/thromboembolic outcome",
  model_type = "Logistic regression",
  formula = c(
    "AF.CA.LAE.PAD ~ Cluster",
    "AF.CA.LAE.PAD ~ Age + Sex + BMI + DMDuration + Cluster",
    "AF.CA.LAE.PAD ~ Age + Sex + BMI + DMDuration + ScrABIRightValue + ScrABILeftValue",
    "AF.CA.LAE.PAD ~ Age + Sex + BMI + DMDuration + ScrABIRightValue + ScrABILeftValue + Cluster"
  ),
  adjustment = c(
    "Unadjusted",
    "Adjusted for age, sex, BMI and diabetes duration",
    "Adjusted for age, sex, BMI, diabetes duration and bilateral ABI",
    "Adjusted for age, sex, BMI, diabetes duration, bilateral ABI and temperature variability cluster"
  ),
  stringsAsFactors = FALSE
)

################################################################################
# 3.3.7 Export Table 3 outputs
################################################################################

table3.output.path <- file.path( table.output.folder, "Table_3_logistic_regression.xlsx" )  # Build a platform-independent file path.

table3.export.list <- list("Table 3 main" = table3_main, "All model terms" = table3_all_terms, "Model descriptions" = table3_model_descriptions)  # Public-safe Table 3 export list.
if (export.model.datasets) table3.export.list[["Logistic dataset"]] <- log.df  # Optional restricted patient-level model dataset; do not commit publicly.
openxlsx::write.xlsx(x = table3.export.list, file = table3.output.path, overwrite = TRUE, rowNames = FALSE)  # Export Table 3 workbook.

message("Table 3 exported to:")
message(table3.output.path)

message("Analysis completed up to Table 3.")

################################################################################
# 3.4 Table 4: Cox proportional hazards regression analysis
################################################################################

## This section generates Table 4:
## Cox proportional hazards models for thromboembolic time-to-event outcomes
## in the intervention-group analysis cohort.
##
## Outcome:
##   First occurrence of PAD, LAE/PE, or CA/stroke.
##
## Required objects from previous sections:
##   full.df
##   int.full.df
##   full.ABI.df or full.ABI.path
##   table.output.folder

################################################################################
# 3.4.1 Check required objects and input files
################################################################################

required.objects.table4 <- c("full.df", "int.full.df", "table.output.folder" )  # Define a compact constant vector used below.

missing.objects.table4 <- required.objects.table4[  # Store an intermediate object used in the following analysis step.
  !sapply(required.objects.table4, exists)
]

if (length(missing.objects.table4) > 0) {  # Run a conditional validation or processing step.
  stop(
    "The following required objects are missing before Table 4: ",
    paste(missing.objects.table4, collapse = ", "),
    ". Please run the setup, cohort construction, Table 1, Table 2 and Table 3 sections first."
  )
}

if (!exists("full.ABI.df")) {  # Run a conditional validation or processing step.
  if (!exists("full.ABI.path")) {  # Run a conditional validation or processing step.
    stop("Neither full.ABI.df nor full.ABI.path exists. Please define full.ABI.path before Table 4.")
  }

  if (!file.exists(full.ABI.path)) {  # Run a conditional validation or processing step.
    stop("ABI data file not found. Please check full.ABI.path: ", full.ABI.path)
  }

  full.ABI.df <- read.xlsx( file = full.ABI.path, sheetIndex = 1 )  # Read an Excel input sheet.
}

required.abi.vars <- c("PatientId", "Group", "ScrABIRightValue", "ScrABILeftValue" )  # minimum ABI columns required for regression models.

missing.abi.vars <- setdiff(required.abi.vars, names(full.ABI.df))  # Store an intermediate object used in the following analysis step.

if (length(missing.abi.vars) > 0) {  # Run a conditional validation or processing step.
  stop( "The ABI data are missing required columns: ", paste(missing.abi.vars, collapse = ", ") )
}

################################################################################
# 3.4.2 Prepare time-to-event variables
################################################################################

## PAD.LAE.CA.Event:
##   TRUE if the patient had PAD, LAE/PE, or CA/stroke with an available event date.
##
## DaysToEvent:
##   For patients with a thromboembolic event, time from StudyStartDate to earliest
##   PAD/LAE/CA event date.
##   For censored patients, time from StudyStartDate to FollowupEndDate.

required.event.date.vars <- c("PatientId", "StudyStartDate", "FollowupEndDate", "FollowedUpDays", "PAD.Date", "LAE.Date", "CA.Date" )  # date variables needed to derive the time-to-event outcome.

missing.event.date.vars <- setdiff(required.event.date.vars, names(full.df))  # Store an intermediate object used in the following analysis step.

if (length(missing.event.date.vars) > 0) {  # Run a conditional validation or processing step.
  stop( "full.df is missing required variables for time-to-event analysis: ", paste(missing.event.date.vars, collapse = ", ") )
}

full.df$DaysToEvent <- full.df$FollowedUpDays  # Store an intermediate object used in the following analysis step.
full.df$DateToEvent <- full.df$FollowupEndDate  # Store an intermediate object used in the following analysis step.
full.df$PAD.LAE.CA.Event <- FALSE  # Store an intermediate object used in the following analysis step.

for (n.pat in seq_len(nrow(full.df))) {  # Iterate over patients or model outputs.

  event.dates <- as.Date(character())  # Convert values to Date format.

  if (!is.na(full.df$PAD.Date[n.pat])) {  # Run a conditional validation or processing step.
    event.dates <- c(event.dates, as.Date(full.df$PAD.Date[n.pat]))  # Define a compact constant vector used below.
  }

  if (!is.na(full.df$LAE.Date[n.pat])) {  # Run a conditional validation or processing step.
    event.dates <- c(event.dates, as.Date(full.df$LAE.Date[n.pat]))  # Define a compact constant vector used below.
  }

  if (!is.na(full.df$CA.Date[n.pat])) {  # Run a conditional validation or processing step.
    event.dates <- c(event.dates, as.Date(full.df$CA.Date[n.pat]))  # Define a compact constant vector used below.
  }

  if (length(event.dates) > 0) {  # Run a conditional validation or processing step.
    earliest.date <- min(event.dates, na.rm = TRUE)  # Calculate an extreme value used for selection or plotting.

    full.df$PAD.LAE.CA.Event[n.pat] <- TRUE  # Update selected rows/columns according to the analysis definition.
    full.df$DateToEvent[n.pat] <- earliest.date  # Update selected rows/columns according to the analysis definition.
    full.df$DaysToEvent[n.pat] <- as.numeric( as.Date(full.df$DateToEvent[n.pat]) - as.Date(full.df$StudyStartDate[n.pat]) )  # Convert values to numeric format.
  }
}

message("Time-to-event outcome for Table 4:")
print(table(full.df$PAD.LAE.CA.Event, useNA = "ifany"))

################################################################################
# 3.4.3 Build Cox regression dataset
################################################################################

abi.cox.df <- full.ABI.df[  # Store an intermediate object used in the following analysis step.
  full.ABI.df$Group == "Intervention",
  c("PatientId", "ScrABIRightValue", "ScrABILeftValue")
]

required.int.vars.table4 <- c("PatientId", "Sex", "Age", "BMI", "DMDuration", "Cluster" )  # intervention-cohort variables required for Cox regression.

missing.int.vars.table4 <- setdiff(required.int.vars.table4, names(int.full.df))  # Store an intermediate object used in the following analysis step.

if (length(missing.int.vars.table4) > 0) {  # Run a conditional validation or processing step.
  stop( "int.full.df is missing required variables for Table 4: ", paste(missing.int.vars.table4, collapse = ", ") )
}

cox.df <- merge( int.full.df[, required.int.vars.table4], abi.cox.df, by = "PatientId", all.x = FALSE, all.y = FALSE )  # Merge source tables using their shared patient identifier.

cox.df <- merge( cox.df, full.df[, c("PatientId", "PAD.LAE.CA.Event", "DaysToEvent")], by = "PatientId", all.x = FALSE, all.y = FALSE )  # Merge source tables using their shared patient identifier.

## Format variables.
cox.df$Sex <- factor( cox.df$Sex, levels = c("Male", "Female") )  # Standardize factor levels for consistent tables, models, or plots.

cox.df$Cluster <- factor( cox.df$Cluster, levels = c("C1 (n=71)", "C2 (n=47)") )  # Standardize factor levels for consistent tables, models, or plots.

cox.df$Age <- as.numeric(cox.df$Age)  # Convert values to numeric format.
cox.df$BMI <- as.numeric(cox.df$BMI)  # Convert values to numeric format.
cox.df$DMDuration <- as.numeric(cox.df$DMDuration)  # Convert values to numeric format.
cox.df$ScrABIRightValue <- as.numeric(cox.df$ScrABIRightValue)  # Convert values to numeric format.
cox.df$ScrABILeftValue <- as.numeric(cox.df$ScrABILeftValue)  # Convert values to numeric format.
cox.df$DaysToEvent <- as.numeric(cox.df$DaysToEvent)  # Convert values to numeric format.
cox.df$event <- as.numeric(cox.df$PAD.LAE.CA.Event)  # Convert values to numeric format.

message("Table 4 Cox regression dataset:")
message("n before complete-case filtering = ", nrow(cox.df))
print(table(cox.df$Cluster, useNA = "ifany"))
print(table(cox.df$event, useNA = "ifany"))

################################################################################
# 3.4.4 Fit Cox proportional hazards models
################################################################################

require_pkg("survival")  # Check that the required package is available.
library(survival)  # Attach package functions used below.

## Model 1: unadjusted
cox.model1 <- coxph( Surv(DaysToEvent, event) ~ Cluster, data = cox.df )  # Fit a Cox proportional hazards model.

## Model 2: adjusted for demographic and clinical covariates
cox.model2 <- coxph(  # Fit a Cox proportional hazards model.
  Surv(DaysToEvent, event) ~ Age + Sex + BMI + DMDuration + Cluster,
  data = cox.df
)

## Model 3: ABI model without cluster
cox.model3 <- coxph(  # Fit a Cox proportional hazards model.
  Surv(DaysToEvent, event) ~ Age + Sex + BMI + DMDuration + ScrABIRightValue + ScrABILeftValue,
  data = cox.df
)

## Model 4: fully adjusted model
cox.model4 <- coxph(  # Fit a Cox proportional hazards model.
  Surv(DaysToEvent, event) ~ Age + Sex + BMI + DMDuration + ScrABIRightValue + ScrABILeftValue + Cluster,
  data = cox.df
)

################################################################################
# 3.4.5 Helper function to extract HR, 95% CI, p values, n and events
################################################################################

extract_cox_results <- function(model, model_name, data, time_var, event_var) {  # Store an intermediate object used in the following analysis step.

  model_summary <- summary(model)  # Store an intermediate object used in the following analysis step.

  ## Get model terms
  beta <- stats::coef(model)  # Extract model coefficient estimates.
  se <- sqrt(diag(stats::vcov(model)))  # Store an intermediate object used in the following analysis step.

  ## Identify variables used in the model formula
  model_vars <- all.vars(stats::formula(model))  # Store an intermediate object used in the following analysis step.

  ## Complete-case dataset actually used by the model
  model_data <- data[complete.cases(data[, model_vars]), ]  # Apply complete-case filtering for the relevant variables.

  n_observations <- nrow(model_data)  # Count rows for reporting.
  n_events <- sum(model_data[[event_var]] == 1, na.rm = TRUE)  # Store an intermediate object used in the following analysis step.

  ## Get p values from summary(model)
  coef_table <- as.data.frame(model_summary$coefficients)  # Convert summary output to a data frame for export.
  coef_table$term <- rownames(coef_table)  # Store an intermediate object used in the following analysis step.
  rownames(coef_table) <- NULL  # Store an intermediate object used in the following analysis step.

  p_col <- grep("Pr\\(", names(coef_table), value = TRUE)  # Store an intermediate object used in the following analysis step.
  if (length(p_col) == 0) {  # Run a conditional validation or processing step.
    p_col <- grep("p", names(coef_table), value = TRUE, ignore.case = TRUE)  # Store an intermediate object used in the following analysis step.
  }
  if (length(p_col) == 0) {  # Run a conditional validation or processing step.
    stop("Could not identify p-value column in Cox model summary.")
  }
  p_col <- p_col[1]  # Store an intermediate object used in the following analysis step.

  ## Confidence intervals on log-HR scale
  ci_log <- tryCatch(  # Store an intermediate object used in the following analysis step.
    {
      suppressMessages(stats::confint(model))
    },
    error = function(e) {
      cbind(
        beta - 1.96 * se,
        beta + 1.96 * se
      )
    }
  )

  ci_log <- as.data.frame(ci_log)  # Convert summary output to a data frame for export.
  ci_log$term <- rownames(ci_log)  # Store an intermediate object used in the following analysis step.
  rownames(ci_log) <- NULL  # Store an intermediate object used in the following analysis step.
  names(ci_log)[1:2] <- c("ci_lower_logHR", "ci_upper_logHR")  # Define a compact constant vector used below.

  out <- data.frame( term = names(beta), estimate_logHR = as.numeric(beta), standard_error = as.numeric(se), HR = exp(as.numeric(beta)), stringsAsFactors = FALSE )  # Create a labelled source-data table for validation/export.

  out <- merge( out, ci_log, by = "term", all.x = TRUE, sort = FALSE )  # Merge source tables using their shared patient identifier.

  out <- merge( out, coef_table[, c("term", p_col)], by = "term", all.x = TRUE, sort = FALSE )  # Merge source tables using their shared patient identifier.

  names(out)[names(out) == p_col] <- "p_value"  # Update selected rows/columns according to the analysis definition.

  out$model <- model_name  # Store an intermediate object used in the following analysis step.
  out$n_observations <- n_observations  # Store an intermediate object used in the following analysis step.
  out$n_events <- n_events  # Store an intermediate object used in the following analysis step.

  out$CI_lower <- exp(out$ci_lower_logHR)  # Store an intermediate object used in the following analysis step.
  out$CI_upper <- exp(out$ci_upper_logHR)  # Store an intermediate object used in the following analysis step.

  out$p_value_formatted <- ifelse( is.na(out$p_value), NA, ifelse(out$p_value < 0.0001, "<0.0001", sprintf("%.4f", out$p_value)) )  # Store an intermediate object used in the following analysis step.

  out$HR_95CI <- paste0( sprintf("%.2f", out$HR), " (", sprintf("%.2f", out$CI_lower), "–", sprintf("%.2f", out$CI_upper), ")" )  # Create a readable label for plotting or reporting.

  out <- out[  # Store an intermediate object used in the following analysis step.
    ,
    c("model", "n_observations", "n_events", "term", "HR", "CI_lower", "CI_upper", "HR_95CI", "p_value", "p_value_formatted", "estimate_logHR", "standard_error" )
  ]

  out
}

################################################################################
# 3.4.6 Extract and assemble Table 4
################################################################################

table4_model1 <- extract_cox_results( model = cox.model1, model_name = "Model 1", data = cox.df, time_var = "DaysToEvent", event_var = "event" )  # Store an intermediate object used in the following analysis step.

table4_model2 <- extract_cox_results( model = cox.model2, model_name = "Model 2", data = cox.df, time_var = "DaysToEvent", event_var = "event" )  # Store an intermediate object used in the following analysis step.

table4_model3 <- extract_cox_results( model = cox.model3, model_name = "Model 3", data = cox.df, time_var = "DaysToEvent", event_var = "event" )  # Store an intermediate object used in the following analysis step.

table4_model4 <- extract_cox_results( model = cox.model4, model_name = "Model 4", data = cox.df, time_var = "DaysToEvent", event_var = "event" )  # Store an intermediate object used in the following analysis step.

table4_all_terms <- rbind( table4_model1, table4_model2, table4_model3, table4_model4 )  # Store an intermediate object used in the following analysis step.

## Main manuscript Table 4 can focus on the cluster term and ABI terms,
## while the full model output is retained in a separate sheet.
table4_main <- table4_all_terms[  # Store an intermediate object used in the following analysis step.
  table4_all_terms$term %in% c("ClusterC2 (n=47)", "ScrABIRightValue", "ScrABILeftValue" ),
]

message("Table 4 main results:")
print(table4_main)

################################################################################
# 3.4.7 Add model descriptions
################################################################################

table4_model_descriptions <- data.frame(  # Create a labelled source-data table for validation/export.
  model = c("Model 1", "Model 2", "Model 3", "Model 4"),
  outcome = "First thromboembolic event: PAD, LAE/PE or CA/stroke",
  model_type = "Cox proportional hazards regression",
  formula = c(
    "Surv(DaysToEvent, event) ~ Cluster",
    "Surv(DaysToEvent, event) ~ Age + Sex + BMI + DMDuration + Cluster",
    "Surv(DaysToEvent, event) ~ Age + Sex + BMI + DMDuration + ScrABIRightValue + ScrABILeftValue",
    "Surv(DaysToEvent, event) ~ Age + Sex + BMI + DMDuration + ScrABIRightValue + ScrABILeftValue + Cluster"
  ),
  adjustment = c(
    "Unadjusted",
    "Adjusted for age, sex, BMI and diabetes duration",
    "Adjusted for age, sex, BMI, diabetes duration and bilateral ABI",
    "Adjusted for age, sex, BMI, diabetes duration, bilateral ABI and temperature variability cluster"
  ),
  stringsAsFactors = FALSE
)

################################################################################
# 3.4.8 Export Table 4 outputs
################################################################################

table4.output.path <- file.path( table.output.folder, "Table_4_cox_regression.xlsx" )  # Build a platform-independent file path.

table4.export.list <- list("Table 4 main" = table4_main, "All model terms" = table4_all_terms, "Model descriptions" = table4_model_descriptions)  # Public-safe Table 4 export list.
if (export.model.datasets) table4.export.list[["Cox dataset"]] <- cox.df  # Optional restricted patient-level model dataset; do not commit publicly.
openxlsx::write.xlsx(x = table4.export.list, file = table4.output.path, overwrite = TRUE, rowNames = FALSE)  # Export Table 4 workbook.

message("Table 4 exported to:")
message(table4.output.path)

message("Analysis completed up to Table 4.")

################################################################################
# 4. Figure 1 source data
################################################################################

## Figure 1 source data:
##   Fig1a: timeline and follow-up periods
##   Fig1b: cohort flow diagram
##   Fig1c: data analytic workflow

required.objects.fig1 <- c("full.df", "patients.df", "study.days.df", "temp.data", "output.folder")  # objects required before exporting Figure 1 source data.
missing.objects.fig1 <- required.objects.fig1[!sapply(required.objects.fig1, exists)]  # Store an intermediate object used in the following analysis step.
if (length(missing.objects.fig1) > 0) {  # Run a conditional validation or processing step.
  stop("Missing objects before Figure 1 source data: ", paste(missing.objects.fig1, collapse = ", "))
}

################################################################################
# 4.1 Figure 1a: timeline and follow-up periods
################################################################################

fig1a_followup_summary <- data.frame(  # Create a labelled source-data table for validation/export.
  metric = c("Median follow-up", "Follow-up range lower", "Follow-up range upper", "Earliest StudyStartDate", "Latest StudyStartDate", "Earliest StudyEndDate", "Latest StudyEndDate", "Follow-up end date"),
  value = c(median(full.df$FollowedUpMonths, na.rm = TRUE), min(full.df$FollowedUpMonths, na.rm = TRUE), max(full.df$FollowedUpMonths, na.rm = TRUE),
            format(min(full.df$StudyStartDate, na.rm = TRUE), "%Y-%m"), format(max(full.df$StudyStartDate, na.rm = TRUE), "%Y-%m"),
            format(min(full.df$StudyEndDate, na.rm = TRUE), "%Y-%m"), format(max(full.df$StudyEndDate, na.rm = TRUE), "%Y-%m"),
            format(unique(full.df$FollowupEndDate)[1], "%Y-%m")),
  unit = c("months", "months", "months", "YYYY-MM", "YYYY-MM", "YYYY-MM", "YYYY-MM", "YYYY-MM"),
  notes = c("Median follow-up duration in the full follow-up cohort.",
            "Lower bound of observed follow-up duration in the full follow-up cohort.",
            "Upper bound of observed follow-up duration in the full follow-up cohort.",
            "Cohort-level first patient study start month.",
            "Cohort-level last patient study start month.",
            "Cohort-level first patient SPDF observation end month.",
            "Cohort-level last patient SPDF observation end month.",
            "Fixed follow-up end date used for the telephone follow-up analysis."),
  stringsAsFactors = FALSE
)

fig1a_timeline <- data.frame(  # Create a labelled source-data table for validation/export.
  timeline_element = c("Total follow-up", "SPDF observation period", "Recruitment period", "Observation time", "Post-SPDF observation period", "Telephone interview period"),
  start_date = c("2018-01", "2018-01", "2018-01", "2018-01", "2021-03", "2023-09"),
  end_date = c("2024-05", "2021-03", "2020-02", "2021-03", "2024-05", "2024-05"),
  displayed_label = c(paste0("Median follow-up: ", sprintf("%.1f", median(full.df$FollowedUpMonths, na.rm = TRUE)), " months (range: ",
                             sprintf("%.1f", min(full.df$FollowedUpMonths, na.rm = TRUE)), "–",
                             sprintf("%.1f", max(full.df$FollowedUpMonths, na.rm = TRUE)), ")"),
                      "SPDF observation period", "recruitment", "observation time", "Post-SPDF observation period", "telephone interview"),
  notes = c("Overall follow-up duration shown above the timeline for the full follow-up cohort.",
            "Cohort-level SPDF observation period shown in the timeline.",
            "Dates indicate the cohort-level range from first to last recruited patient.",
            "Observation period shown within the SPDF study timeline.",
            "Cohort-level post-SPDF follow-up period during which AF and thromboembolic events were recorded.",
            "Period during which structured telephone interviews were conducted."),
  stringsAsFactors = FALSE
)

message("Figure 1a timeline source data:")
print(fig1a_timeline)

################################################################################
# 4.2 Figure 1b: cohort flow diagram
################################################################################

fig1b_flow <- data.frame(  # Create a labelled source-data table for validation/export.
  step_order = 1:18,
  flow_section = c("Screening", "Screening", "Randomization", "Randomization", "Randomization", "SPDF allocation", "SPDF allocation", "SPDF allocation", "SPDF allocation",
                   "SPDF observation period", "SPDF observation period", "SPDF observation period", "SPDF observation period",
                   "Post-SPDF follow-up", "Post-SPDF follow-up", "Post-SPDF follow-up", "Post-SPDF follow-up", "Final follow-up cohort"),
  group = c("Total", "Total", "Total", "Control", "Intervention", "Control", "Intervention", "Control", "Intervention",
            "Control", "Control", "Intervention", "Intervention", "Control", "Intervention", "Control", "Intervention", "Total"),
  n_patients = c(351, 68, 283, 143, 140, 6, 7, 137, 133, 2, 2, 6, 127, 10, 6, 121, 118, 239),
  displayed_label = c("351 screened", "68 screening failures", "283 enrolled and randomized", "143 control", "140 intervention",
                      "6 withdrew before intervention", "7 withdrew before intervention", "137 received standard care", "133 received SPDF intervention",
                      "2 died during SPDF observation", "2 with <30 days observation", "6 with <30 days temperature recordings", "127 entered post-SPDF follow-up",
                      "10 died during follow-up", "6 died during follow-up", "121 completed telephone visit", "118 completed telephone visit",
                      "239 included in final follow-up cohort"),
  notes = c("Number of screened patients.", "Patients not enrolled after screening.", "Patients enrolled and randomized.",
            "Patients randomized to the control group.", "Patients randomized to the intervention group.",
            "Control-group patients who withdrew before intervention start.", "Intervention-group patients who withdrew before intervention start.",
            "Control-group patients available after early withdrawal.", "Intervention-group patients available after early withdrawal.",
            "Control-group deaths during SPDF observation period.", "Control-group patients excluded because observation duration was <30 days.",
            "Intervention-group patients excluded because temperature recordings were available for <30 days.",
            "Intervention-group patients entering post-SPDF follow-up after early exclusions.",
            "Control-group deaths during post-SPDF follow-up.", "Intervention-group deaths during post-SPDF follow-up.",
            "Control-group patients with telephone follow-up data.", "Intervention-group patients with telephone follow-up data.",
            "Patients included in the final full follow-up cohort."),
  stringsAsFactors = FALSE
)

message("Figure 1b flow source data:")
print(fig1b_flow)

################################################################################
# 4.3 Figure 1c: data analytic workflow
################################################################################

fig1c_workflow <- data.frame(  # Create a labelled source-data table for validation/export.
  workflow_step = 1:8,
  input_data = c("Sensor-derived plantar temperature recordings", "Sensor-derived plantar temperature recordings", "Sensor-derived plantar temperature recordings",
                 "Plantar infrared images", "Plantar infrared images", "Follow-up event ascertainment", "Follow-up event ascertainment", "Statistical analysis"),
  analysis_method = c("Post-hoc identification of plantar temperature downshifts", "Classification by PTD status",
                      "Unsupervised clustering of bilateral plantar temperature variability", "Manual image classification",
                      "Classification by PIRI symmetry/asymmetry", "Structured telephone interview and available follow-up data",
                      "Composite outcome definition", "Association analyses"),
  output_or_group = c("PTD assessment", "PTD yes/no", "Lower- and higher-variability clusters", "PIRI assessment",
                      "Symmetric/asymmetric PIRI patterns", "AF, stroke/CA, PE/LAE and PAD outcomes",
                      "Composite AF/thromboembolic outcome", "Associations between plantar temperature asymmetry indicators and AF/thromboembolic outcomes"),
  notes = c("Sensor recordings were analysed to identify plantar temperature downshifts.",
            "Patients were classified as with or without PTDs.",
            "Patients were classified into temperature variability clusters.",
            "Plantar infrared images were manually classified.",
            "PIRI-based symmetry groups were used as an independent plantar thermometry indicator.",
            "AF and thromboembolic outcomes were ascertained during follow-up.",
            "Composite outcome included AF, stroke/CA, PE/LAE and PAD.",
            "Associations were evaluated using group comparisons and regression models."),
  stringsAsFactors = FALSE
)

message("Figure 1c workflow source data:")
print(fig1c_workflow)

################################################################################
# 4.4 Export Figure 1 source data for validation
################################################################################

fig1.output.path <- file.path(figure.output.folder, "Figure_1_source_data_validation.xlsx")  # Build a platform-independent file path.

openxlsx::write.xlsx(  # Export analysis output to an Excel workbook.
  x = list("Fig1a_timeline" = fig1a_timeline, "Fig1a_followup_summary" = fig1a_followup_summary, "Fig1b_flow" = fig1b_flow, "Fig1c_workflow" = fig1c_workflow),
  file = fig1.output.path, overwrite = TRUE, rowNames = FALSE
)

message("Figure 1 source data exported to:")
message(fig1.output.path)
message("Figure 1 source data generation completed.")

################################################################################
# 5. Figure 2 source data and figure generation
################################################################################

## Figure 2:
## Fig2a: post-hoc PTD classification summary
## Fig2b: plantar sensor map metadata
## Fig2c: representative longitudinal temperature recordings
## Fig2d: CVD event overview by PTD group

required.objects.fig2 <- c("int.full.df", "temp.data", "alarms.data.df", "figure.output.folder", "output.ppt.path", "save.as.image.in.ppt", "ppt.template.path")  # objects required before generating Figure 2 source data.
missing.objects.fig2 <- required.objects.fig2[!sapply(required.objects.fig2, exists)]  # Store an intermediate object used in the following analysis step.
if (length(missing.objects.fig2) > 0) stop("Missing objects before Figure 2: ", paste(missing.objects.fig2, collapse = ", "))  # Run a conditional validation or processing step.

################################################################################
# 5.1 Figure 2a: PTD classification summary
################################################################################

if (!"PostHocGroup" %in% names(int.full.df)) stop("int.full.df$PostHocGroup is missing. Please run Table 2 section first.")  # Run a conditional validation or processing step.
if (!all(c("PatientId", "PostHocGroup") %in% names(int.full.df))) stop("int.full.df must contain PatientId and PostHocGroup.")  # Run a conditional validation or processing step.

fig2a_ptd_counts <- as.data.frame(table(int.full.df$PostHocGroup, useNA = "ifany"))  # Convert summary output to a data frame for export.
names(fig2a_ptd_counts) <- c("PTD_group", "n_patients")  # Define a compact constant vector used below.
fig2a_ptd_counts$percentage <- round(100 * fig2a_ptd_counts$n_patients / sum(fig2a_ptd_counts$n_patients), 1)  # Round numeric values for reporting.

fig2a_patient_ptd_summary <- int.full.df[, c("PatientId", "PostHocGroup")]  # Store an intermediate object used in the following analysis step.
fig2a_patient_ptd_summary$PostHocGroup <- as.character(fig2a_patient_ptd_summary$PostHocGroup)  # Convert factor-like values to character strings before relabelling.
fig2a_patient_ptd_summary <- fig2a_patient_ptd_summary[order(fig2a_patient_ptd_summary$PatientId), ]  # Store an intermediate object used in the following analysis step.

################################################################################
# 5.2 Figure 2b: plantar sensor map metadata
################################################################################

fig2b_sensor_map <- data.frame(  # Create a labelled source-data table for validation/export.
  sensor_label = c("D1", "MTK1", "MTK3", "MTK5", "Lat.", "Cal.", "Amb."),
  anatomical_site = c("Hallux / digit 1", "First metatarsal head", "Third metatarsal head", "Fifth metatarsal head", "Lateral midfoot", "Calcaneus / heel", "Ambient temperature"),
  foot_region = c("Forefoot", "Forefoot", "Forefoot", "Forefoot", "Midfoot", "Hindfoot", "Ambient"),
  used_for_PTD_detection = c(TRUE, TRUE, TRUE, TRUE, TRUE, TRUE, FALSE),
  notes = c("Plantar insole temperature sensor.", "Plantar insole temperature sensor.", "Plantar insole temperature sensor.", "Plantar insole temperature sensor.", "Plantar insole temperature sensor.", "Plantar insole temperature sensor.", "Ambient reference temperature."),
  stringsAsFactors = FALSE
)

message("Figure 2b sensor map:")
print(fig2b_sensor_map)

################################################################################
# 5.3 Figure 2c: representative longitudinal temperature recording
################################################################################

## Final Fig2c:
##   Left panel:  P181, MTK5, date window derived from selected frame range 99:147
##   Right panel: P181, MTK5, predefined date window 2019-11-29 to 2019-12-27
##   Both panels are now selected by dates for consistency in source data.

require_pkg("patchwork")  # Check that the required package is available.
library(patchwork)  # Attach package functions used below.

required.objects.fig2c <- c("temp.data", "output.ppt.path", "save.as.image.in.ppt")  # Define a compact constant vector used below.
missing.objects.fig2c <- required.objects.fig2c[!sapply(required.objects.fig2c, exists)]  # Store an intermediate object used in the following analysis step.
if (length(missing.objects.fig2c) > 0) stop("Missing objects before Fig2c: ", paste(missing.objects.fig2c, collapse = ", "))  # Run a conditional validation or processing step.

required.vars.fig2c <- c("PatientId", "Date", "MeasureNo", "R.MTK5", "L.MTK5", "Amb", "Abs.Diff.MTK5")  # temperature-recording columns required for Figure 2c.
missing.vars.fig2c <- setdiff(required.vars.fig2c, names(temp.data))  # Store an intermediate object used in the following analysis step.
if (length(missing.vars.fig2c) > 0) stop("temp.data is missing variables for Fig2c: ", paste(missing.vars.fig2c, collapse = ", "))  # Run a conditional validation or processing step.

temp.data$Date <- as.Date(temp.data$Date)  # Convert values to Date format.
fig2c.patient.id <- "P181"  # Store an intermediate object used in the following analysis step.
fig2c.df0 <- subset(temp.data, PatientId == fig2c.patient.id)  # Subset rows or columns for this analysis step.
fig2c.df0 <- fig2c.df0[order(fig2c.df0$MeasureNo), ]  # Store an intermediate object used in the following analysis step.

fig2c.make.plot <- function(plot.df, plot.title) {  # Store an intermediate object used in the following analysis step.
  ggplot(data = plot.df) +
    geom_line(aes(x = MeasureNo, y = R.MTK5 - 10, group = 1), color = "#00BA38") +
    geom_line(aes(x = MeasureNo, y = L.MTK5 - 10, group = 1), color = "#F564E3") +
    geom_line(aes(x = MeasureNo, y = Amb - 10, group = 1), color = "#111111") +
    geom_bar(aes(x = MeasureNo, y = Abs.Diff.MTK5, fill = Abs.Diff.MTK5), stat = "identity") +
    scale_y_continuous(breaks = seq(10, 20, 5), labels = seq(20, 30, 5)) +
    scale_fill_gradient(limits = c(0, 5)) +
    labs(title = plot.title, x = "Measurement number", y = "Temperature (°C) / absolute difference") +
    m.normal.theme
}

## Panel 1: get dates from selected frame range 99:147, then select data by these dates
fig2c.selected.range <- 99:147  # Store an intermediate object used in the following analysis step.
fig2c_selected_range_frame <- fig2c.df0[fig2c.selected.range, ]  # Store an intermediate object used in the following analysis step.
fig2c.panel1.date.start <- min(fig2c_selected_range_frame$Date, na.rm = TRUE)  # Calculate an extreme value used for selection or plotting.
fig2c.panel1.date.end <- max(fig2c_selected_range_frame$Date, na.rm = TRUE)  # Calculate an extreme value used for selection or plotting.

fig2c_selected_range <- subset(fig2c.df0, Date >= fig2c.panel1.date.start & Date <= fig2c.panel1.date.end)  # Subset rows or columns for this analysis step.
fig2c_selected_range <- fig2c_selected_range[order(fig2c_selected_range$MeasureNo), ]  # Store an intermediate object used in the following analysis step.
fig2c_selected_range$Fig2c_version <- "Panel1"  # Store an intermediate object used in the following analysis step.
fig2c_range_dates <- paste(fig2c.panel1.date.start, fig2c.panel1.date.end, sep = " ~ ")  # Create a readable label for plotting or reporting.
fig2c_range_title <- paste(fig2c.patient.id, "temp. diff. MTK5", fig2c_range_dates)  # Create a readable label for plotting or reporting.
fig2c_range_fig <- fig2c.make.plot(fig2c_selected_range, fig2c_range_title)  # Store an intermediate object used in the following analysis step.

## Panel 2: predefined date-window version
fig2c.panel2.date.start <- as.Date("2019-11-29")  # Convert values to Date format.
fig2c.panel2.date.end <- as.Date("2019-12-27")  # Convert values to Date format.

fig2c_date_window <- subset(fig2c.df0, Date >= fig2c.panel2.date.start & Date <= fig2c.panel2.date.end)  # Subset rows or columns for this analysis step.
fig2c_date_window <- fig2c_date_window[order(fig2c_date_window$MeasureNo), ]  # Store an intermediate object used in the following analysis step.
fig2c_date_window$Fig2c_version <- "Panel2"  # Store an intermediate object used in the following analysis step.
fig2c_date_title <- paste(fig2c.patient.id, "temp. diff. MTK5", paste(fig2c.panel2.date.start, fig2c.panel2.date.end, sep = " ~ "))  # Create a readable label for plotting or reporting.
fig2c_date_fig <- fig2c.make.plot(fig2c_date_window, fig2c_date_title)  # Store an intermediate object used in the following analysis step.

## Combined Fig2c: panel 1 first, panel 2 second
fig2c_combined_fig <- fig2c_range_fig + fig2c_date_fig + plot_layout(ncol = 2, widths = c(1, 1))  # Store an intermediate object used in the following analysis step.
fig2c_combined_fig

save.fig.to.pptx(ggfig = fig2c_combined_fig, path = output.ppt.path, append = TRUE, vertical = FALSE, img.width = 9, img.height = 2.5, save.as.image.in.ppt, ppt_temp_path = ppt.template.path)  # Export figure to the validation PowerPoint file.

fig2c_source_cols <- c("Fig2c_version", required.vars.fig2c)  # Define a compact constant vector used below.
fig2c_source_data <- rbind(fig2c_selected_range[, fig2c_source_cols], fig2c_date_window[, fig2c_source_cols])  # Store an intermediate object used in the following analysis step.

fig2c_panel_summary <- data.frame(  # Create a labelled source-data table for validation/export.
  panel = c("Panel 1", "Panel 2"),
  plotted_order = c("First / left panel", "Second / right panel"),
  source_object = c("fig2c_selected_range", "fig2c_date_window"),
  patient_id = c(fig2c.patient.id, fig2c.patient.id),
  sensor_region = c("MTK5", "MTK5"),
  original_selection = c("Frame range 99:147 used only to derive the date window", "Predefined date window"),
  date_selection = c(paste(fig2c.panel1.date.start, fig2c.panel1.date.end, sep = " to "), paste(fig2c.panel2.date.start, fig2c.panel2.date.end, sep = " to ")),
  n_rows = c(nrow(fig2c_selected_range), nrow(fig2c_date_window)),
  plotted_variables = "R.MTK5 - 10; L.MTK5 - 10; Amb - 10; Abs.Diff.MTK5",
  stringsAsFactors = FALSE
)

message("Final Fig2c order: Panel 1 = frame-derived date window; Panel 2 = predefined date window")
message("Panel 1 date window derived from frames 99:147: ", fig2c.panel1.date.start, " to ", fig2c.panel1.date.end)
message("Panel 2 date window: ", fig2c.panel2.date.start, " to ", fig2c.panel2.date.end)
message("Fig2c combined figure exported to PPT: ", output.ppt.path)

################################################################################
# 5.4 Figure 2d: CVD event overview by PTD group
################################################################################

## Final Fig2d:
##   Patient-level timeline of SPDF observation, PTD status and telephone follow-up
##   with AF, CA, LAE and PAD events overlaid.

required.objects.fig2d <- c("int.full.df", "alarms.data.df", "output.ppt.path", "save.as.image.in.ppt")  # Define a compact constant vector used below.
missing.objects.fig2d <- required.objects.fig2d[!sapply(required.objects.fig2d, exists)]  # Store an intermediate object used in the following analysis step.
if (length(missing.objects.fig2d) > 0) stop("Missing objects before Fig2d: ", paste(missing.objects.fig2d, collapse = ", "))  # Run a conditional validation or processing step.

required.int.vars.fig2d <- c("PatientId", "PostHocGroup", "StudyStartDate", "StudyEndDate", "StudyDays", "AF", "CA.Date", "PAD.Date", "LAE.Date", "AF.CA.LAE.PAD")  # intervention-cohort columns required for Figure 2d.
missing.int.vars.fig2d <- setdiff(required.int.vars.fig2d, names(int.full.df))  # Store an intermediate object used in the following analysis step.
if (length(missing.int.vars.fig2d) > 0) stop("int.full.df is missing variables for Fig2d: ", paste(missing.int.vars.fig2d, collapse = ", "))  # Run a conditional validation or processing step.

required.alarm.vars.fig2d <- c("PatientId", "AlarmTypeEN", "StartDate", "EndDate")  # alarm-table columns required for Figure 2d.
missing.alarm.vars.fig2d <- setdiff(required.alarm.vars.fig2d, names(alarms.data.df))  # Store an intermediate object used in the following analysis step.
if (length(missing.alarm.vars.fig2d) > 0) stop("alarms.data.df is missing variables for Fig2d: ", paste(missing.alarm.vars.fig2d, collapse = ", "))  # Run a conditional validation or processing step.

fig2d.df <- int.full.df  # Store an intermediate object used in the following analysis step.
fig2d.df$StudyStartDate <- as.Date(fig2d.df$StudyStartDate)  # Convert values to Date format.
fig2d.df$StudyEndDate <- as.Date(fig2d.df$StudyEndDate)  # Convert values to Date format.
fig2d.df$CA.Date <- as.Date(fig2d.df$CA.Date)  # Convert values to Date format.
fig2d.df$PAD.Date <- as.Date(fig2d.df$PAD.Date)  # Convert values to Date format.
fig2d.df$LAE.Date <- as.Date(fig2d.df$LAE.Date)  # Convert values to Date format.
alarms.data.df$StartDate <- as.Date(alarms.data.df$StartDate)  # Convert values to Date format.
alarms.data.df$EndDate <- as.Date(alarms.data.df$EndDate)  # Convert values to Date format.

follow.up.end <- as.Date("2024-05-27")  # Convert values to Date format.

normal.temp.strip.text <- paste0("Normal Temp. (n=", length(unique(fig2d.df[fig2d.df$PostHocGroup == "Normal Temp.", "PatientId"])), ")")  # Create a readable label for plotting or reporting.
lowered.temp.strip.text <- paste0("Lowered Temp. (n=", length(unique(fig2d.df[fig2d.df$PostHocGroup == "Lowered Temp.", "PatientId"])), ")")  # Create a readable label for plotting or reporting.
fig2d.df$PostHocGroupFig <- factor(as.character(fig2d.df$PostHocGroup), levels = c("Normal Temp.", "Lowered Temp."), labels = c(normal.temp.strip.text, lowered.temp.strip.text))  # Standardize factor levels for consistent tables, models, or plots.

fig2d_status_df <- data.frame()  # Create a labelled source-data table for validation/export.
fig2d_events_df <- data.frame()  # Create a labelled source-data table for validation/export.

for (pat.id in unique(fig2d.df$PatientId)) {  # Iterate over patients or model outputs.
  per.pat.alarm.df <- subset(alarms.data.df, PatientId == pat.id)  # Subset rows or columns for this analysis step.
  per.pat.int.df <- subset(fig2d.df, PatientId == pat.id, select = c("PatientId", "PostHocGroupFig", "StudyStartDate", "StudyEndDate", "StudyDays", "AF", "CA.Date", "PAD.Date", "LAE.Date"))  # Subset rows or columns for this analysis step.

  if (nrow(per.pat.int.df) > 0) {  # Run a conditional validation or processing step.
    per.pat.summary.df <- per.pat.int.df[, 1:3]  # Store an intermediate object used in the following analysis step.
    per.pat.summary.df$Date <- follow.up.end  # Store an intermediate object used in the following analysis step.
    per.pat.summary.df$Status <- "Tel. follow-up"  # Store an intermediate object used in the following analysis step.

    spdf.end.df <- per.pat.int.df[, 1:3]  # Store an intermediate object used in the following analysis step.
    spdf.end.df$Date <- per.pat.int.df$StudyEndDate[1]  # Store an intermediate object used in the following analysis step.
    spdf.end.df$Status <- "Normal Temp."  # Store an intermediate object used in the following analysis step.
    per.pat.summary.df <- rbind(per.pat.summary.df, spdf.end.df)  # Store an intermediate object used in the following analysis step.

    if (nrow(per.pat.alarm.df) > 0 && "Ischemia" %in% per.pat.alarm.df$AlarmTypeEN) {  # Run a conditional validation or processing step.
      for (ischemia.start.date in per.pat.alarm.df[per.pat.alarm.df$AlarmTypeEN == "Ischemia", "StartDate"]) {  # Iterate over patients or model outputs.
        ischemia.start.summary <- per.pat.int.df[, 1:3]  # Store an intermediate object used in the following analysis step.
        ischemia.start.summary$Date <- ischemia.start.date  # Store an intermediate object used in the following analysis step.
        ischemia.start.summary$Status <- "Normal Temp."  # Store an intermediate object used in the following analysis step.
        per.pat.summary.df <- rbind(per.pat.summary.df, ischemia.start.summary)  # Store an intermediate object used in the following analysis step.
      }
      for (ischemia.end.date in per.pat.alarm.df[per.pat.alarm.df$AlarmTypeEN == "Ischemia", "EndDate"]) {  # Iterate over patients or model outputs.
        ischemia.end.summary <- per.pat.int.df[, 1:3]  # Store an intermediate object used in the following analysis step.
        ischemia.end.summary$Date <- ischemia.end.date  # Store an intermediate object used in the following analysis step.
        ischemia.end.summary$Status <- "Lowered Temp."  # Store an intermediate object used in the following analysis step.
        per.pat.summary.df <- rbind(per.pat.summary.df, ischemia.end.summary)  # Store an intermediate object used in the following analysis step.
      }
    }

    per.pat.summary.df <- per.pat.summary.df[order(per.pat.summary.df$Date), ]  # Store an intermediate object used in the following analysis step.
    per.pat.summary.df$Days <- as.numeric(per.pat.summary.df$Date - per.pat.summary.df$StudyStartDate)  # Convert values to numeric format.
    per.pat.summary.df$DaysDiff <- c(per.pat.summary.df$Days[1], diff(per.pat.summary.df$Days))  # Define a compact constant vector used below.
    per.pat.summary.df$Years <- per.pat.summary.df$Days / 365.25  # Store an intermediate object used in the following analysis step.
    per.pat.summary.df$YearsDiff <- c(per.pat.summary.df$Years[1], diff(per.pat.summary.df$Years))  # Define a compact constant vector used below.
    per.pat.summary.df$StatusLayer <- as.factor(seq(nrow(per.pat.summary.df), 1, -1))  # Store an intermediate object used in the following analysis step.
    fig2d_status_df <- rbind(fig2d_status_df, per.pat.summary.df)  # Store an intermediate object used in the following analysis step.

    per.pat.CVD.df <- data.frame()  # Create a labelled source-data table for validation/export.
    if (!is.na(per.pat.int.df$CA.Date[1])) {  # Run a conditional validation or processing step.
      CA.summary <- per.pat.int.df[, 1:3]  # Store an intermediate object used in the following analysis step.
      CA.summary$Date <- per.pat.int.df$CA.Date[1]  # Store an intermediate object used in the following analysis step.
      CA.summary$Event <- "CA"  # Store an intermediate object used in the following analysis step.
      per.pat.CVD.df <- rbind(per.pat.CVD.df, CA.summary)  # Store an intermediate object used in the following analysis step.
    }
    if (!is.na(per.pat.int.df$PAD.Date[1])) {  # Run a conditional validation or processing step.
      PAD.summary <- per.pat.int.df[, 1:3]  # Store an intermediate object used in the following analysis step.
      PAD.summary$Date <- per.pat.int.df$PAD.Date[1]  # Store an intermediate object used in the following analysis step.
      PAD.summary$Event <- "PAD"  # Store an intermediate object used in the following analysis step.
      per.pat.CVD.df <- rbind(per.pat.CVD.df, PAD.summary)  # Store an intermediate object used in the following analysis step.
    }
    if (!is.na(per.pat.int.df$LAE.Date[1])) {  # Run a conditional validation or processing step.
      LAE.summary <- per.pat.int.df[, 1:3]  # Store an intermediate object used in the following analysis step.
      LAE.summary$Date <- per.pat.int.df$LAE.Date[1]  # Store an intermediate object used in the following analysis step.
      LAE.summary$Event <- "LAE"  # Store an intermediate object used in the following analysis step.
      per.pat.CVD.df <- rbind(per.pat.CVD.df, LAE.summary)  # Store an intermediate object used in the following analysis step.
    }
    if (!is.na(per.pat.int.df$AF[1]) && per.pat.int.df$AF[1] == "Yes") {  # Run a conditional validation or processing step.
      AF.summary <- per.pat.int.df[, 1:3]  # Store an intermediate object used in the following analysis step.
      AF.summary$Date <- per.pat.int.df$StudyStartDate[1] - 100  # Store an intermediate object used in the following analysis step.
      AF.summary$Event <- "AF"  # Store an intermediate object used in the following analysis step.
      per.pat.CVD.df <- rbind(per.pat.CVD.df, AF.summary)  # Store an intermediate object used in the following analysis step.
    }

    if (nrow(per.pat.CVD.df) > 0) {  # Run a conditional validation or processing step.
      per.pat.CVD.df$Days <- as.numeric(per.pat.CVD.df$Date - per.pat.CVD.df$StudyStartDate)  # Convert values to numeric format.
      per.pat.CVD.df$Years <- per.pat.CVD.df$Days / 365.25  # Store an intermediate object used in the following analysis step.
      fig2d_events_df <- rbind(fig2d_events_df, per.pat.CVD.df)  # Store an intermediate object used in the following analysis step.
    }
  }
}

fig2d_status_df$Status <- factor(fig2d_status_df$Status, levels = c("Normal Temp.", "Lowered Temp.", "Tel. follow-up"))  # Standardize factor levels for consistent tables, models, or plots.
fig2d_events_df$Event <- factor(fig2d_events_df$Event, levels = c("AF", "CA", "LAE", "PAD"))  # Standardize factor levels for consistent tables, models, or plots.

fig2d_bar.fill.colors <- c("#66C1A4", "#01B0F0", "#FCB2AF")
fig2d_bar.edge.colors <- rep("#ffffff", length(levels(fig2d_status_df$StatusLayer)))
fig2d_title.text <- paste0("Overview of CVD events between posthoc analysis groups (n=", nrow(fig2d.df), ")")  # Create a readable label for plotting or reporting.
fig2d_caption.text <- "AF: atrial fibrillation; CA: cerebral apoplexy; LAE: lung arterial embolism; PAD: peripheral artery disease"  # Store an intermediate object used in the following analysis step.

fig2d_fig <- ggplot(data = fig2d_status_df) +  # Initialize the ggplot object for this panel.
  geom_bar(aes(y = PatientId, x = YearsDiff, color = StatusLayer, fill = Status), stat = "identity", width = 0.6, linewidth = 0.01) +
  scale_y_discrete(guide = guide_axis(n.dodge = 2)) +
  scale_color_manual(values = fig2d_bar.edge.colors) +
  scale_fill_manual(values = fig2d_bar.fill.colors) +
  geom_point(data = fig2d_events_df, aes(y = PatientId, x = Years, shape = Event), color = "#000000", fill = "#222222", alpha = 0.6, size = 4) +
  scale_shape_manual(values = c("AF" = 21, "CA" = 22, "LAE" = 23, "PAD" = 24)) +
  facet_wrap(. ~ PostHocGroupFig, scales = "free_y") +
  ylab(label = "Patient") +
  xlab(label = "Time (years)") +
  labs(title = fig2d_title.text, caption = fig2d_caption.text) +
  guides(color = "none") +
  m.top.legend.theme

fig2d_fig
save.fig.to.pptx(ggfig = fig2d_fig, path = output.ppt.path, append = TRUE, vertical = FALSE, img.width = 10, img.height = 7, save.as.image.in.ppt, ppt_temp_path = ppt.template.path)  # Export figure to the validation PowerPoint file.

fig2d_event_counts <- as.data.frame(table(fig2d_events_df$PostHocGroupFig, fig2d_events_df$Event))  # Convert summary output to a data frame for export.
names(fig2d_event_counts) <- c("PostHocGroup", "Event", "n_events")  # Define a compact constant vector used below.
fig2d_patient_level <- fig2d.df[, c("PatientId", "PostHocGroup", "PostHocGroupFig", "StudyStartDate", "StudyEndDate", "AF", "CA.Date", "LAE.Date", "PAD.Date", "AF.CA.LAE.PAD")]  # Store an intermediate object used in the following analysis step.

message("Fig2d status rows: ", nrow(fig2d_status_df))
message("Fig2d event rows: ", nrow(fig2d_events_df))
message("Fig2d figure exported to PPT: ", output.ppt.path)

################################################################################
# 5.5 Export Figure 2 source data for validation
################################################################################

## Export style:
##   One Excel file for Figure 2
##   One worksheet per panel: Fig2a, Fig2b, Fig2c, Fig2d
##   Fig2c includes plotted raw temperature data.
##   Fig2d includes minimal patient-level plotting data.

fig2.required.objects <- c("fig2a_ptd_counts", "fig2a_patient_ptd_summary", "fig2b_sensor_map",  # Figure 2 source-data objects required before workbook export.
                           "fig2c_panel_summary", "fig2c_source_data",
                           "fig2d_status_df", "fig2d_events_df", "fig2d_event_counts", "fig2d_patient_level")

fig2.missing.objects <- fig2.required.objects[!sapply(fig2.required.objects, exists)]  # Store an intermediate object used in the following analysis step.
if (length(fig2.missing.objects) > 0) stop("Missing Figure 2 source data objects before export: ", paste(fig2.missing.objects, collapse = ", "))  # Run a conditional validation or processing step.

fig2.validation.path <- file.path(figure.output.folder, "Figure_2_source_data_validation_by_panel.xlsx")  # Build a platform-independent file path.

wb.fig2 <- openxlsx::createWorkbook()  # Store an intermediate object used in the following analysis step.
openxlsx::addWorksheet(wb.fig2, "Fig2a")
openxlsx::addWorksheet(wb.fig2, "Fig2b")
openxlsx::addWorksheet(wb.fig2, "Fig2c")
openxlsx::addWorksheet(wb.fig2, "Fig2d")

header.style <- openxlsx::createStyle(textDecoration = "bold", fontSize = 12)  # Store an intermediate object used in the following analysis step.
subheader.style <- openxlsx::createStyle(textDecoration = "bold", fontSize = 10)  # Store an intermediate object used in the following analysis step.

write_block <- function(wb, sheet, title, x, start_row) {  # Store an intermediate object used in the following analysis step.
  openxlsx::writeData(wb, sheet = sheet, x = title, startRow = start_row, startCol = 1)  # Write a labelled table block to the Excel worksheet.
  openxlsx::addStyle(wb, sheet = sheet, style = subheader.style, rows = start_row, cols = 1, gridExpand = TRUE)
  openxlsx::writeData(wb, sheet = sheet, x = x, startRow = start_row + 1, startCol = 1, rowNames = FALSE)  # Write a labelled table block to the Excel worksheet.
  return(start_row + nrow(x) + 4)
}

## Fig2a worksheet
openxlsx::writeData(wb.fig2, sheet = "Fig2a", x = "Figure 2a source data: post-hoc PTD classification", startRow = 1, startCol = 1)  # Write a labelled table block to the Excel worksheet.
openxlsx::addStyle(wb.fig2, sheet = "Fig2a", style = header.style, rows = 1, cols = 1, gridExpand = TRUE)
row.pos <- 3  # Store an intermediate object used in the following analysis step.
row.pos <- write_block(wb.fig2, "Fig2a", "A1. PTD group counts", fig2a_ptd_counts, row.pos)  # Store an intermediate object used in the following analysis step.
row.pos <- write_block(wb.fig2, "Fig2a", "A2. Patient-level PTD group summary", fig2a_patient_ptd_summary, row.pos)  # Store an intermediate object used in the following analysis step.

## Fig2b worksheet
openxlsx::writeData(wb.fig2, sheet = "Fig2b", x = "Figure 2b source data: plantar sensor map metadata", startRow = 1, startCol = 1)  # Write a labelled table block to the Excel worksheet.
openxlsx::addStyle(wb.fig2, sheet = "Fig2b", style = header.style, rows = 1, cols = 1, gridExpand = TRUE)
openxlsx::writeData(wb.fig2, sheet = "Fig2b", x = fig2b_sensor_map, startRow = 3, startCol = 1, rowNames = FALSE)  # Write a labelled table block to the Excel worksheet.

## Fig2c worksheet
fig2c_temp_cols <- c("Fig2c_version", "PatientId", "Date", "MeasureNo", "R.MTK5", "L.MTK5", "Amb", "Abs.Diff.MTK5")  # Define a compact constant vector used below.
fig2c_source_data_public <- fig2c_source_data[, fig2c_temp_cols]  # Store an intermediate object used in the following analysis step.

openxlsx::writeData(wb.fig2, sheet = "Fig2c", x = "Figure 2c source data: plotted raw temperature recordings", startRow = 1, startCol = 1)  # Write a labelled table block to the Excel worksheet.
openxlsx::addStyle(wb.fig2, sheet = "Fig2c", style = header.style, rows = 1, cols = 1, gridExpand = TRUE)
row.pos <- 3  # Store an intermediate object used in the following analysis step.
row.pos <- write_block(wb.fig2, "Fig2c", "C1. Panel summary", fig2c_panel_summary, row.pos)  # Store an intermediate object used in the following analysis step.
row.pos <- write_block(wb.fig2, "Fig2c", "C2. Combined plotted raw temperature data", fig2c_source_data_public, row.pos)  # Store an intermediate object used in the following analysis step.

## Fig2d worksheet
fig2d_status_public <- fig2d_status_df[, intersect(c("PatientId", "PostHocGroupFig", "StudyStartDate", "Date", "Status", "Days", "DaysDiff", "Years", "YearsDiff", "StatusLayer"), names(fig2d_status_df))]  # Store an intermediate object used in the following analysis step.
fig2d_events_public <- fig2d_events_df[, intersect(c("PatientId", "PostHocGroupFig", "StudyStartDate", "Date", "Event", "Days", "Years"), names(fig2d_events_df))]  # Store an intermediate object used in the following analysis step.
fig2d_patient_public <- fig2d_patient_level[, intersect(c("PatientId", "PostHocGroup", "PostHocGroupFig", "StudyStartDate", "StudyEndDate", "AF", "CA.Date", "LAE.Date", "PAD.Date", "AF.CA.LAE.PAD"), names(fig2d_patient_level))]  # Store an intermediate object used in the following analysis step.

fig2d_panel_summary <- data.frame(  # Create a labelled source-data table for validation/export.
  source_table = c("Timeline status segments", "CVD event points", "Event counts", "Patient-level plotting data"),
  object_name = c("fig2d_status_public", "fig2d_events_public", "fig2d_event_counts", "fig2d_patient_public"),
  n_rows = c(nrow(fig2d_status_public), nrow(fig2d_events_public), nrow(fig2d_event_counts), nrow(fig2d_patient_public)),
  description = c("Bar segments used to draw normal temperature, lowered temperature and telephone follow-up periods.",
                  "Event markers used to draw AF, CA, LAE and PAD points on the timeline.",
                  "Aggregated event counts by PTD group and event type.",
                  "Minimal patient-level variables required to reproduce Fig2d."),
  stringsAsFactors = FALSE
)

openxlsx::writeData(wb.fig2, sheet = "Fig2d", x = "Figure 2d source data: patient-level timeline and CVD event plotting data", startRow = 1, startCol = 1)  # Write a labelled table block to the Excel worksheet.
openxlsx::addStyle(wb.fig2, sheet = "Fig2d", style = header.style, rows = 1, cols = 1, gridExpand = TRUE)
row.pos <- 3  # Store an intermediate object used in the following analysis step.
row.pos <- write_block(wb.fig2, "Fig2d", "D1. Panel summary", fig2d_panel_summary, row.pos)  # Store an intermediate object used in the following analysis step.
row.pos <- write_block(wb.fig2, "Fig2d", "D2. Timeline status segments used for plotted bars", fig2d_status_public, row.pos)  # Store an intermediate object used in the following analysis step.
row.pos <- write_block(wb.fig2, "Fig2d", "D3. CVD event points used for plotted markers", fig2d_events_public, row.pos)  # Store an intermediate object used in the following analysis step.
row.pos <- write_block(wb.fig2, "Fig2d", "D4. Event counts by PTD group and event type", fig2d_event_counts, row.pos)  # Store an intermediate object used in the following analysis step.
row.pos <- write_block(wb.fig2, "Fig2d", "D5. Minimal patient-level plotting data", fig2d_patient_public, row.pos)  # Store an intermediate object used in the following analysis step.

## Formatting and save
for (sheet.name in c("Fig2a", "Fig2b", "Fig2c", "Fig2d")) {  # Iterate over patients or model outputs.
  openxlsx::setColWidths(wb.fig2, sheet = sheet.name, cols = 1:40, widths = "auto")
  openxlsx::freezePane(wb.fig2, sheet = sheet.name, firstRow = TRUE)
}

openxlsx::saveWorkbook(wb.fig2, file = fig2.validation.path, overwrite = TRUE)  # Save the assembled Excel workbook.

message("Figure 2 source data validation file exported to:")
message(fig2.validation.path)
message("Figure 2 source data export completed.")

################################################################################
# 6. Figure 3 source data and figure generation
################################################################################

## Figure 3:
##   Fig3a: temperature variability cluster classification
##   Fig3b: representative P183 MTK temperature recording
##   Fig3c: MTK variability distribution and feature scatter plot
##   Fig3d: clustering validation / silhouette analysis
##   Fig3e: composite AF/thromboembolic outcome by cluster
##   Fig3f: Kaplan-Meier analysis by cluster
##   Fig3g: CVD event overview by cluster

required.objects.fig3 <- c("int.full.df", "output.folder", "data.followup.folder")  # objects required before generating Figure 3 source data.
missing.objects.fig3 <- required.objects.fig3[!sapply(required.objects.fig3, exists)]  # Store an intermediate object used in the following analysis step.
if (length(missing.objects.fig3) > 0) stop("Missing objects before Figure 3: ", paste(missing.objects.fig3, collapse = ", "))  # Run a conditional validation or processing step.

################################################################################
# 6.0 Load Figure 3 temperature and clustering source files
################################################################################

## These Figure 3 input files were generated from the temperature preprocessing
## and patient-level feature extraction workflow.
## Paths are defined centrally in Section 0.2.

required.objects.fig3.input.paths <- c("fig3_temp_data.path", "fig3_temp_features.path")  # Figure 3 input path objects required before file loading.
missing.objects.fig3.input.paths <- required.objects.fig3.input.paths[!sapply(required.objects.fig3.input.paths, exists)]  # Store an intermediate object used in the following analysis step.
if (length(missing.objects.fig3.input.paths) > 0) stop("Missing Figure 3 input path objects: ", paste(missing.objects.fig3.input.paths, collapse = ", "))  # Run a conditional validation or processing step.

if (is.na(fig3_temp_data.path) || !file.exists(fig3_temp_data.path)) stop("Figure 3 cleaned temperature data file not found: ", fig3_temp_data.path)  # Run a conditional validation or processing step.
if (!file.exists(fig3_temp_features.path)) stop("Figure 3 temperature feature file not found: ", fig3_temp_features.path)  # Run a conditional validation or processing step.

fig3_temp_data <- read.csv(fig3_temp_data.path, header = TRUE, sep = ",")  # Read a CSV input file.
fig3_temp_features <- read.csv(fig3_temp_features.path, header = TRUE, sep = ",")  # Read a CSV input file.

## Minimum columns needed here; Amb and Abs.Diff.MTK are harmonized below.
required.fig3.temp.vars <- c("PatientId", "Timestamp", "Date", "R.MTK", "L.MTK")  # minimum cleaned-temperature columns required for Figure 3.
check_vars(fig3_temp_data, required.fig3.temp.vars, "fig3_temp_data", "Fig3b/c")  # Validate required columns before this analysis step.

fig3_temp_data$Date <- as.Date(fig3_temp_data$Date)  # Convert values to Date format.
fig3_temp_data <- fig3_temp_data[order(fig3_temp_data$PatientId, fig3_temp_data$Timestamp), ]  # Store an intermediate object used in the following analysis step.

## Re-normalize MeasureNo per patient, same as the original temp_analysis_v6 logic.
fig3_temp_data2 <- data.frame()  # Create a labelled source-data table for validation/export.
for (pat.id in unique(fig3_temp_data$PatientId)) {  # Iterate over patients or model outputs.
  per.pat.df <- subset(fig3_temp_data, PatientId == pat.id)  # Subset rows or columns for this analysis step.
  per.pat.df <- per.pat.df[order(per.pat.df$Timestamp), ]  # Store an intermediate object used in the following analysis step.
  per.pat.df$MeasureNo <- seq(1, nrow(per.pat.df), 1)  # Store an intermediate object used in the following analysis step.
  fig3_temp_data2 <- rbind(fig3_temp_data2, per.pat.df)  # Store an intermediate object used in the following analysis step.
}
fig3_temp_data <- fig3_temp_data2  # Store an intermediate object used in the following analysis step.

required.fig3.feature.vars <- c("PatientId", "MTK.Asym.Max", "MTK.Asym.Mean")  # patient-level MTK feature columns required for Figure 3.
missing.fig3.feature.vars <- setdiff(required.fig3.feature.vars, names(fig3_temp_features))  # Store an intermediate object used in the following analysis step.
if (length(missing.fig3.feature.vars) > 0) stop("fig3_temp_features is missing variables: ", paste(missing.fig3.feature.vars, collapse = ", "))  # Run a conditional validation or processing step.

message("Figure 3 source files loaded:")
message("fig3_temp_data rows = ", nrow(fig3_temp_data))
message("fig3_temp_features rows = ", nrow(fig3_temp_features))
print(head(fig3_temp_features[, required.fig3.feature.vars]))

################################################################################
# 6.1 Figure 3b and 3c: representative MTK variability and feature distribution
################################################################################

require_pkg("patchwork")  # Check that the required package is available.
require_pkg("ggrepel")  # Check that the required package is available.
library(patchwork)  # Attach package functions used below.
library(ggrepel)  # Attach package functions used below.

required.objects.fig3bc <- c("fig3_temp_data", "fig3_temp_features")  # objects required before Figure 3b/c generation.
missing.objects.fig3bc <- required.objects.fig3bc[!sapply(required.objects.fig3bc, exists)]  # Store an intermediate object used in the following analysis step.
if (length(missing.objects.fig3bc) > 0) stop("Missing objects before Fig3b/c: ", paste(missing.objects.fig3bc, collapse = ", "))  # Run a conditional validation or processing step.

## Harmonize variables.
if (!"Amb" %in% names(fig3_temp_data)) {  # Run a conditional validation or processing step.
  if (all(c("L.Amb", "R.Amb") %in% names(fig3_temp_data))) {  # Run a conditional validation or processing step.
    fig3_temp_data$Amb <- rowMeans(fig3_temp_data[, c("L.Amb", "R.Amb")], na.rm = TRUE)  # Store an intermediate object used in the following analysis step.
  } else {
    stop("Cannot derive Amb because fig3_temp_data has neither Amb nor both L.Amb and R.Amb.")
  }
}

if (!"Abs.Diff.MTK" %in% names(fig3_temp_data)) {  # Run a conditional validation or processing step.
  if ("Diff.MTK" %in% names(fig3_temp_data)) {  # Run a conditional validation or processing step.
    fig3_temp_data$Abs.Diff.MTK <- abs(fig3_temp_data$Diff.MTK)  # Store an intermediate object used in the following analysis step.
  } else if (all(c("R.MTK", "L.MTK") %in% names(fig3_temp_data))) {
    fig3_temp_data$Diff.MTK <- fig3_temp_data$R.MTK - fig3_temp_data$L.MTK  # Store an intermediate object used in the following analysis step.
    fig3_temp_data$Abs.Diff.MTK <- abs(fig3_temp_data$Diff.MTK)  # Store an intermediate object used in the following analysis step.
  } else {
    stop("Cannot derive Abs.Diff.MTK because fig3_temp_data has neither Diff.MTK nor both R.MTK and L.MTK.")
  }
}

required.temp.vars.fig3b <- c("PatientId", "MeasureNo", "R.MTK", "L.MTK", "Amb", "Abs.Diff.MTK")  # temperature columns required for Figure 3b/c plotting.
missing.temp.vars.fig3b <- setdiff(required.temp.vars.fig3b, names(fig3_temp_data))  # Store an intermediate object used in the following analysis step.
if (length(missing.temp.vars.fig3b) > 0) stop("fig3_temp_data is missing variables for Fig3b/c: ", paste(missing.temp.vars.fig3b, collapse = ", "))  # Run a conditional validation or processing step.

required.feature.vars.fig3c <- c("PatientId", "MTK.Asym.Mean", "MTK.Asym.Max")  # feature columns required for Figure 3c scatter plot.
missing.feature.vars.fig3c <- setdiff(required.feature.vars.fig3c, names(fig3_temp_features))  # Store an intermediate object used in the following analysis step.
if (length(missing.feature.vars.fig3c) > 0) stop("fig3_temp_features is missing variables for Fig3c: ", paste(missing.feature.vars.fig3c, collapse = ", "))  # Run a conditional validation or processing step.

## Fig3b: P183 recording around largest Abs.Diff.MTK.
fig3bc.patient.id <- "P183"  # Store an intermediate object used in the following analysis step.
fig3b.df0 <- subset(fig3_temp_data, PatientId == fig3bc.patient.id)  # Subset rows or columns for this analysis step.
fig3b.df0 <- fig3b.df0[order(fig3b.df0$MeasureNo), ]  # Store an intermediate object used in the following analysis step.

if (nrow(fig3b.df0) == 0) stop("No temperature rows found for Fig3b/c patient: ", fig3bc.patient.id)  # Run a conditional validation or processing step.

fig3b.largest.delta.frame <- which(fig3b.df0$Abs.Diff.MTK == max(fig3b.df0$Abs.Diff.MTK, na.rm = TRUE))[1]  # Store an intermediate object used in the following analysis step.
fig3b.range.start <- max(1, fig3b.largest.delta.frame - 30)  # Calculate an extreme value used for selection or plotting.
fig3b.range.end <- min(nrow(fig3b.df0), fig3b.largest.delta.frame + 45)  # Calculate an extreme value used for selection or plotting.
fig3b.selected.range <- fig3b.range.start:fig3b.range.end  # Store an intermediate object used in the following analysis step.

fig3b_source_data <- fig3b.df0[fig3b.selected.range, c("PatientId", "MeasureNo", "R.MTK", "L.MTK", "Amb", "Abs.Diff.MTK")]  # Store an intermediate object used in the following analysis step.
if ("Date" %in% names(fig3b.df0)) fig3b_source_data$Date <- fig3b.df0[fig3b.selected.range, "Date"]  # Store an intermediate object used in the following analysis step.
if ("Timestamp" %in% names(fig3b.df0)) fig3b_source_data$Timestamp <- fig3b.df0[fig3b.selected.range, "Timestamp"]  # Store an intermediate object used in the following analysis step.
fig3b_source_data$Fig3_panel <- "Fig3b"  # Store an intermediate object used in the following analysis step.
fig3b_source_data$Selection <- paste0("P183 records around largest Abs.Diff.MTK; frame range ", fig3b.range.start, ":", fig3b.range.end)  # Create a readable label for plotting or reporting.

fig3b_fig <- ggplot(data = fig3b_source_data) +  # Initialize the ggplot object for this panel.
  geom_line(aes(x = MeasureNo, y = L.MTK - 10, group = 1), color = "#F564E3", linewidth = 0.8) +
  geom_line(aes(x = MeasureNo, y = R.MTK - 10, group = 1), color = "#00BA38", linewidth = 0.8) +
  geom_line(aes(x = MeasureNo, y = Amb - 10, group = 1), color = "#111111", linewidth = 0.6) +
  geom_bar(aes(x = MeasureNo, y = Abs.Diff.MTK, fill = Abs.Diff.MTK), stat = "identity") +
  scale_y_continuous(breaks = seq(10, 20, 5), labels = seq(20, 30, 5)) +
  scale_fill_gradient(limits = c(0, 8)) +
  labs(title = paste0(fig3bc.patient.id, "   Sensor: L-MTK  R-MTK  Ambient"),
       x = "Recording (n)", y = expression(T~"("*degree*C*")")) +
  m.normal.theme

## Fig3c left: P183 Abs.Diff.MTK distribution.
fig3c_boxplot_source_data <- fig3b.df0[, c("PatientId", "MeasureNo", "Abs.Diff.MTK")]  # Store an intermediate object used in the following analysis step.
if ("Date" %in% names(fig3b.df0)) fig3c_boxplot_source_data$Date <- fig3b.df0$Date  # Store an intermediate object used in the following analysis step.
if ("Timestamp" %in% names(fig3b.df0)) fig3c_boxplot_source_data$Timestamp <- fig3b.df0$Timestamp  # Store an intermediate object used in the following analysis step.
fig3c_boxplot_source_data$Fig3_panel <- "Fig3c_left_boxplot"  # Store an intermediate object used in the following analysis step.
fig3c_boxplot_source_data$Highlighted <- fig3c_boxplot_source_data$Abs.Diff.MTK == max(fig3c_boxplot_source_data$Abs.Diff.MTK, na.rm = TRUE)  # Store an intermediate object used in the following analysis step.

fig3c_boxplot <- ggplot(data = fig3c_boxplot_source_data, aes(x = 1, y = Abs.Diff.MTK)) +  # Initialize the ggplot object for this panel.
  geom_boxplot(size = 0.3, width = 0.45, outlier.size = 0.7) +
  stat_summary(fun = mean, geom = "point", shape = 23, size = 2.5, fill = "white", color = "black") +
  geom_point(data = subset(fig3c_boxplot_source_data, Highlighted), aes(x = 1, y = Abs.Diff.MTK), shape = 24, size = 3, fill = "#2C7FB8", color = "black") +
  scale_x_continuous(breaks = c(1), labels = "") +
  labs(x = NULL, y = expression("|"*Delta*T[MTK]*"|"~"("*degree*C*")")) +
  m.normal.theme

## Fig3c right: all-patient feature scatter.
fig3c_scatter_source_data <- fig3_temp_features[, c("PatientId", "MTK.Asym.Mean", "MTK.Asym.Max")]  # Store an intermediate object used in the following analysis step.
fig3c_scatter_source_data$Fig3_panel <- "Fig3c_right_scatter"  # Store an intermediate object used in the following analysis step.
fig3c_scatter_source_data$Highlighted <- fig3c_scatter_source_data$PatientId == fig3bc.patient.id  # Store an intermediate object used in the following analysis step.
fig3c_highlighted_point <- subset(fig3c_scatter_source_data, Highlighted)  # Subset rows or columns for this analysis step.

if (nrow(fig3c_highlighted_point) == 0) stop("No feature row found for highlighted Fig3c patient: ", fig3bc.patient.id)  # Run a conditional validation or processing step.

fig3c_scatter <- ggplot(data = fig3c_scatter_source_data, aes(x = MTK.Asym.Mean, y = MTK.Asym.Max)) +  # Initialize the ggplot object for this panel.
  geom_point(size = 0.7, color = "black") +
  geom_point(data = fig3c_highlighted_point, size = 2.2, color = "#2C7FB8") +
  geom_text_repel(data = fig3c_highlighted_point, aes(label = PatientId), vjust = -0.5, color = "black", size = 3) +
  labs(title = "MTK",
       x = expression(italic(mean)*"(|"*Delta*T[MTK]*"|)"~"("*degree*C*")"),
       y = expression(italic(max)*"(|"*Delta*T[MTK]*"|)"~"("*degree*C*")")) +
  m.normal.theme

fig3c_fig <- fig3c_boxplot + fig3c_scatter + plot_layout(ncol = 2, widths = c(0.65, 1.8))  # Store an intermediate object used in the following analysis step.
fig3bc_combined_fig <- fig3b_fig / fig3c_fig + plot_layout(heights = c(1.05, 1))  # Store an intermediate object used in the following analysis step.

fig3bc_combined_fig
save.fig.to.pptx(ggfig = fig3bc_combined_fig, path = fig3.output.ppt.path, img.width = 6, img.height = 5, append = TRUE, save.as.image.in.ppt, ppt_temp_path = ppt.template.path)  # Export figure to the validation PowerPoint file.

fig3bc_panel_summary <- data.frame(  # Create a labelled source-data table for validation/export.
  panel = c("Fig3b", "Fig3c left", "Fig3c right"),
  source_object = c("fig3b_source_data", "fig3c_boxplot_source_data", "fig3c_scatter_source_data"),
  patient_id = c(fig3bc.patient.id, fig3bc.patient.id, fig3bc.patient.id),
  description = c("P183 MTK recording around largest Abs.Diff.MTK.", "Distribution of P183 Abs.Diff.MTK values.", "Patient-level MTK.Asym.Mean vs MTK.Asym.Max with P183 highlighted."),
  n_rows = c(nrow(fig3b_source_data), nrow(fig3c_boxplot_source_data), nrow(fig3c_scatter_source_data)),
  stringsAsFactors = FALSE
)

message("Fig3b/c figure exported to PPT: ", fig3.output.ppt.path)
print(fig3bc_panel_summary)

################################################################################
# 6.2 Figure 3a and 3d: clustering using original script logic
################################################################################

require_pkg("factoextra")  # Check that the required package is available.
require_pkg("cluster")  # Check that the required package is available.
library(factoextra)  # Attach package functions used below.
library(cluster)  # Attach package functions used below.

cluster.vars1 <- c("MTK.Asym.Max", "MTK.Asym.Mean")  # two MTK asymmetry features used for Figure 3 k-means clustering.
missing.cluster.vars1 <- setdiff(cluster.vars1, names(fig3_temp_features))  # Store an intermediate object used in the following analysis step.
if (length(missing.cluster.vars1) > 0) stop("fig3_temp_features is missing clustering variables: ", paste(missing.cluster.vars1, collapse = ", "))  # Run a conditional validation or processing step.

outcome.vars <- c("AF", "PAD", "PAD.LAE.CA", "AF.CA.LAE.PAD", "Gout", "Trauma", "Arthrosis", "Gout.Trauma.Arthrosis", "Gout.Trauma.CVD.Arthrosis")  # optional clinical outcome columns retained if present.
covariate.vars <- c("Age", "Sex")  # optional covariates retained if present.

available.outcome.vars <- intersect(outcome.vars, names(int.full.df))  # Store an intermediate object used in the following analysis step.
available.covariate.vars <- intersect(covariate.vars, names(int.full.df))  # Store an intermediate object used in the following analysis step.

fig3_clinical_for_cluster <- int.full.df[, unique(c("PatientId", available.outcome.vars, available.covariate.vars))]  # Store an intermediate object used in the following analysis step.
full.df.fig3 <- merge(fig3_temp_features, fig3_clinical_for_cluster, by = "PatientId", all = FALSE)  # Merge source tables using their shared patient identifier.
full.df.fig3 <- full.df.fig3[order(full.df.fig3$PatientId), ]  # Store an intermediate object used in the following analysis step.

temp.clustering.df1 <- subset(full.df.fig3, select = unique(c(cluster.vars1, "PatientId", available.outcome.vars, available.covariate.vars)))  # Subset rows or columns for this analysis step.
cluster.df <- temp.clustering.df1  # Store an intermediate object used in the following analysis step.
cluster.features <- cluster.vars1  # Store an intermediate object used in the following analysis step.

if (any(!complete.cases(cluster.df[, cluster.features]))) {  # Run a conditional validation or processing step.
  stop("Missing values found in clustering features. Please check MTK.Asym.Max and MTK.Asym.Mean before reproducing Fig3d.")
}

## Same feature scaling logic as the original cluster_temp_by_pat_v1 script.
data_scaled <- scale(cluster.df[, c(1:length(cluster.vars1))])  # Store an intermediate object used in the following analysis step.

message("Fig3 clustering input reconstructed:")
message("n patients = ", length(unique(cluster.df$PatientId)))
print(summary(cluster.df[, cluster.features]))

################################################################################
# 6.2.1 Fig3d: silhouette method with fixed seed for reproducibility
################################################################################

## Seed 15 reproduces the final Fig3d curve used for the manuscript figure.
fig3.seed.silhouette <- 15  # Store an intermediate object used in the following analysis step.
set.seed(fig3.seed.silhouette)
Silhouette.fig.raw <- factoextra::fviz_nbclust(data_scaled, kmeans, method = "silhouette")  # Store an intermediate object used in the following analysis step.
Silhouette.fig.raw

fig3d_k_selection_source_data <- Silhouette.fig.raw$data  # Store an intermediate object used in the following analysis step.
fig3d_k_selection_source_data$clusters <- as.numeric(as.character(fig3d_k_selection_source_data$clusters))  # Convert values to numeric format.
fig3d_k_selection_source_data$y <- as.numeric(as.character(fig3d_k_selection_source_data$y))  # Convert values to numeric format.
fig3d_k_selection_source_data$Fig3_panel <- "Fig3d"  # Store an intermediate object used in the following analysis step.
fig3d_k_selection_source_data$method <- "fviz_nbclust(data_scaled, kmeans, method = 'silhouette')"  # Store an intermediate object used in the following analysis step.
fig3d_k_selection_source_data$seed <- fig3.seed.silhouette  # Store an intermediate object used in the following analysis step.
fig3d_k_selection_source_data$R_version <- R.version.string  # Store an intermediate object used in the following analysis step.

Silhouette.fig <- ggplot(data = fig3d_k_selection_source_data, aes(x = clusters, y = y)) +  # Initialize the ggplot object for this panel.
  geom_line(color = "black", linewidth = 0.7) +
  geom_point(color = "black", size = 2) +
  geom_vline(xintercept = 2, linetype = "dashed", color = "black", linewidth = 0.5) +
  scale_x_continuous(breaks = 1:10, limits = c(0.7, 10.3)) +
  scale_y_continuous(limits = c(0, 0.5), breaks = seq(0, 0.4, 0.1)) +
  labs(title = "Optimal number of clusters",
       x = "Number of clusters k",
       y = "Average silhouette width") +
  theme_classic() +
  theme(text = element_text(size = 10, family = "sans", color = "black"),
        plot.title = element_text(size = 10, face = "plain", hjust = 0),
        axis.title = element_text(size = 10, face = "bold", color = "black"),
        axis.text = element_text(size = 10, color = "black"),
        axis.line = element_line(color = "black", linewidth = 0.5),
        axis.ticks = element_line(color = "black", linewidth = 0.5),
        legend.position = "none")

Silhouette.fig
save.fig.to.pptx(ggfig = Silhouette.fig, path = fig3.output.ppt.path, img.width = 2.2, img.height = 2.4, append = TRUE, save.as.image.in.ppt, ppt_temp_path = ppt.template.path)  # Export figure to the validation PowerPoint file.

################################################################################
# 6.2.2 Fig3a: final k-means clustering using original script logic
################################################################################

optimal_k <- 2  # Store an intermediate object used in the following analysis step.
set.seed(42)
kmeans_result <- kmeans(data_scaled, centers = optimal_k, nstart = 100)  # Store an intermediate object used in the following analysis step.

cluster.df$Kmeans.Cluster.Raw <- as.factor(kmeans_result$cluster)  # Store an intermediate object used in the following analysis step.
freq.df <- as.data.frame(table(cluster.df$Kmeans.Cluster.Raw))  # Convert summary output to a data frame for export.
freq.df$Label <- paste("C", freq.df$Var1, " (n=", freq.df$Freq, ")", sep = "")  # Create a readable label for plotting or reporting.

cluster.df$Kmeans.Cluster <- factor(cluster.df$Kmeans.Cluster.Raw, levels = freq.df$Var1, labels = freq.df$Label)  # Standardize factor levels for consistent tables, models, or plots.

fig3a_source_data <- cluster.df[, unique(c("PatientId", cluster.features, "Kmeans.Cluster.Raw", "Kmeans.Cluster", available.outcome.vars, available.covariate.vars))]  # Store an intermediate object used in the following analysis step.

## Check consistency between regenerated k-means clusters and imported manuscript cluster labels.
fig3_cluster_check <- merge( fig3a_source_data[, c("PatientId", "Kmeans.Cluster")], int.full.df[, c("PatientId", "Cluster")], by = "PatientId", all = FALSE )  # Merge source tables using their shared patient identifier.

fig3_cluster_check$Kmeans.Cluster <- as.character(fig3_cluster_check$Kmeans.Cluster)  # Convert factor-like values to character strings before relabelling.
fig3_cluster_check$Cluster <- as.character(fig3_cluster_check$Cluster)  # Convert factor-like values to character strings before relabelling.
fig3_cluster_check$cluster_match <- fig3_cluster_check$Kmeans.Cluster == fig3_cluster_check$Cluster  # Store an intermediate object used in the following analysis step.

if (!all(fig3_cluster_check$cluster_match, na.rm = TRUE)) {  # Run a conditional validation or processing step.
  print(fig3_cluster_check[!fig3_cluster_check$cluster_match, ])
  stop("Regenerated Fig3a k-means clusters do not match imported manuscript cluster labels. Please check Fig3 clustering inputs before continuing.")
} else {
  message("Regenerated Fig3a k-means clusters match imported manuscript cluster labels.")
}

fig3a_cluster_counts <- freq.df  # Store an intermediate object used in the following analysis step.
names(fig3a_cluster_counts) <- c("Kmeans.Cluster.Raw", "n_patients", "ClusterLabel")  # Define a compact constant vector used below.

fig3d_silhouette_summary <- data.frame(  # Create a labelled source-data table for validation/export.
  method = "fviz_nbclust + kmeans",
  silhouette_seed = fig3.seed.silhouette,
  kmeans_seed = 42,
  clustering_features = paste(cluster.vars1, collapse = "; "),
  selected_k = optimal_k,
  n_patients = nrow(cluster.df),
  stringsAsFactors = FALSE
)

message("Fig3 k-means cluster counts:")
print(fig3a_cluster_counts)

fig3a_fig <- ggplot(data = fig3a_source_data, aes(x = MTK.Asym.Mean, y = MTK.Asym.Max, color = Kmeans.Cluster)) +  # Initialize the ggplot object for this panel.
  geom_point(size = 1.5, alpha = 0.85) +
  labs(title = "Temperature variability clusters",
       x = expression(italic(mean)*"(|"*Delta*T[MTK]*"|)"~"("*degree*C*")"),
       y = expression(italic(max)*"(|"*Delta*T[MTK]*"|)"~"("*degree*C*")"),
       color = "Cluster") +
  m.normal.theme

fig3a_fig
save.fig.to.pptx(ggfig = fig3a_fig, path = fig3.output.ppt.path, img.width = 4, img.height = 3.5, append = TRUE, save.as.image.in.ppt, ppt_temp_path = ppt.template.path)  # Export figure to the validation PowerPoint file.

fig3ad_panel_summary <- data.frame(  # Create a labelled source-data table for validation/export.
  panel = c("Fig3a", "Fig3d"),
  source_object = c("fig3a_source_data", "fig3d_k_selection_source_data"),
  description = c("Patient-level MTK.Asym.Mean and MTK.Asym.Max feature space with k-means cluster labels.",
                  "Average silhouette width across candidate k values generated by fviz_nbclust using kmeans."),
  n_rows = c(nrow(fig3a_source_data), nrow(fig3d_k_selection_source_data)),
  stringsAsFactors = FALSE
)

message("Fig3a/d figures exported to PPT: ", fig3.output.ppt.path)
print(fig3ad_panel_summary)
print(fig3d_silhouette_summary)

################################################################################
# 6.3 Figure 3e-g: outcome summaries and Kaplan-Meier analysis by cluster
################################################################################

## Fig3e:
##   Composite AF/thromboembolic outcome by temperature variability cluster.
##
## Fig3f:
##   Kaplan-Meier event-free survival for first thromboembolic event
##   PAD, LAE/PE, or CA/stroke by cluster.
##
## Fig3g:
##   Patient-level CVD event overview timeline by temperature variability cluster.

required.objects.fig3eg <- c("int.full.df", "cox.df", "fig3.output.ppt.path")  # objects required before Figure 3e-g generation.
missing.objects.fig3eg <- required.objects.fig3eg[!sapply(required.objects.fig3eg, exists)]  # Store an intermediate object used in the following analysis step.
if (length(missing.objects.fig3eg) > 0) stop("Missing objects before Fig3e-g: ", paste(missing.objects.fig3eg, collapse = ", "))  # Run a conditional validation or processing step.

required.vars.fig3eg <- c("PatientId", "Cluster", "AF", "CA", "LAE", "PAD", "AF.CA.LAE.PAD")  # intervention-cohort columns required for Figure 3e-g.
missing.vars.fig3eg <- setdiff(required.vars.fig3eg, names(int.full.df))  # Store an intermediate object used in the following analysis step.
if (length(missing.vars.fig3eg) > 0) stop("int.full.df is missing variables for Fig3e-g: ", paste(missing.vars.fig3eg, collapse = ", "))  # Run a conditional validation or processing step.

################################################################################
# 6.3.1 Fig3e: composite AF/thromboembolic outcome by cluster
################################################################################

fig3e_source_data <- int.full.df[, c("PatientId", "Cluster", "AF.CA.LAE.PAD")]  # Store an intermediate object used in the following analysis step.
fig3e_source_data$Cluster <- factor(fig3e_source_data$Cluster, levels = c("C1 (n=71)", "C2 (n=47)"))  # Standardize factor levels for consistent tables, models, or plots.
fig3e_source_data$AF.CA.LAE.PAD <- factor(fig3e_source_data$AF.CA.LAE.PAD, levels = c("No", "Yes"))  # Standardize factor levels for consistent tables, models, or plots.

fig3e_summary <- as.data.frame(table(fig3e_source_data$Cluster, fig3e_source_data$AF.CA.LAE.PAD))  # Convert summary output to a data frame for export.
names(fig3e_summary) <- c("Cluster", "CompositeOutcome", "n_patients")  # Define a compact constant vector used below.
fig3e_cluster_n <- as.data.frame(table(fig3e_source_data$Cluster))  # Convert summary output to a data frame for export.
names(fig3e_cluster_n) <- c("Cluster", "cluster_n")  # Define a compact constant vector used below.
fig3e_summary <- merge(fig3e_summary, fig3e_cluster_n, by = "Cluster", all.x = TRUE)  # Merge source tables using their shared patient identifier.
fig3e_summary$percentage <- round(100 * fig3e_summary$n_patients / fig3e_summary$cluster_n, 1)  # Round numeric values for reporting.

fig3e_plot_data <- subset(fig3e_summary, CompositeOutcome == "Yes")  # Subset rows or columns for this analysis step.

fig3e_p_value <- fisher.test(table(fig3e_source_data$Cluster, fig3e_source_data$AF.CA.LAE.PAD))$p.value  # Run Fisher’s exact test for group comparison.
fig3e_p_label <- ifelse(fig3e_p_value < 0.0001, "p < 0.0001", paste0("p = ", sprintf("%.4f", fig3e_p_value)))  # Store an intermediate object used in the following analysis step.

fig3e_fig <- ggplot(data = fig3e_plot_data, aes(x = Cluster, y = percentage, fill = Cluster)) +  # Initialize the ggplot object for this panel.
  geom_bar(stat = "identity", width = 0.55, color = "black", linewidth = 0.2) +
  geom_text(aes(label = paste0(n_patients, "/", cluster_n, "\n", percentage, "%")), vjust = -0.4, size = 3) +
  labs(title = "Composite AF/thromboembolic outcome", subtitle = fig3e_p_label, x = NULL, y = "Patients with event (%)") +
  scale_y_continuous(limits = c(0, max(fig3e_plot_data$percentage, na.rm = TRUE) * 1.25)) +
  guides(fill = "none") +
  m.normal.theme

fig3e_fig
save.fig.to.pptx(ggfig = fig3e_fig, path = fig3.output.ppt.path, img.width = 3.2, img.height = 3.2, append = TRUE, save.as.image.in.ppt, ppt_temp_path = ppt.template.path)  # Export figure to the validation PowerPoint file.

################################################################################
# 6.3.2 Fig3f: Kaplan-Meier analysis by cluster, original script logic
################################################################################

## Original logic:
##   cox.df$YearsToEvent <- cox.df$DaysToEvent / 365.25
##   surv.df2 <- subset(cox.df, select = c("PatientId", "Cluster", "PAD.LAE.CA.Event", "YearsToEvent"))
##   surv.obj2 <- draw.survial.plot(..., title.text = "Event by cluster group", ...)

required.vars.fig3f <- c("PatientId", "Cluster", "PAD.LAE.CA.Event", "DaysToEvent")  # Cox dataset columns required for Figure 3f KM analysis.
missing.vars.fig3f <- setdiff(required.vars.fig3f, names(cox.df))  # Store an intermediate object used in the following analysis step.
if (length(missing.vars.fig3f) > 0) stop("cox.df is missing variables for Fig3f: ", paste(missing.vars.fig3f, collapse = ", "))  # Run a conditional validation or processing step.

cox.df$DaysToEvent <- as.numeric(cox.df$DaysToEvent)  # Convert values to numeric format.
cox.df$PAD.LAE.CA.Event <- as.numeric(cox.df$PAD.LAE.CA.Event)  # Convert values to numeric format.
cox.df$YearsToEvent <- cox.df$DaysToEvent / 365.25  # Store an intermediate object used in the following analysis step.
cox.df$Cluster <- factor(cox.df$Cluster, levels = c("C1 (n=71)", "C2 (n=47)"))  # Standardize factor levels for consistent tables, models, or plots.

fig3f_km_dataset <- subset(cox.df, select = c("PatientId", "Cluster", "PAD.LAE.CA.Event", "YearsToEvent"))  # Subset rows or columns for this analysis step.
fig3f_km_dataset <- fig3f_km_dataset[complete.cases(fig3f_km_dataset), ]  # Apply complete-case filtering for the relevant variables.
fig3f_km_dataset$Cluster <- factor(fig3f_km_dataset$Cluster, levels = c("C1 (n=71)", "C2 (n=47)"))  # Standardize factor levels for consistent tables, models, or plots.

surv.obj2 <- draw.survial.plot(  # Store an intermediate object used in the following analysis step.
  surv.df = fig3f_km_dataset,
  xlab.text = "Time (years)",
  ylab.text = "Free of events (%)",
  title.text = "Event by cluster group",
  font.size = 10,
  legend.labs.texts = levels(fig3f_km_dataset$Cluster),
  legend.colors = c("#F8766D", "#00BFC4"),
  show.censor.points = FALSE
)

surv.obj2

fig3f_fig <- surv.obj2$plot  # Store an intermediate object used in the following analysis step.
fig3f_risk_table_fig <- surv.obj2$table  # Store an intermediate object used in the following analysis step.

save.fig.to.pptx( ggfig = fig3f_fig, path = fig3.output.ppt.path, append = TRUE, vertical = FALSE, img.width = 2.5, img.height = 3.0, save.as.image.in.ppt, ppt_temp_path = ppt.template.path )  # Export figure to the validation PowerPoint file.

save.fig.to.pptx(  # Export figure to the validation PowerPoint file.
  ggfig = fig3f_risk_table_fig,
  path = fig3.output.ppt.path,
  append = TRUE,
  vertical = FALSE,
  img.width = 2.5,
  img.height = 2.0,
  save.as.image.in.ppt,
  ppt_temp_path = ppt.template.path
)

## Source data for validation/export.
fig3f_surv_fit <- survival::survfit( survival::Surv(YearsToEvent, PAD.LAE.CA.Event) ~ Cluster, data = fig3f_km_dataset )  # Fit Kaplan-Meier survival curves.

fig3f_surv_summary <- summary(fig3f_surv_fit)  # Store an intermediate object used in the following analysis step.

fig3f_km_curve_source_data <- data.frame(  # Create a labelled source-data table for validation/export.
  time_years = fig3f_surv_summary$time,
  survival_probability = fig3f_surv_summary$surv,
  free_of_events_percentage = 100 * fig3f_surv_summary$surv,
  standard_error = fig3f_surv_summary$std.err,
  lower_95CI = fig3f_surv_summary$lower,
  upper_95CI = fig3f_surv_summary$upper,
  n_risk = fig3f_surv_summary$n.risk,
  n_event = fig3f_surv_summary$n.event,
  n_censor = fig3f_surv_summary$n.censor,
  strata = fig3f_surv_summary$strata,
  stringsAsFactors = FALSE
)

fig3f_km_curve_source_data$Cluster <- gsub("Cluster=", "", fig3f_km_curve_source_data$strata)  # Set or standardize analysis group labels.

fig3f_logrank <- survival::survdiff( survival::Surv(YearsToEvent, PAD.LAE.CA.Event) ~ Cluster, data = fig3f_km_dataset )  # Run a log-rank test.

fig3f_logrank_p <- 1 - pchisq(fig3f_logrank$chisq, df = length(fig3f_logrank$n) - 1)  # Store an intermediate object used in the following analysis step.

fig3f_logrank_summary <- data.frame(  # Create a labelled source-data table for validation/export.
  method = "Log-rank test",
  grouping_variable = "Cluster",
  outcome = "First PAD/LAE/CA thromboembolic event",
  chisq = fig3f_logrank$chisq,
  df = length(fig3f_logrank$n) - 1,
  p_value = fig3f_logrank_p,
  p_value_label = ifelse(fig3f_logrank_p < 0.0001, "p < 0.0001", paste0("p = ", sprintf("%.4f", fig3f_logrank_p))),
  stringsAsFactors = FALSE
)

message("Fig3f survival plot and risk table exported to PPT: ", fig3.output.ppt.path)
print(fig3f_logrank_summary)

################################################################################
# 6.3.3 Fig3g: CVD event overview by temperature variability cluster
################################################################################

## Original Fig3g logic:
##   Patient-level SPDF and telephone follow-up timeline by cluster
##   with AF, CA, LAE and PAD event markers overlaid.

required.objects.fig3g <- c("int.full.df", "fig3.output.ppt.path")  # objects required before Figure 3g generation.
missing.objects.fig3g <- required.objects.fig3g[!sapply(required.objects.fig3g, exists)]  # Store an intermediate object used in the following analysis step.
if (length(missing.objects.fig3g) > 0) stop("Missing objects before Fig3g: ", paste(missing.objects.fig3g, collapse = ", "))  # Run a conditional validation or processing step.

required.vars.fig3g <- c("PatientId", "Cluster", "StudyStartDate", "StudyEndDate", "StudyDays", "AF", "CA.Date", "PAD.Date", "LAE.Date")  # intervention-cohort columns required for Figure 3g timeline.
missing.vars.fig3g <- setdiff(required.vars.fig3g, names(int.full.df))  # Store an intermediate object used in the following analysis step.
if (length(missing.vars.fig3g) > 0) stop("int.full.df is missing variables for Fig3g: ", paste(missing.vars.fig3g, collapse = ", "))  # Run a conditional validation or processing step.

fig3g.df <- int.full.df  # Store an intermediate object used in the following analysis step.
fig3g.df$StudyStartDate <- as.Date(fig3g.df$StudyStartDate)  # Convert values to Date format.
fig3g.df$StudyEndDate <- as.Date(fig3g.df$StudyEndDate)  # Convert values to Date format.
fig3g.df$CA.Date <- as.Date(fig3g.df$CA.Date)  # Convert values to Date format.
fig3g.df$PAD.Date <- as.Date(fig3g.df$PAD.Date)  # Convert values to Date format.
fig3g.df$LAE.Date <- as.Date(fig3g.df$LAE.Date)  # Convert values to Date format.

follow.up.end <- as.Date("2024-05-27")  # Convert values to Date format.

## Use original-style strip labels.
cluster1.strip.text <- paste0("Cluster 1 (n=", length(unique(fig3g.df[fig3g.df$Cluster == "C1 (n=71)", "PatientId"])), ")")  # Create a readable label for plotting or reporting.
cluster2.strip.text <- paste0("Cluster 2 (n=", length(unique(fig3g.df[fig3g.df$Cluster == "C2 (n=47)", "PatientId"])), ")")  # Create a readable label for plotting or reporting.

fig3g.df$ClusterFig <- factor( as.character(fig3g.df$Cluster), levels = c("C1 (n=71)", "C2 (n=47)"), labels = c(cluster1.strip.text, cluster2.strip.text) )  # Standardize factor levels for consistent tables, models, or plots.

fig3g_status_df <- data.frame()  # Create a labelled source-data table for validation/export.
fig3g_events_df <- data.frame()  # Create a labelled source-data table for validation/export.

for (pat.id in unique(fig3g.df$PatientId)) {  # Iterate over patients or model outputs.
  per.pat.int.df <- subset( fig3g.df, PatientId == pat.id, select = c("PatientId", "ClusterFig", "StudyStartDate", "StudyEndDate", "StudyDays", "AF", "CA.Date", "PAD.Date", "LAE.Date") )  # Subset rows or columns for this analysis step.

  if (nrow(per.pat.int.df) > 0) {  # Run a conditional validation or processing step.
    per.pat.summary.df <- per.pat.int.df[, 1:3]  # Store an intermediate object used in the following analysis step.
    per.pat.summary.df$Date <- follow.up.end  # Store an intermediate object used in the following analysis step.
    per.pat.summary.df$Status <- "Tel. follow-up"  # Store an intermediate object used in the following analysis step.

    spdf.end.df <- per.pat.int.df[, 1:3]  # Store an intermediate object used in the following analysis step.
    spdf.end.df$Date <- per.pat.int.df$StudyEndDate[1]  # Store an intermediate object used in the following analysis step.
    spdf.end.df$Status <- "SPDF"  # Store an intermediate object used in the following analysis step.

    per.pat.summary.df <- rbind(per.pat.summary.df, spdf.end.df)  # Store an intermediate object used in the following analysis step.
    per.pat.summary.df <- per.pat.summary.df[order(per.pat.summary.df$Date), ]  # Store an intermediate object used in the following analysis step.

    per.pat.summary.df$Days <- as.numeric(per.pat.summary.df$Date - per.pat.summary.df$StudyStartDate)  # Convert values to numeric format.
    per.pat.summary.df$DaysDiff <- c(per.pat.summary.df$Days[1], diff(per.pat.summary.df$Days))  # Define a compact constant vector used below.
    per.pat.summary.df$Years <- per.pat.summary.df$Days / 365.25  # Store an intermediate object used in the following analysis step.
    per.pat.summary.df$YearsDiff <- c(per.pat.summary.df$Years[1], diff(per.pat.summary.df$Years))  # Define a compact constant vector used below.
    per.pat.summary.df$StatusLayer <- as.factor(seq(nrow(per.pat.summary.df), 1, -1))  # Store an intermediate object used in the following analysis step.

    fig3g_status_df <- rbind(fig3g_status_df, per.pat.summary.df)  # Store an intermediate object used in the following analysis step.

    per.pat.CVD.df <- data.frame()  # Create a labelled source-data table for validation/export.

    if (!is.na(per.pat.int.df$CA.Date[1])) {  # Run a conditional validation or processing step.
      CA.summary <- per.pat.int.df[, 1:3]  # Store an intermediate object used in the following analysis step.
      CA.summary$Date <- per.pat.int.df$CA.Date[1]  # Store an intermediate object used in the following analysis step.
      CA.summary$Event <- "CA"  # Store an intermediate object used in the following analysis step.
      per.pat.CVD.df <- rbind(per.pat.CVD.df, CA.summary)  # Store an intermediate object used in the following analysis step.
    }

    if (!is.na(per.pat.int.df$PAD.Date[1])) {  # Run a conditional validation or processing step.
      PAD.summary <- per.pat.int.df[, 1:3]  # Store an intermediate object used in the following analysis step.
      PAD.summary$Date <- per.pat.int.df$PAD.Date[1]  # Store an intermediate object used in the following analysis step.
      PAD.summary$Event <- "PAD"  # Store an intermediate object used in the following analysis step.
      per.pat.CVD.df <- rbind(per.pat.CVD.df, PAD.summary)  # Store an intermediate object used in the following analysis step.
    }

    if (!is.na(per.pat.int.df$LAE.Date[1])) {  # Run a conditional validation or processing step.
      LAE.summary <- per.pat.int.df[, 1:3]  # Store an intermediate object used in the following analysis step.
      LAE.summary$Date <- per.pat.int.df$LAE.Date[1]  # Store an intermediate object used in the following analysis step.
      LAE.summary$Event <- "LAE"  # Store an intermediate object used in the following analysis step.
      per.pat.CVD.df <- rbind(per.pat.CVD.df, LAE.summary)  # Store an intermediate object used in the following analysis step.
    }

    if (!is.na(per.pat.int.df$AF[1]) && per.pat.int.df$AF[1] == "Yes") {  # Run a conditional validation or processing step.
      AF.summary <- per.pat.int.df[, 1:3]  # Store an intermediate object used in the following analysis step.
      AF.summary$Date <- per.pat.int.df$StudyStartDate[1] - 100  # Store an intermediate object used in the following analysis step.
      AF.summary$Event <- "AF"  # Store an intermediate object used in the following analysis step.
      per.pat.CVD.df <- rbind(per.pat.CVD.df, AF.summary)  # Store an intermediate object used in the following analysis step.
    }

    if (nrow(per.pat.CVD.df) > 0) {  # Run a conditional validation or processing step.
      per.pat.CVD.df$Days <- as.numeric(per.pat.CVD.df$Date - per.pat.CVD.df$StudyStartDate)  # Convert values to numeric format.
      per.pat.CVD.df$Years <- per.pat.CVD.df$Days / 365.25  # Store an intermediate object used in the following analysis step.
      fig3g_events_df <- rbind(fig3g_events_df, per.pat.CVD.df)  # Store an intermediate object used in the following analysis step.
    }
  }
}

fig3g_status_df$Status <- factor(fig3g_status_df$Status, levels = c("SPDF", "Tel. follow-up"))  # Standardize factor levels for consistent tables, models, or plots.
fig3g_events_df$Event <- factor(fig3g_events_df$Event, levels = c("AF", "CA", "LAE", "PAD"))  # Standardize factor levels for consistent tables, models, or plots.

fig3g_bar.fill.colors <- c("#66C1A4", "#FCB2AF")
fig3g_bar.edge.colors <- rep("#ffffff", length(levels(fig3g_status_df$StatusLayer)))
fig3g_title.text <- paste0("Overview of CVD events between clustering analysis groups (n=", nrow(fig3g.df), ")")  # Create a readable label for plotting or reporting.
fig3g_caption.text <- "AF: atrial fibrillation; CA: cerebral apoplexy; LAE: lung arterial embolism; PAD: peripheral artery disease"  # Store an intermediate object used in the following analysis step.

fig3g_fig <- ggplot(data = fig3g_status_df) +  # Initialize the ggplot object for this panel.
  geom_bar(aes(y = PatientId, x = YearsDiff, color = StatusLayer, fill = Status), stat = "identity", width = 0.6, linewidth = 0.01) +
  scale_y_discrete(guide = guide_axis(n.dodge = 2)) +
  scale_color_manual(values = fig3g_bar.edge.colors) +
  scale_fill_manual(values = fig3g_bar.fill.colors) +
  geom_point(data = fig3g_events_df, aes(y = PatientId, x = Years, shape = Event), color = "#000000", fill = "#222222", alpha = 0.6, size = 4) +
  scale_shape_manual(values = c("AF" = 21, "CA" = 22, "LAE" = 23, "PAD" = 24)) +
  facet_wrap(. ~ ClusterFig, scales = "free_y") +
  ylab(label = "Patient") +
  xlab(label = "Time (years)") +
  labs(title = fig3g_title.text, caption = fig3g_caption.text) +
  guides(color = "none") +
  m.top.legend.theme

fig3g_fig
save.fig.to.pptx( ggfig = fig3g_fig, path = fig3.output.ppt.path, append = TRUE, vertical = FALSE, img.width = 10, img.height = 7, save.as.image.in.ppt, ppt_temp_path = ppt.template.path )  # Export figure to the validation PowerPoint file.

fig3g_event_counts <- as.data.frame(table(fig3g_events_df$ClusterFig, fig3g_events_df$Event))  # Convert summary output to a data frame for export.
names(fig3g_event_counts) <- c("Cluster", "Event", "n_events")  # Define a compact constant vector used below.

fig3g_patient_level <- fig3g.df[, c("PatientId", "Cluster", "ClusterFig", "StudyStartDate", "StudyEndDate", "AF", "CA.Date", "LAE.Date", "PAD.Date", "AF.CA.LAE.PAD")]  # Store an intermediate object used in the following analysis step.

fig3g_panel_summary <- data.frame(  # Create a labelled source-data table for validation/export.
  panel = "Fig3g",
  source_object = "fig3g_status_df / fig3g_events_df / fig3g_event_counts / fig3g_patient_level",
  description = "Patient-level SPDF and telephone follow-up timeline by temperature variability cluster, with AF, CA, LAE and PAD event markers.",
  n_status_rows = nrow(fig3g_status_df),
  n_event_rows = nrow(fig3g_events_df),
  n_patient_rows = nrow(fig3g_patient_level),
  stringsAsFactors = FALSE
)

message("Fig3g figure exported to PPT: ", fig3.output.ppt.path)
print(fig3g_panel_summary)
print(fig3g_event_counts)

################################################################################
# 6.3.4 Panel summaries for Fig3e-g
################################################################################

fig3eg_panel_summary <- data.frame(  # Create a labelled source-data table for validation/export.
  panel = c("Fig3e", "Fig3f", "Fig3g"),
  source_object = c(
    "fig3e_source_data / fig3e_summary",
    "fig3f_km_dataset / fig3f_km_curve_source_data / fig3f_logrank_summary",
    "fig3g_status_df / fig3g_events_df / fig3g_event_counts / fig3g_patient_level"
  ),
  description = c(
    "Composite AF/thromboembolic outcome by temperature variability cluster.",
    "Kaplan-Meier event-free survival for first PAD/LAE/CA thromboembolic event by cluster.",
    "Patient-level SPDF and telephone follow-up timeline by temperature variability cluster, with AF, CA, LAE and PAD event markers."
  ),
  n_rows = c( nrow(fig3e_source_data), nrow(fig3f_km_dataset), nrow(fig3g_patient_level) ),
  stringsAsFactors = FALSE
)

message("Fig3e-g figures exported to PPT: ", fig3.output.ppt.path)
print(fig3eg_panel_summary)
print(fig3e_summary)
print(fig3f_logrank_summary)
print(fig3g_panel_summary)
print(fig3g_event_counts)

################################################################################
# 6.4 Export Figure 3 source data for validation: panels a-g
################################################################################

fig3.required.objects <- c("fig3a_source_data", "fig3a_cluster_counts", "fig3b_source_data", "fig3c_boxplot_source_data", "fig3c_scatter_source_data", "fig3d_k_selection_source_data", "fig3d_silhouette_summary", "fig3e_source_data", "fig3e_summary", "fig3f_km_dataset", "fig3f_km_curve_source_data", "fig3f_logrank_summary", "fig3g_status_df", "fig3g_events_df", "fig3g_event_counts", "fig3g_patient_level", "fig3g_panel_summary", "fig3bc_panel_summary", "fig3ad_panel_summary", "fig3eg_panel_summary")  # Figure 3 source-data objects required before workbook export.

fig3.missing.objects <- fig3.required.objects[!sapply(fig3.required.objects, exists)]  # Store an intermediate object used in the following analysis step.
if (length(fig3.missing.objects) > 0) {  # Run a conditional validation or processing step.
  stop("Missing Figure 3 source data objects before export: ", paste(fig3.missing.objects, collapse = ", "))
}

fig3.validation.path <- file.path( figure.output.folder, "Figure_3_source_data_validation_by_panel.xlsx" )  # Build a platform-independent file path.

wb.fig3 <- openxlsx::createWorkbook()  # Store an intermediate object used in the following analysis step.

for (sheet.name in c("Fig3a", "Fig3b", "Fig3c", "Fig3d", "Fig3e", "Fig3f", "Fig3g")) {  # Iterate over patients or model outputs.
  openxlsx::addWorksheet(wb.fig3, sheet.name)
}

header.style <- openxlsx::createStyle(textDecoration = "bold", fontSize = 12)  # Store an intermediate object used in the following analysis step.
subheader.style <- openxlsx::createStyle(textDecoration = "bold", fontSize = 10)  # Store an intermediate object used in the following analysis step.

write_block <- function(wb, sheet, title, x, start_row) {  # Store an intermediate object used in the following analysis step.
  openxlsx::writeData(wb, sheet = sheet, x = title, startRow = start_row, startCol = 1)  # Write a labelled table block to the Excel worksheet.
  openxlsx::addStyle(wb, sheet = sheet, style = subheader.style, rows = start_row, cols = 1, gridExpand = TRUE)
  openxlsx::writeData(wb, sheet = sheet, x = x, startRow = start_row + 1, startCol = 1, rowNames = FALSE)  # Write a labelled table block to the Excel worksheet.
  return(start_row + nrow(x) + 4)
}

################################################################################
# Fig3a worksheet
################################################################################

openxlsx::writeData(wb.fig3, sheet = "Fig3a", x = "Figure 3a source data: temperature variability cluster classification", startRow = 1, startCol = 1)  # Write a labelled table block to the Excel worksheet.
openxlsx::addStyle(wb.fig3, sheet = "Fig3a", style = header.style, rows = 1, cols = 1, gridExpand = TRUE)

row.pos <- 3  # Store an intermediate object used in the following analysis step.
row.pos <- write_block(wb.fig3, "Fig3a", "A1. Panel summary", subset(fig3ad_panel_summary, panel == "Fig3a"), row.pos)  # Store an intermediate object used in the following analysis step.
row.pos <- write_block(wb.fig3, "Fig3a", "A2. Cluster counts", fig3a_cluster_counts, row.pos)  # Store an intermediate object used in the following analysis step.
row.pos <- write_block(wb.fig3, "Fig3a", "A3. Patient-level cluster feature data", fig3a_source_data, row.pos)  # Store an intermediate object used in the following analysis step.

################################################################################
# Fig3b worksheet
################################################################################

openxlsx::writeData(wb.fig3, sheet = "Fig3b", x = "Figure 3b source data: P183 representative MTK temperature recording", startRow = 1, startCol = 1)  # Write a labelled table block to the Excel worksheet.
openxlsx::addStyle(wb.fig3, sheet = "Fig3b", style = header.style, rows = 1, cols = 1, gridExpand = TRUE)

row.pos <- 3  # Store an intermediate object used in the following analysis step.
row.pos <- write_block(wb.fig3, "Fig3b", "B1. Panel summary", subset(fig3bc_panel_summary, panel == "Fig3b"), row.pos)  # Store an intermediate object used in the following analysis step.
row.pos <- write_block(wb.fig3, "Fig3b", "B2. Plotted raw temperature data", fig3b_source_data, row.pos)  # Store an intermediate object used in the following analysis step.

################################################################################
# Fig3c worksheet
################################################################################

openxlsx::writeData(wb.fig3, sheet = "Fig3c", x = "Figure 3c source data: P183 distribution and patient-level MTK feature scatter", startRow = 1, startCol = 1)  # Write a labelled table block to the Excel worksheet.
openxlsx::addStyle(wb.fig3, sheet = "Fig3c", style = header.style, rows = 1, cols = 1, gridExpand = TRUE)

row.pos <- 3  # Store an intermediate object used in the following analysis step.
row.pos <- write_block(wb.fig3, "Fig3c", "C1. Panel summary", subset(fig3bc_panel_summary, panel %in% c("Fig3c left", "Fig3c right")), row.pos)  # Store an intermediate object used in the following analysis step.
row.pos <- write_block(wb.fig3, "Fig3c", "C2. P183 Abs.Diff.MTK distribution source data", fig3c_boxplot_source_data, row.pos)  # Store an intermediate object used in the following analysis step.
row.pos <- write_block(wb.fig3, "Fig3c", "C3. Patient-level MTK feature scatter source data", fig3c_scatter_source_data, row.pos)  # Store an intermediate object used in the following analysis step.

################################################################################
# Fig3d worksheet
################################################################################

openxlsx::writeData(wb.fig3, sheet = "Fig3d", x = "Figure 3d source data: silhouette-based optimal cluster number", startRow = 1, startCol = 1)  # Write a labelled table block to the Excel worksheet.
openxlsx::addStyle(wb.fig3, sheet = "Fig3d", style = header.style, rows = 1, cols = 1, gridExpand = TRUE)

row.pos <- 3  # Store an intermediate object used in the following analysis step.
row.pos <- write_block(wb.fig3, "Fig3d", "D1. Panel summary", subset(fig3ad_panel_summary, panel == "Fig3d"), row.pos)  # Store an intermediate object used in the following analysis step.
row.pos <- write_block(wb.fig3, "Fig3d", "D2. Silhouette summary", fig3d_silhouette_summary, row.pos)  # Store an intermediate object used in the following analysis step.
row.pos <- write_block(wb.fig3, "Fig3d", "D3. Average silhouette width by k", fig3d_k_selection_source_data, row.pos)  # Store an intermediate object used in the following analysis step.

################################################################################
# Fig3e worksheet
################################################################################

openxlsx::writeData(wb.fig3, sheet = "Fig3e", x = "Figure 3e source data: composite AF/thromboembolic outcome by cluster", startRow = 1, startCol = 1)  # Write a labelled table block to the Excel worksheet.
openxlsx::addStyle(wb.fig3, sheet = "Fig3e", style = header.style, rows = 1, cols = 1, gridExpand = TRUE)

row.pos <- 3  # Store an intermediate object used in the following analysis step.
row.pos <- write_block(wb.fig3, "Fig3e", "E1. Panel summary", subset(fig3eg_panel_summary, panel == "Fig3e"), row.pos)  # Store an intermediate object used in the following analysis step.
row.pos <- write_block(wb.fig3, "Fig3e", "E2. Patient-level composite outcome source data", fig3e_source_data, row.pos)  # Store an intermediate object used in the following analysis step.
row.pos <- write_block(wb.fig3, "Fig3e", "E3. Composite outcome counts and percentages", fig3e_summary, row.pos)  # Store an intermediate object used in the following analysis step.

################################################################################
# Fig3f worksheet
################################################################################

openxlsx::writeData(wb.fig3, sheet = "Fig3f", x = "Figure 3f source data: Kaplan-Meier event-free survival by cluster", startRow = 1, startCol = 1)  # Write a labelled table block to the Excel worksheet.
openxlsx::addStyle(wb.fig3, sheet = "Fig3f", style = header.style, rows = 1, cols = 1, gridExpand = TRUE)

row.pos <- 3  # Store an intermediate object used in the following analysis step.
row.pos <- write_block(wb.fig3, "Fig3f", "F1. Panel summary", subset(fig3eg_panel_summary, panel == "Fig3f"), row.pos)  # Store an intermediate object used in the following analysis step.
row.pos <- write_block(wb.fig3, "Fig3f", "F2. Patient-level survival dataset", fig3f_km_dataset, row.pos)  # Store an intermediate object used in the following analysis step.
row.pos <- write_block(wb.fig3, "Fig3f", "F3. Kaplan-Meier curve source data", fig3f_km_curve_source_data, row.pos)  # Store an intermediate object used in the following analysis step.
row.pos <- write_block(wb.fig3, "Fig3f", "F4. Log-rank test summary", fig3f_logrank_summary, row.pos)  # Store an intermediate object used in the following analysis step.

################################################################################
# Fig3g worksheet
################################################################################

fig3g_status_public <- fig3g_status_df[, intersect( c("PatientId", "ClusterFig", "StudyStartDate", "Date", "Status", "Days", "DaysDiff", "Years", "YearsDiff", "StatusLayer"), names(fig3g_status_df) )]  # Store an intermediate object used in the following analysis step.

fig3g_events_public <- fig3g_events_df[, intersect( c("PatientId", "ClusterFig", "StudyStartDate", "Date", "Event", "Days", "Years"), names(fig3g_events_df) )]  # Store an intermediate object used in the following analysis step.

fig3g_patient_public <- fig3g_patient_level[, intersect(  # Store an intermediate object used in the following analysis step.
  c("PatientId", "Cluster", "ClusterFig", "StudyStartDate", "StudyEndDate", "AF", "CA.Date", "LAE.Date", "PAD.Date", "AF.CA.LAE.PAD"),
  names(fig3g_patient_level)
)]

openxlsx::writeData(wb.fig3, sheet = "Fig3g", x = "Figure 3g source data: patient-level timeline and CVD event plotting data by cluster", startRow = 1, startCol = 1)  # Write a labelled table block to the Excel worksheet.
openxlsx::addStyle(wb.fig3, sheet = "Fig3g", style = header.style, rows = 1, cols = 1, gridExpand = TRUE)

row.pos <- 3  # Store an intermediate object used in the following analysis step.
row.pos <- write_block(wb.fig3, "Fig3g", "G1. Panel summary", fig3g_panel_summary, row.pos)  # Store an intermediate object used in the following analysis step.
row.pos <- write_block(wb.fig3, "Fig3g", "G2. Timeline status segments used for plotted bars", fig3g_status_public, row.pos)  # Store an intermediate object used in the following analysis step.
row.pos <- write_block(wb.fig3, "Fig3g", "G3. CVD event points used for plotted markers", fig3g_events_public, row.pos)  # Store an intermediate object used in the following analysis step.
row.pos <- write_block(wb.fig3, "Fig3g", "G4. Event counts by cluster and event type", fig3g_event_counts, row.pos)  # Store an intermediate object used in the following analysis step.
row.pos <- write_block(wb.fig3, "Fig3g", "G5. Minimal patient-level plotting data", fig3g_patient_public, row.pos)  # Store an intermediate object used in the following analysis step.

################################################################################
# Formatting and save
################################################################################

for (sheet.name in c("Fig3a", "Fig3b", "Fig3c", "Fig3d", "Fig3e", "Fig3f", "Fig3g")) {  # Iterate over patients or model outputs.
  openxlsx::setColWidths(wb.fig3, sheet = sheet.name, cols = 1:40, widths = "auto")
  openxlsx::freezePane(wb.fig3, sheet = sheet.name, firstRow = TRUE)
}

openxlsx::saveWorkbook(wb.fig3, file = fig3.validation.path, overwrite = TRUE)  # Save the assembled Excel workbook.

message("Figure 3 source data validation file exported to:")
message(fig3.validation.path)
message("Figure 3 source data export completed for panels a-g.")

################################################################################
# 7. Figure 4 source data and figure generation
################################################################################

## Figure 4:
##   PIRI symmetric vs asymmetric pattern and composite AF/thromboembolic outcome.
##
## Main comparison:
##   Symmetric:    5 / 81 patients with AF/thromboembolic event
##   Asymmetric: 18 / 37 patients with AF/thromboembolic event
##   Fisher's exact test p < 0.0001

required.objects.fig4 <- c("int.full.df", "output.folder", "figure.output.folder")  # objects required before Figure 4 generation.
missing.objects.fig4 <- required.objects.fig4[!sapply(required.objects.fig4, exists)]  # Store an intermediate object used in the following analysis step.
if (length(missing.objects.fig4) > 0) {  # Run a conditional validation or processing step.
  stop("Missing objects before Figure 4: ", paste(missing.objects.fig4, collapse = ", "))
}

required.vars.fig4 <- c("PatientId", "PIRI.Group", "AF", "CA", "LAE", "PAD", "AF.CA.LAE.PAD" )  # intervention-cohort columns required for Figure 4.

missing.vars.fig4 <- setdiff(required.vars.fig4, names(int.full.df))  # Store an intermediate object used in the following analysis step.
if (length(missing.vars.fig4) > 0) {  # Run a conditional validation or processing step.
  stop("int.full.df is missing variables for Figure 4: ", paste(missing.vars.fig4, collapse = ", "))
}

################################################################################
# 7.1 Prepare Figure 4 patient-level source data
################################################################################

fig4_source_data <- int.full.df[, required.vars.fig4]  # Store an intermediate object used in the following analysis step.

fig4_source_data$PIRI.Group <- factor( fig4_source_data$PIRI.Group, levels = c("Symmetric", "Asymmetric") )  # Standardize factor levels for consistent tables, models, or plots.

fig4_source_data$AF.CA.LAE.PAD <- factor( fig4_source_data$AF.CA.LAE.PAD, levels = c("No", "Yes") )  # Standardize factor levels for consistent tables, models, or plots.

fig4_source_data <- fig4_source_data[complete.cases(fig4_source_data[, c("PIRI.Group", "AF.CA.LAE.PAD")]), ]  # Apply complete-case filtering for the relevant variables.

message("Figure 4 PIRI group counts:")
print(table(fig4_source_data$PIRI.Group, useNA = "ifany"))

message("Figure 4 composite outcome counts by PIRI group:")
print(table(fig4_source_data$PIRI.Group, fig4_source_data$AF.CA.LAE.PAD, useNA = "ifany"))

################################################################################
# 7.2 Summarize composite outcome by PIRI group
################################################################################

fig4_counts <- as.data.frame( table(fig4_source_data$PIRI.Group, fig4_source_data$AF.CA.LAE.PAD) )  # Convert summary output to a data frame for export.

names(fig4_counts) <- c("PIRI.Group", "CompositeOutcome", "n_patients")  # Define a compact constant vector used below.

fig4_group_n <- as.data.frame(table(fig4_source_data$PIRI.Group))  # Convert summary output to a data frame for export.
names(fig4_group_n) <- c("PIRI.Group", "group_n")  # Define a compact constant vector used below.

fig4_counts <- merge( fig4_counts, fig4_group_n, by = "PIRI.Group", all.x = TRUE )  # Merge source tables using their shared patient identifier.

fig4_counts$percentage <- round( 100 * fig4_counts$n_patients / fig4_counts$group_n, 1 )  # Round numeric values for reporting.

fig4_plot_data <- subset(fig4_counts, CompositeOutcome == "Yes")  # Subset rows or columns for this analysis step.

## Sanity check against manuscript counts.
fig4_expected <- data.frame( PIRI.Group = c("Symmetric", "Asymmetric"), expected_event_n = c(5, 18), expected_group_n = c(81, 37), stringsAsFactors = FALSE )  # Create a labelled source-data table for validation/export.

fig4_observed_check <- merge( fig4_expected, fig4_plot_data[, c("PIRI.Group", "n_patients", "group_n", "percentage")], by = "PIRI.Group", all.x = TRUE )  # Merge source tables using their shared patient identifier.

names(fig4_observed_check)[names(fig4_observed_check) == "n_patients"] <- "observed_event_n"  # Update selected rows/columns according to the analysis definition.
names(fig4_observed_check)[names(fig4_observed_check) == "group_n"] <- "observed_group_n"  # Update selected rows/columns according to the analysis definition.

if ( any(fig4_observed_check$expected_event_n != fig4_observed_check$observed_event_n) || any(fig4_observed_check$expected_group_n != fig4_observed_check$observed_group_n) ) {  # Run a conditional validation or processing step.
  warning("Figure 4 observed PIRI/outcome counts differ from expected manuscript counts.")
  print(fig4_observed_check)
}

################################################################################
# 7.3 Fisher's exact test
################################################################################

fig4_contingency_table <- table( fig4_source_data$PIRI.Group, fig4_source_data$AF.CA.LAE.PAD )  # Tabulate counts for quality control or plotting.

fig4_fisher_test <- fisher.test(fig4_contingency_table)  # Run Fisher’s exact test for group comparison.

fig4_p_value <- fig4_fisher_test$p.value  # Store an intermediate object used in the following analysis step.
fig4_p_label <- ifelse( fig4_p_value < 0.0001, "Fisher's exact p < 0.0001", paste0("Fisher's exact p = ", sprintf("%.4f", fig4_p_value)) )  # Store an intermediate object used in the following analysis step.

fig4_test_summary <- data.frame(  # Create a labelled source-data table for validation/export.
  comparison = "PIRI asymmetric vs symmetric",
  outcome = "Composite AF/thromboembolic outcome",
  test = "Fisher's exact test",
  p_value = fig4_p_value,
  p_value_label = fig4_p_label,
  odds_ratio = unname(fig4_fisher_test$estimate),
  conf_low = fig4_fisher_test$conf.int[1],
  conf_high = fig4_fisher_test$conf.int[2],
  symmetric_events = fig4_plot_data[fig4_plot_data$PIRI.Group == "Symmetric", "n_patients"],
  symmetric_total = fig4_plot_data[fig4_plot_data$PIRI.Group == "Symmetric", "group_n"],
  asymmetric_events = fig4_plot_data[fig4_plot_data$PIRI.Group == "Asymmetric", "n_patients"],
  asymmetric_total = fig4_plot_data[fig4_plot_data$PIRI.Group == "Asymmetric", "group_n"],
  stringsAsFactors = FALSE
)

message("Figure 4 Fisher's exact test:")
print(fig4_test_summary)

################################################################################
# 7.4 Generate Figure 4
################################################################################

fig4_fig <- ggplot(  # Initialize the ggplot object for this panel.
  data = fig4_plot_data,
  aes(x = PIRI.Group, y = percentage, fill = PIRI.Group)
) +
  geom_bar(
    stat = "identity",
    width = 0.55,
    color = "black",
    linewidth = 0.2
  ) +
  geom_text(
    aes(label = paste0(n_patients, "/", group_n, "\n", percentage, "%")),
    vjust = -0.4,
    size = 3
  ) +
  labs(
    title = "Composite AF/thromboembolic outcome by PIRI group",
    subtitle = fig4_p_label,
    x = NULL,
    y = "Patients with event (%)"
  ) +
  scale_y_continuous(
    limits = c(0, max(fig4_plot_data$percentage, na.rm = TRUE) * 1.25),
    breaks = seq(0, 60, 10)
  ) +
  scale_fill_manual(values = c("Symmetric" = "#66C1A4", "Asymmetric" = "#FCB2AF")) +
  guides(fill = "none") +
  m.normal.theme

fig4_fig

save.fig.to.pptx( ggfig = fig4_fig, path = fig4.output.ppt.path, append = TRUE, vertical = FALSE, img.width = 3.5, img.height = 3.2, save.as.image.in.ppt, ppt_temp_path = ppt.template.path )  # Export figure to the validation PowerPoint file.

message("Figure 4 exported to PPT:")
message(fig4.output.ppt.path)

################################################################################
# 7.5 Export Figure 4 source data for validation
################################################################################

fig4.validation.path <- file.path( figure.output.folder, "Figure_4_source_data_validation.xlsx" )  # Build a platform-independent file path.

fig4_panel_summary <- data.frame(  # Create a labelled source-data table for validation/export.
  panel = "Fig4",
  source_object = "fig4_source_data / fig4_counts / fig4_test_summary",
  description = "Composite AF/thromboembolic outcome by PIRI symmetric/asymmetric group.",
  n_patient_rows = nrow(fig4_source_data),
  n_summary_rows = nrow(fig4_counts),
  stringsAsFactors = FALSE
)

wb.fig4 <- openxlsx::createWorkbook()  # Store an intermediate object used in the following analysis step.
openxlsx::addWorksheet(wb.fig4, "Fig4")

header.style <- openxlsx::createStyle(textDecoration = "bold", fontSize = 12)  # Store an intermediate object used in the following analysis step.
subheader.style <- openxlsx::createStyle(textDecoration = "bold", fontSize = 10)  # Store an intermediate object used in the following analysis step.

write_block <- function(wb, sheet, title, x, start_row) {  # Store an intermediate object used in the following analysis step.
  openxlsx::writeData(wb, sheet = sheet, x = title, startRow = start_row, startCol = 1)  # Write a labelled table block to the Excel worksheet.
  openxlsx::addStyle(wb, sheet = sheet, style = subheader.style, rows = start_row, cols = 1, gridExpand = TRUE)
  openxlsx::writeData(wb, sheet = sheet, x = x, startRow = start_row + 1, startCol = 1, rowNames = FALSE)  # Write a labelled table block to the Excel worksheet.
  return(start_row + nrow(x) + 4)
}

openxlsx::writeData( wb.fig4, sheet = "Fig4", x = "Figure 4 source data: PIRI group and composite AF/thromboembolic outcome", startRow = 1, startCol = 1 )  # Write a labelled table block to the Excel worksheet.

openxlsx::addStyle( wb.fig4, sheet = "Fig4", style = header.style, rows = 1, cols = 1, gridExpand = TRUE )

row.pos <- 3  # Store an intermediate object used in the following analysis step.
row.pos <- write_block(wb.fig4, "Fig4", "A1. Panel summary", fig4_panel_summary, row.pos)  # Store an intermediate object used in the following analysis step.
row.pos <- write_block(wb.fig4, "Fig4", "A2. Patient-level plotting source data", fig4_source_data, row.pos)  # Store an intermediate object used in the following analysis step.
row.pos <- write_block(wb.fig4, "Fig4", "A3. Composite outcome counts and percentages", fig4_counts, row.pos)  # Store an intermediate object used in the following analysis step.
row.pos <- write_block(wb.fig4, "Fig4", "A4. Fisher's exact test summary", fig4_test_summary, row.pos)  # Store an intermediate object used in the following analysis step.
row.pos <- write_block(wb.fig4, "Fig4", "A5. Expected vs observed manuscript count check", fig4_observed_check, row.pos)  # Store an intermediate object used in the following analysis step.

openxlsx::setColWidths(wb.fig4, sheet = "Fig4", cols = 1:40, widths = "auto")
openxlsx::freezePane(wb.fig4, sheet = "Fig4", firstRow = TRUE)

openxlsx::saveWorkbook( wb.fig4, file = fig4.validation.path, overwrite = TRUE )  # Save the assembled Excel workbook.

message("Figure 4 source data validation file exported to:")
message(fig4.validation.path)
message("Figure 4 source data export completed.")

################################################################################
# 8. Final quality-control summary
################################################################################

message("Final QC summary")
message("Full follow-up cohort n = ", nrow(full.df))
message("Intervention analysis cohort n = ", nrow(int.full.df))

message("Randomized group counts:")
print(table(full.df$Group, useNA = "ifany"))

message("PTD / post-hoc temperature group counts:")
print(table(int.full.df$PostHocGroup, useNA = "ifany"))

message("Temperature variability cluster counts:")
print(table(int.full.df$Cluster, useNA = "ifany"))

message("PIRI group counts:")
print(table(int.full.df$PIRI.Group, useNA = "ifany"))

message("Composite AF/thromboembolic outcome in intervention cohort:")
print(table(int.full.df$AF.CA.LAE.PAD, useNA = "ifany"))

message("Figure 4 PIRI-by-outcome table:")
print(table(fig4_source_data$PIRI.Group, fig4_source_data$AF.CA.LAE.PAD, useNA = "ifany"))

message("Figure 4 Fisher test:")
print(fig4_test_summary[, c("comparison", "p_value", "p_value_label", "odds_ratio", "conf_low", "conf_high")])

message("Table output folder:")
message(table.output.folder)

message("Figure source data output folder:")
message(figure.output.folder)

message("Session information:")
print(sessionInfo())

message("Session information:")
print(sessionInfo())
message("Main analysis script completed successfully.")
