### =====================================================================
### Supplementary Code (Annotated)
### Study: Temperature sensor–derived patient clustering & CVD associations
### File: cluster_temp_by_pat_v1 (ORIGINAL CODE, LOGIC UNCHANGED)
### Author: Antao Ming | Contact: antao.ming@wmu.edu.cn
### Date: 2025-09-14
### ---------------------------------------------------------------------
### Purpose of this script
###   • Merge temperature-sensor data with patient-level information
###     (including cardiovascular outcomes) for the intervention group.
###   • Engineer temperature features (e.g., absolute MTK asymmetry) at the
###     patient level.
###   • Cluster patients based on these temperature features.
###   • Explore associations between the resulting clusters and CVD events
###     using logistic regression models.
###
### Notes for editors/reviewers
###   • This is the authors' ORIGINAL analysis code with added comments only.
###     No functional edits have been made (variable names, paths, and logic
###     are preserved exactly as provided by the authors).
###   • The code relies on several custom helper scripts (sourced below) that
###     define functions such as theming for figures; these are kept as-is.
###   • The script uses absolute Windows paths (C:/...). These are left
###     unchanged to reflect the authors' working environment.
###   • Randomness: a fixed seed (42) is used for k-means to make the cluster
###     assignment reproducible.
###   • Plots of cluster diagnostics (Elbow, Silhouette) and clustering
###     results are exported to a PowerPoint deck via `officer`/`rvg` (and
###     helper `save.fig.to.pptx`).
###
### Input expectations (as used by this script)
###   • Temperature CSV (per-row readings) with, at minimum:
###       PatientId, Timestamp, Date, R.Cal, R.Lat, R.MTK, L.Cal, L.Lat,
###       L.MTK, Diff.MTK
###   • Patient info Excel with, at minimum:
###       PatientId, Group (Intervention/Control), AF, LAE, CA, PAD,
###       AF.CA.LAE.PAD, Age, Sex, BMI, DMDuration, ScrABIRightValue,
###       ScrABILeftValue, StudyStartDate, StudyEndDate, FollowupEndDate
###
### Outputs (as produced by the original code)
###   • Augmented temperature CSV exported to the output folder
###   • A PPTX file containing figures (elbow, silhouette, cluster plots)
###   • Excel files with cluster-level summaries and patient cluster mapping
###   • Console output for model summaries and exponentiated coefficients
###
### Reproducibility & limitations
###   • Dependencies are loaded via sourced custom scripts (see below) and
###     explicit libraries (`officer`, `rvg`, etc.).
###   • ABI full dataset is read early (full.ABI.df) but not used later; kept
###     intact since this is a verbatim workflow.
###   • Confidence intervals use `confint(glm)` defaults (profile likelihood);
###     heavy models may take time or warn. This behavior remains unchanged.
### =====================================================================

### This script initially merged temperature sensor recordings and patient info and CVD events/outcomes, thereafter extract temperature features for each patient and
### clusters the patients into subgroups, finally explores the association between clusters and cardiovarscular events 

Script.Name<-c("cluster_temp_by_pat_v1")   # Script label used when constructing output filenames

#### -------- define global variables -----------------------------------------------------------------
# Root directory of the project workspace (machine-specific absolute path)
workspace.root <- 'C:/local_work_ming/workspaces/r_workspace/smart_prevent_diabetic_feet_study_V1'
# Local working folder for intermediate artifacts
local.workspace <- paste(workspace.root, "local_workspace", sep = "/")
# Central datasets folder (not all items are necessarily used in this script)
datasets.folder <- paste(workspace.root, "datasets", sep = "/")
# Folder containing custom helper scripts used throughout the project
functional.scripts.folder <- paste(workspace.root, "scripts", "my_functional_scripts", sep = "/")

## iq-trial data save path
# Row-level temperature recordings aggregated for clustering
temp.recordings.path <- "C:/local_work_ming/workspaces/r_workspace/smart_prevent_diabetic_feet_study_V1/scripts/follow_up_study/nature cardiovascular research/temperature recordings for clustering.csv"

## information and CVD events of the intervention group
# Patient-level table (cohort overview) including outcomes and covariates
patient.info.path ="C:/local_work_ming/workspaces/r_workspace/smart_prevent_diabetic_feet_study_V1/scripts/follow_up_study/nature cardiovascular research/cohort full overview.xlsx"

## output folder
# Output directory for figures, tables, and exported data. Created if missing.
output.folder <- paste(workspace.root, "outcomes", "follow-up study", "NCR", sep = "/")
if (!dir.exists(output.folder)){
  dir.create(output.folder, showWarnings = T)
}

# PowerPoint path (figures appended here during the run)
output.ppt.path <- paste(output.folder, paste(Script.Name, Sys.Date(), ".pptx"), sep = "/")

# If TRUE, figures are embedded as images in PPTX; otherwise vector objects
save.as.image <- TRUE

#### -------- load scripts and data -----------------------------------------------------------------
## attach functional scripts
# Custom plotting, spatial/data helpers, and global libraries/functions.
# These scripts are part of the authors' internal toolbox and define
# functions referenced later (e.g., save.fig.to.pptx, build.df.summary,
# ggplot themes). They may also attach required CRAN packages.
source("C:/local_work_ming/workspaces/r_workspace/R_customized_global_functions/my_functional_scripts/my_libraries.R")
source("C:/local_work_ming/workspaces/r_workspace/R_customized_global_functions/my_functional_scripts/my_global_functions.R")
source("C:/local_work_ming/workspaces/r_workspace/R_customized_global_functions/my_functional_scripts/my_ggplot_themes.R")
library(officer)   # for creating and writing to PowerPoint files
library(rvg)       # for vector graphics support in PPTX

#### -------- load study participation days table to normalize the observation time  -----------------------------------------------------------------
# Read patient-level cohort overview (Excel), sort by PatientId, and restrict
# to the intervention group (analysis target). Then coerce variables into
# factor/numeric/date types expected by downstream analyses.
pat.df <- read.xlsx(file=patient.info.path, sheetIndex = 1, header=TRUE)
pat.df <- pat.df[order(pat.df$PatientId), ]
pat.df <- subset(pat.df, Group == "Intervention")
head(pat.df)

## format variables
# Binary endpoints and composite are encoded as factors with explicit levels
pat.df$AF <- factor(pat.df$AF , levels = c("No", "Yes"))
pat.df$LAE <- factor(pat.df$LAE , levels = c("No", "Yes"))
pat.df$CA <- factor(pat.df$CA , levels = c("No", "Yes"))
pat.df$PAD <- factor(pat.df$PAD , levels = c("No", "Yes"))

pat.df$AF.CA.LAE.PAD <- factor(pat.df$AF.CA.LAE.PAD , levels = c("No", "Yes"))

# Demographics/covariates
default.na.suppress <- suppressWarnings   # alias used below only for comments clarity
pat.df$Age <- as.numeric(pat.df$Age)
pat.df$Sex <- factor(pat.df$Sex, levels = c("Male", "Female"))

# Study window variables coerced to Date for potential normalization
pat.df$StudyStartDate <- as.Date(pat.df$StudyStartDate)
pat.df$StudyEndDate <- as.Date(pat.df$StudyEndDate)
pat.df$FollowupEndDate <- as.Date(pat.df$FollowupEndDate)


#### -------- load temperature data  -----------------------------------------------------------------
# Row-level temperature recordings (CSV). Each row is a timestamped reading.
final.temp.df <- read.csv(temp.recordings.path, header=TRUE, sep=',')

#### -------- Genarate more temperature parameters -----------------------------------------------------------------
# Sort by patient and time; create row index; compute derived contrasts and
# absolute asymmetry on MTK channel. Also ensure Date is of class Date.
final.temp.df <- final.temp.df[order(final.temp.df$PatientId, final.temp.df$Timestamp), ]
final.temp.df$Index <- seq(1, nrow(final.temp.df), 1)
final.temp.df$Abs.Diff.MTK <- abs(final.temp.df$Diff.MTK)
final.temp.df$R.Cal.Lat.Diff <- final.temp.df$R.Cal - final.temp.df$R.Lat
final.temp.df$R.Cal.MTK.Diff <- final.temp.df$R.Cal - final.temp.df$R.MTK
final.temp.df$R.Lat.MTK.Diff <- final.temp.df$R.Lat - final.temp.df$R.MTK
final.temp.df$L.Cal.Lat.Diff <- final.temp.df$L.Cal - final.temp.df$L.Lat
final.temp.df$L.Cal.MTK.Diff <- final.temp.df$L.Cal - final.temp.df$L.MTK
final.temp.df$L.Lat.MTK.Diff <- final.temp.df$L.Lat - final.temp.df$L.MTK
final.temp.df$Date <- as.Date(final.temp.df$Date)

# Export the augmented row-level dataset for auditability
write.csv(final.temp.df, file = paste(output.folder, paste("temp for clustering (n =", nrow(final.temp.df), ").CSV"), sep = "/"), row.names = F)

#### -------- visualize the events of patients again based on temperature recordings -----------------------------------------------------------------
# Quick check of the first few rows; re-order by PatientId and Date for any
# subsequent per-day inspection or plotting the user may add later.
head(final.temp.df)
# Ensure data is sorted by PatientId and Date
final.temp.df <- final.temp.df[order(final.temp.df$PatientId, final.temp.df$Date), ]

#### -------- extract temp features for each patient -----------------------------------------------------------------
# For each patient, compute per-subject summary features:
#   - N.Measures: number of unique timestamps (total readings)
#   - N.Measured.Days: number of unique observation days
#   - Start/End: first and last recording dates
#   - MTK.Asym.Max: maximum absolute asymmetry on MTK
#   - MTK.Asym.Mean: mean absolute asymmetry on MTK
final.temp.df <- final.temp.df[order(final.temp.df$PatientId, final.temp.df$Timestamp), ]
final.temp.df$PatientId <- factor(final.temp.df$PatientId)
head(final.temp.df)
names(final.temp.df)

all.pat.features <- data.frame()
for (n.pat in unique(final.temp.df$PatientId)) {
  pat.temp.df <- subset(final.temp.df, PatientId==n.pat)
  pat.temp.df <- pat.temp.df[order(pat.temp.df$Timestamp), ]
  pat.temp.df$MeasureNo <- seq(1, nrow(pat.temp.df), 1)
  N.Measures <- length(unique(pat.temp.df$Timestamp))
  N.Measured.Days <- length(unique(pat.temp.df$Date))
  Start.Date <- min(as.Date(pat.temp.df$Date))
  End.Date <- max(as.Date(pat.temp.df$Date))
  
  MTK.Asym.Max <- max(pat.temp.df$Abs.Diff.MTK)
  MTK.Asym.Mean <- mean(pat.temp.df$Abs.Diff.MTK)
  
  pat.features <- data.frame(PatientId=n.pat, Rec.Start.Date=Start.Date, Rec.End.Date=End.Date, N.Rec=N.Measures, Rec.Days=N.Measured.Days, MTK.Asym.Max, MTK.Asym.Mean)
  
  all.pat.features <- rbind(all.pat.features, pat.features)
}


# Merge patient-level features with cohort info (restrict to matched IDs)
full.df <- merge(all.pat.features, pat.df, by="PatientId", all = F)

#### -------- prepare for clustering methods -----------------------------------------------------------------
# Define outcome variables (binary endpoints) used for descriptive summaries
outcome.vars <- c("AF", "PAD", "AF.CA.LAE.PAD")

# Define covariates of interest for later descriptive tables/plots
covariate.vars <- c("Age", "Sex")

## clustering type 1: absolute temp features
# Feature set used for clustering (absolute MTK asymmetry)
cluster.vars1 <- c("MTK.Asym.Max", "MTK.Asym.Mean")
# Dataset handed to clustering: features + identifiers + outcomes + covariates
temp.clustering.df1 <- subset(full.df, select=c(cluster.vars1, "PatientId", outcome.vars, covariate.vars))


#### -------- start clustering methods -----------------------------------------------------------------
# Choose the dataset and feature list for clustering
cluster.df <- temp.clustering.df1
cluster.features <- cluster.vars1
length(unique(cluster.df$PatientId))   # Sanity check: number of unique patients

# Standardize features to zero mean / unit variance prior to k-means
data_scaled <- scale(cluster.df[, c(1:length(cluster.vars1))])

#### ---- K-means clustering -------------
# Load required libraries (expected to be available via sourced scripts or installed in the environment)
library(dplyr)
library(cluster)
library(broom)
library(factoextra)

# NOTE: `fviz_nbclust` and `fviz_cluster` below require the `factoextra` package.
# These are assumed to be attached by the user's environment or helper scripts.

# Step 1: Determine the optimal number of clusters (k) for K-Means using Elbow Method and Silhouette Method
# 1.1. Elbow Method — plots total within-cluster sum of squares (WSS) vs k
elbow.fig <- fviz_nbclust(data_scaled, kmeans, method = "wss") +
  #geom_vline(xintercept = 6, linetype = 2) +  # Example reference line (disabled by default)
  labs(subtitle = "Elbow Method")
print(elbow.fig)

# 1.2. Silhouette Method — average silhouette width vs k (higher is better)
Silhouette.fig <- fviz_nbclust(data_scaled, kmeans, method = "silhouette") +
  labs(subtitle = "Silhouette Method")
print(Silhouette.fig)

# Run K-means with a fixed k (set below)
optimal_k <- 2                   # Chosen number of clusters for downstream analysis
set.seed(42) # ensure reproducible initialization
kmeans_result <- kmeans(data_scaled, centers = optimal_k, nstart = 100)

# Visualize K-Means Clustering (feature space projection)
k.means.result.fig <- fviz_cluster(kmeans_result, data = data_scaled, geom = "point") +
  labs(title = paste("K-Means Clustering with k =", optimal_k))
print(k.means.result.fig)

# Attach cluster labels back to the analysis dataframe
cluster.df$Kmeans.Cluster <- as.factor(kmeans_result$cluster)
# Create labeled frequencies for clearer legends and reporting
freq.df <- as.data.frame(table(cluster.df$Kmeans.Cluster))
freq.df$Label <- paste("C", freq.df$Var1, " (n=", freq.df$Freq, ")", sep = "")
cluster.df$Kmeans.Cluster <- factor(cluster.df$Kmeans.Cluster, levels = freq.df$Var1, labels = freq.df$Label)
table(cluster.df$Kmeans.Cluster)

## visualize the temp features by obtained cluster
# Scatter of MTK asymmetry features colored by cluster (no legend theme is
# applied via a custom theme `m.no.legend.theme` sourced earlier)
post.clustering.mtk <- 
  ggplot(data = cluster.df, aes(y=MTK.Asym.Max, x=MTK.Asym.Mean, color=Kmeans.Cluster))+
  geom_point(size=0.6)+
  labs(title = "Clusters of SeCoSe Sensor MTK")+
  m.no.legend.theme
post.clustering.mtk
save.fig.to.pptx(ggfig = post.clustering.mtk, path=output.ppt.path, append=T,  vertical = F, img.width = 2.5, img.height = 2.5, as.image=save.as.image)


# Overlay CVD composite outcome (AF or LAE or CA or PAD):
#   • shape/size encode Yes vs No, legend moved to top via `m.top.legend.theme`.
post.clustering.mtk.CVD.fig <- 
  ggplot(data = cluster.df, aes(y=MTK.Asym.Max, x=MTK.Asym.Mean, color=Kmeans.Cluster, shape=AF.CA.LAE.PAD, size=AF.CA.LAE.PAD))+
  geom_point(alpha=0.6)+
  scale_shape_manual(values = c(1, 19))+ 
  scale_size_manual(values = c(0.8, 2))+ 
  labs(title = "Clusters of SeCoSe Sensor MTK CVD", size="AF or LAE or CA or PAD", shape="AF or LAE or CA or PAD")+
  m.top.legend.theme
post.clustering.mtk.CVD.fig


#### ---- export clustering results to local file --------------------------------------------------------
# Keep a compact PatientId ↔ Cluster mapping for downstream analyses
pat.df.cluster <- subset(cluster.df, select=c("PatientId", "Kmeans.Cluster"))
pat.df.cluster$Cluster <- pat.df.cluster$Kmeans.Cluster

write.xlsx(pat.df.cluster, file = paste(output.folder, "clustering results.xlsx", sep = "/"),
           sheetName = "List", append = F, row.names = FALSE)


#### ---- Association between clusters and clinical events --------------------------------------------------------

# Logistic regressions assessing association between Cluster and the composite
# outcome (AF.CA.LAE.PAD), with/without adjustment for standard covariates and
# ABI. Exponentiated coefficients (ORs) and 95% CIs are printed to console.
log.df <- subset(pat.df, select=c("PatientId", "Sex", "Age", "BMI", "DMDuration", "ScrABIRightValue", "ScrABILeftValue", "AF.CA.LAE.PAD"))
log.df <- merge(pat.df.cluster[, c("PatientId", "Cluster")], log.df, by="PatientId", all = F)
head(log.df)

# Model 1: Cluster-only
log.model1 <- glm(AF.CA.LAE.PAD~Cluster, family = binomial(link = "logit"), data = log.df)
summary(log.model1)
exp(coef(log.model1))            # Odds ratios
exp(confint(log.model1))         # 95% CIs for ORs (profile likelihood)

# Model 2: Adjusted for age, sex, BMI, DM duration + Cluster
log.model2 <- glm(AF.CA.LAE.PAD~Age + Sex + BMI + DMDuration + Cluster, family = binomial(link = "logit"), data = log.df)
summary(log.model2)
exp(coef(log.model2))
exp(confint(log.model2))

# Model 3: Clinical covariates only (no Cluster), incl. ABI (right & left)
log.model3 <- glm(AF.CA.LAE.PAD~Age + Sex + BMI + DMDuration + ScrABIRightValue + ScrABILeftValue, family = binomial(link = "logit"), data = log.df)
summary(log.model3)
exp(coef(log.model3))
exp(confint(log.model3))

# Model 4: Full model = Model 3 + Cluster
log.model4 <- glm(AF.CA.LAE.PAD~Age + Sex + BMI + DMDuration + ScrABIRightValue + ScrABILeftValue + Cluster, family = binomial(link = "logit"), data = log.df)
summary(log.model4)
exp(coef(log.model4))
exp(confint(log.model4))

