### get the color hue (code in HEX) according to the provided number of categories
gg_color_hue <- function(n) {
  library(scales)
  show_col(hue_pal()(n))
  hues = seq(15, 375, length = n + 1)
  return(hcl(h = hues, l = 65, c = 100)[1:n])
}


## a function to check and set the workspace.root
check.wd <- function(){
  current.workspace.root <- getwd()
  if (current.workspace.root != workspace.root){
    setwd(workspace.root)
    current.workspace.root <- getwd()
    print("change workspace.root to: ")
    current.workspace.root
  }
}

detach_package <- function(pkg, character.only = FALSE)
{
  if(!character.only)
  {
    pkg <- deparse(substitute(pkg))
  }
  search_item <- paste("package", pkg, sep = ":")
  while(search_item %in% search())
  {
    detach(search_item, unload = TRUE, character.only = TRUE)
  }
}

setup.wd <- function(wd){
  current.workspace.root <- getwd()
  if (current.workspace.root != wd){
    setwd(wd)
    current.workspace.root <- getwd()
    print("change workspace.root to: ")
    wd
  }
}


logout.model.score <- function(cross.table){
  if (nrow(cross.table) != 2 || ncol(cross.table) != 2){
    print("invalid inputed cross table")
    return()
  }
  
  tn <- cross.table[1, 1]
  fp <- cross.table[1, 2]
  fn <- cross.table[2, 1]
  tp <- cross.table[2, 2]
  
  accuracy <- (tp + tn)/ (tn + fp + fn + tp)
  specificity <- tn/(tn + fp)
  sensitivity <- tp/(tp + fn)
  print("CrossTable")
  print(cross.table) 
  print("accuracy")
  print(accuracy)
  print("sensitivity")
  print(sensitivity)
  print("specificity")
  print(specificity)
}


## IO function: save a table locally
save.csv <- function(table, name, path){
  csv.file.path <- paste(path, name, sep = "/")
  write.csv(table, csv.file.path, row.names = FALSE)
}

## IO function: save a table locally
save.table <- function(table, name, path){
  csv.file.path <- paste(path, name, sep = "/")
  write.csv(table, csv.file.path, row.names = FALSE)
}

## IO function: save a table locally
save.xlsx <- function(table, name, path, append.save){
  csv.file.path <- paste(path, name, sep = "/")
  write.table(table, csv.file.path, append = append.save, row.names = FALSE)
}

## convert PN levels (schwere, maessige, leichte, or keine) to numeric values (3, 2, 1, or 0)
numeric.PN <- function(PN_level){
  if (PN_level == "schwere"){
    return(3) 
  } else  if (PN_level == "maessige"){
    return(2) 
  }  else  if (PN_level == "leichte"){
    return(1) 
  } else {
    return(0)
  }
}

## convert Gender levels (w, m) to numeric values (1 or 2)
numeric.gender <- function(Gender){
  if (Gender == "w"){
    return(1) 
  } else  if (Gender == "m"){
    return(2) 
  } else {
    return(0)
  }
}

## convert Group levels (Gesund, Diabetiker) to numeric values (1 or 2)
numeric.group <- function(Group){
  if (Group == "Gesund"){
    return(1) 
  } else  if (Group == "Diabetiker"){
    return(2) 
  } else {
    return(0)
  }
}

## convert Vibration levels (schwere, maessige, leichte, or normal) to numeric values (3, 2, 1, or 0)
numeric.Vibration <- function(vibration_level){
  if (vibration_level == "schwer"){
    return(3) 
  } else  if (vibration_level == "maessig"){
    return(2) 
  }  else  if (vibration_level == "leicht"){
    return(1) 
  } else {
    return(0)
  }
}

## convert Sensation levels (fehlend, vermindert， or normal) to numeric values (2, 1, or 0)
numeric.Sensation <- function(Sensation){
  if (Sensation == "fehlend"){
    return(2) 
  } else  if (Sensation == "vermindert"){
    return(1) 
  } else {
    return(0)
  }
}

## convert PN levels (schwere/maessige/leichte or keine) to TWO factors values (1 or 0)
factor.PN <- function(PN_level){
  if (PN_level == "keine"){
    value <- 0
  } else {
    value <- 1
  }
  return(factor(value, levels = c(0, 1), labels = c("False", "True")))
}

## convert Vibration levels (schwer/maessig/leicht or normal) to TWO factors values (1 or 0)
factor.Vibration <- function(vibration_level){
  if (vibration_level == "normal"){
    value <- 0
  } else {
    value <- 1
  }
  return(factor(value, levels = c(0, 1), labels = c("False", "True")))
}

## convert Sensation levels (fehlend/vermindert or normal) to numeric values (1 or 0)
factor.Sensation <- function(sensation_level){
  if (sensation_level == "normal"){
    value <- 0
  } else {
    value <- 1
  }
  return(factor(value, levels = c(0, 1), labels = c("False", "True")))
}

## convert Vibration code 0-8 to a factor("Non", "Mild", "Moderate", "Severe")
convert.vibration.to.factor <- function(vibration){
  if (!is.numeric(vibration)){
    print("Warning: the input vibration is not a numeric variable, it will be converted to a numeric variable")
    vibration <- as.numeric(vibration)
  }
  value <- 0
  if (vibration > 5){
    value <- 0
  } else if (vibration > 2){
    value <- 1
  } else if (vibration > 0){
    value <- 2
  }  else {
    value <- 3
  }
  return(factor(value, levels = c(0, 1, 2, 3), labels = c("Non", "Mild", "Moderate", "Severe")))
}

## convert Vibration code 0-8 to a factor("Non", "Mild", "Moderate", "Severe")
convert.sensation.to.factor <- function(sensation){
  if (!is.numeric(sensation)){
    print("Warning: the input sensation is not a numeric variable, it will be converted to a numeric variable")
    sensation <- as.numeric(sensation)
  }
  value <- 0
  if (sensation == 0){
    value <- 0
  } else if (sensation == 1){
    value <- 1
  } else {
    value <- 2
  }
  return(factor(value, levels = c(0, 1, 2), labels = c("Non", "Mild", "Severe")))
}


## convert NSS score (10-0) to character levels (schwere, maessige, leichte, or keine) 
evaluate.NSS <- function(NSS){
  if (NSS > 6){
    return ("schwere")
  } else if (NSS > 4){
    return ("maessige")
  } else if (NSS > 0){
    return ("leichte")
  }  else {
    return ("keine")
  }
}

## convert NSS score (10-0) to numric levels (3, 2, 1, or 0) 
level.NSS <- function(NSS){
  if (NSS > 6){
    return (3)
  } else if (NSS > 4){
    return (2)
  } else if (NSS > 0){
    return (1)
  }  else {
    return (0)
  }
}

## convert NDS score (10-0) to character levels (schwere, maessige, leichte, or keine) 
evaluate.NDS <- function(NDS){
  if (NDS > 8){
    return ("schwere")
  } else if (NDS > 5){
    return ("maessige")
  } else if (NDS > 0){
    return ("leichte")
  }  else {
    return ("keine")
  }
}

## convert NDS score (10-0) to numric levels (3, 2, 1, or 0)  
factor.NDS.to.four.levels <- function(NDS){
  value <- 0
  if (NDS > 8){
    value <- 3
  } else if (NDS > 5){
    value <- 2
  } else if (NDS > 0){
    value <- 1
  }  else {
    value <- 0
  }
  
  return(factor(value, levels = c(3, 2, 1, 0), labels = c("Severe", "Moderate","Mild", "Non")))
}

## convert NDS score (10-0) to numric levels (3, 2, 1, or 0)  
factor.NDS.to.three.levels <- function(NDS){
  value <- 0
  if (NDS > 5){
    value <- 2
  } else if (NDS > 0){
    value <- 1
  }  else {
    value <- 0
  }
  
  return(factor(value, levels = c(2, 1, 0), labels = c("Severe", "Light", "Non")))
}

factor.age.to.three.levels <- function(age){
  if (age <=30){
    value <- 1
  } else if (age <=50){
    value <- 2
  } else {
    value <- 3
  }
  return(factor(value, levels = c(1, 2, 3), labels = c("CON1", "CON2","CON3")))
}

factor.MoCA.to.two.levels <- function(MoCA, cutoff){
  value <- 0
  if (MoCA > cutoff){
    value <- 0
  } else {
    value <- 1
  }
  return(factor(value, levels = c(1, 0), labels = c("Reduced", "Normal")))
}

factor.NDS.to.two.levels <- function(NDS){
  value <- 0
  if (NDS > 0){
    value <- 1
  } else {
    value <- 0
  }
  return(factor(value, levels = c(1, 0), labels = c("PN", "Non")))
}


factor.NDS.to.two.levels.Non.as.positive <- function(NDS){
  value <- 0
  if (NDS > 0){
    value <- 1
  } else {
    value <- 0
  }
  return(factor(value, levels = c(0, 1), labels = c("Non", "PN")))
}

factor.diabetes.group <- function(group.text){
  group.text <- as.character(group.text)
  if (group.text == "Diabetiker"){
    group <- "diabetes"
  } else {
    group <- "control"
  }
  return(factor(group, levels = c("control", "diabetes"), labels = c("Control", "Diabetes")))
}


paper.Gender.to.levels <- function(gender){
  return(factor(gender, levels = c("m", "w"), labels = c("Male", "Female")))
}

paper.Diabetes.Type.to.levels <- function(type){
  return(factor(type, levels = c(0, 1, 2, 3), labels = c("Non", "Type1", "Type2", "Type3")))
}

## factor Nervensystem (1:Polyneuropathie, 2: Depression, 3: Schädelhirntrauma, 4: Epilepsie, 5: Schlaganfall,6: Tremor, 7: transitorische Ischämie, 8: Depressionen, 9: Fußheberschwäche, 10: Karpaltunnel, 11: Embolie Auge, 12: Hinblutung, 13: Sonstige)
paper.Nerves.System.to.levels <- function(type){
  return(factor(type, levels = seq(1, 13, 1), labels = c("Polyneuropathie", "Depression", "Schädelhirntrauma", "Epilepsie", "Schlaganfall", "Tremor", "Ischämie", "Depressionen", "Fußheberschwäche", "Karpaltunnel", "Embolie Auge", "Hinblutung", "Sonstige")))
}

## convert NSS score (10-0) to numric levels (3, 2, 1, or 0)  
paper.NSS.four.levels <- function(NSS){
  value <- 0
  if (is.na(NSS) || NSS=="NA"|| NSS=="" ){
    value <- 4
  } else if (NSS > 6){ ## 7-10
    value <- 3
  } else if (NSS > 4){ ## 5-6
    value <- 2
  } else if (NSS > 2){ ## 3-4
    value <- 1
  }  else {
    value <- 0 ## 0-2
  }
  
  return(factor(value, levels = c(0, 1, 2, 3, 4), labels = c( "Normal (0-2)", "Mild (3-4)", "Moderate (5-6)", "Severe (7-10)", "N/A"), ordered = T))
}

## convert NDS score (10-0) to numric levels (3, 2, 1, or 0)  
paper.NDS.four.levels <- function(NDS){
  value <- 0
  if (is.na(NDS) || NDS=="NA"|| NDS=="" ){
    value <- 4
  } else if (NDS > 8){ ## 9-10
    value <- 3
  } else if (NDS > 5){ ## 6-8
    value <- 2
  } else if (NDS > 2){ ## 3-5
    value <- 1
  }  else {
    value <- 0 ## 0-2
  }
  
  return(factor(value, levels = c(0, 1, 2, 3, 4), labels = c( "Normal (0-2)", "Mild (3-5)", "Moderate (6-8)", "Severe (9-10)", "N/A"), ordered = T))
}

paper.diagnose.PN <- function(NSS, NDS){
  value <- 0
  if (is.na(NSS) || NSS=="NA"|| NSS=="" || is.na(NDS) || NDS=="NA"|| NDS=="" ){
    value <- 2
  } else if (NDS > 5){ ## NDS:6-10
    value <- 1
  } else if (NDS > 2 && NSS>5){ ## NDS>=3 with NSS>=6
    value <- 1
  } else {
    value <- 0 ## 0-2
  }
  
  return(factor(value, levels = c(0, 1, 2), labels = c( "Normal", "PN", "N/A"), ordered = T))
}

## convert Achillessehnenreflex to numeric levels (3, 2, 1, or 0)  
paper.reflex.to.levels <- function(value){
  return(factor(value, levels = c(0, 1, 2), labels = c( "Normal", "Reduced", "Absent")))
}

## convert Schmerzempfinden (Pin-prick) to numeric levels (1 or 0)  
paper.pinprick.to.levels <- function(value){
  return(factor(value, levels = c(0, 1), labels = c( "Present", "Reduced/Absent")))
}

## convert Temperature sensation (cold tuning fork)	to numeric levels (1 or 0)  
paper.temperature.to.levels <- function(value){
  return(factor(value, levels = c(0, 1), labels = c( "Present", "Reduced/Absent")))
}

## convert Sensibilitätsprüfung Monofilament test result to numeric levels (2, 1 or 0)  
paper.monofilament.to.levels <- function(value){
  return(factor(value, levels = c(0, 1, 2), labels = c( "Present", "Reduced", "Absent")))
}

## convert Vibration code 0-8 to a factor("Non", "Mild", "Moderate", "Severe")
paper.vibration.to.factor <- function(vibration){
  value <- 0
  if (is.na(vibration)){
    value <- 4
  } else {
    if (!is.numeric(vibration)){
      print("Warning: the input vibration is not a numeric variable, it will be converted to a numeric variable")
      vibration <- as.numeric(vibration)
    }
    if (vibration > 5){
      value <- 0
    } else if (vibration == 5){
      value <- 1
    } else if (vibration > 2){
      value <- 2
    }  else {
      value <- 3
    }
  } 
  return(factor(value, levels = c(0, 1, 2, 3, 4), labels = c("Normal (6-8)", "Mild (5)", "Moderate (3-4)", "Severe (0-2)", "N/A"), ordered = T))
}

## convert Vibration score (0-8) to character levels (schwer, maessig, leicht, or normal) 
evaluate.Vibration.Sensation <- function(vibration){
  if (vibration > 5){
    return ("normal")
  } else if (vibration == 5){
    return ("leicht")
  } else if (vibration < 5 && vibration > 2){
    return ("maessig")
  }  else {
    return ("schwer")
  }
}

## convert Vibration score (0-8) to numric levels (3, 2, 1, or 0) 
level.Vibration.Sensation <- function(vibration){
  if (vibration > 5){
    return (0)
  } else if (vibration == 5){
    return (1)
  } else if (vibration < 5 && vibration > 2){
    return (2)
  }  else {
    return (3)
  }
}

## convert Sensation score (2-0) to character levels (fehlend, vermindert, or normal) 
evaluate.Sensation.Test <- function(sensation){
  if (sensation == 0){
    return ("normal")
  } else if (sensation == 1){
    return ("vermindert")
  } else {
    return ("fehlend")
  }
}


## convert Hand/Foot Side (0, 1, 2) to character levels (beide, rechts, or links) 
change.hand.foot.as.chara <- function(hand.foot){
  if (hand.foot == 0){
    return("Both")
  } else if (hand.foot == 1){
    return("Right")
  } else {
    return("Left")
  }
}

# convert NDS value to a binary factor (True: NDS > 0, False: NDS == 0)
nds.to.binary.factors <- function(NDS){
  if (NDS == 0){
    value <- 0
  } else {
    value <- 1
  }
  return(factor(value, levels = c(0, 1), labels = c("Non", "PN")))
}

# convert Vibration value to a binary factor (True: Vibration 0-5, False: Vibration 6-8)
vibration.to.binary.factors <- function(Vibration){
  if (Vibration > 5){
    value <- 0
  } else {
    value <- 1
  }
  return(factor(value, levels = c(0, 1), labels = c("Non", "PN")))
}


# convert Sensation value to a binary factor (True: Sensation 1-2, False: Sensation 0)
sensation.to.binary.factors <- function(Sensation){
  if (Sensation == 0){
    value <- 0
  } else {
    value <- 1
  }
  return(factor(value, levels = c(0, 1), labels = c("Non", "PN")))
}

# convert probandId to a numric value
change.proband.id.to.num <- function(ProbandId){
  if (is.character(ProbandId)){
    probandId <- substr(ProbandId, 2, 4)
    return(as.numeric(probandId))
  } else {
    print("Error: wrong format of the ProbandId")
  }
}


summarize.result <- function(seed.value, game.name, classifier.name, all.features.number, final.features.number, best.trained.model, first.cm){
  
  
  train.accuracy <- round(mean(best.trained.model$resample[, "Accuracy"]), 3)
  train.cv.accuracy.max <- round(max(best.trained.model$resample[, "Accuracy"]), 3)
  train.cv.accuracy.min <- round(min(best.trained.model$resample[, "Accuracy"]), 3)
  
  train.kappa.value <- round(mean(best.trained.model$resample[, "Kappa"]), 3)
  train.cv.kappa.max <- round(max(best.trained.model$resample[, "Kappa"]), 3)
  train.cv.kappa.min <- round(min(best.trained.model$resample[, "Kappa"]), 3)

  # train.highest.accuracy.no.round <- max(best.trained.model$results$Accuracy)
  # best.accuracy.row <- subset(best.trained.model$results, Accuracy==train.highest.accuracy.no.round)
  # print(best.accuracy.row)
  # train.kappa.value <- round(best.accuracy.row$Kappa[1], 3)
  # train.highest.accuracy <- round(train.highest.accuracy.no.round, 3)
  
  predict.accuracy <- round(first.cm$overall["Accuracy"], 3)
  predict.kappa <- round(first.cm$overall["Kappa"], 3)
  
  true.posi <- round(first.cm$table[1, 1], 3)
  true.nega <- round(first.cm$table[2, 2], 3)
  false.posi <- round(first.cm$table[1, 2], 3)
  false.nega <- round(first.cm$table[2, 1], 3)
  sensi <- round(true.posi/(true.posi+false.nega), 3)
  speci <- round(true.nega/(true.nega+false.posi), 3)
  preci <- round(true.posi/(true.posi+false.posi), 3)
  balanced_accu <- round((sensi+speci)*0.5, 3)
  
  model.paras <- ""
  for (col.n in 1:ncol(best.trained.model$bestTune)) {
    col.name <- names(best.trained.model$bestTune)[col.n]
    col.value <- as.character(best.trained.model$bestTune[1, col.n])  
    if (model.paras == ""){
      model.paras <- paste(col.name, ": ", col.value, sep = "")
    } else {
      model.paras <- paste(model.paras, ", ", col.name, ": ", col.value, sep = "")
    }
  }
  
  final.result.row <- data.frame(
    seed.value, 
    game.name, 
    classifier.name, 
    all.features.number, 
    final.features.number,
    train.accuracy,
    train.cv.accuracy.max,
    train.cv.accuracy.min,
    train.kappa.value,
    train.cv.kappa.max,
    train.cv.kappa.min,
    predict.accuracy,
    predict.kappa,
    true.posi, #TP
    true.nega, #TN
    false.posi, #FP
    false.nega, #FN
    sensi,
    speci,
    preci,
    balanced_accu,
    model.paras
  )
  
  names(final.result.row) <- c("Seed", "Game", "Classifier", "N Predictors", "Final Predictors", 
                               "Train Accuracy", "CV Train Accuracy Max", "CV Train Accuracy Min",
                               "Train Kappa", "CV Train Kappa Max", "CV Train Kappa Min",
                              "Predict Accuracy", "Predict Kappa", "TP", "TN", "FP", "FN", "Sensitivity", "Specificity",  "Precision",
                              "Balanced Accuracy", "Best Tune")
  
  return(final.result.row)
}


summarize.ROC.metric.result <- function(seed.value, game.name, classifier.name, all.features.number, final.features.number, best.trained.model, first.cm){
  train.highest.accuracy.no.round <- max(best.trained.model$results$ROC)
  train.highest.accuracy <- round(train.highest.accuracy.no.round, 3)
  predict.accuracy <- round(first.cm$overall["Accuracy"], 3)
  predict.kappa <- round(first.cm$overall["Kappa"], 3)
  
  true.posi <- round(first.cm$table[1, 1], 3)
  true.nega <- round(first.cm$table[2, 2], 3)
  false.posi <- round(first.cm$table[1, 2], 3)
  false.nega <- round(first.cm$table[2, 1], 3)
  sensi <- round(true.posi/(true.posi+false.nega), 3)
  speci <- round(true.nega/(true.nega+false.posi), 3)
  preci <- round(true.posi/(true.posi+false.posi), 3)
  balanced_accu <- round((sensi+speci)*0.5, 3)
  
  model.paras <- ""
  for (col.n in 1:ncol(best.trained.model$bestTune)) {
    col.name <- names(best.trained.model$bestTune)[col.n]
    col.value <- as.character(best.trained.model$bestTune[1, col.n])  
    if (model.paras == ""){
      model.paras <- paste(col.name, ": ", col.value, sep = "")
    } else {
      model.paras <- paste(model.paras, ", ", col.name, ": ", col.value, sep = "")
    }
  }
  
  final.result.row <- data.frame(
    seed.value, 
    game.name, 
    classifier.name, 
    all.features.number, 
    final.features.number,
    train.highest.accuracy,
    predict.accuracy,
    predict.kappa,
    true.posi, #TP
    true.nega, #TN
    false.posi, #FP
    false.nega, #FN
    sensi,
    speci,
    preci,
    balanced_accu,
    model.paras
  )
  
  names(final.result.row) <- c("Seed", "Game", "Classifier", "N Predictors", "Final Predictors", 
                               "Train Accuracy",
                               "Predict Accuracy", "Predict Kappa", "TP", "TN", "FP", "FN", "Sensitivity", "Specificity",  "Precision",
                               "Balanced Accuracy", "Best Tune")
  
  return(final.result.row)
}


summarize.Sens.metric.result <- function(seed.value, game.name, classifier.name, all.features.number, final.features.number, best.trained.model, first.cm){
  train.highest.accuracy.no.round <- max(best.trained.model$results$Sens)
  train.highest.accuracy <- round(train.highest.accuracy.no.round, 3)
  predict.accuracy <- round(first.cm$overall["Accuracy"], 3)
  predict.kappa <- round(first.cm$overall["Kappa"], 3)
  
  true.posi <- round(first.cm$table[1, 1], 3)
  true.nega <- round(first.cm$table[2, 2], 3)
  false.posi <- round(first.cm$table[1, 2], 3)
  false.nega <- round(first.cm$table[2, 1], 3)
  sensi <- round(true.posi/(true.posi+false.nega), 3)
  speci <- round(true.nega/(true.nega+false.posi), 3)
  preci <- round(true.posi/(true.posi+false.posi), 3)
  balanced_accu <- round((sensi+speci)*0.5, 3)
  
  model.paras <- ""
  for (col.n in 1:ncol(best.trained.model$bestTune)) {
    col.name <- names(best.trained.model$bestTune)[col.n]
    col.value <- as.character(best.trained.model$bestTune[1, col.n])  
    if (model.paras == ""){
      model.paras <- paste(col.name, ": ", col.value, sep = "")
    } else {
      model.paras <- paste(model.paras, ", ", col.name, ": ", col.value, sep = "")
    }
  }
  
  final.result.row <- data.frame(
    seed.value, 
    game.name, 
    classifier.name, 
    all.features.number, 
    final.features.number,
    train.highest.accuracy,
    predict.accuracy,
    predict.kappa,
    true.posi, #TP
    true.nega, #TN
    false.posi, #FP
    false.nega, #FN
    sensi,
    speci,
    preci,
    balanced_accu,
    model.paras
  )
  
  names(final.result.row) <- c("Seed", "Game", "Classifier", "N Predictors", "Final Predictors", 
                               "Train Sensitivity",
                               "Predict Accuracy", "Predict Kappa", "TP", "TN", "FP", "FN", "Sensitivity", "Specificity",  "Precision",
                               "Balanced Accuracy", "Best Tune")
  
  return(final.result.row)
}



# x is a matrix containing the data
# method : correlation method. "pearson"" or "spearman"" is supported
# removeTriangle : remove upper or lower triangle
# results :  if "html" or "latex"
# the results will be displayed in html or latex format
corstars <-function(x, method=c("pearson", "spearman"), removeTriangle=c("upper", "lower"),
                    result=c("none", "html", "latex")){
  #Compute correlation matrix
  require(Hmisc)
  x <- as.matrix(x)
  correlation_matrix<-rcorr(x, type=method[1])
  R <- correlation_matrix$r # Matrix of correlation coeficients
  p <- correlation_matrix$P # Matrix of p-value 
  
  ## Define notions for significance levels; spacing is important.
  mystars <- ifelse(p < .0001, "****", ifelse(p < .001, "*** ", ifelse(p < .01, "**  ", ifelse(p < .05, "*   ", "    "))))
  
  ## trunctuate the correlation matrix to two decimal
  R <- format(round(cbind(rep(-1.11, ncol(x)), R), 2))[,-1]
  
  ## build a new matrix that includes the correlations with their apropriate stars
  Rnew <- matrix(paste(R, mystars, sep=""), ncol=ncol(x))
  diag(Rnew) <- paste(diag(R), " ", sep="")
  rownames(Rnew) <- colnames(x)
  colnames(Rnew) <- paste(colnames(x), "", sep="")
  
  ## remove upper triangle of correlation matrix
  if(removeTriangle[1]=="upper"){
    Rnew <- as.matrix(Rnew)
    Rnew[upper.tri(Rnew, diag = TRUE)] <- ""
    Rnew <- as.data.frame(Rnew)
  }
  
  ## remove lower triangle of correlation matrix
  else if(removeTriangle[1]=="lower"){
    Rnew <- as.matrix(Rnew)
    Rnew[lower.tri(Rnew, diag = TRUE)] <- ""
    Rnew <- as.data.frame(Rnew)
  }
  
  ## remove last column and return the correlation matrix
  Rnew <- cbind(Rnew[1:length(Rnew)-1])
  if (result[1]=="none") return(Rnew)
  else{
    if(result[1]=="html") print(xtable(Rnew), type="html")
    else print(xtable(Rnew), type="latex") 
  }
} 


# Calculate the table correlation
# x.cols the column indexs of x variable
# y.col : the column index of y variable
# table : the table to perform correlation test
# the results will be returned within a table including outcome, correlation coffecient, p.value, and significant level ***
get.correlation.table <- function(x.cols, y.col, table){
  if (y.col>ncol(table)){
    print("Error: invalid col index of y")
    return()
  }
  if (max(x.cols)>ncol(table) || min(x.cols)<1){
    print("Error: invalid col index of x")
    return()   
  }
  
  results <- c()
  y <- table[, y.col]

  for (x.index in x.cols) {
    x <- table[, x.index]
    
    cor.result <- cor.test(x, y)
    outcome <- names(table)[x.index]
    cor.coff <- round(cor.result$estimate, 2) 
    p.value <- round(cor.result$p.value, 4) 
    if (is.na(p.value)){
      print(paste("!!!ERROR: Correlation p value is NA between", outcome, "and", names(table)[y.col]))
    } else {
      if (p.value < 0.001){
        sign.level <- "***"
      } else if (p.value < 0.01){
        sign.level <- "**"
      } else if (p.value < 0.05){
        sign.level <- "*"
      } else {
        sign.level <- ""
      }
      if (length(results)==0){
        results <- data.frame(outcome, cor.coff, p.value, sign.level)
      } else {
        results <- rbind(results, data.frame(outcome, cor.coff, p.value, sign.level))
      }
    }
  }
  return(results)
}



# convert accuracy or kappa to factor levels
# the input accuracy should be ranged from 0 to 1
convert.accruracy.to.factors <- function(accuracy){
  if (accuracy > 0.9){
    value <- 5
  } else if (accuracy > 0.8){
    value <- 4
  } else if (accuracy > 0.7){
    value <- 3
  } else if (accuracy > 0.6){
    value <- 2
  } else  if (accuracy > 0.5){
    value <- 1
  } else {
    value <- 0
  }
  return(factor(value, levels = seq(5, 0, -1), labels = c(">90", "80-90", "70-80", "60-70", "50-60",  "<50")))
}


get.max.label <- function(values, max){
  labels <- c()
  has.found <- FALSE
  for (value in values) {
    if (max==value && !has.found){
      has.found <- TRUE
      labels <- c(labels, paste(as.character(value), "%", sep = "") )
    } else {
      labels <- c(labels, "")
    }
  }
  return(labels)
}



# convert  kappa to factor levels
# the input kappa should be ranged from -1 to 1
convert.kappa.to.factors <- function(kappa){
  if (kappa > 0.7){
    value <- 5
  } else if (kappa > 0.6){
    value <- 4
  } else if (kappa > 0.5){
    value <- 3
  } else if (kappa > 0.4){
    value <- 2
  } else  if (kappa > 0.3){
    value <- 1
  } else {
    value <- 0
  }
  return(factor(value, levels = seq(5, 0, -1), labels = c(">0.7", "0.6-0.7", "0.5-0.6", "0.4-0.5", "0.3-0.4",  "<0.3")))
}



get.max.kappa.label <- function(values, max){
  labels <- c()
  has.found <- FALSE
  for (value in values) {
    if (max==value && !has.found){
      has.found <- TRUE
      labels <- c(labels, paste(as.character(round(value, 2)), sep = "") )
    } else {
      labels <- c(labels, "")
    }
  }
  return(labels)
}


# I really liked the beautiful confusion matrix visualization from @Cybernetic and made two tweaks to hopefully improve it further.
# https://stackoverflow.com/questions/23891140/r-how-to-visualize-confusion-matrix-using-the-caret-package
draw_confusion_matrix <- function(cm) {
  
  total <- sum(cm$table)
  res <- as.numeric(cm$table)
  
  # Generate color gradients. Palettes come from RColorBrewer.
  greenPalette <- c("#F7FCF5","#E5F5E0","#C7E9C0","#A1D99B","#74C476","#41AB5D","#238B45","#006D2C","#00441B")
  redPalette <- c("#FFF5F0","#FEE0D2","#FCBBA1","#FC9272","#FB6A4A","#EF3B2C","#CB181D","#A50F15","#67000D")
  getColor <- function (greenOrRed = "green", amount = 0) {
    if (amount == 0)
      return("#FFFFFF")
    palette <- greenPalette
    if (greenOrRed == "red")
      palette <- redPalette
    colorRampPalette(palette)(100)[10 + ceiling(90 * amount / total)]
  }
  
  # set the basic layout
  layout(matrix(c(1,1,2)))
  par(mar=c(2,2,2,2))
  plot(c(100, 345), c(300, 450), type = "n", xlab="", ylab="", xaxt='n', yaxt='n')
  title('CONFUSION MATRIX', cex.main=2)
  
  # create the matrix 
  classes = colnames(cm$table)
  rect(150, 430, 240, 370, col=getColor("green", res[1]))
  text(195, 435, classes[1], cex=1.2)
  rect(250, 430, 340, 370, col=getColor("red", res[3]))
  text(295, 435, classes[2], cex=1.2)
  text(125, 370, 'Predicted', cex=1.3, srt=90, font=2)
  text(245, 450, 'Actual', cex=1.3, font=2)
  rect(150, 305, 240, 365, col=getColor("red", res[2]))
  rect(250, 305, 340, 365, col=getColor("green", res[4]))
  text(140, 400, classes[1], cex=1.2, srt=90)
  text(140, 335, classes[2], cex=1.2, srt=90)
  
  # add in the cm results
  text(195, 400, res[1], cex=1.6, font=2, col='white')
  text(195, 335, res[2], cex=1.6, font=2, col='white')
  text(295, 400, res[3], cex=1.6, font=2, col='white')
  text(295, 335, res[4], cex=1.6, font=2, col='white')
  
  # add in the specifics 
  plot(c(100, 0), c(100, 0), type = "n", xlab="", ylab="", main = "DETAILS", xaxt='n', yaxt='n')
  text(10, 85, names(cm$byClass[1]), cex=1.2, font=2)
  text(10, 70, round(as.numeric(cm$byClass[1]), 3), cex=1.2)
  text(30, 85, names(cm$byClass[2]), cex=1.2, font=2)
  text(30, 70, round(as.numeric(cm$byClass[2]), 3), cex=1.2)
  text(50, 85, names(cm$byClass[5]), cex=1.2, font=2)
  text(50, 70, round(as.numeric(cm$byClass[5]), 3), cex=1.2)
  text(70, 85, names(cm$byClass[6]), cex=1.2, font=2)
  text(70, 70, round(as.numeric(cm$byClass[6]), 3), cex=1.2)
  text(90, 85, names(cm$byClass[7]), cex=1.2, font=2)
  text(90, 70, round(as.numeric(cm$byClass[7]), 3), cex=1.2)
  
  # add in the accuracy information 
  text(30, 35, names(cm$overall[1]), cex=1.5, font=2)
  text(30, 20, round(as.numeric(cm$overall[1]), 3), cex=1.4)
  text(70, 35, names(cm$overall[2]), cex=1.5, font=2)
  text(70, 20, round(as.numeric(cm$overall[2]), 3), cex=1.4)
}


## Function to filter out features that are complete zeros or nearly zero variances
## features.table: a feature matrix without label/class!!!
filter.out.zero.variance.features <- function(features.table){
  ## numeric all columns
  non.numeric.cols <- c()
  for (col_n in 1: ncol(features.table)) {
    if (!is.numeric(features.table[, col_n])){
      if (is.character(features.table[, col_n])){
        features.table[, col_n] <- as.numeric(factor(features.table[, col_n]))
      } else if (is.factor(features.table[, col_n])){
        features.table[, col_n] <- as.numeric(features.table[, col_n])
      } else {
        features.table[, col_n] <- as.numeric(features.table[, col_n])
      }
      non.numeric.cols <- c(non.numeric.cols, names(features.table)[col_n])
    }
  }
  if (length(non.numeric.cols)>0){
    print(non.numeric.cols)
    print(paste("------ Convert", length(non.numeric.cols), "variables to numeric vars --------"))
  }
  
  # rename the label column
  # names(features.table)[ncol(features.table)] <- "class"
  # remove columns which are complete zeros
  select_cols <- c()
  all.zero.features <-c()
  same.values.features <-c()
  for(col_n in 1: ncol(features.table)){
    col_name <- names(features.table)[col_n]
    if (all(features.table[, col_n] == 0)){
      all.zero.features <- c(all.zero.features, col_name)
    } else if (length(unique(features.table[, col_n]))==1){
      same.values.features <- c(same.values.features, col_name)
    } else {
      select_cols <- c(select_cols, col_name)
    }
  }
  
  if (length(all.zero.features)>0){
    print(all.zero.features)
    print(paste("-------------------- Remove", length(all.zero.features), "variables with all zero value----------"))
  }
  
  if (length(same.values.features)>0){
    print(same.values.features)
    print(paste("-------------------- Remove", length(same.values.features), "variables with same values----------"))
  }
  
  nonzero.features.table <- subset(features.table, select = select_cols)
  features.table <- nonzero.features.table
  
  ## check column complete zeros or zero variances
  cols.zero.var <- nearZeroVar(features.table)
  if (length(cols.zero.var) > 0){
    featuresNearZeroVar <- features.table[, -cols.zero.var]
    print(names(features.table)[cols.zero.var])
    print(paste("------------------ delete", ncol(features.table)-ncol(featuresNearZeroVar), "features with near zero variances -------------------"))
    features.table <- featuresNearZeroVar
  } else {
    print("no columns with near zero variance were detected!")
  }
  
  
  
  return(features.table)
}

## Function to convert non-numeric features to numric values
## features.table: a feature matrix without label/class!!!
numeric.features.table <- function(features.table){
  for (col.index in 1: ncol(features.table)) {
    if (is.character(features.table[, col.index])){
      features.table[, col.index] <- as.numeric(factor(features.table[, col.index]))
      print(paste("convert character features to numeric ", names(features.table)[col.index]))
    } else if (!is.numeric(features.table[, col.index])){
      features.table[, col.index] <- as.numeric(features.table[, col.index])
      print(paste("convert features to numeric ", names(features.table)[col.index]))
    }
  }
  return(features.table)
}

## A function to get full information of features according to the provided feature name
## i.e. the official name, source task or group, and extracted game will be returned as a dataframe with the row length same with the number of features
## Input: a vector contains names of all features
get.feature.full.info <- function(features.vector){
  afs.feature.labels <- c("reactionTime", "anticipationTime", "inIdealAreaTime", "timeUnderIdealArea", "timeOverIdealArea", "carMoveError", "errorTime", "finalDistance", 
                          "preSum", "preMax", "preMin", "preSd", "preMean",
                          "preDiffSum", "preDiffMax", "preDiffMin", "preDiffSd", "preDiffMean",
                          "preGradAbsSum", "preGradSum", "preGradMax", "preGradMin", "preGradSd", "preGradMean", 
                          "preIntegrals", "appleCollected")
  
  afs.feature.names <- c("Reaction time (s)", "Anticipation time (s)", "Time inside catching area (s)", "Time outside catching area (s)", "Time outside catching area (s)",
                         "Frequency outside catching area (n)", "Time outside catching area (s)", "Final virtual distance", 
                         "Normalized pressure", "Normalized pressure", "Normalized pressure", "Normalized pressure", "Normalized pressure",
                         "Pressure differences between successive frames", "Pressure differences between successive frames", "Pressure differences between successive frames", "Pressure differences between successive frames", "Pressure differences between successive frames",
                         "Pressure gradients between successive frames", "Pressure gradients between successive frames", "Pressure gradients between successive frames", "Pressure gradients between successive frames", "Pressure gradients between successive frames", "Pressure gradients between successive frames", 
                         "Pressure-time integrals", "Apple caught (yes/no)")
  
  bfs.feature.labels <- c("smiley", "collisions", "mainSmileyDistance", "subSmiley1Distance", "subSmiley2Distance", "subSmiley3Distance",
                          "posDev", "preSum", "preMax", "preMin", "preSd", "preMean",
                          "preDiff", "preGrad", "preIntegrals")
  
  bfs.feature.names <- c("Smiley count (n)", "Collision frequency (n)",
                         "Minimal virtual distance smiley 1", "Minimal virtual distance smiley 2", "Minimal virtual distance smiley 3", "Minimal virtual distance smiley 4", "Virtual deviation of ideal flying route", 
                         "Normalized pressure", "Normalized pressure", "Normalized pressure", "Normalized pressure", "Normalized pressure",
                         "Pressure differences between successive frames", "Pressure gradients between successive frames", "Pressure-time integrals")
  
  fbs.feature.labels <- c("anticipationTime",  "errorTime", "reactionTime", 
                          "preReaIntegralsLeft", "preReaIntegralsRight",
                          "preSumLeft", "preMaxLeft", "preMinLeft", "preSdLeft", "preMeanLeft",
                          "preDiffSumLeft", "preDiffMaxLeft", "preDiffMinLeft", "preDiffSdLeft", "preDiffMeanLeft",
                          "preGradAbsSumLeft", "preGradSumLeft", "preGradMaxLeft", "preGradMinLeft", "preGradSdLeft", "preGradMeanLeft",
                          "preExeIntegralsLeft",
                          "preSumRight", "preMaxRight", "preMinRight", "preSdRight", "preMeanRight",
                          "preDiffSumRight", "preDiffMaxRight", "preDiffMinRight", "preDiffSdRight", "preDiffMeanRight",
                          "preGradAbsSumRight", "preGradSumRight", "preGradMaxRight", "preGradMinRight", "preGradSdRight", "preGradMeanRight",
                          "preExeIntegralsRight") 
  fbs.feature.names <- c("Anticipation time (s)", "Time outside ideal pressure zone (s)", "Reaction time (s)",
                         "Pressure-time integrals of the reaction phase (L)", "Pressure-time integrals of the reaction phase (R)", 
                         "Normalized pressure (L)", "Normalized pressure (L)", "Normalized pressure (L)", "Normalized pressure (L)", "Normalized pressure (L)",
                         "Pressure differences between successive frames (L)", "Pressure differences between successive frames (L)", "Pressure differences between successive frames (L)", "Pressure differences between successive frames (L)", "Pressure differences between successive frames (L)",
                         "Pressure gradients between successive frames (L)", "Pressure gradients between successive frames (L)", "Pressure gradients between successive frames (L)", "Pressure gradients between successive frames (L)", "Pressure gradients between successive frames (L)", "Pressure gradients between successive frames (L)", 
                         "Pressure-time integrals of the execution phase (L)",
                         "Normalized pressure (R)", "Normalized pressure (R)", "Normalized pressure (R)", "Normalized pressure (R)", "Normalized pressure (R)",
                         "Pressure differences between successive frames (R)", "Pressure differences between successive frames (R)", "Pressure differences between successive frames (R)", "Pressure differences between successive frames (R)", "Pressure differences between successive frames (R)",
                         "Pressure gradients between successive frames (R)", "Pressure gradients between successive frames (R)", "Pressure gradients between successive frames (R)", "Pressure gradients between successive frames (R)", "Pressure gradients between successive frames (R)", "Pressure gradients between successive frames (R)", 
                         "Pressure-time integrals of the execution phase (R)")
  
  jps.feature.labels <- c("attempts", "preDevi", "anticipationTime", "executionTime", 
                          "executionPre", "pti", "preGrad", "preDiff", "preIntegrals")
  
  jps.feature.names <- c("Attempt count (n)", "Deviation from ideal pressure", "Anticipation time (s)", "Execution time (s)", 
                         "Mean pressure of execution phase", 
                         "Pressure-time integrals", 
                         "Pressure gradients between successive frames", 
                         "Pressure differences between successive frames", "Pressure-time integrals")
  
  
  numbers10 <- c("1", "2", "3", "4", "5", "6", "7", "8", "9", "0")
  numbers20 <- c( "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20")
  labels <- c()
  names <- c()
  groups <- c()
  games <- c()
  for (feature.name in features.vector) {
    name1 <- substr(feature.name, 1, nchar(feature.name)-3)
    game <- ""
    label <- name1
    name <- ""
    group = ""
    if (grepl("Afs", feature.name, fixed = T)){
      game <- "AC"
      for (afs.label.index in 1:length(afs.feature.labels)) {
        afs.label <- afs.feature.labels[afs.label.index]
        if (grepl(afs.label, name1, fixed = T)){
          name <- afs.feature.names[afs.label.index]
        }
      }
      group.str1 <- substr(name1, nchar(name1)-2, nchar(name1)-2)
      group.str2 <- substr(name1, nchar(name1)-1, nchar(name1)-1)
      group.str3 <- substr(name1, nchar(name1), nchar(name1))
      group.str23 <- substr(name1, nchar(name1)-1, nchar(name1))
      if (group.str1=="G"){
        group = paste("Group", group.str23)
      } else if (group.str2=="G"){
        group = paste("Group", group.str3)
      } else if (group.str2 %in% numbers10){
        group = paste("Task", group.str23)
      } else if (group.str3 %in% numbers10){
        group = paste("Task", group.str3)
      }
    } else if (grepl("Bfs", feature.name, fixed = T)){
      foot.side <- substr(name1, nchar(name1), nchar(name1))
      game <- paste( "BF (", foot.side, ")", sep = "")
      for (bfs.label.index in 1:length(bfs.feature.labels)) {
        bfs.label <- bfs.feature.labels[bfs.label.index]
        if (grepl(bfs.label, name1, fixed = T)){
          name <- bfs.feature.names[bfs.label.index]
        }
      }
      group.str1 <- substr(name1, nchar(name1)-1, nchar(name1)-1)
      group.str2 <- substr(name1, nchar(name1)-2, nchar(name1)-1)
      if (group.str2 %in% numbers20){
        group = paste("Task", group.str2)
      } else if (group.str1 %in% numbers10){
        group = paste("Task", group.str1)
      } else{
        group.str9 <- substr(name1, nchar(name1)-9, nchar(name1)-1)
        if (grepl("All", group.str9, fixed = T)){
          group = "Group 1"
        } else if (grepl("LowPre", group.str9, fixed = T)){
          group = "Group 2"
        } else if (grepl("MiddlePre", group.str9, fixed = T)){
          group = "Group 3"
        }  else if (grepl("HighPre", group.str9, fixed = T)){
          group = "Group 4"
        } 
      }
    } else if (grepl("Fbs", feature.name, fixed = T)){
      game <- "CP"
      for (fbs.label.index in 1:length(fbs.feature.labels)) {
        fbs.label <- fbs.feature.labels[fbs.label.index]
        if (grepl(fbs.label, name1, fixed = T)){
          name <- fbs.feature.names[fbs.label.index]
        }
      }
      group.str1 <- substr(name1, nchar(name1), nchar(name1))
      group.str2 <- substr(name1, nchar(name1)-1, nchar(name1))
      if (group.str2 %in% numbers20){
        group = paste("Task", group.str2)
      } else if (group.str1 %in% numbers10){
        group = paste("Task", group.str1)
      } else{
        group.str9 <- substr(name1, nchar(name1)-8, nchar(name1))
        if (grepl("All", group.str9, fixed = T)){
          group = "Group 1"
        } else if (grepl("LowPre", group.str9, fixed = T)){
          group = "Group 2"
        } else if (grepl("HighPre", group.str9, fixed = T)){
          group = "Group 3"
        }  else if (grepl("owPreLeft", group.str9, fixed = T)){
          group = "Group 4"
        }  else if (grepl("wPreRight", group.str9, fixed = T)){
          group = "Group 5"
        }  else if (grepl("owPreBoth", group.str9, fixed = T)){
          group = "Group 6"
        }  else if (grepl("ghPreLeft", group.str9, fixed = T)){
          group = "Group 7"
        } else if (grepl("hPreRight", group.str9, fixed = T)){
          group = "Group 8"
        } else if (grepl("ghPreBoth", group.str9, fixed = T)){
          group = "Group 9"
        } 
      }
      
    } else if (grepl("Jps", feature.name, fixed = T)){
      game <- "IJ"
      for (jps.label.index in 1:length(jps.feature.labels)) {
        jps.label <- jps.feature.labels[jps.label.index]
        if (grepl(jps.label, name1, fixed = T)){
          name <- jps.feature.names[jps.label.index]
        }
      }
      group.str1 <- substr(name1, nchar(name1), nchar(name1))
      group.str2 <- substr(name1, nchar(name1)-1, nchar(name1))
      if (group.str2 %in% numbers20){
        group = paste("Task", group.str2)
      } else if (group.str1 %in% numbers10){
        group = paste("Task", group.str1)
      } else{
        group.str9 <- substr(name1, nchar(name1)-8, nchar(name1))
        if (grepl("All", group.str9, fixed = T)){
          group = "Group 1"
        } else if (grepl("Left", group.str9, fixed = T)){
          group = "Group 2"
        } else if (grepl("Right", group.str9, fixed = T)){
          group = "Group 3"
        }  else if (grepl("Front", group.str9, fixed = T)){
          group = "Group 4"
        }  else if (grepl("LowPre", group.str9, fixed = T)){
          group = "Group 5"
        }  else if (grepl("MiddlePre", group.str9, fixed = T)){
          group = "Group 6"
        }  else if (grepl("HighPre", group.str9, fixed = T)){
          group = "Group 7"
        } 
      }
    }
    
    labels <- c(labels, label)
    names <- c(names, name)
    groups <- c(groups, group)
    games <- c(games, game)
    
  }
  output.table <- data.frame(Label = labels, Name = names, Group = groups, Game = games)
  print(output.table)
  return(output.table)
  
}

## A function to get full information of features according to the provided feature name
## i.e. the official name, source task or group, and extracted game will be returned as a dataframe with the row length same with the number of features
## Input: a vector contains names of all features
get.feature.full.info.V2 <- function(features.vector){
  afs.feature.labels <- c("reactionTime", "anticipationTime", "inIdealAreaTime", "timeUnderIdealArea", "timeOverIdealArea", "carMoveError", "errorTime", "finalDistance", 
                          "preSum", "preMax", "preMin", "preSd", "preMean",
                          "preDiffSum", "preDiffMax", "preDiffMin", "preDiffSd", "preDiffMean",
                          "preGradAbsSum", "preGradSum", "preGradMax", "preGradMin", "preGradSd", "preGradMean", 
                          "preIntegrals", "appleCollected")
  
  afs.feature.names <- c("Reaction time", "Anticipation time", "Time inside catching area", "Time outside catching area", "Time outside catching area",
                         "Frequency outside catching area", "Time outside catching area", "Final virtual distance", 
                         "Normalized pressure", "Normalized pressure", "Normalized pressure", "Normalized pressure", "Normalized pressure",
                         "Pressure difference", "Pressure difference", "Pressure difference", "Pressure difference", "Pressure difference",
                         "Pressure gradient", "Pressure gradient", "Pressure gradient", "Pressure gradient", "Pressure gradient", "Pressure gradient", 
                         "Pressure-time integral", "Apple")
  afs.feature.units <- c("(sec)", "(sec)", "(sec)", "(sec)", "(sec)",
                         "(n)", "(sec)", "(arb. unit)", 
                         "(arb. unit)", "(arb. unit)", "(arb. unit)", "(arb. unit)", "(arb. unit)",
                         "(arb. unit)", "(arb. unit)", "(arb. unit)", "(arb. unit)", "(arb. unit)",
                         "(arb. unit)", "(arb. unit)", "(arb. unit)", "(arb. unit)", "(arb. unit)", "(arb. unit)", 
                         "(arb. unit)", "(n)")
  
  bfs.feature.labels <- c("smiley", "collisions", "mainSmileyDistance", "subSmiley1Distance", "subSmiley2Distance", "subSmiley3Distance",
                          "posDev", "preSum", "preMax", "preMin", "preSd", "preMean",
                          "preDiff", "preGrad", "preIntegrals")
  
  bfs.feature.names <- c("Smiley count", "Collision frequency",
                         "Minimal virtual distance smiley 1", "Minimal virtual distance smiley 2", "Minimal virtual distance smiley 3", "Minimal virtual distance smiley 4", "Virtual deviation of ideal flying route", 
                         "Normalized pressure", "Normalized pressure", "Normalized pressure", "Normalized pressure", "Normalized pressure",
                         "Pressure difference", "Pressure gradient", "Pressure-time integral")
  
  bfs.feature.units <- c("(n)", "(n)", "(arb. unit)", "(arb. unit)", "(arb. unit)", "(arb. unit)",
                          "(arb. unit)", "(arb. unit)", "(arb. unit)", "(arb. unit)", "(arb. unit)", "(arb. unit)",
                          "(arb. unit)", "(arb. unit)", "(arb. unit)")
  
  fbs.feature.labels <- c("anticipationTime",  "errorTime", "reactionTime", 
                          "preReaIntegralsLeft", "preReaIntegralsRight",
                          "preSumLeft", "preMaxLeft", "preMinLeft", "preSdLeft", "preMeanLeft",
                          "preDiffSumLeft", "preDiffMaxLeft", "preDiffMinLeft", "preDiffSdLeft", "preDiffMeanLeft",
                          "preGradAbsSumLeft", "preGradSumLeft", "preGradMaxLeft", "preGradMinLeft", "preGradSdLeft", "preGradMeanLeft",
                          "preExeIntegralsLeft",
                          "preSumRight", "preMaxRight", "preMinRight", "preSdRight", "preMeanRight",
                          "preDiffSumRight", "preDiffMaxRight", "preDiffMinRight", "preDiffSdRight", "preDiffMeanRight",
                          "preGradAbsSumRight", "preGradSumRight", "preGradMaxRight", "preGradMinRight", "preGradSdRight", "preGradMeanRight",
                          "preExeIntegralsRight", "succeedTasksLeft", "succeedTasksRight") 
  fbs.feature.names <- c("Anticipation time", "Time outside ideal pressure zone", "Reaction time",
                         "Pressure-time integral of the reaction phase", "Pressure-time integral of the reaction phase", 
                         "Normalized pressure", "Normalized pressure", "Normalized pressure", "Normalized pressure", "Normalized pressure",
                         "Pressure difference", "Pressure difference", "Pressure difference", "Pressure difference", "Pressure difference",
                         "Pressure gradient", "Pressure gradient", "Pressure gradient", "Pressure gradient", "Pressure gradient", "Pressure gradient", 
                         "Pressure-time integral of the execution phase",
                         "Normalized pressure", "Normalized pressure", "Normalized pressure", "Normalized pressure", "Normalized pressure",
                         "Pressure difference", "Pressure difference", "Pressure difference", "Pressure difference", "Pressure difference",
                         "Pressure gradient", "Pressure gradient", "Pressure gradient", "Pressure gradient", "Pressure gradient", "Pressure gradient", 
                         "Pressure-time integral of the execution phase", "Successful tasks", "Successful tasks")
  fbs.feature.units <- c("(sec)", "(sec)", "(sec)",
                         "(arb. unit)", "(arb. unit)", 
                         "(arb. unit)", "(arb. unit)", "(arb. unit)", "(arb. unit)", "(arb. unit)",
                         "(arb. unit)", "(arb. unit)", "(arb. unit)", "(arb. unit)", "(arb. unit)",
                         "(arb. unit)", "(arb. unit)", "(arb. unit)", "(arb. unit)", "(arb. unit)", "(arb. unit)", 
                         "(arb. unit)",
                         "(arb. unit)", "(arb. unit)", "(arb. unit)", "(arb. unit)", "(arb. unit)",
                         "(arb. unit)", "(arb. unit)", "(arb. unit)", "(arb. unit)", "(arb. unit)",
                         "(arb. unit)", "(arb. unit)", "(arb. unit)", "(arb. unit)", "(arb. unit)", "(arb. unit)", 
                         "(arb. unit)", "(n)", "(n)")
  
  jps.feature.labels <- c("attempts", "preDevi", "anticipationTime", "executionTime", 
                          "executionPre", "pti", "preGrad", "preDiff", "preIntegrals", "AntiTimeDiff")
  
  jps.feature.names <- c("Attempts", "Deviation from ideal pressure", "Anticipation time", "Execution time", 
                         "Mean pressure of execution phase", 
                         "Pressure-time integral", 
                         "Pressure gradient", 
                         "Pressure difference", "Pressure-time integral", "Difference of anticipation times (s)")
  jps.feature.units <- c("(n)", "(arb. unit)", "(sec)", "(sec)", 
                         "(arb. unit)", 
                         "(arb. unit)", 
                         "(arb. unit)", 
                         "(arb. unit)", "(arb. unit)", "(sec)")
  
  cal.feature.labels <- c("CalLeftMtkMin", "CalLeftMtkMax", "CalRightMtkMin", "CalRightMtkMax", 
                          "CalLeftCalMin", "CalLeftCalMax", "CalRightCalMin", "CalRightCalMax")
  cal.feature.names <- c("Minimum pressure on the left forefoot", "Maximum pressure on the left forefoot",
                         "Minimum pressure on the right forefoot", "Maximum pressure on the right forefoot",
                         "Minimum pressure on the left heel", "Maximum pressure on the left heel",
                         "Minimum pressure on the right heel", "Maximum pressure on the right heel")
  cal.feature.units <- c("(arb. unit)", "(arb. unit)", "(arb. unit)", "(arb. unit)", 
                          "(arb. unit)", "(arb. unit)", "(arb. unit)", "(arb. unit)")       
  
  radar.feature.labels <- c("Skill", "Skill.AC", "Skill.BFL", "Skill.BFR", "Skill.CP", "Skill.IJ",
                            "Reaction", "Reaction.Finger", "Reaction.AC1", "Reaction.AC2", "Reaction.CP", 
                            "Sensation", "Sensation.AC", "Sensation.BFL", "Sensation.BFR", "Sensation.CP", "Sensation.IJ", "Strength", "Strength.AC",     
                            "Strength.BFL", "Strength.BFR", "Strength.CP", "Strength.IJ", "Endurance", "Endurance.AC", "Endurance.CP", "Endurance.IJ",
                            "Balance",  "Balance.Two.Feet", "Balance.One.Foot")
  radar.feature.names <- c("Skill", "Skill (AC)", "Skill (BFL)", "Skill (BFR)", "Skill (CP)", "Skill (IJ)",
                           "Reaction", "Reaction (finger)", "Reaction (AC1)", "Reaction (AC2)", "Reaction (CP)", 
                           "Sensation", "Sensation (AC)", "Sensation (BFL)", "Sensation (BFR)", "Sensation (CP)", "Sensation (IJ)", "Strength", "Strength (AC)",     
                           "Strength (BFL)", "Strength (BFR)", "Strength (CP)", "Strength (IJ)", "Endurance", "Endurance (AC)", "Endurance (CP)", "Endurance (IJ)",
                           "Balance",  "Balance (both feet)", "Balance (single foot)")
  radar.feature.units <- rep("(arb. unit)", length(radar.feature.names))   
  
  
  numbers10 <- c("1", "2", "3", "4", "5", "6", "7", "8", "9", "0")
  numbers20 <- c( "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "20")
  labels <- c()
  names <- c()
  groups <- c()
  foot.sides <- c()
  games <- c()
  feature.units <- c()
  calculations <- c()
  is.game.features.label <- c()
  for (feature.name in features.vector) {
    game <- ""
    name <- ""
    group <- ""
    foot.side <- substr(feature.name, 3, 3)
    unit <- ""
    calculation <- ""
    is.game.feature <- T
    
    if (grepl("ACL", feature.name, fixed = T)||grepl("ACR", feature.name, fixed = T)){
      game <- "AC"
      for (afs.label.index in 1:length(afs.feature.labels)) {
        afs.label <- afs.feature.labels[afs.label.index]
        if (grepl(afs.label, feature.name, fixed = T)){
          name <- afs.feature.names[afs.label.index]
          unit <- afs.feature.units[afs.label.index]
        }
      }
      
      if (grepl("TC", feature.name, fixed = T)){
        ## a feature of the task combination
        group.str1 <- substr(feature.name, 3, 3)
        group.str2 <- substr(feature.name, nchar(feature.name), nchar(feature.name))
        group <- paste("TC", group.str1, group.str2, sep = "")
      } else {
        ## a feature of the task
        group.str1 <- substr(feature.name, nchar(feature.name)-1, nchar(feature.name))
        if (!group.str1 %in% seq(10, 100, 1)){
          group.str1 <- substr(feature.name, nchar(feature.name), nchar(feature.name))
        }
        group <- paste("Task", group.str1, sep = "")
      }
    } else if (grepl("BFL", feature.name, fixed = T)||grepl("BFR", feature.name, fixed = T)){
      game <- "BF"
      for (bfs.label.index in 1:length(bfs.feature.labels)) {
        bfs.label <- bfs.feature.labels[bfs.label.index]
        if (grepl(bfs.label, feature.name, fixed = T)){
          name <- bfs.feature.names[bfs.label.index]
          unit <- bfs.feature.units[bfs.label.index]
        }
      }
      
      if (grepl("TC", feature.name, fixed = T)){
        ## a feature of the task combination
        group.str1 <- substr(feature.name, 3, 3)
        group.str2 <- substr(feature.name, nchar(feature.name), nchar(feature.name))
        group <- paste("TC", group.str1, group.str2, sep = "")
      } else {
        ## a feature of the task
        group.str1 <- substr(feature.name, nchar(feature.name)-1, nchar(feature.name))
        if (!group.str1 %in% seq(10, 100, 1)){
          group.str1 <- substr(feature.name, nchar(feature.name), nchar(feature.name))
        }
        group <- paste("Task", group.str1, sep = "")
      }
    } else if (grepl("CPL", feature.name, fixed = T)||grepl("CPR", feature.name, fixed = T)){
      game <- "CP"
      for (fbs.label.index in 1:length(fbs.feature.labels)) {
        fbs.label <- fbs.feature.labels[fbs.label.index]
        if (grepl(fbs.label, feature.name, fixed = T)){
          name <- fbs.feature.names[fbs.label.index]
          unit <- fbs.feature.units[fbs.label.index]
        }
      }
      
      if (grepl("succeedTasksLeft", feature.name, fixed = T)){
        group <- paste("TCL", sep = "")
      } else if (grepl("succeedTasksRight", feature.name, fixed = T)){
        group <- paste("TCR", sep = "")
      } else if (grepl("TCB", feature.name, fixed = T)){
        ## a feature of the task combination
        group.str1 <- substr(feature.name, 3, 3)
        group.str2 <- substr(feature.name, nchar(feature.name), nchar(feature.name))
        group <- paste("TCB", group.str2, sep = "")
      } else if (grepl("TC", feature.name, fixed = T)){
        ## a feature of the task combination
        group.str1 <- substr(feature.name, 3, 3)
        group.str2 <- substr(feature.name, nchar(feature.name), nchar(feature.name))
        group <- paste("TC", group.str1, group.str2, sep = "")
      } else {
        ## a feature of the task
        group.str1 <- substr(feature.name, nchar(feature.name)-1, nchar(feature.name))
        if (!group.str1 %in% seq(10, 100, 1)){
          group.str1 <- substr(feature.name, nchar(feature.name), nchar(feature.name))
        }
        group <- paste("Task", group.str1, sep = "")
      }
      
    } else if (grepl("IJL", feature.name, fixed = T)||grepl("IJR", feature.name, fixed = T)){
      game <- "IJ"
      for (jps.label.index in 1:length(jps.feature.labels)) {
        jps.label <- jps.feature.labels[jps.label.index]
        if (grepl(jps.label, feature.name, fixed = T)){
          name <- jps.feature.names[jps.label.index]
          unit <- jps.feature.units[jps.label.index]
        }
      }
      
      if (grepl("AntiTimeDiff", feature.name, fixed = T)){
        group <- ""
      } else if (grepl("TC", feature.name, fixed = T)){
        ## a feature of the task combination
        group.str1 <- substr(feature.name, 3, 3)
        group.str2 <- substr(feature.name, nchar(feature.name), nchar(feature.name))
        group <- paste("TC", group.str1, group.str2, sep = "")
      } else {
        ## a feature of the task
        group.str1 <- substr(feature.name, nchar(feature.name)-1, nchar(feature.name))
        if (!group.str1 %in% seq(10, 100, 1)){
          group.str1 <- substr(feature.name, nchar(feature.name), nchar(feature.name))
        }
        group <- paste("Task", group.str1, sep = "")
      }
    } else if (grepl("CalLeft", feature.name, fixed = T)||grepl("CalRight", feature.name, fixed = T)){
      game <- "CAL"
      for (cal.label.index in 1:length(cal.feature.labels)) {
        cal.label <- cal.feature.labels[cal.label.index]
        if (grepl(cal.label, feature.name, fixed = T)){
          name <- cal.feature.names[cal.label.index]
          unit <- cal.feature.units[cal.label.index]
        }
      }
      
      group <- "2"
    } else if (grepl("Radar", feature.name, fixed = T)||grepl("radar", feature.name, fixed = T)){
      game <- "Hypothesis"
      for (radar.label.index in 1:length(radar.feature.labels)) {
        radar.label <- radar.feature.labels[radar.label.index]
        if (grepl(radar.label, feature.name, fixed = T)){
          name <- radar.feature.names[radar.label.index]
          unit <- radar.feature.units[radar.label.index]
        }
      }
      
      group <- "2"
    } else {
      print(paste("Unable to find info for the feature:", feature.name))
      is.game.feature <- F
    }
    
    if (grepl("Sum", feature.name, fixed = T)){
      calculation <- "sum"
    } else if (grepl("Max", feature.name, fixed = T)){
      calculation <- "max"
    } else if (grepl("Min", feature.name, fixed = T)){
      calculation <- "min"
    } else if (grepl("Sd", feature.name, fixed = T)){
      calculation <- "SD"
    } else if (grepl("Mean", feature.name, fixed = T)){
      calculation <- "mean"
    } else {
      calculation <- ""
    }
    
    labels <- c(labels, feature.name)
    names <- c(names, name)
    groups <- c(groups, group)
    games <- c(games, game)
    foot.sides <- c(foot.sides, foot.side)
    feature.units <- c(feature.units, unit)
    calculations <- c(calculations, calculation)
    is.game.features.label <- c(is.game.features.label, is.game.feature)
  }
  output.table <- data.frame(Label = labels, Name = names, Group = groups, Game = games, Foot=foot.sides, Calculation= calculations, Unit=feature.units, Is.Game.Feature = is.game.features.label)
  #print(output.table)
  return(output.table)
  
}


## do intergroup difference test (non parametric test)
# input.data: data.frame with features and class label (named "Class" on the last column)
# return.signif.vars: return full result by FALSE, and only significant variables if TRUE
intergroup.difference.test <- function(input.data, return.signif.vars){
  label.position <- match("class",  names(input.data))
  if (!is.na(label.position)){
    names(input.data)[label.position] <- "Class"
    print("fix the label column name from class to Class...")
  }

  group.different.features <- c()
  game.feature.names <- names(input.data)[1:(ncol(input.data)-1)]
  for (feature.index in 1: length(game.feature.names) ) {
    feature <- game.feature.names[feature.index]
    game.name <- substr(feature, nchar(feature)-2, nchar(feature))
    feature.factor <- factor(input.data[[feature]])
    if (length(levels(feature.factor))==2){
      print(paste("1. --- Perform chi-square test on feature: ", feature))
      test.table <- table(input.data$Class, feature.factor)
      chisq.test.result <- chisq.test(test.table)
      result <- c(game.name, feature, chisq.test.result$statistic, chisq.test.result$p.value, "Chi-square test", FALSE)
    } else {
      normality.test.data <- input.data[[feature]]
      normality.test <- shapiro.test(normality.test.data)
      #print(normality.test)
      #print(paste("2. --- Normality test of ", feature, "-----------"))

      diff.test.df <- subset(input.data, select=c(feature, "Class"))
      names(diff.test.df)[1] <- "Score"
      if (normality.test$p.value > 0.05){
        print(paste("3. --- T.test of parametic variable:", feature, "-------"))
        t.test.result <- t.test(Score~Class, data = diff.test.df)
        #print(t.test.result)
        result <- c(game.name, feature, t.test.result$statistic, t.test.result$p.value, "T-test", TRUE)
        # do t-test
      } else {
        print(paste("4. --- Wilcox.test of non-parametic variable:", feature, "-------"))
        # do wilcox.test
        wilcox.test.result <- wilcox.test(Score~Class, data = diff.test.df)
        #print(wilcox.test.result)
        result <- c(game.name, feature, wilcox.test.result$statistic, wilcox.test.result$p.value, "Wilcoxon test", FALSE)
      }
    }
    if (length(group.different.features) == 0){
      group.different.features <- result
    } else {
      group.different.features <- rbind(group.different.features, result)
    }
  }
  group.different.features <- as.data.frame(group.different.features)
  names(group.different.features) <- c("Game", "Feature", "Statistic", "P.Value", "Method", "Normality")

  feature.info.df <- get.feature.full.info(group.different.features$Feature)

  group.different.features.final <- cbind(group.different.features$Feature, feature.info.df, group.different.features[, 3:6])
  names(group.different.features.final)[1] <- "Feature"
  group.different.features.final$P.Value <- as.numeric(group.different.features.final$P.Value)
  group.different.features.final$Statistic <- as.numeric(group.different.features.final$Statistic)
  group.different.features.final$Signif <- sapply(group.different.features.final$P.Value, convert.p.value.to.signif)

  if (return.signif.vars){
    ## only return variables with a significant difference
    group.different.features.final <- subset(group.different.features.final, P.Value < 0.05)
  }
  group.different.features.final$P.Value <- specify_decimal(as.numeric(group.different.features.final$P.Value), 5)
  group.different.features.final$Statistic <- specify_decimal(as.numeric(group.different.features.final$Statistic), 2)
  rownames(group.different.features.final) <- seq(1, nrow(group.different.features.final), 1)
  return(group.different.features.final)
}

## a function that compares the mean of more than two groups
# Categorical Data Variables is excluded!!
# the last column is the label, which should be named Class/class
# only.return.signif.vars: return full result by FALSE, and only significant variables if TRUE
moregroups.difference.test <- function(input.data, only.return.signif.vars){
  label.position <- match("class",  names(input.data))
  if (!is.na(label.position)){
    names(input.data)[label.position] <- "Class"
    print("fix the label column name from class to Class...")
  }

  group.different.features <- c()
  game.feature.names <- names(input.data)[-label.position]
  test.results.full <- c()
  for (feature.index in 1: length(game.feature.names) ) {
    #for (feature.index in 1: 100 ) {
    feature <- game.feature.names[feature.index]
    game.name <- substr(feature, nchar(feature)-2, nchar(feature))
    feature.factor <- factor(input.data[[feature]])

    if (length(levels(feature.factor))==2){
      print(paste(feature.index, "   1. --- Perform chi-square test on feature: ", feature))
      # test.table <- table(input.data$Class, feature.factor)
      # chisq.test.result <- chisq.test(test.table)
      # chisq.test.p <- chisq.test.result$p.value
      # if(is.na(chisq.test.p))
      # {
      #   test.result <- c(feature, game.name, chisq.test.p, 0, 0, convert.p.value.to.signif(chisq.test.p), "Chi-square test", FALSE)
      # }
    } else {
      normality.test.data <- input.data[[feature]]
      normality.test <- shapiro.test(normality.test.data)

      diff.test.df <- subset(input.data, select=c(feature, "Class"))
      names(diff.test.df)[1] <- "Score"
      if (normality.test$p.value > 0.05){
        print(paste(feature.index, "   3. --- anova test of parametic variable:", feature, "-------"))
        test.result <- compare_means(Score ~ Class,  data = diff.test.df, method="anova", TRUE)
        #print(test.result)
      } else {
        print(paste(feature.index, "   4. --- kruskal test of non-parametic variable:", feature, "-------"))
        test.result <- compare_means(Score ~ Class,  data = diff.test.df, method="kruskal.test", FALSE)
        #print(test.result)
      }

      if (length(test.results.full)==0){
        test.results.full <- cbind(feature, game.name, test.result)
      } else {
        test.results.full <- rbind(test.results.full,  cbind(feature, game.name, test.result))
      }
    }
  }

  if (only.return.signif.vars){
    signif.diff.features <- subset(test.results.full, p.signif != "ns")
    return(signif.diff.features)
  } else {
    return(test.results.full)
  }
}

# build a function to summarize the table by the "Group" automatically according to the variable type
# report.table: should have at least two columns, the name of the first column should be "Group"!!!
# effective.number: an integer value to define how many decimals should be calculated.
build.report <- function(report.table, effective.number){
  # build a function to summarize the table by the "Group" automatically according to the variable type
  report.column.count <- length(levels(report.table$Group))
  report.column.names <- levels(report.table$Group)
  report.row.names <- c("")
  report.first.column <- c("")
  
  # determine the size, column names, and row names of report data frame according to levels of the "Group" and number of variables
  for (col.index in 2:ncol(report.table)) {
    var <- report.table[, col.index]
    var.name <- names(report.table)[col.index]
    if (is.numeric(var)){
      report.row.names <- c(report.row.names, var.name)
      report.first.column <- c(report.first.column, var.name)
    } else if (is.factor(var)){
      report.row.names <- c(report.row.names, var.name)
      report.first.column <- c(report.first.column, var.name)
      var.levels <- levels(var)
      report.row.names <- c(report.row.names, paste(var.name, var.levels) )
      report.first.column <- c(report.first.column, var.levels)
    }
  }
  report.column.names
  
  # add last row to show some comments of the calculation
  report.row.names <- c(report.row.names, "comment")
  report.first.column <- c(report.first.column, "comment")
  
  # create the report data frame
  report <- data.frame(matrix(nrow = length(report.row.names), ncol = length(report.column.names)))
  names(report) <- report.column.names
  rownames(report) <- report.row.names
  
  # summarize different variable and write mean (sd) or freq (prop) values in the table
  for (group.index in 1: length(levels(report.table$Group))) {
    group.name <- levels(report.table$Group)[group.index]
    group.data <- subset(report.table, Group == group.name)
    summary.comments <- "NAs:"
    
    reprot.row.index <- 1
    samples.n <- nrow(group.data)
    report[reprot.row.index, group.index] <- paste("(n = ", samples.n, ")", sep = "") 
    reprot.row.index <- reprot.row.index + 1
    
    for (col.index in 2:ncol(group.data)) {
      var <- group.data[, col.index]
      var.name <- names(group.data)[col.index]
      
      count.NAs <- sum(is.na(var))
      if (count.NAs>0){
        print(paste("--------!!!--- WARNING: ", var.name, "has", count.NAs, " NA values!! These data were excluded ----!!!----"))
        print(var)
        var <- var[!is.na(var)]
        print(paste("Remove NAs from ", var.name))
        print(var)
        summary.comments <- paste(summary.comments, var.name, "(n=", count.NAs, ")")
      }
      
      # print(paste(group.name, var.name, "Col.index ->", reprot.row.index))
      if (is.numeric(var)){
        output <- paste(specify_decimal(mean(var), effective.number), " (", specify_decimal(sd(var), effective.number), ")", sep = "")
        report[reprot.row.index, group.index] <- output
        reprot.row.index <- reprot.row.index + 1
      } else if (is.factor(var)){
        var.levels <- levels(var)
        var.levels.n <- length(var.levels)
        
        # add the var headline with empty value
        output <- ""
        report[reprot.row.index, group.index] <- output
        reprot.row.index <- reprot.row.index + 1
        
        # add summary of different levels
        var.length <- length(var)
        var.summmary <- summary(var)
        for (var.level.freq in var.summmary) {
          if (var.level.freq==0){
            var.level.output <- "-"
          } else {
            var.level.prop <- var.level.freq/var.length
            var.level.output <- paste(var.level.freq, " (",  specify_decimal(var.level.prop*100, effective.number), "%)", sep = "")          
          }
          report[reprot.row.index, group.index] <- var.level.output
          reprot.row.index <- reprot.row.index + 1
        }
      } else {
        print(paste("Empty parameter:", var.name))
      }
    }
    
    # add comments to the end of the column
    report[reprot.row.index, group.index] <- summary.comments
  }
  
  # add the variable name on the first column
  if (length(report.first.column) != nrow(report)){
    print(paste("Error: length of rownames (", length(report.first.column), ") are different to the size of report table (", nrow(report), ")"))
    return(report)
  } else {
    final.report <- cbind(report.first.column, report)
    names(final.report)[1] <- "Feature"
    return(final.report)
  }
}


# build a function to summarize the table by the "Group" automatically according to the variable type
# report.table: should have at least two columns, the name of the first column should be "Group"!!!
# for numeric variables: report means and standard deviations (SD) for normally distributed data and medians and interquartile ranges (IQR) for non-normally distributed data.
# for categorical variables: report number and proportion of each level
# effective.number: an integer value to define how many decimals should be calculated.
new.build.report <- function(report.table, effective.number, only.count){
  if(missing(only.count)) {
    only.count <- F
  }
  
  # build a function to summarize the table by the "Group" automatically according to the variable type
  report.column.count <- length(levels(report.table$Group))
  report.column.names <- c()
  for (col.name in levels(report.table$Group)) {
    report.column.names <- c(report.column.names, col.name, "Property")
  }
  
  #report.column.names <- levels(report.table$Group)
  report.row.names <- c("")
  report.first.column <- c("")
  
  # determine the size, column names, and row names of report data frame according to levels of the "Group" and number of variables
  for (col.index in 2:ncol(report.table)) {
    var <- report.table[, col.index]
    var.name <- names(report.table)[col.index]
    if (is.numeric(var)){
      report.row.names <- c(report.row.names, var.name)
      report.first.column <- c(report.first.column, var.name)
    } else if (is.factor(var)){
      report.row.names <- c(report.row.names, var.name)
      report.first.column <- c(report.first.column, var.name)
      var.levels <- levels(var)
      report.row.names <- c(report.row.names, paste(var.name, var.levels) )
      report.first.column <- c(report.first.column, var.levels)
    }
  }
  report.column.names
  
  # add last row to show some comments of the calculation
  report.row.names <- c(report.row.names, "comment")
  report.first.column <- c(report.first.column, "comment")
  
  # create the report data frame
  report <- data.frame(matrix(nrow = length(report.row.names), ncol = length(report.column.names)))
  names(report) <- report.column.names
  rownames(report) <- report.row.names
  
  # summarize different variable and write mean (sd) or freq (prop) values in the table
  for (group.index in 1: length(levels(report.table$Group))) {
    group.name <- levels(report.table$Group)[group.index]
    group.data <- subset(report.table, Group == group.name)
    summary.comments <- "NAs:"
    
    reprot.row.index <- 1
    samples.n <- nrow(group.data)
    report[reprot.row.index, group.index*2-1] <- paste("(n = ", samples.n, ")", sep = "") 
    report[reprot.row.index, group.index*2] <- "" 
    reprot.row.index <- reprot.row.index + 1
    
    for (col.index in 2:ncol(group.data)) {
      var <- group.data[, col.index]
      var.name <- names(group.data)[col.index]
      
      count.NAs <- sum(is.na(var))
      if (count.NAs>0){
        print(paste("--------!!!--- WARNING: ", var.name, "has", count.NAs, " NA values!! These data were excluded ----!!!----"))
        print(var)
        var <- var[!is.na(var)]
        print(paste("Remove NAs from ", var.name))
        print(var)
        summary.comments <- paste(summary.comments, var.name, "(n=", count.NAs, ")")
      }
      
      # print(paste(group.name, var.name, "Col.index ->", reprot.row.index))
      if (is.numeric(var)&& length(unique(var))!=1){
        ## Shapiro-Wilk test, when and only when both p-values are greater than 0.05, indicates that the data are normally distributed
        sw.test <- tapply(report.table[, col.index], report.table$Group, shapiro.test)
        print(paste("------------", var.name , "----------"))
        print(sw.test)
        p.larger.than.005 <- T
        for (n.class in 1:length(sw.test)) {
          pValue <- sw.test[[n.class]][2]$p.value
          if (pValue<=0.05){
            p.larger.than.005 <- F
            break()
          }
        }
        normal.distributed <- p.larger.than.005
        #Report means and standard deviations (SD) for normally distributed data and medians and interquartile ranges (IQR) for non-normally distributed data.
        if (normal.distributed){
          # use mean (sd) for variables with normal distribution
          output <- paste(specify_decimal(mean(var), effective.number), " (", specify_decimal(sd(var), effective.number), ")", sep = "")
          report[reprot.row.index, group.index*2-1] <- output
          report[reprot.row.index, group.index*2] <- "mean(SD)"
        } else {
          # use median (IQR range) for variables without normal distribution
          output <- paste(specify_decimal(median(var), effective.number), " (", specify_decimal(IQR(var), effective.number), ")", sep = "")
          report[reprot.row.index, group.index*2-1] <- output
          report[reprot.row.index, group.index*2] <- "median(IQR)"
        }
        
        #report[reprot.row.index, group.index] <- output
        reprot.row.index <- reprot.row.index + 1
      } else if (is.factor(var)){
        var.levels <- levels(var)
        var.levels.n <- length(var.levels)
        
        # add the var headline with empty value
        output <- ""
        report[reprot.row.index, group.index*2-1] <- output
        report[reprot.row.index, group.index*2] <- "n(%)"
        reprot.row.index <- reprot.row.index + 1
        
        # add summary of different levels
        var.length <- length(var)
        var.summmary <- summary(var)
        for (var.level.freq in var.summmary) {
          if (var.level.freq==0){
            var.level.output <- "-"
          } else {
            if (only.count){
              var.level.output <- paste(var.level.freq, sep = "")          
            } else {
              var.level.prop <- var.level.freq/var.length
              var.level.output <- paste(var.level.freq, " (",  specify_decimal(var.level.prop*100, effective.number), "%)", sep = "")          
            }
          }
          report[reprot.row.index, group.index*2-1] <- var.level.output
          report[reprot.row.index, group.index*2] <- ""
          reprot.row.index <- reprot.row.index + 1
        }
      } else {
        print(paste("Empty parameter:", var.name))
      }
      
    }
    
    # add comments to the end of the column
    report[reprot.row.index, group.index*2-1] <- summary.comments
    report[reprot.row.index, group.index*2] <- ""
  }
  
  # add the variable name on the first column
  if (length(report.first.column) != nrow(report)){
    print(paste("Error: length of rownames (", length(report.first.column), ") are different to the size of report table (", nrow(report), ")"))
    return(report)
  } else {
    final.report <- cbind(report.first.column, report)
    names(final.report)[1] <- "VariableName"
    return(final.report)
  }
}



## draw a series of box plots to overview the selected features
draw.features.with.boxplot.groups <- function(group.test.data, selected.features, my_comparisons, save.path, save.name){
  font.size <- 16
  title.font.size <- 16
  fig.title <- "Example game features and their differences between CT and DPN subgroups"
  fig.caption <- paste(
    "CON: controls; CON1: controls aged 18-30; CON2: controls aged 31-50; CON3: controls aged 51-80;","\n", 
    "DPN: diabetic peripheral neuropathy; DPN1: mild DPN (NDS: 1-5); DPN2: moderate DPN (NDS: 6-8); DPN3: severe DPN (NDS: 9-10);","\n", 
    "Differences between groups were measured by t-test, Mann-Whitney U test or chi-square test as appropriate. ", "\n",
    "Significance levels: ns (p > 0.05), * (p 0.01-0.05), ** (p 0.001-0.01) , *** (p 0.0001-0.001), **** (p < 0.0001)."
  )
  selected.data <- subset(group.test.data, select=c("Group", "Class", selected.features))
  selected.data.melt <- reshape2::melt(selected.data, id.vars = c("Group", "Class"), variable.name = "Feature", value.name = "Value")
  
  combined.fig <-
    ggplot(data = selected.data.melt, aes(Group, Value, color=Class))+
    geom_boxplot()+
    facet_wrap(.~Feature, scales = "free", nrow = 2)+
    stat_compare_means(label="p.signif", method="wilcox.test",
                       comparisons = my_comparisons)+
    labs(title = fig.title, caption = fig.caption, color="Group")+
    xlab("Groups")+
    ylab("Feature Value")+
    # scale_x_discrete(labels=c("CT (Age: 18-30)", "CT (Age: 31-50)", "CT (Age: 51-80)", "CT", 
    #                           "DPN", "Mild DPN", "Moderate DPN", "Severe DPN"))+
    scale_x_discrete(labels=c("CON", "CON1", "CON2", "CON3", 
                              "DPN", "DPN1", "DPN2", "DPN3"))+
    theme(
      text = element_text(size = title.font.size), # set the text size generally
      plot.title = element_text(hjust = 0, size = title.font.size, face = "bold"),    # Center title position and size
      plot.subtitle = element_text(hjust = 1, size = font.size),            # Center subtitle
      plot.caption = element_text(size = font.size), 
      legend.text = element_text(size = font.size),
      # axis.text.x = element_text(size = 10, angle = 75, vjust = 0.55),
      axis.text.x = element_text(size = font.size, angle = 45, vjust = 0.65),
      strip.text.x =  element_text(size = title.font.size, angle = 0), # 修改分面标签横向的文本大小
      strip.text.y =  element_text(size = title.font.size, angle = 0), # 修改分面标签纵向的文本大小
      legend.position = "top"
    )
  
  
  # combined.fig.save.name <- paste("groups_difference_selected_features.jpeg", sep = "")
  # combined.fig.save.path <- paste(workspace.root, "IQ-Game_Paper1", "groups_difference_test", sep = "/")
  ggsave(filename = save.name, plot = combined.fig, path=save.path, width = 48, height = 32, units = "cm", dpi = 400)
  return(combined.fig)
}

## function to convert p value to significant levels presented by stars
convert.p.value.to.signif <- function(P){
  if(!is.na(P)){
    if (P < 0.0001){
      return("****")
    } else if (P < 0.001){
      return("***")
    } else if (P < 0.01){
      return("**")
    } else if (P < 0.05){
      return("*")
    } else {
      return("ns")
    }
  } else {
    return("na")
  }
}

format.p.value <- function(p){
  formatted.p <- ""
  if (is.na(p) || p=="" || p=="NA"){
    print("Skip format this p value")
  } else {
    formatted.p <- p
    if (!is.numeric(p)){
      formatted.p <- as.numeric(p)
    }
    if (formatted.p>0.99){
      formatted.p <- 0.99
    } else if (formatted.p>0.05){
      formatted.p <- specify_decimal(formatted.p, 2)
    } 
    else if (formatted.p>0.01){
      formatted.p <- specify_decimal(formatted.p, 3)
    } else if (formatted.p>0.001){
      formatted.p <- specify_decimal(formatted.p, 3)
    } else if (formatted.p>0.0001){
      formatted.p <- specify_decimal(formatted.p, 4)
    } else {
      formatted.p <- "<.0001"
    }
  }
  return(formatted.p)
}



convert.official.game.name <- function(game.name){
  if (game.name=="Afs"){
    return("Apple-Catch")
  } else if (game.name=="Bfs"){
    return("Balloon-Flying")
  } else if (game.name=="Fbs"){
    return("Cross-Pressure")
  } else if (game.name=="Jps"){
    return("Island-Jump")
  } else {
    return("")
  }  
}

convert.official.short.game.name <- function(game.name){
  if (game.name=="Afs"){
    return("AC")
  } else if (game.name=="Bfs"){
    return("BF")
  } else if (game.name=="Fbs"){
    return("CP")
  } else if (game.name=="Jps"){
    return("IJ")
  } else {
    return("")
  }  
}

## replace NA values with column mean values 
## please make sure the input data contains only numeric values
replace.NA.with.coulmn.mean <- function(input.df){
  if (sum(is.na(input.df))>0){
    print(paste("Found", sum(is.na(input.df)), "values in the data frame"))
    for(i in 1:ncol(input.df)){
      for(j in 1:nrow(input.df)){
        if (is.na(input.df[j,i])){
          input.df[j,i] <- mean(input.df[,i], na.rm = TRUE)
          print(paste("Replace a [", names(input.df)[i], "] value on the row [", j, "] with the column mean", mean(input.df[,i], na.rm = TRUE)))
        }
      }
    }
  } else {
    print("No NA values were found!")
  }
  return(input.df)
}

find.NA.positions <- function(input.df){
  if (sum(is.na(input.df))>0){
    print(paste("Found", sum(is.na(input.df)), "values in the data frame"))
    for(i in 1:ncol(input.df)){
      for(j in 1:nrow(input.df)){
        if (is.na(input.df[j,i])){
          print(paste("found a [", names(input.df)[i], "] value on the row [", j, "] with the NA"))
        }
      }
    }
  } else {
    print("No NA values were found!")
  }
}

## A more general function is as follows where x is the number and k is the number of decimals to show. trimws removes any leading white space which can be useful if you have a vector of numbers.
specify_decimal <- function(x, k) trimws(format(round(x, k), nsmall=k))


#https://stackoverflow.com/questions/25124895/no-outliers-in-ggplot-boxplot-with-facet-wrap
calc_boxplot_stat <- function(x) {
  coef <- 1.5
  n <- sum(!is.na(x))
  # calculate quantiles
  stats <- quantile(x, probs = c(0.0, 0.25, 0.5, 0.75, 1.0))
  names(stats) <- c("ymin", "lower", "middle", "upper", "ymax")
  iqr <- diff(stats[c(2, 4)])
  # set whiskers
  outliers <- x < (stats[2] - coef * iqr) | x > (stats[4] + coef * iqr)
  if (any(outliers)) {
    stats[c(1, 5)] <- range(c(stats[2:4], x[!outliers]), na.rm = TRUE)
  }
  return(stats)
}


#https://stackoverflow.com/questions/25124895/no-outliers-in-ggplot-boxplot-with-facet-wrap
calc_stat <- function(x) {
  coef <- 1.5
  n <- sum(!is.na(x))
  # calculate quantiles
  stats <- quantile(x, probs = c(0.1, 0.25, 0.5, 0.75, 0.9))
  names(stats) <- c("ymin", "lower", "middle", "upper", "ymax")
  return(stats)
}

## 
replace.outliers.with.NAs <- function(df){
  if (is.data.frame(df)){
    ## replace outliers with NA
    for (feature.n in 1:ncol(df)) {
      if (is.numeric(df[, feature.n])){
        feature.name <- names(df)[feature.n]
        feature.col <- df[[feature.name]]
        ## calculate limits
        # ymin    lower   middle    upper     ymax 
        #-0.18600  0.03900  0.13050  0.19175  0.37400
        stat <- calc_boxplot_stat(feature.col) 
        ## replace outliers (>ymax or <ymin) with NA
        NA.values <- c()
        for (index in 1:length(feature.col)) {
          if (feature.col[index]<stat[1] || feature.col[index]>stat[5] ){
            NA.values <- c(NA.values, feature.col[index])
            feature.col[index] <- NA
          }
        }
        if (length(NA.values)>0){
          print(paste("------------ Replace outliers in the column", feature.name, "to NA:"))
          print(NA.values)     
        }
        df[[feature.name]] <- feature.col
      }
    }
    return(df) 
  } else if (is.vector(df)){
    feature.col <- df
    stat <- calc_boxplot_stat(feature.col) 
    ## replace outliers (>ymax or <ymin) with NA
    NA.values <- c()
    for (index in 1:length(feature.col)) {
      if (feature.col[index]<stat[1] || feature.col[index]>stat[5] ){
        NA.values <- c(NA.values, feature.col[index])
        feature.col[index] <- NA
      }
    }
    if (length(NA.values)>0){
      print(paste("------------ Replace outliers in the column", feature.name, "to NA:"))
      print(NA.values)     
    }
    return(feature.col)
  } else {
    print(paste("------------ unsupported data type", typeof(df), "to NA:"))
    return()
  }
}


## keep effective.number
keep.effective.number <- function(df, effective.number){
  for (variable in names(df)) {
    if (is.numeric(df[[variable]])){
      df[[variable]] <- round(df[[variable]], effective.number)
    }
  }
  return(df)
}



save.fig <- function(fig, save.path, save.name, save.width, save.height, save.dpi){
  if (!dir.exists(save.path)){
    dir.create(save.path, showWarnings = T)
  }
  ggsave(filename = save.name, plot = fig, path = save.path, width = save.width, height = save.height, units = "cm", dpi = save.dpi, limitsize = F)
}

## sort Panaritium to Inflammation
### change the alarm text to a factor without considering "Abnormality" and "False Alarm"
factor.alarm.type <- function(type, language){
  if (!is.character(type)){
    type <- as.character(type)
  }
  
  if ( grepl("Isch", type, fixed = T)){
    value <- 1
  } else if (grepl("Gicht", type, fixed = T)){
    value <- 2
  } else if (grepl("Hautentz", type, fixed = T)){
    value <- 5
  } else if (grepl("Entz", type, fixed = T)){
    value <- 3
  } else if (grepl("Panaritium", type, fixed = T)){
    value <- 3
  } else if (grepl("Trauma", type, fixed = T)){
    value <- 4
  } else if (grepl("arthose", type, fixed = T)){
    value <- 6
  } else if (grepl("unklar", type, fixed = T)){
    value <- 7
  } else {
    value <- 7
  }   
  
  if (language=="DE"){
    return(factor(value, levels = seq(1, 7, 1), labels = c("Ischämie", "Gichtattackt", "Entzündung",  "Trauma/Verletzung", "Hautentzündung", "Retropatellararthose", "Unklar"), ordered = T))
  } else {
    return(factor(value, levels = seq(1, 7, 1), labels = c("Ischemia", "Gout attack", "Inflammation",  "Trauma/Injury", "Skin infection", "Retropatellar arthrosis", "Unclassified"), ordered = T))
  }
}


factor.handedness <- function(input){
  if (input==1){
    value <- 1
  } else if (input==2){
    value <- 2
  } else if (input==0){
    value <- 0
  } else {
    print(paste("Error: wrong type by factoring handedness:", input))
  }
  
  return(factor(value, levels = seq(0, 2, 1), labels = c("Both", "Right", "Left"), ordered = T))
}


factor.customary.foot <- function(input){
  if (input==1){
    value <- 1
  } else if (input==2){
    value <- 2
  } else if (input==0){
    value <- 0
  } else {
    print(paste("Error: wrong type by factoring customary foot:", input))
  }
  
  return(factor(value, levels = seq(0, 2, 1), labels = c("Both", "Right", "Left"), ordered = T))
}


factor.sport.frequency <- function(input){
  if (input==1){
    value <- 1
  } else if (input==2){
    value <- 2
  } else if (input==0){
    value <- 0
  } else {
    print(paste("Error: wrong type by factoring sport frequency:", input))
  }
  
  return(factor(value, levels = seq(0, 2, 1), labels = c("kein", "unregelmäßig", "regelmäßig"), ordered = T))
}

##TODO to be edited
factor.sport.type <- function(input){
  if (is.numeric(input)){
    input <- as.character(input)
  }
  value <- 0
  if (grepl("1", input, fixed = T)){
    value <- 1
  } else if (grepl("2", input, fixed = T)){
    value <- 2
  } else if (grepl("3", input, fixed = T)){
    value <- 3
  } else if (!is.na(input)){
    value <- 3
  } else {
    value <- 0
    print(paste("Error: wrong type by factoring sport type:", input))
  }
  
  return(factor(value, levels = c(0, 1, 2, 3), labels = c("non", "biking", "jogging", "other"), ordered = T))
}

factor.app.experience <- function(input){
  if (is.numeric(input)){
    input <- as.character(input)
  }
  value <- 0
  if (grepl("2", input, fixed = T) || grepl("3", input, fixed = T) || grepl("4", input, fixed = T)){
    value <- 1
  } else if (grepl("0", input, fixed = T) || grepl("1", input, fixed = T) ){
    value <- 0
  } else {
    print(paste("Error: wrong type by factoring app.experience:", input))
  }
  
  return(factor(value, levels = c(0, 1), labels = c("Nein", "Ja"), ordered = T))
}

factor.driving.distance <- function(input){
  value <- 0
  if (is.na(input)){
    value <- 4
  } else {
    if (is.character(input)){
      if (nchar(input)>1){
        input=substr(input, 1, 1)
      } else {
        input <- as.numeric(input)
      }
    }
    if (input==0){
      ## no driving license
      value <- 0
    } else if (input==1 || input==2){
      value <- 1
    } else if (input==3){
      value <- 2
    } else if (input==4 || input==5 || input==6|| input==7){
      value <- 3
    } else {
      print(paste("Error: wrong type by factoring handedness:", input))
    } 
  }

  return(factor(value, levels = seq(0, 4, 1), labels = c("0km/year", "<5,000km/year", "5,000-\n10,000km/year", ">10,000km/year", "N/A"), ordered = T))
}

factor.age.to.three.levels <- function(input){
  if (input<=32){
    value <- 1
  } else if (input>=33 && input<=59){
    value <- 2
  } else if (input>=60){
    value <- 3
  } else {
    print(paste("Error: wrong type by factoring handedness:", input))
  }
  return(factor(value, levels = seq(1, 3, 1), labels = c("18-32", "33-59", "60-80"), ordered = T))
  
}

factor.words.fluency <- function(input){
  if (input==1){
    value <- 1
  } else if (input==0){
    value <- 0
  } else {
    print(paste("Error: wrong type by factoring words.fluency:", input))
  }
  
  return(factor(value, levels = c(0, 1), labels = c("<11", "\u2265 11"), ordered = T))
}

factor.words.memory <- function(input){
  if (input==1 || input==2 || input==3 || input==4 || input==0){
    value <- 0
  } else if (input==5){
    value <- 1
  } else {
    print(paste("Error: wrong type by factoring words.memory:", input))
  }
  
  return(factor(value, levels = c(0, 1), labels = c("0-4", "5"), ordered = T))
}

factor.education.years <- function(input){
  if (input<15){
    value <- 0
  } else if (input<=17.5 && input >=15){
    value <- 1
  }  else if (input<40 && input >17.5){
    value <- 2 
  } else {
    print(paste("Error: wrong type by factoring words.fluency:", input))
  }
  
  return(factor(value, levels = c(0, 1, 2), labels = c("<15", "15-17.5", ">17.5"), ordered = T))
}

## Bein-Halteversuch (0: normal, 1: rechts pathologisch, 2: links pathologisch)
factor.leg.hold.test <- function(input){
  return(factor(input, levels = c(0, 1, 2), labels = c("normal", "right pathological", "left pathological"), ordered = T))
}

## Knie-Hacke-Versuch (0: normal, 1: rechts pathologisch, 2: links pathologisch)
factor.knee.hook.test <- function(input){
  return(factor(input, levels = c(0, 1, 2), labels = c("normal", "right pathological", "left pathological"), ordered = T))
}

## Romberg Stand (0: normal, 1: Fallneigung)
factor.romberg.test <- function(input){
  return(factor(input, levels = c(0, 1), labels = c("Normal", "Fallneigung"), ordered = T))
}

## Unterberger Tretversuch (0: normal, 1: Abweichung nach rechts, 2: Abweichung nach links)
factor.unterberger.test <- function(input){
  return(factor(input, levels = c(0, 1, 2), labels = c("Normal", "Abweichung nach rechts", "Abweichung nach links"), ordered = T))
}


##### a function test different age thresholds and try to build an age-matched cohort
## df: a data frame contains the cohort data, especially age and group label columns
## age.col.name: name of the age column, e.g. "Age"
## group.col.name: name of the group label column, e.g. "PNGroup1"
## age.min.values: all min threshold to be tested, e.g. seq(50, 65, 1)
## age.max.values: all max threshold to be tested, e.g. seq(70, 95, 1)
## return: a list with three elements: (1) a ggplot of age distribution before matching; 
## return: a list with three elements: (2) a ggplot of age distribution after matching; 
## return: a list with three elements: (3) the age matched data frame 
age.matching.func <- function(df, age.col.name, group.col.name, age.min.values, age.max.values){
  text.font.size <- 24
  m.theme <- theme(text = element_text(size = text.font.size,  family="sans", color = "black"),
                   plot.title = element_text(size = text.font.size*1.2, face="bold"),
                   axis.text.x = element_text(size = text.font.size, color = "black"),
                   axis.title.x = element_text(size = text.font.size, color = "black", face="bold"),
                   axis.text.y = element_text(size = text.font.size, color = "black"),
                   axis.title.y = element_text(size = text.font.size, color = "black", face="bold"),
                   legend.text = element_text(size = text.font.size, color = "black"),
                   panel.grid.major.y = element_line(colour = "grey"), # add horizontal grid
                   panel.grid.major.x = element_line(colour = "grey"), # remove vertical grid
                   panel.grid.minor = element_blank(), # remove grid
                   panel.background = element_blank(), # remove background
                   axis.line = element_line(colour = "black"), # draw axis line in black
                   legend.position = "none")
  
  age.match.result <- c()
  test.df.all <- subset(df, select=c(age.col.name, group.col.name))
  names(test.df.all) <- c("Age", "Group")
  
  ## visualize age distribution before matching
  n.proband.per.group <- as.data.frame(table(test.df.all$Group))
  group.labels <- paste(n.proband.per.group$Var1, "\n(n=", n.proband.per.group$Freq, ")", sep = "")
  fig.title <-paste("Age distribution by ", group.col.name, " (", 0, " - ", 100, " years)", sep = "")
  before.match.fig <-ggplot(data = test.df.all, aes(x=Group, y=Age, color=Group))+
    geom_boxplot()+
    geom_point(position = position_jitter(width = 0.1, height = 0.1))+
    stat_compare_means(label.y = max(test.df.all$Age)+1, size=text.font.size/.pt*1.1)+
    scale_x_discrete(labels = group.labels)+
    labs(title = fig.title)+
    m.theme
  print(before.match.fig)
  
  for (age.min in age.min.values) {
    for (age.max in age.max.values) {
      test.df <- subset(test.df.all, Age>=age.min&Age<=age.max)
      normal.distributed <- T
      for (group.n in levels(test.df$Group)) {
        sha.test <- shapiro.test(test.df[test.df$Group==group.n, "Age"])
        #print(sha.test$p.value)
        if (sha.test$p.value<=0.05){
          normal.distributed <- F
          break()
        }
      }
      if (normal.distributed){
        test <- compare_means(Age~Group, data = test.df, method = "anova")
        p.value <- test$p.adj
      } else {
        test <- compare_means(Age~Group, data = test.df, method = "kruskal.test")
        p.value <- test$p.adj
      }
      if (p.value>0.05){
        result <- data.frame(Age.Min=age.min, Age.Max=age.max, P=p.value, N=nrow(test.df))
        count.n <- data.frame(table(test.df$Group))
        count.n <- dcast(count.n, . ~ Var1, value.var = "Freq")
        result <- cbind(result, count.n[, -1])
        if (length(age.match.result)==0){
          age.match.result <- result
        } else {
          age.match.result <- rbind(age.match.result, result)
        }
      }
    }
  }
  age.match.result$AgeDiff <- age.match.result$Age.Max-age.match.result$Age.Min
  age.match.result <- age.match.result[order(age.match.result$N, age.match.result$AgeDiff, decreasing = T), ]
  selected.age.min <- age.match.result$Age.Min[1]
  selected.age.max <- age.match.result$Age.Max[1]
  df.matched <- df[df[[age.col.name]]>=selected.age.min&df[[age.col.name]]<=selected.age.max, ]
  print(paste("!-------------Age matched cohort:", selected.age.min, "~", selected.age.max, "years -------------"))
  
  test.df <- subset(test.df.all, Age>=selected.age.min&Age<=selected.age.max)
  ## visualize age distribution after matching
  n.proband.per.group <- as.data.frame(table(test.df$Group))
  group.labels <- paste(n.proband.per.group$Var1, "\n(n=", n.proband.per.group$Freq, ")", sep = "")
  fig.title <-paste("Age distribution by ", group.col.name, " (", selected.age.min, " - ", selected.age.max, " years)", sep = "")
  after.match.fig <-ggplot(data = test.df, aes(x=Group, y=Age, color=Group))+
    geom_boxplot()+
    geom_point(position = position_jitter(width = 0.1, height = 0.1))+
    stat_compare_means(label.y = max(test.df$Age)+1, size=text.font.size/.pt*1.1)+
    scale_x_discrete(labels = group.labels)+
    labs(title = fig.title)+
    m.theme
  print(after.match.fig)
  
  return(list(before.match.fig, after.match.fig, df.matched, age.match.result))
}

## find rows and colums including NA, -Inf, or Inf values
check.NA.Inf.positions <- function(df){
  output.df <- data.frame()
  
  NA.rows.cols <- as.data.frame(which(is.na(df), arr.ind=TRUE))
  if (nrow(NA.rows.cols)>0){
    NA.rows.cols$id <- df[NA.rows.cols$row, 1]
    NA.rows.cols$col.name <- names(df)[NA.rows.cols$col]
    NA.rows.cols$type="NA"
    output.df <- rbind(output.df, NA.rows.cols)
  }
  
  Inf.rows.cols <- as.data.frame(which(df=="Inf", arr.ind=TRUE))
  if (nrow(Inf.rows.cols)>0){
    Inf.rows.cols$id <- df[Inf.rows.cols$row, 1]
    Inf.rows.cols$col.name <- names(df)[Inf.rows.cols$col]
    Inf.rows.cols$type="Inf"
    output.df <- rbind(output.df, Inf.rows.cols)
  }
  
  neg.Inf.rows.cols <- as.data.frame(which(df=="-Inf", arr.ind=TRUE))
  if (nrow(neg.Inf.rows.cols)>0){
    neg.Inf.rows.cols$id <- df[neg.Inf.rows.cols$row, 1]
    neg.Inf.rows.cols$col.name <- names(df)[neg.Inf.rows.cols$col]
    neg.Inf.rows.cols$type="-Inf"
    output.df <- rbind(output.df, neg.Inf.rows.cols)
  }
  return(output.df)
}



get.curve.length <- function(x.values, y.values){
  if (length(x.values)!=length(y.values)){
    print("Error: invalid length of values for the PCL calculation")
  } else {
    pcl <- 0
    for (n.point in 2:length(x.values)) {
      pcl <- pcl + sqrt((x.values[n.point]-x.values[n.point-1])*(x.values[n.point]-x.values[n.point-1])+(y.values[n.point]-y.values[n.point-1])*(y.values[n.point]-y.values[n.point-1]))
    }
    return(pcl)
  }
}

get.pressure.time.integral <- function(time.values, pre.values){
  if (length(time.values)!=length(pre.values)){
    print("Error: invalid length of values for the PCL calculation")
  } else {
    pti <- 0
    for (n.point in 2:length(time.values)) {
      pti <- pti + (pre.values[n.point] + pre.values[n.point-1])*(time.values[n.point] - time.values[n.point-1])/2.0
    }
    return(pti)
  }
}

get.peak.pressure <- function(time.values, pre.values){
  if (length(time.values)!=length(pre.values)){
    print("Error: invalid length of values for the PCL calculation")
  } else {
    return(max(pre.values))
  }
}



#### Global-Function: Format the p value  -----------------------------------------####
### Input: 
#- p.value: the numeric p value from statistical tests
## p>0.01 two digits, 0.01-0.001 three digits, and p<0.001 outputs "p<.001"
###------ Created on 2024-03-27
Format.P.Value <- function(p.value){
  if(!is.numeric(p.value)){
    p.value <- as.numeric(p.value)
  }
  if (p.value>=0.01){
    p.value.formatted <- sprintf("%.2f", p.value)
  } else if (p.value>=0.001){
    p.value.formatted <- sprintf("%.3f", p.value)
  } else {
    p.value.formatted <- "<.001"
  } 
  
  ## format p value again if it is equal to 0.05, 0.01, 0.001, or 0.0001
  if (p.value.formatted=="0.05"){
    p.value.formatted <- sprintf("%.3f", p.value)
  }
  if (p.value.formatted=="0.01"){
    p.value.formatted <- sprintf("%.3f", p.value)
  }
  if (p.value.formatted=="0.001"){
    p.value.formatted <- sprintf("%.4f", p.value)
  }
  return(p.value.formatted)
}



#### Global-Function: Visualize variables of the provided data frame  -----------------------------------------####
### Input: 
#- df: the data table for visualization
#- group.var.name: the name of the group label variable, it should be a categorized variable, e.g., Fallers vs. Non-Fallers
#- test.vars.names: the names of variables that are selected for visualization, e.g., c("BMI", "Age", "TypeDiabetes")
#- fig.output.path: the path that the figures will be exported, e.g., ""C:/local_work_ming/.../cohort_analysis_v1""
#- [optional] fig.x.colors: the color Platte indicating the levels of the group variable
#- [optional] fig.fill.colors: the color Platte indicating the levels of the categorized variable
#- [optional] text.font.size: the font size in the output figure, e.g., 14
#- [optional] p.value.text.color: the font color of the p value in the output figure, e.g., black
#- [optional] categorized.var.fig.save.width: the width of the output figure visualizing the categorized variable, e.g., 14
#- [optional] continuous.var.fig.save.width: the width of the output figure visualizing the continuous variable, e.g., 12
### Output: 
#- Numbers of figures [test.vars.names] will be saved at the path [fig.output.path]
###------ Created on 2024-03-27
#- ------------------------------------------------------------------------------------------------------------------------------
Visualize.Variables.New <- function(df, group.var.name, test.vars.names, fig.output.path, fig.x.colors, fig.fill.colors, text.font.size, p.value.text.color, categorized.var.fig.save.width, continuous.var.fig.save.width){
  # 
  # 
  # df <- Radar.df
  # group.var.name <- "Falls"
  # test.vars.names <- c("BMI", "Age", "TypeDiabetes")
  # fig.colors <- gg_color_hue(2)
  # text.font.size <- 14
  # p.value.text.color <- "black"
  #   fig.output.path <- output.folder
  library(tidyverse)
  library(gghalves)
  
  if (is.null(df)|| missing(df) || nrow(df)==0){
    print("Error: The provided data frame is empty!")
    return()
  }
  
  if (missing(group.var.name) || is.null(group.var.name)){
    print("Error: The provided group.var.name is empty!")
    return()
  }
  
  if (missing(test.vars.names) || is.null(test.vars.names)){
    print("Error: The provided test.vars.names is empty!")
    return()
  }
  
  if (missing(fig.output.path) || is.null(fig.output.path)){
    print("Error: The provided fig.output.path is empty!")
    return()
  }
  
  if (!group.var.name %in% names(df)){
    print(paste("Error: The provided group variable name [", group.var.name, "] is not included in the data frame!"))
    return()
  }
  
  if (length(unique(test.vars.names %in% names(df)))!=1){
    print(paste("Error: The provided test variables names [", test.vars.names, "] are not included in the data frame!"))
    return()
  }
  
  if (!dir.exists(fig.output.path)){
    dir.create(fig.output.path, showWarnings = T)
  }
  
  if (missing(text.font.size)){
    text.font.size <- 14
  } 
  
  if (missing(p.value.text.color)){
    p.value.text.color <- "black"
  }
  
  n.group.levels <- length(unique(df[[group.var.name]]))
  
  if (missing(categorized.var.fig.save.width)){
    if (n.group.levels<3){
      categorized.var.fig.save.width <- 8
    } else {
      categorized.var.fig.save.width <- 10 + (n.group.levels-2)*2
    }
  }  
  
  if (missing(continuous.var.fig.save.width)){
    continuous.var.fig.save.width <- 8 + (n.group.levels-2)*2
  }  
  
  if (missing(fig.x.colors) || length(fig.x.colors) !=n.group.levels){
    fig.x.colors <- gg_color_hue(length(unique(df[[group.var.name]])))
    print(paste("Warning: Fix invalid color platte,", length(fig.x.colors), "colors has been provided for indicating the Group variable with", n.group.levels, "levels"))
  }
  
  test.df <- subset(df, select=c(group.var.name, test.vars.names))
  test.df.summary <- build.df.summary(target.df = test.df, group.label.name = group.var.name, skip.column.names = c(), effective.digits = 1, debug.activated = F, only.mean.sd = F, group.test.enabled=T)
  fig.index <- 0
  for (var.name in test.vars.names) {
    fig.index <- fig.index + 1
    fig.df <- subset(test.df, select=c(group.var.name, var.name))
    names(fig.df) <- c("Group", "Value")
    na.rows <- nrow(fig.df) - nrow(na.omit(fig.df))
    fig.df <- na.omit(fig.df)
    original.p <- as.numeric(test.df.summary[test.df.summary[[group.var.name]]==var.name, "P"])
    if (length(original.p)>1){
      print(paste("Warning: multiple p values were found for", var.name, ", only the first one will be presented:", original.p))
      original.p <- original.p[1]
    }
    p.value <- paste("p=", Format.P.Value(original.p), sep = "")
    
    var.info <- get.feature.full.info.V2(var.name)
    if (nrow(var.info)>0&&var.info$Unit[1]!=""){
      fig.title <- paste(var.info$Name[1], " (", var.info$Game[1], ", ", var.info$Group[1], ")",  sep = "")
      fig.y.label <- var.info$Unit
    } else {
      fig.title <- var.name
      fig.y.label <- ""
    }
    fig.caption <- ""
    
    group.freq.df <- as.data.frame(table(fig.df$Group))
    group.freq.df$GroupLabel <- paste(group.freq.df$Var1, "\n(n=", group.freq.df$Freq, ")", sep = "")
    fig.caption <- paste(fig.caption, "NA: n=", na.rows)
    
    if (is.factor(fig.df$Value)){
      n.value.levels <- length(levels(fig.df$Value))
      if (missing(fig.fill.colors) || length(fig.fill.colors) !=n.value.levels){
        fig.fill.colors <- gg_color_hue(n.value.levels)
        print(paste("Warning: Fix invalid color platte,", length(fig.fill.colors), "colors has been provided for indicating a categorized variable with", n.value.levels, "levels"))
      }
      
      wrapped_title <- str_wrap(fig.title, width = categorized.var.fig.save.width*3.7) # width 调整为适合的值
      ## show categorized variable
      var.fig <- ggplot(data = fig.df, aes(x=Group, fill=Value)) +
        geom_bar( position = "fill", width = 0.4)+
        scale_y_continuous(breaks = seq(0, 1, 0.25), labels = seq(0, 100, 25))+
        scale_x_discrete(labels =group.freq.df$GroupLabel )+
        scale_fill_manual(values = fig.fill.colors) +
        annotate("text", x = length(levels(fig.df$Group))/2+0.5, y = 1.17, label=p.value, size=text.font.size/.pt, vjust=0.5, hjust=0.5, color=p.value.text.color)+
        annotate("segment", x = 1, xend = length(levels(fig.df$Group)), y = 1.05, yend = 1.05, colour = "black", linewidth=0.6)+
        #labs(title = wrapped_title, caption = fig.caption, fill=var.name, x="Group", y="%") +
        labs(title = wrapped_title, caption = fig.caption, fill="Value", x="Group", y="%") +
        theme_classic() +
        theme(legend.position = "right",
              axis.title.y = element_text(size = 16, color = "black"),
              axis.title.x = element_blank(),
              axis.text = element_text(size=13, color = "black"))
      save.fig(fig = var.fig, save.path = fig.output.path, save.name = paste(fig.index, var.name, "~", group.var.name, ".jpeg"), save.width = categorized.var.fig.save.width, save.height = 8, save.dpi = 600)
    } else {
      wrapped_title <- str_wrap(fig.title, width = continuous.var.fig.save.width*3.7) # width 调整为适合的值
      ## show continuous variable
      p.value.pos.y <- max(fig.df$Value) + 0.2*(max(fig.df$Value)-min(fig.df$Value))
      p.underline.pos.y <- max(fig.df$Value) + 0.1*(max(fig.df$Value)-min(fig.df$Value))
      #y.axis.max <- p.underline.pos.y
      #y.axis.min <- min(fig.df$Value) - 0.05*(max(fig.df$Value)-min(fig.df$Value))
      
      var.fig <- ggplot(data = fig.df, aes(x=Group, y=Value, fill=Group)) +
        geom_half_violin(side = "r", color=NA, alpha=0.35) +
        geom_half_boxplot(side = "r", errorbar.draw = FALSE, width=0.2, linewidth=0.5) +
        geom_half_point_panel(side = "l", shape=21, size=2.5, color="white", seed=1) +
        stat_summary(fun.y=mean, geom="point", shape=20, size=4, color="black", fill="black") +
        annotate("text", x = length(levels(fig.df$Group))/2+0.5, y = p.value.pos.y, label=p.value, size=text.font.size/.pt, vjust=0.5, hjust=0.5, color=p.value.text.color)+
        scale_fill_manual(values = fig.x.colors) +
        annotate("segment", x = 1, xend = length(levels(fig.df$Group)), y = p.underline.pos.y, yend = p.underline.pos.y, colour = "black", linewidth=0.6) +
        scale_x_discrete(labels =group.freq.df$GroupLabel )+
        #scale_y_continuous(limits = c(y.axis.min, y.axis.max)) +
        labs(title = wrapped_title, caption = fig.caption, y=fig.y.label, x="Group") +
        theme_classic() +
        theme(legend.position = "none",
              axis.title.y = element_text(size = 16, color = "black"),
              axis.title.x = element_blank(),
              axis.text = element_text(size=13, color = "black"))
      save.fig(fig = var.fig, save.path = fig.output.path, save.name = paste(fig.index, var.name, "~", group.var.name, ".jpeg"), save.width =continuous.var.fig.save.width, save.height = 8, save.dpi = 600)
    }
    #print(var.fig)
  }
}


#### Global-Function: Visualize variables of the provided data frame  -----------------------------------------####
### Input: 
#- df: the data table for visualization
#- group.var.name: the name of the group label variable, it should be a categorized variable, e.g., Fallers vs. Non-Fallers
#- test.vars.names: the names of variables that are selected for visualization, e.g., c("BMI", "Age", "TypeDiabetes")
#- fig.output.path: the path that the figures will be exported, e.g., ""C:/local_work_ming/.../cohort_analysis_v1""
#- [optional] fig.x.colors: the color Platte indicating the levels of the group variable
#- [optional] fig.fill.colors: the color Platte indicating the levels of the categorized variable
#- [optional] text.font.size: the font size in the output figure, e.g., 14
#- [optional] p.value.text.color: the font color of the p value in the output figure, e.g., black
#- [optional] categorized.var.fig.save.width: the width of the output figure visualizing the categorized variable, e.g., 14
#- [optional] continuous.var.fig.save.width: the width of the output figure visualizing the continuous variable, e.g., 12
### Output: 
#- Numbers of figures [test.vars.names] will be saved at the path [fig.output.path]
###------ Created on 2024-03-27
#- ------------------------------------------------------------------------------------------------------------------------------
Visualize.Variables.in.PPT <- function(df, group.var.name, test.vars.names, fig.output.path, append.save = F, fig.x.colors, fig.fill.colors, p.value.text.color, var.fig.save.width, var.fig.save.height, save.as.image=T){
  # 
  # 
  # df <- Radar.df
  # group.var.name <- "Falls"
  # test.vars.names <- c("BMI", "Age", "TypeDiabetes")
  # fig.colors <- gg_color_hue(2)
  # text.font.size <- 14
  # p.value.text.color <- "black"
  #   fig.output.path <- output.folder
  library(tidyverse)
  library(gghalves)
  
  if (is.null(df)|| missing(df) || nrow(df)==0){
    print("Error: The provided data frame is empty!")
    return()
  }
  
  if (missing(group.var.name) || is.null(group.var.name)){
    print("Error: The provided group.var.name is empty!")
    return()
  }
  
  if (missing(test.vars.names) || is.null(test.vars.names)){
    print("Error: The provided test.vars.names is empty!")
    return()
  }
  
  if (missing(fig.output.path) || is.null(fig.output.path)){
    print("Error: The provided fig.output.path is empty!")
    return()
  }
  
  if (!group.var.name %in% names(df)){
    print(paste("Error: The provided group variable name [", group.var.name, "] is not included in the data frame!"))
    return()
  }
  
  if (length(unique(test.vars.names %in% names(df)))!=1){
    print(paste("Error: The provided test variables names [", test.vars.names, "] are not included in the data frame!"))
    return()
  }
  
  m.text.font.size <- 10
  
  if (missing(p.value.text.color)){
    p.value.text.color <- "black"
  }
  
  n.group.levels <- length(unique(df[[group.var.name]]))
  
  # if (n.group.levels<3){
  #   categorized.var.fig.save.width <- 3
  # } else {
  #   categorized.var.fig.save.width <- 3 + (n.group.levels-2)*var.fig.save.width
  # } 
  
  categorized.var.fig.save.width <- var.fig.save.width
  continuous.var.fig.save.width <- var.fig.save.width
  

  if (missing(fig.x.colors) || length(fig.x.colors) !=n.group.levels){
    fig.x.colors <- gg_color_hue(length(unique(df[[group.var.name]])))
    print(paste("Warning: Fix invalid color platte,", length(fig.x.colors), "colors has been provided for indicating the Group variable with", n.group.levels, "levels"))
  }
  
  if (!append.save){
    save.fig.to.pptx(ggfig = ggplot(), path = fig.output.path, append = F, vertical = F, img.width = 4, img.height = 4, as.image = save.as.image)
  }
  
  test.df <- subset(df, select=c(group.var.name, test.vars.names))
  test.df.summary <- build.df.summary(target.df = test.df, group.label.name = group.var.name, skip.column.names = c(), effective.digits = 1, debug.activated = F, only.mean.sd = F, group.test.enabled=T)
  fig.index <- 0
  for (var.name in test.vars.names) {
    fig.index <- fig.index + 1
    fig.df <- subset(test.df, select=c(group.var.name, var.name))
    names(fig.df) <- c("Group", "Value")
    na.rows <- nrow(fig.df) - nrow(na.omit(fig.df))
    fig.df <- na.omit(fig.df)
    original.p <- as.numeric(test.df.summary[test.df.summary[[group.var.name]]==var.name, "P"])
    if (length(original.p)>1){
      print(paste("Warning: multiple p values were found for", var.name, ", only the first one will be presented:", original.p))
      original.p <- original.p[1]
    }
    p.value <- paste("p=", Format.P.Value(original.p), sep = "")
    
    var.info <- get.feature.full.info.V2(var.name)
    if (nrow(var.info)>0&&var.info$Unit[1]!=""){
      fig.title <- paste(var.info$Name[1], " (", var.info$Game[1], ", ", var.info$Group[1], ")",  sep = "")
      fig.y.label <- var.info$Unit
    } else {
      fig.title <- var.name
      fig.y.label <- ""
    }
    fig.caption <- ""
    
    group.freq.df <- as.data.frame(table(fig.df$Group))
    group.freq.df$GroupLabel <- paste(group.freq.df$Var1, "\n(n=", group.freq.df$Freq, ")", sep = "")
    fig.caption <- paste(fig.caption, "NA: n=", na.rows)
    
    if (is.factor(fig.df$Value)){
      n.value.levels <- length(levels(fig.df$Value))
      if (missing(fig.fill.colors) || length(fig.fill.colors) !=n.value.levels){
        fig.fill.colors <- gg_color_hue(n.value.levels)
        print(paste("Warning: Fix invalid color platte,", length(fig.fill.colors), "colors has been provided for indicating a categorized variable with", n.value.levels, "levels"))
      }
      
      wrapped_title <- str_wrap(fig.title, width = categorized.var.fig.save.width*8) # width 调整为适合的值
      ## show categorized variable
      var.fig <- ggplot(data = fig.df, aes(x=Group, fill=Value)) +
        geom_bar( position = "fill", width = 0.4)+
        scale_y_continuous(breaks = seq(0, 1, 0.25), labels = seq(0, 100, 25))+
        scale_x_discrete(labels =group.freq.df$GroupLabel )+
        scale_fill_manual(values = fig.fill.colors) +
        annotate("text", x = length(levels(fig.df$Group))/2+0.5, y = 1.17, label=p.value, size=m.text.font.size/.pt, vjust=0.5, hjust=0.5, color=p.value.text.color)+
        annotate("segment", x = 1, xend = length(levels(fig.df$Group)), y = 1.05, yend = 1.05, colour = "black", linewidth=0.6)+
        #labs(title = wrapped_title, caption = fig.caption, fill=var.name, x="Group", y="%") +
        labs(title = wrapped_title, caption = fig.caption, fill="Value", x="Group", y="%") +
        m.no.legend.theme
      save.fig.to.pptx(ggfig = var.fig, path = fig.output.path, append = append.save, vertical = F, img.width = categorized.var.fig.save.width, img.height = var.fig.save.height, as.image = save.as.image)
    } else {
      wrapped_title <- str_wrap(fig.title, width = continuous.var.fig.save.width*8) # width 调整为适合的值
      ## show continuous variable
      p.value.pos.y <- max(fig.df$Value) + 0.2*(max(fig.df$Value)-min(fig.df$Value))
      p.underline.pos.y <- max(fig.df$Value) + 0.1*(max(fig.df$Value)-min(fig.df$Value))
      #y.axis.max <- p.underline.pos.y
      #y.axis.min <- min(fig.df$Value) - 0.05*(max(fig.df$Value)-min(fig.df$Value))
      
      var.fig <- ggplot(data = fig.df, aes(x=Group, y=Value, fill=Group)) +
        geom_half_violin(side = "r", color=NA, alpha=0.35) +
        geom_half_boxplot(side = "r", errorbar.draw = FALSE, width=0.2, linewidth=0.5) +
        geom_half_point_panel(side = "l", shape=21, size=2.5, color="white", seed=1) +
        stat_summary(fun.y=mean, geom="point", shape=20, size=4, color="black", fill="black") +
        annotate("text", x = length(levels(fig.df$Group))/2+0.5, y = p.value.pos.y, label=p.value, size=m.text.font.size/.pt, vjust=0.5, hjust=0.5, color=p.value.text.color)+
        scale_fill_manual(values = fig.x.colors) +
        annotate("segment", x = 1, xend = length(levels(fig.df$Group)), y = p.underline.pos.y, yend = p.underline.pos.y, colour = "black", linewidth=0.6) +
        scale_x_discrete(labels =group.freq.df$GroupLabel )+
        #scale_y_continuous(limits = c(y.axis.min, y.axis.max)) +
        labs(title = wrapped_title, caption = fig.caption, y=fig.y.label, x="Group") +
        m.no.legend.theme
      save.fig.to.pptx(ggfig = var.fig, path = fig.output.path, append = append.save, vertical = F, img.width = continuous.var.fig.save.width, img.height = var.fig.save.height, as.image = save.as.image)
    }
    #print(var.fig)
  }
}

## deprecated on 19.04.2024, please use Visualize.Variables.New() function of this script. 
# ## visualize different variables between groups
# visualize.variables.betw.groups <- function(target.df, skip.column.names, group.label.name, fig.title, figure.output.folder){
#   text.font.size <- 24
#   m.theme <- theme(text = element_text(size = text.font.size,  family="sans", color = "black"),
#                    plot.title = element_text(size = text.font.size*1.2, face="bold"),
#                    axis.text.x = element_text(size = text.font.size, color = "black", angle = 60, vjust = 0.5, hjust = 0.5),
#                    axis.title.x = element_text(size = text.font.size, color = "black", face="bold"),
#                    axis.text.y = element_text(size = text.font.size, color = "black"),
#                    axis.title.y = element_text(size = text.font.size, color = "black", face="bold"),
#                    strip.text = element_text(size = text.font.size, color = "black", face="bold"),
#                    legend.text = element_text(size = text.font.size, color = "black"),
#                    panel.grid.major.y = element_blank(), # add horizontal grid
#                    panel.grid.major.x = element_blank(), # remove vertical grid
#                    panel.grid.minor = element_blank(), # remove grid
#                    panel.background = element_blank(), # remove background
#                    axis.line = element_line(colour = "black"), # draw axis line in black
#                    legend.position = "right")
#   
#   if (missing(group.label.name)){
#     print("Warning: no group factor was provided, suppose it is named 'Group'!")
#     group.label.name <- "Group"
#   }
#   
#   if (!group.label.name%in%names(target.df)){
#     return()
#   }
#   
#   if (!is.factor(target.df[[group.label.name]])){
#     target.df[[group.label.name]] <- factor(target.df[[group.label.name]])
#   }
#   
#   df.summary <- build.df.summary(target.df, group.label.name, skip.column.names, effective.digits = 1, debug.activated = F, only.mean.sd = F, group.test.enabled=T)
#   print(df.summary)
#   
#   vars.for.fig <- names(target.df)[!names(target.df) %in% c(group.label.name, skip.column.names)]
#   m.fig.list <- vector(mode = "list", length = length(vars.for.fig))
#   fig.count <- 1
#   for (var.name in vars.for.fig) {
#     fig.df <- subset(target.df, select=c(group.label.name, var.name))
#     names(fig.df) <- c("Group", "Variable")
#     if (nrow(fig.df[!complete.cases(fig.df), ])>0){
#       print(paste(var.name, "has NA values, remove them before visualization"))
#       fig.df <- na.omit(fig.df)
#     }
#     
#     p.value <- df.summary[df.summary[[group.label.name]]==var.name, "All"]
#     
#     if( is.numeric(fig.df$Variable)&length(unique(fig.df$Variable))>5){
#       m.fig <- ggplot(data = fig.df, aes(x=Group, y=Variable))+
#         geom_boxplot()+
#         geom_point(position = position_jitter(width = 0.1, height = 0.1, seed=1))+
#         annotate(geom="text", x=mean(as.numeric(unique(as.factor(fig.df$Group)))), y=max(fig.df$Variable), label=p.value,
#                  color="red", vjust=0.5, hjust=0.5, size=text.font.size/.pt*1.2)
#     } else {
#       if (!is.factor(fig.df$Variable )){
#         fig.df$Variable  <- factor(fig.df$Variable)
#       }
#       fig.df.freq <- data.frame(table(fig.df$Group, fig.df$Variable))
#       names(fig.df.freq) <- c("Group", "Variable", "Freq")
#       for (n.row in 1:nrow(fig.df.freq)) {
#         fig.df.freq$N.Per.Group[n.row] <- sum(fig.df.freq[fig.df.freq$Group==fig.df.freq$Group[n.row], "Freq"])
#         fig.df.freq$Prop[n.row] <- round(fig.df.freq$Freq[n.row]/fig.df.freq$N.Per.Group[n.row]*100)
#       }
#       m.fig <- ggplot(data = fig.df.freq, aes(x=Group, y=Variable))+
#         geom_point(aes(size = Prop), color="#00BFC4", alpha=0.6) +
#         scale_size(range = c(0.5, 12))+  # Adjust the range of points size
#         annotate(geom="text", x=mean(as.numeric(unique(as.factor(fig.df.freq$Group)))), y=mean(as.numeric(unique(as.factor(fig.df.freq$Variable)))), label=p.value,
#                  color="red", vjust=0.5, hjust=0.5, size=text.font.size/.pt*1.2)
#     }
#     m.fig <- m.fig + labs(title = paste(var.name, "Group", sep = " ~ "))+ylab(label = var.name)+m.theme
#     print(m.fig)
#     
#     m.fig.list[[fig.count]] <- m.fig
#     fig.count <- fig.count + 1
#   }
#   full.figure <- ggarrange(plots = m.fig.list, top = text_grob(fig.title, face = "bold", size = text.font.size*1.2))
#   save.fig(fig = full.figure, save.path = figure.output.folder, save.name = paste(fig.title,  ".jpeg", sep = ""), save.width = text.font.size*3, save.height = text.font.size*2, save.dpi = 600)
#   
#   
# }



get_combinations_with_length <- function(list, length){
  library("sets")
  my.sets <- set_power(list)
  output <- vector(mode = "list", length = choose(length(list), length))
  n.ele <- 0
  for (ele in my.sets) {
    comb <- c()
    for (ele_sub in ele) {
      comb <- c(comb, ele_sub)
    }
    if (length(comb)==length){
      n.ele <- n.ele + 1
      print("------------------------")
      print(comb)
      output[[n.ele]] <- comb
    }
  }
  return(output)
}


## draw a odds ratio plot
draw.odds.ratio.plot2 <- function (x, x_label = "Variables", y_label = "Odds Ratio", 
                                   title = NULL, subtitle = NULL, point_colors = c("#7f7f7f", "black"), error_bar_colour = "black", 
                                   point_size = 2, error_bar_width = 0.2, h_line_color = "#bbbbbb", reorder.vars.by="OR", reordered.vars) 
{
  OR <- NULL
  lower <- NULL
  upper <- NULL
  tmp <- data.frame(cbind(exp(coef(x)), exp(confint(x))))
  odds <- tmp[-1, ]
  names(odds) <- c("OR", "lower", "upper")
  odds$vars <- row.names(odds)
  odds_save <- as_tibble(odds)
  # ticks <- c(seq(0.1, 1, by = 0.1), seq(0, 10, by = 1), seq(10,
  #                                                           100, by = 10))
  # ticks.labels <- c(seq(0.1, 1, by = 0.1), seq(0, 10, by = 1), seq(10,
  #                                                                  100, by = 10))
  
  

  #ticks <- c(0.1, 0.5, 1, 5, 10, 50, 100)
  OR.demicals = 2
  odds$Signif <- "No"
  odds$upper.round <- sapply(odds$upper, round, OR.demicals)
  odds$lower.round <- sapply(odds$lower, round, OR.demicals)
  #odds[(odds$lower.round<1&odds$upper.round<1)|(odds$lower.round>1&odds$upper.round>1), "Signif"] <- "Yes"
  odds[(odds$lower<1&odds$upper<1)|(odds$lower>1&odds$upper>1), "Signif"] <- "Yes"
  
  text.font.size <- 14
  paper.theme <- theme(text = element_text(size = text.font.size,  family="sans", color = "black"),
                       plot.title = element_text(size = text.font.size*1.2, face="bold"),
                       axis.text.x = element_text(size = text.font.size, color = "black"),
                       axis.title.x = element_text(size = text.font.size, color = "black", face="bold"),
                       axis.text.y = element_text(size = text.font.size, color = "black"),
                       axis.title.y = element_text(size = text.font.size, color = "black", face="bold"),
                       strip.text = element_text(size = text.font.size, color = "black", face="bold"),
                       legend.text = element_text(size = text.font.size, color = "black"),
                       panel.grid.major.y = element_blank(), # add horizontal grid
                       panel.grid.major.x = element_blank(), # remove vertical grid
                       panel.grid.minor = element_blank(), # remove grid
                       panel.background = element_blank(), # remove background
                       axis.line = element_line(colour = "black"), # draw axis line in black
                       legend.position = "none")

  odds.max <- ceiling(max(odds$upper))
  odds.min <- floor(min(odds$lower)*10)/10
  if (odds.min==0){
    odds.min <- floor(min(odds$lower)*100)/100
  }
  
  if (odds.min==0.9||odds.min==0.8){
    odds.min <- 0.7
  }
  
  if (odds.max>1 && odds.min <1){
    ticks <- c(odds.min, 1, odds.max)
    ticks.labels <- c(odds.min, 1, odds.max)
  } else if (odds.max>1 && odds.min >=1){
    ticks <- c(1, odds.max)
    ticks.labels <- c(1, odds.max)
  } else{
    ticks <- c(odds.min, 1)
    ticks.labels <- c(odds.min, 1)
  } 
  print(paste("Odds range:", odds.min, " - ", odds.max, "with ticks:", ticks.labels))
  
  odds$OR.labels <- rep("", nrow(odds))
  odds$OR.label.pos <- rep(max(odds$lower), nrow(odds))
  odds.plot.y.limit.min <- min(odds$lower)
  odds.plot.y.limit.max <- max(odds$upper)
  ticks.modified <- c(ticks, odds.plot.y.limit.min)
  ticks.label.modified <- c(ticks, "")
  for(n.row in 1:nrow(odds)){
    #if (odds$Signif[n.row]=="Yes"){
    odds$OR.labels[n.row] <- paste(specify_decimal(odds$OR[n.row], OR.demicals), " (",  specify_decimal(odds$lower[n.row], OR.demicals), "-",  specify_decimal(odds$upper[n.row], OR.demicals), ")", sep = "")
    #}
  }
  print(odds)
  
  if(missing(reordered.vars)){
    reordered.vars <- odds$vars
  }
  
  if (reorder.vars.by=="OR"){
    plot <- ggplot(odds, aes(y = OR, x = reorder(vars, OR), color=Signif)) + 
      geom_hline(yintercept = 1, linetype = 3, color = h_line_color) + 
      geom_errorbar(aes(ymin = lower, ymax = upper, color=Signif), width = error_bar_width) + 
      geom_point(size = point_size, alpha=1) + 
      geom_point(aes(x=length(unique(odds$vars))+1, y=1), color="white") + 
      #geom_text(aes(x=reorder(vars, OR), y=0, label=OR.labels), size=text.font.size/.pt, vjust=0.5, hjust=0)+
      #geom_text(aes(label=OR.labels), size=text.font.size/.pt*0.9, vjust==-0.4, hjust=0.5, color=error_bar_colour)+
      geom_text(aes(y=odds.max, label=OR.labels), size=text.font.size/.pt*0.9, vjust==-0.4, hjust=1)+
      #scale_y_continuous(limits = c(odds.min, odds.max), breaks = ticks, labels = ticks.labels)+
      scale_y_log10(limits = c(odds.min, odds.max), breaks = ticks, labels = ticks.labels)+
      #scale_y_log10(breaks = ticks, labels = ticks.labels, guide = guide_axis(check.overlap = TRUE)) +
      scale_color_manual(values = point_colors)+
      coord_flip() + 
      labs(title = title, subtitle = subtitle, x = x_label, 
           y = y_label) 
    
    OR.plot <- ggplot(odds, aes(y = OR, x = reorder(vars, OR)))+
      geom_text(aes(x=reorder(vars, OR), y=0, label=OR.labels), size=text.font.size/.pt, vjust=0.5, hjust=0.5)+
      #scale_y_continuous(limits = c(-3, 3))+
      labs(title = "", subtitle = subtitle, x = x_label, y = y_label) + coord_flip() + 
      theme(text = element_text(size = text.font.size,  family="sans", color = "black"),
            plot.title = element_text(size = text.font.size, face="bold"),
            axis.text.x = element_blank(),
            axis.title.x = element_blank(),
            axis.text.y =element_blank(),
            axis.title.y = element_blank(),
            axis.ticks = element_blank(),
            panel.grid.major = element_blank(), # add horizontal grid
            panel.grid.minor = element_blank(), # remove grid
            panel.background = element_blank(), # remove background
            axis.line = element_blank(), # draw axis line in black
            legend.position = "none")
    
    
  } else {
    odds$vars <- factor(odds$vars, levels = reordered.vars)
    plot <- ggplot(odds, aes(y = OR, x = vars, color=Signif)) + 
      geom_hline(yintercept = 1, linetype = 3, color = h_line_color) + 
      geom_errorbar(aes(ymin = lower, ymax = upper, color=Signif), width = error_bar_width) + 
      geom_point(size = point_size, alpha=1) + 
      geom_point(aes(x=length(unique(odds$vars))+1, y=1), color="white") + 
      #geom_text(aes(x=reorder(vars, OR), y=0, label=OR.labels), size=text.font.size/.pt, vjust=0.5, hjust=0)+
      #geom_text(aes(label=OR.labels), size=text.font.size/.pt, vjust=-0.4, hjust=0.5, color=error_bar_colour)+
      geom_text(aes(y=odds.max, label=OR.labels), size=text.font.size/.pt, vjust=-0.4, hjust=1)+
      #scale_y_log10(breaks = ticks, labels = ticks.labels, guide = guide_axis(check.overlap = TRUE)) +
      #scale_y_continuous(limits = c(odds.min, odds.max), breaks = ticks, labels = ticks.labels)+
      scale_y_log10(limits = c(odds.min, odds.max), breaks = ticks, labels = ticks.labels)+
      scale_color_manual(values = point_colors)+
      #scale_x_discrete(breaks = reordered.vars, labels=seq(1, nrow(odds), 1))+
      coord_flip() + 
      labs(title = title, subtitle = subtitle, x = x_label, 
           y = y_label) 
    OR.plot <- ggplot(odds, aes(y = OR, x = reorder(vars, OR)))+
      geom_text(aes(x=vars, y=0, label=OR.labels), size=text.font.size/.pt, vjust=0.5, hjust=0.5)+
      #scale_y_continuous(limits = c(-3, 3))+
      labs(title = "", subtitle = subtitle, x = x_label, y = y_label) + coord_flip() + 
      theme(text = element_text(size = text.font.size,  family="sans", color = "black"),
            plot.title = element_text(size = text.font.size*1.2, face="bold"),
            axis.text.x = element_blank(),
            axis.title.x = element_blank(),
            axis.text.y =element_blank(),
            axis.title.y = element_blank(),
            axis.ticks = element_blank(),
            panel.grid.major = element_blank(), # add horizontal grid
            panel.grid.minor = element_blank(), # remove grid
            panel.background = element_blank(), # remove background
            axis.line = element_blank(), # draw axis line in black
            legend.position = "none")
  }
  plot <- plot + paper.theme
  
  # returns <- list(odds_data = odds_save, odds_plot = ggarrange(plots = list(plot, OR.plot), widths = c(3, 2), heights = c(1)))
  returns <- list(odds_data = odds_save, odds_plot =plot)
  return(returns)
}

## a function that performs logistic regression analysis
fit.logistic.reg <- function(df, outcome.name, vars.name, output.path){
  if (!outcome.name %in% names(df)){
    print(paste("Error: the outcome variable is not included in the data frame", outcome.name))
    return()
  }
  
  if (FALSE %in% (vars.name %in% names(df))){
    print(paste("Error: the provided variables are not included in the data frame", vars.name))
    return()
  }
  
  if (!is.factor(df[, outcome.name])){
    df[, outcome.name] <- as.factor(df[, outcome.name])
    if (length(unique(df[, outcome.name]))>2){
      print(paste("Error: the outcome variable is invalid, with more than two levels:", paste(unique(df[, outcome.name]), collapse = ", ")))
      return()  
    }
  }
  
  if (!dir.exists(output.path)){
    dir.create(output.path, showWarnings = T)
  }
  
  reg.df <- subset(df, select=c(outcome.name, vars.name))
  names(reg.df)[1] <- c("Outcome")
  
  if (sum(is.na(reg.df))>0){
    print(paste("!!!!!!------------------- NA values with following positions were excluded in the analysis:----------"))
    print(which(is.na(reg.df)))
    reg.df <- na.omit(reg.df)
    print(paste("!!!!!!------------------- ----------------------------------:----------"))
  }
  
  reg.df.report <- build.df.summary(target.df = reg.df, group.label.name = "Outcome", skip.column.names = c(), effective.digits = 1, debug.activated = F, only.mean.sd = F, group.test.enabled=T)
  
  fit1 <- glm(Outcome~., family = binomial(link = "logit"), data = reg.df)
  fit1.report <- as.data.frame(summary(fit1)$coefficients)
  fit1.report$Signif <-sapply(fit1.report$`Pr(>|z|)`, convert.p.value.to.signif)
  library(pscl)
  #adj.R2 <- round(rsq(fit1, adj = T), 2)
  ##fit1.report$AdjR2 <-adj.R2
  fit1.report$AdjR2 <-pR2(fit1)['McFadden']
  fit1.report <- cbind(fit1.report, data.frame(exp(cbind(coef(fit1), confint(fit1)))))
  names(fit1.report)[7:9] <- c("OR", "OR2.5%", "OR97.5%")
  #print(fit1.report)
  
  
  
  plotty <- draw.odds.ratio.plot2(fit1, title = str_wrap(paste("Adjusted OR for", outcome.name), width = 35), reorder.vars.by = "Feature",  reordered.vars = rev(names(fit1$coefficients)[-1]))
  print(plotty$odds_plot)
  
  output.file.name <- paste(outcome.name, "~", paste(vars.name[1:2], collapse = " + "),  "logistic fit")
  
  fit1.report2 <- fit1.report
  fit1.report2[, -5] <- sapply(fit1.report2[, -5], round, 2)
  
  write.xlsx(fit1.report2,
             file = paste(output.path, paste(output.file.name, ".xlsx"), sep = "/"),
             sheetName = "Formatted", append = F, row.names = T)
  write.xlsx(fit1.report,
             file = paste(output.path, paste(output.file.name, ".xlsx"), sep = "/"),
             sheetName = "All", append = T, row.names = T)
  write.xlsx(reg.df.report,
             file = paste(output.path, paste(output.file.name, ".xlsx"), sep = "/"),
             sheetName = "Summary", append = T, row.names = T)
  write.xlsx(reg.df,
             file = paste(output.path, paste(output.file.name, ".xlsx"), sep = "/"),
             sheetName = "Fit Data", append = T, row.names = T)
  
  save.fig(plotty$odds_plot, save.path = output.path, save.name = paste(output.file.name, ".jpeg"), save.width = 14, save.height = 8, save.dpi = 600)
  return(fit1.report)
}


############################################################################################################################################
## perform logistic regression analysis 
## df: data frame contains the response and variates
## target.var: name of the response variable
## skip.vars: name of variables to be excluded in the analysis
## covariate.vars: name of covariate variables to adjust the impact on the response variable, e.g., age, gender, BMI, ...
## only.return.signif.tab: if true, only variables with significant impact will be returned, otherwise, complete results will be returned
## debug.activated: if yes, detailed information will be printed during the analysis
## return: a data.frame consists of the results of analysis
############################################################################################################################################
build.logistic.regession <- function(df, target.var, skip.vars, covariate.vars, only.return.signif.tab, debug.activated){
  if (is.null(df) || nrow(df)<1){
    print("invalid df, stop")
    return()
  }
  if (sum(is.na(df))>0){
    print("follwing NA vluaes will be excluded in analysis")
    print(colSums(is.na(df)))
    df <- na.omit(df)
  }
  if (!target.var %in% names(df)){
    print(paste("target name [", target.var, "] is not included in the df, stop!"))
    return() 
  }
  
  if (missing(covariate.vars)){
    covariate.vars <- c()
  }
  
  test.vars <-  setdiff(names(df), c(target.var, skip.vars, covariate.vars))
  df <- subset(df, select=c(target.var, setdiff(names(df), c(target.var, skip.vars))))
  names(df)[1] <- "Outcome"
  if (!is.factor(df$Outcome)){
    df$Outcome <- factor(df$Outcome)
  }
  
  ## change covariates into factor if necessary
  final.covarites <- c()
  for (cova in covariate.vars) {
    if (cova %in% names(df)){
      final.covarites <- c(final.covarites, cova)
      if (length(unique(df[[cova]]))<5){
        df[[cova]] <- as.factor(df[[cova]])
      } 
    }
  }
  
  if (missing(only.return.signif.tab)){
    only.return.signif.tab <- F
  }
  
  if (missing(debug.activated)){
    debug.activated <- F
  }
  
  library(broom)
  library(dplyr)  
  output <- data.frame()
  skip.vars <- c()
  for (n.var in 1:length(test.vars)) {
    var.name <- test.vars[n.var]
    reg.df <- subset(df, select=c("Outcome", final.covarites, var.name))
    ## change var into factor if necessary
    if (length(unique(reg.df[[var.name]]))<5){
      reg.df[[var.name]] <- as.factor(reg.df[[var.name]])
    }
    print(paste(n.var, var.name))
    model.fit <- glm(Outcome~., family = binomial(link = "logit"), data = reg.df)
    ## get model statistics
    coff.df <- tidy(model.fit)
    names(coff.df)[1] <- "var"
    coff.df$p.format <- sapply(coff.df$p.value, format.p.value)
    coff.df$p.signif <- sapply(coff.df$p.value, convert.p.value.to.signif)
    
    # print(coff.df)
    # print("---------------------")
    
    ## get odds ratio
    if (sum(is.na(confint(model.fit)))>0||sum(is.na(coef(model.fit)))>0){
      skip.vars <- c(skip.vars, var.name)
    } else {
      odds.ratio.df <- as.data.frame(cbind(exp(coef(model.fit)), exp(confint(model.fit))))
      names(odds.ratio.df) <- c("or", "conf.low", "conf.high")
      odds.ratio.df <- cbind(data.frame(var=rownames(odds.ratio.df)), odds.ratio.df)
      odds.ratio.df[, -1] <- sapply(odds.ratio.df[, -1], round, 2)
      
      ## merge them together
      fit.out.df <- merge(coff.df, odds.ratio.df, by="var", all = F)
      fit.out.df.sorted <- fit.out.df[order(match(fit.out.df$var, coff.df$var)), ]
      fit.out.df.sorted$VarIndex <- n.var
      fit.out.df.sorted$VarName <- var.name
      fit.out.df.sorted$Target <- T
      fit.out.df.sorted$Target[1:(length(final.covarites)+1)] <- F
      rownames(fit.out.df.sorted) <- seq(1, nrow(fit.out.df.sorted), 1)
      output <- rbind(output, fit.out.df.sorted) 
      if (debug.activated){
        print(fit.out.df.sorted)
        print("-----------------------------------------------------------------------------------------------------------")
      }
    }
  }
  
  if (length(skip.vars)>0){
    print(paste("Warning: ", length(skip.vars), "variables were excluded for the regression analysis due to NA values from OR or CI calculation!!"))
    print(skip.vars)
  }
  
  if (only.return.signif.tab){
    return(output)
  } else {
    return(output)
  }
}


############################################################################################################################################
## 进行线性回归分析的函数
## df: 数据框，包含响应变量和自变量
## target.var: 响应变量名（字符串）
## skip.vars: 需要在分析中排除的变量名（向量）
## covariate.vars: 要在模型中控制的协变量名（向量），例如 age, gender, BMI, ...
## only.return.signif.tab: 如果为 TRUE，只返回（或仅在返回前自行筛选）显著的结果；否则返回所有
## debug.activated: 如果为 TRUE，会在分析中打印更详细的信息
## return: 一个 data.frame，包含分析结果
############################################################################################################################################
build.linear.regression <- function(df, 
                                    target.var, 
                                    skip.vars, 
                                    covariate.vars, 
                                    only.return.signif.tab, 
                                    debug.activated){
  # 1) 检查数据合法性
  if (is.null(df) || nrow(df) < 1){
    print("invalid df, stop")
    return()
  }
  
  # 2) 检查 NA 并移除
  if (sum(is.na(df)) > 0){
    print("以下 NA 值所在的行将在分析中被移除：")
    print(colSums(is.na(df)))
    df <- na.omit(df)
  }
  
  # 3) 确保目标变量存在
  if (!target.var %in% names(df)){
    print(paste("target name [", target.var, "] is not included in the df, stop!"))
    return() 
  }
  
  # 4) 设置缺省参数
  if (missing(covariate.vars)){
    covariate.vars <- c()
  }
  if (missing(only.return.signif.tab)){
    only.return.signif.tab <- FALSE
  }
  if (missing(debug.activated)){
    debug.activated <- FALSE
  }
  
  # 5) 划定需要检测的变量
  test.vars <- setdiff(names(df), c(target.var, skip.vars, covariate.vars))
  
  # 6) 子集并将目标变量重命名为 "Outcome"
  df <- subset(df, select = c(target.var, setdiff(names(df), c(target.var, skip.vars))))
  names(df)[1] <- "Outcome"
  
  # （可选）若需要，强制 Outcome 为 numeric
  # df$Outcome <- as.numeric(df$Outcome)
  
  # 7) 处理协变量（若某个协变量不到 5 个 unique 值，视为分类变量）
  final.covarites <- c()
  for (cova in covariate.vars) {
    if (cova %in% names(df)){
      final.covarites <- c(final.covarites, cova)
      if (length(unique(df[[cova]])) < 5){
        df[[cova]] <- as.factor(df[[cova]])
      } 
    }
  }
  
  # 8) 载入需要的包
  #    如果您的环境中已加载可不重复加载
  library(broom)
  library(dplyr)
  
  # 存放最终结果的 data.frame
  output <- data.frame()
  
  # 记录因为 confint 或 coef 计算失败而被跳过的变量
  skip.vars.internal <- c()
  
  # 9) 循环对每一个 test.var 做回归
  for (n.var in seq_along(test.vars)) {
    var.name <- test.vars[n.var]
    
    # 只取目标变量 + 协变量 + 当前测试变量
    reg.df <- subset(df, select = c("Outcome", final.covarites, var.name))
    
    # 若当前变量小于 5 个 unique 值，转成 factor（可酌情去除）
    if (length(unique(reg.df[[var.name]])) < 5){
      reg.df[[var.name]] <- as.factor(reg.df[[var.name]])
    }
    
    if (debug.activated){
      print(paste("Now fitting variable #", n.var, ":", var.name))
    }
    
    # 使用线性回归
    model.fit <- lm(Outcome ~ ., data = reg.df)
    
    # 提取回归系数、p 值等
    coff.df <- tidy(model.fit)
    # 重命名第 1 列以保持一致
    names(coff.df)[1] <- "var"
    
    # 格式化 p-value
    coff.df$p.format <- sapply(coff.df$p.value, format.p.value)
    # 若有 convert.p.value.to.signif() 等自定义函数，也可使用
    coff.df$p.signif <- sapply(coff.df$p.value, convert.p.value.to.signif)
    
    # 检查是否有 NA coef 或者 confint 出错
    conf_int <- tryCatch(confint(model.fit), error = function(e) return(NA))
    if (any(is.na(conf_int)) || any(is.na(coef(model.fit)))){
      # 若出现 NA，则跳过该变量
      skip.vars.internal <- c(skip.vars.internal, var.name)
    } else {
      # 计算置信区间
      conf.df <- as.data.frame(conf_int)
      colnames(conf.df) <- c("conf.low", "conf.high")
      conf.df <- cbind(var = rownames(conf.df), conf.df)
      rownames(conf.df) <- NULL
      
      # 合并到 broom 的 tidy 表格
      fit.out.df <- merge(coff.df, conf.df, by = "var", all = FALSE)
      
      # 按原始顺序进行排序
      fit.out.df.sorted <- fit.out.df[order(match(fit.out.df$var, coff.df$var)), ]
      fit.out.df.sorted$VarIndex <- n.var
      fit.out.df.sorted$VarName <- var.name
      # 标识该行是否是主变量（即本次 loop 的测试变量），而非协变量
      fit.out.df.sorted$Target <- TRUE
      fit.out.df.sorted$Target[1:(length(final.covarites) + 1)] <- FALSE
      
      # 重新设置 rownames
      rownames(fit.out.df.sorted) <- seq(1, nrow(fit.out.df.sorted), 1)
      
      # 若开启 debug，打印结果
      if (debug.activated){
        print(fit.out.df.sorted)
        print("-----------------------------------------------------------------------------------------------------------")
      }
      
      # 将结果合并进总表
      output <- rbind(output, fit.out.df.sorted)
    }
  }
  
  # 如果有变量被跳过，打印警告信息
  if (length(skip.vars.internal) > 0){
    print(paste(
      "Warning:", length(skip.vars.internal), 
      "variables were excluded from the regression analysis due to NA values in coefficient or confint calculation!"
    ))
    print(skip.vars.internal)
  }
  
  # 若仅想返回显著结果，可在此处自行做筛选
  if (only.return.signif.tab) {
    # 例如：可筛选 p < 0.05 的行
    # output <- subset(output, p.value < 0.05)
    
    return(output)
  } else {
    return(output)
  }
}




############################################################################################################################################
## save a figure to a local PPT file
## ggfig: the ggplot figure
## path: the local path to save the figure
## vertical: the format of the PPT file, vertically or horizontally 
## img.width: the width of the figure in inches, maximal 8.3
## img.height: the height of the figure in inches, maximal 11.7
## as.image: if TRUE, save the ggplot as a raster image; if FALSE, save it as editable PowerPoint elements
############################################################################################################################################
save.fig.to.pptx <- function(ggfig, path, append = FALSE, vertical = FALSE, img.width = 6, img.height = 4, as.image=TRUE) {
  library(officer)
  library(ggplot2)
  library(rvg)
  
  # Ensure the object is a ggplot object
  if (!inherits(ggfig, "ggplot")) {
    stop("The object passed as 'ggfig' is not a ggplot object.")
  }
  
  # Load or create PowerPoint based on the append parameter
  if (append && file.exists(path)) {
    ppt <- read_pptx(path)
  } else {
    # Load the appropriate custom template
    if (missing(vertical) || !vertical) {
      ppt <- read_pptx(path = "C:/local_work_ming/workspaces/r_workspace/A4 horizontal template.pptx")
    } else {
      ppt <- read_pptx(path = "C:/local_work_ming/workspaces/r_workspace/A4 vertical template.pptx")
    }
  }
  
  # Ensure the object is a ggplot object
  if (!inherits(ggfig, "ggplot")) {
    stop("The object passed as 'ggfig' is not a ggplot object.")
  }
  
  # Optional: Print layout summary for debugging
  layout_summary(ppt)
  
  # Retrieve slide dimensions
  slide_size <- slide_size(ppt)
  slide_width <- slide_size$width
  slide_height <- slide_size$height
  
  # Calculate position to center the image
  left_position <- (slide_width - img.width) / 2
  top_position <- (slide_height - img.height) / 2
  
  # Add a blank slide (adjust based on your template)
  ppt <- add_slide(ppt, layout = "Leer", master = "Office")
  
  # Add the plot to the slide, either as an image or as editable elements
  if (as.image) {
    # Render the ggplot as a raster image in memory using a png device
    img <- tempfile(fileext = ".png")
    png(img, width = img.width, height = img.height, units = "in", res = 300)
    print(ggfig)
    dev.off()
    
    # Read the in-memory PNG file and embed it as an image
    ppt <- ph_with(ppt, external_img(img), 
                   location = ph_location(left = left_position, top = top_position, width = img.width, height = img.height))
    
    # Remove the temporary file
    unlink(img)
  } else {
    # Add as editable vector graphic (PowerPoint elements)
    ppt <- ph_with(ppt, dml(ggobj = ggfig), 
                   location = ph_location(left = left_position, top = top_position, width = img.width, height = img.height))
  }
  
  # Save the PowerPoint file
  print(ppt, target = path)
  
  message(paste("A figure has been saved to", path))
}

# save.fig.to.pptx <- function(ggfig, path, append, vertical, img.width, img.height){
#   library(officer)
#   library(ggplot2)
#   library(rvg)
#   
#   if (append&& file.exists(path)){
#     ppt <- read_pptx(path)
#   } else {
#     if (missing(vertical)||!vertical){
#       # Load the custom A4 template
#       ppt <- read_pptx(path = "C:/local_work_ming/workspaces/r_workspace/A4 horizontal template.pptx")
#     } else {
#       # Load the custom A4 template
#       ppt <- read_pptx(path = "C:/local_work_ming/workspaces/r_workspace/A4 vertical template.pptx")
#     }
#   }
#   
# 
#   #ppt.path <- paste(output.folder, "figure.pptx", sep = "/")
#   
#   # Check available layouts and masters
#   layout_summary(ppt)
#   
#   # Retrieve slide dimensions
#   slide_size <- slide_size(ppt)
#   slide_width <- slide_size$width
#   slide_height <- slide_size$height
#   
#   # Calculate position to center the image
#   left_position <- (slide_width - img.width) / 2
#   top_position <- (slide_height - img.height) / 2
#   
#   # Add a slide to the PowerPoint
#   ppt <- add_slide(ppt, layout = "Leer", master = "Office")
#   
#   # Add the plot to the slide, centered
#   ppt <- ph_with(ppt, dml(ggobj = ggfig), 
#                  location = ph_location(left = left_position, top = top_position, width = img.width, height = img.height))
#   
#   # Save the PowerPoint file
#   print(ppt, target = path)
#   
#   print(paste("A figure has been saved to", path))
# }

############################################################################################################################################
## save a data frame to a local PPT file
## df: the data frame
## path: the local path to save the figure
## vertical: the format of the PPT file, vertically or horizontally 
## img.width: the width of the figure in inches, maximal 8.3
## img.height: the height of the figure in inches, maximal 11.7
############################################################################################################################################
save.df.to.pptx <- function(df, path, append = FALSE, vertical = FALSE, df.width = 10.5, df.height = 7, title.text = "Data Frame") {
  library(ggplot2)
  library(rvg)
  library(officer)
  library(flextable)
  
  # Handle appending or creating a new PowerPoint file
  if (append && file.exists(path)) {
    ppt <- read_pptx(path)
  } else {
    if (vertical) {
      ppt <- read_pptx(path = "C:/local_work_ming/workspaces/r_workspace/A4 vertical template.pptx")
    } else {
      ppt <- read_pptx(path = "C:/local_work_ming/workspaces/r_workspace/A4 horizontal template.pptx")
    }
  }
  layout <- layout_summary(ppt)
  #print(layout)
  
  # Retrieve slide dimensions
  slide_size <- slide_size(ppt)
  slide_width <- slide_size$width
  slide_height <- slide_size$height
  
  # Calculate position to center the image
  left_position <- (slide_width - df.width) / 2
  top_position <- (slide_height - df.height) / 2
  
  # Add a slide to the PowerPoint
  ppt <- add_slide(ppt, layout = "Titel und Inhalt", master = "Office")
  
  # Add a title
  ppt <- ph_with(ppt, value = title.text, location = ph_location_type(type = "title"))
  
  # Convert the data frame to a flextable
  ft <- regulartable(df)
  
  # Add the flextable to the slide
  ppt <- ph_with(ppt, value = ft, location = ph_location_type(type = "body"))
  
  # Save the PowerPoint file
  print(ppt, target = path)
  
  print(paste("A data frame has been saved to", path))
}

save.df.to.pptx2 <- function(df, path, append = FALSE, vertical = FALSE, df.width = 10.5, df.height = 7, title.text = "Data Frame") {
  library(ggplot2)
  library(rvg)
  library(officer)
  library(flextable)
  
  # Handle appending or creating a new PowerPoint file
  if (append && file.exists(path)) {
    ppt <- read_pptx(path)
  } else {
    if (vertical) {
      ppt <- read_pptx(path = "C:/local_work_ming/workspaces/r_workspace/A4 vertical template.pptx")
    } else {
      ppt <- read_pptx(path = "C:/local_work_ming/workspaces/r_workspace/A4 horizontal template.pptx")
    }
  }
  
  # Retrieve slide dimensions
  slide_size <- slide_size(ppt)
  slide_width <- slide_size$width
  slide_height <- slide_size$height
  
  # Calculate position to center the table
  left_position <- (slide_width - df.width) / 2
  top_position <- (slide_height - df.height) / 2
  
  # Add a slide to the PowerPoint
  ppt <- add_slide(ppt, layout = "Titel und Inhalt", master = "Office")
  
  # Add a title
  ppt <- ph_with(ppt, value = title.text, location = ph_location_type(type = "title"))
  
  # Convert the data frame to a flextable
  ft <- regulartable(df)

  # Add the flextable to the slide with specified width and height
  ppt <- ph_with(ppt, value = ft, location = ph_location(left = left_position, top = top_position, width = df.width, df.height = 7, newlabel = "test label", ))

  # Save the PowerPoint file
  print(ppt, target = path)
  
  print(paste("A data frame has been saved to", path))
}


#### format a data frame according to the provided column properties
format.df.func <- function(df, numeric_cols, factor_cols){
  if (is.null(df)){
    return()
  }
  
  if (!all(numeric_cols %in% colnames(df))){
    print(paste("Following numeric columns are not included in the data frame:", numeric_cols[!numeric_cols %in% colnames(df)]))
    return()
  }
  
  if (!all(factor_cols %in% colnames(df))){
    print(paste("Following factor columns are not included in the data frame:", numeric_cols[!numeric_cols %in% colnames(df)]))
    return()
  }
  
  print("------ Pre Formatting --------------------")
  print(str(df))
  
  # 使用 lapply 批量转换为 numeric
  df[numeric_cols] <- lapply(df[numeric_cols], as.numeric)
  
  # 使用 lapply 批量转换为 factor
  df[factor_cols] <- lapply(df[factor_cols], as.factor)
  
  # 查看转换后的数据框
  print("------ Post Formatting --------------------")
  print(str(df))
  return(df)
}



# Function to calculate the difference between two timestamp strings in hours
timestamp_diff_in_hours <- function(timestamp1, timestamp2) {
  # Convert the timestamp strings to POSIXct objects
  time1 <- as.POSIXct(timestamp1, format="%Y-%m-%d %H:%M")
  time2 <- as.POSIXct(timestamp2, format="%Y-%m-%d %H:%M")
  
  # Calculate the difference in hours
  diff_hours <- as.numeric(difftime(time1, time2, units = "hours"))
  
  return(diff_hours)
}

find.rows.contain.NAs <- function(data){
  # Find rows with NA values
  rows_with_na <- data[!complete.cases(data), ]
  
  # Display rows with NA values
  return(rows_with_na)
}

# Define the function to find rows and columns with NAs
find_na_positions <- function(df) {
  # Initialize an empty data frame to store the results
  na_positions <- data.frame(
    row = integer(),
    column = character(),
    stringsAsFactors = FALSE
  )
  
  # Loop through each element to find NA values
  for (i in seq_len(nrow(df))) {
    for (j in seq_len(ncol(df))) {
      if (is.na(df[i, j])) {
        # Add row and column information to na_positions data frame
        na_positions <- rbind(na_positions, data.frame(row = i, column = names(df)[j]))
      }
    }
  }
  
  # Return the data frame with NA positions
  return(na_positions)
}





