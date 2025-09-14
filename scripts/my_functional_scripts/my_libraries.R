library(rJava)
library("sets")

library(gridExtra)
library(reshape2) # to use melt
library(ggpubr) # to use ggarrange
library(png) # to read images
library(jpeg) # to read images Please make sure the package 'jpeg' was installed on the PC with command install.packages('jpeg')
library(imager) # for load.image()
library(grid)
library(ggnewscale) # for generate another legend based on the same demonision
library (egg) # make consistent-width plots in ggplot (even with -/out legends)
library("readxl") #read xls/xlsx file
library(ggrepel)
library("corrplot")
library("psych") # for corr.test()
library("PerformanceAnalytics") # for chart.Correlation()
library("Hmisc") # for rcorr()
library("mlogit") #mlogit
library("car") #scatterplotMatrix()
library(gmodels)
library(caret)# for confusionMatrix, train, predict...
library(gbm) # for Stochastic Gradient Boosting model
library(party) # for C5.0 model
library(mboost)# for C5.0 model
library(plyr)# for C5.0 model
library(partykit)# for C5.0 model
library(rpart)# for rpart model
library(MLeval) # ROC
#install.packages("VIM")
library(VIM)# library to check the NA values
library(plotROC)
library(ROCR)
library(pROC)
library("GGally")                      # Load GGally package
library(knitr)
library(xlsx)
library(ggdist) # for rain cloud plot
library(tidyquant) # for rain cloud plot
library(tidyverse) # for rain cloud plot
library(ggtext)
library(gridExtra)
library(grid)
library(ragg)
library(gghalves)
library(pracma)

###----- for between-group co-variates matching----------------
# #library("MatchIt")
# library(optmatch)
# library(rgenoud)
# #library(Matching)
# library(Rglpk)
# library(quickmatch)

###----- for data visualization ----------------
library(ggplot2) ## basic
#library(fmsb) ## radar chart
library(ggbeeswarm) ## bee swarm chart
library(ggdist) ## rain cloud chart

###----- for machine learning ----------------
library(randomForest)
library("glmnet") # LASSO logistic regression

###----- for ROC-AUC calculation and plot ----------------
library(MLeval)

## visualize odds ratios
library(OddsPlotty)
## library to get R square
library(rsq)


# 加载必要的包
library(dplyr)
library(lubridate)