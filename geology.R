#!/usr/bin/env Rscript
suppressPackageStartupMessages(library(dplyr))
suppressPackageStartupMessages(library(argparse))
suppressPackageStartupMessages(library(readxl))
suppressPackageStartupMessages(library(xlsx))

# parser <- ArgumentParser()
# parser$add_argument("file", help = "Name of file output from IPC")
# parser$add_argument("weights", help = "Name of file containing weights of foams")
# parser$add_argument("-l", "--limits", default = "/Users/fordfishman/Documents/Geology_program/MDL-LOQ limits.xlsx",
#                     help = "Name of file containing MDL, PQL,LOQ,MRL, and IDL information [default %(default)s]")
# parser$add_argument("-b", "--background", default = "/Users/fordfishman/Documents/Geology_program/Background foam weights and concentrations.xlsx",
#                     help = "Name of file containing background foam weight and concentration information [default %(default)s]")
# parser$add_argument("-o", "--output", default = "IPC_analysis.xlsx", type="character",
#                     help = "Name desired for output file (don't include extension) [default %(default)s]")
# parser$add_argument("-r", "--rsquared", default = 0.95, type="double",
#                     help = "Cutoff rsquared for drift correction [default %(default)s]")
# parser$add_argument("-s", "--sample_start", default = "A", type="character",
#                     help = "Set of characters that each sample ID always starts with [default %(default)s]")
# parser$add_argument("-e", "--sample_end", default = "C", type="character",
#                     help = "Set of characters that each sample ID always ends with [default %(default)s]")
# 
# args <- parser$parse_args()
# file <- args$file
# limits <- args$limits
# background <- args$background
# weights <- args$weights
# output <- args$output
# if (! endsWith(output, ".xlsx")) {
#   output <- paste(output, "xlsx", sep = ".")
# }
# rsquare_cutoff <- args$rsquared
# sample_start <- args$sample_start
# sample_end <- args$sample_end

# CHANGE THIS TO READ IN ARGUMENTS ON THE COMMAND LINE
path = "~/Documents/Geology_program/" 
file <- "/Users/fordfishman/Documents/Geology_program/Field Foams 4,12,13,14 rerun,145,147,156,157,158,160,161.xlsx"
limits <- "/Users/fordfishman/Documents/Geology_program/MDL-LOQ limits.xlsx"
weights <- "/Users/fordfishman/Documents/Geology_program/Masses of dried field foams (mod 1-06.27.19).xlsx"
rsquare_cutoff <- 0.94
sample_start <- "A"
sample_end <- "C"
output <- "/Users/fordfishman/Documents/Geology_program/IPC_analysis.xlsx"
load("/Users/fordfishman/Documents/Geology_program/IPC-analysis-app/data/IPC_data.Rdata") # This data includes background foam concentrations and weights


# Change directory to where your file is
# setwd(path)
# Saves the sheet with concentration data into a dataframe
field_foams_concentrations_df <- read_excel(path = file, sheet = "Conc. in Calib Units")

# Rename some columns to remove ambiguous characters
field_foams_concentrations_df$Sample <- as.integer(field_foams_concentrations_df$..1)
colnames(field_foams_concentrations_df) <- gsub("\r", "", x = colnames(field_foams_concentrations_df))
colnames(field_foams_concentrations_df) <- gsub("\n", "", x = colnames(field_foams_concentrations_df))
colnames(field_foams_concentrations_df) <- gsub(" ", "", x = colnames(field_foams_concentrations_df))
colnames(field_foams_concentrations_df) <- gsub("mg", "µg", x = colnames(field_foams_concentrations_df))
colnames(field_foams_concentrations_df) <- gsub("As193\\.69\\(", "As193\\.696\\(", x = colnames(field_foams_concentrations_df))

# Saves sheet with Y standard information
internal_standards_df <- read_excel(path = file, sheet = "Internal Standards")

# Adds Y information to concentration dataframe
colnames(internal_standards_df) <- c("RowIndex", "Y_levels")
field_foams_concentrations_df <- mutate(field_foams_concentrations_df, Y_levels = internal_standards_df$Y_levels)

# Alter sample names to remove " " to remove inconsistent naming.
field_foams_concentrations_df$SampleId <- gsub(" ", "", field_foams_concentrations_df$SampleId)

# Grab Y values for IPC-500 runs
important_Y <- subset(field_foams_concentrations_df, 
                      startsWith(SampleId, "IPC-500"), 
                      select = c(Sample, Y_levels))

# The following is just making variable names for the drift correction regressions
reg_group <- c() # Will contain names of subsets of data
lm_name <- c() # Will contains names of regressions
for (i in 1:(nrow(important_Y)-1)){
  group_name <- paste("group", as.character(i), sep = "")
  model <- paste("lm", as.character(i), sep = "")
  lm_name <- c(lm_name, model)
  reg_group <- c(reg_group, group_name)
}

# Getting linear models for the drift corrections
functions_list <- list() # A list of model names and the sample numbers they correspond with
START <-  1
END <- nrow(important_Y) 

for (i in 1:length(lm_name)){
  # Using all data for the regression at first
  assign(reg_group[i], important_Y[START:END,])
  assign(lm_name[i], lm(data = get(reg_group[i]), formula = Y_levels ~ `Sample`))
  # Removes the last entry in the set of data until the r-squared is above the cutoff
  # Can end up only having 2 data points, where r-squared is 1
  below_sig <- summary(get(lm_name[i]))$r.squared < rsquare_cutoff
  if (is.na(below_sig)) {
    below_sig <- F
  }
  while (below_sig) {
    END <- END - 1
    assign(reg_group[i], important_Y[START:END,])
    assign(lm_name[i], lm(data = get(reg_group[i]), formula = Y_levels ~ `Sample`))
    below_sig <- summary(get(lm_name[i]))$r.squared < rsquare_cutoff
    if (is.na(below_sig)) {
      below_sig <- F
    }
  }
  # Creates a linear formula when r-squared meets cutoff
  assign(paste("b", i, sep = ""), coef(get(lm_name[i]))[1])
  assign(paste("m", i, sep = ""), coef(get(lm_name[i]))[2])
  assign(paste("y", i, sep = ""), 
         function(x) 
           get(paste("m", i, sep = "")) * x + get(paste("b", i, sep = "")))
  ##Case 1: linear model covers all data
  if (START == 1 && END == nrow(important_Y)){
    functions_list[[paste("y", i, sep = "")]] <- c(1:nrow(field_foams_concentrations_df))
    ##Case 2: linear model starts at beginning of data but stops before end
  } else if (START == 1 && END != nrow(important_Y)) {
    functions_list[[paste("y", i, sep = "")]] <- c(1:as.integer(important_Y[END, 1]))
    ##Case 3: linear model doesn't start at beginning and it stops before end
  } else if (START != 1 && END != nrow(important_Y)){
    functions_list[[paste("y", i, sep = "")]] <- c((as.integer(important_Y[START, 1]) + 1):as.integer(important_Y[END, 1]))
    ##Case 4: linear model doesn't start at beginning but stops at end
  } else {
    functions_list[[paste("y", i, sep = "")]] <- c((as.integer(important_Y[START, 1]) + 1):nrow(field_foams_concentrations_df))
  }
  # The loop will start again where the last model left off
  # If the last model made it to the end, the loop will break off prematurely
  START <- END
  END <- nrow(important_Y)
  if (START == END){
    break
  }
}
drift_multiplier <- c()

# Iteratres through and gives the correct drift multiplier to the correct sample
for (i in 1:length(functions_list)) {
  for (num in functions_list[i]){
    drift_multiplier <- c(drift_multiplier, get(names(functions_list[i]))(num))
  }
}
# Reorganization of dataframe
field_foams_concentrations_df <- mutate(field_foams_concentrations_df, drift_multiplier = drift_multiplier)
drift_corrections <-  data.frame(n = field_foams_concentrations_df$..1, id = field_foams_concentrations_df$SampleId)
for (i in names(field_foams_concentrations_df)) {
  if (is.numeric(field_foams_concentrations_df[[i]]) & endsWith(i, ")")) {
    drift_corrections[[i]] <-  field_foams_concentrations_df[[i]] * (field_foams_concentrations_df$drift_multiplier)^-1
  }
}

# Create a separate dataframe with just CCB data
CCB_data <-  subset(drift_corrections, id == "CCB")

# Averaging the CCB numbers
background_corrections <- data_frame(n = drift_corrections$n, id = drift_corrections$id)
for (i in names(CCB_data)) {
  if (is.numeric(CCB_data[[i]]) && i != "n") {
    mean_i <- mean(CCB_data[[i]])
    background_corrections[[i]] = drift_corrections[[i]] - mean_i
  }
}

# Remove Y column
ids <-  background_corrections$id
corrected_data <- subset(background_corrections, select = -c(3))
id_info <- subset(background_corrections, select = c(1,2))
corrected_data <- corrected_data[,3:ncol(corrected_data)]
# corrected_data <- corrected_data %>% select(`As188.979(µg/L)`:`Cu324.752(mg/L)`)
corrected_data <- corrected_data[,order(colnames(corrected_data))]
analytes = colnames(corrected_data)
corrected_data <- cbind(id_info, corrected_data)

## QA/QC

# Make sure that only analytes present are used in the analysis
IDL <- IDL[analytes]
MDL <- MDL[analytes]
PQL_LOQ_MRL <- PQL_LOQ_MRL[analytes]

# General function that can compare several quanitities to its desired range
# Ex. n = 83:86, comp = 100, lower = 0.85, upper = 1.15
# Invalid Invalid Valid Valid
quality <-  function(vec, comp, lower, upper){
  validity <- c()
  lower_limit <- comp * lower
  upper_limit <- comp * upper
  if (length(comp) > 1){
    for (i in 1:length(vec)){
      if (vec[i] >= lower_limit[i] && vec[i] <= upper_limit[i]){
        validity <- c(validity, "Valid")
      } else{
        validity <- c(validity, "Invalid")
      }
    }
  } else{
    for (n in vec){
      if (n >= lower_limit && n <= upper_limit){
        validity <- c(validity, "Valid")
      } else{
        validity <- c(validity, "Invalid")
      }
    }
  }
  return(validity)
}

# To temporarily store information for each check
QA_QC <- list()

# Column number where actual concentration data begins
data_range <- 3:length(corrected_data)
IPC500 <- 1
LFM <- 1
CCB <- 1
IPC_comp <- c()
QCS_comp <-  c()
SIC_comp <- c()
LFB_comp <- c()
for (i in colnames(corrected_data[,data_range])) {
  if (startsWith(i, "Fe")) {
    IPC_comp <- c(IPC_comp, 2500)
    QCS_comp <- c(QCS_comp, 1250)
    SIC_comp <- c(SIC_comp, 125 *5)
    LFB_comp <- c(LFB_comp, 125 *5)
    
  } else if (startsWith(i, "Zn")){
    IPC_comp <- c(IPC_comp, 1250)
    QCS_comp <- c(QCS_comp, 625)
    SIC_comp <- c(SIC_comp, 125 * 2.5)
    LFB_comp <- c(LFB_comp, 125 *2.5)
    
  } else {
    IPC_comp <- c(IPC_comp, 500)
    QCS_comp <- c(QCS_comp, 250)
    SIC_comp <- c(SIC_comp, 125)
    LFB_comp <- c(LFB_comp, 125)
  }
}
# Actual QA/QC procress
for (i in 1:nrow(corrected_data)){
  id <- corrected_data$id[i]
  # IPC-500 (1st one) must be 95-105% of 500 ppb of analyte
  if (id == "IPC-500" && IPC500 == 1) {
    QA_QC[["IPC1"]] <- quality(vec = corrected_data[i, data_range], comp = IPC_comp, lower = 0.95, upper = 1.05)
    IPC500 <- IPC500 + 1
    # IPC-500 (2nd, 3rd, 4th, etc.) must be 90-110% of 500 ppb of analyte
  } else if (id == "IPC-500") {
    name <- paste("IPC", IPC500, sep = "")
    QA_QC[[name]] <- quality(vec = corrected_data[i, data_range], 
                             comp = IPC_comp, 
                             lower = 0.90, 
                             upper = 1.10)
    IPC500 <- IPC500 + 1
    # LFM(-1,-2,-3) must be 85-115% of predicted spike concentration
    # sample id must contain LFM
  } else if (grepl("LFM", id)){
    name <- paste("LFM", LFM, sep = "")
    comp <- corrected_data[i-1, data_range]/2 + 125
    comp["Zn206.200(µg/L)"] <- comp["Zn206.200(µg/L)"] - 125 + (125 * 2.5)
    comp["Fe238.204(µg/L)"] <- comp["Fe238.204(µg/L)"] - 125 + (125 * 5)
    QA_QC[[name]] <- quality(vec = corrected_data[i, data_range], 
                             comp = comp, 
                             lower = 0.85, 
                             upper = 1.15)
    LFM <- LFM + 1
    # CCB must be < IDL
  } else if(grepl("CCB", id)){
    name <- paste("CCB", CCB, sep = "")
    result <- as.character(corrected_data[i, data_range] < IDL)
    result <- gsub("FALSE", "Invalid", x = result)
    result <- gsub("TRUE", "Valid", x = result)
    QA_QC[[name]] <- result
    CCB <- CCB + 1
    # LRB must be < 2.2 xMDL
  } else if(grepl("LRB", id)){
    result <- as.character(corrected_data[i, data_range] < (2.2*MDL))
    result <- gsub("FALSE", "Invalid", x = result)
    result <- gsub("TRUE", "Valid", x = result)
    QA_QC[["LRB"]] <- result
    # SIC-125 must be 80-120% of 125 ppb of analyte.
  } else if (grepl("SIC", id)) {
    QA_QC[["SIC"]] <- quality(vec = corrected_data[i, data_range], 
                              comp = SIC_comp, 
                              lower = 0.80, 
                              upper = 1.20)
    # QCS-250 must be 95-105% of 250 ppb of analyte
  } else if (grepl("QCS", id)) {
    QA_QC[["QCS"]] <- quality(vec = corrected_data[i, data_range], 
                              comp = QCS_comp, 
                              lower = 0.95, 
                              upper = 1.05)
    # LFB must be 85-115% of 125 ppb of analyte
  } else if (grepl("LFB", id)) {
    QA_QC[["LFB"]] <- quality(vec = corrected_data[i, data_range], 
                              comp = LFB_comp, 
                              lower = 0.85, 
                              upper = 1.15)
  }
}

# Finalizing the QA/QC dataframe
QA_QC_df <- as.data.frame(t(as.data.frame(QA_QC)))
colnames(QA_QC_df) <- colnames(corrected_data[data_range])
QA_QC_df$id <- rownames(QA_QC_df)
IPCQ <- subset(QA_QC_df, startsWith(id, "IPC"), select = -c(id))
LFMQ <- subset(QA_QC_df, startsWith(id, "LFM"), select = -c(id))
CCBQ <- subset(QA_QC_df, startsWith(id, "CCB"), select = -c(id))
QA_QC_final <- subset(QA_QC_df, !startsWith(id, "IPC") & !startsWith(id, "LFM") & !startsWith(id, "CCB"), select = -c(id))

# Combine IPC, LFM, and CCB runs
IPC_tot <- c()
for (i in names(IPCQ)){
  if (length(unique(IPCQ[[i]])) == 1 && unique(IPCQ[[i]]) == "Valid"){ 
    IPC_tot <- c(IPC_tot, "Valid")
  } else {
    IPC_tot <- c(IPC_tot, "Invalid")
  }
}

LFM_tot <- c()
for (i in names(LFMQ)){
  if (length(unique(LFMQ[[i]])) == 1 && unique(LFMQ[[i]]) == "Valid"){
    LFM_tot <- c(LFM_tot, "Valid")
  } else {
    LFM_tot <- c(LFM_tot, "Invalid")
  }
}
CCB_tot <- c()
for (i in names(CCBQ)){
  if (length(unique(CCBQ[[i]])) == 1 && unique(CCBQ[[i]]) == "Valid"){
    CCB_tot <- c(CCB_tot, "Valid")
  } else {
    CCB_tot <- c(CCB_tot, "Invalid")
  }
}
# Final output for QA/QC
QA_QC_final <- rbind(QA_QC_final, IPC = IPC_tot, LFM = LFM_tot, CCB = CCB_tot)

############################## ANALYSIS #############################################

weights_df <- read_excel(weights)
# Fixing column names to match weights with appropriate samples
weights_df$`Foam ID` <- gsub(" ", "-", weights_df$`Foam ID`)
weights_df$`Foam ID` <- paste("A", weights_df$`Foam ID`, sep = "")
corrected_data$id <- as.character(corrected_data$id)
sample_df <- subset(corrected_data, startsWith(id, sample_start) & endsWith(id, sample_end), select = -c(n))
sample_df <- sample_df[order(sample_df$id),]
sample_weights <- subset(weights_df, `Foam ID` %in% sample_df$id)
sample_num <- c()
for (name in sample_df$id){
  parts <- strsplit(name, "-")
  parts2 <- strsplit(parts[[1]][1], "A")
  sample_num <- c(sample_num, parts2[[1]][2])
}
# Divide by analyzed foam weight
sample_weights$sample_num <- as.factor(sample_num)
Aweights_means <- with(sample_weights, tapply(`Weight of half of foam analyzed (A) (g)`, sample_num, mean))
Tweights_means <- with(sample_weights, tapply(`Total weight of foam (g)`, sample_num, mean))
Aweights_sds <- with(sample_weights, tapply(`Weight of half of foam analyzed (A) (g)`, sample_num, sd))
Tweights_sds <- with(sample_weights, tapply(`Total weight of foam (g)`, sample_num, sd))
sample_norm <- cbind(sample = as.numeric(sample_num),sample_df[,2:ncol(sample_df)]/sample_weights$`Weight of half of foam analyzed (A) (g)`)
sample_df <- cbind(sample = as.numeric(sample_num), sample_df[,2:ncol(sample_df)])

# Divide background foam concentrations by foam weight
bg1_weights <- subset(background1, 
                      select = c(`Ahalfweight(g)`,
                                 `Bhalfweight(g)`,
                                 `Totalweight(g)`)
)
bg1_conc <- subset(background1, select = analytes)
bg2_weights <- subset(background2, 
                      select = c(`Ahalfweight(g)`,
                                 `Bhalfweight(g)`,
                                 `Totalweight(g)`)
)
bg2_conc <- subset(background2, select = analytes)
bg3_weights <- subset(background3, 
                      select = c(`Ahalfweight(g)`,
                                 `Bhalfweight(g)`,
                                 `Totalweight(g)`)
)
bg3_conc <- subset(background3, select = analytes)
# Grab concentrations, means and SDs for the background data
bg1_conc <- bg1_conc[,order(colnames(bg1_conc))]
bg1_conc_mean <- colMeans(bg1_conc, na.rm = T)
bg1_conc_sd <- apply(bg1_conc, 2, sd, na.rm=TRUE)
bg2_conc <- bg2_conc[,order(colnames(bg2_conc))]
bg2_conc_mean <- colMeans(bg2_conc, na.rm = T)
bg2_conc_sd <- apply(bg2_conc, 2, sd, na.rm=TRUE)
bg3_conc <- bg1_conc[,order(colnames(bg3_conc))]
bg3_conc_mean <- colMeans(bg3_conc, na.rm = T)
bg3_conc_sd <- apply(bg3_conc, 2, sd, na.rm=TRUE)
bg1_norm <- bg1_conc/bg1_weights$`Bhalfweight(g)`
bg2_norm <- bg2_conc/bg2_weights$`Bhalfweight(g)`
bg3_norm <- bg3_conc/bg3_weights$`Bhalfweight(g)`

# t-test: sample concentration compared to background foam concentrations
sample_list <- c(unique(sample_norm$sample))

# Defines a function to compare values
sig_test <- function(sample_df, background_df) {
  p_values <- c()
  for (i in colnames(sample_df)){
    if (length(sample_df[[i]]) < 2){
      p_values <- c(p_values, 1)
    } else {
      p_values <- c(p_values, t.test(sample_df[[i]], background_df[[i]], alternative = "greater")$p.value)
    }
  }
  return(p_values)
}

# How to label each analyte of each sample for the final main spreadsheet
final_label <- function(signif_v, nondetect, only_detect, give_value, final_conc) {
  final <- c()
  for (i in 1:length(signif_v)) {
    if (!signif_v[i]) {
      final <- c(final, "NotSignificant")
    } else if (nondetect[i]) {
      final <- c(final, "Non-detect")
    } else if (only_detect[i]) {
      final <- c(final, "Present")
    } else if (give_value[i]){
      final <- c(final, final_conc[i])
    }
  }
  return(final)
}

# If concentration <= MDL = non-detect. 
# If concentration > MDL, analyte is 99% certain to be present but not quantifiable (i.e., J-values).  
# Quantifiable values are > PQL,LOQ,MRL.

for (i in 1:length(sample_list)) {
  sample_number <- sample_list[i]
  norm_df <- paste(sample_start, sample_number, "_norm" ,sep = "")
  conc_df <- paste(sample_start, sample_number, "_conc" ,sep = "")
  assign(norm_df, subset(sample_norm, sample == sample_number, select = -c(sample)))
  assign(conc_df, subset(sample_df, sample == sample_number, select = -c(sample)))
  conc_mean <- apply(get(conc_df), 2, mean)
  conc_sd <- apply(get(conc_df), 2, sd)
  nondetect <- conc_mean <= MDL
  only_detect <- (conc_mean > MDL) & (conc_mean <= PQL_LOQ_MRL)
  give_value <- conc_mean > PQL_LOQ_MRL
  
  if (sample_number < 198) {
    p_values <- sig_test(sample_df = get(norm_df), background_df = bg1_norm)
    final_conc <- (conc_mean - bg1_conc_mean)
    final_sd <- sqrt(conc_sd^2 + bg1_conc_sd^2 - 2 * conc_sd * bg1_conc_sd)
  } else if (sample_number >= 198 & sample_number < 213) {
    p_values <- sig_test(sample_df = get(norm_df), background_df = bg2_norm)
    final_conc <- (conc_mean - bg2_conc_mean)
    final_sd <- sqrt(conc_sd^2 + bg2_conc_sd^2 - 2 * conc_sd * bg2_conc_sd)
  } else if (sample_number >= 213 & sample_number < 317){
    p_values <- sig_test(sample_df = get(norm_df), background_df = bg3_norm)
    final_conc <- (conc_mean - bg3_conc_mean)
    final_sd <- sqrt(conc_sd^2 + bg3_conc_sd^2 - 2 * conc_sd * bg3_conc_sd)
  }
  weight_ratio <- 3/7 * (Tweights_means[as.character(sample_number)]/Aweights_means[as.character(sample_number)])
  weight_sd <-  3/7 * sqrt((Tweights_sds[as.character(sample_number)]/Tweights_means[as.character(sample_number)])^2 + (Aweights_sds[as.character(sample_number)]/Aweights_means[as.character(sample_number)])^2 - 2 * (Aweights_sds[as.character(sample_number)] * Tweights_sds[as.character(sample_number)])/(Aweights_means[as.character(sample_number)] * Tweights_means[as.character(sample_number)]))
  final_sd <- sqrt((final_sd/final_conc)^2 + (weight_sd/weight_ratio)^2 + 2 * (weight_sd * final_sd)/(weight_ratio * final_conc))
  final_conc <- final_conc * weight_ratio
  signif_v <- p_values < 0.05
  final_important <- final_label(signif_v, nondetect, only_detect, give_value, final_conc)
  if (i == 1){
    all_p_values <- matrix(p_values, ncol = length(p_values))
    signif_m <- matrix(signif_v, ncol = length(p_values))
    nondetect_m <- matrix(nondetect, ncol = length(nondetect))
    only_detect_m <- matrix(only_detect, ncol = length(only_detect))
    give_value_m <- matrix(give_value, ncol = length(give_value))
    final_conc_m <- matrix(final_conc, ncol = length(final_conc))
    final_sd_m <- matrix(final_sd, ncol = length(final_sd))
    final_important_m <- matrix(final_important, ncol = length(final_important))
  } else {
    all_p_values <- rbind(all_p_values, p_values)
    signif_m <- rbind(signif_m, signif_v)
    nondetect_m <- rbind(nondetect_m, nondetect)
    only_detect_m <- rbind(only_detect_m, only_detect)
    give_value_m <- rbind(give_value_m, give_value)
    final_conc_m <- rbind(final_conc_m, final_conc)
    final_important_m <- rbind(final_important_m, final_important)
    final_sd_m <- rbind(final_sd_m, final_sd)
  }
}

# convert many matrices to data frames; makes it easier to export as an excel file 
all_p_values <- as.data.frame(all_p_values, row.names = paste(sample_start,as.character(sample_list), sep = ""))
colnames(all_p_values) <- colnames(bg1_norm)
signif_df <- as.data.frame(signif_m, row.names = paste(sample_start,as.character(sample_list), sep = ""))
colnames(signif_df) <- colnames(bg1_norm)
nondetect_df <- as.data.frame(nondetect_m, row.names = paste(sample_start,as.character(sample_list), sep = ""))
only_detect_df <- as.data.frame(only_detect_m, row.names = paste(sample_start,as.character(sample_list), sep = ""))
give_value_df <- as.data.frame(give_value_m, row.names = paste(sample_start,as.character(sample_list), sep = ""))
final_conc_df <- as.data.frame(final_conc_m, row.names = paste(sample_start,as.character(sample_list), sep = ""))
final_important_df <- as.data.frame(final_important_m, row.names = paste(sample_start,as.character(sample_list), sep = ""))
colnames(final_important_df) <- colnames(bg1_norm)
final_sd_df <- as.data.frame(final_sd_m, row.names = paste(sample_start,as.character(sample_list), sep = ""))
drift_corrections_temp <- subset(drift_corrections, select = -c(n, id))
drift_corrections_temp <- drift_corrections_temp[,order(colnames(drift_corrections_temp))]
drift_corrections_df <- cbind(id = ids, drift_corrections_temp)
corrected_data_df <- subset(corrected_data, select = -c(n))

write.xlsx(final_important_df, file = output, sheetName = "FinalAnalysis", append = FALSE, col.names = T, row.names = T)
write.xlsx(QA_QC_final, file = output, sheetName = "QA_QC", append = TRUE, col.names = T, row.names = T)
write.xlsx(drift_corrections_df, file = output, sheetName = "DriftCorrection", append = TRUE, col.names = T, row.names = F)
write.xlsx(corrected_data_df, file = output, sheetName = "Drift_CCBCorrection", append = TRUE, col.names = T, row.names = F)
write.xlsx(all_p_values, file = output, sheetName = "PvaluesComparingToBackground", append = TRUE, col.names = T, row.names = T)
write.xlsx(final_conc_df, file = output, sheetName = "AllFieldConcentrations", append = TRUE, col.names = T, row.names = T)
write.xlsx(final_sd_df, file = output, sheetName = "AllFieldSDs", append = TRUE, col.names = T, row.names = T)

