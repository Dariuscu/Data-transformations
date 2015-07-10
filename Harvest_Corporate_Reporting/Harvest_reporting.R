# Author: Dario Rodriguez
# Data started: 22/06/2015
# Date last modified: 22/06/2015
# Scope: harvest the corporarte reporting information into an Access database

# Clean up any previous data
rm(list = ls())

# Libraries
library(xlsx)
library(dplyr)
library(RODBC)

# Import the txt file with all the paths
filePaths <- "Y:/Data Services/Dario Rodriguez/Work Space/PROJECTS/R Programming/Repositories/Data-transformations/Harvest_Corporate_Reporting/Paths.txt"
paths <- read.csv(filePaths, sep = '\t', header = TRUE, stringsAsFactors = FALSE)

# Create the function getSheets
getSheets <- function (programme, sheet) {
  projectTargets <- read.xlsx2(file = sheet, sheetName = 'ProjectTargets', header = TRUE, startRow = 7, as.data.frame = TRUE, colClasses = 'character') 
  EQA <- read.xlsx2(file = sheet, sheetName = 'EQA', header = TRUE, startRow = 6, as.data.frame = TRUE, colClasses = 'character')
  PPMs <- read.xlsx2(file = sheet, sheetName = 'PPMs', header = TRUE, startRow = 16, as.data.frame = TRUE, colClasses = 'character') 
  programmeRisks <- read.xlsx2(file = sheet, sheetName = 'ProgRiskProfiles', header = TRUE, startRow = 32, as.data.frame = TRUE, colClasses = 'character') 
  
  projectTargets$Prog <- programme
  EQA$Prog <- programme
  PPMs$Prog <- programme
  programmeRisks$Prog <- programme
  
  projectTargetsAll <<- bind_rows(projectTargetsAll, projectTargets)
  EQAall <<- bind_rows(EQAall, EQA)
  PPMsAll <<- bind_rows(PPMsAll, PPMs)
  programmeRisksAll <<- bind_rows(programmeRisksAll, programmeRisks)
}


# Create an empty dataframe
projectTargetsAll <- data.frame()
EQAall <- data.frame()
PPMsAll <- data.frame()
programmeRisksAll <- data.frame()


# Call the function to harvest the information from the 14 paths
mapply(getSheets, paths$Programme, paths$Path)


# Remove blanks and nulls from the final dataframes
projectTargetsAll <- projectTargetsAll[!(is.na(projectTargetsAll$Project.number) | projectTargetsAll$Project.number == ""),]
EQAall <- EQAall[!(is.na(EQAall$Project.number.and.name) | EQAall$Project.number.and.name == ""),]
PPMsAll <- PPMsAll[!(is.na(PPMsAll$PPM.number) | PPMsAll$PPM.number == ""),]
programmeRisksAll <- programmeRisksAll[!(is.na(programmeRisksAll$Programme.risk) | programmeRisksAll$Programme.risk == ""),]
programmeRisksAll <- programmeRisksAll[, -c(12:46)]


# Connect to the Access database
#channel <- odbcConnectAccess2007("Y:/Data Services/Dario Rodriguez/Work Space/PROJECTS/R Programming/Repositories/Data-transformations/Harvest_Corporate_Reporting/CorparteReporting_2015_2016.accdb")


# Save the different tables
#sqlSave(channel = channel, dat = projectTargetsAll, tablename = 'Targets')

server <- "Driver=SQL Server;Server=DARIO-RODRIGUEZ\\SQLEXPRESS;Database=Test;Trusted_Connection=True;"
channel <- odbcDriverConnect(server)
sqlSave(channel = channel, dat = projectTargetsAll, tablename = 'Targets', append = TRUE, rownames = FALSE, colnames = FALSE)
odbcClose(conn)