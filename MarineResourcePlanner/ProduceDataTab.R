# Author: Dario Rodriguez
# Date started: 04/02/2015
# Date finalised: xxxx
# Scope: to replicate the same VBA code from the Marine Resource planner
# Constraints:
# --------------------------------------------------------------------------------------------------
# Load libraries to work with Excel
library(xlsx)
library(reshape2)
library(data.table)

# Clear any previous data
rm(list=ls())

# Set the working directory to the folder where I have the Marine Resource Planner
setwd("Y:/Data Services/Dario Rodriguez/Work Space/PROJECTS/R Programming/Scripts/MarineResourcePlanner")
marineResourcePlanner <- "Resource Planner Master 2015-16.xlsx"

# Create the empty data frame in which to accumulate the result
dataTab <- data.frame()

# Marine Programmes
marineProgrammes <- c("MPA", "Evidence", "Fisheries and Species", "MEAA", "Monitoring", "OIA")

# Loop through each marine tab
for (i in 1:length(marineProgrammes)) {
  print(paste((marineProgrammes[i]), "programme"))

  a <- read.xlsx(marineResourcePlanner, sheetName = marineProgrammes[i], startRow = 3, header = TRUE)

  # Remove columns of Star and End
  a <- a[, -c(6:7)]
  
  # Replace blank projects with 0
  colPositionProjectNumber <- which(names(a) == "PROJECT.NUMBER")
  a[is.na(a[, colPositionProjectNumber]), colPositionProjectNumber] <- 0
  
  # Remove lines withouth a programme name
  a <- a[which(!is.na(a[, 1])), ]
  
  # Normalise the dataframe so all the columns with staff names become an only column
  b <- melt(a, id.vars=names(a)[1:5], na.rm = TRUE, value.name = "TIME", variable.name = "STAFF", factorsAsStrings = FALSE, rownames = FALSE)
  
  # Add an extra column with the programme's name
  b <- cbind(ProgrammeName = rep(marineProgrammes[i], nrow(b)), b)

  # Append the dataframe to the previous one
  dataTab <- rbind(dataTab, b)
}

# Convert project numbers into numberic
colPositionProjectNumber <- which(names(dataTab) == "PROJECT.NUMBER")
dataTab[, colPositionProjectNumber] <- as.numeric(dataTab[, colPositionProjectNumber])

# Column of staff time will always have something in it, but it must be converted into numeric
colPositionStaffTime <- which(names(dataTab )== "TIME")
dataTab[, colPositionStaffTime] <- as.numeric(dataTab[, colPositionStaffTime])

# Make sure "Data" tab does not already exist
wb <- loadWorkbook(marineResourcePlanner)
sheets <- getSheets(wb)
if("Data" %in% names(sheets)){
    removeSheet(wb, sheetName="Data")
}
saveWorkbook(wb, marineResourcePlanner)

# Rename colum names
names(dataTab) <- c("PROGRAMME","PROJECT NAME","PROJECT NUMBER","PROJECT TARGET","TASK","CONTINGENCY","STAFF","TIME")

# Write the data frame dataTab in the Resource Planner
write.xlsx(dataTab, file = marineResourcePlanner, sheetName = "Data", col.names = TRUE, row.names = FALSE, append = TRUE)
