# Author: Dario Rodriguez
# Date started: 04/02/2015
# Date finalised: xxxx
# Scope: to replicate the same VBA code from the Marine Resource planner
# Constraints:
# --------------------------------------------------------------------------------------------------
# Load libraries to work with Excel
library(xlsx)

# Clear any previous data
rm(list=ls())

# Set the working directory to the folder where I have the Marine Resource Planner
setwd("Y:/Data Services/Dario Rodriguez/Work Space/PROJECTS/R Programming/Scripts/MarineResourcePlanner")
marineResourcePlanner <- "Resource Planner Master 2015-2016.xlsx"

# Create the empty data frame in which to accumulate the result
dataTab <- data.frame()

# Marine Programmes
marineProgrammes <- c("MPA", "Evidence", "Fisheries and Species", "MEAA", "Monitoring", "OIA")


for (i in 1:length(marineProgrammes)) {
  # Loop through each tab
  a <- read.xlsx(marineResourcePlanner, sheetName = marineProgrammes[i], startRow = 3, header=TRUE)
  
 
  # Print the name of the programme so we know where we are
  print(paste((marineProgrammes[i]), "programme"))

  # Go through the columns with the staff names
  for (j in 8:(ncol(a))) {
    # Skip columns with no staff names
    if(!substr(names(a)[j],1,2)=="NA") {
      # For each member of the staff, go through each cell of the column
      for (g in 1:nrow(a)) {
        # Skip cells with no data
        if (!is.na(a[g,j])&a[g,j]>0) {
          dataTab <-rbind(dataTab, data.frame(
                                  ProgrammeName = marineProgrammes[i],
                                  ProjectName = a[g,1],
                                  ProjectNumber = a[g,2],
                                  ProjectTarget = a[g,3],
                                  Task = a[g,4],
                                  Contingency = a[g,5],
                                  StaffName = names(a)[j],
                                  StaffTime = a[g,j],
                                  stringsAsFactors=FALSE))
        }

      }

    }

  }
}


# When a project number doesn't have a number, it will have NA. We need to
#change these NAs to 0 so that the column can be converted into numeric
colPositionProjectNumber <- which(names(dataTab)=="ProjectNumber")
dataTab[is.na(dataTab[,colPositionProjectNumber]),colPositionProjectNumber] <- 0
dataTab[, colPositionProjectNumber] <- as.numeric(dataTab[, colPositionProjectNumber])

# Column of staff time will always have something in it, but it must be converted into numeric
colPositionStaffTime <- which(names(dataTab)=="StaffTime")
dataTab[, colPositionStaffTime] <- as.numeric(dataTab[, colPositionStaffTime])

# Make sure "Data" tab does not already exist
wb <- loadWorkbook(marineResourcePlanner)
sheets <- getSheets(wb)
if("Data" %in% names(sheets)){
  removeSheet(wb, sheetName="Data")
}
saveWorkbook(wb, marineResourcePlanner)

# Make sure column names are as expected
names(dataTab) <- c("PROGRAMME","PROJECT NAME","PROJECT NUMBER","Project Target","TASK","CONTINGENCY","START","END","STAFF","TIME")

# Write the data frame DataTab in the Data tab from the Marine Resource Planner
write.xlsx(dataTab, file=marineResourcePlanner, sheetName="Data", col.names = TRUE,
           row.names=FALSE, append=TRUE)