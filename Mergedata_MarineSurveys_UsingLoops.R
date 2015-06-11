## Stuart Ball
## 09/10/2014
## Aim: to amalgamate in a single sheet all the surveys information provided by a contractor after a marine survey

# Delete any previous objects
rm(list=ls())

# Load the xlxs package
library(xlsx)

# Store the location of the folder with the 289 xls files
inputPath <- "\\\\jncc-corpfile\\JNCC Corporate Data\\Marine\\Evidence\\SurveyAndContracts\\Deepwater_QA\\ader_to_merge"

# Get a list of Excel files in the the inputPath. These files have the extension .xls or xslx
allFiles <- list.files(inputPath, pattern = "*.xls$|*.xlsx$", full.names = TRUE, ignore.case = TRUE)

# Create an empty data frame in order to acumulate the rsult
mainDataFrame <- data.frame()

for (a in 1:length(allFiles)) {
  # Read the first sheet of this file
    fileName <- read.xlsx2(allFiles[a], 1, header = FALSE, stringAsFactors = FALSE)
  
  # Loop through the columns, starting in 4, where the ocurrences and survey information is
  # Then number of columns from 4 to so on is variable
  for (b in 4:ncol(fileName)) {
  # Check there is a value in the first cell. Some sheets have extra blank columns
    if(!is.na(fileName[1, b])) {
      # Loop through the taxa to obtain their ocurrences. Taxa information always starts at row 7
      for (c in 7:nrow(fileName)) {
        
        # Check a taxon name exists since some sheets have extra blank rows
        if(!is.na(fileName[c, 1])) {
          
          # Collect the data and add it to the mainDataFrame
          mainDataFrame <- rbind(mainDataFrame, data.frame(
                                              surveyID = fileName[1, b],
                                              surveyEventID = fileName[2, b],
                                              sampleReference = fileName[3, b],
                                              replicateReference = fileName[4, b],
                                              coordinate = fileName[5, b],
                                              habitatName = fileName[6, b],
                                              taxa = fileName[c, 1],
                                              taxaQualifier = fileName[c, 2],
                                              dataType = fileName[c, 3],
                                              abundance = fileName[c, 4],
                                              stringsAsFactors = FALSE))        
                                    }
                                  }
                               }
                              }    
                             }

# Fix up the abundance column
d <- which(names(mainDataFrame) == "abundance") # Find the column number
mainDataFrame[is.na(mainDataFrame[, d]), d] <- 0 # Fix cases where abundance is NA and replace NAs with zero
mainDataFrame[, d] <- as.numeric(mainDataFrame[, d]) # Make sure they are treated as numbers

## write the results as an Excel spreadsheet file in our working directory
write.xlsx2(mainDataFrame, "merged_data.xlsx", sheetName="Merged", col.names=TRUE, row.names=FALSE)