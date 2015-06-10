# Author: Dario Rodriguez
# Date started: 21/11/2014
# Date finalised: 24/11/2014
# Scope: to import data from a txt delimited file to an Access database
# Constraints: either the txt and Access files don't need a specific name because they'll be chosen.
# --------------------------------------------------------------------------------------------------

# Load the library need for the connexion
library(RODBC)

# Clear any previous data
rm(list=ls())

# Choose the data to import file
data_to_import <- read.csv(file.choose(), sep = '\t', as.is = TRUE, colClasses = "character")

# Quality checks (they aren't necessary)
sapply(data_to_import, class) # Check that everything has been imported as string
paste("No of row in original file:", nrow(data_to_import)) # See the number of rows imported into R

# Connect to the Access database by choosing it
conn.access.database <- odbcConnectAccess2007(file.choose())

# Save the data_to_import dataframe into the Access database
sqlSave(conn.access.database, data_to_import, tablename="dataset", rownames=FALSE)

# Query the number or rows of the table imported in Access
sql.count <- "SELECT COUNT(*) FROM dataset"
count_rows_access <- sqlQuery(conn.access.database, sql.count)
if ((count_rows_access == nrow(data_to_import)) == TRUE) {
  paste("Everything has been imported")
}

# Close connection
odbcClose(conn.access.database)

# Clear everthing after the process has been run
rm(list=ls())