# ---------------------------- #
# LOADS DATASET INTO SQLSERVER
# 2014-18-11 GRAHAM FRENCH
# ---------------------------- #

# LOAD REQUIRED PACKAGES
library(RODBC)

# CLEAR R DATA
rm(list=ls())

# INPUT FILES
inputFile = file.choose()
headset <-read.csv(inputFile, header=TRUE, nrows = 1, sep='\t')
classes <- sapply(headset, class)
classes[names(classes)] <- "character"
data <- read.csv(inputFile, header=TRUE, sep='\t', colClasses = classes)
print(paste("Number of rows =", nrow(data)))

# CONNECT TO SQL SERVER
SQLExpress <- "Driver=SQL Server;Server=DARIO-RODRIGUEZ\\SQLEXPRESS;Database=NBNServices;Trusted_Connection=True;"
conn <- odbcDriverConnect(SQLExpress)
sqlSave(conn, data, tablename="data", rownames=FALSE) # CHANGE TABLENAME TO filename TO PRESERVE NAME
odbcClose(conn)

# OUTPUT MESSAGE
print(paste("Number of rows imported =", nrow(data)))

# CLEAR R DATA
#rm(list=ls())