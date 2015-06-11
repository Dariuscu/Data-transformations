# Author: Dario Rodriguez
# Date started: 31/03/2015
# Date finalised: xxxx
# Scope: practive pivot tables in R
# --------------------------------------------------------------------------------------------------

rm(list=ls())

library(xlsx)
library(dplyr)
library(tidyr)
library(stringr)


setwd("Y:/Data Services/Dario Rodriguez/Work Space/PROJECTS/R Programming/Repositories/Data-transformations/MarineResourcePlanner/")


marineResourcePlanner <- "Resource Planner Master 2015-16.xlsx"


dataMaster <- read.xlsx(marineResourcePlanner, sheetName = 'Data', header = TRUE, stringsAsFactors = FALSE)
dataMaster$STAFF <- gsub("[.]", " ", dataMaster$STAFF)
dataMaster$STAFF <- gsub("  ", " ", dataMaster$STAFF)
dataMaster$STAFF <- str_trim(dataMaster$STAFF)
dataMaster$PROGRAMME <- gsub("Fisheries and Species", "Fisheries", dataMaster$PROGRAMME)

# INDIVIDUAL ALLOCATION
dataMaster2 <- dataMaster[, c(1, 6, 7, 8)] %>%
              group_by(PROGRAMME, CONTINGENCY, STAFF)  %>%
              summarise(totalTime = sum(TIME)) %>%
              spread(PROGRAMME, totalTime, fill = 0) %>%
              arrange(STAFF) %>%
              mutate(grandTotat = Evidence + Fisheries + MEAA + Monitoring + MPA + OIA)
              
dataMaster2 <- dataMaster2[, c(2, 1, 3, 4, 5, 6, 7, 8, 9)]

# PROGRAMME ALLOCATION
dataMaster3 <- dataMaster[, c(1, 2, 5, 6, 7, 8)] %>%
              group_by(PROGRAMME, PROJECT.NAME , TASK, CONTINGENCY, STAFF) %>%
              summarise(totalTime = sum(TIME))