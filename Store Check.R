#start <- Sys.time()

library(stringr)
library(readxl)
library(tidyverse)

#function to get username8
user <- function(wd){
  user_name <- wd %>%
    str_remove("C:/Users/") 
  
  substring(user_name,1, regexpr("/", user_name)-1)
}

##### Set Main paths
username<-  user(getwd())
main_path <- paste0("C:/Users/",username,"/OneDrive - Shoprite Checkers (Pty) Limited/Reporting Files and Documents/Tableau Dashboard Automation/Customer Care Weekly Dashboard/")

StoreList <- read_csv(paste0(main_path,"WIP/DatariteStoreExtract.csv"), col_names = TRUE, col_types = "iccccc") %>%
            select(LocationID, Brand, Division, Region, DescriptionLong, DescriptionShort)

DashboardData <- as_tibble(read_xlsx(paste0(main_path,"csc weekly report.xlsx"), sheet = "Case Advanced Find View")) %>%
                select(LocationID, `Brand (Store)`, `Division (Store)`, `Region (Store)`, Store) 

#Comp <- DashboardData %>% 
#  full_join(StoreList, by = "LocationID") 
#write.csv(Comp, paste0(main_path,"WIP/StoreComp.csv"))

MissingStores <- Comp %>% 
    filter(is.na(Brand))

write.csv(MissingStores, paste0(main_path,"WIP/MissingStores.csv"))

UnqStores <- Comp %>% 
  filter(!is.na(Brand)) %>%
  select(LocationID, `Brand (Store)`, `Division (Store)`, `Region (Store)`) %>%
  distinct() 
  
DuplicateStores <- UnqStores %>%
  select(LocationID) %>%
  group_by(LocationID) %>%
  summarise(StoreCount = n()) %>%
  filter(StoreCount > 1) %>%
  inner_join(StoreList, by = "LocationID")

write.csv(DuplicateStores, paste0(main_path,"WIP/DuplicateStores.csv"))


  
        

