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

StoreList <- read_csv(paste0(main_path,"WIP/DatariteStoreExtract.csv"), col_names = TRUE, col_types = "cccccc") %>%
            select(LocationID, Brand, Division, Region, DescriptionLong, DescriptionShort)

### Read in List of Divisions to be excluded
Div_Excl_List <- read_csv("Input/Div Excl List.csv", col_names = TRUE, col_types = "c")
#read_csv(paste0(main_path,"Div Excl List.csv"), col_names = TRUE, col_types = "c")

### Read in List of Incident Categories to be excluded
Incident_Excl_List <- read_csv("Input/Incident Excl List.csv", col_names = TRUE, col_types = "c")

### Read in list of service categories that we want to include
Service_Cats <- read_csv("Input/Service Cats.csv", col_names = TRUE, col_types = "c")

### Read in list of Non-Careline agents
Agents <- read_csv("Input/AgentInclList.csv", col_names = TRUE, col_types = "c")

### Read in list extract from C4C (SAP service cloud)
SapRawData <- read_csv("Input/SAPAnalyticsReport.csv", col_names = TRUE)

SapData <- SapRawData %>% 
  anti_join(Div_Excl_List, by = "Division")  %>% 
  inner_join(Service_Cats, by = "Service Category") %>% 
  anti_join(Incident_Excl_List, by = "Incident Category") %>%
  inner_join(Agents, by = "Agent") %>%
  rename('Brand (Store)' = Brand,
         'Division (Store)' = Division,
         Store = `Store Name`,
         'Category (Source Item)' = `Service Category`,
         'Type (Source Item)' = `Incident Category`,
         'Source Item' = Object,
         'Region (Store)' = Region,
         'Case Number' = `Ticket ID`
        ) %>%
  select(LocationID, `Brand (Store)`, `Division (Store)`,`Region (Store)`, Store , `Case Number`,`Created On`, `Ticket Type`) 

Comp <- SapData %>% filter(LocationID != "#") %>%
 left_join(StoreList, by = "LocationID")

write.csv(Comp, paste0(main_path,"WIP/StoreComp.csv"))

MissingBrand <- Comp %>% 
    filter(is.na(Brand)) %>% 
    distinct()
    
write.csv(MissingBrand, paste0(main_path,"WIP/MissingBrand.csv"))

MissingDivision <- Comp %>% 
  filter(is.na(Division)) %>% 
  distinct()

write.csv(MissingDivision, paste0(main_path,"WIP/MissingDivision.csv"))

MissingRegion <- Comp %>% 
  filter(is.na(Region)) %>% 
  distinct()

write.csv(MissingRegion, paste0(main_path,"WIP/MissingRegion.csv"))

MissingStore <- Comp %>% 
  filter(is.na(Region)) %>% 
  distinct()

write.csv(MissingStore, paste0(main_path,"WIP/MissingStore.csv"))

UnqStores <- Comp %>% 
  select(LocationID, `Brand (Store)`, `Division (Store)`, `Region (Store)`, Store) %>%
  distinct() 
  
DuplicateStores <- UnqStores %>%
  select(LocationID) %>%
  group_by(LocationID) %>%
  summarise(StoreCount = n()) %>%
  filter(StoreCount > 1) %>%
  inner_join(StoreList, by = "LocationID")

write.csv(DuplicateStores, paste0(main_path,"WIP/DuplicateStores.csv"))


  
        

