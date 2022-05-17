#start <- Sys.time()

library(stringr)
library(xlsx)
library(tidyverse)

#get current date from system
dt <- Sys.Date()

#function to clean up comments
#Improve function 
CleanText <- function(htmlString) {
  clean1 <- str_replace_all(htmlString,"&nbsp;", " ")
  clean2 <- gsub("<.*?>", "", clean1) 
}

#function to get username8
user <- function(wd){
  user_name <- wd %>%
    str_remove("C:/Users/") 
  
  substring(user_name,1, regexpr("/", user_name)-1)
}

##### Set Main paths
username <-  user(getwd())
#main_path <- paste0("C:/Users/",username,"/OneDrive - Shoprite Checkers (Pty) Limited/GIT Repo/Customer Care Complaints Weekly Dashboard/")
report_path <- paste0("C:/Users/",username,"/OneDrive - Shoprite Checkers (Pty) Limited/Reporting Files and Documents/Tableau Dashboard Automation/Customer Care Weekly Dashboard/")

### Read in List of Divisions to be included
Div_Incl_List <- read_csv("Input/Div Incl List.csv", col_names = TRUE, col_types = "c")
  #read_csv(paste0(main_path,"Div Excl List.csv"), col_names = TRUE, col_types = "c")

### Read in List of Incident Categories to be excluded
Incident_Excl_List <- read_csv("Input/Incident Excl List.csv", col_names = TRUE, col_types = "c")

### Read in list of Service Categories that we want to include
Service_Cats <- read_csv("Input/Service Cats.csv", col_names = TRUE, col_types = "c")

### Read in list of Careline agents
Agents <- read_csv("Input/AgentInclList.csv", col_names = TRUE, col_types = "c")

### Read in list extract from C4C (SAP service cloud)
SapRawData <- read_csv("Input/SAPAnalyticsReport.csv", col_names = TRUE)

### This bit of code filters the data extract down to what is applicable for Careline
CCReportingData <- SapRawData %>% 
  filter(`Ticket Type` == "Careline") %>% 
  inner_join(Div_Incl_List, by = "Division")  %>% 
  inner_join(Service_Cats, by = "Service Category") %>% 
  anti_join(Incident_Excl_List, by = "Incident Category") %>%
  inner_join(Agents, by = "Agent")

### This bit of code filters the data extract down to what is applicable for Ecommerce
LSReportingData <- SapRawData %>%
  filter(`Ticket Type` == "E-commerce") %>%
  inner_join(Service_Cats, by = "Service Category") %>% 
  anti_join(Incident_Excl_List, by = "Incident Category") %>%
  inner_join(Agents, by = "Agent")

#CCReportingData %>% write_csv("CCReportingData.csv")

### Rename
LSReportingData$Brand <- "LiquorShop Online"
LSReportingData$Division <- "LiquorShop Online"
LSReportingData$Region <- "LiquorShop Online"
LSReportingData$`Store Name` <- "LiquorShop Online"

### Combine the Reporting Data for Careline and Liquor Shop
ReportingData <- rbind(CCReportingData, LSReportingData)

csc_weekly_report <- ReportingData %>%
  mutate(CleanCaseDesc = CleanText(ReportingData$'Case Description')) %>%
  select(-'Case Description') %>%
  mutate(`Calendar Day Date` = `Created On`) %>%
  rename('Brand (Store)' = Brand,
         'Division (Store)' = Division,
         Store = `Store Name`,
         'Category (Source Item)' = `Service Category`,
         'Type (Source Item)' = `Incident Category`,
         'Source Item' = Object,
         'Case Origin' = Origin,
         'Region (Store)' = Region,
         'Case Description' = CleanCaseDesc,
         'Case Number' = `Ticket ID`,
         Owner = Agent,
         'First Name' = `First Name (Account)`,
         'Last Name' = `Last Name (Account)`,
         'Cell Phone' = `Mobile (Account)`
  ) %>%
  select(`Brand (Store)`, `Division (Store)`,`Region (Store)`, Store , `Category (Source Item)`, `Type (Source Item)`, `Source Item`,
         `Case Origin`, `Case Title`, `Case Description`, `Case Number`, 
         `Product Brand`, `Product Description`, `Created On`, `Status`, `Owner`, `First Name`, `Last Name`, 
         `Cell Phone`, `Calendar Day Date`, #`Incident Description`,
         `Ticket Type`, LocationID) 
  
#This part of the code is to specifically rename Head Office to Home office

#Filter to only the Home Office Store using the store, because of the DataRite issue
temp <- csc_weekly_report %>%
  filter(Store == "Head Office") 

#Rename 
temp$`Brand (Store)` <- str_replace_all(temp$`Brand (Store)`, "#", "Home Office")
temp$`Division (Store)` <- str_replace_all(temp$`Division (Store)`, "Western Cape Division", "Western Cape Division (Home Office)")
temp$Store <- str_replace_all(temp$Store, "Head Office", "Home Office")
  
#Filter original data to exclude  the Home Office Store
ReportingData <- csc_weekly_report %>% 
  filter(Store != "Head Office") 

#Rowbind the 2 data sets
csc_weekly_report <- rbind(ReportingData,temp) %>% 
    filter(`Division (Store)` != "Western Cape Division" ) # this piece of code was added into handle stores that were not renamed, because of the DataRite issue

csc_weekly_report$`Product Brand` <- sub("#", "" , csc_weekly_report$`Product Brand`)
csc_weekly_report$`Product Description` <- sub("#", "" , csc_weekly_report$`Product Description`)
csc_weekly_report$`First Name` <- sub("#", "" , csc_weekly_report$`First Name`)
csc_weekly_report$`Last Name` <- sub("#", "" , csc_weekly_report$`Last Name`)
csc_weekly_report$`Cell Phone` <- sub("#", "" , csc_weekly_report$`Cell Phone`)

write_csv(csc_weekly_report, paste0(report_path,"WIP/csc_weekly_report ",dt,".csv")) 

#write_csv(csc_weekly_report, "csc_weekly_report.csv")

#write.xlsx(as.data.frame(csc_weekly_report), paste0(report_path,"Archive CC Dashboard/","csc_weekly_report_",dt,".xlsx"), row.names = FALSE)

#OUtput for KPI Dashboard
#KPI <- ReportingData %>%
#  select(Brand, Division, `Service Category`, `Service Category`, `Incident Category`, 
#        Object, Region, `Product Brand`, `Ticket ID`, `Created On`, Status, LocationID)

#write_csv(KPIdb, paste0(report_path,"KPI DB csc weekly report.csv"))

#end <-  Sys.time()

#duration <- end - start

#print(duration)

