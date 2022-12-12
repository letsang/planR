
# Data Wrangling
library(readxl)
library(dplyr)
library(tidyr)
library(lubridate)
library(aweek)
library(stringr)
library(rebus)
# Office365
library(Microsoft365R)

calendar <- data.frame("Week" = c("Wk. 1 January 2022","Wk. 2 January 2022","Wk. 3 January 2022","Wk. 4 January 2022","Wk. 5 January 2022","Wk. 1 February 2022","Wk. 2 February 2022","Wk. 3 February 2022","Wk. 4 February 2022","Wk. 1 March 2022","Wk. 2 March 2022","Wk. 3 March 2022","Wk. 4 March 2022","Wk. 1 April 2022","Wk. 2 April 2022","Wk. 3 April 2022","Wk. 4 April 2022","Wk. 5 April 2022","Wk. 1 May 2022","Wk. 2 May 2022","Wk. 3 May 2022","Wk. 4 May 2022","Wk. 1 June 2022","Wk. 2 June 2022","Wk. 3 June 2022","Wk. 4 June 2022","Wk. 1 July 2022","Wk. 2 July 2022","Wk. 3 July 2022","Wk. 4 July 2022","Wk. 5 July 2022","Wk. 1 August 2022","Wk. 2 August 2022","Wk. 3 August 2022","Wk. 4 August 2022","Wk. 1 September 2022","Wk. 2 September 2022","Wk. 3 September 2022","Wk. 4 September 2022","Wk. 1 October 2022","Wk. 2 October 2022","Wk. 3 October 2022","Wk. 4 October 2022","Wk. 5 October 2022","Wk. 1 November 2022","Wk. 2 November 2022","Wk. 3 November 2022","Wk. 4 November 2022","Wk. 1 December 2022","Wk. 2 December 2022","Wk. 3 December 2022","Wk. 4 December 2022","Wk. 5 December 2022"),
                       "Week#" = 1:53)

############################################################################################################################
############################################################################################################################
############################################################################################################################
############################################################################################################################

sales2 <- read_excel("db\\STOCK PRO - SALES.xlsx", skip = 6) %>%
  mutate(SMC = paste0(`Style Color MODEL_CODE`,`Style Color MATERIAL_CODE`, `Style Color COLOR_CODE`)) %>%
  select(SMC,
         Department = `Department DESC`,
         Line = `KPI 1`,
         Season = `KPI 18`,
         Year,
         Week = `Week of Year Number`,
         Sales = `Sales FP (Qty)`) %>%
  group_by(SMC, Department, Line, Season, Year, Week) %>%
  summarise(Sales = sum(Sales))
sales2[is.na(sales2)] <- 0

inv2 <- read_excel("db\\STOCK PRO - INV.xlsx", skip = 5) %>%
  merge(calendar, by="Week") %>%
  select(SMC = `Style Color CODE`,
         Department = `Department DESC`,
         Line = `Kpi 1`,
         Season = `Kpi 18`,
         Year,
         Week = Week.,
         OH) %>%
  as_tibble() %>%
  group_by(SMC, Department, Line, Season, Year, Week) %>%
  summarise(OH = sum(OH))
inv2[is.na(inv2)] <- 0

pfp2 <- read_excel("db\\STOCK PRO - PFP.xlsx", skip = 4) %>%
  mutate(SMC = paste0(`Style Color MODEL_CODE`,`Style Color MATERIAL_CODE`, `Style Color COLOR_CODE`)) %>%
  mutate(Year = year(`Start Delivery Window`), Week = week(`Start Delivery Window`)) %>%
  select(SMC,
         Department = `Department DESC`,
         Line = `KPI 01`,
         Season = `KPI 18`,
         Year,
         Week,
         PFP = `Pending From Production  (Qty)`) %>%
  filter(SMC != "NANANA") %>%
  group_by(SMC, Department, Line, Season, Year, Week) %>%
  summarise(PFP = sum(PFP))
pfp2[is.na(pfp2)] <- 0

############################################################################################################################
############################################################################################################################
############################################################################################################################
############################################################################################################################

### MAIN DATASET ###
collection <- unique(rbind(sales2[,1:4], inv2[,1:4], pfp2[,1:4]))
w22 <- tibble('Week' = 1:52) %>% mutate('Year' = 2022)
w23 <- tibble('Week' = 1:52) %>% mutate('Year' = 2023)
main <- merge(rbind(w22, w23), collection)

base <- full_join(main, sales2, by = c("Week", "Year", "SMC", "Department", "Line", "Season")) %>%
  full_join(inv2, by = c("Week", "Year", "SMC", "Department", "Line", "Season")) %>%
  full_join(pfp2, by = c("Week", "Year", "SMC", "Department", "Line", "Season")) %>%
  mutate(Time = paste(Week,Year,sep = "-"))
base[is.na(base)] <- 0

saveRDS(base, "base.rds")