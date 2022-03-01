'
---
  title: "Business project"
subtitle: "Group project for business project class"
authors: "Jakub Zapasnik (38401), Daniel Lilla (38963), Michał Kloska (46341)"
---
'
Sys.setlocale("LC_CTYPE", "Polish")
Sys.setenv("LANGUAGE"="PL")
library(tidyverse)
library(readxl)
library(xlsx)
library(lubridate)
library(pivottabler)
library(openxlsx)

# Data import and transformation to tidy data
############################################################################

files <- list.files("Data", pattern="*.xlsx", full.names = TRUE, recursive = TRUE)
data_xlsx <- list()

all_sheets <- function(files){
  list_countries <- list()
  for (j in 1:length(files)){
    test_temp <- files[j]
    sheets <- readxl::excel_sheets(test_temp)
    tybble <- lapply(sheets, function(x) read_excel(test_temp, sheet = x))
    list_df <- list() # deleting everything from the list while importing new sheet
    for (i in 1:length(tybble)){
      df <- tybble[i]
      df <- as.data.frame(df)
      df <- df[ , colSums(is.na(df)) < nrow(df)] # delet columns with NA
      df <- df %>% add_column(Yield = sheets[i])
      if (ncol(df) > 3 & ncol(df) < 6){ # 4 columns means two dates
        df1 <- df %>% select(1,2,Yield) %>% drop_na() # extracting first date and yield
        df2 <- df %>% select(3,4,Yield) %>% drop_na() # extracting second date and yield
        names(df1)[1] <- 'Date' # renaming to merge files together
        names(df2)[1] <- 'Date'
        df <- merge(df1,df2, all.x = TRUE)
        list_df[[i]] <- df
        # print('Two dates,', i) # optional
      } else if (ncol(df) == 6) {
        df1 <- df %>% select(1,2,Yield) %>% drop_na() # extracting first date and yield
        df2 <- df %>% select(4,5,Yield) %>% drop_na() # extracting second date and yield
        names(df1)[1] <- 'Date' # renaming to merge files together
        names(df2)[1] <- 'Date'
        df <- merge(df1,df2, all.x = TRUE)
        list_df[[i]] <- df
        print('Poland 10Y')
      } else if (ncol(df) > 6 & ncol(df) < 8) {
        df1 <- df %>% select(1,2,Yield) %>% drop_na() # extracting first date and yield
        df2 <- df %>% select(3,4,Yield) %>% drop_na() # extracting second date and yield
        df3 <- df %>% select(5,6,Yield) %>% drop_na() # extracting third date and yield
        names(df1)[1] <- 'Date' # renaming to merge files together
        names(df2)[1] <- 'Date'
        names(df3)[1] <- 'Date'
        dfs <- merge(df1,df2, all.x = TRUE)
        df <- merge(dfs,df3, all.x = TRUE)
        list_df[[i]] <- df
        # print('Three dates')
      } else if (ncol(df) > 8) {
        list_df[[i]] <- df
        print('More than 3 dates')
      } else {
        list_df[[i]] <- df
        # print('One date')
      }
    }
    file_data <- dplyr::bind_rows(list_df) # binding all sheets into one
    list_countries[[j]] <- file_data
    print(j) # optional
  }
  for (i in 1:length(list_countries)){ #extract from date month and year and create new columns
    df <- list_countries[i]
    df <- as.data.frame(df)
    names(df)[1] <- 'Date'
    mutate(df, Date=as.Date(Date, format = "%d.%m.%Y"))
    df <- df %>% dplyr::mutate(year = lubridate::year(Date),
                               month = lubridate::month(Date))
    df <- df %>% select(!c(1))
    df <- df %>% select(month, everything())
    df <- df %>% select(year, everything())
    df <- df %>% select(Yield, everything())
    list_countries[[i]] <- df
    print(i)
  }
  for (i in 1:length(list_countries)){ # pivot data ino tidy way
    df <- list_countries[i]
    df <- as.data.frame(df)
    df <- pivot_longer(df, cols = (!c(year, month, Yield)), 
                               names_to = "Bond", values_drop_na = TRUE ) 
    list_countries[[i]] <- df
    print(i)
  } 
  for (i in 1:length(list_countries)){
    df <- list_countries[i]
    df <- as.data.frame(df)
    df1 <- df %>% select("Yield", "value") %>% filter(str_detect(Yield, "Y|y"))
    df1$Yield <- parse_number(df1$Yield)
    df1$Yield <- df1$Yield * 12
    df2 <- df %>% select("Yield", "value") %>% filter(str_detect(Yield, "M|m"))
    df2$Yield <- parse_number(df2$Yield)
    dfs <- union_all(df2,df1, all.x = TRUE)
    df <- df %>% mutate(month_yield = dfs$Yield)
    colnames(df) <- c('Yield to maturity in Years','Date','Month', 'Bond name', 'Value', 'Yield to maturity in months')
    df <- df[,c(4,2,3,1,6,5)]
    list_countries[[i]] <- df
    print(i)
    
  }
  data_xlsx <- set_names(list_countries, files) # create list of all files
  # organize names in readable manner
  names(data_xlsx) = gsub(pattern = ".xlsx.*", replacement = "", x = names(data_xlsx))
  names(data_xlsx) = gsub(pattern = "Data/", replacement = "", x = names(data_xlsx))
  return(data_xlsx)
}

#list of countries 
all_together <- all_sheets(files)

# Pivot table and print it to excel workbook
##############################################################################

excel_test1 <- function(all_together){
  wb <- createWorkbook()
  for (i in 1:length(all_together)){
    test <- all_together[[i]]
    test <- as.data.frame(test)
    clean <- PivotTable$new(evaluationMode = "batch", processingLibrary = "data.table", argumentCheckMode = "none")
    clean$addData(test)
    clean$addColumnDataGroups("Yield to maturity in months", addTotal = FALSE)
    clean$addColumnDataGroups("Bond name", addTotal=FALSE)
    clean$addColumnDataGroups("Yield to maturity in Years", addTotal=FALSE)
    clean$addColumnDataGroups("Yield to maturity in months", addTotal = FALSE)
    
    clean$addRowDataGroups("Date", addTotal=FALSE)
    clean$addRowDataGroups("Month", addTotal=FALSE)
    clean$defineCalculation(calculationName="Value", summariseExpression="Value")
    clean$evaluatePivot()
    
    addWorksheet(wb, sheetName = names(all_together[i]))
    clean$writeToExcelWorksheet(wb=wb, wsName=names(all_together[i]),
                                topRowNumber=1, leftMostColumnNumber=1, applyStyles=FALSE)
    
    print(i)
  }
  saveWorkbook(wb, file="Output/business_project.xlsx", overwrite = TRUE, returnValue = TRUE)
}
excel_test1(all_together)

# SPECIAL FUNCTIONS, NOT NECCESARY!!!
############################################################################

#merged big frame
data_xlsx_df <- map2_df(all_together, files, ~update_list(.x, file = .y)) 

# this extract tibbles from lists to seprate objects
lapply(names(data_xlsx), function(x) assign(x, data_xlsx[[x]], envir = .GlobalEnv))
