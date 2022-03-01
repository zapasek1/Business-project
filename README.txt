title: "Business project"
authors: "Jakub Zapasnik (38401), Daniel Lilla (38963), Michał Kloska (46341)"

- The structure of the project is done in such manner:
1. Files are listed and then uploaded from subfolder "Data".
2. All computation takes place inside main "R_proj_script" file. 
3. Output which in that case is one xlsx file is then printed inside "Output" subfolder. 

- Additional libraries used inside main R file:
1. library(tidyverse)
2. library(readxl)
3. library(xlsx)
4. library(lubridate)
5. library(pivottabler)
6 .library(openxlsx)

- Key features of the project:
1. List to get all the xslx files inside "Data" folder.
2. One function that reads all xlsx files inside list, transform them into tidy data and then create final list with multiple data frames that represends each country data. 
3. Second function that use this list to pivot it in desired manner and print it as excel workbook with multiple worksheets.
  
