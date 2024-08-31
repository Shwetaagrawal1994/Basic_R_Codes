# Code for writing excel file (.xlsx) through R code

###########################
# INSTALLING LIBRARIES
###########################

packages <- c("openxlsx")
installed_packages <- packages %in% rownames(installed.packages() )
if ( any(installed_packages == FALSE) ) {
   installed.packages( packages[!installed_packages] )
}

# loading packages
invisible(lapply(packages, library, character.only = TRUE))

###########################
# DEFINE PATH
###########################
setwd( dirname(rstudioapni::getSourceEditorContext()$path) )
setwd('..')
input_Path = paste0( getwd(), "/Input/")

Data = data.frame( Name = c('shweta', 'Vanishika', 'Shreya', 'Shristhi' ), Number = c(100, 200, 300, 400) )

wb2 <- loadWorkbook( paste0(input_Path, "Trial.xlsx") )
writeData(wb2, "TrialTab", Data, startCol = 1, startRow = 1, rownames = F)

protectWorksheet(wb2,
                protect = TRUE,
                password = "ABC",
                loackFormattingCells = TRUE,
                loackFormattingColumns = TRUE)

saveWorkbook(wb2, paste0(input_Path, "Trial.xlsx"), overwrite = TRUE )