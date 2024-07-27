library(openxlsx)

# Function to add a new sheet to an existing workbook
addSheetToWorkbook <- function(workbookPath, sheetName, data = NULL, rowNames = FALSE, colNames = TRUE) {
  # Check if the workbook exists
  if (!file.exists(workbookPath)) {
    stop("The specified workbook does not exist.")
  }
  
  # Load the workbook
  wb <- loadWorkbook(workbookPath)
  
  # Add a new sheet
  addWorksheet(wb, sheetName)
  
  # If data is provided, write it to the new sheet
  if (!is.null(data)) {
    writeData(wb, sheet = sheetName, x = data, rowNames = rowNames, colNames = colNames)
  }
  
  # Save the workbook
  saveWorkbook(wb, workbookPath, overwrite = TRUE)
  
  cat("Sheet added successfully to the workbook.\n")
}

