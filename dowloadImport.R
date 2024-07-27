
source("Credentials.R")

# Get the list of datasets
datasets <- kobotools_api(url=url1, simplified=TRUE, uname=username, pwd=password)

# Filter the dataset for the desired name and extract the assetId
assetId <- datasets %>%
  filter(name == desired_name) %>%
  pull(asset)

# Download the data
data <- kobo_xls_dl(url = url1, 
                  uname = username,
                  pwd = password,
                  assetid = assetId,
                  lang = "_default",
                  all = "false",
                  hierarchy = "false",
                  include_grp = "true",
                  grp_sep = "/",
                  multi_sel = "both",
                  media_url = "true",
                  fields = NULL,
                  sub_ids = NULL,
                  sleep = 2
) 

# Extract each table from data as an external data frame object
# The first table is named 'main', the rest retain their original names
table_names <- names(data)
for (i in seq_along(data)) {
  table_name <- if (i == 1) "main" else table_names[i]
  assign(table_name, data[[i]])
}

# Save objects to an Excel file
save_all_objects_to_excel <- function(file_path, tablenames) {
  # Create a new workbook
  wb <- createWorkbook()
  
  # Loop through each object and add it to a new sheet
  for (obj_name in tablenames) {
    if (exists(obj_name)) {
      obj <- get(obj_name)
      addWorksheet(wb, sheetName = obj_name)
      writeData(wb, sheet = obj_name, x = obj)
    } else {
      warning(paste("Object", obj_name, "does not exist in the global environment."))
    }
  }
  
  # Save the workbook to the specified file path
  saveWorkbook(wb, file_path, overwrite = TRUE)
  cat("Raw data has been successfully added to the Rawdata folder")
}

# Define the table names
tabnames <- table_names
tabnames[1] <- "main"
country <- unique(main$countryName)[1]  # Assuming there's only one unique country

# Save the objects to an Excel file
save_all_objects_to_excel(paste0(rawdata, "/",country,"_",year, "_raw.xlsx"), tabnames)

