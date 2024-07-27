rm(list=ls())
source("Pack.R")
library(KoboconnectR)
# Folder / Put the path name after downloading 
project <- "/Users/mouhamadoudiallo/Library/CloudStorage/OneDrive-AKADEMIYA2063/EATM/Application/EATM Scripts /EATM-S"
rawdata <- file.path(paste0(project,"/Rawdata"))
output <- file.path(paste0(project,"/Outputs"))

# Set up KoboToolbox credentials
username <- "username" # Kobo username
password <- "yourpassword" # Kobo password 
desired_name <- "Niger : Tableau de Bord du Commerce et des Marchés Agricoles de la CEDEAO - (EATM) Scorecard"
#desired_name <- "Tableau de Bord du Commerce et des Marchés Agricoles de la CEDEAO - [The ECOWAS Agriculture Trade and Market (EATM) Scorecard]"

url1 <- "kf.kobotoolbox.org"
year <- "2022" # Replace with the actual year if it's a variable
