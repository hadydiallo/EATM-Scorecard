
#' First set credentials >> Credentials.R
#' Second Install all package >> Pack
#' Third : Run "RUN.R" 

# Load data 
source("dowloadImport.R")
source("S0_annex.R") # Meta

sink("Summury/console_output.txt")
## Objective A 
source("S1.R")

## Objective B
source("S2.R")

## Objective C
source("S3.R")


# Check the saved output
file.show("Summury/console_output.txt")