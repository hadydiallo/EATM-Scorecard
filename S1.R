# Process 
## Years 
# Remove NA values from the vectors
Q1.4.1.0_MReferenceYear1_clean <- na.omit(main$Q1.4.1.0_MReferenceYear1)
Q1.4.2.0_MReferenceYear1_clean <- na.omit(main$Q1.4.2.0_MReferenceYear1)

# Assign the first unique value to year1 and year2
year1 <- unique(Q1.4.1.0_MReferenceYear1_clean)[1] # reference year
year2 <- unique(Q1.4.2.0_MReferenceYear1_clean)[1] # base year

cat("Reference year :", year1, " Base year : ", year2)
cat("    ")
cat("Parameters for importations")

## IMPORT
### import year 1 
if (!is.na(year1)) {
  ai1My1 <- table_1_4_1 %>%
    filter(Q1.4.1.e2_ifYear1==1) %>%
    reframe(Year = year1, 
            HS4_code = substr(Q1.4.1.4_agriProducths4, 1, 4),
            Goods  =  trimws(sub(".+?-\\s*", "", prodlabM141)),
            CountryOrigin = Q1.4.1.e_CountryOrigin,
            AppliedTax = Q1.4.1.j_AppliedCustomDuty, ## droit de douane
            AppliedRest = Q1.4.1.m_Applied_other_import_restriction,
            AppliedOtherRest = Q1.4.1.k_Applied_other_charges,
            ImportRestriction = Q1.4.1.n_Applied_other_import_restrictionS, ## Text
            Value_M = Q1.4.1.h_Import_value,
            UnitValue_M = Q1.4.1.i_Value_Unit,
            Quantity_M = as.numeric(Q1.4.1.f_Import_quantity),
            UnitQuatity_M = Q1.4.1.g_Quantity_Unit,
            Quantity_M_KG = ifelse(UnitQuatity_M %in% c(2,4), Quantity_M, 
                                   ifelse(UnitQuatity_M==1, Quantity_M*1000, 
                                          ifelse(UnitQuatity_M==3, Quantity_M*100,NA) ))
            
    )
  ### import year2 
  ai1My2 <- table_1_4_1 %>% 
    filter(Q1.4.2.e2_ifYear2==1) %>%
    reframe(Year = year2, 
            HS4_code = substr(Q1.4.1.4_agriProducths4, 1, 4),
            Goods  =  trimws(sub(".+?-\\s*", "", prodlabM141)),
            CountryOrigin = Q1.4.1.e_CountryOrigin,
            Value_M = Q1.4.2.h_Import_value,
            UnitValue_M = Q1.4.2.i_Value_Unit,
            Quantity_M = as.numeric(Q1.4.2.f_Import_quantity),
            UnitQuatity_M = Q1.4.2.g_Quantity_Unit,
            Quantity_M_KG = ifelse(UnitQuatity_M %in% c(2,4), Quantity_M, 
                                   ifelse(UnitQuatity_M==1, Quantity_M*1000, 
                                          ifelse(UnitQuatity_M==3, Quantity_M*100,NA) ))
            
    )
  
  ## For the outputs (Sample from Babacar)
  Ag_goods_and_inputs_imports_v1 <- data.table(
    chapter = substr(ai1My1$HS4_code, 1, 2), ## HS2
    heading = ai1My1$HS4_code , ## HS4
    ag_prod_desc = paste0(ai1My1$HS4_code,"-", ai1My1$Goods),
    year1 = ai1My1$Year ,
    year2 = ai1My2$Year ,
    
    ecowas_trading_partner_country_year1 = ai1My1$CountryOrigin,
    ecowas_trading_partner_country_year2 = ai1My2$CountryOrigin,
    subject_to_tarif_other_charges_year1=ai1My1$AppliedOtherRest,
    tarif_or_other_charges_year1 = ai1My1$AppliedTax, ## Not htis variable instead use :: Q1.5.2.k_Applied_other_charges_other,
    affected_by_import_restrict_year1 = ai1My1$AppliedRest,
    import_restrict_or_prohibitions_year1 = ai1My1$ImportRestriction,
    value_base_year_year2 = ai1My2$Value_M,
    quantity_base_year_year2 = ai1My2$Quantity_M_KG,
    unit_quantity_base_year_year2 = ai1My2$UnitQuatity_M,
    
    value_ref_year_year1 = ai1My1$Value_M,
    value_unit_ref_year_year1 = ai1My1$UnitQuatity_M,
    quantity_ref_year_year1 = ai1My1$Quantity_M,
    quantity_unit_ref_year_year1 = ai1My1$UnitQuatity_M,
    quantity_KG_ref_year_year1 = ai1My1$Quantity_M_KG,
    
    value_base_year_year2 = ai1My2$Value_M,
    value_unit_base_year_year2 = ai1My2$UnitQuatity_M,
    quantity_base_year_year2 = ai1My2$Quantity_M,
    quantity_unit_base_year_year2 = ai1My2$UnitQuatity_M,
    quantity_KG_base_year_year2= ai1My2$Quantity_M_KG
  )
   
  cat("output variables for import ")
  # glimpse(Ag_goods_and_inputs_imports_v1)
  
  cat("Parameters for exportation")
  ## EXPORT 
  ### year 1
  # # glimpse(S1_IntraExport)
  ai1Xy1 <- table_1_5_1 %>% 
    filter(Q1.5.1.e2_ifYear1==1) %>% 
    reframe(Year = year1, 
            HS4_code = substr(Q1.5.1.4_agriProducths4, 1, 4),
            Goods  =  trimws(sub(".+?-\\s*", "", prodlabX151)),
            CountryDestination = Q1.5.1.e_CountryOrigin,
            AppliedTax = Q1.5.1.j_AppliedCustomDuty, ## droit de douane
            AppliedRest = Q1.5.1.m_Applied_other_Export_restriction,
            ExportRestriction = Q1.5.1.n_Applied_other_Export_restrictionS, ## Text
            Value_X = Q1.5.1.h_Export_value, # Assume currency identique 
            UnitValue_X = Q1.5.1.i_Value_Unit,
            Quantity_X = as.numeric(Q1.5.2.f_Export_quantity),
            UnitQuatity_X = Q1.5.1.g_Quantity_Unit,
            Quantity_X_KG = ifelse(UnitQuatity_X %in% c(2,4), Quantity_X, 
                                   ifelse(UnitQuatity_X==1, Quantity_X*1000, 
                                          ifelse(UnitQuatity_X==3, Quantity_X*100,NA) ))
    )
  ai1Xy1 <- data.table(ai1Xy1)
  ## year 2, 
  ai1Xy2 <- table_1_5_1 %>% 
    filter(Q1.5.2.e2_ifYear2==1) %>%  
    reframe(Year = year2, 
            HS4_code = substr(Q1.5.1.4_agriProducths4, 1, 4),
            Goods  =  trimws(sub(".+?-\\s*", "", prodlabX151)),
            CountryDestination = Q1.5.1.e_CountryOrigin,
            Value_X = Q1.5.2.h_Export_value,
            UnitValue_X = Q1.5.2.g_Quantity_Unit,
            Quantity_X = as.numeric(Q1.5.2.f_Export_quantity),
            UnitQuatity_X = Q1.5.2.g_Quantity_Unit,
            Quantity_X_KG = ifelse(UnitQuatity_X %in% c(2,4), Quantity_X, 
                                   ifelse(UnitQuatity_X==1, Quantity_X*1000, 
                                          ifelse(UnitQuatity_X==3, Quantity_X*100,NA) ))
    )
  ai1Xy2 <- data.table(ai1Xy2)
  ## For parameters, sample from Babacar
  Ag_goods_and_inputs_exports_v1 <- data.table(
    chapter = substr(ai1Xy1$HS4_code, 1, 2), ## HS2
    heading = ai1Xy1$HS4_code , ## HS4
    ag_prod_desc = paste0(ai1Xy1$HS4_code,"-", ai1Xy1$Goods),
    year1 = ai1Xy1$Year ,
    year2 = ai1Xy2$Year ,
    ecowas_trading_partner_country_year1 = ai1Xy1$CountryDestination,
    ecowas_trading_partner_country_year2 = ai1Xy2$CountryDestination,
    export_tax_applied_if_any_year1 = ai1Xy1$AppliedTax,
    affected_by_export_restrict_year1 = ai1Xy1$AppliedRest,
    export_restrict_or_prohibitions_year1 = ai1Xy1$ExportRestriction,
    
    value_ref_year_year1 = ai1Xy1$Value_X,
    value_unit_ref_year_year1 = ai1Xy1$UnitQuatity_X,
    quantity_ref_year_year1 = ai1Xy1$Quantity_X,
    quantity_unit_ref_year_year1 = ai1Xy1$UnitQuatity_X,
    quantity_KG_ref_year_year1 = ai1Xy1$Quantity_X_KG,
    
    value_base_year_year2 = ai1Xy2$Value_X,
    value_unit_base_year_year2 = ai1Xy2$UnitQuatity_X,
    quantity_base_year_year2 = ai1Xy2$Quantity_X,
    quantity_unit_base_year_year2 = ai1Xy2$UnitQuatity_X,
    quantity_KG_base_year_year2= ai1Xy2$Quantity_X_KG
  )
  
  cat("output variables for export ")
  # glimpse(Ag_goods_and_inputs_exports_v1) 
  
  ## Save file into outputs 
  #write_csv(Ag_goods_and_inputs_exports_v1, paste0(output,"/Ag_goods_and_inputs_exports_", country,".csv"))
  #write_csv(Ag_goods_and_inputs_imports_v1, paste0(output,"/Ag_goods_and_inputs_imports_", country,".csv"))
  
  # OA
  # Create a new Excel workbook
  overView <- "Objectif A : data for computing parameters"
  wb <- createWorkbook()
  Annex <- as.list(Annex)
  df <- data.frame(Over = overView, Sep = "=>>==>>==>>=", Annex)
  # Add "about" sheet and write data
  addWorksheet(wb, sheetName = "about")
  writeData(wb, sheet = "about", df) 
  
  # Add "Ag_goods_and_inputs_exports" sheet and write data
  addWorksheet(wb, sheetName = "Ag_goods_and_inputs_exports")
  writeData(wb, sheet = "Ag_goods_and_inputs_exports", Ag_goods_and_inputs_exports_v1) 
  
  # Add "Ag_goods_and_inputs_imports" sheet and write data
  addWorksheet(wb, sheetName = "Ag_goods_and_inputs_imports")
  writeData(wb, sheet = "Ag_goods_and_inputs_imports", Ag_goods_and_inputs_imports_v1) 
  
  # Remove spaces from the country variable
  country <- gsub(" ", "", country)
  # Define the output path
  output_path <- paste0(output, "/Objective_A_", country, ".xls")
  
  # Save the workbook
  saveWorkbook(wb, output_path, overwrite = TRUE)
  
  # Print the output path for confirmation
  print(paste("Workbook saved at:", output_path))
  
} else {
  print("Data does not exist for this objective")
}
