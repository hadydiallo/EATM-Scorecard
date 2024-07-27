#S3 
# OC
#source("getData.R") # raw data  
source("Functions.R") # Add sheet 
# Create a new Excel workbook
wb <-NULL
wb <- createWorkbook()
overView <- "Data for parameters Objective C, see sheets"
addWorksheet(wb, sheetName ="about")
writeData(wb, sheet = "about",overView) 
saveWorkbook(wb, paste0(output,"/Objective_C_", country, ".xls"), overwrite = TRUE)

## Parameters for Objective C6, import

# # glimpse(table_2_1_1)
if (is.null(table_2_1_1)==F ) {
  OCsix <- table_2_1_1 %>% 
    mutate(  m_chapter = substr(Q2.1.1.b_importProducths4, 1, 2), ## HS2
             m_heading = substr(Q2.1.1.b_importProducths4, 1, 4) , ## HS4
             m_ag_prod_desc = prodlabM211,
             m_border_b = Q2.1.1.d_importMostWideBorder,
             m_country_origin = Q2.1.1.f_borderCountryEcowas,
             m_quantity =as.numeric(Q2.1.1.g_importQuantity),
             m_quantity_unit =Q2.1.1.h_importQuantityUnit,
             m_quantity_ton =ifelse(Q2.1.1.h_importQuantityUnit %in% c(2,4), m_quantity/1000, 
                                    ifelse(Q2.1.1.h_importQuantityUnit %in% c(1,3), m_quantity,NA)),
             m_value =as.numeric(Q2.1.1.i_importValue),
             m_value_unit =Q2.1.1.j_importValueUnit,
             
             m_appliedDuty = Q2.1.1.k_IfAppliedDuty,
             m_appliedOtherDuty = Q2.1.1.l_IfAppliedOtherDuty) %>% 
    select(starts_with("m_"))
  
  cat("Objective C : Results 6 >> ")
  # glimpse(OCsix)
  addSheetToWorkbook(paste0(output,"/Objective_C_", country, ".xls"), "OCsix", OCsix)
  
  ## Parameters for Objective C7, import
  # Rezsults 7 
  OCsept_1 <- table_2_1_1 %>% 
    mutate(  m_chapter = substr(Q2.1.1.b_importProducths4, 1, 2), ## HS2
             m_heading = substr(Q2.1.1.b_importProducths4, 1, 4) , ## HS4
             m_ag_prod_desc = prodlabM211,
             m_border_b = Q2.1.1.d_importMostWideBorder,
             m_country_origin = Q2.1.1.f_borderCountryEcowas,
             m_quantity =as.numeric(Q2.1.1.g_importQuantity),
             m_quantity_unit =Q2.1.1.h_importQuantityUnit,
             m_quantity_ton =ifelse(Q2.1.1.h_importQuantityUnit %in% c(2,4), m_quantity/1000, 
                                    ifelse(Q2.1.1.h_importQuantityUnit %in% c(1,3), m_quantity,NA)),
             m_value =as.numeric(Q2.1.1.i_importValue),
             m_value_unit =Q2.1.1.j_importValueUnit,
             m_appliedQuota = Q2.1.1.m_IfAppliedQuota ,
             m_appliedOthRestr = Q2.1.1.n_IfAppliedOtherRestriction
    ) %>% 
    select(starts_with("m_"))
  
  # # glimpse(OCsept_1)
  addSheetToWorkbook(paste0(output,"/Objective_C_", country, ".xls"), "OCsept_1", OCsept_1)
  
} else {
  cat("no data for this table 2.1.1")
}
## Parameters for Objective C7, export
## Sept 2 
# # glimpse(table_2_2_1)
if (is.null(table_2_2_1)==F) {
  OCsept_2 <- table_2_2_1 %>% 
    mutate(  x_chapter = substr(Q2.2.1.b_exportProducths4, 1, 2), ## HS2
             x_heading = substr(Q2.2.1.b_exportProducths4, 1, 4) , ## HS4
             x_ag_prod_desc = paste0(Q2.2.1.b_exportProducths4,"-", Q2.2.1.b_exportProducths4),
             x_border_b = Q2.2.1.d_exportMostWideBorder,
             x_country_destination =Q2.2.1.f_borderCountryEcowas,
             x_quantity =as.numeric(Q2.2.1.g_exportQuantity),
             x_quantity_unit =Q2.2.1.h_exportQuantityUnit,
             x_quantity_ton =ifelse(Q2.2.1.h_exportQuantityUnit %in% c(2,4), x_quantity/1000, 
                                    ifelse(Q2.2.1.h_exportQuantityUnit %in% c(1,3), x_quantity,NA)),
             x_value =as.numeric(Q2.2.1.i_exportValue),
             x_value_unit =Q2.2.1.j_exportValueUnit,
             
             x_appliedTax = Q2.2.1.k_IfAppliedDuty,
             x_appliedQuota = Q2.2.1.l_IfAppliedOtherDuty,
             x_appliedOtherRest = Q2.2.1.m_IfAppliedQuota) %>% 
    select(starts_with("x_"))
  
  # # glimpse(OCsept_2)
  addSheetToWorkbook(paste0(output,"/Objective_C_", country, ".xls"), "OCsept_2", OCsept_2)
  
} else {
  cat("no data for this table 2.2.1")
}

## Parameters for Objective C8
## C huit , from country
# # glimpse(table_1_3_1)
if (sum(!is.na(main$Q1.3.1.a_tariffNumber))>0) { 
  OChuit_NoEcowas <- table_1_3_1 %>% 
    mutate(number_tariff_Band = length(Q1.3.1.0_tariffBand),
           TBand_NonEcowas = Q1.3.1.0_tariffBand,
           PercentageD_NonEcowas = Q1.3.1.2_tarrifBand_PD,
           Category_NonEcowas = Q1.3.1.3_tarffBand_Category,
    ) %>% select(number_tariff_Band,TBand_NonEcowas,PercentageD_NonEcowas,
                 Category_NonEcowas)
  # # glimpse(OChuit_NoEcowas)
  addSheetToWorkbook(paste0(output,"/Objective_C_", country, ".xls"), "OChuit_NoEcowas", OChuit_NoEcowas)
  
  
  ## Ecowas templete 
  OChuit_Ecowas <- data.table( number_tariff_BandECOWAS = 5,
                               TBand_ecowas = c(0,1,2,3,4),
                               PercentageDD_ecowas = c(0,5,10,20,35),         
                               Category_ecowas = c("Essential social goods"
                                                   ,"Goods of primary necessity, raw goods, and capital goods"
                                                   ,"Intermediate goods and inputs", "Final Consumption goods or finished goods",
                                                   "Specific goods for economic development")    )
  
  # # glimpse(OChuit_Ecowas)
  addSheetToWorkbook(paste0(output,"/Objective_C_", country, ".xls"), "OChuit_Ecowas", OChuit_Ecowas)
  
} else {
  cat("no data for this table of tariff band")
}
## Parameters for Objective C9
## C neuf 

if (sum(!is.na(main$Q3.1.a_numberWidelyBorder))>0) {
  Cneuf_1 <- table_3 %>% 
    mutate(neuf_number_border = length(unique(Q3.1.b_nameMostWidely)),
           neuf_border_name   = Q3.1.b_nameMostWidely,
           neuf_country_border = Q3.1.a_Neighboring_Country,
           neuf_if_mostWidelyBorder = Q3.1.d_IfMostWidelyUsed)  %>% 
    select(starts_with("neuf"))
  
  # # glimpse(Cneuf_1)
  Cneuf_1$Neuf_NMWBorder <- main$Q3.1.a_numberWidelyBorder
  addSheetToWorkbook(paste0(output,"/Objective_C_", country, ".xls"), "Cneuf_1", Cneuf_1)
  
} else {
  cat("no data for this table section 3")
}

## Parameters for Objective C6, documents 
### C nefu 2 
# # glimpse(table_2_1_1_SPSdocs) # document table f(x)
# # glimpse(table_2_1_1) # {products} -x 

if (is.null(table_2_1_1_SPSdocs)==F) {
  ## select relevant variables 
  table_2_1_1L <- table_2_1_1 %>% 
    select(c(1:21),c("_submission__id", "_submission__uuid"))
  # # glimpse(table_2_1_1L)
  
  ## Reshape 
  table_2_1_1LL <- table_2_1_1L %>% 
    pivot_longer(
      cols = starts_with("Q2.1.1.p_importNameMandatoryDoc/"),
      names_to = "DocumentType",
      values_to = "IfPresence"
    ) %>%
    mutate(DocumentType = str_replace(DocumentType, "Q2.1.1.p_importNameMandatoryDoc/", ""))
  
  # # glimpse(table_2_1_1LL)
  # # glimpse(table_2_1_1_SPSdocs)
  ## Combine docs and products, border 
  table_2_1_1_SPSdocss <- merge(table_2_1_1_SPSdocs,table_2_1_1LL,
                                       by.y= c("_submission__id","DocumentType"),
                                       by.x = c("_submission__id","posit_Mdoc"),
                                       all.x = T)
  # # glimpse(table_2_1_1_SPSdocss)
  
  Cneuf_2  <- table_2_1_1_SPSdocss %>% 
    #filter(`_submission__submission_time` == "06/06/2024  10:13:51") %>% 
    mutate( m_chapter = substr(Q2.1.1.b_importProducths4, 1, 2), ## HS2
            m_heading = substr(Q2.1.1.b_importProducths4, 1, 4) , ## HS4
            m_ag_prod_desc = prodlabM211,
            m_border_b = Q2.1.1.d_importMostWideBorder,
            m_country_origin = Q2.1.1.f_borderCountryEcowas,
            m_quantity =as.numeric(Q2.1.1.g_importQuantity),
            m_quantity_unit =Q2.1.1.h_importQuantityUnit,
            m_quantity_ton =ifelse(Q2.1.1.h_importQuantityUnit %in% c(2,4), m_quantity/1000, 
                                   ifelse(Q2.1.1.h_importQuantityUnit %in% c(1,3), m_quantity,NA)),
            m_value =as.numeric(Q2.1.1.i_importValue),
            m_value_unit =Q2.1.1.j_importValueUnit,
            
            m_document_SPS = ifelse( posit_Mdoc =="SPSCertificates", 1,0),
            m_document_CertOrigin = ifelse( posit_Mdoc =="CertificateofOrigin", 1,0) ,
            m_document_OtherDoc = ifelse( posit_Mdoc =="OtherDocs", 1,0),
            m_REQUESTED = Q2.1.1.r_numberDocREQUESTED,
            m_ISSUED = Q2.1.1.S_numberDocISSUED,
            m_ISSUEDless15days = Q2.1.1.t_numberDocISSUEDless15days,
            m_ELECTRONICFORMAT = Q2.1.1.u_propElectronicFormat,
            m_REJECTION = Q2.1.1.v_numRejection,
            m_TIME_DAYS = Q2.1.1.V_importTimeIssuing,
            m_COST_LC = Q2.1.1.w_importCostDoc
    ) %>% 
    select(starts_with("m_"))
  cat("data for parameters 9, import ")
  # glimpse(Cneuf_2)
  addSheetToWorkbook(paste0(output,"/Objective_C_", country, ".xls"), "Cneuf_2", Cneuf_2)
  
} else {
  cat("no data for document table section  2.1.1")
}
## Parameters for Objective C9, export-doc
## Neuf 3 , export 

# # glimpse(table_2_2_1_SPSdocs)
# # glimpse(table_2_2_1)

if(is.null(table_2_2_1_SPSdocs)==F) { 
  
  table_2_2_1L <- table_2_2_1 %>% 
    select(c(1:21),c("_submission__id", "_submission__uuid"))
  # # glimpse(table_2_2_1L)
  
  ## Reshape
  table_2_2_1LL <- table_2_2_1L %>% 
    pivot_longer(
      cols = starts_with("Q2.2.1.p_exportNameMandatoryDoc/"),
      names_to = "DocumentType",
      values_to = "IfPresence"
    ) %>%
    mutate(DocumentType = str_replace(DocumentType, "Q2.2.1.p_exportNameMandatoryDoc/", ""))
  
  # # glimpse(table_2_1_1LL)
  # # glimpse(table_2_1_1_SPSdocs)
  ## Combine docs and products, border 
  table_2_2_1_SPSdocss <- merge(table_2_2_1_SPSdocs,table_2_2_1LL,
                                       by.y= c("_submission__id","DocumentType"),
                                       by.x = c("_submission__id","posit_Xdoc"),
                                       all.x = T)
  # # glimpse(table_2_2_1_SPSdocss)
  
  Cneuf_3  <- table_2_2_1_SPSdocss %>% 
    #filter(`_submission__submission_time` == "06/06/2024  10:13:51") %>% 
    mutate(  x_chapter = substr(Q2.2.1.b_exportProducths4, 1, 2), ## HS2
             x_heading = substr(Q2.2.1.b_exportProducths4, 1, 4) , ## HS4
             x_ag_prod_desc = prodlabX221,
             x_border_b = Q2.2.1.d_exportMostWideBorder,
             x_country_destination =Q2.2.1.f_borderCountryEcowas,
             x_quantity =as.numeric(Q2.2.1.g_exportQuantity), 
             x_quantity_unit = Q2.2.1.h_exportQuantityUnit,
             x_quantity_ton =ifelse(Q2.2.1.h_exportQuantityUnit %in% c(2,4),x_quantity /1000, 
                                    ifelse(Q2.2.1.h_exportQuantityUnit %in% c(1,3), x_quantity,NA)),
             x_value =as.numeric(Q2.2.1.i_exportValue),
             x_value_unit =Q2.2.1.j_exportValueUnit,
             
             x_document_SPS = ifelse( posit_Xdoc =="SPSCertificates", 1,0),
             x_document_CertOrigin = ifelse( posit_Xdoc =="CertificateofOrigin", 1,0) ,
             x_document_OtherDoc = ifelse( posit_Xdoc =="OtherDocs", 1,0),
             x_REQUESTED = Q2.2.1.r_numberDocREQUESTED,
             x_ISSUED = Q2.2.1.S_numberDocISSUED,
             x_ISSUEDless15days = Q2.2.1.t_numberDocISSUEDless15days,
             x_ELECTRONICFORMAT = Q2.2.1.u_propElectronicFormat,
             x_REJECTION = Q2.2.1.v_numRejection,
             x_TIME_DAYS = Q2.2.1.V_exportTimeIssuing,
             x_COST_LC = Q2.2.1.w_exportCostDoc
    ) %>% 
    select(starts_with("x_"))
  cat("data for parameters objective 9, export ")
  # glimpse(Cneuf_3)
  addSheetToWorkbook(paste0(output,"/Objective_C_", country, ".xls"), "Cneuf_3", Cneuf_3)
  
} else {
  cat("no data for document table section  2.2.1")
}
## Parameters for Objective C11, border
## Oonze 
# # glimpse(table_3)
if (sum(!is.na(main$Q3.1.a_numberWidelyBorder))>0) { 
  Conze <- table_3 %>% 
    mutate(onze_number_border = length(unique(Q3.1.b_nameMostWidely)),
           onze_border_name   = Q3.1.b_nameMostWidely,
           onze_country_border = Q3.1.a_Neighboring_Country,
           onze_if_mostWidelyBorder = Q3.1.d_IfMostWidelyUsed,
           onze_if_coordJoinHours =Q3.1.f_BorderWithCoord_OHC,
           onze_if_jointPhysiInspect=Q3.1.e_BorderWithCoord_PI,
           onze_if_jointInspWithNeigb = Q3.1.g_BorderWithCoord_JPI,
           
    )  %>% 
    select(starts_with("onze"))
  # glimpse(Conze)
  
  Conze$Onze_NMWBorder <- main$Q3.1.a_numberWidelyBorder
  addSheetToWorkbook(paste0(output,"/Objective_C_", country, ".xls"), "Conze", Conze)
  
} else {
  cat("no data for table section 3")
}

## Parameters for Objective C12
## Douze 
# # glimpse(table_5_2)

if (is.null("table_5_2")) { 
  Cdouze_1 <- table_5_2 %>% 
    mutate(douze_legislation = Q5.2.b_legislationInfos,
           douze_numberLegislation = length(Q5.2.b_legislationInfos),
           douze_legislation5years = Q5.2.c_legislation5years
    )  %>% 
    select(starts_with("douze"))
  # # glimpse(Cdouze_1) 
  addSheetToWorkbook(paste0(output,"/Objective_C_", country, ".xls"), "Cdouze_1", Cdouze_1)
  
} else {
  cat("no data for table section 5.2")
}
## douze 2 

# # glimpse(table_5_1)
if(is.null("table_5_1")) {
  Cdouze_2 <- table_5_1 %>% 
    mutate(douze_number = length(Q5.1.b_committeeDetails),
      douze_committeeDetail = Q5.1.b_committeeDetails,
           douze_committeeSPSTBT = Q5.1.c_committeeSPSTBT,
    )  %>% 
    select(starts_with("douze"))
  # glimpse(Cdouze_2)
  addSheetToWorkbook(paste0(output,"/Objective_C_", country, ".xls"), "Cdouze_2", Cdouze_2)
  
} else {
  cat("no data for table section 5.1")
}

cat("End Objectif C; Please have look to the output folder ")
cat(" #Warning :: this a compilation, not a treatment. ")
cat(" #Attention :: ceci est une simple compilation, y'a aucun traitement. ")
