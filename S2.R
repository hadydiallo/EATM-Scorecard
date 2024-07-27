#S2
## Objective B
#source("getData.R") # raw data  
cat("data for import ")
# IMPORT 

### IMPORT data
# # glimpse(table_2_1_1_SPSdocs)
# # glimpse(table_2_1_1)

if (is.null(table_2_1_1_SPSdocs)==F ) {
  table_2_1_1L <- table_2_1_1 %>% 
    select(c(1:21),c("_submission__id", "_submission__uuid"))
  # # glimpse(table_2_1_1L)
  
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
  
  objectiveBm  <- table_2_1_1_SPSdocss %>% 
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
  # # glimpse(objectiveBm)
  
} else {
  cat("no data for document table section  2.1.1")
}

cat("data for export ")
## Export 
## Neuf 3 , export 
# # glimpse(table_2_2_1_SPSdocs)
# # glimpse(table_2_2_1)

if (is.null(table_2_2_1_SPSdocs)==F ) {
  
  table_2_2_1L <- table_2_2_1 %>% 
    select(c(1:21),c("_submission__id", "_submission__uuid"))
  # # glimpse(table_2_2_1L)
  
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
  
  objectiveBx  <- table_2_2_1_SPSdocss %>% 
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
  # # glimpse(objectiveBx)
  
} else {
  cat("no data for document table section  2.1.1")
}

# OB
# Create a new Excel workbook
overView <- "Objectif B : data for computing parameters"
# Combine data frames into a single data frame
Annex <- as.list(Annex)
df <- data.frame(Over = overView, Sep = "=>>==>>==>>=", Annex)
if (is.null(objectiveBm)==F & is.null(objectiveBx)==F ) { 
  wb <-NULL
  wb <- createWorkbook()
  addWorksheet(wb, sheetName ="about")
  writeData(wb, sheet = "about",df) 
  addWorksheet(wb, sheetName ="IMPORTS")
  writeData(wb, sheet = "IMPORTS",objectiveBm) 
  addWorksheet(wb, sheetName ="EXPORTS")
  writeData(wb, sheet = "EXPORTS",objectiveBx) 
  saveWorkbook(wb, paste0(output,"/Objective_B_", country, ".xls"), overwrite = TRUE)
  
} else {
  cat("no data for Objective B")
}
