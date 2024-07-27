# S0_annex

# Unit quantity and transport
UnitTransport <- c(1, 2, 3, 4, 5)
English <- c("Ton", "Kilogram", "Hectolitre", "Liter", "Head")
French <- c("Tonne", "Kilogramme", "Hectolitre", "Litre", "Tête")
UnitTransport <- data.table(UnitTransport, English, French)

# Unit value
Unit_Value <- c(1, 2, 3, 4, 5, 6, 7, 8, 9, 10)
English <- c("West African CFA franc (XOF)", "Sierra Leonean leone (SLL)", 
             "Central African CFA franc (XAF)", "Ghanaian cedi (GHS)",
             "Guinean franc (GNF)", "Liberian dollar (LRD)",
             "Nigerian naira (NGN)", "Mauritanian Ouguiya (MRU)", 
             "Gambian Dalasi (GMD)", "Cape Verdean Escudo (CVE)")
French <- c("Franc CFA d'Afrique de l'ouest (XOF)", "Sierra Leone (SLL)", 
            "Franc CFA d'Afrique centrale (XAF)", "Cedi ghanéen (GHS)", 
            "Franc guinéen (GNF)", "Dollar libérien (LRD)", "Naira nigérian (NGN)", 
            "Ouguiya Mauritanien (MRU)", "Dalasi Gambie (GMD)",
            "Escudo Cabo Verde (CVE)")
Unit_Value <- data.table(Unit_Value, English, French)

# QUantity
Unit_Quantity <- c(1, 2, 3, 4, 5)
English <- c("Ton", "Kilogram", "Hectolitre", "Liter", "Head")
French <- c("Tonne", "Kilogramme", "Hectolitre", "Litre", "Tête")
Unit_Quantity <- data.frame(Unit_Quantity, English, French)

separator <- data.frame(matrix(ncol = ncol(Unit_Value), nrow = 1))
colnames(separator) <- colnames(Unit_Value)
Annex <- bind_rows(Unit_Value, separator, Unit_Quantity, separator, UnitTransport)

