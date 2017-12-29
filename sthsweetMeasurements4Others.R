#install.packages('httr')
# Get the excel file and read the tab where the size column is located
# install.packages('XLConnect')
library('XLConnect')
# install.packages('gsubfn')
# install.packages('httr')
library('gsubfn')
library('httr')

sourceFiles <- dir("source4Others/")
graphqlServer <- "http://shopify.ppb.st/graphql"
for (file in sourceFiles) {
    sthExcel = loadWorkbook(paste("source4Others/", file, sep = "/"))
    sthSizeSheet = readWorksheet(sthExcel, sheet = 2)
    print(paste("============ Measurement Start: ", file, " ============", sep = ""))
    # Get the size from sthSize
    numOfRowsInSthSizeSheet = nrow(sthSizeSheet)
    measurementMatrix = matrix(nrow = numOfRowsInSthSizeSheet, ncol = 5, dimnames = list(c(), c('id', 'name', 'size', 'subtype', 'color')))

    for (i in 1:numOfRowsInSthSizeSheet) {
        # add values
        measurementMatrix[i, 'name'] <- sthSizeSheet[i, 'item_name2']
        measurementMatrix[i, 'size'] <- sthSizeSheet[i, 'item_option2']
        measurementMatrix[i, 'subtype'] <- sthSizeSheet[i, 'SUBTYPE']
        measurementMatrix[i, 'color'] <- sthSizeSheet[i, 'item_option']

        # id should be fetched from shopify api by the product name
        productName <- sthSizeSheet[i, 'item_name2']
        numOfProductRowIndexes <- which(grepl(productName, measurementMatrix[, 'name']))
        # check if there are products more than one
        if (length(numOfProductRowIndexes) > 1) {
            productIdRow <- numOfProductRowIndexes[1]
            measurementMatrix[i, 'id'] <- measurementMatrix[productIdRow, 'id']
        } else {
            print(paste("============ Fetching ID Start: ", productName, " ============", sep = ""))
            resQuery <- capture.output(cat(c("{products(productParam:{fields:\"id\",title:\"", productName, "\"}) {id}}"), sep = ""))
            res <- GET(graphqlServer, query = list(query = resQuery))

            if (res$status_code == 200) {
                resParsed <- content(res, "parsed")
                measurementMatrix[i, 'id'] <- resParsed$data$products[[1]]$id
                print(paste("============ Fetching ID Success End: ", productName, " ============", sep = ""))
            } else {
                measurementMatrix[i, 'id'] <- 0
                print(paste("============ Fetching ID Error End: ", productName, " ============", sep = ""))
            }
        }

        # extractMeasure <- strapplyc(sthSizeSheet[i, 'size'], '<h3>(.*?)</h3>', simplify = c)
        # if (length(extractMeasure)) {
        #     key <- c()
        #     val <- c()
        #     for (j in 1:length(extractMeasure)) {
        #         # Devide key and value between :
        #         match <- regexec("-?\\s*?\\d*\\.?\\d*?\\s*cm",extractMeasure[j]) # regex patter for the value ex - 43.23 cm, 23 cm etc
        #         key <- gsub(x=trimws(regmatches(extractMeasure[j],match, invert = T)[[1]][1]),pattern = ":",replacement = "") # extract the key w/o colon ex :
        #         val <- strsplit(trimws(regmatches(extractMeasure[j],match, invert = F)), split = "cm")[[1]][1] # extract the value only w/o unit ex) cm
        #         # check if the key is in sizeMatrix column names
        #         # if not
        #         if (!(key %in% colnames(measurementMatrix))) {
        #             # create the column and add the val on the row
        #             measurementMatrix = cbind(measurementMatrix, c(NA))
        #             # add the val on the row
        #             colnames(measurementMatrix)[ncol(measurementMatrix)] <- c(key)
        #         }
        # 
        #         measurementMatrix[i, key] <- val
        #     }
        # }
    }
    
    # check the dist folder where the result file stored
    if(!file.exists("dist")) {
      dir.create("dist")
    }
    # create an excel file
    destFileName <- "dist/sthsweetMeasurementsOthers.xlsx"
    fileXls <- paste(getwd(), destFileName, sep = '/')
    NewXls <- loadWorkbook(fileXls, create = TRUE) # create only if the file name doesn't exist
    # supplier name
    supplier <- sthSizeSheet[1, 'item_supplier']
    # check the sheet
    if (!existsSheet(NewXls, supplier)) {
        print(paste("============ Create Sheet Start: ", supplier, " ============", sep = ""))
        # if no supplier sheet, create tab with the supplier
        createSheet(NewXls, supplier)
        writeWorksheet(NewXls, measurementMatrix, sheet = supplier)
        print(paste("============ Create Sheet End: ", supplier, " ============", sep = ""))
    } else {
        appendWorksheet(NewXls, measurementMatrix, sheet = supplier)
    }

    saveWorkbook(NewXls)
    print(paste("============ Measurement End: ", file, " ============", sep = ""))
}