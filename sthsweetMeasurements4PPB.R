#install.packages('httr')
# Get the excel file and read the tab where the size column is located
# install.packages('XLConnect')
library('XLConnect')
# install.packages('gsubfn')
# install.packages('httr')
#install.packages('xml2')
library('jsonlite')
library('gsubfn')
library('httr')

sourceFiles <- dir("source4PPB/")
for (file in sourceFiles) {
    sthExcel = loadWorkbook(paste("source4PPB/", file, sep = "/"))
    sthPPBSheet = readWorksheet(sthExcel, sheet = 2)
    print(paste("============ Measurement Start: ", file, " ============", sep = ""))
    # Get the size from sthSize
    numOfRowsInSthSizeSheet = nrow(sthPPBSheet)
    measurementMatrix = matrix(nrow = numOfRowsInSthSizeSheet, ncol = 5, dimnames = list(c(), c('id', 'name', 'size', 'subtype', 'color')))

    for (i in 1:numOfRowsInSthSizeSheet) {
        # add values
        measurementMatrix[i, 'name'] <- sthPPBSheet[i, 'ENG']
        measurementMatrix[i, 'size'] <- sthPPBSheet[i, 'SIZE']
        measurementMatrix[i, 'subtype'] <- sthPPBSheet[i, 'SUBTYPE']
        measurementMatrix[i, 'color'] <- sthPPBSheet[i, 'COLOR']

        # id should be fetched from shopify api by the product name
        productName <- sthPPBSheet[i, 'ENG']
        numOfProductRowIndexes <- which(grepl(productName, measurementMatrix[, 'name']))
        # check if there are products more than one
        if (length(numOfProductRowIndexes) > 1) {
            productIdRow <- numOfProductRowIndexes[1]
            measurementMatrix[i, 'id'] <- measurementMatrix[productIdRow, 'id']
        } else {
            print(paste("============ Fetching ID Start: ", productName, " ============", sep = ""))
            resQuery <- capture.output(cat(c("{products(productParam:{fields:\"id,vendor\",title:\"", productName, "\"}) {id,vendor}}"), sep = ""))
            res <- GET("http://192.168.0.23:3002/graphql", query = list(query = resQuery))

            if (res$status_code == 200) {
                resParsed <- content(res, "parsed")
                measurementMatrix[i, 'id'] <- resParsed$data$products[[1]]$id
                print(paste("============ Fetching ID Success End: ", productName, " ============", sep = ""))
            } else {
                measurementMatrix[i, 'id'] <- 0
                print(paste("============ Fetching ID Error End: ", productName, " ============", sep = ""))
            }
        }

        # the size should be fetched from ppbapps.com/jbkAdmin/api with the handle
        productHandle <- sthPPBSheet[i, 'HANDLE']
        print(paste("============ Fetching Size Start: ", productHandle, " - ", productName, " ============", sep = ""))
        resURL <- paste("http://ppbapps.com/jbkAdmin/api/product", productHandle, sep = "/")
        res <- readLines(resURL, warn = F, encoding = "UTF-8")
        resParsed <- fromJSON(res, simplifyVector=T)

        if (resParsed$message == "success") {
            extractMeasure <- resParsed$result$size$size_data[[1]][,2:3] #info_name(key) and size(value)
            if (nrow(extractMeasure)) {
                key <- c()
                val <- c()
                for (j in 1:nrow(extractMeasure)) {
                    # Devide key and value between :
                    key <- trimws(extractMeasure[j,1])
                    val <- trimws(extractMeasure[j,2])
                    # check if the key is in sizeMatrix column names
                    # if not
                    if (!(key %in% colnames(measurementMatrix))) {
                        # create the column and add the val on the row
                        measurementMatrix = cbind(measurementMatrix, c(NA))
                        # add the val on the row
                        colnames(measurementMatrix)[ncol(measurementMatrix)] <- c(key)
                    }

                    measurementMatrix[i, key] <- val
                }
            }
            print(paste("============ Fetching Size Success End: ", productHandle, " - ", productName, " ============", sep = ""))
        } else {
            print(paste("============ Fetching Size Error End: ", productHandle, " - ", productName, " ============", sep = ""))
        }
    }
    
    # check the dist folder where the result file stored
    if(!file.exists("dist")) {
      dir.create("dist")
    }
    # create an excel file
    destFileName <- "dist/sthsweetMeasurementsPPB.xlsx"
    fileXls <- paste(getwd(), destFileName, sep = '/')
    NewXls <- loadWorkbook(fileXls, create = TRUE) # create only if the file name doesn't exist
    # supplier name
    supplier <- resParsed$result$product$project_no
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
