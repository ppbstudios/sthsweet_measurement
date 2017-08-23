#install.packages('httr')
# Get the excel file and read the tab where the size column is located
library('XLConnect')
sthExcel = loadWorkbook("20170816_CH_new_15.xlsx")
sthSizeSheet = readWorksheet(sthExcel, sheet = 2)

# Get the size from sthSize
library('gsubfn')
library('httr')
numOfRowsInSthSizeSheet = nrow(sthSizeSheet)
measurementMatrix = matrix(nrow = numOfRowsInSthSizeSheet, ncol = 4, dimnames = list(c(),c('id','name','size','subtype')) )

for(i in 1:numOfRowsInSthSizeSheet) {
  measurementMatrix[i,'name']<-sthSizeSheet[i,'item_name2']
  measurementMatrix[i,'size']<-sthSizeSheet[i,'item_option2']
  measurementMatrix[i,'subtype']<-sthSizeSheet[i,'SUBTYPE']
  # id should be fetched from shopify api
  productName <- sthSizeSheet[i,'item_name2']
  reqQuery<-paste("{products(productParam:{fields:\"id\",title:",productName)
  reqQuery<-paste(reqQuery,"}) {id}}")
  res<-GET("http://localhost:3001/graphql", query=list(query=reqQuery))
  if(exists(res)) {
    resParsed<-content(res,"parsed")
    measurementMatrix[i,'id']<-resParsed$data$products[[1]]$id
  } else {
    measurementMatrix[i,'id']<-0
  }
  
  extractMeasure<-strapplyc(sthSizeSheet[i,'size'], '<h3>(.*?)</h3>', simplify = c)
  if(length(extractMeasure)) {
    key<-c()
    val<-c()
    for(j in 1:length(extractMeasure)) {
      # Devide key and value between :
      divideStr<-unlist(strsplit(extractMeasure[j],":"))
      key<-trimws(divideStr[1])
      val<-trimws(divideStr[2])
      # check if the key is in sizeMatrix column names
        # if not
      if(!(key %in% colnames(measurementMatrix))) {
        # create the column and add the val on the row
        measurementMatrix = cbind(measurementMatrix, c(NA))
        # add the val on the row
        colnames(measurementMatrix)[ncol(measurementMatrix)]<-c(key)
      }
      
      measurementMatrix[i,key]<-val
    }  
  }
}

# create an excel file
destFileName <- "sthsweetMeasurements.xlsx"
fileXls <- paste(getwd(), destFileName, sep='/')
NewXls <- loadWorkbook(fileXls, create = TRUE) # create only if the file name doesn't exist
# supplier name
supplier<-sthSizeSheet[1,'item_supplier']
# check the sheet
if(!existsSheet(NewXls,supplier)) {
  # if no supplier sheet, create tab with the supplier
  createSheet(NewXls,supplier)  
}
appendWorksheet(NewXls, measurementMatrix, sheet = supplier)
#writeWorksheet(NewXls, measurementMatrix, sheet = supplier)
saveWorkbook(NewXls)