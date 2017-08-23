# Get the excel file and read the tab where the size column is located
library('XLConnect')
sthExcel = loadWorkbook("20170816_CH_new_15.xlsx")
sthSizeSheet = readWorksheet(sthExcel, sheet = 2)

# Get the size from sthSize
library('gsubfn')
numOfRowsInSthSizeSheet = nrow(sthSizeSheet)
measurementMatrix = matrix(nrow = numOfRowsInSthSizeSheet, ncol = 3, dimnames = list(c(),c('id','name','size')) )

for(i in 1:numOfRowsInSthSizeSheet) {
  measurementMatrix[i,'name']<-sthSizeSheet[i,'item_name2']
  measurementMatrix[i,'size']<-sthSizeSheet[i,'item_option2']
  # id should be fetched from shopify api
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
