#' Runs two diagnostics on multiple choice leads from multiple choice data
#' 
#' This function runs two diagnostics, difficulty index and discrimination index, on each lead of a multiple choice question matrix produced by a Datalink Connect software
#' @param input .xlsx exported from Datalink Connect read into R by using the command: read.xlsx2(file,sheetName="Results Grid")
#' @return An .xlsx spread sheet is written into the working directory 
#' @export
#' @importFrom xlsx saveWorkbook createWorkbook addDataFrame CellStyle


diagnostic=function(input){
  
  answer.cor=matrix(nrow=(ncol(input)-6),ncol=5)
  poss=c("A","B","C","D","E")
  colnames(answer.cor)=poss
  for (j in 1:(ncol(input)-6)){
    for (i in 1:5){
      if (grepl(poss[i],as.character(input[1,j+6]))){
        answer.cor[j,i]=round(cor(grades$missed,as.numeric(!grepl(poss[i],as.character(input[2:nrow(input),j+6])))),2)
      } else {
        answer.cor[j,i]=round(cor(grades$missed,as.numeric(grepl(poss[i],as.character(input[2:nrow(input),j+6])))),2)
      }
    }  
  }

  answer.per=matrix(nrow=(ncol(input)-6),ncol=5)
  colnames(answer.per)=poss
  for (j in 1:(ncol(input)-6)){
    for (i in 1:5){
      if (grepl(poss[i],as.character(input[1,j+6]))){
        answer.per[j,i]=round(sum(grepl(poss[i],as.character(input[2:nrow(input),j+6])))/nrow(grades),2)
      } else {
        answer.per[j,i]=round(sum(!grepl(poss[i],as.character(input[2:nrow(input),j+6])))/nrow(grades),2)
      }
    }  
  }
  answer.cor[is.na(answer.cor)]=0
  per.data=data.frame(answer.per)
  cor.data=data.frame(answer.cor)
  diagnostics=createWorkbook(type="xlsx")
  percent=createSheet(diagnostics,sheetName = "Percent Correct")
  addDataFrame(per.data,percent,col.names=F,row.names = F)
  discrim=createSheet(diagnostics,sheetName = "Discrimination Index")
  addDataFrame(cor.data,discrim,col.names=F,row.names = F)
  
  cs = CellStyle(diagnostics) +
    Fill(backgroundColor="indianred1", foregroundColor="indianred1",
         pattern="SOLID_FOREGROUND")
  
  
  ##Highlight negative Discrimination Indices##
  discrim.rows=getRows(discrim)
  cells <- getCells(discrim.rows)       # get cells
  # in the wb I import with loadWorkbook, numeric data starts in column 3
  # and the first two columns are row number and label number
  values <- lapply(cells, getCellValue) # extract the values
  # find cells meeting conditional criteria 
  highlight <- "test"
  for (i in names(values)) {
    x <- as.numeric(values[i])
    if (x<0 & !is.na(x)) {
      highlight <- c(highlight, i)
    }    
  }
  highlight <- highlight[-1]
  lapply(names(cells[highlight]),function(ii)setCellStyle(cells[[ii]],cs))
  
  ##Highlight negative Percent Correct##
  per.rows=getRows(percent)
  cells <- getCells(per.rows)       # get cells
  # in the wb I import with loadWorkbook, numeric data starts in column 3
  # and the first two columns are row number and label number
  values <- lapply(cells, getCellValue) # extract the values
  # find cells meeting conditional criteria 
  highlight <- "test"
  for (i in names(values)) {
    x <- as.numeric(values[i])
    if (x<0.50 & !is.na(x)) {
      highlight <- c(highlight, i)
    }    
  }
  highlight <- highlight[-1]
  lapply(names(cells[highlight]),function(ii)setCellStyle(cells[[ii]],cs))
  
  saveWorkbook(diagnostics,file = "Diagnostics.xlsx")
}

