#' Add names by ID number to a Datalink Connect output .xlsx file
#' 
#' uses the .xlsx output from Datalink Connect software and will score responses by assigning a grade value to each possible response
#' @param input .xlsx exported from Datalink Connect read into R by using the command: read.xlsx2(file,sheetName="Results Grid")
#' @param id a dataframe containing IDs and Names of students appearing in input file.  Column with ID numbers should be named "ID" and column with names should be named "Name"
#' @return A the same dataframe read in as object with names added to the appropriate column
#' @export
addnames=function(input, id){
  input[,c(1,7:ncol(input))]=lapply(input[,c(1,7:ncol(input))], as.character)
  input[,2]=as.numeric(as.character(input[,2]))
  id$Name=as.character(id$Name)
  for (i in 2:nrow(input)){
    input$Student.Name[i]=id$Name[which(id$ID==input$ID.Number[i])]
  }
return(input)
}