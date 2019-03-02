#' Grades output from Datalink and reports number of earned points
#' 
#' uses the .xlsx output from Datalink Connect software and will score responses by assigning a grade value to each possible response
#' @param input .xlsx exported from Datalink Connect read into R by using the command: read.xlsx2(file,sheetName="Results Grid")
#' @param answer.value a single value or a vector equal to number of answers assigning values to each possible response
#' @param possible a single value of total possible points
#' @return A dataframe consisting of student names, ID numbers and number of points missed on scanned questions
#' @export
#' @importFrom stringdist stringdist
#' @importFrom xlsx read.xlsx2

grade.earned=function(input,answer.value=0.2, possible){
  input[,c(1,7:ncol(input))]=lapply(input[,c(1,7:ncol(input))], as.character)
  input[,2]=as.numeric(as.character(input[,2]))
  spread=input[-1,]
  grades=cbind(input[2:length(input[,1]),1:2],rep(0,length(input[,1])-1))
  colnames(grades)=c("name","id","earned")
  for (i in 1:length(input[,1])-1) {
    key=as.character(input[1,7:length(input[1,])])
    name=which(grades$name==spread$Student.Name[i])
    answers=as.character(spread[name,7:length(input[1,])])
    total=possible-sum(stringdist(key,answers,method="lcs")*answer.value)
    grades$earned[i]=total
  }
  
  return(grades)
}