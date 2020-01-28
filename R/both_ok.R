#' Will allow either marked or unmarked answers for a particular question and lead will be accepted.
#' 
#' Uses the exam answers read in from DataLink software, and import using read.xlsx function, and the grades output from grade.earned() or grade.missed() functions.
#' @param lead The multiple choice question and answer for which either marked or unmarked would be acceptable.  This should be a character string with the question number immediately followed by the lead. For example, if either answer (marked or unmarked) is acceptable for lead B of question 3, the string should be "3B", or the argument should be lead = "3B". This function only works for one lead.  
#' @param grades Output grades dataframe from grade.earned() or grade.missed() functions
#' @param exam Answer matrix from DataLink Excel format export
#' @param answer.value a single value or a vector equal to number of answers assigning values to each possible response

#' @return grades dataframe is updated
#' @export
#' @importFrom stringdist stringdist
#' @importFrom xlsx read.xlsx2

both.ok="17E"

both.ok=function(lead,grades,exam,answer.value=0.6){
  questions=as.numeric(gsub("(\\d+)\\w","\\1",both.ok))
  leads=gsub("\\d+(\\w)","\\1",both.ok)
  poss=c("A","B","C","D","E")
  i=1
  for (i in 1:length(questions)){
    if (grepl(leads[i],as.character(exam[1,questions[i]+6]))){
      grades$missed[!grepl(leads[i],as.character(exam[2:nrow(exam),questions[i]+6]))]=
        grades$missed[!grepl(leads[i],as.character(exam[2:nrow(exam),questions[i]+6]))]-answer.value
    } else {
      grades$missed[grepl(leads[i],as.character(exam[2:nrow(exam),questions[i]+6]))]=
        grades$missed[grepl(leads[i],as.character(exam[2:nrow(exam),questions[i]+6]))]-answer.value
    }
  }
}