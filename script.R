#Saurav Kharb
#Research Assistant - Foster School of Business
#This script automates data download from google trends and process files
# Libraries required for the program
require("XLConnect")
library(readxl)
mkey <- read_excel("ac1e.xlsx")
library(gtrendsR)
ch <- gconnect("redbullsaurav@gmail.com", "abhishekborah")
library(xlsx)


# nea <- mkey$`Table 1`[2]
# for (i in 3:702){
#   s <- mkey$`Table 1`[i]
#   a <- grepl("cl",s)
#   
#   if(a){
#     nea <- c(nea,mkey$`Table 1`[i])
#   }
# }
# #print(cl_filtered)
# 
# nea_1k_nohundred <- mkey$`Table 1`[2]
# for (i in 3:702){
#   s <- mkey$`Table 1`[i]
#   a <- grepl("cl",s)
#   k <- grepl("1K",s)
#   h <- grepl("\t100",s)
#   if(a && k && !h){
#     nea_1k_nohundred <- c(nea_1k_nohundred,mkey$`Table 1`[i])
#   }
# }
# 
# 
# #final work 
# 
# 
# 
# ultimate <- read_excel("newwords.xlsx")
# 
# studentdata[studentdata$Drink == 'water',]
# 
# newList <- ultimate[ultimate$Status == 'new' , "Master"]
# 
# 
# finalData <- ultimate[[2]][[2]]

########################################################################################
# topTen <- finalData
# print(finalData)
# for( i in 3:700){
#   boolName <- grepl("cl", ultimate[[2]][[i]])
#   if(boolName){
#     topTen <- c(topTen,ultimate[[2]][[i]])
#   }
# }
# 
# print(topTen)
# print(length(topTen))
# # topTen is a list of all relevant words 
# 
# 
# #count = 0;
# for(i in 1:length(topTen)){
#   boolSearch_one_to_ten <- grepl("1K",ultimate[[4]][[i]])
#   boolSearch_ten_to_hun <- grepl("100K",ultimate[[4]][[i]])
#   boolSearchHundred <- grepl("100 ",ultimate[[4]][[i]])
#   if(boolSearch_ten_to_hun) {
#     finalData <- c(finalData,ultimate[[2]][[i]] )
#     
#   }
#   else if((!boolSearchHundred && boolSearch_one_to_ten)) {
#     finalData <- c(finalData,ultimate[[2]][[i]] )
#    
#   }
# }
# print(finalData)
########################################################################################

# Accepts file and a keyword and sorts the top 10 keywords from Google adwords csv file

ultimate <- read_excel("xc90.xlsx")
keyword <- "xc90"

finalData <- ultimate[[2]][[2]]
print(finalData)
leftOver <- "waste" 
for(i in 3:700){
  boolName <- grepl(keyword,ultimate[[2]][[i]])
  boolSearch_one_to_ten <- grepl("1K",ultimate[[4]][[i]])
  boolSearch_ten_to_hun <- grepl("100K",ultimate[[4]][[i]])
  boolSearchHundred <- grepl("100 ",ultimate[[4]][[i]])
  #print(ultimate[[2]][[i]])
  
  if(boolSearch_ten_to_hun && boolName){
    finalData <- c(finalData,ultimate[[2]][[i]] )
  }
  else if((!boolSearchHundred && boolSearch_one_to_ten)&&boolName){
    finalData <- c(finalData,ultimate[[2]][[i]] )
  }else{
    leftOver <- c(leftOver, ultimate[[2]][[i]])
  }
}
print(finalData)

cleanLeftOver <- "waste"
for( i in 1:length(leftOver)){
  boolName <- grepl(keyword, leftOver[i])
  if(boolName){
    cleanLeftOver <- c(cleanLeftOver,leftOver[i])
  }
}

#print(cleanLeftOver)
hash <- FALSE

if(length(finalData) > 10 ) {
  top <- finalData[1]
  for(i in 2:10){
    top <- c(top, finalData[i])
  }
  write(top, file = "listof10", append = TRUE)
  write("\n", file = "listof10", append = TRUE)
  print(top)
  
}else if(length(finalData)<10) {
  hash <- TRUE
  required <- 10-length(finalData);
  if(required == 1){
    finalData <- c(finalData,cleanLeftOver[2])
  }else{
    for(i in 2: (required+1)){
      finalData <- c(finalData,cleanLeftOver[i])
    }
  } 
  write(finalData, file = "listof10", append = TRUE)
  write("\n", file = "listof10", append = TRUE)
  print(finalData)
}else{
  write(finalData, file = "listof10", append = TRUE)
  write("\n", file = "listof10", append = TRUE)
  print(finalData)
}


###############################################################z##########

# Code for downloading the data for each keyword with Use in it

ultimate <- read_excel("usage2.xlsx")
for(i in 3369:3656){
  if(ultimate[[4]][[i]] == "Use"){
    Sys.sleep(8)
    tryCatch({
      lang_trend <- gtrends(ultimate[[3]][[i]], geo="US", start_date = as.Date("2004-01-1"),end_date = as.Date("2016-9-1"))
      print(str(lang_trend[[3]]))
      lol<- paste(ultimate[[3]][[i]],".csv",sep="")
      write.csv(lang_trend[[3]], file = lol)
    }, error=function(e){
      print(i)
      print(ultimate[[3]][[i]])
      print("not found")
      write(ultimate[[3]][[i]], file = "notFound", append = TRUE)
      write("\n", file = "notFound", append = TRUE)
    })
  }
  
}
#######################################################################################

ultimate <- read_excel("usage2.xlsx")

count  = 1;
name = 1
#length(ultimate[[2]])
term  = ultimate[[3]][[1]]
for(i in 1:3660){
  count = count + 1
  if(i %% 10 != 1){
    if(ultimate[[4]][[i]] == "Use"){
      term <- paste(term,"+",sep="")
      term <- paste(term,ultimate[[3]][[i]],sep="")
      #print(ultimate[[3]][[i]])
    }
    
  }else{     
    print(term)
    print(name)
    
    tryCatch({
      lang_trend <- gtrends(term, geo="US", start_date = as.Date("2004-01-1"),end_date = as.Date("2016-9-1"))
      print(str(lang_trend[[3]]))
      lol<- paste(name,".csv",sep="")
      write.csv(lang_trend[[3]], file = lol)
    }, error=function(e){
      print("not found")
    })
    #write(term, file = "summation", append = TRUE)
    #write("\n", file = "summation", append = TRUE)
    name = name + 1
    count = 1
    term <- ultimate[[3]][[i]]
    Sys.sleep(3)
  }
}


######################################################

ultimate =  read_excel("leftOver.xlsx")


for(i in 269:length(ultimate[[2]])){
  if()
    tryCatch({
      lang_trend <- gtrends(query= ultimate[[2]][[i]], geo="US", start_date =  as.Date("2004-01-1"),end_date = as.Date("2016-9-1"))
      print(i)
      print(str(lang_trend[[3]]))
      name<- paste(ultimate[[2]][[i]],".csv",sep="")
      write.csv(lang_trend[[3]], file = name)
      
    }, error=function(e){
      print(ultimate[[2]][[i]])
      print("not found")
      print(i)
      write(ultimate[[2]][[i]], file = "notFound", append = TRUE)
      write("\n", file = "notFound", append = TRUE)
    })
  Sys.sleep(7)
}

length(finalData)

minus <- "-used-parts-recalls-repair-recalls"




queryName <- "acura ilx+ilx+ilx acura+acura ilx review-used"

queryName <- paste(queryName,minus,sep="")

tryCatch({
  lang_trend <- gtrends(queryName, geo="US", start_date = as.Date("2004-01-1"),end_date = as.Date("2016-9-1"))
  print(str(lang_trend[[3]]))
  name<- paste("trendstry",".csv",sep="")
  write.csv(lang_trend[[3]], file = name)
}, error=function(e){
  print("not found")
})

tryCatch({
  lang_trend <- gtrends("hi", geo="US", start_date = as.Date("2004-01-1"),end_date = as.Date("2016-9-1"))
  print(str(lang_trend[[3]]))
  name<- paste("trendstry",".csv",sep="")
  write.csv(lang_trend[[3]], file = name)
}, error=function(e){
  print("not found")
})
############################
times = 0
ultimate <- read_excel("usage2.xlsx")
for(i in 1:3656){
  if(ultimate[[4]][[i]] == "Use"){
    times = times + 1
  }
  
}
