##################################################################################################
#  Script générant les plot drc des relations dose réponse à partir des lignes
#  de codes collées dans les feuilles Excel.
# Pour fonctionner la commande library(drc) doit être présente, ainsi que summary
# Les feuilles Excel ne doivent pas contenir de bordure

##################################################################################################


library(xlsx)
library(drc)
library(openxlsx)
#library(gdata)



### Function commande plot search : return the line used to make plot of dose reponse data  
###
commande_plot_search <- function (sheet)
{
  beg <- grep("library",sheet [,1]) # recherche de la première ligne d'intérêt
  end <- grep("summary",sheet [,1])[1]

  comm_tmp <- sheet[c((beg+1):(end-1)),1]
  #comm_vect <- comm_tmp[-which(is.na(comm_tmp))]
  nb_na_row <- sum(is.na(comm_tmp))
  if(nb_na_row==0){
      comm_vect <-  comm_tmp
      } else {
         comm_tmp[-which(is.na(comm_tmp))]
      }
}




improve_plot <- function(comm_vect){
 
  comm_char <- as.character(comm_vect)
  ind <- grep("plot",comm_vect)
  string_tmp <- as.character(comm_vect[ind])
  comm_char[ind] <- paste(substr(string_tmp, 1, nchar(string_tmp)-1), ",cex=1.5, cex.axis=1.5,cex.lab=1.5, col=1, pch=19, xlab=my_xlab, ylab=my_ylab)")
  return(comm_char)
  }




commande_plot_execute <- function(comm_char, sheet_name)
{
  for (i in 1:length(comm_char))
    {
    jpeg(paste0('.\\plots\\', sheet_name,".jpg"),width = 480, height = 480, units = "px", pointsize = 12,
         quality = 500)
      eval(parse(text=comm_char[i]))
      dev.off()
  }
}
  


################ 
#EXECUTION
#  Certaines paramètres sont à modifier
################# 

filename <- "RawObsChronicFish.xlsx" # à modifier
#wb <- read.xlsx(filename, sheet=1, startRow = 1, colNames=TRUE)
names_ws <- getSheetNames(filename)

# wb <- loadWorkbook(filename)
# ws <- getSheets(wb)
#names_ws <- names(ws)                    

ind_start <- 3 # à modifier

my_ylab<- "Response (in effect unit)"
my_xlab<-"Dose (Gy)" # à modifier

for (name in names_ws[ind_start:length(names_ws)]){
  mysheet <- openxlsx::read.xlsx(filename,sheet = name ) # besoin d'enlever les bordures dans la feuille Excel
  my_comm <- commande_plot_search(mysheet)
  new_comm <- improve_plot(my_comm)
  commande_plot_execute(new_comm, name)
}





