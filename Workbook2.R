if(!require(rvest)) install.packages("rvest", quiet=TRUE); library(rvest, quietly=TRUE)
if(!require(openxlsx)) install.packages("openxlsx", quiet=TRUE); library(openxlsx, quietly=TRUE)
if(!require(ggplot2)) install.packages("ggplot2", quiet=TRUE); library(ggplot2, quietly=TRUE)
if(!require(tidyverse)) install.packages("tidyverse", quiet=TRUE); library(tidyverse, quietly=TRUE)
if(!require(dslabs)) install.packages("dslabs", quiet=TRUE); library(dslabs, quietly=TRUE)
if(!require (ggpubr)) install.packages("ggpubr",  quiet=TRUE) ; library("ggpubr")

setwd("C:/Users/Artur/Desktop/Batch")


#-------------------------------------------------------------
#download and wrangle some real data
#-------------------------------------------------------------
url <- "https://edgar.jrc.ec.europa.eu/overview.php?v=booklet2019&dst=CO2emi"
page1 <- read_html(url)
nodes <- html_nodes(page1, "table")
columns <- c("Country", "1990", "2000", "2005", "2010", "2015", "2017", "2018", "Proportion2018")
provisionally_table <- data.frame(html_table(nodes[1], fill=TRUE))
colnames(provisionally_table) <- columns
provisionally_table <- provisionally_table[-c(1,2),]

provisionally_table$Proportion2018 <- as.numeric(provisionally_table$Proportion2018)
true_table <- provisionally_table%>%
  arrange(desc(Proportion2018))%>%
  head(7)

final_table <- true_table%>%
  gather(Year, Emissions_in_tons, -Country)%>%
  filter(Year!="Proportion2018", Country!="GLOBAL TOTAL")

write.xlsx(final_table, "Emissions.xlsx", overwrite=TRUE)



#-------------------------------------------------------------
#create some other data
#-------------------------------------------------------------
data(CO2)
write.xlsx(CO2, file="CO2.xlsx", overwrite=TRUE)
data(mtcars)
write.xlsx(mtcars, file="cars.xlsx", overwirte=TRUE)



#-------------------------------------------------------------
#create the workbook, its sheets, fill it with data and plots
#-------------------------------------------------------------
path <- "C:/Users/Artur/Desktop/Batch"  #create a path name

if (file.exists("Presentation.xlsx")) {
  file.remove("Presentation.xlsx")
}

filenames <- list.files(path=path, pattern="*.xlsx")  # list all XLSX-files
names <- str_replace(filenames, ".xlsx", "")    # proper names for unknown sheets

hs1 <- createStyle(fgFill= "#c1cefa", halign="CENTER", textDecoration="bold", fontSize= 11, fontName="TimesNewRoman",
                   border= "TopBottomLeftRight")

hs2 <- createStyle(fgFill=  "#faedc1", halign="CENTER", textDecoration="bold",fontSize= 11, fontName="TimesNewRoman",
                   border= "TopBottomLeftRight ")

hs3 <- createStyle(fgFill= "#a1ffc8", halign="CENTER", textDecoration="bold", fontSize= 11, fontName="Calibri",
                   border= "TopBottomLeftRight")

wb <- openxlsx::loadWorkbook("Agenda.xlsx")
for (i in 1:length(filenames)){
  if (filenames[i]=="Agenda.xlsx"){
    next
  } else if (filenames[i] == "Emissions.xlsx") {
  addWorksheet(wb, "Global Emission Contributors")
  GEC <- read.xlsx("Emissions.xlsx")
  writeData(wb, "Global Emission Contributors", x=GEC, borders="columns", headerStyle = hs2)
  setColWidths(wb, "Global Emission Contributors", cols=1:3, widths=c(13,10,17))
  GEC_plot <- GEC%>%ggplot(aes(as.numeric(Year), as.numeric(Emissions_in_tons), col=Country)) + 
    geom_line(size=1.3) + 
    ggtitle("Alarming role of China in Global Warming") + 
    xlab("Years") + 
    ylab("Emissions in metric Mtons") + 
    theme(plot.title = element_text(hjust=0.5, vjust=0.1))
  print(GEC_plot)
  insertPlot(wb, sheet="Global Emission Contributors", xy=c("F", 1), width = 8, height=6, fileType = "png")
  saveWorkbook(wb, file="Presentation.xlsx", overwrite=T)
  } else if (filenames[i] == "CO2.xlsx") {
  addWorksheet(wb, "Plants - A Silver Lining")
  CO2 <- read.xlsx("CO2.xlsx", colNames = TRUE)
  writeData(wb, sheet="Plants - A Silver Lining", x=CO2, borders="columns", headerStyle = hs3)
  setColWidths(wb, "Plants - A Silver Lining", cols=1:5, widths=10)
  CO2_plot <- CO2%>%
    ggplot+geom_point(aes(x=uptake, y=conc, col=Plant, size=4)) + 
    ggtitle("Diffenre in C02-uptake of Plant Species in Northamerica") + 
    theme(plot.title = element_text(hjust=0.5, vjust=0.1))
  print(CO2_plot)
  insertPlot(wb, sheet="Plants - A Silver Lining", xy=c("G", 1), width=8, height=6, fileType="png")
  saveWorkbook(wb, file="Presentation.xlsx", overwrite=T)
  } else if (filenames[i] == "cars.xlsx") {
    addWorksheet(wb, "Connection Carweigth and Range")
    cars <- read.xlsx("cars.xlsx", colNames = TRUE)
    writeData(wb, sheet="Connection Carweigth and Range", x=cars, borders="columns", headerStyle = hs1)
    setColWidths(wb, "Connection Carweigth and Range", cols=1:11, widths=7)
    cars_plot <- mtcars%>%ggplot(aes(mpg, wt)) + 
      geom_point()+stat_cor(method="pearson", label.x=27.5, label.y=5) + 
      geom_smooth(method=lm)+ggtitle("Drive Smaller Cars: Simple as that!") + 
      theme(plot.title= element_text(hjust=0.5, vjust=0.1))
    print(cars_plot)
    insertPlot(wb, sheet="Connection Carweigth and Range", xy=c("M",1), fileType="png", width=8, height=6)
    saveWorkbook(wb, file="Presentation.xlsx", overwrite=T)
  } else { #for any additional sheets that should not be wrangled
  addWorksheet(wb, sheetName=names[i])
  writeData(wb, sheet=names[i], x=read.xlsx(filenames[i]))
  saveWorkbook(wb, file="Presentation.xlsx", overwrite=T)
  }
}
