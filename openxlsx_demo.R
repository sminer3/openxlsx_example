library(openxlsx)

#-----------------------------------
#Create Workbook
#-----------------------------------
wb = createWorkbook()
addWorksheet(wb, 'Summary', tabColour = 'green')
addWorksheet(wb, 'Iris Dataset', tabColour = 'blue')
openXL(wb)

#Create Reusable styles
TITLE_STYLE = createStyle(fontSize = 16, fontColour = 'blue', textDecoration = c('bold', 'underline'))

HEADER_STYLE = createStyle(halign = 'CENTER', textDecoration = c('bold'), border = 'bottom', borderStyle = getOption("openxlsx.borderStyle", "thick"), borderColour = 'black')

#-----------------------------------
#Add titles
#-----------------------------------
writeData(wb,
          sheet = 'Summary',
          x = 'Summary of Iris Dataset',
          startCol = 1,
          startRow = 1)
addStyle(wb,
         sheet = 'Summary',
         style = TITLE_STYLE,
         rows = 1,
         cols = 1)

writeData(wb,
          sheet = 'Iris Dataset',
          x = 'Iris Dataset (Raw data)',
          startCol = 1,
          startRow = 1)
addStyle(wb,
         sheet = 'Iris Dataset',
         style = TITLE_STYLE,
         rows = 1,
         cols = 1)
openXL(wb)

#-----------------------------------
#Add dataframes
#-----------------------------------
iris_summary = data.frame(Column = names(iris)[1:4],
                          Mean = sapply(iris[,1:4], mean),
                          Std.Err = sapply(iris[,1:4], sd))

writeData(wb, 
          sheet = 'Summary', 
          x = iris_summary, 
          xy = c(2, 3), 
          headerStyle = HEADER_STYLE)

writeDataTable(wb,
               sheet = 'Iris Dataset',
               x = iris,
               xy = c(2,3),
               headerStyle = HEADER_STYLE,
               tableStyle = "TableStyleLight9")
openXL(wb)

#-----------------------------------
#Formatting
#-----------------------------------
setColWidths(wb,
             sheet = 'Summary',
             cols = 2:4,
             widths = 'auto')
setColWidths(wb,
             sheet = 'Iris Dataset',
             cols = 2:6,
             widths = 'auto')
addStyle(wb,
         sheet = 'Summary',
         style = createStyle(textDecoration = 'bold'),
         cols = 2,
         rows = 4:7,
         stack = T)
openXL(wb)

#-----------------------------------
#Insert Plot
#-----------------------------------
library(tidyverse)
library(gridExtra)

sl_plot = ggplot(iris, aes(x = Species, y = Sepal.Length)) + geom_boxplot()
sw_plot = ggplot(iris, aes(x = Species, y = Sepal.Width)) + geom_boxplot()
pl_plot = ggplot(iris, aes(x = Species, y = Petal.Length)) + geom_boxplot()
pw_plot = ggplot(iris, aes(x = Species, y = Petal.Width)) + geom_boxplot()

png("summary_plot.png", height=1200, width=1200, res=250, pointsize=40)
  grid.arrange(sl_plot, sw_plot, pl_plot, pw_plot, ncol = 2, nrow = 2)
dev.off()

insertImage(wb,
           sheet = 'Summary',
           file = "summary_plot.png",
           startRow = 3,
           startCol = 7,
           heigh = 6,
           width = 6)
openXL(wb)





