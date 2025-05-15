# Read me

# Data file format specifications
# You must provide a xlsx file with one or more sheets
# The sheets names must be different from each other
# In each sheet:
#   1. The first row must contains the names of two drugs in the first two cells
#   2. Drug name order: row drug, column drug
#   3. The second row must be blank
#   4. If two or more data matrices : always separated by a blank row

# The name of figure will be composed by:
#   1. the name of the corresponding sheet
#   2. the rank of the corresponding matrix in the sheet

# Create a folder named "output" in the wd, in order to collect the figure files
# Modify the settings in the corresponding section here below and it's done


######################################################################################################
# Changes since the previous version: 
# - save the index as .xlsx


## set wd
wd <- "C:/users/jean-dupont/Documents/" 
filename <- paste0(wd, "matrix-R.xlsx") 
output <- paste0(wd, "output/")
setwd(output)

## loading libraries
library(readxl)
library(gplots)
library(dplyr)
library(gtools)
library(gridExtra)
library(gridGraphics)
library(ggplot2)
library(ggpubr)
library(plot3D)
library(reshape2)
library(openxlsx)


sheets <- excel_sheets(filename)
for (sheet_n in 1:length(sheets)) {
  
  excel <- read_excel(filename, sheet = sheets[sheet_n], col_names = FALSE)
  sep <- c(which(rowSums(!is.na(excel)) == 0), nrow(excel)+1)
  
  # initialisation
  matrices.data <- matrices.Diff <- index <- c()
  
  for (sep_n in 1:(length(sep)-1)) {
    
    ## data formating
    data <- as.matrix(excel[(sep[sep_n]+1):(sep[sep_n+1]-1),])
    class(data) <- "numeric"
    dimnames(data) <- list(data[,1], data[1,])
    data <- round(data[-1,-1], 0)
    data[which(data > 100)] <- 100
    
    drugs <- list(drugA = list(name = as.character(excel[1,1]), dose = as.numeric(rownames(data))),
                  drugB = list(name = as.character(excel[1,2]), dose = as.numeric(colnames(data))))
    
    drugs$drugA$name <- paste(drugs$drugA$name, "\n")
    drugs$drugB$name <- paste("\n", drugs$drugB$name)
    
    ## Bliss additive effect estimation
    fua <- data[1,]
    fub <- data[,1]
    
    fu <- vector()
    for(a in fua){
      for(b in fub){
        fu <- append(fu, c(a, b))
      }
    } 
    
    fu <- matrix(fu, ncol = 2, byrow = T, dimnames = list(c(), c("a", "b")))/100
    Bliss <- apply(fu, 1, prod) * 100
    Bliss <- matrix(Bliss, dim(data), dimnames = dimnames(data))
    
    
    ## difference matrix 
    Diff <- round(Bliss - data, 1)
    
    
    ## calculation of syntetic indexes according to Lehar's method
    ## (combination + additivity + efficacy indexes)
    
    dfa <- tail(drugs$drugA$dose, 1)/tail(drugs$drugA$dose, 2)[1]
    dfb <- tail(drugs$drugB$dose, 1)/tail(drugs$drugB$dose, 2)[1]
    
    AI <- log(dfa) * log(dfb) * sum(100 - data - Diff, na.rm = TRUE) / 100
    CI <- log(dfa) * log(dfb) * sum(Diff, na.rm = TRUE) / 100
    #CI <- mean(Diff[-1,-1])
    EI <- log(dfa) * log(dfb) * sum(100 - data[-1,-1], na.rm = TRUE) / 100
    #EI <- mean(100 - data[-1,-1])
    index <- rbind(index, c(CI, EI, AI))
    
    
    ## plots 
    colbreaks <- seq(-100, 100, 20)
    
    # plot data 
    data.plot <- as.data.frame(data)
    data.plot$doseA <- rownames(data.plot)
    data.plot <- reshape2::melt(data.plot, id.vars = "doseA")
    
    plot.data <- ggplot(data.plot, aes(x = variable, y = doseA, fill = value)) +
      geom_tile() +
      scale_y_discrete(position = "left", limits = rev(as.character(drugs$drugA$dose))) +
      scale_x_discrete(position = "bottom", limits = as.character(drugs$drugB$dose)) +
      scale_fill_gradient(low = "dodgerblue1", high = "navy", 
                          limits = c(0, 100),
                          breaks = c(0, 50, 100)) +
      labs(x = drugs$drugB$name, y = drugs$drugA$name) +
      geom_text(aes(label = round(value)), color = "white") +
      theme_minimal() +
      theme(axis.title = element_text(size = 14),
            axis.text = element_text(size = 12))
    
    matrices.data <- c(matrices.data, list(plot.data))
    
    
    # plot Bliss
    Bliss.plot <- as.data.frame(Bliss)
    Bliss.plot$doseA <- rownames(Bliss.plot)
    Bliss.plot <- reshape2::melt(Bliss.plot, id.vars = "doseA")
    
    plot.Bliss <- ggplot(Bliss.plot, aes(x = variable, y = doseA, fill = value)) +
      geom_tile() +
      scale_y_discrete(position = "left", limits = rev(as.character(drugs$drugA$dose))) +
      scale_x_discrete(position = "bottom", limits = as.character(drugs$drugB$dose)) +
      scale_fill_gradient(low = "dodgerblue1", high = "navy", 
                          limits = c(0, 100),
                          breaks = c(0, 50, 100)) +
      labs(x = drugs$drugB$name, y = drugs$drugA$name) +
      geom_text(aes(label = round(value)), color = "white") +
      theme_minimal() +
      theme(axis.title = element_text(size = 14),
            axis.text = element_text(size = 12))
    
    
    # plot Diff
    Diff.plot <- as.data.frame(Diff)
    Diff.plot$doseA <- rownames(Diff.plot)
    Diff.plot <- reshape2::melt(Diff.plot, id.vars = "doseA")
    
    plot.Diff <- ggplot(Diff.plot, aes(x = variable, y = doseA, fill = value)) +
      geom_tile() +
      scale_y_discrete(position = "left", limits = rev(as.character(drugs$drugA$dose))) +
      scale_x_discrete(position = "bottom", limits = as.character(drugs$drugB$dose)) +
      scale_fill_gradientn(colors = c("#00FF00", "#004e00", "#000000", "#000000", "#000000", "#4e0000", "#FF0000"),
                           values = c(0, 0.425, 0.426, 0.5, 0.575, 0.576, 1),
                           limits = c(colbreaks[1], colbreaks[11]),
                           breaks = c(colbreaks[1], 0, colbreaks[11])) +
      labs(x = drugs$drugB$name, y = drugs$drugA$name) +
      geom_text(aes(label = round(value)), color = "white") +
      theme_minimal() +
      theme(axis.title = element_text(size = 14),
            axis.text = element_text(size = 12))
    
    matrices.Diff <- c(matrices.Diff, list(plot.Diff))
    
    
    # PDF manip
    pdf(paste0(output, sheets[sheet_n], "_", sep_n, ".pdf"))
    print(plot.data)
    print(plot.Bliss)
    print(plot.Diff)
    
    x <- c(1:length(drugs$drugA$dose))
    y <- c(1:length(drugs$drugB$dose))
    grid <- mesh(x, y)
    hist3D(x, y, 100 - data, border = "black", xlab = drugs$drugB$name, ylab = drugs$drugA$name, zlab = "effect, %")
    hist3D(x, y, Diff, border = "black", xlab = drugs$drugB$name, ylab = drugs$drugA$name, zlab = "effect, %")
    dev.off()
  }
  
  ## save index table
  index <- as.data.frame(index)
  colnames(index) <- c("CI", "EI", "AI")
  # write.csv2(index, paste0(sheets[sheet_n], "_index.csv"), row.names = FALSE, quote = FALSE)
  write.xlsx(index, file = paste0(sheets[sheet_n], "_index.xlsx"), row.names = FALSE)
  
  
  ## PDF matrices 
  row1 <- matrices.data[[1]]
  if (length(matrices.data) > 1) {
    for (i in 2:length(matrices.data)) {
      len <- 1/i
      if (i < length(matrices.data)) row1 <- ggarrange(row1, matrices.data[[i]], widths = c(1-len, len), legend = "none")
      if (i == length(matrices.data)) row1 <- ggarrange(row1, matrices.data[[i]], widths = c(1-len, len), legend = "right", common.legend = TRUE)
    }
  }
  
  row2 <- matrices.Diff[[1]]
  if (length(matrices.Diff) > 1) {
    for (i in 2:length(matrices.Diff)) {
      len <-  1/i
      if (i < length(matrices.Diff)) row2 <- ggarrange(row2, matrices.Diff[[i]], widths = c(1-len, len), legend = "none")
      if (i == length(matrices.Diff)) row2 <- ggarrange(row2, matrices.Diff[[i]], widths = c(1-len, len), legend = "right", common.legend = TRUE)
    }
  }
  
  p <- ggarrange(row1, NULL, row2, nrow = 3, heights = c(1, 0.05, 1))
  width <- length(drugs$drugB$dose) * length(matrices.data)
  height <- length(drugs$drugA$dose) * 2
  ggsave(paste0(output, sheets[sheet_n], "_matrices.pdf"), plot = p, width = width, height = height, units = "in")
}

