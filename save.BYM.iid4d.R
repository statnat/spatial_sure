#' save.BYM.iid4d Function
#'
#' Save INLA BYM model summary
#' @param model the file name of the model
#' @param modelname in quotation marks is the of the file you want the CSV file to be called
#' @param modelname2 a subtitle for the Excel spreadsheet
#' @param outcome the outcome (point to data set and variable name)
#' @param n the number of space-time units
#' @param randomToSD specify as TRUE or FALSE (FALSE is the default). This converts precisions to SDs. Note that this takes some time.
#' @keywords save BYM model results
#' @export
#' @examples
#' save.BYM.iid4d(model=model2, modelname="Model 2", outcome=df$y, expVal=df$E, n=100, randomToSD=FALSE)

## This function is for BYM models with an additional 4d IID non-spatial random effect that takes the 
# correlation among four seemningly unrelated regression equations (i.e., corresponding to a multivariate
# model with four outcomes) into account.

save.BYM.iid4d <- function(model, modelname, modelname2, outcome, expVal, n, randomToSD) {
  library("INLA"); library("xlsx"); library("expss")
  
  ##############################################################################
  # THIS NEEDS TO BE SPECIFIED FOR EACH MODEL.
  ##############################################################################
  if (missing(model)) {print("Please specify a model object.")} else {model = model}
  if (missing(modelname)) {modelname="Model"} else {modelname = modelname}
  if (missing(randomToSD)) {randomToSD=FALSE} else {randomToSD = randomToSD}
  if (missing(modelname2)) {modelname2="Model2"} else {modelname2 = modelname2}
  if (missing(outcome)) {print("Please specify the outcome.")} else {y = outcome} # outcome for MAE, RMSE calculations
  if (missing(n)) {print("Please specify the number of space-time units.")} else {Nareas = n}  # number of space-time units
  if (missing(expVal)) {expVal=rep(1,n)}
  ##############################################################################
  
  ##############################################################################
  # FORMAT R-INLA OUTPUT INTO DATA FRAME
  ##############################################################################
  # these can be modified to change the starting row of different components of the model output.
  startRowTitle=1; startRowSubtitle=2; startRowData=3; startColModel=2
  
  names <- rownames(model$summary.fixed)
  WS <- ifelse(model$summary.fixed$`0.025quant`>0, "High", NA)
  WS <- ifelse(model$summary.fixed$`0.975quant`<0, "Low", WS)
  med <- sprintf("%.4f", exp(model$summary.fixed$`0.5quant`))
  low <- sprintf("%.4f", exp(model$summary.fixed$`0.025quant`))
  high <- sprintf("%.4f", exp(model$summary.fixed$`0.975quant`))
  
  sapply(1:length(low), function(x) if (low[x]=="1.000") {
    low[x]<<- sprintf("%.4f", exp(model$summary.fixed$`0.025quant`)[x])
    med[x]<<- sprintf("%.4f", exp(model$summary.fixed$`0.5quant`)[x])
    high[x]<<- sprintf("%.4f", exp(model$summary.fixed$`0.975quant`)[x])
  })
  sapply(1:length(high), function(x) if (high[x]=="1.000") {
    high[x]<<- sprintf("%.4f", exp(model$summary.fixed$`0.975quant`)[x])
    med[x]<<- sprintf("%.4f", exp(model$summary.fixed$`0.5quant`)[x])
    low[x]<<- sprintf("%.4f", exp(model$summary.fixed$`0.025quant`)[x])
  })
  
  ci <- noquote(paste(med, " [", low, ", ", high, "]", sep=""))
  df <- data.frame(cbind(names, ci))
  colnames(df) <- c("Fixed effects", "Median [95% CI]")
  
  # ADD/MODIFY MANUALLY AS NEEDED. THESE ARE THE STANDARD VARIABLES USED FOR THE spatial_sure PROJECT.
  levels(df$`Fixed effects`) <- sub("^age1$", "Ages 0-19, %", levels(df$`Fixed effects`))
  levels(df$`Fixed effects`) <- sub("^age2$", "Ages 20-24, %", levels(df$`Fixed effects`))
  levels(df$`Fixed effects`) <- sub("^age3$", "Ages 25-44, %", levels(df$`Fixed effects`))
  levels(df$`Fixed effects`) <- sub("^age4$", "Ages 45-64, %", levels(df$`Fixed effects`))
  levels(df$`Fixed effects`) <- sub("^hosprate$", "Hospitalization rate", levels(df$`Fixed effects`))
  levels(df$`Fixed effects`) <- sub("^arth$", "Arthritis hosp.", levels(df$`Fixed effects`))
  levels(df$`Fixed effects`) <- sub("^retail$", "Retail density", levels(df$`Fixed effects`))
  levels(df$`Fixed effects`) <- sub("^retailr$", "Retail density", levels(df$`Fixed effects`))
  levels(df$`Fixed effects`) <- sub("^manual$", "Manual labor density", levels(df$`Fixed effects`))
  levels(df$`Fixed effects`) <- sub("^cancer$", "Cancer hosp.", levels(df$`Fixed effects`))
  levels(df$`Fixed effects`) <- sub("^year2$", "Year", levels(df$`Fixed effects`))
  levels(df$`Fixed effects`) <- sub("^unemprate$", "Unemployment rate", levels(df$`Fixed effects`))
  levels(df$`Fixed effects`) <- sub("^hisp$", "Hispanic, %", levels(df$`Fixed effects`))
  levels(df$`Fixed effects`) <- sub("^medinc$", "Median household income", levels(df$`Fixed effects`))
  levels(df$`Fixed effects`) <- sub("factor\\(year\\)", "Year ", levels(df$`Fixed effects`))
  levels(df$`Fixed effects`) <- sub("\\(Intercept\\)", "Intercept", levels(df$`Fixed effects`))
  levels(df$`Fixed effects`) <- sub("factor\\(catdens\\)", "Density ", levels(df$`Fixed effects`))
  levels(df$`Fixed effects`) <- sub("dens", "Density ", levels(df$`Fixed effects`))
  levels(df$`Fixed effects`) <- sub("white", "White, %", levels(df$`Fixed effects`))
  levels(df$`Fixed effects`) <- sub("black", "Black, %", levels(df$`Fixed effects`))
  levels(df$`Fixed effects`) <- sub("^manualr$", "Manual labor density", levels(df$`Fixed effects`))
  levels(df$`Fixed effects`) <- sub("males", "Male, %", levels(df$`Fixed effects`))
  levels(df$`Fixed effects`) <- sub("bl150", "Poverty (below 150%), %", levels(df$`Fixed effects`))
  levels(df$`Fixed effects`) <- sub("instab", "Geographic instability", levels(df$`Fixed effects`))
  
  ###########################################################
  # convert precisions to SDs
  ###########################################################
  numcols <- 100000
  if (randomToSD==TRUE) {
    
    mat.marg1 <- matrix(NA, nrow=Nareas, ncol=100000)
    mat.marg2 <- matrix(NA, nrow=Nareas, ncol=100000)
    mat.marg3 <- matrix(NA, nrow=Nareas, ncol=100000)
    mat.marg4 <- matrix(NA, nrow=Nareas, ncol=100000)
    
    # m takes on the value of the a list of Nareas*2 lists of 75 objects with x,y values on each row
    # in which there is a row around the middle giving the x value that maximizes y, which seems like a likelihood
    # This loops through the second half of the Nareas*2 lists, which contain the CAR posteriors "u" only
    m1 <- model$marginals.random$bgID1
    m2 <- model$marginals.random$bgID2
    m3 <- model$marginals.random$bgID3
    m4 <- model$marginals.random$bgID4
    
    
    for (i in 1:Nareas){
      u1 <- m1[[i]]
      u2 <- m2[[i]]
      u3 <- m3[[i]]
      u4 <- m4[[i]]
      
      mat.marg1[i,] <- inla.rmarginal(100000, u1)
      mat.marg2[i,] <- inla.rmarginal(100000, u2)
      mat.marg3[i,] <- inla.rmarginal(100000, u3)
      mat.marg4[i,] <- inla.rmarginal(100000, u4)
    }
    
    
    # Calculate 100,000 estimates of variance of u (CAR random intercepts) and v (noise random intercept)
    var.u1 <- apply(mat.marg1, 2, var)
    var.u2 <- apply(mat.marg2, 2, var)
    var.u3 <- apply(mat.marg3, 2, var)
    var.u4 <- apply(mat.marg4, 2, var)
    
    # Calculate 100,000 estimates of variance of v (iid, non-spatial random intercepts)
    var.v1 <- inla.rmarginal(100000,inla.tmarginal(function(x) 1/x,
                                                   model$marginals.hyperpar$`Precision for ID (component 1)`))
    var.v2 <- inla.rmarginal(100000,inla.tmarginal(function(x) 1/x,
                                                   model$marginals.hyperpar$`Precision for ID (component 2)`))
    var.v3 <- inla.rmarginal(100000,inla.tmarginal(function(x) 1/x,
                                                   model$marginals.hyperpar$`Precision for ID (component 3)`))
    var.v4 <- inla.rmarginal(100000,inla.tmarginal(function(x) 1/x,
                                                   model$marginals.hyperpar$`Precision for ID (component 4)`))
    
    
    # report quantiles
    u1<-noquote(paste(sprintf("%.4f", (quantile(sqrt(var.u1[var.u1>=0]), c(0.5)))), " [",
                      sprintf("%.4f", (quantile(sqrt(var.u1[var.u1>=0]), c(0.025)))), ", ",
                      sprintf("%.4f", (quantile(sqrt(var.u1[var.u1>=0]), c(0.975)))), "]", sep=""))
    u2<-noquote(paste(sprintf("%.4f", (quantile(sqrt(var.u2[var.u2>=0]), c(0.5)))), " [",
                      sprintf("%.4f", (quantile(sqrt(var.u2[var.u2>=0]), c(0.025)))), ", ",
                      sprintf("%.4f", (quantile(sqrt(var.u2[var.u2>=0]), c(0.975)))), "]", sep=""))
    u3<-noquote(paste(sprintf("%.4f", (quantile(sqrt(var.u3[var.u3>=0]), c(0.5)))), " [",
                      sprintf("%.4f", (quantile(sqrt(var.u3[var.u3>=0]), c(0.025)))), ", ",
                      sprintf("%.4f", (quantile(sqrt(var.u3[var.u3>=0]), c(0.975)))), "]", sep=""))
    u4<-noquote(paste(sprintf("%.4f", (quantile(sqrt(var.u4[var.u4>=0]), c(0.5)))), " [",
                      sprintf("%.4f", (quantile(sqrt(var.u4[var.u4>=0]), c(0.025)))), ", ",
                      sprintf("%.4f", (quantile(sqrt(var.u4[var.u4>=0]), c(0.975)))), "]", sep=""))
    v1<-noquote(paste(sprintf("%.4f", (quantile(sqrt(var.v1[var.v1>=0]), c(0.5)))), " [",
                      sprintf("%.4f", (quantile(sqrt(var.v1[var.v1>=0]), c(0.025)))), ", ",
                      sprintf("%.4f", (quantile(sqrt(var.v1[var.v1>=0]), c(0.975)))), "]", sep=""))
    v2<-noquote(paste(sprintf("%.4f", (quantile(sqrt(var.v2[var.v2>=0]), c(0.5)))), " [",
                      sprintf("%.4f", (quantile(sqrt(var.v2[var.v2>=0]), c(0.025)))), ", ",
                      sprintf("%.4f", (quantile(sqrt(var.v3[var.v2>=0]), c(0.975)))), "]", sep=""))
    v3<-noquote(paste(sprintf("%.4f", (quantile(sqrt(var.v3[var.v3>=0]), c(0.5)))), " [",
                      sprintf("%.4f", (quantile(sqrt(var.v3[var.v3>=0]), c(0.025)))), ", ",
                      sprintf("%.4f", (quantile(sqrt(var.v3[var.v3>=0]), c(0.975)))), "]", sep=""))
    v4<-noquote(paste(sprintf("%.4f", (quantile(sqrt(var.v4[var.v4>=0]), c(0.5)))), " [",
                      sprintf("%.4f", (quantile(sqrt(var.v4[var.v4>=0]), c(0.025)))), ", ",
                      sprintf("%.4f", (quantile(sqrt(var.v4[var.v4>=0]), c(0.975)))), "]", sep=""))
    
    a1<-noquote(paste(sprintf("%.4f", (quantile(var.u1/(var.u1+var.v1), c(0.5)))), " [",
                      sprintf("%.4f", (quantile(var.u1/(var.u1+var.v1), c(0.025)))), ", ",
                      sprintf("%.4f", (quantile(var.u1/(var.u1+var.v1), c(0.975)))), "]", sep=""))
    a2<-noquote(paste(sprintf("%.4f", (quantile(var.u2/(var.u2+var.v2), c(0.5)))), " [",
                      sprintf("%.4f", (quantile(var.u2/(var.u2+var.v2), c(0.025)))), ", ",
                      sprintf("%.4f", (quantile(var.u2/(var.u2+var.v2), c(0.975)))), "]", sep=""))
    a3<-noquote(paste(sprintf("%.4f", (quantile(var.u3/(var.u3+var.v3), c(0.5)))), " [",
                      sprintf("%.4f", (quantile(var.u3/(var.u3+var.v3), c(0.025)))), ", ",
                      sprintf("%.4f", (quantile(var.u3/(var.u3+var.v3), c(0.975)))), "]", sep=""))
    a4<-noquote(paste(sprintf("%.4f", (quantile(var.u4/(var.u4+var.v4), c(0.5)))), " [",
                      sprintf("%.4f", (quantile(var.u4/(var.u4+var.v4), c(0.025)))), ", ",
                      sprintf("%.4f", (quantile(var.u4/(var.u4+var.v4), c(0.975)))), "]", sep=""))
    
    
    
    
    
    rho12Med <- model$summary.hyperpar$`0.5quant`[9]
    rho12Lo <- model$summary.hyperpar$`0.025quant`[9]
    rho12Hi <- model$summary.hyperpar$`0.975quant`[9]
    rho13Med <- model$summary.hyperpar$`0.5quant`[10]
    rho13Lo <- model$summary.hyperpar$`0.025quant`[10]
    rho13Hi <- model$summary.hyperpar$`0.975quant`[10]
    rho14Med <- model$summary.hyperpar$`0.5quant`[11]
    rho14Lo <- model$summary.hyperpar$`0.025quant`[11]
    rho14Hi <- model$summary.hyperpar$`0.975quant`[11]
    rho23Med <- model$summary.hyperpar$`0.5quant`[12]
    rho23Lo <- model$summary.hyperpar$`0.025quant`[12]
    rho23Hi <- model$summary.hyperpar$`0.975quant`[12]
    rho24Med <- model$summary.hyperpar$`0.5quant`[13]
    rho24Lo <- model$summary.hyperpar$`0.025quant`[13]
    rho24Hi <- model$summary.hyperpar$`0.975quant`[13]
    rho34Med <- model$summary.hyperpar$`0.5quant`[14]
    rho34Lo <- model$summary.hyperpar$`0.025quant`[14]
    rho34Hi <- model$summary.hyperpar$`0.975quant`[14]
    rho12 <- paste(sprintf("%.4f", rho12Med), " [", sprintf("%.4f", rho12Lo), ", ", sprintf("%.4f", rho12Hi), "]", sep="")
    rho13 <- paste(sprintf("%.4f", rho13Med), " [", sprintf("%.4f", rho13Lo), ", ", sprintf("%.4f", rho13Hi), "]", sep="")
    rho14 <- paste(sprintf("%.4f", rho14Med), " [", sprintf("%.4f", rho14Lo), ", ", sprintf("%.4f", rho14Hi), "]", sep="")
    rho23 <- paste(sprintf("%.4f", rho23Med), " [", sprintf("%.4f", rho23Lo), ", ", sprintf("%.4f", rho23Hi), "]", sep="")
    rho24 <- paste(sprintf("%.4f", rho24Med), " [", sprintf("%.4f", rho24Lo), ", ", sprintf("%.4f", rho24Hi), "]", sep="")
    rho34 <- paste(sprintf("%.4f", rho34Med), " [", sprintf("%.4f", rho34Lo), ", ", sprintf("%.4f", rho34Hi), "]", sep="")
    
    rand <<- data.frame(noquote(rbind(u1,u2,u3,u4,v1,v2,v3,v4,a1,a2,a3,a4, rho12, rho13, rho14, rho23, rho24, rho34)))
    namesRandom <<- c("var of bgID1","var of bgID2","var of bgID3","var of bgID4",
                      "iid1","iid2","iid3","iid4","alpha1","alpha2","alpha3","alpha4",
                      "Rho 1:2","Rho 1:3","Rho 1:4","Rho 2:3","Rho 2:4","Rho 3:4")
    
    dfRandom <<- data.frame(cbind(namesRandom, rand))
    colnames(dfRandom) <<- c("Random effects", "Median [95% CI]")
    
  }
  
  else if (randomToSD==FALSE) {
    bg1Med <- model$summary.hyperpar$`0.5quant`[1]
    bg1Lo <- model$summary.hyperpar$`0.025quant`[1]
    bg1Hi <- model$summary.hyperpar$`0.975quant`[1]
    bg2Med <- model$summary.hyperpar$`0.5quant`[2]
    bg2Lo <- model$summary.hyperpar$`0.025quant`[2]
    bg2Hi <- model$summary.hyperpar$`0.975quant`[2]
    rho12Med <- model$summary.hyperpar$`0.5quant`[9]
    rho12Lo <- model$summary.hyperpar$`0.025quant`[9]
    rho12Hi <- model$summary.hyperpar$`0.975quant`[9]
    rho13Med <- model$summary.hyperpar$`0.5quant`[10]
    rho13Lo <- model$summary.hyperpar$`0.025quant`[10]
    rho13Hi <- model$summary.hyperpar$`0.975quant`[10]
    rho14Med <- model$summary.hyperpar$`0.5quant`[11]
    rho14Lo <- model$summary.hyperpar$`0.025quant`[11]
    rho14Hi <- model$summary.hyperpar$`0.975quant`[11]
    rho23Med <- model$summary.hyperpar$`0.5quant`[12]
    rho23Lo <- model$summary.hyperpar$`0.025quant`[12]
    rho23Hi <- model$summary.hyperpar$`0.975quant`[12]
    rho24Med <- model$summary.hyperpar$`0.5quant`[13]
    rho24Lo <- model$summary.hyperpar$`0.025quant`[13]
    rho24Hi <- model$summary.hyperpar$`0.975quant`[13]
    rho34Med <- model$summary.hyperpar$`0.5quant`[14]
    rho34Lo <- model$summary.hyperpar$`0.025quant`[14]
    rho34Hi <- model$summary.hyperpar$`0.975quant`[14]
    bg1CI <- paste(sprintf("%.4f", bg1Med), " [", sprintf("%.4f", bg1Lo), ", ", sprintf("%.4f", bg1Hi), "]", sep="")
    bg2CI <- paste(sprintf("%.4f", bg2Med), " [", sprintf("%.4f", bg2Lo), ", ", sprintf("%.4f", bg2Hi), "]", sep="")
    rho12 <- paste(sprintf("%.4f", rho12Med), " [", sprintf("%.4f", rho12Lo), ", ", sprintf("%.4f", rho12Hi), "]", sep="")
    rho13 <- paste(sprintf("%.4f", rho13Med), " [", sprintf("%.4f", rho13Lo), ", ", sprintf("%.4f", rho13Hi), "]", sep="")
    rho14 <- paste(sprintf("%.4f", rho14Med), " [", sprintf("%.4f", rho14Lo), ", ", sprintf("%.4f", rho14Hi), "]", sep="")
    rho23 <- paste(sprintf("%.4f", rho23Med), " [", sprintf("%.4f", rho23Lo), ", ", sprintf("%.4f", rho23Hi), "]", sep="")
    rho24 <- paste(sprintf("%.4f", rho24Med), " [", sprintf("%.4f", rho24Lo), ", ", sprintf("%.4f", rho24Hi), "]", sep="")
    rho34 <- paste(sprintf("%.4f", rho34Med), " [", sprintf("%.4f", rho34Lo), ", ", sprintf("%.4f", rho34Hi), "]", sep="")
    rand <<- data.frame(noquote(rbind(bg1CI, bg2CI, rho12, rho13, rho14, rho23, rho24, rho34)))
    namesRandom <<- c("Precision of non-spatial bgID1", "Precision of non-spatial bgID2", 
                      "Rho 1:2","Rho 1:3","Rho 1:4","Rho 2:3","Rho 2:4","Rho 3:4")
    
    dfRandom <<- data.frame(cbind(namesRandom, rand))
    colnames(dfRandom) <<- c("Random effects", "Median [95% CI]")
  }
  
  ### Diagnostics
  # DIC
  dic <- sprintf("%.2f", model$dic$dic)
  dev <- sprintf("%.2f", model$dic$mean.deviance)
  pd <- sprintf("%.2f", model$dic$p.eff)
  
  # RMSE,MAE
  mae <- function(x1, x2) { mean(abs(x1-x2)) }
  maeVal <- sprintf("%.4f",mae(outcome[1:n], expVal[1:n]*model$summary.fitted.values$`0.5quant`[1:n]))
  
  rmse <- function(x1, x2) { sqrt(mean((x1-x2)^2)) }
  rmseVal <- sprintf("%.4f", rmse(outcome[1:n], expVal[1:n]*model$summary.fitted.values$`0.5quant`[1:n]))
  
  #  diagNames <- c("Deviance", "Effective number of parameters", "DIC", "MAE", "RMSE")
  # just want deviance and DIC
  diagNames <- c("Deviance", "DIC")
  diagVals <- c(dev, dic)
  
  diagnostics <- data.frame(cbind(diagNames, diagVals))
  colnames(diagnostics) <- c("Diagnostics")
  
  
  #######################################################################################################
  # CREATE WORKBOOK
  #######################################################################################################
  wb <- createWorkbook(type="xlsx")
  CellStyle(wb, dataFormat=NULL, alignment=NULL, border=NULL, fill=NULL, font=NULL)
  
  #######################################################################################################
  # DEFINE STYLES
  #######################################################################################################
  TITLE_STYLE <- CellStyle(wb) + Font(wb,  heightInPoints=14, color="lightsteelblue4", isBold=TRUE, underline=0)+
    Alignment(wrapText=TRUE, horizontal="ALIGN_CENTER")
  SUB_TITLE_STYLE <- CellStyle(wb) + Font(wb,  heightInPoints=12, isItalic=TRUE, isBold=FALSE) +
    Alignment(wrapText=TRUE, horizontal="ALIGN_CENTER")
  HEADING_STYLE <- CellStyle(wb) + Font(wb,  heightInPoints=11, isItalic=FALSE, isBold=TRUE) +
    Alignment(wrapText=TRUE, horizontal="ALIGN_LEFT")
  BOLD <- CellStyle(wb) + Font(wb, isBold=TRUE)
  
  TABLE_ROWNAMES_STYLE <- CellStyle(wb) + Font(wb, isBold=TRUE)
  TABLE_COLNAMES_STYLE <- CellStyle(wb) + Font(wb, isBold=TRUE) +
    Alignment(wrapText=TRUE, horizontal="ALIGN_LEFT") +
    Border(color="black", position=c("TOP", "BOTTOM"), pen=c("BORDER_THIN", "BORDER_THICK"))
  
  #######################################################################################################
  # CREATE SHEETS WITH DATA
  #######################################################################################################
  sheet <- createSheet(wb, sheetName = modelname)
  xlsx.addTitle<-function(sheet, rowIndex, colIndex, title, titleStyle) {
    rows <-createRow(sheet, rowIndex=rowIndex)
    sheetTitle <-createCell(rows, colIndex=colIndex)
    setCellValue(sheetTitle[[1,1]], title); setCellStyle(sheetTitle[[1,1]], titleStyle)
  }
  xlsx.addTitle(sheet, rowIndex=startRowTitle, colIndex=startColModel, title=modelname,
                titleStyle = TITLE_STYLE)
  xlsx.addTitle(sheet, rowIndex=startRowSubtitle, colIndex=startColModel, title=modelname2,
                titleStyle = SUB_TITLE_STYLE)
  
  # Add a table
  addDataFrame(df, sheet, startRow=startRowData, startColumn=1, row.names = FALSE,
               colnamesStyle = TABLE_COLNAMES_STYLE, rownamesStyle = TABLE_ROWNAMES_STYLE)
  
  addDataFrame(dfRandom, sheet, startRow=startRowData+2+nrow(df), startColumn=1, row.names = FALSE,
               colnamesStyle = TABLE_COLNAMES_STYLE, rownamesStyle = TABLE_ROWNAMES_STYLE)
  
  addDataFrame(diagnostics, sheet, startRow=startRowData+4+nrow(df)+nrow(dfRandom), startColumn=1, row.names = FALSE,
               colnamesStyle = TABLE_COLNAMES_STYLE, rownamesStyle = TABLE_ROWNAMES_STYLE)
  
  setColumnWidth(sheet, 1, 30)
  setColumnWidth(sheet, 2, 22)
  
  #######################################################################################################
  # CONDITIONAL FORMATTING
  #######################################################################################################
  fLow <- Fill(foregroundColor="lightsteelblue1")  # create fill object
  csLow <- CellStyle(wb, fill=fLow)          # create cell style
  fHigh <- Fill(foregroundColor="bisque1")  # create fill object
  csHigh <- CellStyle(wb, fill=fHigh)          # create cell style
  sheets <- getSheets(wb)               # get all sheets
  sheet <- sheets[[modelname]]          # get specific sheet
  whichLow <- which(WS=="Low")
  whichHigh <- which(WS=="High")
  rows <- getRows(sheet, rowIndex=4:(nrow(df)+3))    # get rows
  cells <- getCells(rows, colIndex=2)
  
  lapply(cells[whichLow], function(i) setCellStyle(i, csLow))
  lapply(cells[whichHigh], function(i) setCellStyle(i, csHigh))
  
  #######################################################################################################
  # SAVE WORKBOOK
  #######################################################################################################
  saveWorkbook(wb, paste(modelname,".xlsx",sep=""))
}

