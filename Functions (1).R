

# FUNCTIONS *****************************************************************************

library(data.table)
library(dplyr)
library(RDCOMClient)
library(lubridate)
library(readr)
library(stringr)
library(readxl)
library(formattable)

# switch values for other values or bin continuous data
swit = function(vec, stuffstart, stuffFin, contin = F, left = F){
  # takes vector to be changed, 
  # values to be changed, and 
  # corresponding values to become
  # Continuous true will bin data 
  # stuffstart are the boundaries
  # stuffFin are the value mapping for the corresponding bins
  # bins including right boundary unless left = T
  newvec = vec
  if(contin){
    if(left){
      # binning [l, u)
      newvec[vec < stuffstart[1]] = stuffFin[1]
      for (i in 1:(length(stuffstart)-1) ){
        newvec[(vec >= stuffstart[i]) & (vec < stuffstart[i+1])] = stuffFin[i+1]
      }
      newvec[vec >= stuffstart[length(stuffstart)]] = stuffFin[length(stuffFin)]
    }
    else {
      # binning (l, u]
      newvec[vec <= stuffstart[1]] = stuffFin[1]
      for (i in 1:(length(stuffstart)-1) ){
        newvec[(vec > stuffstart[i]) & (vec <= stuffstart[i+1])] = stuffFin[i+1]
      }
      newvec[vec > stuffstart[length(stuffstart)]] = stuffFin[length(stuffFin)]
    }
  }
  else {
    for(i in 1:length(stuffstart)){
      newvec[vec == stuffstart[i]] = stuffFin[i]
    }
  }
  
  return(newvec)
}

# upper case and trim white space
strcln <- function(names){
  gsub("[ ]{2,}", " ", trimws(toupper(names)))
}

# dollar string vector to all numeric (NA to 0)
clDol <- function(dollars, na0 = T, curren = F){
  num <- as.numeric(gsub("[$,]", "", dollars))
  if(na0){
    num[is.na(num)] <- 0
  }
  if(curren){
    num <- currency(num)
  }
  return(num)
}

# input numeric(able) output Accounting format text
formDol <- function(dollars, digs = 2){
  if(suppressWarnings(sum(is.na(as.numeric(dollars[!is.na(dollars)]))) != 0)){
    stop("Need a numeric(able) vector")
  }
  dollars = round(as.numeric(dollars), digs)
  fdol <- as.character(dollars)
  fdol[dollars == 0] = "-   "
  
  fdol[dollars < 0 & !is.na(dollars)] = 
    paste0('(',format(dollars[dollars < 0 &!is.na(dollars)]*-1, big.mark = ",", nsmall = digs), ')')
  
  fdol[dollars > 0 & !is.na(dollars)] = 
    paste0(format(dollars[dollars > 0 & !is.na(dollars)], big.mark = ",", nsmall = digs), " ")
  
  wid = max(str_length(fdol), na.rm = T)
  fdol[!is.na(fdol)] = paste0(" $", str_pad(fdol[!is.na(fdol)], width = wid))
  fdol[is.na(fdol)] <- ""
  return(fdol)
}

# formats DataFrame as nice string
DF_to_Text <- function(dfrm){
  paste(format(dfrm)[-c(1,3)], collapse = "\n")
}

# Cleans data for export changes all to character and NA to empty strings
ExportDF <- function(dfst){
  
  if(nrow(dfst) == 0){
    
    print(paste(nrow(dfst), "Observation(s) in Data.  No cleaning possible"))
    return(dfst)
    
  } else if(nrow(dfst) == 1){
    
    for(i in 1:length(dfst[1, ])){
      dfst[[i]] <- as.character(dfst[[i]])
    }
    return(dfst)
    
  } else {
    dfExport <- as.data.frame(sapply(dfst, as.character), stringsAsFactors = F)
    dfExport[is.na(dfExport)] <- ""
    return(dfExport)
  }
}

# Read Info files on demand into a list
readInfo <- function(InfoTypes = c("Div", "Reg", "CC", "HC")){
  curdir <- getwd()
  types <- c('Div', 'Reg', 'CC', 'HC', 'FHC', 'CHC')
  files <- c("Branch-Division_Info/Divisions.csv", 
             "Branch-Division_Info/Regions.csv", "Branch-Division_Info/Cost Centers.csv",
             "Loan_Officers/HeadCt.csv", "Loan_Officers/FullHeadCt.csv", 
             "Loan_Officers/CurHeadCt.csv")
  Info <- list()
  setwd("~/Clean_Data")
  
  for(itype in InfoTypes){
    if(itype %in% types){
      dfm <- fread(files[which(itype == types)], sep = ',')
      Info[[itype]] <- dfm
    }
  }
  setwd(curdir)
  if(length(Info) == 1){
    return(Info[[1]])
  } else {
    return(Info)
  }
}

# Read Cleaner files on demand into a list
readcleaners <- function(ClnTypes = c("CCName", "Name", "CCNum")){
  curdir <- getwd()
  codes <- c("CCName", "Name", "CCNum")
  files <- c("CCNameClean.csv", "NameClean.csv", "CCNumClean.csv")
  
  Cleaners <- list()
  setwd("~/Clean_Data/Cleaners")
  for(type in ClnTypes){
    dfm <- fread(files[codes == type], sep = ',')
    Cleaners[[type]] <- dfm
  }
  setwd(curdir)
  if(length(Cleaners) == 1){
    return(Cleaners[[1]])
  } else {
    return(Cleaners)
  }
}

# Read in Connectors
readConnectors <- function(ConnTypes = c("EtoC", "CtoD")){
  curdir <- getwd()
  codes <- c("CIDtoC", "CtoD", "CtoR", "EtoD", "EtoC", "RtoD", "CIDtoD")
  files <- c("BCCIDtoCC.csv", "CCDivConn.csv", "BCCtoRegion.csv", "EmpDivCon.csv", 
             "BEmpToCC.csv", "BRegToDiv.csv", "CidDivCon.csv")
  Conns <- list()
  setwd("~/Clean_Data/Connectors")
  for(conn in ConnTypes){
    dfm <- fread(files[codes == conn], sep = ',')
    Conns[[conn]] <- dfm
  }
  setwd(curdir)
  if(length(Conns) == 1){
    return(Conns[[1]])
  } else {
    return(Conns)
  }
}

# general read and clean dataframe
ReadandClean <- function(flpth, Varnms1 = NULL, Varnms2 = NULL, clnTypes = NULL, getNames = F){
  # Need to add date clean
  
  # Check for mistakes
  # file exists and is a .csv file
  if(!file.exists(flpth) & !grepl("\\.csv$", flpth) & 
     !grepl("\\.xlsx$", flpth) & !grepl("\\.xls$", flpth)){
    stop("File does not exist or is not a csv/excel file")
  }
  
  # read basic
  if(grepl("\\.csv$", flpth)){
    dfm <- fread(flpth, sep = ',', na.strings = c("", "NA", "n/a"))
  } else {
    dfm <- read_excel(flpth)
  }
  
  
  # return names for use
  if(getNames){
    return(names(dfm))
  }
  
  # Varnms2 exists
  if(!is.null(Varnms2)){
    # Varnms1 doesn't exist
    if(is.null(Varnms1)){
      stop("Only have replacement variables if original variables are given")
      
    # Varnms2 and Varnms1 not the same length
    } else if(length(Varnms2) != length(Varnms1)){
        stop("Start and finish variable names need to have the same length")
    }
  }
  
  # Clean Types exists
  if(!is.null(clnTypes)){
    if(sum(clnTypes %in% c("string", "dollar", "none")) != length(clnTypes)){
      stop("Clean Types need to be string, dollar, or none")
    }
    # Varnms1 exists and isn't the same length
    if(!is.null(Varnms1)){
      if(length(clnTypes) != length(Varnms1)){
        stop("Clean types need to be the same length as variable names")
      }
    # Varnms1 doesn't exist and clean types not the same length as dataframe
    } else {
      if(length(clnTypes) != length(dfm)){
        stop("Clean Types need to be the same length as the number of variables")
      }
    }
  }
  
  # Finished with checks and completing read and clean
  if(!is.null(Varnms1)){
    dfm <- dfm %>% select(Varnms1)
  }
  
  if(!is.null(Varnms2)){
    names(dfm) <- Varnms2
  }
  
  if(!is.null(clnTypes)){
    varOrd <- names(dfm)
    
    dfmstr <- dfm %>% select(which(clnTypes == "string")) %>% 
      mutate_all(strcln)
    dfmdol <- dfm %>% select(which(clnTypes == "dollar")) %>%
      mutate_all(clDol)
    dfm <- dfm %>% select(which(clnTypes == "none")) %>%
      bind_cols(dfmstr, dfmdol) %>% select(varOrd)
  }
  return(dfm)
}


#test = ReadandClean("~/Important/ConcurRef.csv", getNames = T)
#replnms <- c("Vendor", "Vendor#", "GL", "UIName", "UI", "Owner", "ExpenseType",
#             "Descrp", "ToConcur", "Distr", "Zip", "Auto")
#cltyp <- c("string", "none", "dollar", "none", "string", "string", "string", 
#           rep("none", 5))
#test1 = ReadandClean("~/Important/ConcurRef.csv", test, replnms)


# Send Outlook Emails
# sending from current user
SendEmail <- function(to, subject, message, attachment = NA){
  # to is a vector of emails to send to
  # Subject and message are both strings
  # attachment is a complete file path to the desired attachment
  
  OutApp <- COMCreate("Outlook.Application")
  ## create an email 
  outMail = OutApp$CreateItem(0)
  
  ## configure  email parameter 
  outMail[["To"]] = paste(to, collapse = ";")
  outMail[["subject"]] = subject
  outMail[["body"]] = message
  
  # Attachments need full path
  if(!is.na(attachment)){
    outMail[["Attachments"]]$Add(attachment)
  }
  
  ## send it
  return(outMail$Send())
}

# Returns/ prints missing data
Missing <- function(dfm, missvar, othervars, prin = T){
  if(!(missvar %in% names(dfm))){
    stop("Missing Variable name misspelled or not in dataframe")
  }
  if(sum(othervars %in% names(dfm)) != length(othervars)){
    stop("One or more of the other variable names not in dataframe")
  }
  
  MissStuff <- dfm %>% select(missvar, othervars) %>% 
    filter(is.na(get(missvar))) %>%
    filter(!is.na(get(othervars))) %>%
    unique()
  
  nmiss <- nrow(MissStuff)
  if(nmiss == 0){
    print(paste("No missing", missvar, "values!!"))
  } else {
    print(paste(nmiss, "missing", missvar, "values :("))
    if(prin){
      print(MissStuff)
    } else {
      return(MissStuff)
    }
  }
}

