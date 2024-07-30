### Performance analysis reporting script: Wingate, Spiroergometry, Anthropometry, Kistler force plates ###

wd <- "C:/Users/DKolinger/Documents/R/Wingate + Spiro" #select path to files


# Read READ_ME.txt file before running      
# !DO NOT CHANGE CODE BEYOND THIS LINE!
#-----------------------------------------------------------------------------#

setwd(wd)
##### Reference values for GUI of sport selection ####
sports_data <- data.frame(
  Sport = c("hokej-dospělí","hokej-junioři", "hokej-dorost", "gymnastika"),
  FVC = c(4.7, 4.7, 4.7, 3.8),    
  FEV1 = c(4.7, 4.7, 4.7, 3.1),   
  VO2max = c(58, 58, 58, 48),   
  PowerVO2 = c(4.5, 4.4, 4.4 ,3.3),
  Pmax = c(15.7, 15.3, 14.6, 13),
  SJ = c(45, 42, 42, 35),
  ANC = c(365, 330, 320 ,300)
)

#### Automatic check and instalation of missing packages ####

instalace <- function() {
  files <- list.files(pattern='[.](R|rmd)$', all.files=T, recursive=T, full.names = T, ignore.case=T) # find all source code files in (sub)folders
  code=unlist(sapply(files, scan, what = 'character', quiet = TRUE))   # read in source code
  
  code <- code[grepl('^library', code, ignore.case=T)]   # retain only source code starting with library
  code <- gsub('^library[(]', '', code)
  code <- gsub('[)]', '', code)
  code <- gsub('^library$', '', code)
  
  uniq_packages <- unique(code)   # retain unique packages
  uniq_packages <- uniq_packages[!uniq_packages == '']   # kick out "empty" package names
  uniq_packages <- uniq_packages[order(uniq_packages)]   # order alphabetically
  
  cat('Required packages: \n')
  cat(paste0(uniq_packages, collapse= ', '),fill=T)
  cat('\n\n\n')
  
  installed_packages <- installed.packages()[, 'Package']   # retrieve list of already installed packages
  
  to_be_installed <- setdiff(uniq_packages, installed_packages)   # identify missing packages
  
  if (length(to_be_installed)==length(uniq_packages)) cat('Vsechny balicky musi byt nainstalovany - instaluji.\n')
  if (length(to_be_installed)>0) cat('Instaluji chybejici balicky.\n')
  if (length(to_be_installed)==0) cat('Vsechny balicky jsou jiz nainstalovany!\n')
  
  if (length(to_be_installed)>0) install.packages(to_be_installed, repos = 'https://cloud.r-project.org')   # install missing packages
}

instalace()

if(!require("tinytex")){
  tinytex::install_tinytex()
}

#### Type of execution, Raw data control, check IDs ####
dotaz_spiro <- winDialog("yesno", "Přidat spiro report?")
srovnani <- winDialog("yesno", "Přidat srovnání s předchozími výsledky?")


file.path <- paste(wd, "/wingate", sep = "")  #Wingate folder
file.path.an <- paste(wd, "/antropometrie", sep = "")  #Anthropometry folder
report.path <- paste(wd, "/reporty", sep = "") #Reports folder
database.path <- paste(wd, "/databaze", sep = "") #Database folder
spiro.path <- paste(wd, "/spiro", sep = "") #Spiroergometry folder
list.files(path = file.path, pattern = '*.txt')

if (srovnani == "YES") {
  winDialog("ok", "Vyberte složku pro nejnovější srovnání")
  comparison.path <-  rstudioapi::selectDirectory(path = getwd(), caption = "Vyberte složku pro nejnovější srovnání")
  file.list.compar <- list.files(path = comparison.path, pattern = '*.txt')
  dalsi_srov <- winDialog("yesno", "Přidat další srovnání?")
  if (dalsi_srov == "YES") {
    comparison.path_2 <- rstudioapi::selectDirectory(path = getwd(), caption = "Vyberte složku pro další srovnání")
    file.list.compar_2 <- list.files(path = comparison.path_2, pattern = '*.txt')
    tri_graf <- winDialog("yesno", "Zahrnout nejstarší srovnání do grafu?")
  }
}


# Create file lists
file.list <- list.files(path = file.path, pattern = '*.txt')
file.list.an <- list.files(path = file.path.an, pattern = '\\.xls[xm]$', ignore.case = TRUE)
file.list.dat <- list.files(path = database.path, pattern = c('*.xls'))
file.list.bez <- sort(tools::file_path_sans_ext(file.list))
file.list.spiro <- list.files(path = spiro.path, pattern = c('*.xls'))
file.list.spiro.bez <- sort(tools::file_path_sans_ext(file.list.spiro))


# Check type of anthropometry file

if (length(file.list.an)<2) {
  antropo <- readxl::read_excel(paste(file.path.an, "/", file.list.an[1], sep=""), sheet = "Data_Sheet") 
  if (nrow(antropo)<2) {
    an.input <- "solo"
  } else if (antropo$Surname[1] == antropo$Surname[2]) {
    an.input <- "solo"
  } else {
    an.input <- "batch"
  }
} else {
  an.input <- "solo"
}



if (an.input == "batch") {
  file.list.an.bez <- antropo$ID
  lactate.bez <- antropo$ID[!is.na(antropo$LA)]
} else {
  file.list.an.bez <- sort(tools::file_path_sans_ext(file.list.an))
}


# ID check
kontrola <- sapply(list(unique(sort(file.list.an.bez))), FUN = identical, file.list.bez)
kontrola.spiro <- ifelse(dotaz_spiro == "YES", sapply(list(file.list.spiro.bez), FUN = identical, file.list.bez), NA)
kontrola.s2 <- NA


if (srovnani == "YES") {
  if (exists("file.list.compar_2")) {
    file.list.compar_2.bez <- sort(tools::file_path_sans_ext(file.list.compar_2))
    kontrola.s2 <- sapply(list(file.list.compar_2.bez), FUN = identical, file.list.bez)
  }
  file.list.compar.bez <- c()
  file.list.compar.bez <- sort(tools::file_path_sans_ext(file.list.compar))  
  kontrola.s <- sapply(list(file.list.compar.bez), FUN = identical, file.list.bez)
  
  
}

duplikaty <- duplicated(file.list.an.bez)

# Raw data Error reporting
if ("TRUE" %in% duplikaty) {
  print("Duplikaty v souboru s antropometrii:")
  print(file.list.an.bez[which(duplikaty=="TRUE")])
  if (srovnani != "YES") {
    stop(winDialog("ok", paste("Export přerušen: Nalezeno", length(duplikaty), "duplikátů v souboru s Antropometrií", sep = " ")))
  }
  else if (srovnani == "YES") {
    print(paste("Nalezeno celkem", length(which(duplikaty=="TRUE")), "duplikátù", sep = " "))
    duplikaty_check <- winDialog("yesno", paste("Nalezeno celkem", length(which(duplikaty=="TRUE")), "duplikátů v antropometrii, pokračovat v exportu?", sep = " "))
    pocty.duplikatu.an <- dplyr::as_tibble(table(file.list.an.bez))
    pocty.duplikatu.an <- pocty.duplikatu.an[pocty.duplikatu.an$n != 1, ]
    names(pocty.duplikatu.an )[1] <- "ID"
    names(pocty.duplikatu.an )[2] <- "antropa"
    merged.file.list <- c(file.list.bez, file.list.compar.bez, file.list.compar_2.bez)
    pocty.file.list <- dplyr::as_tibble(table(merged.file.list))
    pocty.file.list <- pocty.file.list[pocty.file.list$n != 1, ]
    names(pocty.file.list)[1] <- "ID"
    names(pocty.file.list)[2]<- "wingate"
    pocty.file.list <- dplyr::left_join(pocty.file.list, pocty.duplikatu.an, by = "ID")
    pocty.file.list$diff <- pocty.file.list$wingate - pocty.file.list$antropa
    chybne.pocty <- as.vector(pocty.file.list$ID[pocty.file.list$diff > 0])
    if (duplikaty_check == "NO") {
      stop("OPERACE PRERUSENA UZIVATELEM", call. = F)
    } else if (duplikaty_check == "YES" & length(chybne.pocty) > 0) {
      pocty_check <- winDialog("yesno", paste("Nalezeno celkem", length(chybne.pocty), "chybějících záznamů pro srovnání v antropě, pokračovat?", sep = " "))
      print("Chybí antropa u srovnávacího wingate pro:")
      cat('\n')
      print(chybne.pocty)
      if (pocty_check == "YES") {
        spatne.compar <- chybne.pocty
        spatne.compar_2 <- chybne.pocty
      } else if (pocty_check == "NO") {
        view(pocty.file.list)
        stop("OPERACE PRERUSENA UZIVATELEM", call. = F)
      }
    }
  }
  
}
if ("FALSE" %in% kontrola & !pracma::isempty(setdiff(file.list.bez, file.list.an.bez))) {
  cat('\n\n')
  print("Nazvy Wingate jsou rozdilne oproti Antropometrii - ZKONTROLUJ:")
  cat('\n')
  if (!pracma::isempty(setdiff(file.list.bez, file.list.an.bez))) {print(setdiff(file.list.bez, file.list.an.bez))}
  input_x <- winDialog("yesno", paste("Nalezeno celkem", length(setdiff(file.list.bez, file.list.an.bez)), "neshod ve Wingate:Antropometrii, pokračovat v exportu?", sep = " "))
  if (input_x == "NO") {
    stop("OPERACE PRERUSENA UZIVATELEM", call. = F)}
  else if (input_x == "YES") {
    spatne.wingate <- setdiff(file.list.bez, file.list.an.bez)
    spatne.wingate <- paste(spatne.wingate, ".txt", sep = "")
    file.list <-  file.list[! file.list %in% spatne.wingate] 
  }
}
if  ("FALSE" %in% kontrola.spiro & !pracma::isempty(setdiff(file.list.spiro.bez, file.list.an.bez))) {
  cat('\n\n')
  print("Nazvy Antropa jsou rozdilne oproti Spiro - ZKONTROLUJ:")
  cat('\n')
  diff_str <- paste(setdiff(file.list.spiro.bez, file.list.an.bez), collapse = ", ")
  print(diff_str)
  if (!pracma::isempty(setdiff(file.list.spiro.bez, file.list.an.bez))) {pokracovat <- winDialog("yesno", paste("Rozdíly v názvech Spiro/Antropo:", diff_str, "POKRAČOVAT?", sep= " "))}
  if (pokracovat == "NO") {
    stop("EXPORT PRERUSEN", call. = F)
  } else if (pokracovat == "YES") {
    spatne.spiro <- setdiff(file.list.spiro.bez, file.list.an.bez)
    spatne.spiro <- paste(spatne.spiro, ".xlsx", sep = "")
    file.list.spiro <-  file.list.spiro[!file.list.spiro %in% spatne.spiro] 
  }
  
} else if (srovnani == "YES") {
  if ("FALSE" %in% kontrola || "FALSE" %in% kontrola.s) {
    if ("FALSE" %in% kontrola & !pracma::isempty(setdiff(file.list.bez, file.list.an.bez))) {
      cat('\n\n')
      print("Nazvy Wingate jsou rozdilne oproti Antropometrii - ZKONTROLUJ:")
      cat('\n')
      if (!pracma::isempty(setdiff(file.list.bez, file.list.an.bez))) {print(setdiff(file.list.bez, file.list.an.bez))}
    }
    if  (dotaz_spiro == "YES" & "FALSE" %in% kontrola.spiro & !pracma::isempty(setdiff(file.list.bez, file.list.spiro.bez))) {
      cat('\n\n')
      print("Nazvy Wingate jsou rozdilne oproti Spiro - ZKONTROLUJ:")
      cat('\n')
      if (!pracma::isempty(setdiff(file.list.bez, file.list.spiro.bez))) {
        print(setdiff(file.list.bez, file.list.spiro.bez))
        spatne.spiro <- setdiff(file.list.bez, file.list.spiro.bez)
      }
    }
    if ("FALSE" %in% kontrola.s) {
      cat('\n\n')
      print("Nazvy Wingate jsou rozdilne (chybi) oproti Srovnani - ZKONTROLUJ:")
      cat('\n')
      if (!pracma::isempty(setdiff(file.list.bez, file.list.compar.bez))) {
        print(setdiff(file.list.bez, file.list.compar.bez))
        input_y <- winDialog("yesno", paste("Srovnání 1 chybí pro", length(setdiff(file.list.bez, file.list.compar.bez)),"Wingate, pokračovat bez srovnání?", sep = " "))
      }  }
    if ("FALSE" %in% kontrola.s2) {
      cat('\n\n')
      print("Nazvy Wingate jsou rozdilne (chybi) oproti Srovnani 2 - ZKONTROLUJ:")
      cat('\n')
      if (!pracma::isempty(setdiff(file.list.bez, file.list.compar_2.bez))) {print(setdiff(file.list.bez, file.list.compar_2.bez))}
      input_2 <- winDialog("yesno", paste("Srovnání 2 chybí pro", length(setdiff(file.list.bez, file.list.compar_2.bez)),"Wingate, pokračovat bez srovnání?", sep = " "))
    }
    if (exists("input_y")) {
      if (input_y == "NO") {
        stop("OPERACE PRERUSENA UZIVATELEM", call. = F)
      } else if (input_y == "YES") {
        if ("FALSE" %in% kontrola) {
          spatne.wingate <- setdiff(file.list.bez, file.list.an.bez)
          spatne.wingate <- paste(spatne.wingate, ".txt", sep = "")
          file.list <-  file.list[! file.list %in% spatne.wingate]
        } else if ("FALSE" %in% kontrola.s) {
          spatne.compar <- c(spatne.compar, setdiff(file.list.bez, file.list.compar.bez))
          spatne.compar <- unique(spatne.compar)
          
        }
      }
    }
    if(exists("input_2")) {
      if (input_2 == "NO") {
        stop("OPERACE PRERUSENA UZIVATELEM", call. = F) 
      } else if (input_2 == "YES") {
        if ("FALSE" %in% kontrola.s2) {
          spatne.compar_2 <- c(spatne.compar_2, setdiff(file.list.bez, file.list.compar_2.bez))
          spatne.compar_2 <- unique(spatne.compar_2)
          
        }
      }
    }
    
  }
  if (exists("spatne.compar_2")) {
    file.list.compar_2.bez <- file.list.compar_2.bez[!file.list.compar_2.bez %in% spatne.compar_2]
  }
  if (exists("spatne.compar")) {
    file.list.compar.bez <- file.list.compar.bez[!file.list.compar.bez %in% spatne.compar]   
  }
  
}







library(tcltk2)

#### Sport selection GUI ####
onSelect <- function() {
  selected_index <- tclvalue(tkcurselection(lst))
  if (as.numeric(selected_index) >= 0) {
    selected_sport <- sports_data$Sport[as.numeric(selected_index) + 1]
    
    assign("sport", selected_sport, envir = .GlobalEnv)     # Global variable for sport selection
    
    tkdestroy(top)
    createTeamNameWindow()  
  }
}

createTeamNameWindow <- function() {
  win <- tktoplevel()
  tkwm.title(win, "Výběr teamu")
  
  
  label <- tklabel(win, text = "Napište jméno týmu:")   # Create a label
  tkpack(label, side = "top", padx = 10, pady = 10)
  
  
  textEntry <- tkentry(win, width = 50)   # Create a text entry widget             
  tkpack(textEntry, side = "top", padx = 20, pady = 20)
  
  # Function to handle button click
  onSubmit <- function() {
    team <<- tclvalue(tkget(textEntry))  
    tkdestroy(win)  # Optionally close the window after submitting
  }
  
  # submit button
  submitButton <- tkbutton(win, text = "Potvrdit", command = onSubmit)
  tkpack(submitButton, side = "top", padx = 10, pady = 10)
  
  # Run the GUI
  tkfocus(win)
  tkwait.window(win)
}

# Create main window
top <- tktoplevel()
tkwm.title(top, "Vyberte sport")

# listbox for sports
lst <- tk2listbox(top, height = 5)
tkpack(lst, fill = "both", expand = TRUE)

# Insert each sport into the listbox
for (sport in sports_data$Sport) {
  tkinsert(lst, "end", sport)
}

# Add button to confirm selection
btn_select <- tk2button(top, text = "Potvrdit", command = onSelect)
tkpack(btn_select, side = "bottom", fill = "x")

# Run the GUI
tkfocus(top)
tkwait.window(top)


library(finalfit)
library(dplyr)
library(ggplot2)
library(lubridate)
library(gt)
library(cowplot)
library(tidyverse)
library(tinytex)
library(rapportools)
library(scales)
library(webshot)
library(patchwork)

i <- 1
#### MAIN FOR LOOP ####
for(i in 1:length(file.list)) {
  # data Wingate
  if (exists("df2")) {
    rm(df2)
  }
  id <- tools::file_path_sans_ext(file.list[i])
  print(paste("Tisknu", i,"/", length(file.list), ": ", id))
  writeLines(iconv(readLines(paste(file.path, "/", file.list[i], sep = "")), from = "ANSI_X3.4-1986", to = "UTF8"), 
             file(paste(file.path, "/", file.list[i], sep = ""), encoding="UTF-8"))
  df <- read.delim(paste(file.path, "/", file.list[i], sep = ""))
  if (an.input == "solo") {
    file.an <- file.list.an[which(file.list.an.bez == id)]
    antropo <- readxl::read_excel(paste(file.path.an, "/", file.an, sep=""), sheet = "Data_Sheet")
    if (srovnani == "YES") {
      antropo <- antropo[order(as.Date(antropo$Date_measurement, format="%d/%m/%Y")),]
    }
  }
  
  if (an.input == "batch") {
    antropo <- readxl::read_excel(paste(file.path.an, "/", file.list.an[1], sep=""), sheet = "Data_Sheet")   
  }
  
  
  # 3rd comparison (oldest)
  if(exists("file.list.compar_2.bez")) {
    if (srovnani == "YES" & id %in% file.list.compar_2.bez) {
      writeLines(iconv(readLines(paste(comparison.path_2, "/", id, ".txt", sep = "")), from = "ANSI_X3.4-1986", to = "UTF8"), 
                 file(paste(comparison.path_2, "/", id, ".txt", sep = ""), encoding="UTF-8"))
      compare.wingate_2 <- read.delim(paste(comparison.path_2, "/", id, ".txt", sep = ""))
      compare.wingate_2$Work.total..KJ. <- as.numeric(gsub(",",".", compare.wingate_2$Work.total..KJ.))
      compare.wingate_2$Elapsed.time.total..h.mm.ss.hh. <- period_to_seconds(hms(compare.wingate_2$Elapsed.time.total..h.mm.ss.hh.))
      if (compare.wingate_2$Elapsed.time.total..h.mm.ss.hh.[1] > 3) {
        compare.wingate_2$Elapsed.time.total..h.mm.ss.hh. <- compare.wingate_2$Elapsed.time.total..h.mm.ss.hh.- compare.wingate_2$Elapsed.time.total..h.mm.ss.hh.[1]
      } else if (tail(compare.wingate_2$Elapsed.time.total..h.mm.ss.hh.,1) > 31) {
        cas.zacatku <- tail(compare.wingate_2$Elapsed.time.total..h.mm.ss.hh.,1) - 30
        radek.zacatku <- which.min(abs(compare.wingate_2$Elapsed.time.total..h.mm.ss.hh. - cas.zacatku)) - 1
        compare.wingate_2 <- compare.wingate_2[-(1:radek.zacatku),]
      }
      if (compare.wingate_2$Elapsed.time.total..h.mm.ss.hh.[1] > 3) {
        compare.wingate_2$Elapsed.time.total..h.mm.ss.hh. <- compare.wingate_2$Elapsed.time.total..h.mm.ss.hh.- compare.wingate_2$Elapsed.time.total..h.mm.ss.hh.[1] 
      }
      
      s5 <- as.numeric(head(which(compare.wingate_2$Elapsed.time.total..h.mm.ss.hh. >= 5), n=1))
      s10 <- as.numeric(head(which(compare.wingate_2$Elapsed.time.total..h.mm.ss.hh. >= 10), n=1))
      s15 <- as.numeric(head(which(compare.wingate_2$Elapsed.time.total..h.mm.ss.hh. >= 15), n=1))
      s20 <- as.numeric(head(which(compare.wingate_2$Elapsed.time.total..h.mm.ss.hh. >= 20), n=1))
      s25 <- as.numeric(head(which(compare.wingate_2$Elapsed.time.total..h.mm.ss.hh. >= 25), n=1))
      s30 <- as.numeric(length(compare.wingate_2$Elapsed.time.total..h.mm.ss.hh.))
      radky5s <- round(base::mean(s5, (s10-s5), (s15-s10), (s20-s15), (s25-s20), (s30-s25)),0)
      compare.wingate_2$RM5_Power <- zoo::rollmean(compare.wingate_2$Power..W., k=radky5s, fill = NA, align = "right")
      
      for (j in 1:length(compare.wingate_2$Power..W.)) {
        compare.wingate_2$AvP_dopocet[j] <- mean(compare.wingate_2$Power..W.[1:j] )
      }
    }
  }
  
  
  # The latest comparison
  if(exists("file.list.compar.bez")) {
    if (srovnani == "YES" & id %in% file.list.compar.bez) {
      writeLines(iconv(readLines(paste(comparison.path, "/", id, ".txt", sep = "")), from = "ANSI_X3.4-1986", to = "UTF8"), 
                 file(paste(comparison.path, "/", id, ".txt", sep = ""), encoding="UTF-8"))
      compare.wingate <- read.delim(paste(comparison.path, "/", id, ".txt", sep = ""))
      compare.wingate$Work.total..KJ. <- as.numeric(gsub(",",".", compare.wingate$Work.total..KJ.))
      compare.wingate$Elapsed.time.total..h.mm.ss.hh. <- period_to_seconds(hms(compare.wingate$Elapsed.time.total..h.mm.ss.hh.))
      if (compare.wingate$Elapsed.time.total..h.mm.ss.hh.[1] > 3) {
        compare.wingate$Elapsed.time.total..h.mm.ss.hh. <- compare.wingate$Elapsed.time.total..h.mm.ss.hh.- compare.wingate$Elapsed.time.total..h.mm.ss.hh.[1]
      } else if (tail(compare.wingate$Elapsed.time.total..h.mm.ss.hh.,1) > 31) {
        cas.zacatku <- tail(compare.wingate$Elapsed.time.total..h.mm.ss.hh.,1) - 30
        radek.zacatku <- which.min(abs(compare.wingate$Elapsed.time.total..h.mm.ss.hh. - cas.zacatku)) - 1
        compare.wingate <- compare.wingate[-(1:radek.zacatku),]
      }
      if (compare.wingate$Elapsed.time.total..h.mm.ss.hh.[1] > 3) {
        compare.wingate$Elapsed.time.total..h.mm.ss.hh. <- compare.wingate$Elapsed.time.total..h.mm.ss.hh.- compare.wingate$Elapsed.time.total..h.mm.ss.hh.[1] 
      }
      
      s5 <- as.numeric(head(which(compare.wingate$Elapsed.time.total..h.mm.ss.hh. >= 5), n=1))
      s10 <- as.numeric(head(which(compare.wingate$Elapsed.time.total..h.mm.ss.hh. >= 10), n=1))
      s15 <- as.numeric(head(which(compare.wingate$Elapsed.time.total..h.mm.ss.hh. >= 15), n=1))
      s20 <- as.numeric(head(which(compare.wingate$Elapsed.time.total..h.mm.ss.hh. >= 20), n=1))
      s25 <- as.numeric(head(which(compare.wingate$Elapsed.time.total..h.mm.ss.hh. >= 25), n=1))
      s30 <- as.numeric(length(compare.wingate$Elapsed.time.total..h.mm.ss.hh.))
      radky5s <- round(base::mean(s5, (s10-s5), (s15-s10), (s20-s15), (s25-s20), (s30-s25)),0)
      compare.wingate$RM5_Power <- zoo::rollmean(compare.wingate$Power..W., k=radky5s, fill = NA, align = "right")
      
      for (j in 1:length(compare.wingate$Power..W.)) {
        compare.wingate$AvP_dopocet[j] <- mean(compare.wingate$Power..W.[1:j] )
      }
    }
  }
  
  
  df$Work.total..KJ. <- as.numeric(gsub(",",".", df$Work.total..KJ.))
  
  df <- df[!apply(is.na(df) | df == "", 1, all),]
  df$Elapsed.time.total..h.mm.ss.hh. <- period_to_seconds(hms(df$Elapsed.time.total..h.mm.ss.hh.))
  
  
  df$Elapsed.time.total..h.mm.ss.hh. <- as.numeric(df$Elapsed.time.total..h.mm.ss.hh.)  
  
  
  # cut the original Wingate to 30s
  if (df$Elapsed.time.total..h.mm.ss.hh.[1] > 3) {
    df$Elapsed.time.total..h.mm.ss.hh. <- df$Elapsed.time.total..h.mm.ss.hh.- df$Elapsed.time.total..h.mm.ss.hh.[1]
  } else if (tail(df$Elapsed.time.total..h.mm.ss.hh.,1) > 31) {
    cas.zacatku <- tail(df$Elapsed.time.total..h.mm.ss.hh.,1) - 30
    radek.zacatku <- which.min(abs(df$Elapsed.time.total..h.mm.ss.hh. - cas.zacatku)) - 1
    df <- df[-(1:radek.zacatku),]
    if (df$Elapsed.time.total..h.mm.ss.hh.[1] > 3) {
      df$Elapsed.time.total..h.mm.ss.hh. <- df$Elapsed.time.total..h.mm.ss.hh.- df$Elapsed.time.total..h.mm.ss.hh.[1]
    }
  }
  
  # find number of rows for 5 seconds of test
  s5 <- as.numeric(head(which(df$Elapsed.time.total..h.mm.ss.hh. >= 5), n=1))
  s10 <- as.numeric(head(which(df$Elapsed.time.total..h.mm.ss.hh. >= 10), n=1))
  s15 <-as.numeric(head(which(df$Elapsed.time.total..h.mm.ss.hh. >= 15), n=1))
  s20 <- as.numeric(head(which(df$Elapsed.time.total..h.mm.ss.hh. >= 20), n=1))
  s25 <- as.numeric(head(which(df$Elapsed.time.total..h.mm.ss.hh. >= 25), n=1))
  s30 <- as.numeric(length(df$Elapsed.time.total..h.mm.ss.hh.))
  differences <- c(s5, s10, s15, s20, s25, s30) - c(0, s5, s10, s15, s20, s25)
  radky5s <- round(mean(differences), 0)
  df$RM5_Power <- zoo::rollmean(df$Power..W., k=radky5s, fill = NA, align = "right")
  
  for (j in 1:length(df$Power..W.)) {
    df$AvP_dopocet[j] <- mean(df$Power..W.[1:j] )
  }
  
  ####  Anthropometry values calculation ####
  
  if (an.input == "batch") {
    datum_mer <- head(as.Date(na.omit(antropo$Date_measurement[antropo$ID == id],"%d/%m/%Y")),1)
    sj <- NA
    if("SJ" %in% colnames(antropo)) {sj <- head(na.omit(antropo$SJ[antropo$ID == id]),1)}
    sj <- ifelse(is.empty(sj), NA, sj)
    vaha <- head(na.omit(antropo$Weight[antropo$ID == id]),1)
    vaha <- head(ifelse(length(vaha) == 0, NA, vaha),1)
    vyska <- head(na.omit(antropo$Height[antropo$ID == id]),1)
    vyska <- head(ifelse(length(vyska) == 0, NA, vyska),1)
    fat <- head(round(na.omit(antropo$Fat[antropo$ID == id]),1),1)
    fat <- ifelse(length(fat) == 0, NA, fat)
    ath <- head(round(na.omit(antropo$ATH[antropo$ID == id]),1),1)
    ath <- ifelse(length(ath) == 0, NA, ath)
    birth <- head(na.omit(antropo$Birth[antropo$ID == id]),1)
    fullname <- head(paste(na.omit(antropo$Name[antropo$ID == id]), na.omit(antropo$Surname[antropo$ID == id]), sep = " ", collapse = NULL),1)
    fullname.rev <- head(paste(na.omit(antropo$Surname[antropo$ID == id]), na.omit(antropo$Name[antropo$ID == id]), sep = " ", collapse = NULL),1)
    datum_nar <- head(format(as.Date(na.omit(antropo$Birth[antropo$ID == id])),"%d/%m/%Y"),1)
    datum_mer <- head(format(as.Date(na.omit(antropo$Date_measurement[antropo$ID == id])),"%d/%m/%Y"),1)
    la <- ifelse(!is.empty(head(antropo$LA[antropo$ID == id],1)), head(antropo$LA[antropo$ID == id],1), NA)
    age <- ifelse(!is.empty(antropo$Age[antropo$ID == id]), 
                  round(na.omit(antropo$Age[antropo$ID == id]),2), 
                  ifelse(!is.na(datum_nar) & datum_nar != "",
                         floor(as.integer(difftime(as.Date(datum_mer, format = "%d/%m/%Y"), as.Date(datum_nar, format = "%d/%m/%Y"), units = "days") / 365.25)),
                         NA))
    age <- head(age,1)
  } else {
    datum_mer <- na.omit(tail(antropo$Date_measurement, n=1))
    vaha <- na.omit(tail(antropo$Weight, n = 1))
    sj <- NA
    if("SJ" %in% colnames(antropo)) {sj <- na.omit(tail(antropo$SJ, n = 1))}
    sj <- ifelse(is.empty(sj), NA, sj)
    vaha <- ifelse(length(vaha) == 0, NA, vaha)
    vyska <- na.omit(tail(antropo$Height, n = 1))
    vyska <- ifelse(length(vyska) == 0, NA, vyska)
    fat <- round(na.omit(tail(antropo$Fat, n =1)),1)
    fat <- ifelse(length(fat) == 0, NA, fat)
    ath <- round(na.omit(tail(antropo$ATH, n =1)),1)
    ath <- ifelse(length(ath) == 0, NA, ath)
    birth <- na.omit(tail(antropo$Birth, n = 1))
    fullname <- paste(na.omit(tail(antropo$Name, n =1)), na.omit(tail(antropo$Surname, n = 1)), sep = " ", collapse = NULL)
    fullname.rev <- paste(tail(antropo$Surname, n = 1), tail(antropo$Name, n = 1), sep = " ", collapse = NULL)
    datum_nar <- format(as.Date(na.omit(tail(antropo$Birth, n =1))),"%d/%m/%Y")
    datum_mer <- format(as.Date(na.omit(tail(antropo$Date_measurement, n =1))),"%d/%m/%Y")
    la <- tail(antropo$LA, n=1)
    age <- ifelse(!is.empty(antropo$Age[antropo$ID == id]), 
                  round(na.omit(tail(antropo$Age, n =1)),2), 
                  ifelse(!is.na(datum_nar) & datum_nar != "",
                         floor(as.integer(difftime(as.Date(datum_mer, format = "%d/%m/%Y"), as.Date(datum_nar, format = "%d/%m/%Y"), units = "days") / 365.25)),
                         NA))
  }
  
  
  pp <- max(df$Power..W.)
  minp <- min(df$Power..W.[round((length(df$Power..W.)/2),0):length(df$Power..W.)])
  pp5s <- round(max(df$RM5_Power, na.rm = T),1)
  minp5s <- round(base::min(df$RM5_Power, na.rm = T),1)
  drop <- pp - minp
  iu <- round(((pp5s-minp5s)/pp5s)*100,1)
  avrp <- round(mean(df$Power..W.),1)
  totalw <- round(mean(df$AvP_dopocet*30),0)
  pp5sradek <- which(df$RM5_Power == pp5s)
  an.cap <- round(totalw/vaha,2)
  
  
  # Line smoothing
  vyhlazeno <- smooth.spline(df$Elapsed.time.total..h.mm.ss.hh., df$Power..W., spar = 0.4)
  y <- predict(vyhlazeno, newdata = df)
  y <- y$y
  
  if(exists("file.list.compar.bez")) {
    if(srovnani == "YES" & id %in% file.list.compar.bez) {
      compare.wingate <- compare.wingate[rowSums(is.na(compare.wingate)) < 3, ]
      vyhlazeno.compare <- smooth.spline(compare.wingate$Elapsed.time.total..h.mm.ss.hh., compare.wingate$Power..W., spar = 0.4)
      y2 <- predict(vyhlazeno.compare, newdata = compare.wingate)
      y2 <- y2$y
    }
  }
  
  if(exists("file.list.compar_2.bez")) {
    if(srovnani == "YES" & id %in% file.list.compar_2.bez) {
      vyhlazeno.compare_2 <- smooth.spline(compare.wingate_2$Elapsed.time.total..h.mm.ss.hh., compare.wingate_2$Power..W., spar = 0.4)
      y3 <- predict(vyhlazeno.compare_2, newdata = compare.wingate_2)
      y3 <- y3$y
    }
  }
  
  
  
  # Wingate reporting table
  columns <- c("Antropometrie", "V2", "Wingate test", "Absolutní", "Relativní (TH)", "Relativní (ATH)", "Pětivteřinový průměr")
  df1 = data.frame(matrix(nrow = 8, ncol = length(columns))) 
  colnames(df1) <- columns
  df1$Antropometrie <-  `length<-`(c("Datum narození", "Věk", "Výška", "TH - Hmotnost", "Tuk", "ATH - Aktivní tělesná hmota"), nrow(df1))
  df1$`Wingate test` <- `length<-`(c("Maximální výkon", "Minimální výkon", "Pokles výkonu", "Celková práce", "Index Únavy", "Celkový počet otáček", "Laktát", "Maximální Tepová frekvence (BPM)"),nrow(df1))
  
  df1[1,2] <- ifelse(!is.empty(datum_nar), datum_nar, NA)
  df1[2,2] <- ifelse(!is.empty(age), 
                     floor(age), 
                     ifelse(!is.na(datum_nar) & datum_nar != "",
                            floor(as.integer(difftime(as.Date(datum_mer, format = "%d/%m/%Y"), as.Date(datum_nar, format = "%d/%m/%Y"), units = "days") / 365.25)),
                            NA))
  df1[3,2] <- paste(vyska, "cm", sep = " ")
  df1[4,2] <- paste(vaha, "kg", sep = " ")
  df1[5,2] <- paste(fat, "%", sep = " ")
  df1[6,2] <- paste(ath, "kg", sep = " ")
  
  df1[1,4] <- paste(pp, "W", sep = " ")
  df1[2,4] <- paste(minp, "W", sep = " ")
  df1[3,4] <- paste(drop, "W", sep = " ")
  df1[4,4] <- paste(round(totalw/1000, 1), "kJ", sep = " ")
  df1[5,4] <- paste(iu, "%", sep = " ")
  df1[6,4] <- round(tail(df$Turns.number..Nr.,1),0)
  df1[7,4] <- ifelse(!is.na(la), paste(la, "mmol/l", sep = " "), NA)
  df1[8,4] <- as.numeric(max(df$Heart.rate..bpm., na.rm =T))
  
  df1$Absolutní[df1$Absolutní == 0] <- NA
  df1$`Relativní (TH)`[1] <- paste(round(pp/vaha,1), "W/kg", sep = " ")
  df1$`Relativní (TH)`[2] <- paste(round(minp/vaha,1), "W/kg", sep = " ")
  df1$`Relativní (TH)`[3] <- paste(round(drop/vaha,1), "W/kg", sep = " ")
  df1$`Relativní (TH)`[4] <- paste(round(totalw/vaha,0), "J/kg", sep = " ")
  df1$`Relativní (ATH)`[1] <- paste(round(pp/ath,1), "W/kg", sep = " ")
  df1$`Relativní (ATH)`[2] <- paste(round(minp/ath,1), "W/kg", sep = " ")
  df1$`Relativní (ATH)`[3] <- paste(round(drop/ath,1), "W/kg", sep = " ")
  df1$`Relativní (ATH)`[4] <- paste(round(totalw/ath,0), "J/kg", sep = " ")
  
  df1[5,5:6] <- NA
  df1[6,5:6] <- NA
  df1[6,5:6] <- NA
  df1[7:8,5:6] <- NA
  
  df1[1,7] <- paste(pp5s, "W", sep = " ")
  df1[2,7] <- paste(minp5s, "W", sep = " ")
  
  
  dotaz_spiro_original <- dotaz_spiro
  
  if (exists("spatne.spiro")) {
    if (id %in% spatne.spiro) {
      dotaz_spiro <- "NO"  
    }
  }
  
  
  k <- which(file.list.spiro.bez==id)
  
  #### Spiroergomatery values  ####
  if (dotaz_spiro == "YES") {
    if (length(file.list.spiro) != 0) {
      spiro <- readxl::read_excel(paste(spiro.path, "/", file.list.spiro[k], sep=""))
      spiro <- spiro[rowSums(!is.na(spiro)) > 0,]
      rows_with_bf <- which(spiro[[1]] == "BF")
      spiro_info <- spiro[1:rows_with_bf, ]
      
      spiro <- spiro[(rows_with_bf + 1):nrow(spiro), ]
      colnames(spiro) <- as.character(spiro[1, ])
      spiro <- spiro[-c(1,2), ]
      spiro <- subset(spiro, Fáze == "Zátěž")
      spiro[4:ncol(spiro)] <- lapply(spiro[4:ncol(spiro)], as.numeric)
      spiro$t <- gsub(",", ".", spiro$t)
      spiro$t <- as.POSIXct(spiro$t, format = "%H:%M:%OS", tz = "UTC")
      spiro$t <- spiro$t - min(spiro$t)
      time_intervals <- diff(spiro$t)
      median_interval <- as.numeric(median(time_intervals))
      closest_rows <- round(5 / median_interval)
      spiro$VT_5s <- zoo::rollmean(spiro$VT, k = closest_rows, align = "right", fill = NA)
    }
    
    s.datum.mer <- format(as.Date(spiro_info[[3]][which(spiro_info[[1]] == "Počáteční čas")], format = "%d.%m.%Y %H:%M"), "%d/%m/%Y")
    s.datum.nar <- format(as.Date(spiro_info[[3]][which(spiro_info[[1]] == "Datum narození")], format = "%d.%m.%Y"),"%d.%m.%Y")
    s.age <- floor(as.integer(difftime(as.Date(spiro_info[[3]][which(spiro_info[[1]] == "Počáteční čas")], format = "%d.%m.%Y %H:%M"), as.Date(spiro_info[[3]][which(spiro_info[[1]] == "Datum narození")], format = "%d.%m.%Y")), units = "days") / 365.25)
    s.height <- as.numeric(gsub("cm", "", gsub(",", ".", spiro_info[[3]][which(spiro_info[[1]] == "Výška")][1])))  
    FVC <- as.numeric(gsub("L", "", gsub(",", ".", spiro_info[[3]][which(spiro_info[[1]] == "VC")][1]))) 
    FEV1 <-  as.numeric(gsub("L", "", gsub(",", ".", spiro_info[[3]][which(spiro_info[[1]] == "FEV1")][1])))
    s.pomer <- round(FEV1/FVC*100,1)
    VO2max <- round(max(spiro$`V'O2`), 1)
    VO2_kg <- as.numeric(spiro_info[which(spiro_info[,1] == "V'O2/kg"), 12])
    vykon <- as.numeric(round(max(spiro$WR),0))
    vykon_kg <- as.numeric(vykon/vaha)
    hrmax <- round(max(spiro$TF),0)
    tep.kyslik <- as.numeric(spiro_info[which(spiro_info[,1] == "V'O2/HR"), 12])
    anp <- as.numeric(ifelse(spiro_info[which(spiro_info[,1] == "TF"),9]!="-", as.numeric(spiro_info[which(spiro_info[,1] == "TF"),9]), as.numeric(spiro_info[which(spiro_info[,1] == "TF"),15])*0.85))
    min.ventilace <- as.numeric(gsub(",", ".", spiro_info[which(spiro_info[,1] == "V'E"), 12]))
    dech.frek <- spiro_info[which(spiro_info[,1] == "BF"), 12]
    dech.objem <- round(max(spiro$VT_5s[which.max(as.numeric(format(spiro$t, "%H%M%S")) > 100):nrow(spiro)]),1)
    dech.objem.per <- round((dech.objem/FVC)*100,0)
    rer <- round(max(spiro$RER),2)
    LaMax <- la
    
    
    #Spiroergometry table
    columns_s <- c("Antropometrie", "Hodnota", "V3", "Spiroergometrie", "Absolutní", "Relativní (TH)", "% z nál. hodnoty" ,"Relativní (ATH)")
    df3 = data.frame(matrix(nrow = 14, ncol = length(columns_s))) 
    colnames(df3) <- columns_s
    df3$Antropometrie <-  `length<-`(c("Datum narození", "Věk", "Výška", "TH - Hmotnost", "Tuk", "ATH - Aktivní tělesná hmota", NA, "Klidová ventilace", "FVC", "FEV1", "Poměr FVC/FEV1"), nrow(df3))
    df3$Spiroergometrie <- `length<-`(c("VO2max", "Dosažený výkon", "HrMax", "Tepový kyslík", "ANP", NA, NA, "Ventilace", "Minutová ventilace", "Dechová frekvence", "Dechový objem", "Dechový objem %", "RER", "LaMax"),nrow(df3))
    df3$V3[8] <- "% z nál. hodnoty"
    df3$`Relativní (TH)`[8] <- "Optimum"
    
    df3[1,2] <- datum_nar
    df3[2,2] <- age
    df3[3,2] <-  paste(s.height, "cm", sep=" ")
    df3[4,2] <- paste(vaha, "kg", sep=" ")
    df3[5,2] <- paste(fat, "%", sep=" ")
    df3[6,2] <- paste(ath, "kg", sep=" ")
    df3[9,2] <- paste(FVC,"l", sep=" ")
    df3[10,2] <- paste(FEV1,"l", sep=" ")
    df3[11,2] <- paste(s.pomer,"%", sep=" ")
    df3[9,3] <- round((FVC/sports_data$FVC[sports_data$Sport==sport])*100,0)
    df3[10,3] <- round((FEV1/sports_data$FEV1[sports_data$Sport==sport])*100,0)
    df3[1,7] <- round((VO2_kg/sports_data$VO2max[sports_data$Sport==sport])*100,0)
    df3[2,7] <- round(((vykon/vaha)/sports_data$PowerVO2[sports_data$Sport==sport])*100,0)
    df3[1,5] <- paste(VO2max,"l", sep=" ")
    df3[2,5] <- paste(vykon,"W", sep=" ")
    df3[3,5] <- paste(hrmax,"BPM", sep=" ")
    df3[4,5] <- paste(tep.kyslik,"ml/kg", sep=" ")
    df3[5,5] <- paste(round(anp,0),"BPM", sep=" ")
    df3[9,5] <- paste(min.ventilace,"l/min", sep=" ")
    df3[10,5] <- paste(dech.frek,"d/min", sep=" ")
    df3[11,5] <- paste(dech.objem,"l", sep=" ")
    df3[12,5] <- paste(dech.objem.per,"%", sep=" ")
    df3[13,5] <- rer
    df3[14,5] <- ifelse(is.empty(la), NA, paste(la, "mmol/l", sep = " "))
    df3[1,6] <- paste(VO2_kg, "ml/min/kg", sep=" ")
    df3[2,6] <- paste(round(vykon/vaha, 1), "W/kg", sep=" ")
    df3[1,8] <- paste(round((VO2max*1000)/ath,1), "ml/min/kg", sep=" ")
    df3[2,8] <- paste(round(vykon/ath, 1), "W/kg", sep=" ")
    df3[9,6] <- round(30*FVC,0)
    df3[10,6] <- "50-60"
    df3[12,6] <- "50-60"
    df3[13,6] <- "1.08-1.18"
    df3 <- df3[-7,]
    
    
    # Training zones definition
    columns_tz <- c("Zóna", "od (BPM)", "do (BPM)")
    tz = data.frame(matrix(nrow = 3, ncol = length(columns_tz))) 
    colnames(tz) <- columns_tz
    tz$Zóna <- c("Aerobní", "Smíšená", "Anaerobní")
    tz$`od (BPM)`[2] <- round((anp / 0.85)*0.76,0)
    tz$`od (BPM)`[3] <- round(anp,0)
    tz$`do (BPM)`[1] <- round((anp / 0.85)*0.75,0)
    tz$`do (BPM)`[2] <- round((anp / 0.85)*0.84,0)
    
    
    vo2_range <- range(spiro$`V'O2`)
    coeff <- vo2_range * 10 
    
    
    
    #### Spiroergometry history table ####
    if (!exists("df4")) {
      df4 <- data.frame()
      df4 <- df4[nrow(df4) + 1,]
    }
    
    
    df4$Date_meas. <- append(na.omit(df4$Date_meas.), datum_mer)
    df4$Name <- append(na.omit(df4$Name), fullname)
    df4$Weight <- append(na.omit(df4$Weight), vaha)
    df4$Fat <- append(na.omit(df4$Fat), fat)
    df4$VO2max_l <- append(na.omit(df4$VO2max), VO2max)
    df4$`VO2_kg/l` <- append(na.omit(df4$VO2_kg), VO2_kg)
    df4$vykon_W <- append(na.omit(df4$vykon_W), vykon)
    df4$`vykon_l/kg` <- append(na.omit(df4$`vykon_l/kg`), round(vykon_kg,1))
    df4$hrmax_BPM<- append(na.omit(df4$hrmax), hrmax)
    df4$anp_BPM <- append(na.omit(df4$anp), round(anp,0))
    df4$tep_kys_ml <- append(na.omit(df4$tep_kys), tep.kyslik)
    df4$vt_l <- append(na.omit(df4$vt), dech.objem)
    df4$RER <- append(na.omit(df4$RER), rer)
    df4$`LaMax_mmol/l` <- append(na.omit(df4$LaMax), la)
    df4$FEV1_l <- append(df4$FEV1_l[1:length(df4$FEV1_l)-1], FEV1)
    df4$FVC_l <- append(na.omit(df4$FVC), FVC)
    df4$aerobni_Z_do <- append(na.omit(df4$aerobni_Z_do), round((anp / 0.85)*0.75,0))
    df4$smisena_Z_od <- append(na.omit(df4$smisena_Z_od), round((anp / 0.85)*0.76,0))
    df4$smisena_Z_do <- append(na.omit(df4$smisena_Z_do), round((anp / 0.85)*0.84,0))
    df4$anaerobni_Z_od <- append(na.omit(df4$anaerobni_Z_od), anp)
    df4 <- add_row(df4)
    
    
    hrmin <- min(spiro$TF)
    
    # Y axis limits calculation
    ylim.prim <- c(0, min.ventilace*1.1)
    ylim.sec <- c(hrmin*0.9, hrmax*1.05)  
    
    b <- diff(ylim.prim)/diff(ylim.sec)
    a <- ylim.prim[1] - b*ylim.sec[1]
    
    
    #### Spiroergometry plots ####
    plot2 <- ggplot(spiro, aes(x = t)) +
      geom_line(aes(y = a + TF*b, color = "Tepová frekvence (BPM)", linetype = "Tepová frekvence (BPM)"), size = 1) +
      geom_line(aes(y = `V'E`, color = "Minutová ventilace (l/min)", linetype = "Minutová ventilace (l/min)"), size = 0.5, alpha = 0.2, linetype = "dashed") +
      geom_smooth(aes(y = `V'E`), color = "black", linetype = "solid", size = 1, method = "loess", se = FALSE, span = 0.2) +
      scale_y_continuous(
        name = "Minutová ventilace (l/min)", 
        labels = scales::comma,
        sec.axis = sec_axis(~ (. - a)/b, name = "Tepová frekvence (BPM)")
      ) +
      scale_x_datetime(labels = scales::date_format("%M:%S"), date_breaks = "1 min") +
      theme_classic() +
      theme(
        text = element_text(size = 30),
        panel.grid.major = element_blank(),
        panel.grid.minor = element_blank(),
        axis.line = element_line(colour = "black", linewidth = 1.5),
        axis.text.x = element_text(angle = -45, hjust = 0),
        axis.title.y = element_text(size = 25),
        legend.position = "bottom"
      ) +
      scale_linetype_manual(values = c('solid', 'solid'), 
                            labels = c("Minutová ventilace (l/min)", "Tepová frekvence (BPM)")) +
      scale_color_manual(values = c("Minutová ventilace (l/min)" = "black", "Tepová frekvence (BPM)" = "orange")) +
      xlab(expression("Čas")) +
      ylab("Hodnota") +
      labs(colour="") +
      guides(
        color = guide_legend(override.aes = list(linetype = c("solid", "solid"))), 
        linetype = "none"
      )
    
    plot3 <- ggplot(spiro, aes(x = t)) +
      geom_line(aes(y = `V'O2`), color = "darkmagenta", size = 0.5, alpha = 0.2, linetype = "dashed") +
      geom_smooth(aes(y = `V'O2`), color = "darkmagenta", linetype = "solid", size = 1, method = "loess", se = FALSE, span = 0.3) +
      scale_x_datetime(labels = scales::date_format("%M:%S"), date_breaks = "1 min") +
      scale_y_continuous(
        name = "Spotřeba kyslíku (l)", 
        labels = scales::comma
      ) +
      theme_classic() +
      theme(
        text = element_text(size = 30),
        panel.grid.major = element_blank(),
        panel.grid.minor = element_blank(),
        axis.line = element_line(colour = "black", linewidth = 1.5),
        axis.text.x = element_text(angle = -45, hjust = 0),
        axis.title.y = element_text(size = 25),
        legend.position = "none"
      ) +
      xlab(expression("Čas")) +
      ylab("Spotřeba kyslíku (l)") 
    
    
    
    combined_plot <- plot2 + plot3 +
      plot_annotation(
        title = 'Spiroergometrie',
        caption = '',
        theme = theme(plot.title = element_text(size = 30, hjust = 0.5))
      )
  }
  
  
  #### Wingate values calculation ####
  if(exists("file.list.compar_2.bez")) {
    if (srovnani == "YES" & id %in% file.list.compar_2.bez) {
      if (exists("duplikaty_check")) {
        antropo <- readxl::read_excel(paste(file.path.an, "/", file.list.an[1], sep=""), sheet = "Data_Sheet") 
        antropo <- antropo %>% 
          filter(ID==id)
        antropo <- antropo[order(as.Date(antropo$Date_measurement, format="%d/%m/%Y")),]
      }
      df2 <- data.frame()
      df2 <- df2[nrow(df2) + 1,]
      df2$Date_meas. <- format(as.Date(antropo$Date_measurement[length(antropo$Date_measurement)-2]), "%d/%m/%Y")
      df2$Weight <- antropo$Weight[length(antropo$Weight)-2]
      df2$`Fat (%)` <- round(antropo$Fat[length(antropo$Fat)-2], 1)
      df2$ATH <- round(antropo$ATH[length(antropo$ATH)-2], 1)
      df2$`Pmax (W)` <- max(compare.wingate_2$Power..W.)
      df2$`Pmax/kg (W/kg)` <- round(max(compare.wingate_2$Power..W.)/antropo$Weight[length(antropo$Weight)-2],2)
      df2$`Pmin (W)` <- min(compare.wingate_2$Power..W.[round((length(compare.wingate_2$Power..W.)/2),0):length(compare.wingate_2$Power..W.)])
      df2$`Avg_power (W)` <- round(mean(compare.wingate_2$Power..W.),1)
      df2$`Pdrop (W)` <- df2$`Pmax (W)` - df2$`Pmin (W)`
      df2$`Pmax_5s (W)` <- round(max(compare.wingate_2$RM5_Power, na.rm = T),1)
      df2$`IU (%)` <- round((max(na.omit(compare.wingate_2$RM5_Power))-round(min(na.omit(compare.wingate_2$RM5_Power)),1))/max(na.omit(compare.wingate_2$RM5_Power))*100,1)
      df2$`Work (kJ)` <- round(mean(compare.wingate_2$AvP_dopocet*30),0)/1000
      df2$`HR_max (BPM)` <- max(compare.wingate_2$Heart.rate..bpm., na.rm =T)
      df2$`La_max (mmol/l)` <- antropo$LA[length(antropo$LA)-1]
      df2[nrow(df2) + 1,] <- NA
    }
  }
  
  if(exists("file.list.compar.bez")) {
    if (srovnani == "YES" & id %in% file.list.compar.bez) {
      if (exists("duplikaty_check")) {
        antropo <- readxl::read_excel(paste(file.path.an, "/", file.list.an[1], sep=""), sheet = "Data_Sheet") 
        antropo <- antropo %>% 
          filter(ID==id)
        antropo <- antropo[order(as.Date(antropo$Date_measurement, format="%d/%m/%Y")),]
      }
      if (!exists("df2")) {
        df2 <- data.frame()
        df2 <- df2[nrow(df2) + 1,]
      }
      df2$Date_meas. <- append(na.omit(df2$Date_meas.), format(as.Date(antropo$Date_measurement[length(antropo$Date_measurement)-1]), "%d/%m/%Y"))
      df2$Weight <- append(na.omit(df2$Weight), antropo$Weight[length(antropo$Weight)-1])
      df2$`Fat (%)` <- append(na.omit(df2$`Fat (%)`), round(antropo$Fat[length(antropo$Fat)-1], 1))
      df2$ATH <- append(na.omit(df2$ATH), round(antropo$ATH[length(antropo$ATH)-1], 1))
      df2$`Pmax (W)` <- append(na.omit(df2$`Pmax (W)`), max(compare.wingate$Power..W.))
      df2$`Pmax/kg (W/kg)` <- append(na.omit(df2$`Pmax/kg (W/kg)`), round(max(compare.wingate$Power..W.)/antropo$Weight[length(antropo$Weight)-1],2))
      df2$`Pmin (W)` <- append(na.omit(df2$`Pmin (W)`), min(compare.wingate$Power..W.[round((length(compare.wingate$Power..W.)/2),0):length(compare.wingate$Power..W.)]))
      df2$`Avg_power (W)` <- append(na.omit(df2$`Avg_power (W)`), round(mean(compare.wingate$Power..W.),1))
      df2$`Pdrop (W)` <- append(na.omit(df2$`Pdrop (W)`), (df2$`Pmax (W)`[nrow(df2)] - df2$`Pmin (W)`[nrow(df2)]))
      df2$`Pmax_5s (W)` <- append(na.omit(df2$`Pmax_5s (W)`), round(max(compare.wingate$RM5_Power, na.rm = T),1))
      df2$`IU (%)` <- append(na.omit(df2$`IU (%)`), round((max(na.omit(compare.wingate$RM5_Power))-round(min(na.omit(compare.wingate$RM5_Power)),1))/max(na.omit(compare.wingate$RM5_Power))*100,1))
      df2$`Work (kJ)` <- append(na.omit(df2$`Work (kJ)`), round(mean(compare.wingate$AvP_dopocet*30),0)/1000)
      df2$`HR_max (BPM)` <- append(na.omit(df2$`HR_max (BPM)`), max(compare.wingate$Heart.rate..bpm., na.rm =T))
      df2$`La_max (mmol/l)` <- append(na.omit(df2$`La_max (mmol/l)`), antropo$LA[length(antropo$LA)-1])
      df2[nrow(df2) + 1,] <- NA
    }
  }
  
  
  #Wingate history table
  if (!exists("df2")) {
    df2 <- data.frame()
    df2 <- df2[nrow(df2) + 1,]
  }
  
  df2$Date_meas. <- append(na.omit(df2$Date_meas.), datum_mer)
  df2$Weight <- append(na.omit(df2$Weight), vaha)
  df2$`Fat (%)` <- append(na.omit(df2$`Fat (%)`), fat)
  df2$ATH <- append(na.omit(df2$ATH), ath)
  df2$`Pmax (W)` <- append(na.omit(df2$`Pmax (W)`), pp)
  df2$`Pmax/kg (W/kg)` <- append(na.omit(df2$`Pmax/kg (W/kg)`), round(pp/vaha,2))
  df2$`Pmin (W)` <- append(na.omit(df2$`Pmin (W)`), minp)
  df2$`Avg_power (W)` <- append(na.omit(df2$`Avg_power (W)`), avrp)
  df2$`Pdrop (W)` <- append(na.omit(df2$`Pdrop (W)`), drop)
  df2$`Pmax_5s (W)` <- append(na.omit(df2$`Pmax_5s (W)`), pp5s)
  df2$`IU (%)` <- append(na.omit(df2$`IU (%)`), iu)
  df2$`Work (kJ)` <- append(na.omit(df2$`Work (kJ)`), totalw/1000)
  df2$`HR_max (BPM)` <- as.numeric(append(na.omit(df2$`HR_max (BPM)`), max(df$Heart.rate..bpm., na.rm =T)))
  df2$`La_max (mmol/l)` <- append(na.omit(df2$`La_max (mmol/l)`), la)
  
  
  
  #### Wingate plots ####
  my_breaks <- seq(0, tail(df$Elapsed.time.total..h.mm.ss.hh.,n=1)+5, by = 5)
  
  
  if(exists("file.list.compar_2.bez")) {
    if (srovnani == "YES" & id %in% file.list.compar_2.bez & tri_graf == "YES") {
      plot1 <- ggplot2::ggplot(df, aes(x = Elapsed.time.total..h.mm.ss.hh., y = Power..W.)) + 
        theme_classic()  + 
        theme(text = element_text(size = 30),
              panel.grid.major = element_line(color = "grey90",
                                              linewidth = 0.1,
                                              linetype = 1),
              panel.grid.minor.y = element_line(color = "grey90",
                                                linewidth = 0.1,
                                                linetype = 1),
              axis.line = element_line(colour = "black",
                                       linewidth = 1.5)) +
        geom_line(data = compare.wingate_2, aes(y = y3, color = "darkmagenta", linetype = "dashed"), lwd = 0.5) +
        geom_line(data = compare.wingate, aes(y = y2, color = "orange", linetype = "solid"), lwd = 0.9) + 
        geom_line(aes(y = mean(Power..W.), color= "blue", linetype = "dashed"), lwd = 0.9) +
        geom_line(aes(color = "grey", linetype = "dashed")) +
        geom_line(data = df, aes(y=y, x = Elapsed.time.total..h.mm.ss.hh., color = "black", linetype = "solid"), lwd = 0.9) + 
        scale_linetype_manual(values = c('dashed', 'solid')) +
        scale_colour_identity(guide = "legend", breaks = c("black", "grey", "blue", "orange", "darkmagenta"), labels = c("black"=paste("vyhlazená data,", datum_mer, sep = " "), "blue"="průměr měření","grey"="hrubá data","orange"=datum_mer, "darkmagenta"=format(as.Date(antropo$Date_measurement[length(antropo$Date_measurement)-2]), "%d/%m/%Y"))) + guides(color = guide_legend(override.aes = list(linetype = c("solid", "dashed", "dashed", "solid", "dashed")))) + 
        guides(linetype = "none") + 
        xlab(expression("Čas (s)")) + 
        ylab("Výkon (W)") + 
        ggtitle("Anaerobní Wingate test - 30 s")  + 
        labs(colour="") +
        scale_x_continuous(breaks = my_breaks)  
    }
    else if (tri_graf == "NO" | !(id %in% file.list.compar_2.bez) & id %in% file.list.compar.bez) {
      plot1 <- ggplot2::ggplot(df, aes(x = Elapsed.time.total..h.mm.ss.hh., y = Power..W.)) + 
        theme_classic()  + 
        theme(text = element_text(size = 30),
              panel.grid.major = element_line(color = "grey90",
                                              linewidth = 0.1,
                                              linetype = 1),
              panel.grid.minor.y = element_line(color = "grey90",
                                                linewidth = 0.1,
                                                linetype = 1),
              axis.line = element_line(colour = "black",
                                       linewidth = 1.5)) +
        geom_line(data = compare.wingate, aes(y = y2, color = "orange", linetype = "solid"), lwd = 0.9) + 
        geom_line(aes(y = mean(Power..W.), color= "blue", linetype = "dashed"), lwd = 0.9) +
        geom_line(aes(color = "grey", linetype = "dashed")) +
        geom_line(data = df, aes(y=y, x = Elapsed.time.total..h.mm.ss.hh., color = "black", linetype = "solid"), lwd = 0.9) + 
        scale_linetype_manual(values = c('dashed', 'solid')) +
        scale_colour_identity(guide = "legend", breaks = c("black", "grey", "blue", "orange"), labels = c("black"=paste("vyhlazená data,", datum_mer, sep = " "), "blue"="průměr měření","grey"="hrubá data","orange"=format(as.Date(antropo$Date_measurement[length(antropo$Date_measurement)-1]), "%d/%m/%Y"))) + guides(color = guide_legend(override.aes = list(linetype = c("solid", "dashed", "dashed", "solid")))) + 
        guides(linetype = "none") + 
        xlab(expression("Čas (s)")) + 
        ylab("Výkon (W)") + 
        ggtitle("Anaerobní Wingate test - 30 s")  + 
        labs(colour="") +
        scale_x_continuous(breaks = my_breaks)  
    } 
    else if (exists("file.list.compar.bez") & !exists("file.list.compar_2.bez")) {
      if (srovnani == "YES" & id %in% file.list.compar.bez) {
        plot1 <- ggplot2::ggplot(df, aes(x = Elapsed.time.total..h.mm.ss.hh., y = Power..W.)) + 
          theme_classic()  + 
          theme(text = element_text(size = 30),
                panel.grid.major = element_line(color = "grey90",
                                                linewidth = 0.1,
                                                linetype = 1),
                panel.grid.minor.y = element_line(color = "grey90",
                                                  linewidth = 0.1,
                                                  linetype = 1),
                axis.line = element_line(colour = "black",
                                         linewidth = 1.5)) +
          geom_line(aes(color = "grey", linetype = "dashed")) +
          geom_line(data = compare.wingate, aes(y = y2, color = "orange", linetype = "solid"), lwd = 0.9) + 
          geom_line(data = df, aes(y=y, x = Elapsed.time.total..h.mm.ss.hh., color = "black", linetype = "solid"), lwd = 0.9) + 
          geom_line(aes(y = mean(Power..W.), color= "blue", linetype = "dashed"), lwd = 0.9) +
          scale_linetype_manual(values = c('dashed', 'solid')) +
          scale_colour_identity(guide = "legend", labels = c(paste("vyhlazená data,", datum_mer, sep = " "), "průměr měření", "hrubá data", format(as.Date(antropo$Date_measurement[length(antropo$Date_measurement)-1]), "%d/%m/%Y"))) + guides(color = guide_legend(override.aes = list(linetype = c("solid", "solid", "dashed", "solid")))) + 
          guides(linetype = "none") + 
          xlab(expression("Čas (s)")) + 
          ylab("Výkon (W)") + 
          ggtitle("Anaerobní Wingate test - 30 s")  + 
          labs(colour="") +
          scale_x_continuous(breaks = my_breaks)
      }
    } else { 
      plot1 <- ggplot2::ggplot(df, aes(x = Elapsed.time.total..h.mm.ss.hh., y = Power..W.)) + 
        theme_classic()  + 
        theme(text = element_text(size = 30),
              panel.grid.major = element_line(color = "grey90",
                                              linewidth = 0.1,
                                              linetype = 1),
              panel.grid.minor.y = element_line(color = "grey90",
                                                linewidth = 0.1,
                                                linetype = 1),
              axis.line = element_line(colour = "black",
                                       linewidth = 1.5)) +
        geom_line(aes(color = "grey", linetype = "dashed")) +
        geom_line(aes(y=y, color = "black", linetype = "solid"), lwd = 0.9) + 
        geom_line(aes(y = mean(Power..W.), color= "blue", linetype = "dashed"), lwd = 0.9) +
        scale_linetype_manual(values = c('dashed', 'solid')) +
        scale_colour_identity(guide = "legend", labels = c(paste("vyhlazená data,", datum_mer, sep = " "), "průměr měření", "hrubá data")) + guides(color = guide_legend(override.aes = list(linetype = c("solid", "dashed", "dashed")))) + 
        guides(linetype = "none") + 
        xlab(expression("Čas (s)")) + 
        ylab("Výkon (W)") + 
        ggtitle("Anaerobní Wingate test - 30 s")  + 
        labs(colour="") +
        scale_x_continuous(breaks = my_breaks)  
    }
  }
  else {
    plot1 <- ggplot2::ggplot(df, aes(x = Elapsed.time.total..h.mm.ss.hh., y = Power..W.)) + 
      theme_classic()  + 
      theme(text = element_text(size = 30),
            panel.grid.major = element_line(color = "grey90",
                                            linewidth = 0.1,
                                            linetype = 1),
            panel.grid.minor.y = element_line(color = "grey90",
                                              linewidth = 0.1,
                                              linetype = 1),
            axis.line = element_line(colour = "black",
                                     linewidth = 1.5)) +
      geom_line(aes(color = "grey", linetype = "dashed")) +
      geom_line(aes(y=y, color = "black", linetype = "solid"), lwd = 0.9) + 
      geom_line(aes(y = mean(Power..W.), color= "blue", linetype = "dashed"), lwd = 0.9) +
      scale_linetype_manual(values = c('dashed', 'solid')) +
      scale_colour_identity(guide = "legend", labels = c(paste("vyhlazená data,", datum_mer, sep = " "), "průměr měření", "hrubá data")) + guides(color = guide_legend(override.aes = list(linetype = c("solid", "dashed", "dashed")))) + 
      guides(linetype = "none") + 
      xlab(expression("Čas (s)")) + 
      ylab("Výkon (W)") + 
      ggtitle("Anaerobní Wingate test - 30 s")  + 
      labs(colour="") +
      scale_x_continuous(breaks = my_breaks)  
  }
  
  
  
  #### Export of Reporting tables ####
  table1 <- df1 %>% gt() %>% 
    tab_header(title = fullname, subtitle = paste("Datum měření: ",datum_mer, sep = "")) %>% 
    cols_label(V2 = "") %>% sub_missing(columns = everything(),
                                        rows = everything(),
                                        missing_text = "")   %>%
    tab_style(
      style = cell_text(weight = "bold"),
      locations = cells_column_labels()) %>%
    tab_style(
      style = cell_text(weight = "bold"),
      locations = cells_body(
        columns = c("Antropometrie", "Wingate test") 
      )
    ) %>%
    tab_style(
      style = "padding-right:30px",
      locations = cells_column_labels()
    )   %>%
    tab_style(
      style = "padding-right:30px",
      locations = cells_body()) %>% 
    tab_options(
      table.width = pct(c(100)),
      container.width = 1500,
      table.font.size = px(18L),
      heading.padding = pct(0)) %>%
    gtExtras::gt_add_divider(columns = "V2", style = "solid", color = "#808080") %>% 
    opt_table_lines(extent = "none") %>%
    tab_style(
      style = cell_borders(
        sides = "top",
        color = "#808080",
        style = "solid"
      ),
      locations = 
        cells_body(
          rows = 1
        )
    ) %>% tab_style(
      style = cell_text(align = "right"),
      locations = cells_column_labels(columns = c("Absolutní", "Relativní (TH)", "Relativní (ATH)", "Pětivteřinový průměr"))) %>% 
    cols_align(
      align = c("center"),
      columns = c("V2", "Absolutní", "Relativní (TH)", "Relativní (ATH)", "Pětivteřinový průměr")
    ) %>% tab_style(
      style = cell_text(align = "center"),
      locations = cells_body(columns = c("V2", "Absolutní", "Relativní (TH)", "Relativní (ATH)", "Pětivteřinový průměr")))
  
  
  #History table
  table2 <- df2 %>% gt() %>% tab_header(title = "Historie měření") %>% 
    cols_align(
      align = c("center")) %>% 
    tab_options(
      table.width = pct(c(100)),
      container.width = 1500,
      table.font.size = px(18L)) %>% 
    cols_label(
      Date_meas. = "Datum",
      Weight = "Váha (kg)",
      ATH = "ATH (kg)",
      `Fat (%)`  = "Tuk (%)",
      `Avg_power (W)` = "Průměr P (W)",
      `Pdrop (W)` = "Pokles (W)",
      `Pmax_5s (W)` = "Pmax 5s (W)",
      `Pmax/kg (W/kg)` = "Pmax (W/kg)",
      `IU (%)` = "IU (%)",
      `HR_max (BPM)` = "TF (BPM)",
      `La_max (mmol/l)` = "LA (mmol/l)",
      `Pmax (W)` = "Pmax (W)",
      `Pmin (W)` = "Pmin (W)",
      `Work (kJ)` = "Práce (kJ)"
    ) %>% 
    cols_width(c("Fat (%)", "ATH") ~ pct(6)) %>% 
    cols_width(c("Avg_power (W)", "La_max (mmol/l)") ~ pct(10.5)) %>% 
    cols_width(c("Pmax_5s (W)", "Pdrop (W)", "Pmax (W)", "Weight", "Work (kJ)", "IU (%)") ~ pct(8)) %>% 
    cols_width(c("Date_meas.") ~ pct(7)) %>% cols_width(c("Pmin (W)", "Pmax (W)", "HR_max (BPM)") ~ pct(7)) %>%  sub_missing(columns = everything(), rows = everything(), missing_text = "") 
  
  
  #tabulka spiro
  if (dotaz_spiro == "YES") {
    table3 <- df3 %>% 
      gt() %>% 
      tab_header(
        title = fullname, 
        subtitle = paste("Datum měření: ", datum_mer, sep = "")
      ) %>% 
      cols_label(Hodnota = "") %>% 
      cols_label(V3 = "")  %>% 
      sub_missing(columns = everything(), rows = everything(), missing_text = "")  %>%
      tab_style(
        style = cell_text(weight = "bold"),
        locations = cells_column_labels()
      ) %>%
      tab_style(
        style = cell_text(weight = "bold"),
        locations = cells_body(
          columns = c("Antropometrie", "Spiroergometrie")
        )
      ) %>%
      tab_style(
        style = cell_text(weight = "bold"),
        locations = cells_body(
          rows = 7,
          columns = everything()  # Apply bold style to all columns in row 9
        )
      ) %>%
      tab_style(
        style = "padding-right:30px",
        locations = cells_column_labels()
      ) %>%
      tab_style(
        style = "padding-right:30px",
        locations = cells_body()
      ) %>% 
      tab_options(
        table.width = pct(c(100)),
        container.width = 1600,
        container.height = 670,
        table.font.size = px(18L),
        heading.padding = pct(0)
      ) %>%
      gtExtras::gt_add_divider(columns = "V3", style = "solid", color = "#808080") %>% 
      opt_table_lines(extent = "none") %>%
      tab_style(
        style = cell_borders(
          sides = "top",
          color = "#808080",
          style = "solid"
        ),
        locations = cells_body(
          rows = c(1,8)
        )
      )  %>% 
      tab_style(
        style = cell_text(align = "right"),
        locations = cells_column_labels(columns = c("Absolutní", "Relativní (TH)", "% z nál. hodnoty", "Relativní (ATH)"))
      ) %>% 
      cols_align(
        align = c("center"),
        columns = c("V3", "Absolutní", "Relativní (TH)", "% z nál. hodnoty", "Relativní (ATH)")
      ) %>% 
      tab_style(
        style = cell_text(align = "center"),
        locations = cells_body(columns = c("V3", "Absolutní", "Relativní (TH)", "% z nál. hodnoty", "Relativní (ATH)"))
      ) %>%
      tab_style(
        style = "padding-bottom:20px",
        locations = cells_body(rows = 6)
      )
    
    
    table4  <- tz %>% gt() %>% 
      tab_header(title = "Tréninkové zóny") %>% 
      cols_label(Zóna = "")  %>% sub_missing(columns = everything(),
                                             rows = everything(),
                                             missing_text = "") %>% 
      tab_options(
        table.width = pct(c(100)),
        container.width = 300,
        container.height = 600,
        table.font.size = px(18L),
        heading.padding = pct(0))
    
    table4 %>% gtsave(
      "t4.png",
      vwidth = 350,  
      vheight = 800  
    )
    
    png("p2.png", width = 1500, height = 600)
    plot(combined_plot)
    dev.off()
    
    table3 %>%
      gtsave("t3.png", vwidth = 1700, vheight = 1000) 
  }
  
  table2 %>%
    gtsave("t2.png", vwidth = 1500)  
  
  table1 %>%
    gtsave("t1.png", vwidth = 1500, vheight = 1000)
  
  png("p1.png", width = 2000, height = 600)
  plot(plot1)
  dev.off()
  
  
  #### Radarchart ####
  
  hg_anckg <- round(((an.cap/sports_data$ANC[sports_data$Sport==sport])*100),1)
  hg_ppkg <- round((((pp/vaha)/sports_data$Pmax[sports_data$Sport==sport])*100),1)
  
  
  
  if (!is.na(sj)) {
    hg_sj <- round(((sj/sports_data$SJ[sports_data$Sport==sport])*100),1)
    data_radar <- data.frame(
      `Anaerobní kapacita (J/kg) %` = hg_anckg,
      `Výkon Wingate  (W/kg) %` = hg_ppkg,
      `Squat Jump (cm) %`= hg_sj,
      check.names = FALSE
    )} else {
      data_radar <- data.frame(
        `Anaerobní kapacita (J/kg) %` = hg_anckg,
        `Výkon Wingate  (W/kg) %` = hg_ppkg,
        check.names = FALSE
      )
    }
  
  
  
  if (dotaz_spiro == "YES") {
    hg_Vo2max <- round(((VO2_kg / sports_data$VO2max[sports_data$Sport==sport]) * 100), 1)
    hg_VO2vykon <- round(((vykon_kg / sports_data$PowerVO2[sports_data$Sport==sport]) * 100), 1)
    data_radar <- cbind( `VO2Max (ml/min/kg) %` = hg_Vo2max, data_radar, `Výkon VO2Max (W/kg) %` = hg_VO2vykon)
  }
  
  
  max_values <- rep(100, ncol(data_radar))
  min_values <- rep(0, ncol(data_radar))
  
  data_radar <- rbind(max_values, min_values, data_radar)
  
  
  
  radarchart2 <- function(df, axistype=0, seg=4, pty=16, pcol=1:8, plty=1:6, plwd=1,
                          pdensity=NULL, pangle=45, pfcol=NA, cglty=3, cglwd=1,
                          cglcol="navy", axislabcol="blue", title="", maxmin=TRUE,
                          na.itp=TRUE, centerzero=FALSE, vlabels=NULL, vlcex=NULL,
                          caxislabels=NULL, calcex=NULL,
                          paxislabels=NULL, palcex=NULL, ...) {
    if (!is.data.frame(df)) { cat("The data must be given as dataframe.\n"); return() }
    if ((n <- length(df))<3) { cat("The number of variables must be 3 or more.\n"); return() }
    if (maxmin==FALSE) { # when the dataframe does not include max and min as the top 2 rows.
      dfmax <- apply(df, 2, max)
      dfmin <- apply(df, 2, min)
      df <- rbind(dfmax, dfmin, df)
    }
    plot(c(-1.2, 1.2), c(-1.2, 1.2), type="n", frame.plot=FALSE, axes=FALSE, 
         xlab="", ylab="", main=title, cex.main=2, asp=1, ...) # define x-y coordinates without any plot
    theta <- seq(90, 450, length=n+1)*pi/180
    theta <- theta[1:n]
    xx <- cos(theta)
    yy <- sin(theta)
    CGap <- ifelse(centerzero, 0, 1)
    for (i in 0:seg) { # complementary guide lines, dotted navy line by default
      polygon(xx*(i+CGap)/(seg+CGap), yy*(i+CGap)/(seg+CGap), lty=cglty, lwd=cglwd, border=cglcol)
      if (axistype==1|axistype==3) CAXISLABELS <- paste(i/seg*100,"(%)")
      if (axistype==4|axistype==5) CAXISLABELS <- sprintf("%3.2f",i/seg)
      if (!is.null(caxislabels)&(i<length(caxislabels))) CAXISLABELS <- caxislabels[i+1]
      if (axistype==1|axistype==3|axistype==4|axistype==5) {
        if (is.null(calcex)) text(-0.05, (i+CGap)/(seg+CGap), CAXISLABELS, col=axislabcol) else
          text(-0.05, (i+CGap)/(seg+CGap), CAXISLABELS, col=axislabcol, cex=calcex)
      }
    }
    if (centerzero) {
      arrows(0, 0, xx*1, yy*1, lwd=cglwd, lty=cglty, length=0, col=cglcol)
    }
    else {
      arrows(xx/(seg+CGap), yy/(seg+CGap), xx*1, yy*1, lwd=cglwd, lty=cglty, length=0, col=cglcol)
    }
    PAXISLABELS <- df[1,1:n]
    if (!is.null(paxislabels)) PAXISLABELS <- paxislabels
    if (axistype==2|axistype==3|axistype==5) {
      if (is.null(palcex)) text(xx[1:n], yy[1:n], PAXISLABELS, col=axislabcol) else
        text(xx[1:n], yy[1:n], PAXISLABELS, col=axislabcol, cex=palcex)
    }
    VLABELS <- colnames(df)
    if (!is.null(vlabels)) VLABELS <- vlabels
    for (i in 1:n) {
      if (i == 1) {
        x_offset <- xx[i] * 0.1  # Example offset value for the top value
        y_offset <- yy[i] * 0.01
      } else if (i == 2) {
        x_offset <- xx[i] * 0.4  # Example offset value for the 2nd and 4th columns
        y_offset <- yy[i] * 0.3
      } else if (i == 4) {
        x_offset <- xx[i] * 0.4  # Example offset value for the 2nd and 4th columns
        y_offset <- yy[i] * 0.3
      } else if (i == 3 || i == 5) {
        x_offset <- xx[i] * 0.1  # Example offset value for the 3rd and 5th columns
        y_offset <- yy[i] * 0.03
      } else {
        x_offset <- 0
        y_offset <- 0
      }
      if (is.null(vlcex)) text(xx[i] * 1.2 + x_offset, yy[i] * 1.2 + y_offset, VLABELS[i]) else
        text(xx[i] * 1.2 + x_offset, yy[i] * 1.2 + y_offset, VLABELS[i], cex=vlcex, font = 2)
    }
    series <- length(df[[1]])
    SX <- series-2
    if (length(pty) < SX) { ptys <- rep(pty, SX) } else { ptys <- pty }
    if (length(pcol) < SX) { pcols <- rep(pcol, SX) } else { pcols <- pcol }
    if (length(plty) < SX) { pltys <- rep(plty, SX) } else { pltys <- plty }
    if (length(plwd) < SX) { plwds <- rep(plwd, SX) } else { plwds <- plwd }
    if (length(pdensity) < SX) { pdensities <- rep(pdensity, SX) } else { pdensities <- pdensity }
    if (length(pangle) < SX) { pangles <- rep(pangle, SX)} else { pangles <- pangle }
    if (length(pfcol) < SX) { pfcols <- rep(pfcol, SX) } else { pfcols <- pfcol }
    for (i in 3:series) {
      xxs <- xx
      yys <- yy
      scale <- CGap/(seg+CGap)+(df[i,]-df[2,])/(df[1,]-df[2,])*seg/(seg+CGap)
      if (sum(!is.na(df[i,]))<3) { cat(sprintf("[DATA NOT ENOUGH] at %d\n%g\n",i,df[i,])) # for too many NA's (1.2.2012)
      } else {
        for (j in 1:n) {
          if (is.na(df[i, j])) { # how to treat NA
            if (na.itp) { # treat NA using interpolation
              left <- ifelse(j>1, j-1, n)
              while (is.na(df[i, left])) {
                left <- ifelse(left>1, left-1, n)
              }
              right <- ifelse(j<n, j+1, 1)
              while (is.na(df[i, right])) {
                right <- ifelse(right<n, right+1, 1)
              }
              xxleft <- xx[left]*CGap/(seg+CGap)+xx[left]*(df[i,left]-df[2,left])/(df[1,left]-df[2,left])*seg/(seg+CGap)
              yyleft <- yy[left]*CGap/(seg+CGap)+yy[left]*(df[i,left]-df[2,left])/(df[1,left]-df[2,left])*seg/(seg+CGap)
              xxright <- xx[right]*CGap/(seg+CGap)+xx[right]*(df[i,right]-df[2,right])/(df[1,right]-df[2,right])*seg/(seg+CGap)
              yyright <- yy[right]*CGap/(seg+CGap)+yy[right]*(df[i,right]-df[2,right])/(df[1,right]-df[2,right])*seg/(seg+CGap)
              if (xxleft > xxright) {
                xxtmp <- xxleft; yytmp <- yyleft;
                xxleft <- xxright; yyleft <- yyright;
                xxright <- xxtmp; yyright <- yytmp;
              }
              xxs[j] <- xx[j]*(yyleft*xxright-yyright*xxleft)/(yy[j]*(xxright-xxleft)-xx[j]*(yyright-yyleft))
              yys[j] <- (yy[j]/xx[j])*xxs[j]
            } else { # treat NA as zero (origin)
              xxs[j] <- 0
              yys[j] <- 0
            }
          }
          else {
            xxs[j] <- xx[j]*CGap/(seg+CGap)+xx[j]*(df[i, j]-df[2, j])/(df[1, j]-df[2, j])*seg/(seg+CGap)
            yys[j] <- yy[j]*CGap/(seg+CGap)+yy[j]*(df[i, j]-df[2, j])/(df[1, j]-df[2, j])*seg/(seg+CGap)
          }
        }
        if (is.null(pdensities)) {
          polygon(xxs, yys, lty=pltys[i-2], lwd=plwds[i-2], border=pcols[i-2], col=pfcols[i-2])
        } else {
          polygon(xxs, yys, lty=pltys[i-2], lwd=plwds[i-2], border=pcols[i-2], 
                  density=pdensities[i-2], angle=pangles[i-2], col=pfcols[i-2])
        }
        points(xx*scale, yy*scale, pch=ptys[i-2], col=pcols[i-2])
        
        ## Line added to add textvalues to points
        for (j in 1:n) {
          if (j == 1) {
            if (df[i, j] >= 100) {
              color <- "green4"
            } else if (df[i, j] >= 90) {
              color <- "orange"
            } else {
              color <- "red"
            }
          } else if (j == 2) {
            if (df[i, j] >= 100) {
              color <- "green4"
            } else if (df[i, j] >= 90) {
              color <- "orange"
            } else {
              color <- "red"
            }
          } else if (j == 3) {
            if (df[i, j] >= 100) {
              color <- "green4"
            } else if (df[i, j] >= 90) {
              color <- "orange"
            } else {
              color <- "red"
            }
          } else if (j == 4) {
            if (df[i, j] >= 105) {
              color <- "green4"
            } else if (df[i, j] >= 95) {
              color <- "orange"
            } else {
              color <- "red"
            }
          } else if (j == 5) {
            if (df[i, j] >= 104) {
              color <- "green4"
            } else if (df[i, j] >= 95) {
              color <- "orange"
            } else {
              color <- "red"
            }
          } else {
            color <- "black"  # Default color for any additional values
          }
          text(xx[j]*scale[j]*0.8, yy[j]*scale[j]*0.8, df[i, j], cex = 2, font = 2,col = color)
        }
      }
    }
  }
  
  
  if (!is.na(sj) | dotaz_spiro == "YES")  {
    png("radar.png", width = 1124, height = 797)
    
    
    # Create radar chart
    
    radarchart2(
      data_radar,
      axistype = 1,
      pcol = rgb(0.2, 0.5, 0.5, 0.9),
      pfcol = NA, 
      cglcol = "grey",
      cglty = 1,
      axislabcol = "grey",
      caxislabels = seq(0, 100, 25),
      cglwd = 0.8,
      vlcex = 1,
      title = "Rozložení parametrů"
    )
    
    dev.off()
    
  }
  if (!is.na(sj) | dotaz_spiro == "YES") {
    rmarkdown::render("export_NEOTVIRAT.Rmd", params = list(dotaz_spiro = dotaz_spiro), output_file = paste(id, ".pdf", sep= ""), clean = T, quiet = F)
  } else {
    rmarkdown::render("export_NEOTVIRAT_wingate.Rmd", params = list(dotaz_spiro = dotaz_spiro), output_file = paste(id, ".pdf", sep= ""), clean = T, quiet = F)
  }
  
  fs::file_move(path = paste(wd, "/", id, ".pdf", sep = ""), new_path = paste(wd, "/reporty/", id, ".pdf", sep = ""))
  fs::file_move(path = paste(wd, "/", id, ".tex", sep = ""), new_path = paste(wd, "/vymazat/", id, ".tex", sep = ""))
  file.rename(from = paste(wd, "/", id, "_files", sep = ""), to = paste(wd, "/vymazat/", id, "_files", sep = ""))
  
  
  
  df2$Name <- fullname.rev
  df2 <- df2 %>% relocate(Name, .before = Date_meas.)
  df2$Age <- ifelse(!is.na(age), floor(age), 
                    ifelse(!is.na(s.age), floor(s.age), NA))
  df2 <- df2 %>% relocate(Age, .after = Date_meas.)
  df2$Turns <- round(tail(df$Turns.number..Nr.,1),0)
  
  if (srovnani == "YES") {
    if (id %in% file.list.compar.bez) {
      df2$Turns[length(df2$Turns)-1] <- round(tail(compare.wingate$Turns.number..Nr.,1),0) 
    }
  }
  
  df2 <- df2 %>% relocate(Turns, .after = ATH)
  df2$Sport <- sport
  df2 <- df2 %>% relocate(Sport, .after = Date_meas.)
  df2$Team <- team
  df2 <- df2 %>% relocate(Team, .after = Sport)
  df2$`HR_max (BPM)` <- as.numeric(df2$`HR_max (BPM)`)
  df2$`Pmax/kgATH (W/kg)` <- df2$`Pmax (W)`/df2$ATH
  df2$id <- id
  df2 <- df2 %>% relocate(id, .before = Name)
  df2$`Anc/kg (J/kg)` <- (df2$`Work (kJ)`*1000)/df2$Weight
  
  
  
  if(!exists("databaze")) {
    databaze <- data.frame()
    databaze <- df2
  } else {
    databaze <- rbind(databaze, df2)
  }
  
  dotaz_spiro <- dotaz_spiro_original
}

# end of MAIN for loop


if (exists("spatne.wingate")) {
  if (!rapportools::is.empty(spatne.wingate)) {
    print(paste("Nevyhodnocene soubory (chyba nazvu):", spatne.wingate, sep = " "))
  }
}

if (!dir.exists(paste(wd, "/vysledky/spiro/", sep=""))) {
  dir.create(paste(wd, "/vysledky/spiro/", sep=""))
}

#### Report printing ####
if (dotaz_spiro == "YES") {
  df4 <- df4[rev(order(df4$`VO2_kg/l`)), ]
  writexl::write_xlsx(df4, paste(wd, "/vysledky/spiro/spiro_vysledek", "_", team, "_", format(Sys.time(), "_%d-%m %H-%M"), ".xlsx", sep = ""), col_names = T, format_headers = T)
}

writexl::write_xlsx(databaze, paste(wd, "/vysledky/vysledek", "_", team, "_", format(Sys.time(), "_%d-%m %H-%M"), ".xlsx", sep = ""), col_names = T, format_headers = T)

if (srovnani == "YES" & length(chybne.pocty) > 0) {
  print("Srovnání nevyhodnoceno (chybí antropa):")
  print(chybne.pocty)
}

input <- readline("Nahrat vysledky do celkove databaze? (A/N): ")
if(input == "A") {
  zmeny.cas <- c()
  for (i in 1:length(file.list.dat)) {
    zmeny.cas <- append(zmeny.cas, file.info(paste(database.path, "/", file.list.dat[i], sep=""))$mtime)
  }
  
  
  datab_old <- readxl::read_excel(paste(database.path, "/", file.list.dat[which.min(abs(zmeny.cas-Sys.time()))], sep=""))
  datab_new <- rbind(datab_old, databaze)
  
  writexl::write_xlsx(datab_new, paste(database.path, "/databaze", "_", format(Sys.time(), "_%d-%m %H-%M"), ".xlsx", sep = ""), col_names = T, format_headers = T)  
}



# fs::file_move(path = paste(wd, "/", "out.rda", sep = ""), new_path = paste(wd, "/vymazat/", "out.rda", sep = ""))
# fs::file_move(path = paste(wd, "/", "t1.png", sep = ""), new_path = paste(wd, "/vymazat/", "t1.png", sep = ""))
# fs::file_move(path = paste(wd, "/", "t2.png", sep = ""), new_path = paste(wd, "/vymazat/", "t2.png", sep = ""))
# fs::file_move(path = paste(wd, "/", "p1.png", sep = ""), new_path = paste(wd, "/vymazat/", "p1.png", sep = ""))