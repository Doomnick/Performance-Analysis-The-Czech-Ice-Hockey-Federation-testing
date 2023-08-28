# Wingate test script

# cesta k vychozi slozce, ktera obsahuje podslozky - wingate(.txt exporty z wattbike); antropometrie(.xls nebo .xlsx soubor s antropometrii obsahujici list "Data_Sheet") - vsechny nazvy jsou senzitivni na velikost pismen
wd <- "C:/Users/Dominik Kolinger/Documents/R/Wingate"


########## PRED SPUSTENIM CTI SOUBOR READ_ME! #################################################
################################################################################################

############### DALSI PARAMETRY NEMENIT! #########################################################

setwd(wd)
#instalace potrebnych balicku
instalace <- function() {
  # find all source code files in (sub)folders
  files <- list.files(pattern='[.](R|rmd)$', all.files=T, recursive=T, full.names = T, ignore.case=T)
  
  # read in source code
  code=unlist(sapply(files, scan, what = 'character', quiet = TRUE))
  
  # retain only source code starting with library
  code <- code[grepl('^library', code, ignore.case=T)]
  code <- gsub('^library[(]', '', code)
  code <- gsub('[)]', '', code)
  code <- gsub('^library$', '', code)
  
  # retain unique packages
  uniq_packages <- unique(code)
  
  # kick out "empty" package names
  uniq_packages <- uniq_packages[!uniq_packages == '']
  
  # order alphabetically
  uniq_packages <- uniq_packages[order(uniq_packages)]
  
  cat('Required packages: \n')
  cat(paste0(uniq_packages, collapse= ', '),fill=T)
  cat('\n\n\n')
  
  # retrieve list of already installed packages
  installed_packages <- installed.packages()[, 'Package']
  
  # identify missing packages
  to_be_installed <- setdiff(uniq_packages, installed_packages)
  
  if (length(to_be_installed)==length(uniq_packages)) cat('Vsechny balicky musi byt nainstalovany - instaluji.\n')
  if (length(to_be_installed)>0) cat('Instaluji chybejici balicky.\n')
  if (length(to_be_installed)==0) cat('Vsechny balicky jsou jiz nainstalovany!\n')
  
  # install missing packages
  if (length(to_be_installed)>0) install.packages(to_be_installed, repos = 'https://cloud.r-project.org')
}

instalace()

sport <- readline("Sport?: ")
team <- readline("Tym?: ")
srovnani <- readline("Pridat srovnani s predchozimi vysledky? (A/N): ")


file.path <- paste(wd, "/wingate", sep = "") #slozka s wingate
file.path.an <- paste(wd, "/antropometrie", sep = "") #slozka s antropometrii
report.path <- paste(wd, "/reporty", sep = "") #slozka pro reporty
database.path <- paste(wd, "/databaze", sep = "") #slozka pro databazi
list.files(path = file.path, pattern = '*.txt')
if (srovnani == "A") {
  comparison.path <- paste(wd, "/wingate/srovnani", sep = "")
  file.list.compar <- list.files(path = comparison.path, pattern = '*.txt')
}

#pracovni slozka
file.list <- list.files(path = file.path, pattern = '*.txt')
file.list.an <- list.files(path = file.path.an, pattern = c('*.xls'))
file.list.dat <- list.files(path = database.path, pattern = c('*.xls'))
file.list.bez <- sort(tools::file_path_sans_ext(file.list))



#analyza antropy
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


kontrola <- sapply(list(file.list.an.bez), FUN = identical, file.list.bez)
file.list.compar.bez <- c()

if (srovnani == "A") {
file.list.compar.bez <- sort(tools::file_path_sans_ext(file.list.compar))  
kontrola.s <- sapply(list(file.list.compar.bez), FUN = identical, file.list.bez)
}

duplikaty<- duplicated(file.list.an.bez)

if ("TRUE" %in% duplikaty) {
  print("Duplikaty v souboru s antropometrii:")
  print(file.list.an.bez[which(duplikaty=="TRUE")])
  if (srovnani != "A") {
    stop("EXPORT PRERUSEN - opravte duplikaty v souboru antropometrie!")
  }
    else if (srovnani == "A") {
      print(paste("Nalezeno celkem", length(which(duplikaty=="TRUE")), "duplikátù", sep = " "))
      duplikaty_check <- readline("Pokracovat v exportu? (A/N): ")
      if (duplikaty_check == "N") {
        stop("OPERACE PRERUSENA UZIVATELEM")
      }
    }
    
}


if (srovnani != "A") {
if ("FALSE" %in% kontrola & !pracma::isempty(setdiff(file.list.bez, file.list.an.bez))) {
  cat('\n\n')
  print("Nazvy Wingate jsou rozdilne oproti Antropometrii - ZKONTROLUJ:")
  cat('\n')
  if (!pracma::isempty(setdiff(file.list.bez, file.list.an.bez))) {print(setdiff(file.list.bez, file.list.an.bez))}
  input <- readline("Chcete pokracovat v exportu? (A/N): ")
  if (input == "N") {
    stop("OPERACE PRERUSENA UZIVATELEM")}
    else if (input == "A") {
      spatne.wingate <- setdiff(file.list.bez, file.list.an.bez)
      spatne.wingate <- paste(spatne.wingate, ".txt", sep = "")
      file.list <-  file.list[! file.list %in% spatne.wingate] 
    }
}
} else if (srovnani == "A") {
  if ("FALSE" %in% kontrola || "FALSE" %in% kontrola.s) {
    if ("FALSE" %in% kontrola & !pracma::isempty(setdiff(file.list.bez, file.list.an.bez))) {
    cat('\n\n')
    print("Nazvy Wingate jsou rozdilne oproti Antropometrii - ZKONTROLUJ:")
    cat('\n')
    if (!pracma::isempty(setdiff(file.list.bez, file.list.an.bez))) {print(setdiff(file.list.bez, file.list.an.bez))}
    }
    if ("FALSE" %in% kontrola.s) {
      cat('\n\n')
      print("Nazvy Wingate jsou rozdilne (chybi) oproti Srovnani - ZKONTROLUJ:")
      cat('\n')
      if (!pracma::isempty(setdiff(file.list.bez, file.list.compar.bez))) {print(setdiff(file.list.bez, file.list.compar.bez))}  
    }
    input <- readline("Chcete pokracovat v exportu? (A/N): ")
    if (input == "N") {
      stop("OPERACE PRERUSENA UZIVATELEM")
    } else if (input == "A") {
      if ("FALSE" %in% kontrola) {
        spatne.wingate <- setdiff(file.list.bez, file.list.an.bez)
        spatne.wingate <- paste(spatne.wingate, ".txt", sep = "")
        file.list <-  file.list[! file.list %in% spatne.wingate]
      } else if ("FALSE" %in% kontrola.s) {
        spatne.compar <- setdiff(file.list.bez, file.list.compar.bez)
      }
    }
  }
  
}

library(finalfit)
library(dplyr)
library(ggplot2)
library(lubridate)
library(gt)
library(cowplot)
library(tidyverse)
library(tinytex)

#smycka pro export
for(i in 1:length(file.list)) {
# data Wingate
if (exists("df2")) {
  rm(df2)
}
id <- tools::file_path_sans_ext(file.list[i])
writeLines(iconv(readLines(paste(file.path, "/", file.list[i], sep = "")), from = "ANSI_X3.4-1986", to = "UTF8"), 
           file(paste(file.path, "/", file.list[i], sep = ""), encoding="UTF-8"))
df <- read.delim(paste(file.path, "/", file.list[i], sep = ""))
if (an.input == "solo") {
file.an <- file.list.an[which(file.list.an.bez == id)]
antropo <- readxl::read_excel(paste(file.path.an, "/", file.an, sep=""), sheet = "Data_Sheet")
if (srovnani == "A") {
  antropo <- antropo[order(as.Date(antropo$Date_measurement, format="%d/%m/%Y")),]
}
}

if (an.input == "batch") {
  antropo <- readxl::read_excel(paste(file.path.an, "/", file.list.an[1], sep=""), sheet = "Data_Sheet")   
}

#duplikaty v antrope u batch exportu
if (an.input == "batch" & exists("duplikaty_check")) {
  antropo <- antropo %>% 
    group_by(ID) %>% 
    filter(Date_measurement==max(as.Date(Date_measurement)))
}

if(exists("file.list.compar.bez")) {
if (srovnani == "A" & id %in% file.list.compar.bez) {
  writeLines(iconv(readLines(paste(comparison.path, "/", id, ".txt", sep = "")), from = "ANSI_X3.4-1986", to = "UTF8"), 
             file(paste(comparison.path, "/", id, ".txt", sep = ""), encoding="UTF-8"))
  compare.wingate <- read.delim(paste(comparison.path, "/", id, ".txt", sep = ""))
  compare.wingate$Work.total..KJ. <- as.numeric(gsub(",",".", compare.wingate$Work.total..KJ.))
  compare.wingate$Elapsed.time.total..h.mm.ss.hh. <- period_to_seconds(hms(compare.wingate$Elapsed.time.total..h.mm.ss.hh.))
  if (compare.wingate$Elapsed.time.total..h.mm.ss.hh.[1] > 5) {
    compare.wingate$Elapsed.time.total..h.mm.ss.hh. <- compare.wingate$Elapsed.time.total..h.mm.ss.hh.- compare.wingate$Elapsed.time.total..h.mm.ss.hh.[1]
  } else if (tail(compare.wingate$Elapsed.time.total..h.mm.ss.hh.,1) > 31) {
    cas.zacatku <- tail(compare.wingate$Elapsed.time.total..h.mm.ss.hh.,1) - 30
    radek.zacatku <- which.min(abs(compare.wingate$Elapsed.time.total..h.mm.ss.hh. - cas.zacatku)) - 1
    compare.wingate <- compare.wingate[-(1:radek.zacatku),]
  }
  s5 <- head(which(compare.wingate$Elapsed.time.total..h.mm.ss.hh. >= 5), n=1)
  s10 <- head(which(compare.wingate$Elapsed.time.total..h.mm.ss.hh. >= 10), n=1)
  s15 <- head(which(compare.wingate$Elapsed.time.total..h.mm.ss.hh. >= 15), n=1)
  s20 <- head(which(compare.wingate$Elapsed.time.total..h.mm.ss.hh. >= 20), n=1)
  s25 <- head(which(compare.wingate$Elapsed.time.total..h.mm.ss.hh. >= 25), n=1)
  s30 <- length(compare.wingate$Elapsed.time.total..h.mm.ss.hh.)
  radky5s <- round(mean(s5, (s10-s5), (s15-s10), (s20-s15), (s25-s20), (s30-s25)),0)
  compare.wingate$RM5_Power <- zoo::rollmean(compare.wingate$Power..W., k=radky5s, fill = NA, align = "right")
  
  for (i in 1:length(compare.wingate$Power..W.)) {
    compare.wingate$AvP_dopocet[i] <- mean(compare.wingate$Power..W.[1:i] )
  }
}
}


df$Work.total..KJ. <- as.numeric(gsub(",",".", df$Work.total..KJ.))

df$Elapsed.time.total..h.mm.ss.hh. <- period_to_seconds(hms(df$Elapsed.time.total..h.mm.ss.hh.))

#orez Wingate na 30s
if (df$Elapsed.time.total..h.mm.ss.hh.[1] > 5) {
df$Elapsed.time.total..h.mm.ss.hh. <- df$Elapsed.time.total..h.mm.ss.hh.- df$Elapsed.time.total..h.mm.ss.hh.[1]
} else if (tail(df$Elapsed.time.total..h.mm.ss.hh.,1) > 31) {
  cas.zacatku <- tail(df$Elapsed.time.total..h.mm.ss.hh.,1) - 30
  radek.zacatku <- which.min(abs(df$Elapsed.time.total..h.mm.ss.hh. - cas.zacatku)) - 1
  df <- df[-(1:radek.zacatku),]
  if (df$Elapsed.time.total..h.mm.ss.hh.[1] > 4) {
    df$Elapsed.time.total..h.mm.ss.hh. <- df$Elapsed.time.total..h.mm.ss.hh.- df$Elapsed.time.total..h.mm.ss.hh.[1]
  }
  }


s5 <- head(which(df$Elapsed.time.total..h.mm.ss.hh. >= 5), n=1)
s10 <- head(which(df$Elapsed.time.total..h.mm.ss.hh. >= 10), n=1)
s15 <- head(which(df$Elapsed.time.total..h.mm.ss.hh. >= 15), n=1)
s20 <- head(which(df$Elapsed.time.total..h.mm.ss.hh. >= 20), n=1)
s25 <- head(which(df$Elapsed.time.total..h.mm.ss.hh. >= 25), n=1)
s30 <- length(df$Elapsed.time.total..h.mm.ss.hh.)
radky5s <- round(mean(s5, (s10-s5), (s15-s10), (s20-s15), (s25-s20), (s30-s25)),0)
df$RM5_Power <- zoo::rollmean(df$Power..W., k=radky5s, fill = NA, align = "right")

for (i in 1:length(df$Power..W.)) {
  df$AvP_dopocet[i] <- mean(df$Power..W.[1:i] )
}

#vypocty hodnot

if (an.input == "batch") {
  datum.mer <- na.omit(antropo$Date_measurement[antropo$ID == id])
  vaha <- na.omit(antropo$Weight[antropo$ID == id])
  vyska <- na.omit(antropo$Height[antropo$ID == id])
  fat <- round(na.omit(antropo$Fat[antropo$ID == id]),1)
  ath <- round(na.omit(antropo$ATH[antropo$ID == id]),1)
  birth <- na.omit(antropo$Birth[antropo$ID == id])
  fullname <- paste(na.omit(antropo$Name[antropo$ID == id]), na.omit(antropo$Surname[antropo$ID == id]), sep = " ", collapse = NULL)
  fullname.rev <- paste(na.omit(antropo$Surname[antropo$ID == id]), na.omit(antropo$Name[antropo$ID == id]), sep = " ", collapse = NULL)
  datum_nar <- format(as.Date(na.omit(antropo$Birth[antropo$ID == id])),"%d/%m/%Y")
  datum_mer <- format(as.Date(na.omit(antropo$Date_measurement[antropo$ID == id])),"%d/%m/%Y")
  la <- antropo$LA[antropo$ID == id][1]
  age <- round(na.omit(antropo$Age[antropo$ID == id]),2)
} else {
  datum.mer <- na.omit(tail(antropo$Date_measurement, n=1))
  vaha <- na.omit(tail(antropo$Weight, n = 1))
  vyska <- na.omit(tail(antropo$Height, n = 1))
  fat <- round(na.omit(tail(antropo$Fat, n =1)),1)
  ath <- round(na.omit(tail(antropo$ATH, n =1)),1)
  birth <- na.omit(tail(antropo$Birth, n = 1))
  fullname <- paste(na.omit(tail(antropo$Name, n =1)), na.omit(tail(antropo$Surname, n = 1)), sep = " ", collapse = NULL)
  fullname.rev <- paste(tail(antropo$Surname, n = 1), tail(antropo$Name, n = 1), sep = " ", collapse = NULL)
  datum_nar <- format(as.Date(na.omit(tail(antropo$Birth, n =1))),"%d/%m/%Y")
  datum_mer <- format(as.Date(na.omit(tail(antropo$Date_measurement, n =1))),"%d/%m/%Y")
  la <- tail(antropo$LA, n=1)
  age <- round(na.omit(tail(antropo$Age, n =1)),2)
}


pp <- max(df$Power..W.)
minp <- min(df$Power..W.[round((length(df$Power..W.)/2),0):length(df$Power..W.)])
pp5s <- round(max(df$RM5_Power, na.rm = T),1)
minp5s <- round(min(df$RM5_Power, na.rm = T),1)
drop <- pp - minp
iu <- round(((pp5s-minp5s)/pp5s)*100,1)
avrp <- round(mean(df$Power..W.),1)
totalw <- round(mean(df$AvP_dopocet*30),0)
pp5sradek <- which(df$RM5_Power == pp5s)
an.cap <- round(totalw/vaha,2)


#smoothing krivky
vyhlazeno <- smooth.spline(df$Elapsed.time.total..h.mm.ss.hh., df$Power..W., spar = 0.4)
y <- predict(vyhlazeno, newdata = df)
y <- y$y

if(srovnani == "A" & id %in% file.list.compar.bez) {
  vyhlazeno.compare <- smooth.spline(compare.wingate$Elapsed.time.total..h.mm.ss.hh., compare.wingate$Power..W., spar = 0.4)
  y2 <- predict(vyhlazeno.compare, newdata = compare.wingate)
  y2 <- y2$y
}

#tabulka report
columns <- c("Antropometrie", "V2", "Wingate test", "Absolutní", "Relativní (TH)", "Relativní (ATH)", "Pìtivteøinový prùmìr")
df1 = data.frame(matrix(nrow = 8, ncol = length(columns))) 
colnames(df1) <- columns
df1$Antropometrie <-  `length<-`(c("Datum narození", "Vìk", "Výška", "TH - Hmotnost", "Tuk", "ATH - Aktivní tìlesná hmota"), nrow(df1))
df1$`Wingate test` <- `length<-`(c("Maximální výkon", "Minimální výkon", "Pokles výkonu", "Celková práce", "Index Únavy", "Celkový poèet otáèek", "Laktát", "Maximální tepová frekvence"),nrow(df1))

df1[1,2] <- datum_nar
df1[2,2] <- age
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
df1[7,4] <- if (!is.na(la)) {paste(la, "mmol/l", sep = " ")} else NA
df1[8,4] <- max(df$Heart.rate..bpm., na.rm =T)
as.numeric(la)
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


if(exists("file.list.compar.bez")) {
if (srovnani == "A" & id %in% file.list.compar.bez) {
  if (exists("duplikaty_check")) {
  antropo <- readxl::read_excel(paste(file.path.an, "/", file.list.an[1], sep=""), sheet = "Data_Sheet") 
  antropo <- antropo %>% 
    filter(ID==id)
  antropo <- antropo[order(as.Date(antropo$Date_measurement, format="%d/%m/%Y")),]
  }
  df2 <- data.frame()
  df2 <- df2[nrow(df2) + 1,]
  df2$Date_meas. <- format(as.Date(antropo$Date_measurement[length(antropo$Date_measurement)-1]), "%d/%m/%Y")
  df2$Weight <- antropo$Weight[length(antropo$Weight)-1]
  df2$Fat <- round(antropo$Fat[length(antropo$Fat)-1], 1)
  df2$ATH <- round(antropo$ATH[length(antropo$ATH)-1], 1)
  df2$Pmax_W <- max(compare.wingate$Power..W.)
  df2$PP_BW <- round(max(compare.wingate$Power..W.)/antropo$Weight[length(antropo$Weight)-1],2)
  df2$Pmin_W <- min(compare.wingate$Power..W.[round((length(compare.wingate$Power..W.)/2),0):length(compare.wingate$Power..W.)])
  df2$Average_power_W <- round(mean(compare.wingate$Power..W.),1)
  df2$Pdrop_W <- df2$Pmax_W - df2$Pmin_W
  df2$Pmax_5s <- round(max(compare.wingate$RM5_Power, na.rm = T),1)
  df2$IU_pct <- round((max(na.omit(compare.wingate$RM5_Power))-round(min(na.omit(compare.wingate$RM5_Power)),1))/max(na.omit(compare.wingate$RM5_Power))*100,1)
  df2$Work_kJ <- round(mean(compare.wingate$AvP_dopocet*30),0)/1000
  df2$HR_max_BPM <- max(compare.wingate$Heart.rate..bpm., na.rm =T)
  df2$La_max_mmol_l <- antropo$LA[length(antropo$LA)-1]
}
}


#tabulka historie
if (!exists("df2")) {
df2 <- data.frame()
df2 <- df2[nrow(df2) + 1,]
}
if(exists("file.list.compar.bez")) {
if (srovnani == "A" & id %in% file.list.compar.bez) {
  df2[nrow(df2) + 1 , ] <- NA  
}
}
df2$Date_meas. <- append(na.omit(df2$Date_meas.), datum_mer)
df2$Weight <- append(na.omit(df2$Weight), vaha)
df2$Fat <- round(append(na.omit(df2$Fat), fat), 1)
df2$ATH <- round(append(na.omit(df2$ATH), ath), 1)
df2$Pmax_W <- append(na.omit(df2$Pmax_W), pp)
df2$PP_BW <- append(na.omit(df2$PP_BW), round(pp/vaha,2))
df2$Pmin_W <- append(na.omit(df2$Pmin_W), minp)
df2$Average_power_W <- append(na.omit(df2$Average_power_W), avrp)
df2$Pdrop_W <- append(na.omit(df2$Pdrop_W), drop)
df2$Pmax_5s <- append(na.omit(df2$Pmax_5s), pp5s)
df2$IU_pct <- append(na.omit(df2$IU_pct), iu)
df2$Work_kJ <- append(na.omit(df2$Work_kJ), totalw/1000)
df2$HR_max_BPM <- append(na.omit(df2$HR_max_BPM), max(df$Heart.rate..bpm., na.rm =T))
df2$La_max_mmol_l <- append(na.omit(df2$La_max_mmol_l), la)


#graf1
my_breaks <- seq(0, tail(df$Elapsed.time.total..h.mm.ss.hh.,n=1)+5, by = 5)

if(exists("file.list.compar.bez")) {
  if (srovnani == "A" & id %in% file.list.compar.bez) {
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
      scale_colour_identity(guide = "legend", labels = c("vyhlazená data", "prùmìr mìøení", "hrubá data", format(as.Date(antropo$Date_measurement[length(antropo$Date_measurement)-1]), "%d/%m/%Y"))) + guides(color = guide_legend(override.aes = list(linetype = c("solid", "solid", "dashed", "solid")))) + 
      guides(linetype = "none") + 
      xlab(expression("Èas (s)")) + 
      ylab("Výkon (W)") + 
      ggtitle("Anaerobní Wingate test - 30 s")  + 
      labs(colour="") +
      scale_x_continuous(breaks = my_breaks)
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
      scale_colour_identity(guide = "legend", labels = c("vyhlazená data", "prùmìr mìøení", "hrubá data")) + guides(color = guide_legend(override.aes = list(linetype = c("solid", "dashed", "dashed")))) + 
      guides(linetype = "none") + 
      xlab(expression("Èas (s)")) + 
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
    scale_colour_identity(guide = "legend", labels = c("vyhlazená data", "prùmìr mìøení", "hrubá data")) + guides(color = guide_legend(override.aes = list(linetype = c("solid", "dashed", "dashed")))) + 
    guides(linetype = "none") + 
    xlab(expression("Èas (s)")) + 
    ylab("Výkon (W)") + 
    ggtitle("Anaerobní Wingate test - 30 s")  + 
    labs(colour="") +
    scale_x_continuous(breaks = my_breaks)  
}
  

#tabulka report 
table1 <- df1 %>% gt() %>% 
  tab_header(title = fullname, subtitle = paste("Datum mìøení: ", format(as.Date(datum.mer), "%d/%m/%Y"), sep = "")) %>% 
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
    locations = cells_column_labels(columns = c("Absolutní", "Relativní (TH)", "Relativní (ATH)", "Pìtivteøinový prùmìr"))) %>% 
  cols_align(
      align = c("center"),
      columns = c("V2", "Absolutní", "Relativní (TH)", "Relativní (ATH)", "Pìtivteøinový prùmìr")
    ) %>% tab_style(
      style = cell_text(align = "center"),
      locations = cells_body(columns = c("V2", "Absolutní", "Relativní (TH)", "Relativní (ATH)", "Pìtivteøinový prùmìr")))
 

#tabulka historie
table2 <- df2 %>% gt() %>% tab_header(title = "Historie mìøení") %>% 
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
    Fat = "Tuk (%)",
    Average_power_W = "Prùmìr P (W)",
    Pdrop_W = "Pokles (W)",
    Pmax_5s = "Pmax 5s (W)",
    PP_BW = "Pmax/váha",
    IU_pct = "IU (%)",
    HR_max_BPM = "TF (BPM)",
    La_max_mmol_l = "LA (mmol/l)",
    Pmax_W = "Pmax (W)",
    Pmin_W = "Pmin (W)",
    Work_kJ = "Práce (kJ)"
  ) %>% 
  cols_width(c("Fat", "ATH") ~ pct(6)) %>% 
  cols_width(c("Average_power_W", "La_max_mmol_l") ~ pct(10.5)) %>% 
  cols_width(c("Pmax_5s", "Pdrop_W", "PP_BW", "Weight", "Work_kJ", "IU_pct") ~ pct(8)) %>% 
  cols_width(c("Date_meas.") ~ pct(7)) %>% cols_width(c("Pmin_W", "Pmax_W", "HR_max_BPM") ~ pct(7)) %>%  sub_missing(columns = everything(), rows = everything(), missing_text = "") 
                                                           

table2 %>%
  gtsave("t2.png", vwidth = 1500)  

table1 %>%
  gtsave("t1.png", vwidth = 1500, vheight = 1000)

png("p1.png", width = 2000, height = 600)
plot(plot1)
dev.off()

rmarkdown::render("export_NEOTVIRAT.Rmd", output_file = paste(id, ".pdf", sep= ""), clean = T, quiet = F)

fs::file_move(path = paste(wd, "/", id, ".pdf", sep = ""), new_path = paste(wd, "/reporty/", id, ".pdf", sep = ""))
fs::file_move(path = paste(wd, "/", id, ".tex", sep = ""), new_path = paste(wd, "/vymazat/", id, ".tex", sep = ""))
file.rename(from = paste(wd, "/", id, "_files", sep = ""), to = paste(wd, "/vymazat/", id, "_files", sep = ""))

df2$Name <- fullname.rev
df2 <- df2 %>% relocate(Name, .before = Date_meas.)
df2$Age <- age
df2 <- df2 %>% relocate(Age, .after = Date_meas.)
df2$Turns <- round(tail(df$Turns.number..Nr.,1),0)
if (srovnani == "A" & id %in% file.list.compar.bez) {
  df2$Turns[length(df2$Turns)-1] <- round(tail(compare.wingate$Turns.number..Nr.,1),0) 
}

df2 <- df2 %>% relocate(Turns, .after = ATH)
df2$Sport <- sport
df2 <- df2 %>% relocate(Sport, .after = Date_meas.)
df2$Team <- team
df2 <- df2 %>% relocate(Team, .after = Sport)
df2$HR_max_BPM <- as.numeric(df2$HR_max_BPM)
df2$PP_ATH <- df2$Pmax_W/df2$ATH
df2$id <- id
df2 <- df2 %>% relocate(id, .before = Name)
df2$Anc_kg <- (df2$Work_k*1000)/df2$Weight


if(!exists("databaze")) {
   databaze <- data.frame()
   databaze <- df2
} else {
  databaze <- rbind(databaze, df2)
}
}

if (exists("spatne.wingate")) {
if (!rapportools::is.empty(spatne.wingate)) {
print(paste("Nevyhodnocene soubory (chyba nazvu):", spatne.wingate, sep = " "))
}
}

writexl::write_xlsx(databaze, paste(wd, "/vysledky/vysledek", "_", team, "_", format(Sys.time(), "_%d-%m %H-%M"), ".xlsx", sep = ""), col_names = T, format_headers = T)


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

fs::file_move(path = paste(wd, "/", "out.rda", sep = ""), new_path = paste(wd, "/vymazat/", "out.rda", sep = ""))
fs::file_move(path = paste(wd, "/", "t1.png", sep = ""), new_path = paste(wd, "/vymazat/", "t1.png", sep = ""))
fs::file_move(path = paste(wd, "/", "t2.png", sep = ""), new_path = paste(wd, "/vymazat/", "t2.png", sep = ""))
fs::file_move(path = paste(wd, "/", "p1.png", sep = ""), new_path = paste(wd, "/vymazat/", "p1.png", sep = ""))