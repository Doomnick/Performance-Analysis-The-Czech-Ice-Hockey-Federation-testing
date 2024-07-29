wd <- "C:/Users/DKolinger/Documents/R/Wingate + Spiro"
setwd(wd)

library(tcltk2)
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




sports_data <- data.frame(
  Sport = c("hokej-dospělí","hokej-junioři", "hokej-dorost", "gymnastika"),
  FVC = c(4.7, 4.7, 4.7, 3.8),    # Example values for Forced Vital Capacity
  FEV1 = c(4.7, 4.7, 4.7, 3.1),   # Example values for Forced Expiratory Volume in 1 second
  VO2max = c(58, 58, 58, 48),   # Example maximal oxygen consumption values
  Power = c(4.5, 4.4, 4.4 ,3.3),
  Pmax = c(15.7, 15.3, 14.6, 13),
  SJ = c(45, 42, 42, 35),
  ANC = c(365, 330, 320 ,300)
)



file.path.an <- paste(wd, "/antropometrie", sep = "") 
spiro.path <- paste(wd, "/spiro", sep = "") 
report.path <- paste(wd, "/reporty", sep = "") #slozka pro reporty
database.path <- paste(wd, "/databaze", sep = "")

file.list.spiro <- list.files(path = spiro.path, pattern = c('*.xls'))
file.list.spiro.bez <- sort(tools::file_path_sans_ext(file.list.spiro))
file.list.an <- list.files(path = file.path.an, pattern = c('*.xls'))
file.list.dat <- list.files(path = database.path, pattern = c('*.xls'))

#### analyza antropy ####
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


kontrola.spiro <- sapply(list(file.list.spiro.bez), FUN = identical, file.list.an.bez)

duplikaty<- duplicated(file.list.an.bez)

if ("TRUE" %in% duplikaty) {
  print("Duplikaty v souboru s antropometrii:")
  print(file.list.an.bez[which(duplikaty=="TRUE")])
}

if  ("FALSE" %in% kontrola.spiro & !pracma::isempty(setdiff(file.list.spiro.bez, file.list.an.bez))) {
  cat('\n\n')
  print("Nazvy Wingate jsou rozdilne oproti Spiro - ZKONTROLUJ:")
  cat('\n')
  diff_str <- paste(setdiff(file.list.spiro.bez, file.list.an.bez), collapse = ", ")
  if (!pracma::isempty(setdiff(file.list.spiro.bez, file.list.an.bez))) {pokracovat <- winDialog("yesno", paste("Rozdíly v názvech Spiro/Antropo:", diff_str, "POKRAČOVAT?", sep= " "))}
  if (pokracovat == "NO") {
    stop("EXPORT PRERUSEN")
  } else if (pokracovat == "YES") {
    spatne.spiro <- setdiff(file.list.spiro.bez, file.list.an.bez)
    spatne.spiro <- paste(spatne.spiro, ".xlsx", sep = "")
    file.list.spiro <-  file.list.spiro[!file.list.spiro %in% spatne.spiro] 
  }
}


onSelect <- function() {
  selected_index <- tclvalue(tkcurselection(lst))
  if (as.numeric(selected_index) >= 0) {
    selected_sport <- sports_data$Sport[as.numeric(selected_index) + 1]
    
    # Store the selection in the global variable
    assign("sport", selected_sport, envir = .GlobalEnv)
    
    tkdestroy(top)
    createTeamNameWindow()  # Open team name entry window after sport selection
  }
}

createTeamNameWindow <- function() {
  win <- tktoplevel()
  tkwm.title(win, "Výběr teamu")
  
  # Create a label
  label <- tklabel(win, text = "Napište jméno týmu:")
  tkpack(label, side = "top", padx = 10, pady = 10)
  
  # Create a text entry widget
  textEntry <- tkentry(win, width = 50)
  tkpack(textEntry, side = "top", padx = 20, pady = 20)
  
  # Function to handle button click
  onSubmit <- function() {
    team <<- tclvalue(tkget(textEntry))  # Use '<<-' for global assignment
    tkdestroy(win)  # Optionally close the window after submitting
  }
  
  # Create a submit button
  submitButton <- tkbutton(win, text = "Potvrdit", command = onSubmit)
  tkpack(submitButton, side = "top", padx = 10, pady = 10)
  
  # Run the GUI
  tkfocus(win)
}

# Create main window
top <- tktoplevel()
tkwm.title(top, "Vyberte sport")

# Add listbox for sports
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

i<-1
for(i in 1:length(file.list.spiro)) {
  id <- tools::file_path_sans_ext(file.list.spiro.bez[i])
  if (an.input == "solo") {
    file.an <- file.list.an[which(file.list.an.bez == id)]
    antropo <- readxl::read_excel(paste(file.path.an, "/", file.an, sep=""), sheet = "Data_Sheet")
  }
  
  if (an.input == "batch") {
    antropo <- readxl::read_excel(paste(file.path.an, "/", file.list.an[1], sep=""), sheet = "Data_Sheet")   
  }
  if (an.input == "batch" & exists("duplikaty_check")) {
    antropo <- antropo %>% 
      group_by(ID) %>% 
      filter(Date_measurement==max(as.Date(Date_measurement)))
  }
  if (an.input == "batch") {
    datum_mer <- na.omit(antropo$Date_measurement[antropo$ID == id])
    vaha <- na.omit(antropo$Weight[antropo$ID == id])
    vaha <- ifelse(length(vaha) == 0, NA, vaha)
    vyska <- na.omit(antropo$Height[antropo$ID == id])
    vyska <- ifelse(length(vyska) == 0, NA, vyska)
    fat <- round(na.omit(antropo$Fat[antropo$ID == id]),1)
    fat <- ifelse(length(fat) == 0, NA, fat)
    ath <- round(na.omit(antropo$ATH[antropo$ID == id]),1)
    ath <- ifelse(length(ath) == 0, NA, ath)
    birth <- na.omit(antropo$Birth[antropo$ID == id])
    fullname <- paste(na.omit(antropo$Name[antropo$ID == id]), na.omit(antropo$Surname[antropo$ID == id]), sep = " ", collapse = NULL)
    fullname.rev <- paste(na.omit(antropo$Surname[antropo$ID == id]), na.omit(antropo$Name[antropo$ID == id]), sep = " ", collapse = NULL)
    datum_nar <- format(as.Date(na.omit(antropo$Birth[antropo$ID == id])),"%d/%m/%Y")
    datum_mer <- format(as.Date(na.omit(antropo$Date_measurement[antropo$ID == id])),"%d/%m/%Y")
    la <- antropo$LA[antropo$ID == id][1]
    age <- ifelse(!is.empty(antropo$Age[antropo$ID == id]), 
                  round(na.omit(antropo$Age[antropo$ID == id]),2), 
                  ifelse(!is.na(datum_nar) & datum_nar != "",
                         floor(as.integer(difftime(as.Date(datum_mer, format = "%d/%m/%Y"), as.Date(datum_nar, format = "%d/%m/%Y"), units = "days") / 365.25)),
                         NA))
  } else {
    datum_mer <- na.omit(tail(antropo$Date_measurement, n=1))
    vaha <- na.omit(tail(antropo$Weight, n = 1))
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
  if (length(file.list.spiro) != 0) {
    spiro <- readxl::read_excel(paste(spiro.path, "/", file.list.spiro[i], sep=""))
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
  FEV1 <- as.numeric(gsub("L", "", gsub(",", ".", spiro_info[[3]][which(spiro_info[[1]] == "FEV1")][1])))
  s.pomer <- round(FEV1/FVC*100,1)
  VO2max <- round(max(spiro$`V'O2`), 1)
  VO2_kg <- as.numeric(spiro_info[which(spiro_info[,1] == "V'O2/kg"), 12])
  vykon <- round(as.numeric(max(spiro$WR)),0)
  vykon_kg <- as.numeric(vykon/vaha)
  hrmax <- round(max(spiro$TF),0)
  tep.kyslik <- as.numeric(spiro_info[which(spiro_info[,1] == "V'O2/HR"), 12])
  anp <- as.numeric(ifelse(spiro_info[which(spiro_info[,1] == "TF"),9]!="-", as.numeric(spiro_info[which(spiro_info[,1] == "TF"),9]), as.numeric(spiro_info[which(spiro_info[,1] == "TF"),15])*0.85))
  min.ventilace <- as.numeric(gsub(",", ".", spiro_info[which(spiro_info[,1] == "V'E"), 12]))
  dech.frek <- as.numeric(spiro_info[which(spiro_info[,1] == "BF"), 12])
  dech.objem <-   dech.objem <- round(max(spiro$VT_5s[which.max(as.numeric(format(spiro$t, "%H%M%S")) > 100):nrow(spiro)]),1)
  dech.objem.per <- round((dech.objem/FVC)*100,0)
  rer <- round(max(spiro$RER),2)
  LaMax <- la
  
  
  #spiro tabulka
  columns_s <- c("Antropometrie", "Hodnota", "V3", "Spiroergometrie", "Absolutní", "Relativní (TH)", "% z nál. hodnoty" ,"Relativní (ATH)")
  df3 = data.frame(matrix(nrow = 14, ncol = length(columns_s))) 
  colnames(df3) <- columns_s
  df3$Antropometrie <-  `length<-`(c("Datum narození", "Věk", "Výška", "TH - Hmotnost", "Tuk", "ATH - Aktivní tělesná hmota", NA, "Klidová ventilace", "FVC", "FEV1", "Poměr FVC/FEV1"), nrow(df3))
  df3$Spiroergometrie <- `length<-`(c("VO2max", "Dosažený výkon", "HrMax", "Tepový kyslík", "ANP", NA, NA, "Ventilace", "Minutová ventilace (l/min)", "Dechová frekvence", "Dechový objem", "Dechový objem %", "RER", "LaMax"),nrow(df3))
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
  df3[9,3] <- round((FVC/sports_data$FVC[sports_data$Sport==sport])*100,1)
  df3[10,3] <- round((FEV1/sports_data$FEV1[sports_data$Sport==sport])*100,1)
  df3[1,7] <- round((VO2_kg/sports_data$VO2max[sports_data$Sport==sport])*100,1)
  df3[2,7] <- round(((vykon/vaha)/sports_data$Power[sports_data$Sport==sport])*100,1)
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
  
  
  #treninkove zony
  columns_tz <- c("Zóna", "od (BPM)", "do (BPM)")
  tz = data.frame(matrix(nrow = 3, ncol = length(columns_tz))) 
  colnames(tz) <- columns_tz
  tz$Zóna <- c("Aerobní", "Smíšená", "Anaerobní")
  tz$`od (BPM)`[2] <- round((anp / 0.85)*0.76,0)
  tz$`od (BPM)`[3] <- round(anp,0)
  tz$`do (BPM)`[1] <- round((anp / 0.85)*0.75,0)
  tz$`do (BPM)`[2] <- round((anp / 0.85)*0.84,0)
  
  tf_range <- min(spiro$`V'E`)
  ve_range <- max(spiro$TF)
  min(spiro$TF)
  
  #### spiro historie ####
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
  df4$FEV1_l <- append(na.omit(df4$FEV1), FEV1)
  df4$FVC_l <- append(na.omit(df4$FVC), FVC)
  df4$aerobni_Z_do <- append(na.omit(df4$aerobni_Z_do), round((anp / 0.85)*0.75,0))
  df4$smisena_Z_od <- append(na.omit(df4$smisena_Z_od), round((anp / 0.85)*0.76,0))
  df4$smisena_Z_do <- append(na.omit(df4$smisena_Z_do), round((anp / 0.85)*0.84,0))
  df4$anaerobni_Z_od <- append(na.omit(df4$anaerobni_Z_od), anp)
  df4 <- add_row(df4)
  
  coeff <- 50 
  ;
  hrmin <- min(spiro$TF)
  
  ylim.prim <- c(0, min.ventilace*1.1)# in this example, precipitation
  ylim.sec <- c(hrmin*0.9, hrmax*1.05)  
  
  b <- diff(ylim.prim)/diff(ylim.sec)
  a <- ylim.prim[1] - b*ylim.sec[1]
  
  # Plot with TF and VE
  plot1 <- ggplot(spiro, aes(x = t)) +
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
  
  plot2 <- ggplot(spiro, aes(x = t)) +
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
  
  library(patchwork)
  combined_plot <- plot1 + plot2 +
    plot_annotation(
      title = 'Spiroergometrie',
      caption = '',
      theme = theme(plot.title = element_text(size = 30, hjust = 0.5))
    )

  
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
  
  gtsave(
    table4,
    "t4.png",
    vwidth = 350,  
    vheight = 800  
  )
  
  table3 %>%
    gtsave("t3.png", vwidth = 1700, vheight = 1000) 


  
  png("p2.png", width = 1500, height = 600)
  plot(combined_plot)
  dev.off()

  rmarkdown::render("export_spiro_NEOTVIRAT.Rmd", output_file = paste(id, ".pdf", sep= ""), clean = T, quiet = F)
  
  fs::file_move(path = paste(wd, "/", id, ".pdf", sep = ""), new_path = paste(wd, "/reporty/", id, ".pdf", sep = ""))
  fs::file_move(path = paste(wd, "/", id, ".tex", sep = ""), new_path = paste(wd, "/vymazat/", id, ".tex", sep = ""))
  file.rename(from = paste(wd, "/", id, "_files", sep = ""), to = paste(wd, "/vymazat/", id, "_files", sep = ""))
}

if (!dir.exists(paste(wd, "/vysledky/spiro/", sep=""))) {
  dir.create(paste(wd, "/vysledky/spiro/", sep=""))
}


df4 <- df4[rev(order(df4$`VO2_kg/l`)), ]
writexl::write_xlsx(df4, paste(wd, "/vysledky/spiro/spiro_vysledek", "_", team, "_", format(Sys.time(), "_%d-%m %H-%M"), ".xlsx", sep = ""), col_names = T, format_headers = T)
