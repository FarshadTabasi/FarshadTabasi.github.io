library_list <- c("tidyverse", "readxl", "lubridate", "data.table", "rebus")

for (package in library_list) {
   if (!(package %in% installed.packages()[, "Package"])) {
      install.packages(package)
   }
}

for (package in library_list) {
   if (!(require(package,character.only = TRUE))) {
      require(package,character.only = TRUE)
   }
}


## OFFICE FOLDER
base_address <- "C:\\Users\\FARSHAD\\OneDrive\\DataBase\\Alberta\\Oilsands"
setwd(base_address)

## HOME FOLDER
# base_address <- "C:\\Users\\ftn60\\OneDrive\\DataBase\\Alberta\\Oilsands"
# setwd(base_address)



start_time <- Sys.time()

lst.file <- list.files(path = ".", recursive = TRUE, pattern = "\\.xlsx$", full.names = TRUE)
field_data <- list()
year_pattern <- one_or_more(DGT)
operator_pattern <- SPC %R% char_class("(") %R% "?" %R% char_class(")")

for (i in seq_along(lst.file)) {
   tab_names <- excel_sheets(lst.file[i])
   lt <- length(tab_names)
   annual_data <- list()
   year <- as.numeric(str_match(lst.file[i], pattern = year_pattern))[]
   for (j in (1:lt)) {
      annual_data[[tab_names[j]]] <- read_excel(lst.file[i], sheet = j, skip = 3, col_names = TRUE)
      annual_data[[tab_names[j]]]$Operator <- str_remove(annual_data[[tab_names[j]]]$Operator, pattern = operator_pattern)
      annual_data[[tab_names[j]]] <- annual_data[[tab_names[j]]] %>% mutate(Operator = ifelse(Operator == "Blackpearl Resources Inc.", "BlackPearl Resources Inc.",Operator))
      annual_data[[tab_names[j]]] <- annual_data[[tab_names[j]]] %>% mutate(Operator = ifelse(Operator == "Nexen Inc.", "CNOOC Petroleum North America ULC",Operator))
      annual_data[[tab_names[j]]] <- annual_data[[tab_names[j]]] %>% mutate(Operator = ifelse(Operator == "Nexen Energy ULC", "CNOOC Petroleum North America ULC",Operator))
      annual_data[[tab_names[j]]] <- annual_data[[tab_names[j]]] %>% mutate(Operator = ifelse(Operator == "Meg Energy Corp.", "MEG Energy Corp.",Operator))
      index <- which(is.na(annual_data[[tab_names[j]]][,1]))
      annual_data[[tab_names[j]]] <- annual_data[[tab_names[j]]][1:index[1] - 1,]
      annual_data[[tab_names[j]]]$`Recovery Method` <- as.factor(annual_data[[tab_names[j]]]$`Recovery Method`)
      factor_names <- unique(annual_data[[tab_names[j]]]$`Recovery Method`)
      levels(annual_data[[tab_names[j]]]$`Recovery Method`) <- factor_names
      y <- unlist(str_split(lst.file[i], pattern = "_"))[2]
      year <- as.numeric(str_extract(y, pattern = year_pattern))
      annual_data[[tab_names[j]]]$year <- year
      annual_data[[tab_names[j]]]$data <- tab_names[j]
      annual_data[[tab_names[j]]] <- annual_data[[tab_names[j]]] %>% dplyr::select(data,year,everything())
      if (j == 4) {
         well_month <- annual_data[[tab_names[j]]] %>% dplyr::select(c(Jan:Dec),`Monthly Average`)
         annual_data[[tab_names[j]]] <- annual_data[[tab_names[j]]] %>% dplyr::select(data:`Recovery Method`)
         col_names <- colnames(well_month)
         for (k in seq_along(col_names)) {
          df_temp <- well_month[,k] %>% separate(col_names[k], c(str_c(col_names[k],"_actual"), str_c(col_names[k],"_capable")), sep = "/", extra = 'drop')
          df_temp <- as.data.frame(apply(df_temp,2,as.numeric))
          annual_data[[tab_names[j]]] <- bind_cols(annual_data[[tab_names[j]]],df_temp)
         }
      }
   }
   # All data for each year
   field_data[[i]] <- annual_data
}

saveRDS(field_data,file = "oilsands_production_data.rds")
end_time <- Sys.time()
duration <- end_time - start_time

oil_prod_data <- field_data[[1]][[1]]   # first year(2010)
water_prod_data <- field_data[[1]][[2]]   # first year(2010)
steam_inj_data <- field_data[[1]][[3]]   # first year(2010)
well_data <- field_data[[1]][[4]]   # first year(2010)
SOR_data <- field_data[[1]][[5]]   # first year(2010)
WSR_data <- field_data[[1]][[6]]   # first year(2010)

for (i in (2:length(lst.file))) {
   oil_prod_data <- bind_rows(oil_prod_data,field_data[[i]][[1]])   # adding oil production data from other years
   water_prod_data <- bind_rows(water_prod_data,field_data[[i]][[2]])   # adding water production data from other years
   steam_inj_data <- bind_rows(steam_inj_data,field_data[[i]][[3]])   # adding steam injection datat from other years
   well_data <- bind_rows(well_data,field_data[[i]][[4]])   # adding well data from other years
   SOR_data <- bind_rows(SOR_data,field_data[[i]][[5]])   # adding SOR data from other years
   WSR_data <- bind_rows(WSR_data,field_data[[i]][[6]])   # adding WSR data from other years
}




# *************** CNRL's CSS Operation

oil_data_cnrl <- oil_prod_data %>% filter(Operator == "Canadian Natural Resources Limited", `Recovery Method` == "Commercial-CSS")
oil_data_cnrl_long <- oil_data_cnrl %>% gather(month, prod, Jan:Dec) %>% arrange(year) %>% mutate(date = format(ymd(paste0(year,month,15)), "%d-%m-%Y")) %>% mutate(date = dmy(date))
oil_data_cnrl_long %>%
   ggplot(aes(x = date, y = prod)) +
   geom_line()

SOR_data_cnrl <- SOR_data %>% filter(Operator == "Canadian Natural Resources Limited", `Recovery Method` == "Commercial-CSS")
SOR_data_cnrl_long <- SOR_data_cnrl %>% gather(month, sor, Jan:Dec) %>% arrange(year) %>% mutate(date = format(ymd(paste0(year,month,15)), "%d-%m-%Y")) %>% mutate(date = dmy(date))
SOR_data_cnrl_long %>%
   ggplot(aes(x = date, y = sor)) +
   geom_line() +
   geom_hline(yintercept = mean(SOR_data_cnrl_long$sor, na.rm = TRUE), col = "red")

wSR_data_cnrl <- WSR_data %>% filter(Operator == "Canadian Natural Resources Limited", `Recovery Method` == "Commercial-CSS")
wSR_data_cnrl_long <- wSR_data_cnrl %>% gather(month, wsr, Jan:Dec) %>% arrange(year) %>% mutate(date = format(ymd(paste0(year,month,15)), "%d-%m-%Y")) %>% mutate(date = dmy(date))
wSR_data_cnrl_long %>%
   ggplot(aes(x = date, y = wsr)) +
   geom_line() +
   geom_hline(yintercept = mean(wSR_data_cnrl_long$wsr, na.rm = TRUE), col = "red")



# *************** IMPERIAL's CSS Operation

oil_data_imperial <- oil_prod_data %>% filter(Operator == "Imperial Oil Resources", `Recovery Method` == "Commercial-CSS")
oil_data_imperial_long <- oil_data_imperial %>% gather(month, prod, Jan:Dec) %>% arrange(year) %>% mutate(date = format(ymd(paste0(year,month,15)), "%d-%m-%Y")) %>% mutate(date = dmy(date))
oil_data_imperial_long %>%
   ggplot(aes(x = date, y = prod)) +
   geom_line()

SOR_data_imperial <- SOR_data %>% filter(Operator == "Imperial Oil Resources", `Recovery Method` == "Commercial-CSS")
SOR_data_imperial_long <- SOR_data_imperial %>% gather(month, sor, Jan:Dec) %>% arrange(year) %>% mutate(date = format(ymd(paste0(year,month,15)), "%d-%m-%Y")) %>% mutate(date = dmy(date))
SOR_data_imperial_long %>%
   ggplot(aes(x = date, y = sor)) +
   geom_line() +
   geom_hline(yintercept = mean(SOR_data_imperial_long$sor, na.rm = TRUE), col = "red")

wSR_data_imperial <- WSR_data %>% filter(Operator == "Imperial Oil Resources", `Recovery Method` == "Commercial-CSS")
wSR_data_imperial_long <- wSR_data_imperial %>% gather(month, wsr, Jan:Dec) %>% arrange(year) %>% mutate(date = format(ymd(paste0(year,month,15)), "%d-%m-%Y")) %>% mutate(date = dmy(date))
wSR_data_imperial_long %>%
   ggplot(aes(x = date, y = wsr)) +
   geom_line() +
   geom_hline(yintercept = mean(wSR_data_imperial_long$wsr, na.rm = TRUE), col = "red")



# *************** All Operators

alberta_oil_prod_plot  <- oil_prod_data %>% mutate(year_f = fct_rev(as.factor(year))) %>% filter(`Recovery Method` == "Commercial-SAGD") %>%
   ggplot(aes(x = Operator, y = `Monthly Average`/1000, fill = year_f)) +
   geom_bar(stat = "identity") +
   xlab("Operator") +
   ylab("Cumulative Monthly Oil Production, Mm3/day") +
   labs(title = "Commercial SAGD Operation in Alberta", caption = "source: https://aer.ca") +
   theme_bw() +
   theme(
      plot.title = element_text(color = "black", size = 20, face = "bold"),
      plot.caption = element_text(color = "black", size = 18, face = "italic"),
      axis.text.x = element_text(angle = 90, hjust = 1, face = "bold", size = 16),
      legend.text = element_text(size = 16),
      legend.title = element_text(size = 16, face = "bold"),
      axis.text.y = element_text(face = "bold", size = 16),
      axis.title = element_text(size = 16, face = "bold")
   )
ggsave(filename = str_c(base_address, "\\alberta_oil_prod_plot.png", sep = ""), plot = alberta_oil_prod_plot, scale = 3, dpi = 300)
