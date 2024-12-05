# Copyright 2024 Province of British Columbia
#
# Licensed under the Apache License, Version 2.0 (the "License");
# you may not use this file except in compliance with the License.
# You may obtain a copy of the License at
#
# http://www.apache.org/licenses/LICENSE-2.0
#
# Unless required by applicable law or agreed to in writing, software distributed under the License is distributed on an "AS IS" BASIS,
# WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
# See the License for the specific language governing permissions and limitations under the License.

#### Country Trade Profiles App ----
## https://bcstats.shinyapps.io/CountryTradeApp/
## https://www2.gov.bc.ca/gov/content/data/statistics/business-industry-trade/trade/trade-data/country-trade-profile
## Version: V2 (added countries and regions; added/changed some sources & notes)
## Updated: Jan 2021, May 2024
## Coder: Julie Hawkins (and Stephanie Yurchak for app)

## ** This file creates the country and state nested datasets. **


## load packages ----
library(tidyverse)
library(openxlsx)
library(janitor)

options(scipen=999)


## functions ----

### * get.last.year() function pulls the specified row's last column to get the last year in the data
get.last.year <- function(data, colVar, rowVar) {
  data %>%
    rename(colVar = {{colVar}}) %>%
    filter(colVar == rowVar) %>%
    select(ncol(data)) %>%
    pull() %>%
    as.character()
}
# yearHT <- get.last.year(data = dataHighTech, colVar = `BC.Trade.in.High.Technology.Goods`, rowVar = "Domestic Exports")


### * make.numeric() function drops unnecessary text rows and ensures remainder are numeric
make.numeric <- function(data, colVar) {
  ## find total row in colVar column
  endRow <- which(tolower(data[[{{colVar}}]]) == "total")
  ## drop all rows past total row
  data <- data[1:endRow, ]
  ## convert columns to numeric
  data <- data %>%
    mutate_at(vars(c(names(.)[2]:(which(names(.) == "Label")-1))), as.numeric)
  return(data)
}
# dataProv <- make.numeric(data = dataProv, colVar = "Province.of.Origin")


### * rows.to.colnames() function pulls rows specified by text match in specified searchCol, and
##   elevates first row to column names if rowAsNames = TRUE
rows.to.colnames <- function(data, rowAsNames = TRUE, searchCol, startRow, endRow) {

  ## temporarily rename searchCol to use in `slice()`
  data <- data %>% rename(Col = {{searchCol}})

  ## get index of specified rows: from startRow to row BEFORE endRow
  startRowa <- (which(data$Col == startRow))[1]  ## get first instance, in case multiple instances
  endRowa <- (which(data$Col == endRow)-1)[1]    ## get first instance, in case multiple instances

  ## subset data for only specified rows, set searchCol name back to original
  data <- data %>%
    slice(startRowa:endRowa) %>%
    rename({{searchCol}} := Col)

  ## unless FALSE, elevate first row to be the column names
  if(rowAsNames){
    data <- data %>% janitor::row_to_names(row_number = 1)
  }
  data
}


### * make.row.in.Services() function makes a combined countries row in dataServices
make.row.in.Services <- function(data, countryVar, grpdCountry, individualCountries) {
  temp <- data %>%
    rename(tempVar = {{countryVar}}) %>%
    ## pull just individualCountries rows
    filter(tempVar %in% individualCountries) %>%
    ## pull just relevant transaction rows: "Total services, receipts" & "Total services, payments"
    filter(X2 == "Total services, receipts" | X2 == "Total services, payments") %>%
    ## gather year columns as temp
    pivot_longer(-c(names(.)[1], "X2", "tempVar"), names_to = "temp", values_to = "value") %>%
    ## spread by country
    pivot_wider(names_from = "tempVar", values_from = "value") %>%
    ## convert country cols to numeric for rowSums() and create new grpdCountry (as.character)
    mutate(across(4:(3+length(individualCountries)),
                  ~case_when(.x == ".." ~ NA_real_, TRUE ~ as.numeric(.x))),
           Grp = as.character(rowSums(across(4:(3+length(individualCountries))), na.rm = TRUE))) %>%
    rename({{grpdCountry}} := Grp) %>%
    ## drop original individual countries
    select(-all_of(individualCountries)) %>%
    ## gather country as countryVar column
    pivot_longer(-c(names(.)[1], "X2", "temp"), names_to = "cV", values_to = "value") %>%
    rename({{countryVar}} := cV) %>%
    ## spread by year (aka, temp)
    pivot_wider(names_from = "temp", values_from = "value")
  data <- data %>% bind_rows(temp)
}
# test <- make.row.in.Services(data = dataServices, countryVar = "X3", grpdCountry = "China & Hong Kong",
#                              individualCountries = c("Hong Kong", "Mainland China"))


### * make.row.in.Travel() function makes a combined countries row in dataTravel
make.row.in.Travel <- function(data, countryVar, grpdCountry, individualCountries) {
  temp <- data %>%
    rename(tempVar = {{countryVar}}) %>%
    ## pull just individualCountries rows
    filter(tempVar %in% individualCountries) %>%
    ## convert country cols to numeric
    mutate(across(-c(tempVar), as.numeric)) %>%
    ## gather year columns as temp
    pivot_longer(-tempVar, names_to = "temp", values_to = "value") %>%
    ## spread by country
    pivot_wider(names_from = "tempVar", values_from = "value") %>%
    ## create new grpdCountry (as.character)
    mutate(Grp = rowSums(across(2:(1+length(individualCountries))), na.rm = TRUE),
           Country = grpdCountry) %>%
    rename({{grpdCountry}} := Grp, {{countryVar}} := Country) %>%
    ## drop original individual countries
    select(-all_of(individualCountries)) %>%
    ## spread by year (aka, temp)
    pivot_wider(names_from = "temp", values_from = grpdCountry)
  if(class(data[,2]) == "character") {
    temp[,2] <- as.character(temp[,2])
  }
  data <- data %>%
    mutate(across(-c(all_of(countryVar), names(.)[2]), as.numeric)) %>%
    bind_rows(temp)
}
# test <- make.row.in.Travel(data = dataTravel, countryVar = "Country of residence",
#                            grpdCountry = "China & Hong Kong",
#                            individualCountries = c("Hong Kong", "Mainland China"))


### * make.row.in.Investment() function makes a combined countries row in dataInvestment
make.row.in.Investment <- function(data, countryVar, grpdCountry, individualCountries) {
  temp <- data %>%
    rename(tempVar = {{countryVar}}) %>%
    ## pull just individualCountries rows
    filter(tempVar %in% individualCountries) %>%
    ## convert cols to character
    mutate(across(where(is.numeric), as.character)) %>%
    ## gather year columns as temp
    pivot_longer(-c(names(.)[1], "X2", tempVar), names_to = "temp", values_to = "value") %>%
    ## re-word X2 column so that individual country rows have matching info
    mutate(X2 = case_when(str_detect(X2, pattern = "abroad") ~ "Canadian direct investment abroad",
                          str_detect(X2, pattern = "Foreign") ~ "Foreign direct investment in Canada")) %>%
    ## spread by country
    pivot_wider(names_from = "tempVar", values_from = "value") %>%
    ## convert country cols to numeric temporarily for rowSums(), then back to character
    mutate(across(4:(3+length(individualCountries)),
                  ~case_when(.x == "x" ~ NA_real_, TRUE ~ as.numeric(gsub("x", "", .x)))),  ## gsub() suppresses NA coercion warning: https://statisticsglobe.com/warning-message-nas-introduced-by-coercion-in-r#:~:text=As%20you%20can%20see%2C%20the,converted%20to%20the%20numeric%20class
           Grp = rowSums(across(4:(3+length(individualCountries))), na.rm = FALSE),
           Grp = case_when(is.na(Grp) ~ "x", TRUE ~ as.character(Grp))) %>%
    rename({{grpdCountry}} := Grp) %>%
    ## drop original individual countries
    select(-all_of(individualCountries)) %>%
    ## gather country as X3 column
    pivot_longer(-c(names(.)[1], "X2", "temp"), names_to = countryVar, values_to = "value") %>%
    ## spread by year (aka, temp)
    pivot_wider(names_from = "temp", values_from = "value")
  data <- data %>%
    mutate(across(where(is.numeric), as.character)) %>%
    bind_rows(temp)
}
# test <- make.row.in.Investment(data = dataInvestment, countryVar = "X3",
#                                grpdCountry = "China & Hong Kong",
#                                individualCountries = c("Hong Kong", "Mainland China"))


### * top5.Cda.table() function pulls top 5 commodities in CDN $Millions and Share of Total %, for specified country
top5.Cda.table <- function(data, country) {

  ## get TOTAL row temporarily
  temp <- data %>%
    rename(country = {{country}}) %>%
    select(Commodity, country) %>%
    filter(tolower(Commodity) == "total") %>%    # filter(Commodity == "TOTAL") %>%
    mutate(Commodity = "Total")

  data %>%
    rename(country = {{country}}) %>%
    ## get just Commodity and country columns
    select(Commodity, country) %>%
    ## temporarily drop TOTAL row
    filter(tolower(Commodity) != "total") %>%    # filter(Commodity != "TOTAL") %>%
    ## sort descending
    arrange(desc(country)) %>%
    ## get just top 5 commodities
    slice(1:5) %>%
    ## add number (`row.names(.)`) and remove 4 digits and dash at beginning of Commodity: e.g., 1. Crude petroleum...
    mutate(Commodity = paste0(row.names(.), ". ", str_sub(Commodity, start = 6))) %>%
    ## bring back TOTAL row
    bind_rows(temp) %>%
    ## divide by TOTAL, multiply by 100, round to whole number and add percent sign
    mutate(Share.of.Total = paste0(janitor::round_half_up(100*(country / temp$country), digits = 0), "%")) %>%
    ## convert to millions, one decimal place, add comma separator and dollar sign
    mutate(country = paste0("$", prettyNum(round_half_up(country/1000000, digits = 1), big.mark = ","))) %>%
    rename(`Cdn $Millions` = country, `Share of Total` = Share.of.Total) %>%
    ## add in country name as first column
    mutate(Country = rlang::quo_text(enquo(country))) %>%
    select(Country, everything())

}


### * top5.US.table() function pulls top 5 "var"s in US $Millions and Share of Total %, for specified country
top5.US.table <- function(data, var, country) {

  ## get TOTAL row temporarily
  temp <- data %>%
    rename(var = {{var}}, country = {{country}}) %>%
    select(var, country) %>%
    filter(var == "TOTAL") %>%
    mutate(var = "Total")

  data %>%
    rename(var = {{var}}, country = {{country}}) %>%
    ## get just var and country columns
    select(var, country) %>%
    ## temporarily drop TOTAL row
    filter(var != "TOTAL") %>%
    ## sort descending
    arrange(desc(country)) %>%
    ## get just top 5 commodities
    slice(1:5) %>%
    ## add number (`row.names(.)`) and remove 2 digits, colon and space at beginning of var
    mutate(var = paste0(row.names(.), ". ", str_sub(var, start = 5))) %>%
    ## bring back TOTAL row
    bind_rows(temp) %>%
    ## divide by TOTAL, multiply by 100, round to whole number and add percent sign
    mutate(Share.of.Total = paste0(janitor::round_half_up(100*(country / temp$country), digits = 0), "%")) %>%
    ## convert to millions, one decimal place, add comma separator and dollar sign
    mutate(country = paste0("$", prettyNum(round_half_up(country/1000, digits = 1), big.mark = ","))) %>%
    # mutate(country = paste0("$", prettyNum(round_half_up(country/1000000, digits = 1), big.mark = ","))) %>%
    rename(`US $Millions` = country, `Share of Total` = Share.of.Total) %>%
    ## add in country name as first column
    mutate(Country = rlang::quo_text(enquo(country))) %>%
    select(Country, everything())

}


### * prov.table() function gets provincial distribution of exports to specified country
prov.table <- function(data, country) {
  countryTotal <- data %>% select({{country}}) %>% sum(na.rm = TRUE)
  data %>%
    ## rename country column
    rename(country = {{country}}) %>%
    ## get just Province.of.Origin and country columns
    select(Province.of.Origin, country) %>%
    ## drop any NA rows
    filter(!is.na(Province.of.Origin)) %>%
    ## fix some provinces names (abbreviated)
    mutate(Province.of.Origin = case_when(str_detect(Province.of.Origin, pattern = "Lab") ~ "Newfoundland & Labrador",
                                          str_detect(Province.of.Origin, pattern = "Terr") ~ "Northwest Territories",
                                          str_detect(Province.of.Origin, pattern = "Is") ~ "Prince Edward Island",
                                          TRUE ~ as.character(Province.of.Origin))) %>%
    arrange(desc(country)) %>%
    ## add number (`row.names(.)`) before Province name
    mutate(Province.of.Origin = paste0(row.names(.), ". ", Province.of.Origin)) %>%
    ## add in total row
    add_row(Province.of.Origin = "Canada Total", country = countryTotal) %>%
    ## divide by TOTAL, multiply by 100, round to whole number and add percent sign
    mutate(Share.of.Total = paste0(janitor::round_half_up(100*(country / countryTotal), digits = 1), "%")) %>%
    ## convert to millions, one decimal place, add comma separator and dollar sign
    mutate(country = paste0("$", prettyNum(round_half_up(country/1000000, digits = 1), big.mark = ","))) %>%
    rename(`Province of Origin` = Province.of.Origin, `Cdn $Millions` = country, `Share of Total` = Share.of.Total) %>%
    ## add in country name as first column
    mutate(Country = rlang::quo_text(enquo(country))) %>%
    select(Country, everything())
}


### * gen.table() function gets general information, for specified country
gen.table <- function(data, country) {
  temp <- data %>% filter(X2 == {{country}})

  ## pull stats
  pop <- temp %>% select(Population) %>% pull() %>% prettyNum(big.mark = ",")
  apgr <- temp %>% select(`Pop.Growth.%`) %>% pull() %>% paste0(., "%")
  gdp <- temp %>% select(`GDP.($US.billions).(Purchasing.Power.Parity)`) %>% pull() %>% prettyNum(big.mark = ",") %>% paste0("$", .)
  gdp2 <- temp %>% select(`Per.Capita.GDP.($US.2017)`) %>% pull() %>% prettyNum(big.mark = ",") %>% paste0("$", .)
  gdp3 <- temp %>% select(`Real.GDP.Growth.%`) %>% pull() %>% paste0(., "%")

  ## pull last year data was available
  yrs <- data %>%
    slice(1) %>%
    select(-contains("Source"), -X2) %>%
    mutate_if(is.numeric, as.character) %>%
    pivot_longer(everything(), names_to = "Var", values_to = "Year", values_drop_na = TRUE) %>%
    mutate(year = paste0("(", Year, ")"))

  ## put all into table
  tibble::tibble(value = c(pop, apgr, gdp, gdp2, gdp3), yrs) %>%
    select(Var, value, Year) %>%
    mutate(Var = case_when(Var == "Pop.Growth.%" ~ "Annual Population Growth Rate",
                           Var == "GDP.($US.billions).(Purchasing.Power.Parity)" ~ "GDP (US$ billions)*",
                           Var == "Per.Capita.GDP.($US.2017)" ~ "Per Capita GDP (US$ 2017)*",
                           Var == "Real.GDP.Growth.%" ~ "GDP Real Growth",
                           TRUE ~ as.character(Var))) %>%
    ## add in country name as first column
    mutate(Country = {{country}}) %>%
    select(Country, everything())

}


### * high.tech.table() function gets High Technology Trade data, for specified country and years
high.tech.table <- function(data, country, year) {

  yearStart <- as.character(as.numeric(year) - 9)

  ## get Re-Exports info
  RX <- rows.to.colnames(data = data, searchCol = BC.Trade.in.High.Technology.Goods,
                         startRow = "Re-Exports", endRow = "Imports") %>%
    filter(`Re-Exports` == {{country}}) %>%
    select(-`Re-Exports`) %>%
    pivot_longer(everything(), names_to = "Year", values_to = "Re-Exports")

  ## get Imports info
  Im <- data %>%
    add_row(BC.Trade.in.High.Technology.Goods = "zzz") %>%
    rows.to.colnames(searchCol = BC.Trade.in.High.Technology.Goods,
                     startRow = "Imports", endRow = "zzz") %>%
    filter(`Imports` == {{country}}) %>%
    select(-`Imports`) %>%
    pivot_longer(everything(), names_to = "Year", values_to = "Imports")

  ## get Domestic Exports info
  rows.to.colnames(data = data, searchCol = BC.Trade.in.High.Technology.Goods,
                   startRow = "Domestic Exports", endRow = "Re-Exports") %>%
    filter(`Domestic Exports` == {{country}}) %>%
    select(-`Domestic Exports`) %>%
    pivot_longer(everything(), names_to = "Year", values_to = "Domestic Exports") %>%
    ## add in Re-Exports and Imports info, calculate Trade Balance
    left_join(RX, by = "Year") %>%
    left_join(Im, by = "Year") %>%
    mutate(`Trade Balance` = `Domestic Exports` + `Re-Exports` - `Imports`) %>%
    ## get specified years only
    # filter(Year %in% yearStart:year) %>%
    filter(Year >= yearStart) %>%
    ## add in country name as first column
    mutate(Country = {{country}}) %>%
    select(Country, everything())

}
## ignore warning message: "Column `Year` has different attributes on LHS and RHS of join"


### * invest.table() function gets investment position data, for specified country
invest.table <- function(data, country) {
 data %>%
    ## get country and year rows
    filter(X3 == {{country}} | X3 == "Countries or regions") %>%
    select(-X3) %>%
    ## elevate row 1 as column names
    janitor::row_to_names(row_number = 1) %>%
    select(-Geography) %>%
    ## re-word what will become data column names
    rename(X2 = `Canadian and foreign direct investment1`) %>%
    mutate(X2 = case_when(str_detect(X2, "abroad") ~ "Canadian Direct Investment",   ## paste0("Canadian Direct Investment in ", country)
                          str_detect(X2, "Foreign") ~ "Direct Investment in Canada")) %>%
    ## gather year data
    pivot_longer(-X2, names_to = "Year", values_to = "stat") %>%
    ## spread by X2
    pivot_wider(names_from = X2, values_from = stat) %>%
    # ## get specified years only
    # filter(Year %in% yearStart:year) %>%
    ## add in country name as first column
    mutate(Country = {{country}}) %>%
    select(Country, everything())
}


### * travel.table() function gets travel and immigration data, for specified country
travel.table <- function(dataTravel, dataIm, country) {
  ## get Immigration info
  Im <- dataIm %>%
    rename(Country = `Country of Citizenship`) %>%
    mutate(Country = case_when(Country == "United States of America" ~ "United States",
                               TRUE ~ as.character(Country))) %>%
    filter(Country == {{country}}) %>%
    mutate_at(vars(-Country), as.numeric) %>%
    pivot_longer(-Country, names_to = "Year", values_to = "Immigrants (Persons)")

  ## get Travel info
  dataTravel %>%
    ## get column names from row 3
    #janitor::row_to_names(row_number = 3) %>%
    ## get country row
    filter(`Country of residence` == {{country}}) %>%
    select(-`Country of residence`) %>%
    mutate_if(is.character, as.numeric) %>%
    ## gather data by Year
    pivot_longer(everything(), names_to = "Year", values_to = "Travellers (Persons)") %>%
    ## add in Immigrants info
    left_join(Im, by = "Year") %>%
    ## add in country name as first column
    mutate(Country = {{country}}) %>%
    select(Country, everything())

}
## ignore warning message: "Column `Year` has different attributes on LHS and RHS of join"


### * time.trade() function gets trade data for specified country
time.trade <- function(BCX, CX, CTX, CM, country) {
  ## get exports and imports data for specified country
  BCX <- BCX  %>% mutate(Var = "BC Exports") %>%
    rename(Country = `BC Origin Exports`) %>% filter(Country == {{country}})
  CX <- CX %>% mutate(Var = "Canada Exports") %>%
    rename(Country = `Canada Origin Exports`) %>% filter(Country == {{country}})
  CTX <- CTX %>% mutate(Var = names(CTX)[1]) %>%
    rename(Country = `Canada Total Exports`) %>% filter(Country == {{country}})
  CM <- CM %>% mutate(Var = names(CM)[1]) %>%
    rename(Country = `Canada Imports`) %>% filter(Country == {{country}})

  ## bind above data together
  BCX %>%
    bind_rows(CX) %>%
    bind_rows(CTX) %>%
    bind_rows(CM) %>%
    select(Var, everything()) %>%
    mutate_at(vars(-"Var", -"Country"), as.numeric) %>%
    ## flip data so Year is a column and Var becomes multiple columns
    pivot_longer(-c("Var", "Country"), names_to = "Year", values_to = "value") %>%
    pivot_wider(names_from = "Var", values_from = "value") %>%
    ## calculate `Trade Balance` and `BC %Canada` columns
    mutate(`Trade Balance` = `Canada Total Exports` - `Canada Imports`,
           `BC %Canada` = 100 * `BC Exports` / `Canada Exports`) %>%
    ## add in country name as first column
    mutate(Country = {{country}}) %>%
    select(Country, everything())

}


### * time.service() function gets service trade data for specified country
time.service <- function(data, country) {

  ## get service exports and imports data for specified country
  data %>%
    janitor::row_to_names(row_number = 3) %>%
    rename(Country = `Countries or regions`, Var = `Type of transaction`) %>%
    filter(Country == {{country}}) %>%
    select(-Geography) %>%
    ## gather data so Year is a column
    pivot_longer(-c("Var", "Country"), names_to = "Year", values_to = "value") %>%
    ## change Var names, and drop unneeded rows
    mutate(Var = case_when(Var == "Total services, receipts" ~ "Service Exports",
                           Var == "Total services, payments" ~ "Service Imports",
                           TRUE ~ NA_character_),
           value = as.numeric(value)) %>%
    filter(!is.na(Var)) %>%
    ## spread data so Var becomes multiple columns
    pivot_wider(names_from = "Var", values_from = "value") %>%
    ## calculate `Service Trade Balance`
    mutate(`Service Trade Balance` = `Service Exports` - `Service Imports`) %>%
    ## add in country name as first column
    mutate(Country = {{country}}) %>%
    select(Country, everything())

}


### * time.gdp() function gets GDP % growth for specified country
time.gdp <- function(data, country) {

  data %>%
    ## get specified country row
    filter(Country == {{country}}, Subject.Descriptor == "Gross domestic product, constant prices") %>%
    select(-Subject.Descriptor, -Units) %>%
    ## ensure all Year columns are numeric
    mutate_at(vars(-c("Country")), as.numeric) %>%
    mutate(Var = "GDP % Growth") %>%
    select(Country, Var, everything()) %>%
    ## gather data so Year is a column
    pivot_longer(-c("Country", "Var"), names_to = "Year", values_to = "value") %>%
    ## add in country name as first column
    mutate(Country = {{country}}) %>%
    select(Country, everything())

}


### * USxchg.table() function gets Canada Exports and Canada Imports in US$ for specified country
USxchg.table <- function(time_trade, TimeXchg, country) {

  ## get {{country}} info of `Canada Exports` and `Canada Imports` by running `time.trade()`
  time_trade <- time_trade %>% select(Country, Year, `Canada Exports`, `Canada Imports`)

  TimeXchg %>%
    mutate_at(vars(-Year), as.numeric) %>%
    mutate(Country = {{country}}) %>%
    select(Country, everything(), -Year) %>%
    ## gather by Year
    pivot_longer(-c("Country"), names_to = "Year", values_to = "Canada/US Exchange") %>%
    ## add in time.trade info of that country of Canada's Exports and Imports (in Cdn$)
    left_join(time_trade, by = c("Country", "Year")) %>%
    ## convert Cdn$ into US$
    mutate(`Canada Exports $US` = `Canada Exports` / `Canada/US Exchange`,
           `Canada Imports $US` = `Canada Imports` / `Canada/US Exchange`) %>%
    select(-`Canada Exports`, -`Canada Imports`) %>%
    ## spread by Year
    pivot_longer(-c("Country", "Year"), names_to = "Var", values_to = "value") %>%
    pivot_wider(names_from = "Year", values_from = "value")

}


### * rank.data() function gets importance of specified country relative to BC and Canada
rank.data <- function(data, cols, var, country, grpdCountries = "", grpdCountry = FALSE){

  ## get vector of grouped countries
  # grpdCountries <- data %>% select("X22") %>%
  #   filter(!is.na(X22) & str_detect(X22, pattern = "Country", negate = TRUE)) %>%
  #   pull() %>% unique()

  ## prep data
  temp <- data[, {{cols}}] %>%
    janitor::row_to_names(row_number = 1) %>%
    rename(Country = names(.)[1], `Val+` = names(.)[3]) %>%
    filter(!(Country %in% grpdCountries))

  valSum <- sum(as.numeric(temp$`Val+`), na.rm = T)

  ## if a grpdCountry, add data as a row
  if(grpdCountry == TRUE) {
    ## get just percentage if grpdCountry (i.e., don't include grpdCountries in ranking)
    out <- data[, {{cols}}] %>%
      janitor::row_to_names(row_number = 1) %>%
      rename(Country = names(.)[1], `Val+` = names(.)[3]) %>%
      filter(Country %in% grpdCountries) %>%
      select(Country, Value = {{var}}, `Val+`) %>%
      mutate(Var = rlang::quo_text(enquo(var)),
             `Val+` = as.numeric(`Val+`),
             Rank = NA_integer_,
             Perc = 100 * `Val+` / valSum) %>%
      filter(Country == {{country}}) %>%
      select(Country, Var, Rank, Perc)
    ## old way, where grouped countries data was in a column (X22) at far right of data
    # aa <- data %>% select("X22", "X23") %>% filter(!is.na(X22))
    # getRow <- which(aa$X23 == rlang::as_name(enquo(var)))
    # aa <- aa %>% slice(getRow:(getRow + length(grpdCountries))) %>% filter(X22 == {{country}}) %>%
    #   rename(Country = X22, `Val+` = X23)
    # temp <- bind_rows(temp, aa)
  }

  if(grpdCountry == FALSE) {
    ## re-calc rank (may change depending on which index Excel book was on when data read in)
    out <- temp %>%
      select(Country, Value = {{var}}, `Val+`) %>%
      mutate(Var = rlang::quo_text(enquo(var)),
             `Val+` = as.numeric(`Val+`),
             Rank = rank(desc(as.numeric(`Val+`)), ties.method = "average"),
             # Perc = 100 * `Val+` / sum(`Val+`, na.rm = TRUE)) %>%
             Perc = 100 * `Val+` / valSum) %>%
      filter(Country == {{country}}) %>%
      select(Country, Var, Rank, Perc)
  }

  return(out)

}
# rank.data(dataRank, cols = 1:4, var = `BC Exports`, country = "ASEAN", grpdCountries = grpdCountries, grpdCountry = TRUE)
# rank.data(dataRank, cols = 1:4, var = `BC Exports`, country = "Australia")
# rank.data(dataRank, cols = 8:11, var = `Canada Exports`, country = "Australia")
# rank.data(dataRank, cols = 15:18, var = `Canada Imports`, country = "Australia")
## old version
# rank.data <- function(data, cols, var, country, grpdCountry = FALSE){
#
#   ## get vector of grouped countries
#   grpdCountries <- data %>% select("X22") %>%
#     filter(!is.na(X22) & str_detect(X22, pattern = "Country", negate = TRUE)) %>%
#     pull() %>% unique()
#
#   ## prep data
#   temp <- data[, {{cols}}] %>% janitor::row_to_names(row_number = 1) %>%
#     rename(Country = names(.)[1], `Val+` = names(.)[3]) %>%
#     filter(!(Country %in% grpdCountries))
#
#   valSum <- sum(as.numeric(temp$`Val+`), na.rm = T)
#
#   ## if a grpdCountry, add data as a row
#   if(grpdCountry == TRUE) {
#     aa <- data %>% select("X22", "X23") %>% filter(!is.na(X22))
#     getRow <- which(aa$X23 == rlang::as_name(enquo(var)))
#     aa <- aa %>% slice(getRow:(getRow + length(grpdCountries))) %>% filter(X22 == {{country}}) %>%
#       rename(Country = X22, `Val+` = X23)
#     temp <- bind_rows(temp, aa)
#   }
#
#   ## re-calc rank (may change depending on which index Excel book was on when data read in)
#   temp %>%
#     select(Country, Value = {{var}}, `Val+`) %>%
#     mutate(Var = rlang::quo_text(enquo(var)),
#            `Val+` = as.numeric(`Val+`),
#            Rank = rank(desc(as.numeric(`Val+`)), ties.method = "average"),
#            # Perc = 100 * `Val+` / sum(`Val+`, na.rm = TRUE)) %>%
#            Perc = 100 * `Val+` / valSum) %>%
#     filter(Country == {{country}}) %>%
#     select(Country, Var, Rank, Perc)
#
# }


### * rank.bind() function runs rank.data() and binds info to gets importance of specified country relative to BC and Canada
rank.bind <- function(data, country, grpdCountries) {
  ## run rank info for BC Exports
  rank.data(data, cols = 1:4, var = `BC Exports`, country, grpdCountries = grpdCountries) %>%
    ## rank info for Canada Exports
    bind_rows(rank.data(data, cols = 8:11, var = `Canada Exports`, country, grpdCountries = grpdCountries)) %>%
    ## rank info for Canada Imports
    bind_rows(rank.data(data, cols = 15:18, var = `Canada Imports`, country, grpdCountries = grpdCountries)) %>%
    mutate(Perc = paste0(janitor::round_half_up(Perc, digits = 1), "%"),
           Var = case_when(Var == "`BC Exports`" ~ "BC Exports",
                           Var == "`Canada Exports`" ~ "Canada Exports",
                           Var == "`Canada Imports`" ~ "Canada Imports"))

}
# rank.bind(data = dataRank, country = "Australia", grpdCountries = grpdCountries)


### * imp.table() function gets importance of Canada to specified country for specified year
imp.table <- function(World, t14, nested = TRUE, year, country) {

  ## get TOTAL exports and imports for {{country}}
  temp <- World %>% filter(Country == {{country}})

  ## get Canadian imports and exports from/to {{country}}
  if(nested == TRUE) {
    tab <- t14 %>% unnest(cols = c(t14))
  } else {
    tab <- t14
  }

  tab %>%
    filter(Country == {{country}}) %>%
    select(Var, all_of({{year}})) %>%
    pivot_wider(names_from = "Var", values_from = {{year}}) %>%
    ## add in {{country}} total exports and imports
    mutate(CtyX = temp$WorldX * 100000,  ## convert from Millions
           CtyM = temp$WorldM * 100000) %>%
    ## divide {{country}} exports with Canadian imports, and {{country}} imports with Canadian exports to get share
    mutate(`Imports into Canada` = 100 * `Canada Exports $US` / CtyM,
           `Exports from Canada` = 100 * `Canada Imports $US` / CtyX) %>%
    select(`Imports into Canada`, `Exports from Canada`) %>%
    pivot_longer(everything(), names_to = "Var", values_to = "Share") %>%
    mutate(Country = {{country}},
           Share = janitor::round_half_up(100*Share, digits = 1)) %>%
    select(Country, Var, Share)

}


## load data ----

inputs <- openxlsx::loadWorkbook("../CountryFacts.xlsm")  ## "../" finds file one folder up

Index <- openxlsx::readWorkbook(inputs, sheet = "Index")  ## first column is list of all countries
dataGeneral <- openxlsx::readWorkbook(inputs, sheet = "General", cols = 1:7)
dataTime <- openxlsx::readWorkbook(inputs, sheet = "Time")                ## Time data - used in t11 & t14
dataCdaX <- openxlsx::readWorkbook(inputs, sheet = "CdaX", startRow = 2) %>%
  rename(Commodity = `Commodity.(note:.don't.overwrite.these.labels)`)    ## Canadian Origin Exports ($Cdn)
dataBCX <- openxlsx::readWorkbook(inputs, sheet = "BCX", startRow = 2) %>%
  rename(Commodity = `Commodity.(note:.don't.overwrite.these.labels)`)    ## British Columbia Origin Exports ($Cdn)
dataCdaM <- openxlsx::readWorkbook(inputs, sheet = "CdaM", startRow = 2) %>%
  rename(Commodity = `Commodity.(note:.don't.overwrite.these.labels)`)    ## Canadian Imports ($Cdn)
dataRank <- openxlsx::readWorkbook(inputs, sheet = "Rank")                ## Exports and Imports (3 sets of tables in here)
dataProv <- openxlsx::readWorkbook(inputs, sheet = "Prov", startRow = 2)  ## Exports by Province of Origin ($Cdn)
dataWorldX <- openxlsx::readWorkbook(inputs, sheet = "WorldX", startRow = 2) %>%
  rename(SITC = `SITC.(Note:.Don't.overwrite.these.labels)`)              ## Export by Country ($US)
dataWorldM <- openxlsx::readWorkbook(inputs, sheet = "WorldM", startRow = 2) %>%
  rename(SITC = `SITC.(Note:.Don't.overwrite.these.labels)`)              ## Import by Country ($US)
dataUSXM <- openxlsx::readWorkbook(inputs, sheet = "US XM")         ## US Trade ($US); for US Country profile only
dataUSXall <- openxlsx::readWorkbook(inputs, sheet = "US Xall")     ## US Exports ($US); for US Country profile only
dataUSMall <- openxlsx::readWorkbook(inputs, sheet = "US Mall")     ## US Imports ($US); for US Country profile only
dataServices <- openxlsx::readWorkbook(inputs, sheet = "Services")
## (above) International transactions in services, by selected countries, annual (dollars x 1,000,000)
dataHighTech <- openxlsx::readWorkbook(inputs, sheet = "HighTech")        ## BC Trade in High Technology Goods
dataInvestment <- openxlsx::readWorkbook(inputs, sheet = "Investment")
## (above) International investment position, Canadian direct investment abroad and foreign direct investment in Canada, by country, annual (dollars x 1,000,000)
dataTravel <- openxlsx::readWorkbook(inputs, sheet = "Travel")
## (above) Number of non-resident travellers entering Canada through BC, by country of residence (excluding the United States), annual (persons)(1)
dataImmigration <- openxlsx::readWorkbook(inputs, sheet = "Immigration") %>%
  mutate_all(~(ifelse(is.na(.), 0, .)))  ## replace all NAs with 0
## (above) British Columbia (Intended Province of Destination) - Admissions of Permanent Residence by Select Country of Citizenship
dataGDP <- openxlsx::readWorkbook(inputs, sheet = "GDP Growth", na.strings = c("n/a", "--")) %>%
  filter(Units == "Percent change")  ## drop duplicate rows used to calculate grouped regions


## get year variables ----
years <- openxlsx::readWorkbook(inputs, "Page1", rows = 1, cols = 3:5)
yearSC <- names(years)[1]         ## year for Stats Can data: BCX/t01, CdaX/t02, CdaM/t03, Rank/t15, Prov/t06
yearUS <- names(years)[2]         ## year for US Bureau of Census data: USXM/t04&5, USXall/t16, USMall/t16
yearITC <- names(years)[3]        ## year for International Trade Centre data: WorldX/t04, WorldM/t05, World
yearHT <- get.last.year(data = dataHighTech, colVar = `BC.Trade.in.High.Technology.Goods`, rowVar = "Domestic Exports") ## t08
# yearOth <- get.last.year(data = dataTime, colVar = `TIME.SERIES.DATA`, rowVar = "BC Origin Exports")
years <- tibble::tibble(year = c(yearSC, yearUS, yearITC, yearHT),   ## yearOth
                        tab = c("yearSC", "yearUS", "yearITC", "yearHT"))  ## "yearOth"
saveRDS(years, "data/years.rds")
## others (e.g., General/t07, Investment/t09, Travel/t10, Immigration/t10, Service/t12, GDP/t13) are embedded in data


## determine Countries ----

## get list of Countries
Countries <- Index %>% filter(!is.na(CANADIAN.TRADE)) %>% select(CANADIAN.TRADE) %>% pull()
## drop non-country "World", and "Extra" if they exist
if(any(str_detect(Countries, "Extra"))) { Countries <- Countries[-which(str_detect(Countries, "Extra"))] }
if(any(str_detect(Countries, "World"))) { Countries <- Countries[-which(str_detect(Countries, "World"))] }
if(any(Countries == "China + Hong Kong")) {
  Countries[which(Countries == "China + Hong Kong")] <- "China & Hong Kong"
}
# if(any(!str_detect(Countries, "-"))) {
#   Countries[which(!str_detect(Countries, "-"))] <- paste0("xxx-", Countries[which(!str_detect(Countries, "-"))])
# }
## for use in tables where Countries are column names (spaces were replaced with dots)
# CountriesDot <- Countries %>% str_replace_all(pattern = " ", replacement = ".")
## remove three digits and dash from beginning of country names
Countries <- str_replace_all(Countries, pattern = "[0-9][0-9][0-9]-", replacement = "")
# if(any(str_detect(Countries, "xxx-"))) {
#   Countries <- str_replace_all(Countries, pattern = "xxx-", replacement = "")
# }


## for use in tables where Countries are column names (spaces were replaced with dots)
CountriesDot <- Countries %>% str_replace_all(pattern = " ", replacement = ".")
if(any(CountriesDot == "EU.+.UK")) {
  CountriesDot[which(CountriesDot == "EU.+.UK")] <- "EU.&.UK"
}


## list of US States comes from StateFacts.xlsm file


## data wrangling ----

### * drop unnecessary text rows, ensure remainder are numeric
dataBCX <- make.numeric(data = dataBCX, colVar = "Commodity")
# endRow <- which(tolower(dataBCX$Commodity) == "total")
# dataBCX <- dataBCX[1:endRow, ]; rm(endRow)
# dataBCX <- dataBCX %>% mutate_at(vars(c(names(.)[2]:(which(names(.) == "Label")-1))), as.numeric)

dataCdaM <- make.numeric(data = dataCdaM, colVar = "Commodity")
# endRow <- which(tolower(dataCdaM$Commodity) == "total")
# dataCdaM <- dataCdaM[1:endRow, ]; rm(endRow)
# dataCdaM <- dataCdaM %>% mutate_at(vars(c(names(.)[2]:(which(names(.) == "Label")-1))), as.numeric)

dataCdaX <- make.numeric(data = dataCdaX, colVar = "Commodity")
# endRow <- which(tolower(dataCdaX$Commodity) == "total")
# dataCdaX <- dataCdaX[1:endRow, ]; rm(endRow)
# dataCdaX <- dataCdaX %>% mutate_at(vars(c(names(.)[2]:(which(names(.) == "Label")-1))), as.numeric)

# dataProv <- make.numeric(data = dataProv, colVar = "Province.of.Origin")
endRow <- which(tolower(dataProv$Province.of.Origin) == "total") - 1
dataProv <- dataProv[1:endRow, ]; rm(endRow)
dataProv <- dataProv %>% mutate_at(vars(c(names(.)[2]:(which(names(.) == "Label")-1))), as.numeric)

# drop total rows from dataRank
endRow <- which(tolower(dataRank[,1]) == "total") - 1
dataRank <- dataRank[1:endRow, ]; rm(endRow)

colVar <- names(dataUSXM)[which(str_starts(tolower(names(dataUSXM)), "us"))]
endRow <- which(tolower(dataUSXM[[{{colVar}}]]) == "total")
dataUSXM <- dataUSXM[1:endRow, ]; rm(endRow, colVar)

dataWorldX <- make.numeric(data = dataWorldX, colVar = "SITC")
# endRow <- which(tolower(dataWorldX$SITC) == "total")
# dataWorldX <- dataWorldX[1:endRow, ]; rm(endRow)
# dataWorldX <- dataWorldX %>% mutate_at(vars(c(names(.)[2]:(which(names(.) == "Label")-1))), as.numeric)


### * separate out time series datasets
## if you get this error "Error in initialize(...) : attempt to use zero-length variable name", you
##   likely have a column in dataTime that has no name :(
TimeBCX <- rows.to.colnames(data = dataTime, rowAsNames = TRUE, searchCol = TIME.SERIES.DATA,
                            startRow = "BC Origin Exports", endRow = "Canada Origin Exports")
## if get error "Error in initialize(...) : attempt to use zero-length variable name", do this:
# if(any(is.na(names(TimeBCX)))) {  TimeBCX <- TimeBCX %>% select(!which(is.na(names(TimeBCX)))) }
TimeBCX <- TimeBCX %>%
  mutate(`BC Origin Exports` = case_when(`BC Origin Exports` == "United Arab Emirates" ~ "United Arab Emir.",
                                         `BC Origin Exports` == "EU + United Kingdom" ~ "EU + UK",
                                      TRUE ~ as.character(`BC Origin Exports`)))

TimeCX <- rows.to.colnames(data = dataTime, rowAsNames = TRUE, searchCol = TIME.SERIES.DATA,
                           startRow = "Canada Origin Exports", endRow = "Canada Total Exports") %>%
  mutate(`Canada Origin Exports` = case_when(`Canada Origin Exports` == "United Arab Emirates" ~ "United Arab Emir.",
                                             `Canada Origin Exports` == "EU + United Kingdom" ~ "EU + UK",
                                      TRUE ~ as.character(`Canada Origin Exports`)))

TimeCTX <- rows.to.colnames(data = dataTime, rowAsNames = TRUE, searchCol = TIME.SERIES.DATA,
                            startRow = "Canada Total Exports", endRow = "Canada Imports") %>%
  mutate(`Canada Total Exports` = case_when(`Canada Total Exports` == "United Arab Emirates" ~ "United Arab Emir.",
                                            `Canada Total Exports` == "EU + United Kingdom" ~ "EU + UK",
                                            TRUE ~ as.character(`Canada Total Exports`)))

TimeCM <- rows.to.colnames(data = dataTime, rowAsNames = TRUE, searchCol = TIME.SERIES.DATA,
                           startRow = "Canada Imports", endRow = "BC Exports") %>%
  mutate(`Canada Imports` = case_when(`Canada Imports` == "United Arab Emirates" ~ "United Arab Emir.",
                                      `Canada Imports` == "EU + United Kingdom" ~ "EU + UK",
                                            TRUE ~ as.character(`Canada Imports`)))

getRow <- which(dataTime$TIME.SERIES.DATA == "Canada/US Exchange")
TimeXchg <- dataTime %>%
  slice((getRow-1):getRow) %>%
  mutate(TIME.SERIES.DATA = case_when(is.na(TIME.SERIES.DATA) ~ "Year", TRUE ~ as.character(TIME.SERIES.DATA))) %>%
  janitor::row_to_names(row_number = 1)
rm(getRow)


### * remove three digit number and dash in front of country names
names(dataCdaX)[2:(which(names(dataCdaX) == "Label")-1)] <- names(dataCdaX)[2:(which(names(dataCdaX) == "Label")-1)] %>%
  str_sub(., start = 5)
names(dataCdaX)[which(names(dataCdaX) == "rld")] <- "World"

names(dataBCX)[2:(which(names(dataBCX) == "Label")-1)] <- names(dataBCX)[2:(which(names(dataBCX) == "Label")-1)] %>%
  str_sub(., start = 5)
names(dataBCX)[which(names(dataBCX) == "rld")] <- "World"

if(any(names(dataCdaM) == "Extra")) {
  names(dataCdaM)[2:(which(names(dataCdaM) == "Extra")-1)] <- names(dataCdaM)[2:(which(names(dataCdaM) == "Extra")-1)] %>%
    str_sub(., start = 5)
} else {
  names(dataCdaM)[2:(which(names(dataCdaM) == "Label")-1)] <- names(dataCdaM)[2:(which(names(dataCdaM) == "Label")-1)] %>%
    str_sub(., start = 5)
}
names(dataCdaM)[which(names(dataCdaM) == "rld")] <- "World"

if(any(names(dataProv) == "Extra")) {
  names(dataProv)[2:(which(names(dataProv) == "Extra")-1)] <- names(dataProv)[2:(which(names(dataProv) == "Extra")-1)] %>%
    str_sub(., start = 5)
} else {
  names(dataProv)[2:(which(names(dataProv) == "Label")-1)] <- names(dataProv)[2:(which(names(dataProv) == "Label")-1)] %>%
    str_sub(., start = 5)
}
names(dataProv)[which(names(dataProv) == "rld")] <- "World"


### * manually make country names match Index$CANADIAN.TRADE
dataGeneral <- dataGeneral %>%
  mutate(X2 = case_when(X2 == "Czech Republic (Czechia)" ~ "Czech Republic",
                        X2 == "United Arab Emirates" ~ "United Arab Emir.",
                        TRUE ~ as.character(X2)))

dataInvestment <- dataInvestment %>%
  mutate(X3 = case_when(X3 == "People's Republic of China" ~ "Mainland China",
                        X3 == "Russian Federation" ~ "Russia",
                        X3 == "United Arab Emirates" ~ "United Arab Emir.",
                        X3 == "Viet Nam" ~ "Vietnam",
                        TRUE ~ as.character(X3)))

dataServices <- dataServices %>%
  ## remove all footnote digits from country name, and a few other fixes
  mutate(X3 = str_trim(str_replace_all(X3, pattern = "[:digit:]", "")),
         X3 = case_when(X3 == "China" ~ "Mainland China",
                        X3 == "Hong Kong, China" ~ "Hong Kong",
                        X3 == "Republic of Korea" ~ "South Korea",
                        X3 == "Viet Nam" ~ "Vietnam",
                        X3 == "United Arab Emirates" ~ "United Arab Emir.",
                        TRUE ~ as.character(X3)))

dataGDP <- dataGDP %>%
  mutate(Country = case_when(Country == "Hong Kong SAR" ~ "Hong Kong",
                             Country == "China" ~ "Mainland China",
                             Country == "Slovak Republic" ~ "Slovakia",
                             Country == "Korea" ~ "South Korea",
                             Country == "Taiwan Province of China" ~ "Taiwan",
                             Country == "United Arab Emirates" ~ "United Arab Emir.",
                             Country == "CHINA & HONG KONG" ~ "China & Hong Kong",
                             TRUE ~ as.character(Country)))

dataImmigration <- dataImmigration %>% janitor::row_to_names(row_number = 1)
## last column might end up with no name (i.e., be NA) or as 0
names(dataImmigration)[(which(is.na(names(dataImmigration))))] <- "zzz"
names(dataImmigration)[(which(names(dataImmigration) == "0"))] <- "zz0"
if(any(str_detect(names(dataImmigration), "zz"))) {
  dataImmigration <- dataImmigration %>% select(-contains("zz"))
}
dataImmigration <- dataImmigration %>%
  mutate(`Country of Citizenship` = str_replace(`Country of Citizenship`, pattern = ", Republic of", replacement = ""),
         `Country of Citizenship` =
           case_when(`Country of Citizenship` == "United States of America" ~ "United States",
                     `Country of Citizenship` == "China and Hong Kong" ~ "China & Hong Kong",
                     `Country of Citizenship` == "Brunei" ~ "Brunei Darussalam",
                     `Country of Citizenship` == "China, People's Republic of" ~ "Mainland China",
                     `Country of Citizenship` == "Germany, Federal Republic of" ~ "Germany",
                     `Country of Citizenship` == "Hong Kong SAR" ~ "Hong Kong",
                     `Country of Citizenship` == "Netherlands, The" ~ "Netherlands",
                     `Country of Citizenship` == "Slovak Republic" ~ "Slovakia",
                     `Country of Citizenship` == "Korea" ~ "South Korea",
                     `Country of Citizenship` == "United Arab Emirates" ~ "United Arab Emir.",
                     `Country of Citizenship` == "United Kingdom and Overseas Territories" ~ "United Kingdom",
                     `Country of Citizenship` == "Vietnam, Socialist Republic of" ~ "Vietnam",
                     TRUE ~ as.character(`Country of Citizenship`)))

dataTravel <- dataTravel %>% janitor::row_to_names(row_number = 3)
## remove any columns that have name `NA` or are empty
if(any(is.na(names(dataTravel)))) {  dataTravel <- dataTravel[,-which(is.na(names(dataTravel)))]  }
dataTravel <- dataTravel %>%
  mutate(`Country of residence` = str_replace(`Country of residence`, pattern = " 3", replacement = ""),
         `Country of residence` =
           case_when(`Country of residence` == "China" ~ "Mainland China",
                     `Country of residence` == "Germany 6" ~ "Germany",
                     `Country of residence` == "Russian Federation" ~ "Russia",
                     `Country of residence` == "South Africa, Republic of" ~ "South Africa",
                     `Country of residence` == "Korea, South" ~ "South Korea",
                     `Country of residence` == "Turkey 7" ~ "Turkey",
                     `Country of residence` == "United Arab Emirates" ~ "United Arab Emir.",
                     `Country of residence` == "United States residents entering Canada (2)" ~ "United States",
                     `Country of residence` == "Viet Nam 4" ~ "Vietnam",
                     `Country of residence` == "Viet Nam" ~ "Vietnam",
                     TRUE ~ as.character(`Country of residence`)))
## create combined country "China & Hong Kong"  and "EU + UK" rows in dataTravel
if(!any(dataTravel$`Country of residence` == "China & Hong Kong", na.rm = TRUE)) {
  dataTravel <- make.row.in.Travel(data = dataTravel, countryVar = "Country of residence",
                                   grpdCountry = "China & Hong Kong",
                                   individualCountries = c("Hong Kong", "Mainland China"))
}
if(!any(dataTravel$`Country of residence` == "EU + UK", na.rm = TRUE)) {
  dataTravel <- make.row.in.Travel(data = dataTravel, countryVar = "Country of residence",
                                   grpdCountry = "EU + UK",
                                   individualCountries = c("European Union", "United Kingdom"))
}


dataWorldM <- dataWorldM %>%
  rename("Brunei.Darussalam" = "Brunei",
         "United.Arab.Emir." = "United.Arab.Emirates",
         "China.&.Hong.Kong" = "China+HK",
         "European.Union" = "EU",
         "EU.&.UK" = "EU+UK")

dataWorldX <- dataWorldX %>%
  rename("Brunei.Darussalam" = "Brunei",
         "United.Arab.Emir." = "United.Arab.Emirates",
         "China.&.Hong.Kong" = "China+HK",
         "European.Union" = "EU",
         "EU.&.UK" = "EU+UK")


### * create World (spread TOTAL row by country)
temp <- dataWorldM %>% filter(SITC == "TOTAL") %>% mutate(Var = "WorldM")
World <- dataWorldX %>% filter(SITC == "TOTAL") %>% mutate(Var = "WorldX") %>%
  bind_rows(temp) %>%
  select(Var, SITC:`EU.&.UK`, -SITC) %>%
  pivot_longer(-c("Var"), names_to = "Country", values_to = "value") %>%
  pivot_wider(names_from = "Var", values_from = "value") %>%
  mutate(Country = str_replace_all(Country, pattern = "[.]", replacement = " "),
         Country = str_replace(Country, pattern = "Emir ", replacement = "Emir."))
rm(temp)


### * create combined country "China & Hong Kong"  and "EU + UK" rows in dataInvestment
if(!any(dataInvestment$X3 == "China & Hong Kong", na.rm = TRUE)) {
  dataInvestment <- make.row.in.Investment(data = dataInvestment, countryVar = "X3",
                                           grpdCountry = "China & Hong Kong",
                                           individualCountries = c("Hong Kong", "Mainland China"))
}
if(!any(dataInvestment$X3 == "EU + UK", na.rm = TRUE)) {
  dataInvestment <- make.row.in.Investment(data = dataInvestment, countryVar = "X3",
                                           grpdCountry = "EU + UK",
                                           individualCountries = c("European Union", "United Kingdom"))
}


### * create combined country "China & Hong Kong"  and "EU + UK" rows in dataServices
if(!any(dataServices$X3 == "China & Hong Kong", na.rm = TRUE)) {
  dataServices <- make.row.in.Services(data = dataServices, countryVar = "X3",
                                       grpdCountry = "China & Hong Kong",
                                       individualCountries = c("Hong Kong", "Mainland China"))
}
if(!any(dataServices$X3 == "EU + UK", na.rm = TRUE)) {
  dataServices <- make.row.in.Services(data = dataServices, countryVar = "X3",
                                       grpdCountry = "EU + UK",
                                       individualCountries = c("European Union", "United Kingdom"))
}
## Note from Dan S: we can't do ASEAN services. You will have to suppress that chart like you
##                  do with the others that don't have data.


### * round ("China & Hong Kong" and "EU + UK") rows in dataGeneral
dataGeneral <- dataGeneral %>%
  mutate(`Pop.Growth.%` = janitor::round_half_up(`Pop.Growth.%`, digits = 2),
         `Real.GDP.Growth.%`= janitor::round_half_up(`Real.GDP.Growth.%`, digits = 1))


## build Aggregate Data for B.C. & Canada main page tables ----
table1 <- top5.Cda.table(data = dataBCX, country = `World`)     ## Top five B.C. origin exports to world
table2 <- top5.Cda.table(data = dataCdaX, country = `World`)    ## Top five Canadian exports to world
table3 <- top5.Cda.table(data = dataCdaM, country = `World`)    ## Top five Canadian imports from world
table4 <- prov.table(data = dataProv, country = `World`)        ## Provincial distribution of exports to world

saveRDS(table1, "data/table1.rds")
saveRDS(table2, "data/table2.rds")
saveRDS(table3, "data/table3.rds")
saveRDS(table4, "data/table4.rds")


## testing ----

### * Australia tables
### * Page 1 tables
# t01Aus <- top5.Cda.table(data = dataBCX, country = `Australia`)   ## Top 5 BC Origin Exports to Australia
# t02Aus <- top5.Cda.table(data = dataCdaX, country = `Australia`)  ## Top 5 Canadian Exports to Australia
# t03Aus <- top5.Cda.table(data = dataCdaM, country = `Australia`)  ## Top 5 Canadian Imports from Australia
# t04Aus <- top5.US.table(data = dataWorldX, var = SITC, country = `Australia`) ## Top 5 Exports from Australia to the Rest of the World
# t05Aus <- top5.US.table(data = dataWorldM, var = SITC, country = `Australia`) ## Top 5 Imports into Australia from the Rest of the World
# t06Aus <- prov.table(data = dataProv, country = `Australia`)      ## Provincial Distribution of Exports to Australia

### * Page 2 tables
# t07Aus <- gen.table(data = dataGeneral, country = "Australia")         ## Australia General Information
# t08Aus <- high.tech.table(data = dataHighTech, country = "Australia",  ## BC's High Technology Trade with Australia
#                          year = yearHT)
# t09Aus <- invest.table(data = dataInvestment, country = "Australia")   ## Canada's Investment Position with Australia
# t10Aus <- travel.table(dataTravel, dataIm = dataImmigration, country = "Australia") ## Travellers from Australia Entering Canada Through BC and Immigration to BC from Australia
# t11Aus <- time.trade(BCX = TimeBCX, CX = TimeCX, CTX = TimeCTX, CM = TimeCM, country = "Australia")
# t12Aus <- time.service(dataServices, country = "Australia")
# t13Aus <- time.gdp(dataGDP, country = "Australia")
# t14Aus <- USxchg.table(time_trade = t11Aus, TimeXchg = TimeXchg, country = "Australia")

### * importance of Canada and country to each other (Page 1)
## How important is {{country}} to BC and Canada?
# t15Aus <- rank.bind(data = dataRank, country = "Australia")
## How important is Canada to {{country}}?
## Canada was the source of approximately XX1% of imports into {{country}} in 2018.
## Approximately XX2% of exports from {{country}} were destined for Canada in that year.
## XX1 = 100 * Time!{{year=2018}}{{Canada Exports $US}} / total of top 5 Imports into {{country}} from rest of world
## XX2 = 100 * Time!{{year=2018}}{{Canada Imports $US}} / total of top 5 Exports from {{country}} to rest of world
## Note: Canada EXPORTS are IMPORTS to other country, and vice versa.
# t16Aus <- imp.table(World, t14Aus, nested = FALSE, year = yearITC, country = "Australia")


### * ASEAN tables
### * Page 1 tables
# t01AS <- top5.Cda.table(data = dataBCX, country = `ASEAN`)   ## Top 5 BC Origin Exports to ASEAN
# t02As <- top5.Cda.table(data = dataCdaX, country = `ASEAN`)  ## Top 5 Canadian Exports to ASEAN
# t03As <- top5.Cda.table(data = dataCdaM, country = `ASEAN`)  ## Top 5 Canadian Imports from ASEAN
# t04As <- top5.US.table(data = dataWorldX, var = SITC, country = `ASEAN`) ## Top 5 Exports from ASEAN to the Rest of the World
# t05As <- top5.US.table(data = dataWorldM, var = SITC, country = `ASEAN`) ## Top 5 Imports into ASEAN from the Rest of the World
# t06As <- prov.table(data = dataProv, country = `ASEAN`)      ## Provincial Distribution of Exports to ASEAN

### * Page 2 tables
# t07As <- gen.table(data = dataGeneral, country = "ASEAN")         ## ASEAN General Information
# t08As <- high.tech.table(data = dataHighTech, country = "ASEAN",  ## BC's High Technology Trade with ASEAN
#                          year = yearHT)
# t09As <- invest.table(data = dataInvestment, country = "ASEAN")   ## Canada's Investment Position with ASEAN
# t10As <- travel.table(dataTravel, dataIm = dataImmigration, country = "ASEAN") ## Travellers from ASEAN Entering Canada Through BC and Immigration to BC from ASEAN
# t11As <- time.trade(BCX = TimeBCX, CX = TimeCX, CTX = TimeCTX, CM = TimeCM, country = "ASEAN")
# t12As <- time.service(dataServices, country = "")                 ## no such data for ASEAN
# t13As <- time.gdp(dataGDP, country = "ASEAN")                     ## needed for chart
# t14As <- USxchg.table(time_trade = t11As, TimeXchg = TimeXchg, country = "ASEAN")

### * importance of Canada and country to each other (Page 1)
# t15As <- rank.bind(data = dataRank, country = "ASEAN")
# t16As <- imp.table(World, t14As, nested = FALSE, year = yearITC, country = "ASEAN")


### * China & Hong Kong tables
# t01CHK <- top5.Cda.table(data = dataBCX, country = `China.&.Hong.Kong`)   ## Top 5 BC Origin Exports to China & Hong Kong
# t02CHK <- top5.Cda.table(data = dataCdaX, country = `China.&.Hong.Kong`)  ## Top 5 Canadian Exports to China & Hong Kong
# t03CHK <- top5.Cda.table(data = dataCdaM, country = `China.&.Hong.Kong`)  ## Top 5 Canadian Imports from China & Hong Kong
# t04CHK <- top5.US.table(data = dataWorldX, var = SITC, country = `China.&.Hong.Kong`) ## Top 5 Exports from China & Hong Kong to the Rest of the World
# t05CHK <- top5.US.table(data = dataWorldM, var = SITC, country = `China.&.Hong.Kong`) ## Top 5 Imports into China & Hong Kong from the Rest of the World
# t06CHK <- prov.table(data = dataProv, country = `China.&.Hong.Kong`)      ## Provincial Distribution of Exports to China & Hong Kong
# t07CHK <- gen.table(data = dataGeneral, country = "China & Hong Kong")         ## China & Hong Kong General Information
# t08CHK <- high.tech.table(data = dataHighTech, country = "China & Hong Kong",  ## BC's High Technology Trade with China & Hong Kong
#                          year = yearHT)
# t09CHK <- invest.table(data = dataInvestment, country = "China & Hong Kong")   ## Canada's Investment Position with China & Hong Kong
# t10CHK <- travel.table(dataTravel, dataIm = dataImmigration, country = "China & Hong Kong") ## Travellers from China & Hong Kong Entering Canada Through BC and Immigration to BC from China & Hong Kong
# t11CHK <- time.trade(BCX = TimeBCX, CX = TimeCX, CTX = TimeCTX, CM = TimeCM, country = "China & Hong Kong")
# t12CHK <- time.service(dataServices, country = "China & Hong Kong")
# t13CHK <- time.gdp(dataGDP, country = "China & Hong Kong")
# t14CHK <- USxchg.table(time_trade = t11CHK, TimeXchg = TimeXchg, country = "China & Hong Kong"); rm(t11CHK)
## t15CHK <- rank.bind(data = dataRank, country = "China & Hong Kong")  ## different columns, so run separately
# t16CHK <- imp.table(World, t14, nested = TRUE, year = yearITC, country = "China & Hong Kong")


## NOT NEEDED ANYMORE: different US country data ----

## Some US country data is in a few different tabs, so run separately and replace incorrect data pulled from dataWorldX/M

# dataUS <- dataUSXM %>%
#   select(names(.)[1:3]) %>%
#   janitor::row_to_names(row_number = 1) %>%
#   mutate(Exports = as.numeric(Exports), Imports = as.numeric(Imports))
#
# ## Top 5 Exports from US to the Rest of the World
# t04US <- top5.US.table(data = dataUS, var = Commodity, country = "Exports") %>%
#   mutate(Country = "\"United.States\"") %>%
#   mutate(var = str_replace(var, pattern = " -", replacement = " "))
#
# ## Top 5 Imports into US from the Rest of the World
# t05US <- top5.US.table(data = dataUS, var = Commodity, country = "Imports") %>%
#   mutate(Country = "\"United.States\"") %>%
#   mutate(var = str_replace(var, pattern = " -", replacement = " "))
#
# ## Canada amount of US Exports and Imports, for t16
# USX <- dataUSXall %>%
#   janitor::row_to_names(row_number = 1) %>%
#   filter(Country == "CANADA") %>%
#   mutate(Exports = as.numeric(Exports)) %>%
#   pull(Exports)
#
# USM <- dataUSMall %>%
#   janitor::row_to_names(row_number = 1) %>%
#   filter(Country == "CANADA") %>%
#   mutate(Imports = as.numeric(Imports)) %>%
#   pull(Imports)
#
# t16US <- dataUS %>%
#   ## get TOTAL exports and imports for US
#   filter(Commodity == "TOTAL") %>%
#   mutate(Commodity = "United States",                                ## will become Country col
#          ## divide {{country}} exports with Canadian imports, and {{country}} imports with Canadian exports to get share
#          `Imports into Canada` = 100 * USX / Exports,        ## USX = US exports to Canada
#          `Exports from Canada` = 100 * USM / Imports) %>%    ## USM = US imports from Canada
#   ## drop unnecessary vars
#   select(-Exports, -Imports) %>%
#   ## gather by Var
#   pivot_longer(-c("Commodity"), names_to = "Var", values_to = "Share") %>%
#   rename(Country = Commodity)                                        ## properly name Country col
#
# rm(dataUS, USM, USX)


## create country-nested datafile ----
## build mega-data with tables and chart info, nested by country

## 1. run tables for all countries
t01 <- map_df(CountriesDot, top5.Cda.table, data = dataBCX) %>% nest(t01 = -Country)
t02 <- map_df(CountriesDot, top5.Cda.table, data = dataCdaX) %>% nest(t02 = -Country)
t03 <- map_df(CountriesDot, top5.Cda.table, data = dataCdaM) %>% nest(t03 = -Country)
t04 <- map_df(CountriesDot, top5.US.table, data = dataWorldX, var = SITC) %>%
  # ## US data now (2024) in with all others, so do not remove and add from separate sheet
  # ## remove incorrect US data
  # filter(str_detect(Country, pattern = "United.States", negate = TRUE)) %>%
  # ## add in correct US data
  # bind_rows(t04US) %>%
  nest(t04 = -Country)
t05 <- map_df(CountriesDot, top5.US.table, data = dataWorldM, var = SITC) %>%
  # ## US data now (2024) in with all others, so do not remove and add from separate sheet
  # ## remove incorrect US data
  # filter(str_detect(Country, pattern = "United.States", negate = TRUE)) %>%
  # ## add in correct US data
  # bind_rows(t05US) %>%
  nest(t05 = -Country)
t06 <- map_df(CountriesDot, prov.table, data = dataProv) %>% nest(t06 = -Country)
t07 <- map_df(Countries, gen.table, data = dataGeneral %>% select(-X1)) %>% nest(t07 = -Country)
t08 <- map_df(Countries, high.tech.table, data = dataHighTech, year = yearHT) %>%
  nest(t08 = -Country)  ## ignore warnings: Column `Year` has different attributes on LHS and RHS of join
t09 <- map_df(Countries, invest.table, data = dataInvestment) %>% nest(t09 = -Country)
t10 <- map_df(Countries, travel.table, dataTravel = dataTravel, dataIm = dataImmigration) %>% nest(t10 = -Country) ## ignore warnings: Column `Year` has different attributes on LHS and RHS of join
t11 <- map_df(Countries, time.trade, BCX = TimeBCX, CX = TimeCX, CTX = TimeCTX, CM = TimeCM) %>% nest(t11 = -Country)
temp <- dataServices$X3 %>% unique()
CountriesSvcs <- intersect(Countries, temp); rm(temp)
t12 <- map_df(CountriesSvcs, time.service, data = dataServices) %>%
  nest(t12 = -Country)  ## ignore warnings: NAs introduced by coercion
t13 <- map_df(Countries, time.gdp, data = dataGDP) %>% nest(t13 = -Country)
## run for first country, then bind in remaining countries, then nest
temp <- t11 %>% unnest(cols = c(t11)) %>% filter(Country == Countries[1])
t14 <- USxchg.table(time_trade = temp, TimeXchg = TimeXchg, country = Countries[1]); rm(temp)
for(i in 2:length(Countries)) {
  temp <- t11 %>% unnest(cols = c(t11)) %>% filter(Country == Countries[i])
  t14 <- t14 %>% bind_rows(USxchg.table(time_trade = temp, TimeXchg = TimeXchg, country = Countries[i]))
}        ## ignore Warning: Column `Year` has different attributes on LHS and RHS of join
rm(i, temp)
t14 <- t14 %>% nest(t14 = -Country)
# old way (when grouped countries were in column at very end of sheet)
## run grouped Countries separately for t15, b/c data is in different columns
# grpdCountries <- dataRank %>% select("X22") %>%
#   filter(!is.na(X22) & str_detect(X22, pattern = "Country", negate = TRUE)) %>%
#   unique() %>% rename(Country = X22) %>% pull()   ## get vector of grouped countries
grpdCountries <- dataRank %>% filter(!is.na(X4), is.na(X5)) %>%
  select(names(dataRank)[1]) %>% pull()  ## get vector of grouped countries
saveRDS(grpdCountries, "data/grpdCountries.rds")
t15grpdCs <- map_df(grpdCountries, rank.data, data = dataRank, cols = 1:4, var = `BC Exports`, grpdCountries = grpdCountries, grpdCountry = TRUE) %>%
  bind_rows(map_df(grpdCountries, rank.data, data = dataRank, cols = 8:11, var = `Canada Exports`, grpdCountries = grpdCountries, grpdCountry = TRUE)) %>%
  bind_rows(map_df(grpdCountries, rank.data, data = dataRank, cols = 15:18, var = `Canada Imports`, grpdCountries = grpdCountries, grpdCountry = TRUE)) %>%
  mutate(Perc = paste0(janitor::round_half_up(Perc, digits = 1), "%")) %>%
  mutate(Var = case_when(Var == "`BC Exports`" ~ "BC Exports",
                         Var == "`Canada Exports`" ~ "Canada Exports",
                         Var == "`Canada Imports`" ~ "Canada Imports"))
t15 <- map_df(Countries, rank.bind, data = dataRank, grpdCountries = grpdCountries) %>%
  bind_rows(rank.bind(data = dataRank, country = "United Arab Emirates", grpdCountries = grpdCountries)) %>%
  bind_rows(t15grpdCs) %>%
  mutate(Country = case_when(Country == "United Arab Emirates" ~ "United Arab Emir.",
                             Country == "EU & UK" ~ "EU + UK",
                             TRUE ~ as.character(Country))) %>%
  nest(t15 = -Country)
rm(grpdCountries)
rm(t15grpdCs)
## run for first country, then bind in remaining countries, then nest
t16 <- imp.table(World, t14, nested = TRUE, year = yearITC, country = Countries[1])
temp <- World; temp$Country <- str_replace(temp$Country, pattern = "EU & UK", replacement = "EU + UK")
for(i in 2:length(Countries)) {
  t16 <- t16 %>% bind_rows(imp.table(temp, t14, nested = TRUE, year = yearITC, country = Countries[i]))
}; rm(i, temp)
t16 <- t16 %>%
  # ## US data now (2024) in with all others, so do not remove and add from separate sheet
  # ## remove incorrect US data
  # filter(str_detect(Country, pattern = "United States", negate = TRUE)) %>%
  # ## add in correct US data
  # bind_rows(t16US) %>%
  nest(t16 = -Country)#; rm(t16US)


## 2. merge into mega-data
nested_data_countries <- t01 %>%
  ## take t01 and add in t02 through t06 (built with CountriesDot)
  left_join(t02, by = "Country") %>%
  left_join(t03, by = "Country") %>%
  left_join(t04, by = "Country") %>%
  left_join(t05, by = "Country") %>%
  left_join(t06, by = "Country") %>%
  ## replace dots with spaces in country names
  mutate(Country = str_replace_all(Country, pattern = "\"", replacement = ""),
         Country = str_replace_all(Country, pattern = "[.]", replacement = " "),
         Country = str_replace_all(Country, pattern = "Emir ", replacement = "Emir."),
         Country = str_replace_all(Country, pattern = "EU.&.UK", replacement = "EU + UK")) %>%
  ## add in remaining tables (built with Countries and CountriesSvcs)
  left_join(t07, by = "Country") %>%
  left_join(t08, by = "Country") %>%
  left_join(t09, by = "Country") %>%
  left_join(t10, by = "Country") %>%
  left_join(t11, by = "Country") %>%
  left_join(t12, by = "Country") %>%  ## built with CountriesSvcs
  left_join(t13, by = "Country") %>%
  left_join(t14, by = "Country") %>%
  left_join(t15, by = "Country") %>%
  left_join(t16, by = "Country")

saveRDS(nested_data_countries, "data/nested_data_countries.rds")

