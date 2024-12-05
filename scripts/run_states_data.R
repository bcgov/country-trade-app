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
## Date: Sept. 2020
## Coder: Julie Hawkins (and Stephanie Yurchak for app)

## ** This file creates the US states nested data. **


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


### * divide.by.million() function divides column by 1,000,000
divide.by.million <- function(var) { var = var/1000000  }


### * multiply.by.thousand() function divides column by 1,000,000
multiply.by.thousand <- function(var) { var = var*1000  }


### * top5.Cda.table() function pulls top 5 commodities in CDN $Millions and Share of Total %, for specified state
top5.Cda.table <- function(data, state) {

  ## get TOTAL row temporarily
  temp <- data %>%
    rename(state = {{state}}) %>%
    select(Commodity, state) %>%
    filter(Commodity == "TOTAL") %>%
    mutate(Commodity = "Total")

  data %>%
    rename(state = {{state}}) %>%
    ## get just Commodity and state columns
    select(Commodity, state) %>%
    ## temporarily drop TOTAL row
    filter(Commodity != "TOTAL") %>%
    ## sort descending
    arrange(desc(state)) %>%
    ## get just top 5 commodities
    slice(1:5) %>%
    ## add number (`row.names(.)`) and remove 4 digits and dash at beginning of Commodity: e.g., 1. Crude petroleum...
    mutate(Commodity = paste0(row.names(.), ". ", str_sub(Commodity, start = 6))) %>%
    ## bring back TOTAL row
    bind_rows(temp) %>%
    ## divide by TOTAL, multiply by 100, round to whole number and add percent sign
    mutate(Share.of.Total = paste0(janitor::round_half_up(100*(state / temp$state), digits = 0), "%")) %>%
    ## convert to millions, one decimal place, add comma separator and dollar sign
    mutate(state = paste0("$", prettyNum(round_half_up(state/1000000, digits = 1), big.mark = ","))) %>%
    rename(`Cdn $Millions` = state, `Share of Total` = Share.of.Total) %>%
    ## add in state name as first column
    mutate(state = rlang::quo_text(enquo(state))) %>%
    select(state, everything())

}


### * top5.US.table() function pulls top 5 "var"s in US $Millions and Share of Total %, for specified state
top5.US.table <- function(data, var, state) {

  ## get TOTAL row temporarily
  temp <- data %>%
    rename(var = {{var}}, state = {{state}}) %>%
    select(var, state) %>%
    filter(var == "TOTAL") %>%
    mutate(var = "Total")

  data %>%
    rename(var = {{var}}, state = {{state}}) %>%
    ## get just var and state columns
    select(var, state) %>%
    ## temporarily drop TOTAL row
    filter(var != "TOTAL") %>%
    ## sort descending
    arrange(desc(state)) %>%
    ## get just top 5 commodities
    slice(1:5) %>%
    ## add number (`row.names(.)`) and remove 3 digits, space and underscore at beginning of var
    mutate(var = paste0(row.names(.), ". ", str_sub(var, start = 6))) %>%
    ## bring back TOTAL row
    bind_rows(temp) %>%
    ## divide by TOTAL, multiply by 100, round to whole number and add percent sign
    mutate(Share.of.Total = paste0(janitor::round_half_up(100*(state / temp$state), digits = 0), "%")) %>%
    ## convert to millions, one decimal place, add comma separator and dollar sign
    mutate(state = paste0("$", prettyNum(round_half_up(state/1000000, digits = 1), big.mark = ","))) %>%
    rename(`US $Millions` = state, `Share of Total` = Share.of.Total) %>%
    ## add in state name as first column
    mutate(state = rlang::quo_text(enquo(state))) %>%
    select(state, everything())

}


### * prov.table() function gets provincial distribution of exports to specified state
prov.table <- function(data, state) {
  stateTotal <- data %>% select({{state}}) %>% sum(na.rm = TRUE)

  data %>%
    ## rename state column
    rename(state = {{state}}) %>%
    ## get just Province.of.Origin and state columns
    select(Province.of.Origin, state) %>%
    ## drop any NA rows
    filter(!is.na(Province.of.Origin)) %>%
    arrange(desc(state)) %>%
    ## add number (`row.names(.)`) before Province name
    mutate(Province.of.Origin = paste0(row.names(.), ". ", Province.of.Origin)) %>%
    ## add in total row
    add_row(Province.of.Origin = "Canada Total", state = stateTotal) %>%
    ## divide by TOTAL, multiply by 100, round to whole number and add percent sign
    mutate(Share.of.Total = paste0(janitor::round_half_up(100*(state / stateTotal), digits = 1), "%")) %>%
    ## convert to millions, one decimal place, add comma separator and dollar sign
    mutate(state = paste0("$", prettyNum(round_half_up(state/1000000, digits = 1), big.mark = ","))) %>%
    rename(`Province of Origin` = Province.of.Origin, `Cdn $Millions` = state, `Share of Total` = Share.of.Total) %>%
    ## add in state name as first column
    mutate(state = rlang::quo_text(enquo(state))) %>%
    select(state, everything())
}


### * gen.table() function gets general information, for specified state
gen.table <- function(data, state) {

  data <- data %>%
    rename(GDP_chained = starts_with("GDP.(US.$Millions.Chained"),
           real_chained = starts_with("Real.Per.Capita.GDP.(Chained"))

  temp <- data %>% filter(X2 == {{state}})

  ## pull stats
  pop <- temp %>% select(Population) %>% pull() %>% prettyNum(big.mark = ",")
  apgr <- temp %>% select(`Pop.Growth.%`) %>%
    mutate(`Pop.Growth.%` = janitor::round_half_up(as.numeric(`Pop.Growth.%`), digits = 1)) %>%
    pull() %>% paste0(., "%")
  # gdp <- temp %>% select(`GDP.(US.$Millions.Chained.2012)`) %>%
  #   mutate(`GDP.(US.$Millions.Chained.2012)` = janitor::round_half_up(as.numeric(`GDP.(US.$Millions.Chained.2012)`), digits = 0)) %>%
  #   pull() %>% prettyNum(big.mark = ",") %>% paste0("$", .)
  # gdp2 <- temp %>% select(`Real.Per.Capita.GDP.(Chained.2012)`) %>%
  #   mutate(`Real.Per.Capita.GDP.(Chained.2012)` = janitor::round_half_up(as.numeric(`Real.Per.Capita.GDP.(Chained.2012)`), digits = 0)) %>%
  #   pull() %>% prettyNum(big.mark = ",") %>% paste0("$", .)
  gdp <- temp %>% select(GDP_chained) %>%
    mutate(GDP_chained = janitor::round_half_up(as.numeric(GDP_chained), digits = 0)) %>%
    pull() %>% prettyNum(big.mark = ",") %>% paste0("$", .)
  gdp2 <- temp %>% select(real_chained) %>%
    mutate(real_chained = janitor::round_half_up(as.numeric(real_chained), digits = 0)) %>%
    pull() %>% prettyNum(big.mark = ",") %>% paste0("$", .)
  gdp3 <- temp %>% select(`GDP.Growth.%`) %>%
    mutate(`GDP.Growth.%` = janitor::round_half_up(as.numeric(`GDP.Growth.%`), digits = 1)) %>%
    pull() %>% paste0(., "%")

  ## pull last year data was available
  yrs <- data %>%
    slice(1) %>%
    select(Population:`GDP.Growth.%`) %>%
    mutate_if(is.numeric, as.character) %>%
    pivot_longer(everything(), names_to = "Var", values_to = "Year", values_drop_na = TRUE)

  ## put all into table, and add in state name as first column
  tibble::tibble(value = c(pop, apgr, gdp, gdp2, gdp3), yrs) %>%
    select(Var, value, Year) %>%
    mutate(Var = case_when(Var == "Pop.Growth.%" ~ "Annual Population Growth Rate",
                           Var == "GDP_chained" ~ "GDP (Chained US$ Millions)",
                           # Var == "GDP.(US.$Millions.Chained.2012)" ~ "GDP (Chained US$ Millions)",
                           Var == "real_chained" ~ "Real Per Capita GDP (US$)",
                           # Var == "Real.Per.Capita.GDP.(Chained.2012)" ~ "Real Per Capita GDP (US$)",
                           Var == "GDP.Growth.%" ~ "GDP Real Growth",
                           TRUE ~ as.character(Var))) %>%
    mutate(state = {{state}}) %>%
    select(state, everything())

}


### * high.tech.table() function gets High Technology Trade data, for specified state and years
high.tech.table <- function(data, state) {

  ## get Re-Exports info
  RX <- rows.to.colnames(data = data, searchCol = BC.Trade.in.High.Technology.Goods,
                         startRow = "Re-Exports", endRow = "Imports") %>%
    filter(`Re-Exports` == {{state}}) %>%
    select(-`Re-Exports`) %>%
    mutate_if(is.character, as.numeric) %>%
    pivot_longer(everything(), names_to = "Year", values_to = "Re-Exports")

  ## get Imports info
  Im <- data %>%
    add_row(BC.Trade.in.High.Technology.Goods = "zzz") %>%
    rows.to.colnames(searchCol = BC.Trade.in.High.Technology.Goods,
                     startRow = "Imports", endRow = "zzz") %>%
    filter(`Imports` == {{state}}) %>%
    select(-`Imports`) %>%
    mutate_if(is.character, as.numeric) %>%
    pivot_longer(everything(), names_to = "Year", values_to = "Imports")

  ## get Domestic Exports info
  rows.to.colnames(data = data, searchCol = BC.Trade.in.High.Technology.Goods,
                   startRow = "Domestic Exports", endRow = "Re-Exports") %>%
    filter(`Domestic Exports` == {{state}}) %>%
    select(-`Domestic Exports`) %>%
    mutate_if(is.character, as.numeric) %>%
    pivot_longer(everything(), names_to = "Year", values_to = "Domestic Exports") %>%
    ## add in Re-Exports and Imports info, calculate Trade Balance
    left_join(RX, by = "Year") %>%
    left_join(Im, by = "Year") %>%
    mutate(`Trade Balance` = `Domestic Exports` + `Re-Exports` - `Imports`) %>%
    mutate_if(is.numeric, divide.by.million) %>%
    mutate_if(is.numeric, janitor::round_half_up, digits = 1) %>%
    ## get specified years only
    # filter(Year >= yearStart) %>%
    filter(Year >= max(as.numeric(Year))-9) %>%
    ## add in state name as first column
    mutate(state = {{state}}) %>%
    select(state, everything())

}


### * travel.table() function gets travel data, for specified state
travel.table <- function(data, state) {

  endRow <- (which(data$`TRIPS.TO.CANADA.(000)` == "CANADIAN TRIPS TO US (000)")-1)
  startRow <- which(data$`TRIPS.TO.CANADA.(000)` == "CANADIAN TRIPS TO US (000)")
  maxYear <- names(data)[ncol(data)]

  ## pull trips from {{state}} to Canada data
  TripsToCda <- data[1:endRow, 1:which(names(data) == "Rank")] %>%
    filter(`TRIPS.TO.CANADA.(000)` == {{state}}) %>%
    mutate(Var = "TripsToCanada") %>%
    mutate_if(is.numeric, multiply.by.thousand)

  ## pull money spent on trips from {{state}} to Canada data
  SpentInCda <- data[1:endRow, which(names(data) == "SPENDING.IN.CANADA.($Cda.Millions)"):ncol(data)] %>%
    filter(`SPENDING.IN.CANADA.($Cda.Millions)` == {{state}}) %>%
    mutate(Var = "SpendingInCanada") %>%
    select(-`SPENDING.IN.CANADA.($Cda.Millions)`)

  ## pull trips to {{state}} from Canada
  TripsToState <- data[startRow:nrow(data), 1:which(names(data) == "Rank")] %>%
    janitor::row_to_names(row_number = 1) %>%
    filter(`CANADIAN TRIPS TO US (000)` == {{state}}) %>%
    mutate(Var = "TripsToState") %>%
    select(-`CANADIAN TRIPS TO US (000)`) %>%
    mutate_if(is.numeric, multiply.by.thousand)

  ## pull money spent on trips to {{state}} from Canada data
  SpentInState <- data[startRow:nrow(data), which(names(data) == "SPENDING.IN.CANADA.($Cda.Millions)"):ncol(data)] %>%
    janitor::row_to_names(row_number = 1) %>%
    filter(`SPENDING IN US ($Cda Millions)` == {{state}}) %>%
    mutate(Var = "SpendingInState") %>%
    select(-`SPENDING IN US ($Cda Millions)`)

  ## pull all together
  bind_rows(TripsToCda, SpentInCda, TripsToState, SpentInState) %>%
    ## add in state name, re-arrange columns, only include maxYear col of data
    mutate(state = {{state}}) %>%
    select(state, Var, Stat = maxYear, Rank) %>%
    add_row(state = {{state}}, Var = "Year", Stat = as.numeric(maxYear), Rank = NA)

}


### * time.trade() function gets trade data for specified state
time.trade <- function(BCX, CX, CTX, CM, state) {

  ## get exports and imports data for specified state
  BCX <- BCX  %>% mutate(Var = "BC Exports") %>%
    rename(state = `BC Origin Exports`) %>% filter(state == {{state}})
  CX <- CX %>% mutate(Var = "Canada Exports") %>%
    rename(state = `Canada Origin Exports`) %>% filter(state == {{state}})
  CTX <- CTX %>% mutate(Var = names(CTX)[1]) %>%
    rename(state = `Canada Total Exports`) %>% filter(state == {{state}})
  CM <- CM %>% mutate(Var = names(CM)[1]) %>%
    rename(state = `Canada Imports`) %>% filter(state == {{state}})

  ## bind above data together
  BCX %>%
    bind_rows(CX) %>%
    bind_rows(CTX) %>%
    bind_rows(CM) %>%
    select(Var, everything()) %>%
    mutate_at(vars(-"Var", -"state"), as.numeric) %>%
    ## flip data so Year is a column and Var becomes multiple columns
    pivot_longer(-c("Var", "state"), names_to = "Year", values_to = "value") %>%
    pivot_wider(names_from = "Var", values_from = "value") %>%
    ## calculate `Trade Balance` and `BC %Canada` columns
    mutate(`Trade Balance` = `Canada Total Exports` - `Canada Imports`,
           `BC %Canada` = 100 * `BC Exports` / `Canada Exports`) %>%
    # ## get specified years only
    # filter(Year >=  yearStart) %>%
    ## add in state name as first column
    mutate(state = {{state}}) %>%
    select(state, everything())

}


### * time.gdp() function gets GDP % growth for specified state
time.gdp <- function(data, state) {

  data[, which(names(data) == "Percent.Change"):ncol(data)] %>%
    ## get specified state row
    filter(Percent.Change == {{state}}) %>%
    rename(state = Percent.Change) %>%
    ## gather data so Year is a column
    pivot_longer(-c("state"), names_to = "Year", values_to = "Percent.Change")

}


### * rank.data() function gets importance of specified state relative to BC and Canada
rank.data <- function(data, cols, var, state){
  data[, {{cols}}] %>%
    janitor::row_to_names(row_number = 1) %>%
    select(state = names(.)[1], Value = {{var}}, Rank) %>%
    mutate(Var = rlang::quo_text(enquo(var)),
           Value = as.numeric(Value),
           Perc = 100 * Value / sum(Value, na.rm = TRUE)) %>%
    filter(state == {{state}}) %>%
    select(state, Var, Rank, Perc)

}
# rank.data(dataRankState, cols = 1:5, var = `BC Exports`, state = "Washington")
# rank.data(dataRankState, cols = 9:13, var = `Canada Exports`, state = "Washington")
# rank.data(dataRankState, cols = 17:21, var = `Canada Imports`, state = "Washington")


### * rank.as.country() function gets state's value for variable and adds it to "dataRank"
rank.as.country <- function(data, dataAsCountry, cols1, cols2, var, state) {

  stateVal <- data[, {{cols1}}] %>%
    janitor::row_to_names(row_number = 1) %>%
    select(state = names(.)[1], Value = {{var}}) %>%
    mutate(Value = as.numeric(Value)) %>%
    filter(state == {{state}}) %>%
    select(Value) %>%
    pull()

  ## pull relevant var and cols2
  temp <- dataAsCountry[, {{cols2}}] %>%
    janitor::row_to_names(row_number = 1) %>%
    select(Country = names(.)[1], Value = {{var}})

  ## get row index of row before "total" row
  lastRow <- min(which(str_detect(tolower(temp$Country), "total")))

  ## replace last row with correct State name and replace old/wrong value with state's value
  temp[lastRow, "Country"] <- paste0(state, "_State")  ## b/c Georgia is BOTH a Country and a US State
  temp[lastRow, "Value"]   <- stateVal

  temp %>%
    mutate(Value = as.numeric(Value)) %>%
    filter(str_detect(Country, pattern = "TOTAL", negate = TRUE))

}
# rank.as.country(dataRankState, dataAsCountry = dataRank, cols1 = 1:5, cols2 = 1:4, var = `BC Exports`, state = "Washington")


### * rank.bind() function runs rank.data() and binds info to gets importance of specified state relative to BC and Canada
rank.bind <- function(data, dataAsCountry, state) {

  ## run rank info for BC Exports
  tab <- rank.data(data, cols = 1:5, var = `BC Exports`, {{state}}) %>%
    ## rank info for Canada Exports
    bind_rows(rank.data(data, cols = 9:13, var = `Canada Exports`, {{state}})) %>%
    ## rank info for Canada Imports
    bind_rows(rank.data(data, cols = 17:21, var = `Canada Imports`, {{state}})) %>%
    mutate(Perc = paste0(janitor::round_half_up(Perc, digits = 1), "%"),
           Var = case_when(Var == "`BC Exports`" ~ "BC Exports",
                           Var == "`Canada Exports`" ~ "Canada Exports",
                           Var == "`Canada Imports`" ~ "Canada Imports"))

  ## create new ranks as though state was a separate country
  ## for BC Exports
  BCX <- rank.as.country(data, dataAsCountry, cols1 = 1:5, cols2 = 1:4, var = `BC Exports`, {{state}}) %>%
    arrange(desc(Value)) %>%
    mutate(Rank = 1:nrow(.)) %>%
    filter(Country == paste0(state, "_State")) %>%
    select(Rank) %>%
    pull()

  ## for Canada Exports
  CX <- rank.as.country(data, dataAsCountry, cols1 = 9:13, cols2 = 8:12, var = `Canada Exports`, {{state}}) %>%
    arrange(desc(Value)) %>%
    mutate(Rank = 1:nrow(.)) %>%
    filter(Country == paste0(state, "_State")) %>%
    select(Rank) %>%
    pull()

  ## for Canada Imports
  CM <- rank.as.country(data, dataAsCountry, cols1 = 17:21, cols2 = 15:18, var = `Canada Imports`, {{state}}) %>%
    arrange(desc(Value)) %>%
    mutate(Rank = 1:nrow(.)) %>%
    filter(Country == paste0(state, "_State")) %>%
    select(Rank) %>%
    pull()

  ## add RanksAsCountry into main tab
  tab %>%
    mutate(RankAsCountry = case_when(Var == "BC Exports" ~ BCX,
                                     Var == "Canada Exports" ~ CX,
                                     Var == "Canada Imports" ~ CM))

}


## load data ----

inputs <- openxlsx::loadWorkbook("../StateFacts.xlsm")

Index <- openxlsx::readWorkbook(inputs, sheet = "Index", colNames = F)    ## list of US states
dataGeneral <- openxlsx::readWorkbook(inputs, sheet = "General")
dataTime <- openxlsx::readWorkbook(inputs, sheet = "Time")                ## Time Series data
dataCdaX <- openxlsx::readWorkbook(inputs, sheet = "CdaX", startRow = 2) %>%
  rename(Commodity = `Commodity.(note:.don't.overwrite.these.labels)`)    ## Canadian Origin Exports ($Cdn)
dataBCX <- openxlsx::readWorkbook(inputs, sheet = "BCX", startRow = 2) %>%
  rename(Commodity = `Commodity.(note:.don't.overwrite.these.labels)`)    ## British Columbia Origin Exports ($Cdn)
dataCdaM <- openxlsx::readWorkbook(inputs, sheet = "CdaM", startRow = 2) %>%
  rename(Commodity = `Commodity.(note:.don't.overwrite.these.labels)`)    ## Canadian Imports ($Cdn)
dataRank <- openxlsx::readWorkbook(inputs, sheet = "Rank")                ## Exports and Imports (3 sets of tables in here)
dataRankState <- openxlsx::readWorkbook(inputs, sheet = "Rank-State")     ## Exports and Imports (3 sets of tables in here)
dataProv <- openxlsx::readWorkbook(inputs, sheet = "Prov", startRow = 2)  ## Exports by Province of Origin ($Cdn)
dataStateX <- openxlsx::readWorkbook(inputs, sheet = "StateX", startRow = 2, na.strings = "--") %>%
  rename(SITC = `SITC.(Note:.Don't.overwrite.these.labels)`)              ## Export by State ($US)
dataHighTech <- openxlsx::readWorkbook(inputs, sheet = "HighTech")        ## BC Trade in High Technology Goods (3 sets of tables in here)
dataGDP <- openxlsx::readWorkbook(inputs, sheet = "GDP", na.strings = c("n/a", "--"))
dataTravel <- openxlsx::readWorkbook(inputs, sheet = "Travel", startRow = 2)  ## 4 sets of tables in here
## (above) any extra notes HAVE To be below last rows, NOT as new columns, else `travel.table()` won't work
## (above) Travel by Canadians to the United States, top 15 states visited
## (above) Travellers to Canada from the United States by state of origin, top 15 states of origin


## get year variables ----
years <- openxlsx::readWorkbook(inputs, "Page1", rows = 1, cols = 3:4)
yearSC <- names(years)[1]         ## data year for: BCX/t01, CdaX/t02, CdaM/t03, Rank/t15, Prov/t06
yearUS <- names(years)[2]         ## data year for: StateX/t04
yearHT <- get.last.year(data = dataHighTech, colVar = `BC.Trade.in.High.Technology.Goods`, rowVar = "Domestic Exports") ## t08
years <- tibble::tibble(year = c(yearSC, yearUS, yearHT), tab = c("yearSC", "yearUS", "yearHT"))
saveRDS(years, "data/yearsStates.rds")
## others (e.g., General/t07, Travel/t10, GDP/t13) are embedded in data


## determine States ----

## get list of States
States <- Index %>% select(X2) %>% pull()
## drop #34 "NON-US", #41 "Puerto Rico", #45 "State Unknown", #48 "U.S. Virgin Is."
drops <- c("NON-US", "Puerto Rico", "State Unknown", "U.S. Virgin Is.")
States <- States[-which(States %in% drops)]

## for use in tables where States are column names (spaces were replaced with dots)
StatesDot <- States %>% str_replace_all(pattern = " ", replacement = ".")


## data wrangling ----

### * drop unnecessary text rows, ensure remainder are numeric
dataBCX <- make.numeric(data = dataBCX, colVar = "Commodity")
# endRow <- which(dataBCX$Commodity == "TOTAL")
# dataBCX <- dataBCX[1:endRow, ]; rm(endRow)
# dataBCX <- dataBCX %>% mutate_at(vars(c(names(.)[2]:(which(names(.) == "Label")-1))), as.numeric)

dataCdaX <- make.numeric(data = dataCdaX, colVar = "Commodity")
# endRow <- which(dataCdaX$Commodity == "TOTAL")
# dataCdaX <- dataCdaX[1:endRow, ]; rm(endRow)
# dataCdaX <- dataCdaX %>% mutate_at(vars(c(names(.)[2]:(which(names(.) == "Label")-1))), as.numeric)

dataCdaM <- make.numeric(data = dataCdaM, colVar = "Commodity")

### * manually make state names match Index$X2
dataGeneral <- dataGeneral %>%
  mutate(X2 = case_when(X2 == "Washington, state" ~ "Washington",
                        TRUE ~ as.character(X2)))

dataHighTech <- dataHighTech %>%
  mutate(BC.Trade.in.High.Technology.Goods = case_when(
    BC.Trade.in.High.Technology.Goods == "Washington, state" ~ "Washington",
    TRUE ~ as.character(BC.Trade.in.High.Technology.Goods)))

dataProv <- dataProv %>% rename(Washington = `Washington,.state`)
endRow <- min(which(is.na(dataProv$`Province.of.Origin`)))
dataProv <- dataProv[1:endRow, ]; rm(endRow)
dataProv <- dataProv %>% mutate_at(vars(c(names(.)[2]:(which(names(.) == "Label")-1))), as.numeric)


endRow <- which(str_detect(tolower(dataRankState[,1]), "total")) - 1
dataRankState <- dataRankState[1:endRow, ]; rm(endRow)
dataRankState <- dataRankState %>%
  mutate(across(where(is.character), ~case_when(.x == "Washington, state" ~ "Washington",
                              TRUE ~ as.character(.x))))

dataStateX <- dataStateX %>% rename(Washington = `Washington,.state`) %>%
  make.numeric(colVar = "SITC")

dataTime <- dataTime %>%
  mutate(TIME.SERIES.DATA = case_when(TIME.SERIES.DATA == "Washington, state" ~ "Washington",
                                      TRUE ~ as.character(TIME.SERIES.DATA)))

myRows <- which(dataGDP[, "GDP.(Millions.2017.Chained.$)"] == "District of Columbia")
dataGDP[myRows, "GDP.(Millions.2017.Chained.$)"] <- "Dist. of Columbia"
myRows <- which(dataGDP[, "Percent.Change"] == "District of Columbia")
dataGDP[myRows, "Percent.Change"] <- "Dist. of Columbia"
rm(myRows)

myRows <- which(dataTravel[, "TRIPS.TO.CANADA.(000)"] == "District of Columbia")
dataTravel[myRows, "TRIPS.TO.CANADA.(000)"] <- "Dist. of Columbia"
myRows <- which(dataTravel[, "SPENDING.IN.CANADA.($Cda.Millions)"] == "District of Columbia")
dataTravel[myRows, "SPENDING.IN.CANADA.($Cda.Millions)"] <- "Dist. of Columbia"
rm(myRows)


### * separate out time series datasets
TimeBCX <- rows.to.colnames(data = dataTime, rowAsNames = TRUE, searchCol = TIME.SERIES.DATA,
                            startRow = "BC Origin Exports", endRow = "Canada Origin Exports")

TimeCX <- rows.to.colnames(data = dataTime, rowAsNames = TRUE, searchCol = TIME.SERIES.DATA,
                           startRow = "Canada Origin Exports", endRow = "Canada Total Exports")

TimeCTX <- rows.to.colnames(data = dataTime, rowAsNames = TRUE, searchCol = TIME.SERIES.DATA,
                            startRow = "Canada Total Exports", endRow = "Canada Imports")

TimeCM <- rows.to.colnames(data = dataTime, rowAsNames = TRUE, searchCol = TIME.SERIES.DATA,
                           startRow = "Canada Imports", endRow = "BC Exports")


## testing ----

### * Washington tables
### * Page 1 tables (t01 - t04, t06)
# t01WA <- top5.Cda.table(data = dataBCX, state = `Washington`)    ## Top 5 BC Origin Exports to Washington
# t02WA <- top5.Cda.table(data = dataCdaX, state = `Washington`)   ## Top 5 Canadian Exports to Washington
# t03WA <- top5.Cda.table(data = dataCdaM, state = `Washington`)   ## Top 5 Canadian Imports from Washington
# t04WA <- top5.US.table(data = dataStateX, var = SITC, state = `Washington`) ## Top 5 Exports from Washington to the Rest of the World
## there is no t05 "Top 5 Imports into {{state}} from the Rest of the World", as there is for Countries
# t06WA <- prov.table(data = dataProv, state = `Washington`)       ## Provincial Distribution of Exports to Washington

### * Page 2 tables (t07, t08, t10 for text)
# t07WA <- gen.table(data = dataGeneral, state = "Washington")     ## Washington General Information
# t08WA <- high.tech.table(data = dataHighTech, state = "Washington") ## BC's High Technology Trade with Washington
## there is no t09 "Canada's Investment Position with {{state}}", as there is for Countries
# t10WA <- travel.table(data = dataTravel, state = "Washington")  ## Washington Travel to Canada & Canadian Travel to Washington
## t10 for Countries is used on c4 & c5; for states, it is just text, no table or charts

### * charts data (t11 for c1 & c2, t13 for c4)
# t11WA <- time.trade(BCX = TimeBCX, CX = TimeCX, CTX = TimeCTX, CM = TimeCM, state = "Washington")  ## c1, c2
## there is no t12 (needed for c3 "Canada's Balance of Trade in Services with {{state}}"), as there is for Countries
# t13WA <- time.gdp(data = dataGDP, state = "Washington")                  ## c4, NEED THIS
## t14 is only for Time series data, which does not need to be done after all

### * importance of state to Canada among states and as "country" (Page 1)

## How does {{state}} compare with other US states?
## How important is {{state}} to BC and Canada?
## t15 for states is slightly different than t15 for Countries
# t15WA <- rank.bind(data = dataRankState, dataAsCountry = dataRank, state = "Washington")
# there is no t16 "How important is Canada to {{state}}?", as there is for Countries


## create state-nested datafile ----
## build mega-data with tables and chart info, nested by state

## 1. run tables for all states
t01 <- map_df(StatesDot, top5.Cda.table, data = dataBCX) %>% nest(t01 = -state)
t02 <- map_df(StatesDot, top5.Cda.table, data = dataCdaX) %>% nest(t02 = -state)
t03 <- map_df(StatesDot, top5.Cda.table, data = dataCdaM) %>% nest(t03 = -state)
t04 <- map_df(StatesDot, top5.US.table, data = dataStateX, var = SITC) %>% nest(t04 = -state)
## there is no t05 table like for Countries
t06 <- map_df(StatesDot, prov.table, data = dataProv) %>% nest(t06 = -state)
t07 <- map_df(States, gen.table, data = dataGeneral) %>% nest(t07 = -state)
t08 <- map_df(States, high.tech.table, data = dataHighTech) %>% nest(t08 = -state)
## there is no t09 table like for Countries
t10 <- map_df(States, travel.table, data = dataTravel) %>% nest(t10 = -state)
t11 <- map_df(States, time.trade, BCX = TimeBCX, CX = TimeCX, CTX = TimeCTX, CM = TimeCM) %>% nest(t11 = -state)
## there is no t12 table like for Countries
t13 <- map_df(States, time.gdp, data = dataGDP) %>% nest(t13 = -state)
## there is no t14 table after all
t15 <- map_df(States, rank.bind, data = dataRankState, dataAsCountry = dataRank) %>% nest(t15 = -state)


## 2. merge into mega-data
nested_data_states <- t01 %>%
  ## take t01 and add in t02 through t06 (built with StatesDot)
  left_join(t02, by = "state") %>%
  left_join(t03, by = "state") %>%
  left_join(t04, by = "state") %>%
  ## there is no t05 table like for Countries
  left_join(t06, by = "state") %>%
  ## replace dots with spaces in state names
  mutate(state = str_replace_all(state, pattern = "\"", replacement = ""),
         state = str_replace_all(state, pattern = "[.]", replacement = " "),
         state = str_replace_all(state, pattern = "  ", replacement = ". ")) %>%
  ## add in remaining tables (built with States)
  left_join(t07, by = "state") %>%
  left_join(t08, by = "state") %>%
  ## there is no t09 table like for Countries
  left_join(t10, by = "state") %>%
  left_join(t11, by = "state") %>%
  ## there is no t12 table like for Countries
  left_join(t13, by = "state") %>%
  ## there is no t14 for States (or Countries)
  left_join(t15, by = "state")

saveRDS(nested_data_states, "data/nested_data_states.rds")

