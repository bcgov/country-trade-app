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
## Date: June/Sept/Oct 2020; Feb 2021; May 2024
## Coder: Julie Hawkins (and Stephanie Yurchak for app)
## KUDOS: Nasim Taba trouble-shooting with getting it to run on shinyapps.io
## ** To re-publish app: global.R, server.R, ui.R, .Rmds, columns.tex, and www and data folders (rds files)
##            Click on blue eye icon to Publish.
##            You may need to login to the BC Stats account through "Manage Accounts".
##            The App ID is 2402517, and the name is "CountryTradeApp".
## ** Private test version at: ID 3645251, name "CountryTradeTest".


## set value ----

## set "Last updated" date, for ui.R
last_update <- "February 18, 2026" ## "February 17, 2021"
# last_update <- format(Sys.time(), "%B %d %Y")  ## gives Month Date Year (e.g., November 12 2020)


## load packages ----
library(shiny) 
library(lubridate) 
library(DT) 
library(htmlwidgets) 
library(shinythemes) 
library(shinydashboard) 
library(plotly) 
library(shinyjs) 
library(shinycssloaders)
library(markdown)
library(tidyr)
library(tidyverse)
library(openxlsx)
library(janitor)
library(kableExtra)
options(scipen = 999)


## load data ----

### * read in tables, nested by country
nested_data_countries <- readRDS("data/nested_data_countries.rds")

## get list of countries (to call in drop down list)
countries <- unique(nested_data_countries$Country) %>% sort()

## years for data titles and filtering
years <- readRDS("data/years.rds")

## set year for titles in ui.R
year <- years$year[years$tab == "yearSC"]

## get names of grouped countries (to add special note on page 1)
grpdCountries <- readRDS("data/grpdCountries.rds") %>% #pull() %>%
  str_replace(pattern = "EU & UK", replacement = "EU + UK")

### * read in tables, nested by state
nested_data_states <- readRDS("data/nested_data_states.rds")

## get list of states (to call in drop down list)
states <- unique(nested_data_states$state) %>% sort()

## years for data titles and filtering
yearsStates <- readRDS("data/yearsStates.rds")


### * read in main page tables
## Aggregate Data for B.C. & Canada main page tables

## Top five B.C. origin exports to world
table1 <- readRDS("data/table1.rds") %>% select(-Country)

## Top five Canadian exports to world
table2 <- readRDS("data/table2.rds") %>% select(-Country)

## Top five Canadian imports from world
table3 <- readRDS("data/table3.rds") %>% select(-Country)

## Provincial distribution of exports to world
table4 <- readRDS("data/table4.rds") %>% select(-Country)
## get row number of BC row, so we can bold it in ui.R DT::datatable
BCrow <- which(str_detect(table4$`Province of Origin`, "British Columbia"))


## table functions ----

## kable_styling() requires bootstrap_options if format=html (instead of latex_options, only for pdf)
format.for.html <- function(tab, alignVector, boldRow) {
  tab %>% 
    kbl(align = {{alignVector}}) %>%
    kable_styling(full_width = FALSE, position = "left", bootstrap_options = c("striped", "condensed"))%>%
    column_spec(2:3, width = "5em") %>%     ## an "em" ~ the size of the M character
    row_spec(0, color = "#2A64AB") %>%
    row_spec(boldRow, bold = TRUE)
}

format.for.html.footnote <- function(tab, alignVector, footNote) {
  tab %>% 
    kbl(align = {{alignVector}}) %>%
    kable_styling(full_width = FALSE, position = "left", bootstrap_options = c("striped", "condensed"))%>%
    footnote(general = {{footNote}}, general_title = "")
}

## kable_styling() requires latex_options if format=latex (instead of bootstrap_options, only for html)
format.for.pdf <- function(tab, alignVector, boldRow) {
  tab %>% 
    kbl(align = {{alignVector}}, booktabs = TRUE, linesep = "") %>%
    kable_styling(full_width = FALSE, position = "left", font_size = 6.4, latex_options = "striped") %>%
    column_spec(1, width = "25em") %>%
    column_spec(2:3, width = "4em") %>%
    # column_spec(1, width = "28em") %>%   ## if font change
    # column_spec(2:3, width = "6em") %>%  ## if font change
    row_spec(0, color = "#2A64AB") %>%
    row_spec(boldRow, bold = TRUE)
}

format.for.pdf.footnote <- function(tab, alignVector, footNote) {
  tab %>% 
    kbl(align = {{alignVector}}, booktabs = TRUE, linesep = "") %>%
    kable_styling(full_width = FALSE, position = "left", font_size = 6.4, latex_options = "striped") %>%
    footnote(general = {{footNote}}, general_title = "")
}


## chart functions ----

## c1: BC Origin Exports to {{country}}:
##       blue bars = $Millions, t11 `BC Exports`
##       line = share, t11 `BC Exports` / t11 `Canada Exports`
c1 <- function(tab, country, nested = TRUE) {
  
  ## if tab is nested, unnest it
  if(nested == TRUE) {
    tab <- tab %>% unnest(cols = c(t11))
  }
  
  tab <- tab %>% 
    ## get just {{country}}'s rows
    filter(Country == {{country}}) %>%
    ## get just specified years
    filter(Year >= max(as.numeric(Year))-9)

  ## set ylim maximum to next rounded UP value past max, set sequence in breaks
  ## https://stackoverflow.com/questions/42486079/ggplot2-how-to-end-y-axis-on-a-tick-mark
  #ylim_max_rounded <- plyr::round_any(ggplot_build(p)$layout$panel_params[[1]][["y.range"]][[2]], 100000000, f = ceiling)
  ymax_rnd <- plyr::round_any(max(tab$`BC Exports`), 100000000, f = ceiling)  ## rounded max of y-axis
  break_seq <- 100000000                                                          ## set breaks sequence (in millions)
  if(max(tab$`BC Exports`) > 1000000000) {
    ## for countries with Exports in 1,000 Millions, want larger rounded accuracy and fewer breaks
    ymax_rnd <- plyr::round_any(max(tab$`BC Exports`), accuracy = 1000000000, f = ceiling)
    break_seq <- 1000000000
  }
  if(max(tab$`BC Exports`) > 10000000000) {
    ## for countries with Exports in 10,000 Millions, want larger rounded accuracy and fewer breaks
    ymax_rnd <- plyr::round_any(max(tab$`BC Exports`), accuracy = 5000000000, f = ceiling)
    break_seq <- 5000000000
  }
  if(max(tab$`BC Exports`) < 250000000) {
    ## for countries with Exports < 250 Million, want more frequent breaks
    break_seq <- 50000000
  }
  if(max(tab$`BC Exports`) < 50000000) {
    ## for countries with Exports < 50 Million, want smaller rounded accuracy and more frequent breaks
    ymax_rnd <- plyr::round_any(max(tab$`BC Exports`), accuracy = 5000000, f = ceiling)
    break_seq <- 5000000
  }
  if(max(tab$`BC Exports`) < 10000000) {
    ## for countries with Exports < 10 Million, want smaller rounded accuracy and many more frequent breaks
    ymax_rnd <- plyr::round_any(max(tab$`BC Exports`), accuracy = 1000000, f = ceiling)
    break_seq <- 500000
  }
  if(ymax_rnd %% break_seq != 0) {
    ## if ymax_rnd is not a multiple of break_seq, make next break_seq the max instead
    ymax_rnd <- break_seq*(ceiling(ymax_rnd / break_seq))
  }
  ysec_max_rnd <- plyr::round_any(max(tab$`BC %Canada`), accuracy = 5, f = ceiling)  ## rounded max of secondary y-axis
  ysec_trans <- ymax_rnd / ysec_max_rnd                           ## scaling value needed for transformation
  if(ysec_max_rnd %% 5 != 0) {
    ## if ysec_max_rnd is not a multiple of secondary axis break seq, make next sequence the max instead
    ysec_max_rnd <- 5*(ceiling(ysec_max_rnd / 5))
  }
  breaks_seq_sec <- 5
  if(ysec_max_rnd <= 10) {
    ## if max `BC %Canada` is less than 10%, sequence breaks by 1 instead of 5
    breaks_seq_sec <- 1
  }

  tab %>%
    ggplot2::ggplot(mapping = aes(x = Year)) +

    ### * add bars of $$, need fill here for legend, black outline
    geom_col(aes(y = `BC Exports`, fill = "BC Exports"), color = "black") +   ## stat = "identity",  if use geom_bar
    ## change fill to #17A1E3 (RGB 23 161 227); if added, name = "" removes legend title
    scale_fill_manual(values = c("#17A1E3")) +

    ### * add line of %; multiply by ysec_trans to scale up, color here for legend
    geom_line(aes(y = `BC %Canada` * ysec_trans, group = 1, color = "BC Share of Canada Exports"), size = 1) +
    ## change color to #234275 (RGB 35 66 117); if added, name = "" removes legend title
    scale_color_manual(values = c("#234275")) +
    ## format y-axis
    scale_y_continuous(limits = c(0, ymax_rnd),
                       breaks = seq(0, ymax_rnd, break_seq),
                       expand = c(0,0),  ## "tight layout": no space between axis and plot; https://stackoverflow.com/questions/22945651/how-to-remove-space-between-axis-area-plot-in-ggplot2
                       labels = function(x)x / 1000000,  ## display labels in millions; https://stackoverflow.com/questions/4646020/ggplot2-axis-transformation-by-constant-factor
                       ## add secondary axis, need to transform it by scaling down accordingly
                       sec.axis = sec_axis(trans = ~ . / ysec_trans,
                                           #labels = scales::unit_format(unit = "%", sep = ""),
                                           breaks = seq(0, ysec_max_rnd, breaks_seq_sec),
                                           name = "Percent")) +
    ### * format
    labs(#title = paste0("BC Origin Exports to ", country),
         ## 2-part caption requires this code to work: `theme(plot.caption = element_text(hjust = c(1, 0))`
         caption = c("BC Stats", expression('Note: Exports'~bold('exclude')~'re-exports')),
         #subtitle = expression('Note: Exports'~bold('exclude')~'re-exports'),  ## placed as caption now
         #caption = "BC Stats",
         fill = "",  color = "",            ## removes legend title(s), https://stackoverflow.com/questions/14622421/how-to-change-legend-title-in-ggplot
         x = "", y = "$Millions") +
    theme_classic() +                       ## x and y axis, but no gridlines
    theme(plot.title = element_text(color = "blue", face = "bold"),
          #plot.subtitle = element_text(size = rel(0.8)),
          plot.caption = element_text(hjust = c(1, 0)),  ## needed for 2-part caption ** this MIGHT cause problems later**
          legend.position = c("top"),       ## places legend OUTSIDE plot
          #legend.position = c(0.10, 0.95),          ## places legend INSIDE plot (place at 0:1,0:1)
          #legend.justification = c("left", "top"),  ## if legend INSIDE plot
          #legend.box = "horizontal",                ## if legend INSIDE plot
          ## moves "$Millions" axis title to top and centered over axis; smaller font size
          # axis.title.y.left = element_text(size = rel(0.8), angle = 0, vjust = 1.10, hjust = 450000000),
          axis.title.y.left = element_text(size = rel(0.8), angle = 0, vjust = 1.10, margin = margin(t=0, r=0, b=0, l=0)),
          ## moves "Percent" axis title to top; margin piece centers over axis; smaller font size; https://github.com/tidyverse/ggplot2/issues/1435
          axis.title.y.right = element_text(size = rel(0.8), angle = 0, vjust = 1.10, margin = margin(t=0, r=0, b=0, l=-32)),
    )
}
#c1(tab = t11, country = "Australia", nested = TRUE)


## c2: Canada's Balance of Trade in Goods with {{country}}:
##       blue bars = Exports, t11 `Canada Exports`
##       gray bars = Imports, t11 `Canada Imports`
##       line = Balance of Trade, t11 `Trade Balance`
c2 <- function(tab, country, nested = TRUE) {
  
  ## if tab is nested, unnest it
  if(nested == TRUE) {
    tab <- tab %>% unnest(cols = c(t11))
  }
  tab <- tab %>% 
    ## get just {{country}}'s rows
    filter(Country == {{country}}) %>%
    ## get just specified years
    filter(Year >= max(as.numeric(Year))-9)
  ## get line data
  tab_line <- tab %>% select(Year, `Trade Balance`)
  ## get bar data, gather long
  tab <- tab %>%
    select(Year, Exports = `Canada Total Exports`, Imports = `Canada Imports`) %>%
    ## gather long
    pivot_longer(-c("Year"), names_to = "Var", values_to = "value")
  
  ## set ylim maximum to next rounded UP value past max, ylim min to rounded past lowest Trade Balance
  max_value <- max(tab$value)
  ymax_rnd <- plyr::round_any(max_value, 100000000, f = ceiling)
  break_seq <- 500000000
  if(max_value > 5000000000) {
    ## for countries with max_value > 5 Billion, want smaller rounded accuracy and more frequent breaks
    ymax_rnd <- plyr::round_any(max_value, 5000000000, f = ceiling)
    break_seq <- 5000000000
  }
  if(max_value > 10000000000) {
    ## for countries with max_value in 10+ Billions, want larger rounded accuracy and fewer breaks
    ymax_rnd <- plyr::round_any(max_value, accuracy = 10000000000, f = ceiling)
    break_seq <- 10000000000
  }
  if(max_value > 100000000000) {
    ## for countries with max_value in 100+ Billions, want larger rounded accuracy and fewer breaks
    ymax_rnd <- plyr::round_any(max_value, accuracy = 100000000000, f = ceiling)
    break_seq <- 100000000000
  }
  if(max_value < 500000000) {
    ## for countries with max_value < 0.5 Billion, want smaller rounded accuracy and more frequent breaks
    ymax_rnd <- plyr::round_any(max_value, 500000000, f = ceiling)
    break_seq <- 100000000
  }
  if(max_value < 100000000) {
    ## for countries with max_value < 0.1 Billion, want smaller rounded accuracy and more frequent breaks
    ymax_rnd <- plyr::round_any(max_value, 100000000, f = ceiling)
    break_seq <- 100000000
  }
  if(max_value < 50000000) {
    ## for countries with max_value < 0.05 Billion, want smaller rounded accuracy and more frequent breaks
    ymax_rnd <- plyr::round_any(max_value, 10000000, f = ceiling)
    break_seq <- 10000000
  }
  if(ymax_rnd %% break_seq != 0) {
    ## if ymax_rnd is not a multiple of break_seq, make next break_seq the max instead
    ymax_rnd <- break_seq*(ceiling(ymax_rnd / break_seq))
  }
  
  ymin_rnd <- plyr::round_any(min(tab_line$`Trade Balance`), 100000000, f = floor)
  if(min(tab_line$`Trade Balance`) > -50000000) {
    ## for countries with much smaller min value, want smaller rounded accuracy
    ymin_rnd <- plyr::round_any(min(tab_line$`Trade Balance`), 10000000, f = floor)
  }
  if(ymin_rnd >= 0) {
    ## set ylim minimum to 0 (all numbers are positive)
    ylim_min_rnd <- 0
  } else if(ymin_rnd > (0-break_seq)) {
    ## minimum `Trade Balance` between 0 and one break_seq away, so set ylim minimum to one break_seq away (e.g., -0.5B)
    ylim_min_rnd <- (0-break_seq)
  } else {
    ## minimum `Trade Balance` more than one break_seq less than zero, find how many times more, round up, multiply by break_seq
    #ylim_min_rnd <- (-1)*(break_seq*(janitor::round_half_up(ymin_rnd / break_seq*(-1))))
    ylim_min_rnd <- break_seq*(floor(ymin_rnd / break_seq))
  }
  
  tab %>%
    ggplot2::ggplot(mapping = aes(x = Year)) +
    ### * add bars (Exports & Imports), need fill here for legend, black outline; http://www.sthda.com/english/wiki/ggplot2-barplots-quick-start-guide-r-software-and-data-visualization
    geom_col(aes(y = value, fill = Var), color = "black", position = position_dodge()) +
    ## change fill to #17A1E3 (RGB 23 161 227) and #DDE2ED (RGB 221 226 237)
    scale_fill_manual(values = c("#17A1E3", "#DDE2ED"), name = "") +
    
    ### * add line of %, need color here for legend
    geom_line(data = tab_line, aes(x = Year, y = `Trade Balance`, group = 1, color = "Balance of Trade"), size = 1) +
    ## change color to #234275 (RGB 35 66 117)
    scale_color_manual(values = c("#234275"), name = "") +
    
    ### * format y-axis
    scale_y_continuous(limits = c(ylim_min_rnd, ymax_rnd),
                       breaks = c(seq(ylim_min_rnd, 0, break_seq), seq(0, ymax_rnd, break_seq)),
                       expand = c(0,0),                  ## "tight layout": no space between axis and plot
                       labels = function(x)x / 1000000000,  ## display labels in billions
    ) +
    ### * format
    labs(#title = paste0("Canada's Balance of Trade in Goods with ", country),
         ## 2-part caption requires this code to work: `theme(plot.caption = element_text(hjust = c(1, 0))`
         caption = c("BC Stats", expression('Note: Exports'~bold('exclude')~'re-exports')),
         #subtitle = expression('Note: Exports'~bold('exclude')~'re-exports'),  ## placed as caption now
         #caption = "BC Stats",
         x = "", y = "$Billions") +
    theme_classic() +                       ## x and y axis, but no gridlines
    theme(plot.title = element_text(color = "blue", face = "bold"),
          plot.caption = element_text(hjust = c(1, 0)),  ## needed for 2-part caption
          #plot.subtitle = element_text(size = rel(0.8)),
          legend.position = c("top"),       ## places legend OUTSIDE plot
          ## moves y-axis title to top of axis; smaller font size
          axis.title.y = element_text(size = rel(0.8), angle = 0, vjust = 1.10)  ## hjust = 0, margin = margin(t=0, l=0, r=0, b=0)
    )
}
#c2(tab = t11, country = "Australia", nested = TRUE)


## c3: Canada's Balance of Trade in Services with {{country}}:
##       blue bars = Exports, t12 `Service Exports` (in $M)
##       gray bars = Imports, t12 `Service Imports` (in $M)
##       line = Balance of Trade, t12 `Service Trade Balance`
## ** NOTE: c3 should NOT appear when data is unavailable (i.e., when Index$X8 == "XX") **
c3 <- function(tab, country, nested = TRUE) {
  
  ## if tab is nested, unnest it
  if(nested == TRUE) {
    tab <- tab %>% unnest(cols = c(t12))
  }
  tab <- tab %>% 
    ## get just {{country}}'s rows
    filter(Country == {{country}}) %>%
    ## get just specified years
    filter(Year >= max(as.numeric(Year))-9)
  
  ## if tab is empty, i.e., no Service results for that country, skip this chart
  if(dim(tab)[1] > 0) {
    ## get line data
    tab_line <- tab %>% select(Year, `Trade Balance` = `Service Trade Balance`)
    ## get bar data, gather long
    tab <- tab %>%
      select(Year, Exports = `Service Exports`, Imports = `Service Imports`) %>%
      ## gather long
      pivot_longer(-c("Year"), names_to = "Var", values_to = "value")
    
    ## set ylim maximum to next rounded UP value past max, ylim min to rounded past lowest Trade Balance
    max_value <- max(tab$value)
    ymax_rnd <- plyr::round_any(max_value, 200, f = ceiling)
    break_seq <- 200
    if(max_value > 2000) {
      ## for countries with max_value > 2 Billion, want larger rounded accuracy and breaks
      ymax_rnd <- plyr::round_any(max_value, accuracy = 500, f = ceiling)
      break_seq <- 500
    }
    if(max_value > 5000) {
      ## for countries with max_value > 5 Billion, want larger rounded accuracy and breaks
      ymax_rnd <- plyr::round_any(max_value, accuracy = 1000, f = ceiling)
      break_seq <- 1000
    }
    if(max_value > 10000) {
      ## for countries with max_value in 10+ Billions, want much larger rounded accuracy and breaks
      ymax_rnd <- plyr::round_any(max_value, accuracy = 10000, f = ceiling)
      break_seq <- 10000
    }
    if(max_value < 400) {
      ## for countries with max_value < 0.400 Billion, want smaller rounded accuracy and more frequent breaks
      ymax_rnd <- plyr::round_any(max_value, 100, f = ceiling)
      break_seq <- 50
    }
    if(ymax_rnd %% break_seq != 0) {
      ## if ymax_rnd is not a multiple of break_seq, make next break_seq the max instead
      ymax_rnd <- break_seq*(ceiling(ymax_rnd / break_seq))
    }
    
    ymin_rnd <- plyr::round_any(min(tab_line$`Trade Balance`), 100, f = floor)
    if(ymin_rnd >= 0) {
      ## set ylim minimum to 0 (all numbers are positive)
      ylim_min_rnd <- 0
    } else if(ymin_rnd > (0-break_seq)) {
      ## minimum `Trade Balance` between 0 and one break_seq away, so set ylim minimum to one break_seq away (e.g., -0.5B)
      ylim_min_rnd <- (0-break_seq)
    } else {
      ## minimum `Trade Balance` more than one break_seq less than zero, find how many times more, round up, multiply by break_seq
      ylim_min_rnd <- break_seq*(floor(ymin_rnd / break_seq))
    }
    
    p <- tab %>%
      ggplot2::ggplot(mapping = aes(x = Year)) +
      ### * add bars (Exports & Imports), need fill here for legend, black outline; http://www.sthda.com/english/wiki/ggplot2-barplots-quick-start-guide-r-software-and-data-visualization
      geom_col(aes(y = value, fill = Var), color = "black", position = position_dodge()) +
      ## change fill to #17A1E3 (RGB 23 161 227) and #DDE2ED (RGB 221 226 237)
      scale_fill_manual(values = c("#17A1E3", "#DDE2ED"), name = "") +
      
      ### * add line of %, need color here for legend
      geom_line(data = tab_line, aes(x = Year, y = `Trade Balance`, group = 1, color = "Balance of Trade"), size = 1) +
      ## change color to #234275 (RGB 35 66 117)
      scale_color_manual(values = c("#234275"), name = "") +
      
      ### * format y-axis
      scale_y_continuous(limits = c(ylim_min_rnd, ymax_rnd),
                         breaks = c(seq(ylim_min_rnd, 0, break_seq), seq(0, ymax_rnd, break_seq)),
                         expand = c(0,0),                  ## "tight layout": no space between axis and plot
                         labels = function(x)x / 1000,     ## display labels in billions (already down to millions)
      ) +
      ### * format
      labs(#title = paste0("Canada's Balance of Trade in Services with ", country),
           caption = "BC Stats",
           x = "", y = "$Billions") +
      theme_classic() +                       ## x and y axis, but no gridlines
      theme(plot.title = element_text(color = "blue", face = "bold"),
            plot.subtitle = element_text(size = rel(0.8)),
            legend.position = c("top"),       ## places legend OUTSIDE plot
            axis.title.y = element_text(size = rel(0.8), angle = 0, vjust = 1.10)  ## moves y-axis title to top of axis; smaller font size
      )
    p
  } else {
    
  }
}
#c3(tab = t12, country = "Australia", nested = TRUE)


## c4: Year-over-Year % Growth in GDP of {{country}}:
##       line = % change, t13 value
c4 <- function(tab, country, nested = TRUE, source_text) {
  
  ## if tab is nested, unnest it
  if(nested == TRUE) {
    tab <- tab %>% unnest(cols = c(t13))
  }
  tab <- tab %>% 
    ## get just {{country}}'s rows
    filter(Country == {{country}}) %>%
    ## get just specified years
    filter(Year >= max(as.numeric(Year))-9)
  
  ## set y maximum to next rounded UP value past max, y minimum to rounded past lowest value
  max_value <- max(tab$value)
  ymax_rnd <- plyr::round_any(max_value, 1, f = ceiling)
  break_seq <- 0.5
  if(ymax_rnd >= 5) {
    ## if rounded y max is 5+, break at every percent
    break_seq <- 1
  }
  if(ymax_rnd >= 10) {
    ## if rounded y max is 10+, break at every second percent (and make sure ymax is multiply of 2)
    break_seq <- 2
    if(ymax_rnd %% 2 != 0) {
      ymax_rnd <- ymax_rnd + 1
    }
  }
  
  min_value <- min(tab$value)
  ymin_rnd <- plyr::round_any(min_value, 1, f = floor)
  if(ymin_rnd >= 0) {
    ## set y minimum to 0 when all values are positive
    ymin_rnd <- 0
  } else if(ymin_rnd > (0-break_seq)) {
    ## else: minimum value between 0 and one break_seq away, so set y min to one break_seq away
    ymin_rnd <- (0-break_seq)
  } else {
    ## minimum value more than one break_seq less than zero, set y min to next break_seq
    ymin_rnd <- break_seq*(floor(ymin_rnd / break_seq))
  }
  
  p <- tab %>%
    ggplot2::ggplot(mapping = aes(x = Year)) +
    ### * add line of %, color #234275 (RGB 35 66 117)
    geom_line(aes(y = value, group = 1), size = 1, color = "#234275") +
    
    ### * format y-axis
    scale_y_continuous(limits = c(ymin_rnd, ymax_rnd),
                       breaks = c(seq(ymin_rnd, 0, break_seq), seq(0, ymax_rnd, break_seq)),
                       expand = c(0,0)
    ) +
    ### * format
    labs(#title = paste0("Year-Over-Year % Growth in GDP of ", country),
         subtitle = "\n",  ## one extra break so y-axis moves up
         ## 2-part caption requires this code to work: `theme(plot.caption = element_text(hjust = c(1, 0))`
         caption = c("BC Stats", {{source_text}}),
         # caption = "BC Stats",
         x = "", y = "% Change") +
    theme_classic() +                       ## x and y axis, but no gridlines
    theme(plot.title = element_text(color = "blue", face = "bold"),
          plot.caption = element_text(hjust = c(1, 0)),  ## needed for 2-part caption
          #plot.subtitle = element_text(size = rel(0.8)),
          legend.position = c("top"),       ## places legend OUTSIDE plot
          ## moves y-axis title to top of axis; smaller font size
          axis.title.y = element_text(size = rel(0.8), angle = 0, vjust = 1.1)
    )
  
  if(ymin_rnd < 0) {
    ## if ymin is less than zero, add light grey dashed line at y=0
    p <- p + geom_hline(yintercept = 0, linetype = 2, color = "grey")
  }
  
  p
  
}
#c4(tab = t13, country = "Australia", nested = TRUE, source_text = "")


## c5: Immigrants to BC from {{country}}:
##       line = Persons, t10 `Immigrants (Persons)`
c5 <- function(tab, country, nested = TRUE) {
  
  ## if tab is nested, unnest it
  if(nested == TRUE) {
    tab <- tab %>% unnest(cols = c(t10))
  }
  tab <- tab %>% 
    ## get just {{country}}'s rows
    filter(Country == {{country}}) %>%
    ## get just specified years
    filter(Year >= max(as.numeric(Year))-9) %>%
    rename(Persons = `Immigrants (Persons)`)
  
  ## set y maximum to next rounded UP Persons past max
  max_value <- max(tab$Persons, na.rm = TRUE)
  #if(max_value == -Inf) { max_value <- 0 }
  ymax_rnd <- plyr::round_any(max_value, 100, f = ceiling)
  break_seq <- 100
  if(ymax_rnd >= 1000) {
    ## if rounded y max is 1000+, break at 200
    break_seq <- 200
    ymax_rnd <- plyr::round_any(max_value, 200, f = ceiling)
  }
  if(ymax_rnd >= 2000) {
    ## if rounded y max is 2000+, break at 500
    break_seq <- 500
    ymax_rnd <- plyr::round_any(max_value, 500, f = ceiling)
  }
  if(ymax_rnd >= 5000) {
    ## if rounded y max is 5000+, break at 1000
    break_seq <- 1000
    ymax_rnd <- plyr::round_any(max_value, 1000, f = ceiling)
  }
  if(ymax_rnd <= 300) {
    ## if rounded y max is 300 or less, break by 50s, and reduce accuracy
    break_seq <- 50
    ymax_rnd <- plyr::round_any(max_value, 50, f = ceiling)
  }
  if(ymax_rnd <= 100) {
    ## if rounded y max is 100 or less, break by 10s, and reduce accuracy
    break_seq <- 10
    ymax_rnd <- plyr::round_any(max_value, 10, f = ceiling)
  }
  if(ymax_rnd <= 50) {
    ## if rounded y max is 100 or less, break by 5s, and reduce accuracy
    break_seq <- 5
    ymax_rnd <- plyr::round_any(max_value, 10, f = ceiling)
  }
  
  tab %>%
    ## drop NA values so don't get warning
    filter(!is.na(Persons)) %>%
    ggplot2::ggplot(mapping = aes(x = Year)) +
    ### * add line of %, color #234275 (RGB 35 66 117)
    geom_line(aes(y = Persons, group = 1), size = 1, color = "#234275") +
    
    ### * format y-axis
    scale_y_continuous(limits = c(0, ymax_rnd),
                       breaks = seq(0, ymax_rnd, break_seq),
                       expand = c(0,0)
    ) +
    ### * format
    labs(#title = paste0("Immigrants to BC from ", country),
         subtitle = "\n",  ## one extra break so y-axis moves up
         ## 2-part caption requires this code to work: `theme(plot.caption = element_text(hjust = c(1, 0))`
         caption = c("BC Stats", "Source: Immigration, Refugees and Citizenship Canada"),
         #subtitle = "Source: Immigration, Refugees and Citizenship Canada \n\n",  ## two extra breaks so y-axis moves up
         #caption = "BC Stats",
         x = "") +
    theme_classic() +                       ## x and y axis, but no gridlines
    theme(plot.title = element_text(color = "blue", face = "bold"),
          plot.caption = element_text(hjust = c(1, 0)),  ## needed for 2-part caption
          #plot.subtitle = element_text(size = rel(0.8)),
          legend.position = c("top"),       ## places legend OUTSIDE plot
          ## moves y-axis title to top of axis; smaller font size
          axis.title.y = element_text(size = rel(0.8), angle = 0, vjust = 1.1)
    )
  
}
#c5(tab = t10, country = "Australia", nested = TRUE)


## c6: Travellers from {{country}} Entering Through BC:
##       line = Thousands, t10 `Travellers (Persons)`
c6 <- function(tab, country, nested = TRUE) {
  
  ## if tab is nested, unnest it
  if(nested == TRUE) {
    tab <- tab %>% unnest(cols = c(t10))
  }
  tab <- tab %>% 
    ## get just {{country}}'s rows
    filter(Country == {{country}}) %>%
    ## get just specified years
    filter(Year >= max(as.numeric(Year))-9) %>%
    rename(Persons = `Travellers (Persons)`)
  
  ## set y maximum to next rounded UP Persons past max
  max_value <- max(tab$Persons, na.rm = TRUE)
  ymax_rnd <- plyr::round_any(max_value, 50000, f = ceiling)
  break_seq <- 50000
  if(ymax_rnd >= 1000000) {
    ## if rounded y max is 1M+, break at 1000000
    break_seq <- 1000000
    ymax_rnd <- plyr::round_any(max_value, 1000000, f = ceiling)
  }
  if(max_value < 50000) {
    ## if rounded y max is < 50,000, break by 10000s, and reduce accuracy
    break_seq <- 5000
    ymax_rnd <- plyr::round_any(max_value, 10000, f = ceiling)
  }
  if(max_value < 10000) {
    ## if rounded y max is < 10,000, break by 1000s, and reduce accuracy
    break_seq <- 1000
    ymax_rnd <- plyr::round_any(max_value, 1000, f = ceiling)
  }
  if(max_value < 1000) {
    ## if rounded y max is < 1000, break by 100s, and reduce accuracy
    break_seq <- 100
    ymax_rnd <- plyr::round_any(max_value, 100, f = ceiling)
  }
  
  if(max_value < 1000) {
    ## for countries with few travellers (< 1,000) to BC
    tab %>%
      ggplot2::ggplot(mapping = aes(x = Year)) +
      ### * add line of %, color #234275 (RGB 35 66 117)
      geom_line(aes(y = Persons, group = 1), size = 1, color = "#234275") +
      
      ### * format y-axis
      scale_y_continuous(limits = c(0, ymax_rnd),
                         breaks = seq(0, ymax_rnd, break_seq),
                         expand = c(0,0)
      ) +
      ### * format
      labs(#title = paste0("Travellers from ", country, " Entering Canada Through BC"),
           subtitle = "\n",  ## one extra break so y-axis moves up
           ## 2-part caption requires this code to work: `theme(plot.caption = element_text(hjust = c(1, 0))`
           caption = c("BC Stats", "Source: Statistics Canada"),
           #subtitle = "Source: Statistics Canada \n\n",  ## two extra breaks so y-axis moves up  ## placed as caption now
           #caption = "BC Stats",
           x = "", y = "") +
      theme_classic() +                       ## x and y axis, but no gridlines
      theme(plot.title = element_text(color = "blue", face = "bold"),
            plot.caption = element_text(hjust = c(1, 0)),  ## needed for 2-part caption
            #plot.subtitle = element_text(size = rel(0.8)),
            legend.position = c("top"),       ## places legend OUTSIDE plot
            ## moves y-axis title to top of axis; smaller font size
            axis.title.y = element_text(size = rel(0.8), angle = 0, vjust = 1.1)
      )
  } else {
    tab %>%
      ggplot2::ggplot(mapping = aes(x = Year)) +
      ### * add line of %, color #234275 (RGB 35 66 117)
      geom_line(aes(y = Persons / 1000, group = 1), size = 1, color = "#234275") +
      
      ### * format y-axis
      scale_y_continuous(limits = c(0, ymax_rnd / 1000),
                         breaks = seq(0, ymax_rnd / 1000, break_seq / 1000),
                         expand = c(0,0)
      ) +
      ### * format
      labs(#title = paste0("Travellers from ", country, " Entering Canada Through BC"),
           subtitle = "\n",  ## one extra break so y-axis moves up
           ## 2-part caption requires this code to work: `theme(plot.caption = element_text(hjust = c(1, 0))`
           caption = c("BC Stats", "Source: Statistics Canada"),
           #subtitle = "Source: Statistics Canada \n\n",  ## two extra breaks so y-axis moves up  ## placed as caption now
           #caption = "BC Stats",
           x = "", y = "Thousands") +
      theme_classic() +                       ## x and y axis, but no gridlines
      theme(plot.title = element_text(color = "blue", face = "bold"),
            plot.caption = element_text(hjust = c(1, 0)),  ## needed for 2-part caption
            #plot.subtitle = element_text(size = rel(0.8)),
            legend.position = c("top"),       ## places legend OUTSIDE plot
            ## moves y-axis title to top of axis; smaller font size
            axis.title.y = element_text(size = rel(0.8), angle = 0, vjust = 1.1)
      )
  }
}
#c6(tab = t10, country = "Australia", nested = TRUE)


## c7: BC's Balance of Trade in High Tech Goods with {{state}}:
##       *** this chart is only for STATES ***
##       blue bars = Exports, t08 `Domestic Exports` + t08 `Re-Exports` (most re-exports are 0, but not all)
##       gray bars = Imports, t08 `Imports`
##       line = Balance of Trade, t08 `Trade Balance`
c7 <- function(tab, state, nested = TRUE) {
  
  ## if tab is nested, unnest it
  if(nested == TRUE) {
    tab <- tab %>% unnest(cols = c(t08))
  }
  tab <- tab %>% 
    ## get just {{state}}'s rows
    filter(state == {{state}}) %>%
    ## get just specified years
    filter(Year >= max(as.numeric(Year))-9)
  ## get line data
  tab_line <- tab %>% select(Year, `Trade Balance`)
  ## get bar data, gather long
  tab <- tab %>%
    mutate(Exports = `Domestic Exports` + `Re-Exports`) %>%
    select(Year, Exports, Imports) %>%
    ## gather long
    pivot_longer(-c("Year"), names_to = "Var", values_to = "value")
  
  ## set ylim maximum to next rounded UP value past max, ylim min to rounded past lowest Trade Balance
  max_value <- max(tab$value)
  ymax_rnd <- plyr::round_any(max_value, 5, f = ceiling)
  break_seq <- 5            ## e.g., Iowa, Oregon
  if(max_value > 30) {      ## e.g., kentucky
    ## for countries with max_value > 30 Million, want larger rounded accuracy and fewer breaks
    ymax_rnd <- plyr::round_any(max_value, accuracy = 10, f = ceiling)
    break_seq <- 10
  }
  if(max_value > 100) {      ## e.g., Illinois
    ## for countries with max_value > 100 Million, want larger rounded accuracy and fewer breaks
    ymax_rnd <- plyr::round_any(max_value, accuracy = 20, f = ceiling)
    break_seq <- 20
  }
  if(max_value > 150) {      ## e.g., California, Texas
    ## for countries with max_value > 150 Million, want larger rounded accuracy and fewer breaks
    ymax_rnd <- plyr::round_any(max_value, accuracy = 50, f = ceiling)
    break_seq <- 50
  }
  if(max_value > 400) {      ## e.g., Washington
    ## for countries with max_value > 150 Million, want larger rounded accuracy and fewer breaks
    ymax_rnd <- plyr::round_any(max_value, accuracy = 200, f = ceiling)
    break_seq <- 200
  }
  if(max_value < 3) {        ## e.g., Wyoming
    ## for countries with max_value < 3 Million, want smaller rounded accuracy and more frequent breaks
    # ymax_rnd <- plyr::round_any(max_value, 1, f = ceiling)
    ymax_rnd <- 3
    break_seq <- 0.5
  }
  if(ymax_rnd %% break_seq != 0) {
    ## if ymax_rnd is not a multiple of break_seq, make next break_seq the max instead
    ymax_rnd <- break_seq*(ceiling(ymax_rnd / break_seq))
  }
  
  ymin_rnd <- plyr::round_any(min(tab_line$`Trade Balance`), 5, f = floor)
  if(min(tab_line$`Trade Balance`) > -3) {
    ## for countries with smaller min value, want smaller rounded accuracy
    ymin_rnd <- plyr::round_any(min(tab_line$`Trade Balance`), 0.5, f = floor)
  }
  if(ymin_rnd >= 0) {
    ## set ylim minimum to 0 (all numbers are positive)
    ylim_min_rnd <- 0
  } else if(ymin_rnd > (0-break_seq)) {
    ## minimum `Trade Balance` between 0 and one break_seq away, so set ylim minimum to one break_seq away (e.g., -0.5B)
    ylim_min_rnd <- (0-break_seq)
  } else {
    ## minimum `Trade Balance` more than one break_seq less than zero, find how many times more, round up, multiply by break_seq
    #ylim_min_rnd <- (-1)*(break_seq*(janitor::round_half_up(ymin_rnd / break_seq*(-1))))
    ylim_min_rnd <- break_seq*(floor(ymin_rnd / break_seq))
  }
  
  tab %>%
    ggplot2::ggplot(mapping = aes(x = Year)) +
    ### * add bars (Exports & Imports), need fill here for legend, black outline; http://www.sthda.com/english/wiki/ggplot2-barplots-quick-start-guide-r-software-and-data-visualization
    geom_col(aes(y = value, fill = Var), color = "black", position = position_dodge()) +
    ## change fill to #17A1E3 (RGB 23 161 227) and #DDE2ED (RGB 221 226 237)
    scale_fill_manual(values = c("#17A1E3", "#DDE2ED"), name = "") +
    
    ### * add line of %, need color here for legend
    geom_line(data = tab_line, aes(x = Year, y = `Trade Balance`, group = 1, color = "Balance of Trade"), size = 1) +
    ## change color to #234275 (RGB 35 66 117)
    scale_color_manual(values = c("#234275"), name = "") +
    
    ### * format y-axis
    scale_y_continuous(limits = c(ylim_min_rnd, ymax_rnd),
                       breaks = c(seq(ylim_min_rnd, 0, break_seq), seq(0, ymax_rnd, break_seq)),
                       expand = c(0,0),     ## "tight layout": no space between axis and plot
                       ) +
    ### * format
    labs(## 2-part caption requires this code to work: `theme(plot.caption = element_text(hjust = c(1, 0))`
         caption = c("BC Stats", expression('Note: Exports'~bold('exclude')~'re-exports')),
         #subtitle = expression('Note: Exports'~bold('include')~'re-exports'),  ## placed as caption now
         #caption = "BC Stats",
         x = "", y = "$Millions") +
    theme_classic() +                       ## x and y axis, but no gridlines
    theme(plot.title = element_text(color = "blue", face = "bold"),
          plot.caption = element_text(hjust = c(1, 0)),  ## needed for 2-part caption
          #plot.subtitle = element_text(size = rel(0.8)),
          legend.position = c("top"),       ## places legend OUTSIDE plot
          ## moves y-axis title to top of axis; smaller font size
          axis.title.y = element_text(size = rel(0.8), angle = 0, vjust = 1.10)  ## hjust = 0, margin = margin(t=0, l=0, r=0, b=0)
    )
}
# c7(tab = t08, state = "Washington", nested = TRUE)


## text styling functions ----

## https://bookdown.org/yihui/rmarkdown-cookbook/font-color.html
colorize <- function(x, color) {
  if (knitr::is_latex_output()) {
    ## in: \\textcolor{%s}{%s}; first {} is color, second {} is text to color
    # sprintf("\\textcolor{%s}{%s}", color, x)             ## pound sign in hex color crashes PDF
    ## trick it into using color='blue' instead of a hexcode with # that breaks the LaTeX code
    # sprintf("\\textcolor{%s}{%s}", color = 'blue', x)    ## works, but isn't right blue
    sprintf("\\textcolor[HTML]{2A64AB}{%s}", x)
  } else if (knitr::is_html_output()) {
    sprintf("<span style = 'color: %s;'>%s</span>", color, x)
  } else x
}


bold.center.color.text <- function(x, color) {
  ## this function formats text as bold, centered, colored (blue) font
  if (knitr::is_latex_output()) {
    ## need double backslash to escape commands in sprintf
    ## \textbf for bold; \textcolor{}{} for color; \begin{center} \end{center} to center
    sprintf("\\textcolor[HTML]{2A64AB}{\\begin{center}{\\textbf{%s}}\\end{center}}", x)
  } else if (knitr::is_html_output()) {
    sprintf("<strong><center><span style = 'color: %s;'>%s</span></center></strong>", color, x)
  } else x
}


## xlsx settings ----

library(openxlsx)

### * source notes for countries
# source_pg1 <- c("Data Source: International Trade Centre")                        ## t04 & t05
source_pg2 <- c("* Purchasing Power Parity",                                        ## t07
                "Sources: International Monetary Fund",                             ## t07
                "Source: BC Stats",                                                 ## t08
                "Source: Statistics Canada",                                        ## t09
                "Sources:",                                                         ## t10
                "Travellers - Statistics Canada",                                   ## t10
                "Immigrants - Immigration, Refugees and Citizenship Canada")        ## t10


### * source notes for states
source_pg1_states <- c("* These are origin of movement figures and may not represent origin of",
                       "manufacture; data is by industry (NAICS) rather than commodity.",
                       "Source: US Department of Commerce")                            ## t04
source_pg2_states <- c("Source: US Census Bureau and US Bureau of Economic Analysis",  ## t07
                       "Source: BC Stats")                                             ## t08


### * set rows for Page 1
## hard-coded b/c always top 5 and province table
row_t01 <- 2       ## start at row2 because row1 = title row
row_t02 <- 10      ## title row + t01's header row, 5 commodity rows, total row (1+5+1=7)
row_t03 <- 18      ## above + blank row after t01 + t02's header row, 5 commodity rows, total row
row_t04 <- 26      ## as above
row_t05 <- 35      ## as above + data source row
row_t06 <- 44      ## as above + data source row
rows_total_p1 <- c(8, 16, 24, 32, 41, 58)        ## "Total" rows
rows_source_p1 <- c(33, 42)                      ## "Source" rows
rows_blank_p1 <- c(9, 17, 25, 34, 43, 59)        ## blank rows
rows_page1 <- max(rows_total_p1)                 ## number of rows in page1

row_t06_states <- row_t05 + 1      ## b/c states do not have a t05, and t04 has extra source row
rows_total_p1_states <- c(8, 16, 24, 32, 50)     ## "Total" rows for states
rows_source_p1_states <- c(33, 34)               ## "Source" rows for states
rows_blank_p1_states <- c(9, 17, 25, 35, 51)     ## blank rows for states
rows_page1_states <- max(rows_total_p1_states)   ## number of rows in page1 for states


### * create styles
sTitle <- createStyle(fontSize = 48, textDecoration = "bold", valign = "center")
sCountry <- createStyle(fontSize = 36, textDecoration = "bold", valign = "center", halign = "right", fontColour = "#2A64AB")  ## "#2A64AB;" #114F8B = dark blue (17-79-139)
sBorder <- createStyle(border = "Bottom", borderStyle = getOption("openxlsx.borderStyle", "slantDashDot"))
sContinued <- createStyle(fontSize = 18, textDecoration = "bold", valign = "center", border = "Bottom", borderStyle = getOption("openxlsx.borderStyle", "slantDashDot"))

sBlue <- createStyle(fontColour = "#2A64AB", textDecoration = "bold")
sText <- createStyle(wrapText = TRUE, valign = "center")
sStat <- createStyle(wrapText = TRUE, valign = "center", halign = "right")
sStat2 <- createStyle(wrapText = TRUE, valign = "center", halign = "center")
sCenter <- createStyle(valign = "center")
sBold <- createStyle(textDecoration = "bold")
sChart <- createStyle(fontColour = "#2A64AB", textDecoration = "bold", fontSize = 14, valign = "center", halign = "center")
sSource <- createStyle(valign = "center", fontSize = 8)  ## wrapText = TRUE, 
sHeader <- createStyle(wrapText = TRUE, valign = "center", halign = "center", border = "TopBottom", borderStyle = getOption("openxlsx.borderStyle", "thin"))
sNum    <- createStyle(valign = "center", halign = "center", numFmt = "0")  ## Perc: numFmt = "0%"

### * functions
add.table.styled <- function(wb, sheet, table, row, style) {
  writeData(wb, sheet = {{sheet}}, x = {{table}}, startRow = {{row}})
  addStyle(wb, sheet = {{sheet}}, style = {{style}}, rows = {{row}}, cols = 1, stack = TRUE)
}
add.table.styled.titled <- function(wb, sheet, title, table, row) {
  writeData(wb, sheet = {{sheet}}, x = {{title}}, startRow = {{row}})
  addStyle(wb, sheet = {{sheet}}, style = sBlue, rows = {{row}}, cols = 1, stack = TRUE)
  writeData(wb, sheet = {{sheet}}, x = {{table}}, startRow = {{row}}+1, headerStyle = sHeader)
}
add.chart.titled <- function(wb, sheet, c_title, chart, row, col) {
  writeData(wb, sheet = {{sheet}}, x = {{c_title}}, startRow = {{row}}, startCol = {{col}})
  addStyle(wb, sheet = {{sheet}}, style = sChart, rows = {{row}}, cols = {{col}}, stack = TRUE)
  insertImage(wb, sheet = {{sheet}}, file = {{chart}}, width = 5, height = 3.5, units = "in", startRow = {{row}}+1, startCol = {{col}})
}


