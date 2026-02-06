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

## This is a Shiny web application. You can run the application by clicking the 'Run App' button above.

## Define UI for application
shinyUI(tagList(tags$head(includeHTML("www/google-analytics.html")),
                fluidRow( bcsapps::bcsHeaderUI(id = 'header', appname = "Trade Profiles", github = "https://github.com/bcgov/country-trade-app")),

                dashboardPage(title = "Trade Profiles",


                              # dashboard header
                              dashboardHeader(
                                disable = TRUE),

                              # dashboard sidebar
                              dashboardSidebar(
                                br(),
                                width = 310,
                               tags$head(
                                  tags$style(HTML(".sidebar {
                                        height: 90vh; overflow-y: auto; padding-top: 15px
                                    }"
                                                  ) # close HTML
                                             )      # close tags$style
                                ),                  # close tags$Head
                                sidebarMenu(id = "sidebar",
                                            menuItem(uiOutput("last_date"), icon = NULL),
                                            menuItem("B.C. & Canada", tabName = "hometab", icon = NULL),
                                            menuItem("Country Fact Sheets", tabName = "profilestab", icon = NULL,
                                                     uiOutput("note"),
                                                     menuSubItem(tabName = "profilestab", icon = NULL,
                                                                 selectInput("Country",
                                                                             label = "Country",
                                                                             selected = "Click here to choose a country",
                                                                             choices = c("Click here to choose a country", countries))),
                                                     br(),br(),br(),br(),br(),br(),br(),br(),br(),br(),
                                                     menuSubItem(tabName = "profilestab", icon = NULL,
                                                                 downloadButton("DownloadXLSX", label = "Download XLSX")
                                                                 ),
                                                     menuSubItem(tabName = "profilestab", icon = NULL,
                                                                 downloadButton("DownloadPDF", label = "Download PDF")
                                                     )
                                            ),
                                            menuItem("USA State Fact Sheets", tabName = "statestab", icon = NULL,
                                                     uiOutput("note_states"),
                                                     menuSubItem(tabName = "statestab", icon = NULL,
                                                                 selectInput("state", ## this must match params in .Rmd
                                                                             label = "USA State",
                                                                             selected = "Click here to choose a state",
                                                                             choices = c("Click here to choose a state", states))),
                                                     br(),br(),br(),br(),br(),br(),br(),br(),br(),br(),
                                                     menuSubItem(tabName = "statestab", icon = NULL,
                                                                 downloadButton("DownloadXLSX_states", label = "Download State XLSX")
                                                     ),
                                                     menuSubItem(tabName = "statestab", icon = NULL,
                                                                 downloadButton("DownloadPDF_states", label = "Download State PDF")
                                                     )
                                            )
                                )
                              ),

                              # dashboard body
                              dashboardBody(

                                tags$style(type = "text/css",
                                           ".shiny-output-error {visibility: hidden;}",
                                           ".shiny-output-error:before {visibility: hidden;}",
                                           ".content-wrapper {margin-top:15px;}"),
                                tags$script(HTML("$('body').addClass('fixed');")),
                                tags$head(tags$meta(name = "viewport", content = "width=1500")), # mobile friendly version of the app
                                tags$link(rel = "stylesheet", type = "text/css", href = "style.css"),
                                tabItems(
                                  tabItem(tabName = "hometab",
                                          fluidRow(
                                            box(title = "Aggregate Data for B.C. & Canada",
                                                width = 12,
                                                status = "success",
                                                solidHeader = TRUE,
                                                collapsible = FALSE,
                                                ## table 1
                                                h4(paste0("Top five B.C. origin exports to world, ", year),
                                                   style = "color:#2A64AB; font-weight: bold"),
                                                DT::datatable(table1, rownames = T,   ## set rownames TRUE, so can reference them in formatStyle
                                                              options = list(dom = 't',
                                                                             columnDefs = list(list(className = 'dt-right', targets = 2:3),  ## right alighn columns 2&3
                                                                list(visible = F, targets = 0))), ## hide the rownames
                                                              height = "100%") %>%
                                                  formatStyle(0, target = "row", fontWeight = styleEqual(levels = dim(table1)[1], values = 'bold')),
                                                br(),br(),
                                                ## table 2
                                                h4(paste0("Top five Canadian exports to world, ", year),
                                                   style = "color:#2A64AB; font-weight: bold"),
                                                DT::datatable(table2, rownames = T,
                                                              options = list(dom = 't', columnDefs = list(list(className = 'dt-right', targets = 2:3), list(visible = F, targets = 0))),
                                                              height = "100%") %>%
                                                  formatStyle(0, target = "row", fontWeight = styleEqual(levels = dim(table2)[1], values = 'bold')),
                                                br(),br(),
                                                ## table 3
                                                h4(paste0("Top five Canadian imports from world, ", year),
                                                   style = "color:#2A64AB; font-weight: bold"),
                                                DT::datatable(table3, rownames = T,
                                                              options = list(dom = 't',  columnDefs = list(list(className = 'dt-right', targets = 2:3), list(visible = F, targets = 0))),
                                                              height = "100%") %>%
                                                  formatStyle(0, target = "row", fontWeight = styleEqual(levels = dim(table3)[1], values = 'bold')),
                                                br(),br(),
                                                ## table 4
                                                h4(paste0("Provincial distribution of exports to world, ", year),
                                                   style = "color:#2A64AB; font-weight: bold"),
                                                DT::datatable(table4,
                                                              rownames = T,
                                                              options = list(dom = 't',
                                                                             pageLength = -1, ## pageLength = -1 shows ALL rows
                                                                             columnDefs = list(list(className = 'dt-right', targets = 2:3), list(visible = F, targets = 0)))
                                                              ) %>%
                                                  formatStyle(0, target = "row", fontWeight = styleEqual(levels = c(BCrow, dim(table4)[1]), values = c('bold', 'bold'))),
                                                br(),br(),br(),
                                                "Data source: Statistics Canada",
                                                br(),
                                                paste0("Data last updated: ", last_update)
                                            ))),
                                  tabItem(tabName = "profilestab",
                                          fluidRow(
                                            box(title = "Country Trade Profiles",
                                                width = 12,
                                                status = "success",
                                                solidHeader = TRUE,
                                                collapsible = FALSE,
                                                uiOutput("md_file")
                                                ))),
                                  tabItem(tabName = "statestab",
                                          fluidRow(
                                            box(title = "USA State Trade Profiles",
                                                width = 12,
                                                status = "success",
                                                solidHeader = TRUE,
                                                collapsible = FALSE,
                                                uiOutput("md_file_states")
                                            )))
                                  )  # close tabItems

                                ,
                                # BC gov required footer
                                tags$footer(class="footer",
                                            tags$div(class="container", style="display:flex; justify-content:center; flex-direction:column; text-align:center; height:46px;",
                                                     tags$ul(style="display:flex; flex-direction:row; flex-wrap:wrap; margin:0; list-style:none; align-items:center; height:100%;",
                                                             tags$li(a(href="https://www2.gov.bc.ca/gov/content/home", "Home", style="font-size:1em; font-weight:normal; color:white; padding-left:5px; padding-right:5px; border-right:1px solid #4b5e7e;")),
                                                             tags$li(a(href="https://www2.gov.bc.ca/gov/content/home/disclaimer", "Disclaimer", style="font-size:1em; font-weight:normal; color:white; padding-left:5px; padding-right:5px; border-right:1px solid #4b5e7e;")),
                                                             tags$li(a(href="https://www2.gov.bc.ca/gov/content/home/privacy", "Privacy", style="font-size:1em; font-weight:normal; color:white; padding-left:5px; padding-right:5px; border-right:1px solid #4b5e7e;")),
                                                             tags$li(a(href="https://www2.gov.bc.ca/gov/content/home/accessibility", "Accessibility", style="font-size:1em; font-weight:normal; color:white; padding-left:5px; padding-right:5px; border-right:1px solid #4b5e7e;")),
                                                             tags$li(a(href="https://www2.gov.bc.ca/gov/content/home/copyright", "Copyright", style="font-size:1em; font-weight:normal; color:white; padding-left:5px; padding-right:5px; border-right:1px solid #4b5e7e;")),
                                                             tags$li(a(href="https://www2.gov.bc.ca/StaticWebResources/static/gov3/html/contact-us.html", "Contact", style="font-size:1em; font-weight:normal; color:white; padding-left:5px; padding-right:5px; border-right:1px solid #4b5e7e;"))
                                                             )
                                                   )
                                       ) # close footer tag
                              ) # close dashboardBody
               ) # close dashboardPage
       ) # close tagList
) # close shinyUI
