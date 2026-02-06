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

## Define server logic
shinyServer(function(input, output, session) {
  bcsapps::bcsHeaderServer(id = 'header', links = TRUE)

  ## Adds text to sidebar
  output$last_date <- renderUI({
    div(HTML("Data last updated: ", last_update)#, style = "font-style: italic; "
        )
  })
  output$note <- renderUI({
    div(HTML("Choose one of the countries from <br />
             the list below. The Fact Sheet for <br />
             that country will load within 1 minute.<br />
             <br />
             Once the Fact Sheet has loaded, you <br />
             can click on the `Download XLSX` <br />
             button below to generate an Excel file <br />
             or the `Download PDF` button for a <br />
             PDF version."),
        style = "padding: 10px 10px 10px 30px;")
  })
  ### * STATE VERSION ** ##
  output$note_states <- renderUI({
    div(HTML("Choose one of the USA states from <br />
             the list below. The Fact Sheet for <br />
             that state will load within 1 minute.<br />
             <br />
             Once the Fact Sheet has loaded, you <br />
             can click on the `Download State XLSX` <br />
             button below to generate an Excel file <br />
             or the `Download State PDF` button for a <br />
             PDF version."),
        style = "padding: 10px 10px 10px 30px;")
  })


  ## Runs .RMD file to create html webpage
  output$md_file <- renderUI({
    includeHTML(rmarkdown::render("CountryFactSheet.Rmd",
                                  output_format = "html_document",
                                  params = list(Country = input$Country)))
  })
  ### * STATE VERSION ** ##
  output$md_file_states <- renderUI({
    includeHTML(rmarkdown::render("StateFactSheet.Rmd",
                                  output_format = "html_document",
                                  params = list(state = input$state)))
  })

  ## Adds XLSX download functionality to sidebar
  output$DownloadXLSX <- downloadHandler(

    filename = function() { paste0("Fact", input$Country, ".xlsx") },

    content = function(file) {

      ## send download event tracking to google analytics
      session$sendCustomMessage("trackDownload", list(filename = paste0("Fact", input$Country, ".xlsx")))

      ## Set up parameters to pass to Rmd document
      params <- list(Country = input$Country)

      ## Read in xlsx workbook made/knit in .Rmd, then save it for download
      wb <- loadWorkbook(file = paste0("Fact", input$Country, ".xlsx"))
      saveWorkbook(wb, file = file, overwrite = TRUE)

    }
  )
  ### * STATE VERSION ** ##
  output$DownloadXLSX_states <- downloadHandler(

    filename = function() { paste0("Fact", input$state, ".xlsx") },

    content = function(file) {

      ## send download event tracking to google analytics
      session$sendCustomMessage("trackDownload", list(filename = paste0("Fact", input$state, ".xlsx")))

      ## Set up parameters to pass to Rmd document
      params <- list(state = input$state)

      ## Read in xlsx workbook made/knit in .Rmd, then save it for download
      wb <- loadWorkbook(file = paste0("Fact", input$state, ".xlsx"))
      saveWorkbook(wb, file = file, overwrite = TRUE)

    }
  )

  ## Adds PDF download functionality to sidebar
  output$DownloadPDF <- downloadHandler(

    filename = function() { paste0("Fact", input$Country, ".pdf") },

    content = function(file) {

      ## send download event tracking to google analytics
      session$sendCustomMessage("trackDownload", list(filename = paste0("Fact", input$Country, ".pdf")))

      # this appears to do nothing
      # ## trying to get fontspec & setmainfont to work
      # tempsty <- file.path(tempdir(), "unicode-math.sty")
      # file.copy("unicode-math.sty", tempsty, overwrite = TRUE)
      # tempxesty <- file.path(tempdir(), "unicode-math-xetex.sty")
      # file.copy("unicode-math-xetex.sty", tempxesty, overwrite = TRUE)
      # tempxesty1 <- file.path(tempdir(), "unicode-math-table.tex")
      # file.copy("unicode-math-table.tex", tempxesty1, overwrite = TRUE)
      # tempxesty2 <- file.path(tempdir(), "unicode-math-luatex.sty")
      # file.copy("unicode-math-luatex.sty", tempxesty, overwrite = TRUE)

      out = rmarkdown::render("CountryFactSheet.Rmd",
                              output_file = file,
                              # output_dir = tempdir(),
                              output_format = "pdf_document",
                              # intermediates_dir = tempdir(),
                              envir = new.env(),
                              params = list(Country = input$Country)
      )
    }
  )
  ### * STATE VERSION ** ##
  output$DownloadPDF_states <- downloadHandler(

    filename = function() { paste0("Fact", input$state, ".pdf") },

    content = function(file) {

      ## send download event tracking to google analytics
      session$sendCustomMessage("trackDownload", list(filename = paste0("Fact", input$state, ".pdf")))

      out = rmarkdown::render("StateFactSheet.Rmd",
                              output_file = file,
                              output_format = "pdf_document",
                              envir = new.env(),
                              params = list(state = input$state)
      )
    }

  )

})
