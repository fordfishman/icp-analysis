load("data/IPC_data.Rdata")
library(dplyr)
library(argparse)
library(readxl)
library(xlsx)

# User interface ----
ui <- fluidPage(
  titlePanel("ICP Analysis V1.07"),
  
  sidebarLayout(
    sidebarPanel(
      helpText(p("App by Ford Fishman for the Peterson Lab. Compares sample analyte concentration to background concentrations."),
               p("Upload files before running program. Download file afterwards."),
               p("* Required Manual Entry")
    ),
      
      fileInput("file", h4("Name of file output from IPC (.xlsx)*"), multiple = FALSE, 
                accept = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
      fileInput("weights", h4("Name of file containing weights of foams (.xlsx)*"), multiple = FALSE,
                accept = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"),
      textInput("output", h4("Desired name for output file"), value = "IPC_analysis.xlsx"),
      actionButton("do", "Run analysis"),
      downloadButton("downloadData", "Download results"),
      textOutput("check")
    ),
    mainPanel(
      column(4,
        h3("Optional parameter changes:"),
        sliderInput("rsquare_cutoff", 
                    h4("R-Squared Cutoff for Drift Correction Regressions:"),
                    min = 0, max = 1, value = 0.95),
        textInput("sample_start", h4("Starting character for sample IDs"), 
                  value = "A"),
        textInput("sample_end", h4("Ending character for sample IDs"), 
                value = "C"))
    
    )
  )
)

# Server logic ----
server <- function(input, output) {
  data <- reactive({
    list(file = input$file,
    weights = input$weights,
    output = input$output,
    rsquare_cutoff = input$rsquare_cutoff,
    sample_start = input$sample_start,
    sample_end = input$sample_end)
  })
  dfs <- eventReactive(input$do, {
    withProgress(message = "Running Analysis", value = 0, min = 0, max = 1, {
      if (is.null(data()[["file"]]) | is.null(data()[["weights"]]))
        return(NULL)
      incProgress(0.1, detail = "Stating up")
      data_ <- data()
      file <- data_[["file"]]$datapath
      weights <- data_[["weights"]]$datapath
      output_thing <- data_[["output"]]
      rsquare_cutoff <- data_[["rsquare_cutoff"]]
      sample_start <- data_[["sample_start"]]
      sample_end <- data_[["sample_end"]]
      incProgress(0.1, detail = "Running main script")
      source("analysis.R", local = TRUE)
      incProgress(0.7, detail = "Finishing")
      list(final_important_df, QA_QC_final, drift_corrections_df, corrected_data_df, all_p_values, final_conc_df, final_sd_df)
    })
  })
  output$check <- renderText({
    if (is.null(dfs()))
      return("Missing files")
    "Analysis complete"
    
  })
  output$downloadData <- downloadHandler(
    filename = function() {
      output_name <- data()[["output"]]
      # Ensure that the desired output name has the correct file extension
      if (! endsWith(output_name, ".xlsx")) {
        output_name <- paste(output_name, "xlsx", sep = ".")
      }
      output_name
      },
    content = function(file) {
      write.xlsx(dfs()[1], file = file, sheetName = "FinalAnalysis", append = FALSE, col.names = T, row.names = T)
      write.xlsx(dfs()[2], file = file, sheetName = "QA_QC", append = TRUE, col.names = T, row.names = T)
      write.xlsx(dfs()[3], file = file, sheetName = "DriftCorrection", append = TRUE, col.names = T, row.names = F)
      write.xlsx(dfs()[4], file = file, sheetName = "Drift_CCBCorrection", append = TRUE, col.names = T, row.names = F)
      write.xlsx(dfs()[5], file = file, sheetName = "PvaluesComparingToBackground", append = TRUE, col.names = T, row.names = T)
      write.xlsx(dfs()[6], file = file, sheetName = "AllFieldConcentrations", append = TRUE, col.names = T, row.names = T)
      write.xlsx(dfs()[7], file = file, sheetName = "AllFieldSDs", append = TRUE, col.names = T, row.names = T)
    }
  )
}

# Run app ----
shinyApp(ui, server)