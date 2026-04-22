#
# This is a Shiny web application. You can run the application by clicking
# the 'Run App' button above.
#

library(shiny)
library(stringr)
library(officer)

# Function to process one VTT file
vtt_to_docx <- function(vtt_path, docx_path = NULL) {
  lines <- readLines(vtt_path, encoding = "UTF-8")
  time_pattern <- "^\\d{2}:\\d{2}:\\d{2}\\.\\d{3} -->"
  doc <- read_docx()
  speaker_text <- list()
  
  i <- 1
  current_speaker <- ""
  start_time <- ""
  end_time <- ""
  
  while (i <= length(lines)) {
    line <- str_trim(lines[i])
    
    if (str_detect(line, time_pattern)) {
      time_range <- str_split_fixed(line, " --> ", 2)
      start_time <- ifelse(start_time == "", time_range[1], start_time)
      end_time <- time_range[2]
      i <- i + 1
      
      if (i <= length(lines)) {
        next_line <- str_trim(lines[i])
        
        if (str_detect(next_line, "^[^:]+:")) {
          speaker_info <- str_split_fixed(next_line, ": ", 2)
          current_speaker <- speaker_info[1]
          text <- speaker_info[2]
        } else {
          text <- next_line
        }
        
        if (!is.null(speaker_text[[current_speaker]])) {
          speaker_text[[current_speaker]]$text <- paste(speaker_text[[current_speaker]]$text, text, sep = " ")
        } else {
          speaker_text[[current_speaker]] <- list(text = text, start_time = start_time, end_time = end_time)
        }
        
        i <- i + 1
      }
    } else {
      i <- i + 1
    }
  }
  
  for (speaker in names(speaker_text)) {
    doc <- body_add_par(doc, paste0(speaker, " (", speaker_text[[speaker]]$start_time, "-", speaker_text[[speaker]]$end_time, "):"), style = "Normal")
    doc <- body_add_par(doc, speaker_text[[speaker]]$text, style = "Normal")
    doc <- body_add_par(doc, "\n", style = "Normal")
  }
  
  # If no output path specified, use same as input with .docx extension
  if (is.null(docx_path)) {
    docx_path <- sub("\\.vtt$", ".docx", vtt_path)
  }
  print(doc, target = docx_path)
  return(docx_path)
}

# Define UI
ui <- fluidPage(
  titlePanel("Batch VTT to DOCX Converter"),
  sidebarLayout(
    sidebarPanel(
      fileInput("pathfiles", "Upload text file(s) listing VTT paths", accept = ".txt", multiple = TRUE),
      actionButton("convert", "Convert VTT Files to DOCX")
    ),
    mainPanel(
      verbatimTextOutput("log_output")
    )
  )
)

# Define server
server <- function(input, output) {
  observeEvent(input$convert, {
    req(input$pathfiles)
    logs <- c()
    
    for (file in input$pathfiles$datapath) {
      # Read and clean up the VTT paths
      vtt_paths <- readLines(file, encoding = "UTF-8")
      vtt_paths <- str_trim(vtt_paths)
      vtt_paths <- vtt_paths[vtt_paths != ""]
      vtt_paths <- str_replace_all(vtt_paths, "^\"|\"$", "")  # strip quotes
      
      # DEBUG HERE
      cat("Loaded paths:\n", paste(vtt_paths, collapse = "\n"), "\n")
      
      for (vtt_path in vtt_paths) {
        vtt_path <- str_trim(vtt_path)
        
        if (file.exists(vtt_path)) {
          docx_path <- sub("\\.vtt$", ".docx", vtt_path)
          tryCatch({
            out <- vtt_to_docx(vtt_path, docx_path)
            logs <- c(logs, paste("✅ Converted:", vtt_path, "→", out))
          }, error = function(e) {
            logs <- c(logs, paste("❌ Error processing", vtt_path, ":", e$message))
          })
        } else {
          logs <- c(logs, paste("📂 File not found:", vtt_path))
        }
      }
    }
    
    output$log_output <- renderText({
      paste(logs, collapse = "\n")
    })
  })
}

shinyApp(ui = ui, server = server)
