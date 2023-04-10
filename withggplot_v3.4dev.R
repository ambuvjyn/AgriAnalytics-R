library(shiny)
library(readxl)
library(markdown)
library(agricolae)

ui <- fluidPage(
  navbarPage(title = div(img(src = "favicon.png", height = 30), "AgriAnalytics@R Ver. 0.1"),
             header = list(
               tags$div(
                 class = "navbar-header",
                 HTML(paste0("<a class='navbar-brand' href='#'>",
                             "<img src='ICAR.png' height='30'></a>")),
                 style = "float: right;"
               )
             ),
             tabPanel("Upload",
                      fluidRow(
                        column(width = 4,
                               fileInput("file", "Choose Excel File",
                                         accept = c(".xlsx", ".csv")),
                               tableOutput("data_preview")
                        ),
                        column(width = 8)
                      )
             ),
             tabPanel("Descriptive Statistics",
                      fluidRow(
                        column(width = 4,
                               selectInput("desc_xvar", "Select a Variable",
                                           choices = NULL)
                        ),
                        column(width = 8,
                               verbatimTextOutput("summary"),
                               plotOutput("boxplot")
                        )
                      )
             ),
             tabPanel("Anova_CRD",
                      fluidRow(
                        column(width = 4,
                               selectInput("xvar_crd", "Select X Variable",
                                           choices = NULL),
                               selectInput("yvar_crd", "Select Y Variable",
                                           choices = NULL),
                               actionButton("run_crd", "Run ANOVA")
                        ),
                        column(width = 8,
                               verbatimTextOutput("anova_crd"),
                               plotOutput("plot_crd")
                        )
                      )
             ),
             tabPanel("Anova_RBD",
                      fluidRow(
                        column(width = 4,
                               selectInput("xvar_rbd", "Select X Variable",
                                           choices = NULL),
                               selectInput("yvar_rbd", "Select Y Variable",
                                           choices = NULL),
                               selectInput("rep_var_rbd", "Select Replication Factor",
                                           choices = NULL),
                               actionButton("run_rbd", "Run ANOVA")
                        ),
                        column(width = 8,
                               verbatimTextOutput("anova_rbd"),
                               verbatimTextOutput("anova_rbd_2"),
                               downloadButton("download_anova", "Download ANOVA Results as RTF"),
                               plotOutput("plot_rbd"),
                               downloadButton("download_plot", "Download ANOVA PLOT as PNG")
                        )
                      )
             ),
             tabPanel("About", 
                      uiOutput("about_ui")
             )
             
  ))

server <- function(input, output, session) {
  df <- reactive({
    req(input$file)
    file <- input$file
    ext <- tools::file_ext(file$datapath)
    if (ext %in% c("xlsx", "xls")) {
      read_excel(file$datapath)
    } else if (ext == "csv") {
      read.csv(file$datapath)
    } else {
      stop("Invalid file type")
    }
  })
  
  # Data preview
  output$data_preview <- renderTable({
    df()
  })
  
  # Descriptive statistics output
  observeEvent(df(), {
    updateSelectInput(session, "desc_xvar", "Select a Variable", choices = names(df()))
  })
  
  # Descriptive statistics output
  output$summary <- renderPrint({
    x <- input$desc_xvar
    summary(df()[[x]])
  })
  
  output$boxplot <- renderPlot({
    y <- input$desc_xvar
    boxplot(df()[[y]], main = paste("Boxplot of", y))
  })
  
  # ANOVA_rbd results output
  observeEvent(df(), {
    numeric_cols <- sapply(df(), is.numeric)
    updateSelectInput(session, "xvar_crd", "Choose the Treatment variable", choices = names(df())[numeric_cols])
    updateSelectInput(session, "yvar_crd", "Choose the Response variable", choices = names(df())[numeric_cols])
  })
  
  observeEvent(input$run_crd, {
    if (!is.null(df())) {
      x <- input$xvar_crd
      y <- input$yvar_crd
      model_crd <- lm(paste(y, "~", x), df())
      model_lsd_crd <- LSD.test(model_crd, x, p.adj = "none")
      anova_output <- anova(model_crd)
      output$anova_crd <- renderPrint({
        anova_output
      })
      output$plot_crd <- renderPlot({
        plot(model_lsd_crd)
      })
    }
  })
  
  # ANOVA_crd results output
  observeEvent(df(), {
    numeric_cols <- sapply(df(), is.numeric)
    updateSelectInput(session, "xvar_rbd", "Choose the Treatment variable", choices = names(df())[numeric_cols])
    updateSelectInput(session, "yvar_rbd", "Choose the Response variable", choices = names(df())[numeric_cols])
    updateSelectInput(session, "rep_var_rbd", "Choose the Replication factor", choices = names(df())[numeric_cols])
  })
  
  observeEvent(input$run_rbd, {
    if (!is.null(df())) {
      x <- input$xvar_rbd
      y <- input$yvar_rbd
      r <- input$rep_var_rbd
      model_rbd <- lm(paste(y, "~", x, "+", r), data = df())
      model_lsd_rbd <- LSD.test(model_rbd, x, p.adj = "none")
      
      
      #Tab ANOVA First row
      output$anova_rbd <- renderPrint({
        model_lsd_rbd
      })
      
      #Tab ANOVA Second row
      output$anova_rbd_2 <- renderPrint({
        model_rbd
      })
      
      # Download ANOVA results as RTF
      filename <- paste0("anova_results_", Sys.Date(), ".rtf")
      sink(filename)
      print(model_lsd_rbd)
      print(model_rbd)
      sink()
      output$download_anova <- downloadHandler(
        filename = function() {
          filename
        },
        content = function(file) {
          file.copy(filename, file)
        }
      )
      
      #Tab ANOVA Third row
      output$plot_rbd <- renderPlot({
        plot(model_lsd_rbd)
      })
      
      # Download ANOVA plot as PNG
      output$download_plot <- downloadHandler(
        filename = function() {
          paste("anova_plot", ".png", sep = "")
        },
        content = function(file) {
          png(file)
          plot(model_lsd_rbd)
          dev.off()
        }
      )
      
    }
  })
  
  # Read the contents of the README.md file
  about_text <- readLines("README.md")
  
  # Render the contents of the README.md file as Markdown
  output$about_ui <- renderText({
    HTML(renderMarkdown(about_text))
  })
}

shinyApp(ui, server)
