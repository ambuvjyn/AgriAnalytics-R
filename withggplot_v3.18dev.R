##############################Packages_Needed###################################

library(shiny)
library(readxl)
library(agricolae)
library(markdown)
library(bslib)
library(pastecs)
library(corrplot)
library(GGally)
library(ggplot2)
library(correlation)
library(flextable)
library(fixr)


#####################################UI#########################################

ui <- fluidPage(theme = bs_theme(version = 4, bootswatch = "litera"),
                navbarPage(title = div(img(src = "favicon.png", height = 30), "AgriAnalytics@R Ver. 3.17", img(src = "ICAR.png", height = 30)),

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
                                      column(width = 3,
                                             selectInput("desc_xvar", "Select a Variable",
                                                         choices = NULL),
                                             selectInput("desc_grvar", "Select a Group Variable",
                                                         choices = NULL)
                                      ),
                                      column(width = 9,
                                             plotOutput("boxplot"),
                                             uiOutput("summary")
                                      )
                                    )
                           ),
                           
                           tabPanel("ANOVA",
                                    tabsetPanel(
                                      id = "anova_tabs",
                                      tabPanel("Anova CRD",
                                               fluidRow(
                                                 column(width = 4,
                                                        selectInput("xvar_crd", "Select Treatment Variable",
                                                                    choices = NULL),
                                                        selectInput("yvar_crd", "Select Response Variable",
                                                                    choices = NULL, multiple = TRUE),
                                                        actionButton("run_crd", "Run ANOVA")
                                                 ),
                                                 column(width = 8,
                                                        div(
                                                          verbatimTextOutput("anova_crd"),
                                                          style = "max-height: 500px; overflow-y: scroll;"),
                                                        downloadButton("download_anova_crd", "Download ANOVA Results as RTF")
                                                 )
                                               )
                                      ),
                                      
                                      tabPanel("Plot CRD",
                                               fluidRow(
                                                 column(width = 4,
                                                        selectInput("xvar_crd_plot", "Select Treatment Variable",
                                                                    choices = NULL),
                                                        selectInput("yvar_crd_plot", "Select Response Variable",
                                                                    choices = NULL),
                                                        numericInput("plot_crd_width", "Plot Width", 800),
                                                        numericInput("plot_crd_height", "Plot Height", 800),
                                                        numericInput("plot_crd_resolution", "Plot Resolution", 120),
                                                        actionButton("run_crd_plot", "Plot ANOVA")
                                                        
                                                 ),
                                                 column(width = 8,
                                                        plotOutput("plot_crd"),
                                                        downloadButton("download_plot_crd", "Download ANOVA PLOT as PNG")
                                                 )
                                               )
                                      ),
                                      
                                      tabPanel("Anova RBD",
                                               fluidRow(
                                                 column(width = 4,
                                                        selectInput("xvar_rbd", "Select Treatment Variable",
                                                                    choices = NULL),
                                                        selectInput("rep_var_rbd", "Select Replication Variable",
                                                                    choices = NULL),
                                                        selectizeInput("yvar_rbd", "Select Response Variable",
                                                                       choices = NULL, multiple = TRUE),
                                                        actionButton("run_rbd", "Run ANOVA")
                                                 ),
                                                 column(width = 8,
                                                        div(
                                                          verbatimTextOutput("anova_rbd"),
                                                          style = "max-height: 500px; overflow-y: scroll;"),
                                                        downloadButton("download_anova_rbd", "Download ANOVA Results as RTF")
                                                 )
                                               )
                                      ),
                                      
                                      tabPanel("Plot RBD",
                                               fluidRow(
                                                 column(width = 4,
                                                        selectInput("xvar_rbd_plot", "Select Treatment Variable",
                                                                    choices = NULL),
                                                        selectInput("rep_var_rbd_plot", "Select Replication Variable",
                                                                    choices = NULL),
                                                        selectInput("yvar_rbd_plot", "Select Response Variable",
                                                                    choices = NULL),
                                                        numericInput("plot_rbd_width", "Plot Width", 800),
                                                        numericInput("plot_rbd_height", "Plot Height", 800),
                                                        numericInput("plot_rbd_resolution", "Plot Resolution", 120),
                                                        actionButton("run_rbd_plot", "Plot ANOVA"),
                                                         
                                                 ),
                                                 column(width = 8,
                                                        plotOutput("plot_rbd"),
                                                        downloadButton("download_plot_rbd", "Download ANOVA PLOT as PNG")
                                                 )
                                               )
                                      )
                                    )
                           ),
                           
                           tabPanel("Correlation",
                                    fluidRow(
                                      column(width = 4,
                                             selectizeInput("xvar_cor", "Select Variables",
                                                            choices = NULL, multiple = TRUE
                                             ),
                                             uiOutput("cor_output_selector"),
                                             actionButton("run_cor", "Run Correlation")
                                             
                                      ),
                                      column(width = 8,
                                             conditionalPanel(condition = "input.cor_output=='cor_out'",
                                                              verbatimTextOutput("cor_out"),
                                                              downloadButton("download_cor", "Download Correlation Results as RTF")
                                             ),
                                             conditionalPanel(condition = "input.cor_output=='cor_out_2'",
                                                              verbatimTextOutput("cor_out_2"),
                                                              downloadButton("download_cor_2", "Download Individual Correlation Results as RTF")
                                             ),
                                             conditionalPanel(condition = "input.cor_output=='plot_cor'",
                                                              plotOutput("plot_cor"),
                                                              downloadButton("download_plot_cor", "Download Correlogram as PNG")
                                             ),
                                             conditionalPanel(condition = "input.cor_output=='plot_cor_2'",
                                                              plotOutput("plot_cor_2"),
                                                              downloadButton("download_plot_cor_2", "Download Scatterplot Matrix as PNG")
                                             )
                                      )
                                    )
                           ),
                           
                           tabPanel("About",
                                    uiOutput("about_ui")
                           )
                           
                )
)


################################################################################

####################################Server######################################

server <- function(input, output, session) {
  df_ori <- reactive({
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
  
  df <- reactive({
    fix_column_names(df_ori())
  })
  
  
  ###################################UPLOAD#######################################
  # Data preview
  output$data_preview <- renderTable({
    df()
  })
  
  ############################DESCRIPTIVE STATISTICS##############################
  # Descriptive statistics output
  observeEvent(df(), {
    updateSelectInput(session, "desc_xvar", "Select Response Variable", choices = names(df()))
    updateSelectInput(session, "desc_grvar", "Select Group Variable", choices = c("None", names(df())))
  })
  
  output$summary <- renderUI({
    dat <- data.frame(
      Descriptors = c("nbr.val", "nbr.null", "nbr.na", "min", "max", "range", "sum", "median", "mean", "SE.mean", "CI.mean.0.95", "var", "std.dev", "coef.var"),
      stat.desc(df()) %>% round(2))
    tbl <- flextable(dat)
    htmltools_value(tbl)
  })
  
  output$boxplot <- renderPlot({
    y <- input$desc_xvar
    grv <- input$desc_grvar
    
    if (grv == "None") {
      boxplot(df()[[y]], main = paste("Boxplot of", y), xlab = "",
              ylab = paste(y))
    } else {
      boxplot(df()[[y]]~df()[[grv]], main = paste("Boxplot of", y), xlab = paste(grv),
              ylab = paste(y))
    }
  })
  
  
  ###################################ANOVA CRD####################################
  # ANOVA_crd results output
  
  observeEvent(df(), {
    numeric_cols <- sapply(df(), is.numeric)
    updateSelectInput(session, "xvar_crd", "Choose the Treatment variable", choices = names(df()))
    updateSelectInput(session, "yvar_crd", "Choose the Response variable", choices = names(df())[numeric_cols])
  })
  
  observeEvent(input$run_crd, {
    if (!is.null(df())) {
      x <- input$xvar_crd
      y <- input$yvar_crd
      if (is.null(y)) {
        return()
      }
      if (length(y) > 1) {
        # create a list to store the results for each response variable
        anova_list <- list()
        for (i in 1:length(y)) {
          # fit the ANOVA model for each response variable
          model_crd <- lm(paste(y[i], "~", x), data = df())
          model_lsd_crd <- LSD.test(model_crd, x, p.adj = "none")
          anova_list[[y[i]]] <- list(anova = anova(model_crd), lsd_groups = model_lsd_crd$groups)
        }
      } else {
        # fit the ANOVA model for a single response variable
        model_crd <- lm(paste(y, "~", x), data = df())
        model_lsd_crd <- LSD.test(model_crd, x, p.adj = "none")
        # store the results in a list
        anova_list <- list(anova = anova(model_crd), lsd_groups = model_lsd_crd$groups)
      }
      
      #Tab ANOVA First row
      output$anova_crd <- renderPrint({
        anova_list
      })
      
      # Download ANOVA results as RTF
      filename <- paste0("anova_crd_results_", Sys.Date(), ".rtf")
      sink(filename)
      print(anova_list)
      sink()
      output$download_anova_crd <- downloadHandler(
        filename = function() {
          filename
        },
        content = function(file) {
          file.copy(filename, file)
        }
      )
      
      #Tab ANOVA Third row
      output$plot_crd <- renderPlot({
        plot(model_lsd_crd)
      })
      
      #Download ANOVA plot as PNG
      output$download_plot_crd <- downloadHandler(
        filename = function() {
          paste("anova_plot", ".png", sep = "")
        },
        content = function(file) {
          png(file, width = 800, height = 800, res = 120)
          par(cex.lab = 1.5, cex.main = 1.5, mar = c(2,2,2,2))
          plot(model_lsd_crd)
          dev.off()
        }
      )
    }
  })
  
  ###################################ANOVA RBD####################################
  # ANOVA_rbd results output
  
  observeEvent(df(), {
    numeric_cols <- sapply(df(), is.numeric)
    updateSelectInput(session, "xvar_rbd", "Choose the Treatment variable", choices = names(df())[numeric_cols])
    updateSelectInput(session, "rep_var_rbd", "Choose the Replication factor", choices = names(df()))
    updateSelectInput(session, "yvar_rbd", "Choose the Response variable", choices = names(df())[numeric_cols])
  })
  
  observeEvent(input$run_rbd, {
    if (!is.null(df())) {
      x <- input$xvar_rbd
      y <- input$yvar_rbd
      r <- input$rep_var_rbd
      if (is.null(y)) {
        return()
      }
      if (length(y) > 1) {
        # create a list to store the results for each response variable
        anova_list <- list()
        for (i in 1:length(y)) {
          # fit the ANOVA model for each response variable
          model_rbd <- lm(paste(y[i], "~", r, "+", x), data = df())
          model_lsd_rbd <- LSD.test(model_rbd, x, p.adj = "none")
          anova_list[[y[i]]] <- list(anova = anova(model_rbd), lsd_groups = model_lsd_rbd$groups)
        }
      } else {
        # fit the ANOVA model for a single response variable
        model_rbd <- lm(paste(y, "~", r, "+", x), data = df())
        model_lsd_rbd <- LSD.test(model_rbd, x, p.adj = "none")
        # store the results in a list
        anova_list <- list(anova = anova(model_rbd), lsd_groups = model_lsd_rbd$groups)
      }
      
      #Tab ANOVA First row
      output$anova_rbd <- renderPrint({
        anova_list
      })
      
      #Download ANOVA results as RTF
      filename <- paste0("anova_results_", Sys.Date(), ".rtf")
      sink(filename)
      print(anova_list)
      sink()
      output$download_anova_rbd <- downloadHandler(
        filename = function() {
          filename
        },
        content = function(file) {
          file.copy(filename, file)
        }
      )
    }
  })
  
  ###################################ANOVA CRD PLOT####################################
  # ANOVA_crd results plot output
  
  observeEvent(df(), {
    numeric_cols <- sapply(df(), is.numeric)
    updateSelectInput(session, "xvar_crd_plot", "Choose the Treatment variable", choices = names(df()))
    updateSelectInput(session, "yvar_crd_plot", "Choose the Response variable", choices = names(df())[numeric_cols])
  })
  
  observeEvent(input$run_crd_plot, {
    if (!is.null(df())) {
      x <- input$xvar_crd_plot
      y <- input$yvar_crd_plot
      if (is.null(y)) {
        return()
      }
      # fit the ANOVA model for a single response variable
      model_crd_plot <- lm(paste(y, "~", x), data = df())
      model_lsd_crd_plot <- LSD.test(model_crd_plot, x, p.adj = "none")
      
      #Tab ANOVA Third row
      output$plot_crd <- renderPlot({
        plot(model_lsd_crd_plot, xlab = input$xvar_crd_plot, ylab = input$yvar_crd_plot, 
             width = input$plot_crd_width, height = input$plot_crd_height)
      })
      
      #Download ANOVA plot as PNG
      output$download_plot_crd <- downloadHandler(
        filename = function() {
          paste("anova_plot", ".png", sep = "")
        },
        content = function(file) {
          png(file, width = input$plot_crd_width, height = input$plot_crd_height, res = input$plot_crd_resolution)
          par(cex.lab = 1.5, cex.main = 1.5, mar = c(2, 2, 2, 2))
          plot(model_lsd_crd_plot, xlab = input$xvar_crd_plot, ylab = input$yvar_crd_plot,
               width = input$plot_crd_width, height = input$plot_crd_height)
          dev.off()
        }
      )
    }
  })
  
  ###################################ANOVA RBD PLOT####################################
  # ANOVA_rbd results plot output
  
  observeEvent(df(), {
    numeric_cols <- sapply(df(), is.numeric)
    updateSelectInput(session, "xvar_rbd_plot", "Choose the Treatment variable", choices = names(df())[numeric_cols])
    updateSelectInput(session, "rep_var_rbd_plot", "Choose the Replication factor", choices = names(df()))
    updateSelectInput(session, "yvar_rbd_plot", "Choose the Response variable", choices = names(df())[numeric_cols])
  })
  
  observeEvent(input$run_rbd_plot, {
    if (!is.null(df())) {
      x <- input$xvar_rbd_plot
      y <- input$yvar_rbd_plot
      r <- input$rep_var_rbd_plot
      if (is.null(y)) {
        return()
      }
      # fit the ANOVA model for a single response variable
      model_rbd_plot <- lm(paste(y, "~", r, "+", x), data = df())
      model_lsd_rbd_plot <- LSD.test(model_rbd_plot, x, p.adj = "none")
      
      # Tab ANOVA Third row
      output$plot_rbd <- renderPlot({
        plot(model_lsd_rbd_plot, xlab = input$xvar_rbd_plot, ylab = input$yvar_rbd_plot)
      })
      
      # Download ANOVA plot as PNG
      output$download_plot_rbd <- downloadHandler(
        filename = function() {
          paste("anova_plot", ".png", sep = "")
        },
        content = function(file) {
          png(file, width = input$plot_rbd_width, height = input$plot_rbd_height, res = input$plot_rbd_resolution)
          par(cex.lab = 1.5, cex.main = 1.5, mar = c(5, 4, 4, 2))
          plot(model_lsd_rbd_plot, xlab = input$xvar_rbd_plot, ylab = input$yvar_rbd_plot,
               width = input$plot_rbd_width, height = input$plot_rbd_height)
          dev.off()
        }
      )
      
    }
  })
  
  
  #################################CORRELATION####################################
  # Correlation results output
  
  observeEvent(df(), {
    numeric_cols <- sapply(df(), is.numeric)
    updateSelectizeInput(session, "xvar_cor", "Choose Variables",
                         choices = names(df())[numeric_cols])
  })
  
  observeEvent(input$run_cor, {
    if (!is.null(df())) {
      x <- input$xvar_cor
      y <- input$yvar_cor
      model_cor <- round(cor(df()[, c(x, y = NULL)]), 2)
      model_cor_2 <- correlation::correlation((df()[, c(x, y = NULL)]),include_factors = TRUE, method = "auto")
      
      #Tab Correlation First row
      output$cor_out <- renderPrint({
        model_cor
      })
      
      
      #Download Correlation  as RTF
      filename <- paste0("Correlation_results_", Sys.Date(), ".rtf")
      sink(filename)
      print(model_cor)
      sink()
      output$download_cor <- downloadHandler(
        filename = function() {
          filename
        },
        content = function(file) {
          file.copy(filename, file)
        }
      )
      
      #Tab Correlation Second row
      output$cor_out_2 <- renderPrint({
        model_cor_2
      })
      
      
      #Download Correlation  as RTF
      filename <- paste0("Correlation_results_withp_", Sys.Date(), ".rtf")
      sink(filename)
      print(model_cor_2)
      sink()
      output$download_cor_2 <- downloadHandler(
        filename = function() {
          filename
        },
        content = function(file) {
          file.copy(filename, file)
        }
      )
      
      #Tab Correlation Third row
      output$plot_cor <- renderPlot({
        corrplot(model_cor)
      })
      
      #Download corrplot as PNG
      output$download_plot_cor <- downloadHandler(
        filename = function() {
          paste("cor_plot", ".png", sep = "")
        },
        content = function(file) {
          png(file, width = 800, height = 800, res = 120)
          par(cex.lab = 1.5, cex.main = 1.5, mar = c(2,2,2,2))
          corrplot(model_cor)
          dev.off()
        }
      )
      
      #Tab Correlation Third row
      output$plot_cor_2 <- renderPlot({
        ggpairs(df()[, x])
      })
      
      # Download corrplot as PNG_2
      output$download_plot_cor_2 <- downloadHandler(
        filename = function() {
          paste("scatter_plot", ".png", sep = "")
        },
        content = function(file) {
          ggsave(file, plot = ggpairs(df()[, x]))
        }
      )
      
    }
  })
  
  # Correlation output selector
  output$cor_output_selector <- renderUI({
    selectInput(
      "cor_output",
      "Select correlation output",
      choices = c("Correlation Matrix" = "cor_out",
                  "Correlation Matrix with P value" = "cor_out_2",
                  "Correlogram" = "plot_cor",
                  "Scatterplot Matrix" = "plot_cor_2")
    )
  })
  
  # Display selected correlation output
  output$selected_cor_output <- renderPrint({
    eval(parse(text = input$cor_output))
  })
  
  ###################################ABOUT########################################
  #Read the contents of the README.md file
  about_text <- readLines("README.md")
  
  #Render the contents of the README.md file as Markdown
  output$about_ui <- renderText({
    HTML(renderMarkdown(about_text))
  })
  
}

#################################App_Exec#######################################

shinyApp(ui, server)
