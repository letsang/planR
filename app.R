#
# This is a Shiny web application. You can run the application by clicking
# the 'Run App' button above.
#
# Find out more about building applications with Shiny here:
#
#    http://shiny.rstudio.com/
#

# Shiny
library(shiny)
library(shinyjs)
library(shinyWidgets)
library(shinycssloaders)
# Time series & Forecast
library(forecast)
# Data Visualization
library(ggplot2)
library(shiny)
library(shinyjs)
library(shinyWidgets)
library(htmltools)
# Data Wrangling
library(readxl)
library(dplyr)
library(tidyr)
library(lubridate)
library(aweek)
library(stringr)
library(rebus)
# Data Visualization
library(ggplot2)
library(dygraphs)
library(plotly)
# Table
library(DT)
library(formattable)
library(RCurl)
# Office365
library(Microsoft365R)

### NETWORK ###
# https://frankzheng.me/en/2022/06/run-shiny-apps-locally/
x <- system("ipconfig", intern=TRUE)
z <- x[grep("IPv4", x)]
ip <- gsub(".*? ([[:digit:]])", "\\1", z)
port <- 8888

### ONE DRIVE ###
odb <- get_business_onedrive()
items <- odb$list_items("planR/seasonality")
odb$download_file(src = paste0("planR/seasonality/",items$name[1]), paste0("seasonality\\",items$name[1]), overwrite = TRUE)

### UI FONT CONSTANTS
t <- list(
  family = "Helvetica",
  size = 10,
  color = "black")

t1 <- list(
  family = "Helvetica",
  size = 8,
  color = "black")

### DATABASE ###
base <- readRDS("base.rds")

### SERVER ###
server <- function(input, output, session) {
  
  values <- reactiveValues(
    smc = "698651AAANG1000",
    week0 = 50,
    year0 = 2022,
    items = odb$list_items("planR/seasonality"),
    seasonality = read_excel(paste0("seasonality\\",items$name[1]))
  )
  
  observe({
    values$smc <- input$smc
    # updateSelectizeInput(session, "smc", choices = unique(base$SMC[base$SMC %in% input$dep]))
  })
  
  observeEvent(input$update,{
    odb$download_file(src = paste0("planR/seasonality/", input$weight), paste0("seasonality\\",input$weight), overwrite = TRUE)
    values$seasonality <- read_excel(paste0("seasonality\\",input$weight))
    updateSelectizeInput(session, "weight", choices = odb$list_items("planR/seasonality")$name, selected = NULL)
  })
  
  ### UPDATE PACKSHOT ###
  output$picture <- renderUI({
    htmltools::tags$img(src=paste0("https://getmedia.pdi.keringapps.com/image/pijqmiaszv/", values$smc), height="200px")
  })

  res <- reactive({
    tmp <- base %>% full_join(values$seasonality, by = "Week") %>% filter(SMC == input$smc) %>% mutate(AWS = 0, Forecast = 0, WOC = 4, Closing = 0, Need = 0)
    ### AWS ###  
    for(i in 1:nrow(tmp)){
      if(i %in% 1:4){
        tmp$AWS[i] <- 0
      }else{
        tmp$AWS[i] <- mean(c(tmp$Sales[i-1], tmp$Sales[i-2], tmp$Sales[i-3], tmp$Sales[i-4]))
      }
    }
    
    ### FORECAST & NEED ###
    for(i in 1:nrow(tmp)){
      if(tmp$Week[i] < values$week0 & tmp$Year[i] <= values$year0){
        tmp$Forecast[i] <- 0
        tmp$Closing[i] <- 0
        tmp$Need[i] <- 0
      }else{
        tmp$Forecast[i] <- round(tmp[tmp$Week == values$week0 & tmp$Year == values$year0, "AWS"]*52*tmp$Weight[i])
        tmp$Closing[i] <- round(max(0,(tmp$Closing[i-1] + tmp$OH[i] + tmp$PFP[i] - tmp$Forecast[i])))
        tmp$Need[i] <- round(min(0,(tmp$Closing[i] - tmp$Forecast[i] - tmp$Forecast[i+1] - tmp$Forecast[i+2] - tmp$Forecast[i+3]))*-1)
      }
    }
    tmp$Time <- factor(tmp$Time, levels = unique(tmp$Time), ordered = TRUE)
    tmp %>% select(Time, Week, Year, Weight, SMC, Department, Line, Season, OH, PFP, Sales, AWS, WOC, Forecast, Closing, Need)
  })
  
  output$table = DT::renderDataTable({
    datatable(res(),
              options = list(searching = FALSE, paging = FALSE, columnDefs = list(list(className = 'dt-center', targets = "_all"))),
              extensions = c('Scroller'),
              rownames = FALSE,
              editable = "cell",
              escape = FALSE)
  })
  
  output$plot <- renderPlotly({
    
    plot_ly(res()) %>% 
      add_trace(x = ~Time, y = ~OH, type = "bar", name ="OH") %>%
      add_trace(x = ~Time, y = ~Closing, type = "bar", name ="Stock Proj") %>%
      add_trace(x = ~Time, y = ~Need, type = "bar", name ="Reorder") %>%
      add_trace(x = ~Time, y = ~Forecast, type = "scatter", name ="Forecast", mode='markers+lines', yaxis = "y2", marker = list(color = "red"), opacity=1) %>%
      add_trace(x = ~Time, y = ~Sales, type = "scatter", name ="Sales", mode='markers+lines', yaxis = "y2", marker = list(color = "blue"), opacity=1) %>%
      layout(xaxis = list(title = list(text=""), dtick=1, tickfont = list(size = 8), font = t),
             yaxis = list(title = list(text=""), tickfont = list(size = 8), font = t, rangemode = "nonnegative"),
             yaxis2 = list(title = list(text=""), tickfont = list(size = 8), font = t, rangemode = "nonnegative", overlaying = "y", side = "right"),
             margin = list(r = 50),
             showlegend = FALSE)
    
  })
  
}

### CLIENT ###
ui <- navbarPage(
  htmltools::tags$img(src = "https://www.ysl.com/on/demandware.static/-/Library-Sites-Library-SLP/default/dw86be354f/images/logo.svg", height = "15px"),
  tabPanel("★",
           fluidPage(
             htmltools::tags$style(HTML("body {font-family:'Helvetica',sans-serif; font-size:10px
                                              }
             
                                        .js-irs-0 .irs-single, .js-irs-0 .irs-bar-edge, .js-irs-0 .irs-bar {
                                                  background: black;
                                                  border-top: 1px solid black;
                                                  border-bottom: 1px solid black;
                                                  }

                                        .irs-from, .irs-to, .irs-single { 
                                                  background: black
                                                  }"
             )),
             fluidRow(
               column(6,
                      # selectizeInput("dep", "Department :", choices = unique(base$Department), selected = "HANDBAGS", multiple = TRUE),
                      selectizeInput("smc", "Style Material Color :", choices = unique(base$SMC), selected = "698651AAANG1000", multiple = FALSE, options = list(maxOptions = 10)),
                      # selectizeInput("weight", "Seasonality :", choices = items$name, selected = items$name[1], multiple = FALSE),
                      div(style="display:inline-block;",
                          fluidRow(
                            column(10, selectizeInput("weight", "Seasonality :", choices = items$name, selected = items$name[1], multiple = FALSE)),
                            column(2, actionButton("update", "", icon = icon("refresh"), style="color: #fff; background-color: #000; border-color: #000"))
                            )
                          )
                      ),
               column(6, shinycssloaders::withSpinner(uiOutput("picture"), type = 7)),
               column(12, shinycssloaders::withSpinner(plotlyOutput("plot"), type = 7)),
               column(12, shinycssloaders::withSpinner(DT::dataTableOutput("table"), type = 7))
             ),
           )
  )
)

# Run the application 
app <- shinyApp(ui = ui, server = server)
runApp(app, launch.browser = FALSE, port = port, host = ip)
