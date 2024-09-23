library(tidyverse)
library(prophet)
library(readxl)
library(writexl)
library(stringr)
library(tictoc)

setwd(
  'A:/Shared drives/Forecasting/Planning & Analytics/Projects/2022/R/Forecasting')

# Start timer
tic()

# Specify usage file name
usage_file <- 'SB.xlsx'

# Specify number of sheets in Excel workbook
sheets_n <- 1

# Specify floor and cap (if logistic)
floor <- 4000
cap <- 28000

# Read in data from .xlsx
df_all <- list()
for (i in 1:sheets_n) {
  df_all[[excel_sheets(usage_file)[i]]] <- read_excel(usage_file, sheet = i)
  df_all[[i]]$floor <- floor
  df_all[[i]]$cap <- cap
}

# Create empty list
sheets <- list()
plots <- list()

# Forecast engine
for (i in 1:length(df_all)) {
  
  # Prophet
  m <- prophet(df_all[[i]], yearly.seasonality = TRUE,
               changepoint.range = 0.95,
               seasonality.mode = 'multiplicative',
               growth = 'linear',
               changepoint.prior.scale = 0.05,
               seasonality.prior.scale = 1
               )
  future <- make_future_dataframe(m, freq = 'month', periods = 24)
  future$floor <- floor
  future$cap <- cap
  forecast <- predict(m, future)
  forecast$yhat <- ceiling(forecast$yhat)
  
  # Add forecast to Sheets list
  sheets[[excel_sheets(usage_file)[i]]] <- 
    tail(forecast[, c('ds', 'yhat', 'yearly')], n = 24)
  
  # Add plot to Plots list
  plots[[excel_sheets(usage_file)[i]]] <- 
    dyplot.prophet(m, forecast, main = excel_sheets(usage_file)[i])
  
}

# Show plots
plots

# Write output file
write_xlsx(sheets, 'Prophet Forecasts.xlsx')

# End timer
toc()