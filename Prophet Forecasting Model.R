######################## SCM Prophet Forecasting Model #########################
#                                                                              #
# Created 3/11/2022 by Alex McDonald                                           #
# Last updated 7/14/2022                                                       #
#                                                                              #
################################################################################

library(tidyverse)
library(prophet)
library(RODBC)
library(readxl)
library(writexl)
library(stringr)
library(tictoc)

setwd('A:/Shared drives/Forecasting/Planning & Analytics/Projects/2022/R/Forecasting/Master')


############################ Read in data - 10 min #############################

# Start timer
tic()

# Connect to COMM, run query, and store results in actuals data frame
db_conn <- odbcConnect('COMM')
if (db_conn == -1) {
  print('Connection failed')
} else {
  print('Connection successful')
}
query <- paste0(readLines(paste0('Usage.sql')), collapse='\n')
actuals <- sqlQuery(db_conn, query)
odbcCloseAll()
remove(db_conn)

# Stop timer
toc()


################################ Run forecasts #################################

# Start timer
tic()

# Read in Part List
part_list <- read_excel('Part List.xlsx', col_names = TRUE)

# Settings to consider for tuning
at_cp_ps <- c(0.01,0.02, 0.03, 0.04, 0.05)
at_seas_ps <- c(0.1, 1, 4, 7, 10)
param_grid <- data.frame(at_cp_ps, at_seas_ps)

# Create empty lists
df <- list()
sheets <- list()
plots <- list()

# Forecast engine
for (i in 1:nrow(part_list)) {
  
  if (part_list$`Run?`[i] == 'Y' | part_list$`Run?`[i] == 'y') {
    
    # Pull part name from list
    part_name <- part_list$`Part Name`[i]
    
    # Format and clean data
    df[[part_name]] <- filter(actuals, PART_GROUP2 == part_name)
    df[[part_name]] <- setNames(aggregate(df[[part_name]]$QUANTITY, 
                                          by = list(df[[part_name]]$WEEK_DATE), 
                                          FUN = sum), c('ds', 'y'))
    df[[part_name]] <- slice(df[[part_name]], 
                  (part_list$`Initial data points to exclude`[i] + 1):
                    (n() - part_list$`Latest data points to exclude`[i]))
    df[[part_name]]$y[
           as.Date(df[[part_name]]$ds) == as.Date(part_list$`Outlier 1`[i]) |
           as.Date(df[[part_name]]$ds) == as.Date(part_list$`Outlier 2`[i]) |
           as.Date(df[[part_name]]$ds) == as.Date(part_list$`Outlier 3`[i]) |
          as.Date(df[[part_name]]$ds) == as.Date(part_list$`Outlier 4`[i])] = NA
    df[[part_name]]$y <- df[[part_name]]$y * -1
    if (part_list$Growth[i] == 'logistic') {
      if (is.na(part_list$Floor[i]) & is.na(part_list$Cap[i])) {
        df[[part_name]]$floor <- min(df[[part_name]]$y, na.rm = TRUE) - 1
        df[[part_name]]$cap <- max(df[[part_name]]$y, na.rm = TRUE) + 1
      } else {
        df[[part_name]]$floor <- part_list$Floor[i]
        df[[part_name]]$cap <- part_list$Cap[i]
      }
    }

    # Hyperparameter tuning
    if (part_list$`Auto tune?`[i] == 'Y' | part_list$`Auto tune?`[i] == 'y') {
      all_param_comb <- expand.grid(param_grid)
      all_param_comb[, c('MAPE', 'RMSE')] <- NA
      at_horizon <- part_list$`Auto tuning horizon`[i]
      
      for (j in 1:nrow(all_param_comb)) {
        m_tuning <- prophet(df[[part_name]], yearly.seasonality = TRUE, 
                            changepoint.range = 0.95,
                            changepoint.prior.scale = all_param_comb[j, 1],
                            seasonality.prior.scale = all_param_comb[j, 2],
                            seasonality.mode = 'multiplicative',
                            growth = part_list$Growth[i])
        df_cv <- cross_validation(m_tuning, 
              cutoffs = df[[part_name]]$ds[nrow(df[[part_name]]) - at_horizon],
                                  initial = nrow(df[[part_name]]) - at_horizon,
                                  horizon = at_horizon,
                                  units = 'weeks')
        df_p <- performance_metrics(df_cv, rolling_window = 1)
        all_param_comb[j, 3] <- df_p[1, 'mape']
        all_param_comb[j, 4] <- df_p[1, 'rmse']
        
      }
      
      tuning_results <- all_param_comb[order(all_param_comb$MAPE),]
      print(paste0(part_name, ' TUNING RESULTS'))
      print(tuning_results)
      
      cp_ps <- tuning_results[1, 1]
      seas_ps <- tuning_results[1, 2]
    } else {
      cp_ps <- part_list$cp_ps[i]
      seas_ps <- part_list$seas_ps[i]
    }
    
    # Run forecast
    m <- prophet(df[[part_name]], yearly.seasonality = TRUE, 
                 changepoint.range = 0.95,
                 changepoint.prior.scale = cp_ps,
                 seasonality.prior.scale = seas_ps,
                 seasonality.mode = 'multiplicative',
                 growth = part_list$Growth[i])
    future <- make_future_dataframe(m, freq = 'week', 
                                    periods = part_list$`Forecast horizon`[i])
    if (part_list$Growth[i] == 'logistic') {
      if (is.na(part_list$Floor[i]) & is.na(part_list$Cap[i])) {
        future$floor <- min(df[[part_name]]$y, na.rm = TRUE) - 1
        future$cap <- max(df[[part_name]]$y, na.rm = TRUE) + 1
      } else {
        future$floor <- part_list$Floor[i]
        future$cap <- part_list$Cap[i]
      }
    }
    forecast <- predict(m, future)
    forecast$yhat <- ceiling(forecast$yhat)
    
    # Add forecast to Sheets list
    sheets[[str_replace_all(part_name, '/', ' ')]] <- 
                                     tail(forecast[, c('ds', 'yhat', 'yearly')], 
                                            n = part_list$`Forecast horizon`[i])
    
    # Add plot to Plots list
    plots[[part_name]] <- dyplot.prophet(m, forecast, main = paste0(part_name,
                                                           ' | cp_ps = ', cp_ps, 
                                                      ' | seas_ps = ', seas_ps))
    
  }
  
}

# Show plots
plots

# Write output file
write_xlsx(sheets, 'Prophet Forecasts.xlsx')

# Stop timer
toc()
