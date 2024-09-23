######################### SCM Safety Stock Calculator ##########################
#                                                                              #
# Created 6/10/2022 by Alex McDonald                                           #
# Last updated 10/25/2022                                                      #
#                                                                              #
################################################################################

library(tidyverse)
library(RODBC)
library(readxl)
library(writexl)
library(stringr)
library(tictoc)

setwd(
  'A:/Shared drives/Forecasting/Planning & Analytics/Projects/2022/R/Safety Stock Calculator')

############################ Read in data - 1 min #############################

# Start timer
tic()

# Connect to COMM, run query, and store results in actuals data frame
db_conn <- odbcConnect('COMM')
if (db_conn == -1) {
  print('Connection failed')
} else {
  print('Connection successful')
}
query <- paste0(readLines(paste0('Usage(SS).sql')), collapse='\n')
daily_actuals <- sqlQuery(db_conn, query)
odbcCloseAll()
remove(db_conn)

# Stop timer
toc()

############################# Run SS calculation ###############################

# Start timer
tic()

# Specify part name
part_name <- 'DPP44'

# Specify confidence level
cl <- 0.975

# Specify lead time (days)
lt <- (7 * 68)

# Format data
df <- filter(daily_actuals, PART_GROUP2 == part_name)
df$QUANTITY <- df$QUANTITY * -1
df_final <- list()
for (i in 1:length(unique(df$ORG))) {
  df_final[[unique(df$ORG)[i]]] <- filter(df, ORG == unique(df$ORG)[i])
}

# Calculate safety stocks by ORG
SS <- list()
for (i in 1:length(df_final)) {
  
  daily_usage <- df_final[[i]]$QUANTITY
  
  if (length(daily_usage) < length(unique(df$TRANSACTION_DATE))) {
    daily_usage <- c(rep(0, 
                    length(unique(df$TRANSACTION_DATE)) - length(daily_usage)),
                     daily_usage)
  }
  
  SS[[names(df_final)[i]]] <- ceiling(qnorm(cl) * sd(daily_usage) * sqrt(lt))
  remove(daily_usage)
}

# Group safety stocks by IHS, Suwanee, Denver
IHS_SS_list <- list()
Suwanee_SS_list <- list()
Denver_SS_list <- list()
for (i in 1:length(SS)) {
  if (names(SS)[i] == 'S02' | names(SS)[i] == 'S09' | names(SS)[i] == 'M02') {
    Suwanee_SS_list[[names(SS)[i]]] <- SS[[i]]
  }
  else if (names(SS)[i] == 'S07' | names(SS)[i] == 'S08' |
           names(SS)[i] == 'M01') {
    Denver_SS_list[[names(SS)[i]]] <- SS[[i]]
  }
  else {
    IHS_SS_list[[names(SS)[i]]] <- SS[[i]]
  }
}

# Print results
cat(paste0('\n',
           'IHS SS = ', format(Reduce('+', IHS_SS_list), big.mark = ','), '\n', 
           'Suwanee SS = ', format(Reduce('+', Suwanee_SS_list), big.mark = ','), '\n',
           'Denver SS = ', format(Reduce('+', Denver_SS_list), big.mark = ','), '\n', 
           'Total SS = ', 
           format(sum(Reduce('+', IHS_SS_list), Reduce('+', Suwanee_SS_list), Reduce('+', Denver_SS_list)), big.mark = ','), '\n'),
           '\n')

# Write IHS output file
IHS <- do.call('rbind', lapply(IHS_SS_list, as.data.frame))
write.csv(IHS, 'IHS Safety Stock.csv')
 
# Stop timer
toc()