library(readxl)
library(dplyr)
library(lubridate)
library(ggplot2)
library(tidyverse)
library(tidyr)
library(writexl) 
library(lubridate)
library(openxlsx)
library(kableExtra)
library(gridExtra)
library(flextable)
library(officer)
library(psych)
library(DescTools)


##################################################### STARTING FROM END OF CARS SUMMARY STATISTICS #############################################

##################################### NEXT : REPORT SUMMARY STATS PER QUARTER, RELATIVE TO ANNOUNCEMENT ########################################

# HOW ARE MISSING VALUES HANDLED?

# Import balance sheet data and transaction_for_regression data
balance_sheet_data <- read_excel("./input/balance_sheet_data_Compustat.xlsx")
annual_balance_sheet_data <- read_excel("./input/annual_bs_data_Compustat.xlsx")
transactions_for_regression <- read_excel("./input/transactions_for_regression.xlsx")

balance_sheet_data$`Data Date` <- as.Date(balance_sheet_data$`Data Date`)
annual_balance_sheet_data$`Data Date` <- as.Date(annual_balance_sheet_data$`Data Date`)
transactions_for_regression$`Announced date` <- as.Date(transactions_for_regression$`Announced date`)
transactions_for_regression$`Completed date` <- as.Date(transactions_for_regression$`Completed date`)
summary(balance_sheet_data)
summary(annual_balance_sheet_data)
summary(transactions_for_regression)

################################## UNIVARIATE ANALYSES ########################################################################
################## by tercile

# 1) For each Acquiror ticker symbol - Deal Number group in transactions_for_regression, filter to
#    include only balance_sheet_data$Data Date corresponding to the 4 quarters preceding and 12
#    quarters following  transactions_for_regression$Announced date. 
# 2) Calculate new variables:
#         asset_turnover = sales/total_assets,
#         qoq_change_in_sales = (sales - lag(sales)),
#         yoy_change_in_sales = (sales - lag(sales, n=4)),
#         scaled_change_in_sales = change_in_sales/total_assets,
#         ln(mve) = ln(market_value),
#         mtb = market_value / book_value,
#         working_capital = (current_assets - cash_st_investments) - (current_liabilities - short_term_debt),
#         awca = working_capital - (lag4(working_capital)/lag4(sales))*sales,
#         scaled_AWCA = AWCA / lag4(total assets),
#         DUMMIES?
# 3) Winsorize the variables at the 5% and 95% levels.
# 4) Calculate summary statistics (Mean, st dev, 1st quar., median, 3rd quar.) on the variables.
# 5) Calculate t-statistic for means and Wilcoxon signed-rank test for medians.

##############################################################################################################
##############################################################################################################
######################################### QUARTERLY DATA ################################################
##############################################################################################################
##############################################################################################################

# Add a quarter identifier to the balance sheet data
balance_sheet_data <- balance_sheet_data %>%
  mutate(fiscal_quarter_id = year(`Data Date`) * 4 + quarter(`Data Date`))

# Add a quarter identifier to the transactions data 
# This step correctly assigns quarter identifiers to each transaction for serial acquirers.
transactions_for_regression <- transactions_for_regression %>%
  mutate(announcement_quarter_id = year(`Announced date`) * 4 + quarter(`Announced date`))

# Merge the datasets and filter for the periods Q[-2] to Q[+12], with Q[0] the quarter of announcement
# Many-to-many relationship expected because some firms engage in serial transactions, and in some quarters, multiple firms engage in transactions
merged_data <- balance_sheet_data %>%
  inner_join(transactions_for_regression, by = c("Ticker Symbol" = "Acquiror ticker symbol"), relationship = "many-to-many") %>%
  # Calculate the difference in quarters
  mutate(quarter_diff = fiscal_quarter_id - announcement_quarter_id) %>%
  # Filter for the desired time window around each announcement
  filter(quarter_diff >= -6 & quarter_diff <= 12) # >= -6 ensures there is sufficient data to calculate lagged variables in next step

# Write data to Excel to investigate NAs
# write_xlsx(merged_data, "./output/merged_data_with_NAs.xlsx")

# Calculate new variables
merged_data <- merged_data%>%
  group_by(`Deal Number`, CAR_Tercile)%>%
  mutate(
    asset_turnover = sales/total_assets,
#    qoq_change_in_sales = (sales - lag(sales)),
    seasonally_adjusted_change_in_sales = (sales - lag(sales, n=4)),
#    scaled_qoq_change_in_sales = qoq_change_in_sales/lag(total_assets), # Absolute change in sales from the previous quarter, as a proportion of assets last quarter
    scaled_adj_change_in_sales = seasonally_adjusted_change_in_sales/lag(total_assets, n=4), # Absolute change in sales over the same quarter last year, as a proportion of assets in the same quarter the previous year
    ln_mve = log(market_value),
    mtb = market_value / book_value,
    working_capital = (current_assets - cash_st_investments) - (current_liabilities - short_term_debt),
    awca = working_capital - (lag(working_capital, n=4)/lag(sales, n=4))*sales,
    scaled_awca = awca / lag(total_assets, n=4)
  )%>%
  ungroup()


temp_winsorized <- merged_data %>%
  filter(!is.na(total_assets), !is.na(asset_turnover),   # filter out NA values
              !is.na(seasonally_adjusted_change_in_sales),
              !is.na(scaled_adj_change_in_sales), !is.na(ln_mve),
              !is.na(mtb), !is.na(awca)) %>%
  mutate(
    total_assets_winsorized = Winsorize(total_assets, probs = c(0.05, 0.95), na.rm = TRUE),
    asset_turnover_winsorized = Winsorize(asset_turnover, probs = c(0.05, 0.95), na.rm = TRUE),
#    qoq_change_in_sales_winsorized = Winsorize(qoq_change_in_sales, probs = c(0.05, 0.95), na.rm = TRUE), # if this is added back, check NA values and add to filter above if necessary
    seasonally_adjusted_change_in_sales_winsorized = Winsorize(seasonally_adjusted_change_in_sales, probs = c(0.05, 0.95), na.rm = TRUE),
#    scaled_qoq_change_in_sales_winsorized = Winsorize(scaled_qoq_change_in_sales, probs = c(0.05, 0.95), na.rm = TRUE),
    scaled_adj_change_in_sales_winsorized = Winsorize(scaled_adj_change_in_sales, probs = c(0.05, 0.95), na.rm = TRUE),
    ln_mve_winsorized = Winsorize(ln_mve, probs = c(0.05, 0.95), na.rm = TRUE),
    mtb_winsorized = Winsorize(mtb, probs = c(0.05, 0.95), na.rm = TRUE),
    awca_winsorized = Winsorize(awca, probs = c(0.05, 0.95), na.rm = TRUE)
  ) %>%
  ungroup()

# Join the Winsorized data back to the original dataset
merged_data <- merged_data %>%
  left_join(temp_winsorized %>%
              select(`Deal Number`, quarter_diff, total_assets_winsorized, asset_turnover_winsorized,
                     seasonally_adjusted_change_in_sales_winsorized,
                     scaled_adj_change_in_sales_winsorized,
                     ln_mve_winsorized, mtb_winsorized, awca_winsorized),
            by = c("Deal Number", "quarter_diff"))

# Calculate summary statistics for the winsorized variables
summary_stats <- merged_data %>%
  filter(quarter_diff >= -2 & quarter_diff <= 12) %>%
  group_by(quarter_diff, CAR_Tercile) %>%
  summarize(
    n = n_distinct(`Deal Number`),
    Mean_total_assets = mean(total_assets_winsorized, na.rm = TRUE),
    SD_total_assets = sd(total_assets_winsorized, na.rm = TRUE),
    First_quartile_total_assets = quantile(total_assets_winsorized, 0.25, na.rm = TRUE),
    Median_total_assets = median(total_assets_winsorized, na.rm = TRUE),
    Third_quartile_total_assets = quantile(total_assets_winsorized, 0.75, na.rm = TRUE),
    Mean_asset_turnover = mean(asset_turnover_winsorized, na.rm = TRUE),
    SD_asset_turnover = sd(asset_turnover_winsorized, na.rm = TRUE),
    First_quartile_asset_turnover = quantile(asset_turnover_winsorized, 0.25, na.rm = TRUE),
    Median_asset_turnover = median(asset_turnover_winsorized, na.rm = TRUE),
    Third_quartile_asset_turnover = mean(asset_turnover_winsorized, 0.75, na.rm = TRUE),
#    Mean_qoq_sales = mean(qoq_change_in_sales_winsorized, na.rm = TRUE),
#    SD_qoq_sales = sd(qoq_change_in_sales_winsorized, na.rm = TRUE),
#    First_quartile_qoq_sales = quantile(qoq_change_in_sales_winsorized, 0.25, na.rm = TRUE),
#    Median_qoq_sales = median(qoq_change_in_sales_winsorized, na.rm = TRUE),
#    Third_quartile_qoq_sales = quantile(qoq_change_in_sales_winsorized, 0.75, na.rm = TRUE),
    Mean_seasonally_adjusted_change_in_sales = mean(seasonally_adjusted_change_in_sales_winsorized, na.rm = TRUE),
    SD_seasonally_adjusted_sales = sd(seasonally_adjusted_change_in_sales_winsorized, na.rm = TRUE),
    First_quartile_seasonally_adjusted_sales = quantile(seasonally_adjusted_change_in_sales_winsorized, 0.25, na.rm = TRUE),
    Median_seasonally_adjusted_sales = median(seasonally_adjusted_change_in_sales_winsorized, na.rm = TRUE),
    Third_quartile_seasonally_adjusted_sales = quantile(seasonally_adjusted_change_in_sales_winsorized, 0.75, na.rm = TRUE),
#    Mean_scaled_qoq_sales = mean(scaled_qoq_change_in_sales_winsorized, na.rm = TRUE),
#    SD_scaled_qoq_sales = sd(scaled_qoq_change_in_sales_winsorized, na.rm = TRUE),
#    First_quartile_scaled_qoq_sales = quantile(scaled_qoq_change_in_sales_winsorized, 0.25, na.rm = TRUE),
#    Median_scaled_qoq_sales = median(scaled_qoq_change_in_sales_winsorized, na.rm = TRUE),
#    Third_quartile_scaled_qoq_sales = quantile(scaled_qoq_change_in_sales_winsorized, 0.75, na.rm = TRUE),
    Mean_scaled_yoy_sales = mean(scaled_adj_change_in_sales_winsorized, na.rm = TRUE),
    SD_scaled_yoy_sales = sd(scaled_adj_change_in_sales_winsorized, na.rm = TRUE),
    First_quartile_scaled_yoy_sales = quantile(scaled_adj_change_in_sales_winsorized, 0.25, na.rm = TRUE),
    Median_scaled_yoy_sales = median(scaled_adj_change_in_sales_winsorized, na.rm = TRUE),
    Third_quartile_scaled_yoy_sales = quantile(scaled_adj_change_in_sales_winsorized, 0.75, na.rm = TRUE),
    Mean_ln_mve = mean(ln_mve_winsorized, na.rm = TRUE),
    SD_ln_mve = sd(ln_mve_winsorized, na.rm = TRUE),
    First_quartile_ln_mve = quantile(ln_mve_winsorized, 0.25, na.rm = TRUE),
    Median_ln_mve = median(ln_mve_winsorized, na.rm = TRUE),
    Third_quartile_ln_mve = quantile(ln_mve_winsorized, 0.75, na.rm = TRUE)
  )%>%
#  mutate(quarters_from_announcement = quarter_diff)%>%
  ungroup()

####################################### check above ^ fewer than 100 transactions are being taken into account !

# Transpose summary_stats data to have quarters as column headers and stats as rows
transposed_summary_stats <- summary_stats %>%
  pivot_longer(
    cols = -quarter_diff, # Exclude the quarter_diff column from transposing
    names_to = "Statistic",
    values_to = "Value"
  ) %>%
  pivot_wider(
    names_from = "quarter_diff",
    values_from = "Value"
  )

write_xlsx(transposed_summary_stats, "./output/summary_stats_quarterly_transposed.xlsx")

# Create a table to be edited in Word
summary_stats_table <- flextable(transposed_summary_stats)

summary_stats_table <- summary_stats_table %>%
  theme_apa()%>%
  bold(part = "header")

doc <- read_docx() %>%
  body_add_flextable(value = summary_stats_table) %>%
  body_add_par("Summary Statistics Table", style = "heading 1")  # Adds a heading above the table

print(doc, target = "./output/Summary_Statistics_Table_transposed.docx")  # Save the Word document


##############################################################################################################
##############################################################################################################
######################################### ANNUAL DATA ################################################
##############################################################################################################
##############################################################################################################

# Step 1: Extract the M&A announcement year for each firm
ma_announcement_years <- transactions_for_regression %>%
  group_by(firm_id) %>%
  summarise(MA_Announcement_Year = min(year(`Announced date`))) %>%
  ungroup()

# Step 2: Join the M&A announcement year to the balance sheet data
annual_balance_sheet_data <- annual_balance_sheet_data %>%
  left_join(ma_announcement_years, by = "firm_id") %>%
  mutate(year_id = year(`Data Date`) - MA_Announcement_Year)

# Merge the datasets and filter for the periods Q[-2] to Q[+12], with Q[0] the quarter of announcement
# Many-to-many relationship expected because some firms engage in serial transactions, and in some quarters, multiple firms engage in transactions
merged_data <- balance_sheet_data %>%
  inner_join(transactions_for_regression, by = c("Ticker Symbol" = "Acquiror ticker symbol"), relationship = "many-to-many") %>%
  # Calculate the difference in quarters
  mutate(quarter_diff = fiscal_quarter_id - announcement_quarter_id) %>%
  # Filter for the desired time window around each announcement
  filter(quarter_diff >= -6 & quarter_diff <= 12) # >= -6 ensures there is sufficient data to calculate lagged variables in next step

######################################################################################################################################################
######################################################################################################################################################
#################################### THIS WOULD BE A GOOD PLACE TO WRITE ##############################################################################
#################################### MERGED_DATA TO EXCEL TO INVESTIGATE ##############################################################################
################################################# NA VALUES  ######################################################################################
######################################################################################################################################################

write_xlsx(merged_data, "./output/merged_data_with_NAs.xlsx")

# Calculate new variables
merged_data <- merged_data%>%
  group_by(`Deal Number`)%>%
  mutate(
    asset_turnover = sales/total_assets,
    qoq_change_in_sales = (sales - lag(sales)),
    yoy_change_in_sales = (sales - lag(sales, n=4)),
    scaled_qoq_change_in_sales = qoq_change_in_sales/lag(total_assets), # Absolute change in sales from the previous quarter, as a proportion of assets in this quarter last year
    scaled_yoy_change_in_sales = yoy_change_in_sales/lag(total_assets), # Absolute change in sales from the same quarter last year, as a proportion of assets in this quarter last year
    ln_mve = log(market_value),
    mtb = market_value / book_value,
    working_capital = (current_assets - cash_st_investments) - (current_liabilities - short_term_debt),
    awca = working_capital - (lag(working_capital, n=4)/lag(sales, n=4))*sales,
    scaled_awca = awca / lag(total_assets, n=4)
  )%>%
  ungroup()


temp_winsorized <- merged_data %>%
  filter(quarter_diff >= -2 & quarter_diff <= 12) %>%
  group_by(`Deal Number`) %>%
  # Only apply Winsorize if there are no NA values in the columns of interest
  filter(!any(is.na(total_assets), is.na(asset_turnover), is.na(qoq_change_in_sales),
              is.na(yoy_change_in_sales), is.na(scaled_qoq_change_in_sales),
              is.na(scaled_yoy_change_in_sales), is.na(ln_mve), is.na(mtb), is.na(awca))) %>%
  mutate(
    total_assets_winsorized = Winsorize(total_assets, probs = c(0.05, 0.95), na.rm = TRUE),
    asset_turnover_winsorized = Winsorize(asset_turnover, probs = c(0.05, 0.95), na.rm = TRUE),
    qoq_change_in_sales_winsorized = Winsorize(qoq_change_in_sales, probs = c(0.05, 0.95), na.rm = TRUE),
    yoy_change_in_sales_winsorized = Winsorize(yoy_change_in_sales, probs = c(0.05, 0.95), na.rm = TRUE),
    scaled_qoq_change_in_sales_winsorized = Winsorize(scaled_qoq_change_in_sales, probs = c(0.05, 0.95), na.rm = TRUE),
    scaled_yoy_change_in_sales_winsorized = Winsorize(scaled_yoy_change_in_sales, probs = c(0.05, 0.95), na.rm = TRUE),
    ln_mve_winsorized = Winsorize(ln_mve, probs = c(0.05, 0.95), na.rm = TRUE),
    mtb_winsorized = Winsorize(mtb, probs = c(0.05, 0.95), na.rm = TRUE),
    awca_winsorized = Winsorize(awca, probs = c(0.05, 0.95), na.rm = TRUE)
  ) %>%
  ungroup()

# Join the Winsorized data back to the original dataset
merged_data <- merged_data %>%
  left_join(temp_winsorized %>%
              select(`Deal Number`, quarter_diff, total_assets_winsorized, asset_turnover_winsorized, 
                     qoq_change_in_sales_winsorized, yoy_change_in_sales_winsorized,
                     scaled_qoq_change_in_sales_winsorized, scaled_yoy_change_in_sales_winsorized,
                     ln_mve_winsorized, mtb_winsorized, awca_winsorized),
            by = c("Deal Number", "quarter_diff"))

# Calculate summary statistics for the winsorized variables
summary_stats <- merged_data %>%
  group_by(`Deal Number`, quarter_diff)%>%
  summarise(
    Mean_total_assets = mean(total_assets_winsorized, na.rm = TRUE),
    SD_total_assets = sd(total_assets_winsorized, na.rm = TRUE),
    First_quartile_total_assets = quantile(total_assets_winsorized, 0.25, na.rm = TRUE),
    Median_total_assets = median(total_assets_winsorized, na.rm = TRUE),
    Third_quartile_total_assets = quantile(total_assets_winsorized, 0.75, na.rm = TRUE),
    Mean_asset_turnover = mean(asset_turnover_winsorized, na.rm = TRUE),
    SD_asset_turnover = sd(asset_turnover_winsorized, na.rm = TRUE),
    First_quartile_asset_turnover = quantile(asset_turnover_winsorized, 0.25, na.rm = TRUE),
    Median_asset_turnover = median(asset_turnover_winsorized, na.rm = TRUE),
    Third_quartile_asset_turnover = mean(asset_turnover_winsorized, 0.75, na.rm = TRUE),
    Mean_qoq_sales = mean(qoq_change_in_sales_winsorized, na.rm = TRUE),
    SD_qoq_sales = sd(qoq_change_in_sales_winsorized, na.rm = TRUE),
    First_quartile_qoq_sales = quantile(qoq_change_in_sales_winsorized, 0.25, na.rm = TRUE),
    Median_qoq_sales = median(qoq_change_in_sales_winsorized, na.rm = TRUE),
    Third_quartile_qoq_sales = quantile(qoq_change_in_sales_winsorized, 0.75, na.rm = TRUE),
    Mean_yoy_sales = mean(yoy_change_in_sales_winsorized, na.rm = TRUE),
    SD_yoy_sales = sd(yoy_change_in_sales_winsorized, na.rm = TRUE),
    First_quartile_yoy_sales = quantile(yoy_change_in_sales_winsorized, 0.25, na.rm = TRUE),
    Median_yoy_sales = median(yoy_change_in_sales_winsorized, na.rm = TRUE),
    Third_quartile_yoy_sales = quantile(yoy_change_in_sales_winsorized, 0.75, na.rm = TRUE),
    Mean_scaled_qoq_sales = mean(scaled_qoq_change_in_sales_winsorized, na.rm = TRUE),
    SD_scaled_qoq_sales = sd(scaled_qoq_change_in_sales_winsorized, na.rm = TRUE),
    First_quartile_scaled_qoq_sales = quantile(scaled_qoq_change_in_sales_winsorized, 0.25, na.rm = TRUE),
    Median_scaled_qoq_sales = median(scaled_qoq_change_in_sales_winsorized, na.rm = TRUE),
    Third_quartile_scaled_qoq_sales = quantile(scaled_qoq_change_in_sales_winsorized, 0.75, na.rm = TRUE),
    Mean_scaled_yoy_sales = mean(scaled_yoy_change_in_sales_winsorized, na.rm = TRUE),
    SD_scaled_yoy_sales = sd(scaled_yoy_change_in_sales_winsorized, na.rm = TRUE),
    First_quartile_scaled_yoy_sales = quantile(scaled_yoy_change_in_sales_winsorized, 0.25, na.rm = TRUE),
    Median_scaled_yoy_sales = median(scaled_yoy_change_in_sales_winsorized, na.rm = TRUE),
    Third_quartile_scaled_yoy_sales = quantile(scaled_yoy_change_in_sales_winsorized, 0.75, na.rm = TRUE),
    Mean_ln_mve = mean(ln_mve_winsorized, na.rm = TRUE),
    SD_ln_mve = sd(ln_mve_winsorized, na.rm = TRUE),
    First_quartile_ln_mve = quantile(ln_mve_winsorized, 0.25, na.rm = TRUE),
    Median_ln_mve = median(ln_mve_winsorized, na.rm = TRUE),
    Third_quartile_ln_mve = quantile(ln_mve_winsorized, 0.75, na.rm = TRUE)
  )%>%
  mutate(quarters_from_announcement = quarter_diff)%>%
  ungroup()

# Create a table to be edited in Word
summary_stats_table <- flextable(summary_stats)

summary_stats_table <- summary_stats_table %>%
  set_table_properties(layout = "autofit")%>%
  theme_apa()%>%
  bold(part = "header")%>%
  autofit()

doc <- read_docx() %>%
  body_add_flextable(value = summary_stats_table) %>%
  body_add_par("Summary Statistics Table", style = "heading 1")  # Adds a heading above the table

print(doc, target = "Summary_Statistics_Table.docx")  # Save the Word document


########################################################### ROBUSTNESS CHECKS #####################################################################################


# End of script