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
library(stats)
library(car)
library(ggpubr)
library(broom)


############### STARTING FROM END OF CARS SUMMARY STATISTICS ###################

######### REPORT SUMMARY STATS PER YEAR, RELATIVE TO ANNOUNCEMENT ##############

# Start from "cleaned_annual_balance_sheet_data" ! 

# Import balance sheet data and transaction_for_regression data
#annual_balance_sheet_data <- read_excel("./input/annual_balance_sheet_data2_Compustat.xlsx")
transactions_for_regression <- read_excel("./output/transactions_for_analysis.xlsx") 

# balance_sheet_data$`Data Date` <- as.Date(balance_sheet_data$`Data Date`) # FOR QUARTERLY DATA
#annual_balance_sheet_data$`Data Date` <- as.Date(annual_balance_sheet_data$`Data Date`)
#transactions_for_regression$`Announced date` <- as.Date(transactions_for_regression$`Announced date`)

#annual_balance_sheet_data$fiscal_year <- annual_balance_sheet_data$`Data Year - Fiscal` # rename column
#annual_balance_sheet_data <- annual_balance_sheet_data %>%
#  select(-`Global Company Key`, -`Data Date`, -`Data Year - Fiscal`) # remove unnecessary columns

#transactions_for_regression$`Acquiror ISIN number` <- transactions_for_regression$`Acquiror ISIN number.x` # clean up df
#transactions_for_regression$`Acquiror ticker symbol` <- transactions_for_regression$`Acquiror ticker symbol.x`
#transactions_for_regression <- transactions_for_regression %>%
#  select(`Deal Number`, `Acquiror ISIN number`, `Acquiror ticker symbol`,`Acquiror name`, `Announced date`, announcement_year,
#         CAR_Tercile, industry_relatedness, industry, serial, stock_deal,
#         cross_border, is_Big4)


#summary(annual_balance_sheet_data)
#summary(transactions_for_regression)

#options(scipen = 999)


# Count unique Acquiror ISIN numbers in annual_balance_sheet_data
#number_of_unique_acquiror_isins <- n_distinct(annual_balance_sheet_data$`Ticker Symbol`)
#number_of_unique_acquiror_isins # 458 expected

# Count Deal Numbers in transactions_for_regression
#number_of_deals <- n_distinct(transactions_for_regression$`Deal Number`)
#number_of_deals
#number_of_acquirors <- n_distinct(transactions_for_regression$`Acquiror ISIN number`)
#number_of_acquirors



################################## UNIVARIATE ANALYSES ########################################################################
################## by tercile

# 1) For each Acquiror ticker symbol - Deal Number group in transactions_for_regression, filter to
#    include only balance_sheet_data$Data Date corresponding to the 2 years preceding and 5
#    years following  transactions_for_regression$Announced date. 
# 2) Calculate new variables:
#         asset_turnover = sales/total_assets,
#         qoq_change_in_sales = (sales - lag(sales)),
#         yoy_change_in_sales = (sales - lag(sales, n=4)),
#         scaled_change_in_sales = change_in_sales/total_assets,
#         ln(mve) = ln(market_value),
#         mtb = market_value / book_value,
#         working_capital = (current_assets - cash_st_investments) - (current_liabilities - short_term_debt),
#         awca = working_capital - (lag4(working_capital)/lag4(sales))*sales,
#         scaled_AWCA = AWCA / lag4(total assets)
# 3) Winsorize the variables at the 5% and 95% levels.
# 4) Calculate summary statistics (Mean, st dev, 1st quar., median, 3rd quar.) on the variables.
# 5) Calculate t-statistic for means and Wilcoxon signed-rank test for medians.
# 6) Dispersion of means/medians: 
#       a. paired samples t-tests to compare same tercile over different years (each year compared to year -1, for example)
#           - Medians: sign test / signed-rank test
#       b. independent sample t-test to compare groups (in same year, central tendency of Tercile 1 and Tercile 3)
#           - Medians: Mann-Whitney-Wilcoxon (if distributions of groups have similar shapes)
#       c. Boxplots
# 7) Assess normality of variables:
#       a. Kolmogorov-Smirnov
#       b. Shapiro-Wilk
#       c. Q-Q plots
# 8) Differences of central tendency across all groups: ANOVA / Kruskal-Wallis
#     - Assumptions for ANOVA: 
#       a. model errors are normally distributed (Kolmogorov-Smirnov, Shapiro-Wilk on model errors)
#       b. error variances are equal across groups (Levene's test)
#     - If assumptions are violated, Kruskal-Wallis is more reliable (assuming distribution of groups are similar)




##############################################################################################################
##############################################################################################################
######################################### ANNUAL DATA ################################################
##############################################################################################################
##############################################################################################################


# Extract the M&A announcement year for each firm and create a new dataframe
# which includes the date range for each deal
#transactions_for_regression <- transactions_for_regression %>%
#  mutate(
#    start_year = announcement_year - 2,
#    end_year = announcement_year + 5
#  )

# Semi-join to filter balance_sheet_data
#filtered_balance_sheet_data <- annual_balance_sheet_data %>%
#  left_join(transactions_for_regression, by = c(`Ticker Symbol` = "Acquiror ticker symbol" )) %>%
#  filter(fiscal_year >= start_year & fiscal_year <= end_year) %>%
#  mutate(years_from_announcement = fiscal_year - announcement_year) %>%
#  select(-start_year, -end_year)

# Count unique Deal Numbers
#n_distinct(filtered_balance_sheet_data$`Deal Number`)

# Count unique Acquiror ISIN numbers
#n_distinct(filtered_balance_sheet_data$`Acquiror ISIN number`)



#write_xlsx(filtered_balance_sheet_data, "./output/annual_bs_to_check_missing_values.xlsx")


#####################################################################################################################
################## TO AVOID REDUCING SAMPLE SIZE HERE, ##############################################################
###############   NA VALUES HAVE BEEN FILLED IN WHERE POSSIBLE #################################################
#####################################################################################################################
#####################################################################################################################
######################### START HERE ################################################################################
#####################################################################################################################
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
library(stats)
library(car)
library(ggpubr)
library(broom)
library(robustHD)
library(nortest)
library(purrr)

options(digits=4) 

# Import manually cleaned annual bs data
cleaned_annual_balance_sheet_data <- read_excel("./input/annual_bs_cleaned.xlsx")
transactions_for_analysis <- read_excel("./output/transactions_for_analysis.xlsx")

# Remove duplicate rows in winsorized df
cleaned_annual_balance_sheet_data <- cleaned_annual_balance_sheet_data %>%
  distinct()

# Retain only Window == [-1,+1]
cleaned_annual_balance_sheet_data <- cleaned_annual_balance_sheet_data %>%
  filter(Window == "[-1,+1]")

# Remove deals with 1 or 2 years missing
# List of Deal Numbers to be removed
deal_numbers_to_remove <- c(1601106346, 1601238103, 1601264581, 1909095461, 1909135551,
                            1909262216, 1909301680, 1909542668, 1909566989, 1941025551,
                            1941205115, 1601193258, 1601238504, 1601246469, 1601421745,
                            1633013996, 1909048527, 1909073908, 1909103603, 1909260332,
                            1941148217, 1941172470)

# Filter out these Deal Numbers from the dataframe
cleaned_annual_balance_sheet_data <- cleaned_annual_balance_sheet_data %>%
  filter(!(`Deal Number` %in% deal_numbers_to_remove), Window == "[-1,+1]")

# Update CAR_winsorized with current data from CARs.R
cleaned_annual_balance_sheet_data <- cleaned_annual_balance_sheet_data %>%
  select(-CAR_winsorized, -is_Big4) %>%
  left_join(select(transactions_for_analysis, `Deal Number`, CAR_winsorized, is_Big4,
                   Shares_percentage, Cash_percentage), by = "Deal Number")

# Variables to test 
variables_to_test <- c(
  "scaled_awca_winsorized",
  "ln_mve_winsorized", # regression model uses lagged ln_mve
  "scaled_yoy_change_in_sales_winsorized",
  "mtb_winsorized", # regression uses lagged mtb
  "roa_winsorized" 

)

# Count remaining Deal Numbers
n_distinct(cleaned_annual_balance_sheet_data$`Deal Number`)

# Count unique Acquiror ISIN numbers
n_distinct(cleaned_annual_balance_sheet_data$`Acquiror ISIN number`)


# Calculate new variables 
cleaned_annual_balance_sheet_data <- cleaned_annual_balance_sheet_data%>%
  arrange(`Deal Number`, fiscal_year) %>%
  group_by(`Deal Number`, Window) %>%
  mutate(
    lagged_total_assets = lag(total_assets),
    asset_turnover = sales/lagged_total_assets,
    lagged_sales = lag(sales),
    yoy_change_in_sales = sales - lagged_sales,
    scaled_yoy_change_in_sales = yoy_change_in_sales/lagged_total_assets, # Absolute change in sales from the same quarter last year, as a proportion of assets in this quarter last year
    ln_mve = log(market_value),
    mtb = market_value / book_value,
    working_capital = (current_assets - cash_st_investments) - (current_liabilities - short_term_debt),
    lagged_working_capital = lag(working_capital),
    awca = working_capital - (lagged_working_capital/lagged_sales)*sales,
    scaled_awca = awca / lagged_total_assets,
    roa = net_income / total_assets
  )%>%
  ungroup()

# Check lagged variables
lagged_vars <- cleaned_annual_balance_sheet_data %>%
  filter(Window == "[-1,+1]", years_from_announcement>=-2 & years_from_announcement <= 5) %>%
  select(`Deal Number`, total_assets, lagged_total_assets, sales, lagged_sales, working_capital, lagged_working_capital)

# Check working capital
working_capital_check <- cleaned_annual_balance_sheet_data %>%
  filter(Window == "[-1,+1]", years_from_announcement>=-2 & years_from_announcement <= 5) %>%
  select(`Deal Number`, current_assets, cash_st_investments, current_liabilities, short_term_debt, working_capital)

# Winsorize variables of interest
winsorized_balance_sheet_data <- cleaned_annual_balance_sheet_data %>% 
  filter(Window == "[-1,+1]", years_from_announcement>=-2 & years_from_announcement <= 5) %>% # winsorized from year -2 to avoid problems with lagged vars later
  group_by(years_from_announcement) %>% # variables winsorized per year from announcement
  mutate(
    scaled_awca_winsorized = Winsorize(scaled_awca, probs = c(0.05, 0.95), na.rm = TRUE), # na.rm = TRUE because lagged vars cause NAs in year -2
    ln_mve_winsorized = Winsorize(ln_mve, probs = c(0.05, 0.95), na.rm = TRUE),
    scaled_yoy_change_in_sales_winsorized = Winsorize(scaled_yoy_change_in_sales, probs = c(0.05, 0.95), na.rm = TRUE),
    mtb_winsorized = Winsorize(mtb, probs = c(0.05, 0.95), na.rm = TRUE),
    roa_winsorized = Winsorize(roa, probs = c(0.05, 0.95), na.rm = TRUE),
    mve_winsorized = Winsorize(market_value, probs = c(0.05, 0.95), na.rm = TRUE) # summary stats should be mve (not logged value)
  ) %>%
  ungroup()

# Add AWCA in year (-1)
  awca_year_before_announcement <- winsorized_balance_sheet_data %>%
    filter(Window == "[-1,+1]", years_from_announcement == -1) %>%
    select(`Deal Number`, awca_year_before_announcement = scaled_awca_winsorized)

# Left join this helper dataframe back to the original dataframe to spread the awca value across all rows of each deal
  winsorized_balance_sheet_data <- winsorized_balance_sheet_data %>%
    left_join(awca_year_before_announcement, by = "Deal Number")


# Ensure CAR_Tercile is a factor
winsorized_balance_sheet_data$CAR_Tercile <- as.factor(winsorized_balance_sheet_data$CAR_Tercile)

# Write to .xlsx for further processing ################################################# correlations, regression in 
#write_xlsx(winsorized_balance_sheet_data, "./output/winsorized_annual_balance_sheet_data.xlsx") # SPSS/Excel from this point


# Calculate summary statistics for the winsorized variables ########################## full sample

full_sample_summary_stats <- winsorized_balance_sheet_data %>%
  group_by(years_from_announcement) %>%
  filter(years_from_announcement >= -1 & years_from_announcement <= 2, Window == "[-1,+1]") %>%
  summarise(
    Transactions_n = n_distinct(`Deal Number`),
    Acquirors_n = n_distinct(`Acquiror ISIN number`),
    
    Mean_scaled_awca_winsorized = mean(scaled_awca_winsorized, na.rm = TRUE),
    t_stat_awca = t.test(scaled_awca_winsorized, mu = 0, alternative = "two.sided", conf.level = 0.95, na.action = na.exclude)$statistic,
    p_value_t_test_awca = t.test(scaled_awca_winsorized, mu = 0, alternative = "two.sided", conf.level = 0.95, na.action = na.exclude)$p.value,
    SD_scaled_awca_winsorized = sd(scaled_awca_winsorized, na.rm = TRUE),
    First_quartile_scaled_awca_winsorized = quantile(scaled_awca_winsorized, 0.25, na.rm = TRUE),
    Median_scaled_awca_winsorized = median(scaled_awca_winsorized, na.rm = TRUE),
    W_stat_awca = wilcox.test(scaled_awca_winsorized, mu = 0, alternative = "two.sided", conf.int = FALSE, conf.level = 0.95, exact = FALSE, correct = TRUE, na.action = na.exclude)$statistic,
    p_value_Wilcox_awca = wilcox.test(scaled_awca_winsorized, mu = 0, alternative = "two.sided", conf.int = FALSE, conf.level = 0.95, exact = FALSE, correct = TRUE, na.action = na.exclude)$p.value,
    Third_quartile_scaled_awca_winsorized = quantile(scaled_awca_winsorized, 0.75, na.rm = TRUE),
    
    Mean_mve_winsorized = mean(mve_winsorized, na.rm = TRUE),
    SD_mve_winsorized = sd(mve_winsorized, na.rm = TRUE),
    First_quartile_mve_winsorized = quantile(mve_winsorized, 0.25, na.rm = TRUE),
    Median_mve_winsorized = median(mve_winsorized, na.rm = TRUE),
    Third_quartile_mve_winsorized = quantile(mve_winsorized, 0.75, na.rm = TRUE),
    
    Mean_scaled_yoy_sales_wins = mean(scaled_yoy_change_in_sales_winsorized, na.rm = TRUE),
    SD_scaled_yoy_sales_wins = sd(scaled_yoy_change_in_sales_winsorized, na.rm = TRUE),
    First_quartile_scaled_yoy_sales_wins = quantile(scaled_yoy_change_in_sales_winsorized, 0.25, na.rm = TRUE),
    Median_scaled_yoy_sales_wins = median(scaled_yoy_change_in_sales_winsorized, na.rm = TRUE),
    Third_quartile_scaled_yoy_sales_wins = quantile(scaled_yoy_change_in_sales_winsorized, 0.75, na.rm = TRUE),
    
    Mean_roa_winsorized = mean(roa_winsorized, na.rm = TRUE),
    SD_roa_winsorized = sd(roa_winsorized, na.rm = TRUE),
    First_quartile_roa_winsorized = quantile(roa_winsorized, 0.25, na.rm = TRUE),
    Median_roa_winsorized = median(roa_winsorized, na.rm = TRUE),
    Third_quartile_roa_winsorized = quantile(roa_winsorized, 0.75, na.rm = TRUE),
    
    Mean_mtb_winsorized = mean(mtb_winsorized, na.rm = TRUE),
    SD_mtb_winsorized = sd(mtb_winsorized, na.rm = TRUE),
    First_quartile_mtb_winsorized = quantile(mtb_winsorized, 0.25, na.rm = TRUE),
    Median_mtb_winsorized = median(mtb_winsorized, na.rm = TRUE),
    Third_quartile_mtb_winsorized = quantile(mtb_winsorized, 0.75, na.rm = TRUE),
    
    Tercile = "Full sample",
    .groups = 'drop'
  )

# Calculate summary statistics for the winsorized variables per tercile

tercile_summary_stats <- winsorized_balance_sheet_data %>%
  group_by(CAR_Tercile, years_from_announcement) %>%
  filter(years_from_announcement>=-1 & years_from_announcement <=2, Window == "[-1,+1]")%>%
  summarise(
    Transactions_n = n_distinct(`Deal Number`),
    Acquirors_n = n_distinct(`Acquiror ISIN number`),
    
    Mean_scaled_awca_winsorized = mean(scaled_awca_winsorized, na.rm = TRUE),
    t_stat_awca = t.test(scaled_awca_winsorized, mu = 0, alternative = "two.sided", conf.level = 0.95, na.action = na.exclude)$statistic,
    p_value_t_test_awca = t.test(scaled_awca_winsorized, mu = 0, alternative = "two.sided", conf.level = 0.95, na.action = na.exclude)$p.value,
    SD_scaled_awca_winsorized = sd(scaled_awca_winsorized, na.rm = TRUE),
    First_quartile_scaled_awca_winsorized = quantile(scaled_awca_winsorized, 0.25, na.rm = TRUE),
    Median_scaled_awca_winsorized = median(scaled_awca_winsorized, na.rm = TRUE),
    W_stat_awca = wilcox.test(scaled_awca_winsorized, mu = 0, alternative = "two.sided", conf.int = FALSE, conf.level = 0.95, exact = FALSE, correct = TRUE, na.action = na.exclude)$statistic,
    p_value_Wilcox_awca = wilcox.test(scaled_awca_winsorized, mu = 0, alternative = "two.sided", conf.int = FALSE, conf.level = 0.95, exact = FALSE, correct = TRUE, na.action = na.exclude)$p.value,
    Third_quartile_scaled_awca_winsorized = quantile(scaled_awca_winsorized, 0.75, na.rm = TRUE),
    
    Mean_mve_winsorized = mean(mve_winsorized, na.rm = TRUE),
    SD_mve_winsorized = sd(mve_winsorized, na.rm = TRUE),
    First_quartile_mve_winsorized = quantile(mve_winsorized, 0.25, na.rm = TRUE),
    Median_mve_winsorized = median(mve_winsorized, na.rm = TRUE),
    Third_quartile_mve_winsorized = quantile(mve_winsorized, 0.75, na.rm = TRUE),
    
    Mean_scaled_yoy_sales_wins = mean(scaled_yoy_change_in_sales_winsorized, na.rm = TRUE),
    SD_scaled_yoy_sales_wins = sd(scaled_yoy_change_in_sales_winsorized, na.rm = TRUE),
    First_quartile_scaled_yoy_sales_wins = quantile(scaled_yoy_change_in_sales_winsorized, 0.25, na.rm = TRUE),
    Median_scaled_yoy_sales_wins = median(scaled_yoy_change_in_sales_winsorized, na.rm = TRUE),
    Third_quartile_scaled_yoy_sales_wins = quantile(scaled_yoy_change_in_sales_winsorized, 0.75, na.rm = TRUE),
    
    Mean_roa_winsorized = mean(roa_winsorized, na.rm = TRUE),
    SD_roa_winsorized = sd(roa_winsorized, na.rm = TRUE),
    First_quartile_roa_winsorized = quantile(roa_winsorized, 0.25, na.rm = TRUE),
    Median_roa_winsorized = median(roa_winsorized, na.rm = TRUE),
    Third_quartile_roa_winsorized = quantile(roa_winsorized, 0.75, na.rm = TRUE),
    
    Mean_mtb_winsorized = mean(mtb_winsorized, na.rm = TRUE),
    SD_mtb_winsorized = sd(mtb_winsorized, na.rm = TRUE),
    First_quartile_mtb_winsorized = quantile(mtb_winsorized, 0.25, na.rm = TRUE),
    Median_mtb_winsorized = median(mtb_winsorized, na.rm = TRUE),
    Third_quartile_mtb_winsorized = quantile(mtb_winsorized, 0.75, na.rm = TRUE),
    
    Tercile = first(CAR_Tercile),  # Use the first (or any representative) value of CAR_Tercile for the group
    .groups = 'drop'
  )

# Combine the two sets of summary statistics
combined_summary_stats <- bind_rows(full_sample_summary_stats, tercile_summary_stats)

# Write to excel to format table
#write_xlsx(combined_summary_stats, "./output/annual_summary_stats.xlsx")

# Add deal ratio and deal equity value stats
# Winsorize deal equity value
transactions_for_regression <- transactions_for_regression %>%
  mutate( deal_equity_value_usd_winsorized = Winsorize(deal_equity_value_usd, probs = c(0.05, 0.95), na.rm = TRUE))

grouped_summary_stats <- transactions_for_regression %>%
  group_by(CAR_Tercile) %>%
  summarise(
    Mean_deal_equity = mean(deal_equity_value_usd_winsorized, na.rm = TRUE),
    SD_deal_equity = sd(deal_equity_value_usd_winsorized, na.rm = TRUE),
    First_quartile_deal_equity = quantile(deal_equity_value_usd_winsorized, 0.25, na.rm = TRUE),
    Median_deal_equity = median(deal_equity_value_usd_winsorized, na.rm = TRUE),
    Third_quartile_deal_equity = quantile(deal_equity_value_usd_winsorized, 0.75, na.rm = TRUE),
    
    Mean_deal_ratio = mean(deal_ratio, na.rm = TRUE),
    SD_deal_ratio = sd(deal_ratio, na.rm = TRUE),
    First_quartile_deal_ratio = quantile(deal_ratio, 0.25, na.rm = TRUE),
    Median_deal_ratio = median(deal_ratio, na.rm = TRUE),
    Third_quartile_deal_ratio = quantile(deal_ratio, 0.75, na.rm = TRUE)
  )

# Convert CAR_Tercile to character in the grouped summary stats
grouped_summary_stats <- grouped_summary_stats %>%
  mutate(CAR_Tercile = as.character(CAR_Tercile))

# Define overall summary stats with CAR_Tercile as character
overall_summary_stats <- transactions_for_regression %>%
  summarise(
    Mean_deal_equity = mean(deal_equity_value_usd_winsorized, na.rm = TRUE),
    SD_deal_equity = sd(deal_equity_value_usd_winsorized, na.rm = TRUE),
    First_quartile_deal_equity = quantile(deal_equity_value_usd_winsorized, 0.25, na.rm = TRUE),
    Median_deal_equity = median(deal_equity_value_usd_winsorized, na.rm = TRUE),
    Third_quartile_deal_equity = quantile(deal_equity_value_usd_winsorized, 0.75, na.rm = TRUE),
    
    Mean_deal_ratio = mean(deal_ratio, na.rm = TRUE),
    SD_deal_ratio = sd(deal_ratio, na.rm = TRUE),
    First_quartile_deal_ratio = quantile(deal_ratio, 0.25, na.rm = TRUE),
    Median_deal_ratio = median(deal_ratio, na.rm = TRUE),
    Third_quartile_deal_ratio = quantile(deal_ratio, 0.75, na.rm = TRUE)
  ) %>%
  mutate(CAR_Tercile = "Overall")

# Combine the results
deal_summary_stats <- bind_rows(overall_summary_stats, grouped_summary_stats)

# Combining the new summary statistics with a previous dataframe
final_combined_summary_stats <- bind_rows(combined_summary_stats, deal_summary_stats)

# Write to excel to format table
#write_xlsx(final_combined_summary_stats, "./output/summary_statisticss.xlsx")


##################################

# ANOVA and Kruskal-Wallis test statistics and p-values 

# Function to perform ANOVA and return summary
perform_anova <- function(data, formula) {
  anova_result <- aov(formula, data = data)
  summary(anova_result)
}

# Function to perform Kruskal-Wallis test
perform_kruskal <- function(data, formula) {
  kruskal_result <- kruskal.test(formula, data = data)
  kruskal_result
}


# Initialize an empty list to store test results and assumption checks
anova_results_list <- list()
welch_results_list <- list()
kruskal_results_list <- list()
assumption_checks_list <- list()

# Set CAR_Tercile variable to factor for use in Levene's test
winsorized_balance_sheet_data$CAR_Tercile <- as.factor(winsorized_balance_sheet_data$CAR_Tercile)

# test only one year_from_announcement at a time
winsorized_balance_sheet_data <- winsorized_balance_sheet_data %>%
  filter(years_from_announcement == -1) # should be done for each year_from_announcement

# Loop through each variable to test assumptions, perform ANOVA and Kruskal-Wallis tests
for (variable in variables_to_test) {
  # Formulate the model
  formula <- as.formula(paste(variable, "~ CAR_Tercile"))
  
  # Subset the data only with non-NA values for the current variable and CAR_Tercile
  data_subset <- winsorized_balance_sheet_data[, c(variable, "CAR_Tercile")]
  data_subset <- na.omit(data_subset)  # Remove all rows with any NAs in these columns
  
  # Check if data_subset is non-empty
  if (nrow(data_subset) > 0) {
    # Assumption checks for ANOVA
    model <- aov(formula, data = data_subset)
    residuals_data <- residuals(model)
    shapiro_result <- shapiro.test(residuals_data)  # Normality of residuals
    bartlett_result <- bartlett.test(formula = formula, data = data_subset)  # Homogeneity of variances
    levene_result <- leveneTest(residuals_data ~ CAR_Tercile, data = data_subset, center = median)
    
    # Store the assumption check results
    assumption_checks_list[[variable]] <- list(
      Shapiro_Wilk = shapiro_result,
      Bartlett = bartlett_result,
      Levene = levene_result
    )
    
    # Regular ANOVA test
    anova_result <- perform_anova(data_subset, formula)
    anova_results_list[[variable]] <- anova_result
    
    # Welch ANOVA test
    welch_result <- oneway.test(formula, data = data_subset, var.equal = FALSE)
    welch_results_list[[variable]] <- welch_result
    
    # Kruskal-Wallis test
    kruskal_results_list[[variable]] <- perform_kruskal(data_subset, formula)
  } else {
    cat("No complete cases found for variable:", variable, "\n")
  }
}

# Print out the assumption check results
print("Assumption Checks:")
for (variable in names(assumption_checks_list)) {
  print(paste("Results for", variable))
  print(assumption_checks_list[[variable]])
}

# Print results
print("Regular ANOVA Results:")
for (variable in names(anova_results_list)) {
  print(paste("ANOVA for", variable))
  print(anova_results_list[[variable]])
}

# Print out the Welch ANOVA results
print("Welch ANOVA Results:")
for (variable in names(welch_results_list)) {
  print(paste("Welch ANOVA for", variable))
  print(welch_results_list[[variable]])
}

# Print out the Kruskal-Wallis results
print("Kruskal-Wallis Results:")
print(kruskal_results_list)


############-------------------------------------------------------
########################################################### Plots and KS, Shapiro-Wilk tests for all vars.
# Define a function to perform normality tests and generate plots
perform_checks <- function(df, var_name) {
  df %>%
    group_by(years_from_announcement, CAR_Tercile) %>%
    summarise(
      count = n(),
      shapiro_p_value = if(n() >= 3 & n() <= 5000) shapiro.test(.[[var_name]])$p.value else NA,
      ks_test_p_value = if(n() >= 3) lillie.test(.[[var_name]])$p.value else NA
    ) %>%
    ungroup() %>%
    mutate(variable = var_name) %>% # Add the variable name to the dataframe
    list(
      normality_results = .,
      histogram = ggplot(df, aes(x = .data[[var_name]])) +
        geom_histogram(bins = 30, fill = "blue", color = "black") +
        facet_grid(years_from_announcement ~ CAR_Tercile) +
        ggtitle(paste("Histograms of Winsorized", var_name, "by Year and Window")) +
        theme_minimal() +
        theme(axis.text.x = element_text(angle = 45, hjust = 1, family = "Arial", size = 7)),
      qqplot = ggplot(df, aes(sample = .data[[var_name]])) +
        geom_qq() +
        geom_qq_line() +
        ggtitle(paste("Q-Q Plot for", var_name, "by Year and Tercile")) +
        theme_minimal() +
        theme(text = element_text(family = "Arial", size = 8)) # Set font family and size
    )
}

# Apply the function to each variable and create a list of results
results <- map(variables_to_test, ~perform_checks(winsorized_balance_sheet_data, .x))

# To print all results or save plots, you can loop over this list
for (result in results) {
  print(result$normality_results)
  print(result$histogram)
  print(result$qqplot)
}


########################################################### ROBUSTNESS TESTING
# Truncate extreme values # ROBUSTNESS TESTING
# Function to truncate values beyond the specified quantiles
truncate <- function(x, probs = c(0.05, 0.95)) {
  quantiles <- quantile(x, probs = probs, na.rm = TRUE)
  x[x < quantiles[1]] <- NA
  x[x > quantiles[2]] <- NA
  return(x)
}

# Truncate extreme values
truncated_balance_sheet_data <- cleaned_annual_balance_sheet_data %>% 
  arrange(fiscal_year) %>%
  filter(years_from_announcement >= -1 & years_from_announcement <= 2, Window == "[-1,+1]") %>%
  group_by(years_from_announcement) %>%
  mutate(
    # Replace Winsorize function calls with truncate function
    scaled_awca_truncated = truncate(scaled_awca),
    ln_mve_truncated = truncate(ln_mve),
    scaled_yoy_change_in_sales_truncated = truncate(scaled_yoy_change_in_sales),
    mtb_truncated = truncate(mtb)
  ) %>%
  ungroup()

# Conduct normality tests by year
normality_results_truncated <- truncated_balance_sheet_data %>%
  group_by(years_from_announcement) %>%
  summarise(
    count = n(),
    shapiro_p_value = if(n() >= 3 & n() <= 5000) shapiro.test(scaled_awca_truncated)$p.value else NA,
    ks_test_p_value = if(n() >= 3) lillie.test(scaled_awca_truncated)$p.value else NA
  )

# Print results
print(normality_results_truncated)

# Generating histograms for each year and window
truncated_balance_sheet_data %>%
  ggplot(aes(x = scaled_awca_truncated)) +
  geom_histogram(bins = 30, fill = "blue", color = "black") +
  facet_grid(years_from_announcement ~ Window) +  # Use facet_grid to combine two faceting variables
  ggtitle("Histograms of Winsorized Scaled AWCA by Year and Window") +
  theme_minimal() +  # Optional: clean theme for better visibility
  theme(axis.text.x = element_text(angle = 45, hjust = 1))  # Optional: improve x-axis label readability

truncated_balance_sheet_data %>%
  group_by(years_from_announcement) %>%
  do({
    qqplot <- ggplot(., aes(sample = scaled_awca_truncated)) +
      geom_qq() +
      geom_qq_line() +
      ggtitle(paste("Q-Q Plot for Year", unique(.$years_from_announcement)))
    print(qqplot)
  })
rm(truncated_balance_sheet_data, normality_results_truncated)
######################################### END OF ROBUSTNESS TEST


# End of script

#######wilcoxon 
# Assuming you have `tercile_summary_stats` with different CAR Terciles
group1 <- tercile_summary_stats$scaled_awca_winsorized[tercile_summary_stats$CAR_Tercile == "1"]
group2 <- tercile_summary_stats$scaled_awca_winsorized[tercile_summary_stats$CAR_Tercile == "2"]

# Compare medians between low and high CAR Tercile groups
wilcox.test(group1, group2)



####### box plots
library(tidyr)
library(dplyr)
library(ggplot2)


# Reshape the data to long format
long_data <- tercile_summary_stats %>%
  pivot_longer(
    cols = starts_with("Mean_"),  # or specify all columns you need to plot
    names_to = "Variable",
    values_to = "Value"
  )

# You might want to clean up the Variable names to remove the "Mean_" prefix
long_data$Variable <- sub("Mean_", "", long_data$Variable)
# Plotting boxplots
ggplot(long_data, aes(x = Variable, y = Value, fill = CAR_Tercile)) +
  geom_boxplot() +
  labs(title = "Boxplot of Winsorized Variables",
       x = "Variable",
       y = "Value") +
  theme(axis.text.x = element_text(angle = 90, hjust = 1)) +  # Rotate x labels for clarity
  scale_fill_brewer(palette = "Set3")  # Optional: Use a color palette

# for each year from announcement
# Adding facets to separate by years from announcement
ggplot(long_data, aes(x = Variable, y = Value, fill = CAR_Tercile)) +
  geom_boxplot() +
  facet_wrap(~ years_from_announcement) +  # Creates a separate plot for each year
  labs(title = "Boxplot of Winsorized Variables by Year from Announcement",
       x = "Variable",
       y = "Value") +
  theme(axis.text.x = element_text(angle = 90, hjust = 1)) +
  scale_fill_brewer(palette = "Set3")

########QQ plots for each variable by year from announcement
library(dplyr)
library(ggplot2)
library(tidyr)

# Example variables to plot
variables_to_plot <- c( "scaled_awca_winsorized",
                        "CAR_winsorized",
                        "ln_mve_winsorized", # regression model uses lagged ln_mve
                        "scaled_yoy_change_in_sales_winsorized",
                        "mtb_winsorized", # regression uses lagged mtb
                        "asset_turnover_winsorized", # not in regression but seems interesting to describe sample
                        "scaled_working_capital_winsorized" )# not in regression but seems interesting to describe sample

# Reshape the data to long format for easier plotting
long_format_data <- winsorized_balance_sheet_data %>%
  select(years_from_announcement, one_of(variables_to_plot)) %>%
  pivot_longer(cols = -years_from_announcement, names_to = "variable", values_to = "value")

# Create Q-Q plots
qq_plots <- ggplot(long_format_data, aes(sample = value)) +
  facet_wrap(~ variable + years_from_announcement, scales = "free") +
  stat_qq() +
  stat_qq_line() +
  labs(title = "Q-Q Plots for Winsorized Variables by Year from Announcement",
       x = "Theoretical Quantiles",
       y = "Sample Quantiles")

print(qq_plots)




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

#################################### partition on method of payment to check if
#################### +awca is associated more with shares than cash transactions
full_sample_summary_stats_cash_vs_shares <- winsorized_balance_sheet_data %>%
  group_by(years_from_announcement, stock_deal) %>%
  filter(years_from_announcement >= -1 & years_from_announcement <= 2, Window == "[-1,+1]") %>%
  summarise(
    Transactions_n = n_distinct(`Deal Number`),
    Acquirors_n = n_distinct(`Acquiror ISIN number`),
    
    Mean_scaled_awca_winsorized = mean(scaled_awca_winsorized, na.rm = TRUE),
    t_stat_awca = t.test(scaled_awca_winsorized, mu = 0, alternative = "two.sided", conf.level = 0.95, na.action = na.exclude)$statistic,
    p_value_t_test_awca = t.test(scaled_awca_winsorized, mu = 0, alternative = "two.sided", conf.level = 0.95, na.action = na.exclude)$p.value,
    SD_scaled_awca_winsorized = sd(scaled_awca_winsorized, na.rm = TRUE),
    First_quartile_scaled_awca_winsorized = quantile(scaled_awca_winsorized, 0.25, na.rm = TRUE),
    Median_scaled_awca_winsorized = median(scaled_awca_winsorized, na.rm = TRUE),
    W_stat_awca = wilcox.test(scaled_awca_winsorized, mu = 0, alternative = "two.sided", conf.int = FALSE, conf.level = 0.95, exact = FALSE, correct = TRUE, na.action = na.exclude)$statistic,
    p_value_Wilcox_awca = wilcox.test(scaled_awca_winsorized, mu = 0, alternative = "two.sided", conf.int = FALSE, conf.level = 0.95, exact = FALSE, correct = TRUE, na.action = na.exclude)$p.value,
    Third_quartile_scaled_awca_winsorized = quantile(scaled_awca_winsorized, 0.75, na.rm = TRUE),
    
    Mean_car_winsorized = mean(CAR_winsorized, na.rm = TRUE),
    SD_car_winsorized = sd(CAR_winsorized, na.rm = TRUE),
    First_quartile_car_winsorized = quantile(CAR_winsorized, 0.25, na.rm = TRUE),
    Median_car_winsorized = median(CAR_winsorized, na.rm = TRUE),
    Third_quartile_car_winsorized = quantile(CAR_winsorized, 0.75, na.rm = TRUE),
    t_stat_car = t.test(CAR_winsorized, mu = 0, alternative = "two.sided", conf.level = 0.95, na.action = na.exclude)$statistic,
    p_value_t_test_car = t.test(CAR_winsorized, mu = 0, alternative = "two.sided", conf.level = 0.95, na.action = na.exclude)$p.value,
    W_stat_car = wilcox.test(CAR_winsorized, mu = 0, alternative = "two.sided", conf.int = FALSE, conf.level = 0.95, exact = FALSE, correct = TRUE, na.action = na.exclude)$statistic,
    p_value_Wilcox_car = wilcox.test(CAR_winsorized, mu = 0, alternative = "two.sided", conf.int = FALSE, conf.level = 0.95, exact = FALSE, correct = TRUE, na.action = na.exclude)$p.value,
    
    Mean_mve_winsorized = mean(mve_winsorized, na.rm = TRUE),
    SD_ln_mve_winsorized = sd(mve_winsorized, na.rm = TRUE),
    First_quartile_ln_mve_winsorized = quantile(mve_winsorized, 0.25, na.rm = TRUE),
    Median_ln_mve_winsorized = median(mve_winsorized, na.rm = TRUE),
    Third_quartile_ln_mve_winsorized = quantile(mve_winsorized, 0.75, na.rm = TRUE),
    t_stat_mve = t.test(mve_winsorized, mu = 0, alternative = "two.sided", conf.level = 0.95, na.action = na.exclude)$statistic,
    p_value_t_test_mve = t.test(mve_winsorized, mu = 0, alternative = "two.sided", conf.level = 0.95, na.action = na.exclude)$p.value,
    W_stat_mve = wilcox.test(mve_winsorized, mu = 0, alternative = "two.sided", conf.int = FALSE, conf.level = 0.95, exact = FALSE, correct = TRUE, na.action = na.exclude)$statistic,
    p_value_Wilcox_mve = wilcox.test(mve_winsorized, mu = 0, alternative = "two.sided", conf.int = FALSE, conf.level = 0.95, exact = FALSE, correct = TRUE, na.action = na.exclude)$p.value,
    
    Mean_scaled_yoy_sales_wins = mean(scaled_yoy_change_in_sales_winsorized, na.rm = TRUE),
    t_stat_yoy_sales = t.test(scaled_yoy_change_in_sales_winsorized, mu = 0, alternative = "two.sided", conf.level = 0.95, na.action = na.exclude)$statistic,
    p_value_t_test_yoy_sales = t.test(scaled_yoy_change_in_sales_winsorized, mu = 0, alternative = "two.sided", conf.level = 0.95, na.action = na.exclude)$p.value,
    SD_scaled_yoy_sales_wins = sd(scaled_yoy_change_in_sales_winsorized, na.rm = TRUE),
    First_quartile_scaled_yoy_sales_wins = quantile(scaled_yoy_change_in_sales_winsorized, 0.25, na.rm = TRUE),
    Median_scaled_yoy_sales_wins = median(scaled_yoy_change_in_sales_winsorized, na.rm = TRUE),
    W_stat_yoy_sales = wilcox.test(scaled_yoy_change_in_sales_winsorized, mu = 0, alternative = "two.sided", conf.int = FALSE, conf.level = 0.95, exact = FALSE, correct = TRUE, na.action = na.exclude)$statistic,
    p_value_Wilcox_yoy_sales = wilcox.test(scaled_yoy_change_in_sales_winsorized, mu = 0, alternative = "two.sided", conf.int = FALSE, conf.level = 0.95, exact = FALSE, correct = TRUE, na.action = na.exclude)$p.value,
    Third_quartile_scaled_yoy_sales_wins = quantile(scaled_yoy_change_in_sales_winsorized, 0.75, na.rm = TRUE),
    
    Mean_roa_winsorized = mean(roa_winsorized, na.rm = TRUE),
    t_stat_roa = t.test(roa_winsorized, mu = 0, alternative = "two.sided", conf.level = 0.95, na.action = na.exclude)$statistic,
    p_value_t_test_roa = t.test(roa_winsorized, mu = 0, alternative = "two.sided", conf.level = 0.95, na.action = na.exclude)$p.value,
    SD_roa_winsorized = sd(roa_winsorized, na.rm = TRUE),
    First_quartile_roa_winsorized = quantile(roa_winsorized, 0.25, na.rm = TRUE),
    Median_roa_winsorized = median(roa_winsorized, na.rm = TRUE),
    W_stat_roa = wilcox.test(roa_winsorized, mu = 0, alternative = "two.sided", conf.int = FALSE, conf.level = 0.95, exact = FALSE, correct = TRUE, na.action = na.exclude)$statistic,
    p_value_Wilcox_roa = wilcox.test(roa_winsorized, mu = 0, alternative = "two.sided", conf.int = FALSE, conf.level = 0.95, exact = FALSE, correct = TRUE, na.action = na.exclude)$p.value,
    Third_quartile_roa_winsorized = quantile(roa_winsorized, 0.75, na.rm = TRUE),
    
    Mean_mtb_winsorized = mean(mtb_winsorized, na.rm = TRUE),
    t_stat_mtb = t.test(mtb_winsorized, mu = 0, alternative = "two.sided", conf.level = 0.95, na.action = na.exclude)$statistic,
    p_value_t_test_mtb = t.test(mtb_winsorized, mu = 0, alternative = "two.sided", conf.level = 0.95, na.action = na.exclude)$p.value,
    SD_mtb_winsorized = sd(mtb_winsorized, na.rm = TRUE),
    First_quartile_mtb_winsorized = quantile(mtb_winsorized, 0.25, na.rm = TRUE),
    Median_mtb_winsorized = median(mtb_winsorized, na.rm = TRUE),
    W_stat_mtb = wilcox.test(mtb_winsorized, mu = 0, alternative = "two.sided", conf.int = FALSE, conf.level = 0.95, exact = FALSE, correct = TRUE, na.action = na.exclude)$statistic,
    p_value_Wilcox_mtb = wilcox.test(mtb_winsorized, mu = 0, alternative = "two.sided", conf.int = FALSE, conf.level = 0.95, exact = FALSE, correct = TRUE, na.action = na.exclude)$p.value,
    Third_quartile_mtb_winsorized = quantile(mtb_winsorized, 0.75, na.rm = TRUE),
    
    Tercile = "Full sample",
    .groups = 'drop'
  )

##################### Partition sample on positive/negative CAR instead of terciles
winsorized_balance_sheet_data <- winsorized_balance_sheet_data %>%
  mutate(negative_car = ifelse(CAR_winsorized<0, TRUE, FALSE))

summary_stats_pos_vs_neg_car <- winsorized_balance_sheet_data %>%
  group_by(years_from_announcement, negative_car) %>%
  filter(years_from_announcement >= -1 & years_from_announcement <= 2, Window == "[-1,+1]") %>%
  summarise(
    Transactions_n = n_distinct(`Deal Number`),
    Acquirors_n = n_distinct(`Acquiror ISIN number`),
    
    Mean_scaled_awca_winsorized = mean(scaled_awca_winsorized, na.rm = TRUE),
    t_stat_awca = t.test(scaled_awca_winsorized, mu = 0, alternative = "two.sided", conf.level = 0.95, na.action = na.exclude)$statistic,
    p_value_t_test_awca = t.test(scaled_awca_winsorized, mu = 0, alternative = "two.sided", conf.level = 0.95, na.action = na.exclude)$p.value,
    SD_scaled_awca_winsorized = sd(scaled_awca_winsorized, na.rm = TRUE),
    First_quartile_scaled_awca_winsorized = quantile(scaled_awca_winsorized, 0.25, na.rm = TRUE),
    Median_scaled_awca_winsorized = median(scaled_awca_winsorized, na.rm = TRUE),
    W_stat_awca = wilcox.test(scaled_awca_winsorized, mu = 0, alternative = "two.sided", conf.int = FALSE, conf.level = 0.95, exact = FALSE, correct = TRUE, na.action = na.exclude)$statistic,
    p_value_Wilcox_awca = wilcox.test(scaled_awca_winsorized, mu = 0, alternative = "two.sided", conf.int = FALSE, conf.level = 0.95, exact = FALSE, correct = TRUE, na.action = na.exclude)$p.value,
    Third_quartile_scaled_awca_winsorized = quantile(scaled_awca_winsorized, 0.75, na.rm = TRUE),
    
    Mean_car_winsorized = mean(CAR_winsorized, na.rm = TRUE),
    SD_car_winsorized = sd(CAR_winsorized, na.rm = TRUE),
    First_quartile_car_winsorized = quantile(CAR_winsorized, 0.25, na.rm = TRUE),
    Median_car_winsorized = median(CAR_winsorized, na.rm = TRUE),
    Third_quartile_car_winsorized = quantile(CAR_winsorized, 0.75, na.rm = TRUE),
    t_stat_car = t.test(CAR_winsorized, mu = 0, alternative = "two.sided", conf.level = 0.95, na.action = na.exclude)$statistic,
    p_value_t_test_car = t.test(CAR_winsorized, mu = 0, alternative = "two.sided", conf.level = 0.95, na.action = na.exclude)$p.value,
    W_stat_car = wilcox.test(CAR_winsorized, mu = 0, alternative = "two.sided", conf.int = FALSE, conf.level = 0.95, exact = FALSE, correct = TRUE, na.action = na.exclude)$statistic,
    p_value_Wilcox_car = wilcox.test(CAR_winsorized, mu = 0, alternative = "two.sided", conf.int = FALSE, conf.level = 0.95, exact = FALSE, correct = TRUE, na.action = na.exclude)$p.value,
    
    Mean_mve_winsorized = mean(mve_winsorized, na.rm = TRUE),
    SD_ln_mve_winsorized = sd(mve_winsorized, na.rm = TRUE),
    First_quartile_ln_mve_winsorized = quantile(mve_winsorized, 0.25, na.rm = TRUE),
    Median_ln_mve_winsorized = median(mve_winsorized, na.rm = TRUE),
    Third_quartile_ln_mve_winsorized = quantile(mve_winsorized, 0.75, na.rm = TRUE),
    t_stat_mve = t.test(mve_winsorized, mu = 0, alternative = "two.sided", conf.level = 0.95, na.action = na.exclude)$statistic,
    p_value_t_test_mve = t.test(mve_winsorized, mu = 0, alternative = "two.sided", conf.level = 0.95, na.action = na.exclude)$p.value,
    W_stat_mve = wilcox.test(mve_winsorized, mu = 0, alternative = "two.sided", conf.int = FALSE, conf.level = 0.95, exact = FALSE, correct = TRUE, na.action = na.exclude)$statistic,
    p_value_Wilcox_mve = wilcox.test(mve_winsorized, mu = 0, alternative = "two.sided", conf.int = FALSE, conf.level = 0.95, exact = FALSE, correct = TRUE, na.action = na.exclude)$p.value,
    
    Mean_scaled_yoy_sales_wins = mean(scaled_yoy_change_in_sales_winsorized, na.rm = TRUE),
    t_stat_yoy_sales = t.test(scaled_yoy_change_in_sales_winsorized, mu = 0, alternative = "two.sided", conf.level = 0.95, na.action = na.exclude)$statistic,
    p_value_t_test_yoy_sales = t.test(scaled_yoy_change_in_sales_winsorized, mu = 0, alternative = "two.sided", conf.level = 0.95, na.action = na.exclude)$p.value,
    SD_scaled_yoy_sales_wins = sd(scaled_yoy_change_in_sales_winsorized, na.rm = TRUE),
    First_quartile_scaled_yoy_sales_wins = quantile(scaled_yoy_change_in_sales_winsorized, 0.25, na.rm = TRUE),
    Median_scaled_yoy_sales_wins = median(scaled_yoy_change_in_sales_winsorized, na.rm = TRUE),
    W_stat_yoy_sales = wilcox.test(scaled_yoy_change_in_sales_winsorized, mu = 0, alternative = "two.sided", conf.int = FALSE, conf.level = 0.95, exact = FALSE, correct = TRUE, na.action = na.exclude)$statistic,
    p_value_Wilcox_yoy_sales = wilcox.test(scaled_yoy_change_in_sales_winsorized, mu = 0, alternative = "two.sided", conf.int = FALSE, conf.level = 0.95, exact = FALSE, correct = TRUE, na.action = na.exclude)$p.value,
    Third_quartile_scaled_yoy_sales_wins = quantile(scaled_yoy_change_in_sales_winsorized, 0.75, na.rm = TRUE),
    
    Mean_roa_winsorized = mean(roa_winsorized, na.rm = TRUE),
    t_stat_roa = t.test(roa_winsorized, mu = 0, alternative = "two.sided", conf.level = 0.95, na.action = na.exclude)$statistic,
    p_value_t_test_roa = t.test(roa_winsorized, mu = 0, alternative = "two.sided", conf.level = 0.95, na.action = na.exclude)$p.value,
    SD_roa_winsorized = sd(roa_winsorized, na.rm = TRUE),
    First_quartile_roa_winsorized = quantile(roa_winsorized, 0.25, na.rm = TRUE),
    Median_roa_winsorized = median(roa_winsorized, na.rm = TRUE),
    W_stat_roa = wilcox.test(roa_winsorized, mu = 0, alternative = "two.sided", conf.int = FALSE, conf.level = 0.95, exact = FALSE, correct = TRUE, na.action = na.exclude)$statistic,
    p_value_Wilcox_roa = wilcox.test(roa_winsorized, mu = 0, alternative = "two.sided", conf.int = FALSE, conf.level = 0.95, exact = FALSE, correct = TRUE, na.action = na.exclude)$p.value,
    Third_quartile_roa_winsorized = quantile(roa_winsorized, 0.75, na.rm = TRUE),
    
    Mean_mtb_winsorized = mean(mtb_winsorized, na.rm = TRUE),
    t_stat_mtb = t.test(mtb_winsorized, mu = 0, alternative = "two.sided", conf.level = 0.95, na.action = na.exclude)$statistic,
    p_value_t_test_mtb = t.test(mtb_winsorized, mu = 0, alternative = "two.sided", conf.level = 0.95, na.action = na.exclude)$p.value,
    SD_mtb_winsorized = sd(mtb_winsorized, na.rm = TRUE),
    First_quartile_mtb_winsorized = quantile(mtb_winsorized, 0.25, na.rm = TRUE),
    Median_mtb_winsorized = median(mtb_winsorized, na.rm = TRUE),
    W_stat_mtb = wilcox.test(mtb_winsorized, mu = 0, alternative = "two.sided", conf.int = FALSE, conf.level = 0.95, exact = FALSE, correct = TRUE, na.action = na.exclude)$statistic,
    p_value_Wilcox_mtb = wilcox.test(mtb_winsorized, mu = 0, alternative = "two.sided", conf.int = FALSE, conf.level = 0.95, exact = FALSE, correct = TRUE, na.action = na.exclude)$p.value,
    Third_quartile_mtb_winsorized = quantile(mtb_winsorized, 0.75, na.rm = TRUE),
    
    Tercile = "Full sample",
    .groups = 'drop'
  )




print(doc, target = "./output/Summary_Statistics_Table_transposed.docx")  # Save the Word document