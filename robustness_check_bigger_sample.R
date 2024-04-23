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


################################## FOLLOWS FROM END OF CARS.R  ########################################

######################  REPORTS SUMMARY STATS PER YEAR, RELATIVE TO ANNOUNCEMENT ######################

################### FOR REGRESSIONS, START FROM  clean_regression_data BELOW ##########################
 

options(digits=6) 

# Import manually cleand annual bs data
transactions_for_analysis <- read_excel("./output/robust_check_transactions.xlsx")
annual_balance_sheet_data <- read_excel("./input/robustness_balance_sheet.xlsx")

annual_balance_sheet_data$`Data Date` <- as.Date(annual_balance_sheet_data$`Data Date`)
transactions_for_analysis$`Announced date` <- as.Date(transactions_for_analysis$`Announced date`)

annual_balance_sheet_data$fiscal_year <- annual_balance_sheet_data$`Data Year - Fiscal` # rename column
annual_balance_sheet_data <- annual_balance_sheet_data %>%
  select(-`Global Company Key`, -`Data Date`, -`Data Year - Fiscal`) # remove unnecessary columns

transactions_for_analysis <- transactions_for_analysis %>%
  select(`Deal Number`, `Acquiror ISIN number`, `Acquiror ticker symbol`,`Acquiror name`, `Announced date`, year_of_announcement,
         CAR_Tercile, CAR_winsorized, industry_relatedness, industry, serial, stock_deal,
         cross_border, is_Big4, deal_ratio)


summary(annual_balance_sheet_data)
summary(transactions_for_analysis)

options(scipen = 999)


##############################################################################################################
##############################################################################################################
######################################### ANNUAL DATA ################################################
##############################################################################################################
##############################################################################################################


# Extract the M&A announcement year for each firm and create a new dataframe
# which includes the date range for each deal
transactions_for_analysis <- transactions_for_analysis %>%
  mutate(
    start_year = year_of_announcement - 2,
    end_year = year_of_announcement + 2
  )

# Semi-join to filter balance_sheet_data
filtered_balance_sheet_data <- annual_balance_sheet_data %>%
# Join by Deal Number to get the date ranges
  left_join(transactions_for_analysis, by = c(`Ticker Symbol` = "Acquiror ticker symbol" )) %>%
# Filter to include only rows within the date range
  filter(fiscal_year >= start_year & fiscal_year <= end_year) %>%
  mutate(years_from_announcement = fiscal_year - year_of_announcement) %>%
# Select the original balance_sheet_data columns
  select(-start_year, -end_year)

# Count unique Deal Numbers
n_distinct(filtered_balance_sheet_data$`Deal Number`)

summary(filtered_balance_sheet_data)



# Variables to test 
variables_to_test <- c(
  "scaled_awca_winsorized",
  "ln_mve_winsorized",
  "scaled_yoy_change_in_sales_winsorized",
  "mtb_winsorized"
)

# Calculate new variables 
filtered_balance_sheet_data <- filtered_balance_sheet_data %>%
  arrange(`Deal Number`, fiscal_year) %>%
  group_by(`Deal Number`) %>%
  mutate(
    lagged_total_assets = lag(`Assets - Total`),
    lagged_sales = lag(`Sales/Turnover (Net)`),
    yoy_change_in_sales = `Sales/Turnover (Net)` - lagged_sales,
    scaled_yoy_change_in_sales = yoy_change_in_sales / lagged_total_assets,
    ln_mve = log(`Market Value - Total - Fiscal`),
    mtb = `Market Value - Total - Fiscal` / `Common/Ordinary Equity - Total`,
    working_capital = (`Current Assets - Total` - `Cash and Short-Term Investments`) - `Current Liabilities - Total` + `Debt in Current Liabilities - Total`,
    lagged_working_capital = lag(working_capital),
    awca = working_capital - (lagged_working_capital / lagged_sales) * `Sales/Turnover (Net)`,
    scaled_awca = awca / lagged_total_assets,
  ) %>%
  ungroup()

# Check lagged variables
lagged_vars <- filtered_balance_sheet_data %>%
  select(`Deal Number`, `Assets - Total`, lagged_total_assets, `Sales/Turnover (Net)`, lagged_sales, working_capital, lagged_working_capital)

# Check working capital
working_capital_check <- filtered_balance_sheet_data %>%
  select(`Deal Number`, `Current Assets - Total`, `Cash and Short-Term Investments`, `Current Liabilities - Total`, `Debt in Current Liabilities - Total`, working_capital)

# Winsorize variables of interest
winsorized_balance_sheet_data <- filtered_balance_sheet_data %>% 
  group_by(years_from_announcement) %>% # variables winsorized per year from announcement
  mutate(
    scaled_awca_winsorized = Winsorize(scaled_awca, probs = c(0.05, 0.95), na.rm = TRUE),
    ln_mve_winsorized = Winsorize(ln_mve, probs = c(0.05, 0.95), na.rm = TRUE),
    scaled_yoy_change_in_sales_winsorized = Winsorize(scaled_yoy_change_in_sales, probs = c(0.05, 0.95), na.rm = TRUE),
    mtb_winsorized = Winsorize(mtb, probs = c(0.05, 0.95), na.rm = TRUE),
    mve_winsorized = Winsorize(`Market Value - Total - Fiscal`, probs = c(0.05, 0.95), na.rm = TRUE) # summary stats should be mve (not logged value)
  ) %>%
  ungroup()

# Add AWCA in year (-1)
awca_year_before_announcement <- winsorized_balance_sheet_data %>%
  filter(years_from_announcement == -1) %>%
  select(`Deal Number`, awca_year_before_announcement = scaled_awca_winsorized)

# Left join this helper dataframe back to the original dataframe to spread the awca value across all rows of each deal
winsorized_balance_sheet_data <- winsorized_balance_sheet_data %>%
  left_join(awca_year_before_announcement, by = "Deal Number")

# Ensure CAR_Tercile is a factor
winsorized_balance_sheet_data$CAR_Tercile <- as.factor(winsorized_balance_sheet_data$CAR_Tercile)



# Calculate summary statistics for the winsorized variables ################################################## full sample
full_sample_summary_stats <- winsorized_balance_sheet_data %>%
  group_by(years_from_announcement)%>% 
  filter(years_from_announcement>=-1 & years_from_announcement <=2)%>%
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
    t_stat_car = t.test(CAR_winsorized, mu = 0, alternative = "two.sided", conf.level = 0.95, na.action = na.exclude)$statistic,
    p_value_t_test_car = t.test(CAR_winsorized, mu = 0, alternative = "two.sided", conf.level = 0.95, na.action = na.exclude)$p.value,
    SD_car_winsorized = sd(CAR_winsorized, na.rm = TRUE),
    First_quartile_car_winsorized = quantile(CAR_winsorized, 0.25, na.rm = TRUE),
    Median_car_winsorized = median(CAR_winsorized, na.rm = TRUE),
    W_stat_car = wilcox.test(CAR_winsorized, mu = 0, alternative = "two.sided", conf.int = FALSE, conf.level = 0.95, exact = FALSE, correct = TRUE, na.action = na.exclude)$statistic,
    p_value_Wilcox_car = wilcox.test(CAR_winsorized, mu = 0, alternative = "two.sided", conf.int = FALSE, conf.level = 0.95, exact = FALSE, correct = TRUE, na.action = na.exclude)$p.value,
    Third_quartile_car_winsorized = quantile(CAR_winsorized, 0.75, na.rm = TRUE),
    
    Mean_mve_winsorized = mean(mve_winsorized, na.rm = TRUE),
    t_stat_mve = t.test(mve_winsorized, mu = 0, alternative = "two.sided", conf.level = 0.95, na.action = na.exclude)$statistic,
    p_value_t_test_mve = t.test(mve_winsorized, mu = 0, alternative = "two.sided", conf.level = 0.95, na.action = na.exclude)$p.value,
    SD_ln_mve_winsorized = sd(mve_winsorized, na.rm = TRUE),
    First_quartile_ln_mve_winsorized = quantile(mve_winsorized, 0.25, na.rm = TRUE),
    Median_ln_mve_winsorized = median(mve_winsorized, na.rm = TRUE),
    W_stat_mve = wilcox.test(mve_winsorized, mu = 0, alternative = "two.sided", conf.int = FALSE, conf.level = 0.95, exact = FALSE, correct = TRUE, na.action = na.exclude)$statistic,
    p_value_Wilcox_mve = wilcox.test(mve_winsorized, mu = 0, alternative = "two.sided", conf.int = FALSE, conf.level = 0.95, exact = FALSE, correct = TRUE, na.action = na.exclude)$p.value,
    Third_quartile_ln_mve_winsorized = quantile(mve_winsorized, 0.75, na.rm = TRUE),
    
    Mean_scaled_yoy_sales_wins = mean(scaled_yoy_change_in_sales_winsorized, na.rm = TRUE),
    t_stat_yoy_sales = t.test(scaled_yoy_change_in_sales_winsorized, mu = 0, alternative = "two.sided", conf.level = 0.95, na.action = na.exclude)$statistic,
    p_value_t_test_yoy_sales = t.test(scaled_yoy_change_in_sales_winsorized, mu = 0, alternative = "two.sided", conf.level = 0.95, na.action = na.exclude)$p.value,
    SD_scaled_yoy_sales_wins = sd(scaled_yoy_change_in_sales_winsorized, na.rm = TRUE),
    First_quartile_scaled_yoy_sales_wins = quantile(scaled_yoy_change_in_sales_winsorized, 0.25, na.rm = TRUE),
    Median_scaled_yoy_sales_wins = median(scaled_yoy_change_in_sales_winsorized, na.rm = TRUE),
    W_stat_yoy_sales = wilcox.test(scaled_yoy_change_in_sales_winsorized, mu = 0, alternative = "two.sided", conf.int = FALSE, conf.level = 0.95, exact = FALSE, correct = TRUE, na.action = na.exclude)$statistic,
    p_value_Wilcox_yoy_sales = wilcox.test(scaled_yoy_change_in_sales_winsorized, mu = 0, alternative = "two.sided", conf.int = FALSE, conf.level = 0.95, exact = FALSE, correct = TRUE, na.action = na.exclude)$p.value,
    Third_quartile_scaled_yoy_sales_wins = quantile(scaled_yoy_change_in_sales_winsorized, 0.75, na.rm = TRUE),
    
    
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


# Calculate summary statistics for the winsorized variables per tercile

tercile_summary_stats <- winsorized_balance_sheet_data %>%
  group_by(CAR_Tercile, years_from_announcement) %>%
  filter(years_from_announcement>=-1 & years_from_announcement <=2)%>%
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
    t_stat_car = t.test(CAR_winsorized, mu = 0, alternative = "two.sided", conf.level = 0.95, na.action = na.exclude)$statistic,
    p_value_t_test_car = t.test(CAR_winsorized, mu = 0, alternative = "two.sided", conf.level = 0.95, na.action = na.exclude)$p.value,
    SD_car_winsorized = sd(CAR_winsorized, na.rm = TRUE),
    First_quartile_car_winsorized = quantile(CAR_winsorized, 0.25, na.rm = TRUE),
    Median_car_winsorized = median(CAR_winsorized, na.rm = TRUE),
    W_stat_car = wilcox.test(CAR_winsorized, mu = 0, alternative = "two.sided", conf.int = FALSE, conf.level = 0.95, exact = FALSE, correct = TRUE, na.action = na.exclude)$statistic,
    p_value_Wilcox_car = wilcox.test(CAR_winsorized, mu = 0, alternative = "two.sided", conf.int = FALSE, conf.level = 0.95, exact = FALSE, correct = TRUE, na.action = na.exclude)$p.value,
    Third_quartile_car_winsorized = quantile(CAR_winsorized, 0.75, na.rm = TRUE),
    
    Mean_mve_winsorized = mean(mve_winsorized, na.rm = TRUE),
    t_stat_mve = t.test(mve_winsorized, mu = 0, alternative = "two.sided", conf.level = 0.95, na.action = na.exclude)$statistic,
    p_value_t_test_mve = t.test(mve_winsorized, mu = 0, alternative = "two.sided", conf.level = 0.95, na.action = na.exclude)$p.value,
    SD_ln_mve_winsorized = sd(mve_winsorized, na.rm = TRUE),
    First_quartile_ln_mve_winsorized = quantile(mve_winsorized, 0.25, na.rm = TRUE),
    Median_ln_mve_winsorized = median(mve_winsorized, na.rm = TRUE),
    W_stat_mve = wilcox.test(mve_winsorized, mu = 0, alternative = "two.sided", conf.int = FALSE, conf.level = 0.95, exact = FALSE, correct = TRUE, na.action = na.exclude)$statistic,
    p_value_Wilcox_mve = wilcox.test(mve_winsorized, mu = 0, alternative = "two.sided", conf.int = FALSE, conf.level = 0.95, exact = FALSE, correct = TRUE, na.action = na.exclude)$p.value,
    Third_quartile_ln_mve_winsorized = quantile(mve_winsorized, 0.75, na.rm = TRUE),
    
    Mean_scaled_yoy_sales_wins = mean(scaled_yoy_change_in_sales_winsorized, na.rm = TRUE),
    t_stat_yoy_sales = t.test(scaled_yoy_change_in_sales_winsorized, mu = 0, alternative = "two.sided", conf.level = 0.95, na.action = na.exclude)$statistic,
    p_value_t_test_yoy_sales = t.test(scaled_yoy_change_in_sales_winsorized, mu = 0, alternative = "two.sided", conf.level = 0.95, na.action = na.exclude)$p.value,
    SD_scaled_yoy_sales_wins = sd(scaled_yoy_change_in_sales_winsorized, na.rm = TRUE),
    First_quartile_scaled_yoy_sales_wins = quantile(scaled_yoy_change_in_sales_winsorized, 0.25, na.rm = TRUE),
    Median_scaled_yoy_sales_wins = median(scaled_yoy_change_in_sales_winsorized, na.rm = TRUE),
    W_stat_yoy_sales = wilcox.test(scaled_yoy_change_in_sales_winsorized, mu = 0, alternative = "two.sided", conf.int = FALSE, conf.level = 0.95, exact = FALSE, correct = TRUE, na.action = na.exclude)$statistic,
    p_value_Wilcox_yoy_sales = wilcox.test(scaled_yoy_change_in_sales_winsorized, mu = 0, alternative = "two.sided", conf.int = FALSE, conf.level = 0.95, exact = FALSE, correct = TRUE, na.action = na.exclude)$p.value,
    Third_quartile_scaled_yoy_sales_wins = quantile(scaled_yoy_change_in_sales_winsorized, 0.75, na.rm = TRUE),
    
    
    Mean_mtb_winsorized = mean(mtb_winsorized, na.rm = TRUE),
    t_stat_mtb = t.test(mtb_winsorized, mu = 0, alternative = "two.sided", conf.level = 0.95, na.action = na.exclude)$statistic,
    p_value_t_test_mtb = t.test(mtb_winsorized, mu = 0, alternative = "two.sided", conf.level = 0.95, na.action = na.exclude)$p.value,
    SD_mtb_winsorized = sd(mtb_winsorized, na.rm = TRUE),
    First_quartile_mtb_winsorized = quantile(mtb_winsorized, 0.25, na.rm = TRUE),
    Median_mtb_winsorized = median(mtb_winsorized, na.rm = TRUE),
    W_stat_mtb = wilcox.test(mtb_winsorized, mu = 0, alternative = "two.sided", conf.int = FALSE, conf.level = 0.95, exact = FALSE, correct = TRUE, na.action = na.exclude)$statistic,
    p_value_Wilcox_mtb = wilcox.test(mtb_winsorized, mu = 0, alternative = "two.sided", conf.int = FALSE, conf.level = 0.95, exact = FALSE, correct = TRUE, na.action = na.exclude)$p.value,
    Third_quartile_mtb_winsorized = quantile(mtb_winsorized, 0.75, na.rm = TRUE),
    Tercile = first(CAR_Tercile),  # Use the first (or any representative) value of CAR_Tercile for the group
    .groups = 'drop'
  )

# Combine the two sets of summary statistics
combined_summary_stats <- bind_rows(full_sample_summary_stats, tercile_summary_stats)


# ANOVA and Kruskal-Wallis test statistics and p-values ################################################# test statistics and p-values


################## 

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




################################################# LINEAR REGRESSION
regression_data <- winsorized_balance_sheet_data

regression_data <- regression_data %>%
  select(`Ticker Symbol`, `Deal Number`, CAR_Tercile, scaled_awca_winsorized, CAR_winsorized, stock_deal, ln_mve_winsorized,
         scaled_yoy_change_in_sales_winsorized, mtb_winsorized, cross_border, industry_relatedness,
         serial, awca_year_before_announcement, years_from_announcement, industry, is_Big4, fiscal_year, deal_ratio) %>%
mutate(stock_deal = factor(stock_deal),
         cross_border = factor(cross_border),
         industry_relatedness = factor(industry_relatedness),
         serial = factor(serial),
         is_Big4 = factor(is_Big4),
         CAR_Tercile = factor(CAR_Tercile),
         industry = factor(industry)) %>%
  arrange(`Deal Number`, fiscal_year) %>%
  group_by(`Deal Number`) %>%
  mutate(lagged_ln_mve_wins = lag(ln_mve_winsorized),
         lagged_mtb_wins = lag(mtb_winsorized),
         lagged_scaled_yoy_change_sales_wins = lag(scaled_yoy_change_in_sales_winsorized)) %>%
  ungroup() %>%
  filter(years_from_announcement >= -1 & years_from_announcement <= 2)


# Filter out incomplete cases
filtered_data <- regression_data %>%
  filter(years_from_announcement >=0 & years_from_announcement<= 2)
  # Using `complete.cases()` on the subset of data for each group
filtered_data <- filtered_data %>%  
  group_by(`Deal Number`)%>%
  filter(!is.na(scaled_awca_winsorized))%>%
  filter(!is.na(awca_year_before_announcement))%>%
  filter(!is.na(lagged_ln_mve_wins))%>%
  filter(!is.na(lagged_mtb_wins))%>%
  filter(!is.na(lagged_scaled_yoy_change_sales_wins))%>%
  ungroup()


# Define the years for which the model should be run
years <- c(0, 1, 2) 
full_sample_models <- list()

# Loop to build models for each year
for (year in years) {
  year_data <- filter(filtered_data, years_from_announcement == year)
  
  if (nrow(year_data) > 0) {
    model <- lm(scaled_awca_winsorized ~ CAR_winsorized + stock_deal + lagged_ln_mve_wins +
                  lagged_scaled_yoy_change_sales_wins + lagged_mtb_wins + 
                  cross_border + industry_relatedness + serial + is_Big4 + awca_year_before_announcement +
                  deal_ratio + factor(industry) + factor(fiscal_year), data = year_data) 
    full_sample_models[[as.character(year)]] <- model
  } else {
    cat("No data available for year", year, "\n")
  }
}

# Loop for results and diagnostics adapted for the full sample model
for (year in names(full_sample_models)) {
  current_model <- full_sample_models[[year]]
  
  # Ensure the model is a linear model object
  if (!inherits(current_model, "lm")) {
    cat("Model for Year:", year, "is not a linear model object.\n")
    next  # Skip this iteration
  }
  
  model_summary <- summary(current_model)
  
  # Display main regression results
  cat("Results for Year:", year, "\n")
  print(model_summary)  # This prints all main results including coefficients, R-squared, and Adjusted R-squared
  
  # Display VIF values for multicollinearity check
  cat("Variance Inflation Factors:\n")
  print(vif(current_model))
  cat("\n")

  
  # Subsetting data to match the current year for plotting
  current_data <- subset(regression_data, years_from_announcement == as.numeric(year))
  residuals_values <- residuals(current_model)  # Calculate residuals
  predicted_values <- predict(current_model, newdata = current_data)  # Get predicted values

  
  # QQ Plot for Normality of Residuals
  qqnorm(residuals_values, main = paste("Normal Q-Q Plot - Year", year))
  qqline(residuals_values, col = "red")
  cat("\n")
  
  # Shapiro-Wilk Test for Normality
  shapiro_test <- shapiro.test(residuals_values)
  cat("Shapiro-Wilk Test for Normality of Residuals:\n")
  print(shapiro_test)
  cat("\n")
}



######################### for Terciles 1, 2 and 3 
# Define the specific years for which the model should be run
years_announcement <- c(0, 1, 2) 

# CAR Terciles as previously defined
car_terciles <- levels(filtered_data$CAR_Tercile)

# Initialize list to store models
tercile_year_models <- list()

# Loop to build models for each CAR_Tercile and specified years
for (tercile in car_terciles) {
  tercile_year_models[[tercile]] <- list()
  for (year in years_announcement) {
    tercile_year_data <- subset(regression_data, CAR_Tercile == tercile & years_from_announcement == year)
    if (nrow(tercile_year_data) > 0) {  # Ensure there is data for the subset
      model <- lm(scaled_awca_winsorized ~ CAR_winsorized + stock_deal + lagged_ln_mve_wins +
                   lagged_scaled_yoy_change_sales_wins + lagged_mtb_wins + deal_ratio +
                    cross_border + industry_relatedness + serial + is_Big4 +
                    factor(industry) + awca_year_before_announcement , data = tercile_year_data)
      tercile_year_models[[tercile]][[as.character(year)]] <- model  # Store lm object directly for later use
    } else {
      cat(sprintf("No data available for Tercile %s, Year %d.\n", tercile, year))
    }
  }
}
# Diagnostics adjusted to ensure only valid indices are used
for (tercile in names(tercile_year_models)) {
  for (year in names(tercile_year_models[[tercile]])) {
    current_model <- tercile_year_models[[tercile]][[year]]
    if (!inherits(current_model, "lm")) {
      cat("Model for Tercile:", tercile, "Year:", year, "is not a linear model object.\n")
      next
    }
    
    # Prepare data for plotting
    current_data <- subset(regression_data, CAR_Tercile == tercile & years_from_announcement == as.numeric(year))
    residuals_values <- residuals(current_model)
    predicted_values <- predict(current_model, newdata = current_data)
    
    # Remove NA values before plotting
    valid_indices <- !is.na(residuals_values) & !is.na(predicted_values)
    residuals_values <- residuals_values[valid_indices]
    predicted_values <- predicted_values[valid_indices]
    
    # Ensure correct lengths for plotting to avoid warnings
    if (length(residuals_values) != length(predicted_values)) {
      cat("Mismatch in lengths of predicted and residual values for Tercile:", tercile, "Year:", year, "\n")
      next
    }
  }
}

# Loop for results and diagnostics
for (tercile in names(tercile_year_models)) {
  for (year in names(tercile_year_models[[tercile]])) {
    current_model <- tercile_year_models[[tercile]][[year]]
    
    # Display main regression results
    cat("Results for Tercile:", tercile, "Year:", year, "\n")
    print(summary(current_model))  # This prints all main results including coefficients, R-squared, and Adjusted R-squared
    
    # Display VIF values for multicollinearity check
    cat("Variance Inflation Factors:\n")
    print(vif(current_model))
    cat("\n")
    
    # Subsetting data to match current tercile and year for plotting
    current_data <- subset(regression_data, CAR_Tercile == tercile & years_from_announcement == as.numeric(year))
    residuals_values <- residuals(current_model)  # Calculate residuals
    predicted_values <- predict(current_model, newdata = current_data)  # Get predicted values
    
    # QQ Plot for Normality of Residuals
    qqnorm(residuals(current_model), main = paste("Normal Q-Q Plot -", tercile, "Year", year))
    qqline(residuals(current_model), col = "red")
    cat("\n")
    
    # Shapiro-Wilk Test for Normality
    shapiro_test <- shapiro.test(residuals(current_model))
    cat("Shapiro-Wilk Test for Normality of Residuals:\n")
    print(shapiro_test)
    cat("\n")
  }
}
