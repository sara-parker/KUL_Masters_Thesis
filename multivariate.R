library(dplyr)
library(readxl)
library(lmtest)
library(sandwich)
library(car)
library(ggplot2)

regression_data <- read_excel("./output/winsorized_annual_balance_sheet_data.xlsx") 
regression_data <- read_excel("./output/winsorized_robustness_check_data.xlsx") # for robustness checks, adds shares and cash percentages

regression_data <- regression_data %>%
  select(`Ticker Symbol`, `Deal Number`, CAR_Tercile, scaled_awca_winsorized, CAR_winsorized, stock_deal,
         ln_mve_winsorized, scaled_yoy_change_in_sales_winsorized, mtb_winsorized, roa_winsorized, cross_border, industry_relatedness,
         serial, awca_year_before_announcement, years_from_announcement, industry, is_Big4, fiscal_year) %>%
  mutate(stock_deal = factor(stock_deal),
         cross_border = factor(cross_border),
         industry_relatedness = factor(industry_relatedness),
         serial = factor(serial),
         is_Big4 = factor(is_Big4),
         CAR_Tercile = factor(CAR_Tercile, levels = c("2", "1", "3")),  # Sets Tercile 2 as the reference category
         industry = factor(industry)) %>%
  arrange(`Deal Number`, fiscal_year) %>%
  group_by(`Deal Number`) %>%
  mutate(lagged_ln_mve_wins = lag(ln_mve_winsorized),
         lagged_roa_wins = lag(roa_winsorized),
         lagged_mtb_wins = lag(mtb_winsorized),
         lagged_scaled_yoy_change_sales_wins = lag(scaled_yoy_change_in_sales_winsorized)) %>%
  ungroup() %>%
  filter(years_from_announcement >= -1 & years_from_announcement <= 2)

options(scipen = 99999)

################################# MAIN REGRESSION: full sample 

# Define the years for which the model should be run
years <- c(0, 1) 
full_sample_models <- list()

# Loop to build models for each year
for (year in years) {
  year_data <- filter(regression_data, years_from_announcement == year)
  
  if (nrow(year_data) > 0) {
    model <- lm(scaled_awca_winsorized ~ CAR_winsorized + stock_deal + lagged_ln_mve_wins +
                  lagged_roa_wins + lagged_scaled_yoy_change_sales_wins + lagged_mtb_wins + 
                  cross_border + industry_relatedness + serial + is_Big4 + awca_year_before_announcement +
                  factor(industry), data = year_data) 
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
  
  # Plotting Residuals vs. Fitted Values
  par(mfrow=c(1, 2))  # Set plotting area to display two plots side by side
  plot(predicted_values, residuals_values, main = paste("Residuals vs. Fitted - Year", year),
       xlab = "Fitted Values", ylab = "Residuals", pch = 20)
  abline(h = 0, col = "red")  # Adds a horizontal line at zero
  
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



######################### ROBUSTNESS TESTS: Per tercile and year.

regression_data <- read_excel("./output/winsorized_annual_balance_sheet_data.xlsx") 

regression_data <- regression_data %>%
  select(`Ticker Symbol`, `Deal Number`, CAR_Tercile, scaled_awca_winsorized, CAR_winsorized, stock_deal, ln_mve_winsorized,
         scaled_yoy_change_in_sales_winsorized, mtb_winsorized, roa_winsorized, cross_border, industry_relatedness,
         serial, awca_year_before_announcement, years_from_announcement, industry, is_Big4, fiscal_year) %>%
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

# Define the specific years for which the model should be run
years_announcement <- c(0, 1) 

# CAR Terciles as previously defined
car_terciles <- levels(regression_data$CAR_Tercile)

# Initialize list to store models
tercile_year_models <- list()

# Loop to build models for each CAR_Tercile and specified years
for (tercile in car_terciles) {
  tercile_year_models[[tercile]] <- list()
  for (year in years_announcement) {
    tercile_year_data <- subset(regression_data, CAR_Tercile == tercile & years_from_announcement == year)
    if (nrow(tercile_year_data) > 0) {  # Ensure there is data for the subset
      model <- lm(scaled_awca_winsorized ~ CAR_winsorized + stock_deal + lagged_ln_mve_wins +
                    roa_winsorized + lagged_scaled_yoy_change_sales_wins + lagged_mtb_wins + 
                    cross_border + industry_relatedness + serial + is_Big4 +
                    factor(industry) , data = tercile_year_data)
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
    
    # Display F-statistic and its associated p-value
    fstatistic <- summary(current_model)$fstatistic
    fvalue <- fstatistic["value"]
    df1 <- fstatistic["numdf"]  # numerator degrees of freedom
    df2 <- fstatistic["dendf"]  # denominator degrees of freedom
    fpvalue <- pf(fvalue, df1, df2, lower.tail = FALSE)
    cat("F-statistic:", fvalue, "on", df1, "and", df2, "DF, p-value:", fpvalue, "\n\n")
    
    # Display VIF values for multicollinearity check
    cat("Variance Inflation Factors:\n")
    print(vif(current_model))
    cat("\n")
    
    # Subsetting data to match current tercile and year for plotting
    current_data <- subset(regression_data, CAR_Tercile == tercile & years_from_announcement == as.numeric(year))
    residuals_values <- residuals(current_model)  # Calculate residuals
    predicted_values <- predict(current_model, newdata = current_data)  # Get predicted values
    
    # Plotting Residuals vs. Fitted Values
    par(mfrow=c(1, 2))  # Set plotting area to display two plots side by side
    plot(predicted_values, residuals_values, main = paste("Residuals vs. Fitted -", tercile, "Year", year),
         xlab = "Fitted Values", ylab = "Residuals", pch = 20)
    abline(h = 0, col = "red")  # Adds a horizontal line at zero
    
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




