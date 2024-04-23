library(Hmisc) # for rcorr() which computes correlations and p-values
library(xtable) # for formatting tables
library(dplyr)
library(readxl)

annual_balance_sheet_correlation <- read_excel("./output/winsorized_annual_balance_sheet_data.xlsx")


######################################################################################## CORRELATION MATRIX
# Variables to test 
#variables_to_test <- c(
#  "scaled_awca_winsorized",
#  "car_winsorized",
#  "ln_mve_winsorized", # regression model uses lagged ln_mve
#  "scaled_yoy_change_in_sales_winsorized",
#  "mtb_winsorized", # regression uses lagged mtb
#  "asset_turnover_winsorized", # not in regression but seems interesting to describe sample
#  "scaled_working_capital_winsorized" # not in regression but seems interesting to describe sample

# Subset dataframe with only the variables of interest
correlation_variables <- annual_balance_sheet_correlation[, c("scaled_awca_winsorized", "car_winsorized",
                                                              "ln_mve_winsorized", "scaled_yoy_change_in_sales_winsorized",
                                                              "mtb_winsorized", "asset_turnover_winsorized",
                                                              "scaled_working_capital_winsorized")]


# Calculate Spearman correlation matrix
spearman_corr <- cor(correlation_variables, method = "spearman")

# Calculate Pearson correlation matrix
pearson_corr <- cor(correlation_variables, method = "pearson")

# Initialize the combined matrix with the same dimensions as your correlation matrices
combined_corr_matrix <- matrix(nrow = ncol(correlation_variables), ncol = ncol(correlation_variables))

# Assign names to the rows and columns
rownames(combined_corr_matrix) <- colnames(correlation_variables)
colnames(combined_corr_matrix) <- colnames(correlation_variables)

# Fill the upper triangle of the combined matrix with Spearman coefficients
combined_corr_matrix[upper.tri(combined_corr_matrix)] <- spearman_corr[upper.tri(spearman_corr)]

# Fill the lower triangle with Pearson coefficients
combined_corr_matrix[lower.tri(combined_corr_matrix)] <- pearson_corr[lower.tri(pearson_corr)]

# Convert the matrix to a dataframe for easier viewing/manipulation
combined_corr_df <- as.data.frame(combined_corr_matrix)

# View the combined correlation matrix dataframe
print(combined_corr_df)


################################################################################################################OPTION 2
annual_balance_sheet_correlation <- read_excel("./output/winsorized_annual_balance_sheet_data.xlsx")

# Subset dataframe with only the variables of interest
correlation_variables <- annual_balance_sheet_correlation[, c("scaled_awca_winsorized", "car_winsorized",
                                                              "ln_mve_winsorized", "scaled_yoy_change_in_sales_winsorized",
                                                              "mtb_winsorized", "asset_turnover_winsorized",
                                                              "scaled_working_capital_winsorized")]

# Initialize matrices for Spearman and Pearson correlation coefficients and p-values
spearman_corr <- matrix(nrow = ncol(correlation_variables), ncol = ncol(correlation_variables))
pearson_corr <- matrix(nrow = ncol(correlation_variables), ncol = ncol(correlation_variables))
spearman_pval <- matrix(nrow = ncol(correlation_variables), ncol = ncol(correlation_variables))
pearson_pval <- matrix(nrow = ncol(correlation_variables), ncol = ncol(correlation_variables))

# Fill matrices with correlation coefficients and p-values
for (i in 1:ncol(correlation_variables)) {
  for (j in i:ncol(correlation_variables)) {
    if (i != j) {
      # Explicitly extract vectors using double bracket
      vec_i <- correlation_variables[[i]]
      vec_j <- correlation_variables[[j]]
      
      # Spearman correlation test
      spearman_test <- cor.test(vec_i, vec_j, method = "spearman")
      spearman_corr[i, j] <- spearman_test$estimate
      spearman_corr[j, i] <- spearman_test$estimate
      spearman_pval[i, j] <- spearman_test$p.value
      spearman_pval[j, i] <- spearman_test$p.value
      
      # Pearson correlation test
      pearson_test <- cor.test(vec_i, vec_j, method = "pearson")
      pearson_corr[i, j] <- pearson_test$estimate
      pearson_corr[j, i] <- pearson_test$estimate
      pearson_pval[i, j] <- pearson_test$p.value
      pearson_pval[j, i] <- pearson_test$p.value
    }
  }
}

# Combine the results into a single display matrix
combined_corr_matrix <- matrix(nrow = 2 * ncol(correlation_variables), ncol = ncol(correlation_variables))

# Set up row and column names
colnames(combined_corr_matrix) <- colnames(correlation_variables)
rownames(combined_corr_matrix) <- c(paste(colnames(correlation_variables), "Corr"), 
                                    paste(colnames(correlation_variables), "P-val"))

# Fill the matrix
combined_corr_matrix[1:ncol(correlation_variables), ] <- spearman_corr
combined_corr_matrix[(ncol(correlation_variables) + 1):(2 * ncol(correlation_variables)), ] <- spearman_pval

# Convert the matrix to a dataframe for easier viewing
combined_corr_df <- as.data.frame(combined_corr_matrix)
######################################################################################## 
