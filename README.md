# Earnings Management in the Post-Deal Implementation Phase: A U.S. perspective

## Supporting documents: R scripts

### CARs.R
1. Load libraries.
2. Read in data.

- ./input/Zephyr_Definitive_Repeating.xlsx
- ./input/unique_acquiror_isins2.xlsx
- ./input/financial_data_CRSP.csv
- ./input/robustness_tests_balance_sheet.xlsx

3. This script does the following:
- cleans and filters data
- estimates alpha and beta for market model computation of CARs
- computes CARs
- extracts summary statistics for CARs
- creates tables that describe the sample (industry frequency distribution, deal-specific characteristics cross-tab)

4. Output includes:
- ./output/transactions_for_analysis.xlsx
- ./output/CAR_summary_statistics.xlsx
- ./output/deal_characteristics_table2.docx

### univariate.R
1. Load libraries.
2. Read in data.

- ./input/annual_bs_cleaned.xlsx
- ./output/transactions_for_analysis.xlsx

3. This script does the following: 
- filters balance sheet data to retain only necessary years
- calculates new variables for analysis
- winsorizes all variables for analysis
- calculates summary statistics, ANOVA
- assesses normality and homoskedasticity

4. Output includes:
- ./output/winsorized_annual_balance_sheet_data.xlsx
- ./output/annual_summary_stats.xlsx

### multivariate.R
1. Load libraries.
2. Read in data:

- ./output/winsorized_annual_balance_sheet_data.xlsx

3. This script does the following: 
- Creates and runs OLS regression model for full sample and individual CAR terciles
- Presents regression results and diagnostic tests for regression assumptions

* Note that regression output is only in the console (no output file). All regressions
are run on data from ./output/winsorized_annual_balance_sheet_data.xlsx.

