library(readxl)
library(dplyr)
library(lubridate)
library(ggplot2)
library(tidyverse)
library(tidyr)
library(writexl) 
library(lubridate)
library(openxlsx)

#################################################### STEP 1: IMPORT AND CLEAN DATA, REFINE SAMPLE ################################################################

# Import data
# financial_data has total index returns in the first column associated with an ISIN and unadjusted daily closing prices in the ISIN.2 column
transactions <- read_excel("/Users/saraparker/Library/CloudStorage/OneDrive-Personal/KU Leuven/MBA - Accounting/Masters Thesis/data/Zephyr_transaction_data_split_repeating_DEFINITIVE_original.xlsx", sheet = "Results")
financial_data <- read_excel("/Users/saraparker/Library/CloudStorage/OneDrive-Personal/KU Leuven/MBA - Accounting/Masters Thesis/data/unique_acquiror_isins2.xlsx", sheet = "tri and price", .name_repair = "minimal")

# Rename columns with ISIN headers such that repeated ISINs can be uniquely identified
original_colnames <- colnames(financial_data)
new_colnames <- c()
for (col in original_colnames) {
  if (col %in% new_colnames) {
    count <- sum(grepl(paste0("^", col, "\\b"), new_colnames))
    new_name <- paste0(col, ".", count + 1)
  } else {
    new_name <- col
  }
  new_colnames <- c(new_colnames, new_name)
}
colnames(financial_data) <- new_colnames


# Count initial number of transactions and unique ISINs in Zephyr data
# Count unique transactions
starting_unique_transaction_count <- transactions %>%
  summarize(unique_transactions = n_distinct(`Deal Number`))
print( starting_unique_transaction_count) #1612 expected for Zephyr_transaction_data_split_repeating_DEFINITIVE_original.xlsx", sheet = "Results"  

# Count unique firms making transactions
starting_unique_firm_count_transactions <- transactions %>%
  summarize(unique_values = n_distinct(`Acquiror ISIN number`))
print(starting_unique_firm_count_transactions) #1257 expected for Zephyr_transaction_data_split_repeating_DEFINITIVE_original.xlsx", sheet = "Results"  

remove(starting_unique_firm_count_transactions, starting_unique_transaction_count, col, count, new_colnames, new_name, original_colnames)

# Clean and prepare data
# Convert date columns to correct format without timestamp
financial_data$Date <- as.Date(financial_data$Date, format = "%Y/%m/%d")
transactions$`Announced date` <-as.Date(transactions$`Announced date`, format = "%Y/%m/%d")
transactions$`Acquiror IPO date` <-as.Date(transactions$`Acquiror IPO date`, format = "%d/%m/%Y")
transactions$`Completed date` <- as.Date(transactions$`Completed date`, format = "%Y/%m/%d")
transactions$`Assumed completion date` <- as.Date(transactions$`Assumed completion date`, format = "%Y/%m/%d")
transactions$`Target IPO date` <- as.Date(transactions$`Target IPO date`, format = "%d/%m/%Y")

transactions <- transactions %>%
  mutate(year_of_announcement = year(`Announced date`))

# Apply criteria filters to transactions identified by Zephyr
transactions <- transactions %>% filter(`Acquiror country code` == "US") # Filter out non-US acquirers
transactions <- transactions %>% filter(`Deal equity value th USD` >= 10000) ################################## REMOVE $10mil FILTER FOR ROBUSTNESS CHECKS

filtered_unique_transaction_count <- transactions %>%
  summarize(unique_transactions = n_distinct(`Deal Number`))
print( filtered_unique_transaction_count) #1528 expected for Zephyr_transaction_data_split_repeating_DEFINITIVE_original.xlsx", sheet = "Results"  

filtered_unique_firm_count_transactions <- transactions %>%
  summarize(unique_values = n_distinct(`Acquiror ISIN number`))
print(filtered_unique_firm_count_transactions) #1184 expected for Zephyr_transaction_data_split_repeating_DEFINITIVE_original.xlsx", sheet = "Results"  

############################################## NEXT, APPLY SIC CODE FILTER  #########################################################################

# Apply SIC code filter and transform SIC codes to their first 2 digits
transactions <- transactions %>%
  # Create new columns with the first 2 digits of each SIC code
  mutate(
    `Target primary US SIC code` = substring(`Target primary US SIC code`, 1, 2),
    `Acquiror primary US SIC code` = substring(`Acquiror primary US SIC code`, 1, 2)
  ) %>%
  # Filter out rows where the transformed SIC code starts with "49" or "6"
  filter(
    !(`Target primary US SIC code` %in% c("49", "6")) & 
      !(`Acquiror primary US SIC code` %in% c("49", "6"))
  )

# Count unique transactions after SIC code filter
sic_filtered_unique_transaction_count <- transactions %>%
  summarize(unique_transactions = n_distinct(`Deal Number`))
print(sic_filtered_unique_transaction_count) # 1506 expected count after filter

# Count unique firms after SIC code filter
sic_filtered_unique_firm_count_transactions <- transactions %>%
  summarize(unique_values = n_distinct(`Acquiror ISIN number`))
print(sic_filtered_unique_firm_count_transactions) # 1128 expected count after filter

remove(filtered_unique_firm_count_transactions, filtered_unique_transaction_count, sic_filtered_unique_firm_count_transactions, sic_filtered_unique_transaction_count)

############################################## NEXT, APPLY METHOD OF PAYMENT FILTER  #########################################################################

transactions$`Deal method of payment value th USD` <- as.numeric(as.character(transactions$`Deal method of payment value th USD`))

# Process the data
# Combining the relevant categories for cash
percentages <- transactions %>%
  mutate(Cash = case_when(
    `Deal method of payment` %in% c("Cash", "Cash assumed", "Cash Reserves", "Deferred payment", "Bonds", "Dividend") ~ `Deal method of payment value th USD`,
    TRUE ~ 0
  )) %>%
  # Considering only "Shares" for shares payment
  mutate(Shares = ifelse(`Deal method of payment` == "Shares", `Deal method of payment value th USD`, 0)) %>%
  # Remove Liabilities from equity value as they are not included
  filter(!(`Deal method of payment` %in% c("Cash assumed", "Liabilities"))) %>% #########
group_by(`Deal Number`) %>%
  # Summing up all the cash and shares payments for each deal
  summarise(
    Total_Cash = sum(Cash, na.rm = TRUE),
    Total_Shares = sum(Shares, na.rm = TRUE),
    Total_Equity_Value = first(`Deal equity value th USD`)
  ) %>%
  # Calculate the percentages
  mutate(
    Cash_percentage = (Total_Cash / Total_Equity_Value),
    Shares_percentage = (Total_Shares / Total_Equity_Value)
  ) 
transactions <- transactions %>%
  left_join(percentages, by = "Deal Number")  

# Replace NaN with zeros in case of any divisions by zero
transactions <- transactions %>%
  mutate(
    Cash_percentage = round(replace_na(Cash_percentage, 0), digits = 2),
    Shares_percentage = round(replace_na(Shares_percentage, 0), digits = 2)
  )


# Filter out deals with 0% cash and 0% shares
transactions <- transactions %>%
  filter(Cash_percentage != 0 | Shares_percentage != 0) %>%
  filter((Cash_percentage + Shares_percentage) == 1)

transactions <- transactions %>%
  mutate(Payment_Method = case_when(
    Cash_percentage == 1 & Shares_percentage == 0 ~ "Cash",
    Cash_percentage == 0 & Shares_percentage == 1 ~ "Shares",
    Cash_percentage > 0 & Shares_percentage > 0 ~ "Mixed",
    TRUE ~ NA_character_  # For any other case, it will set NA
  ))

# Calculate the frequency of different payment methods ########################### TRANSFORM LATER TO BREAKDOWN PMT STATS BY TERCILE
payment_stats <- transactions %>%
  summarize(
    Cash_Only = n_distinct(`Deal Number`[Cash_percentage == 1 & Shares_percentage == 0]),
    Shares_Only = n_distinct(`Deal Number`[Shares_percentage == 1 & Cash_percentage == 0]),
    Mixed = n_distinct(`Deal Number`[Cash_percentage > 0 & Shares_percentage > 0])
  )

pmt_filtered_unique_transaction_count <- transactions %>%
  summarize(unique_transactions = n_distinct(`Deal Number`))
print( pmt_filtered_unique_transaction_count) #1100 expected for Zephyr_transaction_data_split_repeating_DEFINITIVE_original.xlsx", sheet = "Results"  

pmt_filtered_unique_firm_count_transactions <- transactions %>%
  summarize(unique_values = n_distinct(`Acquiror ISIN number`))
print(pmt_filtered_unique_firm_count_transactions) #896 expected for Zephyr_transaction_data_split_repeating_DEFINITIVE_original.xlsx", sheet = "Results"  

# Exclude transactions with mixed payments where both methods are not at least 10%
transactions <- transactions %>%
  group_by(`Deal Number`) %>%
  filter(!(Payment_Method == "Mixed" & (Shares_percentage < .1 | Cash_percentage < .1))) %>%
  ungroup()

pmt_filter2_unique_transaction_count <- transactions %>%
  summarize(unique_transactions = n_distinct(`Deal Number`))
print( pmt_filter2_unique_transaction_count) #1073 expected for Zephyr_transaction_data_split_repeating_DEFINITIVE_original.xlsx", sheet = "Results"  

pmt_filter2_unique_firm_count_transactions <- transactions %>%
  summarize(unique_values = n_distinct(`Acquiror ISIN number`))
print(pmt_filter2_unique_firm_count_transactions) #879 expected for Zephyr_transaction_data_split_repeating_DEFINITIVE_original.xlsx", sheet = "Results"  

remove(pmt_filter2_unique_firm_count_transactions, pmt_filter2_unique_transaction_count, pmt_filtered_unique_firm_count_transactions, pmt_filtered_unique_transaction_count)

# Extract unique Acquiror ISINs to look for financial data in Datastream
#unique_acquiror_isins <- transactions %>%
# select(`Acquiror ISIN number`) %>%
#  distinct()

# Write to CSV
#write.csv(unique_acquiror_isins, "unique_acquiror_isins.csv", row.names = FALSE)
# Write to Excel
#write.xlsx(unique_acquiror_isins, "unique_acquiror_isins.xlsx")


##################### MATCH TRANSACTION DATA TO FINANCIAL DATA VIA ACQUIROR ISIN ######################

transactions <- transactions %>%
  filter(`Acquiror ISIN number` %in% colnames(financial_data)) # Filter transactions data to keep only those transactions for which acquirer financial data is available


matched_unique_transaction_count <- transactions %>%
  summarize(unique_transactions = n_distinct(`Deal Number`))
print( matched_unique_transaction_count) #991 expected for Zephyr_transaction_data_split_repeating_DEFINITIVE_original.xlsx", sheet = "Results"  

matched_unique_firm_count_transactions <- transactions %>%
  summarize(unique_values = n_distinct(`Acquiror ISIN number`))
print(matched_unique_firm_count_transactions) #820 expected for Zephyr_transaction_data_split_repeating_DEFINITIVE_original.xlsx", sheet = "Results"  

remove(matched_unique_firm_count_transactions, matched_unique_transaction_count)

######################################  CALCULATION OF MARKET CAP & DEAL RATIO  ###############################################################

# Import data
market_cap <- read_excel("/Users/saraparker/Library/CloudStorage/OneDrive-Personal/KU Leuven/MBA - Accounting/Masters Thesis/data/Zephyr_transaction_data_one_line_per_deal_DEFINITIVE_original.xlsx", sheet = "Results")
crsp_data <- read_csv("/Users/saraparker/Library/CloudStorage/OneDrive-Personal/KU Leuven/MBA - Accounting/Masters Thesis/data/financial_data_938_CRSP.csv")

### Calculate market cap 4 weeks before announcement
#  = shares outstanding 4 weeks before announcement * avg share price over previous 4 weeks

calculate_market_cap <- function(market_cap, crsp_data) {
  
  # Data cleaning
  market_cap <- market_cap %>%
    mutate(`Announced date` = as.Date(`Announced date`)) %>%
    group_by(`Acquiror ISIN number`, `Acquiror ticker symbol`, `Deal Number`) %>% #grouping addresses serial acquirers
    mutate(
      end_date = `Announced date` - weeks(4), # Calculates the relevant period for market capitalization calculation
      start_date = end_date - weeks(4),
      deal_equity_value_usd = `Deal equity value th USD` * 1000 # Necessary to preserve the deal equity value
    )%>%
    ungroup()
  
  # Join with crsp_data and calculate market cap, deal ratio
  merged_data <- market_cap %>%
    left_join(crsp_data, by = c("Acquiror ticker symbol" = "TICKER")) %>%
    filter(date >= start_date & date <= end_date) %>%
    group_by( `Acquiror ISIN number`, `Acquiror ticker symbol`, `Deal Number`, end_date) %>%  
    summarize(
      pre_announcement_share_price = mean(PRC, na.rm = TRUE), # Determine the average share price
      shares_outstanding_at_end_date = SHROUT[which.max(date)]*1000,  # and shares outstanding at end_date
      deal_equity_value_usd = first(deal_equity_value_usd), # Ensures deal equity value is carried through the summarize function
      .groups = 'drop' 
    ) %>%
    mutate( 
      market_cap_pre_announcement = pre_announcement_share_price * shares_outstanding_at_end_date, # Calculate market cap
      deal_ratio = deal_equity_value_usd / market_cap_pre_announcement # Calculate deal ratio
    ) %>%
    ungroup()
  
  return(merged_data)
}

final_data <- calculate_market_cap(market_cap, crsp_data)

# Join final_data with transactions df
transactions <- transactions %>%
  left_join(final_data, by = c("Acquiror ISIN number", "Acquiror ticker symbol", "Deal Number")) %>%
  filter(deal_ratio >= 0.05)

deal_ratio_filtered_unique_transaction_count <- transactions %>%
  summarize(unique_transactions = n_distinct(`Deal Number`))
print( deal_ratio_filtered_unique_transaction_count) #313 expected for Zephyr_transaction_data_split_repeating_DEFINITIVE_original.xlsx", sheet = "Results"  

deal_ratio_filtered_unique_firm_count_transactions <- transactions %>%
  summarize(unique_values = n_distinct(`Acquiror ISIN number`))
print(deal_ratio_filtered_unique_firm_count_transactions) #286 expected for Zephyr_transaction_data_split_repeating_DEFINITIVE_original.xlsx", sheet = "Results"  

