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
library(dunn.test)

############## STEP 1: IMPORT AND CLEAN DATA, FILTER SAMPLE ####################

# Import data
# financial_data has total index returns in the first column associated with an ISIN and unadjusted daily closing prices in the ISIN.2 column
transactions <- read_excel("./input/Zephyr_Definitive_Repeating.xlsx", sheet = "Results")
financial_data <- read_excel("./input/unique_acquiror_isins2.xlsx", sheet = "tri and price", .name_repair = "minimal")
auditor_data <- read_excel("./input/robustness_tests_balance_sheet.xlsx")
crsp_data <- read_csv("./input/financial_data_CRSP.csv")

# Count initial number of transactions and unique ISINs in Zephyr data

starting_unique_transaction_count <- transactions %>%
  dplyr::summarize(unique_transactions = n_distinct(`Deal Number`))
print( starting_unique_transaction_count) #1575 expected for Zephyr transactions 

starting_unique_firm_count_transactions <- transactions %>%
  dplyr::summarize(unique_values = n_distinct(`Acquiror ISIN number`))
print(starting_unique_firm_count_transactions) #1230 expected for Zephyr transactions

remove(starting_unique_firm_count_transactions, starting_unique_transaction_count)

# Clean and prepare data

# Avoid displaying values in scientific notation, for readability
options(scipen = 999) 

############# financial_data ####################################################
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
colnames(financial_data) <- new_colnames # ISIN column is total return index; ISIN.2 is unadjusted daily share price

# Set financial data class to numeric
for (i in 2:ncol(financial_data)) {
  if (class(financial_data[[i]]) == "character") {
    
    # Remove non-numeric characters except the decimal point and negative sign
    financial_data[[i]] <- gsub("[^0-9.-]", "", financial_data[[i]])
    
    # Convert the cleaned strings to numeric
    financial_data[[i]] <- as.numeric(financial_data[[i]])
  }
} #Warnings due to NA values expected here. Relevant NAs will be dealt with when financial data is filtered for clean period.

# Convert date columns to correct format without timestamp
financial_data$Date <- as.Date(financial_data$Date, format = "%Y/%m/%d")
 
######################### transaction data #####################################

# Convert date columns to correct format without timestamp
transactions$`Announced date` <-as.Date(transactions$`Announced date`, format = "%Y/%m/%d")
transactions$`Acquiror IPO date` <-as.Date(transactions$`Acquiror IPO date`, format = "%d/%m/%Y")
transactions$`Completed date` <- as.Date(transactions$`Completed date`, format = "%Y/%m/%d")
transactions$`Assumed completion date` <- as.Date(transactions$`Assumed completion date`, format = "%Y/%m/%d")

# Deal method of payment value as numeric
# Warnings (NAs introduced by coercion) refer to empty or n.a. cells with no deal method of payment value information
transactions$`Deal method of payment value th USD` <- as.numeric(transactions$`Deal method of payment value th USD`)

# Indicate announcement year
transactions <- transactions %>%
  mutate(year_of_announcement = as.integer(year(`Announced date`)))


############################################## FIRST, APPLY BASIC CRITERIA FILTER  #########################################################################

# Remove if non-US acquirer country code
transactions <- transactions %>%
  filter(`Acquiror country code` == "US")

# Remove if deal equity value < $10 mil --------------> Consider removing this filter to increase sample size
transactions <- transactions %>%
  filter(`Deal equity value th USD` >= 10000) 

# Count 1 for Selection Process table
original_transaction_count <- transactions %>%
  dplyr::summarize(unique_transactions = n_distinct(`Deal Number`))
original_acquirer_count <- transactions %>%
  dplyr::summarize(unique_acquirers = n_distinct(`Acquiror ISIN number`))
print(original_transaction_count) #1495 expected for Zephyr transactions
print(original_acquirer_count) #1161 expected for Zephyr transactions



############################################## NEXT, APPLY SIC CODE FILTER  #########################################################################

# Remove transaction if acquirer or target SIC code starts with "49" or "6"
transactions <- transactions %>% 
  arrange(`Deal Number`, `Announced date`) %>% 
  group_by(`Deal Number`) %>%
  fill(`Target primary US SIC code`, .direction = "downup") %>% # Fills in missing SIC numbers for the same ISIN
  fill(`Acquiror primary US SIC code`, .direction = "downup") %>% # for both acquirer and target firms
  ungroup() %>%
  mutate( 
    `Target primary US SIC code` = substring(`Target primary US SIC code`, 1, 2), # Reduces 4-digit SIC code to 2 digits
    `Acquiror primary US SIC code` = substring(`Acquiror primary US SIC code`, 1, 2) # for both acquirer and target firms
  )%>%
  filter(
    !(`Target primary US SIC code` == "49") & #----------> ROBUSTNESS: BIGGER SAMPLE
      !(`Acquiror primary US SIC code` == "49") & 
      !(`Target primary US SIC code` >= "60" & `Target primary US SIC code` <= "69") & #----------> ROBUSTNESS: BIGGER SAMPLE
      !(`Acquiror primary US SIC code` >= "60" & `Acquiror primary US SIC code` <= "69")
  )

# Count 2 for Selection Process table
sic_filtered_unique_transaction_count <- transactions %>%
  dplyr::summarize(unique_transactions = n_distinct(`Deal Number`))
sic_filtered_unique_firm_count_transactions <- transactions %>%
  dplyr::summarize(unique_values = n_distinct(`Acquiror ISIN number`))

print(sic_filtered_unique_transaction_count) # 956 expected count after filter
print(sic_filtered_unique_firm_count_transactions) # 774 expected count after filter


############################################## NEXT, APPLY EARN-OUTS FILTER  #########################################################################

# Remove deal number if earnout appears in 'Deal method of payment'
deal_numbers_with_earnout <- transactions %>%
  filter(grepl("Earn-out", `Deal method of payment`, ignore.case = TRUE)) %>%
  distinct(`Deal Number`)

transactions <- transactions %>%
  filter(!(`Deal Number` %in% deal_numbers_with_earnout$`Deal Number`))


# Count 3 for Selection Process table
no_earnouts_transaction_count <- transactions %>%
  dplyr::summarize(unique_transactions = n_distinct(`Deal Number`))
no_earnouts_acquirer_count <- transactions %>%
  dplyr::summarize(unique_acquirers = n_distinct(`Acquiror ISIN number`))
print(no_earnouts_transaction_count) #891 expected for Zephyr transactions
print(no_earnouts_acquirer_count) #730 expected for Zephyr transactions



############################################## NEXT, APPLY METHOD OF PAYMENT FILTER  #########################################################################

# Combining the relevant categories for "Cash"
percentages <- transactions %>%
  mutate(
    Cash = case_when(
     `Deal method of payment` %in% c("Cash", "Cash assumed", "Cash Reserves", "Deferred payment", "Bonds", "Dividend") ~ `Deal method of payment value th USD`,
      TRUE ~ 0),
    Shares = ifelse(`Deal method of payment` == "Shares", `Deal method of payment value th USD`, 0)
    ) %>%
  group_by(`Deal Number`) %>%
  summarise(
    Total_Cash = sum(Cash, na.rm = TRUE),
    Total_Shares = sum(Shares, na.rm = TRUE),
    Total_Equity_Value = first(`Deal equity value th USD`)
  ) %>%
  # Calculate the percentages
  mutate(
    Cash_percentage = (Total_Cash / Total_Equity_Value),
    Shares_percentage = (Total_Shares / Total_Equity_Value)
  ) %>%
  ungroup() 

transactions <- transactions %>%
  left_join(percentages, by = "Deal Number")  

# Replace NaN with zeros in case of any divisions by zero, and round to 2 digits
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

# Calculate the frequency of different payment methods  
payment_stats <- transactions %>%
  dplyr::summarize(
    Cash_Only = n_distinct(`Deal Number`[Payment_Method == "Cash"]),
    Shares_Only = n_distinct(`Deal Number`[Payment_Method == "Shares"]),
    Mixed = n_distinct(`Deal Number`[Payment_Method == "Mixed"])
  )

pmt_filtered_50_50_mixed_transactions <- transactions %>%
  dplyr::summarise(unique_values = n_distinct(`Deal Number`[Shares_percentage == 0.5]))
print(pmt_filtered_50_50_mixed_transactions)



### Mixed payments classified as Cash or Shares
transactions <- transactions %>%
  mutate(Payment_Method = case_when(
    Cash_percentage >= .51 ~ "Cash",  ####################### change to CASH >= .5 for robustness test?
    Shares_percentage >= .50 ~ "Shares", 
        TRUE ~ NA_character_  # For any other case, it will set NA
  ))

# Re-Calculate the frequency of different payment methods  
payment_stats <- transactions %>%
  dplyr::summarize(
    Cash_Only = n_distinct(`Deal Number`[Payment_Method == "Cash"]),
    Shares_Only = n_distinct(`Deal Number`[Payment_Method == "Shares"]),
  )

# Count 4 for Selection Process table
pmt_filter_unique_transaction_count <- transactions %>%
  dplyr::summarize(unique_transactions = n_distinct(`Deal Number`))
print( pmt_filter_unique_transaction_count) #658 expected for Zephyr transactions

pmt_filter_unique_firm_count_transactions <- transactions %>%
  dplyr::summarize(unique_values = n_distinct(`Acquiror ISIN number`))
print(pmt_filter_unique_firm_count_transactions) #571 expected for Zephyr_transactions

transactions <- transactions %>%
  select(-`Deal method of payment value th USD`, -`Deal method of payment`, -`Deal enterprise value th USD`, -`Deal status`,`Deal type`)
############################################################################################################################################


# Write to Excel to retrieve financial data (stock prices, shares outstanding) for market cap calculation
#write.xlsx(unique_acquiror_isins, "unique_acquiror_isins.xlsx")

###################################### MATCH TRANSACTION DATA TO FINANCIAL DATA  #######################################

# Filter transactions data to keep only those transactions for which acquirer financial data is available ---> robustness test: bigger sample
transactions <- transactions %>%
  filter(`Acquiror ISIN number` %in% colnames(financial_data)) #crsp data for deal ratio calculation

## Count 5 for Selection Process table
matched_unique_transaction_count <- transactions %>%
  dplyr::summarize(unique_transactions = n_distinct(`Deal Number`))
print( matched_unique_transaction_count) #601 expected for Zephyr transactions

matched_unique_firm_count_transactions <- transactions %>%
  dplyr::summarize(unique_values = n_distinct(`Acquiror ISIN number`))
print(matched_unique_firm_count_transactions) #520 expected for Zephyr transactions



###########################################  CALCULATION OF MARKET CAP & DEAL RATIO  ###############################################################

### Calculate market cap 4 weeks before announcement
#  = shares outstanding 4 weeks before announcement * avg share price over previous 4 weeks

# Data pre-processing
transactions <- transactions %>%
  mutate(`Announced date` = as.Date(`Announced date`),
         end_date = `Announced date` - weeks(4),
         start_date = end_date - weeks(4),
         deal_equity_value_usd = `Deal equity value th USD` * 1000)

# Create a new df to avoid many-to-many merging issues
date_ranges <- transactions %>%
  select(`Acquiror ticker symbol`, `Deal Number`, start_date, end_date) %>%
  distinct(`Deal Number`, .keep_all = TRUE) %>%
  ungroup()

expanded_crsp <- crsp_data %>% 
  right_join(date_ranges, by = c("TICKER" = "Acquiror ticker symbol")) %>%
  group_by(TICKER) %>%
  mutate(date = as.Date(date)) %>%
  filter(date >= start_date & date <= end_date) # many-to-many relationship expected here

crsp_metrics <- expanded_crsp %>%
  group_by(TICKER, start_date, end_date) %>%
  summarize(
    pre_announcement_share_price = mean(PRC, na.rm = TRUE),
    shares_outstanding_at_end_date = SHROUT[which.max(date)] * 1000,
    .groups = 'drop'
  ) %>%
  mutate(
    market_cap_pre_announcement = pre_announcement_share_price * shares_outstanding_at_end_date
  )

merged_data <- transactions %>%
  left_join(crsp_metrics, by = c("Acquiror ticker symbol" = "TICKER", "start_date", "end_date")) %>%
    mutate( 
      deal_ratio = deal_equity_value_usd / market_cap_pre_announcement # Calculate deal ratio
    )

############### --> robustness testing: bigger sample
#market_cap_robustness <- read.csv("/Users/saraparker/Downloads/cpcsufliin12e66u.csv")

#market_cap_robustness$datadate <- as.Date(market_cap_robustness$datadate, format = "%Y-%m-%d")


# Joining datasets
#robust_check <- merged_data%>%
#  left_join(market_cap_robustness, by = c("Acquiror ticker symbol" = "tic" ,"end_date"  ="datadate" ))

#robust_check <- robust_check %>%
#  mutate( market_cap_pre_announcement = ifelse(is.na(market_cap_pre_announcement),
#                                               cshoc*1000*prccd,
#                                               market_cap_pre_announcement )) %>%
#  mutate( deal_ratio = deal_equity_value_usd / market_cap_pre_announcement )

################################################################################### ->>>> use ROBUST_CHECK for test!
#transactions <- robust_check

# Remove duplicate Deal Number rows, keeping only the first instance
#transactions <- transactions %>%
#  group_by(`Deal Number`) %>%
#  slice(1) %>%
#  ungroup()
############################


deal_ratios <- merged_data %>% #-------_> pause for robustness check
  select(`Acquiror ISIN number`, `Deal Number`, deal_ratio)

# Left join deal_ratios dataframe with the transactions dataframe
transactions <- transactions %>%
  left_join(deal_ratios, by = c("Acquiror ISIN number", "Deal Number")) %>%
  filter(!is.na(deal_ratio) & deal_ratio >= 0.05) 

transactions <- transactions %>%
  filter(!is.na(deal_ratio) )  ######## -------> robustness test: no min deal ratio

# Count 6 for Selection Process table
deal_ratio_filtered_unique_transaction_count <- transactions %>%
  dplyr::summarize(unique_transactions = n_distinct(`Deal Number`))
print( deal_ratio_filtered_unique_transaction_count) #284 expected for Zephyr transactions

deal_ratio_filtered_unique_firm_count_transactions <- transactions %>%
  dplyr::summarize(unique_values = n_distinct(`Acquiror ISIN number`))
print(deal_ratio_filtered_unique_firm_count_transactions) #261 expected for Zephyr transactions

#################################### SAMPLE IS NOW FILTERED FOR ALL DEAL CRITERIA ###############################################################

# Filter the transactions dataframe to include only those entries that are present in the final_sample
# (i.e., retain only those transactions with valid balance sheet data for analysis)
final_sample <- read_excel("./input/final_sample.xlsx")
final_sample$`Deal Number` <- as.character(final_sample$`Deal Number`)
filtered_transactions <- inner_join(transactions, final_sample, by = "Deal Number")

n_distinct(filtered_transactions$`Deal Number`) # 243 expected
n_distinct(filtered_transactions$`Acquiror ISIN number.x`) # 223 expected

# Remove deals with years 1 or 2 missing
# List of Deal Numbers to be removed
deal_numbers_to_remove <- c(1601106346, 1601238103, 1601264581, 1909095461, 1909135551,
                            1909262216, 1909301680, 1909542668, 1909566989, 1941025551,
                            1941205115, 1601193258, 1601238504, 1601246469, 1601421745,
                            1633013996, 1909048527, 1909073908, 1909103603, 1909260332,
                            1941148217, 1941172470)

# Filter out these Deal Numbers from the dataframe
filtered_transactions <- filtered_transactions %>%
  filter(!(`Deal Number` %in% deal_numbers_to_remove))

transactions <- filtered_transactions # would be good to clean up the columns here

n_distinct(transactions$`Deal Number`)
n_distinct(transactions$`Acquiror ISIN number.x`)

############### Prepare categorical variables : serial acquirer, target domicile, method of payment, industry relatedness, Big4, industry fixed effects ###################

# Serial acquirer
transactions <- transactions %>%
  arrange(`Acquiror ISIN number.x`, `Announced date`) %>% # Sort the transactions
  group_by(`Acquiror ISIN number.x`) %>%
  mutate(
    prev_date = lag(`Announced date`), # Get the date of the previous transaction
    next_date = lead(`Announced date`), # Get the date of the next transaction
    within_2_years = ifelse(
      (as.Date(`Announced date`) - as.Date(prev_date)) <= years(2) | 
        (as.Date(next_date) - as.Date(`Announced date`)) <= years(2),
      TRUE, 
      FALSE
    ),
    serial = ifelse(any(within_2_years, na.rm = TRUE), TRUE, FALSE) # Check if any transactions are within 2 years
  ) %>%
  ungroup()

# Check deals where the acquirer is serial
serial_acquisitions <- transactions %>%
  filter(serial) %>%  # Keep only rows where 'serial' is TRUE
  select(`Deal Number`, `Acquiror ticker symbol`, `Announced date`)

# Domestic vs Cross-border deals
transactions <- transactions %>%
  mutate(cross_border = if_else(`Target country code` != "US", TRUE, FALSE))

# Method of payment
transactions <- transactions %>%
  mutate(stock_deal = if_else(Payment_Method == "Shares", TRUE, FALSE))

# Industry relatedness
transactions <- transactions %>%
  mutate(
    industry_relatedness = `Acquiror primary US SIC code` == `Target primary US SIC code`
  )

# Big4 auditor ! 

auditor_data <- auditor_data %>%
  select(`Data Year - Fiscal`, `Ticker Symbol`, `Auditor`)

names(auditor_data) <- c("year_of_announcement", "Acquiror ticker symbol", "Auditor")

transactions <- transactions %>%
  group_by(`Deal Number`) %>%
  left_join(auditor_data, by=c("Acquiror ticker symbol", "year_of_announcement"))

transactions <- transactions %>%
  group_by(`Deal Number`) %>%
  mutate(is_Big4 = ifelse(Auditor %in% c(4, 5, 6, 7), TRUE, FALSE))



# Industry classification by 1 digit for fixed effects
transactions <- transactions %>%
  mutate( 
    industry = substring(`Acquiror primary US SIC code`, 1, 1) ) # single digit industry identifier



################################################ STEP 2: MARKET MODEL ESTIMATION #############################################################################

# Calculate log returns for RUA 3000 and firms
for (i in 2:ncol(financial_data)) {
  ratios <- financial_data[[i]] / lag(financial_data[[i]]) # ratio = (tri / previous-day tri)
  ratios[is.infinite(ratios) | is.na(ratios) | ratios <= 0] <- NA  # Replaces non-positive, NA, or infinite ratios with NA
  financial_data[[i]] <- log(ratios)  # Calculate log returns
}

financial_data <- financial_data %>% filter(`RUSSELL 3000 (EOD) - TOT RETURN IND` != 0) # eliminates holidays from data set


alpha_beta_estimates <- data.frame(ISIN=character(), # Initializes new df for storing alpha and beta coefficients
                                   DealNumber=character(), 
                                   Alpha=numeric(), 
                                   Beta=numeric(), 
                                   BlumeBeta=numeric(), 
                                   stringsAsFactors=FALSE)

transactions_to_eliminate <- character() # Initializes a vector to store deal numbers with no corresponding firm data to be filtered out

# Calculate clean period for each transaction [-250,-51] in relation to announcement date
for (i in 1:nrow(transactions)) { 
  announced_date <- transactions$`Announced date`[i]
  if(!(announced_date %in% financial_data$Date)) {
    adjusted_date <- announced_date   
    while(!(adjusted_date %in% financial_data$Date)) {
      adjusted_date <- adjusted_date + days(1) # Moves announcement date forward if it occurred on a non-trading day
    }
  } else {
    adjusted_date <- announced_date
  }
  match_row_index <- which(financial_data$Date == adjusted_date)
  start_index <- match_row_index - 250
  end_index <- match_row_index - 51
  firm_isin <- transactions$`Acquiror ISIN number.x`[i]  # Keeps ISIN for later analysis
  deal_number <- transactions$`Deal Number`[i]  # Extract Deal Number

  if (length(match_row_index) > 0 && start_index > 0 && end_index > 0 && start_index < end_index) {
    cat("Processing: ", deal_number, "\n") #debugging
    cat("ISIN: ", firm_isin, "\n")
    cat("start_index: ", start_index, ", end_index: ", end_index, "\n") #debugging
    
    if(firm_isin %in% colnames(financial_data)) {
      clean_data <- financial_data[start_index:end_index, ]
  
      # Count the number of non-NA rows available for the regression
      num_complete_cases <- sum(complete.cases(clean_data[[firm_isin]], clean_data[["RUSSELL 3000 (EOD) - TOT RETURN IND"]]))
      
      if (num_complete_cases > 0) {
        model_formula <- as.formula(sprintf("%s ~ `RUSSELL 3000 (EOD) - TOT RETURN IND`", firm_isin))
        model <- lm(model_formula, data=clean_data)
        
        Alpha <- coef(model)[1]
        Beta <- coef(model)[2]
        BlumeBeta <- Beta * (2/3) + (1/3)
        
        new_row <- data.frame(ISIN=firm_isin, 
                              DealNumber=deal_number, 
                              Alpha=Alpha, 
                              Beta=Beta, 
                              BlumeBeta=BlumeBeta, 
                              stringsAsFactors=FALSE)
        
        alpha_beta_estimates <- rbind(alpha_beta_estimates, new_row)
      } else {
        cat("Insufficient data for regression for deal", deal_number, "\n")
        transactions_to_eliminate <- c(transactions_to_eliminate, deal_number)
      }
    } else {
      cat("Column for ISIN", firm_isin, "not found in financial_data.\n")
      transactions_to_eliminate <- c(transactions_to_eliminate, deal_number)
    }
  } else {
    cat("Insufficient data to calculate period for", deal_number, "\n")
    transactions_to_eliminate <- c(transactions_to_eliminate, deal_number)
  }
}
 transactions <- transactions[!transactions$`Deal Number` %in% transactions_to_eliminate,]
 
 alpha_beta_estimates <- alpha_beta_estimates %>% # removes 5 deal numbers with duplicate rows
   distinct()

# Check 
summary(alpha_beta_estimates)

# Count remaining transactions with sufficient financial data for clean period --> to be updated after matching with balance sheet data for regression
sufficient_data_transaction_count <- transactions %>%
  dplyr::summarize(unique_transactions = n_distinct(`Deal Number`))
print(sufficient_data_transaction_count) #221

sufficient_data_firm_count <- transactions %>%
  dplyr::summarize(unique_values = n_distinct(`Acquiror ISIN number.x`))
print(sufficient_data_firm_count) #203

# Create a data frame where event windows are specified
event_window_data <- data.frame(ISIN=character(), # Initialize the data frame for storing event window data
                                DealNumber=character(),
                                Window=character(), 
                                Date=as.Date(character()), 
                                Firm_Return=numeric(), 
                                Market_Return=numeric(), 
                                stringsAsFactors=FALSE)
# Definition of event windows 
event_windows <- list(
  "[-1,+1]" = c(-1, 1)
# "[-1,+2]" = c(-1, 2),
#  "[-2,+2]" = c(-2, 2),
#  "[-3,+1]" = c(-3, 1), 
#  "[-3,+3]" = c(-3, 3),
#  "[-5,+1]" = c(-5, 1),
#  "[-5,+2]" = c(-5, 2), 
#  "[-21,+1]" = c(-21, 1)
)


# Loop through each transaction to extract and append the event window data 
for(i in 1:nrow(transactions)) {  
  firm_isin <- transactions$`Acquiror ISIN number.x`[i]
  announcement_date <- transactions$`Announced date`[i]
  deal_number <- transactions$`Deal Number`[i]

  
  # Check if announcement date is a trading day
  adjusted_date <- announcement_date
  if(!(announcement_date %in% financial_data$Date)) { ######THIS SHOULD ALREADY BE ADDRESSED BY PREVIOUS STEP 
    while(!(adjusted_date %in% financial_data$Date)) {
      adjusted_date <- adjusted_date + days(1) # Adjust for non-trading days
    }
 }  
  # Process event windows based on adjusted_date
  announcement_row_indices <- which(financial_data$Date == adjusted_date)
  if(length(announcement_row_indices) > 0) {
    announcement_row_index <- announcement_row_indices[1]  # Use the first match
    for(window_name in names(event_windows)) {
      window_bounds <- event_windows[[window_name]]
      start_index <- max(1, announcement_row_index + window_bounds[1])
      end_index <- min(nrow(financial_data), announcement_row_index + window_bounds[2])
      
      if(start_index <= end_index) {
        window_data <- financial_data[start_index:end_index, ]
        for(j in 1:nrow(window_data)) {
          date_in_window <- window_data$Date[j]
          market_return <- window_data$`RUSSELL 3000 (EOD) - TOT RETURN IND`[j]
          firm_return <- ifelse(is.na(window_data[[firm_isin]][j]), NA, window_data[[firm_isin]][j])
          event_window_data <- rbind(event_window_data, data.frame(
            ISIN=firm_isin, 
            DealNumber=deal_number, 
            Window=window_name,  
            Date=date_in_window, 
            Firm_Return=firm_return, 
            Market_Return=market_return, 
            stringsAsFactors=FALSE))
        }
      } else {
        cat("Window indices out of bounds for ISIN", firm_isin, "\n")
      }
    }
  } else {
    cat("Announcement date not found in financial_data for ISIN", firm_isin, "\n")
  }
}

# Ensure event_window_data$Firm_Return is numeric before calculations
event_window_data$Firm_Return <- as.numeric(as.character(event_window_data$Firm_Return))


######################################### CARs computations #########################################

# Step 1: Join 'alpha_beta_estimates' with 'event_window_data' to get Alpha and Beta for each ISIN
event_window_data <- event_window_data %>%
  left_join(alpha_beta_estimates, by = c("ISIN", "DealNumber"))

# Step 2: Calculate expected returns for each row now that Alpha and Beta are part of 'event_window_data'
event_window_data <- event_window_data %>%
  mutate(Expected_Return = ifelse(is.na(Alpha) | is.na(Beta), NA, Alpha + Beta * Market_Return))

# check unique firms
transaction_count <- event_window_data %>%
  distinct(DealNumber) %>%
  nrow()
cat("Number of transactions: ", transaction_count, "\n")

# Calculate abnormal returns
event_window_data$AR <- event_window_data$Firm_Return - event_window_data$Expected_Return

# Calculate CAR for each firm in each event window
CARs <- event_window_data %>%
  group_by(ISIN, DealNumber, Window) %>%
  summarise(CAR = sum(AR, na.rm = TRUE), .groups = 'drop')

# Calculate terciles for CAR within each Window
CARs <- CARs %>%
  group_by(Window) %>%
  mutate(CAR_Tercile = ntile(CAR, 3)) %>%
  ungroup()

# Group CARs by event window and winsorize at 5% and 95% levels
CARs <- CARs %>%
  group_by(Window) %>%
  mutate(CAR_winsorized = Winsorize(CAR, probs = c(0.05, 0.95))) %>%
  ungroup()

# Add CAR data to transactions data frame 
CARs_data_from_window_1 <- CARs %>%
  filter(Window == "[-1,+1]")%>%
  select(DealNumber, CAR_Tercile, CAR_winsorized)
transactions <- transactions %>%
  left_join(CARs_data_from_window_1, by = c("Deal Number" = "DealNumber")) %>%
  distinct(`Deal Number`, .keep_all = TRUE) 

# write_xlsx(transactions, path = "./output/transactions_for_analysis.xlsx") # to use in univariate.R

################################################################################
###################### CARS COMPLETE ###########################################
########### DESCRIPTIVE STATISTICS OF CARS, SAMPLE COMPOSITION #################
################################################################################

# Calculating summary statistics for the full sample, for each event window
# (with CARs winsorized at 5% and 95% levels)
stats_full_sample <- CARs %>%
    group_by(Window) %>%
    dplyr::summarize(
      n = n_distinct(DealNumber),
      Mean = mean(CAR_winsorized, na.rm = TRUE),
      t_value = t.test(CAR_winsorized, mu = 0)$statistic, # one-sample t-test 
      t_p_value = t.test(CAR_winsorized, mu = 0)$p.value,
      SD = sd(CAR_winsorized, na.rm = TRUE),
      First_quartile = quantile(CAR_winsorized, 0.25, na.rm = TRUE),
      Median = median(CAR_winsorized, na.rm = TRUE),
      w_statistic = wilcox.test(CAR_winsorized, mu = 0, paired = FALSE)$statistic, #Wilcoxon signed-rank test
      w_p_value = wilcox.test(CAR_winsorized, mu = 0, paired = FALSE)$p.value,
      Third_quartile = quantile(CAR_winsorized,0.75, na.rm = TRUE),
      .groups = 'drop'
    )


# Calculate summary statistics for each tercile within each event window
stats_by_tercile <- CARs %>%
  group_by(Window, CAR_Tercile) %>%
  dplyr::summarize(
    n = n_distinct(DealNumber),
    Mean = mean(CAR_winsorized, na.rm = TRUE),
    t_value = t.test(CAR_winsorized, mu = 0)$statistic, # one-sample t-test 
    t_p_value = t.test(CAR_winsorized, mu = 0)$p.value,
    SD = sd(CAR_winsorized, na.rm = TRUE),
    First_quartile = quantile(CAR_winsorized, 0.25, na.rm = TRUE),
    Median = median(CAR_winsorized, na.rm = TRUE),
    w_statistic = wilcox.test(CAR_winsorized, mu = 0, paired = FALSE)$statistic, #Wilcoxon signed-rank test
    w_p_value = wilcox.test(CAR_winsorized, mu = 0, paired = FALSE)$p.value,
    Third_quartile = quantile(CAR_winsorized,0.75, na.rm = TRUE),
    .groups = 'drop'
  )

# Write both tables to Excel with different sheets
write_xlsx(list(Full_Sample = stats_full_sample, By_Tercile = stats_by_tercile), 
           path = "./output/CAR_summary_statistics.xlsx")

########################################### Z-test statistic computation to report instead of Wilcoxon w-score (for ease of analysis...)
wilcoxon_tercile_results <- CARs %>%
  group_by(Window, CAR_Tercile) %>%
  dplyr::summarize(
    w_statistic = wilcox.test(CAR_winsorized, mu = 0, paired = FALSE)$statistic,
    p_value = wilcox.test(CAR_winsorized, mu = 0, paired = FALSE)$p.value,
    sample_size = length(CAR_winsorized),  # Storing sample size for Z calculation
    .groups = 'drop'
  ) %>%
  mutate(
    mean_W = sample_size * (sample_size + 1) / 4,
    std_dev_W = sqrt(sample_size * (sample_size + 1) * (2 * sample_size + 1) / 24),
    z_statistic = (w_statistic - mean_W) / std_dev_W
  )

#############################################################################################################
# Sample characteristics 



##### For categorical variables: industry_relatedness, target_domicile, method_of_payment, big4

# Payment methods across terciles
payment_method_analysis <- transactions %>%
#  filter(Window == "[-1,+1]") %>%
  group_by(CAR_Tercile, Payment_Method) %>%
  summarise(Count = n_distinct(`Deal Number`), .groups = 'drop') %>%
  mutate(
    Total_Count = sum(Count, na.rm = TRUE),
    Percentage = Count / Total_Count )

# Industry relatedness across terciles
industry_relatedness_analysis <- transactions %>%
#  filter(Window == "[-1,+1]") %>%
  group_by(CAR_Tercile, industry_relatedness) %>%
  summarise(Count = n_distinct(`Deal Number`), .groups = 'drop') %>%
  mutate(
    Total_Count = sum(Count, na.rm = TRUE),
    Percentage = Count / Total_Count )

# Target domicile across terciles
target_domicile_analysis <- transactions %>%
#  filter(Window == "[-1,+1]") %>%
  group_by(CAR_Tercile, cross_border) %>%
  summarise(Count = n_distinct(`Deal Number`), .groups = 'drop') %>%
  mutate(
    Total_Count = sum(Count, na.rm = TRUE),
    Percentage = Count / Total_Count )

# Big4 auditor across terciles
big4_analysis <- transactions %>%
#  filter(Window == "[-1,+1]") %>%
  group_by(CAR_Tercile, is_Big4)%>%
  summarise(Count = n_distinct(`Deal Number`), .groups = 'drop')%>%
  mutate(
    Total_Count = sum(Count, na.rm = TRUE),
    Percentage = Count / Total_Count
  )

# Serial acquirer across terciles
serial_analysis <- transactions %>%
  #  filter(Window == "[-1,+1]") %>%
  group_by(CAR_Tercile, serial)%>%
  summarise(Count = n_distinct(`Deal Number`), .groups = 'drop')%>%
  mutate(
    Total_Count = sum(Count, na.rm = TRUE),
    Percentage = Count / Total_Count
  )


# Convert the payment_method_analysis to a data frame with proper column names
payment_method <- data.frame(
  "Deal Characteristic" = c("Payment Method", ""),
  "Category" = c("Cash", "Shares"),
  "Tercile1_Freq" = payment_method_analysis$Count[c(1, 2)],
  "Tercile1_Perc" = payment_method_analysis$Percentage[c(1, 2)],
  "Tercile2_Freq" = payment_method_analysis$Count[c(3, 4)],
  "Tercile2_Perc" = payment_method_analysis$Percentage[c(3, 4)],
  "Tercile3_Freq" = payment_method_analysis$Count[c(5, 6)],
  "Tercile3_Perc" = payment_method_analysis$Percentage[c(5, 6)]
)
# Convert the industry_relatedness_analysis to a data frame with proper column names
industry_relatedness <- data.frame(
  "Deal Characteristic" = c("Industry Relatedness", ""),
  "Category" = c("Same Industry", "Different Industries"),
  "Tercile1_Freq" = industry_relatedness_analysis$Count[c(2, 1)],
  "Tercile1_Perc" = industry_relatedness_analysis$Percentage[c(2, 1)],
  "Tercile2_Freq" = industry_relatedness_analysis$Count[c(4, 3)],
  "Tercile2_Perc" = industry_relatedness_analysis$Percentage[c(4, 3)],
  "Tercile3_Freq" = industry_relatedness_analysis$Count[c(6, 5)],
  "Tercile3_Perc" = industry_relatedness_analysis$Percentage[c(6, 5)]
)

# Convert the target_domicile_analysis to a data frame with proper column names
cross_border_deal <- data.frame(
  "Deal Characteristic" = c("Target Domicile", ""),
  "Category" = c("Cross-border", "Domestic"),
  "Tercile1_Freq" = target_domicile_analysis$Count[c(2, 1)],
  "Tercile1_Perc" = target_domicile_analysis$Percentage[c(2, 1)],
  "Tercile2_Freq" = target_domicile_analysis$Count[c(4, 3)],
  "Tercile2_Perc" = target_domicile_analysis$Percentage[c(4, 3)],
  "Tercile3_Freq" = target_domicile_analysis$Count[c(6, 5)],
  "Tercile3_Perc" = target_domicile_analysis$Percentage[c(6, 5)]
)

# Convert the big4_analysis to a data frame with proper column names
big4_auditor <- data.frame(
  "Deal Characteristic" = c("Big4 Auditor", ""),
  "Category" = c("Big4", "Non-Big4"),
  "Tercile1_Freq" = big4_analysis$Count[c(2, 1)],
  "Tercile1_Perc" = big4_analysis$Percentage[c(2, 1)],
  "Tercile2_Freq" = big4_analysis$Count[c(4, 3)],
  "Tercile2_Perc" = big4_analysis$Percentage[c(4, 3)],
  "Tercile3_Freq" = big4_analysis$Count[c(6, 5)],
  "Tercile3_Perc" = big4_analysis$Percentage[c(6, 5)]
)



# Add a total column to each of the data frames
payment_method$Total_Freq <- rowSums(payment_method[, c("Tercile1_Freq", "Tercile2_Freq", "Tercile3_Freq")], na.rm = TRUE)
payment_method$Total_Perc <- rowSums(payment_method[, c("Tercile1_Perc", "Tercile2_Perc", "Tercile3_Perc")], na.rm = TRUE)

industry_relatedness$Total_Freq <- rowSums(industry_relatedness[, c("Tercile1_Freq", "Tercile2_Freq", "Tercile3_Freq")], na.rm = TRUE)
industry_relatedness$Total_Perc <- rowSums(industry_relatedness[, c("Tercile1_Perc", "Tercile2_Perc", "Tercile3_Perc")], na.rm = TRUE)

cross_border_deal$Total_Freq <- rowSums(cross_border_deal[, c("Tercile1_Freq", "Tercile2_Freq", "Tercile3_Freq")], na.rm = TRUE)
cross_border_deal$Total_Perc <- rowSums(cross_border_deal[, c("Tercile1_Perc", "Tercile2_Perc", "Tercile3_Perc")], na.rm = TRUE)

big4_auditor$Total_Freq <- rowSums(big4_auditor[, c("Tercile1_Freq", "Tercile2_Freq", "Tercile3_Freq")], na.rm = TRUE)
big4_auditor$Total_Perc <- rowSums(big4_auditor[, c("Tercile1_Perc", "Tercile2_Perc", "Tercile3_Perc")], na.rm = TRUE)



# Combine all the above into one data frame
deal_characteristics <- rbind(payment_method, industry_relatedness, cross_border_deal, big4_auditor)

# Format the percentages
deal_characteristics <- deal_characteristics %>%
  mutate(across(ends_with("Perc"), ~paste0("(", round(. * 100, 2), ")")))

# Create the flextable object with the additional Total columns
flextable_obj <- flextable(deal_characteristics) %>%
  set_header_labels(
    Tercile1_Freq = "Tercile 1",
    Tercile1_Perc = "%",
    Tercile2_Freq = "Tercile 2",
    Tercile2_Perc = "%",
    Tercile3_Freq = "Tercile 3",
    Tercile3_Perc = "%",
    Total_Freq = "Total Count",
    Total_Perc = "Total %"
  ) %>%
  bold(part = "body", i = nrow(deal_characteristics))  # Set total row to bold

# Create a new Word document for the flextable
doc <- read_docx()

# Add the flextable to the Word document ######### This table is not well formatted.
doc <- body_add_flextable(doc, value = flextable_obj) 

# Save the Word document
print(doc, target = "./output/deal_characteristics_table2.docx")






##### SAMPLE DISTRIBUTION OVER TIME, BY METHOD OF PMT, DOMESTIC / CROSS-BORDER ########

######## METHOD OF PMT
# Summarize the data to count the number of each payment type per year ######### window???
transactions_summary <- transactions %>%
  arrange(announcement_year)%>%
  group_by(announcement_year,Payment_Method) %>%
  summarise(Frequency = n_distinct(`Deal Number`), .groups = 'drop')

# Arrange such that cash is on the bottom, mixed in the center, shares on top
transactions_summary$Payment_Method <- factor(transactions_summary$Payment_Method, levels = c("Shares", "Mixed", "Cash"))

p1 <- ggplot(transactions_summary, aes(x = as.factor(announcement_year), y = Frequency, fill = Payment_Method)) +
  geom_bar(stat = "identity", position = "stack", color = "black", width = 0.75) +
  scale_fill_manual(values = c("Shares" = "white", "Mixed" = "white", "Cash" = "grey40")) +
  labs(x = "Year", y = "Frequency", fill = "Payment Type") +
  ggtitle("Annual transaction frequency by payment method") +
  theme_minimal() +
  theme(
    text = element_text(family = "Arial", size = 7),
    panel.grid.major = element_blank(),
    #   panel.grid.minor = element_blank(),
    axis.text.x = element_text(angle = 90, vjust = 0.5, hjust=1),
    legend.key.size = unit(0.7, "lines"),
    plot.title = element_text(size = 8),
    legend.title = element_text(size = 7) 
  )


###### Domestic / Cross-border

# Summarize the data to count the number of each deal type per year
transactions_summary_deals <- transactions %>%
  arrange(announcement_year) %>%
  group_by(announcement_year, cross_border) %>%
  summarise(Frequency = n_distinct(`Deal Number`), .groups = 'drop')

# Arrange such that domestic is on the bottom, cross-border on top
transactions_summary_deals$cross_border <- factor(transactions_summary_deals$cross_border,
                                                  levels = c(TRUE, FALSE),
                                                  labels = c("Cross-border", "Domestic"))

# Now create the ggplot object for the new bar graph
p2 <- ggplot(transactions_summary_deals, aes(x = as.factor(announcement_year), y = Frequency, fill = cross_border)) +
  geom_bar(stat = "identity", position = "stack", color = "black", width = 0.75) +
  scale_fill_manual(values = c("Cross-border" = "white", "Domestic" = "grey40"), 
                    name = "Deal Type", 
                    labels = c("Cross-border", "Domestic")) +
  labs(x = "Year", y = "Frequency", fill = "Deal Type") +
  ggtitle("Annual transaction frequency by target domicile") +
  theme_minimal() +
  theme(
    text = element_text(family = "Arial", size = 7),
    panel.grid.major = element_blank(),
    axis.text.x = element_text(angle = 90, vjust = 0.5, hjust=1),
    legend.key.size = unit(0.7, "lines"),
    plot.title = element_text(size = 8),
    legend.title = element_text(size = 7)
  )

# Combine the graphs
grid.arrange(p1, p2, ncol=1)

# Re-Calculate the frequency of different payment methods  
payment_stats <- transactions %>%
  dplyr::summarize(
    Cash_Only = n_distinct(`Deal Number`[Payment_Method == "Cash"]),
    Shares_Only = n_distinct(`Deal Number`[Payment_Method == "Shares"]),
  )
##################################################### END OF SCRIPT #############################################

# Clean up transactions for further analysis
transactions <- transactions %>%
  select ( -`Deal type`, -`Deal status`, -`Deal enterprise value th USD`, -`Deal value th USD...9`,
          -`Deal value th USD...21`, -`Deal type`, -`Deal status`, -`Deal method of payment`, -`Deal method of payment value th USD`,
          -`Assumed completion date`, -`Completed date`, -`Initial stake (%)`, -`Acquired stake (%)`, -`Final stake (%)`, -start_date,
          -`Deal equity value th USD`, -Total_Cash, -Total_Shares, -Total_Equity_Value, -year_of_announcement, -`COMPANY FKEY`, `EVENT DATE`,...5,
          -prev_date,-end_date,-next_date,-within_2_years, -`EVENT DATE`, -...5, -`Acquiror IPO date`)

################################################################################
################################################################################
################### to be fixed : industry distribution table ##################
################################################################################



##### Table: Terciles by industry 

# 2-digit SIC code descriptions
sic_codes <- read_excel("./input/sic_2_digit_codes.xls")

# Ensure data types are correct
transactions$`Acquiror primary US SIC code` <- as.character(transactions$`Acquiror primary US SIC code`)
sic_codes$SIC_2_digit <- as.character(sic_codes$SIC_2_digit)
transactions$CAR_Tercile <- as.factor(transactions$CAR_Tercile)
#previously, SIC codes and terciles were as factors

# Join 2-digit SIC descriptions
transactions <- transactions %>%
  left_join(sic_codes, by = c("Acquiror primary US SIC code" = "SIC_2_digit"))

# Calculate the frequency of acquirers by industry and tercile
frequency_table <- transactions %>%
  group_by(`Acquiror primary US SIC code`, two_digit_industry_description, CAR_Tercile) %>%
  summarise(Frequency = n(), .groups = 'drop')

# Spread the data to have separate columns for each tercile
frequency_table <- frequency_table %>%
  pivot_wider(names_from = CAR_Tercile, values_from = Frequency, values_fill = list(Frequency = 0))

# Add total counts for each industry
totals <- frequency_table %>%
  summarise(across(c("1", "2", "3"), sum, na.rm = TRUE))

totals$`Acquiror primary US SIC code` <- "Totals"
totals$two_digit_industry_description <- "All Industries"

# Bind the totals row to the original frequency table
frequency_table_with_totals <- bind_rows(frequency_table, totals)


