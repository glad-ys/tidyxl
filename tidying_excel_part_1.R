#' ---
#' title: "Load and clean Excel files using tidyxl and unpivotr part 1"
#' author: "Gladys Wojciechowska"
#' date: "22 June 2021"
#' ---

# * Load libraries -----
library(tidyverse)
library(unpivotr)
library(tidyxl)

# * Load data using tidyxl::xlsx_cells-----

test <- xlsx_cells("sample_data.xlsx")

head(test)
tail(test)

# How many Excel sheets do we have?

xlsx_sheet_names("sample_data.xlsx")

# Load the first sheet using two options

test_1a <- xlsx_cells("sample_data.xlsx", sheets = 1)
test_1b <- xlsx_cells("sample_data.xlsx", sheets = "Sample 1")

identical(test_1a, test_1b)

# * Explore the data (First sheet) ----
data_1 <- xlsx_cells("sample_data.xlsx", sheets = 1)
print(data_1 %>% filter(row == 9), width = Inf)

names(data_1)

# what kind of data types do we have in this sheet?

table(data_1$data_type)

# The selected variables from this sheet
data_1 %>% 
  select(row, col, data_type, numeric, character, local_format_id)

# Move header names to a dedicated column using unpivotr::behead -----

# First beheading
data_1 %>% 
  select(row, col, data_type, numeric, character, local_format_id) %>%
  behead("up", header_1)

# Second beheading

data_1 %>% 
  select(row, col, data_type, numeric, character, local_format_id) %>%
  behead("up", header_1) %>%
  behead("up", header_2) %>% 
  print(width = Inf)

# Last beheading

data_1 %>% 
  select(row, col, data_type, numeric, character, local_format_id) %>%
  behead("up", header_1) %>%
  behead("up", header_2) %>% 
  behead("up", header_3) %>% 
  print(width = Inf)

# Create a header column with the proper header names, then spatter

data_1 %>% 
  select(row, col, data_type, numeric, character, local_format_id) %>%
  behead("up", header_1) %>%
  behead("up", header_2) %>% 
  behead("up", header_3) %>% 
  mutate(header = case_when(header_1 == "ID" ~ "id",
                            header_1 == "History" ~ "history",
                            header_3 == "Test 1" ~ "biochem_1",
                            header_3 == "Test 2" ~ "biochem_2",
                            header_3 == "Test 3" ~ "biochem_3",
                            header_3 == "Test 4" ~ "biochem_4",
                            header_3 == "Test 5" ~ "biochem_5",
                            header_3 == "Test 6" ~ "biochem_6")) %>% 
  print(width = Inf)


data_1 <- data_1 %>% 
  select(row, col, data_type, numeric, character, local_format_id) %>%
  behead("up", header_1) %>%
  behead("up", header_2) %>% 
  behead("up", header_3) %>% 
  mutate(header = case_when(header_1 == "ID" ~ "id",
                            header_1 == "History" ~ "history",
                            header_3 == "Test 1" ~ "biochem_1",
                            header_3 == "Test 2" ~ "biochem_2",
                            header_3 == "Test 3" ~ "biochem_3",
                            header_3 == "Test 4" ~ "biochem_4",
                            header_3 == "Test 5" ~ "biochem_5",
                            header_3 == "Test 6" ~ "biochem_6")) %>% 
  select(row, data_type, numeric, character, header) %>%
  spatter(header) %>%
  select(row, id, history, everything())

# The clean data frame! Save as csv.

print(data_1, width = Inf)

write_csv(data_1, "data_1_part1.csv")
