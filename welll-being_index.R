# 1. Load packages --------------------------------------------------------
pacman::p_load(tidyverse, readxl, magrittr, openxlsx)


# import data -------------------------------------------------------------

bns <- read_excel("G:/My Drive/Data/BAF/BNS/ke_BNS_analysis_2023/ke_BNS_survey_NEW_-_all_versions_-_False_-_2024-02-06-08-39-40.xlsx") %>% 
  rename(household_ID = "_id")

# remove 'bns_matrix_' from column names
names(bns) <- sub("bns_matrix_", "", names(bns))

# subset data to remain with relevant variables
bns_1 <- bns %>% 
  mutate(household_size = male + female + children) %>% 
  select(household_ID, village, membership, gender, age, household_size, education, 
  meals_possess, meals_necessary, stove_possess, stove_necessary, clean_water_possess, clean_water_necessary, 
  fish_market_possess, fish_market_necessary, loans_possess, loans_necessary, school_possess, school_necessary, 
  boat_possess, boat_necessary, business_possess, business_necessary, chair_possess, chair_necessary, internet_possess, 
  internet_necessary, land_possess, land_necessary, livestock_possess, livestock_necessary, health_facility_possess, 
  health_facility_necessary, police_station_possess, police_station_necessary, clothes_possess, clothes_necessary, 
  playground_possess, playground_necessary, mill_possess, mill_necessary, transport_possess, transport_necessary, 
  toilet_possess, toilet_necessary, broom_possess, broom_necessary, mattress_possess, mattress_necessary, pot_possess, 
  pot_necessary, electricity_possess, electricity_necessary, fan_possess, fan_necessary, freezer_possess, freezer_necessary,
  piped_water_possess, piped_water_necessary, roof_possess, roof_necessary, radio_possess, radio_necessary, tv_possess, 
  tv_necessary, dishes_possess, dishes_necessary, iron_box_possess, iron_box_necessary, phone_possess, phone_necessary, phone_possess_, 
  phone_necessary_) %>% 
  rename(cleanwater_possess = clean_water_possess,
         cleanwater_necessary = clean_water_necessary,
         market_possess = fish_market_possess,
         market_necessary = fish_market_necessary,
         healthfacility_possess = health_facility_possess,
         healthfacility_necessary = health_facility_necessary,
         policestation_possess = police_station_possess,
         policestation_necessary = police_station_necessary,
         pipedwater_possess = piped_water_possess,
         pipedwater_necessary = piped_water_necessary,
         ironbox_possess = iron_box_possess, 
         ironbox_necessary = iron_box_necessary,
         phone2_possess_ = phone_possess_,
         phone2_necessary_ = phone_necessary_)


# Pivot the data frame into long format
bns.clean <- pivot_longer(bns_1,
                        cols = starts_with(c("meals_", "stove_", "cleanwater_", "market_", "loans_",
                                             "school_", "boat_", "business_", "chair_", "internet_", "land_", 
                                             "livestock_", "healthfacility_", "policestation_", "clothes_", 
                                             "playground_", "mill_", "transport_", "toilet_", "broom_", "mattress_", 
                                             "pot_", "electricity_", "fan_", "freezer_", "pipedwater_", "roof_", "radio_", 
                                             "tv_", "dishes_", "ironbox_", "phone_", "phone2_", "phone2_")),
                        names_to = c("commodity", ".value"),
                        names_sep = "_")  

 # Calculate weight of each item separately 
bns.clean_1 <- bns.clean %>%
  group_by(commodity) %>% 
  summarise(weighting = sum(possess=='yes')/n())
  
# match up the weightings for the various commodities
bns.clean_2 <- merge(bns.clean, bns.clean_1, by = "commodity") %>% 
  select(2:8, 1,9:11)

# calculate the well-being score for each commodity for each household
bns.summary.hh <- bns.clean_2 %>% 
  filter(weighting > .5) %>% # filter only thos weighte above 50%
  group_by(household_ID, commodity) %>% 
  summarize(village = village,
            have_now = sum(possess == 'yes'),
            weighting = weighting,
            well_being_score = have_now * weighting) 

# calculate the well-being index for each household
bns.wellbeing <- bns.summary.hh %>% 
  summarize(village = village,
            commodity = commodity,
            maximum_score = sum(weighting, na.rm = TRUE),
            well_being_score_total = sum(well_being_score, na.rm = TRUE),
            well_being_index = well_being_score_total/maximum_score)
print(bns.wellbeing)

# overall well being index
overall.index <- bns.wellbeing %>% 
  ungroup() %>% 
  summarise(total_index = mean(well_being_index))

# well being index per village
village.index <- bns.wellbeing %>% 
  group_by(village) %>% 
  summarise(village_index = mean(well_being_index))

# BNS summary
bns.stats <- merge(village.index, overall.index)

# Choose a file path and name for your Excel file
excel_file_1 <- "index_results.xlsx"

excel_file_2 <- "bns_wellbeing_scores.xlsx"

# Create a new Excel workbook
wb <- createWorkbook()

# Add a worksheet to the workbook
addWorksheet(wb, "index_results")

addWorksheet(wb, "bns_wellbeing_scores")


# Write your data to the worksheet
writeData(wb, "index_results", bns.stats)

writeData(wb, "bns_wellbeing_scores", bns.summary.hh)

# Save the workbook to an Excel file
saveWorkbook(wb, "ke_BNS_2023.xlsx")

# # Print a message indicating successful export
# cat("Results exported to", excel_file, "\n")
