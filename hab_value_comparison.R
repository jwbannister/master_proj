library(tidyverse)
library(ggplot2)
library(readxl)
library(ggrepel)

factors_df <- read_excel(paste0("~/airsci/owens/Master Project & Cost Benefit/", 
                         "MP Workbook JWB 04-11-2018.xlsx"), range="A3:G34", 
                  sheet="Generic HV & WD")
colnames(factors_df) <- c('dcm', 'bw', 'mw', 'pl', 'ms', 'md', 'water')
factors_df$Avg_Habitat <- apply(factors_df[ , 2:6], 1, mean)
factors_df <- factors_df %>%
    gather(key="guild", value="value", bw:md)
factors_df$guild <- factor(factors_df$guild, 
                           levels=c("bw", "mw", "pl", "ms", "md"), ordered=T)
factors_df$hab2water <- factors_df$value / factors_df$water
    
dcm_dust <- c("SFP", "Veg 08", "Veg 11", "ENV", "SFL", "SFLS", "Gravel", 
              "Tillage", "Till-Brine", "Sand Fences", "Brine", "None")
dcm_habitat <- c("BWF", "MWF", "SNPL_realistic", "SNPL_with gravel", 
                 "MSB", "Meadow", "MWF and MSB", "MSB and SNPL", 
                 "MSB and SNPL_gravel", "MSB and SNPL_gravel_MWF", 
                 "MWF and SNPL", "MWF and SNPL_with gravel", 
                 "Breeding Waterfowl & Meadow")
dcm_dwm <- c("DWM_Jan", "DWM_Oct", "DWM_Dec", "DWM_Dust Control", 
             "DWM_Plovers", "DWM_Spring_only")

p1 <- factors_df %>% filter(DCM %in% dcm_dust) %>% arrange(guild) %>%
    ggplot(aes(x=guild, y=value, group=DCM)) +
    geom_path(aes(color=DCM))

