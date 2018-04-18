library(tidyverse)
library(ggplot2)
library(readxl)
library(ggrepel)

df1 <- read_excel(paste0("~/airsci/owens/Master Project & Cost Benefit/", 
                         "Master_proj_cleaned.xlsm"), sheet="Generic HV & WD", 
                  skip=2)
colnames(df1) <- c('DCM', 'BW', 'MW', 'PL', 'MS', 'MD', 'H2O', 'Avg_Habitat')
for (a in colnames(df1)[c(2:6, 8)]){
    p1 <- df1 %>%
        ggplot(aes_string(x='H2O', y=a)) +
        geom_point() +
        geom_label_repel(aes(label=DCM), size=2)
    png(paste0("~/Desktop/hab-water plots/", a, " vs. Water Demand.png"), 
        height=6, width=6, units="in", res=300)
    print(p1)
    dev.off()
}

    
                  

