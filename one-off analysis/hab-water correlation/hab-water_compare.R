library(tidyverse)
library(ggplot2)
df1 <- read_csv("~/Desktop/habitat-water comparison.csv")[ , 1:9]

guild <- 'ms'
p1 <- df1 %>% filter(type=='Design (generic)') %>%
    ggplot(aes_string(x='water', y=guild)) +
    geom_point(aes(color=type))

