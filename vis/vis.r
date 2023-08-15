###VISUALISE MODEL OUTPUTS###
# install.packages("tidyverse")
# library(tidyverse)
# install.packages("ggpubr")
library(dplyr)
library(ggplot2)
library(ggpubr)
library(tidyr)
library(readxl)
library(lemon)

####################POPULATION
folder <- dirname(rstudioapi::getSourceEditorContext()$path)

#mobile capex
path = file.path(folder, '..', "Oughton et al. (2023) DICE v1.0.0.xlsx")
data <- read_excel(path, sheet = "Pop", col_names = T)
data$ISO3 = NULL
data$`country_name` = NULL
data$`Population Sum` = NULL
colnames(data)[3] <- "1"
colnames(data)[4] <- "2"
colnames(data)[5] <- "3"
colnames(data)[6] <- "4"
colnames(data)[7] <- "5"
colnames(data)[8] <- "6"
colnames(data)[9] <- "7"
colnames(data)[10] <- "8"
colnames(data)[11] <- "9"
colnames(data)[12] <- "10"

data = data[!(data$`Income Group`=="-"),]
names(data)[names(data) == 'Income Group'] <- 'income_group'

# data$Region = NULL
names(data)[names(data) == 'Region'] <- 'region'
data = data %>% gather(decile, value, (3:12))
data$value <- as.numeric(data$value)

data$decile = factor(data$decile, 
                          levels=c(1,2,3,4,5,6,7,8,9,10),
                          labels=c('1','2','3','4','5','6','7','8','9','10')
)

data$income_group = factor(data$income_group, 
                    levels=c(
                      'Low Income Developing Countries',
                      'Emerging Market Economies',
                      'Advanced Economies'))

data$value = data$value / 1e6

plot1 = 
  ggplot(data, aes(x=decile, y=value, fill=region, group=region)) +
  geom_text(aes(label = round(after_stat(y),0), group = decile), 
    stat = 'summary', fun = sum, vjust = -.5, size=2.5) +
  geom_bar(stat="identity")  +              
  theme(legend.position = "bottom",
        axis.text.x = element_text(angle = 15, hjust=1)) +
  labs(title="(A) Population by Density Deciles for Income Groups and Regions",
       fill="Region",
       subtitle = bquote("Aggregated from WorldPop Global Mosaic 1 km^2 (2020)"),
       x="Population Density Decile \n(Decile 1 is the most densely populated, Decile 10 is the least densely populated)", 
       y='Population (Millions)') +
  scale_y_continuous(expand = c(0, 0), limits=c(0,1800)) +#
  scale_fill_viridis_d() + #direction = -1
  facet_wrap(~income_group) #direction = -1

write.csv(data, file.path(folder, 'final_data', 'fig_5a_area_data.csv')) 

data_pop = data

# perc_pop = data %>%
#   select(decile, income, population) %>% 
#   group_by(decile,income) %>%
#   summarize(population = sum(population)) 
# 
# write.csv(perc_pop, file.path(folder, 'percentages', 'perc_pop.csv')) 

# #export to folder
# path = file.path(folder, 'figures', 'population.tiff')
# tiff(path, units="in", width=10, height=5, res=300)
# print(plot1)
# dev.off()

####################AREA
folder <- dirname(rstudioapi::getSourceEditorContext()$path)

path = file.path(folder, '..', "Oughton et al. (2023) DICE v1.0.0.xlsx")
data <- read_excel(path, sheet = "Area", col_names = T)
data$ISO3 = NULL
data$`country_name` = NULL
data$`area_km2_sum` = NULL
colnames(data)[3] <- "1"
colnames(data)[4] <- "2"
colnames(data)[5] <- "3"
colnames(data)[6] <- "4"
colnames(data)[7] <- "5"
colnames(data)[8] <- "6"
colnames(data)[9] <- "7"
colnames(data)[10] <- "8"
colnames(data)[11] <- "9"
colnames(data)[12] <- "10"

data = data[!(data$`Income Group`=="-"),]
names(data)[names(data) == 'Income Group'] <- 'income_group'

# data$Region = NULL
names(data)[names(data) == 'Region'] <- 'region'
data = data %>% gather(decile, value, (3:12))
data$value <- as.numeric(data$value)

data$decile = factor(data$decile, 
                     levels=c(1,2,3,4,5,6,7,8,9,10),
                     labels=c('1','2','3','4','5','6','7','8','9','10')
)

data$income_group = factor(data$income_group, 
                           levels=c(
                             'Low Income Developing Countries',
                             'Emerging Market Economies',
                             'Advanced Economies'))

data = data[complete.cases(data), ]

data$value = data$value / 1e6

plot2 = ggplot(data, aes(x=decile, y=value, fill=region, group=region)) +
  geom_text(aes(label = round(after_stat(y),1), group = decile), 
            stat = 'summary', fun = sum, vjust = -.5, size=2.5) +
  geom_bar(stat="identity")  +              #
  theme(legend.position = "bottom",
        axis.text.x = element_text(angle = 15, hjust=1)) +
  labs(title="(B) Geographic Area by Density Deciles for Income Groups and Regions",
       fill="Region",
       subtitle = bquote("Aggregated from GADM Adminstrative Area Boundaries"),
       x="Population Density Decile \n(Decile 1 is the most densely populated, Decile 10 is the least densely populated)",
       y='Area (Millions km^2)') +
  scale_y_continuous(expand = c(0, 0), limits=c(0,40)) +#
  scale_fill_viridis_d() +
  facet_wrap(~income_group) #, scales = "free"

write.csv(data, file.path(folder, 'final_data', 'fig_5b_area_data.csv')) 

# #export to folder
# path = file.path(folder, 'figures', 'area.tiff')
# tiff(path, units="in", width=10, height=5, res=300)
# print(plot2)
# dev.off()

# perc_area = data %>%
#   select(decile, income, area) %>% 
#   group_by(decile,income) %>%
#   summarize(area = sum(area)) 
# 
# write.csv(perc_area, file.path(folder, 'percentages', 'perc_area.csv')) 

####################POP DENSITY
folder <- dirname(rstudioapi::getSourceEditorContext()$path)

path = file.path(folder, '..', "Oughton et al. (2023) DICE v1.0.0.xlsx")
data <- read_excel(path, sheet = "Pop", col_names = T)
data$`Population Sum` = NULL
data$`country_name` = NULL
colnames(data)[4] <- "1"
colnames(data)[5] <- "2"
colnames(data)[6] <- "3"
colnames(data)[7] <- "4"
colnames(data)[8] <- "5"
colnames(data)[9] <- "6"
colnames(data)[10] <- "7"
colnames(data)[11] <- "8"
colnames(data)[12] <- "9"
colnames(data)[13] <- "10"
data = data[!(data$`Income Group`=="-"),]
names(data)[names(data) == 'Income Group'] <- 'income_group'
names(data)[names(data) == 'Region'] <- 'region'
data = data %>% gather(decile, pop, (4:13))
data$pop <- as.numeric(data$pop)
data = data[complete.cases(data), ]

path = file.path(folder, '..', "Oughton et al. (2023) DICE v1.0.0.xlsx")
area <- read_excel(path, sheet = "Area", col_names = T)
area$`country_name` = NULL
area$`area_km2_sum` = NULL
colnames(area)[4] <- "1"
colnames(area)[5] <- "2"
colnames(area)[6] <- "3"
colnames(area)[7] <- "4"
colnames(area)[8] <- "5"
colnames(area)[9] <- "6"
colnames(area)[10] <- "7"
colnames(area)[11] <- "8"
colnames(area)[12] <- "9"
colnames(area)[13] <- "10"
area = area[!(area$`Income Group`=="-"),]
names(area)[names(area) == 'Income Group'] <- 'income_group'
names(area)[names(area) == 'Region'] <- 'region'
area = area %>% gather(decile, area, (4:13))
area$area <- as.numeric(area$area)
area = area[complete.cases(area), ]

data_pop_d = merge(data, area, by=c("ISO3", "decile","income_group","region"))
data_pop_d = select(data_pop_d, decile, income_group, region, pop, area)

data_pop_d = data_pop_d %>%
  group_by(decile, income_group, region) %>%
  summarize(pop = sum(pop, na.rm = TRUE),
            area = sum(area, na.rm = TRUE),
            )

data_pop_d$pop_d_km2 = data_pop_d$pop / data_pop_d$area

data_pop_d$decile = factor(data_pop_d$decile, 
                     levels=c(1,2,3,4,5,6,7,8,9,10),
                     labels=c('1','2','3','4','5','6','7','8','9','10')
)

data_pop_d$income_group = factor(data_pop_d$income_group, 
                           levels=c(
                             'Low Income Developing Countries',
                             'Emerging Market Economies',
                             'Advanced Economies'))

plot3 = ggplot(data_pop_d, aes(x=decile, y=pop_d_km2, fill=region, group=region)) +
  geom_bar(stat="identity")  +              #
  geom_text(aes(label = round(after_stat(y),0), group = decile), 
            stat = 'summary', fun = sum, vjust = -.5, size=2.5) +
  theme(legend.position = "bottom",
        axis.text.x = element_text(angle = 15, hjust=1)) +
  labs(title="(C) Population Density by Density Deciles for Income Groups and Regions",
       fill="Region",
       subtitle = bquote("Estimated from WorldPop and GADM Adminstrative Areas"),
       x="Population Density Decile \n(Decile 1 is the most densely populated, Decile 10 is the least densely populated)", 
       y='Persons per km^2') +
  scale_y_continuous(expand = c(0, 0), limits=c(0,6750)) +#
  scale_fill_viridis_d() +
  facet_wrap(~income_group) #, scales = "free"

write.csv(data, file.path(folder, 'final_data', 'fig_5c_pop_density_data.csv')) 

# #export to folder
# path = file.path(folder, 'figures', 'pop_density.tiff')
# tiff(path, units="in", width=10, height=5, res=300)
# print(plot3)
# dev.off()

# perc_pop_d = data_pop_d %>%
#   select(decile, income, population, area) %>% 
#   group_by(decile,income) %>%
#   summarize(
#     population = sum(population),
#     area = sum(area)
#     )
# 
# write.csv(perc_pop_d, file.path(folder, 'percentages', 'perc_pop_d.csv')) 


####################GDP
folder <- dirname(rstudioapi::getSourceEditorContext()$path)

path = file.path(folder, '..', "Oughton et al. (2023) DICE v1.0.0.xlsx")
data <- read_excel(path, sheet = "GDP", col_names = T)
data$ISO3 = NULL
data$`country_name` = NULL
data = data[ , 1:12]
data = data[!(data$`Income Group`=="-"),]
names(data)[names(data) == 'Income Group'] <- 'income_group'

data$`2021` = NULL
names(data)[names(data) == 'Region'] <- 'region'
data = data %>% gather(year, value, (3:11))
data$value <- as.numeric(data$value)
data = data[complete.cases(data), ]

data$income_group = factor(data$income_group, 
                         levels=c(
                           'Low Income Developing Countries',
                           'Emerging Market Economies',
                           'Advanced Economies'))

data = data %>%
  group_by(year, income_group, region) %>%
  summarize(
    value = sum(value, na.rm=TRUE)
  )

data$value = data$value / 1000 

plot4 = ggplot(data, aes(x=year, y=value, fill=region, group=region)) +
        geom_text(aes(label = round(after_stat(y),1), group = year), 
            stat = 'summary', fun = sum, vjust = -.5, size=2.5) +
          geom_bar(stat="identity")  +              #
          theme(legend.position = "bottom",
                axis.text.x = element_text(angle = 15, hjust=1)) +
          labs(title="(D) Gross Domestic Product (GDP) by Income Group and Region",
               fill="Region",
               subtitle = "IMF Forecasts for Member Countries (2022-2030)",
               x=NULL, y='GDP ($Tn)') +
          scale_y_continuous(expand = c(0, 0), limits=c(0,72)) +#
          scale_fill_viridis_d() + 
          facet_wrap(~income_group) #, scales = "free"

# #export to folder
# path = file.path(folder, 'figures', 'gdp_by_income_and_region.tiff')
# tiff(path, units="in", width=10, height=5, res=300)
# print(plot1)
# dev.off()

# income_data = income

# perc_gdp = income %>%
#   select(year, income, gdp) %>% 
#   group_by(year,income) %>%
#   summarize(
#     gdp = sum(gdp)
#   )
# 
# write.csv(perc_gdp, file.path(folder, 'percentages', 'perc_gdp.csv')) 

context <- ggarrange(plot1, plot2, plot3, #plot4, 
                     ncol = 1, nrow = 3, align = c("hv"),
                     common.legend = TRUE, legend='bottom')

path = file.path(folder, 'figures', 'context.svg')
dir.create(file.path(folder, 'figures'), showWarnings = FALSE)
tiff(path, units="in", width=10, height=10, res=300)
# svg(path,width=10, height=10)
print(context)
dev.off()


######################################
####################AGGREGATED RESULTS
folder <- dirname(rstudioapi::getSourceEditorContext()$path)

path = file.path(folder, '..', "Oughton et al. (2023) DICE v1.0.0.xlsx")
data <- read_excel(path, sheet = "Cost_by_IMF_Income_Group", col_names = T)
names(data) <- data[1,]
data <- data[-1,]
data = data[, 1:4]
data = data[1:6,]

data = data %>% gather(income, cost, -Category)
data$cost <- as.numeric(data$cost)

data$Category = factor(data$Category,
                       levels=c('Mobile Infra Capex', 'Metro+backbone fiber',
                                'Mobile Infra Opex','Remote Coverage','Policy/Regulation',
                                'ICT Skills/Content'),
                         labels=c('Mobile Infra Capex', 'Metro+Backbone Fiber',
                                  'Mobile Infra Opex','Remote Coverage','Policy/Regulation',
                                  'ICT Skills/Content')
)

data$income = factor(data$income,
                       levels=c('Low Income Developing Countries',
                                'Emerging Market Economies',
                                'Advanced Economies')
)

plot1 = ggplot(data, aes(x=income, y=cost, fill=Category, group=Category)) +
  geom_bar(stat="identity") + coord_flip() +
  geom_text(aes(label = round(after_stat(y),0), group = income),
            stat = 'summary', fun = sum, hjust = -.5, size=2.5) +
  theme(legend.position = "bottom",
        axis.text.x = element_text(angle = 0, hjust=.5)) +
  labs(title="(A) Aggregate Cost for Income Groups By Asset Cost Type",
       fill="Cost\nType",
       subtitle = "Based on 50 GB/Month in EMEs and AEs, and 40 GB/Month in LIDCs",
       x=NULL, y='Investment Cost ($Bn)') +
  scale_y_continuous(expand = c(0, 0), limits=c(0,325)) +
  scale_fill_viridis_d(direction=-1) +
  guides(fill = guide_legend(reverse=T))

write.csv(data, file.path(folder, 'aggregated_income.csv'))
write.csv(data, file.path(folder, 'final_data', 'fig_6a_aggregated_income.csv'))

folder <- dirname(rstudioapi::getSourceEditorContext()$path)

path = file.path(folder, '..', "Oughton et al. (2023) DICE v1.0.0.xlsx")
data <- read_excel(path, sheet = "Cost_by_IMF_Regions", col_names = T)
names(data) <- data[1,]
data <- data[-1,]
data = data[, 1:8]
data = data[1:6,]

data = data %>% gather(region, cost, -Category)
data$cost <- as.numeric(data$cost)

data$Category = factor(data$Category,
                       levels=c('Mobile Infra Capex', 'Metro+backbone fiber',
                                'Mobile Infra Opex','Remote Coverage','Policy/Regulation',
                                'ICT Skills/Content'),
                       labels=c('Mobile Infra Capex', 'Metro+Backbone Fiber',
                                'Mobile Infra Opex','Remote Coverage','Policy/Regulation',
                                'ICT Skills/Content')
)

data$region = factor(data$region,
                     levels=c("Sub-Sahara Africa",
                              "Middle East, North Africa, Afghanistan, and Pakistan",
                              "Latin America and the Caribbean",
                              "Emerging and Developing Europe",
                              'Emerging and Developing Asia',
                              'Caucasus and Central Asia',
                              'Advanced Economies'
                              ),
                      labels=c("Sub-Saharan Africa",
                              "MENA, AFG and PAK",
                              "Latin America and the Caribbean",
                              "Emerging and Developing Europe",
                              'Emerging and Developing Asia',
                              'Caucasus and Central Asia',
                              'Advanced Economies'
                              )
)

plot2 = ggplot(data, aes(x=region, y=cost, fill=Category, group=Category)) +
  geom_text(aes(label = round(after_stat(y),0), group = region),
            stat = 'summary', fun = sum, hjust = -.5, size=2.5) +
  geom_bar(stat="identity")  +   coord_flip() +
  theme(legend.position = "bottom",
        axis.text.x = element_text(angle = 0, hjust=.5)) +
  labs(title="(B) Aggregate Cost for Regions By Asset Cost Type",
       fill="Cost\nType",
       subtitle = "Based on 50 GB/Month in EMEs and AEs, and 40 GB/Month in LIDCs",
       x=NULL, y='Investment Cost ($Bn)') +
  scale_y_continuous(expand = c(0, 0), limits=c(0,190)) +
  scale_fill_viridis_d(direction=-1) +
  guides(fill = guide_legend(reverse=T))

write.csv(data, file.path(folder, 'final_data', 'fig_6b_aggregated_regions.csv'))

aggregated_costs <- ggarrange(plot1, plot2,
                              ncol = 1, nrow = 2, align = c("hv"),
                              common.legend = TRUE, legend='bottom',
                              heights=c(.65,1))

path = file.path(folder, 'figures', 'aggregated_costs.tiff')
tiff(path, units="in", width=8, height=4.5, res=300)
print(aggregated_costs)
dev.off()

# # perc_regions = data %>%
# #   select(region, Category, cost) %>% 
# #   group_by(region, Category) %>%
# #   summarize(
# #     cost = sum(cost)
# #   )
# # 
# # write.csv(perc_regions, file.path(folder, 'percentages', 'perc_regions.csv')) 

####################Costs by Decile

folder <- dirname(rstudioapi::getSourceEditorContext()$path)

#mobile capex
path = file.path(folder, '..', "Oughton et al. (2023) DICE v1.0.0.xlsx")
capex <- read_excel(path, sheet = "Capex_Per_Decile", col_names = T)
capex$ISO3 = NULL
capex$`Country Name` = NULL
capex$Sum = NULL
capex$type = 'Mobile Infra Capex'

#fiber
metro_and_backbone_fiber <- read_excel(path, sheet = "Fiber_Per_Decile")
metro_and_backbone_fiber$ISO3 = NULL
metro_and_backbone_fiber$`Country Name` = NULL
metro_and_backbone_fiber$Sum = NULL
metro_and_backbone_fiber$type = 'Metro+Backbone Fiber'

#mobile opex
opex <- read_excel(path, sheet = "Opex_Per_Decile")
opex$ISO3 = NULL
opex$`Country Name` = NULL
opex$Sum = NULL
opex$type = 'Mobile Infra Opex'

#remote coverage
remote_coverage <- read_excel(path, sheet = "Remote_Coverage_Per_Decile")
remote_coverage$ISO3 = NULL
remote_coverage$`Country Name` = NULL
remote_coverage$Sum = NULL
remote_coverage$`itu_region` = NULL
remote_coverage$`wb_income` = NULL
remote_coverage$type = 'Remote Coverage'

#policy_regulation
policy_regulation <- read_excel(path, sheet = "Policy_Regulation_Per_Decile")
policy_regulation$ISO3 = NULL
policy_regulation$`Country Name` = NULL
policy_regulation$Sum = NULL
policy_regulation$type = 'Policy/Regulation'

#ict_skills
ict_skills <- read_excel(path, sheet = "ICT_Skills_Per_Decile")
ict_skills$ISO3 = NULL
ict_skills$`Country Name` = NULL
ict_skills$Sum = NULL
ict_skills$type = 'ICT Skills/Content'

data = rbind(capex, metro_and_backbone_fiber, opex, remote_coverage, policy_regulation, ict_skills)
remove(capex, metro_and_backbone_fiber, opex, remote_coverage, policy_regulation, ict_skills)
# write.csv(data, file.path(folder, 'test.csv')) 

data = na.omit(data)

colnames(data)[3] <- "D1"
colnames(data)[4] <- "D2"
colnames(data)[5] <- "D3"
colnames(data)[6] <- "D4"
colnames(data)[7] <- "D5"
colnames(data)[8] <- "D6"
colnames(data)[9] <- "D7"
colnames(data)[10] <- "D8"
colnames(data)[11] <- "D9"
colnames(data)[12] <- "D10"

data = data[!(data$`Income Group`=="-"),]
names(data)[names(data) == 'Income Group'] <- 'income_group'

data$type = factor(data$type,
                   levels=c('Mobile Infra Capex', 'Metro+Backbone Fiber',
                            'Mobile Infra Opex','Remote Coverage','Policy/Regulation',
                            'ICT Skills/Content')
)

income = data
income$Region = NULL
income = income %>% gather(decile, value, (2:11))
income$value <- as.numeric(income$value)
income = income %>%
  group_by(income_group, type, decile) %>%
  summarize(value = sum(value)/1e9)

income$income_group = factor(income$income_group,
                             levels=c('Low Income Developing Countries', 
                                      'Emerging Market Economies',
                                      'Advanced Economies')
)

income$decile = factor(income$decile,
                       levels=c('D1', 
                                'D2',
                                'D3',
                                'D4',
                                'D5',
                                'D6',
                                'D7',
                                'D8',
                                'D9',
                                'D10'
                       )
)

plot1 = ggplot(income, aes(x=decile, y=value, fill=type, group=type)) +
  geom_bar(stat="identity") + 
  geom_text(aes(label = round(after_stat(y),1), group = income_group),
            stat = 'summary', fun = sum, vjust = -.5, size=2.2) +
  theme(legend.position = "None",
        axis.text.x = element_text(angle = 45, vjust=.5, hjust=.5)) +
  labs(title="(A) Decile Cost for Income Groups By Asset Cost Type",
       fill="Cost\nType",
       subtitle = "Based on 50 GB/Month in EMEs and AEs, and 40 GB/Month in LIDCs",
       x=NULL, y='Cost ($Bn)') +
  scale_y_continuous(expand = c(0, 0), limits=c(0,55)) +
  scale_fill_viridis_d(direction=-1) +
  guides(fill = guide_legend(reverse=T)) +
  facet_wrap(~income_group, scales = "free")

write.csv(income, file.path(folder, 'final_data', 'fig_7a_income.csv'))

regions = data
regions$income_group = NULL
regions = regions %>% gather(decile, value, (2:11))
regions$value <- as.numeric(regions$value)
regions = regions %>%
  group_by(Region, type, decile) %>%
  summarize(value = sum(value)/1e9)

regions = regions %>% 
  filter(Region != 'Advanced Economies')

regions$Region = factor(regions$Region,
                        levels=c("Sub-Sahara Africa",
                                 "Middle East, North Africa, Afghanistan, and Pakistan",
                                 "Latin America and the Caribbean",
                                 "Emerging and Developing Europe",
                                 'Emerging and Developing Asia',
                                 'Caucasus and Central Asia'#,
                                 # 'Advanced Economies'
                                 ),
                        labels=c("Sub-Sahara Africa",
                                 "MENA, AFG and PAK",
                                 "Latin America and the Caribbean",
                                 "Emerging and Developing Europe",
                                 'Emerging and Developing Asia',
                                 'Caucasus and Central Asia'#,
                                 # 'Advanced Economies'
                                 )
)

regions$decile = factor(regions$decile,
                        levels=c('D1', 
                                 'D2',
                                 'D3',
                                 'D4',
                                 'D5',
                                 'D6',
                                 'D7',
                                 'D8',
                                 'D9',
                                 'D10'
                        )
)

plot2 = ggplot(regions, aes(x=decile, y=value, fill=type, group=type)) +
  geom_bar(stat="identity") + 
  geom_text(aes(label = round(after_stat(y),1), group = Region),
            stat = 'summary', fun = sum, vjust = -.5, size=2.2) +
  theme(legend.position = "bottom",
        axis.text.x = element_text(angle = 45, vjust=.5, hjust=.5)) +
  labs(title="(B) Decile Cost for Regions By Asset Cost Type",
       fill="Cost\nType",
       subtitle = "Based on 50 GB/Month in EMEs and AEs, and 40 GB/Month in LIDCs",
       x=NULL, y='Cost ($Bn)') +
  scale_y_continuous(expand = c(0, 0), limits=c(0,35)) +
  scale_fill_viridis_d(direction=-1) +
  guides(fill = guide_legend(reverse=T, ncol=3)) + 
  facet_wrap(~Region, scales = "free")

write.csv(regions, file.path(folder, 'final_data', 'fig_7b_regions.csv'))

# shift_legend2 <- function(p) {
#   # to grob
#   gp <- ggplotGrob(p)
#   facet.panels <- grep("^panel", gp[["layout"]][["name"]])
#   empty.facet.panels <- sapply(facet.panels, function(i) "zeroGrob" %in% class(gp[["grobs"]][[i]]))
#   empty.facet.panels <- facet.panels[empty.facet.panels]
#   # establish name of empty panels
#   empty.facet.panels <- gp[["layout"]][empty.facet.panels, ]
#   names <- empty.facet.panels$name
#   # example of names:
#   #[1] "panel-3-2" "panel-3-3"
#   # now we just need a simple call to reposition the legend
#   reposition_legend(p, 'center', panel=names)
# }

# plot2 = shift_legend2(plot2)

aggregated_costs <- ggarrange(plot1, plot2,
                              ncol = 1, nrow = 2, 
                              heights=c(0.45,1))

path = file.path(folder, 'figures', 'decile_panel_costs2.tiff')
tiff(path, units="in", width=8, height=6, res=300)
print(aggregated_costs)
dev.off()

####################GDP
folder <- dirname(rstudioapi::getSourceEditorContext()$path)

path = file.path(folder, '..', "Oughton et al. (2023) DICE v1.0.0.xlsx")
data <- read_excel(path, sheet = "GDP", col_names = T)
data$ISO3 = NULL
data$`country_name` = NULL
data = data[ , c(1,2,13)]
data = data[!(data$`Income Group`=="-"),]
names(data)[names(data) == 'Income Group'] <- 'income'
names(data)[names(data) == 'Region'] <- 'region'
colnames(data)[3] <- "value"
data$value <- as.numeric(data$value)
data = data[complete.cases(data), ]

data$income = factor(data$income, 
                     levels=c(
                       'Low Income Developing Countries',
                       'Emerging Market Economies',
                       'Advanced Economies'))

income = data %>%
  group_by(income) %>%
  summarize(
    value = sum(value, na.rm=TRUE) 
  )

income_results = read.csv(file.path(folder, 'aggregated_income.csv'))
income_results$X = NULL
income = merge(income, income_results, by=c("income"))
income$gdp_perc = (income$cost/income$value) * 100
remove(income_results)
# write.csv(income, file.path(folder,'gdp_perc_by_income_group.csv'))
options(scipen=999)

plot1 = ggplot(income, aes(x=income, y=gdp_perc, fill=Category, group=Category)) +
  geom_text(aes(label = paste(round(after_stat(y),2),"%"), group = income),
            stat = 'summary', fun = sum, hjust = -.2, size=2.5) +
  geom_bar(stat="identity")  +   coord_flip() +
  theme(legend.position = "bottom",
        axis.text.x = element_text(angle = 0, hjust=0)) +
  labs(title="(A) Total Investment as a Percentage of Annual GDP by Income Group",
       fill="Cost\nType",
       subtitle = "Based on 50 GB/Month in EMEs and AEs, and 40 GB/Month in LIDCs",
       x=NULL, y='Percentage of GDP (%)') +
  scale_y_continuous(expand = c(0, 0), limits=c(0,3.8)) +
  scale_fill_viridis_d(direction=-1) +
  guides(fill = guide_legend(reverse=T))

write.csv(income, file.path(folder, 'final_data', 'fig_8a_income.csv'))

data$region = factor(data$region,
                     levels=c("Sub-Sahara Africa",
                              "Middle East, North Africa, Afghanistan, and Pakistan",
                              "Latin America and the Caribbean",
                              "Emerging and Developing Europe",
                              'Emerging and Developing Asia',
                              'Caucasus and Central Asia',
                              'Advanced Economies'),
                     labels=c("Sub-Saharan Africa",
                              "MENA, AFG and PAK",
                              "Latin America and the Caribbean",
                              "Emerging and Developing Europe",
                              'Emerging and Developing Asia',
                              'Caucasus and Central Asia',
                              'Advanced Economies')
)

region = data %>%
  group_by(region) %>%
  summarize(
    value = sum(value, na.rm=TRUE) 
  )

region_results = read.csv(file.path(folder, 'aggregated_regions.csv'))
region_results$X = NULL
region = merge(region, region_results, by=c("region"))
region$gdp_perc = region$cost/region$value * 100
remove(region_results, data)

plot2 = ggplot(region, aes(x=region, y=gdp_perc, fill=Category, group=Category)) +
  geom_text(aes(label = paste(round(after_stat(y),2),"%"), group = region),
            stat = 'summary', fun = sum, hjust = -.2, size=2.5) +
  geom_bar(stat="identity")  +   coord_flip() +
  theme(legend.position = "bottom",
        axis.text.x = element_text(angle = 0, hjust=0)) +
  labs(title="(B) Total Investment as a Percentage of Annual GDP by Region",
       fill="Cost\nType",
       subtitle = "Based on 50 GB/Month in EMEs and AEs, and 40 GB/Month in LIDCs",
       x=NULL, y='Percentage of GDP (%)') +
  scale_y_continuous(expand = c(0, 0), limits=c(0,4.85)) +
  scale_fill_viridis_d(direction=-1) +
  guides(fill = guide_legend(reverse=T))

write.csv(income, file.path(folder, 'final_data', 'fig_8b_regions.csv'))

gdp_perc_costs <- ggarrange(plot1, plot2,
                            ncol = 1, nrow = 2, align = c("hv"),
                            common.legend = TRUE, legend='bottom',
                            heights=c(.65,1))

path = file.path(folder, 'figures', 'gdp_perc_costs.tiff')
tiff(path, units="in", width=8, height=6, res=300)
print(gdp_perc_costs)
dev.off()

######################################
######################################
######################################
####################20/10 GB
folder <- dirname(rstudioapi::getSourceEditorContext()$path)

path = file.path(folder, '..', 'sensitivity', "dice_data_consumption_twenty_and_ten_gig.xlsx")
data <- read_excel(path, sheet = "Cost_Comp_IMF_Income_Group", col_names = T)
names(data) <- data[1,]
data <- data[-1,]
data = data[, 1:4]
data = data[1:6,]

data = data %>% gather(income, cost, -Category)
data$cost <- as.numeric(data$cost)

data$Category = factor(data$Category,
                       levels=c('Mobile Infra Capex', 'Metro+backbone fiber',
                                'Mobile Infra Opex','Remote Coverage','Policy/Regulation',
                                'ICT Skills/Content'),
                       labels=c('Mobile Infra Capex', 'Metro+Backbone Fiber',
                                'Mobile Infra Opex','Remote Coverage','Policy/Regulation',
                                'ICT Skills/Content')
)

data$income = factor(data$income,
                     levels=c('Low Income Developing Countries',
                              'Emerging Market Economies',
                              'Advanced Economies'),
                     labels=c('Low Income\nDeveloping\nCountries',
                              'Emerging\nMarket\nEconomies',
                              'Advanced\nEconomies')
)

plot1 = ggplot(data, aes(x=income, y=cost, fill=Category, group=Category)) +
  geom_bar(stat="identity") + coord_flip() +
  geom_text(aes(label = round(after_stat(y),0), group = income),
            stat = 'summary', fun = sum, hjust = -.5, size=2.5) +
  theme(legend.position = "bottom",
        axis.text.x = element_text(angle = 0, hjust=.5)) +
  labs(title="(A) Lower Data Consumption by Income",
       fill="Cost\nType",
       subtitle = "EMEs/AEs: 20 GB/Month. LIDCs: 10 GB/Month",
       x=NULL, y='Investment Cost ($Bn)') +
  scale_y_continuous(expand = c(0, 0), limits=c(0,650)) +
  scale_fill_viridis_d(direction=-1) +
  guides(fill = guide_legend(reverse=T))

write.csv(data, file.path(folder, 'final_data', 'fig_11a.csv'))

folder <- dirname(rstudioapi::getSourceEditorContext()$path)

path = file.path(folder, '..', 'sensitivity', "dice_data_consumption_twenty_and_ten_gig.xlsx")
data <- read_excel(path, sheet = "Cost_Comp_IMF_Region", col_names = T)
names(data) <- data[1,]
data <- data[-1,]
data = data[, 1:8]
data = data[1:6,]

data = data %>% gather(region, cost, -Category)
data$cost <- as.numeric(data$cost)

data$Category = factor(data$Category,
                       levels=c('Mobile Infra Capex', 'Metro+backbone fiber',
                                'Mobile Infra Opex','Remote Coverage','Policy/Regulation',
                                'ICT Skills/Content'),
                       labels=c('Mobile Infra Capex', 'Metro+Backbone Fiber',
                                'Mobile Infra Opex','Remote Coverage','Policy/Regulation',
                                'ICT Skills/Content')
)

data$region = factor(data$region,
                     levels=c("Sub-Sahara Africa",
                              "Middle East, North Africa, Afghanistan, and Pakistan",
                              "Latin America and the Caribbean",
                              "Emerging and Developing Europe",
                              'Emerging and Developing Asia',
                              'Caucasus and Central Asia',
                              'Advanced Economies'),
                     labels=c("Sub-Saharan\nAfrica",
                              "MENA, AFG\nand PAK",
                              "Latin America\nand the\nCaribbean",
                              "Emerging and\nDeveloping\nEurope",
                              'Emerging and\nDeveloping\nAsia',
                              'Caucasus and\nCentral\nAsia',
                              'Advanced\nEconomies')
)

plot2 = ggplot(data, aes(x=region, y=cost, fill=Category, group=Category)) +
  geom_text(aes(label = round(after_stat(y),0), group = region),
            stat = 'summary', fun = sum, hjust = -.5, size=2.5) +
  geom_bar(stat="identity")  +   coord_flip() +
  theme(legend.position = "bottom",
        axis.text.x = element_text(angle = 0, hjust=.5)) +
  labs(title="(C) Lower Data Consumption by Region",
       fill="Cost\nType",
       subtitle = "EMEs/AEs: 20 GB/Month. LIDCs: 10 GB/Month",
       x=NULL, y='Investment Cost ($Bn)') +
  scale_y_continuous(expand = c(0, 0), limits=c(0,430)) +
  scale_fill_viridis_d(direction=-1) +
  guides(fill = guide_legend(reverse=T))

write.csv(data, file.path(folder, 'final_data', 'fig_11b.csv'))

######################################
####################100/80 GB
folder <- dirname(rstudioapi::getSourceEditorContext()$path)

path = file.path(folder, '..', 'sensitivity', "dice_data_consumption_one_hundred_and_eighty_gig.xlsx")
data <- read_excel(path, sheet = "Cost_Comp_IMF_Income_Group", col_names = T)
names(data) <- data[1,]
data <- data[-1,]
data = data[, 1:4]
data = data[1:6,]

data = data %>% gather(income, cost, -Category)
data$cost <- as.numeric(data$cost)

data$Category = factor(data$Category,
                       levels=c('Mobile Infra Capex', 'Metro+backbone fiber',
                                'Mobile Infra Opex','Remote Coverage','Policy/Regulation',
                                'ICT Skills/Content'),
                       labels=c('Mobile Infra Capex', 'Metro+Backbone Fiber',
                                'Mobile Infra Opex','Remote Coverage','Policy/Regulation',
                                'ICT Skills/Content')
)

data$income = factor(data$income,
                     levels=c('Low Income Developing Countries',
                              'Emerging Market Economies',
                              'Advanced Economies'),
                     labels=c('Low Income\nDeveloping\nCountries',
                              'Emerging\nMarket\nEconomies',
                              'Advanced\nEconomies')
)

plot3 = ggplot(data, aes(x=income, y=cost, fill=Category, group=Category)) +
  geom_bar(stat="identity") + coord_flip() +
  geom_text(aes(label = round(after_stat(y),0), group = income),
            stat = 'summary', fun = sum, hjust = -.5, size=2.5) +
  theme(legend.position = "bottom",
        axis.text.x = element_text(angle = 0, hjust=.5)) +
  labs(title="(B) Higher Data Consumption by Income",
       fill="Cost\nType",
       subtitle = "EMEs/AEs: 100 GB/Month. LIDCs: 80 GB/Month",
       x=NULL, y='Investment Cost ($Bn)') +
  scale_y_continuous(expand = c(0, 0), limits=c(0,650)) +
  scale_fill_viridis_d(direction=-1) +
  guides(fill = guide_legend(reverse=T))

write.csv(data, file.path(folder, 'final_data', 'fig_11c.csv'))

folder <- dirname(rstudioapi::getSourceEditorContext()$path)

path = file.path(folder, '..', 'sensitivity', "dice_data_consumption_one_hundred_and_eighty_gig.xlsx")
data <- read_excel(path, sheet = "Cost_Comp_IMF_Region", col_names = T)
names(data) <- data[1,]
data <- data[-1,]
data = data[, 1:8]
data = data[1:6,]

data = data %>% gather(region, cost, -Category)
data$cost <- as.numeric(data$cost)

data$Category = factor(data$Category,
                       levels=c('Mobile Infra Capex', 'Metro+backbone fiber',
                                'Mobile Infra Opex','Remote Coverage','Policy/Regulation',
                                'ICT Skills/Content'),
                       labels=c('Mobile Infra Capex', 'Metro+Backbone Fiber',
                                'Mobile Infra Opex','Remote Coverage','Policy/Regulation',
                                'ICT Skills/Content')
)

data$region = factor(data$region,
                     levels=c("Sub-Sahara Africa",
                              "Middle East, North Africa, Afghanistan, and Pakistan",
                              "Latin America and the Caribbean",
                              "Emerging and Developing Europe",
                              'Emerging and Developing Asia',
                              'Caucasus and Central Asia',
                              'Advanced Economies'),
                     labels=c("Sub-Saharan\nAfrica",
                              "MENA, AFG\nand PAK",
                              "Latin America\nand the\nCaribbean",
                              "Emerging and\nDeveloping\nEurope",
                              'Emerging and\nDeveloping\nAsia',
                              'Caucasus and\nCentral\nAsia',
                              'Advanced\nEconomies')
)

plot4 = ggplot(data, aes(x=region, y=cost, fill=Category, group=Category)) +
  geom_text(aes(label = round(after_stat(y),0), group = region),
            stat = 'summary', fun = sum, hjust = -.5, size=2.5) +
  geom_bar(stat="identity")  +   coord_flip() +
  theme(legend.position = "bottom",
        axis.text.x = element_text(angle = 0, hjust=.5)) +
  labs(title="(D) Higher Data Consumption by Region",
       fill="Cost\nType",
       subtitle = "EMEs/AEs: 100 GB/Month. LIDCs: 80 GB/Month",
       x=NULL, y='Investment Cost ($Bn)') +
  scale_y_continuous(expand = c(0, 0), limits=c(0,430)) +
  scale_fill_viridis_d(direction=-1) +
  guides(fill = guide_legend(reverse=T))

write.csv(data, file.path(folder, 'final_data', 'fig_11d.csv'))

data_consumption_sensitivity.tiff <- ggarrange(plot1, plot3, plot2, plot4,
                                               ncol = 2, nrow = 2, align = c("hv"),
                                               common.legend = TRUE, legend='bottom',
                                               heights=c(.6,1, .6,1))

path = file.path(folder, 'figures', 'data_consumption_sensitivity.tiff')
tiff(path, units="in", width=9, height=8, res=300)
print(data_consumption_sensitivity.tiff)
dev.off()


######################################
######################################
######################################
####################Reliability 5%
folder <- dirname(rstudioapi::getSourceEditorContext()$path)

path = file.path(folder, '..', 'sensitivity', "dice_reliability_five_percent.xlsx")
data <- read_excel(path, sheet = "Cost_Comp_IMF_Income_Group", col_names = T)
names(data) <- data[1,]
data <- data[-1,]
data = data[, 1:4]
data = data[1:6,]

data = data %>% gather(income, cost, -Category)
data$cost <- as.numeric(data$cost)

data$Category = factor(data$Category,
                       levels=c('Mobile Infra Capex', 'Metro+backbone fiber',
                                'Mobile Infra Opex','Remote Coverage','Policy/Regulation',
                                'ICT Skills/Content'),
                       labels=c('Mobile Infra Capex', 'Metro+Backbone Fiber',
                                'Mobile Infra Opex','Remote Coverage','Policy/Regulation',
                                'ICT Skills/Content')
)

data$income = factor(data$income,
                     levels=c('Low Income Developing Countries',
                              'Emerging Market Economies',
                              'Advanced Economies'),
                     labels=c('Low Income\nDeveloping\nCountries',
                              'Emerging\nMarket\nEconomies',
                              'Advanced\nEconomies')
)

plot1 = ggplot(data, aes(x=income, y=cost, fill=Category, group=Category)) +
  geom_bar(stat="identity") + coord_flip() +
  geom_text(aes(label = round(after_stat(y),0), group = income),
            stat = 'summary', fun = sum, hjust = -.5, size=2.5) +
  theme(legend.position = "bottom",
        axis.text.x = element_text(angle = 0, hjust=.5)) +
  labs(title="(A) Lower QoS Reliability by Income",
       fill="Cost\nType",
       subtitle = "Quality of service reduced to 5% reliability",
       x=NULL, y='Investment Cost ($Bn)') +
  scale_y_continuous(expand = c(0, 0), limits=c(0,190)) +
  scale_fill_viridis_d(direction=-1) +
  guides(fill = guide_legend(reverse=T))

write.csv(data, file.path(folder, 'final_data', 'fig_12a.csv'))

folder <- dirname(rstudioapi::getSourceEditorContext()$path)

path = file.path(folder, '..', 'sensitivity', "dice_reliability_five_percent.xlsx")
data <- read_excel(path, sheet = "Cost_Comp_IMF_Region", col_names = T)
names(data) <- data[1,]
data <- data[-1,]
data = data[, 1:8]
data = data[1:6,]

data = data %>% gather(region, cost, -Category)
data$cost <- as.numeric(data$cost)

data$Category = factor(data$Category,
                       levels=c('Mobile Infra Capex', 'Metro+backbone fiber',
                                'Mobile Infra Opex','Remote Coverage','Policy/Regulation',
                                'ICT Skills/Content'),
                       labels=c('Mobile Infra Capex', 'Metro+Backbone Fiber',
                                'Mobile Infra Opex','Remote Coverage','Policy/Regulation',
                                'ICT Skills/Content')
)

data$region = factor(data$region,
                     levels=c("Sub-Sahara Africa",
                              "Middle East, North Africa, Afghanistan, and Pakistan",
                              "Latin America and the Caribbean",
                              "Emerging and Developing Europe",
                              'Emerging and Developing Asia',
                              'Caucasus and Central Asia',
                              'Advanced Economies'),
                     labels=c("Sub-Saharan\nAfrica",
                              "MENA, AFG\nand PAK",
                              "Latin America\nand the\nCaribbean",
                              "Emerging and\nDeveloping\nEurope",
                              'Emerging and\nDeveloping\nAsia',
                              'Caucasus and\nCentral\nAsia',
                              'Advanced\nEconomies')
)

plot2 = ggplot(data, aes(x=region, y=cost, fill=Category, group=Category)) +
  geom_text(aes(label = round(after_stat(y),0), group = region),
            stat = 'summary', fun = sum, hjust = -.5, size=2.5) +
  geom_bar(stat="identity")  +   coord_flip() +
  theme(legend.position = "bottom",
        axis.text.x = element_text(angle = 0, hjust=.5)) +
  labs(title="(C) Lower QoS Reliability by Region",
       fill="Cost\nType",
       subtitle = "Quality of service reduced to 5% reliability",
       x=NULL, y='Investment Cost ($Bn)') +
  scale_y_continuous(expand = c(0, 0), limits=c(0,80)) +
  scale_fill_viridis_d(direction=-1) +
  guides(fill = guide_legend(reverse=T))

write.csv(data, file.path(folder, 'final_data', 'fig_12b.csv'))

######################################
####################50%
folder <- dirname(rstudioapi::getSourceEditorContext()$path)

path = file.path(folder, '..', 'sensitivity', "dice_reliability_fifty_percent.xlsx")
data <- read_excel(path, sheet = "Cost_Comp_IMF_Income_Group", col_names = T)
names(data) <- data[1,]
data <- data[-1,]
data = data[, 1:4]
data = data[1:6,]

data = data %>% gather(income, cost, -Category)
data$cost <- as.numeric(data$cost)

data$Category = factor(data$Category,
                       levels=c('Mobile Infra Capex', 'Metro+backbone fiber',
                                'Mobile Infra Opex','Remote Coverage','Policy/Regulation',
                                'ICT Skills/Content'),
                       labels=c('Mobile Infra Capex', 'Metro+Backbone Fiber',
                                'Mobile Infra Opex','Remote Coverage','Policy/Regulation',
                                'ICT Skills/Content')
)

data$income = factor(data$income,
                     levels=c('Low Income Developing Countries',
                              'Emerging Market Economies',
                              'Advanced Economies'),
                     labels=c('Low Income\nDeveloping\nCountries',
                              'Emerging\nMarket\nEconomies',
                              'Advanced\nEconomies')
)

plot3 = ggplot(data, aes(x=income, y=cost, fill=Category, group=Category)) +
  geom_bar(stat="identity") + coord_flip() +
  geom_text(aes(label = round(after_stat(y),0), group = income),
            stat = 'summary', fun = sum, hjust = -.5, size=2.5) +
  theme(legend.position = "bottom",
        axis.text.x = element_text(angle = 0, hjust=.5)) +
  labs(title="(B) Lower QoS Reliability by Income",
       fill="Cost\nType",
       subtitle = "Quality of service reduced to 50% reliability",
       x=NULL, y='Investment Cost ($Bn)') +
  scale_y_continuous(expand = c(0, 0), limits=c(0,230)) +
  scale_fill_viridis_d(direction=-1) +
  guides(fill = guide_legend(reverse=T))

write.csv(data, file.path(folder, 'final_data', 'fig_12c.csv'))

folder <- dirname(rstudioapi::getSourceEditorContext()$path)

path = file.path(folder, '..', 'sensitivity', "dice_reliability_fifty_percent.xlsx")
data <- read_excel(path, sheet = "Cost_Comp_IMF_Region", col_names = T)
names(data) <- data[1,]
data <- data[-1,]
data = data[, 1:8]
data = data[1:6,]

data = data %>% gather(region, cost, -Category)
data$cost <- as.numeric(data$cost)

data$Category = factor(data$Category,
                       levels=c('Mobile Infra Capex', 'Metro+backbone fiber',
                                'Mobile Infra Opex','Remote Coverage','Policy/Regulation',
                                'ICT Skills/Content'),
                       labels=c('Mobile Infra Capex', 'Metro+Backbone Fiber',
                                'Mobile Infra Opex','Remote Coverage','Policy/Regulation',
                                'ICT Skills/Content')
)

data$region = factor(data$region,
                     levels=c("Sub-Sahara Africa",
                              "Middle East, North Africa, Afghanistan, and Pakistan",
                              "Latin America and the Caribbean",
                              "Emerging and Developing Europe",
                              'Emerging and Developing Asia',
                              'Caucasus and Central Asia',
                              'Advanced Economies'),
                     labels=c("Sub-Saharan\nAfrica",
                              "MENA, AFG\nand PAK",
                              "Latin America\nand the\nCaribbean",
                              "Emerging and\nDeveloping\nEurope",
                              'Emerging and\nDeveloping\nAsia',
                              'Caucasus and\nCentral\nAsia',
                              'Advanced\nEconomies')
)

plot4 = ggplot(data, aes(x=region, y=cost, fill=Category, group=Category)) +
  geom_text(aes(label = round(after_stat(y),0), group = region),
            stat = 'summary', fun = sum, hjust = -.5, size=2.5) +
  geom_bar(stat="identity")  +   coord_flip() +
  theme(legend.position = "bottom",
        axis.text.x = element_text(angle = 0, hjust=.5)) +
  labs(title="(D) Lower QoS Reliability by Region",
       fill="Cost\nType",
       subtitle = "Quality of service reduced to 50% reliability",
       x=NULL, y='Investment Cost ($Bn)') +
  scale_y_continuous(expand = c(0, 0), limits=c(0,115)) +
  scale_fill_viridis_d(direction=-1) +
  guides(fill = guide_legend(reverse=T))

write.csv(data, file.path(folder, 'final_data', 'fig_12d.csv'))

data_consumption_sensitivity.tiff <- ggarrange(plot1, plot3, plot2, plot4,
                                               ncol = 2, nrow = 2, align = c("hv"),
                                               common.legend = TRUE, legend='bottom',
                                               heights=c(.6,1, .6,1))

path = file.path(folder, 'figures', 'reliability_sensitivity.tiff')
tiff(path, units="in", width=9, height=8, res=300)
print(data_consumption_sensitivity.tiff)
dev.off()

######################################
######################################
######################################
###########Validation 
folder <- dirname(rstudioapi::getSourceEditorContext()$path)

path = file.path(folder, '..', "Oughton et al. (2023) DICE v1.0.0.xlsx")
data <- read_excel(path, sheet = "Validation", col_names = T)

colnames(data)[2] <- "ITU"
colnames(data)[4] <- "DICE"

data$"ITU" <- as.numeric(data$"ITU")
data$"DICE" <- as.numeric(data$"DICE")

######################################
###########cost categories 

asset_type = data[1:6,]
asset_type = select(asset_type, Category, "ITU", "DICE")

asset_type$Category = factor(asset_type$Category,
                             levels=c('ICT Skills/Content', 'Policy/Regulation',
                                      'Remote Coverage', 'Mobile Infra Opex','Metro+backbone fiber',
                                      'Mobile Infra Capex'),
                             labels=c('ICT Skills/Content', 'Policy/Regulation',
                                      'Remote Coverage', 'Mobile Infra Opex','Metro+Backbone Fiber',
                                      'Mobile Infra Capex'
                             )
)

asset_type = gather(asset_type, model, value, "ITU":"DICE")

plot1 = ggplot(asset_type, aes(x=Category, y=value, fill=model, group=model)) +
  geom_text(aes(label = round(after_stat(y),0), group = model), 
            stat = 'summary', fun = sum, hjust = -.5, size=2,
            position = position_dodge(width = .9)) +
  geom_bar(stat="identity", position="dodge") + coord_flip() +
  theme(legend.position = "bottom",
        axis.text.x = element_text(angle = 15, hjust=1)) +
  labs(title="(A) Comparison of DICE Model Baseline to ITU (2019)",
       fill="Model",
       subtitle = bquote("Compared based on investment cost categories."),
       x="",
       y='Estimated Cost (US$ Billions)') +
  scale_y_continuous(expand = c(0, 0), limits=c(0,250)) +
  scale_fill_viridis_d() 

write.csv(asset_type, file.path(folder, 'final_data', 'fig_10a.csv'))

######################################
###########regions

regions = data[11:16,]
regions = select(regions, Category, "ITU", "DICE")

regions$Category = factor(regions$Category,
                          levels=c("SSA",
                                   "South Asia",
                                   "MENA",
                                   'Europe + Central Asia',
                                   'East Asia + Pacific',
                                   'Americas'),
                          labels=c("SSA",
                                   "South Asia",
                                   "MENA",
                                   'Europe + Central Asia',
                                   'East Asia + Pacific',
                                   'Americas')
)

regions = gather(regions, model, value, "ITU":"DICE")

plot2 = ggplot(regions, aes(x=Category, y=value, fill=model, group=model)) +
  geom_text(aes(label = round(after_stat(y),0), group = model), 
            stat = 'summary', fun = sum, hjust = -.5, size=2,
            position = position_dodge(width = .9)) +
  geom_bar(stat="identity", position="dodge") + coord_flip() +
  theme(legend.position = "bottom",
        axis.text.x = element_text(angle = 15, hjust=1)) +
  labs(title="(B) Comparison of DICE Model Baseline to ITU (2019)",
       fill="Model",
       subtitle = bquote("Compared based on ITU geographic regions."),
       x="",
       y='Estimated Cost (US$ Billions)') +
  scale_y_continuous(expand = c(0, 0), limits=c(0,250)) +
  scale_fill_viridis_d() 

write.csv(regions, file.path(folder, 'final_data', 'fig_10b.csv'))

######################################
###########regions

income = data[21:24,]
income = select(income, Category, "ITU", "DICE")

income$Category = factor(income$Category,
                         levels=c("High Income",
                                  "Upper Middle Income",
                                  "Lower Middle Income",
                                  'Low Income'),
                         labels=c("High Income",
                                  "Upper Middle Income",
                                  "Lower Middle Income",
                                  'Low Income')
)

income = gather(income, model, value, "ITU":"DICE")

plot3 = 
  ggplot(income, aes(x=Category, y=value, fill=model, group=model)) +
  geom_text(aes(label = round(after_stat(y),0), group = model), 
            stat = 'summary', fun = sum, hjust = -.5, size=2,
            position = position_dodge(width = .9)) +
  geom_bar(stat="identity", position="dodge") + coord_flip() +
  theme(legend.position = "bottom",
        axis.text.x = element_text(angle = 15, hjust=1)) +
  labs(title="(C) Comparison of DICE Model Baseline to ITU (2019)",
       fill="Model",
       subtitle = bquote("Compared based on World Bank income groups."),
       x="",
       y='Estimated Cost (US$ Billions)') +
  scale_y_continuous(expand = c(0, 0), limits=c(0,250)) +
  scale_fill_viridis_d() 

write.csv(income, file.path(folder, 'final_data', 'fig_10c.csv'))

panel <- ggarrange(plot1, plot2, plot3,
                   ncol = 1, nrow = 3, align = c("hv"),
                   common.legend = TRUE, legend='bottom',
                   heights=c(1,1,.85))

path = file.path(folder, 'figures', 'test_panel.png')
dir.create(file.path(folder, 'figures'), showWarnings = FALSE)
ggsave(
  'panel.png',
  plot = last_plot(),
  device = "png",
  path=file.path(folder, 'figures'),
  units = c("in"),
  width = 8.3,
  height = 8.3,
  bg="white"
)
