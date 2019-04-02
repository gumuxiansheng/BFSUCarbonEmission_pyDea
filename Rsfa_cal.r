library(readxl)
library(frontier)

for (year in 2006:2016) {
  data1 <- read_excel(paste('RFrontierInputFiles/_sfa_in', year, '.xls', sep = ""))
  sfa_data_co2 <- sfa( (-Slack_CO2) ~ Urbanization + Secondary_Industry + Capita_GDP + Coal_Consume | -1, data = data1 )
  sfa_data_capital <- sfa( (-Slack_CAPITAL) ~ Urbanization + Secondary_Industry + Capita_GDP + Coal_Consume | -1, data = data1 )
  sfa_data_work <- sfa( (-Slack_WORK) ~ Urbanization + Secondary_Industry + Capita_GDP + Coal_Consume | -1, data = data1 )

  sink(paste('RFrontierOutputFiles/_sfa_out_xx', year, '.txt', sep = ""))
  cat(paste('--------', year, '---------\n\n**CO2**\n\n', sep = ""))
  print(sfa_data_co2['fitted'][1])
  print(sfa_data_co2['resid'][1])
  cat('**Capital**\n\n')
  print(sfa_data_capital['fitted'][1])
  print(sfa_data_capital['resid'][1])
  cat('**Labour**\n\n')
  print(sfa_data_work['fitted'][1])
  print(sfa_data_work['resid'][1])
  cat(paste('--------', year, ' END---------\n\n', sep = ""))
  sink()
}

# There's something wrong when run this in for.
year <- 2006
data1 <- read_excel(paste('RFrontierInputFiles/_sfa_in', year, '.xls', sep = ""))
sfa_data_co2 <- sfa( (-Slack_CO2) ~ Urbanization + Secondary_Industry + Capita_GDP + Coal_Consume | -1, data = data1 )
sfa_data_capital <- sfa( (-Slack_CAPITAL) ~ Urbanization + Secondary_Industry + Capita_GDP + Coal_Consume | -1, data = data1 )
sfa_data_work <- sfa( (-Slack_WORK) ~ Urbanization + Secondary_Industry + Capita_GDP + Coal_Consume | -1, data = data1 )

sink(paste('RFrontierOutputFiles/_sfa_out', year, '.txt', sep = ""))
print(paste('********', year, '********\n\n**CO2**\n\n', sep = ""))
summary(sfa_data_co2, extraPar=TRUE)
print('\n**Capital**\n\n')
summary(sfa_data_capital, extraPar=TRUE)
print('\n**Labour**\n\n')
summary(sfa_data_work, extraPar=TRUE)
print(paste('********', year, ' END********\n\n', sep = ""))
sink()

# 
# epsilon <- 0.286
# mustar <- 0.229619
# sigmau2 <- 0.09637277
# sigmav2 <- 0.074447
# sigmauv2 <- sigmav2+sigmau2
# # mustar <- -epsilon*sigmau2/sigmauv2
# sigmastar2 <- sigmav2*sigmau2/sigmauv2
# (1-dnorm(sigmastar2^0.5-mustar/(sigmastar2^0.5)))/(1-dnorm(-mustar/(sigmastar2^0.5)))*exp(-mustar+0.5*sigmastar2)
# 
# epsilon_ <- 0.286
# lambda_<-1.137766
# sigma_<-0.413304
# 
# lambda_*sigma_/(1+lambda_^2)*(dnorm(epsilon_*lambda_/sigma_)/pnorm(epsilon_*lambda_/sigma_)+epsilon_*lambda_/sigma_)
