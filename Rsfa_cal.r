library(readxl)
library(frontier)

# There's something wrong when run this for.
# for (year in 2006:2016) {
#   data1 <- read_excel(paste('RFrontierInputFiles/_sfa_in', year, '.xls', sep = ""))
#   sfa_data_co2 <- sfa( Slack_CO2 ~ Urbanization + Secondary_Industry + Capita_GDP + Environmental_Support + Coal_Consume | -1, data = data1 )
#   sfa_data_capital <- sfa( Slack_CAPITAL ~ Urbanization + Secondary_Industry + Capita_GDP + Environmental_Support + Coal_Consume | -1, data = data1 )
#   sfa_data_work <- sfa( Slack_WORK ~ Urbanization + Secondary_Industry + Capita_GDP + Environmental_Support + Coal_Consume | -1, data = data1 )
#   
#   sink(paste('RFrontierOutputFiles/_sfa_out', year, '.txt', sep = ""))
#   cat(paste('--------', year, '---------\n\n**CO2**\n\n', sep = ""))
#   summary(sfa_data_co2, extraPar=FALSE)
#   cat('**Capital**\n\n')
#   summary(sfa_data_capital, extraPar=FALSE)
#   cat('**Labour**\n\n')
#   summary(sfa_data_work, extraPar=FALSE)
#   cat(paste('--------', year, ' END---------\n\n', sep = ""))
#   sink()
# }

year <- 2014
data1 <- read_excel(paste('RFrontierInputFiles/_sfa_in', year, '.xls', sep = ""))
sfa_data_co2 <- sfa( Slack_CO2 ~ Urbanization + Secondary_Industry + Capita_GDP + Environmental_Support + Coal_Consume | -1, data = data1 )
sfa_data_capital <- sfa( Slack_CAPITAL ~ Urbanization + Secondary_Industry + Capita_GDP + Environmental_Support + Coal_Consume | -1, data = data1 )
sfa_data_work <- sfa( Slack_WORK ~ Urbanization + Secondary_Industry + Capita_GDP + Environmental_Support + Coal_Consume | -1, data = data1 )

sink(paste('RFrontierOutputFiles/_sfa_out', year, '.txt', sep = ""))
cat(paste('--------', year, '---------\n\n**CO2**\n\n', sep = ""))
summary(sfa_data_co2, extraPar=FALSE)
cat('\n**Capital**\n\n')
summary(sfa_data_capital, extraPar=FALSE)
cat('\n**Labour**\n\n')
summary(sfa_data_work, extraPar=FALSE)
cat(paste('--------', year, ' END---------\n\n', sep = ""))
sink()

0.9711547+0.3419829*log(17413.769)-0.0221989*log(11193.066)+0.7370436*log(1220.10)
0.9711547+0.350746*log(10250.173)-0.0221989*log(19321.777)+0.7370436*log(902.42)

1.359880+0.3419829*log(10250.173)-0.044416*log(19321.777)+0.738362*log(902.42)

summary(sfa_data, extraPar=TRUE, effic=TRUE)

epsilon <- 0.286
mustar <- 0.229619
sigmau2 <- 0.09637277
sigmav2 <- 0.074447
sigmauv2 <- sigmav2+sigmau2
# mustar <- -epsilon*sigmau2/sigmauv2
sigmastar2 <- sigmav2*sigmau2/sigmauv2
(1-dnorm(sigmastar2^0.5-mustar/(sigmastar2^0.5)))/(1-dnorm(-mustar/(sigmastar2^0.5)))*exp(-mustar+0.5*sigmastar2)

epsilon_ <- 0.286
lambda_<-1.137766
sigma_<-0.413304

lambda_*sigma_/(1+lambda_^2)*(dnorm(epsilon_*lambda_/sigma_)/pnorm(epsilon_*lambda_/sigma_)+epsilon_*lambda_/sigma_)
# 
# library(sfa)
# sfa_data_s <- sfa(log( GDP ) ~ log( CAPITAL ) + log( CO2 ) + log( WORK ), data = data1)
# View(sfa_data_s)
# 
# te.eff.sfa(sfa_data_s)
# u.sfa(sfa_data_s)
# eff(sfa_data_s)
