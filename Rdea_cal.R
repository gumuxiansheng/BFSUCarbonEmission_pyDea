library(deaR)
data("Hua_Bian_2007")
data_example <- read_data(Hua_Bian_2007,
                          ni = 2,
                          no = 3,
                          ud_outputs = 3) 
result <- model_basic(data_example,
                      orientation = "oo",
                      rts = "vrs") 