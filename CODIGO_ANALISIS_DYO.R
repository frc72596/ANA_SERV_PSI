library(readxl)
library(summarytools)
library(psych)
BASE_DEP <- read_excel("C:/Users/felip/OneDrive - UNIVERSIDAD DE CUNDINAMARCA/ANA_OF_DE_SERSALD/BASE_DEP.xlsx")
describeBy(BASE_DEP,BASE_DEP$DPNOM_DANE)
descr(BASE_DEP,
      stats     = c("mean", "sd"),
