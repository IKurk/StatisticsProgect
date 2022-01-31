library(xlsx)
library(lubridate)
library(plyr)
library(ggpubr)

Data <- list()
for (i in 1:6) {
  Data[[i]] <- read.xlsx("C:/Rlib/Var_3.xlsx", sheetIndex = i)
}


DepSum <- list()
for (i in 1:length(Data)) {
  DepSum[[i]] <- numcolwise(sum)(Data[[i]])
}


TotalSum <- list()
for (i in 1:length(Data)) {
  TotalSum[i] <- sum(DepSum[[i]]) 
}

MarketsFrame <- data.frame(vol = as.numeric(TotalSum), market = as.character(list("Магазин 1","Магазин 2","Магазин 3","Магазин 4","Магазин 5","Магазин 6")))
ggbarplot(data = MarketsFrame, x = "market", y = "vol", fill = "market")

###

ggpie(data = MarketsFrame, x = "vol", fill = "market")

###

MarketSum <- list()
for (i in 1:length(Data)) {
  MarketSum[[i]] <- sum(DepSum[[i]])
}


DepFrame <- data.frame(vol = as.numeric(MarketSum), market = as.character(list("Магазин 1","Магазин 2","Магазин 3","Магазин 4","Магазин 5","Магазин 6")))
ggbarplot(data = DepFrame, x = "market", y = "vol", fill = "market")



###############

MarketPerYears2020 <- list()
for (i in 1:6) {
  MarketPerYears2020[[i]]<- Data[[i]][year(Data[[i]]$Дата) == 2020,]
}


TotalDepMarketPerYears2020 <- list()
for (i in 1:6) {
  TotalDepMarketPerYears2020 <- numcolwise(sum)(MarketPerYears2020[[i]])
}



TotalMarketPerYears2020 <- list()
for (i in 1:6) {
  TotalMarketPerYears2020[[i]] <- sum(TotalDepMarketPerYears2020[[i]])
}


MarketPerYears2021 <- list()
for (i in 1:6) {
  
  MarketPerYears2021[[i]]<- Data[[i]][year(Data[[i]]$Дата) == 2021,]
}


TotalDepMarketPerYears2021 <- list()
for (i in 1:6) {
  TotalDepMarketPerYears2021 <- numcolwise(sum)(MarketPerYears2021[[i]])
}



TotalMarketPerYears2021 <- list()
for (i in 1:6) {
  TotalMarketPerYears2021[[i]] <- sum(TotalDepMarketPerYears2021[[i]])
}


GrowRate <- list()
for (i in 1:6) {
  GrowRate[[i]] <- ((TotalMarketPerYears2020[[i]]/ TotalMarketPerYears2021[[i]])*100)-100
}

GrowRateFrame <- data.frame(GrowRate = as.numeric(GrowRate), market = as.character(list("Магазин 1","Магазин 2","Магазин 3","Магазин 4","Магазин 5","Магазин 6")))
ggbarplot(data = GrowRateFrame, x = "market", y = "GrowRate", fill = "market")

###################33

MarketPerYears2018 <- list()
for (i in 1:6) {
  
  MarketPerYears2018[[i]]<- Data[[i]][year(Data[[i]]$Дата) == 2018,]
}

TotalDepMarketPerYears2018 <- list()
for (i in 1:6) {
  TotalDepMarketPerYears2018 <- numcolwise(sum)(MarketPerYears2018[[i]])
}
MarketPerYears2019 <- list()
for (i in 1:6) {
  
  MarketPerYears2019[[i]]<- Data[[i]][year(Data[[i]]$Дата) == 2019,]
}

TotalDepMarketPerYears2019 <- list()
for (i in 1:6) {
  TotalDepMarketPerYears2019 <- numcolwise(sum)(MarketPerYears2019[[i]])
}


TotalDepSum <- list()
for (i in 1:18) {
  TotalDepSum[[i]] <- sum(TotalDepMarketPerYears2018[,i])+sum(TotalDepMarketPerYears2019[,i])+
    sum(TotalDepMarketPerYears2020[,i])+sum(TotalDepMarketPerYears2021[,i])
}
  

TotalDepSumFrame <- data.frame(vol = as.numeric(TotalDepSum), Dep = as.character(names(TotalDepMarketPerYears2018)))

TotalSumNum <- 0
for (i in 1:18) {
  TotalSumNum <- TotalSumNum + TotalDepSum[[i]]
}


TotalDepSumProcent <- list()
for (i in 1:18) {
  TotalDepSumProcent[[i]] <- round(((TotalDepSum[[i]]/TotalSumNum)*100), digits = 2)
}

TotalDepSumFrameProcent <- data.frame(vol = as.numeric(TotalDepSumProcent), Dep = (as.character(names(TotalDepMarketPerYears2018))))

ggpie(data = TotalDepSumFrameProcent, x = "vol", fill = "Dep")


############
 Tea <- list()
for (i in 1:6) {
  Tea[[i]] <- DepSum[[i]]$Чай.Кофе
}

 MarketTeaFrame <- data.frame(MarketTeaFrame = as.numeric(Tea), market = as.character(list("Магазин 1","Магазин 2",
                                                                                           "Магазин 3","Магазин 4","Магазин 5","Магазин 6")))
 ggpie(data = MarketTeaFrame, x = "MarketTeaFrame", fill = "market")

###########
 
 DepSumPerDay <- list()
for (i in 1:18) {
  DepSumPerDay[[i]] <- TotalDepSum[[i]]/(365*4)
}

 
 DepSumPerDayframe <- data.frame(SumPerDay = as.numeric(DepSumPerDay), Dep = as.character(names(TotalDepMarketPerYears2018)))
 ggbarplot(data = DepSumPerDayframe, x = "Dep", y = "SumPerDay", fill = "Dep", x.text.angle = 45)


#########################

 DepSumPerDayMarket1 <- list()
 for (i in 1:18) {
   DepSumPerDayMarket1[[i]] <- DepSum[[1]][1,i]/(365*4)
 }
 
 DepSumPerDayMarket1frame <- data.frame(SumPerDayMarket1 = as.numeric(DepSumPerDayMarket1), Dep = as.character(names(TotalDepMarketPerYears2018)))
 ggbarplot(data = DepSumPerDayMarket1frame, x = "Dep", y = "SumPerDayMarket1", fill = "Dep", x.text.angle = 45)

 
 
 
 
 
 DepSumPerDayMarket2 <- list()
 for (i in 1:18) {
   DepSumPerDayMarket2[[i]] <- DepSum[[2]][1,i]/(365*4)
 }
 
 DepSumPerDayMarket2frame <- data.frame(SumPerDayMarket2 = as.numeric(DepSumPerDayMarket2), Dep = as.character(names(TotalDepMarketPerYears2018)))
 ggbarplot(data = DepSumPerDayMarket2frame, x = "Dep", y = "SumPerDayMarket2", fill = "Dep", x.text.angle = 45)
 
 
 
 
 
 DepSumPerDayMarket3 <- list()
 for (i in 1:18) {
   DepSumPerDayMarket3[[i]] <- DepSum[[3]][1,i]/(365*4)
 }
 
 DepSumPerDayMarket3frame <- data.frame(SumPerDayMarket3 = as.numeric(DepSumPerDayMarket3), Dep = as.character(names(TotalDepMarketPerYears2018)))
 ggbarplot(data = DepSumPerDayMarket3frame, x = "Dep", y = "SumPerDayMarket3", fill = "Dep", x.text.angle = 45)
 
 
 
 
 DepSumPerDayMarket4 <- list()
 for (i in 1:18) {
   DepSumPerDayMarket4[[i]] <- DepSum[[4]][1,i]/(365*4)
 }
 
 DepSumPerDayMarket4frame <- data.frame(SumPerDayMarket4 = as.numeric(DepSumPerDayMarket4), Dep = as.character(names(TotalDepMarketPerYears2018)))
 ggbarplot(data = DepSumPerDayMarket4frame, x = "Dep", y = "SumPerDayMarket4", fill = "Dep", x.text.angle = 45)
 
 
 
 
 DepSumPerDayMarket5 <- list()
 for (i in 1:18) {
   DepSumPerDayMarket5[[i]] <- DepSum[[5]][1,i]/(365*4)
 }
 
 DepSumPerDayMarket5frame <- data.frame(SumPerDayMarket5 = as.numeric(DepSumPerDayMarket5), Dep = as.character(names(TotalDepMarketPerYears2018)))
 ggbarplot(data = DepSumPerDayMarket5frame, x = "Dep", y = "SumPerDayMarket5", fill = "Dep", x.text.angle = 45)
 
 
 
 
 DepSumPerDayMarket6 <- list()
 for (i in 1:18) {
   DepSumPerDayMarket6[[i]] <- DepSum[[6]][1,i]/(365*4)
 }
 
 DepSumPerDayMarket6frame <- data.frame(SumPerDayMarket6 = as.numeric(DepSumPerDayMarket6), Dep = as.character(names(TotalDepMarketPerYears2018)))
 ggbarplot(data = DepSumPerDayMarket6frame, x = "Dep", y = "SumPerDayMarket6", fill = "Dep", x.text.angle = 45)
 
 
#####################темп прироста магазина 1 
 DepSumMarketYears2020 <- list()
 for (i in 1:6) {
   DepSumMarketYears2020[[i]] <- numcolwise(sum)(MarketPerYears2020[[i]])
 }
 
 
 DepSumMarketYears2021 <- list()
 for (i in 1:6) {
   DepSumMarketYears2021[[i]] <- numcolwise(sum)(MarketPerYears2021[[i]])
 }

 GrowRateMarket1Years2021 <-list()
 
 for (i in 1:18) {
   
    GrowRateMarket1Years2021[[i]] <- ((DepSumMarketYears2021[[1]][1,i]/DepSumMarketYears2020[[1]][1,i])*100)-100
   
 }
 

 GrowRateMarket2Years2021 <-list()
 
 for (i in 1:18) {
   
   GrowRateMarket2Years2021[[i]] <- ((DepSumMarketYears2021[[2]][1,i]/DepSumMarketYears2020[[2]][1,i])*100)-100
   
 }

 
 GrowRateMarket3Years2021 <-list()
 
 for (i in 1:18) {
   
   GrowRateMarket3Years2021[[i]] <- ((DepSumMarketYears2021[[3]][1,i]/DepSumMarketYears2020[[3]][1,i])*100)-100
   
 }

 
 GrowRateMarket4Years2021 <-list()
 
 for (i in 1:18) {
   
   GrowRateMarket4Years2021[[i]] <- ((DepSumMarketYears2021[[4]][1,i]/DepSumMarketYears2020[[4]][1,i])*100)-100
   
 }

 
 GrowRateMarket5Years2021 <-list()
 
 for (i in 1:18) {
   
   GrowRateMarket5Years2021[[i]] <- ((DepSumMarketYears2021[[5]][1,i]/DepSumMarketYears2020[[5]][1,i])*100)-100
   
 }

 
 GrowRateMarket6Years2021 <-list()
 
 for (i in 1:18) {
   
   GrowRateMarket6Years2021[[i]] <- ((DepSumMarketYears2021[[6]][1,i]/DepSumMarketYears2020[[6]][1,i])*100)-100
   
 }
 
 
 GrowRateMarket1Years2021Frame <- data.frame(GrowRate = as.numeric(GrowRateMarket1Years2021), Dep = as.character(names(TotalDepMarketPerYears2018)))
 ggbarplot(data = GrowRateMarket1Years2021Frame, x = "Dep", y = "GrowRate", fill = "Dep", x.text.angle = 45)
 
 
 GrowRateMarket2Years2021Frame <- data.frame(GrowRate = as.numeric(GrowRateMarket2Years2021), Dep = as.character(names(TotalDepMarketPerYears2018)))
 ggbarыplot(data = GrowRateMarket2Years2021Frame, x = "Dep", y = "GrowRate", fill = "Dep", x.text.angle = 45)
 
 GrowRateMarket3Years2021Frame <- data.frame(GrowRate = as.numeric(GrowRateMarket3Years2021), Dep = as.character(names(TotalDepMarketPerYears2018)))
 ggbarplot(data = GrowRateMarket3Years2021Frame, x = "Dep", y = "GrowRate", fill = "Dep", x.text.angle = 45)
 
 GrowRateMarket4Years2021Frame <- data.frame(GrowRate = as.numeric(GrowRateMarket4Years2021), Dep = as.character(names(TotalDepMarketPerYears2018)))
 ggbarplot(data = GrowRateMarket4Years2021Frame, x = "Dep", y = "GrowRate", fill = "Dep", x.text.angle = 45)
 
 GrowRateMarket5Years2021Frame <- data.frame(GrowRate = as.numeric(GrowRateMarket5Years2021), Dep = as.character(names(TotalDepMarketPerYears2018)))
 ggbarplot(data = GrowRateMarket5Years2021Frame, x = "Dep", y = "GrowRate", fill = "Dep", x.text.angle = 45)
 
 GrowRateMarket6Years2021Frame <- data.frame(GrowRate = as.numeric(GrowRateMarket6Years2021), Dep = as.character(names(TotalDepMarketPerYears2018)))
 ggbarplot(data = GrowRateMarket6Years2021Frame, x = "Dep", y = "GrowRate", fill = "Dep", x.text.angle = 45)
 
 
 
#####################
 
 Years <-  c(2018:2021)
 
 MarketPerYears <- list()
 n <- list()
 for (i in 1:6) {
   for (j in 1:4) { 
     
     
     n[[j]]<- Data[[i]][year(Data[[i]]$Дата) == Years[j],]
     
   }
   MarketPerYears[[i]]<-n
 } 
 
 
 
 Months <- c(1:12)
 
 MarketPerMonths <- list()
 n <- list()
 for (i in 1:6) {
    for (j in 1:12) { 
   

      n[[j]]<- Data[[i]][month(Data[[i]]$Дата) == Months[j],]
     
    }
   MarketPerMonths[[i]]<-n
 } 
 
 
 weeks <- c(1:52)
 
 MarketPerWeeks <- list()
 n <- list()
 for (i in 1:6) {
   for (j in 1:52) { 
     
     
     n[[j]]<- Data[[i]][week(Data[[i]]$Дата) == weeks[j],]
     
   }
   MarketPerWeeks[[i]]<-n
 } 
 
 
 
 
 
 ###За 3 любые месяца
 
 DepSum3Month <- list()
 for (i in 1:length(MarketPerMonths)) {
   DepSum3Month[[i]] <- numcolwise(sum)(MarketPerMonths[[i]][[9]])+numcolwise(sum)(MarketPerMonths[[i]][[10]])+
     numcolwise(sum)(MarketPerMonths[[i]][[11]])
 }
 
 TotalSum3Month <- list()
 for (i in 1:length(MarketPerMonths)) {
   TotalSum3Month[i] <- sum(DepSum3Month[[i]]) 
 }
 
 MarketPerMonthsFrame <- data.frame(vol = as.numeric(TotalSum3Month), market = as.character(list("Магазин 1","Магазин 2","Магазин 3",
                                                                                                 "Магазин 4","Магазин 5","Магазин 6")))
 ggbarplot(data = MarketPerMonthsFrame, x = "market", y = "vol", fill = "market",palette = "ucscgb")
 
 
 ###Продажи за любую неделю 
 
 
 DepSum3Week <- list()
 for (i in 1:length(MarketPerWeeks)) {
   DepSum3Week[[i]] <- numcolwise(sum)(MarketPerWeeks[[i]][[1]])
 }
 
 TotalSum3Week <- list()
 for (i in 1:length(MarketPerWeeks)) {
   TotalSum3Week[i] <- sum(DepSum3Week[[i]]) 
 }
 
 MarketPerWeeksFrame <- data.frame(vol = as.numeric(TotalSum3Week), market = as.character(list("Магазин 1","Магазин 2","Магазин 3","Магазин 4",
                                                                                               "Магазин 5","Магазин 6")))
 ggbarplot(data = MarketPerWeeksFrame, x = "market", y = "vol", fill = "market")
 
 

 
 