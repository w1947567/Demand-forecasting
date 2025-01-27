library(readxl)
library(tidyr)
library(dplyr)
library(forecast)
library(lubridate)
library(xts)
library(rio)
library(ggplot2)
library(tseries)
library(zoo)
library(Metrics)
library(tidyquant)
library(gridExtra)
library(writexl)
library(openxlsx)
library(xlsx)

#Set data pathway
X1161000 <- read_excel("Desktop/1161000.xlsx")
View(X1161000)
data<-X1161000
data<-as.data.frame(data)

#Drop + combine date columns
df<-data[,-c(1,2)]
df$Date<-as.Date(paste(df$Year, df$Month, "01", sep="-"))
df<-df[,-c(1,2)]
df<-as.data.frame(df)
df<-df[,c("Date", "Volume")]
df

df$Date<-as.Date(df$Date) #convert date to a continuous format

#Create the line plot
ggplot(df, aes(x=Date, y=Volume))+
  geom_line()+
  labs(title="Volume sold", x="Date", y="Volume")

#Define date range
start_date<-ymd("2010-01-01")
end_date<-ymd("2022-12-01")

#Generate sequence of dates
date_sequence<-seq(start_date, end_date, by="1 month")

#Filter data for date range
data_filtered<-df[df$Date>=start_date & df$Date<=end_date,]

#Convert data into time series data
TStest<-ts(data_filtered$'Volume', frequency = 12,
           start=c(year(start_date), month(start_date)),
           end=c(year(end_date), month(end_date)))
plot(TStest, main="Time series data", xlab="Years", ylab="Volume")

#Moving average (MA), trend
df_trend=ma(TStest, order=4, centre=T)
plot(as.ts(TStest), ylab="Moving average")
lines(df_trend, col="red")

#Decompose data
decomposed<-decompose(TStest, type="additive")
plot(decomposed)

decomposed<-data.frame(
  Date=time(TStest),
  Observed=as.numeric(TStest),
  Trend=as.numeric(df_trend),
  Seasonal=as.numeric(decomposed$seasonal),
  Random=as.numeric(decomposed$random)
)

#Fit HW for actual data
HW_fit<-HoltWinters(TStest, seasonal="additive")
plot(HW_fit, main="HW additive model")

#Forecast + visual of predictions
HW_forecast<-forecast(HW_fit, h=12)
plot(HW_forecast, main="HW forecast for 2023", xlab="Years", ylab="Volume sold")

actual_2023<-df[year(df$Date)==2023, "Volume"]
forecast_2023<-subset(HW_forecast, floor(time(HW_forecast$mean))==2023)
dates_2023<-seq(as.Date("2023-01-01"), as.Date("2023-12-01"), by="month")

forecast_table<-data.frame(
  Date=dates_2023,
  Actual=actual_2023,
  HW=forecast_2023$mean,
  Upper_limit=forecast_2023$upper[,1],
  Lower_limit=forecast_2023$lower[,1]
  )

########## TESTING #########
APE<-abs((forecast_table$Actual - forecast_table$HW)/forecast_table$Actual)*100
MAPE<-mean(APE, na.rm=TRUE)
RMSE<-sqrt(mean((forecast_table$Actual-forecast_table$HW)^2))

cat("HW MAPE:", MAPE, "\n")
cat("HW RMSE:", RMSE, "\n")

ggplot(data = forecast_table, aes(x = Date)) +
  geom_line(aes(y = Actual, color = "Actual data"), linewidth = 1) +
  geom_line(aes(y = HW, color = "Forecasted value"), linewidth = 1) +
  geom_ribbon(aes(ymin = Lower_limit, ymax = Upper_limit), fill = "blue", alpha = 0.2) +
  ggtitle("Comparison of actual values vs forecasted 2023 ") +
  xlab("Months") + 
  ylab("Volume sold") +
  scale_color_manual(values = c("Actual data" = "black", "Forecast" = "blue")) +
  guides(colour = guide_legend(title = "Model: Holt-Winters - additive")) +
  theme_minimal()

### Export tables as xlsx###
write.xlsx(forecast_table, "limits_2023.xlsx")
