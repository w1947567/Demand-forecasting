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
library(psych)
library(Hmisc)

#Set data pathway
setwd("C:/Users/lauri/OneDrive/Desktop/Intersurgica/Data")

# Load your data
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
start_date<-ymd("2014-01-01")
end_date<-ymd("2022-12-01")

#Filter data for date range
data_filtered<-df[df$Date>=start_date & df$Date<=end_date,]

#Convert data into time series data
df_ts<-ts(data_filtered$"Volume", frequency = 12,
          start = c(year(min(data_filtered$Date)), month(min(data_filtered$Date))))
plot(df_ts, main="Time series data", xlab="Years", ylab="Volume")

#Difference to make data stationary
diff_df_ts<-diff(df_ts, differences=1)
plot(diff_df_ts, main = "Differenced data", ylab = "Volume")
abline(h=0, col="red", lty="dashed")

#Proved stationarity using ADF test
adf_test<-adf.test(diff_df_ts, alternative="stationary")
adf_test

#Examination of autocorrelation/partial correlation
ggtsdisplay(df_ts)
ggtsdisplay(diff_df_ts)
ggtsdisplay(diff(diff_df_ts))

#Fitting SARIMA
fit<-Arima(df_ts, order=c(3,1,3), seasonal=c(1,1,1))
summary(fit)

#Fitting AURO ARIMA
fit.aa<-auto.arima(df_ts, trace=TRUE)
summary(fit.aa)

#Visual fitted models values vs actual
autoplot(df_ts, series="Data")+
  autolayer(fit$fitted,series="SARIMA")+
  autolayer(fit.aa$fitted,series="AUTO ARIMA")

#Evaluate residuals
checkresiduals(fit)
checkresiduals(fit.aa)

#Forecast for 2023
forecast_values<-forecast(fit, h=12)
forecast_values_aa<-forecast(fit.aa, h=12)

grid.arrange(autoplot(forecast(fit.aa)), autoplot(forecast(fit)))

start_actual<-ymd("2023-01-01")
end_actual<-ymd("2023-12-01")
actual_values<-df[df$Date>=start_actual & df$Date<=end_actual, "Volume"]

#Convet actual values to time series 
actual_ts<-ts(actual_values, frequency = 12,
              star=c(year(start_actual), month(start_actual)))

#Table with upper & lower limits
forecast_df<-data.frame(
  Date=seq(start_actual, by="month", length.out=12),
  Actual=actual_values,
  SARIMA=forecast_values$mean,
  Upper_limit_95=forecast_values$upper[,2],
  Lower_limit_95=forecast_values$lower[,2],
  AUTO_ARIMA=forecast_values_aa$mean,
  Upper_limit_95_A=forecast_values_aa$upper[,2],
  Lower_limit_95_A=forecast_values_aa$lower[,2]
)

#Table without upper & lower limits
forecast_table <- data.frame(
  Date = seq(as.Date("2023-01-01"), by = "month", length.out = 12),
  Actual = actual_values,
  SARIMA = forecast_values$mean,
  AUTO_ARIMA = forecast_values_aa$mean
)

########## TESTING #########
#SARIMA
APE<-abs((forecast_table$Actual-forecast_table$SARIMA)/forecast_table$Actual)*100
MAPE<-mean(APE,na.rm=TRUE)
RMSE<-sqrt(mean((forecast_table$SARIMA-forecast_table$Actual)^2))

cat("SARIMA MAPE:", MAPE, "\n")
cat("SARIMA RMSE:", RMSE, "\n")

#AUTO ARIMA
APE<-abs((forecast_table$Actual-forecast_table$AUTO_ARIMA)/forecast_table$Actual)*100
MAPE<-mean(APE,na.rm=TRUE)
RMSE<-sqrt(mean((forecast_table$AUTO_ARIMA-forecast_table$Actual)^2))

cat("AUTO ARIMA MAPE:", MAPE, "\n")
cat("AUTO ARIMA RMSE:", RMSE, "\n")

#Plot forecast values vs actual (with bounderies)
ggplot(data=forecast_df, aes(x=Date))+
  geom_line(aes(y=actual_values, color="Actual values"), linewidth=1)+
  geom_line(aes(y=SARIMA, color="SARIMA"), linewidth=1)+
  geom_ribbon(aes(ymin=Lower_limit_95, ymax=Upper_limit_95), fill="blue", alpha=0.2)+
  geom_line(aes(y=AUTO_ARIMA, color="AUTO ARIMA"), linewidth=1)+
  geom_ribbon(aes(ymin=Lower_limit_95_A, ymax=Upper_limit_95_A), fill="red", alpha=0.2)+
  ggtitle("Comparison of actual values vs forecasted values for 2023")+
  xlab("Year")+ylab("Volume sold")+
  scale_color_manual(values=c("Actual values"="black", "SARIMA"="blue", "AUTO ARIMA"="red"))+
  guides(color=guide_legend(title = "Models"))+
  theme_minimal()

### Export tables as xlsx###
write.xlsx(forecast_df, "limits_2023.xlsx")
write.xlsx(forecast_table, "forecast_2023.xlsx")


