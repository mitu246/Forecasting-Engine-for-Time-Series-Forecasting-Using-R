# Forecasting-Engine-for-Time-Series-Forecasting-Using-R
Designed an Automated Forecasting Engine to forecast monthly demand for different product items in each of the product family for Comapny X.

# Introduction:
Company is a automobile parts manufacturing company. During Fall, 2018 they came to Texas A&M University with a challenge for graduate students to build an automated forecasting tool for their supply chain division that could forecast monthly demand for a total of 273 different truck part types separately for 6 months into the future. These 273 part types fall under approx. 52 product families. So, the challenge was to design an automated tool that could act as a forecasting engine. I with a group of 5 people designed this tool where my contribution was designing the Arima Tool for forecasting. Where as in total we had different models fit on data for each product type separately and chose the one with best fit to forecast for that specific product type.
This month i came up with an idea to start from the scratch and build the automated tool myself using 5 different techniques in Time Series Forecasting. Idea was to use basic forecasting methods like Simple Moving Average, Weighted Moving Average and Exponential Smoothing for product types for which there was very few data and the time series was relatively simple with no seasonality and a little to no trend. And then try more sophisticated methods like Tbats Algorithm, Holt Winters Methods, Arima and Prophet Modeling Tool by Facebook, on bigger and complex time series.
Thus, i designed this automated tool using R, where the tool takes in thetime series data for each of the product type separately, processes it, cleans it and then fit different models based in the type of series, and then choses the best of them to forecast the monthly demand for next six months for that specific product type.

# Data:
Data was provided by the Company X. I have chosen not to name the company. It consisted of a file with two sheets in it regarding all the data.
-- Arconic Data 1.xlsx

# Forecast Data:
I had created a csv file with the name of "Final_Arconic_Forecast_newewst.csv" which contains the Product Family, Type, Month and the corresponding forecast.
