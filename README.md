# Excel-projects

In the folder "Excel - Projects" there is two projects presented. 

Project number 1 : CoffeOrders

This project is done on dataset called "Coffe Sales", where we have data for coffe sales from 2019 up to 2022, from countries United States, United Kingdom and Ireland. 
In this project we can see which country sales the most amount of coffe, which type of coffe, which size and also top customers. 
There is a timeline which helps us analyze based on years,months or days which is the best selling periods. 

-Using XLOOKUP in "Orders" worksheet to create new column "Customer Name"-"Email"-"Country" based on values found in "Customer" worksheet
-Using INDEX and MATCH in "Orders" worksheet to create column "Coffe Type"-"Roast Type"-"Size"-"Unit Price" based on values found in "Products" worksheet.
-Multiplication formula on columns "Unit Price" and "Quantity" to get "Sales" column.
-Multiple IF functions used on columns "Coffe Type" and "Roast Type" in order to create a new columnswith full names.
-Date Formatting, creating a new forrmated column from "Order Date" with format "dd-mmm-yyyy"
-Number Formatting, changing the decimal places in "Sales" column and setting "dollar" as currency
-Check For Duplicates
-Convert Range to Table
-Pivot Tables and Pivot Charts + Formatting
-Insert Timeline + Formatting
-Insert Slicers + Formatting
-Updating the Pivot Table Data Source
-Building the Dashboard

Project number 2 : divvy-tripdata

-This dashboard is representing the data for company "Divvy Bike-Sharing Chicago" for month December 2022. In this project we can see few crucial points which helps us to understand in which days people tend to rent bikes and what type of bikes people mostly rent.
Also we can see if people tend to rent bikes on longer period or shorter period and if members or casual users mosly rent bikes in this company and what kind of bike they prefere.
Before creating the dashboard, there is few data transformations done, like creating columns :started_at_day_number, ended_at_day_number, started_at_day_name, ended_at_day_name, started_at_timestamp, ended_at_timestamp, ride_duration, ride_duration_minutes,number_of_trips_under_1_minute, number_of_trips_between_1m_and_30m, number_of_trips_between_30m_1h, number_of_trips_more_than_1h.

All of this columns helped later in creating power pivot tables in order to create the charts.
Some of the formulas used :
-CountIFS
-SUM
-IF(nested)
-TEXT
-LARGE
-SMALL
-HOUR
-MINUTE
