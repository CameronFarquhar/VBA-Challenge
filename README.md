# VBA project - The VBA of Wall Street


### Files

* [Test Data](Resources/alphabetical_testing.xlsx) - Use this while developing your scripts.

* [Stock Data](Resources/Multiple_year_stock_data.xlsx) - Run your scripts on this data to generate the final homework report.

### Stock market analyst

![stock Market](Images/stockmarket.jpg)

## Description

* Loop through all the stocks for one year and output the following information.

  * The ticker symbol.

  * Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The percent change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The total stock volume of the stock.

* Conditional formatting highlights positive change in green and negative change in red.

* The result should look as follows.

![moderate_solution](Images/moderate_solution.png)

## Additional Analysis

* The stock with the "Greatest % increase", "Greatest % decrease" and "Greatest total volume".

![hard_solution](Images/hard_solution.png)

* The VBA script that will allow it to run on every worksheet, i.e., every year, just by running the VBA script once.

## Other Considerations

* Using the sheet `alphabetical_testing.xlsx` while developing the code. This data set is smaller and will allow to test faster. The code should run on this file in less than 3-5 minutes.

* Make sure that the script acts the same on each sheet.
