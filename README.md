# Example of how to use the IVolatility Rest API with Excel

## Requirements
- Microsoft Excel 2016 or later
- Microsoft Power Query 
  - To learn more about MS Power Query, please visit Microsoft’s official documentation: <https://learn.microsoft.com/en-us/power-query/>
- test API.xlsx
  - This is IVolatility’s prepared worksheet that demonstrates the use of the API.
## User Guide
   1. Open “test API.xlsx” Excel file and switch to the **PARAMS** worksheet

         ![](001.png)

   For this example, we are using two historical endpoints:

   - **/equities/eod/stock-prices**
     - stock price data
   - **/equities/eod/ivx**
     - volatility (IVX)

   which take username, password, date-range:

   - **username**
   - **password**
   - **From** 
   - **To**

   and stock symbol: 

   - **ticker** 

   parameters.

   The IVX, short for Implied Volatility Index, is calculated by using raw IV data with a proprietary weighting technique, factoring the delta and vega of each option.

   For more information on the IVX, please visit IVolatility’s documentation (<https://redocly.github.io/redoc/?nocors&url=https://restapi.ivolatility.com/api-docs#tag/EOD-Equities/operation/EOD%20Equities%20IVX>), as well as Wikipedia (<https://en.wikipedia.org/wiki/IVX>).

   Within IVolatility’s documentation, you may also find more information on the EOD stock price data (<https://redocly.github.io/redoc/?nocors&url=https://restapi.ivolatility.com/api-docs#tag/EOD-Equities/operation/EOD%20Equity%20Prices>). 

   2. After inputting the necessary parameters, from the menu, select **Data** and click on **Refresh All**

         ![](002.png)

   3. The retrieved stock price and volatility data from the IVolatility API will be loaded into the corresponding **stock-prices** and **table\_ivx** worksheets:

   - **stock-prices**

         ![](003.png)

      Next to the price data, the table also contains logarithmic returns for each of the stocks, indicated by “\_log”.



   - **table-ivx**

         ![](004.png)

   4. Within the **Charts** worksheet, you will find the following helpful metrics:
     - log-scale price chart of all the stocks in the portfolio:

         ![](005.png)

     - correlation matrix:

         ![](006.png)


         The correlation matrix can be calculated with Excel’s **Data Analysis** toolkit:

         ![](007.png)


         After clicking on **Data Analysis**, select **Correlation** and press **OK**:

         ![](008.png)


         The correlation matrix tool will pop up:

         ![](009.png)


         You will need to select the **Input Range** and **Output Range**, then press **OK**:

         ![](010.png)


         Excel will then calculate the correlation matrix and output it where you’ve specified. And, after a bit of formatting, it may look similar to the above correlation matrix.

   - risk-return chart of the stocks in the portfolio:

         ![](011.png)


   5. On **Sheet3**, with the use of regression modelling and the correlation matrix, you will find all of the portfolio calculations on the retrieved datasets. They include the annualized returns (APR), volatility (IVX and standard deviations), R2 scores and correlations of each of the stocks with respect to the Nasdaq index.

         ![](012.png)

   6. And, finally, in the **MAIN** worksheet, you will find the individual stocks’, as well as the aggregate portfolio’s, alpha and beta stats.

         ![](013.png)

