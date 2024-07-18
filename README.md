# VBA-challenge
VBA Challenge Report

Homework Assignment Description

The homework assignment involves using VBA (Visual Basic for Applications) scripting to analyze stock market data. The tasks include:

1. Looping through stock data for each quarter.
2. Calculating and outputting:
   - Ticker symbol.
   - Quarterly change from the opening price at the beginning of the quarter to the closing price at the end of the quarter.
   - Percentage change over the quarter.
   - Total stock volume.
3. Identifying and outputting:
   - The stock with the greatest percentage increase.
   - The stock with the greatest percentage decrease.
   - The stock with the greatest total volume.
4. Applying conditional formatting:
   - Highlighting positive changes in green.
   - Highlighting negative changes in red.
5. Ensuring the script can run on multiple worksheets simultaneously.
6. Submitting the script, results screenshots, and a README file to a GitHub repository.

Process and Implementation

1. **Environment Setup**: 
   - Enabled the Developer tab in Excel.
   - Opened the VBA editor.

2. **Script Development**:
   - Created a VBA script that loops through each worksheet, processes stock data, and calculates the required values.
   - Applied conditional formatting to highlight positive and negative changes.
   - Implemented functionality to determine the greatest percentage increase, decrease, and total volume.

3. **Testing**:
   - Tested the script on a smaller dataset (`alphabetical_testing.xlsx`) to ensure accuracy and performance.
   - Verified that the script correctly processes multiple worksheets and outputs the required data.

Findings

1. **Stock Data Analysis**:
   - Calculated the quarterly change, percentage change, and total stock volume for each ticker.
   - Applied conditional formatting to visually distinguish positive and negative changes.
   
2. **Key Metrics**:
   - Greatest Percentage Increase: MSE with 47.01%.
   - Greatest Percentage Decrease: VNG with -77.83%.
   - Greatest Total Volume: HK with 2.64E+11.

Outputs

The script generated the following key outputs:

1. **Ticker Data Summary**:
   - For each ticker, the quarterly change, percentage change, and total volume were calculated and displayed in a summary table.

2. **Conditional Formatting**:
   - Positive changes were highlighted in green.
   - Negative changes were highlighted in red.

3. **Greatest Values**:
   - The script identified and displayed the tickers with the greatest percentage increase, decrease, and total volume.

Example

An example of the output for ticker AAF:
- Ticker: AAF
- Quarterly Change: 0.05
- Percent Change: 1.00%
- Total Stock Volume: 752812510

Greatest Values:
- Greatest % Increase: Ticker - MSE, Value - 47.01%
- Greatest % Decrease: Ticker - VNG, Value - -77.83%
- Greatest Total Volume: Ticker - HK, Value - 2.64E+11

Submission

All the required files were uploaded to the GitHub repository named `VBA-challenge`, including:
- VBA script files.
- Screenshots of the results.
- README file with instructions and additional information.

Conclusion

The VBA script effectively analyzed the stock market data, calculated the required metrics, and highlighted significant changes through conditional formatting. The results provided insights into the stocks with the most notable performance changes and volumes. The project demonstrated the power of VBA in automating and simplifying data analysis tasks in Excel.
