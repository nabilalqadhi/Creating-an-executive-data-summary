# Overview
In the exercise Creating an Executive Data Summary, you were asked to put into practice a range of formatting and calculation techniques that you have learned over the past couple of weeks. 

In this exercise was to create an executive summary of a business’s month-by-month profit margin performance for Quarter 1 compared to the same period in the previous year. 

You were required to make formatting changes that would make the worksheet easier to read and ensure the requested analysis results stand out. In addition, you were asked to create formulas that had to: 

### 1- Produce customized quarter-one monthly totals across two business years using the SUMIF function.

### 2- Calculate the percentage difference for each month in quarter 1, 2023, from the same period in 2022.

### 3- use a logical function to test the order value and display the correct tax amount.

*** This reading provides you with a step-by-step guide for identifying these results.

### Step 1: Download the File
1- You downloaded and opened the Microsoft Excel workbook Quarter One Report.xlsx. The workbook contains only one worksheet called Summary.

The sheet contained sales information for specific products spread over two years. It included wholesale and retail prices for each product and the numbers sold. You had to reformat information so that the worksheet showed only the required data and displayed it effectively.

### Step 2: Add and format headings.
1- The first task was to widen column A as it was not wide enough to display the month names in cells A12 to A14 correctly. By dragging or double-clicking the vertical line between the initial letters identifying columns A and B, it was possible to resize the column. 

You then added a new blank column to the left of column E. Column E was titled Product ID. When inserting columns, it is important to remember that the new column is added to the left of the cursor position. You highlighted column E and then chose Insert column from the cells group on the Home ribbon or the right-click shortcut menu.
2- Next, you selected cell A4 and typed the heading TOTAL Q1 SALES. Then you selected cell A10 and typed the heading Q1 MONTHLY TOTALS.

3- When the heading had been added to A4 and A10 you applied a range of formatting choices to ensure that the headings had impact. The quickest approach was to apply the formatting manually to cell A4 and then use the Format painter feature to copy the same look to cell A10.

With the heading in A4 selected, you used the Font and Size dropdown choices to adjust the appearance and size of the text. In the same Font section, you selected the Bold option to bold the text. Then you used the Fill color choice to change the background color of the cell.

The Merge & Center choice is in the Alignment section. This option helps to make headings appear professional by centering them above the table of information they relate to. 

You began by typing headings in the cell at the left edge of the area that the heading had to be centered in. Then you highlighted from cell A4 to the right edge. You selected the Merge & center button to merge the selected cells into one cell and display the heading in the center. 

The Format painter feature is a useful tool to copy formats from one cell to another quickly. It copies the font, color, and Merge & center settings.

You positioned the cursor on cell A4, which was the cell that had the formatting you wished to copy. Then you selected the Format painter button. When the mouse pointer assumed a paintbrush shape, you selected cell A10 to copy the formatting.

4- The headings in B5, C5 and D5 also needed to stand out from other content. You selected these cells and then applied Bold formatting and selected the Wrap Text choice to position the headings correctly in the cells. You then used the Format painter to apply the same formatting to B11, C11 and D11. 

### Step 3: Customize and reorganize how the data is displayed.

1- The entries in column G were in block capitals but you were asked to change the case. In cell H2, you created a formula using the PROPER function to copy the product name in cell G2 

The PROPER function is one of a collection of functions that can be used to change the case of text. It will format text as lowercase with a capital at the beginning of each word. In this worksheet, the correct syntax for this formula is:

=PROPER(G2)

When this formula was applied, the result in cell H2 was Mountain Bikes.

his formula had to be copied down column H using the Autofill feature. You first positioned the cursor on cell H2. There was a block of data to the left in column G, so you could use the double-click shortcut on the bottom right corner of the cursor to use the Autofill shortcut to copy the formula down column H. 

Once the formula was copied down column H and had generated results, the formulas were no longer necessary.

You selected the block of results, chose Copy on the Clipboard section of the Home ribbon, and then selected Paste Values from the dropdown on the Paste button. These options would also have been available on the right-click shortcut menu.

Once the formulas had been removed, you deleted column G by selecting the column and then choosing the Delete option from the cells section of the Home ribbon or from the shortcut right-click menu.


2- The block of sales data had to be sorted so that the row order would be suitable for the monthly total calculations required in a later step. The block had to be sorted by date oldest to newest without sorting the data in columns A to D. The presence of a blank column, column E, would normally prevent this from happening if the data is sorted using the Quick sort shortcut choices in the Data ribbon.

However,  you ensured that Excel would sort the correct content by highlighting the data before sorting it. If you were using a standard keyboard, you could have positioned the cursor on cell F2 and then used the Ctrl+Shift+End key combination to highlight the correct block of data quickly.

3- Some of the data was not relevant for the summary required so you highlighted column F and selected the Hide and Unhide choice on the Format dropdown in the Font section of the Home ribbon. Alternatively, you may have used the right-click shortcut menu to hide the column quickly. You repeated this process for columns S to Y.

If there is data in a worksheet that does not need to be displayed, then hiding columns removes the content from view without deleting it. It is important that you remember that this is not a security measure, as the columns can easily be unhidden. However, it is a useful technique when screens are shared in presentations or meetings. 

You selected the Sort choice on the Data ribbon because the data was already highlighted. The Sort dialog should automatically have been aware that row one contained headings. In the Sort By drop-down, you selected Order Date, and in the Order drop-down, Oldest to Newest.

4- To make reading the data a smoother process, you were asked to freeze both rows and columns on the screen.

You positioned the cursor on cell G2 and selected the View tab to display the ribbon. You opened the Freeze dropdown and selected the Freeze Panes option. Columns A to E and row 1 which were above and to the left of the cursor, were now stay frozen on the screen. 

Remember that, with Freeze Panes, the position of the cursor is used to determine the areas of the screen that should remain static.

### Step 4: Use formulas to create new row information.

1- You were asked to create a formula in K2 using MONTH and a formula in L2 using YEAR to extract the two component parts of the date in J2. These formulas also needed to be copied down as far as row 246.

The MONTH and YEAR functions are in the Date and Time category of functions. It will extract the stated element from the cell entry in J2, which is formatted as a date. 

The syntax for the formula in K2 was:

=MONTH(J2)

The result should be 1.
The syntax for the formula in L2 should be:

=YEAR(J2)

The result should read 2022.


You then copied down the two formulas using the Autofill double-click shortcut or Copy and paste.


2- In P2, you created a standard multiplication formula that multiplied the retail price by the order quantity. You then copied the formula using Autofill or Copy and paste.

The formula in P2 should read:

=N2*O2

The result should be 2,400.
3- In cell Q2, you created a formula using an IF function that calculated if tax was due on the amount in P2. The IF function had to check if the amount in P2 was over 2000. If it was, then the amount in P2 had to be multiplied by 5%. If it was not, then cell Q2  should display a 0. 

The IF formula in P2 should read:

=IF(P2>2000,P2*5%,0)

Here the Value if true action is a formula embedded in the larger logical formula. The percentage calculation is processed because the logical test for the IF returns a value of TRUE. It is possible to have calculations as both the Value if true and the Value if false actions. 

The result of this formula should be 120 which is 5% of the value in P2.

### Step 5: Create formulas to calculate and compare the profit margin across two years.

1- In cell B6, you created a SUMIF  formula to sum the sales values for 2022. The sales values were in the range R2 to R246. The criteria range was the range  L2 to L246. You then created a similar formula in cell C6 with the same cell ranges but changed the criteria to 2023.

The formula in B6 should read:

=SUMIF(L2:L246,2022,R2:R246)

The result in B6 should be $330,500.

The formula in C6 should read:

=SUMIF(L2:L246,2023,R2:R246)

The result should be $453,830.


2- In cell B12 you created a SUMIF to sum the range R2 to R103 if there was the number 1 in the criteria range K2 to K103. You also added dollar signs to the R and K cell references so that the formula could be copied down.

Rows 2 to 103 contained the entries for 2022 because the data had already been sorted in date order. To obtain a total for only the January 2022 entries, the SUMIF formula used a criteria range of K2 to K103, and a sum range of R2 to R103. 
You had already created the entries in column K using the MONTH function to extract the month number. Asking Excel to match criteria 1 meant that it only included entries for January in the total. (The dollar signs added to the criteria range and sum range cell references were preparation for the next task which was to copy the formula.)

The formula in B12 should read:

=SUMIF($K$2:$K$103,1,$R$2:$R$103)

The result is $101,595.


3- You copied the formulas from cell B12 into cells B13 and B14. In the B13 copy you changed the criteria to 2 and in the B14 copy changed the criteria to 3.

The formula in B13 should read:

=SUMIF($K$2:$K$103,2,$R$2:$R$103)

The result should read $113,445.

The formula in B14 should read:

=SUMIF($K$2:$K$103,3,$R$2:$R$103)

The result should be $115,460.



4- In cell C12  you created a formula using  SUMIF. This formula summed the range R104 to R246 if it said 1 in the range K104 to K246. You added dollar signs to the R and K cell references.

Rows 104 to 246 contained the entries for 2023 because the data had been sorted in date order. To obtain a total for the January 2023 entries, the SUMIF formula used a criteria range of K104 to K246, the cells holding the month numbers, and a sum range of R104 to R246. 

Matching criteria 1 meant that Excel would only include entries for January in the total. The dollar signs added to the criteria range and sum range cell references allowed for the formula to be copied down. 

The formula in C12 should read: 

=SUMIF($K$104:$K$246,1,$R$104:$R$246)

The result should be $143,555.

5- You then copied the formula to C13 and C14. In the C13 copy, you changed the criteria to 2, and in the C14 copy, you changed the criteria to 3.

The formula in C13 should read:

=SUMIF($K$104:$K$246,2,$R$104:$R$246)

The result should be $145,535.

The formula in C14 should read:

=SUMIF($K$104:$K$246,3,$R$104:$R$246)

The result should be $164,740.

6- u created a Percentage difference formula in D6 which showed the percentage by which sales increased in 2023.

To determine the percentage difference between the results for 2022 and 2023, the total for 2022 first had to be subtracted from the 2023 total. The result had then to be divided by the result for 2022. This formula needed parentheses since the subtraction had to be done first. The cell was already formatted as a percentage, so the result displayed correctly.

The formula in D6 should read:

=(C6-B6)/B6

The result is 37.32%.
7- You created a similar formula in D12 and copied the calculation in D12 down to D14.

The formula in D12 should read:

=(C12-B12)/B12

The result should be 41.30%.

The formula in D13 should read:

=(C13-B13)/B13 
The result should be 28.29%.

The formula in D14 should read:

=(C14-B14)/B14

The result should be 42.68%.

### Conclusion 
In this exercise, you were tasked with using a variety of formatting skills and with creating a range of formulas to create data columns in a spreadsheet and calculate customized totals.

You transformed the standard sales data into a data summary that could be used to inform and drive business decisions. Well done! 
