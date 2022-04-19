# An Analysis of Green Stock Data using Excel and VBA
### *Performing analysis on green stock data to uncover trends*
## Overview of Project
### *Purpose:* 
#### The purpose of this project is to analyze the performance of 12 green stocks, investments associated with companies that are somehow involved in the protection of the environment, during the years 2017 and 2018. From obtaining this information we are given insight into which stocks would likely be the most lucrative to invest in based on past performance. We performed this analysis with our green_stocks dataset file. 

### *Background:*
#### We prepared this workbook for our client, Steve. At first, he was most interested in the Daqo New Energy Corp. stock, DAQO (Ticker: DQ), because his parents were investing in it and he wanted to know whether or not it was a good investment for them. When we presented him with results from the DQAnalysis tab showing that the DAQO stock had dropped over 63% in 2018, he then requested information on the performance of additional stocks in order to provide his parents with some alternative and possibly better investment options.

#### We added a tab for All Stocks Analysis so that Steve can now analyze the entire green stock dataset with the click of a button. He can input whether he would like to see the performance for the year 2017 or the year 2018. We were able to run the VBA code and record the time it took in order to check for efficiency. 

### *Challenge:*
#### The challenge of this project is to refractor our code looping through all the data one time in order to collect the same information and then see if refactoring it made it run faster. 

## Results
### *Analysis:*

#### Analysis of the results showed that while DAQO stock did have a great year in 2017, it did not perform well in 2018 with a 63% drop as compared to the other 11 green stocks listed in this dataset. In fact, 2017 was an up year for all except one of these stocks. Upon further analysis, there were only 2 stocks that performed very well in both years. These 2 stocks were Sunrun Inc. (Ticker: RUN) and Enphase Energy (Ticker: ENPH). Both of these had over 100% return in 2017 and over 80% return in 2018. 

#### Our original VBA code was fairly efficient for the dataset we were working with. However, we found through refactoring it could be made more efficient and scalable. The original 2017 code ran in .72 seconds and the 2018 code ran in .71 seconds. Although the code was running smoothly for these 12 stocks, if our client were to ask us to expand the dataset to include hundreds or thousands of stocks, the efficiency of the code could be significantly slower and take a much longer time to execute. After Refactoring, the 2017 code ran in .12 seconds and the 2018 code ran in .13 seconds. Both of these were a considerable improvement.

#### We included screenshots below from our analysis. 

#### DQAnalysis tab screenshot showing that the DAQO stock dropped over 63% in 2018 - 
![DQ_Analysis_2018](https://user-images.githubusercontent.com/92938054/142785594-77fe0d30-b5c0-4d7a-b8cc-6800bb96a46c.png)
#### DQAnalysis tab VBA code - 
```
Sub DQAnalysis()
    
    Worksheets("DQAnalysis").Activate
    
    Range("A1").Value = "DAQO (Ticker: DQ)"
    
    'Create a header row
    Cells(3, 1).Value = "Year"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    
    Worksheets("2018").Activate
    
    rowStart = 2
    'DELETE: rowEnd = 3013
    'rowEnd code taken from https://stackoverflow.com/questions/18088729/row-count-where-data-exists
    rowEnd = Cells(Rows.Count, "A").End(xlUp).Row
    'set initial volume to zero
    totalVolume = 0
    
    Dim startingPrice As Double
    Dim endingPrice As Double

    'Establish the number of rows to loop over
    rowStart = 2
    rowEnd = Cells(Rows.Count, "A").End(xlUp).Row

    'loop over all the rows
    For i = rowStart To rowEnd

        If Cells(i, 1).Value = "DQ" Then

            'increase totalVolume by the value in the current row
            totalVolume = totalVolume + Cells(i, 8).Value

        End If

        If Cells(i - 1, 1).Value <> "DQ" And Cells(i, 1).Value = "DQ" Then

            startingPrice = Cells(i, 6).Value

        End If

        If Cells(i + 1, 1).Value <> "DQ" And Cells(i, 1).Value = "DQ" Then

            endingPrice = Cells(i, 6).Value

        End If

    Next i

    Worksheets("DQAnalysis").Activate
    Cells(4, 1).Value = 2018
    Cells(4, 2).Value = totalVolume
    Cells(4, 3).Value = (endingPrice / startingPrice) - 1
    

End Sub
```
#### Original All Stocks Analysis tab screenshot - 2017 -
![green_stocks_2017](https://user-images.githubusercontent.com/92938054/142788931-ec5f4d52-1ab9-40cc-9dd8-095fe590a2c1.png)

#### Original All Stocks Analysis tab screenshot - 2018 -
![green_stocks_2018](https://user-images.githubusercontent.com/92938054/142788994-3b47ae18-a8ff-4270-9d6c-45b2f20c22d6.png)

#### Original All Stocks Analysis tab ("yearValueAnalysis") Code -
```
Sub yearValueAnalysis()

    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer


        '1) Format the output sheet on All Stocks Analysis worksheet
    
        Worksheets("All Stocks Analysis").Activate
        
        Range("A1").Value = "All Stocks (" + yearValue + ")"

        'Create a header row
        Cells(3, 1).Value = "Ticker"
        Cells(3, 2).Value = "Total Daily Volume"
        Cells(3, 3).Value = "Return"
      
        '2) Initialize array of all tickers
   
        Dim tickers(11) As String
    
        tickers(0) = "AY"
        tickers(1) = "CSIQ"
        tickers(2) = "DQ"
        tickers(3) = "ENPH"
        tickers(4) = "FSLR"
        tickers(5) = "HASI"
        tickers(6) = "JKS"
        tickers(7) = "RUN"
        tickers(8) = "SEDG"
        tickers(9) = "SPWR"
        tickers(10) = "TERP"
        tickers(11) = "VSLR"

        '3a) Initialize variables for starting price and ending price

        Dim startingPrice As Single
        Dim endingPrice As Single
    

        '3b) Activate data worksheet

        Worksheets(yearValue).Activate

        '3c) Get the number of rows to loop over

        RowCount = Cells(Rows.Count, "A").End(xlUp).Row

        '4) Loop through tickers
   
        For i = 0 To 11
       
        ticker = tickers(i)
        totalVolume = 0

            '5) loop through rows in the data

            Worksheets(yearValue).Activate
            For j = 2 To RowCount
       
                '5a) Get total volume for current ticker
    
                If Cells(j, 1).Value = ticker Then

                    'increase totalVolume by the value in the current row
                    totalVolume = totalVolume + Cells(j, 8).Value

                End If
            
                '5b) get starting price for current ticker
    
                If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

                    startingPrice = Cells(j, 6).Value

                End If
    
                '5c) get ending price for current ticker
    
                If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

                    endingPrice = Cells(j, 6).Value

                End If
        
            Next j
       
       
        '6) Output data for current ticker

        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = (endingPrice / startingPrice) - 1
    
        Next i
    
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub
```
#### After clicking on "Clear Worksheet" button and then clicking "Run Analysis for All Stocks Refactored" on All Stocks Analysis tab - 
![All_Stks_Refact_clrd_popupbox](https://user-images.githubusercontent.com/92938054/142789318-fb5d5bf1-17ee-476b-accb-bd40c2083c55.png)
#### Inputting year of 2017 and clicking "OK" on All Stocks Analysis tab - 
![All_Stks_Refact_clrd_popupbox 2017](https://user-images.githubusercontent.com/92938054/142789559-e353befe-45ed-49cd-baa2-3b5f40f7dbed.png)
#### Refactored All Stocks Analysis tab screenshot - 2017 -
##### *Original code ran in .72 seconds, Refactored code ran in .12 seconds - See also outstanding performance of ENPH and RUN*
![All_Stks_Refact_popupbox with time-121_2017](https://user-images.githubusercontent.com/92938054/142786930-e97231f2-5a78-4646-9b25-4fbd96113161.png)
#### Inputting year of 2018 and clicking "OK" on All Stocks Analysis tab - 
![All_Stks_Refact_clrd_popupbox 2018](https://user-images.githubusercontent.com/92938054/142789647-f567bfbf-dfb2-4286-a757-8bd9ac6ff1ae.png)
#### Refactored All Stocks Analysis tab screenshot - 2018 -
##### *Original code ran in .71 seconds, Refactored code ran in .13 seconds - See also outstanding performance of ENPH and RUN*
![All_Stks_Refact_popupbox with time-136_2018](https://user-images.githubusercontent.com/92938054/142786911-6e2b6f74-04ae-4749-b356-6b1b27c821a6.png)
#### Refactored All Stocks Analysis Code - 
```
Sub AllStocksAnalysisRefactored()

    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stocks (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
    Dim tickers(12) As String
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
      '2a) Create a for loop to initialize the tickerVolumes to zero.
      Worksheets(yearValue).Activate
      
      For i = 0 To 11
      
          tickerVolumes(i) = 0
          tickerStartingPrices(i) = 0
          tickerEndingPrices(i) = 0
          
      Next i
           
          '2b) Loop over all the rows in the spreadsheet.
          For j = 2 To RowCount

            '3a) Increase volume for current ticker
            tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
  
            '3b) Check if the current row is the first row with the selected tickerIndex.
            If Cells(j - 1, 1).Value <> tickers(tickerIndex) And Cells(j, 1).Value = tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(j, 6).Value
        
            End If
        
            '3c) check if the current row is the last row with the selected ticker
            'If the next row's ticker doesn't match, increase the tickerIndex.
            If Cells(j + 1, 1).Value <> tickers(tickerIndex) And Cells(j, 1).Value = tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(j, 6).Value

            End If

            '3d Increase the tickerIndex.
            If Cells(j, 1).Value = tickers(tickerIndex) And Cells(j + 1, 1).Value <> tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1

            End If
        
          Next j
 
      '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
      For i = 0 To 11
             
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = (tickerEndingPrices(i) / tickerStartingPrices(i)) - 1
        
      Next i
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

      For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
        Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
        Cells(i, 3).Interior.Color = vbRed
            
        End If
        
      Next i
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)

End Sub
```
#### Link to Original green_stocks.xlsm
[Green Stocks](https://github.com/ashleydsingleton/stock_analysis/blob/main/green_stocks.xlsm)

#### Link to Refactored VBA_Challenge.xlsm
[VBA Challenge](https://github.com/ashleydsingleton/stock_analysis/blob/main/VBA_Challenge.xlsm)

## Summary
#### By creating the most efficient code possible and without risking the integrity of our data outcomes, developers and analysts can create scalable, more robust and lasting code that can flex to our clients' needs. Refactoring is about making good better. By making your code more efficient, you are improving logic, taking less steps, lowering costs by using less memory and making it easier for future users who could potentially have to pick up where you left off. Refactoring makes code more readable and useable for years to come. 
#### By recording the time it took to run our code before and after refactoring we proved that it was more efficiently ran after refactoring. Upon loading larger datasets with more rows, this will prove to be very beneficial and cut down on time loss. The code was more streamlined after refactoring. It makes sense to me why refactoring is a common and valuable practice when coding. It seems the long-term payoff would far exceed any short-term hassle if there is any. 
