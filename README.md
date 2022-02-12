# VBA_Challenge

##Overview
The overview of this analysis is to refactor the stock information from year 2017 and 2018 using one code to loop through two different data.  And also by refactoring, the data was verified if the code ran faster. 

## Result

The assignment was very similar to the module in terms of steps that I had to take to refactor the data. However, there were few differences in the coding where it required adjustment to the data. 
First, tickerIndex was introduced which was set to zero. Then, setting up three output arrays were slightly different since tickerVolumes needed to set As Long. Because I was given Dim Tickers(12) As String in the beginning because there were twelve tickers, arrays were consistently written. And each tickerVolumes, tickerStartingPrices, tickerEndingPrices, and tickerIndex code was expressed based with tickerindex so that information on each tickers can be obtained. Lastly, by setting the output, ticker, Total Daily Volume, and Return were organized for each year. 

###Coding
    tickerIndex = 0
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
   
    For i = 0 To 11

        tickerVolumes(i) = 0
        tickerStartingPrices(i) = 0
        tickerEndingPrices(i) = 0
    Next i
       For i = 2 To RowCount
    
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
           
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
         
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
        End If
         
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
        
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        End If
          
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
                tickerIndex = tickerIndex + 1
                
              
        End If
        Next i
    For i = 0 To 11
      Worksheets("All Stocks Analysis").Activate
       Cells(4 + i, 1).Value = tickers(i)
       Cells(4 + i, 2).Value = tickerVolumes(i)
       Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
  
    Next i

###Images
Finally, by adding colors to the chart, timer was also created to see how quickly VBA can organize and analyze the data. 

![2017_Stock](https://https://github.com/sangyoo1021/VBA_Challenge/blob/main/Resources/VBA_Challenge_2017.png)

![2018_Stock](https://https://github.com/sangyoo1021/VBA_Challenge/blob/main/Resources/VBA_Challenge_2018.png)
 
##Summary

###Pros and Cons
VBA is very strong tool that can be managed through coding. The biggest advantage refactoring is, if the coding is well organized and clear, all the information is very accessible in one single page. Also, the analyzing data can be faster. However, there are some disadvantages that follow. If the data becomes larger, there needs to be more coding required which sometimes requires more debugging. With simple typo, running the codes will not always work. 
