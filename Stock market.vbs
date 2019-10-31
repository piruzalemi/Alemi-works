Sub TickerStock():

        
        Dim tickerName As String
        Dim nextTickerName As String
        
        Dim beginOpenStock As Double
        Dim beginOpenHold As Double
        Dim endOpenStock As Double
        Dim changeOpen As Double
        
        
        Dim highStock As Integer
        Dim lowStock As Integer
        Dim closeStock As Integer
        Dim volStock As Long
        Dim totalVolStock As Currency
        
        Dim changeYearly As Long
        Dim percentOpenStock As Double
       
        Dim fromDate As String
        Dim toDate As String
        
        Dim dateYYYY As String
        Dim holdYYYY As String
        Dim dateMM As String
        Dim dateDD As String
        
        Dim formatDate1 As Date
        Dim formatDate2 As Date
        
        Dim Date1 As Date
        Dim Date2 As Date
        
        Dim numDays As Long
        
        Dim lastRow As Long
        Dim i As Long
        
         '------------------------------------
        ' Max & Min extracts
        ' -----------------------------------
        
        Dim greatestVolStock As Currency
        Dim minPercentPrice As Double
        Dim maxPercentPrice As Double
    
        
        Dim minTickerName As String
        Dim maxTickerName As String
        Dim greatestNameVol As String
        
        
        
        
        greatestVolStock = 0
        maxPercentPrice = 0#
        minPercentPrice = 0#
        
        
    ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
    For Each ws In Worksheets

        ' --------------------------------------------
        '
        ' --------------------------------------------

        ' Created a Variable to Hold File Name, Last Row, Last Column, and Year
        Dim WorksheetName As String

        ' Determine the Last Row
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

        ' Grabbed the WorksheetName
        WorksheetName = ws.Name
        MsgBox (WorksheetName & " " & lastRow)

        
        ' Initialize the first & next ticker names & hold the first open stock price
          i = 2
          k = 2
          totalVolStock = 0
          
        ' Only for the first time Beginning of Open Stock Price is set here
        ' The next Open Stock is set in the Else Statement
        
          beginOpenStock = ws.Cells(i, 3).Value
       
          
                   fromDate = ws.Cells(i, 2).Value
                            dateYYYY = Left(fromDate, 4)
                            dateMM = Mid(fromDate, 5, 2)
                            dateDD = Right(fromDate, 2)
                            formatDate1 = dateMM + "/" + dateDD + "/" + dateYYYY
                            Date1 = Format(formatDate1, "mm/dd/yy")
                            ' MsgBox ("fromdate is Date1= " & Date1)
         
         ' --------------------------------------------------------------------
         ' Main Loop: for each ticker, summarize the results
         ' --------------------------------------------------------------------
        For i = 2 To lastRow
        '   For i = 2 To 4000
         
            ' Check if we are still within the same ticker, if it is not, select the ending Stock
              tickerName = ws.Cells(i, 1).Value
              nextTickerName = ws.Cells(i + 1, 1).Value
         
                 If tickerName = nextTickerName Then
                       
                         volStock = ws.Cells(i, 7).Value
                         totalVolStock = totalVolStock + volStock
                         ws.Cells(i, 19).Value = totalVolStock
                         
                         endOpenStock = ws.Cells(i + 1, 3).Value
                               toDate = ws.Cells(i + 1, 2)
            
                              dateYYYY = Left(toDate, 4)
                              dateMM = Mid(toDate, 5, 2)
                              dateDD = Right(toDate, 2)
                              formatDate2 = dateMM + "/" + dateDD + "/" + dateYYYY
                              
                              Date2 = Format(formatDate2, "mm/dd/yy")
                             
                             ' MsgBox ("To Date is Date2= " & Date2)
                              
                              numDays = DateDiff("d", Date1, Date2) + 1
                              ' MsgBox ("numDays...............>>>>>: = " & numDays)
                             
                          
                 Else
                      
                    '----------------------------------------------------------------------
                    ' Must have reached a new Ticker, so Summarize the previous ticker and write it out
                    '----------------------------------------------------------------------
                       changeOpen = endOpenStock - beginOpenStock
                       
                                If beginOpenStock > 0 Then
                                   percentOpenStock = changeOpen / beginOpenStock
                                Else
                                   percentOpenStock = 0
                                End If
                       
                                volStock = ws.Cells(i, 7).Value
                                totalVolStock = totalVolStock + volStock
                                
                       ' -----------------------------------------------------------------
                       '       Max & Min % Stock Prices
                       ' -----------------------------------------------------------------
                       
                               If percentOpenStock > maxPercentPrice Then
                                  maxPercentPrice = percentOpenStock
                                  maxTickerName = tickerName
                                  End If
                                  
                               If percentOpenStock < minPercentPrice Then
                                  minPercentPrice = percentOpenStock
                                  minTickerName = tickerName
                                  End If
                               
                               If totalVolStock > greatestVolStock Then
                                  greatestVolStock = totalVolStock
                                  greatestNameVol = tickerName
                                  End If
                                  
                                  ' ----------------------------------------------------
                                  ' Write out the Summaries + the last of thr running sums
                                  ' ----------------------------------------------------
                                     
                                     ws.Cells(k, 9).Value = tickerName
                                     ws.Cells(k, 10).Value = beginOpenStock
                                     ws.Cells(k, 11).Value = endOpenStock
                                     ws.Cells(k, 12).Value = changeOpen
                                     ws.Cells(k, 13).NumberFormat = "0.00%"
                                     ws.Cells(k, 13).Value = percentOpenStock
                                     ws.Cells(k, 14).Value = totalVolStock
                                     ws.Cells(k, 14).NumberFormat = "0"
                                     
                                     ws.Cells(k, 16).Value = Date1
                                     ws.Cells(k, 17).Value = Date2
                                     ws.Cells(k, 18).Value = numDays
                                     ws.Cells(i, 19).Value = totalVolStock
                                     ws.Cells(k, 19).NumberFormat = "0"
                                     ws.Cells(i, 20).Value = tickerName
                      
                                      If percentOpenStock < 0 Then
                                          ws.Cells(k, 12).Interior.ColorIndex = 3
                                      Else
                                          ws.Cells(k, 12).Interior.ColorIndex = 4
                                      End If
                                          
                      
                                      k = k + 1
                                      
                                    
                                      totalVolStock = 0
                                       
                                       '--------------------------------------------------------------------------
                                       ' making sure the index i+1 does not go beyond its max(lastRow) for date type to hold
                                       ' This is an if statement within a previous else statement (where we just hit a new Ticker)
                                       '--------------------------------------------------------------------------
                                                If i < (lastRow) Then
                                                       beginOpenStock = ws.Cells(i + 1, 3).Value
                                                             fromDate = ws.Cells(i + 1, 2).Value
                                                                 dateYYYY = Left(fromDate, 4)
                                                                 dateMM = Mid(fromDate, 5, 2)
                                                                 dateDD = Right(fromDate, 2)
                                                                 formatDate1 = dateMM + "/" + dateDD + "/" + dateYYYY
                                                             Date1 = Format(formatDate1, "mm/dd/yy")
                                                         ' MsgBox ("from date1: " & Date1)
                                                End If
                   
                     End If
                  
       Next i
       ' Next is to loop for the next sheet. It makes no difference if we say Next or Next WS'
Next ws
  
        ' Write out the min & Max values
        
                                  Cells(2, 22).Value = maxTickerName
                                  Cells(2, 23).Value = maxPercentPrice
                              
                                  Cells(3, 22).Value = minTickerName
                                  Cells(3, 23).Value = minPercentPrice
                                  
                                  Cells(4, 22).Value = greatestNameVol
                                  Cells(4, 23).Value = greatestVolStock
                                 
                                  
End Sub


