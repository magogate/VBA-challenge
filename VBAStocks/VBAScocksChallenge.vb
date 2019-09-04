'Written By : Mandar Gogate
'Written On : 09/02/2019
'Modified On : 09/03/2019
	'Reinitialize the value of total sum at line#102
'Referances are as below
'1. To get last row and iterate through differance worksheets and activate worksheet
' https://www.youtube.com/watch?v=Jyls2ZTIqUo&list=PLMM5e6ePXVqK7KKnebxZ4hMRokGTpe0Zw&index=17
' https://docs.microsoft.com/en-us/office/vba/api/excel.worksheet.rows
' https://docs.microsoft.com/en-us/office/vba/api/excel.range.end
' https://docs.microsoft.com/en-us/office/vba/api/excel.xldirection
' https://www.youtube.com/watch?v=rHD21IlKePY&list=PLMM5e6ePXVqK7KKnebxZ4hMRokGTpe0Zw&index=34
' https://docs.microsoft.com/en-us/office/vba/api/excel.worksheet.activate(method)
'2. To sort data
' https://www.youtube.com/watch?v=9mgK9yiAOZs
'3. To color cells
' https://www.youtube.com/watch?v=F29G18GdTAQ&list=PLMM5e6ePXVqK7KKnebxZ4hMRokGTpe0Zw&index=4
'4. ArrayList
' https://docs.microsoft.com/en-us/dotnet/api/system.collections.arraylist?view=netframework-4.8
' Note: ternary operate in VB won't help in case of division 0; as it evaluates both conditions always

Dim myArrayListOuter As Object
    
Sub processData()

    Dim myData As Range
    Dim myColumn As Range
   
    Dim myCurrentTicker As String
    Dim myOldTicker As String
    
    Dim myArrayList As Object
    
    
    Dim tickerSum As Double
    Dim isNewTicker As String
    
    Dim op As Double
    Dim cl As Double
    Dim wsName As String
      
    ' iterating through all worksheets available in workbook
    For ws = 1 To Worksheets.Count
        ' initializing ArrayList for each separate worksheet
        Set myArrayListOuter = CreateObject("System.Collections.ArrayList")
        
        myCurrentTicker = "None"
        myOldTicker = "None"
        isNewTicker = "Yes"
        tickerSum = 0
        
        
        wsName = Worksheets(ws).Name
            'activating each worksheet in order to iterate through data
            Worksheets(wsName).Activate
            'going to very first cell of worksheet to get last row to find out range
            Cells(1, 1).Select
            'Sorting data by Ticker & Dates as data can be unsorted due to which following logic won't work
            Set myData = Range("A:G")
            Set myColumn = Range("A:A")
            Set myColumn2 = Range("B:B")
            myData.Sort key1:=myColumn, order1:=xlAscending, key2:=myColumn2, Order2:=xlAscending, Header:=xlYes
            
            'fetching last row so that we can iterate through each & every cell
            lastRow = Cells(Rows.Count, 1).End(xlUp).Row
            
            'rowCnt = 1 is a header so starting with 2
            For rowCnt = 2 To lastRow
                myCurrentTicker = Cells(rowCnt, 1).Value
                tickerSum = tickerSum + Cells(rowCnt, 7).Value
                
                If (rowCnt = 2) Then
                    op = Cells(rowCnt, 3).Value
                End If
                
                'if ticker changes, then adding last ticker in an array list.
                If ((myCurrentTicker <> myOldTicker And myOldTicker <> "None") Or rowCnt = lastRow) Then
                    'separate arrayList for each ticker
                    Set myArrayList = CreateObject("System.Collections.ArrayList")
                    
                    'if we are not proecssing last row we need to take values from one row before
                    If (rowCnt = lastRow) Then
                        cl = Cells(rowCnt, 6).Value
						tickerSum = tickerSum - 0
                    Else
                        cl = Cells(rowCnt - 1, 6).Value
						tickerSum = tickerSum - Cells(rowCnt, 7).Value
                    End If
                    
                    'Adding elements in ArrayList
                    '1. Ticker Name
                    myArrayList.Add myOldTicker
                    '2.Total Stock Volume
                    myArrayList.Add tickerSum
                    '3. Opening Value
                    myArrayList.Add op
                    '4. Closing Value
                    myArrayList.Add cl
                    
                    myArrayListOuter.Add myArrayList
                    
                    tickerSum = 0
					tickerSum = tickerSum + Cells(rowCnt, 7).Value
                End If
                
                If (myCurrentTicker <> myOldTicker) Then
                    op = Cells(rowCnt, 3).Value
                End If
                
                myOldTicker = myCurrentTicker
            Next
            
            printData
    Next
    'Once all worksheets are processed
    MsgBox ("Process Complete")
End Sub

Sub printData()

    Dim maxPerIncTicker As String
    Dim maxPerInc As Double
    Dim oldPerInc As Double
    Dim minPerInc As Double
    
    maxPerIncTicker = "None"
    minPerIncTicker = "None"
    maxStockVolTicker = "None"
    maxPerInc = 0
    minPerInc = 0
    oldPerInc = 0
    maxStockVol = 0
    oldStockVol = 0

    ' Adding headers
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly Change"
    Cells(1, 11).Value = "Percentage Change"
    Cells(1, 12).Value = "Total Stock Volume"

   For cnt = 0 To myArrayListOuter.Count - 1
   
    yearlyChange = myArrayListOuter(cnt)(3) - myArrayListOuter(cnt)(2)
   
    If (cnt = 0) Then
     maxStockVol = myArrayListOuter(cnt)(1)
         If (myArrayListOuter(cnt)(2) = 0) Then
             maxPerInc = 0
             minPerInc = 0
         Else
             maxPerInc = ((yearlyChange) / myArrayListOuter(cnt)(2)) * 100
             minPerInc = ((yearlyChange) / myArrayListOuter(cnt)(2)) * 100
         End If
    End If
    
    '1. Ticker
    Cells(cnt + 2, 9).Value = myArrayListOuter(cnt)(0)
    
    '2. Yearly Change
    Cells(cnt + 2, 10).Value = yearlyChange
    
    If yearlyChange < 0 Then
        Cells(cnt + 2, 10).Interior.ColorIndex = 3
        Else
        Cells(cnt + 2, 10).Interior.ColorIndex = 4
    End If
    
    '3. Percent Change
    If (myArrayListOuter(cnt)(2) = 0) Then
        Cells(cnt + 2, 11).Value = "0%"
        oldPerInc = 0
    Else
        Cells(cnt + 2, 11).Value = CStr(Round((((yearlyChange) / myArrayListOuter(cnt)(2)) * 100), 2)) + "%"
        oldPerInc = ((yearlyChange) / myArrayListOuter(cnt)(2)) * 100
    End If
    
    If (oldPerInc > maxPerInc) Then
        maxPerInc = oldPerInc
        maxPerIncTicker = Cells(cnt + 2, 9).Value
    End If
    
    If (oldPerInc < minPerInc) Then
        minPerInc = oldPerInc
        minPerIncTicker = Cells(cnt + 2, 9).Value
    End If
    
    '4. Total Stock Volume
    Cells(cnt + 2, 12).Value = myArrayListOuter(cnt)(1)
    oldStockVol = Cells(cnt + 2, 12).Value
    If (maxStockVol < oldStockVol) Then
        maxStockVol = oldStockVol
        maxStockVolTicker = Cells(cnt + 2, 9).Value
    End If
    
   Next
   
   'print row header
   Range("O2").Value = "Greatest % Increase"
   Range("O3").Value = "Greatest % Decrease"
   Range("O4").Value = "Greatest Total Volume"
   
   'print column header
   Range("P1").Value = "Ticker"
   Range("Q1").Value = "Value"
      
   'print actual values
   Range("P2") = maxPerIncTicker
   Range("Q2") = CStr(Round(maxPerInc, 2)) + "%"
   
   'print actual values
   Range("P3") = minPerIncTicker
   Range("Q3") = CStr(Round(minPerInc, 2)) + "%"
   
   'print stock values
   Range("P4") = maxStockVolTicker
   Range("Q4") = maxStockVol
   
End Sub



