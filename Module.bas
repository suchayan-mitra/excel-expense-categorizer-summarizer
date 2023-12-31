Attribute VB_Name = "Module1"
Function GetCategory(description As String) As String
    ' Dictionary to hold the keywords for each category
    Dim categories As Object
    Set categories = CreateObject("Scripting.Dictionary")
    
    ' Add categories and their associated keywords
    With categories
        .Add "Coffee", Array("TIM HORTONS")
        .Add "Grocery", Array("CENTRAL MEAT MARKET", "INSTACART", "NOFRILLS", "AMAZON.CA", "MARCHELEO'S MARKETPLAC", "STOP 2 SHOP", "DOLLARAMA", "ONKAR FOODS & SPICES", "CANADIAN TIRE GAS BAR", "WAL-MART", "CONVENIENCE STORE", "HASTY MARKET")
        .Add "Drink", Array("LCBO")
        .Add "Electronics", Array("LENOVO", "BEST BUY", "APPLE.COM")
        .Add "Housing - Rental", Array("AIRBNB")
        .Add "Subscription", Array("APPLE.COM", "Amazon.ca Prime")
        .Add "Phone and Mobile", Array("TELUS ONLINE PAYMENT")
        .Add "Online Order - Food", Array("DOORDASH")
        .Add "Online Order - AMZN", Array("AMZN Mktp")
        .Add "Car Rental - Taxi", Array("UBER", "LYFT")
        .Add "Gas", Array("PETRO CANADA", "ESSO CIRCLE K")
        .Add "Transport", Array("METROLINX - GO TRANSIT")
        .Add "Restaurant", Array("KELSEYS", "NEW MIRCHI DHABA")
        .Add "Fee and Charges", Array("OVERLIMIT FEE")
        .Add "Travel", Array("AIR CAN*", "EXPEDIA")
        .Add "MISC - Immigration", Array("IMMIGRATION CANADA")
        .Add "Car Rental", Array("CAR RENTAL")
        ' Add more categories and keywords as needed
    End With
    
    ' Check the description against each category
    Dim key As Variant
    For Each key In categories.Keys
        Dim words As Variant
        words = categories(key)
        Dim word As Variant
        For Each word In words
            If InStr(1, description, word) > 0 Then
                GetCategory = key
                Exit Function
            End If
        Next word
    Next key
    
    ' Default category if no match is found
    GetCategory = "Uncategorized"
End Function

Sub CategorizeTransactions()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("INPUT")

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row ' Assuming 'B' is the Description column

    Dim i As Long
    For i = 2 To lastRow ' Assuming row 1 has headers
        Dim description As String
        description = ws.Cells(i, "B").Value ' Get the description
        
        ' Get the category from the GetCategory function
        ws.Cells(i, "E").Value = GetCategory(description)
    Next i
End Sub

Sub SummarizeAndChartExpenses()
    Dim ws As Worksheet, summarySheet As Worksheet
    Dim lastRow As Long
    Dim pTable As PivotTable
    Dim pCache As PivotCache
    Dim dataRange As Range
    Dim pvtTableDestination As Range

    ' Use the existing "INPUT" sheet
    Set ws = ThisWorkbook.Sheets("INPUT")
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row

    ' Create a new summary sheet, delete if it already exists
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Sheets("Summary").Delete
    On Error GoTo 0
    Application.DisplayAlerts = True

    ' Create a summary sheet
    Set summarySheet = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    summarySheet.Name = "Summary"

    ' Define the data range for the Pivot Table
    Set dataRange = ws.Range("A1:E" & lastRow)

    ' Define Pivot Cache
    Set pCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=dataRange)

    ' Define where to place the Pivot Table
    Set pvtTableDestination = summarySheet.Cells(1, 1)

    ' Create the Pivot Table
    Set pTable = pCache.CreatePivotTable(TableDestination:=pvtTableDestination, TableName:="ExpenseSummary")

    ' Set up the Pivot Table's row fields and data fields
    With pTable
        .PivotFields("Category").Orientation = xlRowField
        .PivotFields("Category").Position = 1
        .PivotFields("Date").Orientation = xlColumnField
        .PivotFields("Date").Position = 1

        ' Group by months and years if your data spans multiple years
        .PivotFields("Date").dataRange.Cells(1).Group _
            Start:=True, End:=True, Periods:=Array(False, False, False, False, True, False, True)

        .AddDataField .PivotFields("Debit"), "Sum of Debit", xlSum
    End With

    ' Clear any error that might have occurred during PivotTable setup
    If Err.Number <> 0 Then
        MsgBox "An error occurred while creating the PivotTable. " & _
               "Please check the field names and data range.", vbCritical
        Err.Clear
    End If

    ' Refresh PivotTable to update fields
    pTable.RefreshTable
    
    ' Sort the PivotTable by the Grand Total for "Sum of Debit" in descending order
    With pTable.PivotFields("Category")
        .AutoSort Order:=xlDescending, Field:="Sum of Debit"
    End With
    
    
    ' Sort the PivotTable by the data field in descending order
    ' With pTable
    '    .RowFields("Category").AutoSort Field:=.DataFields("Sum of Debit").SourceName
    ' End With

    ' Create a Pivot Chart
    Dim chartObj As ChartObject
    Set chartObj = summarySheet.ChartObjects.Add(Left:=100, Width:=600, Top:=50, Height:=300)
    With chartObj.Chart
        .SetSourceData Source:=pTable.TableRange2
        .ChartType = xlColumnClustered
        .HasTitle = True
        .ChartTitle.Text = "Monthly Expenses by Category"
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "Category"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Total Expenses"
        .ShowAllFieldButtons = False
    End With

    ' Deselect the Pivot Chart
    ws.Activate
End Sub

