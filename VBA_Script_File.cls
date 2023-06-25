VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Stock_Data()
'declare
    Dim dataSheet As Worksheet
    Dim outputSheet As Worksheet
    Dim lastRow As Long
    Dim dataRange As Range
    Dim tickerCell As Range
    Dim i As Long
    Dim maxPercentIncrease As Double
    Dim maxPercentDecrease As Double
    Dim maxTotalVolume As Double
    Dim tickerMaxPercentIncrease As String
    Dim tickerMaxPercentDecrease As String
    Dim tickerMaxTotalVolume As String
    Dim yearSheetNames As Variant
    
    
    yearSheetNames = Array("2018", "2019", "2020") '
    
    
    
    'Loop Through each sheet
    For Each yearSheetName In yearSheetNames
        Set dataSheet = Sheets(yearSheetName)
        Set outputSheet = Sheets(yearSheetName)
    
        ' Find last row
        lastRow = dataSheet.Cells(dataSheet.Rows.Count, "A").End(xlUp).Row
        Set dataRange = dataSheet.Range("A2:G" & lastRow)
    
        dataSheet.Range("I2:L" & lastRow).ClearContents
        dataSheet.Range("P2:Q4").ClearContents
    
        ' loop through data
        i = 2
    
        For Each tickerCell In dataRange.Columns(1).Cells
            If tickerCell.Value <> "" And (tickerCell.Value <> tickerCell.Offset(-1, 0).Value Or tickerCell.Row = 2) Then
            
                Dim firstOpen As Double
                Dim lastClose As Double
                
                
                firstOpen = tickerCell.Offset(0, 2).Value
                lastClose = tickerCell.Offset(WorksheetFunction.CountIf(dataRange.Columns(1), tickerCell.Value) - 1, 5).Value
                
                
                Dim yearlyChange As Double
                Dim percentChange As Double
                
                
                yearlyChange = lastClose - firstOpen
                'made conditional formatting in excel because I couldnt figure out how to make it a percent answer
                percentChange = (yearlyChange / firstOpen) * 100
                
                
                Dim totalVolume As Double
                
                totalVolume = Application.WorksheetFunction.SumIf(dataRange.Columns(1), tickerCell.Value, dataRange.Columns(7))
                
                
        
                dataSheet.Cells(i, "I").Value = tickerCell.Value
                dataSheet.Cells(i, "J").Value = yearlyChange
                dataSheet.Cells(i, "K").Value = percentChange
                dataSheet.Cells(i, "L").Value = totalVolume
            
    
                If percentChange > maxPercentIncrease Then
                    maxPercentIncrease = percentChange
                    tickerMaxPercentIncrease = tickerCell.Value
                    
                ElseIf percentChange < maxPercentDecrease Then
                    maxPercentDecrease = percentChange
                    tickerMaxPercentDecrease = tickerCell.Value
                    
                    
                End If
            
                
                If totalVolume > maxTotalVolume Then
                    maxTotalVolume = totalVolume
                    tickerMaxTotalVolume = tickerCell.Value
                    
                    
                End If
            
                'increment output/summary row
                i = i + 1
            End If
            
            
            
            
        Next tickerCell
    
        ' final summary section/output area
        outputSheet.Range("P2").Value = tickerMaxPercentIncrease
        outputSheet.Range("P3").Value = tickerMaxPercentDecrease
        outputSheet.Range("P4").Value = tickerMaxTotalVolume
        outputSheet.Range("Q2").Value = maxPercentIncrease
        outputSheet.Range("Q3").Value = maxPercentDecrease
        outputSheet.Range("Q4").Value = maxTotalVolume
    
        ' Reset for following years
        maxPercentIncrease = 0
        maxPercentDecrease = 0
        maxTotalVolume = 0
        tickerMaxPercentIncrease = ""
        tickerMaxPercentDecrease = ""
        tickerMaxTotalVolume = ""
        
        
    Next yearSheetName
    
    
    
End Sub
