Attribute VB_Name = "Module1"
Sub Creditcard()
  Dim ticker As String
  Dim lastrow As Long
  Dim OpenPrice As Double
  Dim Closeprice As Double
  Dim percentage As Variant
  Dim j As Integer
  Dim YearlyChange As Double
  Dim TotalVolume As Variant
  'Initialize the variables
   TotalVolume = 0
   j = 2
   'Find the lastrow
   lastrow = Cells(Rows.Count, 1).End(xlUp).Row
   'Loop through the column A
   For i = 2 To lastrow
    If (Cells(i + 1, 1).Value <> Cells(i, 1).Value) Then
        ticker = Cells(i, "A").Value
        Closeprice = Cells(i, "F").Value
        TotalVolume = TotalVolume + Cells(i, 7).Value
        'calculate yearly change and percentage
        YearlyChange = Closeprice - OpenPrice
        If (YearlyChange < 0) Then
            Range("J" & j).Interior.ColorIndex = 3
        Else
            Range("J" & j).Interior.ColorIndex = 4
        End If
        If (OpenPrice <> 0) Then
            percentage = (1 - (OpenPrice / Closeprice)) * 100
        Else
           percentage = 0
        End If
        percentage = Format(percentage, "0.00%")
        'Print on column I , J , K
        Range("I" & 1).Value = "Ticker"
        Range("I" & j).Value = ticker
        Range("J" & 1).Value = "Yearly Change"
        Range("J" & j).Value = YearlyChange
        Range("K" & 1).Value = "Percentage Change"
        Range("K" & j).Value = percentage
        Range("L" & 1).Value = "Total Volume"
        Range("L" & j).Value = TotalVolume
        'inititalize for next ticker
        TotalVolume = 0
        j = j + 1
    Else
         OpenPrice = Cells(i, "C").Value
         TotalVolume = TotalVolume + Cells(i, 7).Value
    End If
   Next i
   
End Sub
