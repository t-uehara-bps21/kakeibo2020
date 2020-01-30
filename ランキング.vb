Private Type Arr
    TotalValue As Long
    Month As Long
End Type

Sub 月別ランキング()

    Dim RankArr() As Arr
    Dim Val As Arr
    Dim i As Long
    Dim j As Long
    Dim MonthTotalSt As Worksheet
    Dim DataSt As Worksheet
    Dim Senntaku As String
    
    Set MonthTotalSt = Sheets("月別合計")
    Set DataSt = Sheets("データ")
    
    For i = 1 To Rows().Count
        ReDim Preserve RankArr(i)
        
            If MonthTotalSt.Cells(3 + i, "D") = "" Then
                Exit For
            End If
        
            RankArr(i).Month = MonthTotalSt.Cells(3 + i, "C")
            Senntaku = DataSt.Cells(16, "B")
         
            Select Case Senntaku
                Case "収支"
                    RankArr(i).TotalValue = MonthTotalSt.Cells(3 + i, "D")
                Case "収入"
                    RankArr(i).TotalValue = MonthTotalSt.Cells(3 + i, "E")
                Case "支出"
                    RankArr(i).TotalValue = MonthTotalSt.Cells(3 + i, "F")
                Case "貯蓄"
                    RankArr(i).TotalValue = MonthTotalSt.Cells(3 + i, "G")
            End Select
    Next i
  
    For i = 1 To UBound(RankArr) - 1
        For j = 1 To UBound(RankArr) - 1
        
            If RankArr(i).TotalValue > RankArr(j).TotalValue Then
                Val = RankArr(i)
                RankArr(i) = RankArr(j)
                RankArr(j) = Val
            End If
        Next j
    Next i
  
    DataSt.Cells(16, "F") = RankArr(1).TotalValue
    DataSt.Cells(17, "F") = RankArr(2).TotalValue
    DataSt.Cells(18, "F") = RankArr(3).TotalValue
    DataSt.Cells(16, "E") = RankArr(1).Month
    DataSt.Cells(17, "E") = RankArr(2).Month
    DataSt.Cells(18, "E") = RankArr(3).Month
  
End Sub