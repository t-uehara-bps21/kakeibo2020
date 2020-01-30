Private Type Hairetsu
    money As Long
    ym As Long
End Type

Sub ランキング()

    Dim arr() As Hairetsu
    Dim val As Hairetsu
    Dim i As Long
    Dim j As Long
    Dim tsukist As Worksheet
    Dim datast As Worksheet
    Dim senntaku As String
    
    Set tsukist = Sheets("月別合計")
    Set datast = Sheets("データ")
    
    For i = 1 To Rows().Count
    
    ReDim Preserve arr(i)
        If tsukist.Cells(3 + i, "D") = "" Then
            Exit For  
        End If
        
        arr(i).ym = tsukist.Cells(3 + i, "C")
        senntaku = datast.Cells(16, "B")
         
        Select Case senntaku
            Case "収支"
                arr(i).money = tsukist.Cells(3 + i, "D")   
            Case "収入"
                arr(i).money = tsukist.Cells(3 + i, "E")
            Case "支出"
                arr(i).money = tsukist.Cells(3 + i, "F")
            Case "貯蓄"
                arr(i).money = tsukist.Cells(3 + i, "G")
            End Select
    Next i
  
    For i = 1 To UBound(arr) - 1
        For j = 1 To UBound(arr) - 1
        
            If arr(i).money > arr(j).money Then
                val = arr(i)
                arr(i) = arr(j)
                arr(j) = val
            End If
        Next j
    Next i
  
    datast.Cells(16, "F") = arr(1).money
    datast.Cells(17, "F") = arr(2).money
    datast.Cells(18, "F") = arr(3).money
    datast.Cells(16, "E") = arr(1).ym
    datast.Cells(17, "E") = arr(2).ym
    datast.Cells(18, "E") = arr(3).ym
  
End Sub