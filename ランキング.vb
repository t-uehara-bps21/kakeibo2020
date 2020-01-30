Private Type Hairetsu
    money As Long
    ym As Long
End Type

Sub ランキング()
    Dim arr(100) As Hairetsu
    Dim i As Long
    Dim j As Long
    Dim vmax As Long
    Dim val As Long
    Dim ym As Long
    Dim ym2 As Long
    Dim tsukist As Worksheet
    Dim datast As Worksheet
    Dim senntaku As String
    
    Set tsukist = Sheets("月別合計")
    Set datast = Sheets("データ")
    
    For i = 1 To Rows().Count
        
        If tsukist.Cells(3 + i, "D") = "" Then
            Exit For
            
        End If
        
        arr(i) = tsukist.Cells(3 + i, "C") 
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
    
    For i = 1 To 20
        vmax = arr(i)
        ym2 = arr(i)
        
        For j = 1 To 20
            val = arr(j)
            ym = arr(j)
        
        If vmax > val Then
            arr(j) = vmax
            arr(i) = val
            vmax = val
            
            arr(j) = ym2
            arr(i) = ym
            ym2 = ym
            
        End If
        Next j
    Next i
               
    datast.Cells(16, "F") = arr(1)
    datast.Cells(17, "F") = arr(2)
    datast.Cells(18, "F") = arr(3)
    datast.Cells(16, "E") = arr(1)
    datast.Cells(17, "E") = arr(2)
    datast.Cells(18, "E") = arr(3)
 
End Sub
