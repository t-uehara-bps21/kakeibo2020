Sub ランキング()
    Dim arva(100) As Long, arym(100) As Long
    Dim i As Long, j As Long
    Dim vmax As Long, val As Long, ym As Long, ym2 As Long
    Dim st As Worksheet, ast As Worksheet
    Dim senntaku As String
    
    Set st = Sheets("月別合計")
    Set ast = Sheets("データ")
    
    For i = 1 To Rows().Count
        
        If st.Cells(3 + i, "D") = "" Then
            Exit For
            
        End If
        
        arym(i) = st.Cells(3 + i, "C") 
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
        vmax = arva(i)
        ym2 = arym(i)
        
        For j = 1 To 20
            val = arva(j)
            ym = arym(j)
        
        If vmax > val Then
            arva(j) = vmax
            arva(i) = val
            vmax = val
            
            arym(j) = ym2
            arym(i) = ym
            ym2 = ym
            
        End If
        Next j
    Next i
               
    ast.Cells(16, "F") = arva(1)
    ast.Cells(17, "F") = arva(2)
    ast.Cells(18, "F") = arva(3)
    ast.Cells(16, "E") = arym(1)
    ast.Cells(17, "E") = arym(2)
    ast.Cells(18, "E") = arym(3)
 
End Sub
