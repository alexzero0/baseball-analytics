Sub SetPitcher()
'
' SetPitcher Ìàêðîñ
'
' Ñî÷åòàíèå êëàâèø: Ctrl+g
'
    
    If ActiveSheet.Name = "Ñòàòà" Then
        a = ActiveCell.Row
        b = ActiveCell.Column
        If b = 3 Then
            mas = Range("C" & a, "AI" & a).Value
            Play = Sheets("Âûâîä").Range("AI9")
            PRow = Sheets("Âûâîä").Range("AI9").Row
            j = 3
            Do
                PlayWAR = Sheets("WAR").Range("A" & j)
                If PlayWAR = mas(1, 1) Then WAR = Sheets("WAR").Range("B" & j): Exit Do
                j = j + 1
            Loop While PlayWAR <> "Õèòòåð"
            j = 3
            Do
                PlayWAR = Sheets("WAR").Range("D" & j)
                If PlayWAR = mas(1, 1) Then WAR = Sheets("WAR").Range("E" & j): Exit Do
                j = j + 1
            Loop While PlayWAR <> "Õèòòåð"
            If Play = "Player" Then
                For i = PRow To 15
                    If Sheets("Âûâîä").Range("AI" & i) = "" Then
                        Sheets("Âûâîä").Range("AI" & i, "BO" & i) = mas
                        Sheets("Âûâîä").Range("BQ" & i) = WAR
                        Exit For
                    End If
                Next
            End If
        End If
    End If
End Sub
Sub SetXitter()
'
' SetXitter Ìàêðîñ
'
' Ñî÷åòàíèå êëàâèø: Ctrl+h
'
    If ActiveSheet.Name = "Ñòàòà" Then
        a = ActiveCell.Row
        b = ActiveCell.Column
        If b = 3 Then
            mas = Range("C" & a, "AF" & a).Value
            Play = Sheets("Âûâîä").Range("AI17")
            PRow = Sheets("Âûâîä").Range("AI17").Row
            j = 3
            Do
                PlayWAR = Sheets("WAR").Range("A" & j)
                If PlayWAR = "Õèòòåð" Then Exit Do
                j = j + 1
            Loop While True
            Do
                PlayWAR = Sheets("WAR").Range("A" & j)
                If PlayWAR = mas(1, 1) Then WAR = Sheets("WAR").Range("B" & j): GoTo myExit1
                j = j + 1
            Loop While PlayWAR <> ""
            j = 3
            Do
                PlayWAR = Sheets("WAR").Range("D" & j)
                If PlayWAR = "Õèòòåð" Then Exit Do
                j = j + 1
            Loop While True
            Do
                PlayWAR = Sheets("WAR").Range("D" & j)
                If PlayWAR = mas(1, 1) Then WAR = Sheets("WAR").Range("E" & j): Exit Do
                j = j + 1
            Loop While PlayWAR <> ""
            
            
myExit1:
            If Play = "Player" Then
                For i = PRow To 35
                    If Sheets("Âûâîä").Range("AI" & i) = "" Then
                        Sheets("Âûâîä").Range("AI" & i, "BL" & i) = mas
                        Sheets("Âûâîä").Range("BM" & i) = WAR
                        Exit For
                    End If
                Next
            End If
        End If
        
    End If
End Sub

Sub ClearList5()
    Sheets("Âûâîä").Activate
    Sheets("Âûâîä").Range("AI10:BO15").Select
    Selection.ClearContents
    Sheets("Âûâîä").Range("BQ10:BQ15").Select
    Selection.ClearContents
    Sheets("Âûâîä").Range("AI18:BL35").Select
    Selection.ClearContents
    Sheets("Âûâîä").Range("BM18:BM35").Select
    Selection.ClearContents
    Sheets("Âûâîä").Range("C4").Select
End Sub
Sub ClearAllList()
    Application.ScreenUpdating = False
    ClearList1
    ClearList2
    ClearList3
    ClearList4
    ClearList5
    Sheets("Ìñïèñîê").Activate
    Application.ScreenUpdating = True
End Sub
