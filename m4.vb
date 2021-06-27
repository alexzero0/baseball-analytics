Sub SetFormulaD(a As Integer, str As String)
'
' SetFormula Ìàêðîñ
'

'
    Sheets(str).Activate
    Sheets(str).Range("X3").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-19]="""","""",IF(RC[-10]=""X"",RC[-18]+RC[-17]+RC[-16]+RC[-15]+RC[-14]+RC[-13]+RC[-12]+RC[-11],RC[-18]+RC[-17]+RC[-16]+RC[-15]+RC[-14]+RC[-13]+RC[-12]+RC[-11]+RC[-10]))"
    Selection.AutoFill Destination:=Range("X3", "X" & a), Type:=xlFillDefault
    
End Sub
Sub SetFormulaG(a As Integer, str As String)
'
' SetFormulaG Ìàêðîñ
'

'
    Sheets(str).Activate
    Range("Y3").Select
    ActiveCell.FormulaR1C1 = _
        "=IF(RC[-20]="""","""",IF(RC[-2]=""X"",RC[-10]+RC[-9]+RC[-8]+RC[-7]+RC[-6]+RC[-5]+RC[-4]+RC[-3],RC[-10]+RC[-9]+RC[-8]+RC[-7]+RC[-6]+RC[-5]+RC[-4]+RC[-3]+RC[-2]))"
    Selection.AutoFill Destination:=Range("Y3", "Y" & a), Type:=xlFillDefault

End Sub
Sub Set20Game(a As Integer, str As String) 'a As Integer
    
    Sheets(str).Activate
    Dim mas As Variant
    mas = Range("A3:AC23").Value
    j = 1
    For i = 20 To 1 Step -1
        Sheets(str).Range("A" & a) = j
        If mas(i, 5) = "Äîìà" Then
            Sheets(str).Range("C" & a) = mas(i, 3) 'êîìàíäà ñâîÿ
            Sheets(str).Range("B" & a) = mas(i, 2) 'Äàòà
            Sheets(str).Range("E" & a) = mas(i, 5) 'Äîìà-Ãîñòè
            Sheets(str).Range("D" & a) = mas(i, 4) 'êîìàíäà ïðîòèâíèêîâ!
            Sheets(str).Range("X" & a) = mas(i, 24) 'ñ÷åò êîìàíäû(Ä)
            Sheets(str).Range("Y" & a) = mas(i, 25) 'ñ÷åò êîìàíäà ïðîòèâíèêîâ!
            Sheets(str).Range("Z" & a) = mas(i, 26) 'òîòàë õèòîâ êîìàíäû(Ä)
            Sheets(str).Range("AA" & a) = mas(i, 27) 'òîòàë õèòîâ êîìàíäà ïðîòèâíèêîâ!
        ElseIf mas(i, 5) = "Ãîñòè" Then
            Sheets(str).Range("C" & a) = mas(i, 4) 'êîìàíäà ñâîÿ
            Sheets(str).Range("B" & a) = mas(i, 2) 'Äàòà
            Sheets(str).Range("E" & a) = mas(i, 5) 'Äîìà-Ãîñòè
            Sheets(str).Range("D" & a) = mas(i, 3) 'êîìàíäà ïðîòèâíèêîâ!
            Sheets(str).Range("X" & a) = mas(i, 25) 'ñ÷åò êîìàíäû
            Sheets(str).Range("Y" & a) = mas(i, 24) 'ñ÷åò êîìàíäà ïðîòèâíèêîâ!
            Sheets(str).Range("Z" & a) = mas(i, 27) 'òîòàë õèòîâ êîìàíäû
            Sheets(str).Range("AA" & a) = mas(i, 26) 'òîòàë õèòîâ êîìàíäà ïðîòèâíèêîâ!
        End If
        a = a + 1
        j = j + 1
    Next
    Sheets(str).Range("A" & a) = j
    Sheets(str).Range("X" & a).FormulaR1C1 = _
        "=FORECAST.LINEAR(RC[-23],R[-20]C:R[-1]C,R[-20]C[-23]:R[-1]C[-23])"
    Dim str2 As String
    If str = "Äîìà" Then
       str2 = """" & str & "!" & "X" & a & """"
       Formula = "INDIRECT(" & str2 & ",TRUE)"
       Sheets("Âûâîä").Range("E3").Formula = "=" & Formula
    ElseIf str = "Ãîñòè" Then
        str2 = """" & str & "!" & "X" & a & """"
        Formula = "INDIRECT(" & str2 & ",TRUE)"
        Sheets("Âûâîä").Range("F3").Formula = "=" & Formula
    End If
End Sub
Sub SetFormulaChet()
'
' SetFormulaChet Ìàêðîñ
'

'
    ActiveCell.FormulaR1C1 = _
        "=FORECAST.LINEAR(RC[-23],R[-20]C:R[-1]C,R[-20]C[-23]:R[-1]C[-23])"
    Range("X69").Select
End Sub


