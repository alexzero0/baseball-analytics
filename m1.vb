
Sub GetMatch()

On Error Resume Next
    Application.ScreenUpdating = False
    timezone = 3
    dayzone = 1
    sourcer = "myscore.ru/"
    suffix = "_ru_1"
    a = 2
    
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", "https://d." & sourcer & "x/feed/f_6_" & dayzone & "_" & timezone & suffix, False
    http.setRequestHeader "X-Fsign", "SW9D1eZo"
    http.Send
    fs_input = http.ResponseText
    fs_rows = Split(fs_input, "~")
    Dim fs_rows_length As Long
    fs_rows_length = UBound(fs_rows) - LBound(fs_rows)
    Call Show_PrBar_Or_No(fs_rows_length, "Загрузка матчей...")
    For i = 0 To fs_rows_length - 4
        If bShowBar Then Call MyProgresBar
        fs_row = Split(fs_rows(i), "¬")
        fs_row_length = UBound(fs_row) - LBound(fs_row)
        fs_index = Split(fs_row(0), ChrW(&HF7))
        If IsArray(fs_index) Then
            fs_index_name = fs_index(0)
            fs_index_value = fs_index(1)
        End If
        If fs_index_name = "SA" Then
            sport_id = fs_index_value
        ElseIf fs_index_name = "ZA" Then
            tour_name = ""
            home_name = ""
            For j = 0 To fs_row_length - 1
                fs_row_parts = Split(fs_row(j), ChrW(&HF7))
                If fs_row_parts(0) = "ZA" Then tour_name = fs_row_parts(1)
                If fs_row_parts(0) = "ZB" Then country_id = fs_row_parts(1)
            Next j
        ElseIf fs_index_name = "AA" Then
            home_name = "": Dim home_scored_first As Integer: Dim away_scored_first As Integer: Dim home_scored_second As Integer: Dim away_scored_second As Integer:
            For j = 0 To fs_row_length - 1
                fs_row_parts = Split(fs_row(j), ChrW(&HF7))
                If fs_row_parts(0) = "AA" Then match_id = fs_row_parts(1)
                If fs_row_parts(0) = "AD" Then
                    date_match = DateAdd("s", fs_row_parts(1), "01/01/1970")
                    date_match = DateAdd("h", timezone, date_match)
                    date_match = Format(date_match, "yyyy.mm.dd hh:mm")
                End If
                If fs_row_parts(0) = "AE" Then home_name = fs_row_parts(1)
                If fs_row_parts(0) = "AF" Then away_name = fs_row_parts(1)
                If fs_row_parts(0) = "AB" Then status_game = fs_row_parts(1)
                If fs_row_parts(0) = "AC" Then status_game_code = fs_row_parts(1)
                If fs_row_parts(0) = "BA" Then home_scored_first = fs_row_parts(1)
                If fs_row_parts(0) = "BB" Then away_scored_first = fs_row_parts(1)
                If fs_row_parts(0) = "BC" Then home_scored_second = fs_row_parts(1)
                If fs_row_parts(0) = "BD" Then away_scored_second = fs_row_parts(1)
                If fs_row_parts(0) = "BE" Then home_scored_second3 = fs_row_parts(1)
                If fs_row_parts(0) = "BF" Then away_scored_second3 = fs_row_parts(1)
                If fs_row_parts(0) = "BG" Then home_scored_second4 = fs_row_parts(1)
                If fs_row_parts(0) = "BH" Then away_scored_second4 = fs_row_parts(1)
                
                Set objRegExp = CreateObject("VBScript.RegExp")
                objRegExp.Pattern = "\s\(...\)"
                home_name = objRegExp.Replace(home_name, "")
                away_name = objRegExp.Replace(away_name, "")
            Next j
        End If
        If tour_name = "" Or home_name = "" Then
        Else
            Sheets("Мсписок").Range("A" & a).Value = tour_name
            Sheets("Мсписок").Range("B" & a).Value = date_match
            Sheets("Мсписок").Range("C" & a).Value = home_name
            Sheets("Мсписок").Range("D" & a).Value = away_name
            Sheets("Мсписок").Range("E" & a).Value = match_id
            Sheets("Мсписок").Range("F" & a).Value = country_id
        
            Sheets("Мсписок").Hyperlinks.Add Anchor:=Sheets("Мсписок").Range("I" & a), Address:="", TextToDisplay:="Загрузить"
            a = a + 1
        End If
    Next i
    If bShowBar Then Unload UserForm1
    Sheets("Мсписок").Range("A2:K" & a).Sort Key1:=Sheets("Мсписок").Columns("B"), Header:=xlYes, Order1:=xlAscending
    Application.ScreenUpdating = True

End Sub

Sub ClearList1()
    Sheets("Мсписок").Activate
    Sheets("Мсписок").Range("A2:I500").Select
    Selection.ClearContents
    Sheets("Мсписок").Range("A2").Select
End Sub
Sub ClearList2()
    Sheets("Дома").Activate
    Sheets("Дома").Range("A3:AC200").Select
    Selection.ClearContents
    Sheets("Дома").Range("A3").Select
End Sub
Sub ClearList4()
    Sheets("Гости").Activate
    Sheets("Гости").Range("A3:AC200").Select
    Selection.ClearContents
    Sheets("гости").Range("A3").Select
End Sub

Sub GetGame(id As String)
    Application.ScreenUpdating = False
    timezone = 3
    dayzone = 0
    sourcer = "myscore.ru/"
    suffix = "_ru_1"
    a = 2
    
    
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", "https://www." & sourcer & "match/" & id & "/", False
    http.Send
    fs_input = http.ResponseText
    Set objRegExp = CreateObject("VBScript.RegExp")
    objRegExp.Pattern = "participantEncodedIds = \[\'(.*)\',\'(.*)\'\]"
    If objRegExp.test(fs_input) = True Then
        Set objMatches = objRegExp.Execute(fs_input)
        ActiveCell.Offset(, -2).Value = objMatches.Item(0).submatches(0)
        ActiveCell.Offset(, -1).Value = objMatches.Item(0).submatches(1)
        home_id = objMatches.Item(0).submatches(0)
        away_id = objMatches.Item(0).submatches(1)
        country_id = ActiveCell.Offset(, -3).Value
    End If
    objRegExp.Pattern = "tournamentStageEncodedId = \'(.*?)\'"
    If objRegExp.test(fs_input) = True Then
        Set objMatches = objRegExp.Execute(fs_input)
        tournament_stage_id = objMatches.Item(0).submatches(0)
    End If
    objRegExp.Pattern = "tournamentEncodedId = \'(.*?)\'"
    If objRegExp.test(fs_input) = True Then
        Set objMatches = objRegExp.Execute(fs_input)
        tournament_id = objMatches.Item(0).submatches(0)
    End If
    
   
    
    a = 3 'Get Home Matches
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", "https://d." & sourcer & "x/feed/pr_1_" & country_id & "_" & home_id & "_0_0_ru_1", False
    http.setRequestHeader "X-Fsign", "SW9D1eZo"
    http.Send
    fs_input = http.ResponseText
    fs_rows = Split(fs_input, "~")
    Dim fs_rows_length As Long
    fs_rows_length = UBound(fs_rows) - LBound(fs_rows)
    Call Show_PrBar_Or_No(fs_rows_length, "Загрузка домашних...")
    For i = 0 To fs_rows_length - 4
        If bShowBar Then Call MyProgresBar
        fs_row = Split(fs_rows(i), "¬")
        fs_row_length = UBound(fs_row) - LBound(fs_row)
        fs_index = Split(fs_row(0), ChrW(&HF7))
        If IsArray(fs_index) Then
            fs_index_name = fs_index(0)
            fs_index_value = fs_index(1)
        End If
        If fs_index_name = "SA" Then
            sport_id = fs_index_value
        ElseIf fs_index_name = "ZA" Then
            tour_name = ""
            home_name = ""
            For j = 0 To fs_row_length - 1
                fs_row_parts = Split(fs_row(j), ChrW(&HF7))
                If fs_row_parts(0) = "ZA" Then tour_name = fs_row_parts(1)
                If fs_row_parts(0) = "ZB" Then country_id = fs_row_parts(1)
            Next j
        ElseIf fs_index_name = "AA" Then
            home_name = ""
            For j = 0 To fs_row_length - 1
                fs_row_parts = Split(fs_row(j), ChrW(&HF7))
                If fs_row_parts(0) = "AA" Then match_id = fs_row_parts(1)
                If fs_row_parts(0) = "AD" Then
                    date_match = DateAdd("s", fs_row_parts(1), "01/01/1970")
                    date_match = DateAdd("h", timezone, date_match)
                    date_match = Format(date_match, "dd.mm.yyyy hh:mm")
                End If
                If fs_row_parts(0) = "AE" Then home_name = fs_row_parts(1)
                If fs_row_parts(0) = "AF" Then away_name = fs_row_parts(1)
                If fs_row_parts(0) = "BA" Then home_1inning = fs_row_parts(1)
                If fs_row_parts(0) = "BB" Then away_1inning = fs_row_parts(1)
                If fs_row_parts(0) = "BC" Then home_2inning = fs_row_parts(1)
                If fs_row_parts(0) = "BD" Then away_2inning = fs_row_parts(1)
                If fs_row_parts(0) = "BE" Then home_3inning = fs_row_parts(1)
                If fs_row_parts(0) = "BF" Then away_3inning = fs_row_parts(1)
                If fs_row_parts(0) = "BG" Then home_4inning = fs_row_parts(1)
                If fs_row_parts(0) = "BH" Then away_4inning = fs_row_parts(1)
                If fs_row_parts(0) = "BI" Then home_5inning = fs_row_parts(1)
                If fs_row_parts(0) = "BJ" Then away_5inning = fs_row_parts(1)
                If fs_row_parts(0) = "BK" Then home_6inning = fs_row_parts(1)
                If fs_row_parts(0) = "BL" Then away_6inning = fs_row_parts(1)
                If fs_row_parts(0) = "BM" Then home_7inning = fs_row_parts(1)
                If fs_row_parts(0) = "BN" Then away_7inning = fs_row_parts(1)
                If fs_row_parts(0) = "BO" Then home_8inning = fs_row_parts(1)
                If fs_row_parts(0) = "BP" Then away_8inning = fs_row_parts(1)
                If fs_row_parts(0) = "BQ" Then home_9inning = fs_row_parts(1)
                If fs_row_parts(0) = "BR" Then away_9inning = fs_row_parts(1)
                              
                If fs_row_parts(0) = "PX" Then home_id2 = fs_row_parts(1)
                If fs_row_parts(0) = "WO" Then home_Pitcher = fs_row_parts(1)
                If fs_row_parts(0) = "WP" Then away_Pitcher = fs_row_parts(1)
                
                If fs_row_parts(0) = "AG" Then home_TotalHit = fs_row_parts(1)
                If fs_row_parts(0) = "WG" Then away_TotalHit = fs_row_parts(1)
            Set objRegExp = CreateObject("VBScript.RegExp")
            objRegExp.Pattern = "\s\(...\)"
            home_name = objRegExp.Replace(home_name, "")
            away_name = objRegExp.Replace(away_name, "")
            Next j
            'If fs_index_value = id Then home_name = ""
        End If
        If tour_name = "" Or home_name = "" Then
        Else
            Sheets("Дома").Range("A" & a).Value = tour_name
            Sheets("Дома").Range("B" & a).Value = date_match
            Sheets("Дома").Range("C" & a).Value = home_name
            Sheets("Дома").Range("D" & a).Value = away_name
            Sheets("Дома").Range("F" & a).Value = home_1inning
            Sheets("Дома").Range("G" & a).Value = home_2inning
            Sheets("Дома").Range("H" & a).Value = home_3inning
            Sheets("Дома").Range("I" & a).Value = home_4inning
            Sheets("Дома").Range("J" & a).Value = home_5inning
            Sheets("Дома").Range("K" & a).Value = home_6inning
            Sheets("Дома").Range("L" & a).Value = home_7inning
            Sheets("Дома").Range("M" & a).Value = home_8inning
            Sheets("Дома").Range("N" & a).Value = home_9inning
            Sheets("Дома").Range("O" & a).Value = away_1inning
            Sheets("Дома").Range("P" & a).Value = away_2inning
            Sheets("Дома").Range("Q" & a).Value = away_3inning
            Sheets("Дома").Range("R" & a).Value = away_4inning
            Sheets("Дома").Range("S" & a).Value = away_5inning
            Sheets("Дома").Range("T" & a).Value = away_6inning
            Sheets("Дома").Range("U" & a).Value = away_7inning
            Sheets("Дома").Range("V" & a).Value = away_8inning
            Sheets("Дома").Range("W" & a).Value = away_9inning
            Sheets("Дома").Range("Z" & a).Value = home_TotalHit
            Sheets("Дома").Range("AA" & a).Value = away_TotalHit
            Sheets("Дома").Range("AB" & a).Value = home_Pitcher
            Sheets("Дома").Range("AC" & a).Value = away_Pitcher
            
            If home_id2 = home_id Then
                Sheets("Дома").Range("E" & a).Value = "Дома"
            Else
                Sheets("Дома").Range("E" & a).Value = "Гости"
            End If
            a = a + 1
        End If
    Next i
    If bShowBar Then Unload UserForm1
    
    SetFormulaD (a), "Дома"
    SetFormulaG (a), "Дома"
    
    a = a + 5
    Set20Game (a), "Дома" 'подумать над выводом в лист
    
    a = 3 'Get Away Macthes
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", "https://d." & sourcer & "x/feed/pr_1_" & country_id & "_" & away_id & "_0_0_ru_1", False
    http.setRequestHeader "X-Fsign", "SW9D1eZo"
    http.Send
    fs_input = http.ResponseText
    fs_rows = Split(fs_input, "~")
    fs_rows_length = UBound(fs_rows) - LBound(fs_rows)
    Call Show_PrBar_Or_No(fs_rows_length, "Загрузка гостевых...")
    For i = 0 To fs_rows_length - 4
        If bShowBar Then Call MyProgresBar
        fs_row = Split(fs_rows(i), "¬")
        fs_row_length = UBound(fs_row) - LBound(fs_row)
        fs_index = Split(fs_row(0), ChrW(&HF7))
        If IsArray(fs_index) Then
            fs_index_name = fs_index(0)
            fs_index_value = fs_index(1)
        End If
        If fs_index_name = "SA" Then
            sport_id = fs_index_value
        ElseIf fs_index_name = "ZA" Then
            tour_name = ""
            home_name = ""
            For j = 0 To fs_row_length - 1
                fs_row_parts = Split(fs_row(j), ChrW(&HF7))
                If fs_row_parts(0) = "ZA" Then tour_name = fs_row_parts(1)
                If fs_row_parts(0) = "ZB" Then country_id = fs_row_parts(1)
            Next j
        ElseIf fs_index_name = "AA" Then
            home_name = ""
            For j = 0 To fs_row_length - 1
                fs_row_parts = Split(fs_row(j), ChrW(&HF7))
                If fs_row_parts(0) = "AA" Then match_id = fs_row_parts(1)
                If fs_row_parts(0) = "AD" Then
                    date_match = DateAdd("s", fs_row_parts(1), "01/01/1970")
                    date_match = DateAdd("h", timezone, date_match)
                    date_match = Format(date_match, "yyyy.mm.dd hh:mm")
                End If
                If fs_row_parts(0) = "AE" Then home_name = fs_row_parts(1)
                If fs_row_parts(0) = "AF" Then away_name = fs_row_parts(1)
                If fs_row_parts(0) = "BA" Then home_1inning = fs_row_parts(1)
                If fs_row_parts(0) = "BB" Then away_1inning = fs_row_parts(1)
                If fs_row_parts(0) = "BC" Then home_2inning = fs_row_parts(1)
                If fs_row_parts(0) = "BD" Then away_2inning = fs_row_parts(1)
                If fs_row_parts(0) = "BE" Then home_3inning = fs_row_parts(1)
                If fs_row_parts(0) = "BF" Then away_3inning = fs_row_parts(1)
                If fs_row_parts(0) = "BG" Then home_4inning = fs_row_parts(1)
                If fs_row_parts(0) = "BH" Then away_4inning = fs_row_parts(1)
                If fs_row_parts(0) = "BI" Then home_5inning = fs_row_parts(1)
                If fs_row_parts(0) = "BJ" Then away_5inning = fs_row_parts(1)
                If fs_row_parts(0) = "BK" Then home_6inning = fs_row_parts(1)
                If fs_row_parts(0) = "BL" Then away_6inning = fs_row_parts(1)
                If fs_row_parts(0) = "BM" Then home_7inning = fs_row_parts(1)
                If fs_row_parts(0) = "BN" Then away_7inning = fs_row_parts(1)
                If fs_row_parts(0) = "BO" Then home_8inning = fs_row_parts(1)
                If fs_row_parts(0) = "BP" Then away_8inning = fs_row_parts(1)
                If fs_row_parts(0) = "BQ" Then home_9inning = fs_row_parts(1)
                If fs_row_parts(0) = "BR" Then away_9inning = fs_row_parts(1)
                              
                If fs_row_parts(0) = "PX" Then away_id2 = fs_row_parts(1)
                If fs_row_parts(0) = "WO" Then home_Pitcher = fs_row_parts(1)
                If fs_row_parts(0) = "WP" Then away_Pitcher = fs_row_parts(1)
                
                If fs_row_parts(0) = "AG" Then home_TotalHit = fs_row_parts(1)
                If fs_row_parts(0) = "WG" Then away_TotalHit = fs_row_parts(1)
            Set objRegExp = CreateObject("VBScript.RegExp")
            objRegExp.Pattern = "\s\(...\)"
            home_name = objRegExp.Replace(home_name, "")
            away_name = objRegExp.Replace(away_name, "")
            Next j
            'If fs_index_value = id Then home_name = ""
            'MsgBox
        End If
        If tour_name = "" Or home_name = "" Or away_id2 = "" Then
        Else
            Sheets("Гости").Range("A" & a).Value = tour_name
            Sheets("Гости").Range("B" & a).Value = date_match
            Sheets("Гости").Range("C" & a).Value = home_name
            Sheets("Гости").Range("D" & a).Value = away_name
            Sheets("Гости").Range("F" & a).Value = home_1inning
            Sheets("Гости").Range("G" & a).Value = home_2inning
            Sheets("Гости").Range("H" & a).Value = home_3inning
            Sheets("Гости").Range("I" & a).Value = home_4inning
            Sheets("Гости").Range("J" & a).Value = home_5inning
            Sheets("Гости").Range("K" & a).Value = home_6inning
            Sheets("Гости").Range("L" & a).Value = home_7inning
            Sheets("Гости").Range("M" & a).Value = home_8inning
            Sheets("Гости").Range("N" & a).Value = home_9inning
            Sheets("Гости").Range("O" & a).Value = away_1inning
            Sheets("Гости").Range("P" & a).Value = away_2inning
            Sheets("Гости").Range("Q" & a).Value = away_3inning
            Sheets("Гости").Range("R" & a).Value = away_4inning
            Sheets("Гости").Range("S" & a).Value = away_5inning
            Sheets("Гости").Range("T" & a).Value = away_6inning
            Sheets("Гости").Range("U" & a).Value = away_7inning
            Sheets("Гости").Range("V" & a).Value = away_8inning
            Sheets("Гости").Range("W" & a).Value = away_9inning
            Sheets("Гости").Range("Z" & a).Value = home_TotalHit
            Sheets("Гости").Range("AA" & a).Value = away_TotalHit
            Sheets("Гости").Range("AB" & a).Value = home_Pitcher
            Sheets("Гости").Range("AC" & a).Value = away_Pitcher
            
            If away_id2 = away_id Then
                Sheets("Гости").Range("E" & a).Value = "Дома"
            Else
                Sheets("Гости").Range("E" & a).Value = "Гости"
            End If
            a = a + 1
        End If
        Next i
    If bShowBar Then Unload UserForm1
    
    Sheets("Гости").Range("A2:AC" & a).Sort Key1:=Sheets("Гости").Columns("B"), Header:=xlYes, Order1:=xlDescending
    
    
    SetFormulaD (a), "Гости"
    SetFormulaG (a), "Гости"
    
    a = a + 5
    Set20Game (a), "Гости" 'подумать над выводом в лист
    
    SetGraf
    Application.ScreenUpdating = True
        
End Sub
