Sub GetWAR()
    home = Sheets("Ñòàòà").Range("A32").Value
    If home = 1 Then
        home = "ari" 'Arizona Diamondbacks
    ElseIf home = 2 Then
        home = "atl" 'Atlanta Braves
    ElseIf home = 3 Then
        home = "bal" 'Baltimore Orioles
    ElseIf home = 4 Then
        home = "bos" 'Boston Red Sox
    ElseIf home = 5 Then
        home = "chc" 'Chicago Cubs
    ElseIf home = 6 Then
        home = "chw" 'Chicago White Sox
    ElseIf home = 7 Then
        home = "cin" 'Cincinnati Reds
    ElseIf home = 8 Then
        home = "cle" 'Cleveland Indians
    ElseIf home = 9 Then
        home = "col" 'Colorado Rockies
    ElseIf home = 10 Then
        home = "det" 'Detroit Tigers
    ElseIf home = 11 Then
        home = "hou" 'Houston Astros
    ElseIf home = 12 Then
        home = "kan" 'Kansas City Royals
    ElseIf home = 13 Then
        home = "laa" 'Los Angeles Angels
    ElseIf home = 14 Then
        home = "lad" 'Los Angeles Dodgers
    ElseIf home = 15 Then
        home = "mia" 'Miami Marlins
    ElseIf home = 16 Then
        home = "mil" 'Milwaukee Brewers
    ElseIf home = 17 Then
        home = "min" 'Minnesota Twins
    ElseIf home = 18 Then
        home = "nym" 'New York Mets
    ElseIf home = 19 Then
        home = "nyy" 'New York Yankees
    ElseIf home = 20 Then
        home = "oak" 'Oakland Athletics
    ElseIf home = 21 Then
        home = "phi" 'Philadelphia Phillies
    ElseIf home = 22 Then
        home = "pit" 'Pittsburgh Pirates
    ElseIf home = 23 Then
        home = "sd" 'San Diego Padres
    ElseIf home = 24 Then
        home = "sf" 'San Francisco Giants
    ElseIf home = 25 Then
        home = "sea" 'Seattle Mariners
    ElseIf home = 26 Then
        home = "stl" 'St. Louis Cardinals
    ElseIf home = 27 Then
        home = "tam" 'Tampa Bay Rays
    ElseIf home = 28 Then
        home = "tex" 'Texas Rangers
    ElseIf home = 29 Then
        home = "tor" 'Toronto Blue Jays
    ElseIf home = 30 Then
        home = "was" 'Washington Nationals
    ElseIf home = False Then
        Exit Sub
    End If
       
    away = Sheets("Ñòàòà").Range("B32").Value
    
    If away = 1 Then
        away = "ari" 'Arizona Diamondbacks
    ElseIf away = 2 Then
        away = "atl" 'Atlanta Braves
    ElseIf away = 3 Then
        away = "bal" 'Baltimore Orioles
    ElseIf away = 4 Then
        away = "bos" 'Boston Red Sox
    ElseIf away = 5 Then
        away = "chc" 'Chicago Cubs
    ElseIf away = 6 Then
        away = "chw" 'Chicago White Sox
    ElseIf away = 7 Then
        away = "cin" 'Cincinnati Reds
    ElseIf away = 8 Then
        away = "cle" 'Cleveland Indians
    ElseIf away = 9 Then
        away = "col" 'Colorado Rockies
    ElseIf away = 10 Then
        away = "det" 'Detroit Tigers
    ElseIf away = 11 Then
        away = "hou" 'Houston Astros
    ElseIf away = 12 Then
        away = "kan" 'Kansas City Royals
    ElseIf away = 13 Then
        away = "laa" 'Los Angeles Angels
    ElseIf away = 14 Then
        away = "lad" 'Los Angeles Dodgers
    ElseIf away = 15 Then
        away = "mia" 'Miami Marlins
    ElseIf away = 16 Then
        away = "mil" 'Milwaukee Brewers
    ElseIf away = 17 Then
        away = "min" 'Minnesota Twins
    ElseIf away = 18 Then
        away = "nym" 'New York Mets
    ElseIf away = 19 Then
        away = "nyy" 'New York Yankees
    ElseIf away = 20 Then
        away = "oak" 'Oakland Athletics
    ElseIf away = 21 Then
        away = "phi" 'Philadelphia Phillies
    ElseIf away = 22 Then
        away = "pit" 'Pittsburgh Pirates
    ElseIf away = 23 Then
        away = "sd" 'San Diego Padres
    ElseIf away = 24 Then
        away = "sf" 'San Francisco Giants
    ElseIf away = 25 Then
        away = "sea" 'Seattle Mariners
    ElseIf away = 26 Then
        away = "stl" 'St. Louis Cardinals
    ElseIf away = 27 Then
        away = "tam" 'Tampa Bay Rays
    ElseIf away = 28 Then
        away = "tex" 'Texas Rangers
    ElseIf away = 29 Then
        away = "tor" 'Toronto Blue Jays
    ElseIf away = 30 Then
        away = "was" 'Washington Nationals
    ElseIf away = False Then
        Exit Sub
    End If
    
    

    Set http = CreateObject("WinHttp.WinHttpRequest.5.1") 'WinHttp.WinHttpRequest.5.1   MSXML2.XMLHTTP
    http.Open "GET", "https://www.espn.com/mlb/team/stats/_/type/pitching/name/" & home
    http.setRequestHeader "Upgrade-Insecure-Requests", "1"
    http.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/75.0.3770.142 Safari/537.36 OPR/62.0.3331.116"
    http.setRequestHeader "DNT", "1"
    http.setRequestHeader "Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8"
    http.Send
    http.waitForResponse = True
    fs_input = http.ResponseText

    fs_rows = Split(fs_input, "athlete")
    a = 3
    Dim fs_rows_length As Long
    fs_rows_length = UBound(fs_rows) - LBound(fs_rows)
    Call Show_PrBar_Or_No(fs_rows_length, "Çàãðóçêà ïèò÷åðîâ Äîìà(WAR)...")
    For i = 2 To fs_rows_length
        If bShowBar Then Call MyProgresBar
        fs_row = Split(fs_rows(i), ",")
        fs_row_length = UBound(fs_row) - LBound(fs_row)
        For j = 0 To fs_row_length

            fs_nrow = Replace(fs_row(j), """", "")
            fs_nrow = Replace(fs_nrow, "{", "")
            If j = 0 Then fs_nrow = Replace(fs_nrow, ":", "", 1, 1)
            fs_nrow2 = Split(fs_nrow, ":")
            If j = 0 Then Name = fs_nrow2(1): j = 85
                If a = 3 Then
                    ProvName = Name
                ElseIf ProvName = Name Then GoTo MyExit
                End If
            If j = 88 Then
                WARvalue = fs_nrow2(1)
                Exit For
            End If
            
        Next
        
        Sheets("WAR").Range("A" & a).Value = Name
        Sheets("WAR").Range("B" & a).Value = WARvalue
        a = a + 1
    Next
MyExit:
    
    
    Sheets("WAR").Range("A" & a).Value = "Õèòòåð"
    If bShowBar Then Unload UserForm1
    a = a + 1
    b = a
    ' Çàãðóçêà õèòòåðîâ äîìà
    
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1") 'WinHttp.WinHttpRequest.5.1   MSXML2.XMLHTTP
    http.Open "GET", "https://www.espn.com/mlb/team/stats/_/name/" & home
    http.setRequestHeader "Upgrade-Insecure-Requests", "1"
    http.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/75.0.3770.142 Safari/537.36 OPR/62.0.3331.116"
    http.setRequestHeader "DNT", "1"
    http.setRequestHeader "Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8"
    http.Send
    http.waitForResponse = True
    fs_input = http.ResponseText
    
    fs_rows = Split(fs_input, "athlete")

    
    fs_rows_length = UBound(fs_rows) - LBound(fs_rows)
    Call Show_PrBar_Or_No(fs_rows_length, "Çàãðóçêà õèòòåðîâ Äîìà(WAR)...")
    For i = 2 To fs_rows_length
        If bShowBar Then Call MyProgresBar
        fs_row = Split(fs_rows(i), ",")
        fs_row_length = UBound(fs_row) - LBound(fs_row)
        For j = 0 To fs_row_length

            fs_nrow = Replace(fs_row(j), """", "")
            fs_nrow = Replace(fs_nrow, "{", "")
            If j = 0 Then fs_nrow = Replace(fs_nrow, ":", "", 1, 1)
            fs_nrow2 = Split(fs_nrow, ":")
            If j = 0 Then Name = fs_nrow2(1): j = 92
                If a = b Then
                    ProvName = Name
                ElseIf ProvName = Name Then GoTo MyExit2
                End If
            If j = 93 Then
                WARvalue = fs_nrow2(1)
                Exit For
            End If
            
        Next
        
        Sheets("WAR").Range("A" & a).Value = Name
        Sheets("WAR").Range("B" & a).Value = WARvalue
        a = a + 1
    Next
MyExit2:
    If bShowBar Then Unload UserForm1
    ' çàãðóçêà ãîñòåé
    
     Set http = CreateObject("WinHttp.WinHttpRequest.5.1") 'WinHttp.WinHttpRequest.5.1   MSXML2.XMLHTTP
    http.Open "GET", "https://www.espn.com/mlb/team/stats/_/type/pitching/name/" & away
    http.setRequestHeader "Upgrade-Insecure-Requests", "1"
    http.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/75.0.3770.142 Safari/537.36 OPR/62.0.3331.116"
    http.setRequestHeader "DNT", "1"
    http.setRequestHeader "Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8"
    http.Send
    http.waitForResponse = True
    fs_input = http.ResponseText

    fs_rows = Split(fs_input, "athlete")
    a = 3
    
    fs_rows_length = UBound(fs_rows) - LBound(fs_rows)
    Call Show_PrBar_Or_No(fs_rows_length, "Çàãðóçêà ïèò÷åðîâ Ãîñòè(WAR)...")
    For i = 2 To fs_rows_length
        If bShowBar Then Call MyProgresBar
        fs_row = Split(fs_rows(i), ",")
        fs_row_length = UBound(fs_row) - LBound(fs_row)
        For j = 0 To fs_row_length

            fs_nrow = Replace(fs_row(j), """", "")
            fs_nrow = Replace(fs_nrow, "{", "")
            If j = 0 Then fs_nrow = Replace(fs_nrow, ":", "", 1, 1)
            fs_nrow2 = Split(fs_nrow, ":")
            If j = 0 Then Name = fs_nrow2(1): j = 85
                If a = 3 Then
                    ProvName = Name
                ElseIf ProvName = Name Then GoTo MyExit3
                End If
            If j = 88 Then
                WARvalue = fs_nrow2(1)
                Exit For
            End If
            
        Next
        
        Sheets("WAR").Range("D" & a).Value = Name
        Sheets("WAR").Range("E" & a).Value = WARvalue
        a = a + 1
    Next
MyExit3:
    
    
    Sheets("WAR").Range("D" & a).Value = "Õèòòåð"
    If bShowBar Then Unload UserForm1
    a = a + 1
    b = a
    ' Çàãðóçêà õèòòåðîâ äîìà
    
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1") 'WinHttp.WinHttpRequest.5.1   MSXML2.XMLHTTP
    http.Open "GET", "https://www.espn.com/mlb/team/stats/_/name/" & away
    http.setRequestHeader "Upgrade-Insecure-Requests", "1"
    http.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/75.0.3770.142 Safari/537.36 OPR/62.0.3331.116"
    http.setRequestHeader "DNT", "1"
    http.setRequestHeader "Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8"
    http.Send
    http.waitForResponse = True
    fs_input = http.ResponseText
    
    fs_rows = Split(fs_input, "athlete")

    
    fs_rows_length = UBound(fs_rows) - LBound(fs_rows)
    Call Show_PrBar_Or_No(fs_rows_length, "Çàãðóçêà õèòòåðîâ Ãîñòè(WAR)...")
    For i = 2 To fs_rows_length
        If bShowBar Then Call MyProgresBar
        fs_row = Split(fs_rows(i), ",")
        fs_row_length = UBound(fs_row) - LBound(fs_row)
        For j = 0 To fs_row_length

            fs_nrow = Replace(fs_row(j), """", "")
            fs_nrow = Replace(fs_nrow, "{", "")
            If j = 0 Then fs_nrow = Replace(fs_nrow, ":", "", 1, 1)
            fs_nrow2 = Split(fs_nrow, ":")
            If j = 0 Then Name = fs_nrow2(1): j = 92
                If a = b Then
                    ProvName = Name
                ElseIf ProvName = Name Then GoTo MyExit4
                End If
            If j = 93 Then
                WARvalue = fs_nrow2(1)
                Exit For
            End If
            
        Next
        
        Sheets("WAR").Range("D" & a).Value = Name
        Sheets("WAR").Range("E" & a).Value = WARvalue
        a = a + 1
    Next
MyExit4:
    If bShowBar Then Unload UserForm1
End Sub

Sub ClearList6()
    Sheets("WAR").Activate
    Sheets("WAR").Range("A3:F150").Select
    Selection.ClearContents
    Sheets("WAR").Range("A3").Select
End Sub

Sub SetGraf()
'
' SetGraf Ìàêðîñ
'
'
    Sheets("Âûâîä").Activate
    ActiveSheet.ChartObjects("Äèàãðàììà 1").Activate
    
    If ActiveChart.SeriesCollection.Count > 1 Then
        ActiveChart.SeriesCollection(1).Delete
        ActiveChart.SeriesCollection(1).Delete
    End If
    
    For i = 3 To 200
        If Sheets("Äîìà").Range("A" & i) = 1 Then Exit For
    Next
    Dim Xstr1, Ystr1, Xstr2, Ystr2 As String
    a = i + 20
    
    Xstr1 = "Äîìà!$A$" & i
    Xstr1 = Xstr1 & ":$A$"
    Xstr1 = Xstr1 & a
    
    Ystr1 = "Äîìà!$X$" & i
    Ystr1 = Ystr1 & ":$X$"
    Ystr1 = Ystr1 & a
    
    For i = 3 To 200
        If Sheets("Ãîñòè").Range("A" & i) = 1 Then Exit For
    Next
    a = i + 20
    
    Xstr2 = "Ãîñòè!$A$" & i
    Xstr2 = Xstr2 & ":$A$"
    Xstr2 = Xstr2 & a
    
    Ystr2 = "Ãîñòè!$X$" & i
    Ystr2 = Ystr2 & ":$X$"
    Ystr2 = Ystr2 & a
    
    Sheets("Âûâîä").Activate
    ActiveSheet.ChartObjects("Äèàãðàììà 1").Activate
    
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(1).XValues = "=" & Xstr1
    ActiveChart.FullSeriesCollection(1).Values = "=" & Ystr1
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(2).XValues = "=" & Xstr2
    ActiveChart.FullSeriesCollection(2).Values = "=" & Ystr2
End Sub
