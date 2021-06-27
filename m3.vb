Private Declare Sub Sleep Lib "kernel32" (ByVal milliseconds As Long)
Sub GetPitchers()
    Application.ScreenUpdating = False
    home_team = Sheets("Ñòàòà").Range("A32").Value
    If home_team = 1 Then
        home_team = "109" 'Arizona Diamondbacks
    ElseIf home_team = 2 Then
        home_team = "144" 'Atlanta Braves
    ElseIf home_team = 3 Then
        home_team = "110" 'Baltimore Orioles
    ElseIf home_team = 4 Then
        home_team = "111" 'Boston Red Sox
    ElseIf home_team = 5 Then
        home_team = "112" 'Chicago Cubs
    ElseIf home_team = 6 Then
        home_team = "145" 'Chicago White Sox
    ElseIf home_team = 7 Then
        home_team = "113" 'Cincinnati Reds
    ElseIf home_team = 8 Then
        home_team = "114" 'Cleveland Indians
    ElseIf home_team = 9 Then
        home_team = "115" 'Colorado Rockies
    ElseIf home_team = 10 Then
        home_team = "116" 'Detroit Tigers
    ElseIf home_team = 11 Then
        home_team = "117" 'Houston Astros
    ElseIf home_team = 12 Then
        home_team = "118" 'Kansas City Royals
    ElseIf home_team = 13 Then
        home_team = "108" 'Los Angeles Angels
    ElseIf home_team = 14 Then
        home_team = "119" 'Los Angeles Dodgers
    ElseIf home_team = 15 Then
        home_team = "146" 'Miami Marlins
    ElseIf home_team = 16 Then
        home_team = "158" 'Milwaukee Brewers
    ElseIf home_team = 17 Then
        home_team = "142" 'Minnesota Twins
    ElseIf home_team = 18 Then
        home_team = "121" 'New York Mets
    ElseIf home_team = 19 Then
        home_team = "147" 'New York Yankees
    ElseIf home_team = 20 Then
        home_team = "133" 'Oakland Athletics
    ElseIf home_team = 21 Then
        home_team = "143" 'Philadelphia Phillies
    ElseIf home_team = 22 Then
        home_team = "134" 'Pittsburgh Pirates
    ElseIf home_team = 23 Then
        home_team = "135" 'San Diego Padres
    ElseIf home_team = 24 Then
        home_team = "137" 'San Francisco Giants
    ElseIf home_team = 25 Then
        home_team = "136" 'Seattle Mariners
    ElseIf home_team = 26 Then
        home_team = "138" 'St. Louis Cardinals
    ElseIf home_team = 27 Then
        home_team = "139" 'Tampa Bay Rays
    ElseIf home_team = 28 Then
        home_team = "140" 'Texas Rangers
    ElseIf home_team = 29 Then
        home_team = "141" 'Toronto Blue Jays
    ElseIf home_team = 30 Then
        home_team = "120" 'Washington Nationals
    ElseIf home_team = False Then
        Exit Sub
    End If
       
    away_team = Sheets("Ñòàòà").Range("B32").Value
    If away_team = 1 Then
        away_team = "109" 'Arizona Diamondbacks
    ElseIf away_team = 2 Then
        away_team = "144" 'Atlanta Braves
    ElseIf away_team = 3 Then
        away_team = "110" 'Baltimore Orioles
    ElseIf away_team = 4 Then
        away_team = "111" 'Boston Red Sox
    ElseIf away_team = 5 Then
        away_team = "112" 'Chicago Cubs
    ElseIf away_team = 6 Then
        away_team = "145" 'Chicago White Sox
    ElseIf away_team = 7 Then
        away_team = "113" 'Cincinnati Reds
    ElseIf away_team = 8 Then
        away_team = "114" 'Cleveland Indians
    ElseIf away_team = 9 Then
        away_team = "115" 'Colorado Rockies
    ElseIf away_team = 10 Then
        away_team = "116" 'Detroit Tigers
    ElseIf away_team = 11 Then
        away_team = "117" 'Houston Astros
    ElseIf away_team = 12 Then
        away_team = "118" 'Kansas City Royals
    ElseIf away_team = 13 Then
        away_team = "108" 'Los Angeles Angels
    ElseIf away_team = 14 Then
        away_team = "119" 'Los Angeles Dodgers
    ElseIf away_team = 15 Then
        away_team = "146" 'Miami Marlins
    ElseIf away_team = 16 Then
        away_team = "158" 'Milwaukee Brewers
    ElseIf away_team = 17 Then
        away_team = "142" 'Minnesota Twins
    ElseIf away_team = 18 Then
        away_team = "121" 'New York Mets
    ElseIf away_team = 19 Then
        away_team = "147" 'New York Yankees
    ElseIf away_team = 20 Then
        away_team = "133" 'Oakland Athletics
    ElseIf away_team = 21 Then
        away_team = "143" 'Philadelphia Phillies
    ElseIf away_team = 22 Then
        away_team = "134" 'Pittsburgh Pirates
    ElseIf away_team = 23 Then
        away_team = "135" 'San Diego Padres
    ElseIf away_team = 24 Then
        away_team = "137" 'San Francisco Giants
    ElseIf away_team = 25 Then
        away_team = "136" 'Seattle Mariners
    ElseIf away_team = 26 Then
        away_team = "138" 'St. Louis Cardinals
    ElseIf away_team = 27 Then
        away_team = "139" 'Tampa Bay Rays
    ElseIf away_team = 28 Then
        away_team = "140" 'Texas Rangers
    ElseIf away_team = 29 Then
        away_team = "141" 'Toronto Blue Jays
    ElseIf away_team = 30 Then
        away_team = "120" 'Washington Nationals
    ElseIf away_team = False Then
        Exit Sub
    End If
       
    ' Çàãðóçêà ïèò÷åðîâ Äîìà
       
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1") 'WinHttp.WinHttpRequest.5.1   MSXML2.XMLHTTP
    http.Open "GET", "http://mlb.mlb.com/pubajax/wf/flow/stats.splayer?season=2019&sort_order=%27asc%27&sort_column=%27era%27&stat_type=pitching&page_type=SortablePlayer&team_id=" & home_team & "&game_type=%27R%27&player_pool=ALL&season_type=ANY&sport_code=%27mlb%27&results=1000&position=%271%27&recSP=1&recPP=50"
    http.setRequestHeader "Accept", "application/json, text/javascript, */*; q=0.01"
    http.setRequestHeader "DNT", "1"
    http.setRequestHeader "X-Requested-With", "XMLHttpRequest"
    http.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/75.0.3770.100 Safari/537.36 OPR/62.0.3331.99"
    http.Send
    http.waitForResponse = True
    'Sleep (20000)
    fs_input = http.ResponseText
    
    fs_rows = Split(fs_input, "gidp")
    a = 5
    Dim fs_rows_length As Long
    fs_rows_length = UBound(fs_rows) - LBound(fs_rows)
    Call Show_PrBar_Or_No(fs_rows_length, "Çàãðóçêà ïèò÷åðîâ Äîìà...")
    GetTable 4
    For i = 1 To fs_rows_length
        If bShowBar Then Call MyProgresBar
        fs_row = Split(fs_rows(i), ",")
        fs_row_length = UBound(fs_row) - LBound(fs_row)
        For j = 0 To fs_row_length
            fs_nrow = Replace(fs_row(j), """", "")
            fs_nrow2 = Split(fs_nrow, ":")
            
            If fs_nrow2(0) = "name_display_first_last" Then Name = fs_nrow2(1)
            If fs_nrow2(0) = "w" Then w = fs_nrow2(1)
            If fs_nrow2(0) = "l" Then l = fs_nrow2(1)
            If fs_nrow2(0) = "era" Then era = fs_nrow2(1)
            If fs_nrow2(0) = "g" Then g = fs_nrow2(1)
            If fs_nrow2(0) = "gs" Then gs = fs_nrow2(1)
            If fs_nrow2(0) = "sv" Then sv = fs_nrow2(1)
            If fs_nrow2(0) = "svo" Then svo = fs_nrow2(1)
            If fs_nrow2(0) = "ip" Then ip = fs_nrow2(1)
            If fs_nrow2(0) = "h" Then h = fs_nrow2(1)
            If fs_nrow2(0) = "r" Then r = fs_nrow2(1)
            If fs_nrow2(0) = "er" Then fs_nrow2(1) = Replace(fs_nrow2(1), "}", ""): fs_nrow2(1) = Replace(fs_nrow2(1), "]", ""): er = fs_nrow2(1)
            If fs_nrow2(0) = "hr" Then hr = fs_nrow2(1)
            If fs_nrow2(0) = "bb" Then bb = fs_nrow2(1)
            If fs_nrow2(0) = "so" Then so = fs_nrow2(1)
            If fs_nrow2(0) = "avg" Then avg = fs_nrow2(1)
            If fs_nrow2(0) = "whip" Then whip = fs_nrow2(1)
            If fs_nrow2(0) = "cg" Then cg = fs_nrow2(1)
            If fs_nrow2(0) = "sho" Then sho = fs_nrow2(1)
            If fs_nrow2(0) = "hb" Then hb = fs_nrow2(1)
            If fs_nrow2(0) = "ibb" Then ibb = fs_nrow2(1)
            If fs_nrow2(0) = "gf" Then gf = fs_nrow2(1)
            If fs_nrow2(0) = "hld" Then hld = fs_nrow2(1)
            If fs_nrow2(0) = "" Then gidp = fs_nrow2(1)
            If fs_nrow2(0) = "go" Then go = fs_nrow2(1)
            If fs_nrow2(0) = "ao" Then ao = fs_nrow2(1)
            If fs_nrow2(0) = "wp" Then wp = fs_nrow2(1)
            If fs_nrow2(0) = "bk" Then bk = fs_nrow2(1)
            If fs_nrow2(0) = "sb" Then sb = fs_nrow2(1)
            If fs_nrow2(0) = "cs" Then cs = fs_nrow2(1)
            If fs_nrow2(0) = "pk" Then pk = fs_nrow2(1)
            If fs_nrow2(0) = "tbf" Then tbf = fs_nrow2(1)
            If fs_nrow2(0) = "np" Then np = fs_nrow2(1)
        Next
        
        Sheets("Ñòàòà").Range("C" & a).Value = Name
        Sheets("Ñòàòà").Range("D" & a).Value = w
        Sheets("Ñòàòà").Range("E" & a).Value = l
        Sheets("Ñòàòà").Range("F" & a).Value = era
        Sheets("Ñòàòà").Range("G" & a).Value = g
        Sheets("Ñòàòà").Range("H" & a).Value = gs
        Sheets("Ñòàòà").Range("I" & a).Value = sv
        Sheets("Ñòàòà").Range("J" & a).Value = svo
        Sheets("Ñòàòà").Range("K" & a).Value = ip
        Sheets("Ñòàòà").Range("L" & a).Value = h
        Sheets("Ñòàòà").Range("M" & a).Value = r
        Sheets("Ñòàòà").Range("N" & a).Value = er
        Sheets("Ñòàòà").Range("O" & a).Value = hr
        Sheets("Ñòàòà").Range("P" & a).Value = bb
        Sheets("Ñòàòà").Range("Q" & a).Value = so
        Sheets("Ñòàòà").Range("R" & a).Value = avg
        Sheets("Ñòàòà").Range("S" & a).Value = whip
        Sheets("Ñòàòà").Range("T" & a).Value = cg
        Sheets("Ñòàòà").Range("U" & a).Value = sho
        Sheets("Ñòàòà").Range("V" & a).Value = hb
        Sheets("Ñòàòà").Range("W" & a).Value = ibb
        Sheets("Ñòàòà").Range("X" & a).Value = gf
        Sheets("Ñòàòà").Range("Y" & a).Value = hld
        Sheets("Ñòàòà").Range("Z" & a).Value = gidp
        Sheets("Ñòàòà").Range("AA" & a).Value = go
        Sheets("Ñòàòà").Range("AB" & a).Value = ao
        Sheets("Ñòàòà").Range("AC" & a).Value = wp
        Sheets("Ñòàòà").Range("AD" & a).Value = bk
        Sheets("Ñòàòà").Range("AE" & a).Value = sb
        Sheets("Ñòàòà").Range("AF" & a).Value = cs
        Sheets("Ñòàòà").Range("AG" & a).Value = pk
        Sheets("Ñòàòà").Range("AH" & a).Value = tbf
        Sheets("Ñòàòà").Range("AI" & a).Value = np
        a = a + 1
    Next
    If bShowBar Then Unload UserForm1
    
    
    
    ' Çàãðóçêà ïèò÷åðîâ ãîñòè
    fs_input = ""
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", "http://mlb.mlb.com/pubajax/wf/flow/stats.splayer?season=2019&sort_order=%27asc%27&sort_column=%27era%27&stat_type=pitching&page_type=SortablePlayer&team_id=" & away_team & "&game_type=%27R%27&player_pool=ALL&season_type=ANY&sport_code=%27mlb%27&results=1000&position=%271%27&recSP=1&recPP=50"
    http.setRequestHeader "Accept", "application/json, text/javascript, */*; q=0.01"
    http.setRequestHeader "DNT", "1"
    http.setRequestHeader "X-Requested-With", "XMLHttpRequest"
    http.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/75.0.3770.100 Safari/537.36 OPR/62.0.3331.99"
    http.Send
    http.waitForResponse = True
    'Sleep (20000)
    fs_input = http.ResponseText
    
    fs_rows = Split(fs_input, "gidp")
    
    GetTable (a)
    a = a + 1
    'Dim fs_rows_length As Long
    fs_rows_length = UBound(fs_rows) - LBound(fs_rows)
    Call Show_PrBar_Or_No(fs_rows_length, "Çàãðóçêà ïèò÷åðîâ Ãîñòè...")
    For i = 1 To fs_rows_length
        If bShowBar Then Call MyProgresBar
        fs_row = Split(fs_rows(i), ",")
        fs_row_length = UBound(fs_row) - LBound(fs_row)
        For j = 0 To fs_row_length
            fs_nrow = Replace(fs_row(j), """", "")
            fs_nrow2 = Split(fs_nrow, ":")
            
            If fs_nrow2(0) = "name_display_first_last" Then Name = fs_nrow2(1)
            If fs_nrow2(0) = "w" Then w = fs_nrow2(1)
            If fs_nrow2(0) = "l" Then l = fs_nrow2(1)
            If fs_nrow2(0) = "era" Then era = fs_nrow2(1)
            If fs_nrow2(0) = "g" Then g = fs_nrow2(1)
            If fs_nrow2(0) = "gs" Then gs = fs_nrow2(1)
            If fs_nrow2(0) = "sv" Then sv = fs_nrow2(1)
            If fs_nrow2(0) = "svo" Then svo = fs_nrow2(1)
            If fs_nrow2(0) = "ip" Then ip = fs_nrow2(1)
            If fs_nrow2(0) = "h" Then h = fs_nrow2(1)
            If fs_nrow2(0) = "r" Then r = fs_nrow2(1)
            If fs_nrow2(0) = "hr" Then hr = fs_nrow2(1)
            If fs_nrow2(0) = "bb" Then bb = fs_nrow2(1)
            If fs_nrow2(0) = "so" Then so = fs_nrow2(1)
            If fs_nrow2(0) = "avg" Then avg = fs_nrow2(1)
            If fs_nrow2(0) = "whip" Then whip = fs_nrow2(1)
            If fs_nrow2(0) = "cg" Then cg = fs_nrow2(1)
            If fs_nrow2(0) = "sho" Then sho = fs_nrow2(1)
            If fs_nrow2(0) = "hb" Then hb = fs_nrow2(1)
            If fs_nrow2(0) = "ibb" Then ibb = fs_nrow2(1)
            If fs_nrow2(0) = "gf" Then gf = fs_nrow2(1)
            If fs_nrow2(0) = "hld" Then hld = fs_nrow2(1)
            If fs_nrow2(0) = "" Then gidp = fs_nrow2(1)
            If fs_nrow2(0) = "go" Then go = fs_nrow2(1)
            If fs_nrow2(0) = "ao" Then ao = fs_nrow2(1)
            If fs_nrow2(0) = "wp" Then wp = fs_nrow2(1)
            If fs_nrow2(0) = "bk" Then bk = fs_nrow2(1)
            If fs_nrow2(0) = "sb" Then sb = fs_nrow2(1)
            If fs_nrow2(0) = "cs" Then cs = fs_nrow2(1)
            If fs_nrow2(0) = "pk" Then pk = fs_nrow2(1)
            If fs_nrow2(0) = "tbf" Then tbf = fs_nrow2(1)
            If fs_nrow2(0) = "np" Then np = fs_nrow2(1)
            
        Next
        
        Sheets("Ñòàòà").Range("C" & a).Value = Name
        Sheets("Ñòàòà").Range("D" & a).Value = w
        Sheets("Ñòàòà").Range("E" & a).Value = l
        Sheets("Ñòàòà").Range("F" & a).Value = era
        Sheets("Ñòàòà").Range("G" & a).Value = g
        Sheets("Ñòàòà").Range("H" & a).Value = gs
        Sheets("Ñòàòà").Range("I" & a).Value = sv
        Sheets("Ñòàòà").Range("J" & a).Value = svo
        Sheets("Ñòàòà").Range("K" & a).Value = ip
        Sheets("Ñòàòà").Range("L" & a).Value = h
        Sheets("Ñòàòà").Range("M" & a).Value = r
        Sheets("Ñòàòà").Range("N" & a).Value = er
        Sheets("Ñòàòà").Range("O" & a).Value = hr
        Sheets("Ñòàòà").Range("P" & a).Value = bb
        Sheets("Ñòàòà").Range("Q" & a).Value = so
        Sheets("Ñòàòà").Range("R" & a).Value = avg
        Sheets("Ñòàòà").Range("S" & a).Value = whip
        Sheets("Ñòàòà").Range("T" & a).Value = cg
        Sheets("Ñòàòà").Range("U" & a).Value = sho
        Sheets("Ñòàòà").Range("V" & a).Value = hb
        Sheets("Ñòàòà").Range("W" & a).Value = ibb
        Sheets("Ñòàòà").Range("X" & a).Value = gf
        Sheets("Ñòàòà").Range("Y" & a).Value = hld
        Sheets("Ñòàòà").Range("Z" & a).Value = gidp
        Sheets("Ñòàòà").Range("AA" & a).Value = go
        Sheets("Ñòàòà").Range("AB" & a).Value = ao
        Sheets("Ñòàòà").Range("AC" & a).Value = wp
        Sheets("Ñòàòà").Range("AD" & a).Value = bk
        Sheets("Ñòàòà").Range("AE" & a).Value = sb
        Sheets("Ñòàòà").Range("AF" & a).Value = cs
        Sheets("Ñòàòà").Range("AG" & a).Value = pk
        Sheets("Ñòàòà").Range("AH" & a).Value = tbf
        Sheets("Ñòàòà").Range("AI" & a).Value = np
        a = a + 1
    Next
    If bShowBar Then Unload UserForm1
    
    ' õèòòåð äîìà
    
    fs_input = ""
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", "http://mlb.mlb.com/pubajax/wf/flow/stats.splayer?season=2019&sort_order=%27desc%27&sort_column=%27avg%27&stat_type=hitting&page_type=SortablePlayer&team_id=" & home_team & "&game_type=%27R%27&player_pool=ALL&season_type=ANY&sport_code=%27mlb%27&results=1000&recSP=1&recPP=50"
    http.setRequestHeader "Accept", "application/json, text/javascript, */*; q=0.01"
    http.setRequestHeader "DNT", "1"
    http.setRequestHeader "X-Requested-With", "XMLHttpRequest"
    http.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/75.0.3770.100 Safari/537.36 OPR/62.0.3331.99"
    http.Send
    http.waitForResponse = True
    'Sleep (20000)
    fs_input = http.ResponseText
    
    fs_rows = Split(fs_input, "gidp")
    
    Sheets("Ñòàòà").Range("C" & a).Value = "õèòòåðû"
    Sheets("Ñòàòà").Range("C" & a).HorizontalAlignment = xlCenter
    
    a = a + 1
    GetTable2 (a)
    a = a + 1
    'Dim fs_rows_length As Long
    fs_rows_length = UBound(fs_rows) - LBound(fs_rows)
    Call Show_PrBar_Or_No(fs_rows_length, "Çàãðóçêà õèòòåðîâ Äîìà...")
    For i = 1 To fs_rows_length
        If bShowBar Then Call MyProgresBar
        fs_row = Split(fs_rows(i), ",")
        fs_row_length = UBound(fs_row) - LBound(fs_row)
        For j = 0 To fs_row_length
            fs_nrow = Replace(fs_row(j), """", "")
            fs_nrow2 = Split(fs_nrow, ":")
            
            
            If fs_nrow2(0) = "name_display_first_last" Then Name = fs_nrow2(1)
            If fs_nrow2(0) = "pos" Then pos = fs_nrow2(1)
            If fs_nrow2(0) = "g" Then g = fs_nrow2(1)
            If fs_nrow2(0) = "ab" Then ab = fs_nrow2(1)
            If fs_nrow2(0) = "r" Then r = fs_nrow2(1)
            If fs_nrow2(0) = "h" Then h = fs_nrow2(1)
            If fs_nrow2(0) = "d" Then b2 = fs_nrow2(1)
            If fs_nrow2(0) = "t" Then b3 = fs_nrow2(1)
            If fs_nrow2(0) = "hr" Then hr = fs_nrow2(1)
            If fs_nrow2(0) = "rbi" Then rbi = fs_nrow2(1)
            If fs_nrow2(0) = "bb" Then bb = fs_nrow2(1)
            If fs_nrow2(0) = "so" Then so = fs_nrow2(1)
            If fs_nrow2(0) = "sb" Then sb = fs_nrow2(1)
            If fs_nrow2(0) = "cs" Then cs = fs_nrow2(1)
            If fs_nrow2(0) = "avg" Then avg = fs_nrow2(1)
            If fs_nrow2(0) = "obp" Then obp = fs_nrow2(1)
            If fs_nrow2(0) = "slg" Then slg = fs_nrow2(1)
            If fs_nrow2(0) = "ops" Then ops = fs_nrow2(1)
            If fs_nrow2(0) = "ibb" Then ibb = fs_nrow2(1)
            If fs_nrow2(0) = "hbp" Then hbp = fs_nrow2(1)
            If fs_nrow2(0) = "sac" Then sac = fs_nrow2(1)
            If fs_nrow2(0) = "sf" Then sf = fs_nrow2(1)
            If fs_nrow2(0) = "tb" Then tb = fs_nrow2(1)
            If fs_nrow2(0) = "xbh" Then xbh = fs_nrow2(1)
            If fs_nrow2(0) = "gdp" Then gdp = fs_nrow2(1)
            If fs_nrow2(0) = "go" Then go = fs_nrow2(1)
            If fs_nrow2(0) = "ao" Then ao = fs_nrow2(1)
            If fs_nrow2(0) = "go_ao" Then go_ao = fs_nrow2(1)
            If fs_nrow2(0) = "np" Then np = fs_nrow2(1)
            If fs_nrow2(0) = "tpa" Then tpa = fs_nrow2(1)
            
            
        Next
        
        Sheets("Ñòàòà").Range("C" & a).Value = Name
        Sheets("Ñòàòà").Range("D" & a).Value = pos
        Sheets("Ñòàòà").Range("E" & a).Value = g
        Sheets("Ñòàòà").Range("F" & a).Value = ab
        Sheets("Ñòàòà").Range("G" & a).Value = r
        Sheets("Ñòàòà").Range("H" & a).Value = h
        Sheets("Ñòàòà").Range("I" & a).Value = b2
        Sheets("Ñòàòà").Range("J" & a).Value = b3
        Sheets("Ñòàòà").Range("K" & a).Value = hr
        Sheets("Ñòàòà").Range("L" & a).Value = rbi
        Sheets("Ñòàòà").Range("M" & a).Value = bb
        Sheets("Ñòàòà").Range("N" & a).Value = so
        Sheets("Ñòàòà").Range("O" & a).Value = sb
        Sheets("Ñòàòà").Range("P" & a).Value = cs
        Sheets("Ñòàòà").Range("Q" & a).Value = avg
        Sheets("Ñòàòà").Range("R" & a).Value = obp
        Sheets("Ñòàòà").Range("S" & a).Value = slg
        Sheets("Ñòàòà").Range("T" & a).Value = ops
        Sheets("Ñòàòà").Range("U" & a).Value = ibb
        Sheets("Ñòàòà").Range("V" & a).Value = hbp
        Sheets("Ñòàòà").Range("W" & a).Value = sac
        Sheets("Ñòàòà").Range("X" & a).Value = sf
        Sheets("Ñòàòà").Range("Y" & a).Value = tb
        Sheets("Ñòàòà").Range("Z" & a).Value = xbh
        Sheets("Ñòàòà").Range("AA" & a).Value = gdp
        Sheets("Ñòàòà").Range("AB" & a).Value = go
        Sheets("Ñòàòà").Range("AC" & a).Value = ao
        Sheets("Ñòàòà").Range("AD" & a).Value = go_ao
        Sheets("Ñòàòà").Range("AE" & a).Value = np
        Sheets("Ñòàòà").Range("AF" & a).Value = tpa
        a = a + 1
    Next
    If bShowBar Then Unload UserForm1
    
    ' õèòòåð Ãîñòè
    fs_input = ""
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", "http://mlb.mlb.com/pubajax/wf/flow/stats.splayer?season=2019&sort_order=%27desc%27&sort_column=%27avg%27&stat_type=hitting&page_type=SortablePlayer&team_id=" & away_team & "&game_type=%27R%27&player_pool=ALL&season_type=ANY&sport_code=%27mlb%27&results=1000&recSP=1&recPP=50"
    http.setRequestHeader "Accept", "application/json, text/javascript, */*; q=0.01"
    http.setRequestHeader "DNT", "1"
    http.setRequestHeader "X-Requested-With", "XMLHttpRequest"
    http.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/75.0.3770.100 Safari/537.36 OPR/62.0.3331.99"
    http.Send
    http.waitForResponse = True
    'Sleep (20000)
    fs_input = http.ResponseText
    
    fs_rows = Split(fs_input, "gidp")
    
    GetTable2 (a)
    a = a + 1
    'Dim fs_rows_length As Long
    fs_rows_length = UBound(fs_rows) - LBound(fs_rows)
    Call Show_PrBar_Or_No(fs_rows_length, "Çàãðóçêà õèòòåðîâ Ãîñòè...")
    For i = 1 To fs_rows_length
        If bShowBar Then Call MyProgresBar
        fs_row = Split(fs_rows(i), ",")
        fs_row_length = UBound(fs_row) - LBound(fs_row)
        For j = 0 To fs_row_length
            fs_nrow = Replace(fs_row(j), """", "")
            fs_nrow2 = Split(fs_nrow, ":")
            
            
            If fs_nrow2(0) = "name_display_first_last" Then Name = fs_nrow2(1)
            If fs_nrow2(0) = "pos" Then pos = fs_nrow2(1)
            If fs_nrow2(0) = "g" Then g = fs_nrow2(1)
            If fs_nrow2(0) = "ab" Then ab = fs_nrow2(1)
            If fs_nrow2(0) = "r" Then r = fs_nrow2(1)
            If fs_nrow2(0) = "h" Then h = fs_nrow2(1)
            If fs_nrow2(0) = "d" Then b2 = fs_nrow2(1)
            If fs_nrow2(0) = "t" Then b3 = fs_nrow2(1)
            If fs_nrow2(0) = "hr" Then hr = fs_nrow2(1)
            If fs_nrow2(0) = "rbi" Then rbi = fs_nrow2(1)
            If fs_nrow2(0) = "bb" Then bb = fs_nrow2(1)
            If fs_nrow2(0) = "so" Then so = fs_nrow2(1)
            If fs_nrow2(0) = "sb" Then sb = fs_nrow2(1)
            If fs_nrow2(0) = "cs" Then cs = fs_nrow2(1)
            If fs_nrow2(0) = "avg" Then avg = fs_nrow2(1)
            If fs_nrow2(0) = "obp" Then obp = fs_nrow2(1)
            If fs_nrow2(0) = "slg" Then slg = fs_nrow2(1)
            If fs_nrow2(0) = "ops" Then ops = fs_nrow2(1)
            If fs_nrow2(0) = "ibb" Then ibb = fs_nrow2(1)
            If fs_nrow2(0) = "hbp" Then hbp = fs_nrow2(1)
            If fs_nrow2(0) = "sac" Then sac = fs_nrow2(1)
            If fs_nrow2(0) = "sf" Then sf = fs_nrow2(1)
            If fs_nrow2(0) = "tb" Then tb = fs_nrow2(1)
            If fs_nrow2(0) = "xbh" Then xbh = fs_nrow2(1)
            If fs_nrow2(0) = "gdp" Then gdp = fs_nrow2(1)
            If fs_nrow2(0) = "go" Then go = fs_nrow2(1)
            If fs_nrow2(0) = "ao" Then ao = fs_nrow2(1)
            If fs_nrow2(0) = "go_ao" Then go_ao = fs_nrow2(1)
            If fs_nrow2(0) = "np" Then np = fs_nrow2(1)
            If fs_nrow2(0) = "tpa" Then tpa = fs_nrow2(1)
            
            
        Next
        
        Sheets("Ñòàòà").Range("C" & a).Value = Name
        Sheets("Ñòàòà").Range("D" & a).Value = pos
        Sheets("Ñòàòà").Range("E" & a).Value = g
        Sheets("Ñòàòà").Range("F" & a).Value = ab
        Sheets("Ñòàòà").Range("G" & a).Value = r
        Sheets("Ñòàòà").Range("H" & a).Value = h
        Sheets("Ñòàòà").Range("I" & a).Value = b2
        Sheets("Ñòàòà").Range("J" & a).Value = b3
        Sheets("Ñòàòà").Range("K" & a).Value = hr
        Sheets("Ñòàòà").Range("L" & a).Value = rbi
        Sheets("Ñòàòà").Range("M" & a).Value = bb
        Sheets("Ñòàòà").Range("N" & a).Value = so
        Sheets("Ñòàòà").Range("O" & a).Value = sb
        Sheets("Ñòàòà").Range("P" & a).Value = cs
        Sheets("Ñòàòà").Range("Q" & a).Value = avg
        Sheets("Ñòàòà").Range("R" & a).Value = obp
        Sheets("Ñòàòà").Range("S" & a).Value = slg
        Sheets("Ñòàòà").Range("T" & a).Value = ops
        Sheets("Ñòàòà").Range("U" & a).Value = ibb
        Sheets("Ñòàòà").Range("V" & a).Value = hbp
        Sheets("Ñòàòà").Range("W" & a).Value = sac
        Sheets("Ñòàòà").Range("X" & a).Value = sf
        Sheets("Ñòàòà").Range("Y" & a).Value = tb
        Sheets("Ñòàòà").Range("Z" & a).Value = xbh
        Sheets("Ñòàòà").Range("AA" & a).Value = gdp
        Sheets("Ñòàòà").Range("AB" & a).Value = go
        Sheets("Ñòàòà").Range("AC" & a).Value = ao
        Sheets("Ñòàòà").Range("AD" & a).Value = go_ao
        Sheets("Ñòàòà").Range("AE" & a).Value = np
        Sheets("Ñòàòà").Range("AF" & a).Value = tpa
        a = a + 1
    Next
    If bShowBar Then Unload UserForm1
    
    ClearList6
    GetWAR
    
    
    Sheets("Ñòàòà").Activate
    Application.ScreenUpdating = True
End Sub

Sub GetTable(a As Integer)
    Sheets("Ñòàòà").Range("C" & a).Value = "Player"
    Sheets("Ñòàòà").Range("D" & a).Value = "W"
    Sheets("Ñòàòà").Range("E" & a).Value = "L"
    Sheets("Ñòàòà").Range("F" & a).Value = "ERA"
    Sheets("Ñòàòà").Range("G" & a).Value = "G"
    Sheets("Ñòàòà").Range("H" & a).Value = "GS"
    Sheets("Ñòàòà").Range("I" & a).Value = "SV"
    Sheets("Ñòàòà").Range("J" & a).Value = "SVO"
    Sheets("Ñòàòà").Range("K" & a).Value = "IP"
    Sheets("Ñòàòà").Range("L" & a).Value = "H"
    Sheets("Ñòàòà").Range("M" & a).Value = "R"
    Sheets("Ñòàòà").Range("N" & a).Value = "ER"
    Sheets("Ñòàòà").Range("O" & a).Value = "HR"
    Sheets("Ñòàòà").Range("P" & a).Value = "BB"
    Sheets("Ñòàòà").Range("Q" & a).Value = "SO"
    Sheets("Ñòàòà").Range("R" & a).Value = "AVG"
    Sheets("Ñòàòà").Range("S" & a).Value = "WHIP"
    Sheets("Ñòàòà").Range("T" & a).Value = "CG"
    Sheets("Ñòàòà").Range("U" & a).Value = "SHO"
    Sheets("Ñòàòà").Range("V" & a).Value = "HB"
    Sheets("Ñòàòà").Range("W" & a).Value = "IBB"
    Sheets("Ñòàòà").Range("X" & a).Value = "GF"
    Sheets("Ñòàòà").Range("Y" & a).Value = "HLD"
    Sheets("Ñòàòà").Range("Z" & a).Value = "GIDP"
    Sheets("Ñòàòà").Range("AA" & a).Value = "GO"
    Sheets("Ñòàòà").Range("AB" & a).Value = "AO"
    Sheets("Ñòàòà").Range("AC" & a).Value = "WP"
    Sheets("Ñòàòà").Range("AD" & a).Value = "BK"
    Sheets("Ñòàòà").Range("AE" & a).Value = "SB"
    Sheets("Ñòàòà").Range("AF" & a).Value = "CS"
    Sheets("Ñòàòà").Range("AG" & a).Value = "PK"
    Sheets("Ñòàòà").Range("AH" & a).Value = "TBF"
    Sheets("Ñòàòà").Range("AI" & a).Value = "NP"
    With Sheets("Ñòàòà").Range("C" & a, "AI" & a)
        .Font.Bold = True
        .HorizontalAlignment = xlRight
    End With
    
        
End Sub

Sub GetTable2(a As Integer)
    Sheets("Ñòàòà").Range("C" & a).Value = "Player"
    Sheets("Ñòàòà").Range("D" & a).Value = "POS"
    Sheets("Ñòàòà").Range("E" & a).Value = "G"
    Sheets("Ñòàòà").Range("F" & a).Value = "AB"
    Sheets("Ñòàòà").Range("G" & a).Value = "R"
    Sheets("Ñòàòà").Range("H" & a).Value = "H"
    Sheets("Ñòàòà").Range("I" & a).Value = "2B"
    Sheets("Ñòàòà").Range("J" & a).Value = "3B"
    Sheets("Ñòàòà").Range("K" & a).Value = "HR"
    Sheets("Ñòàòà").Range("L" & a).Value = "RBI"
    Sheets("Ñòàòà").Range("M" & a).Value = "BB"
    Sheets("Ñòàòà").Range("N" & a).Value = "SO"
    Sheets("Ñòàòà").Range("O" & a).Value = "SB"
    Sheets("Ñòàòà").Range("P" & a).Value = "CS"
    Sheets("Ñòàòà").Range("Q" & a).Value = "AVG"
    Sheets("Ñòàòà").Range("R" & a).Value = "OBP"
    Sheets("Ñòàòà").Range("S" & a).Value = "SLG"
    Sheets("Ñòàòà").Range("T" & a).Value = "OPS"
    Sheets("Ñòàòà").Range("U" & a).Value = "IBB"
    Sheets("Ñòàòà").Range("V" & a).Value = "HBP"
    Sheets("Ñòàòà").Range("W" & a).Value = "SAC"
    Sheets("Ñòàòà").Range("X" & a).Value = "SF"
    Sheets("Ñòàòà").Range("Y" & a).Value = "TB"
    Sheets("Ñòàòà").Range("Z" & a).Value = "XBH"
    Sheets("Ñòàòà").Range("AA" & a).Value = "GDP"
    Sheets("Ñòàòà").Range("AB" & a).Value = "GO"
    Sheets("Ñòàòà").Range("AC" & a).Value = "AO"
    Sheets("Ñòàòà").Range("AD" & a).Value = "GO_AO"
    Sheets("Ñòàòà").Range("AE" & a).Value = "NP"
    Sheets("Ñòàòà").Range("AF" & a).Value = "PA"
    With Sheets("Ñòàòà").Range("C" & a, "AF" & a)
        .Font.Bold = True
        .HorizontalAlignment = xlRight
    End With
End Sub

Sub ClearList3()
    Sheets("Ñòàòà").Activate
    Sheets("Ñòàòà").Range("C4:AI500").Select
    Selection.ClearContents
    Sheets("Ñòàòà").Range("C5:AI500").Font.Bold = False
    Sheets("Ñòàòà").Range("C5:AI500").HorizontalAlignment = xlLeft
    Sheets("Ñòàòà").Range("C4").Select
End Sub

Private Sub ColorTest2()
    MsgBox Range("B33").Interior.Color
End Sub
