Sub ZIPbet()
    a = 410
    b = 14
    Do
        a = a + 32
        b = b + 1
        NamberBet = Sheets("Àðõèâ").Range("A" & a)
    Loop While NamberBet <> ""
    
    Sheets("Àðõèâ").Range("A" & a) = b
    Sheets("Àðõèâ").Range("A" & a + 1) = "ñ÷åò"
    Sheets("Àðõèâ").Range("A" & a + 2) = "Ïèò÷åð"
    Sheets("Àðõèâ").Range("A" & a + 5) = "Õèòòåð"
    Sheets("Àðõèâ").Range("A" & a + 6) = "Äîìà"
    Sheets("Àðõèâ").Range("A" & a + 15) = "Ãîñòè"
    Sheets("Àðõèâ").Range("A" & a + 30) = "Îñàäêè"
    Sheets("Àðõèâ").Range("A" & a + 31) = "ñòàâêà"
    
    NameTeam = Sheets("Âûâîä").Range("E2:F2")
    Sheets("Àðõèâ").Range("B" & a, "C" & a) = NameTeam
    
    Pitcher = Sheets("Âûâîä").Range("E13:P15")
    Sheets("Àðõèâ").Range("B" & a + 2, "M" & a + 4) = Pitcher
    
    Hitter = Sheets("Âûâîä").Range("E21:Q45")
    Sheets("Àðõèâ").Range("B" & a + 5, "N" & a + 29) = Hitter
    
    
End Sub
