
Option Explicit
Public bShowBar As Boolean
Public dblProgressWidth As Double, dblStep As Double, dblPercent As Double

Sub MyProgresBar()
    dblProgressWidth = dblProgressWidth + dblStep
    UserForm1.FrameProgress.Width = dblProgressWidth
    If dblProgressWidth > dblPercent Then
        UserForm1.lblPercentWhite.Caption = Format(dblPercent / UserForm1.FramePrgBar.Width, "0%")
        UserForm1.lblPercentBlack.Caption = UserForm1.lblPercentWhite.Caption
        dblPercent = dblPercent + dblStep
        UserForm1.Repaint
        DoEvents
    End If
End Sub

Function Show_PrBar_Or_No(lCnt As Long, Optional sUfCaption As String = "Âûïîëíåíèå...")
    bShowBar = (lCnt > 10)
    If bShowBar = False Then Exit Function
    
    UserForm1.Caption = sUfCaption
    dblStep = UserForm1.FramePrgBar.Width / lCnt
    UserForm1.lblPercentWhite.Left = 96
    UserForm1.lblPercentBlack.Left = UserForm1.lblPercentWhite.Left
    
    UserForm1.Show 0
    dblPercent = 0: dblProgressWidth = 0
End Function

