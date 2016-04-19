VERSION 5.00
Begin VB.MDIForm MDIForm1 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   5340
   ClientLeft      =   2460
   ClientTop       =   2460
   ClientWidth     =   7830
   LinkTopic       =   "MDIForm1"
   Begin VB.Menu mnuFl 
      Caption         =   "File"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

' Constants used in this module.
Private Const CTL_MARGIN = 200
Private Const wis_TOP = 0
Private Const wis_BOTTOM = 1

' Previously selected button.
Private PrevButton As Integer

' Mouse co-ordinates.
Private m_MouseX As Single
Private m_MouseY As Single

' Recurssion suppresent.
Private resizeNow As Boolean


' Declare an application class object.
' This class is the interface between the
' main form and the methods...
'Private WithEvents wisAppObj As wisApp
Dim wisAppobj As Object
' Control Index of the current canvas in focus.
Private m_CanvasIndex As Integer


Private Sub Command1_Click()
pnlSlider(0).Caption = Text1.Text
lbl(0).Caption = Text1.Text
End Sub


Private Sub MDIForm_Activate()
Exit Sub
'Call SetActiveWindow(gWindowHandle)
Static buttonsAligned As Boolean
If Not buttonsAligned Then
    ShowIcons 1
    buttonsAligned = True
End If

End Sub

Private Sub MDIForm_Load()
Call SetKannadaCaption
Dim i As Integer
'MsgBox "Befor picCanvas.BorderStyle "
picCanvas.BorderStyle = 0
'MsgBox "After picCanvas.BorderStyle "

'MsgBox "Before Center the window."
' Center the window.
Me.Left = Screen.Width / 2 - Me.Width / 2
Me.Top = Screen.Height / 2 - Me.Height / 2
'Set icon for the form caption
Me.Icon = LoadResPicture(161, vbResIcon)

' Read the toolbar layout information
' from toolbar.lyt.
Dim strLayoutFile As String
Dim nFile As Integer
Dim Txt As String

' Get a file handle.
 nFile = FreeFile
' Open the layout file
If gLangOffSet = wis_KannadaOffset Then
strLayoutFile = App.Path & "\tbarkan.lyt"
Else
strLayoutFile = App.Path & "\toolbar.lyt"
End If
On Error Resume Next
Open strLayoutFile For Input As nFile
If Err.Number = 53 Then
    MsgBox "Lay out file not found", , wis_MESSAGE_TITLE
    gDbTrans.CloseDB
    End
End If

' Read the contents at once.
Txt = Input(LOF(nFile), #nFile)
Close #nFile

' Create the Picture-toolbar from serialization.
cmdUP.Picture = LoadResPicture(136, vbResBitmap)
cmdDown.Picture = LoadResPicture(137, vbResBitmap)
cmdUP.ZOrder
cmdDown.ZOrder
'MsgBox "Befor serialise"
Serialize picToolbar, Txt
'MsgBox "After Serialise"
pnlSlider_Click 1


' Create an instance of application object.
If wisAppobj Is Nothing Then
    'Set wisAppobj = New wisApp
End If

End Sub


Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If gWindowHandle <> Me.hWnd And gWindowHandle <> 0 Then
    'MsgBox "First close the active window then close appliacation", vbInformation, wis_MESSAGE_TITLE
    MsgBox LoadResString(gLangOffSet + 808), vbInformation, wis_MESSAGE_TITLE
    'call setwindowpos(gwindowhandle,1,
    Call SetActiveWindow(gWindowHandle)
    Cancel = True
    Exit Sub
End If

' Ask for user confirmation.
Dim nRet As Integer
'nRet = MsgBox("Do you want to exit this application?", _
        vbQuestion + vbYesNo, wis_MESSAGE_TITLE)
nRet = MsgBox(LoadResString(gLangOffSet + 750), _
        vbQuestion + vbYesNo, wis_MESSAGE_TITLE)
If nRet = vbNo Then Cancel = True


End Sub


Private Sub MDIForm_Resize()
With PicOut
    .Top = 0
    '.Left = 0
    .Height = Height
    '.Width = Width
End With

On Error Resume Next
With picToolbar
    .Left = 0
    .Top = 0
    .Height = PicOut.ScaleHeight - StatusBar1.Height
    sizeBar.Left = .Left + .Width
End With
sizeBar.Top = 0
sizeBar.Height = PicOut.ScaleHeight - StatusBar1.Height
resizeGuide.Height = sizeBar.Height

With picViewport
    .Left = sizeBar.Left + sizeBar.Width
    .Width = PicOut.ScaleWidth - picToolbar.Width - sizeBar.Width
    .Height = PicOut.ScaleHeight - StatusBar1.Height
    .Top = 0
End With


' Redisplay all the controls.
' Hide all the controls.
Dim Ctl As Control
Dim CtlIndex As Integer
'For Each ctl In Me.Controls
'    ctlIndex = ctl.Index
'    If ctlIndex > 0 Then ctl.Visible = True
'Next

'' Redraw the logo.
'DrawLogo

End Sub


Private Sub Picture1_Click()

End Sub


Private Sub MDIForm_Unload(Cancel As Integer)
Dim CompName As String
CompName = String(255, 0)
Dim Retval As Long
Retval = GetComputerName(CompName, Len(CompName))

On Error Resume Next
If gWindowHandle <> Me.hWnd And gWindowHandle <> 0 Then
    CloseWindow (gWindowHandle)
End If

gDbTrans.CloseDB
Set gDbTrans = Nothing
Set gCurrUser = Nothing
Set wisAppobj = Nothing
'code added
If gLangOffSet = wis_KannadaOffset Then
    PauseApplication (1)
    AppActivate "Akruti Engine 1998", False
    SendKeys "%{F4}Y", True ' i & "{+}", false
End If
   End

End Sub

Private Sub AlignButtons()
Dim i As Integer
Dim Alignment As Integer
Dim TopCount As Integer
Dim BottomCount As Integer

' Align all the sliding panels.
For i = 0 To pnlSlider.Count - 1
    ' Get the alignment property.
    With pnlSlider(i)
        Alignment = Val(ExtractToken(.Tag, "Alignment"))
        If Alignment = wis_TOP Then
            .Top = (.Index - 1) * .Height
            If i <> 0 Then TopCount = TopCount + 1
        Else
            .Top = picToolbar.ScaleHeight - _
                    (pnlSlider.Count - .Index) * .Height
            BottomCount = BottomCount + 1
        End If
        .Width = picToolbar.ScaleWidth
        .Left = 0
    End With
Next


End Sub
' Returns the no. of buttons on the specified canvas,
' identified by canvas index.
Private Function BtnCount(TabPanelIndex As Integer) As Integer
On Error Resume Next
Dim i As Integer
Dim ctIndex As Integer

' Loop throug the collection of image buttons.
For i = 0 To img.Count - 1
    ' Get the Container index for the button.
    ctIndex = Val(ExtractToken(img(i).Tag, "Container"))
    ' If the container index is the same as
    ' the index of the tabpanel, increment the count.
    If ctIndex = TabPanelIndex Then
        BtnCount = BtnCount + 1
    End If
Next
'
End Function
Private Sub DrawBorder(btnindex As Integer, Bevel As Integer)
Const BORDER_MARGIN = 30

If btnindex < 0 Then Exit Sub
With img(btnindex)
    ' Set the bevel.
    If Bevel = 0 Then
        picCanvas.ForeColor = picCanvas.BackColor
    ElseIf Bevel = 1 Then
        picCanvas.ForeColor = vbWhite
    ElseIf Bevel = 2 Then
        picCanvas.ForeColor = vbBlack
    End If
    
    ' Draw the top line.
    picCanvas.Line (.Left - BORDER_MARGIN, .Top - BORDER_MARGIN) _
            -(.Left + .Width + BORDER_MARGIN, .Top - BORDER_MARGIN)
    ' Draw the left line.
    picCanvas.Line (.Left - BORDER_MARGIN, .Top - BORDER_MARGIN) _
            -(.Left - BORDER_MARGIN, .Top + .Height + BORDER_MARGIN)

    If Bevel = 0 Then
        picCanvas.ForeColor = picCanvas.BackColor
    ElseIf Bevel = 1 Then
        picCanvas.ForeColor = vbBlack
    ElseIf Bevel = 2 Then
        picCanvas.ForeColor = vbWhite
    End If
    ' Draw the bottom line.
    picCanvas.Line (.Left - BORDER_MARGIN, .Top + .Height + BORDER_MARGIN) _
            -(.Left + .Width + BORDER_MARGIN, .Top + .Height + BORDER_MARGIN)
    ' Draw the right side line.
    picCanvas.Line (.Left + .Width + BORDER_MARGIN, .Top - BORDER_MARGIN) _
            -(.Left + .Width + BORDER_MARGIN, .Top + .Height + BORDER_MARGIN)
End With
End Sub
Public Sub DrawLogo()
Dim ViewportWidth As Single
Dim ViewportHeight As Single
Dim ViewportLeft As Single
Dim ViewportTop As Single
Dim strBanner As String

strBanner = "INDEX-2000"

' Set the dimensions of the viewport.
ViewportLeft = picToolbar.Width + sizeBar.Width
ViewportTop = 0
ViewportWidth = Me.ScaleWidth - ViewportLeft
ViewportHeight = Me.ScaleHeight - StatusBar1.Height

'Me.Cls
' Print the logo.
Me.CurrentX = ViewportLeft + (ViewportWidth - Me.TextWidth(strBanner)) / 2
Me.CurrentY = ViewportTop + (ViewportHeight - Me.TextHeight(strBanner)) / 2
Me.ForeColor = vbBlue
Me.Print strBanner

' Print the shadow.
Me.CurrentX = ViewportLeft + (ViewportWidth - Me.TextWidth(strBanner)) / 2 - 25
Me.CurrentY = ViewportTop + (ViewportHeight - Me.TextHeight(strBanner)) / 2 - 25
Me.ForeColor = &HC0C0C0
Me.Print strBanner

End Sub

Private Sub LoadIconsToForm()

End Sub

Private Function PanelCount(AlignmentPos As Integer) As Integer
Dim i As Integer
Dim AlignmentVal As Integer

For i = 1 To pnlSlider.Count - 1
    AlignmentVal = Val(ExtractToken(pnlSlider(i).Tag, "Alignment"))
    If AlignmentPos = AlignmentVal Then
        PanelCount = PanelCount + 1
    End If
Next

End Function
Private Sub ResetScrollButtons()
' Decide whether or not to show the scroll buttons.
If picCanvas.Top + picCanvas.Height > picToolbar.ScaleHeight _
            - PanelCount(wis_BOTTOM) * pnlSlider(0).Height Then
    cmdDown.Visible = True
Else
    cmdDown.Visible = False
End If

If picCanvas.Top < PanelCount(wis_TOP) * pnlSlider(0).Height Then
    cmdUP.Visible = True
Else
    cmdUP.Visible = False
End If
picCanvas.Height = Me.Height - picCanvas.Top


End Sub
Private Function Serialize(obj As Object, Txt As String) As Boolean

Dim new_value As String
Dim token_name As String
Dim token_value As String
Dim ctl_Index As Integer

On Error Resume Next

new_value = Txt
' Examine each token in turn.
 Do
     ' Get the token name and value.
     GetToken new_value, token_name, token_value
     If token_name = "" Then Exit Do

     ' Examine each token and initialize.
     Select Case UCase$(token_name)
        Case "TOOLTAB"
            ' Load a copy of panel tab.
            Load pnlSlider(pnlSlider.Count)
            pnlSlider(pnlSlider.UBound).Visible = True
            ' Load a canvas for this tab.
            If Not Serialize(pnlSlider(pnlSlider.UBound), _
                    token_value) Then GoTo end_line

            #If NOT_NEEDED Then
            ' Set the height of the canvas, based on the
            ' no. of buttons loaded on it.
            picCanvas(picCanvas.UBound).Height = (BtnCount(picCanvas.UBound) - 1) _
                * (img(0).Height + lbl(0).Height + CTL_MARGIN)
            #End If

        Case "BUTTON"
            ' Load an img control
            Load img(img.Count)
            If Not Serialize(img(img.UBound), _
                    token_value) Then GoTo end_line
            ' Update the tag of this button with the
            ' index of the tooltab.
            With img(img.UBound)
                .Tag = putToken(.Tag, "Container", pnlSlider.UBound)
            End With
        Case "LABEL"
            ' Load a label.
            Load lbl(lbl.Count)
            If Not Serialize(lbl(lbl.UBound), _
                    token_value) Then GoTo end_line
        Case "ICON"
            img(img.UBound).Picture = LoadResPicture(Val(token_value), vbResIcon)
        Case "BITMAP"
            img(img.UBound).Picture = LoadResPicture(Val(token_value), vbResBitmap)
        Case "CAPTION"
            obj.Font.Name = gFontName
            obj.Font.Size = gFontSize
            obj.Caption = token_value
        Case "ALIGNMENT"
            If StrComp(token_value, "Top", vbTextCompare) = 0 Then
                obj.Tag = "Alignment=" & wis_TOP
            Else
                obj.Tag = "Alignment=" & wis_BOTTOM
            End If
        Case "KEY"
            obj.Tag = putToken(obj.Tag, "Key", token_value)
    End Select
Loop
Set obj = Nothing
Serialize = True

end_line:
    Exit Function

End Function
Private Sub SetKannadaCaption()
Dim ctrl As Control

On Error Resume Next
For Each ctrl In Me
    ctrl.FontName = gFontName
    If Not TypeOf ctrl Is ComboBox Then ctrl.Font.Size = gFontSize

Next

'Me.mnuFile.Caption = LoadResString(gLangOffSet + 158)
'Me.mnuAccounts.Caption = LoadResString(gLangOffSet + 90)
'Me.mnuExit.Caption = LoadResString(gLangOffSet + 30)
'Me.mnuSBAcc.Caption = LoadResString(gLangOffSet + 436)
'Me.mnuCAAcc.Caption = LoadResString(gLangOffSet + 437)
'Me.mnuFDAcc.Caption = LoadResString(gLangOffSet + 423)
'Me.mnuRDAcc.Caption = LoadResString(gLangOffSet + 424)
'Me.mnuPDAcc.Caption = LoadResString(gLangOffSet + 425)
'Me.mnuHelp.Caption = LoadResString(gLangOffSet + 16)
'Me.mnuAbout.Caption = LoadResString(gLangOffSet + 165)
'Me.mnuContents.Caption = LoadResString(gLangOffSet + 166)



End Sub

Private Sub SetMessage(strMsg As String)
If Trim$(strMsg) = "" Then
    strMsg = "Ready."
End If
With StatusBar1
    .Panels(1).Text = strMsg
    .Refresh
End With
End Sub

Private Sub ShowIcons(PanelIndex As Integer)
Dim i As Integer
Dim iconCount As Integer
Dim ctIndex As Integer
Dim PositionSet As Boolean
Dim PrevTop As Single
'PrevTop = PanelCount(wis_TOP) * pnlSlider(0).Height
PrevTop = CTL_MARGIN

' Hide or display icons depending upon the
' paneltab index, to which they are bound.
For i = 0 To img.Count - 1
    With img(i)
        ctIndex = Val(ExtractToken(.Tag, "Container"))
        If ctIndex = PanelIndex Then
            img(i).Visible = True
            lbl(i).Visible = True
            If Not PositionSet Then
                img(i).Top = PrevTop
                PositionSet = True
            Else
                img(i).Top = PrevTop + img(0).Height + lbl(0).Height + CTL_MARGIN
            End If
            lbl(i).Top = img(i).Top + img(i).Height + 10 '+ CTL_MARGIN
            iconCount = iconCount + 1
            PrevTop = img(i).Top

        Else
            img(i).Visible = False
            lbl(i).Visible = False
        End If
    End With
Next

' Set the height of the canvas.
'resizeNow = False
picCanvas.Height = (iconCount) * (img(0).Height + lbl(0).Height + CTL_MARGIN) + CTL_MARGIN
picCanvas.Top = PanelCount(wis_TOP) * pnlSlider(0).Height

End Sub
Private Sub cmdDown_Click()
Dim MoveDistance As Single
Dim ScrollDiff As Single
Dim TopCount As Integer
Dim BottomCount As Integer
Dim i As Integer
Dim Alignment As Integer

' Get the count of panels that are top and bottom aligned.
TopCount = PanelCount(wis_TOP)
BottomCount = PanelCount(wis_BOTTOM)

' See if the canvas can be moved up.
With picCanvas
    ScrollDiff = .Top + .Height - (picToolbar.ScaleHeight _
                - BottomCount * pnlSlider(0).Height)
    ' If no scope for further movement upwards, exit.
    If ScrollDiff <= 0 Then Exit Sub

    ' Set the move distance.
    MoveDistance = img(0).Height + lbl(0).Height + CTL_MARGIN
    If ScrollDiff < MoveDistance Then
        .Top = .Top - ScrollDiff
        cmdDown.Visible = False
    Else
        .Top = .Top - MoveDistance
    End If
    
    ' If the top is less than 0, show the down scroll button.
    If .Top < TopCount * pnlSlider(0).Height Then
        cmdUP.Visible = True
    End If
End With
End Sub
Private Sub cmdUP_Click()
Dim MoveDistance As Single
Dim ScrollDiff As Single
picCanvas.SetFocus
Dim TopCount As Integer
Dim BottomCount As Integer
Dim i As Integer
Dim Alignment As Integer

' Get the count of panels which are TOP and BOTTOM aligned.
TopCount = PanelCount(wis_TOP)
BottomCount = PanelCount(wis_BOTTOM)

' See if the canvas can be moved down.
With picCanvas
    ' If no scope for further downward movement, exit.
    If .Top >= TopCount * pnlSlider(0).Height Then
        cmdUP.Visible = False
        Exit Sub
    End If
    ' Set the move distance.
    MoveDistance = img(0).Height + lbl(0).Height + CTL_MARGIN
    If MoveDistance > (TopCount * pnlSlider(0).Height) - .Top Then
        .Top = (TopCount * pnlSlider(0).Height)
        cmdUP.Visible = False
    Else
        .Top = .Top + MoveDistance
    End If

    ' If the canvas scrolls past the bottom of viewport,
    ' display the scroll button near the bottom.
    If .Top + .Height > picToolbar.ScaleHeight - BottomCount * pnlSlider(0).Height Then
        cmdDown.Visible = True
    Else
        cmdDown.Visible = False
    End If
End With

End Sub

Private Sub Form_Click()
'Call SetActiveWindow(gWindowHandle)
End Sub


Private Sub Form_Load()
Dim i As Integer
'MsgBox "Befor picCanvas.BorderStyle "
picCanvas.BorderStyle = 0
'MsgBox "After picCanvas.BorderStyle "

'MsgBox "Before Center the window."
' Center the window.
Me.Left = Screen.Width / 2 - Me.Width / 2
Me.Top = Screen.Height / 2 - Me.Height / 2
'Set icon for the form caption
Me.Icon = LoadResPicture(161, vbResIcon)

' Read the toolbar layout information
' from toolbar.lyt.
Dim strLayoutFile As String
Dim nFile As Integer
Dim Txt As String

' Get a file handle.
 nFile = FreeFile
' Open the layout file
If gLangOffSet = wis_KannadaOffset Then
strLayoutFile = App.Path & "\tbarkan.lyt"
Else
strLayoutFile = App.Path & "\toolbar.lyt"
End If
On Error Resume Next
Open strLayoutFile For Input As nFile
If Err.Number = 53 Then
    MsgBox "Lay out file not found", , wis_MESSAGE_TITLE
    gDbTrans.CloseDB
    End
End If

' Read the contents at once.
Txt = Input(LOF(nFile), #nFile)
Close #nFile

' Create the Picture-toolbar from serialization.
cmdUP.Picture = LoadResPicture(136, vbResBitmap)
cmdDown.Picture = LoadResPicture(137, vbResBitmap)
cmdUP.ZOrder
cmdDown.ZOrder
'MsgBox "Befor serialise"
Serialize picToolbar, Txt
'MsgBox "After Serialise"
pnlSlider_Click 1


' Create an instance of application object.
If wisAppobj Is Nothing Then
    Set wisAppobj = New wisApp
End If
End Sub
Private Sub img_Click(Index As Integer)

'If  any modules are runnnig then do not Show other module till closing then
If gWindowHandle <> 0 And gWindowHandle <> Me.hWnd Then Exit Sub

'MsgBox Me.ButtonKey(Index)
Select Case UCase$(Me.ButtonKey(Index))
    Case "SBACC"
        'wisAppobj.ShowSBDialog
        'frmSBAcc.Show
        Dim SbClass As New clsSBAcc
        SbClass.Show
    Case "CAACC"
        wisAppobj.ShowCADialog
    Case "FDACC"
        'wisAppobj.ShowFDDialog
        frmFDAcc.Show
    Case "RDACC"
        wisAppobj.ShowRDDialog
    Case "PDACC"
        wisAppobj.ShowPDDialog
    Case "DLACC"
        wisAppobj.ShowDLDialog
    Case "MEMBERS"
        wisAppobj.ShowMemberDialog
    Case "LNACC"
        wisAppobj.ShowLoanDialog
    Case "CUSTINFO"
        wisAppobj.ShowCustInfo
    Case "TRACC"
        wisAppobj.ShowReportDialog (wisTradingAccount)
    Case "PLACC"
        wisAppobj.ShowReportDialog (wisProfitLossStatement)
    Case "BALANCESHEET"
        wisAppobj.ShowReportDialog (wisBalanceSheet)
    Case "DEBITCREDIT"
        wisAppobj.ShowReportDialog (wisDebitCreditStatement)
    Case "BANKBALANCE"
        wisAppobj.ShowReportDialog (wisBankBalance)
    Case "DDC"
        wisAppobj.ShowReportDialog (wisDailyDebitCredit)
    Case "UTILS"
        wisAppobj.ShowUtils
    Case "USERS"
        If gCurrUser Is Nothing Then
            Set gCurrUser = New clsUsers
        End If
        gCurrUser.ShowUserDialog
        'wisappobj.show
    Case "MATERIAL"
        wisAppobj.ShowMaterialDialog
    Case "BANKACC"
        wisAppobj.ShowBankDialog
    Case "CLEARING"
        wisAppobj.ShowClearingDialog
    Case "EXIT"
        Unload Me
        Set MDIForm1 = Nothing
End Select

End Sub

Private Sub img_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)
If Index <= 0 Then Exit Sub
DrawBorder Index, 2

End Sub
Private Sub img_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)
If PrevButton = Index Then Exit Sub
If Index <= 0 Then Exit Sub
' Remove the border for previous button, if present.
If PrevButton > 0 Then DrawBorder PrevButton, 0
DrawBorder Index, 1
PrevButton = Index
End Sub

Private Sub img_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)
If Index <= 0 Then Exit Sub
DrawBorder Index, 1
End Sub


Private Sub mnuExit_Click()
   Unload Me
   Set wisMain = Nothing
  ' End
End Sub

Private Sub picCanvas_KeyDown(KeyCode As Integer, Shift As Integer)

Dim CtrlDown
CtrlDown = (Shift And vbCtrlMask) > 0
Select Case KeyCode


End Select
SetActiveWindow (gWindowHandle)
End Sub

Private Sub picCanvas_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
If PrevButton <= 0 Then Exit Sub

DrawBorder PrevButton, 0
PrevButton = -1

End Sub

Private Sub picCanvas_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
If PrevButton = -1 Then Exit Sub
DrawBorder PrevButton, 1
PrevButton = -1
End Sub
Private Sub picCanvas_Resize()

'If Not resizeNow Then Exit Sub
'resizeNow = False
Dim i As Integer
' Set the position of image buttons.
For i = 1 To img.Count - 1
    With img(i)
        .Left = (picCanvas.ScaleWidth - .Width) / 2
    End With
    With lbl(i)
        .Left = (picCanvas.ScaleWidth - .Width) / 2
        .Height = 450
    End With
Next

End Sub

Private Sub picToolbar_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
Call picCanvas_MouseMove(Button, Shift, X, y)
End Sub

Private Sub picToolbar_Resize()
Dim i As Integer
Dim TopCount As Integer
Dim BottomCount As Integer
Dim Alignment As Integer
On Error Resume Next
' If the form is minimized, exit.
If Me.WindowState = vbMinimized Then Exit Sub


' Align all the sliding panels.
For i = 0 To pnlSlider.Count - 1
    ' Get the alignment property.
    With pnlSlider(i)
        Alignment = Val(ExtractToken(.Tag, "Alignment"))
        If Alignment = wis_TOP Then
            .Top = (.Index - 1) * .Height
            If i <> 0 Then TopCount = TopCount + 1
        Else
            .Top = picToolbar.ScaleHeight - _
                    (pnlSlider.Count - .Index) * .Height
            BottomCount = BottomCount + 1
        End If
        .Width = picToolbar.ScaleWidth
        .Left = 0
    End With
Next

' Set the width of canvas.
With picCanvas
    resizeNow = True
    .Width = picToolbar.ScaleWidth
End With

' Position of the scroll buttons...
Const SCROLLMARGIN = 100
cmdUP.Left = picToolbar.ScaleWidth - cmdUP.Width - SCROLLMARGIN
cmdUP.Top = SCROLLMARGIN + TopCount * pnlSlider(0).Height
With cmdDown
    .Left = cmdUP.Left
    .Top = picToolbar.ScaleHeight - .Height - SCROLLMARGIN _
            - BottomCount * pnlSlider(0).Height
End With

' Hide or display the scroll buttons.
ResetScrollButtons

End Sub

Private Sub pnlSlider_Click(Index As Integer)
Dim Alignment As Integer
Dim strTmp As String
Dim i As Integer

' Get the alignment property.
Alignment = Val(ExtractToken(pnlSlider(Index).Tag, "Alignment"))
If Alignment = wis_TOP Then
    ' Change the alignment property for all the panels
    ' that have index greater than the current index.
    For i = Index + 1 To pnlSlider.Count - 1
        With pnlSlider(i)
            If i > 1 Then .Tag = putToken(.Tag, "Alignment", wis_BOTTOM)
        End With
    Next
    picToolbar_Resize
    If Index > 1 Then
        ShowIcons (Index)
    Else
        ShowIcons (Index)
    End If

Else
    ' Change the alignment property for all the panels
    ' that have index less than the current index.
    For i = 1 To Index
        With pnlSlider(i)
            If i > 1 Then .Tag = putToken(.Tag, "Alignment", wis_TOP)
        End With
    Next
    picToolbar_Resize
    ShowIcons (Index)
End If

' Hide or display the scroll buttons.
ResetScrollButtons
End Sub

Private Sub sizeBar_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
m_MouseX = X
m_MouseY = y
With resizeGuide
    .Left = sizeBar.Left
    .Visible = True
End With
End Sub


Private Sub sizeBar_MouseMove(Button As Integer, Shift As Integer, X As Single, y As Single)
If Button <> vbLeftButton Then Exit Sub
resizeGuide.Left = resizeGuide.Left + X - m_MouseX
m_MouseX = X
End Sub
Private Sub sizeBar_MouseUp(Button As Integer, Shift As Integer, X As Single, y As Single)
On Error Resume Next
resizeGuide.Visible = False

' Check for min width of toolbar.
If resizeGuide.Left <= 2 * img(0).Width Then
    resizeGuide.Left = img(0).Left * 2
End If
sizeBar.Left = resizeGuide.Left
picToolbar.Width = sizeBar.Left '- 25
With picViewport
    .Width = Me.ScaleWidth - _
        picToolbar.Width - sizeBar.Width
    .Left = sizeBar.Left + sizeBar.Width
End With
End Sub

Public Property Get ButtonKey(Indx As Integer) As String
ButtonKey = ExtractToken(img(Indx).Tag, "Key")
End Property
Public Property Let ButtonKey(Indx As Integer, ByVal vNewValue As String)
img(Indx).Tag = putToken(img(Indx).Tag, "Key", vNewValue)
End Property

Private Sub wisAppObj_UpdateStatus(strMsg As String)
SetMessage strMsg
End Sub


Private Sub mnuFl_Click()
Dim SbClass As New clsSBAcc
SbClass.Show

End Sub


