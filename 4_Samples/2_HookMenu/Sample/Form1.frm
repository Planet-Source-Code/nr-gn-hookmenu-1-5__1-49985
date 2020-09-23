VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "*\A..\HookMenu.vbp"
Begin VB.Form Form1 
   BackColor       =   &H80000018&
   Caption         =   "Form1"
   ClientHeight    =   5775
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9075
   LinkTopic       =   "Form1"
   ScaleHeight     =   5775
   ScaleWidth      =   9075
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   1320
      List            =   "Form1.frx":0025
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Test Right To Left"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   2415
   End
   Begin HookMenu.ctxHookMenu ctxHookMenu1 
      Left            =   2040
      Top             =   1560
      _ExtentX        =   900
      _ExtentY        =   900
      BmpCount        =   13
      Bmp:1           =   "Form1.frx":0051
      Key:1           =   "#mnuFile:0"
      Bmp:2           =   "Form1.frx":0479
      Key:2           =   "#mnuFile:1"
      Bmp:3           =   "Form1.frx":08A1
      Key:3           =   "#mnuOpen:0"
      Bmp:4           =   "Form1.frx":0CC9
      Key:4           =   "#mnuOpen:1"
      Bmp:5           =   "Form1.frx":10F1
      Key:5           =   "#mnuEdit:2"
      Bmp:6           =   "Form1.frx":1519
      Key:6           =   "#mnuEdit:3"
      Bmp:7           =   "Form1.frx":1941
      Key:7           =   "#mnuEdit:4"
      Bmp:8           =   "Form1.frx":1D69
      Key:8           =   "#mnuEdit:0"
      Bmp:9           =   "Form1.frx":2191
      Key:9           =   "#mnuEdit:6"
      Bmp:10          =   "Form1.frx":25B9
      Key:10          =   "#mnuStyleID:0"
      Bmp:11          =   "Form1.frx":29E1
      Key:11          =   "#mnuStyleID:1"
      Bmp:12          =   "Form1.frx":2E09
      Key:12          =   "#mnuStyleID:2"
      Bmp:13          =   "Form1.frx":3231
      Key:13          =   "#mnuStyleID:3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   5520
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox Text1 
      Height          =   1104
      Left            =   336
      TabIndex        =   0
      Text            =   "There are NO MORE issues with TextBox context menus :-))"
      Top             =   2520
      Width           =   4968
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4680
      Top             =   1365
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3659
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":376B
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":387D
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":39E9
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3AFB
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3C0D
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3D1F
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3E31
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3F43
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4055
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4167
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4279
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":438B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   1
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   4
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Test"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Another test"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Cancel"
               EndProperty
            EndProperty
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Auto columns"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   5940
      Picture         =   "Form1.frx":473F
      Top             =   600
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&File"
      Index           =   0
      Begin VB.Menu mnuFile 
         Caption         =   "&New"
         Checked         =   -1  'True
         Index           =   0
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Open"
         Index           =   1
         Begin VB.Menu mnuOpen 
            Caption         =   "&Mail"
            Index           =   0
            Shortcut        =   ^M
         End
         Begin VB.Menu mnuOpen 
            Caption         =   "&Note"
            Index           =   1
         End
         Begin VB.Menu mnuOpen 
            Caption         =   "Memo"
            Index           =   2
         End
         Begin VB.Menu mnuOpen 
            Caption         =   "Appointment"
            Index           =   3
         End
         Begin VB.Menu mnuOpen 
            Caption         =   "00"
            Index           =   4
            Begin VB.Menu mnuOpen00 
               Caption         =   "00-00"
               Index           =   0
            End
            Begin VB.Menu mnuOpen00 
               Caption         =   "00-01"
               Index           =   1
            End
         End
         Begin VB.Menu mnuOpen 
            Caption         =   "99"
            Index           =   5
         End
         Begin VB.Menu mnuOpen 
            Caption         =   "88"
            Index           =   6
         End
         Begin VB.Menu mnuOpen 
            Caption         =   "77"
            Index           =   7
         End
         Begin VB.Menu mnuOpen 
            Caption         =   "66"
            Index           =   8
         End
         Begin VB.Menu mnuOpen 
            Caption         =   "55"
            Index           =   9
         End
         Begin VB.Menu mnuOpen 
            Caption         =   "44"
            Index           =   10
         End
         Begin VB.Menu mnuOpen 
            Caption         =   "33"
            Index           =   11
         End
         Begin VB.Menu mnuOpen 
            Caption         =   "22"
            Index           =   12
         End
         Begin VB.Menu mnuOpen 
            Caption         =   "11"
            Index           =   13
         End
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Close"
         Checked         =   -1  'True
         Enabled         =   0   'False
         Index           =   2
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Print..."
         Index           =   4
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Print Preview"
         Checked         =   -1  'True
         Index           =   5
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuFile 
         Caption         =   "Exit"
         Index           =   7
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&Edit"
      Index           =   1
      Begin VB.Menu mnuEdit 
         Caption         =   "Undo"
         Index           =   0
         Begin VB.Menu mnuUndo 
            Caption         =   "1111"
            Index           =   0
         End
         Begin VB.Menu mnuUndo 
            Caption         =   "2222"
            Index           =   1
         End
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Cut"
         Index           =   2
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Copy"
         Index           =   3
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Paste"
         Enabled         =   0   'False
         Index           =   4
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "-"
         Index           =   5
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "Add menu"
         Index           =   6
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "-"
         Index           =   7
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&Icon size"
      Index           =   2
      Begin VB.Menu mnuSize 
         Caption         =   "16x16 px"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuSize 
         Caption         =   "20x20 px"
         Index           =   1
      End
      Begin VB.Menu mnuSize 
         Caption         =   "24x24 px"
         Index           =   2
      End
      Begin VB.Menu mnuSize 
         Caption         =   "28x28 px"
         Index           =   3
      End
      Begin VB.Menu mnuSize 
         Caption         =   "32x32 px"
         Index           =   4
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "popup"
      Index           =   3
      Visible         =   0   'False
      Begin VB.Menu mnuPopup 
         Caption         =   "Properties"
         Index           =   0
      End
      Begin VB.Menu mnuPopup 
         Caption         =   "New"
         Index           =   1
         Begin VB.Menu mnuPopupOpen 
            Caption         =   "Mail"
            Index           =   0
         End
         Begin VB.Menu mnuPopupOpen 
            Caption         =   "Appointement"
            Index           =   1
         End
      End
      Begin VB.Menu mnuPopup 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuPopup 
         Caption         =   "Cancel"
         Index           =   3
      End
   End
   Begin VB.Menu mnuTest 
      Caption         =   "&Test && Help"
   End
   Begin VB.Menu mnuStyle 
      Caption         =   "Style"
      Begin VB.Menu mnuStyleID 
         Caption         =   "Normal"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuStyleID 
         Caption         =   "Office 2003"
         Index           =   1
      End
      Begin VB.Menu mnuStyleID 
         Caption         =   "Green Ivy"
         Index           =   2
      End
      Begin VB.Menu mnuStyleID 
         Caption         =   "Ice Grey"
         Index           =   3
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Enum UcsFileMenu
    ucsFileNew = 0
    ucsFileSave = 2
    ucsFilePrintPreview = 5
    ucsFileExit = 7
    ucsEditUndo = 0
    ucsEditCut = 2
    ucsEditAddMenu = 6
    ucsEditSep = 7
    ucsMainPopup = 3
End Enum


Private Sub Combo1_Change()
    Me.ctxHookMenu1.AutoColumn = Me.Combo1.ListIndex
End Sub

Private Sub Combo1_Click()
Me.ctxHookMenu1.AutoColumn = Me.Combo1.ListIndex
End Sub

Private Sub Command1_Click()
If Command1.Caption = "Test Right To Left" Then
    Me.ctxHookMenu1.RightToLeft = True
    Command1.Caption = "Normal"
Else
    Command1.Caption = "Test Right To Left"
    Me.ctxHookMenu1.RightToLeft = False
End If
    

End Sub

Private Sub ctxHookMenu1_CustomDrawItemFont(Font As stdole.StdFont, Caption As String, ForeColour As stdole.OLE_COLOR)

    If Caption = "Normal" Then
        Font.Bold = True
        Font.Underline = True
        ForeColour = vbRed
    ElseIf Caption = "Exit" Then
        Font.Bold = True
        Font.Size = Font.Size + 2
        Font.Underline = True
        Font.Italic = True
        ForeColour = vbRed
    ElseIf Caption = "Office 2003" Then
        Font.Bold = False
        Font.Italic = True
        ForeColour = vbBlue
    ElseIf Caption = "Properties" Then
        Font.Bold = True
        Font.Italic = True
        ForeColour = &H40C0&
    Else
        Font.Bold = False
        Font.Underline = False
        ForeColour = vbBlack
    End If

End Sub


Private Sub ctxHookMenu1_CustomDrawItemHoverFont(SelectedFont As stdole.StdFont, Caption As String, SelectedForeColour As stdole.OLE_COLOR, SelectedBackColour As stdole.OLE_COLOR, SelectedBorderColour As stdole.OLE_COLOR)
    If Caption = "Normal" Then
        SelectedFont.Bold = True
        SelectedFont.Italic = True
        SelectedForeColour = vbGreen
        SelectedBackColour = vbYellow
        SelectedBorderColour = vbGreen
    End If
End Sub

Private Sub ctxHookMenu1_Highlight(strMenuCaption As String)
    
    Select Case strMenuCaption
        Case "Green Ivy"
        Me.StatusBar1.SimpleText = "Changes The Menu Style To Ivy Green"
        Case "Normal"
        Me.StatusBar1.SimpleText = "Changes The Menu Style To The Default XP Style"
        Case "Exit"
        Me.StatusBar1.SimpleText = "Exit The Application"
        Case "Print Preview"
        Me.StatusBar1.SimpleText = "Print Preview Selected Page...."

        Case Else
        Me.StatusBar1.SimpleText = strMenuCaption
    End Select
End Sub

Private Sub Form_Activate()
    Call ctxHookMenu1.SetBitmap(mnuFile(ucsFileNew), Image1.Picture, &HC0C0C0)
    Me.Combo1 = "Not Set"
    
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        PopupMenu mnuMain(ucsMainPopup), , , , mnuPopup(0)
    End If
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_MouseUp Button, Shift, X, Y
End Sub

Private Sub mnuEdit_Click(Index As Integer)
    Select Case Index
    Case ucsEditUndo
        Call ctxHookMenu1.SetBitmap(mnuFile(ucsFileSave), Image1.Picture, &HC0C0C0)
    Case ucsEditCut
        mnuFile(ucsFileNew).Caption = mnuFile(ucsFileNew).Caption & "1"
    Case ucsEditAddMenu
        mnuEdit(ucsEditSep).Visible = True
        Load mnuEdit(mnuEdit.Count)
        mnuEdit(mnuEdit.UBound).Caption = "Test - " & Timer
    End Select
End Sub

Private Sub mnuFile_Click(Index As Integer)
    Select Case Index
    Case ucsFileNew
        Dim f As New Form1
        f.Show
        mnuFile(ucsFileNew).Checked = Not mnuFile(ucsFileNew).Checked
    Case ucsFileExit
        Unload Me
    Case ucsFilePrintPreview
        mnuFile(ucsFilePrintPreview).Checked = Not mnuFile(ucsFilePrintPreview).Checked
    End Select
End Sub

Private Sub mnuOpen_Click(Index As Integer)
    If Index < 4 Then
        MDIForm1.Show
    End If
End Sub

Private Sub mnuSize_Click(Index As Integer)
    Dim lI As Long
    ctxHookMenu1.BitmapSize = 16 + Index * 4
    For lI = mnuSize.LBound To mnuSize.UBound
        mnuSize(lI).Checked = Index = lI
    Next
End Sub

Private Sub mnuStyleID_Click(Index As Integer)
Dim i As Integer


    If Me.ctxHookMenu1.AutoColumn > 0 Then
        Me.ctxHookMenu1.AutoColumn = 0
    Else
        Me.ctxHookMenu1.AutoColumn = 4
    End If
    

    For i = 0 To 3
        mnuStyleID(i).Checked = False
    Next
    
    mnuStyleID(Index).Checked = True

    '-- Change The Style
    Select Case Index
        Case Is = 0
            Me.ctxHookMenu1.DrawStyle = MS_Default
            Me.ctxHookMenu1.UseSystemFont = True
            Me.ctxHookMenu1.DisplayShadow = True
            Me.ctxHookMenu1.MenuDrawStyle = DS_XP
        Case Is = 1
            Me.ctxHookMenu1.UseSystemFont = False
            Me.ctxHookMenu1.MenuDrawStyle = DS_XP
            Me.ctxHookMenu1.DrawStyle = MS_Custom
            Me.ctxHookMenu1.DisplayShadow = True
            Me.ctxHookMenu1.SetCustomAttributes 16761765, 16769990, 13040639, &H80FF&, vbBlue, &H800000, _
                vbBlue, 8108783, vbBlue, vbWhite, 16761765, vbBlue, &H4000&, True, False
        Case Is = 2
            Me.ctxHookMenu1.MenuDrawStyle = DS_NORMAL
            Me.ctxHookMenu1.UseSystemFont = True
            Me.ctxHookMenu1.DrawStyle = MS_Custom
            Me.ctxHookMenu1.DisplayShadow = False
            Me.ctxHookMenu1.SetCustomAttributes vbYellow, vbRed, &HC0FFC0, &HC0FFC0, vbGreen, &H8000&, _
                vbYellow, &HC000&, &H80FF80, vbWhite, vbGreen, &H8000&, &H4000&, True, True
        Case Is = 3
            Me.ctxHookMenu1.MenuDrawStyle = DS_NORMAL
            Me.ctxHookMenu1.DrawStyle = MS_Custom
            Me.ctxHookMenu1.UseSystemFont = False
            Me.ctxHookMenu1.DisplayShadow = True
            Me.ctxHookMenu1.SetCustomAttributes vbButtonFace, &HE0E0E0, &HE0E0E0, &HE0E0E0, vbButtonShadow, vbApplicationWorkspace, _
                 &HC0C0C0, vbButtonFace, vbApplicationWorkspace, vbWhite, &HC0C0C0, vbBlack, vbBlack, True, True

        End Select
    

End Sub
