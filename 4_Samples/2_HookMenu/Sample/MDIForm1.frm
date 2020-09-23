VERSION 5.00
Object = "*\A..\HookMenu.vbp"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   3990
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7005
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin HookMenu.ctxHookMenu ctxHookMenu1 
      Left            =   1320
      Top             =   1440
      _extentx        =   900
      _extenty        =   900
      bmpcount        =   0
      font            =   "MDIForm1.frx":0000
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuWindows 
      Caption         =   "&Windows"
      WindowList      =   -1  'True
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuNew_Click()
    Dim f As New Form2
    f.Show
End Sub

