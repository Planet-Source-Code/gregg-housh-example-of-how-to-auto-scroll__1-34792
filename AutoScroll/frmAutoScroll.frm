VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmAutoScroll 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Auto Scroll Demo"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   240
      Top             =   2640
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   4260
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmAutoScroll.frx":0000
   End
End
Attribute VB_Name = "frmAutoScroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_USER = &H400
Private Const EM_GETSCROLLPOS = (WM_USER + 221)
Private Const EM_SETSCROLLPOS = (WM_USER + 222)
Private Const EM_LINEFROMCHAR = &HC9
Private Const EM_CHARFROMPOS = &HD7
Private Const EM_GETLINECOUNT = &HBA
Private Type POINTL
x As Long
y As Long
End Type
Private Type RECT
Left As Long
Top As Long
Right As Long
Bottom As Long
End Type
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

Private Sub Timer1_Timer()
    Dim lPos As Long
    Dim pt As POINTL
    Dim r As RECT
    Dim lCount As Long
    Dim sTemp As String
    Dim l As Long

    sTemp = String$(Rnd() * 5 + 5, Rnd() * 26 + Asc("A")) & vbCrLf
    With RichTextBox1
        lCount = SendMessage(.hwnd, EM_GETLINECOUNT, 0, ByVal 0&) - 1
        GetClientRect .hwnd, r
        pt.x = r.Left + 1
        pt.y = r.Bottom - 1
        lPos = SendMessage(.hwnd, EM_CHARFROMPOS, 0, pt)
        lPos = SendMessage(.hwnd, EM_LINEFROMCHAR, lPos, ByVal 0&)
        If lPos < lCount Then 'do not scroll
            l = SendMessage(.hwnd, EM_GETSCROLLPOS, 0, pt)
            .Text = .Text & sTemp
            l = SendMessage(.hwnd, EM_SETSCROLLPOS, 0, pt)
        Else
            .Text = .Text & sTemp
            .SelStart = Len(.Text)
        End If
    End With
End Sub

