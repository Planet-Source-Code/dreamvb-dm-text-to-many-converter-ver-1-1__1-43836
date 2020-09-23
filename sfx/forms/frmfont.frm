VERSION 5.00
Begin VB.Form frmfont 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Font Layout"
   ClientHeight    =   3885
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7395
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3885
   ScaleWidth      =   7395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdcancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   6105
      TabIndex        =   15
      Top             =   3255
      Width           =   1140
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "&OK"
      Height          =   375
      Left            =   4965
      TabIndex        =   14
      Top             =   3255
      Width           =   870
   End
   Begin VB.Frame Frame1 
      Caption         =   "Font Properties"
      Height          =   3000
      Left            =   75
      TabIndex        =   0
      Top             =   135
      Width           =   7245
      Begin VB.PictureBox picsample 
         BackColor       =   &H00FFFFFF&
         Height          =   1005
         Left            =   3675
         ScaleHeight     =   945
         ScaleWidth      =   3000
         TabIndex        =   11
         Top             =   1785
         Width           =   3060
         Begin VB.PictureBox Picture2 
            Height          =   990
            Left            =   -30
            ScaleHeight     =   930
            ScaleWidth      =   270
            TabIndex        =   13
            Top             =   -30
            Width           =   330
         End
         Begin VB.Label lblsample 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sample"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   480
            TabIndex        =   12
            Top             =   345
            Width           =   690
         End
      End
      Begin VB.PictureBox picbkcol 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6780
         ScaleHeight     =   225
         ScaleWidth      =   300
         TabIndex        =   10
         Top             =   1365
         Width           =   330
      End
      Begin VB.PictureBox picforecol 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   6780
         ScaleHeight     =   225
         ScaleWidth      =   300
         TabIndex        =   9
         Top             =   660
         Width           =   330
      End
      Begin VB.PictureBox picbackcolor 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   3690
         Picture         =   "frmfont.frx":0000
         ScaleHeight     =   255
         ScaleWidth      =   3000
         TabIndex        =   6
         Top             =   1365
         Width           =   3000
      End
      Begin VB.PictureBox picforecolor 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   3690
         Picture         =   "frmfont.frx":281A
         ScaleHeight     =   255
         ScaleWidth      =   3000
         TabIndex        =   5
         Top             =   660
         Width           =   3000
      End
      Begin VB.ListBox lstfsize 
         Height          =   1815
         Left            =   2310
         TabIndex        =   4
         Top             =   570
         Width           =   1005
      End
      Begin VB.ListBox lstfont 
         Height          =   1815
         Left            =   270
         TabIndex        =   2
         Top             =   570
         Width           =   1995
      End
      Begin VB.Label lblbk 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Text Back Colour:"
         Height          =   195
         Left            =   3705
         TabIndex        =   8
         Top             =   1125
         Width           =   1275
      End
      Begin VB.Label lblfc 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Text Fore Colour"
         Height          =   195
         Left            =   3690
         TabIndex        =   7
         Top             =   390
         Width           =   1170
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   3495
         X2              =   3495
         Y1              =   540
         Y2              =   2445
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   0
         X1              =   3480
         X2              =   3480
         Y1              =   540
         Y2              =   2445
      End
      Begin VB.Label lblfontsize 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Font size:"
         Height          =   195
         Left            =   2475
         TabIndex        =   3
         Top             =   315
         Width           =   675
      End
      Begin VB.Label lblfont 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Font:"
         Height          =   195
         Left            =   270
         TabIndex        =   1
         Top             =   315
         Width           =   360
      End
   End
End
Attribute VB_Name = "frmfont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcancel_Click()
    Unload frmfont  ' Unload the form
End Sub

Private Sub cmdok_Click()
    frmmain.txtbase.ForeColor = picforecol.BackColor
    frmmain.txtbase.BackColor = picbkcol.BackColor
    frmmain.txtbase.FontName = lstfont.Text
    frmmain.txtbase.FontSize = Val(lstfsize.Text)
    frmmain.txtbase.Font.Charset = 0
    frmmain.txtbase.Font.Weight = 0
    cmdcancel_Click
    
    
End Sub

Private Sub Form_Load()
Dim i As Integer
    For i = 0 To Screen.FontCount - 1
        lstfont.AddItem Screen.Fonts(i)
    Next
    lstfont.ListIndex = 0
    i = 0
    
    lstfsize.AddItem "8"
    lstfsize.AddItem "9"
    lstfsize.AddItem "10"
    lstfsize.AddItem "12"
    lstfsize.AddItem "14"
    lstfsize.AddItem "16"
    lstfsize.AddItem "18"
    lstfsize.AddItem "20"
    lstfsize.AddItem "21"
    lstfsize.AddItem "22"
    lstfsize.AddItem "24"
    lstfsize.AddItem "26"
    lstfsize.AddItem "28"
    lstfsize.AddItem "30"
    lstfsize.AddItem "32"
    lstfsize.AddItem "34"
    lstfsize.AddItem "36"
    lstfsize.AddItem "38"
    lstfsize.AddItem "42"
    lstfsize.AddItem "44"
    lstfsize.AddItem "48"
    lstfsize.AddItem "72"
    lstfsize.ListIndex = 2

    
End Sub

Private Sub lstfont_Click()
    lblsample.FontName = lstfont.Text
    lblsample.Font.Charset = 0
    lblsample.Font.Weight = 0
End Sub

Private Sub lstfsize_Click()
    lblsample.FontSize = Val(lstfsize.Text)
    lblsample.Font.Charset = 0
    lblsample.Font.Weight = 0
End Sub



Private Sub picbackcolor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    If Button = 1 Then
         picbkcol.BackColor = picbackcolor.Point(X, Y)
         picsample.BackColor = picbkcol.BackColor
    End If
    
End Sub

Private Sub picforecolor_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    If Button = 1 Then
         picforecol.BackColor = picforecolor.Point(X, Y)
         lblsample.ForeColor = picforecol.BackColor
    End If
    
End Sub
