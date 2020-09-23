VERSION 5.00
Begin VB.Form frmabout 
   Caption         =   "About"
   ClientHeight    =   2925
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4275
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2925
   ScaleWidth      =   4275
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdok 
      Caption         =   "&OK"
      Height          =   360
      Left            =   3300
      TabIndex        =   7
      Top             =   2370
      Width           =   795
   End
   Begin Project1.Line3D Line3D1 
      Height          =   45
      Left            =   345
      TabIndex        =   3
      Top             =   1065
      Width           =   3765
      _ExtentX        =   6641
      _ExtentY        =   79
   End
   Begin Project1.Line3D Line3D2 
      Height          =   45
      Left            =   345
      TabIndex        =   4
      Top             =   1560
      Width           =   3765
      _ExtentX        =   6641
      _ExtentY        =   79
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Copyright 2003 © Ben G"
      Height          =   195
      Left            =   1110
      TabIndex        =   6
      Top             =   2025
      Width           =   1740
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Written and designed by Ben G"
      Height          =   195
      Left            =   900
      TabIndex        =   5
      Top             =   1725
      Width           =   2220
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "This program will allow a developer to convert text between 10 different formats."
      Height          =   540
      Left            =   405
      TabIndex        =   1
      Top             =   600
      Width           =   3615
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "This Program is Freeware"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   930
      TabIndex        =   2
      Top             =   1215
      Width           =   2205
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DM Text to many converter."
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   405
      TabIndex        =   0
      Top             =   105
      Width           =   3450
   End
End
Attribute VB_Name = "frmabout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdok_Click()
    Unload frmabout ' Unload the form
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmabout = Nothing ' Unload thr form form memory
End Sub

