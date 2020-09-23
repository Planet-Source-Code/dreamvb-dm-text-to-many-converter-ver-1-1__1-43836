VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmappconv 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Text to Exe Converter"
   ClientHeight    =   3165
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6345
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   6345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   5085
      TabIndex        =   10
      Top             =   2595
      Width           =   1110
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "Con&vert"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3720
      TabIndex        =   9
      Top             =   2595
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      Caption         =   "Please select your file below"
      Height          =   1710
      Left            =   135
      TabIndex        =   2
      Top             =   720
      Width           =   6105
      Begin VB.CheckBox chkdelete 
         Caption         =   "Delete this file after it's been converted."
         Height          =   225
         Left            =   1350
         TabIndex        =   8
         Top             =   765
         Width           =   4020
      End
      Begin MSComDlg.CommonDialog CDialog 
         Left            =   5355
         Top             =   795
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdinput 
         Caption         =   "...."
         Height          =   315
         Left            =   5400
         TabIndex        =   7
         Top             =   375
         Width           =   345
      End
      Begin VB.TextBox txtoutput 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1110
         Width           =   3930
      End
      Begin VB.TextBox txtinput 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1305
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   390
         Width           =   3930
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Output File:"
         Height          =   195
         Left            =   300
         TabIndex        =   4
         Top             =   1140
         Width           =   810
      End
      Begin VB.Label lblinp 
         AutoSize        =   -1  'True
         Caption         =   "Input File:"
         Height          =   195
         Left            =   285
         TabIndex        =   3
         Top             =   390
         Width           =   690
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   570
      Left            =   0
      ScaleHeight     =   570
      ScaleWidth      =   6345
      TabIndex        =   0
      Top             =   0
      Width           =   6345
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "frmappconv.frx":0000
         Top             =   15
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DM Text to Exe Converter DM Text to many Add-on"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   675
         TabIndex        =   1
         Top             =   105
         Width           =   5460
      End
   End
   Begin VB.Line lnbar 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   675
      Y1              =   570
      Y2              =   570
   End
End
Attribute VB_Name = "frmappconv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Unload frmappconv   ' Unload this form
    frmmain.Show
End Sub

Private Sub cmdConvert_Click()
Dim sBuffer As String, mCaption As String, ans As Integer

Dim nFile As Long

    If FindFile(Fixpath(App.Path) & "sfx\sfx.exe") = False Then
        ' Above line will check for the exe file that will be used for the text viewer
        MsgBox "Unable to locate data please make sure that the file has not been renamed or deleted by mistake.", vbCritical, "File not found"
        Exit Sub
    Else
        sBuffer = "DM:" & GetFileTitle(txtinput.Text) & "  " & "<!~~" & Encrypt(OpenData(txtinput.Text)) & "~~!>"
        'The top line will Open, Encrypt and add a header info were we need to extract latter

        FileCopy Fixpath(App.Path) & "sfx\sfx.exe", GetPath(txtoutput.Text) & GetFileTitle(txtinput.Text) & ".exe"
        ' Copy the new exe to its new location
        
        nFile = FreeFile ' Free pinter to file
        Open txtoutput.Text For Binary As #nFile ' Open the file in binary mode
            Put #nFile, LOF(nFile) + 1, sBuffer ' Add the new data to the end of the sfx.exe
        Close #nFile ' Close the file
        
        If chkdelete Then
            SetAttr txtinput.Text, vbNormal ' Remove all Attributes
            Kill txtinput.Text  ' Delete the file
        End If
        
        ans = MsgBox("The text file has been converted and saved to " _
        & vbNewLine & txtoutput.Text & vbNewLine & vbNewLine & "Do you want to view this file now?", vbYesNo Or vbQuestion)
        If ans = vbNo Then Exit Sub ' Don't do anything here
        WinExec txtoutput.Text, 10 ' Run the file
        
    End If
    
End Sub

Private Sub cmdinput_Click()
    With CDialog
        .DialogTitle = "Open Text Document" ' Dialog title
        .Filter = "Text Documents(*.txt)|*.txt|" ' Dialog filter
        .ShowOpen   ' Show the open dialog
        If Len(.FileName) <= 0 Then Exit Sub ' don't do anything here cancel button was pressed
        If Not UCase(Right(.FileName, 3)) = "TXT" Then ' Check to see if we have a vaild text file
            MsgBox "Please select a vaild text document", vbExclamation, "Invaild Document" ' Show user error message invaild filename
            Exit Sub ' Stop here
        Else
            txtinput.Text = .FileName ' Assign textbox with returned filename
            txtoutput.Text = Left(.FileName, Len(.FileName) - 3) + "exe" ' Assign text box the new filename
            cmdConvert.Enabled = True ' Enable the convert button
        End If
    End With
    
End Sub

Private Sub Form_Resize()
    lnbar.X2 = frmappconv.ScaleWidth ' Resize the line to the with of the form
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmappconv = Nothing ' Unload thr form from memory
    frmmain.Show ' show the main form
End Sub
