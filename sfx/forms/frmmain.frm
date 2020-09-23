VERSION 5.00
Begin VB.Form frmmain 
   ClientHeight    =   5430
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7740
   LinkTopic       =   "Form1"
   ScaleHeight     =   5430
   ScaleWidth      =   7740
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdclose 
      Caption         =   "&Close"
      Height          =   780
      Left            =   4800
      Picture         =   "frmmain.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   45
      Width           =   1065
   End
   Begin VB.CommandButton cmdfont 
      Caption         =   "&Font layout"
      Height          =   780
      Left            =   3615
      Picture         =   "frmmain.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   45
      Width           =   1065
   End
   Begin VB.CommandButton cmddown 
      Caption         =   "&Decrease"
      Height          =   780
      Left            =   2430
      Picture         =   "frmmain.frx":1994
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   45
      Width           =   1065
   End
   Begin VB.CommandButton cmdup 
      Caption         =   "&Increase"
      Height          =   780
      Left            =   1230
      Picture         =   "frmmain.frx":265E
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   45
      Width           =   1065
   End
   Begin Project1.Line3D Line3D1 
      Height          =   45
      Left            =   30
      TabIndex        =   2
      Top             =   885
      Width           =   6540
      _ExtentX        =   11536
      _ExtentY        =   79
   End
   Begin VB.TextBox txtbase 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1170
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   975
      Width           =   2220
   End
   Begin VB.CommandButton cmdcopy 
      Caption         =   "&Copy Text"
      Enabled         =   0   'False
      Height          =   780
      Left            =   30
      Picture         =   "frmmain.frx":3328
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   45
      Width           =   1065
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub EnableCopy()
    If Len(txtbase.SelText) > 0 Then ' Check to see if the textbox seltext has anything in it
        cmdcopy.Enabled = True ' Enable the copy button
    Else
        cmdcopy.Enabled = False ' Disable the copy button
    End If
End Sub

Private Sub cmdclose_Click()
    Unload frmmain  ' Unload the form
End Sub

Private Sub cmdcopy_Click()
    Clipboard.Clear ' Clear the contents of the clipbaord
    Clipboard.SetText txtbase.SelText ' Assign the clipboard textbos seltext
    MsgBox "The text has now been copied to the clipboard.", vbInformation, "Copy Text" ' A messagebox I am getting board of commenting now.
End Sub

Private Sub cmddown_Click()
    On Error Resume Next
    txtbase.FontSize = txtbase.FontSize - 2 ' Decrease the fontsize by 2
End Sub

Private Sub cmdfont_Click()
    frmfont.Show vbModal, frmmain ' Show the fontstyle form
End Sub

Private Sub cmdup_Click()
    On Error Resume Next
    txtbase.FontSize = txtbase.FontSize + 2 ' Increase the fontsize by 2
End Sub

Private Sub Form_Load()
Dim tFile As Long, iPos As Long, lPos As Long
Dim sBuffer As String, sData As String

    tFile = FreeFile ' Pointer to freefile
    Open Fixpath(App.Path) & App.EXEName & ".exe" For Binary As #tFile
        sBuffer = Space(LOF(tFile)) ' Create a buffer
        Get #tFile, , sBuffer ' Load file cotents
    Close #tFile ' close file
    
    iPos = InStr(1, sBuffer, "DM:", vbTextCompare)
    If iPos = 0 Then
        MsgBox "Data file not found", vbInformation, "No data Found"
        Unload frmmain
    Else
       lPos = InStr(iPos, sBuffer, " ", vbTextCompare)
       frmmain.Caption = Mid(sBuffer, iPos + 3, lPos - iPos - 3) ' Set the forms caption
       iPos = 0 ' Reset var
       sBuffer = Trim$(sBuffer)
       lPos = 0 ' Reset var
       
       iPos = InStr(1, sBuffer, "<!~~", vbTextCompare) ' Start of encrypted text
       lPos = InStr(iPos, sBuffer, "~~!>", vbTextCompare) ' end of encrypted text
       
       If Not (iPos > 0 Or lPos > 0) Then
            MsgBox "Data file not found", vbInformation, "No Data Found"
            Unload frmmain ' Unload the form
       Else
            sData = Mid(sBuffer, iPos + 4, lPos - iPos - 4)
            txtbase.Text = Encrypt(sData)
            sBuffer = "" ' Clear buffer
            sData = ""   ' Clear buffer
            iPos = 0     ' Reset var
            lPos = 0     ' Reset var
       End If
    End If

End Sub

Private Sub Form_Resize()
On Error Resume Next
    Line3D1.Width = frmmain.ScaleWidth - Line3D1.Left ' Resize 3Dline to forms with
    txtbase.Width = frmmain.ScaleWidth - txtbase.Left ' Resize textbox to forms with
    txtbase.Height = frmmain.ScaleHeight - txtbase.Top ' Resize textbox to forms height
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    txtbase.Text = "" ' Clear textbox contents
    Set frmmain = Nothing ' Release the form from memory
End Sub

Private Sub txtbase_KeyPress(KeyAscii As Integer)
    If KeyAscii = 1 Then KeyAscii = 0
End Sub

Private Sub txtbase_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyA And Shift) Then
        txtbase.SelStart = 0
        txtbase.SelLength = Len(txtbase.Text)
        txtbase.SetFocus
        EnableCopy
    End If
    
End Sub

Private Sub txtbase_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    EnableCopy
End Sub
