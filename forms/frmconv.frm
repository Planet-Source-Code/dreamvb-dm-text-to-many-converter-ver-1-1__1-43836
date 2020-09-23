VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmconv 
   ClientHeight    =   3030
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6900
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   6900
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   510
      Left            =   1785
      TabIndex        =   4
      Top             =   1845
      Width           =   735
   End
   Begin MSComDlg.CommonDialog CDialog 
      Left            =   90
      Top             =   1260
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar sBar 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   3
      Top             =   2700
      Width           =   6900
      _ExtentX        =   12171
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11668
         EndProperty
      EndProperty
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
      Height          =   2220
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "frmconv.frx":0000
      Top             =   435
      Width           =   6810
   End
   Begin Project1.Line3D Line3D1 
      Height          =   45
      Left            =   -15
      TabIndex        =   1
      Top             =   375
      Width           =   690
      _ExtentX        =   1217
      _ExtentY        =   79
   End
   Begin MSComctlLib.Toolbar tBar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6900
      _ExtentX        =   12171
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "NEW"
            Object.ToolTipText     =   "New Blank Document"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "OPEN"
            Object.ToolTipText     =   "Open Text Resource"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "SAVE"
            Object.ToolTipText     =   "Save Converted Text"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "CUT"
            Object.ToolTipText     =   "Cut Text"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "COPY"
            Object.ToolTipText     =   "Copy Text"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "PASTE"
            Object.ToolTipText     =   "Paste Text"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "CONV"
            Object.ToolTipText     =   "Convert Now"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "CLOSE"
            Object.ToolTipText     =   "Close Workspace"
            ImageIndex      =   8
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   630
      Top             =   1185
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmconv.frx":0024
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmconv.frx":0376
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmconv.frx":06C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmconv.frx":0A1A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmconv.frx":0D6C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmconv.frx":10BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmconv.frx":1410
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmconv.frx":1762
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmconv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Enum MNU_COMMAND
    mCut = 1
    mCopy
    mPaste
End Enum
Sub SaveFile(sData As String)
    CDialog.DialogTitle = save_dlgTitle ' Assign the dialogs caption
    CDialog.Filter = save_dlgFilter ' Assign the dialog filter value
    CDialog.ShowSave ' Show the save dialog
    If Len(CDialog.FileName) = 0 Then Exit Sub ' Nothing for us to do so just exit the sub
    SaveData CDialog.FileName, sData ' Save the new file and the data
    
End Sub
Function OpenFile() As String
    CDialog.DialogTitle = "Open Text File" ' Assign the dialogs caption
    CDialog.Filter = "Text Document(*.txt)|*.txt|" ' Assign the dialog filter value
    CDialog.ShowOpen ' Show the open dialog
    If Len(CDialog.FileName) = 0 Then Exit Function ' Nothing for us to do so just exit the function
    OpenFile = OpenData(CDialog.FileName) 'Open the selected filename
    
End Function
Sub EditMenu(mTextBox As TextBox, mnuCommand As MNU_COMMAND)
    Select Case mnuCommand
        Case mCut ' Cut command was selected
            Clipboard.SetText mTextBox.Text ' Assign mTextBox value to the clipboard
            mTextBox.SelText = "" ' Clear the mTextBox seltext value
            mTextBox.SetFocus     ' Set focus on the textbox
        Case mCopy ' Copy command was selected
            Clipboard.Clear ' Clear what ever on the clipboard
            Clipboard.SetText mTextBox.SelText 'Assign the clipboard with the mTextBox value
            mTextBox.SetFocus   ' Set focus on the textbox
        Case mPaste ' Paste command was selected
            mTextBox.SelText = Clipboard.GetText(vbCFText)
            mTextBox.SetFocus ' Set focus on the textbox
    End Select

End Sub


Private Sub Form_Load()
On Error Resume Next
    ' The code will enable or disable the paste button depaending on the state of the clipboards data
    If Len(Trim(Clipboard.GetText)) = 0 Then
        tBar1.Buttons(7).Enabled = False
    Else
        tBar1.Buttons(7).Enabled = True
    End If
    
End Sub

Private Sub Form_Paint()
On Error Resume Next
    txtbase.SelStart = Len(txtbase.Text) - 1
    txtbase.SetFocus
    If Err Then Err.Clear
End Sub

Private Sub Form_Resize()
    Line3D1.Width = frmconv.Width
    txtbase.Width = frmconv.ScaleWidth - txtbase.Left
    txtbase.Height = frmconv.ScaleHeight - sBar.Height - txtbase.Top

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form_Caption = ""   ' Clear the forms caption
    save_dlgTitle = ""  ' Clear the dialogs title
    save_dlgFilter = "" ' Clear the dialogs filter
    lstKey = "" ' Clear the user selected choice
    Set frmconv = Nothing ' Unrelease the form from memory
    frmmain.Show ' Show the main form
    
End Sub

Private Sub tBar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim mHtmTitle As String

Dim ans As Integer

    Select Case Button.Key ' Index for the button selected
        '----------------------------------------------------------------------------------------------------------------------------------
        Case "NEW"
            If Len(txtbase.Text) > 0 Then
                ans = MsgBox("There is still text in the workspace." & vbNewLine & "Do you want to save this now", vbYesNo Or vbQuestion)
                If ans = vbNo Then
                    txtbase.Text = ""
                Else
                    SaveFile txtbase.Text
                    txtbase.Text = ""
                End If
            Else
                txtbase.Text = ""
            End If
        '----------------------------------------------------------------------------------------------------------------------------------
        Case "OPEN"
            If Len(txtbase.Text) > 0 Then
                ans = MsgBox("There is still text in the workspace." & vbNewLine & "Do you want to save this now.", vbYesNo Or vbQuestion)
                If ans = vbNo Then
                    txtbase.Text = ""
                    txtbase.Text = OpenFile
                Else
                    SaveFile txtbase.Text
                End If
            Else
                txtbase.Text = OpenFile
            End If
        '----------------------------------------------------------------------------------------------------------------------------------
        Case "SAVE"
            SaveFile txtbase.Text ' Save the contents of the textbox
        '----------------------------------------------------------------------------------------------------------------------------------
        Case "CUT"
            EditMenu txtbase, mCut  ' Call cut command
        '----------------------------------------------------------------------------------------------------------------------------------
        Case "COPY"
            EditMenu txtbase, mCopy ' Call the copy command
        '----------------------------------------------------------------------------------------------------------------------------------
        Case "PASTE"
            EditMenu txtbase, mPaste ' Call the paste command
        '----------------------------------------------------------------------------------------------------------------------------------
        Case "CLOSE"
            ans = MsgBox("Are you sure you want to close the workspace now.", vbYesNo Or vbQuestion, "Close Workspace")
            If ans = vbNo Then Exit Sub ' Exit the code block
            Unload frmconv ' Unload this form
        Case "CONV"
        '----------------------------------------------------------------------------------------------------------------------------------
            Select Case lstKey  ' Convert Options index
                Case "TXT_HTM"  ' Text to html convert
                    mHtmTitle = Trim(InputBox("Please enter in a title for your webpage", "Enter page title", "Your title goes here")) ' Show user inputbox
                    If (mHtmTitle = "" Or Len(mHtmTitle) > 0) Then ' did the user enter anything
                        If Not Right(txtbase.Text, 2) = vbCrLf Then txtbase.Text = txtbase.Text & vbNewLine ' Add a newline if needed
                        txtbase.Text = Text2Html(txtbase.Text, mHtmTitle)
                    End If
                Case "TXT_JAVA" ' Text to JavaScript convert
                    txtbase.Text = Text2JScript(txtbase.Text)
                Case "TXT_ASP"  ' Text to ASP (Active server page) convert
                    txtbase.Text = Text2ASP(txtbase.Text)
                Case "TXT_RTF"  ' Text to Rich Text Format convert
                    txtbase.Text = Text2RTF(txtbase.Text)
                Case "TXT_PHP"  ' Text to PHP Document convert
                    txtbase.Text = Text2PHP(txtbase.Text)
                Case "TXT_LINUX2PC" ' Linux to PC Text convert
                    txtbase.Text = FixLinuxTXT(txtbase.Text)
                Case "TXT_MAC2PC"   ' Mac to PC Text Convert
                    txtbase.Text = FixMacText(txtbase.Text)
                Case "TXT_PERL" ' Text to perl convert
                    txtbase.Text = Text2Perl(txtbase.Text)
                Case "TXT_VBS" ' Text to VBScript convert
                    txtbase.Text = Text2VBSscript(txtbase.Text)
            End Select
    End Select
    
End Sub

Private Sub txtbase_KeyDown(KeyCode As Integer, Shift As Integer)
    ' The code below lets you use CTRL key and A to select all the text in the textbox
    If (KeyCode = vbKeyA And Shift) Then
        txtbase.SelStart = 0
        txtbase.SelLength = Len(txtbase.Text)
        txtbase.SetFocus
    End If
    
End Sub

Private Sub txtbase_KeyPress(KeyAscii As Integer)
    If KeyAscii = 1 Then KeyAscii = 0 ' This stops the beep
    
End Sub

Private Sub txtbase_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Len(txtbase.SelText) = 0 Then
        tBar1.Buttons(5).Enabled = False
        tBar1.Buttons(6).Enabled = False
    Else
        tBar1.Buttons(5).Enabled = True
        tBar1.Buttons(6).Enabled = True
    End If

End Sub
