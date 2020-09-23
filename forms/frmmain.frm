VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmmain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DM Text to many converter"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   6045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdBut 
      Caption         =   "E&xit"
      Height          =   375
      Index           =   2
      Left            =   4500
      TabIndex        =   6
      Top             =   2310
      Width           =   1335
   End
   Begin VB.CommandButton CmdBut 
      Caption         =   "&About"
      Height          =   375
      Index           =   1
      Left            =   4500
      TabIndex        =   5
      Top             =   1620
      Width           =   1335
   End
   Begin VB.CommandButton CmdBut 
      Caption         =   "&Create New"
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   4500
      TabIndex        =   4
      Top             =   1005
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   780
      Left            =   0
      ScaleHeight     =   780
      ScaleWidth      =   6045
      TabIndex        =   2
      Top             =   0
      Width           =   6045
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DM Text to many converter"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   1755
         TabIndex        =   3
         Top             =   75
         Width           =   3375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Personal Ver 1.2"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   4095
         TabIndex        =   7
         Top             =   465
         Width           =   1605
      End
      Begin VB.Image Image1 
         Height          =   630
         Left            =   75
         Picture         =   "frmmain.frx":0000
         Top             =   60
         Width           =   1605
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   150
      Top             =   2940
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":356A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":3E44
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":471E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":4FF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":58D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":65AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":7286
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":7F60
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":8C3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":9914
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lstV 
      Height          =   3780
      Left            =   0
      TabIndex        =   1
      Top             =   810
      Width           =   4260
      _ExtentX        =   7514
      _ExtentY        =   6668
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDragMode     =   1
      HoverSelection  =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      OLEDragMode     =   1
      NumItems        =   0
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   4650
      Width           =   6045
      _ExtentX        =   10663
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10610
         EndProperty
      EndProperty
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      Index           =   1
      X1              =   4335
      X2              =   4335
      Y1              =   840
      Y2              =   4560
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      Index           =   0
      X1              =   4320
      X2              =   4320
      Y1              =   840
      Y2              =   4560
   End
   Begin VB.Line lntop 
      BorderColor     =   &H00808080&
      X1              =   0
      X2              =   1125
      Y1              =   780
      Y2              =   780
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub ShowForm(frm As Form)
    Select Case lstKey
        Case "TXT_HTM" ' HTML option was selected
            Form_Caption = "Text to Html Converter" ' Update the forms caption
            save_dlgTitle = "Save HTM File" ' Update the dialogs caption
            save_dlgFilter = "HTM Document(*.htm)|*.htm|" ' Update the dialogs filter value
        Case "TXT_JAVA" ' JavaScript option was selected
            Form_Caption = "Text to Java Script Converter" ' Update the forms caption
            save_dlgTitle = "Save JavaScript File" ' Update the dialogs caption
            save_dlgFilter = "JavaScript(*.js)|*.js|" ' Update the dialogs filter value
        Case "TXT_ASP"
            Form_Caption = "Text to Active Server Page" ' Update the forms caption
            save_dlgTitle = "Save ASP File" ' Update the dialogs caption
            save_dlgFilter = "Active Server pages(*.asp)|*.asp|" ' Update the dialogs filter value
        Case "TXT_RTF"
            Form_Caption = "Text to Rich Text Format" ' Update the forms caption
            save_dlgTitle = "Save Rich Text File" ' Update the dialogs caption
            save_dlgFilter = "Rich Text Format(*.rtf)|*.rtf|" ' Update the dialogs filter value
        Case "TXT_PHP"
            Form_Caption = "Text to PHP" ' Update the forms caption
            save_dlgTitle = "Save PHP File" ' Update the dialogs caption
            save_dlgFilter = "PHP Document(*.php)|*.php|" ' Update the dialogs filter value
        Case "TXT_LINUX2PC"
            Form_Caption = "Linux Text to PC Text" ' Update the forms caption
            save_dlgTitle = "Save Text Document" ' Update the dialogs caption
            save_dlgFilter = "Linux Text Document(*.txt)|*.txt|" ' Update the dialogs filter value
        Case "TXT_MAC2PC"
            Form_Caption = "Mac Text to PC Text" ' Update the forms caption
            save_dlgTitle = "Save Text Document" ' Update the dialogs caption
            save_dlgFilter = "Mac Text Document(*.txt)|*.txt|" ' Update the dialogs filter value
        Case "TXT_PERL"
            Form_Caption = "Text to Perl" ' Update the forms caption
            save_dlgTitle = "Save Perl Document" ' Update the dialogs caption
            save_dlgFilter = "Perl Document(*.pl)|*.pl|" ' Update the dialogs filter value
        Case "TXT_VBS"
            Form_Caption = "Text to VBScript" ' Update the forms caption
            save_dlgTitle = "Save VBScript Document" ' Update the dialogs caption
            save_dlgFilter = "VBScript Document(*.vbs)|*.vbs|" ' Update the dialogs filter value

    End Select
    
    frmmain.Hide ' Hide the main form
    frm.Caption = Form_Caption ' Update the forms caption
    frm.Show ' Show the convert form
    
End Sub
Private Sub CmdBut_Click(Index As Integer)
Dim ans As Integer

    Select Case Index
        Case 0 ' Create button was pressed
            Select Case lstKey
                Case "TXT_HTM", "TXT_JAVA", "TXT_ASP", _
                "TXT_RTF", "TXT_PHP", "TXT_LINUX2PC", _
                "TXT_MAC2PC", "TXT_PERL", "TXT_VBS"
                ShowForm frmconv ' Show the Converter form
            Case "TXT_EXE"
                frmmain.Hide    ' Hide this form
                frmappconv.Show ' Show the text to exe form
            End Select
        Case 1 ' About button was pressed
            frmabout.Show vbModal, frmmain ' show the about box
        Case 2 ' exit button was pressed
            ans = MsgBox("Are you sure you want to quit now.", vbYesNo Or vbQuestion, "Quit....")
            If ans = vbNo Then Exit Sub
            Unload frmmain  ' Unload the form
    End Select
    
End Sub

Private Sub Form_Load()
    frmmain.Icon = Nothing  ' Remove the forms icon
    lstV.ListItems.Add , "TXT_HTM", "Text To HTM", 1, 1
    lstV.ListItems.Add , "TXT_JAVA", "Text to JavaScript", 2, 2
    lstV.ListItems.Add , "TXT_ASP", "Text to ASP", 3, 3
    lstV.ListItems.Add , "TXT_RTF", "Text to Rich Text Format", 4, 4
    lstV.ListItems.Add , "TXT_PHP", "Text to PHP", 5, 5
    lstV.ListItems.Add , "TXT_LINUX2PC", "Linux Text to PC Text", 6, 6
    lstV.ListItems.Add , "TXT_MAC2PC", "Mac Text to PC Text", 7, 7
    lstV.ListItems.Add , "TXT_PERL", "Text to Perl", 8, 8
    lstV.ListItems.Add , "TXT_VBS", "Text to VBScript", 9, 9
    lstV.ListItems.Add , "TXT_EXE", "Text to Win32 Application", 10, 10
    lstV.ListItems(1).Selected = False
    StatusBar1.Panels(1).Text = "Currently supports" & lstV.ListItems.Count & " formats."
    
End Sub

Private Sub Form_Resize()
    lntop.X2 = frmmain.ScaleWidth - 1 ' Resize the line to fit the form
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmmain = Nothing
End Sub

Private Sub lstV_Click()
    CmdBut(0).Enabled = True       ' Enable the create new button
    lstKey = lstV.SelectedItem.Key ' Get the list view key value
End Sub

Private Sub lstV_DblClick()
    CmdBut_Click 0
End Sub

Private Sub lstV_ItemClick(ByVal Item As MSComctlLib.ListItem)
    lstV_Click  'Raise the listview click event
End Sub

