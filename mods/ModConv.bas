Attribute VB_Name = "ModConv"

Function Text2ASP(txtData As String) As String
Dim mLine As Variant
Dim ASP_Top As String, ASP_Body As String
Dim I As Long
    
    ASP_Top = "<%" & vbNewLine ' Top of ASP code
    
    mLine = Split(txtData, vbNewLine)
    
    For I = LBound(mLine) To UBound(mLine)
        ASP_Body = ASP_Body & "Response.Write " & Chr(34) & FixVB(CStr(mLine(I))) & Chr(34) & " & " & "vbCrLf" & vbNewLine
    Next
    I = 0 ' reset the counter
        
    Text2ASP = ASP_Top & ASP_Body & "%>"  ' Build the ASP code
    ASP_Top = ""    ' Clear var
    ASP_Body = ""   ' Clear var
End Function

Function Text2Html(txtData As String, Optional PageTitle As String) As String
Dim mLine As Variant
Dim Htm_Top As String, htm_Body As String
Dim I As Long

    Htm_Top = Htm_Top & "<head>" & vbNewLine
    Htm_Top = Htm_Top & "<html>" & vbNewLine
    Htm_Top = Htm_Top & "<title>" & PageTitle & "</title>"
    Htm_Top = Htm_Top & "<meta http-equiv=" & Chr(34) & "Content-Type" & Chr(34) _
    & " content=" & Chr(34) & "text/html; charset=ios-8859-1" & Chr(34) & ">" & vbNewLine
    Htm_Top = Htm_Top & "</head>" & vbNewLine
    Htm_Top = Htm_Top & "<body bgcolor=" & Chr(34) & "#FFFFFF" & Chr(34) _
    & " text=" & Chr(34) & "#000000" & Chr(34) & ">" & vbNewLine
    Htm_Top = Htm_Top & "<p><font size=" & Chr(34) & "2 " & Chr(34) & "face=" & Chr(34) _
    & "Verdana, Arial, Helvetica, sans-serif" & Chr(34) & ">"
    
    mLine = Split(txtData, vbNewLine)
    For I = LBound(mLine) To UBound(mLine) - 1
        htm_Body = htm_Body & FixHtml(CStr(mLine(I))) & "<br>"
    Next
    I = 0
    Text2Html = Htm_Top & htm_Body & "</font>" & vbNewLine & "</p>" _
    & vbNewLine & "</body>" & vbNewLine & "</html>"
    Htm_Top = ""
    htm_Body = ""
    
End Function

Function Text2JScript(txtData As String) As String
Dim mLine As Variant
Dim JavaS_Top As String, JavaBody_Body As String
Dim I As Long
    
    JavaS_Top = "<script language=" & Chr(34) & "JavaScript" & Chr(34) & ">" & vbNewLine _
    & "<!--" & vbNewLine
    
    mLine = Split(txtData, vbNewLine)
    
    For I = LBound(mLine) To UBound(mLine)
        JavaBody_Body = JavaBody_Body & "document.write(" & Chr(34) & FixChar34(CStr(mLine(I))) _
        & Chr(34) & " + " & Chr(34) & "<br>" & Chr(34) & ");" & vbNewLine
    Next
    I = 0
        
    Text2JScript = JavaS_Top & JavaBody_Body & "//-->" & vbNewLine & "</script>"
    JavaS_Top = ""
    JavaBody_Body = ""
End Function

Function Text2VBSscript(txtData As String) As String
Dim mLine As Variant
Dim VBS_Top As String, VBS_Body As String
Dim I As Long
    
    VBS_Top = "<script language=" & Chr(34) & "VBScript" & Chr(34) & ">" & vbNewLine
    
    mLine = Split(txtData, vbNewLine)
    
    For I = LBound(mLine) To UBound(mLine)
        VBS_Body = VBS_Body & "document.write(" & Chr(34) & FixVB(CStr(mLine(I))) _
        & Chr(34) & " + " & Chr(34) & "<br>" & Chr(34) & ")" & vbNewLine
    Next
    I = 0
        
    Text2VBSscript = VBS_Top & VBS_Body & "</script>"
    VBS_Top = ""
    VBS_Body = ""
End Function


Function Text2RTF(txtData As String) As String
Dim mLine As Variant
Dim RTF_Top As String, RTF_Body As String
Dim I As Long
    
    RTF_Top = "{\rtf1\ansi\ansicpg1252\deff0\deflang1033{\fonttbl{\f0\fswiss\fcharset0 Arial;}}" _
    & vbNewLine & "\viewkind4\uc1\pard\f0\fs20 " ' Top of RTF code
    
    mLine = Split(txtData, vbNewLine)
    
    For I = LBound(mLine) To UBound(mLine)
        RTF_Body = RTF_Body & mLine(I) & "\par" & vbNewLine
    Next
    I = 0
        
    Text2RTF = RTF_Top & RTF_Body & "}" ' Build the Rich Text Format data
    RTF_Top = ""    ' Clear Var
    RTF_Body = ""   ' Clear Var
End Function

Function Text2PHP(txtData As String) As String
Dim mLine As Variant
Dim PHP_Top As String, PHP_Body As String
Dim I As Long
    
    PHP_Top = "<?php" & vbNewLine ' Top of php code
    
    mLine = Split(txtData, vbNewLine)
    
    For I = LBound(mLine) To UBound(mLine)
        PHP_Body = PHP_Body & "echo " & Chr(34) & FixChar34(CStr(mLine(I))) & " <br>" & Chr(34) & ";" & vbNewLine
    Next
    I = 0 ' reset the counter
        
    Text2PHP = PHP_Top & PHP_Body & "?>"  ' Build the PHP code
    PHP_Top = ""    ' Clear var
    PHP_Body = ""   ' Clear var
End Function

Function Text2Perl(txtData As String) As String
Dim mLine As Variant
Dim Perl_Body As String
Dim I As Long
    
    
    mLine = Split(txtData, vbNewLine)
    
    For I = LBound(mLine) To UBound(mLine)
        Perl_Body = Perl_Body & "print " & Chr(34) & FixChar34(CStr(mLine(I))) & " <br>" & Chr(34) & ";" & vbNewLine
    Next
    I = 0 ' reset the counter
    Text2Perl = Perl_Body  ' Build the perl code
    Perl_Body = ""   ' Clear var
    
End Function

Public Function FixLinuxTXT(StrTxt As String) As String
Dim I As Long, StrCh As String, StrB As String
    Do While I < Len(StrTxt)
        I = I + 1 ' Add one to our counter
        StrCh = Mid$(StrTxt, I, 1) ' Get a char from the string
        If Asc(StrCh) = 10 Then ' Does the char equal 10
            StrB = StrB & vbNewLine ' replace it with the PC newline
        Else
            StrB = StrB & StrCh ' Don't change anything
        End If
    Loop
    I = 0 ' Reset counter
    StrCh = ""  ' Clear string buffer
    FixLinuxTXT = StrB
    StrB = "" ' Clear string buffer
    
End Function

Public Function FixMacText(lzdata As String) As String
Dim I As Long, CH As Long
Dim StrB As String
    For I = 1 To Len(lzdata) ' Loop till we get to the end of the data
        CH = Asc(Mid(lzdata, I, 1)) ' Get a char from the string
        If CH = 13 Then ' If we find 13 add the vbnewline
            StrB = StrB & vbNewLine ' Add a new line to the string
        Else
            StrB = StrB & Chr(CH) ' Don't change anything
        End If
    Next
    CH = 0   ' Reset var
    I = 0    ' Reset counter
    FixMacText = StrB 'Pass new data back
    StrB = ""   ' Clear buffer
End Function

Private Function FixHtml(lzdata As String) As String
Dim sBuffer As String
    sBuffer = lzdata
    sBuffer = Replace(sBuffer, Chr$(34), "&quot;")
    sBuffer = Replace(sBuffer, "&", "&amp;")
    sBuffer = Replace(sBuffer, "<", "&lt;")
    sBuffer = Replace(sBuffer, ">", "&gt;")
    FixHtml = sBuffer
    sBuffer = ""
End Function

Private Function FixChar34(lzdata As String) As String
Dim sBuffer As String
    sBuffer = lzdata
    sBuffer = Replace(lzdata, Chr$(34), "\" & Chr(34))
    FixChar34 = sBuffer
    sBuffer = ""
End Function

Private Function FixVB(lzdata As String)
Dim sBuffer As String
    sBuffer = lzdata
    sBuffer = Replace(sBuffer, Chr(34), Chr(34) & Chr(34))
    FixVB = sBuffer
    sBuffer = ""
End Function
