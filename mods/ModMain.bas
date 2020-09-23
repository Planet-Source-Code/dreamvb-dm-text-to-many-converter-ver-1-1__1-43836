Attribute VB_Name = "ModMain"
Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long

Public Form_Caption As String   ' Holds the convert forms caption
Public save_dlgTitle As String  ' Holds the dialogs caption
Public save_dlgFilter As String ' Holds the Dialogs filter value

Public lstKey As String         ' Holds the key for users convert option
Public Function GetPath(lzPath As String) As String
Dim I As Long
    For I = Len(lzPath) To 1 Step -1
        If InStr(I, lzPath, "\", vbTextCompare) Then
            GetPath = Mid(lzPath, 1, I)
            Exit For
            Exit Function
        End If
    Next
    I = 0
    
End Function
Public Function GetFileTitle(lzFile As String) As String
Dim I As Long, First As String

    For I = Len(lzFile) To 1 Step -1 ' Start looping form the end to 1
        If InStr(I, lzFile, "\", vbTextCompare) Then ' Check for the slash
            First = Mid(lzFile, I + 1, Len(lzFile)) ' Extract the filename
            GetFileTitle = Mid(First, 1, InStr(1, First, ".") - 1) ' Extract the file title we don't need the dot
            Exit For
            Exit Function
        End If
    Next
    I = 0
    First = ""
    
End Function
Public Function Fixpath(lzPath As String) As String
    ' This will add a backslash to a path if needed
    If Right(lzPath, 1) = "\" Then Fixpath = lzPath Else Fixpath = lzPath & "\"
End Function
Public Function FindFile(mFile As String) As Boolean
    ' This check to see if a file can be found or not
    If Dir(mFile) = "" Then FindFile = False Else FindFile = True
End Function

Public Function OpenData(lzFile As String) As String
Dim nFile As Long, fData() As Byte

' This is used to open binary files useing ByteArrays for faster access
    nFile = FreeFile ' Pointer to the file
    Open lzFile For Binary As #nFile ' open the file in binary mode
        ReDim fData(LOF(nFile) - 1) ' resize the byte array
        Get #nFile, , fData()   ' get file contents
    Close #nFile    ' close the file

    OpenData = StrConv(fData, vbUnicode) ' convert byte to string and pass back
    Erase fData ' erase the array we have finished with it now
    
End Function

Public Function SaveData(lzFilename As String, lzdata As String)
Dim nFile As Long
    nFile = FreeFile    ' Pointer to freefile
    Open lzFilename For Output As #nFile ' Open the file for output mode
        Print #nFile, lzdata    ' Print string contents to file
    Close #nFile    ' Close the file
    
End Function

Public Function Encrypt(S As String) As String
Dim I As Long
Dim ch As String

    For I = 1 To Len(S)
        ch = ch & Chr(48 Xor Asc(Mid(S, I, 1)))
    Next
    I = 0
    Encrypt = ch
    ch = ""
End Function
