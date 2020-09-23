Attribute VB_Name = "modFileFunctions"
Function XGetSize(ByVal Filename As String) As Long
On Error GoTo ErrorOccured
freef = FreeFile
XGetSize = VBA.FileSystem.FileLen(Filename)
Exit Function
ErrorOccured:
XGetSize = -1
End Function
Function XGetFileName(ByVal Filename As String) As String
i = InStrRev(Filename, "\")
If i > 0 Then
XGetFileName = Mid(Filename, i + 1)
End If
End Function
Function XGetParentFolder(ByVal Filename As String) As String
i = InStrRev(Filename, "\")
If i > 0 Then
XGetParentFolder = Left(Filename, i)
End If
End Function

Function XBuildPath(ByVal sPath As String, sFileName As String)
If Right(sPath, 1) <> "\" Then sPath = sPath & "\"
XBuildPath = sPath & sFileName
End Function
Function IsDirectoryExist(ByVal SomePath As String) As Boolean
On Error GoTo ErrorOccured
ChDir txtDir
IsDirectoryExist = True
Exit Function
ErrorOccured:
IsDirectoryExist = False
End Function


Function IsFileExist(ByVal Filename As String) As Boolean
On Error GoTo ErrorOccured
freef = FreeFile
Open Filename For Input As freef: Close freef
IsFileExist = True
Exit Function
ErrorOccured:
IsFileExist = False
End Function

Sub DeleteFile(ByVal Filename)
On Error Resume Next
Kill Filename
End Sub
