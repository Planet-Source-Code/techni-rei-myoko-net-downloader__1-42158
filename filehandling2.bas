Attribute VB_Name = "filehandling"
Option Explicit

Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

Public Function uniquefilename(filename As String) As String
    Dim temp1 As String, temp2 As String, temp3 As Long
    uniquefilename = filename
    
    If FileExists(filename) Then
        Dim count As Long
        count = 1
        temp3 = InStrRev(filename, ".")
        temp1 = filename
        If temp3 > 0 Then
            temp1 = Left(filename, temp3 - 1)
            temp2 = Right(filename, Len(filename) - temp3 + 1)
        End If
        Do Until FileExists(temp1 & " (" & count & ")" & temp2) = False
            count = count + 1
        Loop
        uniquefilename = temp1 & " (" & count & ")" & temp2
    End If
End Function
Public Function DownloadFile(URL As String, filename As String) As Boolean
On Error Resume Next

If Len(filename) > 255 Then filename = Left(filename, 255)

filename = Replace(filename, "*", Empty)
filename = Replace(filename, "&", Empty)
filename = Replace(filename, "%", Empty)
filename = Replace(filename, "=", Empty)

If InStrRev(filename, "?") > 0 Then
    Dim temp As String
    temp = Left(filename, InStrRev(filename, "?") - 1)
    If InStr(temp, ".") = 0 Then temp = temp & ".txt"
    filename = uniquefilename(temp)
End If

Debug.Print filename

DownloadFile = False
'Downloads the file from URL and saves it as filename
If URL <> Empty And filename <> Empty Then
    Dim lngRetVal As Long
    lngRetVal = URLDownloadToFile(0, URL, filename, 0, 0)
    If lngRetVal = 0 Then DownloadFile = True
End If
End Function

'I could make these 2 into one function, but I like it seperate
Public Function direxists(directory As String) As Boolean
'Checks to see if a directory exists
On Error Resume Next
If Dir(directory, vbDirectory + vbHidden) = Empty Then direxists = False Else direxists = True
End Function
Public Function FileExists(filename As String) As Boolean
'Checks to see if a file exists
On Error Resume Next
If Dir(filename) = Empty Then FileExists = False Else FileExists = True
End Function

Public Function chkfile(directory As String, filename As String) As String
'Adds the filename to a dir without getting an error if its the root dir
On Error Resume Next
If Right(directory, 1) <> "\" Then chkfile = directory & "\" & filename Else chkfile = directory & filename
End Function

Public Function loadfile(filename As String) As String
On Error Resume Next
Dim intFile As Integer, temp As String, allfile As String
allfile = Empty
If Dir(filename) <> Empty And Right(filename, 1) <> "\" Then
intFile = FreeFile()
Open filename For Input As intFile
Do Until EOF(intFile)
    Line Input #intFile, temp
    allfile = allfile & temp & vbNewLine
Loop
Close intFile
loadfile = Left(allfile, Len(allfile) - 1)
Else
loadfile = Empty
End If
End Function
