VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmmain 
   Caption         =   "Net Downloader"
   ClientHeight    =   4305
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   4695
   Icon            =   "frmmain1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4305
   ScaleWidth      =   4695
   StartUpPosition =   3  'Windows Default
   Tag             =   "Net Downloader"
   Begin VB.TextBox txtubound 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   3240
      TabIndex        =   14
      Text            =   "100"
      ToolTipText     =   "Put the highest number here"
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox txtlbound 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2640
      TabIndex        =   13
      Text            =   "0"
      ToolTipText     =   "Put the lowest number here (Usually 1 or 0)"
      Top             =   1200
      Width           =   495
   End
   Begin VB.CommandButton cmdadd 
      Caption         =   "&Add"
      Height          =   285
      Left            =   3840
      TabIndex        =   12
      ToolTipText     =   "Add a range of files to the list from the URL typed above"
      Top             =   1200
      Width           =   735
   End
   Begin VB.TextBox txtfiles 
      Height          =   285
      Left            =   840
      TabIndex        =   11
      ToolTipText     =   "Type the name of the file range here, replacing the numbers with #. If it's numbered like 001, then replace it with ###"
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox txtpattern 
      Height          =   285
      Left            =   2040
      TabIndex        =   9
      Text            =   "*.htm;*.html;*.asp;*.txt"
      ToolTipText     =   $"frmmain1.frx":0E42
      Top             =   840
      Width           =   2535
   End
   Begin VB.CheckBox chkmain 
      Caption         =   "&Recurse files of type:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      ToolTipText     =   "Rescans files of the following type for more files"
      Top             =   840
      Width           =   1815
   End
   Begin VB.CommandButton cmdclear 
      Caption         =   "&Clear"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton cmdbrowse 
      Caption         =   "&Browse"
      Height          =   285
      Left            =   3840
      TabIndex        =   4
      ToolTipText     =   "Browse your computers directory structure for a place to store the downloaded files"
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox txtdir 
      Height          =   285
      Left            =   840
      OLEDropMode     =   2  'Automatic
      TabIndex        =   3
      ToolTipText     =   "Please select a download directory"
      Top             =   480
      Width           =   2895
   End
   Begin VB.CommandButton cmdgo 
      Caption         =   "&Go"
      Height          =   285
      Left            =   3840
      TabIndex        =   1
      ToolTipText     =   "Search the file (the one in the text box to the left) for links and images"
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox txtURL 
      Height          =   285
      Left            =   840
      OLEDropMode     =   2  'Automatic
      TabIndex        =   0
      ToolTipText     =   "Please select a source page"
      Top             =   120
      Width           =   2895
   End
   Begin VB.CommandButton cmddownload 
      Caption         =   "&Download Checked Files"
      Height          =   375
      Left            =   1320
      TabIndex        =   5
      Top             =   3840
      Width           =   3255
   End
   Begin VB.CommandButton cmdstop 
      Caption         =   "&Stop"
      Height          =   375
      Left            =   1320
      TabIndex        =   6
      Top             =   3840
      Visible         =   0   'False
      Width           =   3255
   End
   Begin MSComctlLib.ListView lstmain 
      Height          =   2175
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   3836
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      OLEDropMode     =   1
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      OLEDropMode     =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Extracted"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "URL"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblpattern 
      Caption         =   "Directory:"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   16
      Top             =   480
      Width           =   735
   End
   Begin VB.Label lblpattern 
      Caption         =   "URL:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblpattern 
      Caption         =   "Pattern:"
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   1200
      Width           =   615
   End
   Begin VB.Menu mnuselect 
      Caption         =   "Select"
      Visible         =   0   'False
      Begin VB.Menu mnusel 
         Caption         =   "Invert Selected"
         Index           =   0
      End
   End
   Begin VB.Menu mnuontop 
      Caption         =   "OnTop"
      Visible         =   0   'False
      Begin VB.Menu mnualwaysontop 
         Caption         =   "&Always on Top"
      End
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim spoth As Integer, spoth2 As Integer, stopped As Boolean

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const HTCAPTION = 2
Private Const WM_NCLBUTTONDOWN = &HA1
Private Declare Function SetWindowPos Lib "User32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    Private Const SWP_NOMOVE = &H2
    Private Const SWP_NOSIZE = &H1
    'Used to set window to always be on top or not
    Private Const HWND_NOTOPMOST = -2
    Private Const HWND_TOPMOST = -1

Public Sub dragform(hwnd As Long)
On Error Resume Next
  ReleaseCapture
  SendMessage hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Public Sub setAlwaysOnTop(hwnd As Long, Optional ontop As Boolean = True)
On Error Resume Next
If ontop = False Then Call SetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 1, 1, SWP_NOMOVE Or SWP_NOSIZE)
If ontop = True Then Call SetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 1, 1, SWP_NOMOVE Or SWP_NOSIZE)
End Sub

Private Sub cmdadd_Click()
If txtURL <> Empty Then
For spoth = Val(txtlbound) To Val(txtubound) 'countchars -1
    If countchars(txtpattern, "#") - 1 = 1 Then
        lstmain.ListItems.add , , Replace(txtfiles, "#", spoth)
    Else
        lstmain.ListItems.add , , Replace(txtfiles, String(countchars(txtfiles, "#") - 1, "#"), Format(spoth, String(countchars(txtfiles, "#") - 1, "0")))
    End If
    lstmain.ListItems.Item(lstmain.ListItems.count).SubItems(1) = "Direct"
    lstmain.ListItems.Item(lstmain.ListItems.count).SubItems(2) = chkurl(txtURL, lstmain.ListItems.Item(lstmain.ListItems.count).text)
    lstmain.ListItems.Item(lstmain.ListItems.count).Checked = True
Next
removedoubles
resizecolumnheaders frmmain.lstmain
Else
MsgBox "I need a site to download the files from first", vbCritical, "No URL given."
End If
End Sub

Private Sub cmdbrowse_Click()
Dim temp As String
BB.ShowCheck = True
BB.ShowButton = True
BB.Prompt = txtdir.ToolTipText
BB.AllowResize = True
BB.EditBoxNew = True
If direxists(txtdir) = True Then BB.InitDir = txtdir
temp = BrowseFF
'temp = BrowseForFolder(Me.hwnd, txtdir.ToolTipText)
If temp <> Empty Then txtdir = temp
End Sub

Private Sub cmdclear_Click()
lstmain.ListItems.Clear
End Sub

Private Sub cmddownload_Click()
Dim filename As String
If txtdir.BackColor = vbRed Then
    MkDir (txtdir)
    txtdir.BackColor = vbWhite
End If
stopped = False
cmdstop.Visible = True
For spoth = 1 To lstmain.ListItems.count
    If lstmain.ListItems.Item(spoth).Checked = True And stopped = False Then
        filename = Right(lstmain.ListItems.Item(spoth).SubItems(2), Len(lstmain.ListItems.Item(spoth).SubItems(2)) - InStrRev(lstmain.ListItems.Item(spoth).SubItems(2), "/"))
        Me.Caption = "Downloading " & filename & " (" & spoth & "/" & lstmain.ListItems.count & ")(" & Round(spoth / lstmain.ListItems.count * 100, 2) & "%)"
        DoEvents
        If DownloadFile(lstmain.ListItems.Item(spoth).SubItems(2), chkfile(txtdir.text, filename)) = True Then Me.Caption = "Download Successful"
    End If
Next
cmdstop.Visible = False
Me.Caption = Me.Tag
End Sub
Public Sub deletefile(filename As String)
    If Right(filename, 1) <> "\" And Dir(filename) <> Empty Then Kill filename
End Sub
Public Sub cmdgo_Click()
Dim localfile As String
txtURL = Replace(txtURL, "\", "/")
If Right(txtURL, 1) = "/" Then txtURL = Left(txtURL, Len(txtURL) - 1)
Me.Caption = "Downloading " & txtURL & " from web"
stopped = False
If txtURL <> Empty And Dir(txtdir.text, vbDirectory) <> Empty Then
localfile = chkfile(txtdir, Right(txtURL, Len(txtURL) - InStrRev(txtURL, "/")))
deletefile localfile
If DownloadFile(txtURL, localfile) = True Then
    txtURL.Tag = Left(txtURL, InStrRev(txtURL, "/"))
    If Len(txtURL.Tag) - Len(Replace(txtURL.Tag, "/", Empty)) = 2 Then txtURL.Tag = txtURL
    Me.Caption = "Extracting URLs"
    Call extracturls(localfile, txtURL.Tag)
End If
Else
    Call MsgBox("I'm sorry but I can't pull the URLs out of thin air" & vbNewLine & "And I can't put them into thin air either" & vbNewLine & "Please give me a URL and an existing folder", vbCritical, "Please give me a URL/Folder")
End If
Me.Caption = Me.Tag
End Sub
Public Sub extracturls(filename As String, basehref As String)
Dim htmlfile As String, temp As String, tempstr As String
htmlfile = loadfile(filename)
'http://techni.keenspace.com/index.html
'search for src tags (images and scripts)

'Search for "<base href="
Me.Caption = "Checking for a base href redefinition"
If countwords(htmlfile, "<base href=") > 1 Then
    basehref = addfrom(htmlfile, findchar(htmlfile, "<base href=", spoth) + Len("<base href="))
End If
'lstmain.ToolTipText = basehref

'Search for "src="
Me.Caption = "Searching for src tags"
For spoth = 1 To countwords(htmlfile, "src=") - 1
DoEvents
    temp = addfrom(htmlfile, findchar(htmlfile, "src=", spoth) + Len("src="))
    If temp <> Empty Then
        lstmain.ListItems.add , , temp
        Select Case LCase(Right(lstmain.ListItems.Item(lstmain.ListItems.count).text, 3))
            Case "vbs", ".vb", ".js", ".inc"
                lstmain.ListItems.Item(lstmain.ListItems.count).SubItems(1) = "Script"
            Case Else
                lstmain.ListItems.Item(lstmain.ListItems.count).SubItems(1) = "Image"
        End Select
    End If
Next

'Search for "href="
Me.Caption = "Searching for href tags"
For spoth = 1 To countwords(htmlfile, "href=") - 1
    DoEvents
    temp = addfrom(htmlfile, findchar(htmlfile, "href=", spoth) + Len("href="))
    If temp <> Empty And LCase(Left(temp, Len("mailto:"))) <> "mailto:" Then 'cancel out blank urls and mailto's
        lstmain.ListItems.add , , temp
        lstmain.ListItems.Item(lstmain.ListItems.count).SubItems(1) = "Link"
    End If
Next

'Resolve the urls
Me.Caption = "Converting added tags to URLs"
For spoth = 1 To lstmain.ListItems.count
    DoEvents
    If lstmain.ListItems.Item(spoth).SubItems(2) = Empty Then 'cause now they can be done before hand (recursive)
        lstmain.ListItems.Item(spoth).SubItems(2) = chkurl(basehref, lstmain.ListItems.Item(spoth).text)
        lstmain.ListItems.Item(spoth).Checked = True
        If Left(lstmain.ListItems.Item(spoth).text, 1) = "#" Then
            lstmain.ListItems.Item(spoth).SubItems(2) = chkurl(basehref, Left(lstmain.ListItems.Item(spoth), InStr(lstmain.ListItems.Item(spoth), "#") - 1))
        End If
    End If
Next

removedoubles

Me.Caption = "Removing blank URLS"
For spoth = lstmain.ListItems.count To 1 Step -1
    DoEvents
    If lstmain.ListItems.Item(spoth).SubItems(2) = Empty Then lstmain.ListItems.Remove spoth
Next

'Recurse the htmls
Me.Caption = "Recursing linked files"
If chkmain.Value = vbChecked Then
    tempstr = txtURL
    
    For spoth = 1 To lstmain.ListItems.count
        DoEvents
        If spoth <= lstmain.ListItems.count Then
        If lstmain.ListItems.Item(spoth).SubItems(1) = "Link" Then
            If islike(txtpattern, lstmain.ListItems.Item(spoth).SubItems(2)) = True Then
            If sameurl(txtURL, lstmain.ListItems.Item(spoth).SubItems(2)) = False Then
            txtURL = lstmain.ListItems.Item(spoth).SubItems(2)
            If lstmain.ListItems.Item(spoth).Tag <> True Then
                lstmain.ListItems.Item(spoth).Tag = True
                cmdgo.Tag = False
                cmdgo_Click
                cmdgo.Tag = Empty
            End If
            End If
            End If
        End If
        End If
    Next
    
    txtURL = tempstr
End If
resizecolumnheaders frmmain.lstmain
End Sub
Public Function islike(filter As String, text As String) As Boolean
    Dim patterns() As String, count As Long
    patterns = Split(filter, ";")
    islike = False
    For count = LBound(patterns) To UBound(patterns)
        If text Like patterns(count) Then islike = True
    Next
End Function
Public Function sameurl(ByVal url1 As String, ByVal url2 As String) As Boolean
'compares urls, accounts for the fact that urls with # in it are only whats before the #
    If InStr(url1, "#") > 0 Then url1 = Left(url1, InStr(url1, "#") - 1)
    If InStr(url2, "#") > 0 Then url2 = Left(url2, InStr(url2, "#") - 1)
    If LCase(url1) = LCase(url2) Then sameurl = True Else sameurl = False
End Function
Public Function addfrom(content As String, location As Long) As String
    Dim temp As Long 'end of the url
    If Mid(content, location, 1) = """" Or Mid(content, location, 1) = "'" Then
        temp = findnext(content, location, ">'""")
        If temp > 0 Then addfrom = Replace(Mid(content, location + 1, temp - location - 1), ">", "")
    Else
        temp = findnext(content, location, " >")
        addfrom = Replace(Mid(content, location, temp - location + 1), ">", "")
    End If
    'MsgBox temp & " " & location
End Function

Private Sub cmdstop_Click()
    stopped = True
    cmdstop.Visible = False
End Sub
Public Sub removedoubles()
'Remove doubles of the resolved urls
Me.Caption = "Removing Doubles"
For spoth = 1 To lstmain.ListItems.count 'To 1
    DoEvents
    If spoth <= lstmain.ListItems.count Then
    If LCase(lstmain.ListItems.Item(spoth).SubItems(2)) = LCase(txtURL) Or lstmain.ListItems.Item(spoth).SubItems(2) = Empty Then
        lstmain.ListItems.Remove spoth
    Else
    For spoth2 = lstmain.ListItems.count To spoth + 1 Step -1
        'If LCase(lstmain.ListItems.Item(spoth).SubItems(2)) = LCase(lstmain.ListItems.Item(spoth2).SubItems(2)) Then
            If sameurl(lstmain.ListItems.Item(spoth).SubItems(2), lstmain.ListItems.Item(spoth2).SubItems(2)) = True Then lstmain.ListItems.Remove spoth2
        'End If
    Next
    End If
    End If
Next
End Sub
Private Sub Form_Load()
    txtdir = GetSetting("Net Downloader", "Main", "Last Path", App.Path)
    WindowState = GetSetting("Net Downloader", "Main", "WindowState", WindowState)
    Width = GetSetting("Net Downloader", "Main", "Width", Width)
    Height = GetSetting("Net Downloader", "Main", "Height", Height)
    Top = GetSetting("Net Downloader", "Main", "Top", Top)
    Left = GetSetting("Net Downloader", "Main", "Left", Left)
    chkmain.Value = GetSetting("Net Downloader", "Main", "Recurse Links", vbUnchecked)
    txtpattern.text = GetSetting("Net Downloader", "Main", "Pattern", "*.htm?;*.html;*.asp;*.txt")
    mnualwaysontop.Checked = GetSetting("Net Downloader", "Main", "Always On Top", False)
    If mnualwaysontop.Checked Then setAlwaysOnTop Me.hwnd, True
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then PopupMenu Me.mnuontop
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    Call SaveSetting("Net Downloader", "Main", "Last Path", txtdir.text)
    Call SaveSetting("Net Downloader", "Main", "WindowState", WindowState)
    WindowState = 0
    Call SaveSetting("Net Downloader", "Main", "Width", Width)
    Call SaveSetting("Net Downloader", "Main", "Height", Height)
    Call SaveSetting("Net Downloader", "Main", "Top", Top)
    Call SaveSetting("Net Downloader", "Main", "Left", Left)
    Call SaveSetting("Net Downloader", "Main", "Recurse Links", chkmain.Value)
    Call SaveSetting("Net Downloader", "Main", "Pattern", txtpattern)
    Call SaveSetting("Net Downloader", "Main", "Always On Top", mnualwaysontop.Checked)
End Sub

Private Sub Form_Resize()
If Width > 2500 And Height > 2500 Then
    cmdgo.Left = Width - 930
    cmdbrowse.Left = cmdgo.Left
    txtURL.Width = cmdgo.Left - txtURL.Left - 120
    txtdir.Width = txtURL.Width
    lstmain.Width = cmdgo.Left + cmdgo.Width - lstmain.Left
    cmddownload.Width = lstmain.Width - cmdclear.Width - 120
    cmddownload.Top = Height - 900
    cmdclear.Top = cmddownload.Top
    lstmain.Height = cmddownload.Top - lstmain.Top - 120
    cmdstop.Move cmddownload.Left, cmddownload.Top, cmddownload.Width, cmddownload.Height
    txtpattern.Width = cmdgo.Left + cmdgo.Width - txtpattern.Left
    cmdadd.Left = cmdgo.Left
    txtubound.Left = cmdadd.Left - txtubound.Width - 120
    txtlbound.Left = txtubound.Left - txtlbound.Width - 120
    txtfiles.Width = txtlbound.Left - txtfiles.Left - 120
End If
End Sub

Public Function chkurl(ByVal basehref As String, URL As String) As String
'check for absolute (is like *://*)
'check for relative (contains ../)
'check for additive (else)
If Left(URL, 1) = "#" Then Exit Function 'is not a file
If Left(URL, 1) = "/" Then URL = Right(URL, Len(URL) - 1)
'If Len(basehref) - Len(Replace(basehref, "/", Empty)) = 2 Then basehref = basehref & "/"
If containsword(basehref, "://") = False Then basehref = "http://" & basehref
If LCase(URL) <> LCase(basehref) And URL <> Empty And basehref <> Empty Then
If URL Like "*://*" Then 'is absolute
    chkurl = URL
Else
    If containsword(URL, "../") Then 'is relative
        If Right(basehref, 1) = "/" And Len(basehref) - Len(Replace(basehref, "/", Empty)) > 2 Then basehref = Left(basehref, Len(basehref) - 1)
        If containsword(Replace(basehref, "://", ""), "/") = True Then
            For spoth = 1 To countwords(basehref, "../")
                URL = Right(URL, Len(URL) - Len("../"))
                basehref = Left(basehref, InStrRev(basehref, "/"))
            Next
        Else
            URL = Replace(URL, "../", "")
        End If
        If Right(basehref, 1) <> "/" Then chkurl = basehref & "/" & URL Else chkurl = basehref & URL
    Else 'is additive
        If Right(basehref, 1) <> "/" Then chkurl = basehref & "/" & URL Else chkurl = basehref & URL
    End If
End If
End If
End Function

Private Sub lstmain_Click()
If lstmain.ListItems.count > 0 And lstmain.SelectedItem.Index > -1 Then If lstmain.SelectedItem.SubItems(1) = "Link" Then txtURL = lstmain.SelectedItem.SubItems(2)
End Sub

Private Sub lstmain_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If lstmain.SelectedItem.SubItems(1) = "Link" And Button = vbLeftButton Then
    lstmain.OLEDrag
End If
End Sub

Private Sub lstmain_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then PopupMenu Me.mnuselect
End Sub

Private Sub lstmain_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim tempURL As String
tempURL = Data.GetData(1)

With lstmain.ListItems
    .add , , Right(tempURL, Len(tempURL) - InStrRev(tempURL, "/"))
    If islike(txtpattern, tempURL) Then
        .Item(.count).SubItems(1) = "Link"
    Else
        .Item(.count).SubItems(1) = "Image"
    End If
    .Item(.count).SubItems(2) = tempURL
    .Item(.count).Checked = True
End With

removedoubles
resizecolumnheaders frmmain.lstmain
Me.Caption = Me.Tag
End Sub

Private Sub lstmain_OLEStartDrag(Data As MSComctlLib.DataObject, AllowedEffects As Long)
Data.Clear
Data.SetData lstmain.SelectedItem.SubItems(2)
End Sub

Private Sub mnualwaysontop_Click()
    mnualwaysontop.Checked = Not mnualwaysontop.Checked
    setAlwaysOnTop Me.hwnd, mnualwaysontop.Checked
End Sub

Private Sub mnusel_Click(Index As Integer)
Dim count As Long
Select Case Index
    Case 0
        For count = 1 To lstmain.ListItems.count
            If lstmain.ListItems.Item(count).Selected = True Then
                lstmain.ListItems.Item(count).Checked = Not lstmain.ListItems.Item(count).Checked
            End If
        Next
End Select
End Sub

Public Sub txtdir_change()
If direxists(txtdir) = False Then txtdir.BackColor = vbRed Else txtdir.BackColor = vbWhite
End Sub

Private Sub txtdir_KeyPress(KeyAscii As Integer)
txtdir_change
End Sub

Private Sub txtdir_KeyUp(KeyCode As Integer, Shift As Integer)
txtdir_change
End Sub

Private Sub txtlbound_Change()
If IsNumeric(txtlbound) = False Then txtlbound = 0
End Sub

Private Sub txtubound_Change()
If IsNumeric(txtubound) = False Then txtubound = 100
End Sub

Private Sub txtURL_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then cmdgo_Click
If KeyAscii = 10 Then
    If LCase(Left(txtURL, Len("http://www."))) <> "http://www." Then txtURL = "http://" & txtURL
    If containsword(txtURL, "/") And containsword(txtURL, ".com") = False Then txtURL = txtURL & ".com/"
End If
End Sub

Private Sub txtURL_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtURL = Data.GetData(1)
    cmdgo_Click
End Sub
