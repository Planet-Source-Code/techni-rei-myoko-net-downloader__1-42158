Attribute VB_Name = "ListViewHandling"
Option Explicit
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
    Const LVM_FIRST = &H1000
    
Public Sub AutoSizeColumnHeader(LView As ListView, Column As ColumnHeader, Optional ByVal SizeToHeader As Boolean = True)
On Error Resume Next
    Dim l As Long
    If SizeToHeader Then l = -2 Else l = -1
    Call SendMessage(LView.hWnd, LVM_FIRST + 30, Column.Index - 1, l)
End Sub

Public Sub resizecolumnheaders(LView As ListView)
On Error Resume Next
Dim temp As Integer
If LView.ListItems.count > 0 Then
    For temp = 1 To LView.ColumnHeaders.count
        AutoSizeColumnHeader LView, LView.ColumnHeaders.Item(temp)
    Next
End If
End Sub
