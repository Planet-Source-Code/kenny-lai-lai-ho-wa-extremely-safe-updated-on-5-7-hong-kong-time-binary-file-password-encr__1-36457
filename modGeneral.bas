Attribute VB_Name = "modGeneral"
Option Explicit
Public Declare Function ShellExecute Lib "shell32.dll" _
        Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal _
        lpOperation As String, ByVal lpFile As String, ByVal _
        lpParameters As String, ByVal lpDirectory As String, _
        ByVal nShowCmd As Long) As Long

Public Function Browser(Adresse As String, FRMhWnd As Long)
  Dim Ret As Long
    Ret = ShellExecute(FRMhWnd, "Open", Adresse, "", App.Path, 1)
End Function
