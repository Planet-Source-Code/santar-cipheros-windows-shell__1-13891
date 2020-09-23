Attribute VB_Name = "OnTopandINI"
Option Explicit
Declare Function SetWindowPos Lib "user32" _
    (ByVal hWnd As Long, _
    ByVal hWndInsertAfter As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal cx As Long, _
    ByVal cy As Long, _
    ByVal wFlags As Long) As Long
Public Declare Function WritePrivateProfileString Lib "KERNEL32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "KERNEL32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public Function GetFromINI(Section As String, Key As String, Directory As String) As String
 Dim strBuffer As String
 strBuffer = String(750, Chr(0))
 Key$ = LCase$(Key$)
 GetFromINI$ = Left(strBuffer, GetPrivateProfileString(Section$, ByVal Key$, "", strBuffer, Len(strBuffer), Directory$))
End Function

Public Sub WriteToINI(Section As String, Key As String, KeyValue As String, Directory As String)
  Call WritePrivateProfileString(Section$, UCase$(Key$), KeyValue$, Directory$)
End Sub

