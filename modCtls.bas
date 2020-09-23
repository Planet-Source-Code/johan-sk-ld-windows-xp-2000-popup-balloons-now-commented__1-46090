Attribute VB_Name = "modCtls"
Option Explicit

Public Declare Function InitCommonControls Lib "comctl32" () As Long                    ' Used to enable xp style
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal _
    hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long      ' Used to display the
                                                                                        ' Balloon for textboxes

Public Type EDITBALLOONTIP          ' Type used to store the balloon data for textboxes
   cbStruct As Long
   pszTitle As String
   pszText As String
   ttiIcon As Long
End Type

Public Const ECM_FIRST As Long = &H1500                     ' Stores the start settings for the two below
Public Const EM_SHOWBALLOONTIP As Long = (ECM_FIRST + 3)    ' Shows a balloon on a textbox
Public Const EM_HIDEBALLOONTIP As Long = (ECM_FIRST + 4)    ' Hides a balloon from a textbox
Public LicenseChange As Boolean

Public Sub SecureProgram()  ' Secures this app so no one just changes the info and calls it his/hers app
                            ' you may remove this when you've changed enough to call it your app
    If Not App.Major = 1 Or Not App.Minor = 0 Or Not App.Revision = 0 Or _
      Not App.CompanyName = "Johan Sköld (http://www.johanskold.host.sk)" Or _
      Not App.LegalCopyright = "Copyright © 1998-2003 Johan Sköld" Or _
      Not App.ProductName = "Windows 2000/XP Popup Balloons" Then
        
        MsgBox "Please restore the Version info in this file (PopupBalloons.ocx) before using." & vbCrLf & vbCrLf & "According to the license you may use this as you like as long as you DON'T CHANGE IT!", vbCritical, "License Break"
        LicenseChange = True
    Else
        LicenseChange = False
    End If
End Sub
