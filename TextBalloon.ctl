VERSION 5.00
Begin VB.UserControl TextBalloon 
   CanGetFocus     =   0   'False
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   495
   InvisibleAtRuntime=   -1  'True
   Picture         =   "TextBalloon.ctx":0000
   ScaleHeight     =   33
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   33
   ToolboxBitmap   =   "TextBalloon.ctx":0D26
End
Attribute VB_Name = "TextBalloon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit     ' Requires Variable Declaration

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

' BalloonStyle Constants
Private Const NIIF_NONE = &H0       ' Nothing special
Private Const NIIF_INFO = &H1       ' Same icon as vbInformation for MsgBox'es
Private Const NIIF_WARNING = &H2    ' Same icon as vbExclamaion for MsgBox'es
Private Const NIIF_ERROR = &H3      ' Same icon as vbCritical for MsgBox'es
Private Const NIIF_GUID = &H5       ' Dunno
Private Const NIIF_ICON_MASK = &HF  ' Dunno
Private Const NIIF_NOSOUND = &H10   ' No sound when balloon pops up

Public Enum BalloonStyle
    bsIconExclamation = NIIF_WARNING
    bsIconCritical = NIIF_ERROR
    bsIconInformation = NIIF_INFO
    bsGuid = NIIF_GUID
    bsIconMask = NIIF_ICON_MASK
    bsNoSound = NIIF_NOSOUND
End Enum

Private Type COMBOBOXINFO       ' Stores info on a particular combobox
   cbSize As Long
   rcItem As RECT
   rcButton As RECT
   stateButton  As Long
   hwndCombo  As Long
   hwndEdit  As Long
   hwndList As Long
End Type

Private Declare Function GetComboBoxInfo Lib "user32" _
  (ByVal hwndCombo As Long, _
   CBInfo As COMBOBOXINFO) As Long                          ' Gets info on a combobox and stores it in a COMBOBOXINFO

Private m_BalloonData As EDITBALLOONTIP             ' Stores the balloon data

Public Sub ShowBalloon(hTextBoxWnd As Long, sMessage As String, sTitle As String, bsStyle As BalloonStyle)
  ' Shows a balloon on a textbox
    If Not LicenseChange Then
        With m_BalloonData
            .cbStruct = Len(m_BalloonData)
            .pszTitle = StrConv(sTitle, vbUnicode)  ' Balloon title (Must be Unicode)
            .pszText = StrConv(sMessage, vbUnicode) ' Balloon text (Must be Unicode)
            .ttiIcon = bsStyle                      ' Balloon style
        End With
        
        SendMessage hTextBoxWnd, EM_SHOWBALLOONTIP, 0&, m_BalloonData   ' Shows the balloon
    End If
End Sub

Public Sub ShowComboBoxBalloon(hComboBoxWnd As Long, sMessage As String, sTitle As String, bsStyle As BalloonStyle)
  ' Shows a balloon on a combobox
  Dim CBI As COMBOBOXINFO, hEdit As Long
    If Not LicenseChange Then
        CBI.cbSize = Len(CBI)
        GetComboBoxInfo hComboBoxWnd, CBI   ' Gets info about the combobox
        hEdit = CBI.hwndEdit                ' Stores the textbox hWnd of a combobox in hEdit
        
        ShowBalloon hEdit, sMessage, sTitle, bsStyle    ' Show a balloon for the textbox in a combobox
    End If
End Sub

Private Sub UserControl_Initialize()
    InitCommonControls  ' Enable XP style
    
    SecureProgram       ' Se sub declaration for info
End Sub

Private Sub UserControl_Resize()
    UserControl.Width = 33 * Screen.TwipsPerPixelX      ' Makes the control fixed-size
    UserControl.Height = 33 * Screen.TwipsPerPixelY     ' Makes the control fixed-size
End Sub
