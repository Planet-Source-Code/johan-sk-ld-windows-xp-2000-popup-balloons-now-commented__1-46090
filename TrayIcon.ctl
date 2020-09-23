VERSION 5.00
Begin VB.UserControl TrayIcon 
   CanGetFocus     =   0   'False
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   495
   InvisibleAtRuntime=   -1  'True
   Picture         =   "TrayIcon.ctx":0000
   ScaleHeight     =   495
   ScaleWidth      =   495
   ToolboxBitmap   =   "TrayIcon.ctx":0D26
End
Attribute VB_Name = "TrayIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit         ' Requires Variable Declaration

Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias _
    "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Long   ' Sets and modifies an icon
    
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long    ' Gets the cursor position

Private Type POINTAPI
        X As Long
        Y As Long
End Type

Private Type NOTIFYICONDATA     ' This is only the structure of this type in 2000/XP or equalent
  cbSize As Long                ' If other operating system used this type is wrong declared
  hWnd As Long
  uID As Long
  uFlags As Long
  uCallbackMessage As Long
  hIcon As Long
  szTip As String * 128         ' * 64 on other operating system
  dwState As Long               ' Does not exist on other operating system
  dwStateMask As Long           ' Does not exist on other operating system
  szInfo As String * 256        ' Does not exist on other operating system
  uTimeoutOrVersion As Long     ' Does not exist on other operating system
  szInfoTitle As String * 64    ' Does not exist on other operating system
  dwInfoFlags As Long           ' Does not exist on other operating system
End Type

Private Const NIF_ICON = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_TIP = &H4
Private Const NIF_STATE = &H8
Private Const NIF_INFO = &H10

Private Const NIS_HIDDEN = &H1
Private Const NIS_SHAREDICON = &H2

Private Const NOTIFYICON_VERSION = &H1

Private Const NIM_ADD = &H0             ' Add icon
Private Const NIM_DELETE = &H2          ' Delete icon
Private Const NIM_MODIFY = &H1          ' Modify icon
Private Const NIM_SETFOCUS = &H4        ' Set focus to icon
Private Const NIM_SETVERSION = &H8      ' Set NotifyIcon version

Private Const WM_LBUTTONDBLCLK = &H203  ' Returned on Left Button Doubleclick
Private Const WM_LBUTTONDOWN = &H201    ' Returned on Left Button MouseDown
Private Const WM_LBUTTONUP = &H202      ' Returned on Left Button MouseUp
Private Const WM_MBUTTONDBLCLK = &H209  ' Returned on Middle Button Doubleclick
Private Const WM_MBUTTONDOWN = &H207    ' Returned on Middle Button MouseDown
Private Const WM_MBUTTONUP = &H208      ' Returned on Middle Button MouseUp
Private Const WM_RBUTTONDBLCLK = &H206  ' Returned on Right Button Doubleclick
Private Const WM_RBUTTONDOWN = &H204    ' Returned on Right Button MouseDown
Private Const WM_RBUTTONUP = &H205      ' Returned on Right Button MouseUp
Private Const WM_MOUSEMOVE = &H200      ' Returned on MouseMove

Private Const NIN_BALLOONSHOW = &H402       ' Returned on BalloonShow
Private Const NIN_BALLOONHIDE = &H403       ' Returned on BalloonHide or BalloonTimeout
Private Const NIN_BALLOONTIMEOUT = &H404    ' Returned on BalloonHide or BalloonTimeout
Private Const NIN_BALLOONUSERCLICK = &H405  ' Returned on BalloonClick

Private m_InTray As Boolean, m_IconData As NOTIFYICONDATA, _
    m_Icon As Long, m_sTip As String

Public Event MouseUp(Button As Integer, X As Single, Y As Single)       ' Event for MouseUp on systray icon
Public Event MouseDown(Button As Integer, X As Single, Y As Single)     ' Event for MouseDown on systray icon
Public Event MouseDblClick(Button As Integer, X As Single, Y As Single) ' Event for DoubleClick on systray icon
Public Event BalloonShow()                                              ' Event for BalloonShow
Public Event BalloonTimeoutOrHide()                                     ' Event for BalloonHide or BalloonTimeout
Public Event BalloonClick()                                             ' Event for BalloonClick

Public Property Get InTray() As Boolean                      ' Gets wheter the icon is shown
Attribute InTray.VB_Description = "Sets whether the icon in the systray should be visible or not."
Attribute InTray.VB_MemberFlags = "400"
    InTray = m_InTray
End Property

Public Property Let InTray(ByVal bNewValue As Boolean)       ' Sets wheter the icon is shown
    If Not LicenseChange Then
        If bNewValue Then
            UserControl.ScaleMode = vbTwips                  ' Set ScaleMode to Twips
            
            With m_IconData
                .cbSize = Len(m_IconData)
                .hIcon = m_Icon                              ' Sets icon
                .hWnd = UserControl.hWnd                     ' Sets hWnd of icon owner
                .szTip = m_sTip & vbNullChar                 ' Sets tooltip text
                .uCallbackMessage = WM_MOUSEMOVE             ' Sets where the callback messages will be sent
                .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP ' Sets that we want an icon, tooltip and messages to be sent
                .uID = 0&
            End With
            
            Shell_NotifyIcon NIM_ADD, m_IconData             ' Adds the icon
            Shell_NotifyIcon NIM_SETVERSION, m_IconData      ' Sets the NotifyIconData version
        Else
            Shell_NotifyIcon NIM_DELETE, m_IconData          ' Deletes the icon
        End If
        
        m_InTray = bNewValue
    End If
End Property

Public Property Get InfoTip() As String                      ' Gets infotip
Attribute InfoTip.VB_Description = "Sets the infotip for the icon in the systray"
Attribute InfoTip.VB_MemberFlags = "400"
    InfoTip = m_sTip
End Property

Public Property Let InfoTip(ByVal sNewValue As String)       ' Sets infotip
    If Not LicenseChange Then
        m_sTip = sNewValue
        
        If m_InTray Then
            m_IconData.szTip = sNewValue
            Shell_NotifyIcon NIM_MODIFY, m_IconData          ' Modifies the icon to contain the new tooltip
        End If
    End If
End Property

Public Property Get TrayIcon() As Long                       ' Gets the TrayIcon handle
Attribute TrayIcon.VB_Description = "Sets the handle for the icon that should appear in the systray."
Attribute TrayIcon.VB_MemberFlags = "400"
    TrayIcon = m_Icon
End Property

Public Property Let TrayIcon(ByVal lNewValue As Long)        ' Sets the TrayIcon
    If Not LicenseChange Then
        m_Icon = lNewValue
        
        If m_InTray Then
            m_IconData.hIcon = m_Icon
            Shell_NotifyIcon NIM_MODIFY, m_IconData          ' Modifies the icon to contain a new icon
        End If
    End If
End Property

Public Sub PopupBalloon(sMessage As String, sTitle As String, bStyle As BalloonStyle) ' Shows a balloon
    If Not LicenseChange Then
        If Not m_InTray Then
            Err.Raise 112, , "The InTray property must be set to True"
            Exit Sub
        End If
        
        With m_IconData
            .uFlags = NIF_INFO
            .uTimeoutOrVersion = 1                           ' Sets that we want to enable timeout
            .dwInfoFlags = bStyle                            ' Sets the buttonstyle
            .szInfo = sMessage                               ' Sets the message of the balloon
            .szInfoTitle = sTitle                            ' Sets the title of the balloon
        End With
        
        Shell_NotifyIcon NIM_MODIFY, m_IconData              ' Shows the balloon
    End If
End Sub

Private Sub UserControl_Initialize()
    SecureProgram       ' Se sub declaration for info
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim P As POINTAPI                                         ' POINTAPI to store the Cursor Position
    GetCursorPos P                                          ' Stores the Cursor position in P
    
    Select Case X / Screen.TwipsPerPixelX
      Case WM_LBUTTONUP                                     ' Looks for Left Button MouseUp
        RaiseEvent MouseUp(1, CSng(P.X), CSng(P.Y))
      Case WM_LBUTTONDOWN                                   ' Looks for Left Button MouseDown
        RaiseEvent MouseDown(1, CSng(P.X), CSng(P.Y))
      Case WM_LBUTTONDBLCLK                                 ' Looks for Left Button Doubleclick
        RaiseEvent MouseDblClick(1, CSng(P.X), CSng(P.Y))

      Case WM_RBUTTONUP                                     ' Looks for Right Button MouseUp
        RaiseEvent MouseUp(2, CSng(P.X), CSng(P.Y))
      Case WM_RBUTTONDOWN                                   ' Looks for Right Button MouseDown
        RaiseEvent MouseDown(2, CSng(P.X), CSng(P.Y))
      Case WM_RBUTTONDBLCLK                                 ' Looks for Right Button Doubleclick
        RaiseEvent MouseDblClick(2, CSng(P.X), CSng(P.Y))

      Case WM_MBUTTONUP                                     ' Looks for Middle Button MouseUp
        RaiseEvent MouseUp(4, CSng(P.X), CSng(P.Y))
      Case WM_MBUTTONDOWN                                   ' Looks for Middle Button MouseDown
        RaiseEvent MouseDown(4, CSng(P.X), CSng(P.Y))
      Case WM_MBUTTONDBLCLK                                 ' Looks for Middle Button Doubleclick
        RaiseEvent MouseDblClick(4, CSng(P.X), CSng(P.Y))

      Case NIN_BALLOONSHOW                                  ' Looks for BalloonShow
        RaiseEvent BalloonShow
      Case NIN_BALLOONHIDE                                  ' Looks for BalloonHide or BalloonTimeout
        RaiseEvent BalloonTimeoutOrHide
      Case NIN_BALLOONTIMEOUT                               ' Looks for BalloonHide or BalloonTimeout
        RaiseEvent BalloonTimeoutOrHide
      Case NIN_BALLOONUSERCLICK                             ' Looks for BalloonClick
        RaiseEvent BalloonClick
    End Select
End Sub

Private Sub UserControl_Resize()
    UserControl.Width = 33 * Screen.TwipsPerPixelX          ' Makes the control fixed-size
    UserControl.Height = 33 * Screen.TwipsPerPixelY         ' Makes the control fixed-size
End Sub

Private Sub UserControl_Terminate()
    Shell_NotifyIcon NIM_DELETE, m_IconData                 ' On UserControl_Terminate, delete the systray icon
End Sub
