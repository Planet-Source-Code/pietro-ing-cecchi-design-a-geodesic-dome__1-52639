VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Hlink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'USAGE
'===================== URL ===================== URL
'     Dim Hlink As New Hlink
'     Hlink.URL = "some URL"
'     Hlink.Maximized = True
'     Hlink.OpenURL
'===================== EMAIL ===================== Email
'     Dim Hlink As New Hlink
'     Hlink.URL = "some email"
'     Hlink.Maximized = True
'     Hlink.Mail

'===================== Hlink class===================== Hlink class
'**************************************
' Name: An easy to use Hypelink class
' Description:URL easy hendling. You don
'     't have to write all the long shell func
'     tion, only enter parameters
' By: Kamen
'
' Inputs:URL/mail, Browser window starti
'     ng mode
'
'This code is copyrighted and has' limited warranties.Please see http://w
'     ww.Planet-Source-Code.com/vb/scripts/Sho
'     wCode.asp?txtCodeId=10871&lngWId=1'for details.'**************************************

'The class is used to launch the browser


'     with the
    'parameters - URL/mailto, Maximized or M
    '     inimized,
    'or Normal;
    'Example :
    '
    'Private Sub cmdLaunch_Click()
    ' Dim Hlink As New clsHlink
    ' Hlink.URL = "http://www.someurl.com" '
    '     write the email address if you want to s
    '     end an email
    ' Hlink.Maximized = True
    ' Hlink.OpenURL
    'End Sub

Private Const SW_SHOWMAXIMIZED = 3
Private Const SW_SHOWMINIMIZED = 2
Private Const SW_SHOWNORMAL = 1

Private mstrURL As String
Private mlngMax As Integer
Private mlngMin As Integer
Private mlngNormal As Integer

'Run declares
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Property Let URL(strURL As String)
    mstrURL = strURL
End Property


Property Let Maximized(bMax As Boolean)


    If bMax Then
        mlngMax = SW_SHOWMAXIMIZED
    End If
End Property


Property Let Minimized(bMin As Boolean)


    If bMin Then
        mlngMin = SW_SHOWMINIMIZED
    End If
End Property


Property Let Normal(bNormal As Boolean)


    If bNormal Then
        mlngNormal = SW_SHOWNORMAL
    End If
End Property


Public Sub OpenURL()
    
    Call ShellExecute(0&, vbNullString, mstrURL, _
    vbNullString, "C:\", mlngMax Or mlngMin Or mlngNormal)
End Sub


Public Sub Mail()
    
    Call ShellExecute(0&, vbNullString, "mailto:" & mstrURL, _
    vbNullString, "C:\", mlngMax Or mlngMin Or mlngNormal)
End Sub






