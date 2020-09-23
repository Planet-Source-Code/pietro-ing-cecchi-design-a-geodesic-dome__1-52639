VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "faces dimensions"
   ClientHeight    =   3720
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   9900
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   9900
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "ENTER"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5520
      TabIndex        =   7
      ToolTipText     =   "updates values of list with new diameter"
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   2400
      TabIndex        =   5
      Top             =   720
      Width           =   3015
   End
   Begin VB.CommandButton SuckInFocus 
      Caption         =   "SuckInFocus"
      Height          =   375
      Left            =   5640
      TabIndex        =   4
      Top             =   0
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "min"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "minimize"
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0FFC0&
      Caption         =   "print list"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "on default printer"
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "verify edges lengths"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   7335
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2460
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   1200
      Width           =   9855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "SPHERE DIAMETER:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   480
      TabIndex        =   6
      ToolTipText     =   "whichever unit"
      Top             =   795
      Width           =   1845
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
 Me.WindowState = 1 'minimize
End Sub

Private Sub Command1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
SuckInFocus.SetFocus
End Sub

Private Sub Command2_Click()
 'display only
End Sub

Private Sub Command3_Click()
   SPHEREdiameter = Text1.Text
   Command3.Enabled = False
   Text1.BackColor = vbWindowBackground
   FillListSpecs
   VerifyEquilateralSides
   List1.Visible = True
   Command2.Visible = True
   MsgBox "List has been updated for sphere's diameter change", vbInformation, "LIST UPDATE"
End Sub

Private Sub Command7_Click()
 'print list
 PrinterFontNameSAVE = Printer.FontName
 Printer.FontName = "Courier New"
 For a = 0 To List1.ListCount - 1
    Printer.Print List1.List(a)
    If ((a + 1) Mod 60) = 0 Then Printer.NewPage
 Next
 Printer.EndDoc
 Printer.FontName = PrinterFontNameSAVE
End Sub

Private Sub Command7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
SuckInFocus.SetFocus
End Sub


Private Sub Form_Load()
   'send this out of screen
   SuckInFocus.Move -2 * SuckInFocus.Width
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = UnloadMode = 0
End Sub

Private Sub Form_Resize()
  Static busy As Boolean
  SuckInFocus.SetFocus
  If WindowState <> vbNormal Then Exit Sub
  If busy Then Exit Sub
  busy = True
  
  Select Case Width
     Case Is < 10020
        Width = 10020
     Case Is > 12000
        Width = 12000
     Case Else
        List1.Width = Width - List1.Left - 10 * Screen.TwipsPerPixelX
  End Select
  
  Select Case Height
     Case Is < 4110
        Height = 4110
     Case Is > 9000
        Height = 9000
     Case Else
        List1.Height = Height - List1.Top - 20 * Screen.TwipsPerPixelY
  End Select
  
  busy = False

End Sub

Private Sub Text1_Change()
   If DoNotExec Then Exit Sub
   Command3.Enabled = True
   Text1.BackColor = vbYellow '&H8080C0
   List1.Visible = False
   Command2.Visible = False
End Sub
