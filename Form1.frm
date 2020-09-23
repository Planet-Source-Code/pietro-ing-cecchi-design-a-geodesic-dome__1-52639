VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "geodesic dome-sphere"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   568
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   794
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "show dome"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   10080
      TabIndex        =   17
      ToolTipText     =   "half sphere (tension ring at bottom needed)"
      Top             =   6720
      Value           =   1  'Checked
      Width           =   1455
   End
   Begin VB.CommandButton help 
      BackColor       =   &H0000FF00&
      Caption         =   "help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10440
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "micro-help"
      Top             =   2760
      Width           =   735
   End
   Begin VB.CommandButton SuckInFocus 
      Caption         =   "SuckInFocus"
      Height          =   375
      Left            =   7800
      TabIndex        =   15
      Top             =   4800
      Width           =   1935
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00C0FFC0&
      Caption         =   "print sphere"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "as shown, on default printer"
      Top             =   7320
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "triangles data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "build a geodesic sphere"
      Top             =   7800
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFFC0&
      Caption         =   "visit geodesic sphere site (Rod_Stephens)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "must connect to internet"
      Top             =   480
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "vote on Planet  (Pietro_Cecchi)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "must connect to internet"
      Top             =   1440
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "tile more"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4200
      Width           =   1695
   End
   Begin VB.Timer MouseTimer 
      Interval        =   60
      Left            =   9120
      Top             =   5520
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   9960
      ScaleHeight     =   2145
      ScaleWidth      =   1665
      TabIndex        =   1
      Top             =   4920
      Width           =   1695
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "rotate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "cw=click left   ccw=click right  (spin=hold down)"
         Top             =   960
         Width           =   1455
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "show points"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1440
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "cw"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   960
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "ccw"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   720
         TabIndex        =   5
         Top             =   960
         Width           =   615
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "around its     Z"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "around its     0  Y"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "around its   X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Value           =   1  'Checked
         Width           =   1575
      End
   End
   Begin VB.CommandButton reset 
      BackColor       =   &H00C0E0FF&
      Caption         =   "icosahedron"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   ". icosahedron ."
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7935
      Left            =   9840
      TabIndex        =   14
      Top             =   360
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'GEODESIC domes, Copyright 2004 by Pietro Cecchi, pietrocecchi@inwind.it
'credits to Rod Stephens, www.vb-helper.com
'posted on www.planet-source-code.com on march 25th 2004
'at http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=52639&lngWId=1
'
'pure VB6, no DirectX
'
'Geodesic Sphere is a graphical application that shows you
'a polyhedron which, as the faces increase, approaches a sphere.
'The basic polyhedron is the icosahedron (20 equal equilateral triangles
'as faces).
'This application is based on an application downloaded from
'at http://www.vb-helper.com/howto_geodesic_sphere.html done by Rod Stephens
'which performs all the calculations for the sphere points in 3D.
'Rotations algorithms are web page http://skal.planet-d.net/demo/maths.html ,paragraph '3D and Perpective transformations'.
'My contribute has been essentially the addition of code for 3D display.
'
'This application should meet the needs and curiosity of many interested
'in geodesic building, such architects, civilian engineers and geometry
'students.
'
'Sorry that the rotation of the sphere happens around its own axes and not
'around the fixed ones (the ones of the observer).
'This is the reason why this is not a 3D virtual reality application.
'I hope to do it in the next future.
'
'NOTE that most of this code may be subject to Copyrights by the many
'who worked in many ways around this application.
'To mention at least the most important of them, here follows the list:
'Pietro Cecchi, pietrocecchi@inwind.it
'Rod Stephens, feedback@vb-helper.com
'
'I do have the posting permission from Rod Stephens, of course.
'
'Even if most credits go to Rod Stephens, I will appreciate if you vote
'for me (or let's say: us) on Planet-source-code.
'
'Thanks everybody and bye
'Pietro Cecchi
'
'KEYWORDS: tensegrity, tension, structures, triangles, tetrahedrons, icosahedron,
'sphere, 3D, virtual reality, Platon, geometry, graphics, poligons, building, domes
'stadiums, botanic, geodesic, homes




Private Function FindPointByName(ByVal point_name As String) As Integer
Dim i As Integer
    For i = 1 To UBound(m_Points)
        If m_Points(i).Name = point_name Then
            FindPointByName = i
            Exit For
        End If
    Next i
End Function

' Create a geodesic tiling for this face.
Private Sub TileFace(the_face As Face)
   Const LEVEL = 2

   Dim x0 As Double
   Dim y0 As Double
   Dim z0 As Double
   Dim v1x As Double
   Dim v1y As Double
   Dim v1z As Double
   Dim v2x As Double
   Dim v2y As Double
   Dim v2z As Double
   Dim base_name As String
   Dim p10 As Long
   Dim p01 As Long
   Dim p11 As Long

    ' Get the Face's origin.
    With m_Points(the_face.Pt1)
        x0 = .X
        y0 = .Y
        z0 = .Z
    End With

    ' Get the tiling's generating vectors.
    v1x = (m_Points(the_face.Pt2).X - m_Points(the_face.Pt1).X) / LEVEL
    v1y = (m_Points(the_face.Pt2).Y - m_Points(the_face.Pt1).Y) / LEVEL
    v1z = (m_Points(the_face.Pt2).Z - m_Points(the_face.Pt1).Z) / LEVEL
    v2x = (m_Points(the_face.Pt3).X - m_Points(the_face.Pt1).X) / LEVEL
    v2y = (m_Points(the_face.Pt3).Y - m_Points(the_face.Pt1).Y) / LEVEL
    v2z = (m_Points(the_face.Pt3).Z - m_Points(the_face.Pt1).Z) / LEVEL
   
    ' Generate the points.
    'make names of points unique
    name10 = IIf(m_Points(the_face.Pt1).Name < m_Points(the_face.Pt2).Name, m_Points(the_face.Pt1).Name & m_Points(the_face.Pt2).Name, m_Points(the_face.Pt2).Name & m_Points(the_face.Pt1).Name)
    name01 = IIf(m_Points(the_face.Pt1).Name < m_Points(the_face.Pt3).Name, m_Points(the_face.Pt1).Name & m_Points(the_face.Pt3).Name, m_Points(the_face.Pt3).Name & m_Points(the_face.Pt1).Name)
    name11 = IIf(m_Points(the_face.Pt2).Name < m_Points(the_face.Pt3).Name, m_Points(the_face.Pt2).Name & m_Points(the_face.Pt3).Name, m_Points(the_face.Pt3).Name & m_Points(the_face.Pt2).Name)
    p10 = MakeNormalizedPoint(name10, x0 + v1x, y0 + v1y, z0 + v1z)
    p01 = MakeNormalizedPoint(name01, x0 + v2x, y0 + v2y, z0 + v2z)
    p11 = MakeNormalizedPoint(name11, x0 + v1x + v2x, y0 + v1y + v2y, z0 + v1z + v2z)

    ' Make the triangles.
    MakeTriangle the_face.Pt1, p10, p01
    MakeTriangle p10, p11, p01
    MakeTriangle p10, the_face.Pt2, p11
    MakeTriangle p01, p11, the_face.Pt3

End Sub

Private Sub TileMore()
    'this routine makes a polyhedron with 4 times the original faces
    'id est, each time you call this routine,
    'starting with 20 faces (icosahedron): 80 faces, 320 faces, 1280 faces and so on
    
    'remake faces from existent triangles
    ReDim m_Faces(0)
    For i = 1 To UBound(m_Triangles)
        MakeFace m_Points(m_Triangles(i).Pt1).Name, m_Points(m_Triangles(i).Pt2).Name, m_Points(m_Triangles(i).Pt3).Name
    Next
    'tile triangles on all faces
    ReDim m_Triangles(0)
    ' Make geodesic tilings of the faces.
    For i = 1 To UBound(m_Faces)
        TileFace m_Faces(i)
    Next i

    'remake faces from existent triangles
    ReDim m_Faces(0)
    For i = 1 To UBound(m_Triangles)
        MakeFace m_Points(m_Triangles(i).Pt1).Name, m_Points(m_Triangles(i).Pt2).Name, m_Points(m_Triangles(i).Pt3).Name
    Next

End Sub

Private Sub Check1_Click(Index As Integer)
   'around sphere's axes wanted rotations (used in Drawsphere routine)
   'no code is needed here
End Sub

Private Sub Check2_Click()
   If DoNotExec Then Exit Sub
   'show polyhedron's points on graph
   DrawSphere True 'True=no rotation
End Sub

Private Sub Check3_Click()
   If DoNotExec Then Exit Sub
   'show polyhedron's points on graph
   DrawSphere True 'True=no rotation
   FillListSpecs
End Sub

Private Sub Command2_Click()
   Form2.WindowState = 0 'normal
   Form2.Show
   Form2.ZOrder
End Sub

Private Sub Command3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
SuckInFocus.SetFocus
End Sub

Private Sub Command7_Click()
'   SavePicture Me.Image, App.Path & "\screenshot.bmp"
   PrintForm
End Sub

Private Sub Command7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
SuckInFocus.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Unload Form2
End Sub

Private Sub help_Click()
   msg = "GEODESIC STRUCTURES: not seismic, strong, light, cheap" & vbNewLine & vbNewLine & _
         "Click on '20 faces' for 'icosahedron', on 'tile' for more faces." & vbNewLine & _
         "Click on 'Triangles' to read triangles data: list of all faces and its sides lengths [side 12 means from first point of face to the second one]." & vbNewLine & _
         "Click on 'Rotate' to rotate the polyhedron, on 'Print sphere' to print out the screen." & vbNewLine & _
         "Check 'show points' to show points." & vbNewLine & _
         "Check 'show dome' to show half sphere [in the case you build the dome, don't forget you need a tension ring/cable at the base of it]." & vbNewLine & vbNewLine & _
         "Don't forget to visit vb-helper site and planet-source-code to vote for this application." & vbNewLine & vbNewLine & _
         "thanks everybody, Pietro and Rod"
   
   MsgBox msg, vbInformation, "micro-help"
End Sub

Private Sub help_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
SuckInFocus.SetFocus
End Sub

Private Sub reset_Click()
   '20 faces, icosahedron
   ReDim m_Faces(0)
   ReDim m_Triangles(0)
   Command3.Enabled = True
   Picture1.Enabled = True

   Command1.Enabled = True

   Dim pi As Double
   pi = Atn(1) * 4 ' 3,14159265358979 '3.14159265

   Dim S As Double
   S = 100! 'icosahedron side length

   Dim t1 As Double
   Dim t2 As Double
   Dim t3 As Double
   Dim t4 As Double
   Dim R As Double
   Dim H As Double
   Dim Cx As Double
   Dim Cy As Double
   Dim H1 As Double
   Dim H2 As Double
   Dim Z2 As Double

   Dim i As Long

    
    'let's calculate the vertexes of the icosahedron
    'that also are points of the sphere
    'see URL: http://www.vb-helper.com/tutorial_platonic_solids.html
    'where all calculations are clearly explained
    
    ' Calculate intermediate values.
    t1 = 2 * pi / 5!
    t2 = pi / 10!
    t4 = pi / 5!
    t3 = -3 * pi / 10!
    R = (S / 2) / Sin(t4)
    H = Cos(t4) * R
    Cx = R * Cos(t2)
    Cy = R * Sin(t2)
    Z2 = 1!
    H1 = Sqr(S * S - R * R)
    H2 = Sqr((H + R) * (H + R) - H * H)
    Z2 = (H2 - H1) / 2!
    Z1 = Z2 + H1
    
    SPHEREdiameter = 200   'the real one is 2 * Z1
    DoNotExec = True
       Form2.Text1.Text = SPHEREdiameter
    DoNotExec = False
    Form2.Command3.Enabled = False

    ' Create the points.
    'see URL: http://www.vb-helper.com/tutorial_platonic_solids.html
    MakePoint "a", 0, 0, Z1
    MakePoint "b", 0, R, Z2
    MakePoint "c", Cx, Cy, Z2
    MakePoint "d", S / 2, -H, Z2
    MakePoint "e", -S / 2, -H, Z2
    MakePoint "f", -Cx, Cy, Z2
    MakePoint "g", 0, -R, -Z2
    MakePoint "h", -Cx, -Cy, -Z2
    MakePoint "i", -S / 2, H, -Z2
    MakePoint "j", S / 2, H, -Z2
    MakePoint "k", Cx, -Cy, -Z2
    MakePoint "l", 0, 0, -Z1

    ' Make the outwardly-oriented faces.
    MakeFace "a", "c", "b"
    MakeFace "a", "b", "f"
    MakeFace "a", "f", "e"
    MakeFace "a", "e", "d"
    MakeFace "a", "d", "c"
    MakeFace "c", "d", "k"
    MakeFace "c", "k", "j"
    MakeFace "c", "j", "b"
    MakeFace "b", "j", "i"
    MakeFace "b", "i", "f"
    MakeFace "f", "i", "h"
    MakeFace "f", "h", "e"
    MakeFace "e", "h", "g"
    MakeFace "e", "g", "d"
    MakeFace "d", "g", "k"
    MakeFace "l", "j", "k"
    MakeFace "l", "k", "g"
    MakeFace "l", "g", "h"
    MakeFace "l", "h", "i"
    MakeFace "l", "i", "j"

'originally it was:
    ' Make geodesic tilings of the faces.
'    For i = 1 To UBound(m_Faces)
'        TileFace m_Faces(i)
'    Next i
    
    'instead of the above tiling
    '(that would have shown 80 faces),
    'I preferred to show
    'the very basic icosahedron
    For i = 1 To UBound(m_Faces)
         MakeTriangle m_Faces(i).Pt1, m_Faces(i).Pt2, m_Faces(i).Pt3
    Next
    
    DrawSphere True  'True = no rotation

    'added
    facesnumber = UBound(m_Triangles)
    reset.Caption = facesnumber & " faces"
    Command3.Caption = "tile more [now " & UBound(m_Triangles) & " faces]"
    
    FillListSpecs
    
End Sub


'added
Private Sub DrawSphere(Optional ByVal NoRotation = False)
   'xyz space:  z
   '            0  y
   '           x
   'form plane:  Y
   '             0 X
   '
   Dim P1 As Point, P2 As Point, P3 As Point
   
   'clear form
   Cls
   'watermark
   FontSize = 25
   ForeColor = &HF0F0F0    'gray &H00C0C0C0&
   isstring = " Geodesic Sphere       "
   For a = 1 To 16
      Print isstring; isstring
   Next
    
   'i accomplish rotation, if allowed, incrementing by inc coordinates
   If Not NoRotation Then
      If Check1(0) Then incx = (incx + incdelta * IIf(Option1(0), 1, -1)) Mod 360
      If Check1(1) Then incy = (incy + incdelta * IIf(Option1(0), 1, -1)) Mod 360
      If Check1(2) Then incz = (incz + incdelta * IIf(Option1(0), 1, -1)) Mod 360
   End If
    
   'rotation around sphere's axes x, y and z
   alfa_a = incx
   alfa_b = incy
   alfa_c = incz
    
   Dim cos_a As Double, cos_b As Double, cos_c As Double, sin_a As Double, sin_b As Double, sin_c As Double
   cos_a = Cos(radiants(alfa_a))
   cos_b = Cos(radiants(alfa_b))
   cos_c = Cos(radiants(alfa_c))
   sen_a = Sin(radiants(alfa_a))
   sen_b = Sin(radiants(alfa_b))
   sen_c = Sin(radiants(alfa_c))
    
   VerifyEquilateralSides
   

   'draw and rotate
   For i = 1 To UBound(m_Triangles)
        
      P1 = m_Points(m_Triangles(i).Pt1)
      P2 = m_Points(m_Triangles(i).Pt2)
      P3 = m_Points(m_Triangles(i).Pt3)
        
  
'this from skal.planet-d.net\index.html
'
'                       cz.Cy , cz.sy.sx - sz.Cx, cz.sy.Cx + sz.sx
'                       sz.Cy , sz.sy.sx + cz.Cx, sz.sy.Cx - cz.sx
'                         -sy,          cy.sx,          cy.cx
'    alfa_a = incx
'    alfa_b = incy
'    alfa_c = incz
'the good one (ma ruota - correttamente - sugli assi della sfera) <-- CHOOSE THIS ONE
'        .X = PX * (cos_c * cos_b) + PY * (cos_c * sen_b * sen_a - cos_a * sen_c) + PZ * (cos_c * sen_b * cos_a + sen_a * sen_c)
'        .Y = PX * (sen_c * cos_b) + PY * (sen_c * sen_b * sen_a + cos_c * cos_a) + PZ * (sen_c * sen_b * cos_a - cos_c * sen_a)
'        .Z = PX * (-sen_b) + PY * (cos_b * sen_a) + PZ * (cos_b * cos_a)
  
  
      
      Dim PX As Double, PY As Double, PZ As Double
  
      With P1
         PX = .X: PY = .Y: PZ = .Z
         .X = PX * (cos_c * cos_b) + PY * (cos_c * sen_b * sen_a - cos_a * sen_c) + PZ * (cos_c * sen_b * cos_a + sen_a * sen_c)
         .Y = PX * (sen_c * cos_b) + PY * (sen_c * sen_b * sen_a + cos_c * cos_a) + PZ * (sen_c * sen_b * cos_a - cos_c * sen_a)
         .Z = PX * (-sen_b) + PY * (cos_b * sen_a) + PZ * (cos_b * cos_a)
      End With
      With P2
         PX = .X: PY = .Y: PZ = .Z
         .X = PX * (cos_c * cos_b) + PY * (cos_c * sen_b * sen_a - cos_a * sen_c) + PZ * (cos_c * sen_b * cos_a + sen_a * sen_c)
         .Y = PX * (sen_c * cos_b) + PY * (sen_c * sen_b * sen_a + cos_c * cos_a) + PZ * (sen_c * sen_b * cos_a - cos_c * sen_a)
         .Z = PX * (-sen_b) + PY * (cos_b * sen_a) + PZ * (cos_b * cos_a)
      End With
      With P3
         PX = .X: PY = .Y: PZ = .Z
         .X = PX * (cos_c * cos_b) + PY * (cos_c * sen_b * sen_a - cos_a * sen_c) + PZ * (cos_c * sen_b * cos_a + sen_a * sen_c)
         .Y = PX * (sen_c * cos_b) + PY * (sen_c * sen_b * sen_a + cos_c * cos_a) + PZ * (sen_c * sen_b * cos_a - cos_c * sen_a)
         .Z = PX * (-sen_b) + PY * (cos_b * sen_a) + PZ * (cos_b * cos_a)
      End With


      'fastest way to color the faces,
      'giving the illusion of 3D
      'is based on the X coordinate of the 3 points of each triangle
      Dim xx1 As Long, xx2 As Long, xx3 As Long
      
'different shadowing
'
'if you use these 3 then you think 'the sphere' rotates while you look at it
'dinamic shadowing
      xx1 = 2 / 3 * (P1.Y - P1.Z)
      xx2 = 2 / 3 * (P2.Y - P2.Z)
      xx3 = 2 / 3 * (P3.Y - P3.Z)
'INSTEAD, if you use these 3 then you think YOU go around the sphere
'solidal shadowing
'      xx1 = 2 / 3 * (m_Points(m_Triangles(i).Pt1).Y - m_Points(m_Triangles(i).Pt1).Z)
'      xx2 = 2 / 3 * (m_Points(m_Triangles(i).Pt2).Y - m_Points(m_Triangles(i).Pt2).Z)
'      xx3 = 2 / 3 * (m_Points(m_Triangles(i).Pt3).Y - m_Points(m_Triangles(i).Pt3).Z)
      
      
      'draw only positive plane, IMPROVE (draw only seen)
      lcolor = RGB(Abs(xx1 - 150), Abs(xx2 - 150), Abs(xx3 - 150))
      If m_Points(m_Triangles(i).Pt1).Name = "a" Then
         'exception:
         'Yellow pentagon (made up of 5 triangles, starting points of icosahedron construction)
         'may like to see it on the sphere
         lcolor = &HC0FFFF    'pale yellow &H00C0FFFF& 'vbYellow
      End If
      Me.FillStyle = 0 '0 solid, 1 transparent
      Me.FillColor = lcolor
        
      'compute points vector for API Polygon
      Dim scalefactx As Double, scalefacty As Double
      scalefacty = 280!
      scalefactx = 0.85 * scalefacty * Abs(ScaleWidth / ScaleHeight)
      fact = 2.8
      
      Dim PNTpoly(1 To 3) As POINTAPI
      'Y of the sphere becomes X of the screen
      '-Z of the sphere becomes Y of the screen
      PNTpoly(1).X = P1.Y * fact + scalefactx
      PNTpoly(1).Y = -P1.Z * fact + scalefacty
      PNTpoly(2).X = P2.Y * fact + scalefactx
      PNTpoly(2).Y = -P2.Z * fact + scalefacty
      PNTpoly(3).X = P3.Y * fact + scalefactx
      PNTpoly(3).Y = -P3.Z * fact + scalefacty
        
      'a tricky way to decide if the face has to be drawn or not
      'only facing faces need to be drawn
      facing = (P1.X >= 0) And (P2.X >= 0) And (P3.X >= 0)
      If Not facing Then
         mean = P1.X + P2.X + P3.X
         facing = mean > 0
      End If
      
      'case dome, show only the upper part of polyhedron
      If Check3 Then
         If UBound(m_Faces) = 20 Then 'icosahedron
            cond = (m_Points(m_Triangles(i).Pt1).Z > 0) Or _
                   (m_Points(m_Triangles(i).Pt2).Z > 0) Or _
                   (m_Points(m_Triangles(i).Pt3).Z > 0)
            facing = facing And cond
         Else
            cond = (m_Points(m_Triangles(i).Pt1).Z < 0) Or _
                   (m_Points(m_Triangles(i).Pt2).Z < 0) Or _
                   (m_Points(m_Triangles(i).Pt3).Z < 0)
            facing = facing And (Not cond)
         End If
      End If
      
      If facing Then
         ForeColor = vbBlack
         'draws the face, id est the triangle
         Polygon Me.hdc, PNTpoly(1), 3
         If Check2 Then
            'print points names at location
            FontSize = 12
            ForeColor = vbRed

            CurrentX = P1.Y * fact + scalefactx
            CurrentY = -P1.Z * fact + scalefacty
            Me.Print m_Points(m_Triangles(i).Pt1).Name
            CurrentX = P2.Y * fact + scalefactx
            CurrentY = -P2.Z * fact + scalefacty
            Me.Print m_Points(m_Triangles(i).Pt2).Name
            CurrentX = P3.Y * fact + scalefactx
            CurrentY = -P3.Z * fact + scalefacty
            
            
            Me.Print m_Points(m_Triangles(i).Pt3).Name
         End If
      End If
   Next
End Sub

Private Function radiants(ByVal alfa_degrees As Double) As Double
   Dim pi As Double
   pi = Atn(1) * 4
   radiants = 2! * pi * alfa_degrees / 360!
End Function


' Add an Face to the m_Faces array.
Private Sub MakeFace(ByVal name1 As String, ByVal name2 As String, ByVal name3 As String)
   Dim num_faces As Long 'Integer

    ' Make room.
    On Error Resume Next
    num_faces = UBound(m_Faces)
    On Error GoTo 0
    num_faces = num_faces + 1
    ReDim Preserve m_Faces(0 To num_faces)

    ' Make the Face
    With m_Faces(num_faces)
        .Pt1 = FindPointByName(name1)
        .Pt2 = FindPointByName(name2)
        .Pt3 = FindPointByName(name3)
    End With
End Sub

' Add a Triangle to the m_Triangles array.
Private Sub MakeTriangle(ByVal Pt1 As Integer, ByVal Pt2 As Integer, ByVal Pt3 As Integer)
   Dim num_triangles As Long 'Integer

    ' Make room.
    On Error Resume Next
    num_triangles = UBound(m_Triangles)
    On Error GoTo 0
    num_triangles = num_triangles + 1
    ReDim Preserve m_Triangles(0 To num_triangles)

    ' Make the Triangle
    With m_Triangles(num_triangles)
        .Pt1 = Pt1
        .Pt2 = Pt2
        .Pt3 = Pt3
    End With
End Sub

' Add a point to the m_Points array.
Private Function MakePoint(ByVal new_name As String, ByVal X As Single, ByVal Y As Single, ByVal Z As Single) As Long 'Integer
   Dim num_points As Long 'Integer

    ' Make room.
    On Error Resume Next
    num_points = UBound(m_Points)
    On Error GoTo 0
    num_points = num_points + 1
    ReDim Preserve m_Points(1 To num_points)

    ' Create the point.
    With m_Points(num_points)
        .Name = new_name
        .X = X
        .Y = Y
        .Z = Z
    End With

    MakePoint = num_points
End Function

' Add a point to the m_Points array.
Private Function MakeNormalizedPoint(ByVal new_name As String, ByVal X As Single, ByVal Y As Single, ByVal Z As Single) As Long 'Integer
   Dim dist As Double 'Single

    ' Make the point distance Z1 from the origin.
    dist = Sqr(X * X + Y * Y + Z * Z)
    
'    k = 1
'    a = k: b = k: c = k
'    dist = Sqr((X * X) / (a * a) + (Y * Y) / (b * b) + (Z * Z) / (c * c))

    X = X / dist * Z1
    Y = Y / dist * Z1
    Z = Z / dist * Z1

    ' Make the point.
    MakeNormalizedPoint = MakePoint(new_name, X, Y, Z)
End Function

Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   'rotate
   Select Case Button
      Case 1
         Option1(0) = True
      Case 2
         Option1(1) = True
   End Select
   Option1(0).Refresh
   Option1(1).Refresh

   DrawSphere
   
   MouseIsDown = True
   MouseTimer.Enabled = True
End Sub

Private Sub Command1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   'rotate
   MouseIsDown = False
   MouseTimer.Enabled = False
SuckInFocus.SetFocus
End Sub

Private Sub Command3_Click()
   'tile more faces
   TileMore
   DrawSphere True  'True = no rotation
   If UBound(m_Faces) >= 20 * 4 * 4 Then Command3.Enabled = False
   Command3.Caption = "tile more [now " & UBound(m_Faces) & " faces]"
   FillListSpecs
End Sub

Private Sub Command4_Click()
   'vote this customized geodesic sphere on Planet-source-code
   Dim Hlink As New Hlink
   Hlink.URL = "http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=52639&lngWId=1"
   Hlink.Maximized = True
   Hlink.OpenURL
End Sub

Private Sub Command5_Click()
   'Geodesic Sphere site URL
   'must visit to understand the matter
   Dim Hlink As New Hlink
   Hlink.URL = "http://www.vb-helper.com/howto_geodesic_sphere.html"
   Hlink.Maximized = True
   Hlink.OpenURL
End Sub

Private Sub Form_Load()
   'initial values
   AutoRedraw = True
   BackColor = &HE0E0E0   'light grey &H00E0E0E0&
   reset.Caption = "start"
   Command1.Enabled = False
   Command3.Caption = "tile more faces"
   Command3.Enabled = False
   Picture1.Enabled = False
   WindowState = vbNormal
   Width = 12000
   Height = 9000 - 1 * Screen.TwipsPerPixelY
 With Screen
   Move (.Width - 12000) / 2, (.Height - 9000) / 2
 End With
 
   'send this out of screen
   SuckInFocus.Move -2 * SuckInFocus.Width
   
   Show
   
   'some other initializations
   Check1(0) = vbUnchecked 'no rotation around x axis
   Check1(1) = vbUnchecked 'no rotation around y axis
   Check1(2) = vbChecked 'rotation preset around Z axis
   
   DoNotExec = True
      Check2 = vbChecked 'show points
      Check3 = vbChecked 'show dome (half sphere, to build a geodesic dome: DO NOT FORGET a tension ring/cable at the bottom)
   DoNotExec = False
   
      
   reset_Click     'draw icosahedron
   Command3_Click  'tile more (to 80 faces)
'   Command3_Click  'tile more (to 320 faces)
   
SuckInFocus.SetFocus
   
End Sub

Private Sub MouseTimer_Timer()
   If MouseIsDown Then
      DrawSphere
   End If
End Sub

Private Sub reset_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
SuckInFocus.SetFocus

End Sub
