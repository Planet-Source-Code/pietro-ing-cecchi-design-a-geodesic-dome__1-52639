Attribute VB_Name = "Module1"
Public SPHEREdiameter As Double
 
'to set form1.check2 to true, in form_load, without executing code
'to write into Form2.Text1, without executing code
Public DoNotExec As Boolean


Public MouseIsDown As Boolean

Public Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long

Public Type POINTAPI
        X As Long
        Y As Long
End Type

Public incx As Single
Public incy As Single
Public incz As Single
Public Const incdelta = 2 '5 '15

Public Type Point
    Name As String
    X As Double
    Y As Double
    Z As Double
End Type

Public Type Face
    Pt1 As Long
    Pt2 As Long
    Pt3 As Long
End Type

Public Type Triangle
    Pt1 As Long
    Pt2 As Long
    Pt3 As Long
End Type

Public m_Points() As Point
Public m_Faces() As Face
Public m_Triangles() As Triangle

Public Z1 As Double

Public facesnumber As Long


Public Sub FillListSpecs()
'fill list triangles specs
   scalingfact = SPHEREdiameter / (2 * Z1)
   Dim the_face As Face
    Form2.List1.Clear
    'fill list, sorting
    For i = 1 To UBound(m_Faces)
         GoSub DOMEcond
         If facing Then
            the_face = m_Faces(i)
            name1 = m_Points(the_face.Pt1).Name
            name2 = m_Points(the_face.Pt2).Name
            name3 = m_Points(the_face.Pt3).Name
            itemstring = "face_points " & name1 & "-" & name2 & "-" & name3
            itemstring = itemstring & Space(45 - Len(itemstring)) '& "SPECS: "
            len12 = Round(Sqr((m_Points(the_face.Pt1).X - m_Points(the_face.Pt2).X) ^ 2 + (m_Points(the_face.Pt1).Y - m_Points(the_face.Pt2).Y) ^ 2 + (m_Points(the_face.Pt1).Z - m_Points(the_face.Pt2).Z) ^ 2) * scalingfact, 2)
            len13 = Round(Sqr((m_Points(the_face.Pt1).X - m_Points(the_face.Pt3).X) ^ 2 + (m_Points(the_face.Pt1).Y - m_Points(the_face.Pt3).Y) ^ 2 + (m_Points(the_face.Pt1).Z - m_Points(the_face.Pt3).Z) ^ 2) * scalingfact, 2)
            len23 = Round(Sqr((m_Points(the_face.Pt2).X - m_Points(the_face.Pt3).X) ^ 2 + (m_Points(the_face.Pt2).Y - m_Points(the_face.Pt3).Y) ^ 2 + (m_Points(the_face.Pt2).Z - m_Points(the_face.Pt3).Z) ^ 2) * scalingfact, 2)
            itemstring = itemstring & " sides:" & " 12=" & len12 & " 13=" & len13 & " 23=" & len23
            Form2.List1.AddItem itemstring
       End If
    Next
    'number the list items
    For i = 0 To Form2.List1.ListCount - 1
         itemstring = Space(1) & Format(i + 1, "000") & Space(2)
         itemstring = itemstring & Form2.List1.List(i)
         Form2.List1.RemoveItem i
         Form2.List1.AddItem itemstring, i
    Next

Exit Sub

DOMEcond:
      If Form1.Check3 Then
         If UBound(m_Faces) = 20 Then 'icosahedron
            cond = (m_Points(m_Faces(i).Pt1).Z > 0) Or _
                   (m_Points(m_Faces(i).Pt2).Z > 0) Or _
                   (m_Points(m_Faces(i).Pt3).Z > 0)
            facing = cond
         Else
            cond = (m_Points(m_Faces(i).Pt1).Z < 0) Or _
                   (m_Points(m_Faces(i).Pt2).Z < 0) Or _
                   (m_Points(m_Faces(i).Pt3).Z < 0)
            facing = Not cond
         End If
      Else
         facing = True
      End If
Return

End Sub


Public Sub VerifyEquilateralSides()
   'check all sides length
   distmax = -1000
   distmin = 1000
    
   For i = 1 To UBound(m_Triangles)
      'each side of each triagle requires the max/min computation
      '1st triangle's side
      X = m_Points(m_Triangles(i).Pt2).X - m_Points(m_Triangles(i).Pt1).X
      Y = m_Points(m_Triangles(i).Pt2).Y - m_Points(m_Triangles(i).Pt1).Y
      Z = m_Points(m_Triangles(i).Pt2).Z - m_Points(m_Triangles(i).Pt1).Z
      dista = Sqr(X * X + Y * Y + Z * Z)
      If dista < distmin Then distmin = dista
      If dista > distmax Then distmax = dista
      '2nd triangle's side
      X = m_Points(m_Triangles(i).Pt3).X - m_Points(m_Triangles(i).Pt1).X
      Y = m_Points(m_Triangles(i).Pt3).Y - m_Points(m_Triangles(i).Pt1).Y
      Z = m_Points(m_Triangles(i).Pt3).Z - m_Points(m_Triangles(i).Pt1).Z
      dista = Sqr(X * X + Y * Y + Z * Z)
      If dista < distmin Then distmin = dista
      If dista > distmax Then distmax = dista
      '3rd triangle's side
      X = m_Points(m_Triangles(i).Pt3).X - m_Points(m_Triangles(i).Pt2).X
      Y = m_Points(m_Triangles(i).Pt3).Y - m_Points(m_Triangles(i).Pt2).Y
      Z = m_Points(m_Triangles(i).Pt3).Z - m_Points(m_Triangles(i).Pt2).Z
      dista = Sqr(X * X + Y * Y + Z * Z)
      If dista < distmin Then distmin = dista
      If dista > distmax Then distmax = dista
   Next
   'indicate this max/min and difference between the two on bottom button
On Error Resume Next
   scalingfact = SPHEREdiameter / (2 * Z1)
   Form2.Command2.Caption = "edges length: " & Round(distmin * scalingfact, 2) & " to " & Round(distmax * scalingfact, 2) & "  diff=" & Round((distmax - distmin) * scalingfact, 2) & " [" & Round((distmax - distmin) / distmin * 100, 1) & "%]"
On Error GoTo 0


End Sub

