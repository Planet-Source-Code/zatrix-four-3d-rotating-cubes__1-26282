VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls3dObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Const PI = 3.141593
Const PS_SOLID = 0

'values used for rendering
Dim HALF_SCREEN_WIDTH As Long
Dim HALF_SCREEN_HEIGHT As Long
Dim HPC As Long
Dim VPC As Long
Dim ASPECT_COMP As Long

'object declaretion
Private obj3dObject As Object3D
Private Render As PictureBox

'API DECLARATIONS
Private Declare Function PolyDraw Lib "gdi32" (ByVal hdc As Long, lppt As POINTAPI, lpbTypes As Byte, ByVal cCount As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long

'type used to store which points make up a triangle
Private Type Triplet
    First As Long
    Second As Long
    Third As Long
End Type

'type to store a 3d point
Private Type Point3d
    X As Double
    Y As Double
    Z As Double
End Type

'type to store a 2d point
Private Type Point2d
    X As Double
    Y As Double
End Type

'data type for storing a 3d object
Private Type Object3D
    Name As String
    Version As String
    NumVertices As Long
    NumTriangles As Long
    Xangle As Long
    Yangle As Long
    Zangle As Long
    ScaleFactor As Double
    CenterofWorld As Point3d
    LocalCoord() As Point3d
    RotatedLocalCoord() As Point3d
    WorldCoord() As Point3d
    CameraCoord() As Point3d
    Triangle() As Triplet
    ScreenCoord() As Point2d
    Isvisible() As Boolean
    Color() As Long
End Type

'type used for calulating the normal for each face
Private Type Face
    Y As Double
    X As Double
End Type

'API type for use when making object solid
Private Type POINTAPI
        X As Long
        Y As Long
End Type
Private Sub CalculateNormals()

    Dim lngIncr As Long
    
    'formula for calculation
    '( face(0).y - face(2).y ) * ( face(1).x - face(0).x ) - ( face(0).x - face(2).x ) * ( face(1).y - face(0).y )
    
    Dim ObjectFace(0 To 2) As Face
    
        For lngIncr = 0 To obj3dObject.NumTriangles - 1
            
            ObjectFace(0).X = obj3dObject.ScreenCoord(obj3dObject.Triangle(lngIncr).First).X
            ObjectFace(0).Y = obj3dObject.ScreenCoord(obj3dObject.Triangle(lngIncr).First).Y
            ObjectFace(1).X = obj3dObject.ScreenCoord(obj3dObject.Triangle(lngIncr).Second).X
            ObjectFace(1).Y = obj3dObject.ScreenCoord(obj3dObject.Triangle(lngIncr).Second).Y
            ObjectFace(2).X = obj3dObject.ScreenCoord(obj3dObject.Triangle(lngIncr).Third).X
            ObjectFace(2).Y = obj3dObject.ScreenCoord(obj3dObject.Triangle(lngIncr).Third).Y
            
            If ((ObjectFace(0).Y - ObjectFace(2).Y) * (ObjectFace(1).X - ObjectFace(0).X)) - _
               ((ObjectFace(0).X - ObjectFace(2).X) * (ObjectFace(1).Y - ObjectFace(0).Y)) > 0 Then
                obj3dObject.Isvisible(lngIncr) = True
            Else
                obj3dObject.Isvisible(lngIncr) = False
            End If
            
        Next

End Sub


Public Sub SetRotations(Optional X As Double, Optional Y As Double, Optional Z As Double)

    If Not (IsMissing(X)) Then
        obj3dObject.Xangle = X
    End If
    
    If Not (IsMissing(Y)) Then
        obj3dObject.Yangle = Y
    End If
    
    If Not (IsMissing(Z)) Then
        obj3dObject.Zangle = Z
    End If

End Sub


Public Sub SetTranslations(Optional XPos As Variant, Optional YPos As Variant, Optional ZPos As Variant)

    If Not (IsMissing(XPos)) Then
        obj3dObject.CenterofWorld.X = XPos
    End If
    
    If Not (IsMissing(YPos)) Then
        obj3dObject.CenterofWorld.Y = YPos
    End If
    
    If Not (IsMissing(ZPos)) Then
        obj3dObject.CenterofWorld.Z = ZPos
    End If

End Sub


Public Sub LoadObject(strFileName As String, DeviceContext As PictureBox, lngCenterofWorldX As Double, lngCenterofWorldY As Double, lngCenterofWorldZ As Double, dblScaleFactor As Double, lngSetXRotation As Long, lngSetYRotation As Long, lngSetZRotation As Long)

    Dim strTemp As String
    Dim lngNumTemp As Long
    Dim lngNumVertices As Long
    Dim lngNumTriangles As Long
    
    'get values needed for rendering calculations
    Set Render = DeviceContext
    HALF_SCREEN_HEIGHT = Render.ScaleHeight / 2
    HALF_SCREEN_WIDTH = Render.ScaleWidth / 2
    ASPECT_COMP = (Render.ScaleHeight) / ((Render.ScaleWidth * 3) / 4)
    HPC = HALF_SCREEN_WIDTH / (Tan((60 / 2) * (PI / 180)))
    VPC = HALF_SCREEN_HEIGHT / (Tan((60 / 2) * (PI / 180)))
    
    'set initial stuff for object
    obj3dObject.CenterofWorld.X = lngCenterofWorldX
    obj3dObject.CenterofWorld.Y = lngCenterofWorldY
    obj3dObject.CenterofWorld.Z = lngCenterofWorldZ
    obj3dObject.ScaleFactor = dblScaleFactor
    obj3dObject.Xangle = lngSetXRotation
    obj3dObject.Yangle = lngSetYRotation
    obj3dObject.Zangle = lngSetZRotation
    
    'get handle to file
    Open strFileName For Input As 1
    
    'make sure file is a valid object file
    Line Input #1, strTemp
    If strTemp <> "3D OBJECT DEFINITION FILE" Then
        MsgBox "Not a valid object file!", vbOKOnly + vbCritical, "Open"
        Exit Sub
    End If

    'get version
    Line Input #1, strTemp
    obj3dObject.Version = Trim(strTemp)

    'get name
    Line Input #1, strTemp
    obj3dObject.Name = Trim(strTemp)

    Line Input #1, strTemp 'skip line
    
    'get vertices
    Line Input #1, strTemp
    Do While strTemp <> ""

        lngNumVertices = lngNumVertices + 1
        ReDim Preserve obj3dObject.LocalCoord(0 To lngNumVertices - 1)
        
        obj3dObject.LocalCoord(lngNumVertices - 1).X = CDbl(Left(strTemp, InStr(1, strTemp, ",", vbTextCompare) - 1))
        lngNumTemp = InStr(1, strTemp, ",", vbTextCompare)
        obj3dObject.LocalCoord(lngNumVertices - 1).Y = CDbl(Mid(strTemp, lngNumTemp + 1, InStr(lngNumTemp + 1, strTemp, ",", vbTextCompare) - lngNumTemp - 1))
        lngNumTemp = InStr(lngNumTemp + 1, strTemp, ",", vbTextCompare)
        obj3dObject.LocalCoord(lngNumVertices - 1).Z = CDbl(Right(strTemp, Len(strTemp) - lngNumTemp))
            
        Line Input #1, strTemp
    Loop
    obj3dObject.NumVertices = lngNumVertices
    
    'get triangles
    Line Input #1, strTemp
    Do While strTemp <> "END"

        lngNumTriangles = lngNumTriangles + 1
        ReDim Preserve obj3dObject.Triangle(0 To lngNumTriangles - 1)
        ReDim Preserve obj3dObject.Color(0 To lngNumTriangles - 1)
        
        obj3dObject.Triangle(lngNumTriangles - 1).First = CDbl(Left(strTemp, InStr(1, strTemp, ",", vbTextCompare) - 1))
        lngNumTemp = InStr(1, strTemp, ",", vbTextCompare)
        obj3dObject.Triangle(lngNumTriangles - 1).Second = CDbl(Mid(strTemp, lngNumTemp + 1, InStr(lngNumTemp + 1, strTemp, ",", vbTextCompare) - lngNumTemp - 1))
        lngNumTemp = InStr(lngNumTemp + 1, strTemp, ",", vbTextCompare)
        obj3dObject.Triangle(lngNumTriangles - 1).Third = CDbl(Mid(strTemp, lngNumTemp + 1, InStr(lngNumTemp + 1, strTemp, ",", vbTextCompare) - lngNumTemp - 1))
        lngNumTemp = InStr(lngNumTemp + 1, strTemp, ",", vbTextCompare)
        obj3dObject.Color(lngNumTriangles - 1) = CLng(Right(strTemp, Len(strTemp) - lngNumTemp))
            
        Line Input #1, strTemp
    Loop
    obj3dObject.NumTriangles = lngNumTriangles

    Close #1

    'redimension arrays
    ReDim Preserve obj3dObject.RotatedLocalCoord(0 To obj3dObject.NumVertices - 1)
    ReDim Preserve obj3dObject.WorldCoord(0 To obj3dObject.NumVertices - 1)
    ReDim Preserve obj3dObject.CameraCoord(0 To obj3dObject.NumVertices - 1)
    ReDim Preserve obj3dObject.ScreenCoord(0 To obj3dObject.NumVertices - 1)
    ReDim Preserve obj3dObject.Isvisible(0 To obj3dObject.NumTriangles - 1)

End Sub
Private Sub LocaltoWorld()

    Dim lngIncr As Long

    'converts local coordinates into world ones
    For lngIncr = 0 To obj3dObject.NumVertices - 1
        obj3dObject.WorldCoord(lngIncr).X = obj3dObject.RotatedLocalCoord(lngIncr).X + obj3dObject.CenterofWorld.X
        obj3dObject.WorldCoord(lngIncr).Y = obj3dObject.RotatedLocalCoord(lngIncr).Y + obj3dObject.CenterofWorld.Y
        obj3dObject.WorldCoord(lngIncr).Z = obj3dObject.RotatedLocalCoord(lngIncr).Z + obj3dObject.CenterofWorld.Z
    Next

End Sub
Private Sub Project3dto2d()
    
    Dim lngIncr As Long
    
    'converts 3d point to a 2d point
    For lngIncr = 0 To obj3dObject.NumVertices - 1
        obj3dObject.ScreenCoord(lngIncr).X = (obj3dObject.WorldCoord(lngIncr).X * HPC / obj3dObject.WorldCoord(lngIncr).Z) + HALF_SCREEN_WIDTH
        obj3dObject.ScreenCoord(lngIncr).Y = (-obj3dObject.WorldCoord(lngIncr).Y * VPC * ASPECT_COMP / obj3dObject.WorldCoord(lngIncr).Z) + HALF_SCREEN_HEIGHT
    Next

End Sub
Public Sub RenderObject()

    Dim lngIncr As Long
    Dim ScreenBuffer(0 To 2) As POINTAPI
    Dim Brush As Long
    Dim Pen As Long
    Dim OldBrush As Long
    Dim OldPen As Long
    
    'calculate rotations around all three axis
    DoRotations
    'calculate translation and convert local rotated coordinate
    'to world coordinate
    LocaltoWorld
    'project 3d coordinates to 2d coordinates
    Project3dto2d
    'calculate the normal vectors
    CalculateNormals

    For lngIncr = 0 To obj3dObject.NumTriangles - 1
        
        'if face is visible then show it else don't
        If obj3dObject.Isvisible(lngIncr) = True Then
            With obj3dObject.ScreenCoord(obj3dObject.Triangle(lngIncr).First)
                ScreenBuffer(0).X = .X
                ScreenBuffer(0).Y = .Y
            End With
            With obj3dObject.ScreenCoord(obj3dObject.Triangle(lngIncr).Second)
                ScreenBuffer(1).X = .X
                ScreenBuffer(1).Y = .Y
            End With
            With obj3dObject.ScreenCoord(obj3dObject.Triangle(lngIncr).Third)
                ScreenBuffer(2).X = .X
                ScreenBuffer(2).Y = .Y
            End With
        
            'UNCOMMENT FOLLOWING TO DRAW SOLID
            Brush = CreateSolidBrush(obj3dObject.Color(lngIncr))
            Pen = CreatePen(PS_SOLID, 1, obj3dObject.Color(lngIncr))
            OldPen = SelectObject(Render.hdc, Pen)
            OldBrush = SelectObject(Render.hdc, Brush)
            Polygon Render.hdc, ScreenBuffer(0), 3
            SelectObject Render.hdc, OldPen
            SelectObject Render.hdc, OldBrush
            DeleteObject Pen
            DeleteObject Brush
        
            'UNCOMMENT FOLLOWING TO DRAW WIRE FRAME
'            Render.Line (ScreenBuffer(0).x, ScreenBuffer(0).y)-(ScreenBuffer(1).x, ScreenBuffer(1).y), 2
'            Render.Line (ScreenBuffer(1).x, ScreenBuffer(1).y)-(ScreenBuffer(2).x, ScreenBuffer(2).y), 2
'            Render.Line (ScreenBuffer(2).x, ScreenBuffer(2).y)-(ScreenBuffer(0).x, ScreenBuffer(0).y), 2
        End If
        
    Next

End Sub


Property Get RotateX() As Long

    RotateX = obj3dObject.Xangle

End Property

Property Get RotateY() As Long

    RotateY = obj3dObject.Yangle

End Property

Property Get RotateZ() As Long

    RotateZ = obj3dObject.Zangle

End Property



Property Get TranslateX() As Double

    TranslateX = obj3dObject.CenterofWorld.X

End Property

Property Get TranslateY() As Double

    TranslateY = obj3dObject.CenterofWorld.Y

End Property

Property Get TranslateZ() As Double

    TranslateZ = obj3dObject.CenterofWorld.Z

End Property

Private Sub DoRotations()

    Dim lngIncr As Long
    Dim RotationBuffer As Point3d

    'ROTATE X
    If obj3dObject.Xangle > 360 Then
        obj3dObject.Xangle = obj3dObject.Xangle - 360
    ElseIf obj3dObject.Xangle < 0 Then
        obj3dObject.Xangle = obj3dObject.Xangle + 360
    End If
    'ROTATE Y
    If obj3dObject.Yangle > 360 Then
        obj3dObject.Yangle = obj3dObject.Yangle - 360
    ElseIf obj3dObject.Yangle < 0 Then
        obj3dObject.Yangle = obj3dObject.Yangle + 360
    End If
    'ROTATE Z
    If obj3dObject.Zangle > 360 Then
        obj3dObject.Zangle = obj3dObject.Zangle - 360
    ElseIf obj3dObject.Zangle < 0 Then
        obj3dObject.Zangle = obj3dObject.Zangle + 360
    End If

    'rotate local coordinates around each axis
    For lngIncr = 0 To obj3dObject.NumVertices - 1
        RotationBuffer = obj3dObject.LocalCoord(lngIncr)
        obj3dObject.RotatedLocalCoord(lngIncr).X = obj3dObject.ScaleFactor * (RotationBuffer.X)
        obj3dObject.RotatedLocalCoord(lngIncr).Y = obj3dObject.ScaleFactor * (RotationBuffer.Y * Cos(DegtoRad(obj3dObject.Xangle)) - RotationBuffer.Z * Sin(DegtoRad(obj3dObject.Xangle)))
        obj3dObject.RotatedLocalCoord(lngIncr).Z = obj3dObject.ScaleFactor * (RotationBuffer.Z * Cos(DegtoRad(obj3dObject.Xangle)) + RotationBuffer.Y * Sin(DegtoRad(obj3dObject.Xangle)))
        
        RotationBuffer = obj3dObject.RotatedLocalCoord(lngIncr)
        obj3dObject.RotatedLocalCoord(lngIncr).X = obj3dObject.ScaleFactor * (RotationBuffer.X * Cos(DegtoRad(obj3dObject.Yangle)) + RotationBuffer.Z * Sin(DegtoRad(obj3dObject.Yangle)))
        obj3dObject.RotatedLocalCoord(lngIncr).Y = obj3dObject.ScaleFactor * (RotationBuffer.Y)
        obj3dObject.RotatedLocalCoord(lngIncr).Z = obj3dObject.ScaleFactor * (RotationBuffer.Z * Cos(DegtoRad(obj3dObject.Yangle)) - RotationBuffer.X * Sin(DegtoRad(obj3dObject.Yangle)))
        
        RotationBuffer = obj3dObject.RotatedLocalCoord(lngIncr)
        obj3dObject.RotatedLocalCoord(lngIncr).X = obj3dObject.ScaleFactor * (RotationBuffer.X * Cos(DegtoRad(obj3dObject.Zangle)) - RotationBuffer.Y * Sin(DegtoRad(obj3dObject.Zangle)))
        obj3dObject.RotatedLocalCoord(lngIncr).Y = obj3dObject.ScaleFactor * (RotationBuffer.Y * Cos(DegtoRad(obj3dObject.Zangle)) + RotationBuffer.X * Sin(DegtoRad(obj3dObject.Zangle)))
        obj3dObject.RotatedLocalCoord(lngIncr).Z = obj3dObject.ScaleFactor * (RotationBuffer.Z)
    Next

End Sub


Private Function DegtoRad(lngDeg As Long) As Double

    'converts degrees to radians
    DegtoRad = (lngDeg * PI) / 180

End Function



