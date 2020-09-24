VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "3D demo by ZATRiX"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7590
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   406
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   506
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pic3d 
      AutoRedraw      =   -1  'True
      Height          =   6060
      Left            =   0
      ScaleHeight     =   400
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   500
      TabIndex        =   0
      Top             =   0
      Width           =   7560
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'declare new objects
Dim obj1 As New cls3dObject
Dim obj2 As New cls3dObject
Dim obj3 As New cls3dObject
Dim obj4 As New cls3dObject
Sub RunDemo()

    Do
        'rotate around x,y,z axises for each object
        obj1.SetRotations obj1.RotateX - 3, obj1.RotateY + 3, obj1.RotateZ + 3
        obj2.SetRotations obj2.RotateX + 2, obj2.RotateY - 2, obj2.RotateZ + 1
        obj3.SetRotations obj3.RotateX + 1, obj3.RotateY + 1, obj3.RotateZ - 1
        obj4.SetRotations obj4.RotateX + 3, obj4.RotateY - 2, obj4.RotateZ + 1

        'clear picture box, render objects and pause for moment
        pic3d.Cls
        obj1.RenderObject
        obj2.RenderObject
        obj3.RenderObject
        obj4.RenderObject
        DoEvents
    Loop

End Sub
Private Sub Form_Load()

    Me.Show
    'load each object
    obj1.LoadObject App.Path & "\cube.odf", pic3d, -20, 0, -75, 2, 0, 0, 0
    obj2.LoadObject App.Path & "\cube.odf", pic3d, 20, 0, -70, 2, 0, 0, 0
    obj3.LoadObject App.Path & "\cube.odf", pic3d, 0, -20, -65, 2, 0, 0, 0
    obj4.LoadObject App.Path & "\cube.odf", pic3d, 0, 20, -60, 2, 0, 0, 0
    RunDemo
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    End

End Sub


