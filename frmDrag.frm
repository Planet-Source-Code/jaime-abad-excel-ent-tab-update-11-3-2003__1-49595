VERSION 5.00
Begin VB.Form frmArrow 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000012&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   90
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   90
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   6
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmArrow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_Poligon(3) As PointApi

Private Sub Form_Load()
    Dim lngRegion As Long, lngRespuesta As Long
    With Me
        m_Poligon(0).X = 0
        m_Poligon(0).Y = 0
        m_Poligon(1).X = .ScaleWidth
        m_Poligon(1).Y = 0
        m_Poligon(2).X = .ScaleWidth / 2
        m_Poligon(2).Y = .ScaleHeight
        lngRegion = CreatePolygonRgn(m_Poligon(0), 3, WINDING)
        lngRespuesta = SetWindowRgn(Me.hWnd, lngRegion, False)
    End With
End Sub

