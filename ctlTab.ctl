VERSION 5.00
Begin VB.UserControl ctlTab 
   AutoRedraw      =   -1  'True
   ClientHeight    =   540
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1650
   FillColor       =   &H80000012&
   FillStyle       =   0  'Solid
   ForeColor       =   &H80000015&
   ScaleHeight     =   36
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   110
   ToolboxBitmap   =   "ctlTab.ctx":0000
End
Attribute VB_Name = "ctlTab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'JAIME ABAD 1/11/2003

Option Explicit

'Default Property Values:
Const m_def_MousePointer = 0
Const m_def_Appearance = 1
Const m_def_BorderStyle = 1
Const m_def_Autosize = True
Const m_def_Caption = "Tab"
Const m_def_Placement = 1
'Property Variables:
Dim m_MousePointer As MousePointerConstants
Dim m_MouseIcon As Picture
Dim m_Autosize As Boolean
Dim m_Caption As String
Dim m_Poligon(4) As PointApi
Dim m_FontUnderline As Boolean
Dim m_FontBold As Boolean
Dim m_FontItalic As Boolean
Dim m_BorderStyle As Integer
Dim m_Appearance As Integer
Dim m_Placement As Integer
Dim lastColorText As Long

Dim m_Dont_Refresh As Boolean

Event Click()
Event DblClick()
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=5
Public Sub Refresh()
Attribute Refresh.VB_Description = "Obliga a volver a dibujar un objeto."
    If m_Dont_Refresh Then Exit Sub
    Dim lngRegion As Long, lngRespuesta As Long
    Dim lngColor As Long, bolBold As Boolean
    With UserControl
        .Cls
        bolBold = .FontBold
        .FontBold = True
        If m_Autosize Then
            m_Dont_Refresh = True
            .Width = (.TextWidth(m_Caption) + 2.5 * BSIDE) * Screen.TwipsPerPixelX
            m_Dont_Refresh = False
        End If
        .FontBold = bolBold
        If m_Placement = 1 Then
            m_Poligon(0).X = 0
            m_Poligon(0).Y = 0
            m_Poligon(1).X = BSIDE
            m_Poligon(1).Y = .ScaleHeight
            m_Poligon(2).X = .ScaleWidth - BSIDE
            m_Poligon(2).Y = m_Poligon(1).Y
            m_Poligon(3).X = .ScaleWidth
            m_Poligon(3).Y = 0
        Else
            m_Poligon(0).X = BSIDE
            m_Poligon(0).Y = 0
            m_Poligon(1).X = .ScaleWidth - BSIDE
            m_Poligon(1).Y = 0
            m_Poligon(2).X = .ScaleWidth
            m_Poligon(2).Y = .ScaleHeight
            m_Poligon(3).X = 0
            m_Poligon(3).Y = m_Poligon(2).Y
        End If
        lngColor = .ForeColor
        lngRegion = CreatePolygonRgn(m_Poligon(0), 4, WINDING)
        lngRespuesta = SetWindowRgn(UserControl.hWnd, lngRegion, True)
        If m_BorderStyle <> None Then
            If m_Appearance = Flat Then
                DrawFlat
            Else
                Draw3D
            End If
        End If
        If UserControl.Enabled Then
            .CurrentX = (UserControl.ScaleWidth - .TextWidth(m_Caption)) / 2
            .CurrentY = ((.ScaleHeight - .TextHeight("A")) / 2) - 1
            .ForeColor = lngColor
            UserControl.Print m_Caption
        Else
            .CurrentX = (UserControl.ScaleWidth - .TextWidth(m_Caption)) / 2 + 1
            .CurrentY = (.ScaleHeight - .TextHeight("A")) / 2
            .ForeColor = vb3DHighlight
            UserControl.Print m_Caption
            .CurrentX = (UserControl.ScaleWidth - .TextWidth(m_Caption)) / 2
            .CurrentY = ((.ScaleHeight - .TextHeight("A")) / 2) - 1
            .ForeColor = vbGrayText
            UserControl.Print m_Caption
        End If
        UserControl.MousePointer = m_MousePointer
    End With
End Sub

Private Sub Draw3D()
    If m_Placement = 1 Then
        UserControl.ForeColor = vb3DHighlight
        UserControl.Line (m_Poligon(0).X + 1, m_Poligon(0).Y)-(m_Poligon(1).X, m_Poligon(1).Y - 1)
        UserControl.ForeColor = vb3DDKShadow
        UserControl.Line -(m_Poligon(2).X - 1, m_Poligon(2).Y - 1)
        UserControl.Line -(m_Poligon(3).X - 1, m_Poligon(3).Y)
        UserControl.ForeColor = vb3DLight
        UserControl.Line (m_Poligon(0).X + 2, m_Poligon(0).Y)-(m_Poligon(1).X + 1, m_Poligon(1).Y - 2)
        UserControl.ForeColor = vb3DShadow
        UserControl.Line -(m_Poligon(2).X - 2, m_Poligon(2).Y - 2)
        UserControl.Line -(m_Poligon(3).X - 2, m_Poligon(3).Y)
    Else
        UserControl.ForeColor = vb3DHighlight
        UserControl.Line (m_Poligon(3).X, m_Poligon(3).Y)-(m_Poligon(0).X, m_Poligon(0).Y)
        UserControl.Line -(m_Poligon(1).X - 1, m_Poligon(1).Y)
        UserControl.ForeColor = vb3DDKShadow
        UserControl.Line -(m_Poligon(2).X - 1, m_Poligon(2).Y)
        UserControl.ForeColor = vb3DLight
        UserControl.Line (m_Poligon(3).X + 1, m_Poligon(3).Y + 1)-(m_Poligon(0).X + 1, m_Poligon(0).Y + 1)
        UserControl.Line -(m_Poligon(1).X - 2, m_Poligon(1).Y + 1)
        UserControl.ForeColor = vb3DShadow
        UserControl.Line -(m_Poligon(2).X - 2, m_Poligon(2).Y)
    End If
End Sub

Private Sub DrawFlat()
    UserControl.ForeColor = vb3DDKShadow
    If m_Placement = 1 Then
        UserControl.Line (m_Poligon(0).X + 1, m_Poligon(0).Y)-(m_Poligon(1).X, m_Poligon(1).Y - 1)
        UserControl.Line -(m_Poligon(2).X - 1, m_Poligon(2).Y - 1)
        UserControl.Line -(m_Poligon(3).X - 1, m_Poligon(3).Y)
    Else
        UserControl.Line (m_Poligon(3).X + 1, m_Poligon(3).Y)-(m_Poligon(0).X + 1, m_Poligon(0).Y)
        UserControl.Line -(m_Poligon(1).X - 1, m_Poligon(1).Y)
        UserControl.Line -(m_Poligon(2).X - 1, m_Poligon(2).Y)
    End If
End Sub

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=13,0,0,Tab
Public Property Get Caption() As String
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    m_Caption = New_Caption
    Refresh
    PropertyChanged "Caption"
End Property

'Inicializar propiedades para control de usuario
Private Sub UserControl_InitProperties()
    m_Caption = m_def_Caption
    Set UserControl.Font = Ambient.Font
    m_Autosize = m_def_Autosize
    m_BorderStyle = m_def_BorderStyle
    m_Appearance = m_def_Appearance
    m_MousePointer = m_def_MousePointer
    Set m_MouseIcon = LoadPicture("")
End Sub

'Cargar valores de propiedad desde el almacén
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Caption = PropBag.ReadProperty("Caption", m_def_Caption)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    m_Autosize = PropBag.ReadProperty("Autosize", m_def_Autosize)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    m_Appearance = PropBag.ReadProperty("Appearance", m_def_Appearance)
    m_Placement = PropBag.ReadProperty("Placement", m_def_Placement)
    UserControl.FontStrikethru = PropBag.ReadProperty("FontStrikethru", Ambient.Font.Strikethrough)
    UserControl.FontSize = PropBag.ReadProperty("FontSize", Ambient.Font.Size)
    UserControl.FontName = PropBag.ReadProperty("FontName", Ambient.Font.Name)
    m_MousePointer = PropBag.ReadProperty("MousePointer", m_def_MousePointer)
    Set m_MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
End Sub

Private Sub UserControl_Resize()
    Refresh
End Sub

'Escribir valores de propiedad en el almacén
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("Autosize", m_Autosize, m_def_Autosize)
    Call PropBag.WriteProperty("Appearance", m_Appearance, m_def_Appearance)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("Placement", m_Placement, m_def_Placement)
    Call PropBag.WriteProperty("FontStrikethru", UserControl.FontStrikethru, 0)
    Call PropBag.WriteProperty("FontSize", UserControl.FontSize, 0)
    Call PropBag.WriteProperty("FontName", UserControl.FontName, "")
    Call PropBag.WriteProperty("MousePointer", m_MousePointer, m_def_MousePointer)
    Call PropBag.WriteProperty("MouseIcon", m_MouseIcon, Nothing)
End Sub
'
''ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
''MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    Refresh
    PropertyChanged "BackColor"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=UserControl,UserControl,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Devuelve o establece el color de primer plano usado para mostrar textos y gráficos en un objeto."
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    UserControl.ForeColor() = New_ForeColor
    Refresh
    PropertyChanged "ForeColor"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Devuelve o establece un valor que determina si un objeto puede responder a eventos generados por el usuario."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    If New_Enabled Then
        UserControl.ForeColor = lastColorText
    Else
        lastColorText = UserControl.ForeColor
        UserControl.ForeColor = &H80000011
    End If
    UserControl.Enabled = New_Enabled
    Refresh
    PropertyChanged "Enabled"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Devuelve un objeto Font."
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    With UserControl.Font
        .Bold = New_Font.Bold
        .Italic = New_Font.Italic
        .Underline = New_Font.Underline
    End With
    Refresh
    PropertyChanged "Font"
End Property

Public Property Get FontBold() As Boolean
    FontBold = m_FontBold
End Property

Public Property Let FontBold(New_Bold As Boolean)
    m_FontBold = New_Bold
    UserControl.FontBold = m_FontBold
    Refresh
    PropertyChanged "FontBold"
End Property

Public Property Get FontItalic() As Boolean
    FontItalic = m_FontItalic
End Property

Public Property Let FontItalic(New_Italic As Boolean)
    m_FontItalic = New_Italic
    UserControl.FontItalic = m_FontItalic
    Refresh
    PropertyChanged "FontItalic"
End Property

Public Property Get FontUnderline() As Boolean
    FontUnderline = m_FontUnderline
End Property

Public Property Let FontUnderline(New_Underline As Boolean)
    m_FontUnderline = New_Underline
    UserControl.FontUnderline = m_FontUnderline
    Refresh
    PropertyChanged "FontUnderline"
End Property

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=0,0,0,True
Public Property Get Autosize() As Boolean
    Autosize = m_Autosize
End Property

Public Property Let Autosize(ByVal New_Autosize As Boolean)
    m_Autosize = New_Autosize
    Refresh
    PropertyChanged "Autosize"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=22,0,0,
Public Property Get BorderStyle() As Integer
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
Attribute BorderStyle.VB_Description = "Devuelve o establece el estilo del borde de un objeto."
    m_BorderStyle = New_BorderStyle
    Refresh
    PropertyChanged "BorderStyle"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=7,0,0,0
Public Property Get Appearance() As Integer
    Appearance = m_Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As Integer)
Attribute Appearance.VB_Description = "Devuelve o establece si los objetos se dibujan en tiempo de ejecución con efectos 3D."
    m_Appearance = New_Appearance
    If m_Appearance = Flat Then
        UserControl.BackColor = vbButtonFace
    Else
        UserControl.BackColor = vbButtonFace
    End If
    Refresh
    PropertyChanged "Appearance"
End Property

Public Property Get Placement() As Integer
    Placement = m_Placement
End Property

Public Property Let Placement(New_Val As Integer)
    m_Placement = New_Val
    Refresh
    PropertyChanged "Placement"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=UserControl,UserControl,-1,FontStrikethru
Public Property Get FontStrikethru() As Boolean
Attribute FontStrikethru.VB_Description = "Devuelve o establece el estilo tachado de una fuente."
    FontStrikethru = UserControl.FontStrikethru
End Property

Public Property Let FontStrikethru(ByVal New_FontStrikethru As Boolean)
    UserControl.FontStrikethru() = New_FontStrikethru
    Refresh
    PropertyChanged "FontStrikethru"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=UserControl,UserControl,-1,FontSize
Public Property Get FontSize() As Single
Attribute FontSize.VB_Description = "Especifica el tamaño (en puntos) de la fuente que aparece en cada fila del nivel especificado."
    FontSize = UserControl.FontSize
End Property

Public Property Let FontSize(ByVal New_FontSize As Single)
    UserControl.FontSize() = New_FontSize
    Refresh
    PropertyChanged "FontSize"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=UserControl,UserControl,-1,FontName
Public Property Get FontName() As String
Attribute FontName.VB_Description = "Especifica el nombre de la fuente que aparece en cada fila del nivel especificado."
    FontName = UserControl.FontName
End Property

Public Property Let FontName(ByVal New_FontName As String)
    UserControl.FontName() = New_FontName
    Refresh
    PropertyChanged "FontName"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=21,0,0,
Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_Description = "Devuelve o establece el tipo de puntero del mouse mostrado al pasar por encima de un objeto."
    MousePointer = m_MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    m_MousePointer = New_MousePointer
    UserControl.MousePointer = m_MousePointer
    PropertyChanged "MousePointer"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=11,0,0,
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Establece un icono personalizado para el mouse."
    Set MouseIcon = m_MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set m_MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property
