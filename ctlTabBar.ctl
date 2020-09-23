VERSION 5.00
Begin VB.UserControl xssTabBar 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   ClientHeight    =   525
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9075
   FillColor       =   &H80000012&
   FillStyle       =   0  'Solid
   ScaleHeight     =   35
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   605
   ToolboxBitmap   =   "ctlTabBar.ctx":0000
   Begin VB.PictureBox pcbLine 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   645
      Left            =   5940
      MousePointer    =   9  'Size W E
      ScaleHeight     =   43
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   6
      TabIndex        =   3
      Top             =   15
      Width           =   90
   End
   Begin VB.PictureBox pcbTab 
      BorderStyle     =   0  'None
      Height          =   420
      Left            =   1770
      ScaleHeight     =   28
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   193
      TabIndex        =   1
      Top             =   15
      Width           =   2895
      Begin VB.TextBox txtEdit 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2400
         MaxLength       =   32
         TabIndex        =   2
         Text            =   "Edit"
         Top             =   120
         Visible         =   0   'False
         Width           =   510
      End
      Begin XssTab.ctlTab ctlTab1 
         DragIcon        =   "ctlTabBar.ctx":0312
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontSize        =   8.25
         FontName        =   "MS Sans Serif"
      End
      Begin XssTab.ctlTab tabSelected 
         DragIcon        =   "ctlTabBar.ctx":061C
         Height          =   255
         Left            =   870
         TabIndex        =   5
         Top             =   0
         Visible         =   0   'False
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontSize        =   8.25
         FontName        =   "MS Sans Serif"
      End
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   285
      Left            =   6180
      Max             =   2
      TabIndex        =   0
      Top             =   45
      Width           =   2790
   End
   Begin VB.PictureBox pcbButtons 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   525
      Left            =   15
      ScaleHeight     =   35
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   87
      TabIndex        =   6
      Top             =   15
      Width           =   1305
      Begin VB.Timer tmrDrag 
         Enabled         =   0   'False
         Interval        =   300
         Left            =   0
         Top             =   0
      End
   End
End
Attribute VB_Name = "xssTabBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
'JAIME ABAD 1/11/2003

Option Explicit
'--------------------------
'*** OffsetScroll is in pixels
'--------------------------

Private WithEvents MouseEv As MouseEvent
Attribute MouseEv.VB_VarHelpID = -1
Private m_ColItems As Collection
'Default Property Values:
Const m_def_BorderStyle = 1
Const m_def_MoveableTabs = True
Const m_def_MultiSelect = False
Const BEGIN_BUTTONS = 6
Const TRI_HEIGHT = 8
Const TRI_WIDHT = 6
Const SEPARATION_BUTTONS = 18
Const MIN_SCROLL = 35
Const m_def_OffsetScroll = 200
Const m_def_ShowHscroll = True
Const m_def_ShowNavigator = True
Const m_def_SelColor = vb3DHighlight
Const m_def_SelectItem = 1
Const m_def_Item = 0
Const m_def_Appearance = 1
Const m_def_MoveLeft = 0
Const m_def_MoveRight = 0
Const m_def_Placement = Bottom
Const m_def_FontBold = False
'Property Variables:
Dim m_BorderStyle As Border
Dim m_MoveableTabs As Boolean
Dim m_FontBold As Boolean
Dim m_Placement As Placem
Dim m_MultiSelect As Boolean
Dim m_OffsetScroll As Single
Dim m_ShowHscroll As Boolean
Dim m_ShowNavigator As Boolean
Dim m_SelColor As OLE_COLOR
Dim m_SelectItem As Integer
Dim m_Item As Integer
Dim m_Appearance As Apparen
Dim m_MoveLeft As Integer
Dim m_MoveRight As Integer
Dim m_EndButtons As Integer
Dim m_MouseLineMove As Boolean
'-----------------------
Dim m_ButtonClickInNav As Integer
Dim X_to_Move As Long
Dim m_itnDrag As Integer
'Dim m_ScaleModeParent As Integer
'-----------------------
'enum
Public Enum Apparen
    Flat = 0
    x3D = 1
End Enum
Public Enum Border
    None = 0
    FixedSingle = 1
End Enum
Public Enum Placem
    plcBottom = Bottom
    plcTop = Top
End Enum
'Event Declarations:
Event ClickTab(Index As Integer)
Event DblClickTab(Index As Integer)
Event ChangeTab(Old_Selection As Integer, New_Selection As Integer)
Event ClickNavigator(Index As Integer)
Event DblClickNavigator(Index As Integer)
Event MouseMoveNavigator(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseDownNavigator(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUpNavigator(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUpTab(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseDownTab(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMoveTab(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseWheelUp()
Event MouseWheelDown()
Event ExitEdit(Text As String)
Event ResizeHScroll(New_Offset As Single)
Event ScrollChange() 'MappingInfo=HScroll1,HScroll1,-1,Change
Event Scroll() 'MappingInfo=HScroll1,HScroll1,-1,Scroll
Attribute Scroll.VB_Description = "Ocurre cuando cambia la posición de un cuadro de desplazamiento en un control."

'-----------------------
'property
'-----------------------

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Devuelve o establece el color de fondo usado para mostrar texto y gráficos en un objeto."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    Dim i As Integer, tmpBackColor As Long
    UserControl.BackColor() = New_BackColor
    tmpBackColor = UserControl.BackColor
    pcbButtons.BackColor = tmpBackColor
    pcbTab.BackColor = tmpBackColor
    pcbLine.BackColor = tmpBackColor
    For i = 0 To ctlTab1.UBound
        ctlTab1(i).BackColor = tmpBackColor
    Next
    DrawButtons
    PropertyChanged "BackColor"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Devuelve o establece un valor que determina si un objeto puede responder a eventos generados por el usuario."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled = New_Enabled
    HScroll1.Enabled = New_Enabled
    Dim i As Integer
    For i = 0 To ctlTab1.UBound
        ctlTab1(i).Enabled = New_Enabled
        pcbButtons.Enabled = New_Enabled
        pcbLine.Enabled = New_Enabled
    Next
    tabSelected.Enabled = New_Enabled
    DrawButtons
    PropertyChanged "Enabled"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=7,0,0,0
Public Property Get Appearance() As Apparen
Attribute Appearance.VB_Description = "Devuelve o establece si los objetos se dibujan en tiempo de ejecución con efectos 3D."
    Appearance = m_Appearance
End Property

Public Property Let Appearance(ByVal New_Appearance As Apparen)
    m_Appearance = New_Appearance
    UserControl.Appearance = New_Appearance
    Me.BackColor = UserControl.BackColor
    Dim i As Integer
    For i = 0 To ctlTab1.UBound
        ctlTab1(i).Appearance = m_Appearance
    Next
    With tabSelected
        .Appearance = m_Appearance
        .BackColor = m_SelColor
    End With
    SelectItem = m_SelectItem
    Me.BorderStyle = m_BorderStyle
    PropertyChanged "Appearance"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=7,0,0,0
Public Property Get SelectItem() As Integer
    SelectItem = m_SelectItem
End Property

Public Property Let SelectItem(ByVal New_SelectItem As Integer)
    Dim i As Integer, Old_Selct As Integer, wasBold As Boolean
    If New_SelectItem < 1 Or New_SelectItem > ctlTab1.UBound Or Not Ambient.UserMode Then
        Exit Property
    End If
    'UserControl.ScaleMode = vbPixels
    pcbTab.Visible = False
    Old_Selct = m_SelectItem
    If Old_Selct <= ctlTab1.UBound Then
        If m_MultiSelect = False Then
            ctlTab1(Old_Selct).BackColor = UserControl.BackColor
        End If
        ctlTab1(Old_Selct).FontBold = m_FontBold
    End If
    For i = 1 To ctlTab1.UBound
        ctlTab1(i).ZOrder 1
    Next
    m_SelectItem = New_SelectItem
    If m_SelectItem > 1 Then
        With ctlTab1(m_SelectItem)
            If m_ShowHscroll Then
                If .Left + .Width + pcbTab.Left + BSIDE > HScroll1.Left Then
                    pcbTab.Left = HScroll1.Left - .Left - .Width - 2 * BSIDE
                End If
            Else
                If .Left + .Width + pcbTab.Left + BSIDE > UserControl.ScaleWidth Then
                    pcbTab.Left = UserControl.ScaleWidth - .Left - .Width - 2 * BSIDE
                End If
            End If
            If pcbTab.Left + .Left < m_EndButtons Then
                pcbTab.Left = m_EndButtons - .Left + BSIDE
            End If
        End With
    ElseIf m_SelectItem = 1 Then
        pcbTab.Left = m_EndButtons
    End If
    With ctlTab1(m_SelectItem)
        .BackColor = SelColor
        .FontBold = True
        .ZOrder 0
    End With
    pcbTab.Visible = True
    If Old_Selct <> New_SelectItem Then
        RaiseEvent ChangeTab(Old_Selct, New_SelectItem)
    End If
    'UserControl.ScaleMode = m_ScaleModeParent
    PropertyChanged "SelectItem"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=14,0,0,121212
Public Property Get SelColor() As OLE_COLOR
    SelColor = m_SelColor
End Property

Public Property Let SelColor(ByVal New_SelColor As OLE_COLOR)
    m_SelColor = New_SelColor
    tabSelected.BackColor = m_SelColor
    SelectItem = m_SelectItem
    PropertyChanged "SelColor"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=0,0,0,true
Public Property Get ShowHscroll() As Boolean
    ShowHscroll = m_ShowHscroll
End Property

Public Property Let ShowHscroll(ByVal New_ShowHscroll As Boolean)
    m_ShowHscroll = New_ShowHscroll
    OrderScroll
    PropertyChanged "ShowHscroll"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=0,0,0,true
Public Property Get ShowNavigator() As Boolean
    ShowNavigator = m_ShowNavigator
End Property

Public Property Let ShowNavigator(ByVal New_ShowNavigator As Boolean)
    m_ShowNavigator = New_ShowNavigator
    If m_ShowNavigator Then
        pcbButtons.Visible = True
        m_EndButtons = pcbButtons.Width
        pcbButtons.ZOrder 0
        DrawButtons
    Else
        pcbButtons.Visible = False
        m_EndButtons = 0
    End If
    If Ambient.UserMode = True Then
        SelectItem = m_SelectItem
    End If
    pcbTab.Left = m_EndButtons
    PropertyChanged "ShowNavigator"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=12,0,0,200
Public Property Get OffsetScroll() As Single
    OffsetScroll = m_OffsetScroll
End Property

Public Property Let OffsetScroll(ByVal New_OffsetScroll As Single)
    m_OffsetScroll = New_OffsetScroll
    OrderScroll
    PropertyChanged "OffsetScroll"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=HScroll1,HScroll1,-1,SmallChange
Public Property Get SmallChange() As Integer
    SmallChange = HScroll1.SmallChange
End Property

Public Property Let SmallChange(ByVal New_SmallChange As Integer)
    HScroll1.SmallChange() = New_SmallChange
    PropertyChanged "SmallChange"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=HScroll1,HScroll1,-1,LargeChange
Public Property Get LargeChange() As Integer
    LargeChange = HScroll1.LargeChange
End Property

Public Property Let LargeChange(ByVal New_LargeChange As Integer)
    HScroll1.LargeChange = New_LargeChange
    PropertyChanged "LargeChange"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=0,0,0,False
Public Property Get MultiSelect() As Boolean
    MultiSelect = m_MultiSelect
End Property

Public Property Let MultiSelect(ByVal New_MultiSelect As Boolean)
    m_MultiSelect = New_MultiSelect
    If m_MultiSelect = False Then
        Dim i As Integer
        For i = 0 To ctlTab1.UBound
            ctlTab1(i).BackColor = UserControl.BackColor
        Next
        SelectItem = m_SelectItem
    End If
    PropertyChanged "MultiSelect"
End Property

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=HScroll1,HScroll1,-1,Min
Public Property Get ScrollMin() As Integer
    ScrollMin = HScroll1.Min
End Property

Public Property Let ScrollMin(ByVal New_ScrollMin As Integer)
    HScroll1.Min() = New_ScrollMin
    PropertyChanged "ScrollMin"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=HScroll1,HScroll1,-1,Max
Public Property Get ScrollMax() As Integer
    ScrollMax = HScroll1.Max
End Property

Public Property Let ScrollMax(ByVal New_ScrollMax As Integer)
    HScroll1.Max() = New_ScrollMax
    PropertyChanged "ScrollMax"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=HScroll1,HScroll1,-1,Value
Public Property Get ScrollValue() As Integer
    ScrollValue = HScroll1.Value
End Property

Public Property Let ScrollValue(ByVal New_ScrollValue As Integer)
    If New_ScrollValue > HScroll1.Max Or New_ScrollValue < HScroll1.Min Then
        Exit Property
    End If
    HScroll1.Value() = New_ScrollValue
    PropertyChanged "ScrollValue"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=pcbLine,pcbLine,-1,Enabled
Public Property Get Resizable() As Boolean
    Resizable = pcbLine.Enabled
End Property

Public Property Let Resizable(ByVal New_Resizable As Boolean)
    pcbLine.Enabled() = New_Resizable
    PropertyChanged "Resizable"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=pcbLine,pcbLine,-1,Enabled
Public Property Get Caption(Index As Integer) As String
    Caption = ctlTab1(Index).Caption
End Property

Public Property Let Caption(Index As Integer, ByVal strCaption As String)
    Dim varCaption As Variant
    varCaption = strCaption
    m_ColItems.Add varCaption, After:=Index
    m_ColItems.Remove Index
    ReorderTabs
    SelectItem = Index
End Property

Public Property Get Count() As Integer
    Count = ctlTab1.UBound
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=7,0,0,0
Public Property Get Placement() As Placem
    Placement = m_Placement
End Property

Public Property Let Placement(ByVal New_Placement As Placem)
    Dim i As Integer
    m_Placement = New_Placement
    For i = 0 To ctlTab1.UBound
        ctlTab1(i).Placement = m_Placement
    Next i
    tabSelected.Placement = m_Placement
    PropertyChanged "Placement"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=UserControl,UserControl,-1,Font
Public Property Get Font() As Font
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    Dim i As Integer
    For i = 0 To ctlTab1.UBound
        Set ctlTab1(i).Font = UserControl.Font
        FontBold = UserControl.FontBold
    Next
    ReorderTabs
    PropertyChanged "Font"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=UserControl,UserControl,-1,FontItalic
Public Property Get FontItalic() As Boolean
    FontItalic = UserControl.FontItalic
End Property

Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
    UserControl.FontItalic = New_FontItalic
    Dim i As Integer
    For i = 0 To ctlTab1.UBound
        ctlTab1(i).FontItalic = UserControl.FontItalic
    Next
    ReorderTabs
    PropertyChanged "FontItalic"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=UserControl,UserControl,-1,FontUnderline
Public Property Get FontUnderline() As Boolean
    FontUnderline = UserControl.FontUnderline
End Property

Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
    UserControl.FontUnderline = New_FontUnderline
    Dim i As Integer
    For i = 0 To ctlTab1.UBound
        ctlTab1(i).FontUnderline = UserControl.FontUnderline
    Next
    ReorderTabs
    PropertyChanged "FontUnderline"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=UserControl,UserControl,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
    ForeColor = UserControl.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    UserControl.ForeColor = New_ForeColor
    Dim i As Integer
    For i = 0 To ctlTab1.UBound
        ctlTab1(i).ForeColor = UserControl.ForeColor
    Next
    DrawButtons
    PropertyChanged "ForeColor"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=UserControl,UserControl,-1,FontStrikethru
Public Property Get FontStrikethru() As Boolean
    FontStrikethru = UserControl.FontStrikethru
End Property

Public Property Let FontStrikethru(ByVal New_FontStrikethru As Boolean)
    UserControl.FontStrikethru() = New_FontStrikethru
    Dim i As Integer
    For i = 0 To ctlTab1.UBound
        ctlTab1(i).FontStrikethru = UserControl.FontStrikethru
    Next
    ReorderTabs
    PropertyChanged "FontStrikethru"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=0,0,0,
Public Property Get FontBold() As Boolean
    FontBold = m_FontBold
End Property

Public Property Let FontBold(ByVal New_FontBold As Boolean)
    m_FontBold = New_FontBold
    Dim i As Integer
    For i = 0 To ctlTab1.UBound
        ctlTab1(i).FontBold = m_FontBold
    Next
    ReorderTabs
    PropertyChanged "FontBold"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=0,0,0,true
Public Property Get MoveableTabs() As Boolean
    MoveableTabs = m_MoveableTabs
End Property

Public Property Let MoveableTabs(ByVal New_MoveableTabs As Boolean)
    m_MoveableTabs = New_MoveableTabs
    PropertyChanged "MoveableTabs"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=UserControl,UserControl,-1,MouseIcon
Public Property Get MouseIcon() As Picture
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=UserControl,UserControl,-1,MousePointer
Public Property Get MousePointer() As MousePointerConstants
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

'---------------------------------------------
'Private sub events
'---------------------------------------------

Private Sub ctlTab1_Click(Index As Integer)
    RaiseEvent ClickTab(Index)
End Sub

Private Sub ctlTab1_DblClick(Index As Integer)
    RaiseEvent DblClickTab(Index)
End Sub

Private Sub ctlTab1_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
    tmrDrag.Enabled = False
    If Index = m_itnDrag Then Exit Sub
    'UserControl.ScaleMode = vbPixels
    If X < (ctlTab1(Index).Width * Screen.TwipsPerPixelX / 2) Then
        MoveTab m_itnDrag, Index
        SelectItem = IIf(Index < m_itnDrag, Index, Index - 1)
    Else
        MoveTab m_itnDrag, After:=Index
        SelectItem = IIf(Index > m_itnDrag, Index, Index + 1)
    End If
    'UserControl.ScaleMode = m_ScaleModeParent
End Sub

Private Sub ctlTab1_DragOver(Index As Integer, Source As Control, X As Single, Y As Single, State As Integer)
    Dim rctEnd As RECT, papCur As PointApi
    Call GetWindowRect(ctlTab1(Index).hWnd, rctEnd)
    Call GetCursorPos(papCur)
    Call SetCursorPos(papCur.X, rctEnd.Top + (rctEnd.Bottom - rctEnd.Top) / 2)
End Sub

Private Sub ctlTab1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    SelectItem = Index
    RaiseEvent MouseDownTab(Index, Button, Shift, X, Y)
    Select Case Button
        Case 1
            If Not m_MoveableTabs Then Exit Sub
            tmrDrag.Enabled = True
            m_itnDrag = Index
        Case 2
        Case 4
    End Select
End Sub

Private Sub ctlTab1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMoveTab(Index, Button, Shift, X, Y)
End Sub

Private Sub ctlTab1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUpTab(Index, Button, Shift, X, Y)
    tmrDrag.Enabled = False
End Sub

Private Sub HScroll1_Change()
    RaiseEvent ScrollChange
End Sub

Private Sub HScroll1_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    'UserControl.ScaleMode = vbPixels
    If X Mod 2 = 0 Then
        Dim rctEnd As RECT
        Call GetWindowRect(HScroll1.hWnd, rctEnd)
        Call SetCursorPos(rctEnd.Left - 2, rctEnd.Top + (rctEnd.Bottom - rctEnd.Top) / 2)
        MoveLeft
    End If
    'UserControl.ScaleMode = m_ScaleModeParent
End Sub

Private Sub MouseEv_MouseWheelDown()
    RaiseEvent MouseWheelDown
End Sub

Private Sub MouseEv_MouseWheelUp()
    RaiseEvent MouseWheelUp
End Sub

Private Sub pcbButtons_Click()
    m_EndButtons = 0
    If m_ShowNavigator Then
        m_EndButtons = pcbButtons.Width
    End If
    RaiseEvent ClickNavigator(m_ButtonClickInNav)
End Sub

Private Sub pcbButtons_DblClick()
    RaiseEvent DblClickNavigator(m_ButtonClickInNav)
End Sub

Private Sub pcbButtons_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    If X Mod 2 = 0 Then
        Dim rctEnd As RECT
        Call GetWindowRect(pcbButtons.hWnd, rctEnd)
        Call SetCursorPos(rctEnd.Right + 2, rctEnd.Top + (rctEnd.Bottom - rctEnd.Top) / 2)
        MoveRight
    End If
End Sub

Private Sub pcbButtons_LostFocus()
    DrawButtons
End Sub

Private Sub pcbButtons_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lng_X As Long
    'UserControl.ScaleMode = vbPixels
    lng_X = TRI_WIDHT + BEGIN_BUTTONS + SEPARATION_BUTTONS / 2
    With pcbButtons
        If X >= 0 And X < lng_X Then
            m_ButtonClickInNav = 0
        ElseIf X >= lng_X And X < 2 * lng_X Then
            m_ButtonClickInNav = 1
        ElseIf X >= 2 * lng_X And X < 3 * lng_X Then
            m_ButtonClickInNav = 2
        ElseIf X >= 3 * lng_X And X < 4 * lng_X Then
            m_ButtonClickInNav = 3
        End If
    End With
    'UserControl.ScaleMode = m_ScaleModeParent
    RaiseEvent MouseDownNavigator(m_ButtonClickInNav, Button, Shift, X, Y)
End Sub

Private Sub pcbButtons_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim x1 As Long, x2 As Long, y1 As Long, y2 As Long
    Dim lngLastColor As Long, lng_X As Long, Index As Integer
    'UserControl.ScaleMode = vbPixels
    lng_X = TRI_WIDHT + BEGIN_BUTTONS + SEPARATION_BUTTONS / 2
    With pcbButtons
        y1 = .ScaleHeight - 1
        y2 = 1
        If X >= 0 And X < lng_X Then
            x1 = 1
            x2 = lng_X
            Index = 0
        ElseIf X >= lng_X And X < 2 * lng_X Then
            x1 = lng_X
            x2 = 2 * lng_X
            Index = 1
        ElseIf X >= 2 * lng_X And X < 3 * lng_X Then
            x1 = 2 * lng_X
            x2 = 3 * lng_X
            Index = 2
        ElseIf X >= 3 * lng_X And X < 4 * lng_X Then
            x1 = 3 * lng_X
            x2 = 4 * lng_X
            Index = 3
        End If
        lngLastColor = .ForeColor
        DrawButtons
        .ForeColor = vb3DHighlight
        pcbButtons.Line (x1, y1)-(x1, y2)
        pcbButtons.Line -(x2, y2)
        .ForeColor = vb3DShadow
        pcbButtons.Line -(x2, y1)
        pcbButtons.Line -(x1, y1)
        .ForeColor = lngLastColor
    End With
    'UserControl.ScaleMode = m_ScaleModeParent
    RaiseEvent MouseMoveNavigator(Index, Button, Shift, X, Y)
End Sub

Private Sub pcbButtons_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUpNavigator(m_ButtonClickInNav, Button, Shift, X, Y)
End Sub

Private Sub pcbButtons_Resize()
    DrawButtons
End Sub

Private Sub pcbLine_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Old_Width As Single, New_Offset As Long
    'UserControl.ScaleMode = m_ScaleModeParent
    Old_Width = HScroll1.Width
    'UserControl.ScaleMode = vbPixels
    m_OffsetScroll = m_OffsetScroll + X
    OrderScroll
    pcbTab.Left = m_EndButtons
    SelectItem = m_SelectItem
    'UserControl.ScaleMode = m_ScaleModeParent
    RaiseEvent ResizeHScroll(m_OffsetScroll)
End Sub

Private Sub HScroll1_Scroll()
    RaiseEvent Scroll
End Sub

Private Sub HScroll1_ScrollChange()
    RaiseEvent ScrollChange
End Sub

Private Sub tmrDrag_Timer()
    ctlTab1(m_itnDrag).Drag 1
    tmrDrag.Enabled = False
End Sub

Private Sub txtEdit_Validate(Cancel As Boolean)
    txtEdit.Visible = False
    RaiseEvent ExitEdit(txtEdit)
End Sub

Private Sub UserControl_EnterFocus()
    SetTheProc UserControl.hWnd, MouseEv
End Sub

Private Sub UserControl_ExitFocus()
    LostTheProc
End Sub

Private Sub UserControl_Paint()
    If Ambient.UserMode = False Then
        IniExample
    End If
End Sub

Private Sub UserControl_Resize()
    Dim i As Integer, tmpHeight As Double
    tmpHeight = UserControl.ScaleHeight - 2
    For i = 0 To ctlTab1.UBound
        ctlTab1(i).Height = tmpHeight
    Next
    pcbTab.Height = tmpHeight
    pcbButtons.Height = tmpHeight
    pcbLine.Height = tmpHeight
    HScroll1.Height = tmpHeight
    tabSelected.Height = tmpHeight
    Refresh
End Sub

Private Sub UserControl_Initialize()
    Set MouseEv = New MouseEvent
    Set m_ColItems = New Collection
End Sub

'Inicializar propiedades para control de usuario
Private Sub UserControl_InitProperties()
    m_Appearance = m_def_Appearance
    m_MoveLeft = m_def_MoveLeft
    m_MoveRight = m_def_MoveRight
    m_SelectItem = m_def_SelectItem
    m_Item = m_def_Item
    m_SelColor = m_def_SelColor
    m_ShowHscroll = m_def_ShowHscroll
    m_ShowNavigator = m_def_ShowNavigator
    m_OffsetScroll = m_def_OffsetScroll
    m_MultiSelect = m_def_MultiSelect
    m_Placement = m_def_Placement
    m_FontBold = m_def_FontBold
    m_MoveableTabs = m_def_MoveableTabs
    m_BorderStyle = m_def_BorderStyle
    IniAll
End Sub

'Cargar valores de propiedad desde el almacén
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    m_Appearance = PropBag.ReadProperty("Appearance", m_def_Appearance)
    m_MoveLeft = PropBag.ReadProperty("MoveLeft", m_def_MoveLeft)
    m_MoveRight = PropBag.ReadProperty("MoveRight", m_def_MoveRight)
    m_SelectItem = PropBag.ReadProperty("SelectItem", m_def_SelectItem)
    m_Item = PropBag.ReadProperty("Item", m_def_Item)
    m_SelColor = PropBag.ReadProperty("SelColor", m_def_SelColor)
    m_ShowHscroll = PropBag.ReadProperty("ShowHscroll", m_def_ShowHscroll)
    m_ShowNavigator = PropBag.ReadProperty("ShowNavigator", m_def_ShowNavigator)
    m_OffsetScroll = PropBag.ReadProperty("OffsetScroll", m_def_OffsetScroll)
    m_MultiSelect = PropBag.ReadProperty("MultiSelect", m_def_MultiSelect)
    m_Placement = PropBag.ReadProperty("Placement", m_def_Placement)
    m_FontBold = PropBag.ReadProperty("FontBold", m_def_FontBold)
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.FontItalic = PropBag.ReadProperty("FontItalic", Ambient.Font.Italic)
    UserControl.FontUnderline = PropBag.ReadProperty("FontUnderline", Ambient.Font.Underline)
    UserControl.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    UserControl.FontStrikethru = PropBag.ReadProperty("FontStrikethru", Ambient.Font.Strikethrough)
    HScroll1.SmallChange = PropBag.ReadProperty("SmallChange", 1)
    HScroll1.Min = PropBag.ReadProperty("ScrollMin", 0)
    HScroll1.Max = PropBag.ReadProperty("ScrollMax", 32767)
    HScroll1.LargeChange = PropBag.ReadProperty("LargeChange", 1)
    HScroll1.Min = PropBag.ReadProperty("ScrollMin", 0)
    HScroll1.Max = PropBag.ReadProperty("ScrollMax", 32767)
    HScroll1.Value = PropBag.ReadProperty("ScrollValue", 0)
    pcbLine.Enabled = PropBag.ReadProperty("Resizable", True)
    m_MoveableTabs = PropBag.ReadProperty("MoveableTabs", m_def_MoveableTabs)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    IniAll
End Sub

'Escribir valores de propiedad en el almacén
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        Call .WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
        Call .WriteProperty("Enabled", UserControl.Enabled, True)
        Call .WriteProperty("Appearance", m_Appearance, m_def_Appearance)
        Call .WriteProperty("MoveLeft", m_MoveLeft, m_def_MoveLeft)
        Call .WriteProperty("MoveRight", m_MoveRight, m_def_MoveRight)
        Call .WriteProperty("SelectItem", m_SelectItem, m_def_SelectItem)
        Call .WriteProperty("Item", m_Item, m_def_Item)
        Call .WriteProperty("SelColor", m_SelColor, m_def_SelColor)
        Call .WriteProperty("ShowHscroll", m_ShowHscroll, m_def_ShowHscroll)
        Call .WriteProperty("ShowNavigator", m_ShowNavigator, m_def_ShowNavigator)
        Call .WriteProperty("OffsetScroll", m_OffsetScroll, m_def_OffsetScroll)
        Call .WriteProperty("ScrollValue", HScroll1.Value, 0)
        Call .WriteProperty("SmallChange", HScroll1.SmallChange, 1)
        Call .WriteProperty("ScrollMin", HScroll1.Min, 0)
        Call .WriteProperty("ScrollMax", HScroll1.Max, 2)
        Call .WriteProperty("LargeChange", HScroll1.LargeChange, 1)
        Call .WriteProperty("MultiSelect", m_MultiSelect, m_def_MultiSelect)
        Call .WriteProperty("ScrollMin", HScroll1.Min, 0)
        Call .WriteProperty("ScrollMax", HScroll1.Max, 32767)
        Call .WriteProperty("ScrollValue", HScroll1.Value, 0)
        Call .WriteProperty("Resizable", pcbLine.Enabled, True)
        Call .WriteProperty("Placement", m_Placement, m_def_Placement)
        Call .WriteProperty("Font", UserControl.Font, Ambient.Font)
        Call .WriteProperty("FontItalic", UserControl.FontItalic, Ambient.Font.Italic)
        Call .WriteProperty("FontUnderline", UserControl.FontUnderline, Ambient.Font.Underline)
        Call .WriteProperty("ForeColor", UserControl.ForeColor, &H80000012)
        Call .WriteProperty("FontStrikethru", UserControl.FontStrikethru, Ambient.Font.Strikethrough)
        Call .WriteProperty("FontBold", m_FontBold, m_def_FontBold)
        Call .WriteProperty("MoveableTabs", m_MoveableTabs, m_def_MoveableTabs)
        Call .WriteProperty("MouseIcon", MouseIcon, Nothing)
        Call .WriteProperty("MousePointer", UserControl.MousePointer, 0)
        Call .WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
    End With
End Sub

'---------------------------------------------
'Public sub
'---------------------------------------------

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=5
Public Sub Refresh()
    Me.BorderStyle = m_BorderStyle
    OrderScroll
    DrawButtons
    SelectItem = m_SelectItem
End Sub

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=14
Public Sub AddNew(ByVal Caption As Variant, Optional Before As Integer = 0, Optional After As Integer = 0)
    Dim intLast As Integer, i As Integer
    intLast = ctlTab1.Count
    Load ctlTab1(intLast)
    Set ctlTab1(intLast).Container = pcbTab
    With ctlTab1(intLast)
        .Caption = Caption
        .FontBold = m_FontBold
        .Top = 0
        .BackColor = UserControl.BackColor
        .ForeColor = UserControl.ForeColor
        .ZOrder 1
        .Visible = True
    End With
    If Before > 0 And After = 0 Then
        m_ColItems.Add Caption, Before:=Before
    ElseIf Before = 0 And After > 0 Then
        m_ColItems.Add Caption, After:=After
    Else
        m_ColItems.Add Caption
    End If
    ReorderTabs
    SelectItem = m_SelectItem
End Sub

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=14
Public Sub Remove(Index As Variant)
    On Error GoTo xExit
    m_ColItems.Remove (Index)
    Unload ctlTab1(ctlTab1.UBound)
    ReorderTabs
xExit:
End Sub

Public Sub SelectNext()
    If m_SelectItem < ctlTab1.UBound Then
        SelectItem = m_SelectItem + 1
    End If
End Sub

Public Sub SelectLast()
    If m_SelectItem > 0 Then
        SelectItem = m_SelectItem - 1
    End If
End Sub

Public Sub EditMode()
    'UserControl.ScaleMode = vbPixels
    If m_SelectItem > 1 Or m_SelectItem <= ctlTab1.UBound Then
        With txtEdit
            .ZOrder 0
            .Text = ctlTab1(m_SelectItem).Caption
            .BackColor = ctlTab1(m_SelectItem).BackColor
            Set .Font = ctlTab1(m_SelectItem).Font
            .FontBold = ctlTab1(m_SelectItem).FontBold
            .Left = ctlTab1(m_SelectItem).Left + BSIDE + 2
            .Height = UserControl.TextHeight(.Text)
            .Top = (ctlTab1(m_SelectItem).Height - .Height) / 2
            .Width = ctlTab1(m_SelectItem).Width - 2 * BSIDE
            .Visible = True
            .SetFocus
            .SelStart = 0
            .SelLength = Len(.Text)
        End With
    End If
    'UserControl.ScaleMode = m_ScaleModeParent
End Sub

Public Sub MoveLeft()
    'UserControl.ScaleMode = vbPixels
    With ctlTab1(ctlTab1.UBound)
        If m_ShowHscroll Then
            If .Left + .Width + pcbTab.Left + BSIDE < HScroll1.Left Then
                'UserControl.ScaleMode = m_ScaleModeParent
                Exit Sub
            End If
        Else
            If .Left + .Width + pcbTab.Left + BSIDE < UserControl.ScaleWidth Then
                'UserControl.ScaleMode = m_ScaleModeParent
                Exit Sub
            End If
        End If
    End With
    If KnowFirst() >= ctlTab1.UBound Then
        'UserControl.ScaleMode = m_ScaleModeParent
        Exit Sub
    End If
    pcbTab.Left = m_EndButtons - ctlTab1(KnowFirst() + 1).Left + BSIDE
    'UserControl.ScaleMode = m_ScaleModeParent
End Sub

Public Sub MoveRight()
    'UserControl.ScaleMode = vbPixels
    If KnowFirst() <= 2 Then
        MoveFirst
        Exit Sub
    End If
    pcbTab.Left = m_EndButtons - ctlTab1(KnowFirst() - 1).Left + BSIDE
    'UserControl.ScaleMode = m_ScaleModeParent
End Sub

Public Sub MoveFirst()
    m_EndButtons = 0
    'UserControl.ScaleMode = vbPixels
    If m_ShowNavigator Then
        m_EndButtons = pcbButtons.Width
    End If
    pcbTab.Left = m_EndButtons
    'UserControl.ScaleMode = m_ScaleModeParent
End Sub

Public Sub MoveEnd()
    'UserControl.ScaleMode = vbPixels
    With ctlTab1(ctlTab1.UBound)
        If m_ShowHscroll Then
            If .Left + .Width + pcbTab.Left + BSIDE < HScroll1.Left Then
                'UserControl.ScaleMode = m_ScaleModeParent
                Exit Sub
            End If
            pcbTab.Left = HScroll1.Left - .Left - .Width - 2 * BSIDE
        Else
            If .Left + .Width + pcbTab.Left + BSIDE < UserControl.ScaleWidth Then
                'UserControl.ScaleMode = m_ScaleModeParent
                Exit Sub
            End If
            pcbTab.Left = UserControl.ScaleWidth - .Left - .Width - 2 * BSIDE
        End If
        If pcbTab.Left + .Left < m_EndButtons Then
            pcbTab.Left = m_EndButtons - .Left + BSIDE
        End If
    End With
    'UserControl.ScaleMode = m_ScaleModeParent
End Sub

Public Sub MoveTab(Index As Integer, Optional Before As Integer = -1, Optional After As Integer = -1)
    Dim intIndex As Integer
    If Not m_MoveableTabs Then Exit Sub
    If Index = Before Or Index = After Then Exit Sub
    If Before < 1 And After < 0 Then Exit Sub
    If Before < 0 And After > m_ColItems.Count Then Exit Sub
    If Before > 0 And After < 0 Then
        m_ColItems.Add m_ColItems(Index), Before:=Before
        m_ColItems.Remove (IIf(Index > Before, Index + 1, Index))
    ElseIf Before < 0 And After > 0 Then
        m_ColItems.Add m_ColItems(Index), After:=After
        m_ColItems.Remove (IIf(Index > After, Index + 1, Index))
    End If
    ReorderTabs
End Sub

'---------------------------------------------
'Private sub events
'---------------------------------------------

Private Sub OrderScroll()
    'UserControl.ScaleMode = vbPixels
    If m_ShowHscroll Then
        With HScroll1
            .Top = 1
            .ZOrder 0
            .Left = m_OffsetScroll + m_EndButtons
            If .Left < m_EndButtons + 2 * BSIDE Then
                m_OffsetScroll = 2 * BSIDE
                .Left = m_OffsetScroll + m_EndButtons
            End If
            If UserControl.ScaleWidth - .Left > MIN_SCROLL Then
                .Width = UserControl.ScaleWidth - .Left
            Else
                .Left = UserControl.ScaleWidth - MIN_SCROLL
                .Width = MIN_SCROLL
                m_OffsetScroll = .Left - m_EndButtons
            End If
            pcbLine.Left = .Left - pcbLine.Width
            pcbLine.Top = 1
            pcbLine.Height = .Height
            pcbLine.ZOrder 0
            pcbLine.Visible = True
            .Visible = True
        End With
    Else
        HScroll1.Visible = False
        pcbLine.Visible = False
    End If
    DrawResizeLine
    'UserControl.ScaleMode = m_ScaleModeParent
End Sub

Private Sub DrawButtons()
    Dim Triangle(3) As PointApi, Y As Integer
    'UserControl.ScaleMode = vbPixels
    With pcbButtons
        .BackColor = UserControl.BackColor
        .Cls
        If UserControl.Enabled = False Then
            pcbButtons.FillColor = UserControl.BackColor
            pcbButtons.ForeColor = vbGrayText
        Else
            pcbButtons.FillColor = UserControl.ForeColor
            pcbButtons.ForeColor = UserControl.ForeColor
        End If
        Y = (.ScaleHeight - TRI_HEIGHT) / 2
        pcbButtons.Line (BEGIN_BUTTONS, Y)-(BEGIN_BUTTONS, Y + TRI_HEIGHT)
        Triangle(0).X = BEGIN_BUTTONS + 3
        Triangle(0).Y = .ScaleHeight / 2
        Triangle(1).X = BEGIN_BUTTONS + TRI_WIDHT + 3
        Triangle(1).Y = Y
        Triangle(2).X = Triangle(1).X
        Triangle(2).Y = Y + TRI_HEIGHT
        Call Polygon(.hdc, Triangle(0), 3)
        Triangle(0).X = Triangle(0).X + SEPARATION_BUTTONS
        Triangle(1).X = Triangle(1).X + SEPARATION_BUTTONS
        Triangle(2).X = Triangle(1).X
        Call Polygon(.hdc, Triangle(0), 3)
        Triangle(0).X = Triangle(1).X + SEPARATION_BUTTONS + TRI_WIDHT
        Triangle(1).X = Triangle(1).X + SEPARATION_BUTTONS
        Triangle(2).X = Triangle(1).X
        Call Polygon(.hdc, Triangle(0), 3)
        Triangle(0).X = Triangle(0).X + SEPARATION_BUTTONS
        Triangle(1).X = Triangle(1).X + SEPARATION_BUTTONS
        Triangle(2).X = Triangle(1).X
        Call Polygon(.hdc, Triangle(0), 3)
        pcbButtons.Line (Triangle(0).X + 3, Y)-(Triangle(0).X + 3, Y + TRI_HEIGHT)
    End With
    'UserControl.ScaleMode = m_ScaleModeParent
End Sub

Private Sub DrawResizeLine()
    'UserControl.ScaleMode = vbPixels
    With pcbLine
        .Cls
        .ForeColor = vb3DLight
        pcbLine.Line (0, .ScaleHeight)-(0, 0)
        pcbLine.Line -(.ScaleWidth, 0)
        .ForeColor = vb3DHighlight
        pcbLine.Line (1, .ScaleHeight - 1)-(1, 1)
        pcbLine.Line -(.ScaleWidth - 1, 1)
        .ForeColor = vb3DShadow
        pcbLine.Line (.ScaleWidth - 2, 1)-(.ScaleWidth - 2, .ScaleHeight - 2)
        pcbLine.Line -(0, .ScaleHeight - 2)
        .ForeColor = vb3DDKShadow
        pcbLine.Line (.ScaleWidth - 1, 0)-(.ScaleWidth - 1, .ScaleHeight - 1)
        pcbLine.Line -(0, .ScaleHeight - 1)
    End With
    'UserControl.ScaleMode = m_ScaleModeParent
End Sub

Private Sub IniFont()
    Dim i As Integer
    For i = 0 To ctlTab1.UBound
        Set ctlTab1(i).Font = UserControl.Font
        ctlTab1(i).ForeColor = UserControl.ForeColor
        ctlTab1(i).FontBold = m_FontBold
        ctlTab1(i).FontItalic = UserControl.FontItalic
        ctlTab1(i).FontUnderline = UserControl.FontUnderline
        ctlTab1(i).FontStrikethru = UserControl.FontStrikethru
    Next
End Sub

Private Function KnowFirst()
    Dim i As Integer
    KnowFirst = 1
    'UserControl.ScaleMode = vbPixels
    For i = 1 To ctlTab1.UBound
        With ctlTab1(i)
            If pcbTab.Left + .Left + BSIDE > m_EndButtons Then
                KnowFirst = i
                Exit For
            End If
        End With
    Next
    'UserControl.ScaleMode = m_ScaleModeParent
End Function

Private Sub ReorderTabs()
    Dim i As Integer, strCaption As Variant
    If m_ColItems.Count < 1 Then
        Exit Sub
    End If
    'UserControl.ScaleMode = vbPixels
    pcbTab.Visible = False
    i = 1
    For Each strCaption In m_ColItems
        If ctlTab1(i).Caption <> CStr(strCaption) Then
            ctlTab1(i).Caption = CStr(strCaption)
        End If
        i = i + 1
    Next
    ctlTab1(1).Left = 0
    For i = 2 To ctlTab1.UBound
        ctlTab1(i).Left = ctlTab1(i - 1).Left + ctlTab1(i - 1).Width - 1.2 * BSIDE
    Next i
    pcbTab.Width = ctlTab1(ctlTab1.UBound).Left + ctlTab1(ctlTab1.UBound).Width + BSIDE
    pcbTab.Visible = True
    'UserControl.ScaleMode = m_ScaleModeParent
End Sub

Private Sub IniAll()
    IniFont
    'm_ScaleModeParent = UserControl.Parent.ScaleMode
    If Ambient.UserMode = False Then
        IniExample
    End If
    With Me
        .Appearance = m_Appearance
        .Placement = m_Placement
        .ShowNavigator = m_ShowNavigator
        .ShowHscroll = m_ShowHscroll
        .BackColor = UserControl.BackColor
        .Enabled = UserControl.Enabled
        .BorderStyle = m_BorderStyle
    End With
    Refresh
    'UserControl.ScaleMode = m_ScaleModeParent
End Sub

Private Sub IniExample()
    'UserControl.ScaleMode = vbPixels
    pcbTab.Left = m_EndButtons
    With ctlTab1(0)
        .Left = 0
        .Visible = True
    End With
    With tabSelected
        .Left = ctlTab1(0).Width - 1.2 * BSIDE
        .Appearance = m_Appearance
        .BackColor = m_SelColor
        .Placement = m_Placement
        .FontName = ctlTab1(0).FontName
        .FontItalic = ctlTab1(0).FontItalic
        .FontSize = ctlTab1(0).FontSize
        .FontStrikethru = ctlTab1(0).FontStrikethru
        .FontUnderline = ctlTab1(0).FontUnderline
        .ForeColor = ctlTab1(0).ForeColor
        .FontBold = True
        .Visible = True
    End With
    'UserControl.ScaleMode = m_ScaleModeParent
End Sub
'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=7,0,0,0
Public Property Get BorderStyle() As Border
Attribute BorderStyle.VB_Description = "Devuelve o establece el estilo del borde de un objeto."
    BorderStyle = m_BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Border)
    Dim tmpForeColor As Long
    m_BorderStyle = New_BorderStyle
    With UserControl
        .Cls
        If m_BorderStyle = FixedSingle Then
            If m_Appearance = x3D Then
                .ForeColor = vb3DHighlight
                UserControl.Line (0, .ScaleHeight - 1)-(0, 0)
                UserControl.Line -(.ScaleWidth - 1, 0)
                .ForeColor = vb3DShadow
                UserControl.Line -(.ScaleWidth - 1, .ScaleHeight - 1)
                UserControl.Line -(0, .ScaleHeight - 1)
            Else
                .ForeColor = vb3DDKShadow
                UserControl.Line (0, .ScaleHeight - 1)-(0, 0)
                UserControl.Line -(.ScaleWidth - 1, 0)
                UserControl.Line -(.ScaleWidth - 1, .ScaleHeight - 1)
                UserControl.Line -(0, .ScaleHeight - 1)
            End If
        End If
        .ForeColor = tmpForeColor
    End With
    PropertyChanged "BorderStyle"
End Property

