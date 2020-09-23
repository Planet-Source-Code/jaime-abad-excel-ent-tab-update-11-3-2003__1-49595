VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Tabs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'variables locales para almacenar los valores de las propiedades
Private m_Tabx As ctlTab
Private m_Selected As Boolean
Private m_Index As Integer
Private m_Key As String

Public Property Let ToolTipText(ByVal vData As String)
    m_ToolTipText = vData
End Property

Public Property Get ToolTipText() As String
    ToolTipText = m_ToolTipText
End Property

Public Property Let Tag(ByVal vData As Variant)
    m_Tag = vData
End Property

Public Property Set Tag(ByVal vData As Variant)
    Set m_Tag = vData
End Property

Public Property Get Tag() As Variant
    Tag = m_Tag
End Property

Friend Property Let Top(ByVal vData As Single)
    m_Tabx.Top = vData
End Property

Public Property Get Top() As Single
    Top = m_Tabx.Top
End Property

Friend Property Let Left(ByVal vData As Single)
    m_Tabx.Left = vData
End Property

Public Property Get Left() As Single
    Left = m_Tabx.Left
End Property

Friend Property Let Key(ByVal vData As String)
    m_Key = vData
End Property

Public Property Get Key() As String
    Key = m_Key
End Property

Friend Property Let Index(ByVal vData As Integer)
    m_Index = vData
End Property

Public Property Get Index() As Integer
    Index = m_Index
End Property

Public Property Let Caption(ByVal vData As String)
    m_Tabx.Caption = vData
End Property

Public Property Get Caption() As String
Attribute Caption.VB_UserMemId = 0
    Caption = m_Tabx.Caption
End Property

Public Property Let Selected(ByVal vData As Boolean)
    m_Selected = vData
End Property

Public Property Get Selected() As Boolean
    Selected = m_Selected
End Property

Friend Property Let Height(ByVal vData As Single)
    m_Tabx.Height = vData
End Property

Public Property Get Height() As Single
    Heightx = m_Tabx.Height
End Property

Friend Property Let Width(ByVal vData As Single)
    m_Tabx.Width = vData
End Property

Public Property Get Width() As Single
    Width = m_Tabx.Width
End Property

Friend Property Set Tabx(ByVal vData As ctlTab)
    Set m_Tabx = vData
End Property

Friend Property Get Tabx() As ctlTab
    Set Tabx = m_Tabx
End Property
