VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "colTabs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"Tabs"
Attribute VB_Ext_KEY = "Member0" ,"Tabs"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'variable local para contener colección
Private mCol As Collection
Private m_Control As ctlTab
Event Ninguno()

Public Function Add(Tabx As ctlTab, Widthx As Single, Heightx As Single, Selected As Boolean, Caption As String, Index As Integer, Key As String, Leftx As Single, Topx As Single, Tag As Variant, ToolTipText As String, Optional sKey As String) As Tabs
    'crear un nuevo objeto
    Dim objNewMember As Tabs
    Set objNewMember = New Tabs

    'establecer las propiedades que se transfieren al método
    If IsObject(Tabx) Then
        Set objNewMember.Tabx = Tabx
    Else
        objNewMember.Tabx = Tabx
    End If
    objNewMember.Widthx = Widthx
    objNewMember.Heightx = Heightx
    objNewMember.Selected = Selected
    objNewMember.Caption = Caption
    objNewMember.Index = Index
    objNewMember.Key = Key
    objNewMember.Leftx = Leftx
    objNewMember.Topx = Topx
    If IsObject(Tag) Then
        Set objNewMember.Tag = Tag
    Else
        objNewMember.Tag = Tag
    End If
    objNewMember.ToolTipText = ToolTipText
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'devolver el objeto creado
    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As Tabs
Attribute Item.VB_UserMemId = 0
  Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
    mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
    Set mCol = New Collection
    Load xssTabBar.ctlTab1(25)
    xssTabBar.ctlTab1(25).Visible = True
    MsgBox ""
End Sub

Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub
