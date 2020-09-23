VERSION 5.00
Object = "*\AXssTab.vbp"
Begin VB.Form frmTest 
   Caption         =   "Tab Test"
   ClientHeight    =   5175
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   6810
   BeginProperty Font 
      Name            =   "MS Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTest.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5175
   ScaleWidth      =   6810
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pcbExample 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
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
      Height          =   1995
      Left            =   0
      ScaleHeight     =   1995
      ScaleWidth      =   6810
      TabIndex        =   23
      Top             =   2910
      Width           =   6810
      Begin VB.TextBox txtLog 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1875
         Left            =   0
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   24
         Text            =   "frmTest.frx":0B66
         Top             =   120
         Width           =   6825
      End
   End
   Begin VB.Frame frmScroll 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Scroll"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   15
      TabIndex        =   6
      Top             =   2070
      Width           =   6735
      Begin VB.TextBox txtScoll 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   5490
         TabIndex        =   20
         Text            =   "1"
         Top             =   450
         Width           =   795
      End
      Begin VB.TextBox txtScoll 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   4410
         TabIndex        =   18
         Text            =   "1"
         Top             =   450
         Width           =   795
      End
      Begin VB.TextBox txtScoll 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   3450
         TabIndex        =   13
         Text            =   "2"
         Top             =   450
         Width           =   795
      End
      Begin VB.TextBox txtScoll 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   2520
         TabIndex        =   11
         Text            =   "10"
         Top             =   450
         Width           =   795
      End
      Begin VB.TextBox txtScoll 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   1560
         TabIndex        =   9
         Text            =   "0"
         Top             =   450
         Width           =   795
      End
      Begin VB.CheckBox chkScroll 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Resizeable"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   210
         TabIndex        =   8
         Top             =   570
         Value           =   1  'Checked
         Width           =   1245
      End
      Begin VB.CheckBox chkScroll 
         BackColor       =   &H00C0FFFF&
         Caption         =   "ShowScroll"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   210
         TabIndex        =   7
         Top             =   300
         Value           =   1  'Checked
         Width           =   1545
      End
      Begin VB.Label lblScroll 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "SmallChange"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   5460
         TabIndex        =   21
         Top             =   180
         Width           =   930
      End
      Begin VB.Label lblScroll 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "LageChange"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   4350
         TabIndex        =   19
         Top             =   180
         Width           =   915
      End
      Begin VB.Label lblScroll 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "Value"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   3570
         TabIndex        =   14
         Top             =   180
         Width           =   405
      End
      Begin VB.Label lblScroll 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "Max"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   2700
         TabIndex        =   12
         Top             =   180
         Width           =   300
      End
      Begin VB.Label lblScroll 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFFF&
         Caption         =   "Min"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   1860
         TabIndex        =   10
         Top             =   210
         Width           =   255
      End
   End
   Begin VB.Frame frmGeneral 
      BackColor       =   &H00FFC0FF&
      Caption         =   "General"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   15
      TabIndex        =   4
      Top             =   1020
      Width           =   6735
      Begin VB.CheckBox chkGeneral 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Enabled"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   120
         TabIndex        =   26
         Top             =   630
         Value           =   1  'Checked
         Width           =   1035
      End
      Begin VB.CheckBox chkGeneral 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Moveable Tabs"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   1620
         TabIndex        =   17
         Top             =   330
         Value           =   1  'Checked
         Width           =   1425
      End
      Begin VB.OptionButton optGeneral 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Placement Bottom"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   4530
         TabIndex        =   16
         Top             =   300
         Value           =   -1  'True
         Width           =   1635
      End
      Begin VB.OptionButton optGeneral 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Placement Top"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   3060
         TabIndex        =   15
         Top             =   300
         Width           =   1485
      End
      Begin VB.CheckBox chkGeneral 
         BackColor       =   &H00FFC0FF&
         Caption         =   "ShowNavigator"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   120
         TabIndex        =   5
         Top             =   330
         Value           =   1  'Checked
         Width           =   1545
      End
   End
   Begin VB.Frame frmMouseWheel 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Mouse Wheel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   15
      TabIndex        =   0
      Top             =   330
      Width           =   6735
      Begin VB.OptionButton optMouseWheel 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Move Selected Tab"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   2940
         TabIndex        =   22
         Top             =   270
         Width           =   1875
      End
      Begin VB.OptionButton optMouseWheel 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Nothing"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   5040
         TabIndex        =   3
         Top             =   270
         Width           =   1575
      End
      Begin VB.OptionButton optMouseWheel 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Move Tab"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   1620
         TabIndex        =   2
         Top             =   270
         Value           =   -1  'True
         Width           =   1185
      End
      Begin VB.OptionButton optMouseWheel 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Move Scroll"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   180
         TabIndex        =   1
         Top             =   270
         Width           =   1575
      End
   End
   Begin XssTab.xssTabBar xssTabBar1 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   25
      Top             =   4905
      Width           =   6810
      _ExtentX        =   12012
      _ExtentY        =   476
      SelColor        =   -2147483634
      ScrollValue     =   2
      ScrollMax       =   10
      ScrollMax       =   10
      ScrollValue     =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
   End
   Begin XssTab.xssTabBar xssTabBar2 
      Align           =   1  'Align Top
      Height          =   285
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   6810
      _ExtentX        =   12012
      _ExtentY        =   503
      Enabled         =   0   'False
      SelColor        =   -2147483643
      ShowHscroll     =   0   'False
      ScrollMax       =   2
      Resizable       =   0   'False
      Placement       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      BorderStyle     =   0
   End
   Begin VB.Menu mnuAll 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuRename 
         Caption         =   "Rename"
      End
      Begin VB.Menu mnuAddBefore 
         Caption         =   "Add Tab Before"
      End
      Begin VB.Menu mnuAddAfter 
         Caption         =   "Add Tab After"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete Tab"
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "Select all"
      End
   End
   Begin VB.Menu mnuButtons 
      Caption         =   "mnuButtons"
      Visible         =   0   'False
      Begin VB.Menu mnuFirst 
         Caption         =   "Got2First"
      End
      Begin VB.Menu mnuBack 
         Caption         =   "GoBack"
      End
      Begin VB.Menu mnuNext 
         Caption         =   "GoNext"
      End
      Begin VB.Menu mnuEnd 
         Caption         =   "Go2End"
      End
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'JAIME ABAD 1/11/2003

Option Explicit

Dim m_Ini_Selection As Integer
Dim m_Selected_AllTab As Boolean

Private Sub chkGeneral_Click(Index As Integer)
    xssTabBar1.shownavigator = chkGeneral(1).Value
    xssTabBar1.MoveableTabs = chkGeneral(0).Value
    xssTabBar1.Enabled = chkGeneral(2).Value
    If xssTabBar1.Enabled Then
        xssTabBar1.SetFocus
    End If
End Sub

Private Sub chkScroll_Click(Index As Integer)
    xssTabBar1.ShowHscroll = chkScroll(0).Value
    xssTabBar1.Resizable = chkScroll(1).Value
    If xssTabBar1.Enabled Then
        xssTabBar1.SetFocus
    End If
End Sub

Private Sub Form_Activate()
    xssTabBar1.SetFocus
End Sub

Private Sub mnuAddAfter_Click()
    Dim strInput As String
    strInput = InputBox("Enter the name of a new tab", "New Tab")
    DoEvents
    If strInput <> "" Then
        xssTabBar1.AddNew strInput, after:=xssTabBar1.SelectItem
    End If
    txtLog.Text = txtLog.Text & vbCrLf & "Tab Count= " & xssTabBar1.Count
End Sub

Private Sub mnuAddBefore_Click()
    Dim strInput As String
    strInput = InputBox("Enter the name of a new tab", "New Tab")
    DoEvents
    If strInput <> "" Then
        xssTabBar1.AddNew strInput, xssTabBar1.SelectItem
    End If
    txtLog.Text = txtLog.Text & vbCrLf & "Tab Count= " & xssTabBar1.Count
End Sub

Private Sub mnuDelete_Click()
    xssTabBar1.Remove xssTabBar1.SelectItem
    txtLog.Text = txtLog.Text & vbCrLf & "Tab Count= " & xssTabBar1.Count
End Sub

Private Sub mnuRename_Click()
    xssTabBar1.EditMode
End Sub

Private Sub mnuSelectAll_Click()
    Dim i As Integer
    xssTabBar1.MultiSelect = True
    For i = 1 To xssTabBar1.Count
        xssTabBar1.SelectItem = i
    Next
    m_Selected_AllTab = True
End Sub

Private Sub xssTabBar1_ChangeTab(Old_Selection As Integer, New_Selection As Integer)
    txtLog.Text = txtLog.Text & vbCrLf & "Change Tab Selection OldIndex= " & Old_Selection _
       & " NewIndex= " & New_Selection
End Sub

Private Sub xssTabBar1_ClickNavigator(Index As Integer)
    txtLog.Text = txtLog.Text & vbCrLf & "Click in Navigator Index= " & Index
    Select Case Index
        Case 0
            xssTabBar1.MoveFirst
        Case 1
            xssTabBar1.MoveRight
        Case 2
            xssTabBar1.MoveLeft
        Case 3
            xssTabBar1.MoveEnd
    End Select
End Sub

Private Sub xssTabBar1_ClickTab(Index As Integer)
    txtLog.Text = txtLog.Text & vbCrLf & "Click Tab Index= " & Index
    If xssTabBar1.MultiSelect Then
        xssTabBar1.MultiSelect = False
    End If
End Sub

Private Sub xssTabBar1_DblClickNavigator(Index As Integer)
    txtLog.Text = txtLog.Text & vbCrLf & "Doble click in Navigator Index= " & Index
End Sub

Private Sub xssTabBar1_MouseDownNavigator(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtLog.Text = txtLog.Text & vbCrLf & "MouseDown in Navigator Index= " & Index & " Button= " & Button & " Shift= " & Shift _
        & " X= " & X & " Y= " & Y
    Select Case Button
        Case 1
        Case 2
            PopupMenu mnuButtons
        Case 4
    End Select
End Sub

Private Sub xssTabBar1_MouseDownTab(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtLog.Text = txtLog.Text & vbCrLf & "MouseDown Tab Index= " & Index & " Button= " & Button _
        & " Shift= " & Shift
    Select Case Button
        Case 1
        Case 2
            PopupMenu mnuAll
        Case 4
    End Select
End Sub

Private Sub xssTabBar1_MouseUpNavigator(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtLog.Text = txtLog.Text & vbCrLf & "MouseUp in Navigator Index= " & Index & " Button= " & Button & " Shift= " & Shift _
        & " X= " & X & " Y= " & Y
End Sub

Private Sub xssTabBar1_MouseUpTab(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    txtLog.Text = txtLog.Text & vbCrLf & "MouseUp Tab Index= " & Index & " Button= " & Button _
        & " Shift= " & Shift
End Sub

Private Sub xssTabBar1_MouseWheelDown()
    txtLog.Text = txtLog.Text & vbCrLf & "Mouse Wheel Down"
    If optMouseWheel(0).Value = True Then
        xssTabBar1.ScrollValue = xssTabBar1.ScrollValue + 1
    ElseIf optMouseWheel(1).Value = True Then
        xssTabBar1.MoveLeft
    ElseIf optMouseWheel(3).Value = True Then
        xssTabBar1.SelectItem = xssTabBar1.SelectItem + 1
    End If
End Sub

Private Sub xssTabBar1_MouseWheelUp()
    txtLog.Text = txtLog.Text & vbCrLf & "Mouse Wheel Up"
    If optMouseWheel(0).Value = True Then
        xssTabBar1.ScrollValue = xssTabBar1.ScrollValue - 1
    ElseIf optMouseWheel(1).Value = True Then
        xssTabBar1.MoveRight
    ElseIf optMouseWheel(3).Value = True Then
        xssTabBar1.SelectItem = xssTabBar1.SelectItem - 1
    End If
End Sub

Private Sub xssTabBar1_DblClickTab(Index As Integer)
    txtLog.Text = txtLog.Text & vbCrLf & "Doble click in Tab Index= " & Index
    xssTabBar1.EditMode
    txtLog.Text = txtLog.Text & vbCrLf & "Start Edit mode"
End Sub

Private Sub xssTabBar1_ExitEdit(Text As String)
    xssTabBar1.Caption(xssTabBar1.SelectItem) = Text
    txtLog.Text = txtLog.Text & vbCrLf & "Exit Edit mode Text= " & Text
End Sub

Private Sub xssTabBar1_ResizeHScroll(New_Offset As Single)
    txtLog.Text = txtLog.Text & vbCrLf & " ResizeScroll  New_Offset= " & New_Offset
End Sub

Private Sub xssTabBar1_Scroll()
    txtScoll(2) = xssTabBar1.ScrollValue
    txtLog.Text = txtLog.Text & vbCrLf & "Scroll Value= " & xssTabBar1.ScrollValue
End Sub

Private Sub xssTabBar1_ScrollChange()
    txtScoll(2) = xssTabBar1.ScrollValue
    txtLog.Text = txtLog.Text & vbCrLf & "Scroll Change= " & xssTabBar1.ScrollValue
End Sub

Private Sub Form_Load()
    Dim i As Integer
    With xssTabBar1
        .AddNew "Jaime"
        .AddNew "Abad"
        For i = 65 To 75
            .AddNew "Tab " & Chr(i)
        Next
        txtLog.Text = txtLog.Text & vbCrLf & "Tab Count= " & .Count
    End With
    With xssTabBar2
        .AddNew "Disable"
        .AddNew "Other Color"
        .AddNew "Without Scroll"
    End With
End Sub

Private Sub optGeneral_Click(Index As Integer)
    If optGeneral(0).Value = True Then
        xssTabBar1.Placement = plcTop
    ElseIf optGeneral(1).Value = True Then
        xssTabBar1.Placement = plcBottom
    End If
    If xssTabBar1.Enabled Then
        xssTabBar1.SetFocus
    End If
End Sub

Private Sub optMouseWheel_Click(Index As Integer)
    If xssTabBar1.Enabled Then
        xssTabBar1.SetFocus
    End If
End Sub

Private Sub txtScoll_Validate(Index As Integer, Cancel As Boolean)
    With xssTabBar1
        .ScrollMin = txtScoll(0).Text
        .ScrollMax = txtScoll(1).Text
        .ScrollValue = txtScoll(2).Text
        .LargeChange = txtScoll(3).Text
        .SmallChange = txtScoll(4).Text
    End With
    xssTabBar1.SetFocus
End Sub
