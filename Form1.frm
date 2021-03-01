VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3660
   ClientLeft      =   5880
   ClientTop       =   4575
   ClientWidth     =   5550
   LinkTopic       =   "Form1"
   ScaleHeight     =   3660
   ScaleWidth      =   5550
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4320
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Print form"
      Height          =   495
      Left            =   2040
      TabIndex        =   6
      Top             =   2400
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   600
      List            =   "Form1.frx":0002
      TabIndex        =   5
      Text            =   "Combo1"
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get time and date"
      Height          =   495
      Left            =   3360
      TabIndex        =   4
      Top             =   480
      Width           =   1695
   End
   Begin VB.ListBox List2 
      Height          =   1230
      ItemData        =   "Form1.frx":0004
      Left            =   1920
      List            =   "Form1.frx":0006
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Height          =   1230
      ItemData        =   "Form1.frx":0008
      Left            =   600
      List            =   "Form1.frx":001B
      MultiSelect     =   1  'Simple
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   915
      Left            =   3405
      TabIndex        =   3
      Top             =   1200
      Width           =   1605
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "DblClick Here!"
      Height          =   195
      Left            =   810
      TabIndex        =   0
      Top             =   480
      Width           =   1035
   End
   Begin VB.Menu MnuFile 
      Caption         =   "&File"
      Begin VB.Menu MnuFile_Open 
         Caption         =   "O&pen"
         Shortcut        =   ^O
      End
   End
   Begin VB.Menu MnuPref 
      Caption         =   "&Preferences"
      NegotiatePosition=   3  'Right
      Begin VB.Menu MnuPref_Backcolor 
         Caption         =   "Back Color"
         Shortcut        =   ^B
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuPref_Font 
         Caption         =   "Font"
         Shortcut        =   ^F
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type Person
    Name As String
    Age As Integer
End Type

user As Person

Private Sub Form_Load()
    user.Name = "Mohammed"
    user.Age = 19
    Static arr(3 To 5) As Integer
    For i = 3 To 5
        Combo1.AddItem arr(i)
    Next i
End Sub

Private Sub Form_MouseMove(btn As Integer, shift As Integer, x As Single, y As Single)
    ' x and y must be type of Single
    Me.Caption = "Hello " & user.Name & ": (" & x & ", " & y & ")"
End Sub

' Private Sub Combo1_Change() emmitted for each new character typed
Private Sub Combo1_LostFocus()
    If Not Cbo1Contains(Combo1.Text) Then
        Combo1.AddItem Combo1.Text
    End If
End Sub

Private Function Cbo1Contains(item As String)
    Dim exists As Boolean
    For i = 0 To Combo1.ListCount - 1
        If Combo1.List(i) = item Then
            exists = True
        End If
    Next i
    Cbo1Contains = exists
End Function

Private Sub Command1_Click()
    ' Label2.Caption = Time & vbCrLf & Date
    Label2.Caption = CStr(Time) & vbCrLf & CStr(Date)
End Sub

Private Sub Command2_Click()
    Form1.PrintForm
End Sub

Private Sub Label1_DblClick()
    inp = InputBox("Input an arbitrary integer")
    msg = IIf(inp < 10, 1, 0)
    Select Case msg
        Case Is = 1
            MsgBox "Good prediction!"
        Case Is = 0
            MsgBox "Try again!"
        Case 3 To 5
        Case Else
            Err.Raise "Unexpected value for ""msg""!"
    End Select
End Sub

Private Sub List1_Click()
    List2.Clear
    For i = 0 To List1.ListCount - 1
        If List1.Selected(i) Then List2.AddItem List1.List(i)
    Next i
    List2.AddItem asd
End Sub

Private Sub MnuFile_Open_Click()
    Call MsgBox("Opening file", vbAbortRetryIgnore, vbInformation)
End Sub

Private Sub MnuPref_Backcolor_Click()
    CommonDialog1.ShowColor
    ' Me.BackColor = CommonDialog1.Color
    Label2.BackColor = CommonDialog1.Color
End Sub

Private Sub MnuPref_Font_Click()
    CommonDialog1.ShowFont
    Label2.FontName = CommonDialog1.FontName
    Label2.FontSize = CommonDialog1.FontSize
    Label2.FontBold = CommonDialog1.FontBold
    Label2.FontItalic = CommonDialog1.FontItalic
End Sub

