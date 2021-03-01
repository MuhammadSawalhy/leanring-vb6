VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmFiles 
   Caption         =   "Files"
   ClientHeight    =   6090
   ClientLeft      =   6075
   ClientTop       =   2730
   ClientWidth     =   9465
   LinkTopic       =   "Form2"
   ScaleHeight     =   6090
   ScaleWidth      =   9465
   Begin VB.CommandButton Command6 
      Caption         =   "Read"
      Height          =   615
      Left            =   6480
      TabIndex        =   15
      Top             =   3600
      Width           =   2535
   End
   Begin VB.TextBox Text4 
      Height          =   2295
      Left            =   6480
      MultiLine       =   -1  'True
      TabIndex        =   14
      Text            =   "FrmFiles.frx":0000
      Top             =   960
      Width           =   2535
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   4200
      TabIndex        =   13
      Text            =   "19"
      Top             =   3240
      Width           =   1455
   End
   Begin VB.ListBox List3 
      Height          =   2010
      Left            =   4200
      TabIndex        =   12
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Add To The List"
      Height          =   495
      Left            =   4200
      TabIndex        =   11
      Top             =   3720
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   840
      TabIndex        =   10
      Text            =   "Eng."
      Top             =   3240
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   840
      TabIndex        =   9
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add To The List"
      Height          =   495
      Left            =   840
      TabIndex        =   8
      Top             =   3720
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Read"
      Height          =   615
      Left            =   3360
      TabIndex        =   4
      Top             =   4680
      Width           =   2295
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Save"
      Height          =   615
      Left            =   840
      TabIndex        =   3
      Top             =   4680
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2520
      TabIndex        =   2
      Text            =   "Mohammed"
      Top             =   3240
      Width           =   1455
   End
   Begin VB.ListBox List2 
      Height          =   2010
      Left            =   2520
      TabIndex        =   1
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Add To The List"
      Height          =   495
      Left            =   2520
      TabIndex        =   0
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Line Line1 
      BorderStyle     =   3  'Dot
      X1              =   6480
      X2              =   6480
      Y1              =   1560
      Y2              =   5880
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Name"
      Height          =   195
      Left            =   2520
      TabIndex        =   7
      Top             =   720
      Width           =   420
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Age"
      Height          =   195
      Left            =   4200
      TabIndex        =   6
      Top             =   720
      Width           =   285
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Title"
      Height          =   195
      Left            =   840
      TabIndex        =   5
      Top             =   720
      Width           =   300
   End
End
Attribute VB_Name = "FrmFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type Person
    Name As String
    Title As String * 5
    Age As Integer
End Type

Private Sub Command1_Click()
   List1.AddItem Text1.Text
End Sub

Private Sub Command2_Click()
   List2.AddItem Text2.Text
End Sub

Private Sub Command3_Click()
   List3.AddItem Text3.Text
End Sub

Private Sub Command4_Click()
    Dim per As Person
    MsgBox Len(per)
    CommonDialog1.ShowSave
    Open CommonDialog1.FileName For Random As #1 Len = Len(per)
        For i = 0 To List1.ListCount - 1
            per.Title = List1.List(i)
            per.Name = List2.List(i)
            per.Age = Val(List3.List(i))
            Put #1, i, per
        Next i
    Close #1
End Sub

Private Sub Command5_Click()
    Dim per As Person
    CommonDialog1.ShowOpen
    Open CommonDialog1.FileName For Random As #1 Len = Len(per)
        While (Not EOF(1))
            Get #1, , per
            List1.AddItem per.Title
            List2.AddItem per.Name
            List3.AddItem per.Age
        Wend
    Close #1
End Sub

Private Sub Command6_Click()
    Dim readed As String
    CommonDialog1.ShowOpen
    Open CommonDialog1.FileName For Input As #1
        Line Input #1, k
        If readed = "" Then
            readed = k
        Else
            readed = readed & vbCrLf & k
        End If
    Close #1
    Text4.Text = readed
End Sub
