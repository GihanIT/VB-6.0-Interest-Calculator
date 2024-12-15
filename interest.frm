VERSION 5.00
Begin VB.Form form1 
   Caption         =   "Interest Calculator"
   ClientHeight    =   6345
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7755
   LinkTopic       =   "Form2"
   ScaleHeight     =   6345
   ScaleWidth      =   7755
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      TabIndex        =   11
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   10
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calculate"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   9
      Top             =   5040
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      Height          =   405
      Left            =   3120
      TabIndex        =   8
      Top             =   3960
      Width           =   2415
   End
   Begin VB.TextBox Text3 
      Height          =   405
      Left            =   3120
      TabIndex        =   7
      Top             =   3360
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      Caption         =   "Interest for Deposits"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   240
      TabIndex        =   4
      Top             =   1680
      Width           =   6975
      Begin VB.OptionButton Option2 
         Caption         =   "12  Month Deposits -(15 % Interest)"
         Height          =   495
         Left            =   600
         TabIndex        =   13
         Top             =   720
         Width           =   2895
      End
      Begin VB.OptionButton Option1 
         Caption         =   "6 Month Deposits -(12 % Interest)"
         Height          =   255
         Left            =   600
         TabIndex        =   12
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.TextBox Text2 
      Height          =   405
      Left            =   3120
      TabIndex        =   3
      Top             =   840
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   405
      Left            =   3120
      TabIndex        =   2
      Top             =   240
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "https://github.com/GihanIT/VB-6.0-Interest-Calculator"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   14
      Top             =   5880
      Width           =   4695
   End
   Begin VB.Label Label6 
      Caption         =   "Total amount with Interest"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   4080
      Width           =   2535
   End
   Begin VB.Label Label5 
      Caption         =   "Interest amount"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   3480
      Width           =   1575
   End
   Begin VB.Label Label2 
      Caption         =   "Amount"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Customer No"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = "" Then
e = MsgBox("Plase enter Customer No", 16, "Customer No is required")
End If

Dim a As Double
Dim b As Double
Dim c As Double

Dim d As Double
Dim x As Double



a = Text2.Text
b = a * 12 / 100
c = a * 15 / 100

 d = a + b
 x = a + c
 
 If Option1 = True Then
Text3.Text = b

Else
Text3.Text = c
End If
 
 
 
 
If Option1 = True Then
Text4.Text = d

Else
Text4.Text = x
End If








End Sub

Private Sub Label4_Click()

End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""


End Sub

Private Sub Command3_Click()
End
End Sub

