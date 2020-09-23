VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Eval 3.0 | http://www.semicolonsoftware.de"
   ClientHeight    =   3660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   4680
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton cmdCalc 
      Caption         =   "Calculate"
      Height          =   255
      Left            =   3600
      TabIndex        =   2
      Top             =   1440
      Width           =   855
   End
   Begin VB.TextBox txtExpression 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "2*(3+4)^2"
      Top             =   1440
      Width           =   3375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      Height          =   495
      Left            =   120
      Top             =   600
      Width           =   4455
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "potentiate"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   4
      Left            =   3240
      TabIndex        =   19
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "divide"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   3
      Left            =   2760
      TabIndex        =   18
      Top             =   840
      Width           =   375
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "multiply"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   2
      Left            =   2160
      TabIndex        =   17
      Top             =   840
      Width           =   495
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "substract"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   1
      Left            =   1485
      TabIndex        =   16
      Top             =   840
      Width           =   585
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "add"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   0
      Left            =   1155
      TabIndex        =   15
      Top             =   840
      Width           =   225
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "^"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   3480
      TabIndex        =   14
      Top             =   600
      Width           =   135
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   2880
      TabIndex        =   13
      Top             =   600
      Width           =   60
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   2400
      TabIndex        =   12
      Top             =   600
      Width           =   90
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   1710
      TabIndex        =   11
      Top             =   600
      Width           =   75
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   1200
      TabIndex        =   10
      Top             =   600
      Width           =   135
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "operator list:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   210
      Left            =   165
      TabIndex        =   9
      Top             =   735
      Width           =   900
   End
   Begin VB.Label Label4 
      Caption         =   "Eval 3.0 | pure vb | no libraries required | 100% by me | supports brackets, basic maths"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label lblInfo 
      Caption         =   "See modEval.bas for information on how to use this in your own programs..."
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   3120
      Width           =   4335
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "expression converted into:"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   1860
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "result:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   420
   End
   Begin VB.Label lblEvalExpression 
      BorderStyle     =   1  'Fest Einfach
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   4335
   End
   Begin VB.Label Label1 
      Caption         =   "enter expression:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   3255
   End
   Begin VB.Label lblResult 
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "0"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   2160
      Width           =   4335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCalc_Click()

lblResult.Caption = Eval(txtExpression.Text)
lblEvalExpression.Caption = BracketsForExpression(txtExpression.Text)
End Sub

Private Sub txtExpression_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then cmdCalc_Click
End Sub
