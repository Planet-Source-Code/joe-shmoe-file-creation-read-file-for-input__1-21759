VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6675
   LinkTopic       =   "Form1"
   ScaleHeight     =   1500
   ScaleWidth      =   6675
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   495
      Left            =   5040
      TabIndex        =   16
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   4680
      TabIndex        =   15
      Top             =   480
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Birth Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   4
      Top             =   840
      Width           =   4935
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   3960
         TabIndex        =   14
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2400
         TabIndex        =   13
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   720
         TabIndex        =   12
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblMonth 
         Caption         =   "Month:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblDay 
         Caption         =   "Day:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   6
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lblYear 
         Caption         =   "Year:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.TextBox txtLastname 
      Height          =   285
      Left            =   3720
      TabIndex        =   3
      Top             =   0
      Width           =   1935
   End
   Begin VB.OptionButton optMarriedY 
      Caption         =   "Yes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   2
      Top             =   480
      Value           =   -1  'True
      Width           =   735
   End
   Begin VB.OptionButton optMarriedN 
      Caption         =   "No"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   1
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox txtFirstname 
      Height          =   285
      Left            =   1080
      TabIndex        =   0
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label lblFirstname 
      AutoSize        =   -1  'True
      Caption         =   "First Name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   1020
   End
   Begin VB.Label lblLastname 
      AutoSize        =   -1  'True
      Caption         =   "Last Name:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2640
      TabIndex        =   10
      Top             =   0
      Width           =   1005
   End
   Begin VB.Label lblMarried 
      AutoSize        =   -1  'True
      Caption         =   "Are You Married?"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   0
      TabIndex        =   9
      Top             =   480
      Width           =   1515
   End
   Begin VB.Label lblIncome 
      AutoSize        =   -1  'True
      Caption         =   "Annual Income"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3240
      TabIndex        =   8
      Top             =   480
      Width           =   1290
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
End 'stops run of program
End Sub

Private Sub Form_Load()
frmMain.Visible = False  'makes the main form invisible
Open App.Path & "\survey.dat" For Input As #1  'opens file so the data
' can be entered

Do While Not EOF(1)  'the EOF() function tests to see
'if its the end of the file and prevents it from trying
'to read past the end of the file

Input #1, firstname, lastname, married, income, BDay_mo, BDay_day, BDay_yr
'/\ assigns the values in the file to variable so that
'they can be put in their proper controls


'\/ puts the right values in the right control
txtFirstname.Text = firstname
txtLastname.Text = lastname
optMarriedY.Value = married
Text4.Text = income
Text1.Text = BDay_mo
Text2.Text = BDay_day
Text3.Text = BDay_yr
Loop

End Sub
