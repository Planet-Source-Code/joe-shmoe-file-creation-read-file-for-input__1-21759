VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Survey Example"
   ClientHeight    =   1470
   ClientLeft      =   3225
   ClientTop       =   4230
   ClientWidth     =   6120
   LinkTopic       =   "Form1"
   ScaleHeight     =   1470
   ScaleWidth      =   6120
   Begin VB.TextBox txtFirstname 
      Height          =   285
      Left            =   1080
      TabIndex        =   15
      Top             =   0
      Width           =   1335
   End
   Begin VB.ComboBox cmbMonth 
      Height          =   315
      ItemData        =   "frmMain.frx":0000
      Left            =   720
      List            =   "frmMain.frx":0028
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   960
      Width           =   975
   End
   Begin VB.ComboBox cmbIncome 
      CausesValidation=   0   'False
      Height          =   315
      ItemData        =   "frmMain.frx":008E
      Left            =   4560
      List            =   "frmMain.frx":00AD
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   480
      Width           =   1215
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
      TabIndex        =   5
      Top             =   480
      Value           =   -1  'True
      Width           =   615
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
      TabIndex        =   4
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox txtLastname 
      Height          =   285
      Left            =   3720
      TabIndex        =   0
      Top             =   0
      Width           =   1935
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
      Height          =   735
      Left            =   0
      TabIndex        =   9
      Top             =   720
      Width           =   4935
      Begin VB.ComboBox cmbYear 
         Height          =   315
         ItemData        =   "frmMain.frx":011F
         Left            =   3960
         List            =   "frmMain.frx":0121
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   240
         Width           =   855
      End
      Begin VB.ComboBox cmbDay 
         Height          =   315
         ItemData        =   "frmMain.frx":0123
         Left            =   2400
         List            =   "frmMain.frx":0125
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   240
         Width           =   855
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
         Height          =   375
         Left            =   3480
         TabIndex        =   12
         Top             =   240
         Width           =   855
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
         TabIndex        =   11
         Top             =   240
         Width           =   735
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
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
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
      TabIndex        =   7
      Top             =   480
      Width           =   1290
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
      TabIndex        =   3
      Top             =   480
      Width           =   1515
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
      TabIndex        =   2
      Top             =   0
      Width           =   1005
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
      TabIndex        =   1
      Top             =   0
      Width           =   1020
   End
   Begin VB.Menu mnuWrite 
      Caption         =   "Write to file"
   End
   Begin VB.Menu mnunull 
      Caption         =   " "
   End
   Begin VB.Menu mnuRead 
      Caption         =   "Read File"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Dim variables
Dim firstname As String
Dim lastname As String
Dim married As Boolean
Dim income As String
Dim BDay_mo As String
Dim BDay_day As Integer
Dim BDay_yr As Single
Dim i As Integer
Dim b As Single

Private Sub Form_Load()
'''''''''''''''''''''''
''''' Create File '''''
'''''''''''''''''''''''
'opens file in the directory the program in run from
Open App.Path & "\survey.dat" For Output As #1  'the app.path & "\survey" opens the file survey in the same directory of the application
'typical format for opening a file goes like this
'Open "File_Name" For FILE_MODE As FILE_NUMBER
'File mode can be input, output or append.  the append
'mode will add to an existing file if you dont't want to
'over write data like the output mode would.
'File name can be what ever you want.  you can also specify
'the file type by adding "survey.txt", "survey.dat" ect..
'File number can be any number fromn 1 to 255

''''''''''''''''''''''''''''''''''''''
''''' Put Numbers in combo boxes '''''
''''''''''''''''''''''''''''''''''''''
'Sets the number in the birth day combo box from 1 to 31
For i = 1 To 31
cmbDay.AddItem i
Next i
'Sets the birth year from 1950 to 2001 in the combo box
For b = 1950 To 2001
cmbYear.AddItem b
Next b
End Sub

Private Sub mnuRead_Click()
Form1.Show 'shows the from that reads the file

End Sub

Private Sub mnuWrite_Click()
''''''''''''''''''''''''''''''
''''' Write data to file '''''
''''''''''''''''''''''''''''''
'assigns variables to the data in the controls
firstname = txtFirstname.Text
lastname = txtLastname.Text
married = optMarriedY.Value
income = cmbIncome.Text
BDay_mo = cmbMonth.Text
BDay_day = cmbDay.Text
BDay_yr = cmbYear.Text

'writes the data to the file named "survey" as seen in the form load
Write #1, firstname, lastname, married, income, BDay_mo, BDay_day, BDay_yr
Close #1 'closes the file
End Sub

