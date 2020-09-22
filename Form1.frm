VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1800
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6660
   LinkTopic       =   "Form1"
   ScaleHeight     =   1800
   ScaleWidth      =   6660
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   735
   End
   Begin MSFlexGridLib.MSFlexGrid wordgrid 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   661
      _Version        =   393216
      Rows            =   1
      Cols            =   19
      FixedRows       =   0
      FixedCols       =   0
      BackColor       =   0
      ForeColor       =   65280
      ForeColorFixed  =   0
      GridColor       =   0
      ScrollBars      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label en 
      Caption         =   "Enter a 16 bit number 0-32,767"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   2
      Top             =   240
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

'setup grid
Dim i As Integer
For i = 0 To 18 Step 1
wordgrid.ColWidth(i) = 300
Next
wordgrid.Width = 300 * 19 + 100
wordgrid.Height = wordgrid.RowHeight(0) + 100

 

End Sub

Private Sub Text1_Change()
If Text1.Text = "" Then Exit Sub

Dim intInput As Integer
intInput = CInt(Text1.Text)
Call SetNibbles(intInput)
End Sub
