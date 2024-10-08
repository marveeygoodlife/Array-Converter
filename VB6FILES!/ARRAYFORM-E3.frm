VERSION 5.00
Begin VB.Form frmArrayConverter 
   BackColor       =   &H8000000A&
   Caption         =   " GroupZ-Exercise3"
   ClientHeight    =   2865
   ClientLeft      =   120
   ClientTop       =   615
   ClientWidth     =   4560
   LinkTopic       =   "Groupz"
   ScaleHeight     =   12225
   ScaleWidth      =   22800
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   4200
      TabIndex        =   6
      ToolTipText     =   "Click to clear inputs"
      Top             =   4200
      Width           =   2100
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "Calculate"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1000
      Left            =   120
      TabIndex        =   5
      ToolTipText     =   "Click to Calculate"
      Top             =   4200
      Width           =   2100
   End
   Begin VB.TextBox txtDivisor 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3120
      TabIndex        =   1
      Text            =   "Enter Divisor"
      Top             =   3240
      Width           =   3255
   End
   Begin VB.TextBox txtArray 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3120
      TabIndex        =   0
      Text            =   "Enter Arrays"
      Top             =   1920
      Width           =   3255
   End
   Begin VB.Label lblOutput 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   120
      TabIndex        =   7
      Top             =   5520
      Width           =   6135
   End
   Begin VB.Label lblDivisor 
      Caption         =   "Enter the divisor ""K"""
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   3240
      Width           =   2775
   End
   Begin VB.Label lblArray 
      Caption         =   "Enter Arrays (, , )"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   2775
   End
   Begin VB.Label lblHeader 
      Alignment       =   2  'Center
      Caption         =   "FIND Divisible Pairs in Arrays"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   6135
   End
End
Attribute VB_Name = "frmArrayConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Label3_Click()

End Sub


Private Sub Command2_Click() 'Clear Button
    ' Clear the input fields and output label
    txtArrays.Text = ""
    txtDivisor.Text = ""
    lblOutput.Caption = ""
End Sub


Private Sub cmdCalculate_Click()
    'Calculate Button
    ' Declare the array and variables
    Dim ar() As Integer
    Dim n As Integer
    Dim k As Integer
    Dim count As Integer
    Dim i As Integer, j As Integer
    Dim inputArray As String
    Dim arrayElements() As String
    
    ' Get input from user
    inputArray = txtArray.Text ' Input array as comma-separated values
    k = Val(txtDivisor.Text)     ' Input value of k
    
    ' Split input array into elements
    arrayElements = Split(inputArray, ",")
    n = UBound(arrayElements) + 1
    ReDim ar(1 To n)
    
    ' Convert string elements to integers
    For i = 1 To n
        ar(i) = Val(arrayElements(i - 1))
    Next i
    
    ' Initialize the counter
    count = 0
    
    ' Iterate over each pair (i, j)
    For i = 1 To n - 1
        For j = i + 1 To n
            ' Check if the sum is divisible by k
            If (ar(i) + ar(j)) Mod k = 0 Then
                ' Increment the counter
                count = count + 1
            End If
        Next j
    Next i
    
    ' Display the result in the label
    lblOutput.Caption = "Divisible pairs: " & count
End Sub

Private Sub cmdClear_Click()
txtArray.Text = ""
txtDivisor.Text = ""
lblOutput.Caption = ""
End Sub

Private Sub Form_Load()

End Sub


