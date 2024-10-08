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
      Left            =   8520
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
      Left            =   4440
      TabIndex        =   5
      ToolTipText     =   "Click to Calculate"
      Top             =   4200
      Width           =   2100
   End
   Begin VB.TextBox txtDivisor 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8040
      TabIndex        =   1
      Text            =   "Enter Divisor"
      Top             =   3000
      Width           =   2535
   End
   Begin VB.TextBox txtArray 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8040
      TabIndex        =   0
      Text            =   "Enter Arrays"
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Label lblOutput 
      Height          =   705
      Left            =   4440
      TabIndex        =   7
      Top             =   5640
      Width           =   6135
   End
   Begin VB.Label lblDivisor 
      Caption         =   "Enter the divisor ""K"""
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4440
      TabIndex        =   4
      Top             =   3000
      Width           =   2775
   End
   Begin VB.Label lblArray 
      Caption         =   "Enter Arrays (, , )"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4440
      TabIndex        =   3
      Top             =   1800
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
      Left            =   4440
      TabIndex        =   2
      Top             =   240
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


Private Sub cmdCalculate_Click()
    ' Declare the array and variables
    Dim ar() As Integer
    Dim n As Integer
    Dim k As Integer
    Dim count As Integer
    Dim i As Integer, j As Integer
    Dim inputArray As String
    Dim arrayElements() As String
    Dim sum As Integer  ' Declaring the 'sum' variable
    
    ' Get input from user
    inputArray = Trim(txtArray.Text) ' Input array as comma-separated values
    k = Val(Trim(txtDivisor.Text)) ' Input value of k
    
    ' Check for valid input
    If inputArray = "" Or k = 0 Then
        MsgBox "Please enter a valid array and divisor."
        Exit Sub
    End If
    
    ' Split input array into elements
    arrayElements = Split(inputArray, ",")
    
    ' Ensure input array has at least two numbers
    If UBound(arrayElements) < 1 Then
        MsgBox "Please enter at least two numbers."
        Exit Sub
    End If
    
    ' Set the length of the array
    n = UBound(arrayElements) + 1
    ReDim ar(1 To n)
    
    ' Convert string elements to integers and validate
    For i = 1 To n
        If IsNumeric(Trim(arrayElements(i - 1))) Then
            ar(i) = Val(Trim(arrayElements(i - 1)))
        Else
            MsgBox "Please enter valid numeric values in the array."
            Exit Sub
        End If
    Next i
    
    ' Initialize the counter
    count = 0
    
    ' Iterate over each pair (i, j)
    For i = 1 To n - 1
        For j = i + 1 To n
            ' Calculate the sum of the pair
            sum = ar(i) + ar(j)  ' Now the 'sum' variable is properly declared
            
            ' Debugging output for checking sums and modulo results
            Debug.Print "Pair (" & ar(i) & ", " & ar(j) & ") => Sum: " & sum & " Mod " & k & " = " & (sum Mod k)
            
            ' Check if the sum is divisible by k
            If sum Mod k = 0 Then
                ' Increment the counter
                count = count + 1
                Debug.Print "Valid pair: (" & ar(i) & ", " & ar(j) & ")"
            End If
        Next j
    Next i
    
    ' Display the result in the label
    lblOutput.Caption = "Number of divisible pairs: " & count
End Sub

Private Sub Command2_Click() 'Clear Button
    ' Clear the input fields and output label
    txtArrays.Text = ""
    txtDivisor.Text = ""
    lblOutput.Caption = ""
End Sub



Private Sub cmdClear_Click()
txtArray.Text = ""
txtDivisor.Text = ""
lblOutput.Caption = ""
End Sub

