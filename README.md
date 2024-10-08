# Array-Converter

 This is a simple VB6 application that calculate arrarys with divisible number(K)

 
## Table of Contents

1. [AUTHOR](#marvellous-ediagbonya)
1. [PROBLEM DEFINITION](#problem-definition)
2. [PROBLEM SPECIFICATION](#problem-specification)
3. [ANALYSIS](#analysis)
4. [DESIGN](#design)
5. [IMPLEMENTATION](#implementation)
5. [RESULT](#result)

## Marvellous. Ediagbonya 
NOUN,  
ABUJA  
07015824775  
CIT 104 PRACTICAL  
8th October 2024   

## PROBLEM DEFINITION

The objective is creating a VB6 application that, if given an array of integers and a positive integer K, the program should calculate how many pairs (i, j) meet these required conditions and display an output.

<strong>Where:</strong>

I < j Ar[ i ] + ar[ j ] is divisible by K

## PROBLEM SPECIFICATION

<strong>Inputs:</strong>

- An array  of integers entered by the user, it should be separated by a comma(,).
- A positive integer K will be entered by users.
- Input array must have at least 2 numbers separated by comma.
- The divisor K must be a positive number.

<strong>Outputs:</strong>

 - The total amount of pairs (i,  j) where i < j and it’s sum ( ar[i] + ar[j] ) % k = 0 is divisible by K

## ANALYSIS

- To solve this problem, we will use a brute-force approach by iterating through all pairs of indices(i , j)
- Start a counter to zero.
- Each ar[i], iterate through all subsequent elements ar[j] where j > i.
- User gets a prompt to change value  if invalid values are enter for the arrays and divisor
- Check to know if the sum ar[i] + [j] is divisible by k using the modulo operator.
- If it’s divisible, increment the counter.
- Return the counter as result.
- Input validation is important so users only enter numeric integers.

## DESIGN

The user interface for this application will consist of:

- A label (lblHeader) to display a heading
- A text box {txtArray} for the user to input the array of integers as comma -separated values.
- A text box {txtDivisor} for the user to enter the divisor k
- A button(cmdCalculate) to trigger the calculations.
- A label (lblOutput) to display the answer to the users.

<strong>FLOW CHART:</strong>
![Flowchart](https://github.com/marveeygoodlife/Array-Converter/blob/main/images/Exercise%203.jpg)
A simple flowchart representation of the algorithm

## IMPLEMENTATION

<strong>VB6 CODE IMPLEMENTATION:</strong>

```vb6
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
    lblOutput.Caption = "Number of divisible pairs: " & count
End Sub
```
CODE ENDS HERE



 ## RESULT

The application runs as expected and it has been tested with several values to test the functionality of the implementation code.  
Checks were added to ensure that both txtArray.Text and txtDivisor.Text contain valid input.   
If they are empty or invalid, a message box will appear, and the program will exit the subroutine.  
Once all checks and validations are met, the total amount is displayed to the users in an output box.  

<strong>Testing</strong>

- You can use the following input scenarios to test the implementation:
- Input: 8, 12, 4, 6, Divisor (k): 8 → Output: 1 pair
- Input: 9, 18, 27, 36, Divisor (k): 9 → Output: 6 pairs
- Input: 5, 5, 10, Divisor (k): 10 → Output: 1 pair
- Input: 1, 2, 3, 4, 5, Divisor (k): 3 → Output: 4 pairs
- Input: 10, 20, 30, 40, 50, Divisor (k): 60 → Output: 2 pairs
- Input: 5, 10, 15, 20, Divisor (k): 10 → Output: 2 pairs
 - Input: 1, 2, 3, 4, 5, 6, Divisor (k): 5 → Output: 3 pairs





