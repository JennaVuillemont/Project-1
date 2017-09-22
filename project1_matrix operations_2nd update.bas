Attribute VB_Name = "matrixminus"
Option Explicit

Function minusmatrix(matrix1() As Double, matrix2() As Double) As Variant

Dim n As Double, m As Double, j As Double, k As Double
Dim num1 As Double, num2 As Double
Dim nrows1 As Double, ncols1 As Double, nrows2 As Double, ncols2 As Double, nrows As Double, ncols As Double
Dim ir As Double, ic As Double
Dim total As Double, matrix3() As Variant


ReDim matrix1(1 To nrows, 1 To ncols)
ReDim matrix2(1 To nrows, 1 To ncols)
ReDim matrix3(1 To nrows, 1 To ncols)
'can make matrix2 nxm because they should equal one another


nrows1 = UBound(matrix1, 1) - LBound(matrix1, 1) + 1
ncols1 = UBound(matrix1, 2) - LBound(matrox1, 2) + 1
nrows2 = UBound(matrix2, 1) - LBound(matrix2, 1) + 1
ncols2 = UBound(matrix2, 2) - LBound(matrix2, 2) + 1

n = nrows1
m = ncols1
j = nrows2
k = ncols2

If n <> j Or m <> k Then
    MsgBox ("matrices not the same size")
    Exit Function
Else
    ncols = ncols1
    nrows = nrows1
End If
' this determines if the matrices are the same size

For ic = 1 To ncols Step 1
    total = 0
    matrix1(ir, ic) = num1
        For ir = 1 To nrows Step 1
                matrix2(ir, ic) = num2
                total = num1 - num2
                matrix3(ir, ic) = total
        Next ir
Next ic

' is this doing what I want? no, come back to this (might have an answer if you only do it one way

End Function

Function MatrixAdd(matrix1() As Double, matrix2() As Double) As Variant

    Dim nrows1 As Integer, ncols1 As Integer, nrows2 As Integer, ncols2 As Integer
    Dim i As Integer, j As Integer, k As Integer

    nrows1 = UBound(matrix1, 1) - LBound(matrix1, 1) + 1
    nrows2 = UBound(matrix2, 1) - LBound(matrix2, 1) + 1
    ncols1 = UBound(matrix1, 2) - LBound(matrix1, 2) + 1
    ncols2 = UBound(matrix2, 2) - LBound(matrix2, 2) + 1

    

    If nrows1 = nrows2 And ncols1 = ncols2 Then
        Dim matrix3() As Double
        ReDim matrix3(1 To nrowsa, 1 To ncolsa)

    
        For i = 1 To nrowsa Step 1

            For j = 1 To ncolsa Step 1

                matrix3(i, j) = matrix1(i, j) + matrix2(i, j)

     '           Range("b1").Cells(i, j).Value = matc(i, j)
     
            Next j

        Next i

        MatrixAdd = matc

    

    k = 2
    Do While Range("b1").Cells(k, 1).Value <> ""
        k = k + 1
    Loop

    Range("b1").Cells(k - 1, 1).Value = "Matrix Addition"

    For i = 1 To nrowsa Step 1
        For j = 1 To ncolsb Step 1
            Range("b1").Cells(k + i, j).Value = matrix3(i, j)
        Next j
    Next i

    Else
        MsgBox ("The matrices aren't the correct dimensions for addition")

        Exit Function

    End If
            

End Function


Function inversionmatrix(matrix4 As Variant, matrix5 As Variant)

' use inverse operation to invert one matrix in order to multiply
Dim n As Double, m As Double, j As Double, k As Double
Dim num1 As Double, num2 As Double, num3 As Double
Dim nrows1 As Double, ncols1 As Double, nrows2 As Double, ncols2 As Double, nrows As Double, ncols As Double
Dim ir As Double, ic As Double, ic2 As Double
Dim total As Double, total2 As Double, matrix6 As Variant



matrix1(1 To nrows1, 1 To ncols1)
matrix2(1 To nrows2, 1 To ncols2)
matrix3(1 To nrows1, 1 To ncols2)
'can make matrix2 nxm because they should equal one another


nrows1 = UBound(matrix4, 1) - LBound(matrix4, 1) + 1
ncols1 = UBound(matrix4, 2) - LBound(matrox4, 2) + 1
nrows2 = UBound(matrix5, 1) - LBound(matrix5, 1) + 1
ncols2 = UBound(matrix5, 2) - LBound(matrix5, 2) + 1

n = nrows1
m = ncols1
j = nrows2
k = ncols2

' can not technically divide matrices, so take the inverse and then multiply
' rule for multiplication means the first and last must equal one another

If m <> j Or n <> k Then
    MsgBox ("matrices will not work")
    Exit Function
Else
    matrix5 = Application.MInverse(matrix2)
End If

' switching column on one and row on another



For ir = 1 To nrows Step 1
    total = 0
    matrix5(ir, ic) = num1
        For ic = 1 To ncols Step 1
            matrix4(ir, ic) = num2
            num3 = num1 * num2
            matrix5(ir + 1, ic) = num1
            total2 = total + num3
        Next ic
    matrix6(ir, ic) = total2
Next ir

For ic2 = 1 To ncols2 Step 1
            For ir = 1 To nrows1 Step 1
                For ic = 1 To ncols1 Step 1
                    matrix6(ir, ic2) = matrix6(ir, ic2) + (matrix4(ir, ic) * matrix5(ic, ic2))
                Next ic
            Next ir
Next ic2

        MatrixMult = matrix6

    ic2 = 2

    Do While Range("b1").Cells(ic2, 1).Value <> ""
        ic2 = ic2 + 1
    Loop

    

    Range("b1").Cells(ic2 + i - 1, 1).Value = "Matrix Multiplication"

    
    For ir = 1 To nrows1 Step 1
        For ic = 1 To ncols2 Step 1
            Range("b1").Cells(ic2 + ir, ic).Value = matric6(ir, ic)
        Next ic
    Next ir



End Function



