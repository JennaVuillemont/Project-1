Attribute VB_Name = "matrixminus"
Option Explicit

Function minusmatrix(matrix1 As Variant, matrix2 As Variant)

Dim n As Double, m As Double, j As Double, k As Double
Dim num1 As Double, num2 As Double
Dim nrows1 As Double, ncols1 As Double, nrows2 As Double, ncols2 As Double, nrows As Double, ncols As Double
Dim ir As Double, ic As Double
Dim total As Double, matrix3 As Variant


matrix1(1 To nrows, 1 To ncols)
matrix2(1 To nrows, 1 To ncols)
matrix3(1 To nrows, 1 To ncols)
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


Function inversionmatrix(matrix4 As Variant, matrix5 As Variant)

' use inverse operation to invert one matrix in order to multiply
Dim n As Double, m As Double, j As Double, k As Double
Dim num1 As Double, num2 As Double, num3 As Double
Dim nrows1 As Double, ncols1 As Double, nrows2 As Double, ncols2 As Double, nrows As Double, ncols As Double
Dim ir As Double, ic As Double
Dim total As Double, total2 As Double, matrix3 As Variant



matrix1(1 To nrows1, 1 To ncols1)
matrix2(1 To nrows2, 1 To ncols2)
matrix3(1 To nrows1, 1 To ncols2)
'can make matrix2 nxm because they should equal one another


nrows1 = UBound(matrix1, 1) - LBound(matrix1, 1) + 1
ncols1 = UBound(matrix1, 2) - LBound(matrox1, 2) + 1
nrows2 = UBound(matrix2, 1) - LBound(matrix2, 1) + 1
ncols2 = UBound(matrix2, 2) - LBound(matrix2, 2) + 1

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
    matrix2 = Application.MInverse(matrix2)
End If

' switching column on one and row on another

For ic = 1 To ncols Step 1
    total = 0
    matrix1(ir, ic) = num1
        For ir = 1 To nrows Step 1
                matrix2(ir, ic) = num2
                total = num1 * num2
        Next ir
        matrix3(ir, ic) = total
Next ic

For ir = 1 To nrows Step 1
    total = 0
    matrix2(ir, ic) = num1
        For ic = 1 To ncols Step 1
            matrix1(ir, ic) = num2
            num3 = num1 * num2
            matrix2(ir + 1, ic) = num1
            total2 = total + num3
        Next ic
    matrix(ir, ic) = total2
Next ir




End Function


Function redooperations(matrix2 As Variant)

Dim diff As Variant

ReDim diff(1 To n, 1 To m)

Dim ir As Double, ic As Double

For ir = 1 To n
    For ic = 1 To j
        diff(ir, ic) = matrix1(ir, ic) - matrix2(ir, ic)
    Next ic
Next ir

'deleted this from the first function

For ic = 1 To m Step 1
    total = 0
    matrix1(ir, ic) = num1
        For ir = 1 To n Step 1
                matrix2(ir, ic) = num2
                total = num1 - num2
                matrix3(ir, ic) = total
        Next ir
    Next ic
    

For ic1 = 1 To m Step 1
    total = 0
    For ic2 = 1 To k Step 1
        For ir1 = 1 To n Step 1
            matrix1(ir1, ic1) = num1
            For ir2 = 1 To j Step 1
                matrix(ir2, ic2) = num2
                total = ir1 - ir2
                matrix3(ir2, ic2) = total
            Next ir2
        Next ir1
    Next ic2
Next ic1
End Function



