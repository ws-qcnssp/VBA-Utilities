Attribute VB_Name = "SetUtils"
Function SUMOFMULT2(rng1 As Range, rng2 As Range) As Double

Dim col1, col2, row1, row2, col_i, row_i As Integer

col1 = rng1.Columns.Count
col2 = rng2.Columns.Count
row1 = rng1.Rows.Count
row2 = rng2.Rows.Count

If col1 = col2 And row1 = row2 Then
    SUMOFMULT2 = 0
    For col_i = 1 To col1
        For row_i = 1 To row1
            If isempty(rng1.Cells(row_i, col_i).Value) = False And isempty(rng2.Cells(row_i, col_i).Value) = False And isNumeric(rng1.Cells(row_i, col_i).Value) And isNumeric(rng2.Cells(row_i, col_i).Value) Then
                SUMOFMULT2 = SUMOFMULT2 + rng1.Cells(row_i, col_i).Value * rng2.Cells(row_i, col_i).Value
            End If
        Next row_i
    Next col_i
Else
    SUMOFMULT2 = CVErr(xlErrRef)
End If

End Function
Function SUMOFMULT(ParamArray args() As Variant) As Double

Dim col1, row1, col_i, row_i, countFactors As Integer
Dim cellMult As Double

Dim i As Long

' check if all args are Range
For i = LBound(args) To UBound(args)
    If Not TypeName(args(i)) = "Range" Then
        SUMOFMULT = CVErr(xlErrRef)
    End If
Next i

col1 = args(LBound(args)).Columns.Count
row1 = args(LBound(args)).Rows.Count

' check if all ranges are identical size
For i = LBound(args) + 1 To UBound(args)
    If args(i).Columns.Count = col1 And args(i).Rows.Count = row1 Then
        SUMOFMULT = 0
    Else
        SUMOFMULT = CVErr(xlErrRef)
    End If
Next i

' calculate sum of multiplications
For col_i = 1 To col1
    For row_i = 1 To row1
        cellMult = 1
        countFactors = 0
        For i = LBound(args) To UBound(args)
            If isempty(args(i).Cells(row_i, col_i).Value) = False And isNumeric(args(i).Cells(row_i, col_i).Value) Then
                countFactors = countFators + 1
                cellMult = cellMult * args(i).Cells(row_i, col_i).Value
            End If
        Next i
        If countFactors > 0 Then
            SUMOFMULT = SUMOFMULT + cellMult
        End If
    Next row_i
Next col_i

End Function

Function SUMOFDIV2(rng1 As Range, rng2 As Range) As Double

Dim col1, col2, row1, row2, col_i, row_i As Integer

col1 = rng1.Columns.Count
col2 = rng2.Columns.Count
row1 = rng1.Rows.Count
row2 = rng2.Rows.Count

If col1 = col2 And row1 = row2 Then
    SUMOFDIV2 = 0
    For col_i = 1 To col1
        For row_i = 1 To row1
            If isempty(rng1.Cells(row_i, col_i).Value) = False And isempty(rng2.Cells(row_i, col_i).Value) = False And isNumeric(rng1.Cells(row_i, col_i).Value) And isNumeric(rng2.Cells(row_i, col_i).Value) Then
                SUMOFDIV2 = SUMOFDIV2 + rng1.Cells(row_i, col_i).Value / rng2.Cells(row_i, col_i).Value
            End If
        Next row_i
    Next col_i
Else
    SUMOFDIV2 = CVErr(xlErrRef)
End If

End Function

Function SUMOFMARKED(rng1 As Range, rng2 As Range, marker As Variant) As Double

Dim col1, col2, row1, row2, col_i, row_i As Integer

col1 = rng1.Columns.Count
col2 = rng2.Columns.Count
row1 = rng1.Rows.Count
row2 = rng2.Rows.Count

If col1 = col2 And row1 = row2 Then
    SUMOFMARKED = 0
    For col_i = 1 To col1
        For row_i = 1 To row1
            If isempty(rng1.Cells(row_i, col_i).Value) = False And isempty(rng2.Cells(row_i, col_i).Value) = False And isNumeric(rng1.Cells(row_i, col_i).Value) Then
                If rng2.Cells(row_i, col_i).Value = marker Then
                    SUMOFMARKED = SUMOFMARKED + rng1.Cells(row_i, col_i).Value
                End If
            End If
        Next row_i
    Next col_i
Else
    SUMOFMARKED = CVErr(xlErrRef)
End If

End Function

Function AVERAGEOFMARKED(rng1 As Range, rng2 As Range, marker As Variant) As Double

Dim col1, col2, row1, row2, col_i, row_i, counter As Integer

col1 = rng1.Columns.Count
col2 = rng2.Columns.Count
row1 = rng1.Rows.Count
row2 = rng2.Rows.Count

If col1 = col2 And row1 = row2 Then
    AVERAGEOFMARKED = 0
    counter = 0
    For col_i = 1 To col1
        For row_i = 1 To row1
            If isempty(rng1.Cells(row_i, col_i).Value) = False And isempty(rng2.Cells(row_i, col_i).Value) = False And isNumeric(rng1.Cells(row_i, col_i).Value) Then
                If rng2.Cells(row_i, col_i).Value = marker Then
                    counter = counter + 1
                    AVERAGEOFMARKED = AVERAGEOFMARKED + rng1.Cells(row_i, col_i).Value
                End If
            End If
        Next row_i
    Next col_i
Else
    AVERAGEOFMARKED = CVErr(xlErrRef)
End If

If counter > 0 Then
    AVERAGEOFMARKED = AVERAGEOFMARKED / counter
End If

End Function

