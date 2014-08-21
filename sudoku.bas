Option Explicit

''Helper functions for handling arrays.
Private Function InArray(x As Integer, A() As Integer)
    'Check if x is in one-dimensional array A.
    Dim i As Integer
    
    On Error Resume Next
    Err.Clear
    i = LBound(A)
    If Err.Number <> 0 Then
        'A is empty
        InArray = False
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    For i = LBound(A) To UBound(A)
        If x = A(i) Then
            InArray = True
            Exit Function
        End If
    Next i
    
    InArray = False
End Function

Private Sub push(x As Integer, A() As Integer)
    Dim i As Integer
    'Add x to one-dimensional array A.
    On Error Resume Next
    Err.Clear
    i = LBound(A)
    If Err.Number = 0 Then
        'A has already been allocated, just add more space.
        ReDim Preserve A(UBound(A) + 1)
        A(UBound(A)) = x
    Else
        'A needs to be initialized.
        ReDim A(0)
        A(0) = x
    End If
    On Error GoTo 0
End Sub



''Main code for solving Sudoku puzzles.
Sub SolveSudoku()
'Solve Sudoku puzzle located in cells B5-G13 of Sheet1.
'Assumes all values in this range are either integers from 1-9 or blank.
'Places any error message in cell B4 of Sheet1.

Dim grid(8, 8) As Integer, i As Integer, j As Integer

Range("B5").Activate
For i = 0 To 8
    For j = 0 To 8
        grid(i, j) = ActiveCell.Offset(i, j).Value
    Next j
Next i

If Not solveGrid(grid) Then
    Range("B4").Value = "ERROR: Unable to solve Sudoku grid."
    Range("B4").Font.ColorIndex = 3
Else
    Range("B4").Value = ""
    Range("B4").ClearFormats
    Range("B5").Activate
    For i = 0 To 8
        For j = 0 To 8
            If grid(i, j) <> 0 And ActiveCell.Offset(i, j).Value = "" Then
                With ActiveCell.Offset(i, j)
                    .Value = grid(i, j)
                    .Font.Bold = True
                    .Interior.ColorIndex = 6
                End With
            End If
        Next j
    Next i
End If

Call DebugGrid(grid)

End Sub

Private Function solveGrid(grid() As Integer)
    'Recursive function for solving Sudoku grids. Takes two-dimensional array array index 0-8 as argument.
    'Places solution in this array if possible. Returns True if grid solved, False otherwise.
    'Uses dynamic programming solution to implement solution by trying all possibilities for first cell
    'not already filled in.
    Dim i As Integer, j As Integer, k As Integer, i1 As Integer, j1 As Integer
    Dim usedRow() As Integer, usedCol() As Integer, usedSgrid() As Integer, used(9) As Boolean
    Dim encounteredZero As Boolean

    Call DebugGrid(grid)
    
    'First, check if it is possible to solve grid, or if it is already broken.
    'Check for duplicates in row.
    For i = 0 To 8
        For j = 0 To 8
            If InArray(grid(i, j), usedRow) Then
                solveGrid = False
                Exit Function
            End If
            If grid(i, j) <> 0 Then
                Call push(grid(i, j), usedRow)
            End If
        Next j
        Erase usedRow
    Next i
    'Check for duplicates in column.
    For j = 0 To 8
        For i = 0 To 8
            If InArray(grid(i, j), usedCol) Then
                solveGrid = False
                Exit Function
            End If
            If grid(i, j) <> 0 Then
                Call push(grid(i, j), usedCol)
            End If
        Next i
        Erase usedCol
    Next j
    'Check for duplicates in subgrid.
    For i = 0 To 2
        For j = 0 To 2
            For i1 = 0 To 2
                For j1 = 0 To 2
                    If InArray(grid(i * 3 + i1, j * 3 + j1), usedSgrid) Then
                        solveGrid = False
                        Exit Function
                    End If
                    If grid(i * 3 + i1, j * 3 + j1) <> 0 Then
                        Call push(grid(i * 3 + i1, j * 3 + j1), usedSgrid)
                    End If
                Next j1
            Next i1
            Erase usedSgrid
        Next j
    Next i

    'Now, look for a zero value (corresponding to blank cell), and try possibilities for filling it in.
    encounteredZero = False
    For i = 0 To 8
        For j = 0 To 8
            If grid(i, j) = 0 Then
                encounteredZero = True
                'Find possible values to fill in.
                For i1 = 1 To 9
                    used(i) = False
                Next i1
                For j1 = 0 To 8
                    If grid(i, j1) <> 0 Then used(grid(i, j1)) = True
                Next j1
                For i1 = 0 To 8
                    If grid(i1, j) <> 0 Then used(grid(i1, j)) = True
                Next i1
                For i1 = i - (i Mod 3) To i - (i Mod 3) + 2
                    For j1 = j - (j Mod 3) To j - (j Mod 3) + 2
                        If grid(i1, j1) <> 0 Then used(grid(i1, j1)) = True
                    Next j1
                Next i1
                'Try each in succession.
                For i1 = 1 To 9
                    If Not used(i1) Then
                        grid(i, j) = i1
                        If solveGrid(grid) Then
                            solveGrid = True
                            Exit Function
                        End If
                    End If
                Next i1
                'Nothing worked.
                grid(i, j) = 0
                solveGrid = False
                Exit Function
            End If
        Next j
    Next i
    
    If encounteredZero Then
        'We found a blank cell but couldn't solve.
        solveGrid = False
    Else
        'No blank cells, no problems found above--grid is already solved.
        solveGrid = True
    End If

End Function

Private Sub DebugGrid(grid() As Integer)
'Display Sudoku grid for debugging purposes. Assumes grid is 2-dimensional array indexed 0-8.
Dim i As Integer

For i = 0 To 8
    Debug.Print (grid(i, 0) & grid(i, 1) & grid(i, 2) & grid(i, 3) & grid(i, 4) & grid(i, 5) _
        & grid(i, 6) & grid(i, 7) & grid(i, 8))
Next i

End Sub
