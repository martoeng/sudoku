Attribute VB_Name = "mSolve"
' Copyright © 2006 by Martin Walter. Alle Rechte vorbehalten. All rights reserved
' Dieser Quellcode und das Programm sind freigegeben zur nicht-kommerziellen Nutzung.
' This source and this program are free for non-commercial use.
' Dieser Copyright-Hinweis darf nicht entfernt werden!
' This copyright notice must not be removed!
' Quelle: http://www.martoeng.de
Option Explicit
Public Enum SOLVE_CONSTANTS
    SOLVE_NO_SOLUTION = 0
    SOLVE_ONE_SOLUTION = 1
    SOLVE_MULTIPLE_SOLUTIONS = 2
End Enum
Private uSolutions As Integer                       'gefundene Lösungen (found solutions)
Private bytLastSolution(0 To 8, 0 To 8) As Byte     'letzte gefundene Lösung (last found solution)

'Kopiert die letzte Lösung in ein anderes Sudoku
'Copies the last solution to another sudoku matrix
Public Sub CopySolution(ByRef bytTo() As Byte)
    mSudoku.CopySudoku bytLastSolution, bytTo
End Sub

'Löst ein Sudoku
'Solves a sudoku
Public Function SolveSudoku(ByRef bytSudoku() As Byte) As SOLVE_CONSTANTS
    uSolutions = 0
    ClearSudoku bytLastSolution
    
    Dim bytNewSudoku(0 To 8, 0 To 8) As Byte
    CopySudoku bytSudoku, bytNewSudoku
    
    'Testen, ob es besser ist bei 0,0 oder 8,8 anzufangen
    'test whether it's better to start with 0,0 or 8,8
    Dim x As Integer, y As Integer
    x = IIf(bytSudoku(0, 0) = 0, 1, 0)
    x = x + IIf(bytSudoku(0, 1) = 0, 1, 0)
    x = x + IIf(bytSudoku(0, 2) = 0, 1, 0)
    y = IIf(bytSudoku(8, 8) = 0, 1, 0)
    y = y + IIf(bytSudoku(8, 7) = 0, 1, 0)
    y = y + IIf(bytSudoku(8, 6) = 0, 1, 0)
    Solve bytNewSudoku, (x <= y)
    
    Select Case uSolutions
        Case 0
            SolveSudoku = SOLVE_NO_SOLUTION
        Case 1
            SolveSudoku = SOLVE_ONE_SOLUTION
        Case Else
            SolveSudoku = SOLVE_MULTIPLE_SOLUTIONS
    End Select
End Function

Private Function Solve(ByRef bytNewSudoku() As Byte, ByRef bFrom0To8 As Boolean)
    'Erstes freies Feld suchen
    'Find first free field
    Dim x As Integer, y As Integer
    If bFrom0To8 Then
        For x = 0 To 8
            For y = 0 To 8
                If bytNewSudoku(x, y) = 0 Then GoTo solve_found_free
            Next y
        Next x
    Else
        For x = 8 To 0 Step -1
            For y = 8 To 0 Step -1
                If bytNewSudoku(x, y) = 0 Then GoTo solve_found_free
            Next y
        Next x
    End If
    
    'Wenn kein freies Feld mehr gefunden wurde abbrechen und Lösung checken
    'If there's no free field: abort and check the solution
    If x = 9 Or y = 9 Or x = -1 Or y = -1 Then
        If mCheck.IsValid(bytNewSudoku) Then
            CopySudoku bytNewSudoku, bytLastSolution
            uSolutions = uSolutions + 1
        End If
        Exit Function
    End If
    
solve_found_free:
    
    'Freies Feld gefunden, mögliche Zahlen durchgehen
    'found a free field, test all possible numbers
    Dim Number As Byte
    Dim bPossible(1 To 9) As Byte
    Call mCheck.GetPossibleValuesArray(x, y, bytNewSudoku, bPossible)
    
    'Mögliche Zahlen durchgehen und Prozedur neu aufrufen
    'Test all possible numbers and make recursive call
    For Number = 1 To 9
        If bPossible(Number) = 0 Then
            bytNewSudoku(x, y) = Number
            Solve bytNewSudoku, bFrom0To8
            If uSolutions = 2 Then Exit Function
        End If
    Next Number
    
    'Wieder zurücksetzen auf 0
    bytNewSudoku(x, y) = 0
End Function
