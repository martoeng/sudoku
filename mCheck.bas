Attribute VB_Name = "mCheck"
' Copyright © 2006 by Martin Walter. Alle Rechte vorbehalten. All rights reserved
' Dieser Quellcode und das Programm sind freigegeben zur nicht-kommerziellen Nutzung.
' This source and this program are free for non-commercial use.
' Dieser Copyright-Hinweis darf nicht entfernt werden!
' This copyright notice must not be removed!
' Quelle: http://www.martoeng.de
Option Explicit

'Prüft, ob ein Wert in einer Spalte bereits vorhanden ist
'Checks a column for a value
Public Function CheckColumnForValue(ByVal col As Integer, ByVal val As Integer, ByRef bytSudoku() As Byte) As Boolean
    Dim i As Integer
    For i = 0 To 8
        If bytSudoku(col, i) = val Then
            'Vorhanden
            CheckColumnForValue = True
            Exit Function
        End If
    Next i
    
    'Die Zahl konnte nicht gefunden werden
    'the number could not be found
    CheckColumnForValue = False
End Function

'Prüft, ob ein Wert in einer Zeile bereits vorhanden ist
'checks a row for a value
Public Function CheckRowForValue(ByVal Row As Integer, ByVal val As Integer, ByRef bytSudoku() As Byte) As Boolean
    Dim i As Integer
    For i = 0 To 8
        If bytSudoku(i, Row) = val Then
            'Vorhanden
            CheckRowForValue = True
            Exit Function
        End If
    Next i
    
    'Der Werte konnte nicht gefunden werden
    'the number could not be found
    CheckRowForValue = False
End Function

'Prüft, ob ein Wert bereits in einer 3x3-Untermatrix vorhanden ist
'checks a 3x3 submatrix for a value
Public Function CheckSubMatrixForValue(ByVal x As Integer, ByVal y As Integer, ByVal value As Integer, ByRef bytSudoku() As Byte) As Boolean
    Dim mX As Integer, mY As Integer
    'Linkes oberes Kästchen der Submatrix bestimmen
    mX = Int(x / 3) * 3
    mY = Int(y / 3) * 3
    
    For x = mX To mX + 2
        For y = mY To mY + 2
            If bytSudoku(x, y) = value Then
                'Vorhanden
                CheckSubMatrixForValue = True
                Exit Function
            End If
        Next y
    Next x
    
    'Nicht vorhanden
    CheckSubMatrixForValue = False
End Function

'Prüft, ob ein Sudoku gültig ist
'Checks whether a sudoku matrix is valid
Public Function IsValid(ByRef bytSudoku() As Byte) As Boolean
    Dim Number As Byte
    Dim x As Byte, i As Byte
    Dim bRow As Byte, bCol As Byte
    
    For Number = 1 To 9
        bCol = 0: bRow = 0
        For x = 0 To 8
            For i = 0 To 8
                If bytSudoku(i, x) = Number Then
                    'Vorhanden
                    bCol = 1
                End If
                If bytSudoku(x, i) = Number Then
                    'Vorhanden
                    bRow = 1
                End If
            Next i
        Next x
        If bRow = 0 Or bCol = 0 Then Exit Function
    Next Number
    
    IsValid = True
End Function

'Gibt die möglichen Zahlen für ein Feld zurück
'returns all possible values for a field
Public Function GetPossibleValues(ByRef x As Integer, ByRef y As Integer, ByRef bytSudoku() As Byte) As String

    Dim bPossible(1 To 9) As Byte
    Dim i As Integer
    
    For i = 0 To 8
        'Zeile und Spalte nach Werten durchsuchen
        If bytSudoku(i, y) <> 0 Then bPossible(bytSudoku(i, y)) = 1
        If bytSudoku(x, i) <> 0 Then bPossible(bytSudoku(x, i)) = 1
    Next i
      
    Dim mX As Integer, mY As Integer
    mX = Int(x / 3) * 3
    mY = Int(y / 3) * 3
      
    Dim j As Integer
    For i = mX To mX + 2
        For j = mY To mY + 2
            If bytSudoku(i, j) <> 0 Then bPossible(bytSudoku(i, j)) = 1
        Next j
    Next i
        
    Dim sPossible As String
    
    For i = 1 To 9
        If bPossible(i) = 0 Then sPossible = sPossible & CStr(i)
    Next i
    
    GetPossibleValues = sPossible
End Function

Public Function GetPossibleValuesArray(ByRef x As Integer, ByRef y As Integer, ByRef bytSudoku() As Byte, ByRef bPossible() As Byte)
    Dim i As Integer
    'Alle auf 0 setzen (zahl kommt nicht vor)
    'Set all to 0 (number doesn't appear)
    For i = 1 To 9
        bPossible(i) = 0
    Next i
    
    'Zeile und Spalte nach Werten durchsuchen
    'Check Row and Column for values
    For i = 0 To 8
        If bytSudoku(i, y) <> 0 Then bPossible(bytSudoku(i, y)) = 1
        If bytSudoku(x, i) <> 0 Then bPossible(bytSudoku(x, i)) = 1
    Next i
      
    'Submatrix durchsuchen
    'search submatrix
    Dim mX As Integer, mY As Integer
    mX = Int(x / 3) * 3
    mY = Int(y / 3) * 3
    Dim j As Integer
    For i = mX To mX + 2
        For j = mY To mY + 2
            If bytSudoku(i, j) <> 0 Then bPossible(bytSudoku(i, j)) = 1
        Next j
    Next i
End Function
