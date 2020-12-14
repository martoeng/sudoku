Attribute VB_Name = "mSudoku"
' Copyright © 2006 by Martin Walter. Alle Rechte vorbehalten. All rights reserved
' Dieser Quellcode und das Programm sind freigegeben zur nicht-kommerziellen Nutzung.
' This source and this program are free for non-commercial use.
' Dieser Copyright-Hinweis darf nicht entfernt werden!
' This copyright notice must not be removed!
' Quelle: http://www.martoeng.de
Option Explicit

Private bytSolution(0 To 8, 0 To 8) As Byte     'Sudoku-Lösungs-Matrix (solution matrix)
Public bytGame(0 To 8, 0 To 8) As Byte          'Sudoku-Spiel-Matrix (game matrix)
Private m_Difficulty As Integer                 'Schwierigkeitsgrad (difficulty level)
Private sGiven As String                        'Beim Erstellen festgelegte Zahlen (given numbers at the beginning)

'Schwierigkeitsgrad-Eigenschaft
'Difficulty level-property
Public Property Get Difficulty() As Integer
    Difficulty = m_Difficulty
End Property

Public Property Let Difficulty(ByVal New_Difficulty As Integer)
    m_Difficulty = New_Difficulty
End Property

'Setzt die gegebenen Felder
'Sets the given fields
Public Sub SetGiven(ByVal psGiven As String)
    sGiven = psGiven
End Sub

'Gibt einen Wert aus der Spiel-Matrix zurück
'Returns a value of the game matrix
Public Function GetValue(ByRef x As Integer, ByRef y As Integer) As Byte
    If x < 0 Or y < 0 Or y > 8 Or x > 8 Then
        GetValue = 0
    Else
        GetValue = bytGame(x, y)
    End If
End Function

'Setzt einen Wert der Spiel-Matrix
'Sets a value of the game matrix
Public Function SetValue(ByRef x As Integer, ByRef y As Integer, ByVal value As Byte) As Boolean
    If x >= 0 And y >= 0 And x <= 8 And y <= 8 Then
        bytGame(x, y) = value
        SetValue = True
    Else
        SetValue = False
    End If
End Function

'Ermittelt einen Wert aus der Lösung
'Returns a value of the solution matrix
Public Function GetSolutionValue(ByRef x As Integer, ByRef y As Integer) As Byte
    If x >= 0 And y >= 0 And x <= 8 And y <= 8 Then
        GetSolutionValue = bytSolution(x, y)
    Else
        GetSolutionValue = 0
    End If
End Function

'Überprüft ob alle Felder ausgefüllt wurden und ob es die richtige Lösung ist
'Checks whether all fields are filled out and verifies the solution
Public Function IsFinished() As Boolean
    Dim x As Integer
    Dim y As Integer
    For x = 0 To 8
        For y = 0 To 8
            If bytGame(x, y) = 0 Then
                IsFinished = False
                Exit Function
            End If
        Next y
    Next x

    If IsValid(bytGame) Then
        MsgBox "Herzlichen Glückwunsch. Sie haben das Sudoku gelöst.", vbInformation
        IsFinished = True
    Else
        MsgBox "Ihre Lösung weist Fehler auf!", vbExclamation
        IsFinished = False
    End If
End Function

'Löscht alle Werte aus einer Sudoku-Matrix
'clears a sudoku matrix
Sub ClearSudoku(ByRef bytSudoku() As Byte)
    Dim x As Integer, y As Integer
    For x = 0 To 8
        For y = 0 To 8
            bytSudoku(x, y) = 0
        Next y
    Next x
End Sub

'Löscht die Spiel- und Lösungsmatrix
'Clears the game and the solution matrix
Public Sub Clear()
    ClearSudoku bytGame
    ClearSudoku bytSolution
    sGiven = ""
End Sub

'Erzeugt ein neues Sudoku-Rätsel und deren Lösung
'Generates a new Sudoku (and the solution)
Public Sub Generate()
generate_start:
    'Zunächst Lösung und Spielfeld leeren
    frmSudoku.Caption = "Erzeuge neues Rätsel..."
    Clear
    
    'Alle Zahlen einzeln verteilen, angefangen bei 1 bis hoch zu 9
    'All numbers are placed after each other
    Dim Number As Integer                   'Aktuelle Zahl (current number)
    Dim rndX As Integer, rndY As Integer    'Die Koordinaten (coordinates)
    Dim i As Integer                        'Laufvariable, 1 bis 9 (man braucht je 9 mal eine Zahl)
    Dim errors As Integer                   'Anzahl Fehler bei der zufälligen Verteilung (errors during random assignment)
    Dim DeleteTries As Integer              'Anzahl Fehler bei iterierter Verteilung (errprs during iterated assignment)
    
    For Number = 1 To 9
        For i = 1 To 9
            errors = 0
generate_new_coordinates:
            'Zufällige Koordinaten bestimmen
            'create random coordinates
            rndX = Int(Rnd * 9)
            rndY = Int(Rnd * 9)
            'Prüfen, ob hier die Zahl eingetragen werden kann _
            Dazu wird die Spalte, Zeile und 3x3-Submatrix überprüft
            'check whether the number can be placed there _
            looking for row, column and 3x3 submatrix
            If CheckColumnForValue(rndX, Number, bytSolution) = False And CheckRowForValue(rndY, Number, bytSolution) = False And CheckSubMatrixForValue(rndX, rndY, Number, bytSolution) = False And bytSolution(rndX, rndY) = 0 Then
                bytSolution(rndX, rndY) = Number
            Else
                '"Fehler" erhöhen, Anzahl der Zufalls-Fehlversuche
                'increment "errors"
                errors = errors + 1
                'Bei weniger als 10 "Fehlern"
                If errors < 10 Then
                    GoTo generate_new_coordinates
                Else
                    'Bei mehr Fehlern iterativ vorgehen
                    'if there are more than 10 errors, iterate
                    For rndX = 0 To 8
                        For rndY = 0 To 8
                            If CheckColumnForValue(rndX, Number, bytSolution) = False And CheckRowForValue(rndY, Number, bytSolution) = False And CheckSubMatrixForValue(rndX, rndY, Number, bytSolution) = False And bytSolution(rndX, rndY) = 0 Then
                                'Lösung gefunden, mit nächstem i weitermachen
                                'found solution, go on with next i
                                bytSolution(rndX, rndY) = Number
                                GoTo generate_next_i
                            End If
                        Next rndY
                    Next rndX
                End If
            End If
generate_next_i:
        Next i
        
        'Vorkomnisse der Ziffer zählen, muss 9 ergeben
        'count occurences of this number, must be 9
        If CountOccurence(Number, bytSolution) <> 9 Then
            'Nicht alle konnten eingetragen werden, deshalb letzte Ziffer wieder löschen
            'und Löschvorgänge erhöhen
            'Not all could be assigned, try again with last number und increase deletetries
            DeleteTries = DeleteTries + 1
            DeleteNumber Number
            'Bei mehr als zwei fehlgeschlagenen Versuchen zur Zahl davor zurückkehren
            'If there are more than 2 tries, go back to the number before
            If DeleteTries > 2 Then
                DeleteNumber Number - 1
                Number = Number - 2
                DeleteTries = 0
            Else
                Number = Number - 1
            End If
        End If
    Next Number
    
    'Spielmatrix mit Werten aus der erstellen Lösungsmatrix auffüllen
    'fill game matrix with some values of the created solution matrix
    Dim uSolutionTries As Integer
    uSolutionTries = 0
    frmSudoku.Caption = "Erzeuge eindeutige Lösung..."
    
    ClearSudoku bytGame
    CreateGameFromSolution
    uSolutionTries = uSolutionTries + 1
    If uSolutionTries > 20 + (Difficulty * 2) Then
        GoTo generate_start
    End If
    
    frmSudoku.Caption = "Sudoku"
End Sub

Private Sub CreateGameFromSolution()
    'Anzahl der zu streichenden Kästchen
    'number of fields to be removed
    Dim uCount As Integer
    'Koordinaten
    'coordinates
    Dim rndX As Integer, rndY As Integer
    'Probierte Kästchen mit anschl. uneindeutiger Lösung
    'fields tried with subsequent not unique solution
    Dim sTried As String
create_start:
    CopySudoku bytSolution, bytGame
    sTried = ""
    uCount = 45 + (m_Difficulty * 2)
    sGiven = ""
    
    Do
create_generate_coords:
        'Fehlversuche zählen und ggf. auf iterative Suche umschwenken
        'Count the not unique tries
        If Len(sTried) < 40 Then
            rndX = Int(Rnd * 9)
            rndY = Int(Rnd * 9)
        Else
            For rndX = 0 To 8
                For rndY = 0 To 8
                    If bytGame(rndX, rndY) <> 0 And InStr(1, sTried, "(" & rndX & "," & rndY & ")") = 0 Then
                        GoTo create_check_coordinates
                    End If
                Next rndY
            Next rndX
        End If
        
create_check_coordinates:
        'Koordinaten überprüfen
        'check coordinates
        If bytGame(rndX, rndY) = 0 Or InStr(1, sTried, "(" & rndX & "," & rndY & ")") Then
            GoTo create_generate_coords
        Else
            bytGame(rndX, rndY) = 0
        End If
        
        'Testen, ob noch eindeutig lösbar
        'check for uniqueness
        If mSolve.SolveSudoku(bytGame) <> SOLVE_ONE_SOLUTION Then
            bytGame(rndX, rndY) = bytSolution(rndX, rndY)
            sTried = sTried & "(" & rndX & "," & rndY & ")"
        Else
            uCount = uCount - 1
        End If
    Loop While uCount > 0
    
    For rndX = 0 To 8
        For rndY = 0 To 8
            If bytGame(rndX, rndY) <> 0 Then
                sGiven = sGiven & "(" & rndX & "," & rndY & ")"
            End If
        Next rndY
    Next rndX
End Sub

'Gibt zurück, ob eine Zahl beim Rätselerstellen aufgedeckt wurde
'returns whether a number was revealed when the game was started
Public Function IsGiven(ByRef x As Integer, ByRef y As Integer) As Boolean
    IsGiven = (InStr(1, sGiven, "(" & CStr(x) & "," & CStr(y) & ")") > 0)
End Function

'Löscht eine Zahl aus der Lösungs-Matrix
'deletes a number from the solution matrix
Public Sub DeleteNumber(ByVal num As Integer)
    Dim x As Integer, y As Integer

    For x = 0 To 8
        For y = 0 To 8
            If bytSolution(x, y) = num Then
                bytSolution(x, y) = 0
            End If
        Next y
    Next x
End Sub

'Zählt, wie oft eine Zahl vorkommt (bei Korrektheit 9)
'counts the occurences of a number
Public Function CountOccurence(ByVal num As Integer, ByRef bytSudoku() As Byte) As Integer
    Dim x As Integer, y As Integer
    Dim uCount As Integer
    
    For x = 0 To 8
        For y = 0 To 8
            If bytSudoku(x, y) = num Then
                uCount = uCount + 1
            End If
        Next y
    Next x
    
    'Anzahl zurückgeben
    CountOccurence = uCount
End Function

'Kopiert ein Sudoku-Array in ein anderes
'copies a sudoku array from one to another
Public Sub CopySudoku(ByRef bytFrom() As Byte, ByRef bytTo() As Byte)
    Dim x As Integer, y As Integer
    For x = 0 To 8
        For y = 0 To 8
            bytTo(x, y) = bytFrom(x, y)
        Next y
    Next x
End Sub

'Übernimmt ein Sudoku-Array als neue Spielmatrix
'Applies a sudoku array as new game matrix
Public Sub ApplyGameMatrix(ByRef bytSudoku() As Byte)
    CopySudoku bytSudoku, bytGame
End Sub

'Übernimmt ein Sudoku-Array als neue Lösungsmatrix
'Applies a sudoku array as new solution matrix
Public Sub ApplySolutionMatrix(ByRef bytSudoku() As Byte)
    CopySudoku bytSudoku, bytSolution
End Sub

'Speichert ein Spiel ab
'Saves a game
Public Sub SaveGame(ByVal sFile As String)
    Dim Free As Integer
    Free = FreeFile
    Open sFile For Binary As #Free
        Put #Free, , bytSolution
        Put #Free, , bytGame
        Put #Free, , sGiven
    Close #Free
End Sub

'Lädt ein Spiel
'Loads a game
Public Sub LoadGame(ByVal sFile As String)
    Dim Free As Integer
    Free = FreeFile
    Open sFile For Binary As #Free
        Get #Free, , bytSolution
        Get #Free, , bytGame
        Get #Free, , sGiven
    Close #Free
    frmSudoku.picGame.Refresh
End Sub
