VERSION 5.00
Begin VB.Form frmSolve 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Sudoku lösen"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6735
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSolve.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   281
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   449
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton cmdApply 
      Caption         =   "Übernehmen..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   3
      ToolTipText     =   "Übernimmt Ihre Eingaben als neues Rätsel"
      Top             =   720
      Width           =   2295
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "S&chließen"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   2
      Top             =   1320
      Width           =   2295
   End
   Begin VB.CommandButton cmdSolve 
      Caption         =   "Lösen"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   1
      Top             =   120
      Width           =   2295
   End
   Begin VB.TextBox txtNumber 
      Alignment       =   2  'Zentriert
      Appearance      =   0  '2D
      Height          =   360
      Index           =   0
      Left            =   120
      MaxLength       =   1
      TabIndex        =   0
      Top             =   120
      Width           =   360
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   8
      X2              =   271
      Y1              =   185
      Y2              =   185
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   8
      X2              =   271
      Y1              =   95
      Y2              =   95
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   185
      X2              =   185
      Y1              =   8
      Y2              =   271
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      X1              =   95
      X2              =   95
      Y1              =   8
      Y2              =   271
   End
End
Attribute VB_Name = "frmSolve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright © 2006 by Martin Walter. Alle Rechte vorbehalten. All rights reserved
' Dieser Quellcode und das Programm sind freigegeben zur nicht-kommerziellen Nutzung.
' This source and this program are free for non-commercial use.
' Dieser Copyright-Hinweis darf nicht entfernt werden!
' This copyright notice must not be removed!
' Quelle: http://www.martoeng.de
Option Explicit

Private Sub cmdApply_Click()
    Dim bytSolve(0 To 8, 0 To 8) As Byte
    Dim uCount As Integer
    Dim sGiven As String
    
    uCount = ConvertBoxesToArray(bytSolve, sGiven)
    
    If uCount = -1 Then
        MsgBox "Es ist ein Fehler aufgetreten. Es wurde ein irreguläres Zeichen gefunden.", vbExclamation, "Fehler"
        Exit Sub
    End If
    
    If uCount < 17 Then
        'Unter 17 Feldern nie lösbar
        'At least 17 fields required for a solution
        MsgBox "Bitte geben Sie mindestens 17 Zahlen ein. Darunter sind Sudokus nie eindeutig.", vbExclamation
        Exit Sub
    End If
    
    mSudoku.ApplyGameMatrix bytSolve
    If mSolve.SolveSudoku(bytSolve) <> SOLVE_ONE_SOLUTION Then
        MsgBox "Das eingegebene Sudoku besitzt nicht genau eine Lösung!", vbExclamation, "Fehler"
        Exit Sub
    End If
    mSolve.CopySolution bytSolve
    mSudoku.ApplySolutionMatrix bytSolve
    mSudoku.SetGiven sGiven
    
    'Aus der Anzahl der gegebenen Felder die Schwierigkeit berechnen
    'Calculate the difficulty from the given fields
    Select Case uCount
        Case Is >= 36
            frmSudoku.optDifficulty(0).value = True
        Case Is >= 34
            frmSudoku.optDifficulty(1).value = True
        Case Is >= 32
            frmSudoku.optDifficulty(2).value = True
        Case Is >= 30
            frmSudoku.optDifficulty(3).value = True
        Case Is >= 28
            frmSudoku.optDifficulty(4).value = True
        Case Is < 28
            frmSudoku.optDifficulty(5).value = True
    End Select
    
    Unload Me
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSolve_Click()
    Dim bytSolve(0 To 8, 0 To 8) As Byte
    Dim uCount As Integer
    
    'Textfeld-Array umwandeln in Byte-Array
    'Convert the text fields' contents into a byte array
    uCount = ConvertBoxesToArray(bytSolve)
    
    If uCount = -1 Then
        MsgBox "Es ist ein Fehler aufgetreten. Es wurde ein irreguläres Zeichen gefunden.", vbExclamation, "Fehler"
        Exit Sub
    End If
    
    
    If uCount < 17 Then
        'Unter 17 Feldern nie lösbar
        'At least 17 fields required for a solution
        MsgBox "Bitte geben Sie mindestens 17 Zahlen ein. Darunter sind Sudokus nie eindeutig.", vbExclamation
        Exit Sub
    End If
    
    'Lösen
    'Solve
    Select Case mSolve.SolveSudoku(bytSolve)
        Case 0
            'Keine Lösung
            'No Solution
            MsgBox "Dieses Sudoku besitzt keine Lösung.", vbExclamation
        Case 1
            'Genau eine Lösung, diese anzeigen
            'Exactly one solution, show it
            mSolve.CopySolution bytSolve
            Dim i As Integer, Column As Integer, Row As Integer
            For i = 0 To 80
                Column = i Mod 9
                Row = Int(i / 9)
                txtNumber(i).Text = CStr(bytSolve(Column, Row))
            Next i
            MsgBox "Dieses Sudoku besitzt genau eine Lösung.", vbInformation
        Case 2
            'Mindestens zwei Lösungen, -> nicht eindeutig
            'At least 2 solutions, -> not unique
            MsgBox "Dieses Sudoku ist nicht eindeutig lösbar.", vbExclamation
    End Select
End Sub

Private Sub Form_Load()
    'Alle 79 fehlenden Textfelder laden und anzeigen
    'Load and show all 79 missing textfields
    Dim i As Integer
    Dim Column As Integer, Row As Integer
    For i = 1 To 80
        Load txtNumber(i)
        Column = i Mod 9
        Row = Int(i / 9)
        txtNumber(i).Left = Column * 30 + txtNumber(0).Left
        txtNumber(i).TOp = Row * 30 + txtNumber(0).TOp
        txtNumber(i).Visible = True
        txtNumber(i).TabIndex = i
    Next i
End Sub

Private Sub Form_Terminate()
    Set frmSolve = Nothing
End Sub

'Wandelt die Textboxen in ein Byte-Array um und gibt die Anzahl der ausgefüllten Felder zurück
'Converts the textboxes into a byte array and returns the number of filled out fields
'Gibt -1 zurück, falls ein Fehler auftritt
'Return -1 if an error occurs
Private Function ConvertBoxesToArray(ByRef bytArray() As Byte, Optional ByRef psGiven As String) As Integer
    Dim i As Integer
    Dim Column As Integer
    Dim Row As Integer
    Dim uCount As Integer
    
    For i = 0 To 80
        Column = i Mod 9
        Row = Int(i / 9)
        If IsNumeric(txtNumber(i).Text) Or txtNumber(i).Text = "" Then
            If txtNumber(i).Text = "" Then
                bytArray(Column, Row) = 0
            Else
                bytArray(Column, Row) = CByte(txtNumber(i).Text)
                uCount = uCount + 1
                If Not IsMissing(psGiven) Then psGiven = psGiven & "(" & Column & "," & Row & ")"
            End If
        Else
            txtNumber(i).SetFocus
            ConvertBoxesToArray = -1
            Exit Function
        End If
    Next i
    
    ConvertBoxesToArray = uCount
End Function
