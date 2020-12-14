VERSION 5.00
Begin VB.Form frmSudoku 
   BorderStyle     =   1  'Fest Einfach
   Caption         =   "Sudoku"
   ClientHeight    =   6975
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9360
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSudoku.frx":0000
   MaxButton       =   0   'False
   ScaleHeight     =   465
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   624
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton cmdOpen 
      Height          =   495
      Left            =   8160
      Picture         =   "frmSudoku.frx":0442
      Style           =   1  'Grafisch
      TabIndex        =   18
      Top             =   3120
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Height          =   495
      Left            =   6960
      Picture         =   "frmSudoku.frx":0544
      Style           =   1  'Grafisch
      TabIndex        =   17
      Top             =   3120
      UseMaskColor    =   -1  'True
      Width           =   1095
   End
   Begin VB.PictureBox picSave 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   960
      ScaleHeight     =   3465
      ScaleWidth      =   3825
      TabIndex        =   16
      Top             =   2640
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.CommandButton cmdCopy 
      Caption         =   "&Bild kopieren"
      Height          =   495
      Left            =   6960
      TabIndex        =   15
      ToolTipText     =   "Kopiert das Rätsel als Bild in die Zwischenablage"
      Top             =   2520
      Width           =   2295
   End
   Begin VB.CommandButton cmdSolve 
      Caption         =   "&Rätsel lösen..."
      Height          =   495
      Left            =   6960
      TabIndex        =   14
      ToolTipText     =   "Geben Sie Ihr Sudoku ein und lassen Sie es lösen"
      Top             =   1920
      Width           =   2295
   End
   Begin VB.OptionButton optDifficulty 
      Caption         =   "Profi"
      Height          =   375
      Index           =   5
      Left            =   6960
      TabIndex        =   13
      Top             =   6525
      Width           =   2295
   End
   Begin VB.OptionButton optDifficulty 
      Caption         =   "Knifflig"
      Height          =   375
      Index           =   4
      Left            =   6960
      TabIndex        =   12
      Top             =   6165
      Width           =   2295
   End
   Begin VB.OptionButton optDifficulty 
      Caption         =   "Schwer"
      Height          =   375
      Index           =   3
      Left            =   6960
      TabIndex        =   11
      Top             =   5805
      Width           =   2295
   End
   Begin VB.OptionButton optDifficulty 
      Caption         =   "Mittel"
      Height          =   375
      Index           =   2
      Left            =   6960
      TabIndex        =   10
      Top             =   5445
      Width           =   2295
   End
   Begin VB.OptionButton optDifficulty 
      Caption         =   "Leicht"
      Height          =   375
      Index           =   1
      Left            =   6960
      TabIndex        =   9
      Top             =   5085
      Width           =   2295
   End
   Begin VB.OptionButton optDifficulty 
      Caption         =   "Sehr leicht"
      Height          =   375
      Index           =   0
      Left            =   6960
      TabIndex        =   8
      Top             =   4725
      Value           =   -1  'True
      Width           =   2295
   End
   Begin VB.CheckBox chkPossibilities 
      Caption         =   "Möglichkeiten anzeigen"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6960
      TabIndex        =   6
      Top             =   3975
      Value           =   1  'Aktiviert
      Width           =   2295
   End
   Begin VB.CheckBox chkAllowWrong 
      Caption         =   "Falscheingaben zulassen"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6960
      TabIndex        =   5
      ToolTipText     =   "Wenn angewählt, können Sie keine falschen Zahlen eintragen"
      Top             =   3705
      Width           =   2295
   End
   Begin VB.CommandButton cmdProve 
      Caption         =   "&Lösungsabgleich"
      Height          =   495
      Left            =   6960
      TabIndex        =   4
      Top             =   720
      Width           =   2295
   End
   Begin VB.CommandButton cmdShowSolution 
      Caption         =   "&Auflösen"
      Height          =   495
      Left            =   6960
      TabIndex        =   3
      Top             =   1320
      Width           =   2295
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "&Neues Rätsel"
      Height          =   495
      Left            =   6960
      TabIndex        =   2
      Top             =   120
      Width           =   2295
   End
   Begin VB.PictureBox picGame 
      Appearance      =   0  '2D
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontTransparent =   0   'False
      ForeColor       =   &H80000008&
      Height          =   6765
      Left            =   120
      ScaleHeight     =   451
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   451
      TabIndex        =   0
      Top             =   120
      Width           =   6765
      Begin VB.TextBox txtInput 
         Alignment       =   2  'Zentriert
         Appearance      =   0  '2D
         BorderStyle     =   0  'Kein
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3120
         MaxLength       =   1
         TabIndex        =   1
         Top             =   1920
         Visible         =   0   'False
         Width           =   450
      End
   End
   Begin VB.Label lblPossibilities 
      Caption         =   "Mögliche Zahlen:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      TabIndex        =   7
      Top             =   4230
      Width           =   2295
   End
End
Attribute VB_Name = "frmSudoku"
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
Private uX As Integer                   'Aktuelles Feld (X)  (current field: X)
Private uY As Integer                   'Aktuelles Feld (Y)  (current field: Y)
Private bProve As Boolean               'Lösungsabgleich     (prove the game)
Private stdpicClipboard As StdPicture
 
'Klick auf "Möglichkeiten anzeigen"
'Click on "show possible values"
Private Sub chkPossibilities_Click()
    If chkPossibilities.value = 0 Then
        lblPossibilities.Caption = "Mögliche Zahlen:" & vbCrLf & "(Option ausgeschaltet)"
    End If
End Sub

'Kopiert das Bild in die Zwischenablage
'Copy the game as a picture to the clipboard
Private Sub cmdCopy_Click()
    picGame.AutoRedraw = True
    picGame_Paint
    picGame.Picture = picGame.Image
    picSave.Picture = picGame.Image
    
    Set stdpicClipboard = picSave.Picture
    Clipboard.Clear
    Clipboard.SetData stdpicClipboard, vbCFBitmap
    
    picGame.AutoRedraw = False
    
    Set picGame.Picture = Nothing
End Sub

'Neues Rätsel erzeugen
'generate new game
Private Sub cmdNew_Click()
    Screen.MousePointer = vbHourglass
    mSudoku.Generate
    picGame.Refresh
    uX = -1: uY = -1
    Me.lblPossibilities.Caption = "Mögliche Zahlen:" & IIf(chkPossibilities.value = 1, "", vbCrLf & "(Option ausgeschaltet)")
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdOpen_Click()
    Dim sFile As String
    If mCommonDialog.VBGetOpenFileName(sFile, "Sudoku", True, False, False, True, "Sudoku-Rätsel (*.sud)|*.sud", , , "Sudoku speichern", "sud", hWnd) Then
        LoadGame sFile
    End If
End Sub

'Mit Lösung abgleichen
'prove all entered numbers
Private Sub cmdProve_Click()
    bProve = True
    picGame.Refresh
End Sub

Private Sub cmdSave_Click()
    Dim sFile As String
    If mCommonDialog.VBGetSaveFileName(sFile, "Sudoku", True, "Sudoku-Rätsel (*.sud)|*.sud", , , "Sudoku speichern", "sud", hWnd) Then
        SaveGame sFile
    End If
End Sub

'Lösung eintragen
'show the solution
Private Sub cmdShowSolution_Click()
    Dim x As Integer, y As Integer
    For x = 0 To 8
        For y = 0 To 8
            SetValue x, y, GetSolutionValue(x, y)
        Next y
    Next x
    picGame.Refresh
End Sub

'Ein Sudoku lösen, frmSolve anzeigen
'show frmSolve, to solve a sudoku
Private Sub cmdSolve_Click()
    frmSolve.Show
End Sub

Private Sub Form_Load()
    'Optionen laden
    'Load options
    optDifficulty(VBA.GetSetting(App.Title, "Options", "Difficulty", 0)).value = True
    chkPossibilities.value = VBA.GetSetting(App.Title, "Options", "Possibilities", 1)
    chkAllowWrong.value = VBA.GetSetting(App.Title, "Options", "AllowWrong", 0)
    
    Randomize Second(Time)
    cmdNew_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Optionen speichern
    'Save Options
    VBA.SaveSetting App.Title, "Options", "Difficulty", mSudoku.Difficulty
    VBA.SaveSetting App.Title, "Options", "Possibilities", chkPossibilities.value
    VBA.SaveSetting App.Title, "Options", "AllowWrong", chkAllowWrong.value
End Sub

'Klick auf einen Schwierigkeitsgrad
'click on a difficulty option button
Private Sub optDifficulty_Click(Index As Integer)
    mSudoku.Difficulty = Index
End Sub

'Mausklick auf die Spielfläche
'click on the game area
Private Sub picGame_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'Evtl. altes Feld updaten
    'eventually update an old field
    If txtInput.Text <> "" And txtInput.Visible = True Then
        mSudoku.SetValue uX, uY, CByte(txtInput.Text)
    ElseIf txtInput.Text = "" Then
        mSudoku.SetValue uX, uY, 0
    End If
    
    'Neues aktuelles Feld
    'new current field
    uX = Int(x / 50)
    uY = Int(y / 50)
    
    'Textfeld positionieren
    'place text field at this position
    txtInput.Left = uX * 50 + 25 - txtInput.Width / 2
    txtInput.TOp = uY * 50 + 25 - txtInput.Height / 2
    
    'Wert auslesen
    'read current value
    If mSudoku.GetValue(uX, uY) <> 0 Then
        txtInput.Text = mSudoku.GetValue(uX, uY)
    Else
        txtInput.Text = ""
    End If
    
    'Möglichkeiten anzeigen
    'show possible numbers
    If chkPossibilities.value = 1 Then
        Dim sPossible As String
        
        'Testen, ob nicht vorgegebene Zahl
        'check whether given number or not
        If IsGiven(uX, uY) = True Then
            sPossible = "(Vorgegebene Zahl)"
        Else
            sPossible = GetPossibleValues(uX, uY, mSudoku.bytGame)
            'String etwas bearbeiten
            'format this string a little
            Dim i As Integer
            For i = 0 To Len(sPossible) - 1
                sPossible = Mid$(sPossible, 1, i * 2) & Mid$(sPossible, i * 2 + 1, 1) & " " & Mid$(sPossible, i * 2 + 2)
            Next i
        End If
        lblPossibilities.Caption = "Mögliche Zahlen:" & vbCrLf & sPossible
    End If
    
    'Textfeld ggf. anzeigen
    'eventually show text field
    If mSudoku.IsGiven(uX, uY) = False Then
        txtInput.Visible = True
        txtInput.SelStart = 0: txtInput.SelLength = 1
        txtInput.SetFocus
    Else
        txtInput.Visible = False
    End If
End Sub

Private Sub picGame_Paint()
    'picGame.Cls
    
    'Gitternetzlinien zeichnen
    'paint the gridlines
    picGame.ForeColor = vbBlack
    Dim i As Integer
    For i = 0 To 9
        '3er-Pakete mit stärkerem Rand versehen
        '3x3-packages get a stronger border
        If i Mod 3 = 0 Then
            picGame.DrawWidth = 3
        Else
            picGame.DrawWidth = 1
        End If
        
        'Horizontale Linie
        'horizontal line
        picGame.Line (0, i * 50)-(450, i * 50)
        'Vertikale Linie
        'vertical line
        picGame.Line (i * 50, 0)-(i * 50, 450)
    Next i
    
    'Jetzt Zahlen reinschreiben
    'now print the numbers
    picGame.ForeColor = vbBlack
    Dim x As Integer, y As Integer
    For x = 0 To 8
        For y = 0 To 8
        
            If bProve = True Then
                If GetValue(x, y) <> GetSolutionValue(x, y) Then
                    picGame.ForeColor = vbRed
                Else
                    picGame.ForeColor = vbBlack
                End If
            End If
            
            picGame.CurrentX = x * 50 + 25 - picGame.TextWidth(GetValue(x, y))
            picGame.CurrentY = y * 50 + 25 - picGame.TextHeight(GetValue(x, y)) / 2
            
            If mSudoku.GetValue(x, y) <> 0 Then
                picGame.Print mSudoku.GetValue(x, y)
            End If
        Next y
    Next x
    
    bProve = False
End Sub

'Nur Zahlen zwischen 1 und 9 als Eingaben zulassen
'Only allow numbers between 1 and 9
Private Sub txtInput_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub

    'Enter abfangen
    'check for return key
    If KeyAscii = 13 Then
        If Len(txtInput.Text) = 1 Then
            If chkAllowWrong.value = 0 And CByte(txtInput.Text) = mSudoku.GetSolutionValue(uX, uY) Then
                mSudoku.SetValue uX, uY, CByte(txtInput.Text)
            Else
                mSudoku.SetValue uX, uY, CByte(txtInput.Text)
            End If
        Else
            mSudoku.SetValue uX, uY, 0
        End If
        txtInput.Visible = False
    End If

    'Alles außer Zahlen abfangen
    'only allow 1 to 9
    If (KeyAscii < 49 Or KeyAscii > 57) Then
        KeyAscii = 0
        Exit Sub
    End If
    
    'Wenn "Falscheingaben zulassen" (safe mode) abgewählt ist, überprüfen
    'if wrong numbers are not permitted (safe mode), check it
    If chkAllowWrong.value = 0 Then
        If CByte(Chr$(KeyAscii)) <> mSudoku.GetSolutionValue(uX, uY) Then
            Beep
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtInput_LostFocus()
    If Not (ActiveControl.Name = "txtInput") Then
        txtInput_KeyPress 13
    End If
    If mSudoku.IsFinished Then
        If MsgBox("Möchten Sie ein neues Rätsel erzeugen?", vbQuestion + vbYesNo, "Neues Rätsel?") = vbYes Then
            cmdNew_Click
        End If
    End If

End Sub

