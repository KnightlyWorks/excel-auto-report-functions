Attribute VB_Name = "GreenOnOK"
' ═══════════════════════════════════════════════════════════
' MODULE: GreenOnOK
' PURPOSE: Automatically color cells green when "OK" is entered
' AUTHOR: KnightlyWorks
' VERSION: 1.0
' INSTALL: Copy this code to Sheet module (double-click Sheet1)
' ═══════════════════════════════════════════════════════════

' IMPORTANT: This code must be placed in SHEET MODULE, not regular module!
' HOW TO INSTALL:
' 1. Press Alt+F11 to open VBA editor
' 2. Double-click the target sheet on the left (Sheet1, Sheet2, etc.)
' 3. Paste this code into the code window
' 4. Press Alt+Q to exit VBA editor

' === CONFIGURATION ===
Const TRIGGER_TEXT As String = "OK"           ' Text that triggers green color
Const TARGET_COLUMN As String = "A"           ' Column to monitor (A, B, C, etc.)
Const GREEN_COLOR As Long = 5287936          ' Green color (RGB: 90, 200, 90)
Const CASE_SENSITIVE As Boolean = False      ' True = "OK" ≠ "ok", False = "OK" = "ok"
Const IS_DEBUG As Boolean = False            ' Enable debug logging

' === AUTO-COLOR ON CELL CHANGE ===
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim cell As Range
    Dim checkText As String
    Dim targetText As String
    
    On Error Resume Next
    
    For Each cell In Target
        ' Check only cells in target column
        If cell.Column = Range(TARGET_COLUMN & "1").Column Then
            
            targetText = TRIGGER_TEXT
            checkText = Trim(cell.Value)
            
            ' Convert to uppercase if case-insensitive
            If Not CASE_SENSITIVE Then
                targetText = UCase(targetText)
                checkText = UCase(checkText)
            End If
            
            ' Apply green color if trigger text matches
            If checkText = targetText Then
                With cell
                    .Interior.Color = GREEN_COLOR
                    .Font.Color = vbWhite
                    .Font.Bold = True
                End With
                If IS_DEBUG Then Debug.Print cell.Address & " colored green"
            Else
                ' Reset formatting if text doesn't match
                With cell
                    .Interior.ColorIndex = xlNone
                    .Font.Color = vbBlack
                    .Font.Bold = False
                End With
                If IS_DEBUG Then Debug.Print cell.Address & " color reset"
            End If
        End If
    Next cell
End Sub