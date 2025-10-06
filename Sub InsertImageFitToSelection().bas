Attribute VB_Name = "InsertImage"
' ═══════════════════════════════════════════════════════════
' MODULE: InsertImage
' PURPOSE: Inserts and stretches image to fill selected range
' AUTHOR: KnightlyWorks
' VERSION: 1.0
' ═══════════════════════════════════════════════════════════

Sub InsertImageFitToSelection()
    On Error GoTo ErrorHandler
    
    ' === CONFIGURATION ===
    Const FOLDER_PATH As String = "C:\Users\Pictures\"
    Const ALLOW_MULTIPLE_IMAGES As Boolean = False  ' True = keep old images
    Const IS_DEBUG As Boolean = False                ' Enable debug logging
    
    ' === VARIABLES ===
    Dim fd As FileDialog
    Dim selectedFile As String
    Dim img As Shape
    Dim targetRange As Range
    Dim targetSheet As Worksheet
    Dim oldShape As Shape
    Dim shapesToDelete As Collection
    
    ' === VALIDATION: Check if range is selected ===
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select a cell range first!", vbExclamation, "No Selection"
        If IS_DEBUG Then Debug.Print "ERROR: No range selected"
        Exit Sub
    End If
    
    Set targetRange = Selection
    Set targetSheet = targetRange.Worksheet
    
    If IS_DEBUG Then Debug.Print "INSERT IMAGE INTO: " & targetRange.Address & " (Sheet: " & targetSheet.Name & ")"
    
    ' === FILE PICKER DIALOG ===
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd
        .Title = "Select an Image"
        .InitialFileName = FOLDER_PATH
        .Filters.Clear
        .Filters.Add "Images", "*.jpg;*.jpeg;*.png;*.bmp;*.gif;*.tiff;*.webp"
        .Filters.Add "All Files", "*.*"
        .AllowMultiSelect = False
        
        If .Show <> -1 Then
            If IS_DEBUG Then Debug.Print "Action: CANCELED by user"
            Exit Sub
        End If
        
        selectedFile = .SelectedItems(1)
    End With
    
    ' === VALIDATION: Check if file exists ===
    If Dir(selectedFile) = "" Then
        MsgBox "File not found:" & vbCrLf & selectedFile, vbCritical, "Error"
        If IS_DEBUG Then Debug.Print "ERROR: File not found - " & selectedFile
        Exit Sub
    End If
    
    If IS_DEBUG Then Debug.Print "Selected file: " & selectedFile
    
    ' === CLEANUP: Remove old shapes in target area ===
    If Not ALLOW_MULTIPLE_IMAGES Then
        If IS_DEBUG Then Debug.Print "Cleaning up old images..."
        Set shapesToDelete = New Collection
        
        ' Collect shapes to delete (can't delete while iterating)
        For Each oldShape In targetSheet.Shapes
            On Error Resume Next
            If oldShape.Type = msoPicture Or oldShape.Type = msoLinkedPicture Then
                ' Check if shape intersects with target range
                If Not Intersect(targetRange, _
                    Range(oldShape.TopLeftCell, oldShape.BottomRightCell)) Is Nothing Then
                    shapesToDelete.Add oldShape
                    If IS_DEBUG Then Debug.Print "  Marked for deletion: " & oldShape.Name
                End If
            End If
            On Error GoTo ErrorHandler
        Next oldShape
        
        ' Delete collected shapes
        For Each oldShape In shapesToDelete
            oldShape.Delete
        Next oldShape
        
        If IS_DEBUG Then Debug.Print "  Deleted " & shapesToDelete.Count & " old image(s)"
    End If
    
    ' === INSERT IMAGE ===
    If IS_DEBUG Then Debug.Print "Inserting image..."
    Set img = targetSheet.Shapes.AddPicture( _
        Filename:=selectedFile, _
        LinkToFile:=msoFalse, _
        SaveWithDocument:=msoTrue, _
        Left:=0, _
        Top:=0, _
        Width:=-1, _
        Height:=-1)
    
    ' === STRETCH IMAGE TO FIT CELL RANGE ===
    If IS_DEBUG Then Debug.Print "Stretching image to fit cell range (no aspect ratio preservation)"
    
    With img
        ' Disable aspect ratio lock to allow stretching
        .LockAspectRatio = msoFalse
        
        ' Fill the entire cell range
        .Width = targetRange.Width
        .Height = targetRange.Height
        .Left = targetRange.Left
        .Top = targetRange.Top
        .ZOrder msoBringToFront
        
        If IS_DEBUG Then
            Debug.Print "Image name: " & .Name
            Debug.Print "Final size: " & Round(.Width, 1) & " x " & Round(.Height, 1) & " px"
            Debug.Print "Position: Left=" & Round(.Left, 1) & ", Top=" & Round(.Top, 1)
        End If
    End With
    
    If IS_DEBUG Then Debug.Print "SUCCESS! Image inserted and stretched to cell range."
    
    Exit Sub
    
ErrorHandler:
    If IS_DEBUG Then Debug.Print "ERROR #" & Err.Number & ": " & Err.Description
    
    MsgBox "Error occurred:" & vbCrLf & vbCrLf & _
           "Error #" & Err.Number & vbCrLf & _
           Err.Description, _
           vbCritical, "Insert Image Error"
End Sub