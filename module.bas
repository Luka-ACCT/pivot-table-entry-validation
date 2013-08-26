'FILL IN THE FIELD FORMAT HERE
Public Const fieldFormat = "[Contact].[Email]"

'Candidate values must start on row 2 of column 1 of the candidateValuesSheet
Public Const candidateValuesSheet = "Emails"

'You can have as many pivots as you want on pivotSheet
Public Const pivotSheet = "Emails"

'=====================================================================================================
'DO NOT CHANGE ANYTHING BELOW
'=====================================================================================================
Sub checkValidPivotEntries()
    Dim rngForArray As Range
    Dim lngArrElements As Long
    Dim arr(0) As Variant
    Dim p As PivotTable
    Dim i As Long
    Dim notValidEntryRow As Long
    
    'generates the longer format of fieldFormat e.g. "[Contact].[Email].[Email]" instead of just "[Contact].[Email]"
    fieldFormat2 = fieldFormat + Right(fieldFormat, Len(fieldFormat) - InStrRev(fieldFormat, "[") + 2)
    
    notValidEntryRow = 2

    'determines how many candidate values to iterate through
    'does not ignore blank cells in the range
    With Worksheets(candidateValuesSheet)
        Set rngForArray = .Range(.Cells(2, "A"), .Cells(.Rows.Count, "A").End(xlUp)) 'set array range, currently starts in A2
        lngArrElements = rngForArray.Cells.Count + 1 'array size & + 1 because number of elements was determined from row 2
    End With
    
    With rngForArray
        For i = 2 To lngArrElements
              'load individual entries into the array
              arr(0) = fieldFormat & ".&[" & _
              Trim(Worksheets(candidateValuesSheet).Cells(i, "A").Value) & "]"
  
              'Debug.Print arr(0)
              
              For Each p In Worksheets(pivotSheet).PivotTables 'updates PivotTables
           
                    p.RefreshTable
                    
                    p.CubeFields(fieldFormat).EnableMultiplePageItems = False
                    
                    p.CubeFields(fieldFormat).EnableMultiplePageItems = True
                    
                    On Error GoTo ErrHandler:
                    
                    p.PivotFields(fieldFormat2).VisibleItemsList = arr()
              
Label1:
                    
              On Error GoTo 0
                    
              Next

        Next
    
    End With

'remove all the blanks from column A now that non-valid entries are removed
removeBlanks (lngArrElements)
        
Exit Sub
    
 
    
    
'if candidate value is not found in EDC
ErrHandler:
    Cells(notValidEntryRow, "B").Value = Cells(i, "A").Value
    Cells(notValidEntryRow, "B").Interior.ColorIndex = 3
    Cells(i, "A") = ""
    notValidEntryRow = notValidEntryRow + 1
Resume Label1:
    
    MsgBox "Error: Code execution should never reach this point."
End Sub


'Resolves column A of the candidateValueSheet to a contiguous column of values without blanks
Sub removeBlanks(lngArrElements As Long)
    Worksheets(candidateValuesSheet).Activate
    For i = 2 To lngArrElements
        If Cells(i, "A") = "" Then
            For a = i + 1 To lngArrElements
                If Cells(a, "A") <> "" Then
                    Cells(i, "A").Value = Cells(a, "A").Value
                    Cells(a, "A") = ""
                    Exit For
                End If
            Next
        End If
    Next
End Sub


Sub loadPivots()

    Dim rngForArray As Range
    Dim lngArrElements As Long
    Dim arr() As Variant
    Dim p As PivotTable
    Dim i As Long
    
    'generates the longer format of fieldFormat e.g. "[Contact].[Email].[Email]" instead of just "[Contact].[Email]"
    fieldFormat2 = fieldFormat + Right(fieldFormat, Len(fieldFormat) - InStrRev(fieldFormat, "[") + 2)
    
    With Sheets(candidateValuesSheet)
        Set rngForArray = .Range(.Cells(2, "A"), .Cells(.Rows.Count, "A").End(xlUp)) 'set array range, currently starts in A2
        lngArrElements = rngForArray.Cells.Count 'array size
    End With
    
    With rngForArray
        For i = 1 To lngArrElements 'does not ignore blank cells in the range
              ReDim Preserve arr(1 To i)   'Re-dimension array as required
              arr(i) = fieldFormat & ".&[" & _
              Trim(Sheets(candidateValuesSheet).Cells(i + 1, "A").Value) & "]" 'modifies entry string - change this!
        Next
    End With
        
    'For i = 1 To lngArrElements
    '    Debug.Print arr(i)
    'Next
    
    'Application.ScreenUpdating = False
    
    Sheets(pivotSheet).Activate
        
    For Each p In Sheets(pivotSheet).PivotTables 'updates pivottables
           
         p.RefreshTable
              
         'ititially setting to False, before setting to True, seems to unselect previously selected items
         p.CubeFields(fieldFormat).EnableMultiplePageItems = False
              
         p.CubeFields(fieldFormat).EnableMultiplePageItems = True
         
         'For i = 1 To UBound(arr)
         '   Debug.Print i & ": " & arr(i)
         'Next
              
         p.PivotFields(fieldFormat2).VisibleItemsList = arr()
            
    Next
    
    'Application.ScreenUpdating = True

End Sub


Sub validateAndLoadPivot()

    checkValidPivotEntries
    loadPivots

End Sub
