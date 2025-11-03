Attribute VB_Name = "Input_Helper"
Option Explicit
' =============================================================================
' Estate Planner Input Helper (In-Sheet Selector + Spacing-Safe Save)
' -----------------------------------------------------------------------------
' What this module provides
'   ¥ In-sheet category selection on the "InputHelper" sheet (no modal).
'   ¥ A structured data-entry table (Field Name, Value, Helper Text).
'   ¥ Safe insertion back into the target sheet while preserving EXACTLY one
'     blank separator row immediately before the next category header.
'   ¥ Helper-text hierarchy lookup and promotion of user overrides.
'
' Sheet assumptions
'   ¥ Data sheets are: Assets, Liabilities, Income, Expenses.
'   ¥ Category headers are bold text in Column B.
'   ¥ Field headers for a category are on the row below the header (categoryRow+1),
'     across columns C..T (3..20). Data rows live below those headers.
'   ¥ "InputHelper" exists and is used for the UI.
'   ¥ "HelperText" exists with columns:
'       A = Sheet, B = Category, C = Field, D = HelperText.
'
' Buttons (Form Controls recommended)
'   ¥ Place a Form Control button named "btnConfirmSelection" on InputHelper; assign macro ConfirmCategory.
'   ¥ Place a Form Control button named "btnSaveAndClose" on InputHelper; assign macro SaveAndClose.
'   ¥ Visibility is toggled automatically to guide the user.
'
' Public entry points
'   ¥ AddNewEntry     ? Prepares in-sheet selector (instructions + drop-down).
'   ¥ ConfirmCategory ? Builds the full InputHelper for the selected category.
'   ¥ SaveAndClose    ? Inserts a new data row safely; promotes helper overrides.
' =============================================================================


' ============================================================================
' PUBLIC ENTRY POINTS
' ============================================================================

' Launches the workflow by preparing the in-sheet category selector on InputHelper.
Public Sub AddNewEntry()
    Dim ws As Worksheet
    Dim categories As Collection

    Set ws = ActiveSheet
    If Not IsValidSheet(ws.Name) Then
        MsgBox "Only valid on Assets, Liabilities, Income, or Expenses tabs.", vbExclamation
        Exit Sub
    End If

    Set categories = FindCategories(ws)
    If categories.Count = 0 Then
        MsgBox "No categories found on '" & ws.Name & "'.", vbExclamation
        Exit Sub
    End If

    ' Build the selector UI on InputHelper and store the target sheet in B2.
    ShowCategorySelector ws.Name, categories
End Sub


' Builds the full InputHelper for the category chosen in InputHelper!B5.
' Hides the confirm button and shows the Save & Close button.
Public Sub ConfirmCategory()
    Dim ih As Worksheet
    Dim targetSheetName As String
    Dim categoryName As String
    Dim wsTarget As Worksheet
    Dim categoryRow As Long
    Dim fields As Collection, helperTexts As Collection

    Set ih = Sheets("InputHelper")

    targetSheetName = Trim$(CStr(ih.Range("B2").value))
    If Len(targetSheetName) = 0 Then
        MsgBox "No target sheet recorded. Click 'Add New Entry' again.", vbExclamation
        Exit Sub
    End If

    categoryName = Trim$(CStr(ih.Range("B5").value))
    If Len(categoryName) = 0 Then
        MsgBox "Please choose a category from the drop-down in cell B5.", vbExclamation
        Exit Sub
    End If

    Set wsTarget = Sheets(targetSheetName)

    categoryRow = FindCategoryRow(wsTarget, categoryName)
    If categoryRow = 0 Then
        MsgBox "Category '" & categoryName & "' not found on '" & targetSheetName & "'.", vbCritical
        Exit Sub
    End If

    Set fields = GetCategoryFields(wsTarget, categoryRow)
    Set helperTexts = GetHelperTexts(wsTarget.Name, categoryName, fields)

    ' Replace the selector view with the full InputHelper table.
    PopulateInputHelper wsTarget.Name, categoryName, fields, helperTexts

    ' Switch the visible button from Confirm ? Save & Close.
    ShowSaveButton

    ih.Activate
    ih.Range("B7").Select
End Sub


' Validates and saves values back to the selected category,
' ensuring proper spacing before the next category header.
Public Sub SaveAndClose()
    Dim ih As Worksheet
    Dim targetSheetName As String, categoryName As String
    Dim ws As Worksheet
    Dim inputData As Collection
    Dim targetRow As Long

    Set ih = Sheets("InputHelper")

    ' These are recorded by PopulateInputHelper when the entry form is built.
    targetSheetName = Trim$(CStr(ih.Range("B2").value))
    categoryName = Trim$(CStr(ih.Range("B3").value))

    If Len(targetSheetName) = 0 Or Len(categoryName) = 0 Then
        MsgBox "No active input session found on InputHelper.", vbExclamation
        Exit Sub
    End If

    Set ws = Sheets(targetSheetName)
    Set inputData = CollectInputData(ih)

    ' Calculate a safe insert row for this category.
    targetRow = FindTargetRowWithBuffer(ws, categoryName)
    If targetRow = 0 Then
        MsgBox "Could not locate a valid row for '" & categoryName & "'.", vbCritical
        Exit Sub
    End If

    ' Insert values into the mapped columns.
    InsertDataToSheet ws, targetRow, inputData

    ' Promote any helper text overrides from InputHelper col C to HelperText.
    UpdateHelperTextFromInput targetSheetName, categoryName

    ' Clear the InputHelper and return to the data sheet.
    WipeInputHelper ih
    ws.Activate

    MsgBox "Entry saved to '" & targetSheetName & "' ? '" & categoryName & "'.", vbInformation
End Sub


' ============================================================================
' IN-SHEET SELECTOR UI
' ============================================================================

' Preps InputHelper with instructions and a drop-down in B5 listing categories.
' Uses a comma-delimited validation list (no temporary range sheet).
' Also toggles buttons to show only the Confirm button at this stage.
Private Sub ShowCategorySelector(ByVal targetSheetName As String, ByVal categories As Collection)
    Dim ih As Worksheet
    Dim i As Long
    Dim catList As String

    Set ih = Sheets("InputHelper")

    ' Clear the sheet first so the validation we apply remains intact.
    WipeInputHelper ih

    ' Header + instructions
    ih.Range("A1").value = "Estate Planner - Data Entry": ih.Range("A1").Font.Bold = True
    ih.Range("A2").value = "Sheet:":  ih.Range("B2").value = targetSheetName
    ih.Range("A3").value = "Select a category to add an entry for:"
    ih.Range("A3").Font.Bold = True
    ih.Range("A5").value = "Category:"
    ih.Range("A5").Font.Bold = True

    ' Build a comma-delimited list for validation.
    ' (Avoid commas in category names; keep display names simple.)
    For i = 1 To categories.Count
        catList = catList & CStr(categories(i)) & ","
    Next
    If Right$(catList, 1) = "," Then catList = Left$(catList, Len(catList) - 1)

    With ih.Range("B5").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:=catList
        .IgnoreBlank = True
        .InCellDropdown = True
    End With

    ' Show only the Confirm button at this stage.
    ShowConfirmButton

    ih.Activate
    ih.Range("B5").Select
End Sub


' Shows the "Confirm Selection" button and hides "Save and Close".
' Form Controls are exposed as Shapes for visibility control.
Private Sub ShowConfirmButton()
    On Error Resume Next
    With Sheets("InputHelper").Shapes
        .Item("btnConfirmSelection").Visible = msoTrue
        .Item("btnSaveAndClose").Visible = msoFalse
    End With
    On Error GoTo 0
End Sub

' Shows the "Save and Close" button and hides "Confirm Selection".
Private Sub ShowSaveButton()
    On Error Resume Next
    With Sheets("InputHelper").Shapes
        .Item("btnConfirmSelection").Visible = msoFalse
        .Item("btnSaveAndClose").Visible = msoTrue
    End With
    On Error GoTo 0
End Sub


' ============================================================================
' VALIDATION & DISCOVERY ON DATA SHEETS
' ============================================================================

' True if the sheet is one of the supported financial tabs.
Private Function IsValidSheet(ByVal sheetName As String) As Boolean
    IsValidSheet = (sheetName = "Assets" Or sheetName = "Liabilities" _
                 Or sheetName = "Income" Or sheetName = "Expenses")
End Function

' Returns all unique category names (bold, non-empty values in column B).
Private Function FindCategories(ByVal ws As Worksheet) As Collection
    Dim cats As New Collection
    Dim lastRow As Long, r As Long, v As String
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).row
    For r = 1 To lastRow
        If ws.Cells(r, 2).Font.Bold Then
            v = Trim$(CStr(ws.Cells(r, 2).value))
            If Len(v) > 0 Then On Error Resume Next: cats.Add v, v: On Error GoTo 0
        End If
    Next
    Set FindCategories = cats
End Function

' Row number of the category header (bold in column B) matching categoryName.
Private Function FindCategoryRow(ByVal ws As Worksheet, ByVal categoryName As String) As Long
    Dim lastRow As Long, r As Long
    lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).row
    For r = 1 To lastRow
        If ws.Cells(r, 2).Font.Bold Then
            If StrComp(Trim$(CStr(ws.Cells(r, 2).value)), categoryName, vbTextCompare) = 0 Then
                FindCategoryRow = r
                Exit Function
            End If
        End If
    Next
    FindCategoryRow = 0
End Function


' ============================================================================
' FIELD & HELPER-TEXT RETRIEVAL
' ============================================================================

' Collects field headers for the category from row (categoryRow+1), columns C..T.
' Returns a Collection of arrays:
'   [0] = Header text
'   [1] = Sample number format (from first data row below)
'   [2] = Sample background color (from first data row below)
Private Function GetCategoryFields(ByVal ws As Worksheet, ByVal categoryRow As Long) As Collection
    Dim fields As New Collection, c As Long, arr As Variant
    For c = 3 To 20
        If Len(Trim$(CStr(ws.Cells(categoryRow + 1, c).value))) > 0 Then
            ReDim arr(0 To 2)
            arr(0) = ws.Cells(categoryRow + 1, c).value
            arr(1) = ws.Cells(categoryRow + 2, c).NumberFormat
            arr(2) = ws.Cells(categoryRow + 2, c).Interior.Color
            fields.Add arr
        End If
    Next
    Set GetCategoryFields = fields
End Function

' Produces helper text for each field using precedence:
'   Specific (Sheet, Category, Field) >
'   Sheet-All (Sheet, "All", Field)  >
'   Global-All ("All", "All", Field)
Private Function GetHelperTexts(ByVal sheetName As String, ByVal categoryName As String, _
                                ByVal fields As Collection) As Collection
    Dim results As New Collection, i As Long, fieldArr As Variant, txt As String, src As String
    For i = 1 To fields.Count
        fieldArr = fields(i)
        txt = ResolveHelperText(sheetName, categoryName, CleanFieldName(CStr(fieldArr(0))), src)
        results.Add txt
    Next
    Set GetHelperTexts = results
End Function

' Removes trailing "*" from field names (used as a required marker in headers).
Private Function CleanFieldName(ByVal s As String) As String
    s = Trim$(s)
    If Len(s) > 0 And Right$(s, 1) = "*" Then s = Left$(s, Len(s) - 1)
    CleanFieldName = s
End Function


' ============================================================================
' INPUTHELPER CONSTRUCTION
' ============================================================================

' Clears InputHelper to a readable baseline and sets column sizes.
Private Sub WipeInputHelper(ByVal inputSheet As Worksheet)
    With inputSheet
        .Cells.Clear
        .Cells.Font.Name = "Calibri": .Cells.Font.Size = 11
        .Columns("A").ColumnWidth = 26
        .Columns("B").ColumnWidth = 28
        .Columns("C").ColumnWidth = 60
        .Rows.RowHeight = 18
        .Range("A:C").VerticalAlignment = xlVAlignTop
        .Range("A:C").HorizontalAlignment = xlLeft
        .Range("C:C").WrapText = True
    End With
End Sub

' Draws light borders around A6:C(last) and bolds the header row.
Private Sub StyleInputHelperRange(ByVal inputSheet As Worksheet, ByVal headerRow As Long, ByVal lastDataRow As Long)
    Dim rng As Range
    If lastDataRow < headerRow Then Exit Sub
    Set rng = inputSheet.Range(inputSheet.Cells(headerRow, 1), inputSheet.Cells(lastDataRow, 3))
    With rng.Borders
        .LineStyle = xlContinuous: .Color = RGB(200, 200, 200): .Weight = xlThin
    End With
    inputSheet.Rows(headerRow).Font.Bold = True
End Sub

' Builds the full InputHelper table (header band, table headers, rows).
' Records the sheet name in B2 and the category name in B3 for later saving.
Private Sub PopulateInputHelper(ByVal sheetName As String, ByVal categoryName As String, _
                                ByVal fields As Collection, ByVal helperTexts As Collection)
    Dim ih As Worksheet, i As Long, r As Long, arr As Variant

    Set ih = Sheets("InputHelper")
    WipeInputHelper ih

    ' Header band
    ih.Range("A1").value = "Estate Planner - Data Entry": ih.Range("A1").Font.Bold = True
    ih.Range("A2").value = "Sheet:":    ih.Range("B2").value = sheetName
    ih.Range("A3").value = "Category:": ih.Range("B3").value = categoryName

    ' Table headers
    ih.Range("A6").value = "Field Name"
    ih.Range("B6").value = "Value"
    ih.Range("C6").value = "Helpful Instructions"
    ih.Range("A6:C6").Font.Bold = True

    ' Data rows
    r = 7
    For i = 1 To fields.Count
        arr = fields(i)
        ih.Cells(r, 1).value = CleanFieldName(CStr(arr(0)))
        With ih.Cells(r, 2)
            .NumberFormat = arr(1)
            .Font.Color = RGB(0, 0, 0)
            .Interior.Pattern = xlNone
        End With
        ih.Cells(r, 3).value = helperTexts(i)
        With ih.Cells(r, 3).Font
            .Italic = True: .Size = 9
        End With
        ih.Cells(r, 3).WrapText = True
        r = r + 1
    Next

    StyleInputHelperRange ih, 6, r - 1

    ' Footer hint
    ih.Cells(r + 2, 2).value = "Click 'Save and Close' when complete"
    ih.Cells(r + 2, 2).Font.Bold = True
End Sub


' ============================================================================
' COLLECT & WRITE DATA
' ============================================================================

' Returns a collection of [FieldName, Value] pairs from InputHelper rows.
Private Function CollectInputData(ByVal ih As Worksheet) As Collection
    Dim data As New Collection, lastRow As Long, r As Long, pair As Variant
    lastRow = ih.Cells(ih.Rows.Count, 1).End(xlUp).row
    For r = 7 To lastRow
        If Len(Trim$(CStr(ih.Cells(r, 1).value))) > 0 Then
            ReDim pair(0 To 1)
            pair(0) = CStr(ih.Cells(r, 1).value)
            pair(1) = ih.Cells(r, 2).value
            data.Add pair
        End If
    Next
    Set CollectInputData = data
End Function

' Writes pairs into the correct columns of targetRow by matching field names
' to the header row at (categoryRow+1).
Private Sub InsertDataToSheet(ByVal ws As Worksheet, ByVal targetRow As Long, ByVal inputData As Collection)
    Dim categoryRow As Long, c As Long, i As Long
    Dim pair As Variant, fieldName As String

    ' Climb up from targetRow to locate the category header (bold in col B).
    categoryRow = targetRow
    Do While categoryRow > 1 And Not (ws.Cells(categoryRow, 2).Font.Bold And Len(Trim$(CStr(ws.Cells(categoryRow, 2).value))) > 0)
        categoryRow = categoryRow - 1
    Loop

    ' Map each field to the right column and write the value.
    For i = 1 To inputData.Count
        pair = inputData(i)
        fieldName = CleanFieldName(CStr(pair(0)))
        For c = 3 To 20
            If CleanFieldName(CStr(ws.Cells(categoryRow + 1, c).value)) = fieldName Then
                ws.Cells(targetRow, c).value = pair(1)
                Exit For
            End If
        Next
    Next
End Sub


' ============================================================================
' ROW PLACEMENT WITH PRESERVED SEPARATOR
' ============================================================================

' Last used row anywhere on the sheet.
Private Function LastUsedDataRow(ByVal ws As Worksheet) As Long
    Dim f As Range
    On Error Resume Next
    Set f = ws.Cells.Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, _
                          SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    On Error GoTo 0
    LastUsedDataRow = IIf(f Is Nothing, 1, f.row)
End Function

' Next category header row (bold in col B) at/after startRow; 0 if none.
Private Function NextCategoryRow(ByVal ws As Worksheet, ByVal startRow As Long) As Long
    Dim r As Long, lastRow As Long
    lastRow = LastUsedDataRow(ws)
    For r = startRow To lastRow
        If ws.Cells(r, 2).Font.Bold And Len(Trim$(CStr(ws.Cells(r, 2).value))) > 0 Then
            NextCategoryRow = r
            Exit Function
        End If
    Next
    NextCategoryRow = 0
End Function

' True if all cells in columns C..T for the row are empty.
Private Function IsRowEmpty(ByVal ws As Worksheet, ByVal rowNum As Long) As Boolean
    Dim c As Long
    For c = 3 To 20
        If Len(Trim$(CStr(ws.Cells(rowNum, c).value))) > 0 Then
            IsRowEmpty = False
            Exit Function
        End If
    Next
    IsRowEmpty = True
End Function

' Returns the last non-empty row (considering columns C..T) between rStart and rEnd, or 0.
Private Function LastDataRowInRange(ByVal ws As Worksheet, ByVal rStart As Long, ByVal rEnd As Long) As Long
    Dim r As Long
    If rEnd < rStart Then
        LastDataRowInRange = 0
        Exit Function
    End If
    For r = rEnd To rStart Step -1
        If Not IsRowEmpty(ws, r) Then
            LastDataRowInRange = r
            Exit Function
        End If
    Next
    LastDataRowInRange = 0
End Function

' Calculates a safe target row for inserting a new entry in categoryName:
'   ¥ First data row is directly under the field headers (categoryRow+2).
'   ¥ Exactly one blank separator row is maintained immediately before the
'     next category header.
'   ¥ If the block is empty, the first entry uses categoryRow+2 (no gap).
Private Function FindTargetRowWithBuffer(ByVal ws As Worksheet, ByVal categoryName As String) As Long
    Dim categoryRow As Long, nextHdr As Long, lastUsed As Long
    Dim dataStart As Long, sepRow As Long, lastData As Long, targetRow As Long

    categoryRow = FindCategoryRow(ws, categoryName)
    If categoryRow = 0 Then Exit Function

    ' First data row lives directly below the field headers.
    dataStart = categoryRow + 2

    ' Locate the next category header; if none, synthesize one just beyond used rows.
    nextHdr = NextCategoryRow(ws, categoryRow + 1)
    If nextHdr = 0 Then
        lastUsed = LastUsedDataRow(ws)
        nextHdr = Application.Max(lastUsed + 2, dataStart + 1)
    End If

    ' Ensure exactly one blank row exists directly before the next header.
    sepRow = nextHdr - 1
    If Not IsRowEmpty(ws, sepRow) Then
        ws.Rows(nextHdr).Insert               ' creates a blank row above the header
        nextHdr = nextHdr + 1
        sepRow = nextHdr - 1
    End If

    ' Determine the last data row within this category block (excluding the separator).
    lastData = LastDataRowInRange(ws, dataStart, sepRow - 1)

    If lastData = 0 Then
        ' No existing rows ? write immediately under the headers.
        targetRow = dataStart
    Else
        ' Append below the last data row.
        targetRow = lastData + 1
    End If

    ' If the target would collide with the separator, make space before it.
    If targetRow >= sepRow Then
        ws.Rows(nextHdr).Insert               ' push header down, separator preserved
        nextHdr = nextHdr + 1
        sepRow = nextHdr - 1
        targetRow = sepRow - 1
    End If

    FindTargetRowWithBuffer = targetRow
End Function


' ============================================================================
' HELPER-TEXT LOOKUP & PROMOTION
' ============================================================================

' Finds the exact (Sheet, Category, Field) row in HelperText; returns 0 if none.
Private Function FindHelperTextRow(ByVal sheetName As String, ByVal categoryName As String, _
                                   ByVal fieldName As String) As Long
    Dim ws As Worksheet, lastRow As Long, r As Long
    Set ws = Sheets("HelperText")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row
    For r = 2 To lastRow
        If ws.Cells(r, 1).value = sheetName _
           And ws.Cells(r, 2).value = categoryName _
           And ws.Cells(r, 3).value = fieldName Then
            FindHelperTextRow = r
            Exit Function
        End If
    Next
    FindHelperTextRow = 0
End Function

' Returns helper text using precedence: Specific > Sheet-All > Global-All.
Private Function ResolveHelperText(ByVal sheetName As String, ByVal categoryName As String, _
                                   ByVal fieldName As String, ByRef sourceLevel As String) As String
    Dim ws As Worksheet, lastRow As Long, r As Long
    Set ws = Sheets("HelperText")
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).row

    ' Specific
    For r = 2 To lastRow
        If ws.Cells(r, 1).value = sheetName _
           And ws.Cells(r, 2).value = categoryName _
           And ws.Cells(r, 3).value = fieldName Then
            sourceLevel = "Specific": ResolveHelperText = ws.Cells(r, 4).value: Exit Function
        End If
    Next
    ' Sheet-All
    For r = 2 To lastRow
        If ws.Cells(r, 1).value = sheetName _
           And ws.Cells(r, 2).value = "All" _
           And ws.Cells(r, 3).value = fieldName Then
            sourceLevel = "SheetAll": ResolveHelperText = ws.Cells(r, 4).value: Exit Function
        End If
    Next
    ' Global-All
    For r = 2 To lastRow
        If ws.Cells(r, 1).value = "All" _
           And ws.Cells(r, 2).value = "All" _
           And ws.Cells(r, 3).value = fieldName Then
            sourceLevel = "GlobalAll": ResolveHelperText = ws.Cells(r, 4).value: Exit Function
        End If
    Next

    sourceLevel = "None": ResolveHelperText = ""
End Function

' Promotes user-edited helper text (InputHelper col C) into HelperText as Specific.
Private Sub UpdateHelperTextFromInput(ByVal sheetName As String, ByVal categoryName As String)
    Dim ih As Worksheet, ht As Worksheet, lastRow As Long, r As Long
    Dim fieldName As String, userText As String, curRow As Long
    Dim inherited As String, src As String, appendRow As Long

    Set ih = Sheets("InputHelper")
    Set ht = Sheets("HelperText")

    lastRow = ih.Cells(ih.Rows.Count, 1).End(xlUp).row
    If lastRow < 7 Then Exit Sub

    For r = 7 To lastRow
        fieldName = Trim$(CStr(ih.Cells(r, 1).value))
        userText = Trim$(CStr(ih.Cells(r, 3).value))
        If Len(fieldName) = 0 Or Len(userText) = 0 Then GoTo NextR

        curRow = FindHelperTextRow(sheetName, categoryName, fieldName)
        If curRow > 0 Then
            If Trim$(CStr(ht.Cells(curRow, 4).value)) <> userText Then
                ht.Cells(curRow, 4).value = userText
            End If
        Else
            inherited = ResolveHelperText(sheetName, categoryName, fieldName, src)
            If Trim$(inherited) <> userText Then
                appendRow = ht.Cells(ht.Rows.Count, 1).End(xlUp).row + 1
                ht.Cells(appendRow, 1).value = sheetName
                ht.Cells(appendRow, 2).value = categoryName
                ht.Cells(appendRow, 3).value = fieldName
                ht.Cells(appendRow, 4).value = userText
            End If
        End If
NextR:
    Next
End Sub

