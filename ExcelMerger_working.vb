' MergeAllWorkbooks
' This macro merges all workbooks in a folder into a single master workbook. The macro opens each target
' workbook in the given folder location, copies all data from each sheet into a corresponding sheet
' in the master workbook. If a sheet contained in a target workbook does not exist in the master workbook
' than a sheet by the same name, with the same field is created in the master workbook.
Sub MergeAllWorkbooks()
    ' Modify These Values accordingly
    Const SavePath As String = ""        ' Set the location to save the master workbook
    Const SaveFile As String = ""        ' Set the name of the master workbook
    Const FolderPath As String = "\"     ' Set folder path to point to the location of the files to be merged
    Const CopyFile As String = "*.xml"   ' Set the file name and extension of the files in FolderPath. filename should always be denoted with a "*" instead of an absolute name
    
    Dim SummaryBook As Workbook         ' the master workbook which will be merged to
    Dim SummarySheet As Worksheet       ' the current sheet in the master workbook being merged to
    Dim SummarySheetLoc As Long         ' the numerical value of the current sheet in the master workbook being merged to

    Dim FileName As String              ' holds the name of a file in FolderPath location
    Dim LastUsedRow As Long             ' the location of the last row filled with data
   
    Dim CopyFrom As Workbook            ' the workbook being copied from
    Dim CopySheet As Worksheet          ' the worksheet being copied from inside the workbook being copied from
    Dim CopySheetName As String         ' the name of the worksheet being copied
    Dim CopySheetCount As Long          ' the number of sheets in the workbook being copied
    Dim CopyRange As String             ' the range of cells to be copied from CopyFrom
    Dim HeaderRange As String           ' the range of header cells to be copied from CopyFrom
    Dim SourceRange As range
    Dim DestRange As range
    
    Dim i As Long                       ' counter
    Dim Alphabet(1 To 26) As String     ' ABCDEFGHIJKLMNOPQRSTUVWXYZ used for detemining column range
    
    ' Fill the alphabet array
    Call FillArray(Alphabet)
    
    ' Set master workbook to the current workbook
    Set SummaryBook = ActiveWorkbook

    ' Get the first file from FolderPath location
    FileName = Dir(FolderPath & CopyFile)
    
    ' loop through Files in FolderPath location
    Do While FileName <> ""
        ' open a workbook in FolderPath
        Set CopyFrom = Workbooks.Open(FolderPath & FileName)
        
        ' get the sheet count of CopyFrom workbook
        CopySheetCount = CopyFrom.Sheets.count
        
        ' loop through sheets is current CopyFrom workbook
        For i = 1 To CopySheetCount
            
            ' get the name of the current sheet from CopyFrom
            CopySheetName = CopyFrom.Worksheets(i).Name
            
            ' set range here
            CopyRange = SetRange(CopyFrom.Sheets(i), Alphabet, "A2", False)
            
            ' check if current sheet exists in master
            SummarySheetLoc = CheckForSheet(SummaryBook, CopySheetName)
            
            ' if sheet not found in mater, then add a new sheet
            If SummarySheetLoc = -1 Then
                ' add new sheet to the end of master
                SummaryBook.Sheets.Add(After:=SummaryBook.Sheets(SummaryBook.Sheets.count)).Name = CopySheetName
                ' set the index of the newly added sheet
                SummarySheetLoc = SummaryBook.Sheets.count
                HeaderRange = SetRange(CopyFrom.Sheets(i), Alphabet, "A1", True)
                
                ' copy headers to newly created sheet in master
                Call CopyCells(CopyFrom.Sheets(i), SummaryBook.Sheets(SummarySheetLoc), HeaderRange, 1)
                            
            End If
            
            ' set SummarySheet to the master sheet with the same name as the current sheet in CopyFrom
            Set SummarySheet = SummaryBook.Sheets(SummarySheetLoc)
            
            ' Set the last row used in master
            LastUsedRow = SummaryBook.Sheets(SummarySheetLoc).Cells(Rows.count, "A").End(xlUp).Row
            LastUsedRow = LastUsedRow + 1
            
            Call CopyCells(CopyFrom.Sheets(i), SummarySheet, CopyRange, LastUsedRow)
                    
        Next i
        
        ' Close the source workbook without saving changes.
        CopyFrom.Close savechanges:=False
        
        ' Use Dir to get the next file name.
        FileName = Dir()
    Loop
    
    Application.DisplayAlerts = False
    SummaryBook.Sheets(1).Delete
    Application.DisplayAlerts = True
    
    ' auto fit the master  so that data is readable
    For i = 1 To SummaryBook.Sheets.count
        SummaryBook.Sheets(i).Columns.AutoFit
    Next i
        
    ' save the master with the predefined name in the predefined location
    SummaryBook.SaveAs (SavePath & SaveFile)

End Sub

' CheckForSheet
' checks to see if a sheet with SheetName exists in Master. If a sheet exists, then the index
' of that sheet (in Master) is returned. If a sheet by SheetName does not exist in Master,
' then -1 is returned.
Function CheckForSheet(ByRef Master As Workbook, ByVal SheetName As String) As Long
    Dim master_count As Long
    Dim master_sheet As String
    Dim i As Long
    
    ' set return val to not found
    CheckForSheet = -1

    ' set master_count to number of sheets in Master
    master_count = Master.Sheets.count

    ' loop through all sheets in Master
    ' if a sheet with a matching name to SheetName is found
    ' then set the return to the index of the sheet in Master
    ' and break out of loop
    For i = 1 To master_count
        master_sheet = Master.Worksheets(i).Name
        If master_sheet = SheetName Then
            CheckForSheet = i
            Exit For
        End If
    Next i

End Function

Sub FillArray(ByRef alpha() As String)
    Dim i As Long
    Dim ch_count As Long
    Dim ch As String
    ch_count = 65 				' ascii 'A' value
    
    For i = 1 To 26
        ch = Chr(ch_count)
        alpha(i) = ch
        ch_count = ch_count + 1
    Next i
End Sub

Function SetRange(ByRef sheet As Worksheet, ByRef alpha() As String, ByVal startCell As String, ByVal isHeader As Boolean)
    Dim colLtr As String
    Dim colInit As Long
    Dim count As Long
    Dim LastRow As Long
    Dim lastCol As Long
    
    LastRow = sheet.Cells(Rows.count, "A").End(xlUp).Row
    lastCol = sheet.Cells(1, Columns.count).End(xlToLeft).Column
    
    colLtr = ""
    colInit = 0
    count = lastCol
    
    Do While count > 26
        count = count - 26
        colInit = colInit + 1
    Loop
    
    If colInit > 0 Then
        colLtr = colLtr + alpha(colInit)
    End If
    
    ' get the col letter
    colLtr = colLtr + alpha(lastCol Mod 26)
    
    If isHeader = False Then
        SetRange = startCell + ":" + CStr(colLtr) + CStr(LastRow)
    Else
        SetRange = startCell + ":" + CStr(colLtr) + "1"
    End If
    
End Function

Sub CopyCells(ByRef SourceSheet As Worksheet, ByRef DestSheet As Worksheet, ByVal CopyRange As String, ByVal LastRow As Long)
    Dim SourceRange As range
    Dim DestRange As range
    
    ' Set the source range
    Set SourceRange = SourceSheet.range(CopyRange)
    
    ' Set the destination range
    Set DestRange = DestSheet.range("A" & LastRow)
    Set DestRange = DestRange.Resize(SourceRange.Rows.count, SourceRange.Columns.count)
    
    ' Copy over the values from the source to the destination.
    DestRange.Value = SourceRange.Value

End Sub


