Attribute VB_Name = "Module1"
Sub Import_vito()

Dim wb As Workbook
Dim wbCSV As Workbook
Dim myPath As String
Dim myFile As Variant
Dim fileType As String
Dim i As Integer
Dim WS As Worksheet
Dim WS_Count As Integer
Dim lRow As Integer
Dim iCntr As Integer
Dim T As Integer
Dim x As Integer
Dim F As Integer
Dim k As Double
Dim f1 As Integer
Dim Dayheatchart As Integer
Dim Hourheatchart As Integer
Dim Daycumulativechart As Integer
Dim Ser As Series
Dim n As Long
Dim Address1 As Range
Dim Address11 As Range
Dim Address2 As Range
Dim Address22 As Range
Dim Address3 As Range
Dim Address33 As Range
Dim ch As chart
Dim Z As Double
Dim Y As Double

Application.ScreenUpdating = False
    
    MsgBox "PLEASE, SELECT A PASTE WITH .CSV FILES.", , "SEARCH SOURCE!"
    
With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select  Folder"
        .AllowMultiSelect = False
        .Show
    myPath = .SelectedItems(1) & "\"
End With

  fileType = "*.csv*"

  myFile = Dir(myPath & fileType)
  
Do While myFile <> ""
    Worksheets.Add(After:=Worksheets("Sheet1")).Name = "Sheet " & i + 2
    With ActiveSheet.QueryTables.Add(Connection:="TEXT;" & myPath & myFile _
            , Destination:=ActiveSheet.Range("$A$1"))
            .Name = myFile
            .FieldNames = True
            .RowNumbers = False
            .FillAdjacentFormulas = False
            .PreserveFormatting = True
            .RefreshOnFileOpen = False
            .RefreshStyle = xlInsertDeleteCells
            .SavePassword = False
            .SaveData = True
            .AdjustColumnWidth = True
            .RefreshPeriod = 0
            .TextFilePromptOnRefresh = False
            .TextFilePlatform = 850
            .TextFileStartRow = 1
            .TextFileParseType = xlDelimited
            .TextFileTextQualifier = xlTextQualifierDoubleQuote
            .TextFileConsecutiveDelimiter = False
            .TextFileTabDelimiter = False
            .TextFileSemicolonDelimiter = False
            .TextFileCommaDelimiter = True
            .TextFileSpaceDelimiter = False
            .TextFileColumnDataTypes = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, _
            1)
            .TextFileTrailingMinusNumbers = True
            .Refresh BackgroundQuery:=False
            
        End With
    i = i + 1
    myFile = Dir
    
Loop
 
Application.ScreenUpdating = True

Sheets("Sheet1").Select
ActiveWindow.SelectedSheets.Delete
 
 Application.Calculation = xlCalculationManual
 Application.ScreenUpdating = False

If Not chartExists("DayXHeatflow.gr") Then
'1.
'3
    Worksheets(1).Select

    ActiveSheet.Shapes.AddChart2(240, xlXYScatterLinesNoMarkers).Select

    Set ch = ActiveChart

With ch

    .Axes(xlValue).MinimumScale = 0
    .Axes(xlCategory).MinimumScale = 0
    .Axes(xlCategory).MaximumScale = 7
    .Location Where:=xlLocationAsNewSheet
End With
 

    Dayheatchart = Charts.Count

Else

    Dayheatchart = Charts("DayXHeatflow.gr").Index

End If


If Not chartExists("HoursXHeatflow.gr") Then

    Worksheets(1).Select
    ActiveSheet.Shapes.AddChart2(240, xlXYScatterLinesNoMarkers).Select
    
    Set ch = ActiveChart

    With ch

    .Axes(xlValue).MinimumScale = 0
    .Axes(xlCategory).MinimumScale = 0
    .Axes(xlCategory).MaximumScale = 200
    .Location Where:=xlLocationAsNewSheet
    
    End With


    Hourheatchart = Charts.Count

Else

    Hourheatchart = Charts("HoursXHeatflow.gr").Index

End If

Worksheets(1).Select

If Not chartExists("DayXCumulativeHeat.gr") Then


    Worksheets(1).Select
    ActiveSheet.Shapes.AddChart2(240, xlXYScatterSmoothNoMarkers).Select
    Set ch = ActiveChart
    
    With ch
    .Axes(xlValue).MinimumScale = 0
    .Axes(xlCategory).MinimumScale = 0
    .Axes(xlCategory).MaximumScale = 7
    .Location Where:=xlLocationAsNewSheet
    
End With
      

    Daycumulativechart = Charts.Count

Else

    Daycumulativechart = Charts("DayXCumulativeHeat.gr").Index

End If
   
   
   Application.ScreenUpdating = True

For Each Worksheet In Application.ActiveWorkbook.Sheets

If Worksheet.Name <> "Chart1" And Worksheet.Name <> "Chart2" And Worksheet.Name <> "Chart3" And Worksheet.Name <> "Chart4" And Worksheet.Name <> "Chart5" And Worksheet.Name <> "Chart6" Then
 
 Application.ScreenUpdating = True
    
    Worksheet.Select
    Worksheet.Name = Worksheet.Range("B10")
    Worksheet.Activate
    
    Cells.Find(What:="ampoule removed", After:=ActiveCell, LookIn:=xlFormulas _
        , LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    ActiveCell.Offset(-2, 0).Range("A1").Select
    Selection.End(xlToLeft).Select
    Selection.End(xlToLeft).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
    
    ' removendo linhas menor que 0
    Cells.Find(What:="reaction start", After:=ActiveCell, LookIn:=xlFormulas _
        , LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
    ActiveCell.Offset(-1, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = "k"
    Selection.End(xlUp).Select
    ActiveCell.Offset(1, 0).Range("A1").Select
    ActiveCell.FormulaR1C1 = "k"
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Selection.EntireRow.Delete
    Range("e1").Select

        
    Rows("1:30").Select
    Application.CutCopyMode = False
    Selection.Insert Shift:=xlDown
    'depois daqui que da erro no b1
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "INFORMATION ABOUT THE MIX:        "
    'aqui embaixo lembrar de colocar a e colar na celula depois
    Range("B1").Value = Range("B40").Value
              
    
    Range("A3:B3").Select
    Range("B3").Activate
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    ActiveCell.FormulaR1C1 = "INFORMATION ABOUT THE MIX IN THE CALORIMETER"
    Range("A4").Select
    ActiveCell.FormulaR1C1 = "START DATE"
    Range("A5").Select
    ActiveCell.FormulaR1C1 = "STOP DATE"
    Range("C6").Select
    ActiveCell.FormulaR1C1 = "min"
    Range("A6").Select
    ActiveCell.FormulaR1C1 = "TIME BETWEEN WATER ADDED TO THE MIX"
    Columns("A:A").Select
        Range("A8").Select
    ActiveCell.FormulaR1C1 = "MASS OF THE SAMPLE (grams)"
    Range("A9").Select
    ActiveCell.FormulaR1C1 = "ISOTHERMAL TEMPERATURE"
    Range("A10").Select
    ActiveCell.FormulaR1C1 = "ISOTHERMAL CALORIMETER USED"
    Range("B10").Select
    ActiveCell.FormulaR1C1 = "TAMAIR 490 - CH1"
    Range("A12").Select
    ActiveCell.FormulaR1C1 = _
        "CALCULATED INFORMATION ABOUT THE MIX IN THE CALORIMETER"
    Range("A12:B12").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    
    Range("B21").Select
    ActiveCell.FormulaR1C1 = "1"
    Range("A22").Select
    ActiveCell.FormulaR1C1 = "BINDERS"
    Range("A26").Select
    ActiveCell.FormulaR1C1 = "NON-BINDERS"
    
    UserForm1.Show
    
    Range("A13").Select
    ActiveCell.FormulaR1C1 = "VOL (m³)"
    Range("B13").Select
    ActiveCell.FormulaR1C1 = "=R[17]C[4]/(R[17]C[1]/(R[17]C[2]/1000))"
    
    Range("F22").Select
    ActiveCell.FormulaR1C1 = "=R[1]C+R[2]C+R[3]C"
    
    Range("A15").Select
    ActiveCell.FormulaR1C1 = "MASS BINDER (g)"
    Range("B15").Select
    ActiveCell.FormulaR1C1 = "=R[7]C[4]"

    'copy and delete data
    Range("B33").Select
    Selection.Copy
    Range("B4").Select
    ActiveSheet.Paste
    Range("B34").Select
    Selection.Copy
    Range("B5").Select
    ActiveSheet.Paste
    Range("B38").Select
    Selection.Copy
    Range("B9").Select
    ActiveSheet.Paste
    Rows("33:33").Select
    Selection.Delete Shift:=xlUp
    Rows("33:33").Select
    Selection.Delete Shift:=xlUp
    Rows("34:34").Select
    Selection.Delete Shift:=xlUp
    Rows("35:35").Select
    Selection.Delete Shift:=xlUp
    Rows("36:36").Select
    Selection.Delete Shift:=xlUp
    Rows("36:36").Select
    Selection.Delete Shift:=xlUp
         Range("A18:F18").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    ActiveCell.FormulaR1C1 = "COMPOSITION OF THE MIX"
    Range("B20").Select
    ActiveCell.FormulaR1C1 = "DENSITY (KG/L)"
    Range("C20").Select
    ActiveCell.FormulaR1C1 = "MASS IN TUBE (g)"
    Range("D20").Select
    ActiveCell.FormulaR1C1 = "VOLUME"
    Range("E20").Select
    ActiveCell.FormulaR1C1 = "VOLUME FRACTION"
    Range("A21").Select
    ActiveCell.FormulaR1C1 = "WATER"

    
    Range("D21").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=(RC[-1]/1000)/RC[-2]"
     Range("D23").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=(RC[-1]/1000)/RC[-2]"
     Range("D24").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=(RC[-1]/1000)/RC[-2]"
     Range("D25").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=(RC[-1]/1000)/RC[-2]"
     Range("D27").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=(RC[-1]/1000)/RC[-2]"
     Range("D28").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=(RC[-1]/1000)/RC[-2]"
     Range("D29").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=(RC[-1]/1000)/RC[-2]"

    Range("A30").Select
    ActiveCell.FormulaR1C1 = "TOTAL"
    
    Range("C30").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=SUM(R[-9]C:R[-1]C)-R[-8]C-R[-4]C"
    
    Range("C30").Select
    Selection.AutoFill Destination:=Range("C30:E30"), Type:=xlFillDefault
    Range("C30:E30").Select
    
    Range("E21").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]/R30C4"
    Range("E23").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]/R30C4"
    Range("E24").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]/R30C4"
    Range("E25").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]/R30C4"
    Range("E27").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]/R30C4"
    Range("E28").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]/R30C4"
    Range("E29").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]/R30C4"
        
        
    Range("F21").Select
    ActiveCell.FormulaR1C1 = "=RC[-3]*(RC[-1]*(R30C6*R30C4)/R30C3)/RC[-2]"
    Range("F23").Select
    ActiveCell.FormulaR1C1 = "=RC[-3]*(RC[-1]*(R30C6*R30C4)/R30C3)/RC[-2]"
    Range("F24").Select
    ActiveCell.FormulaR1C1 = "=RC[-3]*(RC[-1]*(R30C6*R30C4)/R30C3)/RC[-2]"
    Range("F25").Select
    ActiveCell.FormulaR1C1 = "=RC[-3]*(RC[-1]*(R30C6*R30C4)/R30C3)/RC[-2]"
    Range("F27").Select
    ActiveCell.FormulaR1C1 = "=RC[-3]*(RC[-1]*(R30C6*R30C4)/R30C3)/RC[-2]"
    Range("F28").Select
    ActiveCell.FormulaR1C1 = "=RC[-3]*(RC[-1]*(R30C6*R30C4)/R30C3)/RC[-2]"
    Range("F29").Select
    ActiveCell.FormulaR1C1 = "=RC[-3]*(RC[-1]*(R30C6*R30C4)/R30C3)/RC[-2]"
    
    
    Range("F30").Select
    ActiveCell.FormulaR1C1 = "=R[-22]C[-4]"
    
    Range("A31").Select
    Selection.ClearContents
    
    Range("A18:F30").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("A27:F29").Select
    Range("F29").Activate
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("A23:F25").Select
    Range("F25").Activate
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("A21:F21").Select
    Range("F21").Activate
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("A18:F18").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("A20:F20").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With

   
   
    Range("B6").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("B8").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("B13").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
     Range("B15").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("E37:F37").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.ClearContents
    Range("G37").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Cut
    Range("E37").Select
    ActiveSheet.Paste
    'Segunda parte organizando as
    Range("A37").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(1, 0).Range("A1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=ROW(R[-1]C)"
    T = ActiveCell.Value
    ' Macro44 Macro
    Range("G36").Select
    ActiveCell.FormulaR1C1 = "Time Correction Added water in sec."
    Range("G36:G37").Select
    Range("G37").Activate
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    Range("G38").Select
    ActiveCell.FormulaR1C1 = "=R[-32]C[-5]*60"
    Range("H36").Select
    ActiveCell.FormulaR1C1 = "Time"
    Range("H37").Select
    ActiveCell.FormulaR1C1 = "Sec"
    Range("H38").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-5]=""nan"","""",RC[-7]+R38C7)"
    Range("H38").Select
    Selection.AutoFill Destination:=Range("H38:H" & T), Type:=xlFillDefault
    Range("I36").Select
    ActiveCell.FormulaR1C1 = "Heat Flow"
    Range("I37").Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    ActiveCell.FormulaR1C1 = "W"
    Range("I38").Select
    ActiveCell.FormulaR1C1 = "=IF(RC[-6]=""nan"","""",RC[-6])"
    Range("I38").Select
    Selection.AutoFill Destination:=Range("I38:I" & T), Type:=xlFillDefault
 
    Range("G36:G37").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    Range("G38").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Range("K36").Select
    ActiveCell.FormulaR1C1 = "Time"
    Range("K37").Select
    ActiveCell.FormulaR1C1 = "sec"
    Range("L36").Select
    ActiveCell.FormulaR1C1 = "Time"
    Range("M36").Select
    ActiveCell.FormulaR1C1 = "Time"
    Range("N36").Select
    ActiveCell.FormulaR1C1 = "Time"
    Range("L37").Select
    ActiveCell.FormulaR1C1 = "min"
    Range("M37").Select
    ActiveCell.FormulaR1C1 = "hours"
    Range("N37").Select
    ActiveCell.FormulaR1C1 = "days"
    Range("K38").Select
    ActiveCell.FormulaR1C1 = "=RC[-3]"
    Range("K38").Select
    Selection.AutoFill Destination:=Range("K38:K" & T), Type:=xlFillDefault
    Range("L38").Select
    ActiveCell.FormulaR1C1 = "=RC[-1]/60"
    Range("L38").Select
    Selection.AutoFill Destination:=Range("L38:L" & T), Type:=xlFillDefault
    Range("M38").Select
    ActiveCell.FormulaR1C1 = "=RC[-2]/3600"
    Range("M38").Select
    Selection.AutoFill Destination:=Range("M38:M" & T), Type:=xlFillDefault
    Range("N38").Select
    ActiveCell.FormulaR1C1 = "=RC[-3]/86400"
    Range("N38").Select
    Selection.AutoFill Destination:=Range("N38:N" & T), Type:=xlFillDefault
    Range("P35").Select
    ActiveCell.FormulaR1C1 = "HEAT FLOW"
    Range("P36").Select
    Range("P35:Q35").Select
    Range("Q35").Activate
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("P36").Select
    ActiveCell.FormulaR1C1 = " Heat Flow (mW/gr.poz)"
    Range("P37").Select
    ActiveCell.FormulaR1C1 = "mW/gr.poz"
    Range("P38").Select
    ActiveCell.FormulaR1C1 = "=IFERROR((RC[-7]*1000/R15C2),"""")"
    Range("P38").Select
    Selection.AutoFill Destination:=Range("P38:P" & T), Type:=xlFillDefault
    Range("Q37").Select
    ActiveCell.FormulaR1C1 = "mW/m³"
    Range("Q38").Select
    ActiveCell.FormulaR1C1 = "=IFERROR(RC[-8]*1000/R13C2,"""")"
    Selection.AutoFill Destination:=Range("Q38:Q" & T), Type:=xlFillDefault
    Range("S35").Select
    ActiveCell.FormulaR1C1 = "Cumulative Heat calculation"
    Range("S35:y35").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Range("S36").Select
    ActiveCell.FormulaR1C1 = "Heat (J) Cumulative"
    Range("u36").Select
    ActiveCell.FormulaR1C1 = " Heat (J/gr.poz)"
    Range("v36").Select
    ActiveCell.FormulaR1C1 = " Heat(J/m³)"
    Range("S37").Select
    ActiveCell.FormulaR1C1 = "0"
    Range("x36").Select
    ActiveCell.FormulaR1C1 = "Cumulative Heat J/gr.poz"
    Range("y36").Select
    ActiveCell.FormulaR1C1 = "Cumulative Heaat J/m³"
    Range("S37:V37").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
Application.Calculation = xlCalculationAutomatic

Set WS = ActiveSheet

For x = 40 To T

If WS.Range("i" & x) > 0 Then
WS.Range("i" & x).Select
F = ActiveCell.Row

Exit For

End If

Next

Application.Calculation = xlCalculationManual

f1 = T - F

  ActiveCell.Offset(0, 10).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=((R[1]C[-10]+RC[-10])*0.5*(R[1]C[-11]-RC[-11]))+R[-1]C"
  
    
    ActiveCell.Offset(0, 2).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[-2]/R15C2"
    
    ActiveCell.Offset(0, 1).Range("A1").Select
    ActiveCell.FormulaR1C1 = "=RC[-3]/R13C2"
    
    
    ActiveCell.Offset(0, -3).Range("A1").Select
    Selection.AutoFill Destination:=ActiveCell.Range("A1:A" & f1), Type:= _
        xlFillDefault
    ActiveCell.Range("A1:A" & f1).Select
    Selection.End(xlUp).Select
    Selection.End(xlDown).Select
    
    ActiveCell.Offset(0, 2).Range("A1").Select
    Selection.AutoFill Destination:=ActiveCell.Range("A1:A" & f1)
    ActiveCell.Range("A1:A" & f1).Select
    Selection.End(xlUp).Select
    Selection.End(xlDown).Select
    
    ActiveCell.Offset(0, 1).Range("A1").Select
    Selection.AutoFill Destination:=ActiveCell.Range("A1:A" & f1)
    ActiveCell.Range("A1:A" & f1).Select
    Selection.End(xlUp).Select
    Selection.End(xlDown).Select
     
    
    Range("U38").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(0, 3).Range("A1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-3]"
    Selection.AutoFill Destination:=ActiveCell.Range("A1:A" & f1)
    ActiveCell.Range("A1:A" & f1).Select
    Range("U38").Select
    Selection.End(xlDown).Select
    Selection.Copy
    ActiveCell.Offset(0, 3).Range("A1").Select
    ActiveCell.Offset(1, 0).Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlSubtract, _
        SkipBlanks:=False, Transpose:=False
    
        
    Range("u38").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(0, 4).Range("A1").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-6]"
    Selection.AutoFill Destination:=ActiveCell.Range("A1:A" & f1)
    ActiveCell.Range("A1:A" & f1).Select
    Range("v38").Select
    Selection.End(xlDown).Select
    Selection.Copy
    ActiveCell.Offset(0, 3).Range("A1").Select
    ActiveCell.Offset(1, 0).Range("A1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlSubtract, _
        SkipBlanks:=False, Transpose:=False
    
ActiveWindow.DisplayGridlines = False
Columns("A:Y").EntireColumn.AutoFit
    
    Range("L38:N38").Select
    Range("N38").Activate
    Selection.ClearContents
    
     Range("N39").Select
    Range(Selection, Selection.End(xlDown)).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    Range("P39").Select
    Range(Selection, Selection.End(xlDown)).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    
    Range("x39").Select
    Range(Selection, Selection.End(xlDown)).Select
    With Selection.Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    ActiveCell.Offset(1, 0).Range("A1").Select
    Selection.End(xlDown).Select
    Range(Selection, Selection.End(xlDown)).Select
    With Selection.Interior
    .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent4
        .TintAndShade = 0.799981688894314
        .PatternTintAndShade = 0
    End With
    
     Range("Y39").Select
    Selection.End(xlDown).Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range(Selection, Selection.End(xlToLeft)).Select
    Range(Selection, Selection.End(xlToLeft)).Select
    With Selection.Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.599993896298105
        .PatternTintAndShade = 0
    End With
     
    Range("A1:B1").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    
    Rows("36:37").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
  
Set WS = ActiveSheet

WS.Select

With ActiveSheet

.Range("N40:N" & T).Select
Set Address1 = Selection

.Range("P40:P" & T).Select
Set Address11 = Selection

.Range("M40:M" & T).Select
Set Address2 = Selection

.Range("P40:P" & T).Select
Set Address22 = Selection

.Range("N40:N" & T).Select
Set Address3 = Selection

.Range("X40:X" & T).Select
Set Address33 = Selection

End With
    
Charts(Dayheatchart).Select
    
 Set Ser = ActiveChart.SeriesCollection.NewSeries

    With Ser
        .Name = WS.Name
        .XValues = "=" & Address1.Address(False, False, xlA1, xlExternal)
        .Values = "=" & Address11.Address(False, False, xlA1, xlExternal)
    End With
    
Charts(Hourheatchart).Select

    Set Ser = ActiveChart.SeriesCollection.NewSeries

    With Ser
        .Name = WS.Name
        .XValues = "=" & Address2.Address(False, False, xlA1, xlExternal)
        .Values = "=" & Address22.Address(False, False, xlA1, xlExternal)
    End With

Charts(Daycumulativechart).Select

    Set Ser = ActiveChart.SeriesCollection.NewSeries

    With Ser
        .Name = WS.Name
        .XValues = "=" & Address3.Address(False, False, xlA1, xlExternal)
        .Values = "=" & Address33.Address(False, False, xlA1, xlExternal)
    End With

End If

Next Worksheet

Charts(Dayheatchart).Select
ActiveChart.FullSeriesCollection(1).Delete

Charts(Hourheatchart).Select
ActiveChart.FullSeriesCollection(1).Delete

Charts(Daycumulativechart).Select
ActiveChart.FullSeriesCollection(1).Delete
 

Application.Calculation = xlCalculationAutomatic
Application.ScreenUpdating = True

Charts(Dayheatchart).Select
ActiveChart.Axes(xlCategory).Select
ActiveChart.SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
Selection.Caption = "Time (Day)"
ActiveChart.HasLegend = True
ActiveChart.Legend.Select
ActiveChart.SetElement (msoElementLegendTop)
ActiveChart.Axes(xlValue).Select
ActiveChart.SetElement (msoElementPrimaryValueAxisTitleAdjacentToAxis)
With ActiveChart.Axes(xlValue)
.HasTitle = True
.AxisTitle.Caption = "Heat flow (mW/grpoz)"
End With

Charts(Hourheatchart).Select
ActiveChart.Axes(xlCategory).Select
ActiveChart.SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
Selection.Caption = "Time (Hour)"
ActiveChart.HasLegend = True
ActiveChart.Legend.Select
ActiveChart.SetElement (msoElementLegendTop)
ActiveChart.Axes(xlValue).Select
ActiveChart.SetElement (msoElementPrimaryValueAxisTitleAdjacentToAxis)
With ActiveChart.Axes(xlValue)
.HasTitle = True
.AxisTitle.Caption = "Heat flow (mW/grpoz)"
End With

Charts(Daycumulativechart).Select
ActiveChart.Axes(xlCategory).Select
ActiveChart.SetElement (msoElementPrimaryCategoryAxisTitleAdjacentToAxis)
Selection.Caption = "Time (Day)"
ActiveChart.HasLegend = True
ActiveChart.Legend.Select
ActiveChart.SetElement (msoElementLegendTop)
ActiveChart.Axes(xlValue).Select
ActiveChart.SetElement (msoElementPrimaryValueAxisTitleAdjacentToAxis)
With ActiveChart.Axes(xlValue)
.HasTitle = True
.AxisTitle.Caption = "Cumulative Heat (J/grpoz)"
End With

MsgBox "IN CASE A WRONG VALUE IS INSERTED, CORRECT IT MANUALLY OR RESET MACRO.", vbExclamation, "NOTICE!"
              
End Sub
Private Function chartExists(chartToFind As String) As Boolean

   

    Dim ch As chart

    chartExists = False


For Each ch In Charts

        If chartToFind = ch.Name Then

            chartExists = True

            Exit Function

        End If

Next ch

   

End Function


