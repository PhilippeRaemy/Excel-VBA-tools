Attribute VB_Name = "KCalcHelper"
Option Explicit

Sub MergeKCalcResults()
    MergeKCalcResultsImpl "tradercalls.globalgas.s*"
End Sub

Sub MergeKCalcResultsImpl(rootSheetName As String)
    Dim theBook As Workbook
    Dim wb As Workbook
    Dim ws As Worksheet
    
    For Each wb In Application.Workbooks
        If wb.Name Like rootSheetName Then
            If theBook Is Nothing Then
                Set theBook = wb
            Else
                For Each ws In wb.Worksheets
                    ws.Move , theBook.ActiveSheet
                Next ws
            End If
        End If
    Next wb
    
    For Each ws In theBook.Worksheets
        FormatAndPresent ws
    Next ws

End Sub


Sub FormatAndPresent(sh As Worksheet)

Dim periods As Variant, p As Integer, c As Integer
Dim columnNames As Variant
Dim TableName As String

Const KQuery = "=KQuery(""[Data={structure}]|OBS=MID|DateRangeFilter=20220812,20220816,FirstAndLast|horizontal|notempty"",""K|Query Cloud"",$T$2:$W$67)"


Dim pvt As PivotTable

    On Error Resume Next
    sh.Activate
    sh.Select
    On Error GoTo 0
    
    sh.columns("M:CZ").Delete Shift:=xlToLeft ' clear rendering

    periods = sh.UsedRange.columns(3).Value
    For p = LBound(periods, 1) To UBound(periods, 1)
        If IsDate(periods(p, 1)) Then periods(p, 1) = Format(periods(p, 1), "yyyy-mm")
    Next p
    sh.UsedRange.columns(3).Value = periods
    
    columnNames = sh.Range("a1:i1").Value
    For c = LBound(columnNames, 2) To UBound(columnNames, 2)
        If IsEmpty(columnNames(1, c)) Then columnNames(1, c) = "column " & c
    Next c
    sh.Range("a1:i1").Value = columnNames
    
    sh.Cells.EntireColumn.AutoFit
    
    TableName = "Pivot_" & sh.Name
    
    
    Set pvt = ActiveWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=sh.UsedRange.Address, _
        Version:=8) _
            .CreatePivotTable( _
                TableDestination:=sh.Range("m1"), _
                TableName:=TableName, _
                DefaultVersion:=8)
    
    With pvt
        .ColumnGrand = True
        .HasAutoFormat = True
        .DisplayErrorString = False
        .DisplayNullString = True
        .EnableDrilldown = True
        .ErrorString = ""
        .MergeLabels = False
        .NullString = ""
        .PageFieldOrder = 2
        .PageFieldWrapCount = 0
        .PreserveFormatting = True
        .RowGrand = True
        .SaveData = True
        .PrintTitles = False
        .RepeatItemsOnEachPrintedPage = True
        .TotalsAnnotation = False
        .CompactRowIndent = 1
        .InGridDropZones = False
        .DisplayFieldCaptions = True
        .DisplayMemberPropertyTooltips = False
        .DisplayContextTooltips = True
        .ShowDrillIndicators = True
        .PrintDrillIndicators = False
        .AllowMultipleFilters = False
        .SortUsingCustomLists = True
        .FieldListSortAscending = False
        .ShowValuesRow = False
        .CalculatedMembersInFilters = False
        .RowAxisLayout xlCompactRow
    End With
    With pvt.PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    pvt.RepeatAllLabels xlRepeatLabels
    With pvt.PivotFields("series_value_date")
        .Orientation = xlColumnField
        .Position = 1
    End With
    With pvt.PivotFields("period_code")
        .Orientation = xlRowField
        .Position = 1
    End With
    pvt.AddDataField pvt.PivotFields("series_value"), "series_value ", xlSum
    
    On Error Resume Next
    sh.Range("t1").Value = sh.Name
    'sh.Range("t1").Value = Replace(KQuery, "{structure}", sh.Name)
    'If Err.Number <> 0 Then
    '    Debug.Print Err.Number, Err.Description
    'End If
End Sub




Sub AddHeatMap()
    
    Range("AA3:AC3").Select
    columns("A:AC").EntireColumn.AutoFit
    Range(Selection, Selection.End(xlDown)).Select
    Selection.FormatConditions.AddColorScale ColorScaleType:=3
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).ColorScaleCriteria(1).Type = _
        xlConditionValueLowestValue
    With Selection.FormatConditions(1).ColorScaleCriteria(1).FormatColor
        .Color = 7039480
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).ColorScaleCriteria(2).Type = _
        xlConditionValuePercentile
    Selection.FormatConditions(1).ColorScaleCriteria(2).Value = 50
    With Selection.FormatConditions(1).ColorScaleCriteria(2).FormatColor
        .Color = 8711167
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).ColorScaleCriteria(3).Type = _
        xlConditionValueHighestValue
    With Selection.FormatConditions(1).ColorScaleCriteria(3).FormatColor
        .Color = 8109667
        .TintAndShade = 0
    End With
    ActiveWindow.SmallScroll Down:=-93
End Sub

Sub RemapCharts()
    Dim ws As Worksheet
    
    Dim ch As ChartObject
    Dim se As Series
    Dim re As RegExp
    
    Set re = New RegExp
    re.Pattern = "'[^']+'"
    re.Global = True
    
    For Each ws In ActiveWorkbook.Worksheets
        ws.Select
        For Each ch In ws.ChartObjects
            For Each se In ch.Chart.SeriesCollection
                se.Formula = re.Replace(se.Formula, "'" & ws.Name & "'")
                Debug.Print ws.Name, se.Name
                Debug.Print se.Formula
                Debug.Print re.Replace(se.Formula, "'" & ws.Name & "'")
                Debug.Print
            Next se
        Next ch
    Next ws

End Sub


