Attribute VB_Name = "OneOffs"
' ####################
' \\GVA0MS01\RAEMYP\Excel\Copy of PRTools.xlsm.OneOffs.bas
' ####################
Sub Macro2()
'
' Macro2 Macro
'
Dim mkt() As String
ReDim Preserve mkt(0): mkt(UBound(mkt)) = "aoc"
ReDim Preserve mkt(UBound(mkt) + 1): mkt(UBound(mkt)) = "baumgarten"
ReDim Preserve mkt(UBound(mkt) + 1): mkt(UBound(mkt)) = "gaspool"
ReDim Preserve mkt(UBound(mkt) + 1): mkt(UBound(mkt)) = "german_spark_spread_power_price"
ReDim Preserve mkt(UBound(mkt) + 1): mkt(UBound(mkt)) = "german_spark_spread_spark_spread"
ReDim Preserve mkt(UBound(mkt) + 1): mkt(UBound(mkt)) = "german_spark_spread_ttf"
ReDim Preserve mkt(UBound(mkt) + 1): mkt(UBound(mkt)) = "nbp"
ReDim Preserve mkt(UBound(mkt) + 1): mkt(UBound(mkt)) = "ncg"
ReDim Preserve mkt(UBound(mkt) + 1): mkt(UBound(mkt)) = "nordpool"
ReDim Preserve mkt(UBound(mkt) + 1): mkt(UBound(mkt)) = "pegnord"
ReDim Preserve mkt(UBound(mkt) + 1): mkt(UBound(mkt)) = "pegsud"
ReDim Preserve mkt(UBound(mkt) + 1): mkt(UBound(mkt)) = "pegtigf"
ReDim Preserve mkt(UBound(mkt) + 1): mkt(UBound(mkt)) = "psv"
ReDim Preserve mkt(UBound(mkt) + 1): mkt(UBound(mkt)) = "ttf"
ReDim Preserve mkt(UBound(mkt) + 1): mkt(UBound(mkt)) = "vob"
ReDim Preserve mkt(UBound(mkt) + 1): mkt(UBound(mkt)) = "zeebrugge"

Dim mk, sh As Worksheet

For Each mk In mkt
    Set sh = ActiveWorkbook.Sheets.Add()
    sh.Name = Left(mk, 31)
    Helpers.CreateODBCQuery _
        "DECLARE @AsofDate DATETIME=dbo.LastWeekDay(getdate()-7) exec GetSeriesValue @AsofDateFrom=@AsofDate, @AsofGranularityDay=1, @PeriodGranularityDay=1, @orderby='2,1 DESC',    @csvtags ='PROVIDER:MDE,SRC:HEREN,OBTYPE:MID,MKT:" & mk & "'" _
        , CStr(mk)
Next mk
End Sub


Function EncodeExcelOnMacro(ByVal Name, ParamArray args())
EncodeExcelOnMacro = Name
    If IsNull(args) Then Exit Function
    If Not IsArray(args) Then Err.Raise 5, , "Arguments must be null or an array"
    Dim ArgCount: ArgCount = UBound(args) - LBound(args) + 1
    If ArgCount = 0 Then Exit Function
    Dim Index As Integer
    For Index = LBound(args) To UBound(args)
        If VarType(args(Index)) = vbString Then args(Index) = Chr(34) & args(Index) & Chr(34)
    Next
    EncodeExcelOnMacro = "'" & Name & " " & VBA.Join(args, ", ") & "'"

End Function
