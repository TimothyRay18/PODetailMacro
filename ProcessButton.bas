Attribute VB_Name = "ProcessButton"
Function getMaxRow(col As Integer) As Double
    getMaxRow = ActiveSheet.Cells(Rows.Count, col).End(xlUp).row
End Function

Function getMaxCol(row As Integer) As Double
    getMaxCol = ActiveSheet.Cells(row, Columns.Count).End(xlToLeft).Column
End Function

Function getFilenameFromPath(ByVal strPath As String) As String
    If Right$(strPath, 1) <> "\" And Len(strPath) > 0 Then
        getFilenameFromPath = getFilenameFromPath(Left$(strPath, Len(strPath) - 1)) + Right$(strPath, 1)
    End If
End Function

Function findCellInColumn(row As Integer, str As String) As Double
    Dim i As Double
    i = 1
    Dim m As Double
    m = getMaxCol(row)
    While LCase(ActiveSheet.Cells(row, i).Value) <> LCase(str) And i <= m
        i = i + 1
    Wend
    findCellInColumn = i
End Function
Sub addBorder(start As String, last As String)
    Range(start & ":" & last).Select
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
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
End Sub

Sub Process()
    CleanPODetail
    SaveAsPO
    CopyPOToTemplate
    FinishPODetail
End Sub

Sub CleanPODetail()
'   Open file
    Dim dir1 As String
    dir1 = Range("B1").Value
    Workbooks.OpenText Filename:= _
        dir1, Origin:= _
        xlWindows, StartRow:=1, DataType:=xlDelimited, TextQualifier:= _
        xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, _
        Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), _
        Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), Array(7, 1), Array(8, 2), _
        Array(9, 1), Array(10, 1), Array(11, 1), Array(12, 1), Array(13, 1), Array(14, 2), Array(15 _
        , 1), Array(16, 1), Array(17, 1), Array(18, 1), Array(19, 1), Array(20, 1), Array(21, 1), _
        Array(22, 1), Array(23, 4), Array(24, 4), Array(25, 4), Array(26, 4), Array(27, 4), Array( _
        28, 1), Array(29, 1), Array(30, 1), Array(31, 1), Array(32, 1), Array(33, 1), Array(34, 1), _
        Array(35, 1), Array(36, 1), Array(37, 1), Array(38, 1), Array(39, 1), Array(40, 4)), _
        TrailingMinusNumbers:=True
'    Clean
    Rows("1:21").Select
    Selection.Delete Shift:=xlUp
    Rows("2:2").Select
    Selection.Delete Shift:=xlUp
    Columns("A:A").Select
    Selection.Delete Shift:=xlToLeft
    Columns("C:F").Select
    Selection.Delete Shift:=xlToLeft
    Columns("E:E").Select
    Selection.Delete Shift:=xlToLeft
    Columns("F:G").Select
    Selection.Delete Shift:=xlToLeft
    Columns("AC:AD").Select
    Selection.Delete Shift:=xlToLeft
    
    Dim max_row As Double
    max_row = getMaxRow(1)
    Dim max_col As Double
    max_col = getMaxCol(1)
    
    Range(Columns(1), Columns(max_col)).EntireColumn.AutoFit
    
    Range(Cells(1, 1), Cells(max_row, max_col)).Select
    Selection.AutoFilter
    
'   Clean MRP Control
    ActiveSheet.Range(Cells(1, 1), Cells(max_row, max_col)).AutoFilter Field:=findCellInColumn(1, "MRP contro"), Criteria1:= _
        "=MRP contro", Operator:=xlOr, Criteria2:="="
    Rows("2:" + CStr(max_row)).Select
    Selection.Delete Shift:=xlUp
    ActiveSheet.Range(Cells(1, 1), Cells(max_row, max_col)).AutoFilter Field:=findCellInColumn(1, "MRP contro")
    
'   Clean Last AB
    max_row = getMaxRow(1)
    ActiveSheet.Range(Cells(1, 1), Cells(max_row, max_col)).AutoFilter Field:=findCellInColumn(1, "Last AB Dt"), Criteria1:= _
        "=.  .", Operator:=xlOr, Criteria2:="=00.00.0000"
    Range("R2:R" + CStr(max_row)).Select
    Selection.ClearContents
    ActiveSheet.Range("$A$1:$AF$" + CStr(max_row)).AutoFilter Field:=findCellInColumn(1, "Last AB Dt")

'   Clean Last AL
    ActiveSheet.Range(Cells(1, 1), Cells(max_row, max_col)).AutoFilter Field:=findCellInColumn(1, "Last LA Dt"), Criteria1:="00.00.0000"
    Range("S2:S" + CStr(max_row)).Select
    Selection.ClearContents
    ActiveSheet.Range("$A$1:$AF$" + CStr(max_row)).AutoFilter Field:=findCellInColumn(1, "Last LA Dt")
    
    Range(Cells(1, 1), Cells(max_row, max_col)).Select
    Selection.AutoFilter
End Sub

Sub SaveAsPO()
    Workbooks("PO Detail Macro.xlsm").Activate
    Dim dir As String
    dir = Workbooks(getFilenameFromPath(Range("B1").Value)).Path
   ' MsgBox dir
    Dim dir1 As String
    dir1 = Range("B5").Value
    Workbooks.Open Filename:=dir1
    
    Dim name As String
    name = "Y" + CStr(Format(Now(), "yyyy")) + "-" + CStr(Format(Now(), "mm")) + "-" + CStr(Format(Now(), "dd")) + " PO DETAIL.xlsx"
    
    Dim dir2 As String
    
    dir2 = dir + "\" + name
    ActiveWorkbook.SaveAs Filename:= _
        dir2, FileFormat _
        :=xlOpenXMLWorkbook, CreateBackup:=False
        
    Workbooks.Open Filename:=dir1
End Sub

Sub CopyPOToTemplate()
    Workbooks("PO Detail Macro.xlsm").Activate
    Dim po As String
    po = getFilenameFromPath(Range("B1").Value)
    Dim poDetail As String
    poDetail = "Y" + CStr(Format(Now(), "yyyy")) + "-" + CStr(Format(Now(), "mm")) + "-" + CStr(Format(Now(), "dd")) + " PO DETAIL.xlsx"
    
    Workbooks(po).Activate
    Dim mr_po As String
    Dim mc_po As String
    mr_po = getMaxRow(1)
    mc_po = getMaxCol(1)
    
    Range(Cells(2, 1), Cells(mr_po, findCellInColumn(1, "Ord.UOM"))).Select
    Selection.Copy
    Workbooks(poDetail).Activate
    Dim mr_pd As String
    Dim mc_pd As String
    mr_pd = getMaxRow(1)
    mc_pd = getMaxCol(2)
    Range("A3").Select
    ActiveSheet.Paste
    
    Range(Cells(3, findCellInColumn(2, "Final New date")), Cells(3, findCellInColumn(2, "Forbidden List"))).Select
    Selection.AutoFill Destination:=Range(Cells(3, findCellInColumn(2, "Final New date")), Cells(getMaxRow(1), findCellInColumn(2, "Forbidden List")))
    
    If mr_po < mr_pd Then
        Rows(CStr(mr_po + 2) + ":" + mr_pd).Select
        Selection.Delete Shift:=xlUp
    End If
End Sub

Sub FinishPODetail()
    Workbooks("PO Detail Macro.xlsm").Activate
    Dim po As String
    po = getFilenameFromPath(Range("B1").Value)
    Dim poDetailOld As String
    poDetailOld = getFilenameFromPath(Range("B5").Value)
    Dim poDetailNew As String
    poDetailNew = "Y" + CStr(Format(Now(), "yyyy")) + "-" + CStr(Format(Now(), "mm")) + "-" + CStr(Format(Now(), "dd")) + " PO DETAIL.xlsx"
    Dim forbidden As String
    forbidden = getFilenameFromPath(Range("B9").Value)
    Dim jit As String
    jit = getFilenameFromPath(Range("B13").Value)
    
    Dim fb_path As String
    Dim jit_path As String
    fb_path = Range("B9").Value
    jit_path = Range("B13").Value
    
    Workbooks.Open Filename:=fb_path
    Workbooks.Open Filename:=jit_path
    
    Dim jit_sh As String
    Dim sh As Worksheet
    Application.DisplayAlerts = False
    For Each sh In Worksheets
        If LCase(Left(sh.name, 3)) = "jit" Then jit_sh = sh.name
    Next sh
    Application.DisplayAlerts = True
    
    
    Workbooks(poDetailNew).Activate
    
    Dim max_row As Double
    max_row = getMaxRow(1)
    Dim pic As Double
    pic = findCellInColumn(2, "PIC")
    m = findCellInColumn(2, "Material")
    Cells(3, pic).Select
    ActiveCell.FormulaR1C1 = _
        "=IFNA(VLOOKUP(RC[-" + CStr(pic - m) + "],'[PO Detail Macro.xlsm]Master Data'!C1:C2,2,0),VLOOKUP(RC[-" + CStr(pic - m) + "],'[" + poDetailOld + "]PO Detail'!C" + CStr(m) + ":C" + CStr(pic) + "," + CStr(pic - m + 1) + ",0))"
    Selection.AutoFill Destination:=Range(Cells(3, pic), Cells(max_row, pic))
    Range(Cells(3, pic), Cells(max_row, pic)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range(Cells(2, 1), Cells(max_row, getMaxCol(2))).Select
    Selection.AutoFilter
    
    Workbooks(po).Activate
    max_row = getMaxRow(1)
    Range(Cells(2, findCellInColumn(1, "Exception")), Cells(max_row, findCellInColumn(1, "Exception") + 1)).Select
    Selection.Copy
    
    Workbooks(poDetailNew).Activate
    Cells(3, findCellInColumn(2, "Exception")).Select
    ActiveSheet.Paste
    
    Dim pd As Double
    pd = findCellInColumn(2, "Purch. doc")
    Dim it As Double
    it = findCellInColumn(2, "Item")
    Dim poIt As Double
    poIt = findCellInColumn(2, "PO Item")

    Cells(3, poIt).Select
    ActiveCell.FormulaR1C1 = "=RC[-" + CStr(poIt - pd) + "]&RC[-" + CStr(poIt - it) + "]"
    Selection.AutoFill Destination:=Range(Cells(3, poIt), Cells(getMaxRow(2), poIt))
    Range(Cells(3, poIt), Cells(getMaxRow(2), poIt)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.TextToColumns Destination:=Cells(3, poIt), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True

    Dim pcs As Double
    pcs = findCellInColumn(2, "Price/pcs (USD)")
    m = findCellInColumn(2, "Material")
    max_row = getMaxRow(1)
    Cells(3, pcs).Select
    Application.CutCopyMode = False
    
 '   Dim wb As String
  '  wb = "JIT " + CStr(Format(Now(), "dd")) + " " + UCase(Format(Date, "mmmm")) + " " + CStr(Format(Now(), "yy"))
  '  ActiveCell.FormulaR1C1 = _
        "=IFNA(VLOOKUP(RC[-" + CStr(pcs - m) + "],'[" + jit + "]" + jit_sh + "'!C3:C47,45,0),VLOOKUP(RC[-" + CStr(pcs - m) + "],'[" + poDetailOld + "]PO Detail'!C3:C38,36,0))"
    ActiveCell.FormulaR1C1 = _
        "=IFNA(IF(VLOOKUP(RC[-" + CStr(pcs - m) + "],'[" + jit + "]" + jit_sh + "'!C3:C47,45,0)=0,VLOOKUP(RC[-" + CStr(pcs - m) + "],'[" + poDetailOld + "]PO Detail'!C3:C38,36,0),VLOOKUP(RC[-" + CStr(pcs - m) + "],'[" + jit + "]" + jit_sh + "'!C3:C47,45,0)),VLOOKUP(RC[-" + CStr(pcs - m) + "],'[" + poDetailOld + "]PO Detail'!C3:C38,36,0))"
    Selection.AutoFill Destination:=Range(Cells(3, pcs), Cells(getMaxRow(1), pcs))
    Range(Cells(3, pcs), Cells(getMaxRow(1), pcs)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Dim sc As Double
    sc = findCellInColumn(2, "Status Consigment")
    Cells(3, sc).Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = _
        "=IF(LEFT(RC[-" + CStr(sc - m) + "],2)=""OP"",""OP"",IF(RC[-" + CStr(sc - findCellInColumn(2, "      Net price")) + "]="""",""K"",""Non K""))"
    Cells(3, sc).Select
    Selection.AutoFill Destination:=Range(Cells(3, sc), Cells(max_row, sc))
    Range(Cells(3, sc), Cells(max_row, sc)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Dim fl As Double
    fl = findCellInColumn(2, "Forbidden List")
    Cells(3, fl).Select
    ActiveCell.FormulaR1C1 = _
        "=IFNA(VLOOKUP(RC[-" + CStr(fl - m) + "],'[" + forbidden + "]Sheet1'!C1:C3,3,0),"""")"
    Selection.AutoFill Destination:=Range(Cells(3, fl), Cells(max_row, fl))
    Range(Cells(3, fl), Cells(max_row, fl)).Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    Range("AM1").Select
    Cells(1, findCellInColumn(2, "Net Value (USD)")).Select
    ActiveCell.FormulaR1C1 = "=SUBTOTAL(9,R[2]C:R[" + CStr(max_row - 1) + "]C)"
    
    addBorder "A2", Cells(max_row, fl).Address
    
    Application.DisplayAlerts = False
    Windows(forbidden).Close
    Windows(po).Close
    Windows(poDetailOld).Close
    Windows(jit).Close
    Application.DisplayAlerts = True
    
    Dim pt As PivotTable
    Dim source As String
    source = "PO Detail!R2C1:R" + CStr(max_row) + "C37"
    
    For Each sh In Worksheets
        For Each pt In sh.PivotTables
            sh.PivotTables(pt.name).ChangePivotCache ActiveWorkbook. _
                PivotCaches.Create(SourceType:=xlDatabase, SourceData:=source, Version:=6)
            sh.PivotTables(pt.name).PivotCache.Refresh
        Next pt
    Next sh
End Sub


