' button
Sub 新增今日明细单()
    Dim index As Integer
    Dim lineNum As Long
    Dim myFont As Font
    Dim myRange As Range
    LineNum = 1
    For index = 1 To 7
        Rows(1).Insert
    Next index

    lastRow = Worksheets(2).Columns(2).Find("*", searchdirection:=xlPrevious).Row
    For index = 1 To lastRow
        If Worksheets(2).Cells(index, 8).Value = "日期" Then
            lineNum = lineNum + 1
        End If
    Next index
    Set myRange = Cells(1,1)
    Set myFont = myRange.Font
    myRange.Value = "鹊源商贸商品出售明细单"
    with myFont
    .Name = "等线"
    .size = 18
    .Bold = True
    End with
    Worksheets(2).Range("A1:J2").merge
    Worksheets(2).Cells(3,1).Value = "CS"
    Worksheets(2).Cells(3,1).HorizontalAlignment = Excel.xlRight
    Worksheets(2).Cells(3,2).Value = lineNum
    Worksheets(2).Cells(3,2).HorizontalAlignment = Excel.xlLeft
    Worksheets(2).Cells(3,8).Value = "日期"
    Worksheets(2).Cells(3,8).Font.size = 11
    Worksheets(2).Cells(3,9).NumberFormatLocal = "yyyy/m/d"
    Worksheets(2).Cells(3,9).Value = Now()
    Worksheets(2).Range("I3:J3").merge
    Worksheets(2).Cells(4,1).Value = "序号"
    Worksheets(2).Cells(4,2).Value = "商品名称"
    Worksheets(2).Cells(4,3).Value = "商品数量"
    Worksheets(2).Cells(4,4).Value = "商品单价"
    Worksheets(2).Cells(4,5).Value = "收益"
    Worksheets(2).Cells(4,6).Value = "一级分销"
    Worksheets(2).Cells(4,7).Value = "二级分销"
    Worksheets(2).Cells(4,8).Value = "物流费用"
    Worksheets(2).Cells(4,9).Value = "快递费"
    Worksheets(2).Cells(4,10).Value = "净利润"
    '第三行
    Worksheets(2).Cells(5,2).Value = "合计"
    For index = 1 To 10
        Worksheets(2).Cells(4,index).Font.size = 11
        Worksheets(2).Cells(4,index).Font.Bold = True
        Worksheets(2).Cells(4,index).Font.Color = RGB(0, 0, 0)
        Worksheets(2).Cells(4,index).Interior.Color = RGB(226, 226, 226)
        Worksheets(2).Cells(4,index).Borders.LineStyle = xlContinuous

        Worksheets(2).Cells(5,index).Font.Bold = True
        Worksheets(2).Cells(5,index).Font.Color = RGB(255, 135, 22)
        Worksheets(2).Cells(5,index).Borders.LineStyle = xlContinuous
        Worksheets(2).Cells(5,index).Interior.Color = RGB(226, 226, 226)
    Next index
End Sub