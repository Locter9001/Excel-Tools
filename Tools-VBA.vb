Dim Line As Long '选择的商品的行数
Dim This_price As Double '单价
Dim This_profits As Double '利润
Dim This_brokerageOne As Double '一级佣金
Dim This_brokerageTwo As Double '二级佣金
Dim This_logisticsCost As Double '物流费
Dim This_courierCost As Double  '快递费
Dim TheNum As Integer '数量 用于进行乘法运算
Dim TorF As Byte '快递选择判断符 1: 平台配送 2: 快递配送 3: 未选择
Dim DateLine As Integer '选择的日期行数
Dim BoxLine As Long '选定区域的行数

'日期选择器:
Private Sub DateList_Enter()
    Dim lastRow As Long
    Dim index As Long
    lastRow = Worksheets(2).Columns(2).Find("*", searchdirection:=xlPrevious).Row
    For index = 1 To lastRow
        If Worksheets(2).Cells(index, 8).Value = "日期" Then
            DateList.AddItem Worksheets(2).Cells(index, 9).Value
        End If
    Next index
    DateList.ColumnCount = 1
    DateList.ColumnHeads = False
End Sub

'选择日期
Private Sub DateList_Change()
    If DateList.ListIndex <> -1 Then
        Dim Date_Line As Range
        Set Date_Line = Worksheets(2).Cells.Find(DateValue(DateList.Value))
        If Date_Line Is Nothing Then
            MsgBox "失败"
        Else
            DateLine = Date_Line.Row
            BoxLine = [Date_Line].CurrentRegion.Rows.Count
            [Date_Line].CurrentRegion.Select
        End If
    End If
End Sub


'商品选择器:
Private Sub GoodsList_Enter()
    TorF = 3 '默认为未选择配送方式
    Dim lastMsg As Long
    Dim index As Long
    lastMsg = Worksheets(1).Columns(2).Find("*", searchdirection:=xlPrevious).Row
    For index = 5 To lastMsg
        GoodsList.AddItem Worksheets(1).Cells(index, 2).Value
    Next index
    GoodsList.ColumnCount = 1
    GoodsList.ColumnHeads = False
End Sub

Private Sub GoodsList_Change()
'商品窗体数值变化
If GoodsList.ListIndex <> -1 Then
    Line = Worksheets(1).UsedRange.Find(GoodsList.Value).Row
    goodsNum_ipt = 1
    This_price = Worksheets(1).Cells(Line, 5).Value
    price = This_price
    This_profits = Worksheets(1).Cells(Line, 6).Value
    profits = This_profits
    This_brokerageOne = Worksheets(1).Cells(Line, 8).Value
    brokerageOne = This_brokerageOne
    This_brokerageTwo = Worksheets(1).Cells(Line, 9).Value
    brokerageTwo = This_brokerageTwo
    This_logisticsCost = Worksheets(1).Cells(Line, 10).Value
    This_courierCost = Worksheets(1).Cells(Line, 11).Value
End If
End Sub

Private Sub goodsNum_ipt_Change()
    If goodsNum_ipt.Value <> "" Then
        TheNum = goodsNum_ipt.Value
        brokerageOne = This_brokerageOne * goodsNum_ipt.Value
        brokerageTwo = This_brokerageTwo * goodsNum_ipt.Value
        profits = This_profits * goodsNum_ipt.Value
        price = This_price * goodsNum_ipt.Value
        If TorF = 1 Then
            logisticsCost = This_logisticsCost * TheNum
        ElseIf TorF = 2 Then
            logisticsCost = This_courierCost * TheNum
        End If
    End If
    If goodsNum_ipt = "" Then
        brokerageOne = 0
        brokerageTwo = 0
        profits = 0
        price = 0
        logisticsCost = 0
    End If
    
End Sub

Private Sub goodsNum_ipt_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    '控制文本框只能输入数字

    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
    
        KeyAscii = 0

    End If

End Sub

Private Sub Label8_Click()

End Sub

Private Sub OptionButton1_Click()
    TorF = 1
    logisticsCost = This_logisticsCost * TheNum
End Sub

Private Sub OptionButton2_Click()
    TorF = 2
    logisticsCost = This_courierCost * TheNum
End Sub

Private Sub price_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    '控制文本框只能输入数字

    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then

        KeyAscii = 0

    End If

End Sub

Private Sub profits_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    '控制文本框只能输入数字

    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then

        KeyAscii = 0

    End If

End Sub

Private Sub brokerageOne_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    '控制文本框只能输入数字

    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then

        KeyAscii = 0

    End If

End Sub

Private Sub brokerageTwo_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

    '控制文本框只能输入数字

    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then

        KeyAscii = 0

    End If

End Sub

Private Sub submitBtn_Click()
    If brokerageTwo <> "" And brokerageOne <> "" And profits <> "" And price <> "" And TorF <> 3 And goodsNum_ipt <> "" And GoodsList.ListIndex <> -1 And DateList.ListIndex <> -1 Then
    '在第几行增加行
    Dim Line As Long
    '选定区域内有多少个商品行
    Dim Box_Line As Long
    '5是标题和合计行的总计 BoxLine 是选定区域的行数
    Box_Line = BoxLine - 5
    '运算在第几行增加新的商品明细，运算规则：日期的行位置 + 副标题和标题栏行数 + 商品占行数
    Line = DateLine + 2 + Box_Line
    '新增行
    Rows(Line).Insert
    Worksheets(2).Range(Cells(Line,1),Cells(Line,10)).clearformats
    Worksheets(2).Range(Cells(Line,1),Cells(Line,10)).HorizontalAlignment = Excel.xlcenter
    '给单元格赋值 => 序号
    Dim serialNum As Long
    if Cells((Line - 1), 1).Value = "序号" then
        serialNum = 1
    Else
        serialNum = Cells((Line - 1), 1).Value + 1
    End if

    Cells(Line, 1).Value = serialNum
    
    

    '商品名称
    Cells(Line, 2).Value = GoodsList.Value
    '商品数量
    Cells(Line, 3).Value = TheNum
    '商品单价
    Cells(Line, 4).Value = This_price
    '收益
    Cells(Line, 5).Value = This_profits * TheNum
    '一级分销
    Cells(Line, 6).Value = This_brokerageOne * TheNum
    '二级分销
    Cells(Line, 7).Value = This_brokerageTwo * TheNum

    If TorF <> 3 Then
        If TorF = 1 Then
            '物流配送
            Cells(Line, 8).Value = This_logisticsCost * TheNum
            '净利润
            Cells(Line, 10).Value = (This_profits - This_brokerageOne - This_brokerageTwo - This_logisticsCost) * TheNum
        ElseIf TorF = 2 Then
            '快递费用
            Cells(Line, 9).Value = This_courierCost * TheNum
            '净利润
            Cells(Line, 10).Value = (This_profits - This_brokerageOne - This_brokerageTwo - This_courierCost) * TheNum
        End If
    End If
    Cells((Line + 1), 3).Value = Application.WorksheetFunction.Sum(Range(Cells(DateLine + 2, 3), Cells(Line, 3)))
    Cells((Line + 1), 4).Value = Application.WorksheetFunction.Sum(Range(Cells(DateLine + 2, 4), Cells(Line, 4)))
    Cells((Line + 1), 5).Value = Application.WorksheetFunction.Sum(Range(Cells(DateLine + 2, 5), Cells(Line, 5)))
    Cells((Line + 1), 6).Value = Application.WorksheetFunction.Sum(Range(Cells(DateLine + 2, 6), Cells(Line, 6)))
    Cells((Line + 1), 7).Value = Application.WorksheetFunction.Sum(Range(Cells(DateLine + 2, 7), Cells(Line, 7)))
    Cells((Line + 1), 8).Value = Application.WorksheetFunction.Sum(Range(Cells(DateLine + 2, 8), Cells(Line, 8)))
    Cells((Line + 1), 9).Value = Application.WorksheetFunction.Sum(Range(Cells(DateLine + 2, 9), Cells(Line, 9)))
    Cells((Line + 1), 10).Value = Application.WorksheetFunction.Sum(Range(Cells(DateLine + 2, 10), Cells(Line, 10)))
    Label8.Visible = False
    Else
    Label8.Visible = True
    End If
End Sub


