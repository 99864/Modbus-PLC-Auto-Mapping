Attribute VB_Name = "模块1"
Sub 终极一键覆盖并生成完美CSV()
    Dim wsRaw As Worksheet, wsOut As Worksheet
    Dim wbTemp As Workbook
    Dim fileToOpen As Variant, fileToSave As Variant
    Dim fso As Object, ts As Object
    Dim i As Long, lastRow As Long
    Dim csvLine As String
    Dim tempArr As Variant
    Dim val1 As String, val2 As String, val3 As String, val4 As String, val5 As String
    Dim val6 As String, val7 As String, val8 As String, val9 As String
    
    ' 1. 绑定内部工作表
    On Error Resume Next
    Set wsRaw = ThisWorkbook.Sheets("原始数据")
    Set wsOut = ThisWorkbook.Sheets("汇川表_new")
    On Error GoTo 0
    
    If wsRaw Is Nothing Or wsOut Is Nothing Then
        MsgBox "找不到指定的工作表！请检查 '原始数据' 和 '汇川表_new' 页签。"
        Exit Sub
    End If
    
    ' ================= 第一步：导入外部数据覆盖 =================
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "第一步：请选择原始的乱码 CSV 文件"
        .Filters.Clear
        .Filters.Add "表格文件", "*.csv; *.txt"
        If .Show = -1 Then
            fileToOpen = .SelectedItems(1)
        Else
            Exit Sub
        End If
    End With
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    Set wbTemp = Workbooks.Open(fileToOpen)
    tempArr = wbTemp.Sheets(1).UsedRange.Value
    wbTemp.Close False
    
    wsRaw.Cells.Clear
    If IsArray(tempArr) Then
        wsRaw.Range("A1").Resize(UBound(tempArr, 1), UBound(tempArr, 2)).Value = tempArr
    End If
    
    Application.CalculateFull
    Application.DisplayAlerts = True
    
    ' ================= 第二步：生成全新CSV =================
    fileToSave = Application.GetSaveAsFilename( _
        InitialFileName:="完美导入_" & Format(Now, "hhmmss") & ".csv", _
        fileFilter:="CSV 文件 (*.csv), *.csv", _
        Title:="第二步：保存新文件")
        
    If fileToSave = False Then
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile(fileToSave, True, False)
    
    ' 写入表头（逗号分隔）
    ts.WriteLine "序号,变量名,数据类型,隐藏初始值,初始值,掉电保持,网络公开,注释,软元件地址"
    
    lastRow = wsOut.Cells(wsOut.Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To lastRow
        ' 过滤空行和报错行
        If Trim(wsOut.Cells(i, 9).Text) = "" Or wsOut.Cells(i, 9).Text = "#REF!" Then GoTo NextRow
        
        ' 清理所有换行符
        val1 = Replace(Replace(wsOut.Cells(i, 1).Text, vbCr, ""), vbLf, "")
        val2 = Replace(Replace(wsOut.Cells(i, 2).Text, vbCr, ""), vbLf, "")
        val3 = Replace(Replace(wsOut.Cells(i, 3).Text, vbCr, ""), vbLf, "")
        val4 = Replace(Replace(wsOut.Cells(i, 4).Text, vbCr, ""), vbLf, "")
        val5 = Replace(Replace(wsOut.Cells(i, 5).Text, vbCr, ""), vbLf, "")
        val6 = Replace(Replace(wsOut.Cells(i, 6).Text, vbCr, ""), vbLf, "")
        val7 = Replace(Replace(wsOut.Cells(i, 7).Text, vbCr, ""), vbLf, "")
        val8 = Replace(Replace(wsOut.Cells(i, 8).Text, vbCr, ""), vbLf, "")
        val9 = Replace(Replace(wsOut.Cells(i, 9).Text, vbCr, ""), vbLf, "")
        
        ' 用逗号拼接
        csvLine = val1 & "," & val2 & "," & val3 & "," & val4 & "," & val5 & "," & val6 & "," & val7 & "," & val8 & "," & val9
                  
        ts.WriteLine csvLine
NextRow:
    Next i
    
    ts.Close
    Application.ScreenUpdating = True
    MsgBox "处理完成！既包含 X/Y 也包含 D 区，直接导入汇川！", vbInformation
End Sub

