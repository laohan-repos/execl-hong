Attribute VB_Name = "modToolbox"
' ============================================================================
'  Excel 小工具箱 v1.0.0  --  核心功能模块 (modToolbox)
' ----------------------------------------------------------------------------
'  设计约束
'   1. 只操作 ActiveWorkbook（用户当前打开的工作簿），绝不修改 ThisWorkbook(.xlam)。
'   2. 删除空行 / 一键去重属于数据损失型操作：
'        确认 -> 整表备份到新工作表 -> 执行 -> 报告备份位置。
'      Excel VBA 的删除动作不进入撤销栈，所以必须自己留可恢复副本。
'   3. 每个入口都有 Silent 参数：安装后的自动化回归测试用它跳过所有对话框。
'   4. 不使用 Declare / ActiveX 控件 / 外部引用，32 位与 64 位 Office 通用。
'   5. 全部过程为 Public 且无参数依赖，便于 Application.Run 调用与测试。
' ============================================================================
Option Explicit

Public Const TOOLBOX_VERSION As String = "1.0.0"
Public Const TOOLBOX_TITLE   As String = "Excel 小工具箱 v1.0.0"
Public Const BACKUP_PREFIX   As String = "备份_"

' ---------------------------------------------------------------------------
' 1) 删除空行：删除当前工作表中「整行既无值也无公式」的行
' ---------------------------------------------------------------------------
Public Sub Toolbox_DeleteEmptyRows(Optional ByVal Silent As Boolean = False)
    Const PROC As String = "删除空行"
    Dim wb As Workbook, ws As Worksheet, ur As Range, bak As Worksheet
    Dim firstRow As Long, lastRow As Long, firstCol As Long, lastCol As Long
    Dim r As Long, nEmpty As Long, nDeleted As Long
    Dim oldCalc As XlCalculation, appTouched As Boolean
    Dim oldEnableEvents As Boolean, oldScreenUpdating As Boolean
    Dim errText As String, bakName As String

    On Error GoTo EH
    If Not GetTargetWorkbook(wb, Silent) Then Exit Sub
    If Not GetTargetSheet(ws, Silent) Then Exit Sub
    Set ur = ws.UsedRange

    firstRow = ur.Row
    lastRow = ur.Row + ur.Rows.Count - 1
    firstCol = ur.Column
    lastCol = ur.Column + ur.Columns.Count - 1

    For r = firstRow To lastRow
        If RowIsEmpty(ws, r, firstCol, lastCol) Then nEmpty = nEmpty + 1
    Next r

    If nEmpty = 0 Then
        If Not Silent Then MsgBox "当前工作表【" & ws.Name & "】没有完全为空的行，无需处理。", _
                                  vbInformation, TOOLBOX_TITLE
        Exit Sub
    End If

    If Not Silent Then
        If MsgBox("即将删除工作表【" & ws.Name & "】中 " & nEmpty & " 个完全为空的行。" & vbCrLf & vbCrLf & _
                  "· 判定标准：整行既没有值，也没有公式" & vbCrLf & _
                  "· 执行前会自动新建一张【备份】工作表，保存改动前的完整数据" & vbCrLf & _
                  "· 如需恢复：从备份工作表把数据复制回来即可" & vbCrLf & vbCrLf & _
                  "是否继续？", vbYesNo + vbExclamation + vbDefaultButton2, TOOLBOX_TITLE) <> vbYes Then Exit Sub
    End If

    ' 备份也会切换活动表；先保存状态并关闭事件，避免触发其它工作簿宏。
    oldCalc = Application.Calculation
    oldEnableEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    appTouched = True
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Set bak = CreateBackup(wb, ws, PROC)
    If bak Is Nothing Then Err.Raise vbObjectError + 1001, , _
        "无法创建备份工作表（工作簿可能被保护、共享或只读），已放弃操作，未改动任何数据。"
    bakName = bak.Name
    ws.Activate                       ' 备份副本会成为活动表，切回用户原来的表

    For r = lastRow To firstRow Step -1
        If RowIsEmpty(ws, r, firstCol, lastCol) Then
            ws.Rows(r).Delete Shift:=xlUp
            nDeleted = nDeleted + 1
        End If
    Next r

FINALLY:
    On Error Resume Next
    If appTouched Then
        Application.Calculation = oldCalc
        Application.EnableEvents = oldEnableEvents
        Application.ScreenUpdating = oldScreenUpdating
    End If
    On Error GoTo 0
    If Len(errText) > 0 Then
        If Not Silent Then MsgBox "【" & PROC & "】执行出错：" & errText & vbCrLf & vbCrLf & _
            "备份工作表：" & IIf(Len(bakName) > 0, bakName, "(未创建)") & vbCrLf & _
            "如有备份，原始数据仍可从中恢复。", vbCritical, TOOLBOX_TITLE
        Exit Sub
    End If
    If Not Silent Then MsgBox "完成：已删除 " & nDeleted & " 个空行。" & vbCrLf & _
        "改动前的数据已保存在工作表【" & bakName & "】。", vbInformation, TOOLBOX_TITLE
    Exit Sub
EH:
    errText = Err.Description
    Resume FINALLY
End Sub

' ---------------------------------------------------------------------------
' 2) 一键去重：按选中区域所有列组合判重，重复行只保留首次出现的那条
'    交互模式会询问首行是否为表头；Silent 模式默认「全部为数据行」。
' ---------------------------------------------------------------------------
Public Sub Toolbox_RemoveDuplicates(Optional ByVal Silent As Boolean = False)
    Const PROC As String = "一键去重"
    Dim wb As Workbook, ws As Worksheet, rng As Range, bak As Worksheet
    Dim hasHeader As Boolean, resp As VbMsgBoxResult
    Dim c1 As Long, c2 As Long, rStart As Long, rEnd As Long, r As Long, i As Long
    Dim dict As Object, dupRows As Collection, key As String
    Dim nDup As Long, bakName As String
    Dim oldCalc As XlCalculation, appTouched As Boolean
    Dim oldEnableEvents As Boolean, oldScreenUpdating As Boolean
    Dim errText As String

    On Error GoTo EH
    If Not GetTargetWorkbook(wb, Silent) Then Exit Sub
    If Not GetTargetSheet(ws, Silent) Then Exit Sub

    If TypeName(Selection) <> "Range" Then
        If Not Silent Then MsgBox "请先选中一个单元格区域。", vbExclamation, TOOLBOX_TITLE
        Exit Sub
    End If
    Set rng = Selection
    If rng.Areas.Count > 1 Then
        If Not Silent Then MsgBox "请只选择一个连续的矩形区域（不支持多重选区）。", vbExclamation, TOOLBOX_TITLE
        Exit Sub
    End If
    If rng.Rows.Count < 2 Then
        If Not Silent Then MsgBox "选中区域至少需要 2 行数据。", vbExclamation, TOOLBOX_TITLE
        Exit Sub
    End If

    hasHeader = False
    If Not Silent Then
        resp = MsgBox("选中区域的第一行是【表头】吗？" & vbCrLf & vbCrLf & _
              "· 选「是」：第一行作为表头，不参与去重" & vbCrLf & _
              "· 选「否」：所有行都视为数据行" & vbCrLf & _
              "· 选「取消」：放弃本次操作" & vbCrLf & vbCrLf & _
              "规则：按选中区域的所有列组合判断，重复行只保留第一次出现的那条。", _
              vbYesNoCancel + vbQuestion, TOOLBOX_TITLE)
        If resp = vbCancel Then Exit Sub
        hasHeader = (resp = vbYes)
    End If

    c1 = rng.Column
    c2 = rng.Column + rng.Columns.Count - 1
    rStart = rng.Row + IIf(hasHeader, 1, 0)
    rEnd = rng.Row + rng.Rows.Count - 1
    If rStart > rEnd Then
        If Not Silent Then MsgBox "选中区域除表头外没有数据行。", vbExclamation, TOOLBOX_TITLE
        Exit Sub
    End If

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare      ' 与 Excel 自带去重一致：不区分大小写
    Set dupRows = New Collection
    For r = rStart To rEnd
        key = RowKey(ws, r, c1, c2)
        If dict.Exists(key) Then
            dupRows.Add r
        Else
            dict.Add key, r
        End If
    Next r
    nDup = dupRows.Count

    If nDup = 0 Then
        If Not Silent Then MsgBox "未发现重复记录，无需处理。", vbInformation, TOOLBOX_TITLE
        Exit Sub
    End If

    If Not Silent Then
        If MsgBox("共发现 " & nDup & " 条重复记录（每组保留第一条）。" & vbCrLf & vbCrLf & _
              "· 将删除工作表中的整行（不只是选中区域），请先确认选区已覆盖所有要保留的列" & vbCrLf & _
              "· 执行前会自动新建一张【备份】工作表，保存改动前的完整数据" & vbCrLf & vbCrLf & _
              "是否继续？", vbYesNo + vbExclamation + vbDefaultButton2, TOOLBOX_TITLE) <> vbYes Then Exit Sub
    End If

    oldCalc = Application.Calculation
    oldEnableEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating
    appTouched = True
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    Set bak = CreateBackup(wb, ws, PROC)
    If bak Is Nothing Then Err.Raise vbObjectError + 1002, , _
        "无法创建备份工作表（工作簿可能被保护、共享或只读），已放弃操作，未改动任何数据。"
    bakName = bak.Name
    ws.Activate

    For i = dupRows.Count To 1 Step -1   ' 自下而上删除，行号不会错位
        ws.Rows(dupRows(i)).Delete Shift:=xlUp
    Next i

FINALLY:
    On Error Resume Next
    If appTouched Then
        Application.Calculation = oldCalc
        Application.EnableEvents = oldEnableEvents
        Application.ScreenUpdating = oldScreenUpdating
    End If
    On Error GoTo 0
    If Len(errText) > 0 Then
        If Not Silent Then MsgBox "【" & PROC & "】执行出错：" & errText & vbCrLf & vbCrLf & _
            "备份工作表：" & IIf(Len(bakName) > 0, bakName, "(未创建)"), vbCritical, TOOLBOX_TITLE
        Exit Sub
    End If
    If Not Silent Then MsgBox "完成：已删除 " & nDup & " 条重复记录，每组保留第一条。" & vbCrLf & _
        "改动前的数据已保存在工作表【" & bakName & "】。", vbInformation, TOOLBOX_TITLE
    Exit Sub
EH:
    errText = Err.Description
    Resume FINALLY
End Sub

' ---------------------------------------------------------------------------
' 3) 美化表头：选中区域首行 -> 蓝底 / 白色粗体 / 居中 / 自动列宽
' ---------------------------------------------------------------------------
Public Sub Toolbox_BeautifyHeader(Optional ByVal Silent As Boolean = False)
    Const PROC As String = "美化表头"
    Dim wb As Workbook, ws As Worksheet, rng As Range, hdr As Range, bak As Worksheet
    Dim bakName As String, errText As String

    On Error GoTo EH
    If Not GetTargetWorkbook(wb, Silent) Then Exit Sub
    If Not GetTargetSheet(ws, Silent) Then Exit Sub

    If TypeName(Selection) <> "Range" Then
        If Not Silent Then MsgBox "请先选中一个单元格区域。", vbExclamation, TOOLBOX_TITLE
        Exit Sub
    End If
    Set rng = Selection
    If rng.Areas.Count > 1 Then
        If Not Silent Then MsgBox "请只选择一个连续的矩形区域。", vbExclamation, TOOLBOX_TITLE
        Exit Sub
    End If
    Set hdr = rng.Rows(1)

    If Not Silent Then
        If MsgBox("将把选中区域的第一行（" & hdr.Address(False, False) & "）设置为：" & vbCrLf & _
              "· 蓝色背景 + 白色粗体 + 水平/垂直居中" & vbCrLf & _
              "· 自动调整选中区域涉及的列宽" & vbCrLf & vbCrLf & _
              "执行前会自动新建一张【备份】工作表。是否继续？", _
              vbYesNo + vbQuestion + vbDefaultButton2, TOOLBOX_TITLE) <> vbYes Then Exit Sub
    End If

    Set bak = CreateBackup(wb, ws, PROC)
    If bak Is Nothing Then Err.Raise vbObjectError + 1003, , _
        "无法创建备份工作表（工作簿可能被保护、共享或只读），已放弃操作。"
    bakName = bak.Name
    ws.Activate

    With hdr
        .Interior.Pattern = xlSolid
        .Interior.Color = RGB(0, 112, 192)          ' Office 蓝
        With .Font
            .Color = RGB(255, 255, 255)
            .Bold = True
        End With
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = RGB(0, 70, 140)
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With
    hdr.RowHeight = 24
    rng.EntireColumn.AutoFit

FINALLY:
    On Error GoTo 0
    If Len(errText) > 0 Then
        If Not Silent Then MsgBox "【" & PROC & "】执行出错：" & errText & vbCrLf & vbCrLf & _
            "备份工作表：" & IIf(Len(bakName) > 0, bakName, "(未创建)"), vbCritical, TOOLBOX_TITLE
        Exit Sub
    End If
    If Not Silent Then MsgBox "完成：表头已美化，列宽已自动调整。" & vbCrLf & _
        "改动前的数据已保存在工作表【" & bakName & "】。", vbInformation, TOOLBOX_TITLE
    Exit Sub
EH:
    errText = Err.Description
    Resume FINALLY
End Sub

' ---------------------------------------------------------------------------
' 关于 / 自检
' ---------------------------------------------------------------------------
Public Sub Toolbox_About()
    MsgBox "Excel 小工具箱  v" & TOOLBOX_VERSION & vbCrLf & _
           "————————————————————————————" & vbCrLf & _
           "功能：删除空行 / 一键去重 / 美化表头" & vbCrLf & vbCrLf & _
           "安全说明" & vbCrLf & _
           "· 所有操作只作用于你当前打开的工作簿，不修改加载项自身" & vbCrLf & _
           "· 删除/去重执行前会确认，并自动新建【备份_xxx】工作表" & vbCrLf & _
           "· 恢复数据：从备份工作表复制回原工作表即可" & vbCrLf & _
           "· 安装程序只把工具箱目录登记为 Excel 受信任位置，" & vbCrLf & _
           "  不会关闭或降低全局宏安全设置" & vbCrLf & vbCrLf & _
           "重要：Excel 的 VBA 操作不进入撤销栈，请勿依赖 Ctrl+Z 恢复。", _
           vbInformation, TOOLBOX_TITLE
End Sub

' 供 COM 自动化测试确认加载项已加载、宏已启用
Public Function Toolbox_Ping() As String
    Toolbox_Ping = TOOLBOX_VERSION
End Function

' ---------------------------------------------------------------------------
' 内部辅助
' ---------------------------------------------------------------------------

' 取得用户当前工作簿；排除加载项自身
Private Function GetTargetWorkbook(ByRef wb As Workbook, ByVal Silent As Boolean) As Boolean
    On Error GoTo EH
    Set wb = ActiveWorkbook
    If wb Is Nothing Then GoTo NoBook
    If wb Is ThisWorkbook Then GoTo NoBook
    If wb.IsAddin Then GoTo NoBook
    GetTargetWorkbook = True
    Exit Function
NoBook:
    If Not Silent Then MsgBox "请先打开（或激活）一个普通的 Excel 工作簿，再使用本工具。" & vbCrLf & _
        "说明：本工具只处理你当前打开的工作簿，不会修改工具箱加载项自身。", vbExclamation, TOOLBOX_TITLE
    GetTargetWorkbook = False
    Exit Function
EH:
    If Not Silent Then MsgBox "无法获取当前工作簿：" & Err.Description, vbCritical, TOOLBOX_TITLE
    GetTargetWorkbook = False
End Function

' 取得活动工作表；图表工作表会被排除
Private Function GetTargetSheet(ByRef ws As Worksheet, ByVal Silent As Boolean) As Boolean
    On Error Resume Next
    Set ws = ActiveSheet
    On Error GoTo 0
    If ws Is Nothing Then
        If Not Silent Then MsgBox "请切换到一张普通工作表后再执行。", vbExclamation, TOOLBOX_TITLE
        GetTargetSheet = False
    Else
        GetTargetSheet = True
    End If
End Function

' 完全为空 = 既无常量也无公式（仅有格式的单元格算空）
Private Function RowIsEmpty(ByVal ws As Worksheet, ByVal r As Long, _
                            ByVal c1 As Long, ByVal c2 As Long) As Boolean
    RowIsEmpty = (Application.WorksheetFunction.CountA( _
                    ws.Range(ws.Cells(r, c1), ws.Cells(r, c2))) = 0)
End Function

' 生成行指纹：长度前缀 + 值，避免分隔符碰撞；错误值用显示文本
Private Function RowKey(ByVal ws As Worksheet, ByVal r As Long, _
                        ByVal c1 As Long, ByVal c2 As Long) As String
    Dim buf As String, c As Long, s As String, v As Variant
    For c = c1 To c2
        v = ws.Cells(r, c).Value2
        If IsError(v) Then
            s = "#ERR:" & ws.Cells(r, c).Text
        ElseIf IsNull(v) Then
            s = ""
        Else
            s = CStr(v)
        End If
        buf = buf & Len(s) & ":" & s & "|"
    Next c
    RowKey = buf
End Function

' 把整张工作表复制到同一工作簿末尾作为可恢复备份
Private Function CreateBackup(ByVal wb As Workbook, ByVal ws As Worksheet, _
                              ByVal opName As String) As Worksheet
    On Error GoTo EH
    Dim target As Worksheet, sh As Worksheet, nm As String
    Dim beforeCount As Long
    If wb.Worksheets.Count = 0 Then Exit Function
    beforeCount = wb.Worksheets.Count
    Set target = wb.Worksheets(beforeCount)
    ws.Copy After:=target
    Set sh = wb.Worksheets(beforeCount + 1)
    nm = SafeSheetName(BACKUP_PREFIX & opName & "_" & Format(Now, "yymmdd_hhnnss"), wb)
    sh.Name = nm
    On Error Resume Next
    sh.Tab.Color = RGB(255, 192, 0)     ' 橙色标签，方便一眼找到备份
    On Error GoTo EH
    Set CreateBackup = sh
    Exit Function
EH:
    Set CreateBackup = Nothing
End Function

' Excel 工作表名：<=31 字符、不含 : \ / ? * [ ]、不可重名
Private Function SafeSheetName(ByVal base As String, ByVal wb As Workbook) As String
    Dim s As String, bad As Variant, v As Variant, i As Long, cand As String
    s = base
    bad = Array(":", "\", "/", "?", "*", "[", "]")
    For Each v In bad
        s = Replace(s, CStr(v), "_")
    Next v
    If Len(s) > 31 Then s = Left$(s, 31)
    If Not SheetExists(wb, s) Then SafeSheetName = s: Exit Function
    For i = 1 To 999
        cand = Left$(s, 31 - Len(CStr(i))) & CStr(i)
        If Not SheetExists(wb, cand) Then SafeSheetName = cand: Exit Function
    Next i
    SafeSheetName = Left$(s, 26) & Format(Timer * 100, "00000")
End Function

Private Function SheetExists(ByVal wb As Workbook, ByVal nm As String) As Boolean
    Dim sh As Object
    On Error Resume Next
    Set sh = wb.Sheets(nm)
    SheetExists = Not sh Is Nothing
    On Error GoTo 0
End Function
