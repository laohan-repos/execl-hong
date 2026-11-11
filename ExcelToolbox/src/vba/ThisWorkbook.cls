VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerivation = False
' ============================================================================
'  Excel 小工具箱  --  加载项工作簿事件代码
'  构建脚本会读取本文件、剔除所有 Attribute/VERSION/BEGIN/END 头部行后，
'  用 VBComponents("ThisWorkbook").CodeModule.AddFromString 写入 XLAM。
'  这样用户完全不需要打开 VBA 编辑器。
' ============================================================================
Option Explicit

Private Sub Workbook_Open()
    On Error Resume Next
    ThisWorkbook.IsAddin = True
    RefreshRibbon
    On Error GoTo 0
End Sub

Private Sub Workbook_AddinInstall()
    On Error Resume Next
    RefreshRibbon
    On Error GoTo 0
End Sub

Private Sub Workbook_AddinUninstall()
    ' 卸载时不做任何读写，避免影响用户数据
End Sub
