Attribute VB_Name = "modRibbon"
' ============================================================================
'  Excel 小工具箱  --  Ribbon (customUI) 回调模块
'  与 src/ribbon/customUI.xml 中的 onAction / onLoad 一一对应。
'  回调签名必须是 Public Sub，且参数类型严格匹配，否则功能区加载失败。
' ============================================================================
Option Explicit

Private mRibbon As IRibbonUI

' customUI/@onLoad：缓存功能区对象，供后续 Invalidate 使用
Public Sub rxOnLoad(ByVal ribbon As IRibbonUI)
    On Error Resume Next
    Set mRibbon = ribbon
    On Error GoTo 0
End Sub

' 手动刷新功能区（当前版本按钮恒可用，保留给后续扩展）
Public Sub RefreshRibbon()
    On Error Resume Next
    If Not mRibbon Is Nothing Then mRibbon.Invalidate
    On Error GoTo 0
End Sub

' ---------------------------------------------------------------------------
' 三个功能按钮 + 关于按钮
' 每个回调单独兜底错误，避免单个宏异常导致功能区整体失效
' ---------------------------------------------------------------------------
Public Sub rxDeleteEmptyRows(ByVal control As IRibbonControl)
    On Error GoTo EH
    Toolbox_DeleteEmptyRows
    Exit Sub
EH:
    MsgBox "【删除空行】运行失败：" & Err.Description, vbCritical, TOOLBOX_TITLE
End Sub

Public Sub rxRemoveDuplicates(ByVal control As IRibbonControl)
    On Error GoTo EH
    Toolbox_RemoveDuplicates
    Exit Sub
EH:
    MsgBox "【一键去重】运行失败：" & Err.Description, vbCritical, TOOLBOX_TITLE
End Sub

Public Sub rxBeautifyHeader(ByVal control As IRibbonControl)
    On Error GoTo EH
    Toolbox_BeautifyHeader
    Exit Sub
EH:
    MsgBox "【美化表头】运行失败：" & Err.Description, vbCritical, TOOLBOX_TITLE
End Sub

Public Sub rxAbout(ByVal control As IRibbonControl)
    On Error GoTo EH
    Toolbox_About
    Exit Sub
EH:
    MsgBox "运行失败：" & Err.Description, vbCritical, TOOLBOX_TITLE
End Sub
