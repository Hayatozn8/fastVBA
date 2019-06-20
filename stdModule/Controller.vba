Option Explicit

'所有处理的入口，需要将各功能添加到case块中
'该入口会自动处理异常
Public Sub Main(ByVal taskName As String)
    On Error GoTo throw
    
    '开始前处理
    Call CommonModule.PreAllSubAndFunction
    
    '根据taskName来执行操作
    Select Case taskName
        'Case ""
            'Call
        Case Else
            MsgBox "未找到该处理", taskName
    End Select
    
    '结束处理
    Call CommonModule.EndAllSubAndFunction
    
    Exit Sub

throw:
    Call ExceptionModule.ThrowHandle
End Sub

Public Sub DoInit()
    On Error GoTo throw
    nowProjectName = "toolOpen"
    
    If ThisWorkbook.ReadOnly Then
        Call ExceptionModule.InforMationPopup("msg-01-001")
        ActiveWorkbook.Close Savechanges:=False
    End If
    
    Exit Sub
throw:
    Call ExceptionModule.ThrowHandle
End Sub
