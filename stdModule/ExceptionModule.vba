Option Explicit

Private msgIdDict As Object
Private buttonValue As Integer

'删除msg字典
Public Sub DeleteMsgDict()
    msgIdDict.RemoveAll
    Set msgIdDict = Nothing
End Sub

'创建msg字典
Public Sub CreateMsgDict()
    Set msgIdDict = CommonModule.Dict()
    With msgIdDict
        .Add "err-00-000", Array("error from:" & msgArgHolder & vbCrLf & "error:" & msgArgHolder, "未知异常")
        .Add "err-01-001", Array("路径未指定", "error")
		.Add "err-01-002", Array("路径不存在" & vbCrLf & msgArgHolder, "error")
		
		.Add "svn-01-001", Array("请确认是否执行SNV更新", "SVN更新确认")
    End With
End Sub

'显示异常msg (推荐使用)
'@param msgID
'@param endFlg 是否终止程序
'@OP-param msgArg msg中的参数(字符串 or 数组)
'@OP-param titleArg 信息标题中的参数（字符串 or 数组）
'@Usage:
'    Call ExceptionModule.ErrorMsgPopup("xxxx", True, Array("aaa","bbb"), Array("ccc","ddd"))
'    Call ExceptionModule.ErrorMsgPopup("xxxx", True, "aaa", "bbb")
Public Sub ErrorMsgPopup(ByVal msgID As String, ByVal endFlg As Boolean, Optional msgArg As Variant = Nothing, Optional titleArg As Variant = Nothing)
    Call MessagePopup(msgID, vbExclamation, msgArg, titleArg)
    If endFlg = True Then
        Call ThrowHandle
    End If
End Sub

'显示提示msg (推荐使用)
'@param msgID
'@OP-param msgArg msg中的参数（字符串 or 数组）
'@OP-param titleArg 信息标题中的参数（字符串 or 数组）
'@Usage:
'    Call ExceptionModule.InforMationPopup("xxxx", Array("aaa","bbb"), Array("ccc","ddd"))
'    Call ExceptionModule.InforMationPopup("xxxx", "aaa", "bbb")
Public Sub InforMationPopup(ByVal msgID As String, Optional msgArg As Variant = Nothing, Optional titleArg As Variant = Nothing)
    Call MessagePopup(msgID, vbInformation, msgArg, titleArg)
End Sub

'显示可选择msg (推荐使用)
'@param msgID
'@OP-param msgArg msg中的参数（字符串 or 数组）
'@OP-param titleArg 信息标题中的参数（字符串 or 数组）
'@Usage:
'    Call ExceptionModule.QuestionMsgPopup("xxxx", Array("aaa","bbb"), Array("ccc","ddd"))
'    Call ExceptionModule.QuestionMsgPopup("xxxx", "aaa", "bbb")
Public Function QuestionMsgPopup(ByVal msgID As String, Optional msgArg As Variant = Nothing, Optional titleArg As Variant = Nothing) As Boolean
    Call MessagePopup(msgID, vbOKCancel + vbQuestion, msgArg, titleArg)
    If buttonValue = vbCancel Then
        QuestionMsgPopup = False
    Else
        QuestionMsgPopup = True
    End If
End Function

'显示异常msg (shellBox)
'@param msgID
'@param endFlg 是否终止程序
'@OP-param msgArg msg中的参数(字符串 or 数组)
'@OP-param titleArg 信息标题中的参数（字符串 or 数组）
'@Usage:
'    Call ExceptionModule.ErrorShellBoxPopup("xxxx", True, Array("aaa","bbb"), Array("ccc","ddd"))
'    Call ExceptionModule.ErrorShellBoxPopup("xxxx", True, "aaa", "bbb")
Public Sub ErrorShellBoxPopup(ByVal msgID As String, ByVal endFlg As Boolean, Optional msgArg As Variant = Nothing, Optional titleArg As Variant = Nothing)
    Call MessagePopup(msgID, vbExclamation, msgArg, titleArg, True)
    If endFlg = True Then
        Call ThrowHandle
    End If
End Sub

'显示提示msg (shellBox)
'@param msgID
'@OP-param msgArg msg中的参数（字符串 or 数组）
'@OP-param titleArg 信息标题中的参数（字符串 or 数组）
'@Usage:
'    Call ExceptionModule.InfoShellBoxPopup("xxxx", Array("aaa","bbb"), Array("ccc","ddd"))
'    Call ExceptionModule.InfoShellBoxPopup("xxxx", "aaa", "bbb")
Public Sub InfoShellBoxPopup(ByVal msgID As String, Optional msgArg As Variant = Nothing, Optional titleArg As Variant = Nothing)
    Call MessagePopup(msgID, vbInformation, msgArg, titleArg, True)
End Sub

'显示可选择msg (shellBox)
'@param msgID
'@OP-param msgArg msg中的参数（字符串 or 数组）
'@OP-param titleArg 信息标题中的参数（字符串 or 数组）
'@Usage:
'    Call ExceptionModule.QuestionMsgPopup("xxxx", Array("aaa","bbb"), Array("ccc","ddd"))
'    Call ExceptionModule.QuestionMsgPopup("xxxx", "aaa", "bbb")
Public Function QuestionShellBoxPopup(ByVal msgID As String, Optional msgArg As Variant = Nothing, Optional titleArg As Variant = Nothing) As Boolean
    Call MessagePopup(msgID, vbOKCancel + vbQuestion, msgArg, titleArg, True)
    If buttonValue = vbCancel Then
        QuestionMsgPopup = False
    Else
        QuestionMsgPopup = True
    End If
End Function

'msg popup
'@param msgID
'@buttons 按钮类型
'@param msgArg msg中的参数（字符串 or 数组）
'@param titleArg 信息标题中的参数（字符串 or 数组）
'
Public Sub MessagePopup(ByVal msgID As String, ByVal buttons As Integer, msgArg As Variant, ByVal titleArg As Variant, Optional popupType As Boolean = False)
    '创建字典
    Call CreateMsgDict
    Dim msgInfo()
    '获取msg信息
    msgInfo = msgIdDict.item(msgID)
    '删除字典
    Call DeleteMsgDict
    
    '获取msg和title
    Dim outMessage As String, outTitle As String
    outMessage = msgInfo(0)
    outTitle = msgInfo(1)
    
    '替换message中的占位符
    Select Case TypeName(msgArg)
        Case "Variant()"
            Dim mItem As Variant
            For Each mItem In msgArg
                outMessage = Replace(outMessage, msgArgHolder, mItem, , 1)
            Next
        Case "String"
            outMessage = Replace(outMessage, msgArgHolder, msgArg)
    End Select
    
    '替换title中的占位符
    Select Case TypeName(titleArg)
        Case "Variant()"
            Dim tItem As Variant
            For Each tItem In titleArg
                outTitle = Replace(outTitle, msgArgHolder, tItem, , 1)
            Next
        Case "String"
            outTitle = Replace(outTitle, msgArgHolder, titleArg)
    End Select
    
    '构造MsgBox
    If popupType = False Then
        buttonValue = MsgBox(outMessage, buttons, outTitle)
    Else
        Dim obj As Object
        Set obj = CommonModule.WscriptShell()
        buttonValue = obj.Popup(outMessage, ,outTitle, buttons)
        Set obj = Nothing
    End If
End Sub

'异常处理
Public Sub ThrowHandle()
    If Err.Number <> 0 Then
        Call MessagePopup("err-00-000", vbOKOnly + vbExclamation, Array(nowProjectName, Err.Description), "")
    End If
    
    '执行结束处理
    Call CommonModule.EndAllSubAndFunction
    
    End
End Sub
