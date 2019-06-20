Option Explicit

'通过Tortoise Command显示指定路径的SVN—log界面（Windows用）
'@param target 文件路径或文件夹路径
'@parma arg02 指定文件夹路径时，可以指定文件名到arg02，在方法内自动连接成新的目标
Public Sub ShowSVNLogByTortoiseCommand(target As String, Optional arg02 As String = "")
    If arg02 <> "" Then
        If Right(target, 1) = CommonParam.pathSeq Then
            target = target & arg02
        Else
            target = target & CommonParam.pathSeq & arg02
        End If
    End If
            
    '路径连接
    Dim workStr As String, projectStr As String
    workStr = "cmd.exe /c  TortoiseProc.exe /command:log /path:""" & target & """"
    
    Dim objshell As Object
    Set objshell = CommonModule.WscriptShell()
    
    '显示log
    objshell.Run workStr, 0, False
    Set objshell = Nothing
End Sub

'通过Tortoise Command更新SVN文件（Windows用）
'@param target 文件路径或文件夹路径
'@parma arg02 指定文件夹路径时，可以指定文件名到arg02，在方法内自动连接成新的目标
Public Sub FileSVNUpdateByTortoiseCommand(target As String, Optional arg02 As String = "")
    If arg02 <> "" Then
        If Right(target, 1) = CommonParam.pathSeq Then
            target = target & arg02
        Else
            target = target & CommonParam.pathSeq & arg02
        End If
    End If
    
    '防止excel文件在打开状态下更新
    If CommonModule.ExcelFileIsOpened(target) = True Then
        Call ExceptionModule.ErrorMsgPopup("svn-01-001", True)
    End If
    
    Call DoUpdateByTortoise(target)
End Sub

'通过Tortoise Command更新SVN文件夹（Windows用）
'@param target 文件夹路径
Public Sub DirectorySVNUpdateByTortoiseCommand(directory As String)
    Dim exitFlg As Integer
    exitFlg = CommonModule.FolderIsExist(directory)
    If exitFlg = "00" Then
        Call ExceptionModule.ErrorMsgPopup("err-01-001", True)
    ElseIf exitFlg = "02" Then
        Call ExceptionModule.ErrorMsgPopup("err-01-002", True, directory)
    End If
    
    Call DoUpdateByTortoise(directory)

End Sub

'执行SVN更新
Private Sub DoUpdateByTortoise(target As String)
    Dim workStr As String, projectStr As String
    workStr = "cmd.exe /c  TortoiseProc.exe /command:update /path:""" & target & """"
    
    Dim objshell As Object
    Set objshell = CommonModule.WscriptShell()
    
    objshell.Run workStr, 0, False
    Set objshell = Nothing
End Sub
