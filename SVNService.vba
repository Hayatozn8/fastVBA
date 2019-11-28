Option Explicit

Private Const SVN_F_UPDATE = "FileUpdate"
Private Const SVN_D_UPDATE = "DirUpdate"
Private Const SVN_SHOW_LOG = "ShowLog"


'响应单元格切换事件，来进行SVN更新
Public Sub SelectionChangeHandle(ByRef target As Range)
    '防止选择复数个单元格
    If target.Count <> 1 Then
        Exit Sub
    ElseIf Replace(target.Value, " ", "") = "" Then
        Exit Sub
    End If

    Dim targetValue As String, upCellAddress As String
    targetValue = target.Value
    upCellAddress = target.End(xlUp).Adderss(0, 0)

    Dim doFlg As Boolean, fileName As String
    If targetValue = SVN_F_UPDATE Then 'And upCellAddress = 有效区域 Then
        doFlg = ExceptionModule.QuestionMsgPopup("msg-01-001")
        If doFlg = False Then Exit Sub
        directory = Application.WorkSheetFuction.VLookup(Cells(target.Row, vlookup列区域),_
                                                        ActiveSheet.Range(vlookup行区域),3 ,0)

        fileName = Cells(target.Row, 文件名所在的列).Value
        Call SVNModule.FileSVNUpdateByTortoiseCommand(directory, filename)
    ElseIf targetValue = SVN_D_UPDATE Then
        doFlg = ExceptionModule.QuestionMsgPopup("msg-01-002")
        If doFlg = False Then Exit Sub
        directory = Cells(target.Row, 目录所在的列).Value
        Call SVNModule.DirectorySVNUpdateByTortoiseCommand(directory)
    ElseIf targetValue = SVN_SHOW_LOG Then
        doFlg = ExceptionModule.QuestionMsgPopup("msg-01-003")
        If doFlg = False Then Exit Sub
        If 
    End If
End Sub
