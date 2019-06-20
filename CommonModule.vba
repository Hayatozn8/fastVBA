Option Explicit

'所有sub的启动前处理(提高整体性能)
Public Sub PreAllSubAndFunction()
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    ActiveWorkbook.CheckCompatibility = False
    Application.Calculation = xlCalculationManual
  
    nowProjectName = ""
    Err.Clear
  
    Set returnSheet = ActiveSheet
End Sub

'所有sub的结束处理(复原启动前处理的设定)
'1.还原默认值设定
'2.清空error
'3.sheet和book的保护还原
Public Sub EndAllSubAndFunction()
    nowProjectName = ""
    Err.Clear
    
    Application.Calculation = xlCalculationAutomatic
    ActiveWorkbook.CheckCompatibility = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    
End Sub

'弹出文件选择框--> 选择文件--> 返回文件path
'@param moreDiaFlg 多文件选择Flg
'@OP-param defaultPath 默认路径
'@OP-param defaultType 默认文件类型
'@return 选中的文件path
'@tips 返回值的几种形式
'1.选中了文件：返回选中的路径 return new path
'2.未执行选中操作：返回默认路径 return old path(defaultPath)
'    2-1.设定了defaultPath return old path(defaultPath)
'    2-2.未设定defaultPath return ""
Public Function FileDialogAndReturn(ByVal moreDiaFlg As Boolean, Optional defaultPath As String, Optional defaultType As String) As String
    With Application.FileDialog(msoFileDialogFilePicker)
        '多选择
        .AllowMultiSelect = moreDiaFlg
        
        '设定默认路径
        If PathIsExist(defaultPath) = "01" Then .InitialFileName = defaultPath
        
        '设定默认文件类型
        .Filters.Clear
        If defaultType <> "" Then .Filters.Add "defaultType", "*." & defaultType
        .Filters.Add "All Files", "*.*"
        
        '设定选择框标题
        .title = "请选择文件"
        
        '弹出文件选择框
        .Show
        
        '获取选择结果
        If .SelectedItems.Count = 1 Then
            '选中了某个文件
            FileDialogAndReturn = .SelectedItems(1)
        Else
            '取消选择时，返回默认路径
            FileDialogAndReturn = defaultPath
        End If
    End With
End Function
        
'弹出文件夹选择框--> 选择文件夹--> 返回文件夹选择path
'@param moreDiaFlg 多文件选择Flg
'@OP-param defaultFolder 默认文件夹
'@return 选中的文件夹
'tips 返回值的几种形式
'1.选中了文件：返回选中的路径 return new path
'2.未执行选中操作：返回默认路径 return old path(defaultPath)
'    2-1.设定了defaultPath return old path(defaultPath)
'    2-2.未设定defaultPath return ""
Public Function FolderDialogAndReturn(ByVal moreDiaFlg As Boolean, Optional defaultFolder As String) As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        '多选择
        .AllowMultiSelect = moreDiaFlg

        '设定默认文件夹
        If PathIsExist(defaultFolder) = "01" Then .InitialFileName = defaultFolder & pathSeq
        
        '设定选择框标题
        .title = "请选择文件夹"
        
        '弹出文件夹选择框
        .Show
        
        '获取选择结果
        If .SelectedItems.Count = 1 Then
            '选中了某个文件夹
            FolderDialogAndReturn = .SelectedItems(1)
        Else
            '取消选择时，返回默认路径
            FolderDialogAndReturn = defaultFolder
        End If
    End With
End Function

'check路径是否存在
'@param target check对象
'@return check结果:
'    "00" = 未设定check对象
'    "01" = check对象存在
'    "02" = check对象不存在
Public Function PathIsExist(ByVal target As String) As String
    '如果路径是空字符串，则跳过
    If target = "" Then
        PathIsExist = "00"
    Else
        Dim result As String
        result = dir(target, vbDirectory)
        If result = "" Then
            PathIsExist = "02"
        Else
            PathIsExist = "01"
        End If
    End If
End Function

'check文件是否存在
'@param target check对象
'@return check结果:
'    "00" = 未设定check对象
'    "01" = check对象存在
'    "02" = check对象不存在
'    "03" = 文件类型不一致
Public Function FileIsExist(ByVal target As String, Optional fileType As String = "") As String
    If target = "" Then
        FileIsExist = "00"
    Else
        Dim fsObj As Object
        Set fsObj = CreateObject("Scripting.FileSystemObject")
        If Not fsObj.FileExists(target) Then
            FileIsExist = "02"
        Else
            'check文件类型
            If fileType <> "" And fileType <> fsObj.GetExtensionName(target) Then
                FileIsExist = "03"
            Else
                FileIsExist = "01"
            End If
        End If
        Set fsObj = Nothing
    End If
End Function

'check文件夹是否存在
'@param target check对象
'@return check结果:
'    "00" = 未设定check对象
'    "01" = check对象存在
'    "02" = check对象不存在
Public Function FolderIsExist(ByVal target As String) As String
    If target = "" Then
        FolderIsExist = "00"
    Else
        Dim fsObj As Object
        Set fsObj = CreateObject("Scripting.FileSystemObject")
        If Not fsObj.FolderExists(target) Then
            FolderIsExist = "02"
        Else
            FolderIsExist = "01"
        End If
        Set fsObj = Nothing
    End If
End Function

'路径分割： dir+endName
'@param target 分割目标
'@Ref-param directory
'@Ref-param endName
'@return dir 
'@return endName
Public Sub CutPath(ByVal target As String, ByRef directory, ByRef endName As String)
    Dim fsObj As Object
    Set fsObj = CreateObject("Scripting.FileSystemObject")
    directory = fsObj.GetParentFolderName(target)
    endName = fsObj.GetFileName(target)
    
    Set fsObj = Nothing
End Sub

'excel文件是否处于打开状态
'@param target check对象
'@return check结果：
'    true:文件被打开
'    false:文件关闭
Public Function ExcelFileIsOpened(ByVal target As String) As Boolean
    Dim MyXL As Object
    Dim axls As Object
    '返回值初始化
    ExcelFileIsOpened = False
    
    '获取excel_Application对象
    Set MyXL = GetObject(, "Excel.Application")
    For Each axls In MyXL.Workbooks
        If axls.path & pathSeq & axls.NAME = target Then
            ExcelFileIsOpened = True
            Exit For
        End If
    Next axls
    
    Set axls = Nothing
    Set MyXL = Nothing
End Function

'获取文件的Byte数
'@param target 获取目标
'@return 目标的Byte数
Public Function GetFileByteCount(ByVal target As String) As Double
    Dim fsObj As Object
    Dim theFile As Object
    Set fsObj = CreateObject("Scripting.FileSystemObject")
    Set theFile = fsObj.GetFile(target)
    GetFileByteCount = theFile.Size
    
    Set theFile = Nothing
    Set fsObj = Nothing
End Function

'关闭目标excel文件
'@param target 需要关闭的excel文件路径
Public Sub CloseExcelFile(ByVal target As String)
    Dim MyXL As Object
    Dim axls As Object
    
    '获取excel_Application对象
    Set MyXL = GetObject(, "Excel.Application")
    If Err.Number <> 0 Then
        Exit Sub
    End If
    
    For Each axls In MyXL.Workbooks
        If axls.path & pathSeq & axls.NAME = target Then
            axls.Close
            Exit For
        End If
    Next axls
    
    Set axls = Nothing
    Set MyXL = Nothing
End Sub

'创建一个字典对象（用完后，需要手动删除该对象）
Public Function Dict() As Object
    Set Dict = CreateObject("Scripting.Dictionary")
End Function

'创建一个WscriptShell对象（用完后，需要手动删除该对象）
Public Function WscriptShell() As Object
    Set WscriptShell = CreateObject("Wscript.shell")
End Function
