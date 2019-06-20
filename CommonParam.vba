Option Explicit

'执行操作后，需要复原的sheet对象
Public returnSheet As Variant

'当前正在执行的ProjectName
Public nowProjectName As String

'文件分隔符
Public Const pathSeq As String = "\"
