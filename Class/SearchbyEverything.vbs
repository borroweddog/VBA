'ファイル検索ソフトEverythingでキーワード検索
Private Const DEFAULT_EXE_PATH = "C:\Program Files\Everything\Everything.exe"
Private exePath As String


'初期値をセット
Private Sub Class_Initialize()
    exePath = DEFAULT_EXE_PATH
End Sub


'EXEパスを変更
Public Property Set ChangeExePath(ByVal arg_ExePath As String)
    exePath = arg_ExePath
End Property


'Everythingで検索
Public Sub Search(ByVal arg_Text As String)
    Const PARAM As String = "-search"
    Dim searchText As String
    Dim shellCMD As String
    
    searchText = Chr(34) & arg_Text & Chr(34)
    shellCMD = exePath & " " & PARAM & " " & searchText

    Shell shellCMD, vbNormalFocus

End Sub

