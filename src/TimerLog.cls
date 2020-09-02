VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TimerLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'''<classSummary> イミディエイトウィンドウに処理時間に表示するクラス
'''<updateDate> 2020/9/2
'''<version> Ver.100
'''<remarks> 関数の一番最初に `Dim log As New TimerLog: log.Constructor ("test")`とすることで計測できる

Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long

Private FunctionName_ As String
Private Begin As Long
Private Finish As Long

'''<summary> 実質的なコンストラクタ 計測開始
'''<param> 時間と一緒に表示される文字列(関数名など)
'''<returns> なし
Public Function Constructor(ByVal FunctionName As String)
    FunctionName_ = FunctionName
    Begin = GetTickCount
    Debug.Print "[Begin] " & FunctionName
End Function

''<summary> 計測時間の結果を返す 表に表示させたいとき向け
'''<param> なし
'''<returns> 処理時間 [ms]
''<remarks> イミディエイトウィンドウに表示するもの(デストラクタ)とは別に処理を行っている
Public Function FinishTime() As Long
    Dim stopTime As Long
    stopTime = GetTickCount
    FinishTime = stopTime - Begin
End Function

'''<summary> デストラクタ このオブジェクトが破棄されたときに計測終了する
'''<param> なし
'''<returns> なし
Private Sub Class_Terminate()
    Finish = GetTickCount
    Debug.Print "[Finish] " & FunctionName_ & " , " & (Finish - Begin) & "[ms]"
End Sub
