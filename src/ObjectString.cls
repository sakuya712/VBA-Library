VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ObjectString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'''<classSummary> オブジェクト型のString、文字操作をメソッドで行うことができる
'''<updateDate> 2020/10/9
'''<version> Ver.103

Private ObjString As String

'''<summary> 文字をセットする
'''<param> セットする文字
'''<returns> なし
Public Property Let Item(ByVal Arg As String)
Attribute Item.VB_UserMemId = 0
    On Error GoTo err
    ObjString = VBA.CStr(Arg)
    Exit Property
err:
    ObjString = VBA.TypeName(Arg)
End Property

'''<summary> 文字をゲットする
'''<param> なし
'''<returns> 文字
Public Property Get Item() As String
    Item = ObjString
End Property

'''<summary> 文字数を取得する
'''<param> なし
'''<returns> 文字数
Public Property Get Length() As Long
    Length = Len(ObjString)
End Property

'''<summary> 通常のStringに変換する
'''<param> なし
'''<returns> String型
Public Function ToString() As String
    ToString = ObjString
End Function

'''<summary> 数値型に変換する
'''<param> なし
'''<returns> 数値型
Public Function ToNumber() As Variant
    ToNumber = VBA.Val(ObjString)
End Function

'''<summary> 強制的に文字列型にしてObjectStringを作成する(Nullはエラーになります)
'''<param> 文字
'''<returns> ObjectString型
Public Function SetString(ByVal Arg As Variant) As ObjectString
    Dim NewStr As New ObjectString
    NewStr = CStr(Arg)
    Set SetString = NewStr
End Function

'''<summary> 指定された文字列より前の文字を抜き取る
'''<param> 基準となる文字列
'''<returns> ObjectString型
Public Function Rstrip(ByVal char As String) As ObjectString

    Dim NewStr As New ObjectString
    Dim pos As Long
    pos = VBA.InstrRev(ObjString, char)
    If pos = 0 Then
        NewStr = ObjString
    Else
        NewStr = VBA.Left(ObjString, pos - 1)
    End If
    Set Rstrip = NewStr
End Function

'''<summary> 指定された文字列より後の文字を抜き取る
'''<param> 基準となる文字列
'''<returns> ObjectString型
Public Function Lstrip(ByVal char As String) As ObjectString

    Dim NewStr As New ObjectString
    Dim Arr As Variant
    Arr = VBA.Split(ObjString, char, 2)
        Select Case UBound(Arr)
            Case -1, 0
                NewStr = ObjString
            Case Else
                NewStr = Arr(1)
        End Select
    Set Lstrip = NewStr
End Function

'''<summary> 指定された文字列より前後の文字を抜き取る
'''<param> 前の基準となる文字列、後ろの基準となる文字列
'''<returns> ObjectString型
Public Function Strip(ByVal BeginChar As String, ByVal EndChar As String) As ObjectString

    Dim NewStr As New ObjectString
    NewStr = Me.Lstrip(BeginChar).Rstrip(EndChar)
    Set Strip = NewStr
End Function

'''<summary> 特定の文字を置換する
'''<param> 置換対象、置換後の文字
'''<returns> ObjectString型
Public Function Replace(ByVal findString As String, ByVal replaceString As String) As ObjectString
    Dim NewStr As New ObjectString
    NewStr = VBA.Replace(ObjString, findString, replaceString)
    Set Replace = NewStr
End Function

'''<summary> 引数の文字があるかどうかを真偽値で返す
'''<param> 見つける文字
'''<returns> 真偽値
Public Function Contains(ByVal char As String) As Boolean
    If VBA.InStr(ObjString, char) Then
        Contains = True
    Else
        Contains = False
    End If
End Function

'''<summary> 引数の文字の文字位置を返す
'''<param> 見つける文字
'''<returns> 文字位置。ない場合は-1を返す
Public Function IndexOf(ByVal char As String) As Long
    If VBA.InStr(ObjString, char) Then
        IndexOf = VBA.InStr(ObjString, char)
    Else
        IndexOf = -1
    End If
End Function

'''<summary> 引数の文字から先頭にあるかを真偽値で返す
'''<param> 見つける文字
'''<returns> 真偽値
Public Function StartsWith(ByVal char As String) As Boolean
    If ObjString Like char & "*" Then
        StartsWith = True
    Else
        StartsWith = False
    End If
End Function

'''<summary> 引数の文字が末尾にあるかを真偽値で返す
'''<param> 見つける文字
'''<returns> 真偽値
Public Function EndWith(ByVal char As String) As Boolean
    If ObjString Like "*" & char Then
        EndWith = True
    Else
        EndWith = False
    End If
End Function

'''<summary> 文字列を連結させる
'''<param> 区切り文字、連結させたい文字列(複数可能)
'''<returns> ObjectString型
'''<remarks> 連結演算子(&)より処理は遅い、コードの見やすさ重視で使用します。
Public Function Concat(ByVal Delimiter As String, ParamArray Strings() As Variant) As ObjectString
    Dim NewStr As New ObjectString
    Dim Str As Variant
    Dim i As Long
    If ObjString = "" Then
        NewStr = CStr(Strings(0))
        For i = 1 To UBound(Strings)
            NewStr = NewStr & Delimiter & CStr(Strings(i))
        Next
    Else
        NewStr = ObjString
        For Each Str In Strings
            NewStr = NewStr & Delimiter & CStr(Str)
        Next
    End If
    Set Concat = NewStr
End Function

'''<summary> 文字列を分割します
'''<param> 区切り文字、分割数(指定しない場合は全て)
'''<returns> 0から始まる1次元配列
Public Function Split(ByVal Delimiter As String, Optional ByVal limit As Long = -1) As Variant
    Split = VBA.Split(ObjString, Delimiter, limit)
End Function

'''<summary> 文字列を切り抜きます
'''<param> 開始位置、文字の長さ(指定しない場合は全て)
'''<returns> 切り抜いた文字
Public Function Substring(ByVal StartIndex As Long, Optional ByVal Length As Long) As ObjectString
    Dim NewStr As New ObjectString
    If Length < 1 Then
        NewStr = VBA.Mid(ObjString, StartIndex)
    Else
        NewStr = VBA.Mid(ObjString, StartIndex, Length)
    End If
    Set Substring = NewStr
End Function

'''<summary> 文字列の間に文字を挿入する
'''<param> 開始位置、挿入する文字
'''<returns> 挿入された文字
Public Function Insert(ByVal StartIndex As Long, ByVal AddString As String) As ObjectString
    Dim NewStr As New ObjectString
    NewStr = VBA.Left(ObjString, StartIndex - 1) & AddString & VBA.Mid(ObjString, StartIndex)
    Set Insert = NewStr
End Function
