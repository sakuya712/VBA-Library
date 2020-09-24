VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ArrayLib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'''<classSummary> 2次配列を使いやすくするクラス (2次配列専用です。1次配列の場合はListが使いやすいです)
'''<updateDate> 2020/8/31
'''<version> Ver.100
'''<remarks> このクラスはエクセルで使う前提で作成しています。
Private Arr() As Variant

'''<summary> 自分を返すためのもの
'''<param> なし
'''<returns> 自己参照
Public Property Get self() As ArrayLib
    Set self = Me
End Property

'''<summary> セットされた配列があるかどうか
'''<param> なし
'''<returns> 真偽値
Public Property Get IsArray() As Boolean
    On Error GoTo err
    Dim buf
    buf = UBound(Arr)
    IsArray = True
    Exit Function
err:
    IsArray = False
End Property

'''<summary> 配列の値を返す、既定メンバ
'''<param> 1次配列目の番号、2次配列目の番号
'''<returns> 値
Public Property Get Items(ByVal Row As Long, ByVal Col As Long) As Variant
Attribute Items.VB_UserMemId = 0
    Items = Arr(Row, Col)
End Property

'''<summary> 1次配列をセットする (2次配列に変換)
'''<param> 1次配列、行データにするか列データにするか
'''<returns> なし
Public Function SetArray1Dim(ByVal setArr As Variant, ByVal IsRow As Boolean)
    
    Dim low As Long
    Dim max As Long
    low = LBound(setArr)
    max = UBound(setArr)
    
    Dim Row As Long
    Dim Col As Long
    If IsRow Then
        ReDim Arr(low To max, 1 To 1)
        For Row = low To max
            Arr(Row, 1) = setArr(Row)
        Next
    Else
        ReDim Arr(1 To 1, low To max)
        For Col = low To max
            Arr(1, Col) = setArr(Col)
        Next
    End If
    
End Function

'''<summary> 2次配列をセットする
'''<param> 2次配列
'''<returns> なし
Public Function SetArray2Dim(ByVal setArr As Variant)
    Arr() = setArr
End Function

'''<summary> Rangeで範囲を決めて値を配列を設定する
'''<param> 範囲の左上のセル、範囲の右下のセル
'''<returns> なし
'''<remarks> 2つ目の引数を省略した場合CurrentRegionと同じような取得方法になるただし、cell1より、左、上にはならない
Public Function SetRange(ByVal cell1 As Range, Optional ByVal cell2 As Range = Nothing)
    
    If cell2 Is Nothing Then
        Dim BottomRow As Long
        Dim RightColumn As Long
        BottomRow = cell1.End(xlDown).Row
        RightColumn = cell1.End(xlToRight).Column
        Set cell2 = cell1.Parent.Cells(BottomRow, RightColumn)
    End If
    Arr() = Range(cell1, cell2).Value
    
End Function

'''<summary>要素数を取得する
'''<param> 次元の指定 省略した場合はすべての次元の要素数を戻す
'''<returns> 要素数
Public Function Length(Optional ByVal dimension As Long = -1) As Long

    If dimension = -1 Then
        Length = (UBound(Arr, 1) - LBound(Arr, 1) + 1) * (UBound(Arr, 2) - LBound(Arr, 2) + 1)
    Else
        Length = UBound(Arr, dimension) - LBound(Arr, dimension) + 1
    End If
    
End Function

'''<summary>引数と一致する行の配列を返す
'''<param> 一致する値、比較演算子する列
'''<returns> 2次配列
Public Function MatchRow(ByVal Str As String, Optional ByVal Col As Long = -1) As ArrayLib

    If Col = -1 Then Col = LBound(Arr, 1)
    
    Dim fn As WorksheetFunction
    Set fn = Application.WorksheetFunction
    
    Dim buf As New ArrayLib
    Dim Row As Long
    For Row = LBound(Arr) To UBound(Arr)
        If Arr(Row, Col) = Str Then
            buf.SetArray1Dim fn.index(Arr, Row), False
            Set MatchRow = buf
            Exit Function
        End If
    Next
    Set MatchRow = buf
    
End Function

'''<summary>要素の行もしは列を削除する
'''<param> 行番号、列番号 (どちらか一方のみを設定すること)
'''<returns> なし
Public Function Remove(Optional ByVal DelRow As Long = -1, Optional ByVal DelCol As Long = -1)

    If DelRow <> -1 And DelCol <> -1 Then GoTo err1
    
    Dim lowRow As Long
    Dim maxRow As Long
    Dim lowCol As Long
    Dim maxCol As Long
    lowRow = LBound(Arr, 1)
    maxRow = UBound(Arr, 1)
    lowCol = LBound(Arr, 2)
    maxCol = UBound(Arr, 2)
    Dim Row As Long
    Dim Col As Long
    Dim plus As Long
    
    Dim buf() As Variant
    If DelRow <> -1 Then
        ReDim buf(lowRow To maxRow - 1, lowCol To maxCol)
        For Row = lowRow To maxRow - 1
            If Row = DelRow Then plus = 1
            For Col = lowCol To maxCol
                buf(Row, Col) = Arr(Row + plus, Col)
            Next
        Next
    Else
        ReDim buf(lowRow To maxRow, lowCol To maxCol - 1)
        For Row = lowRow To maxRow
            For Col = lowCol To maxCol - 1
                If Col = DelCol Then plus = 1
                buf(Row, Col) = Arr(Row, Col + plus)
            Next
        Next
    End If
    Erase Arr
    Me.SetArray2Dim buf
    Exit Function
err1:
    err.Raise Number:=1201, Description:="クラス：ArrayLib、メソッド：Remove、引数を二つ設定することはできません"

End Function

'''<summary>IndexMatch(行のみ)を使って値を見つける
'''<param> 行側のキーワード、キーにする列、[列の指定、エラーのとき値]
'''<returns> 見つかった値
Public Function IndexMatch(ByVal RowStr As String, ByVal ColIndex As Long, _
                                            Optional ByVal Col As Long = 1, Optional ByVal Iferror As String = "") As String

    On Error GoTo err1
    Dim fn As WorksheetFunction
    Set fn = Application.WorksheetFunction
    Dim ColArr As Variant
    ColArr = fn.index(fn.Transpose(Arr()), ColIndex)
    IndexMatch = fn.index(Arr(), fn.Match(RowStr, ColArr, 0), Col)
    Exit Function
err1:
    IndexMatch = Iferror
    
End Function

'''<summary>IndexMatch(行、列)を使って値を見つける
'''<param> 行側のキーワード、キーにする列、列側のキーワード、キーにする列、[エラーのとき値]
'''<returns> 見つかった値
Public Function IndexMatchMatch(ByVal RowStr As String, ByVal ColIndex As Long, _
                                                    ByVal ColStr As String, ByVal RowIndex As Long, Optional ByVal Iferror As String = "") As String

    On Error GoTo err1
    Dim fn As WorksheetFunction
    Set fn = Application.WorksheetFunction
    Dim ColArr As Variant
    ColArr = fn.index(fn.Transpose(Arr()), ColIndex)
    Dim RowArr As Variant
    RowArr = fn.index(Arr(), RowIndex)
    IndexMatchMatch = fn.index(Arr(), fn.Match(RowStr, ColArr, 0), fn.Match(ColStr, RowArr, 0))
    Exit Function
err1:
    IndexMatchMatch = Iferror
End Function

'''<summary>通常の配列に変換する
'''<param> なし
'''<returns> 2次配列
Public Function ToArray() As Variant()
    ToArray = Arr()
End Function
