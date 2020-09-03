VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'''<classSummary> Collectionにメソッドを拡張したList型、キーは使用できません。
'''<updateDate> 2020/9/3
'''<version> Ver.104

'''<summary> 比較演算子の列挙型
Enum ComparisonOperatorsEnum
    lsEqual = 1
    lsNotEqual = 2
    lsGreater = 3
    lsLess = 4
    lsGreaterEqual = 5
    lsLessEqual = 6
    lsObjectEqual = 7
    lsLike = 8
    lsNotLike = 9
End Enum

'''<summary> 算術演算子の列挙型
Enum ArithmeticOperatorsEnum
    lsSum = 101
    lsDifference = 102
    lsMultiply = 103
    lsDivide = 104
    lsMod = 105
    lsExponent = 106
    lsConcatenate = 107
End Enum

Private Coll As Collection

'''<summary> 自分を返すためのもの
'''<param> なし
'''<returns> 自己参照
Public Property Get self() As List
    Set self = Me
End Property

'''<summary> オブジェクトをセット
'''<param> なし
'''<returns> なし
Private Sub Class_Initialize()
    Set Coll = New Collection
End Sub

'''<summary> 要素を追加する
'''<param> 追加する要素 (オブジェクトや配列そのものが入る)
'''<returns> なし
Public Function Add(ByVal Item As Variant)
    Coll.Add Item
End Function

'''<summary> 配列の値、Rangeの値、オブジェクトの既定メンバーの値を格納する
'''<param> 追加する要素(Range)
'''<returns> なし
Public Function AddValue(ByVal Items As Variant)
    On Error GoTo err
    Dim i
    If TypeName(Items) = "Range" Then
        Dim ItemArr() As Variant
        If Items.Count = 1 Then
            Coll.Add Items.Value
            Exit Function
        End If
        ItemArr() = Items.Value
        For Each i In ItemArr
            Coll.Add i
        Next
        Exit Function
    End If
    For Each i In Items
        Coll.Add i
    Next
    Exit Function
err:
    err.Raise Number:=1001, Description:="クラス：List、メソッド：AddValue、引数の型が適切ではありません。 "
    
End Function

'''<summary> 要素の内容を返す
'''<param> インデックス番号
'''<returns> 内容
'''<remarks> 既定メンバ
Public Function Item(ByVal index As Long) As Variant
Attribute Item.VB_UserMemId = 0
    If VarType(Coll(index)) = vbObject Then
        Set Item = Coll(index)
    Else
        Item = Coll(index)
    End If
End Function

'''<summary> 要素数を返す
'''<param> なし
'''<returns> 要素数
Public Function Count() As Long
    Count = Coll.Count
End Function

'''<summary> 要素数を削除する
'''<param> 削除するインデックス番号
'''<returns> なし
Public Function Remove(ByVal index As Long)
    Coll.Remove (index)
End Function

'''<summary> 差集合のListを作成する
'''<param> 取り除きたい要素のList
'''<returns> 取り除かれたList
Public Function Except(ByVal ExceptList As List) As List
    Dim lst As List
    Set lst = New List
    lst.AddValue Coll
    Dim e As Variant
    Dim index As Long
    For Each e In ExceptList
        index = 1
        Do While index <= lst.Count
            If e = lst(index) Then
                lst.Remove (index)
            Else
                index = index + 1
            End If
        Loop
    Next
    
    Set Except = lst

End Function

'''<summary> 要素に重複があるか調べる
'''<param> なし
'''<returns> 重複がある場合はTrueを返す
Public Function IsOverlap() As Boolean
    Dim buf As Collection
    Set buf = New Collection
    Dim c
    'コレクションのキーが重複するとエラーになるなるのを利用して見つける
    On Error GoTo err
    For Each c In Coll
        buf.Add c, CStr(c)
    Next
    IsOverlap = False
    Exit Function
err:
    IsOverlap = True
End Function

'''<summary> 重複がないListを作成する
'''<param> なし
'''<returns> 重複を削除したList
Public Function Unique() As List
    Dim buf As Collection
    Set buf = New Collection
    Dim c As Variant
    'コレクションのキーが重複するとエラーになるなるのを利用して見つける
    On Error Resume Next
    For Each c In Coll
        buf.Add c, CStr(c)
    Next
    Dim lst As List
    Set lst = New List
    lst.AddValue buf
    Set Unique = lst
End Function

'''<summary> 重複した値をリスト化する
'''<param> なし
'''<returns> 重複した値のList
Public Function OverlapList() As List
    Dim Overlap As List
    Set Overlap = New List
    Dim c1 As Variant
    Dim c2 As Variant
    Dim cnt As Long
    For Each c1 In Coll
        cnt = 0
        For Each c2 In Coll
            If c1 = c2 Then
                cnt = cnt + 1
                If cnt > 1 Then
                    Overlap.Add c1
                    Exit For
                End If
            End If
        Next
    Next
    Set Overlap = Overlap.Unique

    Set OverlapList = Overlap
    
End Function

'''<summary> Listの中に引数とマッチする要素が最初に出現するインデックスを返します。見つからない場合は -1 を返します。
'''<param> 検索したい文字列、完全一致の場合はTrue部分一致の場合はFlase
'''<returns> インデックス番号
Public Function IndexOf(ByVal Str As String, Optional ByVal PerfectMatching As Boolean = True) As Long

    Dim cnt As Long
    Dim Value As Variant
    If PerfectMatching Then
        For Each Value In Me
            cnt = cnt + 1
            If Str = Value Then
                IndexOf = cnt
                Exit Function
            End If
        Next
    Else
        For Each Value In Me
            cnt = cnt + 1
            If InStr(Value, Str) Then
                IndexOf = cnt
                Exit Function
            End If
        Next
    End If
    IndexOf = -1
End Function

'''<summary> Listの中に引数とマッチする要素の有無を真偽値で返す
'''<param> 検索したい文字列、完全一致の場合はTrue部分一致の場合はFlase
'''<returns> ある場合True
Public Function Includes(ByVal Str As String, Optional ByVal PerfectMatching As Boolean = True) As Boolean

    Dim Value As Variant
    If PerfectMatching Then
        For Each Value In Me
            If Str = Value Then
                Includes = True
                Exit Function
            End If
        Next
    Else
        For Each Value In Me
            If InStr(Value, Str) Then
                Includes = True
                Exit Function
            End If
        Next
    End If
    Includes = False
End Function

'''<summary> List同士を結合して新たにListを作成する
'''<param> 結合したいList
'''<returns> 結合したList
Public Function Concat(ParamArray Lists() As Variant) As List

    Dim buf As List
    Set buf = New List
    buf.AddValue Me
    Dim l As Variant
    For Each l In Lists
        buf.AddValue l
    Next
    Set Concat = buf

End Function

'''<summary> 同じ要素数のListを作成する。
'''<param> Listに格納する要素
'''<returns> List型
Public Function ToList(ByVal Value As Variant) As List

    Dim lst As List
    Set lst = New List
    Dim i As Long
    For i = 1 To Me.Count
        lst.Add Value
    Next
    Set ToList = lst

End Function

'''<summary> 条件にあう要素のみを残し、新たにListを作成する
'''<param> 比較演算子の列挙型、比較対象(List型、Collection型、プリミティブ型に対応)
'''<returns> List型
Public Function Where(ByVal operator As ComparisonOperatorsEnum, ByVal compare As Variant) As List

    Dim buf As List
    Set buf = New List
    Dim i As Long
    Dim var1, var2
    
    Select Case TypeName(compare)
        Case "List", "Collection"
            Select Case operator
                Case ComparisonOperatorsEnum.lsEqual
                    For i = 1 To Me.Count
                        var1 = Me(i)
                        var2 = compare(i)
                        If var1 = var2 Then
                            buf.Add var1
                        End If
                    Next
                Case ComparisonOperatorsEnum.lsNotEqual
                    For i = 1 To Me.Count
                        var1 = Me(i)
                        var2 = compare(i)
                        If var1 <> var2 Then
                            buf.Add var1
                        End If
                    Next
                Case ComparisonOperatorsEnum.lsLike
                    For i = 1 To Me.Count
                        var1 = CStr(Me(i))
                        var2 = CStr(compare(i))
                        If var1 Like var2 Then
                            buf.Add var1
                        End If
                    Next
                Case ComparisonOperatorsEnum.lsNotLike
                    For i = 1 To Me.Count
                        var1 = CStr(Me(i))
                        var2 = CStr(compare(i))
                        If Not (var1 Like var2) Then
                            buf.Add var1
                        End If
                    Next
                Case ComparisonOperatorsEnum.lsGreater
                    For i = 1 To Me.Count
                        var1 = Val(Me(i))
                        var2 = Val(compare(i))
                        If var1 > var2 Then
                            buf.Add var1
                        End If
                    Next
                Case ComparisonOperatorsEnum.lsLess
                    For i = 1 To Me.Count
                        var1 = Val(Me(i))
                        var2 = Val(compare(i))
                        If var1 < var2 Then
                            buf.Add var1
                        End If
                    Next
                Case ComparisonOperatorsEnum.lsGreaterEqual
                    For i = 1 To Me.Count
                        var1 = Val(Me(i))
                        var2 = Val(compare(i))
                        If var1 >= var2 Then
                            buf.Add var1
                        End If
                    Next
                Case ComparisonOperatorsEnum.lsLessEqual
                    For i = 1 To Me.Count
                        var1 = Val(Me(i))
                        var2 = Val(compare(i))
                        If var1 <= var2 Then
                            buf.Add var1
                        End If
                    Next
                Case ComparisonOperatorsEnum.lsObjectEqual
                    For i = 1 To Me.Count
                        var1 = Me(i)
                        var2 = compare(i)
                        If var1 Is var2 Then
                            buf.Add var1
                        End If
                    Next
                Case Else
                    err.Raise Number:=1005, Description:="クラス：List、メソッド：Where、引数の型が適切ではありません。 "
            End Select
        Case Else
             Select Case operator
                Case ComparisonOperatorsEnum.lsEqual
                    For i = 1 To Me.Count
                        If Me(i) = compare Then
                            buf.Add Me(i)
                        End If
                    Next
                Case ComparisonOperatorsEnum.lsNotEqual
                    For i = 1 To Me.Count
                        If Me(i) <> compare Then
                            buf.Add Me(i)
                        End If
                    Next
                Case ComparisonOperatorsEnum.lsLike
                    For i = 1 To Me.Count
                        If CStr(Me(i)) Like CStr(compare) Then
                            buf.Add Me(i)
                        End If
                    Next
                Case ComparisonOperatorsEnum.lsNotLike
                    For i = 1 To Me.Count
                        If Not (CStr(Me(i)) Like CStr(compare)) Then
                            buf.Add Me(i)
                        End If
                    Next
                Case ComparisonOperatorsEnum.lsGreater
                    For i = 1 To Me.Count
                        If Val(Me(i)) > Val(compare) Then
                            buf.Add Me(i)
                        End If
                    Next
                Case ComparisonOperatorsEnum.lsLess
                    For i = 1 To Me.Count
                        If Val(Me(i)) < Val(compare) Then
                            buf.Add Me(i)
                        End If
                    Next
                Case ComparisonOperatorsEnum.lsGreaterEqual
                    For i = 1 To Me.Count
                        If Val(Me(i)) >= Val(compare) Then
                            buf.Add Me(i)
                        End If
                    Next
                Case ComparisonOperatorsEnum.lsLessEqual
                    For i = 1 To Me.Count
                        If Val(Me(i)) <= Val(compare) Then
                            buf.Add Me(i)
                        End If
                    Next
                Case ComparisonOperatorsEnum.lsObjectEqual
                    For i = 1 To Me.Count
                        If Me(i) Is compare Then
                            buf.Add Me(i)
                        End If
                    Next
                Case Else
                    err.Raise Number:=1005, Description:="クラス：List、メソッド：Where、引数の型が適切ではありません。 "
            End Select
        End Select

    Set Where = buf

End Function

'''<summary> 要素の中身を計算や結合し、新たにListを作成する
'''<param> 算術演算子の列挙型、算術対象(List型、Collection型、プリミティブ型に対応)
'''<returns> List型
Public Function Map(ByVal operator As ArithmeticOperatorsEnum, ByVal target As Variant) As List

    Dim buf As List
    Set buf = New List
    Dim i As Long
    Dim var1, var2

    Select Case TypeName(target)
        Case "List", "Collection"
            Select Case operator
                Case ArithmeticOperatorsEnum.lsConcatenate
                     For i = 1 To Me.Count
                        var1 = CStr(Me(i))
                        var2 = CStr(target(i))
                        buf.Add var1 & var2
                    Next
                Case ArithmeticOperatorsEnum.lsSum
                    For i = 1 To Me.Count
                        var1 = Val(Me(i))
                        var2 = Val(target(i))
                        buf.Add var1 + var2
                    Next
                Case ArithmeticOperatorsEnum.lsDifference
                    For i = 1 To Me.Count
                        var1 = Val(Me(i))
                        var2 = Val(target(i))
                        buf.Add var1 - var2
                    Next
                Case ArithmeticOperatorsEnum.lsMultiply
                    For i = 1 To Me.Count
                        var1 = Val(Me(i))
                        var2 = Val(target(i))
                        buf.Add var1 * var2
                    Next
                Case ArithmeticOperatorsEnum.lsDivide
                     For i = 1 To Me.Count
                        var1 = Val(Me(i))
                        var2 = Val(target(i))
                        buf.Add var1 / var2
                    Next
                Case ArithmeticOperatorsEnum.lsMod
                     For i = 1 To Me.Count
                        var1 = Val(Me(i))
                        var2 = Val(target(i))
                        buf.Add var1 Mod var2
                    Next
                Case ArithmeticOperatorsEnum.lsExponent
                     For i = 1 To Me.Count
                        var1 = Val(Me(i))
                        var2 = Val(target(i))
                        buf.Add var1 ^ var2
                    Next
                Case Else
                    err.Raise Number:=1006, Description:="クラス：List、メソッド：Map、引数の型が適切ではありません。 "
            End Select
        Case Else
            Select Case operator
                 Case ArithmeticOperatorsEnum.lsConcatenate
                     For i = 1 To Me.Count
                        buf.Add CStr(Me(i)) & CStr(target)
                    Next
                Case ArithmeticOperatorsEnum.lsSum
                    For i = 1 To Me.Count
                        buf.Add Val(Me(i)) + Val(target)
                    Next
                Case ArithmeticOperatorsEnum.lsDifference
                    For i = 1 To Me.Count
                        buf.Add Val(Me(i)) - Val(target)
                    Next
                Case ArithmeticOperatorsEnum.lsMultiply
                    For i = 1 To Me.Count
                        buf.Add Val(Me(i)) * Val(target)
                    Next
                Case ArithmeticOperatorsEnum.lsDivide
                     For i = 1 To Me.Count
                        buf.Add Val(Me(i)) / Val(target)
                    Next
                Case ArithmeticOperatorsEnum.lsMod
                     For i = 1 To Me.Count
                        buf.Add Val(Me(i)) Mod Val(target)
                    Next
                Case ArithmeticOperatorsEnum.lsExponent
                     For i = 1 To Me.Count
                        buf.Add Val(Me(i)) ^ Val(target)
                    Next
                Case Else
                    err.Raise Number:=1006, Description:="クラス：List、メソッド：Map、引数の型が適切ではありません。 "
            End Select
        End Select

    Set Map = buf

End Function

'''<summary> Listを連結した文字列で返します
'''<param> 連結する際に挿入する文字
'''<returns> 連結した文字列
Public Function Join(ByVal separator As String) As String

    Dim buf As String
    Dim Item As Variant
    For Each Item In Me
        If buf = "" Then
            buf = CStr(Item)
        Else
            buf = buf & separator & CStr(Item)
        End If
    Next
    Join = buf
    
End Function

'''<summary> Listを配列に変換します
'''<param> なし
'''<returns> 配列 Variant()型
'''<remarks> インデックス番号は1から始まります
Public Function ToArray() As Variant()
    
    Dim buf() As Variant
    ReDim buf(1 To Me.Count)
    Dim i As Long
    For i = 1 To Me.Count
        If VarType(Me.Item(i)) = vbObject Then
            Set buf(i) = Me.Item(i)
        Else
            buf(i) = Me.Item(i)
        End If
    Next
    ToArray = buf()

End Function

'''<summary> インデックス番号のStartIndex番目からEndIndex番目の要素を返す
'''<param> 先頭のインデックス番号、後尾のインデックス番号(省略した場合は最後まで
'''<returns> List型
Public Function Slice(ByVal StartIndex As Long, Optional ByVal EndIndex As Long = -1) As List

    If EndIndex = -1 Then EndIndex = Me.Count
    Dim buf As List
    Set buf = New List
    
    Dim i As Long
    For i = StartIndex To EndIndex
        buf.Add (Me(i))
    Next
    Set Slice = buf

End Function

'''<summary> イテレータ ※単独では使用できません
'''<param> なし
'''<returns> なし
Public Function NewEnum() As stdole.IUnknown
Attribute NewEnum.VB_UserMemId = -4
        Set NewEnum = Coll.[_NewEnum]
End Function
