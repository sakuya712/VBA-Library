VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Lists"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'''<classSummary> List型のコレクションクラス、キーの設定が必須です
'''<updateDate> 2020/9/14
'''<version> Ver.100
'''<remarks> List型が必要です

Private ListColl As Collection
Private KeyLst As List

'''<summary> 自分を返すためのもの
'''<param> なし
'''<returns> 自己参照
Public Property Get self() As Lists
    Set self = Me
End Property

'''<summary> オブジェクトをセット
'''<param> なし
'''<returns> なし
Private Sub Class_Initialize()
    Set ListColl = New Collection
    Set KeyLst = New List
End Sub

'''<summary> Listを追加する
'''<param> 追加するListもしくは要素、キー
'''<returns> なし
Public Function Add(ByVal Item As Variant, ByVal Key As String)

    If TypeName(Item) = "List" Then
        On Error GoTo err1
        ListColl.Add Item, Key
        KeyLst.Add Key
        Exit Function
    End If

    If Me.IsKey(Key) Then
        ListColl(Key).Add Item
    Else
        Dim newList As New List
        newList.Add Item
        ListColl.Add newList, Key
        KeyLst.Add Key
    End If
    Exit Function
err1:
        err.Raise Number:=1051, Description:="クラス：Lists、メソッド：Add、キーが重複しています。"
End Function

'''<summary> 配列の値、Rangeの値、オブジェクトの既定メンバーの値を格納する
'''<param> 追加する要素(Rangeや配列)、キー
'''<returns> なし
Public Function AddValue(ByVal Items As Variant, ByVal Key As String)
    On Error GoTo err
    Dim Coll As New List
    Dim i As Variant
    Dim ItemArr() As Variant
    If TypeName(Items) = "Range" Then
        If Items.Count = 1 Then
            Coll.Add Items.Value
            Me.Add Coll, Key
            Exit Function
        End If
        ItemArr() = Items.Value
        For Each i In ItemArr
            Coll.Add i
        Next
        Me.Add Coll, Key
        Exit Function
    End If
    For Each i In Items
        Coll.Add i
    Next
    Me.Add Coll, Key
    Exit Function
err:
    err.Raise Number:=1052, Description:="クラス：Lists、メソッド：AddValue、引数の型が適切ではありません"
    
End Function

'''<summary> テーブル(Range型)を読み込んで列ごとにListにする
'''<param> 追加する要素(Range)、[最終行、最終列]
'''<returns> なし
'''<remarks> 一番最初の行をキーとする
Public Function AddTable(ByVal Cell As Range, Optional ByVal RowBottom As Long = -1, Optional ByVal ColRight As Long = -1)
    If RowBottom < 1 Then RowBottom = Cell.End(xlDown).Row
    If ColRight < 1 Then ColRight = Cell.End(xlToRight).Column
    Dim cell2 As Range
    Set cell2 = Cell.Parent.Cells(RowBottom, ColRight)
    Dim CurrentRegion As Range
    Set CurrentRegion = Range(Cell, cell2)
    Dim ItemArr() As Variant
    ItemArr() = CurrentRegion.Value
    Dim Coll As List
    Dim Col As Long
    Dim Row As Long
    Dim Key As String
    For Col = LBound(ItemArr, 2) To UBound(ItemArr, 2)
        Set Coll = New List
        Key = CStr(ItemArr(1, Col))
        For Row = LBound(ItemArr, 1) + 1 To UBound(ItemArr, 1)
            Coll.Add ItemArr(Row, Col)
        Next Row
        Me.Add Coll, Key
    Next Col
    Exit Function
    
End Function

'''<summary> このListsの要素数を返す
'''<param> なし
'''<returns> 要素数
Public Function Count() As Long
    Count = ListColl.Count
End Function

'''<summary> 各Listの要素数を返す
'''<param> キー
'''<returns> 要素数
Public Function CountList(ByVal Key As String) As Long
    CountList = ListColl(Key).Count
End Function

'''<summary> 要素の内容を返す
'''<param> キー
'''<returns> 内容
'''<remarks> 既定メンバ
Public Function Item(ByVal Key As String) As List
Attribute Item.VB_UserMemId = 0
    On Error GoTo err1
    Set Item = ListColl(Key)
    Exit Function
err1:
    err.Raise Number:=1053, Description:="クラス：Lists、メソッド：Item、存在しないキーが選ばれました"
End Function

'''<summary> 要素数を削除する
'''<param> 削除するキー
'''<returns> なし
Public Function Remove(ByVal Key As String)
    On Error GoTo err1
    ListColl.Remove (Key)
    Dim i As Long
    For i = 1 To KeyLst.Count
        If KeyLst(i) = Key Then
            KeyLst.Remove (i)
            Exit Function
        End If
    Next i
    Exit Function
err1:
    Debug.Print "クラス：Lists、メソッド：Remove、存在しないキーが選ばれました"
End Function

'''<summary> キーの一覧を返します
'''<param> なし
'''<returns> キーの一覧(List型)
Public Function KeyList() As List
    Set KeyList = KeyLst
End Function

'''<summary> キーが存在するか調べる
'''<param> 調べたいキー
'''<returns> 真偽値
Public Function IsKey(ByVal Key As String) As Boolean
    Dim k As Variant
    For Each k In KeyLst
        If k = Key Then
            IsKey = True
            Exit Function
        End If
    Next
    IsKey = False
End Function

'''<summary> Listをセルに書き込む
'''<param> 書き込むセル範囲、キー(複数可)
'''<returns> なし
'''<remarks> キーの順番通りに列を追記します
Public Function ToWriteCells(ByVal WriteCells As Range, ParamArray Keys() As Variant)
    Dim KeyArr() As Variant
    If UBound(Keys) = -1 Then
        ReDim KeyArr(0 To KeyLst.Count - 1)
        Dim i As Long
        For i = 0 To KeyLst.Count - 1
            KeyArr(i) = KeyLst(i + 1)
        Next i
    Else
        KeyArr() = Keys()
    End If
    Dim Rows As Long
    Rows = ListColl(KeyArr(0)).Count - 1
    Dim WriteRow As Long
    Dim WriteCol As Long
    WriteRow = WriteCells.Row
    WriteCol = WriteCells.Column
    Dim Key As Variant
    Dim Cell1 As Range
    Dim cell2 As Range
    On Error GoTo err1
    For Each Key In KeyArr
        Set Cell1 = WriteCells.Parent.Cells(WriteRow, WriteCol)
        Set cell2 = WriteCells.Parent.Cells(WriteRow + Rows, WriteCol)
        ListColl(Key).ToWriteCells Range(Cell1, cell2)
        WriteCol = WriteCol + 1
    Next Key
    Exit Function
err1:
    err.Raise Number:=1059, Description:="クラス：Lists、メソッド：ToWriteCells、Listの要素数より大きい範囲を指定しています"
End Function

'''<summary> 同じインデックス番号内で最大値を見つけてListにする
'''<param> 調べたいキー(複数)
'''<returns> 最大値が集まったList
Public Function Max(ParamArray Keys() As Variant) As List
    Dim compare As List
    Dim newList As New List
    Dim Key As Variant
    Dim Index As Long
    Dim Result As Double
    For Index = 1 To Me.CountList(Keys(0))
        Set compare = New List
        For Each Key In Keys()
            compare.Add Me(Key)(Index)
        Next Key
        Result = compare.Aggregate(lsMax)
        newList.Add Result
    Next Index
    Set Max = newList
End Function

'''<summary> 同じインデックス番号内で最小値を見つけてListにする
'''<param> 調べたいキー(複数)
'''<returns> 最大値が集まったList
Public Function Min(ParamArray Keys() As Variant) As List
    Dim compare As List
    Dim newList As New List
    Dim Key As Variant
    Dim Index As Long
    Dim Result As Double
    For Index = 1 To Me.CountList(Keys(0))
        Set compare = New List
        For Each Key In Keys()
            compare.Add Me(Key)(Index)
        Next Key
        Result = compare.Aggregate(lsMin)
        newList.Add Result
    Next Index
    Set Min = newList
End Function

'''<summary> あるキーの比較して条件と一致したインデックス番号のみを残したLists型を作成する
'''<param>  比較演算子の列挙型、比較対象(List型、Collection型、プリミティブ型に対応、比較対象キー)
'''<returns> Lists型
Public Function Where(ByVal operator As ComparisonOperatorsEnum, ByVal compare As Variant, ByVal Key As String) As Lists

    Dim newLists As New Lists
    Dim newList As New List
    Dim indexList As New List
    Dim i As Long
    Dim var1, var2
    Dim Index As Variant
    Dim k As Variant
    Select Case TypeName(compare)
        Case "List", "Collection"
            Select Case operator
                Case ComparisonOperatorsEnum.lsEqual
                    For i = 1 To Me.CountList(Key)
                        var1 = Me(Key)(i)
                        var2 = compare(i)
                        If var1 = var2 Then
                            indexList.Add i
                        End If
                    Next
                    For Each k In Me.KeyList
                        Set newList = New List
                        For Each Index In indexList
                            newList.Add Me(k)(Index)
                        Next Index
                        newLists.Add newList, k
                    Next k
                    Set Where = newLists
                    Exit Function
                Case ComparisonOperatorsEnum.lsNotEqual
                    For i = 1 To Me.CountList(Key)
                        var1 = Me(Key)(i)
                        var2 = compare(i)
                        If var1 <> var2 Then
                            indexList.Add i
                        End If
                    Next
                    For Each k In Me.KeyList
                        Set newList = New List
                        For Each Index In indexList
                            newList.Add Me(k)(Index)
                        Next Index
                        newLists.Add newList, k
                    Next k
                    Set Where = newLists
                    Exit Function
                Case ComparisonOperatorsEnum.lsLike
                    For i = 1 To Me.CountList(Key)
                        var1 = Me(Key)(i)
                        var2 = compare(i)
                        If var1 Like var2 Then
                            indexList.Add i
                        End If
                    Next
                    For Each k In Me.KeyList
                        Set newList = New List
                        For Each Index In indexList
                            newList.Add Me(k)(Index)
                        Next Index
                        newLists.Add newList, k
                    Next k
                    Set Where = newLists
                    Exit Function
                Case ComparisonOperatorsEnum.lsNotLike
                    For i = 1 To Me.CountList(Key)
                        var1 = Me(Key)(i)
                        var2 = compare(i)
                        If Not (var1 Like var2) Then
                            indexList.Add i
                        End If
                    Next
                    For Each k In Me.KeyList
                        Set newList = New List
                        For Each Index In indexList
                            newList.Add Me(k)(Index)
                        Next Index
                        newLists.Add newList, k
                    Next k
                    Set Where = newLists
                    Exit Function
                Case ComparisonOperatorsEnum.lsGreater
                    For i = 1 To Me.CountList(Key)
                        var1 = Me(Key)(i)
                        var2 = compare(i)
                        If var1 > var2 Then
                            indexList.Add i
                        End If
                    Next
                    For Each k In Me.KeyList
                        Set newList = New List
                        For Each Index In indexList
                            newList.Add Me(k)(Index)
                        Next Index
                        newLists.Add newList, k
                    Next k
                    Set Where = newLists
                    Exit Function
                Case ComparisonOperatorsEnum.lsLess
                    For i = 1 To Me.CountList(Key)
                        var1 = Me(Key)(i)
                        var2 = compare(i)
                        If var1 < var2 Then
                            indexList.Add i
                        End If
                    Next
                    For Each k In Me.KeyList
                        Set newList = New List
                        For Each Index In indexList
                            newList.Add Me(k)(Index)
                        Next Index
                        newLists.Add newList, k
                    Next k
                    Set Where = newLists
                    Exit Function
                Case ComparisonOperatorsEnum.lsGreaterEqual
                    For i = 1 To Me.CountList(Key)
                        var1 = Me(Key)(i)
                        var2 = compare(i)
                        If var1 >= var2 Then
                            indexList.Add i
                        End If
                    Next
                    For Each k In Me.KeyList
                        Set newList = New List
                        For Each Index In indexList
                            newList.Add Me(k)(Index)
                        Next Index
                        newLists.Add newList, k
                    Next k
                    Set Where = newLists
                    Exit Function
                Case ComparisonOperatorsEnum.lsLessEqual
                    For i = 1 To Me.CountList(Key)
                        var1 = Me(Key)(i)
                        var2 = compare(i)
                        If var1 <= var2 Then
                            indexList.Add i
                        End If
                    Next
                    For Each k In Me.KeyList
                        Set newList = New List
                        For Each Index In indexList
                            newList.Add Me(k)(Index)
                        Next Index
                        newLists.Add newList, k
                    Next k
                    Set Where = newLists
                    Exit Function
                Case ComparisonOperatorsEnum.lsObjectEqual
                    For i = 1 To Me.CountList(Key)
                        var1 = Me(Key)(i)
                        var2 = compare(i)
                        If var1 Is var2 Then
                            indexList.Add i
                        End If
                    Next
                    For Each k In Me.KeyList
                        Set newList = New List
                        For Each Index In indexList
                            newList.Add Me(k)(Index)
                        Next Index
                        newLists.Add newList, k
                    Next k
                    Set Where = newLists
                    Exit Function
                Case Else
                    err.Raise Number:=1005, Description:="クラス：Lists、メソッド：Where、引数の型が適切ではありません。 "
            End Select
        Case Else
             Select Case operator
                Case ComparisonOperatorsEnum.lsEqual
                    For i = 1 To Me.CountList(Key)
                        If Me(Key)(i) = compare Then
                            indexList.Add i
                        End If
                    Next
                    For Each k In Me.KeyList
                        Set newList = New List
                        For Each Index In indexList
                            newList.Add Me(k)(Index)
                        Next Index
                        newLists.Add newList, k
                    Next k
                    Set Where = newLists
                    Exit Function
                Case ComparisonOperatorsEnum.lsNotEqual
                    For i = 1 To Me.CountList(Key)
                        If Me(Key)(i) <> compare Then
                            indexList.Add i
                        End If
                    Next
                    For Each k In Me.KeyList
                        Set newList = New List
                        For Each Index In indexList
                            newList.Add Me(k)(Index)
                        Next Index
                        newLists.Add newList, k
                    Next k
                    Set Where = newLists
                    Exit Function
                Case ComparisonOperatorsEnum.lsLike
                    For i = 1 To Me.CountList(Key)
                        If Me(Key)(i) Like compare Then
                            indexList.Add i
                        End If
                    Next
                    For Each k In Me.KeyList
                        Set newList = New List
                        For Each Index In indexList
                            newList.Add Me(k)(Index)
                        Next Index
                        newLists.Add newList, k
                    Next k
                    Set Where = newLists
                    Exit Function
                Case ComparisonOperatorsEnum.lsNotLike
                    For i = 1 To Me.CountList(Key)
                        If Not (Me(Key)(i) Like compare) Then
                            indexList.Add i
                        End If
                    Next
                    For Each k In Me.KeyList
                        Set newList = New List
                        For Each Index In indexList
                            newList.Add Me(k)(Index)
                        Next Index
                        newLists.Add newList, k
                    Next k
                    Set Where = newLists
                    Exit Function
                Case ComparisonOperatorsEnum.lsGreater
                    For i = 1 To Me.CountList(Key)
                        If Me(Key)(i) > compare Then
                            indexList.Add i
                        End If
                    Next
                    For Each k In Me.KeyList
                        Set newList = New List
                        For Each Index In indexList
                            newList.Add Me(k)(Index)
                        Next Index
                        newLists.Add newList, k
                    Next k
                    Set Where = newLists
                    Exit Function
                Case ComparisonOperatorsEnum.lsLess
                    For i = 1 To Me.CountList(Key)
                        If Me(Key)(i) < compare Then
                            indexList.Add i
                        End If
                    Next
                    For Each k In Me.KeyList
                        Set newList = New List
                        For Each Index In indexList
                            newList.Add Me(k)(Index)
                        Next Index
                        newLists.Add newList, k
                    Next k
                    Set Where = newLists
                    Exit Function
                Case ComparisonOperatorsEnum.lsGreaterEqual
                    For i = 1 To Me.CountList(Key)
                        If Me(Key)(i) >= compare Then
                            indexList.Add i
                        End If
                    Next
                    For Each k In Me.KeyList
                        Set newList = New List
                        For Each Index In indexList
                            newList.Add Me(k)(Index)
                        Next Index
                        newLists.Add newList, k
                    Next k
                    Set Where = newLists
                    Exit Function
                Case ComparisonOperatorsEnum.lsLessEqual
                    For i = 1 To Me.CountList(Key)
                        If Me(Key)(i) <= compare Then
                            indexList.Add i
                        End If
                    Next
                    For Each k In Me.KeyList
                        Set newList = New List
                        For Each Index In indexList
                            newList.Add Me(k)(Index)
                        Next Index
                        newLists.Add newList, k
                    Next k
                    Set Where = newLists
                    Exit Function
                Case ComparisonOperatorsEnum.lsObjectEqual
                    For i = 1 To Me.CountList(Key)
                        If Me(Key)(i) Is compare Then
                            indexList.Add i
                        End If
                    Next
                    For Each k In Me.KeyList
                        Set newList = New List
                        For Each Index In indexList
                            newList.Add Me(k)(Index)
                        Next Index
                        newLists.Add newList, k
                    Next k
                    Set Where = newLists
                    Exit Function
                Case Else
                    err.Raise Number:=1005, Description:="クラス：Lists、メソッド：Where、引数の型が適切ではありません。 "
            End Select
        End Select
End Function

'''<summary> Lists同士を結合して新たにListを作成する
'''<param> 結合したいLists
'''<returns> 結合したLists
Public Function Concat(ParamArray Lists() As Variant) As Lists
    Dim ele As Variant
    For Each ele In Lists
        If ele Is Me Then err.Raise Number:=1035, Description:="クラス：Lists、メソッド：Concat、自分自身を引数にすることはできません"
    Next ele
    Dim buf As Lists
    Set buf = New Lists
    Dim key As String
    Dim i As Long
    For i = 1 To Me.Count
        key = Me.KeyList(i)
        buf.Add Me(key), key
    Next i
    Dim newLists As New Lists
    Dim newList As New List
    Dim j As Long
    For i = LBound(Lists) To UBound(Lists)
        Set newLists = Lists(i)
        For j = 1 To newLists.Count
            key = newLists.KeyList(j)
            Set newList = newLists(key)
            For Each ele In newList
                buf.Add ele, key
            Next ele
        Next j
    Next i
    Set Concat = buf
End Function

'''<summary> Index用のListをListsに追加する
'''<param> Indexのキー名
'''<returns> なし
Public Function AddIndex(ByVal IndexKey As String)
    Dim key As Variant
    Dim cnt As Long
    For Each key In KeyLst
        if cnt < Me.CountList(key) Then
            cnt = Me.CountList(key)
        End If
    Next key
    Dim ls As new List
    Dim i As Long
    For i = 1 to cnt
        ls.Add i
    Next i
    Me.Add ls, IndexKey
End Function

'''<summary> イテレータ ※単独では使用できません
'''<param> なし
'''<returns> なし
Public Function NewEnum() As stdole.IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = ListColl.[_NewEnum]
End Function

