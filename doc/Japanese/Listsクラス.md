

目次
<!-- @import "[TOC]" {cmd="toc" depthFrom=1 depthTo=6 orderedList=false} -->
<!-- code_chunk_output -->

- [Listsクラス](#listsクラス)
  - [メソッド](#メソッド)
    - [Add](#add)
    - [AddIndex](#addindex)
    - [AddTable](#addtable)
    - [AddValue](#addvalue)
    - [Concat](#concat)
    - [Count](#count)
    - [CountList](#countlist)
    - [IsKey](#iskey)
    - [Item](#item)
    - [KeyList](#keylist)
    - [Max](#max)
    - [Min](#min)
    - [Remove](#remove)
    - [ToWriteCells](#towritecells)
    - [Where](#where)

<!-- /code_chunk_output -->

# Listsクラス

List型のコレクションクラス
キーの設定が必須でディクショナリーのように使用することができる

## メソッド

### Add
引数： 追加するListもしくは要素、キー  
戻値： なし

* 要素を追加します。
* 要素はList型もしくはプリミティブ型のみ可能
* List型の場合はキーを重複させることはできません
```VB
Sub test()
    Dim lsts As New Lists
    Dim ls As New List
    ls.AddValue Array("foo", "bar")
    lsts.Add ls, "key1"
    lsts.Add "baz", "key1"
    Dim Str
    For Each Str In lsts("key1")
        Debug.Print Str
    Next
End Sub
'Print
foo
bar
baz
End Sub
```

### AddIndex
引数： Indexのキー名  
戻値： なし

* インデックス用のListを作成します

```VB
Sub test()
    Dim lsts As New Lists
    lsts.AddValue Array("foo", "bar", "baz"), "key1"
    lsts.AddIndex "index"
    Dim Str
    For Each Str In lsts("index")
        Debug.Print Str
    Next
End Sub
'Print
1
2
3
```


### AddTable
引数： テーブルの左上のセル(range型)[最終行、最終列]  
戻値： なし

* エクセルのみ使用可能
* テーブルの基準となるセルを指定して、最終行、最終列の値を指定する。
* 指定しない場合は、CurrentRegionのようにエクセルのテーブル取得する(ただし、基準となるセルより上、左のセルは取得しない)
* 2次配列になるような範囲でも1列ずつList型が作成される
* 一行目の値をキーとする

### AddValue
引数： 追加する要素  
戻値： なし

* 配列の値やRangeの値、オブジェクトの既定メンバーの値を格納する。
* 2次配列のような範囲にしてもList型は1つしか作成されない。項目ごとにList型を作成したい場合は`AddTable`メソッドの使用する
* Collection型ではないオブジェクトを渡すとエラーが返ります

```VB
Sub test()
    Dim lsts As New Lists
    lsts.AddValue Array("foo", "bar", "baz"), "key1"
    Dim Str
    For Each Str In lsts("key1")
        Debug.Print Str
    Next
End Sub
'Print
foo
bar
baz
```

### Concat
引数： 結合したいLists(複数可能)   
戻値： 結合したLists

* Lists同士を結合して新たにListsを作成する。
* 同じキーがある場合は要素が結合される

```VB
Sub test()
    Dim lsts As New Lists
    lsts.AddValue Array("foo", "bar", "baz"), "key1"
    lsts.AddValue Array("Alice", "Bob", "Charlie"), "key2"
    Dim lsts2 As New Lists
    lsts2.AddValue Array("spam", "ham", "eggs"), "key3"
    lsts2.AddValue Array("qux", "quux"), "key1"
    Dim newLists As New Lists
    Set newLists = lsts.Concat(lsts2)
    Dim lst
    Dim i
    For Each lst In newLists
        For Each i In lst
            Debug.Print i
        Next
    Next
End Sub
'Print
foo
bar
baz
qux
quux
Alice
Bob
Charlie
spam
ham
eggs
```

### Count
引数： なし  
戻値： 要素数

* コレクションの要素数を返します
* 格納されているList型の要素数を返すわけではありません

```VB
Sub test()
    Dim lsts As New Lists
    lsts.Add "foo", "key1"
    lsts.Add "bar", "key2"
    lsts.Add "baz", "key3"
    Debug.Print lsts.Count()
End Sub
'Print
3
```

### CountList
引数： キー  
戻値： 格納されているList型の要素数

* 格納されているList型の要素数を返します
* コレクションの要素数を返すわけではありません

```VB
Sub test()
    Dim lsts As New Lists
    lsts.Add "foo", "key1"
    lsts.Add "bar", "key2"
    lsts.Add "baz", "key3"
    Debug.Print lsts.CountList("key1")
End Sub
'Print
1
```

### IsKey
引数： 調べるキー  
戻値： 真偽値

* 引数のキーが存在するか確認する

```VB
Sub test()
    Dim lsts As New Lists
    lsts.Add "foo", "key1"
    lsts.Add "bar", "key2"
    lsts.Add "baz", "key3"
    Debug.Print lsts.IsKey("key1")
    Debug.Print lsts.IsKey("key4")
End Sub
'Print
True
False
```

### Item
引数： キー  
戻値： List型

* キーに関連付けしたListを返します。
* 既定メンバ

```VB
Sub test()
    Dim lsts As New Lists
    lsts.AddValue Array("foo", "bar", "baz"), "key1"
    Dim ls As New List
    Set ls = lsts.Item("key1")
    Dim Str
    For Each Str In ls
        Debug.Print Str
    Next
End Sub
'Print
foo
bar
baz
```

### KeyList
引数： キー  
戻値： List型

* 存在しているキーのListを返します。

```VB
Sub test()
    Dim lsts As New Lists
    lsts.AddValue Array("foo", "bar", "baz"), "key1"
    lsts.AddValue Array("Alice", "Bob", "Charlie"), "key2"
    Dim ls As New List
    Set ls = lsts.KeyList()
    Dim Str
    For Each Str In ls
        Debug.Print Str
    Next
End Sub
'Print
key1
key2
```

### Max
引数： 比べるキー(複数可)  
戻値： 最大値が集まったList

* 同じインデックス番号内で最大値を見つけて1つのListにする

```VB
Sub test()
    Dim lsts As New Lists
    lsts.AddValue Array(30, 50, 90), "key1"
    lsts.AddValue Array(70, 10, 30), "key2"
    lsts.AddValue Array(70, 80, 20), "key3"
    Dim ls As New List
    Set ls = lsts.Max("key1", "key2", "key3")
    Dim Str
    For Each Str In ls
        Debug.Print Str
    Next
End Sub
'Print
70 
80 
90 
```

### Min
引数： 比べるキー(複数可)  
戻値： 最小値が集まったList

* 同じインデックス番号内で最小値を見つけて1つのListにする

```VB
Sub test()
    Dim lsts As New Lists
    lsts.AddValue Array(30, 50, 90), "key1"
    lsts.AddValue Array(70, 10, 30), "key2"
    lsts.AddValue Array(70, 80, 20), "key3"
    Dim ls As New List
    Set ls = lsts.Min("key1", "key2", "key3")
    Dim Str
    For Each Str In ls
        Debug.Print Str
    Next
End Sub
'Print
30 
10 
20 
```

### Remove
引数： キー   
戻値： なし

* Listを削除する。

```VB
Sub test()
    Dim lsts As New Lists
    lsts.AddValue Array(30, 50, 90), "key1"
    lsts.AddValue Array(70, 10, 30), "key2"
    lsts.AddValue Array(70, 80, 20), "key3"
    lsts.Remove ("key2")
    Dim ls As New List
    Set ls = lsts.KeyList()
    Dim Str
    For Each Str In ls
        Debug.Print Str
    Next
End Sub
'Print
key1
key3
```
### ToWriteCells
引数： 書き込む基点のセル、書き込むキー(複数選択可)    
戻値： なし

* 現在のListsを引数のセルから基点にセルを書き込む
* Excel以外では使えないので削除すること

### Where
引数： 比較演算子の列挙型、比較対象(List型、Collection型、プリミティブ型に対応)、基準となるキー  
戻値： Lists型

* 条件にあう要素のみを残し、新たにListsを作成する
* VBAではアロー演算子が使えないため以下の列挙型で表現する

**ComparisonOperatorsEnum**
|  要素名           |  説明  |
| ------------------| --------|
|  lsEqual          |  等しい(=)|
|  lsNotEqual       |  等しくない(<>)|
|  lsGreater        |  超える(>)|
|  lsLess           |  未満(<)|
|  lsGreaterEqual   |  以上(>=)|
|  lsLessEqual      |  以下(<=)|
|  lsObjectEqual    |  参照比較(Is)|
|  lsLike           |  文字列比較(Like)|
|  lsNotLike        |  文字列比較の否定(Not Like)|


```VB
Sub test()
    Dim lsts As New Lists
    lsts.AddValue Array(30, 50, 90), "key1"
    lsts.AddValue Array(70, 10, 30), "key2"
    lsts.AddValue Array(70, 80, 20), "key3"
    Dim newLsts As New Lists
    Set newLsts = lsts.Where(lsLess, 50, "key2")
    Dim ls As New List
    Set ls = newLsts("key1")
    Dim Str
    For Each Str In ls
        Debug.Print Str
    Next
End Sub
'Print
50 
90
```
