

目次
<!-- @import "[TOC]" {cmd="toc" depthFrom=1 depthTo=6 orderedList=false} -->
<!-- code_chunk_output -->

- [Listクラス](#listクラス)
  - [メソッド](#メソッド)
    - [Add](#add)
    - [AddValue](#addvalue)
    - [Aggregate](#aggregate)
    - [Callback](#callback)
    - [Concat](#concat)
    - [Count](#count)
    - [Except](#except)
    - [Includes](#includes)
    - [IndexOf](#indexof)
    - [IsOverlap](#isoverlap)
    - [Item](#item)
    - [Join](#join)
    - [Lstrip](#lstrip)
    - [Map](#map)
    - [OverlapList](#overlaplist)
    - [Remove](#remove)
    - [RemoveEmpty](#removeempty)
    - [Replace](#replace)
    - [Rstrip](#rstrip)
    - [Slice](#slice)
    - [Split](#split)
    - [ToArray](#toarray)
    - [ToList](#tolist)
    - [ToWriteCells](#towritecells)
    - [TypeConversion](#typeconversion)
    - [Unique](#unique)
    - [Where](#where)

<!-- /code_chunk_output -->

# Listクラス

他の言語にあるようなList型を再現したクラス   
中身はCollection型を拡張したもの  

## メソッド

### Add
引数： 追加する要素  
戻値： なし

* 要素を追加します。Collection型と同じ
* 配列やオブジェクトもそのままコレクションとして格納します。
```VB
Sub test()
    Dim ls As New List
    ls.Add "foo"
    ls.Add "bar"
    ls.Add "baz"
    Dim str
    For Each str In ls
        Debug.Print str
    Next
End Sub
'Print
foo
bar
baz
```

### AddValue
引数： 追加する要素  
戻値： なし

* 配列の値やRangeの値、オブジェクトの既定メンバーの値を格納します。
* Collection型ではないオブジェクトを渡すとエラーが返ります。

```VB
Sub test()
    Dim ls As New List
    ls.AddValue Array("foo", "bar", "baz")
    Dim str
    For Each str In ls
        Debug.Print str
    Next
End Sub
'Print
foo
bar
baz
```

### Aggregate
引数： 集計内容の列挙型、[何番目か(Large、Smallのみ)]  
戻値： Double型

* Listにある数値を計算する。  
* `lsLarge`、`lsSmall`の場合は、2つ目の引数で何番目かを指定する
* VBAではアロー演算子が使えないため以下の列挙型で表現する。

**AggregateEnum**
|  要素名          |  説明  |
| --------------- | --------|
| lsAverage | 平均 |
| lsCount | 要素数(数値のみ) |
| lsCountA | 要素数(空白除く) |
| lsMax | 最大値 |
| lsMin | 最小値 |
| lsProduct | 積 |
| lsTotal | 合計 |
| lsMedian | 中央値 |
| lsLarge | n番目に大きい数値 |
| lsSmall | n番目に小さい数値 |

```VB
Sub test()
    Dim ls As New List
    ls.AddValue Array(10, 60, 30, 40)
    Debug.Print ls.Aggregate(lsMax)
    Debug.Print ls.Aggregate(lsLarge, 2)
End Sub
'Print
60 
40 
```

### Callback
引数： メソッド名、[メソッドの引数]
戻値： List型

* `ListFunc`クラスにメソッドを作成し、そのメソッドの結果を返す
* アロー演算子のようにListにある値とインデックス番号は自動的に引数として渡される
* メソッドの引数は必須ではありません
* `ListFunc`クラスのメソッドの戻値の型は`Varinat`型であること

\<例>リストから35未満の値のみを取り出す
```VB
'Module
Sub test()
    Dim ls As New List
    ls.AddValue Array(10, 60, 30, 40)
    Dim newLs As New List
    Set newLs = ls.Callback("Under", 35)
    Dim i
    For Each i In newLs
        Debug.Print i
    Next
End Sub
'ListFunc Class
Public Function Under(ByVal Args As Variant) As Variant 'Varinat Fixed
    Dim Value As Variant
    Dim Index As Long
    Value = Args(0)    'Arg(0) Fixed value, List value
    Index = Args(1)    'Arg(1) Fixed value, List index
    Dim num As Long
    num = Args(2)
    If Value < num Then
        Under = Value
    End If
End Function
'Print
10 
30 
```

### Concat
引数： 結合したいList(複数可能)   
戻値： 結合したList

* List同士を結合して新たにListを作成する。
* JSのConcatと同じ

```VB
Sub test()
    Dim ls As New List
    ls.AddValue Array(10, 20)
    Dim ls2 As New List
    ls2.AddValue Array("foo", "bar")
    Dim ls3 As New List
    ls3.AddValue Array("Alice", "Bob")
    Dim newLs As New List
    Set newLs = ls.Concat(ls2, ls3)
    Dim i
    For Each i In newLs
        Debug.Print i
    Next
End Sub
'Print
10 
20 
foo
bar
Alice
Bob
```

### Count
引数： なし  
戻値： 要素数

* 要素数を返します。Collection型と同じ
```VB
Sub test()
    Dim ls As New List
    ls.AddValue Array(10, 20)
    Debug.Print ls.Count()
End Sub
'Print
2
```

### Except
引数： 取り除きたい要素のList  
戻値： 取り除かれたList

* 差集合のListを作成する。
* SQLやC#のExceptのように動作する。

```VB
Sub test()
    Dim ls As New List
    ls.AddValue Array("foo", "bar", "baz")
    Dim ls2 As New List
    ls2.AddValue Array("bar", "qux")
    Dim newLs As New List
    Set newLs = ls.Except(ls2)
    Dim i
    For Each i In newLs
        Debug.Print i
    Next
End Sub
'Print
foo
baz
```

### Includes
引数： 文字列、[完全一致かどうか]   
戻値： 真偽値

*  Listの中に引数とマッチする要素の有無を真偽値で返す。
*  JSのIncludesのように動作する。
*  2つ目の引数で完全一致か、部分一致かを設定できます。既定値は完全一致です。
```VB
Sub test()
    Dim ls As New List
    ls.AddValue Array("foo", "bar", "baz")
    Debug.Print ls.Includes("qux")
    Debug.Print ls.Includes("ba", False)
End Sub
'Print
False
True
```

### IndexOf
引数： 文字列、[完全一致かどうか]   
戻値： インデックス番号
  
* Listの中に引数とマッチする要素が最初に出現するインデックスを返します。
*   JSやC#のIndexOfのように動作する。
*  2つ目の引数で完全一致か、部分一致かを設定できます。既定値は完全一致です。
*  見つからない場合は-1を返します。

```VB
Sub test()
    Dim ls As New List
    ls.AddValue Array("foo", "bar", "baz")
    Debug.Print ls.IndexOf("qux")
    Debug.Print ls.IndexOf("ba", False)
End Sub
'Print
-1 
 2 
```

### IsOverlap
引数： なし  
戻値： 真偽値

* 要素に重複があるか調べる。重複がある場合はTrueを返す。

```VB
Sub test()
    Dim ls As New List
    ls.AddValue Array("foo", "bar", "baz")
    Debug.Print ls.IsOverlap()
    ls.Add "foo"
    Debug.Print ls.IsOverlap()
End Sub
'Print
False
True
```

### Item
引数： なし  
戻値： 要素

* 要素を返します。Collection型と同じ。
* 既定メンバ

```VB
Sub test()
    Dim ls As New List
    ls.AddValue Array("foo", "bar", "baz")
    Debug.Print ls.Item(1)
    Debug.Print ls(2) 'default member
End Sub
'Print
foo
bar
```

### Join
引数： 連結する際に挿入する文字  
戻値： 連結した文字列

* Listにある要素を1つの連結した文字列で返します。
* VBAの配列で使えるJoinと同じように動作する。
* 文字列に変換できないオブジェクトの場合、エラーになります
```VB
Sub test()
    Dim ls As New List
    ls.AddValue Array("foo", "bar", "baz")
    Debug.Print ls.Join(",")
End Sub
'Print
foo,bar,baz
```

### Lstrip
引数： 基準となる文字列  
戻値： List型

* 先頭から引数の文字までを除去する
* 引数の文字がない場合は何もしない
* 引数の文字が複数あったとしても1番先頭しか処理しない
* Pythonのlstripと同じように動作する

```VB
Sub test()
    Dim ls As New List
    ls.AddValue Array("foo", "bar", "baz")
    Dim newLs As New List
    Set newLs = ls.Lstrip("a")
    Dim i
    For Each i In newLs
        Debug.Print i
    Next
End Sub
'Print
foo
r
z
```

### Map
引数： 算術演算子の列挙型、算術対象(List型、Collection型、プリミティブ型に対応)  
戻値： List型

* 要素の中身を計算や結合し、新たにListを作成する
* C#のSelectと同じように動作する。
* VBAではアロー演算子が使えないため以下の列挙型で表現する

**ArithmeticOperatorsEnum**
|  要素名          |  説明  |
| --------------- | --------|
|  lsSum          |  足し算(+)|
|  lsDifference   |  引き算(-)|
|  lsMultiply     |  掛け算(*)|
|  lsDivide       |  割り算(/)|
|  lsMod          |  割り算の余り(mod)|
|  lsExponent     |  乗数(^)|
|  lsConcatenate  |  文字列結合(&)|

```VB
Sub test()
    Dim ls As New List
    ls.AddValue Array("foo", "bar", "baz")
    Dim newLs As New List
    Set newLs = ls.Map(lsConcatenate, "100")
    Dim i
    For Each i In newLs
        Debug.Print i
    Next
End Sub
'Print
foo100
bar100
baz100
```

### OverlapList
引数： なし  
戻値： 重複したList

* 重複した要素を取り出したListを作成する。
* 重複した要素同士は重複しない

```VB
Sub test()
    Dim ls As New List
    ls.AddValue Array("foo", "bar", "baz", "foo")
    Dim newLs As New List
    Set newLs = ls.OverlapList()
    Dim i
    For Each i In newLs
        Debug.Print i
    Next
End Sub
'Print
foo
```

### Remove
引数： インデックス番号   
戻値： なし

* 要素を削除する。Collection型と同じ
```VB
Sub test()
    Dim ls As New List
    ls.AddValue Array("foo", "bar", "baz")
    Dim newLs As New List
    ls.Remove (2)
    Dim i
    For Each i In ls
        Debug.Print i
    Next
End Sub
'Print
foo
baz
```

### RemoveEmpty
引数： なし   
戻値： List型

* 空文字(vbNullString)、Empty、Nullを取り除いたListを作成する

```VB
Sub test()
    Dim ls As New List
    ls.AddValue Array("", "foo", "bar", Null, "baz", Empty)
    Dim newLs As New List
    Set newLs = ls.RemoveEmpty()
    Dim i
    For Each i In newLs
        Debug.Print i
    Next
End Sub
'Print
foo
bar
baz
```

### Replace
引数： 置換対象、置換後の文字   
戻値： List型

* Listにある文字列を置換する
* VBAのReplace関数と処理は同じ
* 置換対象が複数ある場合、すべて置換を行います

```VB
Sub test()
    Dim ls As New List
    ls.AddValue Array("foo", "bar", "baz")
    Dim newLs As New List
    Set newLs = ls.Replace("a", "A")
    Dim i
    For Each i In newLs
        Debug.Print i
    Next
End Sub
'Print
foo
bAr
bAz
```

### Rstrip
引数： 基準となる文字列  
戻値： List型

* 後尾から引数の文字までを除去する
* 引数の文字がない場合は何もしない
* 引数の文字が複数あったとしても1番後尾しか処理しない
* Pythonのrstripと同じように動作する

```VB
Sub test()
    Dim ls As New List
    ls.AddValue Array("foo", "bar", "baz")
    Dim newLs As New List
    Set newLs = ls.Rstrip("a")
    Dim i
    For Each i In newLs
        Debug.Print i
    Next
End Sub
'Print
foo
b
b
```

### Slice
引数： 先頭のインデックス番号、[後尾のインデックス番号]    
戻値： 要素を切り取ったList型

* インデックス番号のStartIndex番目からEndIndex番目の要素を返す
* JSのSliceのように動作する。
* 2つ目の引数を省略した場合は最後のインデックス番号まで指定される

```VB
Sub test()
    Dim ls As New List
    ls.AddValue Array("foo", "bar", "baz", "qux")
    Dim newLs As New List
    Set newLs = ls.Slice(2, 3)
    Dim i
    For Each i In newLs
        Debug.Print i
    Next
End Sub
'Print
bar
baz
```

### Split
引数： 区切り文字   
戻値： Lists型

* 引数の文字で区切りListを作成する
* VBAのSplit関数と処理は同じ(ただし、戻り値はLists型)
* Lists型=>List型=>要素のように格納されている。詳細はListsクラスの説明を参照

```VB
Sub test()
    Dim ls As New List
    ls.AddValue Array("foo", "bar", "baz")
    Dim newLists As New Lists
    Set newLists = ls.Split("a")
    Dim lst
    Dim i
    For Each lst In newLists  'lst = List Class
        For Each i In lst
            Debug.Print i
        Next
    Next
End Sub
'Print
foo
b
r
b
z
```

### ToArray
引数： なし    
戻値： 1次配列

* Listを1次配列へ変換する
* インデックス番号は1から始まる

```VB
Sub test()
    Dim ls As New List
    ls.AddValue Array("foo", "bar", "baz")
    Dim Arr() As Variant
    Arr() = ls.ToArray()  'List => Array
End Sub
```

### ToList
引数： Listに格納する要素    
戻値： List型

* 同じ要素数のListを作成する

```VB
Sub test()
    Dim ls As New List
    ls.AddValue Array("foo", "bar", "baz")
    Dim newLs As New List
    Set newLs = ls.ToList("Alice")
    Dim i
    For Each i In newLs
        Debug.Print i
    Next
End Sub
'Print
Alice
Alice
Alice
```

### ToWriteCells
引数： 書き込む範囲    
戻値： なし

* 現在のListを引数の範囲に書き込む
* 2次配列になるようにセル範囲を指定するとエラーになる
* Excel以外では使えないので削除すること

### TypeConversion
引数： データ型変換の列挙型    
戻値： List型

* Listにある要素のデータ型を変換する
* VBAではアロー演算子が使えないため以下の列挙型で表現する


**TypeConversionEnum**
|  要素名          |  説明  |
| --------------- | --------|
| lsBoolean | Boolean型に変換(Falseと0以外はTrue) |
| lsByte | Byte型に変換 (0~255のみ) |
| lsCurrency | Currency型に変換(数値のみ) |
| lsDate | Date型に変換(数値のみ) |
| lsDouble | Double型に変換(数値のみ) |
| lsDecimal | Decimal型に変換(数値のみ) |
| lsLong | Long型に変換(数値のみ) |
| lsString | String型に変換 |
| lsVariant | Variant型に変換 |
| lsVal | 強制的に数値型に変換 *1 |
*1 主にDouble型に変換、VBAのVal関数と同じ処理

```VB
Sub test()
    Dim ls As New List
    ls.AddValue Array(44033, 43633)
    Dim newLs As New List
    Set newLs = ls.TypeConversion(lsDate)
    Dim i
    For Each i In newLs
        Debug.Print i
    Next
End Sub
'Print
2020/07/21 
2019/06/17 
```

### Unique
引数： なし  
戻値： 重複しないList

* 重複しないリストを作成する。
* C#のDistinctと同じように動作する。

```VB
Sub test()
    Dim ls As New List
    ls.AddValue Array("foo", "bar", "baz", "foo")
    Dim newLs As New List
    Set newLs = ls.Unique()
    Dim i
    For Each i In newLs
        Debug.Print i
    Next
End Sub
'Print
foo
bar
baz
```

### Where
引数： 比較演算子の列挙型、比較対象(List型、Collection型、プリミティブ型に対応)、[値になるList]  
戻値： List型

* 条件にあう要素のみを残し、新たにListを作成する
* C#のWhereと同じように動作する。
* VBAではアロー演算子が使えないため以下の列挙型で表現する
* 3つ目の引数は戻される値のListを指定できる。Listを2つ用意すればdictionaryのように使える。省略した場合は自身のListが返る

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
    Dim ls As New List
    ls.AddValue Array("foo", "bar", "baz")
    Dim newLs As New List
    Set newLs = ls.Where(lsLike, "ba*")
    Dim i
    For Each i In newLs
        Debug.Print i
    Next
End Sub
'Print
bar
baz
```