

目次
<!-- @import "[TOC]" {cmd="toc" depthFrom=1 depthTo=6 orderedList=false} -->
<!-- code_chunk_output -->

- [ObjectStringクラス](#objectstringクラス)
  - [メソッド](#メソッド)
    - [Concat](#concat)
    - [Contains](#contains)
    - [EndWith](#endwith)
    - [IndexOf](#indexof)
    - [Insert](#insert)
    - [Lstrip](#lstrip)
    - [Replace](#replace)
    - [Rstrip](#rstrip)
    - [SetString](#setstring)
    - [Split](#split)
    - [StartWith](#startwith)
    - [Strip](#strip)
    - [Substring](#substring)
    - [ToNumber](#tonumber)
    - [ToString](#tostring)
  - [プロパティ](#プロパティ)
    - [Item](#item)
    - [Length](#length)

<!-- /code_chunk_output -->

# ObjectStringクラス

オブジェクト型のString  
モダンな言語の文字型の実態はObject型から継承されたものであり、文字操作はメソッドから行う  
VBAでもそのクラスを作成し、モダンなコードを実現する。  
Itemプロパティを既定メンバにすることで通常のString型と同じように使用できます。  
<使用例>
```VB
Sub test()
    Dim str As New ObjectString
    str = "foo"
    Debug.Print str
    str = str.SetString("bar_baz").Rstrip("_")
    Debug.Print str
End Sub
'Print
foo
bar
```


## メソッド

### Concat
引数： 区切り文字、結合する文字(複数可)  
戻値： ObjectString型  

* 区切り文字をつけて文字同士を結合します。
* JSのconcatのように動作する。
* 区切り文字は""(空文字)も可能
* 処理速度的には連結演算子(&)より遅いです。コードの見やすさ重視で使用します。
```VB
Sub test()
    Dim str As New ObjectString
    str = str.Concat("_", "foo", "bar", "baz")
    Debug.Print str
End Sub
'Print
foo_bar_baz
```

### Contains
引数： 調べたい文字  
戻値： 真偽値

* 引数の文字が存在しているかを真偽値で返す
* C#のContainsのように動作する

```VB
Sub test()
    Dim str As New ObjectString
    Dim bln As Boolean
    str = "foo"
    bln = str.Contains("fo")
    Debug.Print bln
    bln = str.Contains("bar")
    Debug.Print bln
End Sub
'Print
True
False
```

### EndWith
引数： 調べたい文字  
戻値： 真偽値

* 引数の文字が末尾にあるかどうかを真偽値で返す
* C#やJSのEndWithのように動作する

```VB
Sub test()
    Dim str As New ObjectString
    Dim bln As Boolean
    str = "foobar"
    bln = str.EndWith("bar")
    Debug.Print bln
    bln = str.EndWith("foo")
    Debug.Print bln
End Sub
'Print
True
False
```

### IndexOf
引数： 調べたい文字  
戻値： インデックス番号

* 引数の文字が最初に現れたインデックス番号を返します
* 見つからない場合は-1を返す
* C#やJSのIndexOfのように動作する

```VB
Sub test()
    Dim str As New ObjectString
    Dim num As Long
    str = "foobar"
    num = str.IndexOf("bar")
    Debug.Print num
    num = str.IndexOf("baz")
    Debug.Print num
End Sub
'Print
 4 
-1
```

### Insert
引数： 開始位置、挿入する文字  
戻値： ObjectString型

* インスタンスされた文字の間に文字を挿入する
* C#のInsertのように動作する

```VB
Sub tete()
    Dim Str As New ObjectString
    Str = "foobar"
    Str = Str.Insert(4, "baz")
    Debug.Print Str
End Sub
'Print
foobazbar
```

### Lstrip
引数： 基準となる文字列  
戻値： ObjectString型  

* 先頭から引数の文字までを除去する
* 引数の文字がない場合は何もしない
* 引数の文字が複数あったとしても1番先頭しか処理しない
* Pythonのlstripと同じように動作する

```VB
Sub test()
    Dim str As New ObjectString
    str = "foo_bar_baz"
    str = str.Lstrip("_")
    Debug.Print str
End Sub
'Print
bar_baz
```

### Replace
引数： 置換対象、置換後の文字  
戻値： ObjectString型  

* 第1引数を第2引数に置換する
* VBAのReplace関数と処理は同じ
* 置換対象が複数ある場合、すべて置換を行います

```VB
Sub test()
    Dim str As New ObjectString
    str = "foo_bar_baz"
    str = str.Replace("_", ",")
    Debug.Print str
End Sub
'Print
foo,bar,baz
```

### Rstrip
引数： 基準となる文字列  
戻値： ObjectString型  

* 後尾から引数の文字までを除去する
* 引数の文字がない場合は何もしない
* 引数の文字が複数あったとしても1番後尾しか処理しない
* Pythonのrstripと同じように動作する

```VB
Sub test()
    Dim str As New ObjectString
    str = "foo_bar_baz"
    str = str.Rstrip("_")
    Debug.Print str
End Sub
'Print
foo_bar
```

### SetString
引数： セットする要素   
戻値： ObjectString型  

* 引数を強制的に文字列に変換してObjectString型にする
* このメソッドはObjectString型を返すので、通常Stringをセット=>文字操作メソッドを続けて書くためにも使用できます。

```VB
Sub test()
    Dim str As New ObjectString
    Dim pi As Double
    pi = 4 * Atn(1)
    str = str.SetString(pi)
    Debug.Print str
End Sub
'Print
3.14159265358979
```

### Split
引数： 区切り文字、[分割数]  
戻値： 1次元配列

* 引数の文字で区切り1次元配列を作成する
* 1次元配列は0から始まる
* 第2引数を省略すると全て分割する
* VBAのSplit関数と処理は同じ(ただし、戻り値はLists型)

```VB
Sub test()
    Dim str As New ObjectString
    Dim arr As Variant
    str = "foo_bar_baz"
    arr = str.Split("_")
    Dim element As Variant
    For Each element In arr
        Debug.Print element
    Next element
End Sub
'Print
foo
bar
baz
```

### StartWith
引数： 調べたい文字  
戻値： 真偽値

* 引数の文字が先頭にあるかどうかを真偽値で返す
* C#やJSのStartWithのように動作する

```VB
Sub test()
    Dim str As New ObjectString
    Dim bln As Boolean
    str = "foobar"
    bln = str.StartsWith("bar")
    Debug.Print bln
    bln = str.StartsWith("foo")
    Debug.Print bln
End Sub
'Print
False
True
```

### Strip
引数： 先頭の基準となる文字列、後尾の基準となる文字列  
戻値： ObjectString型  

* 先頭と後尾両方の引数の文字までを除去する
* 引数の文字がない場合は何もしない
* 引数の文字が複数あったとしても1番先頭、後尾しか処理しない
* Pythonのstripと同じように動作する

```VB
Sub test()
    Dim str As New ObjectString
    str = "foo_bar_baz"
    str = str.Strip("_", "ba")
    Debug.Print str
End Sub
'Print
bar_
```

### Substring
引数： 開始位置、[文字の長さ]  
戻値： ObjectString型  

* 文字を指定した位置から切り抜く
* 第2引数を省略すると最後も文字まで
* VBAのLeft関数とMid関数と同じ処理ができる
* C#のSubstringと同じように動作する(JSとは異なる)

```VB
Sub test()
    Dim Str As New ObjectString
    Str = "foobar"
    Str = Str.Substring(4)
    Debug.Print Str
End Sub
'Print
bar
```

### ToNumber
引数： なし  
戻値： 数値型(多くはDouble型)  

* インスタンスしている文字を通常の数値型にして返す
* 処理的にはVBAのVal関数と同じ

```VB
Sub test()
    Dim Str As New ObjectString
    Dim num As Long
    Str = "150"
    num = Str.ToNumber() + 300
    Debug.Print num
End Sub
'Print
450
```

### ToString
引数： なし  
戻値： 通常のString型  

* インスタンスしている文字を通常のString型にして返す
* 既定メンバもString型を返すため通常はわざわざ使う必要ない。C#と同じようにポリモーフィズムのために実装しているメソッド

```VB
Sub test()
    Dim str As New ObjectString
    Dim primitiveStr As String
    str = "foo"
    primitiveStr = str.ToString()
    Debug.Print primitiveStr
    'The above and below work in the same way
    str = "bar"
    primitiveStr = str
    Debug.Print primitiveStr
End Sub
'Print
foo
bar
```

## プロパティ

### Item
**Get**  
引数： なし  
戻値： 文字(String型)  
**Set**  
引数： 文字(String型)  
戻値： なし  

* 文字をインスタンスします。
* 既定メンバなので実際はItemプロパティを指定して使うことはない

```VB
Sub test()
    Dim str As New ObjectString
    str.Item = "foo"
    Debug.Print str
    'The above and below work in the same way
    str = "bar"
    Debug.Print str
End Sub
'Print
foo
bar
```

### Length
**Get**  
引数： なし  
戻値： 文字の長さ  
**Set**  
不可

* 文字の長さを返します
* C#やJSのLengthのように動作する

```VB
Sub test()
    Dim str As New ObjectString
    Dim num As Long
    str = "foobarbaz"
    num = str.Length
    Debug.Print num
End Sub
'Print
9
```