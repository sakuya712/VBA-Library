

目次
<!-- @import "[TOC]" {cmd="toc" depthFrom=1 depthTo=6 orderedList=false} -->
<!-- code_chunk_output -->

- [Listクラス](#listクラス)
  - [メソッド](#メソッド)
    - [Add](#add)
    - [AddValue](#addvalue)
    - [Concat](#concat)
    - [Const](#const)
    - [Except](#except)
    - [Includes](#includes)
    - [IndexOf](#indexof)
    - [IsOverlap](#isoverlap)
    - [Item](#item)
    - [Join](#join)
    - [Map](#map)
    - [OverlapList](#overlaplist)
    - [Remove](#remove)
    - [RemoveOverlap](#removeoverlap)
    - [Slice](#slice)
    - [ToArray](#toarray)
    - [ToList](#tolist)
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

### AddValue
引数： 追加する要素  
戻値： なし

* 配列の値やRangeの値、オブジェクトの既定メンバーの値を格納します。
* Collection型ではないオブジェクトを渡すとエラーが返ります。

### Concat
引数： 結合したいList(複数可能)   
戻値： 結合したList

* List同士を結合して新たにListを作成する。
* JSのConcatと同じ

### Const
引数： なし  
戻値： 要素数

* 要素数を返します。Collection型と同じ

### Except
引数： 取り除きたい要素のList  
戻値： 取り除かれたList

* 差集合のListを作成する。
* SQLやC#のExceptのように動作する。

### Includes
引数： 文字列、[完全一致かどうか]   
戻値： 真偽値

*  Listの中に引数とマッチする要素の有無を真偽値で返す。
*  JSのIncludesのように動作する。
*  2つ目の引数で完全一致か、部分一致かを設定できます。既定値は完全一致です。

### IndexOf
引数： 文字列、[完全一致かどうか]   
戻値： インデックス番号
  
* Listの中に引数とマッチする要素が最初に出現するインデックスを返します。
*   JSやC#のIndexOfのように動作する。
*  2つ目の引数で完全一致か、部分一致かを設定できます。既定値は完全一致です。
*  見つからない場合は-1を返します。

### IsOverlap
引数： なし  
戻値： 真偽値

* 要素に重複があるか調べる。ある場合はTrueを返す。

### Item
引数： なし  
戻値： 要素

* 要素を返します。Collection型と同じ。
* 既定メンバ

### Join
引数： 連結する際に挿入する文字  
戻値： 連結した文字列

* Listにある要素を1つの連結した文字列で返します。
* VBAの配列で使えるJoinと同じように動作する。
* 文字列に変換できないオブジェクトの場合、エラーになります

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

### OverlapList
引数： なし  
戻値： 重複したList

* 重複した要素を取り出したListを作成する。
* 重複した要素同士は重複しない

### Remove
引数： インデックス番号   
戻値： なし

* 要素を削除する。Collection型と同じ。

### RemoveOverlap
引数： なし  
戻値： 重複しないList

* 重複しないリストを作成する。

### Slice
引数： 先頭のインデックス番号、[後尾のインデックス番号]    
戻値： 要素を切り取ったList型

* インデックス番号のStartIndex番目からEndIndex番目の要素を返す
* JSのSliceのように動作する。
* 2つ目の引数を省略した場合は最後のインデックス番号まで指定される

### ToArray
引数： なし    
戻値： 1次配列

* Listを1次配列へ変換する
* インデックス番号は1から始まる

### ToList
引数： Listに格納する要素    
戻値： List型

* 同じ要素数のListを作成する

### Where
引数： 比較演算子の列挙型、比較対象(List型、Collection型、プリミティブ型に対応)  
戻値： List型

* 条件にあう要素のみを残し、新たにListを作成する
* C#のWhereと同じように動作する。
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