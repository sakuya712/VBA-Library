

目次
<!-- @import "[TOC]" {cmd="toc" depthFrom=1 depthTo=6 orderedList=false} -->
<!-- code_chunk_output -->

- [TimerLogクラス](#timerlogクラス)
  - [例](#例)
  - [メソッド](#メソッド)
    - [Constructor](#constructor)
    - [FinishTime](#finishtime)
  - [デストラクタ](#デストラクタ)

<!-- /code_chunk_output -->

# TimerLogクラス

イミディエイトウィンドウに処理時間を測定するためのクラス   

## 例
Constructorメソッドで計測開始し、このオブジェトが破棄されたとき(この関数から抜けるとき)計測終了、結果をイミディエイトウィンドウに表記する。
```VB
Sub test()

    Dim log As New TimerLog: log.Constructor ("使用例")
    
    '処理
    Stop
        
End Sub
```
イミディエイトウィンドウ
```VB
[Begin] 使用例
[Finish] 使用例 , 1187[ms] '<=処理時間が表記される
```

## メソッド

### Constructor
引数： 時間と一緒に表示される文字列(関数名など)  
戻値： なし

* 計測が開始される

### FinishTime
引数： なし  
戻値： 処理時間 [ms]

* 計測時間の結果を返す
* イミディエイトウィンドウには表示されない  
* デストラクタとは別に処理を行っている

## デストラクタ

* 測定結果をイミディエイトウィンドウに表示します。


