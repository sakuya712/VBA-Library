VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RangeLib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'''<classSummary>Rangeを使ってよく行うメソッドをまとめた 静的クラス
'''<updateDate>2020/9/1
'''<version> Ver.100
'''<remarks> このクラスはインスタンスをセットしなくても使用できる。エクセルのみで使用可能

'''<summary> 最終行を取得する
'''<param> 基点にするCell
'''<returns> 最終行の番号
'''<remarks>フィルタをしていないこと
Public Function RowBottom(ByVal Cell As Range) As Long
    RowBottom = Cell.End(xlDown).Row
End Function

'''<summary> 最終列を取得する
'''<param> 基点にするCell
'''<returns> 最終列の番号
'''<remarks>フィルタをしていないこと
Public Function RightCol(ByVal Cell As Range) As Long
    RightCol = Cell.End(xlToRight).Column
End Function

'''<summary> RangeのCurrentRegionの改良版、基点にするCellより、左、上には取得しない。列は指定も可能
'''<param> 基点にするCell、[列の指定]
'''<returns> 空白の行と空白の列の任意の組み合わせで囲まれた範囲
'''<remarks>フィルタをしていないこと。列の指定がない場合は空白の列までの範囲
Public Function CurrentRegion(ByVal Cell As Range, Optional ByVal Col As Long = -1) As Range

    If Col < 1 Then Col = Me.RightCol(Cell)
    Dim cell2 As Range
    Set cell2 = Cell.Parent.Cells(Me.RowBottom(Cell), Col)
    Set CurrentRegion = Range(Cell, cell2)
    
End Function

'''<summary> 基点となる行の下に行を追加して、コピーする
'''<param> コピー元となる行、追加する行数
'''<returns> なし
Public Function InsertRowCopy(ByVal SourceRows As Range, ByVal AddRows As Long)
    Dim insertRows As String
    insertRows = CStr(SourceRows.Row + 1) & ":" & CStr(SourceRows.Row + AddRows)
    Cells.Rows(insertRows).Insert
    Cells.Rows(SourceRows.Row).Copy Cells.Rows(insertRows)
End Function

'''<summary> Rangeの列番号(アルファベット)を数値に変換する
'''<param> 列番号(アルファベット)
'''<returns> 列番号(数値)
Public Function ColToInt(ByVal ColAlphabet As String) As Long
    ColToInt = Range(ColAlphabet & "1").Column
End Function
