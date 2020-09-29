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
'''<classSummary> 2���z����g���₷������N���X (2���z���p�ł��B1���z��̏ꍇ��List���g���₷���ł�)
'''<updateDate> 2020/8/31
'''<version> Ver.100
'''<remarks> ���̃N���X�̓G�N�Z���Ŏg���O��ō쐬���Ă��܂��B
Private Arr() As Variant

'''<summary> ������Ԃ����߂̂���
'''<param> �Ȃ�
'''<returns> ���ȎQ��
Public Property Get self() As ArrayLib
    Set self = Me
End Property

'''<summary> �Z�b�g���ꂽ�z�񂪂��邩�ǂ���
'''<param> �Ȃ�
'''<returns> �^�U�l
Public Property Get IsArray() As Boolean
    On Error GoTo err
    Dim buf
    buf = UBound(Arr)
    IsArray = True
    Exit Function
err:
    IsArray = False
End Property

'''<summary> �z��̒l��Ԃ��A���胁���o
'''<param> 1���z��ڂ̔ԍ��A2���z��ڂ̔ԍ�
'''<returns> �l
Public Property Get Items(ByVal Row As Long, ByVal Col As Long) As Variant
Attribute Items.VB_UserMemId = 0
    Items = Arr(Row, Col)
End Property

'''<summary> 1���z����Z�b�g���� (2���z��ɕϊ�)
'''<param> 1���z��A�s�f�[�^�ɂ��邩��f�[�^�ɂ��邩
'''<returns> �Ȃ�
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

'''<summary> 2���z����Z�b�g����
'''<param> 2���z��
'''<returns> �Ȃ�
Public Function SetArray2Dim(ByVal setArr As Variant)
    Arr() = setArr
End Function

'''<summary> Range�Ŕ͈͂����߂Ēl��z���ݒ肷��
'''<param> �͈͂̍���̃Z���A�͈͂̉E���̃Z��
'''<returns> �Ȃ�
'''<remarks> 2�ڂ̈������ȗ������ꍇCurrentRegion�Ɠ����悤�Ȏ擾���@�ɂȂ邽�����Acell1���A���A��ɂ͂Ȃ�Ȃ�
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

'''<summary>�v�f�����擾����
'''<param> �����̎w�� �ȗ������ꍇ�͂��ׂĂ̎����̗v�f����߂�
'''<returns> �v�f��
Public Function Length(Optional ByVal dimension As Long = -1) As Long

    If dimension = -1 Then
        Length = (UBound(Arr, 1) - LBound(Arr, 1) + 1) * (UBound(Arr, 2) - LBound(Arr, 2) + 1)
    Else
        Length = UBound(Arr, dimension) - LBound(Arr, dimension) + 1
    End If
    
End Function

'''<summary>�����ƈ�v����s�̔z���Ԃ�
'''<param> ��v����l�A��r���Z�q�����
'''<returns> 2���z��
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

'''<summary>�v�f�̍s�����͗���폜����
'''<param> �s�ԍ��A��ԍ� (�ǂ��炩����݂̂�ݒ肷�邱��)
'''<returns> �Ȃ�
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
    err.Raise Number:=1201, Description:="�N���X�FArrayLib�A���\�b�h�FRemove�A�������ݒ肷�邱�Ƃ͂ł��܂���"

End Function

'''<summary>IndexMatch(�s�̂�)���g���Ēl��������
'''<param> �s���̃L�[���[�h�A�L�[�ɂ����A[��̎w��A�G���[�̂Ƃ��l]
'''<returns> ���������l
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

'''<summary>IndexMatch(�s�A��)���g���Ēl��������
'''<param> �s���̃L�[���[�h�A�L�[�ɂ����A�񑤂̃L�[���[�h�A�L�[�ɂ����A[�G���[�̂Ƃ��l]
'''<returns> ���������l
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

'''<summary>�ʏ�̔z��ɕϊ�����
'''<param> �Ȃ�
'''<returns> 2���z��
Public Function ToArray() As Variant()
    ToArray = Arr()
End Function