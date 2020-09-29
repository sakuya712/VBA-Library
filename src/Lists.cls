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
'''<classSummary> List�^�̃R���N�V�����N���X�A�L�[�̐ݒ肪�K�{�ł�
'''<updateDate> 2020/9/14
'''<version> Ver.100
'''<remarks> List�^���K�v�ł�

Private ListColl As Collection
Private KeyLst As List

'''<summary> ������Ԃ����߂̂���
'''<param> �Ȃ�
'''<returns> ���ȎQ��
Public Property Get self() As Lists
    Set self = Me
End Property

'''<summary> �I�u�W�F�N�g���Z�b�g
'''<param> �Ȃ�
'''<returns> �Ȃ�
Private Sub Class_Initialize()
    Set ListColl = New Collection
    Set KeyLst = New List
End Sub

'''<summary> List��ǉ�����
'''<param> �ǉ�����List�������͗v�f�A�L�[
'''<returns> �Ȃ�
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
        err.Raise Number:=1051, Description:="�N���X�FLists�A���\�b�h�FAdd�A�L�[���d�����Ă��܂��B"
End Function

'''<summary> �z��̒l�ARange�̒l�A�I�u�W�F�N�g�̊��胁���o�[�̒l���i�[����
'''<param> �ǉ�����v�f(Range��z��)�A�L�[
'''<returns> �Ȃ�
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
    err.Raise Number:=1052, Description:="�N���X�FLists�A���\�b�h�FAddValue�A�����̌^���K�؂ł͂���܂���"
    
End Function

'''<summary> �e�[�u��(Range�^)��ǂݍ���ŗ񂲂Ƃ�List�ɂ���
'''<param> �ǉ�����v�f(Range)�A[�ŏI�s�A�ŏI��]
'''<returns> �Ȃ�
'''<remarks> ��ԍŏ��̍s���L�[�Ƃ���
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

'''<summary> ����Lists�̗v�f����Ԃ�
'''<param> �Ȃ�
'''<returns> �v�f��
Public Function Count() As Long
    Count = ListColl.Count
End Function

'''<summary> �eList�̗v�f����Ԃ�
'''<param> �L�[
'''<returns> �v�f��
Public Function CountList(ByVal Key As String) As Long
    CountList = ListColl(Key).Count
End Function

'''<summary> �v�f�̓��e��Ԃ�
'''<param> �L�[
'''<returns> ���e
'''<remarks> ���胁���o
Public Function Item(ByVal Key As String) As List
Attribute Item.VB_UserMemId = 0
    On Error GoTo err1
    Set Item = ListColl(Key)
    Exit Function
err1:
    err.Raise Number:=1053, Description:="�N���X�FLists�A���\�b�h�FItem�A���݂��Ȃ��L�[���I�΂�܂���"
End Function

'''<summary> �v�f�����폜����
'''<param> �폜����L�[
'''<returns> �Ȃ�
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
    Debug.Print "�N���X�FLists�A���\�b�h�FRemove�A���݂��Ȃ��L�[���I�΂�܂���"
End Function

'''<summary> �L�[�̈ꗗ��Ԃ��܂�
'''<param> �Ȃ�
'''<returns> �L�[�̈ꗗ(List�^)
Public Function KeyList() As List
    Set KeyList = KeyLst
End Function

'''<summary> �L�[�����݂��邩���ׂ�
'''<param> ���ׂ����L�[
'''<returns> �^�U�l
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

'''<summary> List���Z���ɏ�������
'''<param> �������ރZ���͈́A�L�[(������)
'''<returns> �Ȃ�
'''<remarks> �L�[�̏��Ԓʂ�ɗ��ǋL���܂�
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
    err.Raise Number:=1059, Description:="�N���X�FLists�A���\�b�h�FToWriteCells�AList�̗v�f�����傫���͈͂��w�肵�Ă��܂�"
End Function

'''<summary> �����C���f�b�N�X�ԍ����ōő�l��������List�ɂ���
'''<param> ���ׂ����L�[(����)
'''<returns> �ő�l���W�܂���List
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

'''<summary> �����C���f�b�N�X�ԍ����ōŏ��l��������List�ɂ���
'''<param> ���ׂ����L�[(����)
'''<returns> �ő�l���W�܂���List
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

'''<summary> ����L�[�̔�r���ď����ƈ�v�����C���f�b�N�X�ԍ��݂̂��c����Lists�^���쐬����
'''<param>  ��r���Z�q�̗񋓌^�A��r�Ώ�(List�^�ACollection�^�A�v���~�e�B�u�^�ɑΉ��A��r�ΏۃL�[)
'''<returns> Lists�^
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
                    err.Raise Number:=1005, Description:="�N���X�FLists�A���\�b�h�FWhere�A�����̌^���K�؂ł͂���܂���B "
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
                    err.Raise Number:=1005, Description:="�N���X�FLists�A���\�b�h�FWhere�A�����̌^���K�؂ł͂���܂���B "
            End Select
        End Select
End Function

'''<summary> Lists���m���������ĐV����List���쐬����
'''<param> ����������Lists
'''<returns> ��������Lists
Public Function Concat(ParamArray Lists() As Variant) As Lists
    Dim ele As Variant
    For Each ele In Lists
        If ele Is Me Then err.Raise Number:=1035, Description:="�N���X�FLists�A���\�b�h�FConcat�A�������g�������ɂ��邱�Ƃ͂ł��܂���"
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

'''<summary> Index�p��List��Lists�ɒǉ�����
'''<param> Index�̃L�[��
'''<returns> �Ȃ�
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

'''<summary> �C�e���[�^ ���P�Ƃł͎g�p�ł��܂���
'''<param> �Ȃ�
'''<returns> �Ȃ�
Public Function NewEnum() As stdole.IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = ListColl.[_NewEnum]
End Function
