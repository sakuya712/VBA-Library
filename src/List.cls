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
'''<classSummary> Collection�Ƀ��\�b�h���g������List�^�A�L�[�͎g�p�ł��܂���B
'''<updateDate> 2020/9/2
'''<version> Ver.103

'''<summary> ��r���Z�q�̗񋓌^
Enum ComparisonOperatorsEnum
    lsEqual = 1
    lsNotEqual = 2
    lsGreater = 3
    lsLess = 4
    lsGreaterEqual = 5
    lsLessEqual = 6
    lsObjectEqual = 7
    lsLike = 8
End Enum

'''<summary> �Z�p���Z�q�̗񋓌^
Enum ArithmeticOperatorsEnum
    lsSum = 1
    lsDifference = 2
    lsMultiply = 3
    lsDivide = 4
    lsMod = 5
    lsExponent = 6
    lsConcatenate = 7
End Enum

Private Coll As Collection

'''<summary> ������Ԃ����߂̂���
'''<param> �Ȃ�
'''<returns> ���ȎQ��
Public Property Get self() As List
    Set self = Me
End Property

'''<summary> �I�u�W�F�N�g���Z�b�g
'''<param> �Ȃ�
'''<returns> �Ȃ�
Private Sub Class_Initialize()
    Set Coll = New Collection
End Sub

'''<summary> �v�f��ǉ�����
'''<param> �ǉ�����v�f (�I�u�W�F�N�g��z�񂻂̂��̂�����)
'''<returns> �Ȃ�
Public Function Add(ByVal Item As Variant)
    Coll.Add Item
End Function

'''<summary> �z��̒l�ARange�̒l�A�I�u�W�F�N�g�̊��胁���o�[�̒l���i�[����
'''<param> �ǉ�����v�f(Range)
'''<returns> �Ȃ�
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
    err.Raise Number:=1001, Description:="�N���X�FList�A���\�b�h�FAddValue�A�����̌^���K�؂ł͂���܂���B "
    
End Function

'''<summary> �v�f�̓��e��Ԃ�
'''<param> �C���f�b�N�X�ԍ�
'''<returns> ���e
'''<remarks> ���胁���o
Public Function Item(ByVal index As Long) As Variant
Attribute Item.VB_UserMemId = 0
    If VarType(Coll(index)) = vbObject Then
        Set Item = Coll(index)
    Else
        Item = Coll(index)
    End If
End Function

'''<summary> �v�f����Ԃ�
'''<param> �Ȃ�
'''<returns> �v�f��
Public Function Count() As Long
    Count = Coll.Count
End Function

'''<summary> �v�f�����폜����
'''<param> �폜����C���f�b�N�X�ԍ�
'''<returns> �Ȃ�
Public Function Remove(ByVal index As Long)
    Coll.Remove (index)
End Function

'''<summary> ���W����List���쐬����
'''<param> ��菜�������v�f��List
'''<returns> ��菜���ꂽList
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

'''<summary> �v�f�ɏd�������邩���ׂ�
'''<param> �Ȃ�
'''<returns> �d��������ꍇ��True��Ԃ�
Public Function IsOverlap() As Boolean
    Dim buf As Collection
    Set buf = New Collection
    Dim c
    '�R���N�V�����̃L�[���d������ƃG���[�ɂȂ�Ȃ�̂𗘗p���Č�����
    On Error GoTo err
    For Each c In Coll
        buf.Add c, CStr(c)
    Next
    IsOverlap = False
    Exit Function
err:
    IsOverlap = True
End Function

'''<summary> �d�����Ȃ�List���쐬����
'''<param> �Ȃ�
'''<returns> �d�����폜����List
Public Function RemoveOverlap() As List
    Dim buf As Collection
    Set buf = New Collection
    Dim c As Variant
    '�R���N�V�����̃L�[���d������ƃG���[�ɂȂ�Ȃ�̂𗘗p���Č�����
    On Error Resume Next
    For Each c In Coll
        buf.Add c, CStr(c)
    Next
    Dim lst As List
    Set lst = New List
    lst.AddValue buf
    Set RemoveOverlap = lst
End Function

'''<summary> �d�������l�����X�g������
'''<param> �Ȃ�
'''<returns> �d�������l��List
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
    Set Overlap = Overlap.RemoveOverlap

    Set OverlapList = Overlap
    
End Function

'''<summary> List�̒��Ɉ����ƃ}�b�`����v�f���ŏ��ɏo������C���f�b�N�X��Ԃ��܂��B������Ȃ��ꍇ�� -1 ��Ԃ��܂��B
'''<param> ����������������A���S��v�̏ꍇ��True������v�̏ꍇ��Flase
'''<returns> �C���f�b�N�X�ԍ�
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

'''<summary> List�̒��Ɉ����ƃ}�b�`����v�f�̗L����^�U�l�ŕԂ�
'''<param> ����������������A���S��v�̏ꍇ��True������v�̏ꍇ��Flase
'''<returns> ����ꍇTrue
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

'''<summary> List���m���������ĐV����List���쐬����
'''<param> ����������List
'''<returns> ��������List
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

'''<summary> �����v�f����List���쐬����B
'''<param> List�Ɋi�[����v�f
'''<returns> List�^
Public Function ToList(ByVal Value As Variant) As List

    Dim lst As List
    Set lst = New List
    Dim i As Long
    For i = 1 To Me.Count
        lst.Add Value
    Next
    Set ToList = lst

End Function

'''<summary> �����ɂ����v�f�݂̂��c���A�V����List���쐬����
'''<param> ��r���Z�q�̗񋓌^�A��r�Ώ�(List�^�ACollection�^�A�v���~�e�B�u�^�ɑΉ�)
'''<returns> List�^
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
                    err.Raise Number:=1005, Description:="�N���X�FList�A���\�b�h�FWhere�A�����̌^���K�؂ł͂���܂���B "
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
                    err.Raise Number:=1005, Description:="�N���X�FList�A���\�b�h�FWhere�A�����̌^���K�؂ł͂���܂���B "
            End Select
        End Select

    Set Where = buf

End Function

'''<summary> �v�f�̒��g���v�Z�⌋�����A�V����List���쐬����
'''<param> �Z�p���Z�q�̗񋓌^�A�Z�p�Ώ�(List�^�ACollection�^�A�v���~�e�B�u�^�ɑΉ�)
'''<returns> List�^
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
                    err.Raise Number:=1006, Description:="�N���X�FList�A���\�b�h�FMap�A�����̌^���K�؂ł͂���܂���B "
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
                    err.Raise Number:=1006, Description:="�N���X�FList�A���\�b�h�FMap�A�����̌^���K�؂ł͂���܂���B "
            End Select
        End Select

    Set Map = buf

End Function

'''<summary> List��A������������ŕԂ��܂�
'''<param> �A������ۂɑ}�����镶��
'''<returns> �A������������
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

'''<summary> List��z��ɕϊ����܂�
'''<param> �Ȃ�
'''<returns> �z�� Variant()�^
'''<remarks> �C���f�b�N�X�ԍ���1����n�܂�܂�
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

'''<summary> �C���f�b�N�X�ԍ���StartIndex�Ԗڂ���EndIndex�Ԗڂ̗v�f��Ԃ�
'''<param> �擪�̃C���f�b�N�X�ԍ��A����̃C���f�b�N�X�ԍ�(�ȗ������ꍇ�͍Ō�܂�
'''<returns> List�^
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

'''<summary> �C�e���[�^ ���P�Ƃł͎g�p�ł��܂���
'''<param> �Ȃ�
'''<returns> �Ȃ�
Public Function NewEnum() As stdole.IUnknown
Attribute NewEnum.VB_UserMemId = -4
        Set NewEnum = Coll.[_NewEnum]
End Function