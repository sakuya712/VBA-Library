VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ObjectString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'''<classSummary> �I�u�W�F�N�g�^��String�A������������\�b�h�ōs�����Ƃ��ł���
'''<updateDate> 2020/10/9
'''<version> Ver.103

Private ObjString As String

'''<summary> �������Z�b�g����
'''<param> �Z�b�g���镶��
'''<returns> �Ȃ�
Public Property Let Item(ByVal Arg As String)
Attribute Item.VB_UserMemId = 0
    On Error GoTo err
    ObjString = VBA.CStr(Arg)
    Exit Property
err:
    ObjString = VBA.TypeName(Arg)
End Property

'''<summary> �������Q�b�g����
'''<param> �Ȃ�
'''<returns> ����
Public Property Get Item() As String
    Item = ObjString
End Property

'''<summary> ���������擾����
'''<param> �Ȃ�
'''<returns> ������
Public Property Get Length() As Long
    Length = Len(ObjString)
End Property

'''<summary> �ʏ��String�ɕϊ�����
'''<param> �Ȃ�
'''<returns> String�^
Public Function ToString() As String
    ToString = ObjString
End Function

'''<summary> ���l�^�ɕϊ�����
'''<param> �Ȃ�
'''<returns> ���l�^
Public Function ToNumber() As Variant
    ToNumber = VBA.Val(ObjString)
End Function

'''<summary> �����I�ɕ�����^�ɂ���ObjectString���쐬����(Null�̓G���[�ɂȂ�܂�)
'''<param> ����
'''<returns> ObjectString�^
Public Function SetString(ByVal Arg As Variant) As ObjectString
    Dim NewStr As New ObjectString
    NewStr = CStr(Arg)
    Set SetString = NewStr
End Function

'''<summary> �w�肳�ꂽ��������O�̕����𔲂����
'''<param> ��ƂȂ镶����
'''<returns> ObjectString�^
Public Function Rstrip(ByVal char As String) As ObjectString

    Dim NewStr As New ObjectString
    Dim pos As Long
    pos = VBA.InstrRev(ObjString, char)
    If pos = 0 Then
        NewStr = ObjString
    Else
        NewStr = VBA.Left(ObjString, pos - 1)
    End If
    Set Rstrip = NewStr
End Function

'''<summary> �w�肳�ꂽ���������̕����𔲂����
'''<param> ��ƂȂ镶����
'''<returns> ObjectString�^
Public Function Lstrip(ByVal char As String) As ObjectString

    Dim NewStr As New ObjectString
    Dim Arr As Variant
    Arr = VBA.Split(ObjString, char, 2)
        Select Case UBound(Arr)
            Case -1, 0
                NewStr = ObjString
            Case Else
                NewStr = Arr(1)
        End Select
    Set Lstrip = NewStr
End Function

'''<summary> �w�肳�ꂽ��������O��̕����𔲂����
'''<param> �O�̊�ƂȂ镶����A���̊�ƂȂ镶����
'''<returns> ObjectString�^
Public Function Strip(ByVal BeginChar As String, ByVal EndChar As String) As ObjectString

    Dim NewStr As New ObjectString
    NewStr = Me.Lstrip(BeginChar).Rstrip(EndChar)
    Set Strip = NewStr
End Function

'''<summary> ����̕�����u������
'''<param> �u���ΏہA�u����̕���
'''<returns> ObjectString�^
Public Function Replace(ByVal findString As String, ByVal replaceString As String) As ObjectString
    Dim NewStr As New ObjectString
    NewStr = VBA.Replace(ObjString, findString, replaceString)
    Set Replace = NewStr
End Function

'''<summary> �����̕��������邩�ǂ�����^�U�l�ŕԂ�
'''<param> �����镶��
'''<returns> �^�U�l
Public Function Contains(ByVal char As String) As Boolean
    If VBA.InStr(ObjString, char) Then
        Contains = True
    Else
        Contains = False
    End If
End Function

'''<summary> �����̕����̕����ʒu��Ԃ�
'''<param> �����镶��
'''<returns> �����ʒu�B�Ȃ��ꍇ��-1��Ԃ�
Public Function IndexOf(ByVal char As String) As Long
    If VBA.InStr(ObjString, char) Then
        IndexOf = VBA.InStr(ObjString, char)
    Else
        IndexOf = -1
    End If
End Function

'''<summary> �����̕�������擪�ɂ��邩��^�U�l�ŕԂ�
'''<param> �����镶��
'''<returns> �^�U�l
Public Function StartsWith(ByVal char As String) As Boolean
    If ObjString Like char & "*" Then
        StartsWith = True
    Else
        StartsWith = False
    End If
End Function

'''<summary> �����̕����������ɂ��邩��^�U�l�ŕԂ�
'''<param> �����镶��
'''<returns> �^�U�l
Public Function EndWith(ByVal char As String) As Boolean
    If ObjString Like "*" & char Then
        EndWith = True
    Else
        EndWith = False
    End If
End Function

'''<summary> �������A��������
'''<param> ��؂蕶���A�A����������������(�����\)
'''<returns> ObjectString�^
'''<remarks> �A�����Z�q(&)��菈���͒x���A�R�[�h�̌��₷���d���Ŏg�p���܂��B
Public Function Concat(ByVal Delimiter As String, ParamArray Strings() As Variant) As ObjectString
    Dim NewStr As New ObjectString
    Dim Str As Variant
    Dim i As Long
    If ObjString = "" Then
        NewStr = CStr(Strings(0))
        For i = 1 To UBound(Strings)
            NewStr = NewStr & Delimiter & CStr(Strings(i))
        Next
    Else
        NewStr = ObjString
        For Each Str In Strings
            NewStr = NewStr & Delimiter & CStr(Str)
        Next
    End If
    Set Concat = NewStr
End Function

'''<summary> ������𕪊����܂�
'''<param> ��؂蕶���A������(�w�肵�Ȃ��ꍇ�͑S��)
'''<returns> 0����n�܂�1�����z��
Public Function Split(ByVal Delimiter As String, Optional ByVal limit As Long = -1) As Variant
    Split = VBA.Split(ObjString, Delimiter, limit)
End Function

'''<summary> �������؂蔲���܂�
'''<param> �J�n�ʒu�A�����̒���(�w�肵�Ȃ��ꍇ�͑S��)
'''<returns> �؂蔲��������
Public Function Substring(ByVal StartIndex As Long, Optional ByVal Length As Long) As ObjectString
    Dim NewStr As New ObjectString
    If Length < 1 Then
        NewStr = VBA.Mid(ObjString, StartIndex)
    Else
        NewStr = VBA.Mid(ObjString, StartIndex, Length)
    End If
    Set Substring = NewStr
End Function

'''<summary> ������̊Ԃɕ�����}������
'''<param> �J�n�ʒu�A�}�����镶��
'''<returns> �}�����ꂽ����
Public Function Insert(ByVal StartIndex As Long, ByVal AddString As String) As ObjectString
    Dim NewStr As New ObjectString
    NewStr = VBA.Left(ObjString, StartIndex - 1) & AddString & VBA.Mid(ObjString, StartIndex)
    Set Insert = NewStr
End Function