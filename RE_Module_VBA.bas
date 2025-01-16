Attribute VB_Name = "RE_Module_VBA"
Option Explicit

' MIT License
'
' Copyright (c) 2025 Excel-VBA-Diary
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.


'-----------------------------------------------------------------------
'�e�L�X�g�̔C�ӂ̕��������K�\���p�^�[���ƈ�v���邩�ǂ�����������
'������     Text�F�������镶����
'           Pattern�F���K�\���p�^�[��
'           Case_Sensitivity�F�啶���Ə���������ʂ��邩�ǂ����̎w��i�ȗ��j
'                               0: �啶���Ə������̋�ʂ���i�K��l�j
'                               1: �啶���Ə���������ʂ��Ȃ�
'�߂�l     True�F��v����AFalse�F��v���Ȃ�
'-----------------------------------------------------------------------
Public Function RegexTest2(ByVal Text As String, _
                           ByVal Pattern As String, _
                           Optional ByVal Case_Sensitivity As Long = 0) As Boolean
    If Text = "" Or Pattern = "" Then
        RegexTest2 = False
        Exit Function
    End If
    
    With CreateObject("VBScript.RegExp")
        .Global = True
        .IgnoreCase = CBool(Case_Sensitivity)
        .Pattern = Pattern
        RegexTest2 = .Test(Text)
    End With

End Function

'-----------------------------------------------------------------------
'���K�\���̃p�^�[���ƈ�v���镶����̌����J�E���g����
'������     Text�F�������镶����
'           Pattern�F���K�\���p�^�[��
'           Case_Sensitivity�F�啶���Ə���������ʂ��邩�ǂ����̎w��i�ȗ��j
'                               0: �啶���Ə������̋�ʂ���i�K��l�j
'                               1: �啶���Ə���������ʂ��Ȃ�
'�߂�l     ���K�\���p�^�[���Ɉ�v���镶����̌�
'-----------------------------------------------------------------------
Public Function RegexCount2(ByVal Text As String, _
                            ByVal Pattern As String, _
                            Optional ByVal Case_Sensitivity As Long = 0) As Long
    
    If Text = "" Or Pattern = "" Then
        RegexCount2 = 0
        Exit Function
    End If
    
    With CreateObject("VBScript.RegExp")
        .Global = True
        .IgnoreCase = CBool(Case_Sensitivity)
        .Pattern = Pattern
        RegexCount2 = .Execute(Text).Count
    End With

End Function

'-----------------------------------------------------------------------
'�w�肳�ꂽ�e�L�X�g���̕�������A�p�^�[���Ɉ�v���镶�����u���ɒu��������
'������     Text�F�u���O�̕�����
'           Pattern�F���K�\���p�^�[��
'           Replacement�F�u��������
'           Occurrence�F�u��������p�^�[���̃C���X�^���X�i�ȗ��j
'                               0: ���ׂẴC���X�^���X���u��������i�K��l�j
'                               1: ��v�����ŏ��̃C���X�^���X������u��������
'           Case_Sensitivity�F�啶���Ə���������ʂ��邩�ǂ����̎w��i�ȗ��j
'                               0: �啶���Ə������̋�ʂ���i�K��l�j
'                               1: �啶���Ə���������ʂ��Ȃ�
'�߂�l     �u����̕�����
'-----------------------------------------------------------------------
Public Function RegexReplace2(ByVal Text As String, _
                              ByVal Pattern As String, _
                              ByVal Replacement As String, _
                              Optional ByVal Occurrence As Long = 0, _
                              Optional ByVal Case_Sensitivity As Long = 0) As String
    With CreateObject("VBScript.RegExp")
        .Global = Not CBool(Occurrence)
        .IgnoreCase = CBool(Case_Sensitivity)
        .Pattern = Pattern
        RegexReplace2 = .Replace(Text, Replacement)
    End With
End Function

'-----------------------------------------------------------------------
'�w�肳�ꂽ�e�L�X�g���Ő��K�\���p�^�[���Ɉ�v���镶����𒊏o����
'������     Text�F���o�̑Ώۂ̕�����
'           Pattern�F���K�\���p�^�[��
'           Return_Mode�F �Ԃ��l�̎w��i�ȗ��j
'                         0: �p�^�[���Ɉ�v����ŏ��̕������Ԃ��i�K��l�j
'                         1: �p�^�[���Ɉ�v���邷�ׂĂ̕������z��Ƃ��ĕԂ�
'                         2: �ŏ��̈�v����L���v�`���O���[�v��z��Ƃ��ĕԂ�
'                         3: ���ׂĂ̈�v����L���v�`���O���[�v���Q�����z��Ƃ��ĕԂ�
'           Case_Sensitivity�F �啶���Ə���������ʂ��邩�ǂ����̎w��i�ȗ��j
'                              0: �啶���Ə������̋�ʂ���i�K��l�j
'                              1: �啶���Ə���������ʂ��Ȃ�
'�߂�l     ���K�\���p�^�[���Ɉ�v���镶����i������܂��͔z��ŕԂ��j
'-----------------------------------------------------------------------
Public Function RegexExtract2(ByVal Text As String, _
                              ByVal Pattern As String, _
                              Optional Return_Mode As Long = 0, _
                              Optional ByVal Case_Sensitivity As Long = 0) As Variant
   
    If Text = "" Or Pattern = "" Then
        RegexExtract2 = ""
        Exit Function
    End If
    
    Dim matches As Object
    With CreateObject("VBScript.RegExp")
        .Global = True
        .IgnoreCase = CBool(Case_Sensitivity)
        .Pattern = Pattern
        Set matches = .Execute(Text)
    End With
    
    If matches.Count = 0 Then
        RegexExtract2 = ""
        Exit Function
    End If

    Dim tempArray As Variant, i As Long, j As Long
    
    Select Case Return_Mode
        Case 0
            RegexExtract2 = matches(0).Value
        Case 1
            ReDim tempArray(matches.Count - 1)
            For i = 0 To matches.Count - 1
                tempArray(i) = matches(i).Value
            Next
            RegexExtract2 = tempArray
        Case 2
            j = matches(0).submatches.Count
            If j = 0 Then
                RegexExtract2 = ""
                Exit Function
            End If
            ReDim tempArray(j - 1)
            For i = 0 To j - 1
                tempArray(i) = matches(0).submatches(i)
            Next
            RegexExtract2 = tempArray
        Case 3
            i = matches.Count
            j = matches(0).submatches.Count
            If i = 0 Or j = 0 Then
                RegexExtract2 = ""
                Exit Function
            End If
            ReDim tempArray(i - 1, j - 1)
            For i = 0 To i - 1
                For j = 0 To j - 1
                    tempArray(i, j) = matches(i).submatches(j)
                Next
            Next
            RegexExtract2 = tempArray
        Case Else
            RegexExtract2 = ""
    End Select
    
End Function

'-----------------------------------------------------------------------
' End of Source Code
'-----------------------------------------------------------------------
