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
'テキストの任意の部分が正規表現パターンと一致するかどうか検査する
'引き数     Text：検査する文字列
'           Pattern：正規表現パターン
'           Case_Sensitivity：大文字と小文字を区別するかどうかの指定（省略可）
'                               0: 大文字と小文字の区別する（規定値）
'                               1: 大文字と小文字を区別しない
'戻り値     True：一致する、False：一致しない
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
'正規表現のパターンと一致する文字列の個数をカウントする
'引き数     Text：検査する文字列
'           Pattern：正規表現パターン
'           Case_Sensitivity：大文字と小文字を区別するかどうかの指定（省略可）
'                               0: 大文字と小文字の区別する（規定値）
'                               1: 大文字と小文字を区別しない
'戻り値     正規表現パターンに一致する文字列の個数
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
'指定されたテキスト内の文字列を、パターンに一致する文字列を置換に置き換える
'引き数     Text：置換前の文字列
'           Pattern：正規表現パターン
'           Replacement：置換文字列
'           Occurrence：置き換えるパターンのインスタンス（省略可）
'                               0: すべてのインスタンスが置き換える（規定値）
'                               1: 一致した最初のインスタンスだけを置き換える
'           Case_Sensitivity：大文字と小文字を区別するかどうかの指定（省略可）
'                               0: 大文字と小文字の区別する（規定値）
'                               1: 大文字と小文字を区別しない
'戻り値     置換後の文字列
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
'指定されたテキスト内で正規表現パターンに一致する文字列を抽出する
'引き数     Text：抽出の対象の文字列
'           Pattern：正規表現パターン
'           Return_Mode： 返す値の指定（省略可）
'                         0: パターンに一致する最初の文字列を返す（規定値）
'                         1: パターンに一致するすべての文字列を配列として返す
'                         2: 最初の一致からキャプチャグループを配列として返す
'                         3: すべての一致からキャプチャグループを２次元配列として返す
'           Case_Sensitivity： 大文字と小文字を区別するかどうかの指定（省略可）
'                              0: 大文字と小文字の区別する（規定値）
'                              1: 大文字と小文字を区別しない
'戻り値     正規表現パターンに一致する文字列（文字列または配列で返す）
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
