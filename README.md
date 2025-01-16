# Regular-Expression-Functions
Excelユーザーのための正規表現関数 (Regular Expression Functions for Excel Users)   
初回投稿日：2025年1月16日   
最終更新日：2025年1月16日  

## 1. 概要
Microsoft Excelの正規表現関数（REGEXTEST、REGEXREPLACE、REGEXEXTRACT）は、2024年5月にリリースされました。これらの関数は、Excelの365 Insiderプレビュー版で最初に導入されましたが、現在は一般公開されています。  
これらの正規表現関数は、Microsoft 365のサブスクリプションを持っているユーザーが利用できます。具体的には、Excel for Microsoft 365、Excel for the web、Excel for iOS、Excel for Androidなどのバージョンで利用可能です。これ以外の環境ではこれらの正規表現関数を使うことができません。そこでVBAを使ってユーザー定義関数として４つの関数を紹介したいと思います。 

### ここで紹介する正規表現関数
|ユーザー定義関数|概要|
| :---: | :---           |
|RegexTest2|指定されたテキストが正規表現パターンに一致するかどうかを判定します。|   
|RegexCount2| 指定された正規表現パターンに一致する文字列の個数をカウントします。|    
|RegexReplace2|指定された正規表現パターンに一致する文字列を別の文字列で置換します。|    
|RegexExtract2|指定された正規表現パターンに一致する文字列を抽出します。|    
  
これらのユーザー定義関数はExcelのワークシート上でもVBAの中でも使えます。  

### 制約条件
正規表現関数（REGEXTEST、REGEXREPLACE、REGEXEXTRACT）は正規表現エンジン PCRE2 に準拠した仕様になっています。   
一方、VBAで用いる正規表現ライブラリーは「Microsoft VBScript Regular Expressions 5.5」です。このライブラリーはECMA-262第3版に準拠しているため「後読み」ができません。   
PCRE2とECMA-262の構文は基本的に同等ですが、一部で実装上の違いがあります。例えばUNICODEで文字を指定する構文は次のように異なります。    
|エンジン（フレーバー）|UNICODE文字の指定|注記|
| :---: | :---: | :---: |
|PCRE2|\x{nnnn}|nnnnは4桁の16進数|  
|ECMA-262|\unnnn|nnnnは4桁の16進数|   
    
    
## 2. 解説とソースコード   
以下にそれぞれの関数の解説とソースコードを掲載します。   
ソースコードを一括してダウンロードしたい場合は、このリポジトリーにある [RE_Module_VBA.bas](RE_Module_VBA.bas) をお使いください。  
      
### 2.1 RegexTest2  
テキストの任意の部分が正規表現パターンと一致するかどうか検査します。  
Excelの正規表現関数（REGEXTEST）に準拠しています。  
参考： [REGEXTEST関数](https://support.microsoft.com/ja-jp/office/regextest-%E9%96%A2%E6%95%B0-7d38200b-5e5c-4196-b4e6-9bff73afbd31) 
   
```
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
```

### 2.2 RegexCount2  
正規表現のパターンと一致する文字列の個数をカウントします。  
Excelの正規表現関数にはない関数です。   
REGEXTEST関数は一致するかしないかだけを知るこができますが、一致した個数までは知ることができません。 
この関数は一致した個数を調べたり、その個数によって計算の条件を変える場合に使えます。  
   
```
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
```

### 2.3 RegexReplace2
指定されたテキスト内の文字列を、パターンに一致する文字列を置換に置き換えます。  
Excelの正規表現関数（REGEXREPLACE）に準拠していますが、違いはOccurrenceで指定できるのは 0 と 1 だけになります。この制約はVBAで用いる正規表現ライブラリーの実装上の理由によります。   
参考： [REGEXREPLACE関数](https://support.microsoft.com/ja-jp/office/regexreplace-%E9%96%A2%E6%95%B0-9c030bb2-5e47-4efc-bad5-4582d7100897) 
  
```
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
```

### 2.4 RegexExtract2
指定されたテキスト内で正規表現パターンに一致する文字列を抽出します。  
Excelの正規表現関数（REGEXEXTRACT）に準拠しています。  
REGEXEXTRACT関数のReturn_Modeは0～2を指定できますが、RegexExtract2では1～3を指定できます。  
3は「すべての一致からキャプチャグループを２次元配列として返す」という機能で、これはREGEXEXTRACTにはない仕様です。  
参考： [REGEXEXTRACT関数](https://support.microsoft.com/ja-jp/office/regexextract-%E9%96%A2%E6%95%B0-4b96c140-9205-4b6e-9fbe-6aa9e783ff57) 
   
```
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
```
  
# 3. ライセンス
このコードはMITライセンスに基づき利用できます。  
   
■  
