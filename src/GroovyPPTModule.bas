Attribute VB_Name = "GroovyPPTModule"
Sub Execute()
    Dim testName As String

    ' テストクラスとパラメタ用のjsonのパスを生成
    testName = GetTestName(ActivePresentation.Name)

    ' jsonファイルを生成
    WtiteJson testName

    ' テスト実行
    ExecuteTest testName
End Sub

''' ファイル名から拡張子を取り除き、末尾に"Test"を付けて返します
Private Function GetTestName(fileName As String) As String
    Dim tmp As Variant
    tmp = Split(fileName, ".")
    
    GetTestName = tmp(0) & "Test"
    
End Function


''' スライドの内容をjson形式でファイルに書き込みます
Private Function WtiteJson(testName As String)
    Dim os As String
    Dim pSeparator As String
    Dim jsonPath As String
    Dim slideTitle As String
    Dim slideText As String
    Dim myStm As New ADODB.stream
    
    jsonPath = GetJsonFilePath(testName)

    myStm.Open
    myStm.Type = adTypeText
    myStm.Charset = "UTF-8"
    myStm.WriteText "["
    For Each Slide In ActivePresentation.Slides
        slideTitle = Slide.Shapes.Placeholders(1).TextFrame.TextRange.text
        slideTitle = Replace(slideTitle, vbCr, "")
        slideText = Slide.Shapes.Placeholders(2).TextFrame.TextRange.text
        slideText = Replace(slideText, vbCr, "")
        ' 一つのslideのjsonを出力
        myStm.WriteText "{""title"":""" & slideTitle & """, ""text"":""" & slideText & """},"
    Next Slide
    myStm.WriteText "]"
    
    
    myStm.SaveToFile fileName:=jsonPath, Options:=adSaveCreateOverWrite
    
    myStm.Close
    Set myStm = Nothing
End Function


''' jsonのファイルパスを取得します
Private Function GetJsonFilePath(testName As String) As String
    os = Application.OperatingSystem
    If InStr(os, "Windows") Then
        pSeparator = "\\"
    Else
        pSeparator = ":"
    End If
    GetJsonFilePath = ActivePresentation.Path & pSeparator & testName & ".json"
End Function

''' Groovyのテストを実行します
Private Function ExecuteTest(testName As String)
    Dim command As String
    Dim shell As Object
    Set shell = CreateObject("WScript.Shell")
    Dim rc As Integer
    
    ' 現在のファイルパスまで移動して、テストスクリプト実行
    command = "%ComSpec% /c cd " & ActivePresentation.Path & _
                 " & groovy -c UTF-8 " & testName
                 
    rc = shell.Run(command & " & pause")
End Function
