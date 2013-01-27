Attribute VB_Name = "GroovyPPTModule"
''' テストクラスを生成します
Sub GenerateTest()
    Dim testName As String
    Dim testClassPath As String
    Dim preStm As New ADODB.stream
    Dim mainStm As New ADODB.stream
    
    testName = GetTestName(ActivePresentation.Name)
    testClassPath = ActivePresentation.Path & GetPathSeparator & testName & ".groovy"

    ' ファイルの存在チェック
    If Dir(testClassPath) <> "" Then
        MsgBox "既にテストクラスが存在しています"
        Set myStm = Nothing
        Exit Sub
    End If
    
    ' ファイルの書き出し
    preStm.Open
    preStm.Type = adTypeText
    preStm.Charset = "UTF-8"
    
    preStm.WriteText "import org.junit.runner.RunWith", adWriteLine
    preStm.WriteText "import org.junit.Test", adWriteLine
    preStm.WriteText "", adWriteLine
    preStm.WriteText "@RunWith(GroovyPPTTestRunner)", adWriteLine
    preStm.WriteText "class " & testName & " {", adWriteLine
    preStm.WriteText "    " & "PPTPresentation presentation", adWriteLine
    preStm.WriteText "", adWriteLine
    preStm.WriteText "    " & "@Test", adWriteLine
    preStm.WriteText "    " & "void testName() {", adWriteLine
    preStm.WriteText "        " & "assert !'Not yet implemented'", adWriteLine
    preStm.WriteText "    " & "}", adWriteLine
    preStm.WriteText "}"
    
    ' BOM無にするため、最初の3バイト分を飛ばして読む
    preStm.Position = 0
    preStm.Type = adTypeBinary
    preStm.Position = 3
    Dim bin: bin = preStm.Read
    preStm.Close
    
    mainStm.Type = adTypeBinary
    mainStm.Open
    mainStm.Write (bin)
    mainStm.SaveToFile fileName:=testClassPath, Options:=adSaveCreateNotExist
    
    mainStm.Close
    Set mainStm = Nothing
End Sub

''' Groovyのユニットテストを実行します
Sub RunTest()
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
    GetJsonFilePath = ActivePresentation.Path & GetPathSeparator & testName & ".json"
End Function

''' Macでも動くようにするために書いたがMacでは動作未確認
Private Function GetPathSeparator() As String
    Dim pSeparator As String
    Dim os As String
    os = Application.OperatingSystem
    If InStr(os, "Windows") Then
        pSeparator = "\\"
    Else
        pSeparator = ":"
    End If
    GetPathSeparator = pSeparator
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
