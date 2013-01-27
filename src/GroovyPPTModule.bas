<<<<<<< HEAD
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
=======
Sub Execute()
    Dim testName As String

    ' 繝�繧ｹ繝医け繝ｩ繧ｹ縺ｨ繝代Λ繝｡繧ｿ逕ｨ縺ｮjson縺ｮ繝代せ繧堤函謌�
    testName = GetTestName(ActivePresentation.Name)

    ' json繝輔ぃ繧､繝ｫ繧堤函謌�
    WtiteJson testName

    ' 繝�繧ｹ繝亥ｮ溯｡�
    ExecuteTest testName
End Sub

''' 繝輔ぃ繧､繝ｫ蜷阪°繧画僑蠑ｵ蟄舌ｒ蜿悶ｊ髯､縺阪�∵忰蟆ｾ縺ｫ"Test"繧剃ｻ倥¢縺ｦ霑斐＠縺ｾ縺�
>>>>>>> origin/master
Private Function GetTestName(fileName As String) As String
    Dim tmp As Variant
    tmp = Split(fileName, ".")
    
    GetTestName = tmp(0) & "Test"
    
End Function


<<<<<<< HEAD
''' スライドの内容をjson形式でファイルに書き込みます
Private Function WtiteJson(testName As String)
=======
''' 繧ｹ繝ｩ繧､繝峨�ｮ蜀�螳ｹ繧男son蠖｢蠑上〒繝輔ぃ繧､繝ｫ縺ｫ譖ｸ縺崎ｾｼ縺ｿ縺ｾ縺�
Private Function WtiteJson(testName As String)
    Dim os As String
    Dim pSeparator As String
>>>>>>> origin/master
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
<<<<<<< HEAD
        ' 一つのslideのjsonを出力
=======
        ' 荳�縺､縺ｮslide縺ｮjson繧貞�ｺ蜉�
>>>>>>> origin/master
        myStm.WriteText "{""title"":""" & slideTitle & """, ""text"":""" & slideText & """},"
    Next Slide
    myStm.WriteText "]"
    
    
    myStm.SaveToFile fileName:=jsonPath, Options:=adSaveCreateOverWrite
    
    myStm.Close
    Set myStm = Nothing
End Function


<<<<<<< HEAD
''' jsonのファイルパスを取得します
Private Function GetJsonFilePath(testName As String) As String
    GetJsonFilePath = ActivePresentation.Path & GetPathSeparator & testName & ".json"
End Function

''' Macでも動くようにするために書いたがMacでは動作未確認
Private Function GetPathSeparator() As String
    Dim pSeparator As String
    Dim os As String
=======
''' json縺ｮ繝輔ぃ繧､繝ｫ繝代せ繧貞叙蠕励＠縺ｾ縺�
Private Function GetJsonFilePath(testName As String) As String
>>>>>>> origin/master
    os = Application.OperatingSystem
    If InStr(os, "Windows") Then
        pSeparator = "\\"
    Else
        pSeparator = ":"
    End If
<<<<<<< HEAD
    GetPathSeparator = pSeparator
End Function

''' Groovyのテストを実行します
=======
    GetJsonFilePath = ActivePresentation.Path & pSeparator & testName & ".json"
End Function

''' Groovy縺ｮ繝�繧ｹ繝医ｒ螳溯｡後＠縺ｾ縺�
>>>>>>> origin/master
Private Function ExecuteTest(testName As String)
    Dim command As String
    Dim shell As Object
    Set shell = CreateObject("WScript.Shell")
    Dim rc As Integer
    
<<<<<<< HEAD
    ' 現在のファイルパスまで移動して、テストスクリプト実行
=======
    ' 迴ｾ蝨ｨ縺ｮ繝輔ぃ繧､繝ｫ繝代せ縺ｾ縺ｧ遘ｻ蜍輔＠縺ｦ縲√ユ繧ｹ繝医せ繧ｯ繝ｪ繝励ヨ螳溯｡�
>>>>>>> origin/master
    command = "%ComSpec% /c cd " & ActivePresentation.Path & _
                 " & groovy -c UTF-8 " & testName
                 
    rc = shell.Run(command & " & pause")
End Function
