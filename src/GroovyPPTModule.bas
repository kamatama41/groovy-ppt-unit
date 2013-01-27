<<<<<<< HEAD
Attribute VB_Name = "GroovyPPTModule"
''' ƒeƒXƒgƒNƒ‰ƒX‚ð¶¬‚µ‚Ü‚·
Sub GenerateTest()
    Dim testName As String
    Dim testClassPath As String
    Dim preStm As New ADODB.stream
    Dim mainStm As New ADODB.stream
    
    testName = GetTestName(ActivePresentation.Name)
    testClassPath = ActivePresentation.Path & GetPathSeparator & testName & ".groovy"

    ' ƒtƒ@ƒCƒ‹‚Ì‘¶Ýƒ`ƒFƒbƒN
    If Dir(testClassPath) <> "" Then
        MsgBox "Šù‚ÉƒeƒXƒgƒNƒ‰ƒX‚ª‘¶Ý‚µ‚Ä‚¢‚Ü‚·"
        Set myStm = Nothing
        Exit Sub
    End If
    
    ' ƒtƒ@ƒCƒ‹‚Ì‘‚«o‚µ
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
    
    ' BOM–³‚É‚·‚é‚½‚ßAÅ‰‚Ì3ƒoƒCƒg•ª‚ð”ò‚Î‚µ‚Ä“Ç‚Þ
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

''' Groovy‚Ìƒ†ƒjƒbƒgƒeƒXƒg‚ðŽÀs‚µ‚Ü‚·
Sub RunTest()
    Dim testName As String

    ' ƒeƒXƒgƒNƒ‰ƒX‚Æƒpƒ‰ƒƒ^—p‚Ìjson‚ÌƒpƒX‚ð¶¬
    testName = GetTestName(ActivePresentation.Name)

    ' jsonƒtƒ@ƒCƒ‹‚ð¶¬
    WtiteJson testName

    ' ƒeƒXƒgŽÀs
    ExecuteTest testName
End Sub

''' ƒtƒ@ƒCƒ‹–¼‚©‚çŠg’£Žq‚ðŽæ‚èœ‚«A––”ö‚É"Test"‚ð•t‚¯‚Ä•Ô‚µ‚Ü‚·
=======
Sub Execute()
    Dim testName As String

    ' ãƒ†ã‚¹ãƒˆã‚¯ãƒ©ã‚¹ã¨ãƒ‘ãƒ©ãƒ¡ã‚¿ç”¨ã®jsonã®ãƒ‘ã‚¹ã‚’ç”Ÿæˆ
    testName = GetTestName(ActivePresentation.Name)

    ' jsonãƒ•ã‚¡ã‚¤ãƒ«ã‚’ç”Ÿæˆ
    WtiteJson testName

    ' ãƒ†ã‚¹ãƒˆå®Ÿè¡Œ
    ExecuteTest testName
End Sub

''' ãƒ•ã‚¡ã‚¤ãƒ«åã‹ã‚‰æ‹¡å¼µå­ã‚’å–ã‚Šé™¤ãã€æœ«å°¾ã«"Test"ã‚’ä»˜ã‘ã¦è¿”ã—ã¾ã™
>>>>>>> origin/master
Private Function GetTestName(fileName As String) As String
    Dim tmp As Variant
    tmp = Split(fileName, ".")
    
    GetTestName = tmp(0) & "Test"
    
End Function


<<<<<<< HEAD
''' ƒXƒ‰ƒCƒh‚Ì“à—e‚ðjsonŒ`Ž®‚Åƒtƒ@ƒCƒ‹‚É‘‚«ž‚Ý‚Ü‚·
Private Function WtiteJson(testName As String)
=======
''' ã‚¹ãƒ©ã‚¤ãƒ‰ã®å†…å®¹ã‚’jsonå½¢å¼ã§ãƒ•ã‚¡ã‚¤ãƒ«ã«æ›¸ãè¾¼ã¿ã¾ã™
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
        ' ˆê‚Â‚Ìslide‚Ìjson‚ðo—Í
=======
        ' ä¸€ã¤ã®slideã®jsonã‚’å‡ºåŠ›
>>>>>>> origin/master
        myStm.WriteText "{""title"":""" & slideTitle & """, ""text"":""" & slideText & """},"
    Next Slide
    myStm.WriteText "]"
    
    
    myStm.SaveToFile fileName:=jsonPath, Options:=adSaveCreateOverWrite
    
    myStm.Close
    Set myStm = Nothing
End Function


<<<<<<< HEAD
''' json‚Ìƒtƒ@ƒCƒ‹ƒpƒX‚ðŽæ“¾‚µ‚Ü‚·
Private Function GetJsonFilePath(testName As String) As String
    GetJsonFilePath = ActivePresentation.Path & GetPathSeparator & testName & ".json"
End Function

''' Mac‚Å‚à“®‚­‚æ‚¤‚É‚·‚é‚½‚ß‚É‘‚¢‚½‚ªMac‚Å‚Í“®ì–¢Šm”F
Private Function GetPathSeparator() As String
    Dim pSeparator As String
    Dim os As String
=======
''' jsonã®ãƒ•ã‚¡ã‚¤ãƒ«ãƒ‘ã‚¹ã‚’å–å¾—ã—ã¾ã™
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

''' Groovy‚ÌƒeƒXƒg‚ðŽÀs‚µ‚Ü‚·
=======
    GetJsonFilePath = ActivePresentation.Path & pSeparator & testName & ".json"
End Function

''' Groovyã®ãƒ†ã‚¹ãƒˆã‚’å®Ÿè¡Œã—ã¾ã™
>>>>>>> origin/master
Private Function ExecuteTest(testName As String)
    Dim command As String
    Dim shell As Object
    Set shell = CreateObject("WScript.Shell")
    Dim rc As Integer
    
<<<<<<< HEAD
    ' Œ»Ý‚Ìƒtƒ@ƒCƒ‹ƒpƒX‚Ü‚ÅˆÚ“®‚µ‚ÄAƒeƒXƒgƒXƒNƒŠƒvƒgŽÀs
=======
    ' ç¾åœ¨ã®ãƒ•ã‚¡ã‚¤ãƒ«ãƒ‘ã‚¹ã¾ã§ç§»å‹•ã—ã¦ã€ãƒ†ã‚¹ãƒˆã‚¹ã‚¯ãƒªãƒ—ãƒˆå®Ÿè¡Œ
>>>>>>> origin/master
    command = "%ComSpec% /c cd " & ActivePresentation.Path & _
                 " & groovy -c UTF-8 " & testName
                 
    rc = shell.Run(command & " & pause")
End Function
