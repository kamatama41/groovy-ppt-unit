Attribute VB_Name = "GroovyPPTModule"
''' Generate Test Class
Sub GenerateTest()
    Dim testName As String
    Dim testClassPath As String
    Dim preStm As New ADODB.stream
    Dim mainStm As New ADODB.stream
    
    testName = GetTestName(ActivePresentation.Name)
    testClassPath = ActivePresentation.Path & GetPathSeparator & testName & ".groovy"

    If Dir(testClassPath) <> "" Then
        MsgBox "Test class already Exists."
        Set myStm = Nothing
        Exit Sub
    End If
    
    ' Wtire file as text
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
    
    ' non-BOM
    ' @see http://d.hatena.ne.jp/replication/20091117/1258418243
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

''' Run groovy test
Sub RunTest()
    Dim testName As String

    testName = GetTestName(ActivePresentation.Name)

    ' Gererate json file
    WtiteJson testName

    ' Execute test
    ' ExecuteTest testName
End Sub

''' GetTestName
Private Function GetTestName(fileName As String) As String
    Dim tmp As Variant
    tmp = Split(fileName, ".")
    
    GetTestName = tmp(0) & "Test"
    
End Function


''' Write contents as JSON
Private Function WtiteJson(testName As String)
    Dim jsonPath As String
    Dim slideText As String
    Dim myStm As New ADODB.stream
    
    jsonPath = GetJsonFilePath(testName)

    myStm.Open
    myStm.Type = adTypeText
    myStm.Charset = "UTF-8"
    myStm.WriteText "{""slides"":["
    For Each Slide In ActivePresentation.Slides
        myStm.WriteText "{""shapes"":["
        For Each Shape In Slide.Shapes
            slideText = Shape.TextFrame.TextRange.text
            slideText = Replace(slideText, vbCr, "")
            ' Write slide contents as json
            myStm.WriteText "{""text"":""" & slideText & """},"
        Next Shape
        myStm.WriteText "]},"
    Next Slide
    myStm.WriteText "]}"
    
    
    myStm.SaveToFile fileName:=jsonPath, Options:=adSaveCreateOverWrite
    
    myStm.Close
    Set myStm = Nothing
End Function


''' Get json file path
Private Function GetJsonFilePath(testName As String) As String
    GetJsonFilePath = ActivePresentation.Path & GetPathSeparator & testName & ".json"
End Function

''' Processing branches by Mac and windows.
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

''' Execute groovy
Private Function ExecuteTest(testName As String)
    Dim command As String
    Dim shell As Object
    Set shell = CreateObject("WScript.Shell")
    Dim rc As Integer
    
    ' Change directory and execute
    command = "%ComSpec% /c cd " & ActivePresentation.Path & _
                 " & groovy -c UTF-8 " & testName
                 
    rc = shell.Run(command & " & pause")
End Function
