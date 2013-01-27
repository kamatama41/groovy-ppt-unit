Attribute VB_Name = "GroovyPPTModule"
''' �e�X�g�N���X�𐶐����܂�
Sub GenerateTest()
    Dim testName As String
    Dim testClassPath As String
    Dim preStm As New ADODB.stream
    Dim mainStm As New ADODB.stream
    
    testName = GetTestName(ActivePresentation.Name)
    testClassPath = ActivePresentation.Path & GetPathSeparator & testName & ".groovy"

    ' �t�@�C���̑��݃`�F�b�N
    If Dir(testClassPath) <> "" Then
        MsgBox "���Ƀe�X�g�N���X�����݂��Ă��܂�"
        Set myStm = Nothing
        Exit Sub
    End If
    
    ' �t�@�C���̏����o��
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
    
    ' BOM���ɂ��邽�߁A�ŏ���3�o�C�g�����΂��ēǂ�
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

''' Groovy�̃��j�b�g�e�X�g�����s���܂�
Sub RunTest()
    Dim testName As String

    ' �e�X�g�N���X�ƃp�����^�p��json�̃p�X�𐶐�
    testName = GetTestName(ActivePresentation.Name)

    ' json�t�@�C���𐶐�
    WtiteJson testName

    ' �e�X�g���s
    ExecuteTest testName
End Sub

''' �t�@�C��������g���q����菜���A������"Test"��t���ĕԂ��܂�
Private Function GetTestName(fileName As String) As String
    Dim tmp As Variant
    tmp = Split(fileName, ".")
    
    GetTestName = tmp(0) & "Test"
    
End Function


''' �X���C�h�̓��e��json�`���Ńt�@�C���ɏ������݂܂�
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
        ' ���slide��json���o��
        myStm.WriteText "{""title"":""" & slideTitle & """, ""text"":""" & slideText & """},"
    Next Slide
    myStm.WriteText "]"
    
    
    myStm.SaveToFile fileName:=jsonPath, Options:=adSaveCreateOverWrite
    
    myStm.Close
    Set myStm = Nothing
End Function


''' json�̃t�@�C���p�X���擾���܂�
Private Function GetJsonFilePath(testName As String) As String
    GetJsonFilePath = ActivePresentation.Path & GetPathSeparator & testName & ".json"
End Function

''' Mac�ł������悤�ɂ��邽�߂ɏ�������Mac�ł͓��얢�m�F
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

''' Groovy�̃e�X�g�����s���܂�
Private Function ExecuteTest(testName As String)
    Dim command As String
    Dim shell As Object
    Set shell = CreateObject("WScript.Shell")
    Dim rc As Integer
    
    ' ���݂̃t�@�C���p�X�܂ňړ����āA�e�X�g�X�N���v�g���s
    command = "%ComSpec% /c cd " & ActivePresentation.Path & _
                 " & groovy -c UTF-8 " & testName
                 
    rc = shell.Run(command & " & pause")
End Function
