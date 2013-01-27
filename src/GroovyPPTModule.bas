Attribute VB_Name = "GroovyPPTModule"
Sub Execute()
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
    os = Application.OperatingSystem
    If InStr(os, "Windows") Then
        pSeparator = "\\"
    Else
        pSeparator = ":"
    End If
    GetJsonFilePath = ActivePresentation.Path & pSeparator & testName & ".json"
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
