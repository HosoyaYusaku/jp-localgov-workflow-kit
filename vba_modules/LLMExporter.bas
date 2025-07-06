Attribute VB_Name = "LLMExporter"
Option Explicit

' --- �O�����C�u�������g�킸�Ƀe�L�X�g�t�@�C�����o�͂��邽�߂̐ݒ� ---
Private Const adSaveCreateOverWrite = 2
Private Const adTypeText = 2

' --- �ݒ荀�� ---
' �����W���[��1�Œ�`�ς݂̒萔���Q�Ƃ��܂��B
Private Const FILE_PREFIX As String = "�Z���[��t�Ɩ��t���["


'===================================================================================================
' ������ ���C�������iLLM�A�g�t�@�C���o�́j ������
' LLM�A�g�p�̃}�[�N�_�E���t�@�C����2��ޏo�͂��܂��B
'===================================================================================================
Sub ExportMarkdownForLLM()
    Dim wsInput As Worksheet
    Dim lastRow As Long
    Dim dataArr As Variant
    Dim manualMarkdown As String
    Dim analysisMarkdown As String
    Dim saveFolder As String
    Dim manualFilePath As String
    Dim analysisFilePath As String

    On Error GoTo ErrorHandler
    Application.ScreenUpdating = False

    ' 1. ���̓V�[�g�̑��݊m�F�ƃf�[�^�ǂݍ���
    Set wsInput = Nothing
    On Error Resume Next
    ' ���W���[��1�̒萔���Q��
    Set wsInput = ThisWorkbook.Worksheets("�Ɩ����X�g")
    On Error GoTo ErrorHandler

    If wsInput Is Nothing Then
        MsgBox "�u�Ɩ����X�g�v�V�[�g��������܂���B", vbCritical
        GoTo Cleanup
    End If

    lastRow = wsInput.Cells(wsInput.Rows.count, "A").End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "�u�Ɩ����X�g�v�V�[�g�Ƀf�[�^������܂���B", vbExclamation
        GoTo Cleanup
    End If

    ' �w�b�_�[���������f�[�^��z��Ɉꊇ�ǂݍ���
    dataArr = wsInput.Range("A2:J" & lastRow).Value

    ' 2. �e�}�[�N�_�E���R���e���c�̐���
    manualMarkdown = GenerateManualMarkdown(dataArr)
    analysisMarkdown = GenerateAnalysisMarkdown(dataArr)

    ' 3. �t�@�C���̕ۑ� (�K���f�X�N�g�b�v�ɕۑ�����悤�ύX)
    saveFolder = CreateObject("WScript.Shell").SpecialFolders("Desktop")

    manualFilePath = saveFolder & "\" & FILE_PREFIX & "(�}�j���A���݌v�p).md"
    analysisFilePath = saveFolder & "\" & FILE_PREFIX & "(�ۑ蕪�͗p).md"

    If SaveTextToFile(manualMarkdown, manualFilePath) And SaveTextToFile(analysisMarkdown, analysisFilePath) Then
        MsgBox "LLM�A�g�p�̃}�[�N�_�E���t�@�C��2�_���f�X�N�g�b�v�ɏo�͂��܂����B" & vbCrLf & vbCrLf & _
               "�E" & manualFilePath & vbCrLf & _
               "�E" & analysisFilePath, vbInformation, "�o�͊���"
    Else
        MsgBox "�t�@�C���̏o�͒��ɃG���[���������܂����B", vbCritical, "�G���["
    End If

Cleanup:
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    MsgBox "�G���[���������܂����B" & vbCrLf & "�G���[�ԍ�: " & Err.Number & vbCrLf & "�G���[���e: " & Err.Description, vbCritical, "�G���[����"
    Resume Cleanup
End Sub

'===================================================================================================
' �Ɩ��}�j���A���쐬�p�̃}�[�N�_�E���𐶐�����
'===================================================================================================
Private Function GenerateManualMarkdown(dataArr As Variant) As String
    'StringBuilder���g�킸�A�ʏ�̕����񌋍��ɕύX
    Dim markdownContent As String
    Dim i As Long

    ' --- �v�����v�g���� ---
    markdownContent = markdownContent & "# �w��" & vbCrLf
    markdownContent = markdownContent & "���Ȃ��́A�v���̃e�N�j�J�����C�^�[�ł��B�ȉ��̋Ɩ��v���Z�X���Ɋ�Â��āA�V�l�E���◘�p�҂��ǂ�ł������ł���悤�ȁA���J�ŕ�����₷���Ɩ��}�j���A�����쐬���Ă��������B" & vbCrLf & vbCrLf
    markdownContent = markdownContent & "## �쐬����}�j���A���̗v��" & vbCrLf
    markdownContent = markdownContent & "- �Ɩ��̑S�̑����ŏ��ɂ킩��悤�Ɂu�͂��߂Ɂv�Ɓu�Ɩ��̗���v���L�q���Ă��������B" & vbCrLf
    markdownContent = markdownContent & "- �u�ڍׂȎ菇�v�ł́A�e�菇����̓I�ɐ������Ă��������B" & vbCrLf
    markdownContent = markdownContent & "- �S���҂��u���p�ҁv�Ɓu�E���v�ɕ�����Ă��邱�Ƃ𖾊m�ɂ��A���ꂼ��̎��_�ōs����������悤�ɋL�q���Ă��������B" & vbCrLf
    markdownContent = markdownContent & "- ���p��͔����A���ՂȌ��t�ŉ�����Ă��������B" & vbCrLf & vbCrLf

    ' --- �f�[�^���� ---
    markdownContent = markdownContent & "# �Ɩ��v���Z�X���" & vbCrLf
    markdownContent = markdownContent & "| �菇�ԍ� | �S���� | ��Ƃ┻�f�̓��e | �⑫���� |" & vbCrLf
    markdownContent = markdownContent & "|:---|:---|:---|:---|" & vbCrLf

    For i = 1 To UBound(dataArr, 1)
        markdownContent = markdownContent & "| " & CStr(dataArr(i, 1)) '�菇�ԍ�
        markdownContent = markdownContent & " | " & CStr(dataArr(i, 2)) '�S����
        markdownContent = markdownContent & " | " & CStr(dataArr(i, 3)) '��Ƃ┻�f�̓��e
        markdownContent = markdownContent & " | " & CStr(dataArr(i, 7)) '�⑫����
        markdownContent = markdownContent & " |" & vbCrLf
    Next i

    GenerateManualMarkdown = markdownContent
End Function

'===================================================================================================
' �ۑ蕪�͗p�̃}�[�N�_�E���𐶐�����
'===================================================================================================
Private Function GenerateAnalysisMarkdown(dataArr As Variant) As String
    'StringBuilder���g�킸�A�ʏ�̕����񌋍��ɕύX
    Dim markdownContent As String
    Dim i As Long

    ' --- �v�����v�g���� ---
    markdownContent = markdownContent & "# �w��" & vbCrLf
    markdownContent = markdownContent & "���Ȃ��́A�o���L�x�ȋƖ����P�R���T���^���g�ł��B�ȉ��̋Ɩ��v���Z�X���ƁA�e�菇�ŕ񍐂���Ă���u���育�ƁE�ۑ�v�𕪐͂��A��̓I�ȉ��P��Ă����Ă��������B" & vbCrLf & vbCrLf
    markdownContent = markdownContent & "## ���͂ƒ�Ă̗v��" & vbCrLf
    markdownContent = markdownContent & "- �܂��A�񍐂���Ă���ۑ��v�񂵁A���{�������ǂ��ɂ��邩�𕪐͂��Ă��������B" & vbCrLf
    markdownContent = markdownContent & "- �u�f�W�^�����v�u�v���Z�X�̊ȗ����v�u�E���̕��S�y���v�u���p�҂̗��֐�����v�̊ϓ_����A��̓I�ȉ��P�A�N�V�������Ă��Ă��������B" & vbCrLf
    markdownContent = markdownContent & "- ��ẮA�Z���I�Ɏ����\�Ȃ��̂ƁA�������I�Ɏ��g�ނׂ����̂ɕ����Ē񎦂��Ă��������B" & vbCrLf
    markdownContent = markdownContent & "- ��Ăɂ���Ăǂ̂悤�Ȍ��ʁi���ԒZ�k�A�R�X�g�팸�A�����x����Ȃǁj�����҂ł��邩���L�q���Ă��������B" & vbCrLf & vbCrLf

    ' --- �f�[�^���� ---
    markdownContent = markdownContent & "# �Ɩ��v���Z�X�Ɖۑ�" & vbCrLf
    markdownContent = markdownContent & "| �菇�ԍ� | �S���� | ��Ƃ┻�f�̓��e | ���育�ƁE�ۑ� | ���Ԃ⌏�� |" & vbCrLf
    markdownContent = markdownContent & "|:---|:---|:---|:---|:---|" & vbCrLf

    For i = 1 To UBound(dataArr, 1)
        ' �u���育�ƁE�ۑ�v�����͂���Ă���s�݂̂𒊏o
        If Trim(CStr(dataArr(i, 9))) <> "" Then
            markdownContent = markdownContent & "| " & CStr(dataArr(i, 1))  '�菇�ԍ�
            markdownContent = markdownContent & " | " & CStr(dataArr(i, 2))  '�S����
            markdownContent = markdownContent & " | " & CStr(dataArr(i, 3))  '��Ƃ┻�f�̓��e
            markdownContent = markdownContent & " | " & CStr(dataArr(i, 9))  '���育�ƁE�ۑ�
            markdownContent = markdownContent & " | " & CStr(dataArr(i, 8))  '���Ԃ⌏��
            markdownContent = markdownContent & " |" & vbCrLf
        End If
    Next i

    GenerateAnalysisMarkdown = markdownContent
End Function

'===================================================================================================
' �e�L�X�g�R���e���c��UTF-8�Ńt�@�C���ɕۑ�����
'===================================================================================================
Private Function SaveTextToFile(content As String, filePath As String) As Boolean
    Dim adoStream As Object
    On Error GoTo SaveError

    ' ADODB.Stream�I�u�W�F�N�g���g����UTF-8�ŕۑ�
    Set adoStream = CreateObject("ADODB.Stream")
    With adoStream
        .Type = adTypeText
        .Charset = "UTF-8"
        .Open
        .WriteText content
        .SaveToFile filePath, adSaveCreateOverWrite
        .Close
    End With

    SaveTextToFile = True
    Set adoStream = Nothing
    Exit Function

SaveError:
    SaveTextToFile = False
    If Not adoStream Is Nothing Then Set adoStream = Nothing
End Function

