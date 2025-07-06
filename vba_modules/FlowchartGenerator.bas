Attribute VB_Name = "FlowchartGenerator"
Option Explicit

' --- �ݒ荀�� ---
' ���̃c�[�����Q�Ƃ���V�[�g����A�}�̌����ڂ�ݒ肵�܂��B
Private Const INPUT_SHEET_NAME As String = "�Ɩ����X�g"
Private Const OUTPUT_SHEET_NAME As String = "�Ɩ��t���[�}"
Private Const LEGEND_SHEET_NAME As String = "�}��"
Private Const FONT_NAME As String = "Meiryo UI"

' --- ���C�A�E�g�ݒ� ---
' �}�`�̃T�C�Y��Ԋu�ȂǁA�f�U�C���Ɋւ���ڍאݒ�ł��B
Private Const POOL_TITLE_WIDTH As Long = 120    ' ���[�́u���p�ҁv�Ȃǂ̃^�C�g���̕�
Private Const LANE_HEADER_WIDTH As Long = 100   ' �e�S���ҁi���[���j�̃^�C�g���̕�
Private Const SHAPE_WIDTH As Long = 220         ' �e�菇�i�}�`�j�̕�
Private Const SHAPE_HEIGHT As Long = 110        ' �e�菇�i�}�`�j�̍���
Private Const X_START As Long = POOL_TITLE_WIDTH + LANE_HEADER_WIDTH + 20 ' �}�`��`���n�߂�X���W
Private Const Y_START As Long = 30              ' �}�`��`���n�߂�Y���W
Private Const X_STEP As Long = 260              ' �}�`�Ɛ}�`�̐��������̊Ԋu
Private Const Y_LANE_HEIGHT As Long = 280       ' �e�S���ҁi���[���j�̍���
Private Const Y_SHIFT_SPECIAL As Long = 40      ' ����ȂǓ���Ȑ}�`���㉺�ɂ��炷����
Private Const RIGHT_MARGIN As Long = 50         ' �t���[�}�S�̂̉E���̗]��

' --- �O���[�o���ϐ� ---
' �}�N���S�̂ŋ��ʂ��Ďg�p����ϐ��ł��B
Private wsInput As Worksheet, wsOutput As Worksheet, wsLegend As Worksheet
Private laneLayout As Object, taskData As Object, shapeCollection As Object
Private taskOrder() As String
Private lastUsedLane As String
Private CONNECTOR_TYPE As MsoConnectorType

'===================================================================================================
' ������ ���C������ ������
' ���̃}�N�������s����ƁA�t���[�}�̍쐬���n�܂�܂��B
'===================================================================================================
Sub CreateFlowChart()
    On Error GoTo GenericErrorHandler
    Application.ScreenUpdating = False

    InitializeGlobalVariables
    If Not SetupWorksheets() Then GoTo Cleanup
    If Not LoadTaskData() Then
        MsgBox "�u" & INPUT_SHEET_NAME & "�v�V�[�g�Ƀf�[�^������܂���B", vbExclamation
        GoTo Cleanup
    End If

    CreateLegendSheet
    DrawSwimlanes
    
    ' �V���v���Ŕ��������C�A�E�g����������`��E�ڑ�����
    DrawAndConnectShapes_Absolute
    
    AdjustPoolWidths

    If Not wsOutput Is Nothing Then
        wsOutput.Activate
        ActiveWindow.Zoom = 75
    End If

    MsgBox "�Ɩ��t���[�}�A�y�і}��V�[�g�̍쐬�������������܂����B", vbInformation

Cleanup:
    Application.ScreenUpdating = True
    CleanupGlobalVariables
    Exit Sub
GenericErrorHandler:
    MsgBox "�G���[���������܂����B" & vbCrLf & "�G���[�ԍ�: " & Err.Number & vbCrLf & "�G���[���e: " & Err.Description, vbCritical, "�G���[����"
    Resume Cleanup
End Sub

'===================================================================================================
' �������ƃf�[�^�ǂݍ���
'===================================================================================================
Private Sub InitializeGlobalVariables()
    Set laneLayout = CreateObject("Scripting.Dictionary")
    Set taskData = CreateObject("Scripting.Dictionary")
    Set shapeCollection = CreateObject("Scripting.Dictionary")
    ReDim taskOrder(0)
    lastUsedLane = ""
    ' �ڑ����̎�ނ�ݒ肵�܂��BmsoConnectorCurve(�Ȑ�) �܂��� msoConnectorElbow(�J�N�J�N������)���I�ׂ܂��B
    CONNECTOR_TYPE = msoConnectorCurve
End Sub

Private Sub CleanupGlobalVariables()
    Set wsInput = Nothing: Set wsOutput = Nothing: Set wsLegend = Nothing
    Set laneLayout = Nothing: Set taskData = Nothing: Set shapeCollection = Nothing
    Erase taskOrder
End Sub

' �u�Ɩ����X�g�v�V�[�g����f�[�^��ǂݍ��݂܂��B
Private Function LoadTaskData() As Boolean
    LoadTaskData = False: If wsInput Is Nothing Then Exit Function
    Dim lastRow As Long: lastRow = wsInput.Cells(wsInput.Rows.count, "A").End(xlUp).Row
    If lastRow < 2 Then Exit Function

    Dim dataArr As Variant: dataArr = wsInput.Range("A2:J" & lastRow).Value
    ReDim taskOrder(1 To UBound(dataArr, 1))
    Dim taskCount As Long: taskCount = 0
    
    Dim i As Long
    For i = 1 To UBound(dataArr, 1)
        Dim id As String: id = Trim(CStr(dataArr(i, 1)))
        If id <> "" Then
            taskCount = taskCount + 1
            Dim taskInfo As Object: Set taskInfo = CreateObject("Scripting.Dictionary")
            taskInfo.Add "who", Trim(CStr(dataArr(i, 2)))
            taskInfo.Add "summary", Trim(CStr(dataArr(i, 3)))
            taskInfo.Add "flowElement", Trim(CStr(dataArr(i, 4)))
            taskInfo.Add "taskType", Trim(CStr(dataArr(i, 5)))
            taskInfo.Add "toIds", Replace(Replace(Trim(CStr(dataArr(i, 10))), "�A", ","), " ", "")
            taskData.Add id, taskInfo: taskOrder(taskCount) = id
        End If
    Next i
    
    If taskCount = 0 Then Exit Function
    ReDim Preserve taskOrder(1 To taskCount)
    LoadTaskData = True
End Function

' �t���[�}�Ɩ}���`�悷�邽�߂̐V�����V�[�g���������܂��B
Private Function SetupWorksheets() As Boolean
    On Error Resume Next: Set wsInput = ThisWorkbook.Worksheets(INPUT_SHEET_NAME)
    If wsInput Is Nothing Then MsgBox "�u" & INPUT_SHEET_NAME & "�v�V�[�g��������܂���B", vbCritical: Exit Function
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets(OUTPUT_SHEET_NAME).Delete
    ThisWorkbook.Worksheets(LEGEND_SHEET_NAME).Delete
    Application.DisplayAlerts = True: On Error GoTo 0
    Set wsOutput = ThisWorkbook.Worksheets.Add(After:=wsInput): wsOutput.Name = OUTPUT_SHEET_NAME
    Set wsLegend = ThisWorkbook.Worksheets.Add(After:=wsOutput): wsLegend.Name = LEGEND_SHEET_NAME
    SetupWorksheets = True
End Function

'===================================================================================================
' �X�C�����[���i�w�i�̘g�g�݁j�̕`��
'===================================================================================================
Private Sub DrawSwimlanes()
    If wsOutput Is Nothing Then Exit Sub
    Dim currentY As Long: currentY = Y_START
    Dim laneColors As Object: Set laneColors = CreateObject("Scripting.Dictionary")
    laneColors.Add "���p��", RGB(220, 220, 220): laneColors.Add "�E��-�ϑ��E�h��", RGB(200, 220, 255)
    laneColors.Add "�E��-��v�N�x", RGB(180, 200, 255): laneColors.Add "�E��-�ĔC�p", RGB(160, 180, 255)
    laneColors.Add "�E��-���", RGB(140, 160, 255): laneColors.Add "�������E�O���@��", RGB(200, 230, 200)
    laneColors.Add "���̑�", RGB(230, 230, 210)
    Dim pools As Object: Set pools = CreateObject("Scripting.Dictionary")
    pools.Add "���p��", Array("���p��"): pools.Add "�E��", Array("�E��-�ϑ��E�h��", "�E��-��v�N�x", "�E��-�ĔC�p", "�E��-���")
    pools.Add "�������E�O���@��", Array("�������E�O���@��"): pools.Add "���̑�", Array("���̑�")
    Dim poolName As Variant
    For Each poolName In pools.Keys
        Dim lanesInPool As Variant: lanesInPool = pools(poolName)
        Dim poolTop As Long: poolTop = currentY: Dim poolHeight As Long: poolHeight = 0
        Dim i As Long
        For i = LBound(lanesInPool) To UBound(lanesInPool)
            Dim laneName As String: laneName = lanesInPool(i)
            DrawLane laneName, currentY, Y_LANE_HEIGHT, laneColors(laneName)
            laneLayout.Add laneName, Array(currentY, Y_LANE_HEIGHT)
            currentY = currentY + Y_LANE_HEIGHT: poolHeight = poolHeight + Y_LANE_HEIGHT
        Next i
        DrawPool CStr(poolName), poolTop, poolHeight, laneColors(lanesInPool(0))
    Next poolName
End Sub

' �v�[���i�u���p�ҁv��u�E���v�Ȃǂ̑傫�Șg�j��`�悵�܂��B
Private Sub DrawPool(poolName As String, top As Long, height As Long, bgColor As Long)
    Dim poolRect As Shape, poolText As Shape
    Set poolRect = wsOutput.shapes.AddShape(msoShapeRectangle, 0, top, 20000, height)
    With poolRect: .Name = "PoolRect_" & poolName: .Fill.Visible = msoFalse: .Line.Weight = 2: .Line.ForeColor.RGB = RGB(80, 80, 80): End With
    Set poolText = wsOutput.shapes.AddTextbox(msoTextOrientationHorizontal, 0, top, POOL_TITLE_WIDTH, height)
    With poolText: .Fill.ForeColor.RGB = bgColor: .Line.Visible = msoFalse
        With .TextFrame2: .TextRange.Text = poolName: .VerticalAnchor = msoAnchorMiddle
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
            With .TextRange.Font: .Name = FONT_NAME: .Bold = True: .Size = 12: End With
        End With
    End With
End Sub

' ���[���i�e�S���҂̍s�j��`�悵�܂��B
Private Sub DrawLane(laneName As String, top As Long, height As Long, bgColor As Long)
    Dim laneRect As Shape, laneText As Shape
    Set laneRect = wsOutput.shapes.AddShape(msoShapeRectangle, POOL_TITLE_WIDTH, top, 20000 - POOL_TITLE_WIDTH, height)
    With laneRect: .Name = "LaneLine_" & laneName: .Fill.Visible = msoFalse: .Line.ForeColor.RGB = RGB(180, 180, 180): .Line.DashStyle = msoLineDash: End With
    Set laneText = wsOutput.shapes.AddTextbox(msoTextOrientationHorizontal, POOL_TITLE_WIDTH, top, LANE_HEADER_WIDTH, height)
    With laneText: .Fill.ForeColor.RGB = bgColor: .Line.Visible = msoTrue: .Line.ForeColor.RGB = RGB(255, 255, 255)
        With .TextFrame2: .TextRange.Text = laneName: .VerticalAnchor = msoAnchorMiddle
            .Orientation = msoTextOrientationHorizontal
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
            With .TextRange.Font: .Name = FONT_NAME: .Size = 10: End With
        End With
    End With
End Sub

'===================================================================================================
' �}�`�̕`��Ɛڑ�
'===================================================================================================
' �t���[�}�̖{�̂ł���}�`��`�悵�A��������Ō��т܂��B
Private Sub DrawAndConnectShapes_Absolute()
    ' �X�e�b�v1: �S�Ă̐}�`��`�悷��
    Dim id As String
    Dim i As Long
    For i = LBound(taskOrder) To UBound(taskOrder)
        id = taskOrder(i)
        
        If Not taskData.Exists(id) Then GoTo NextTask
        Dim currentTask As Object: Set currentTask = taskData(id)
        
        ' �S���҂����肷��
        Dim who As String: who = currentTask("who")
        If who = "" Then
            ' �S���҂��󗓂̃^�X�N�i�����Ȃǁj�́u���̑��v���[���ɔz�u����
            who = "���̑�"
        ElseIf Not laneLayout.Exists(who) Then
            ' ��`����Ă��Ȃ��S���҂��u���̑��v���[���ɔz�u����
            who = "���̑�"
        Else
            lastUsedLane = who
        End If
        
        If laneLayout.Exists(who) Then
            ' 1. X���W�i���ʒu�j�̌���F�菇�ԍ�(ID)�����ΓI�Ȉʒu���v�Z���܂��B
            Dim numericId As Long: numericId = Val(id)
            Dim xPos As Long: xPos = X_START + (numericId - 1) * X_STEP

            ' 2. Y���W�i�c�ʒu�j�̌���F��{�̓��[���̒����ɔz�u���܂��B
            Dim laneInfo As Variant: laneInfo = laneLayout(who)
            Dim yPos As Long: yPos = laneInfo(0) + (laneInfo(1) - SHAPE_HEIGHT) / 2
            
            ' ����E�����Ȃǂ̓���Ȑ}�`�́A���₷���̂��߂ɏ㉺�ɂ��炵�܂��B
            Dim isSpecialShape As Boolean
            Dim isGateway As Boolean
            isGateway = InStr(currentTask("flowElement"), "����") > 0 Or InStr(currentTask("flowElement"), "����") > 0
            isSpecialShape = isGateway Or currentTask("toIds") <> ""
            
            If isSpecialShape Then
                If numericId Mod 2 = 1 Then ' �菇�ԍ�����Ȃ��ɂ��炷
                    yPos = yPos - Y_SHIFT_SPECIAL
                Else ' �菇�ԍ��������Ȃ牺�ɂ��炷
                    yPos = yPos + Y_SHIFT_SPECIAL
                End If
            End If
            
            ' 3. �}�`��`�悷��
            Dim newShape As Shape: Set newShape = DrawShape(wsOutput, id, currentTask, xPos, yPos)
            shapeCollection.Add id, newShape
        End If
NextTask:
    Next i
    
    ' �X�e�b�v2: �S�Ă̐}�`��ڑ�����
    For i = LBound(taskOrder) To UBound(taskOrder)
        ConnectShape taskOrder(i)
    Next i
End Sub

' ���ۂɈ�̐}�`��`�悷��֐��ł��B
Private Function DrawShape(targetSheet As Worksheet, id As String, currentTask As Object, x As Long, y As Long, Optional isLegendItem As Boolean = False) As Shape
    Dim flowElement As String: flowElement = currentTask("flowElement")
    Dim taskType As String: taskType = currentTask("taskType")
    Dim txt As String: txt = currentTask("summary")
    Dim markerText As String: markerText = ""
    If Not isLegendItem Then
        Select Case True
            Case InStr(taskType, "�p�\�R�����") > 0: markerText = "U: "
            Case InStr(taskType, "����") > 0: markerText = "M: "
            Case InStr(taskType, "���t") > 0: markerText = "S: "
            Case InStr(taskType, "�󂯎��") > 0: markerText = "R: "
            Case InStr(flowElement, "YES/NO") > 0: markerText = "X: "
            Case InStr(flowElement, "���s����") > 0: markerText = "+: "
            Case InStr(flowElement, "�����I��") > 0: markerText = "O: "
        End Select
    End If
    Dim baseShape As Shape
    Dim shpType As MsoAutoShapeType, shpColor As Long, shpWeight As Single
    shpColor = RGB(255, 255, 255): shpWeight = 1.5
    Select Case True
        Case InStr(flowElement, "�J�n") > 0: shpType = msoShapeOval
        Case InStr(flowElement, "�I��") > 0: shpType = msoShapeOval: shpWeight = 3
        Case InStr(flowElement, "����") > 0, InStr(flowElement, "����") > 0: shpType = msoShapeDiamond: shpColor = RGB(240, 240, 240)
        Case Else: shpType = msoShapeRoundedRectangle
    End Select
    Set baseShape = targetSheet.shapes.AddShape(shpType, x, y, SHAPE_WIDTH, SHAPE_HEIGHT)
    With baseShape
        .Name = "Shape_" & id
        .Line.ForeColor.RGB = RGB(0, 0, 0): .Line.Weight = shpWeight: .Fill.ForeColor.RGB = shpColor
        With .TextFrame2
            .VerticalAnchor = msoAnchorMiddle: .WordWrap = msoTrue
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
            If id <> "" Then .TextRange.Text = markerText & id & ". " & vbCrLf & txt Else .TextRange.Text = markerText & txt
            With .TextRange.Font: .Name = FONT_NAME: .Size = 10: .Fill.ForeColor.RGB = RGB(0, 0, 0): End With
            If markerText <> "" Then
                With .TextRange.Characters(1, Len(markerText)).Font: .Bold = msoTrue: End With
            End If
        End With
    End With
    Set DrawShape = baseShape
End Function

'===================================================================================================
' �}��쐬�A�ڑ��A�⏕�֐�
'===================================================================================================
' �}��V�[�g���쐬���܂��B
Private Sub CreateLegendSheet()
    If wsLegend Is Nothing Then Exit Sub
    wsLegend.Cells.Clear
    With wsLegend.Range("B2"): .Value = "�Ɩ��t���[�} �}��": .Font.Bold = True: .Font.Size = 14: .Font.Name = FONT_NAME: End With
    Dim currentY As Long: currentY = 50
    Dim legendItems As Object: Set legendItems = CreateObject("Scripting.Dictionary")
    legendItems.Add "�Ɩ��̊J�n�i�J�n�C�x���g�j", "�v���Z�X�̊J�n�n�_": legendItems.Add "�Ɩ��̏I���i�I���C�x���g�j", "�v���Z�X�̏I���n�_"
    legendItems.Add "��Ɓi�^�X�N�j", "�S���҂��s����̓I�ȍ��": legendItems.Add "���Ɓi�}�j���A���^�X�N�j", "M: �V�X�e�����g��Ȃ�����"
    legendItems.Add "�p�\�R����Ɓi���[�U�^�X�N�j", "U: �l���V�X�e�����g���čs�����": legendItems.Add "���[���╶���̑��t�i���M�^�X�N�j", "S: ���𑗐M������"
    legendItems.Add "���[���╶���̎󂯎��i��M�^�X�N�j", "R: ������M������": legendItems.Add "�������� YES/NO�i�r���Q�[�g�E�F�C�j", "X: �����ɂ���ė��ꂪ��ɕ������"
    legendItems.Add "�����I���ł��镪��i��܃Q�[�g�E�F�C�j", "O: �����ɂ���ė��ꂪ�����ɕ������": legendItems.Add "��Ăɕ��s�����i����Q�[�g�E�F�C�j", "+: �����̍�Ƃ𓯎��ɐi�߂�"
    legendItems.Add "����̍����i�Q�[�g�E�F�C�j", "�����ꂽ���ꂪ��ɖ߂�n�_"
    Dim key As Variant
    For Each key In legendItems.Keys
        Dim dummyTask As Object: Set dummyTask = CreateObject("Scripting.Dictionary")
        dummyTask.Add "flowElement", key: dummyTask.Add "taskType", key: dummyTask.Add "summary", legendItems(key)
        Dim shp As Shape: Set shp = DrawShape(wsLegend, "", dummyTask, 50, currentY, isLegendItem:=True)
        shp.TextFrame2.TextRange.Text = legendItems(key)
        currentY = currentY + shp.height + 20
    Next key
    wsLegend.Columns("B:C").AutoFit
End Sub

' �w�肳�ꂽ�}�`���玟�̐}�`�֐��������܂��B
Private Sub ConnectShape(fromId As String)
    If Not shapeCollection.Exists(fromId) Or Not taskData.Exists(fromId) Then Exit Sub
    Dim fromShape As Shape: Set fromShape = shapeCollection(fromId)
    Dim toIdsRaw As String: toIdsRaw = taskData(fromId)("toIds")
    If InStr(taskData(fromId)("flowElement"), "�I��") > 0 Then Exit Sub
    If toIdsRaw = "" Then
        Dim nextId As String: nextId = GetNextTaskID(fromId)
        If nextId <> "" And shapeCollection.Exists(nextId) Then ConnectTwoShapes fromShape, shapeCollection(nextId)
    Else
        Dim toIdArray As Variant: toIdArray = Split(toIdsRaw, ",")
        Dim j As Long, branchIndex As Long: branchIndex = 0
        For j = LBound(toIdArray) To UBound(toIdArray)
            Dim targetId As String: targetId = Trim(CStr(toIdArray(j)))
            If shapeCollection.Exists(targetId) Then
                ConnectTwoShapes fromShape, shapeCollection(targetId), branchIndex
                branchIndex = branchIndex + 1
            End If
        Next j
    End If
End Sub

' 2�̐}�`����Ō��ԋ�̓I�ȏ����ł��B
Private Sub ConnectTwoShapes(shp1 As Shape, shp2 As Shape, Optional branchIndex As Long = 0)
    If shp1 Is Nothing Or shp2 Is Nothing Then Exit Sub
    If shp1.ConnectionSiteCount = 0 Or shp2.ConnectionSiteCount = 0 Then Exit Sub
    
    Dim conn As Shape
    Set conn = wsOutput.shapes.AddConnector(CONNECTOR_TYPE, 0, 0, 100, 100)
    
    Dim beginConnSite As Long, endConnSite As Long
    
    ' �ڑ����̐}�`���Q�[�g�E�F�C(�Ђ��`)�̏ꍇ�A���͕K���E������o��悤�ɂ��܂��B
    If shp1.AutoShapeType = msoShapeDiamond Then
        beginConnSite = 3 ' �E
        endConnSite = 1   ' ��
    Else
        ' ����ȊO�̐}�`�́A�����ɉ����čœK�Ȑڑ��_��I�т܂��B
        If shp2.top > shp1.top + shp1.height / 2 Then
            beginConnSite = 4 ' ��
            endConnSite = 2   ' ��
        ElseIf shp1.top > shp2.top + shp2.height / 2 Then
            beginConnSite = 2 ' ��
            endConnSite = 4   ' ��
        Else
            beginConnSite = 3 ' �E
            endConnSite = 1   ' ��
        End If
    End If
    
    conn.ConnectorFormat.BeginConnect shp1, beginConnSite
    conn.ConnectorFormat.EndConnect shp2, endConnSite
    
    If conn.ConnectorFormat.BeginConnected And conn.ConnectorFormat.EndConnected Then
        With conn.Line: .EndArrowheadStyle = msoArrowheadTriangle: .ForeColor.RGB = RGB(89, 89, 89): .Weight = 1.5: End With
        conn.RerouteConnections: conn.ZOrder msoSendToBack
    Else
        conn.Delete
    End If
End Sub

' �S�Ă̐}�`�����܂�悤�ɁA�w�i�̘g�̕��𒲐����܂��B
Private Sub AdjustPoolWidths()
    If wsOutput Is Nothing Or shapeCollection.count = 0 Then Exit Sub
    Dim shp As Shape, maxRight As Single: maxRight = 0: Dim shpVar As Variant
    For Each shpVar In shapeCollection.Items
        Set shp = shpVar
        If shp.Left + shp.Width > maxRight Then maxRight = shp.Left + shp.Width
    Next shpVar
    Dim newWidth As Single: newWidth = maxRight + RIGHT_MARGIN
    On Error Resume Next
    For Each shp In wsOutput.shapes
        If Left(shp.Name, 9) = "PoolRect_" Then shp.Width = newWidth
        If Left(shp.Name, 9) = "LaneLine_" Then shp.Width = newWidth - POOL_TITLE_WIDTH
    Next shp
    On Error GoTo 0
End Sub

' �u���̎菇�ԍ��v���󗓂̏ꍇ�ɁA����ID��T���֐��ł��BID����є�тł��Ή��ł��܂��B
Private Function GetNextTaskID(currentId As String) As String
    GetNextTaskID = "" ' �f�t�H���g�͋󕶎�
    Dim currentNumericId As Long: currentNumericId = Val(currentId)
    Dim smallestLargerId As Long: smallestLargerId = -1
    
    ' �S�Ẵ^�X�NID�̒�����A���݂�ID���傫�����̂̂����A�ł����������̂�T���܂��B
    Dim id As Variant
    For Each id In taskData.Keys
        Dim numericId As Long: numericId = Val(id)
        
        If numericId > currentNumericId Then
            If smallestLargerId = -1 Or numericId < smallestLargerId Then
                smallestLargerId = numericId
            End If
        End If
    Next id
    
    ' ���������ꍇ�A����𕶎���Ƃ��ĕԂ��܂��B
    If smallestLargerId <> -1 Then
        GetNextTaskID = CStr(smallestLargerId)
    End If
End Function
