Attribute VB_Name = "FlowchartGenerator"
Option Explicit

' --- 設定項目 ---
' このツールが参照するシート名や、図の見た目を設定します。
Private Const INPUT_SHEET_NAME As String = "業務リスト"
Private Const OUTPUT_SHEET_NAME As String = "業務フロー図"
Private Const LEGEND_SHEET_NAME As String = "凡例"
Private Const FONT_NAME As String = "Meiryo UI"

' --- レイアウト設定 ---
' 図形のサイズや間隔など、デザインに関する詳細設定です。
Private Const POOL_TITLE_WIDTH As Long = 120    ' 左端の「利用者」などのタイトルの幅
Private Const LANE_HEADER_WIDTH As Long = 100   ' 各担当者（レーン）のタイトルの幅
Private Const SHAPE_WIDTH As Long = 220         ' 各手順（図形）の幅
Private Const SHAPE_HEIGHT As Long = 110        ' 各手順（図形）の高さ
Private Const X_START As Long = POOL_TITLE_WIDTH + LANE_HEADER_WIDTH + 20 ' 図形を描き始めるX座標
Private Const Y_START As Long = 30              ' 図形を描き始めるY座標
Private Const X_STEP As Long = 260              ' 図形と図形の水平方向の間隔
Private Const Y_LANE_HEIGHT As Long = 280       ' 各担当者（レーン）の高さ
Private Const Y_SHIFT_SPECIAL As Long = 40      ' 分岐など特殊な図形を上下にずらす距離
Private Const RIGHT_MARGIN As Long = 50         ' フロー図全体の右側の余白

' --- グローバル変数 ---
' マクロ全体で共通して使用する変数です。
Private wsInput As Worksheet, wsOutput As Worksheet, wsLegend As Worksheet
Private laneLayout As Object, taskData As Object, shapeCollection As Object
Private taskOrder() As String
Private lastUsedLane As String
Private CONNECTOR_TYPE As MsoConnectorType

'===================================================================================================
' ■■■ メイン処理 ■■■
' このマクロを実行すると、フロー図の作成が始まります。
'===================================================================================================
Sub CreateFlowChart()
    On Error GoTo GenericErrorHandler
    Application.ScreenUpdating = False

    InitializeGlobalVariables
    If Not SetupWorksheets() Then GoTo Cleanup
    If Not LoadTaskData() Then
        MsgBox "「" & INPUT_SHEET_NAME & "」シートにデータがありません。", vbExclamation
        GoTo Cleanup
    End If

    CreateLegendSheet
    DrawSwimlanes
    
    ' シンプルで美しいレイアウトを実現する描画・接続処理
    DrawAndConnectShapes_Absolute
    
    AdjustPoolWidths

    If Not wsOutput Is Nothing Then
        wsOutput.Activate
        ActiveWindow.Zoom = 75
    End If

    MsgBox "業務フロー図、及び凡例シートの作成が完了いたしました。", vbInformation

Cleanup:
    Application.ScreenUpdating = True
    CleanupGlobalVariables
    Exit Sub
GenericErrorHandler:
    MsgBox "エラーが発生しました。" & vbCrLf & "エラー番号: " & Err.Number & vbCrLf & "エラー内容: " & Err.Description, vbCritical, "エラー発生"
    Resume Cleanup
End Sub

'===================================================================================================
' 初期化とデータ読み込み
'===================================================================================================
Private Sub InitializeGlobalVariables()
    Set laneLayout = CreateObject("Scripting.Dictionary")
    Set taskData = CreateObject("Scripting.Dictionary")
    Set shapeCollection = CreateObject("Scripting.Dictionary")
    ReDim taskOrder(0)
    lastUsedLane = ""
    ' 接続線の種類を設定します。msoConnectorCurve(曲線) または msoConnectorElbow(カクカクした線)が選べます。
    CONNECTOR_TYPE = msoConnectorCurve
End Sub

Private Sub CleanupGlobalVariables()
    Set wsInput = Nothing: Set wsOutput = Nothing: Set wsLegend = Nothing
    Set laneLayout = Nothing: Set taskData = Nothing: Set shapeCollection = Nothing
    Erase taskOrder
End Sub

' 「業務リスト」シートからデータを読み込みます。
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
            taskInfo.Add "toIds", Replace(Replace(Trim(CStr(dataArr(i, 10))), "、", ","), " ", "")
            taskData.Add id, taskInfo: taskOrder(taskCount) = id
        End If
    Next i
    
    If taskCount = 0 Then Exit Function
    ReDim Preserve taskOrder(1 To taskCount)
    LoadTaskData = True
End Function

' フロー図と凡例を描画するための新しいシートを準備します。
Private Function SetupWorksheets() As Boolean
    On Error Resume Next: Set wsInput = ThisWorkbook.Worksheets(INPUT_SHEET_NAME)
    If wsInput Is Nothing Then MsgBox "「" & INPUT_SHEET_NAME & "」シートが見つかりません。", vbCritical: Exit Function
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
' スイムレーン（背景の枠組み）の描画
'===================================================================================================
Private Sub DrawSwimlanes()
    If wsOutput Is Nothing Then Exit Sub
    Dim currentY As Long: currentY = Y_START
    Dim laneColors As Object: Set laneColors = CreateObject("Scripting.Dictionary")
    laneColors.Add "利用者", RGB(220, 220, 220): laneColors.Add "職員-委託・派遣", RGB(200, 220, 255)
    laneColors.Add "職員-会計年度", RGB(180, 200, 255): laneColors.Add "職員-再任用", RGB(160, 180, 255)
    laneColors.Add "職員-常勤", RGB(140, 160, 255): laneColors.Add "他部署・外部機関", RGB(200, 230, 200)
    laneColors.Add "その他", RGB(230, 230, 210)
    Dim pools As Object: Set pools = CreateObject("Scripting.Dictionary")
    pools.Add "利用者", Array("利用者"): pools.Add "職員", Array("職員-委託・派遣", "職員-会計年度", "職員-再任用", "職員-常勤")
    pools.Add "他部署・外部機関", Array("他部署・外部機関"): pools.Add "その他", Array("その他")
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

' プール（「利用者」や「職員」などの大きな枠）を描画します。
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

' レーン（各担当者の行）を描画します。
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
' 図形の描画と接続
'===================================================================================================
' フロー図の本体である図形を描画し、それらを線で結びます。
Private Sub DrawAndConnectShapes_Absolute()
    ' ステップ1: 全ての図形を描画する
    Dim id As String
    Dim i As Long
    For i = LBound(taskOrder) To UBound(taskOrder)
        id = taskOrder(i)
        
        If Not taskData.Exists(id) Then GoTo NextTask
        Dim currentTask As Object: Set currentTask = taskData(id)
        
        ' 担当者を決定する
        Dim who As String: who = currentTask("who")
        If who = "" Then
            ' 担当者が空欄のタスク（合流など）は「その他」レーンに配置する
            who = "その他"
        ElseIf Not laneLayout.Exists(who) Then
            ' 定義されていない担当者も「その他」レーンに配置する
            who = "その他"
        Else
            lastUsedLane = who
        End If
        
        If laneLayout.Exists(who) Then
            ' 1. X座標（横位置）の決定：手順番号(ID)から絶対的な位置を計算します。
            Dim numericId As Long: numericId = Val(id)
            Dim xPos As Long: xPos = X_START + (numericId - 1) * X_STEP

            ' 2. Y座標（縦位置）の決定：基本はレーンの中央に配置します。
            Dim laneInfo As Variant: laneInfo = laneLayout(who)
            Dim yPos As Long: yPos = laneInfo(0) + (laneInfo(1) - SHAPE_HEIGHT) / 2
            
            ' 分岐・合流などの特殊な図形は、見やすさのために上下にずらします。
            Dim isSpecialShape As Boolean
            Dim isGateway As Boolean
            isGateway = InStr(currentTask("flowElement"), "分岐") > 0 Or InStr(currentTask("flowElement"), "合流") > 0
            isSpecialShape = isGateway Or currentTask("toIds") <> ""
            
            If isSpecialShape Then
                If numericId Mod 2 = 1 Then ' 手順番号が奇数なら上にずらす
                    yPos = yPos - Y_SHIFT_SPECIAL
                Else ' 手順番号が偶数なら下にずらす
                    yPos = yPos + Y_SHIFT_SPECIAL
                End If
            End If
            
            ' 3. 図形を描画する
            Dim newShape As Shape: Set newShape = DrawShape(wsOutput, id, currentTask, xPos, yPos)
            shapeCollection.Add id, newShape
        End If
NextTask:
    Next i
    
    ' ステップ2: 全ての図形を接続する
    For i = LBound(taskOrder) To UBound(taskOrder)
        ConnectShape taskOrder(i)
    Next i
End Sub

' 実際に一つの図形を描画する関数です。
Private Function DrawShape(targetSheet As Worksheet, id As String, currentTask As Object, x As Long, y As Long, Optional isLegendItem As Boolean = False) As Shape
    Dim flowElement As String: flowElement = currentTask("flowElement")
    Dim taskType As String: taskType = currentTask("taskType")
    Dim txt As String: txt = currentTask("summary")
    Dim markerText As String: markerText = ""
    If Not isLegendItem Then
        Select Case True
            Case InStr(taskType, "パソコン作業") > 0: markerText = "U: "
            Case InStr(taskType, "手作業") > 0: markerText = "M: "
            Case InStr(taskType, "送付") > 0: markerText = "S: "
            Case InStr(taskType, "受け取り") > 0: markerText = "R: "
            Case InStr(flowElement, "YES/NO") > 0: markerText = "X: "
            Case InStr(flowElement, "並行処理") > 0: markerText = "+: "
            Case InStr(flowElement, "複数選択") > 0: markerText = "O: "
        End Select
    End If
    Dim baseShape As Shape
    Dim shpType As MsoAutoShapeType, shpColor As Long, shpWeight As Single
    shpColor = RGB(255, 255, 255): shpWeight = 1.5
    Select Case True
        Case InStr(flowElement, "開始") > 0: shpType = msoShapeOval
        Case InStr(flowElement, "終了") > 0: shpType = msoShapeOval: shpWeight = 3
        Case InStr(flowElement, "分岐") > 0, InStr(flowElement, "合流") > 0: shpType = msoShapeDiamond: shpColor = RGB(240, 240, 240)
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
' 凡例作成、接続、補助関数
'===================================================================================================
' 凡例シートを作成します。
Private Sub CreateLegendSheet()
    If wsLegend Is Nothing Then Exit Sub
    wsLegend.Cells.Clear
    With wsLegend.Range("B2"): .Value = "業務フロー図 凡例": .Font.Bold = True: .Font.Size = 14: .Font.Name = FONT_NAME: End With
    Dim currentY As Long: currentY = 50
    Dim legendItems As Object: Set legendItems = CreateObject("Scripting.Dictionary")
    legendItems.Add "業務の開始（開始イベント）", "プロセスの開始地点": legendItems.Add "業務の終了（終了イベント）", "プロセスの終了地点"
    legendItems.Add "作業（タスク）", "担当者が行う具体的な作業": legendItems.Add "手作業（マニュアルタスク）", "M: システムを使わない手作業"
    legendItems.Add "パソコン作業（ユーザタスク）", "U: 人がシステムを使って行う作業": legendItems.Add "メールや文書の送付（送信タスク）", "S: 情報を送信する作業"
    legendItems.Add "メールや文書の受け取り（受信タスク）", "R: 情報を受信する作業": legendItems.Add "条件分岐 YES/NO（排他ゲートウェイ）", "X: 条件によって流れが一つに分かれる"
    legendItems.Add "複数選択できる分岐（包含ゲートウェイ）", "O: 条件によって流れが複数に分かれる": legendItems.Add "一斉に並行処理（並列ゲートウェイ）", "+: 複数の作業を同時に進める"
    legendItems.Add "流れの合流（ゲートウェイ）", "分かれた流れが一つに戻る地点"
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

' 指定された図形から次の図形へ線を引きます。
Private Sub ConnectShape(fromId As String)
    If Not shapeCollection.Exists(fromId) Or Not taskData.Exists(fromId) Then Exit Sub
    Dim fromShape As Shape: Set fromShape = shapeCollection(fromId)
    Dim toIdsRaw As String: toIdsRaw = taskData(fromId)("toIds")
    If InStr(taskData(fromId)("flowElement"), "終了") > 0 Then Exit Sub
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

' 2つの図形を線で結ぶ具体的な処理です。
Private Sub ConnectTwoShapes(shp1 As Shape, shp2 As Shape, Optional branchIndex As Long = 0)
    If shp1 Is Nothing Or shp2 Is Nothing Then Exit Sub
    If shp1.ConnectionSiteCount = 0 Or shp2.ConnectionSiteCount = 0 Then Exit Sub
    
    Dim conn As Shape
    Set conn = wsOutput.shapes.AddConnector(CONNECTOR_TYPE, 0, 0, 100, 100)
    
    Dim beginConnSite As Long, endConnSite As Long
    
    ' 接続元の図形がゲートウェイ(ひし形)の場合、線は必ず右側から出るようにします。
    If shp1.AutoShapeType = msoShapeDiamond Then
        beginConnSite = 3 ' 右
        endConnSite = 1   ' 左
    Else
        ' それ以外の図形は、高さに応じて最適な接続点を選びます。
        If shp2.top > shp1.top + shp1.height / 2 Then
            beginConnSite = 4 ' 下
            endConnSite = 2   ' 上
        ElseIf shp1.top > shp2.top + shp2.height / 2 Then
            beginConnSite = 2 ' 上
            endConnSite = 4   ' 下
        Else
            beginConnSite = 3 ' 右
            endConnSite = 1   ' 左
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

' 全ての図形が収まるように、背景の枠の幅を調整します。
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

' 「次の手順番号」が空欄の場合に、次のIDを探す関数です。IDが飛び飛びでも対応できます。
Private Function GetNextTaskID(currentId As String) As String
    GetNextTaskID = "" ' デフォルトは空文字
    Dim currentNumericId As Long: currentNumericId = Val(currentId)
    Dim smallestLargerId As Long: smallestLargerId = -1
    
    ' 全てのタスクIDの中から、現在のIDより大きいもののうち、最も小さいものを探します。
    Dim id As Variant
    For Each id In taskData.Keys
        Dim numericId As Long: numericId = Val(id)
        
        If numericId > currentNumericId Then
            If smallestLargerId = -1 Or numericId < smallestLargerId Then
                smallestLargerId = numericId
            End If
        End If
    Next id
    
    ' 見つかった場合、それを文字列として返します。
    If smallestLargerId <> -1 Then
        GetNextTaskID = CStr(smallestLargerId)
    End If
End Function
