Attribute VB_Name = "LLMExporter"
Option Explicit

' --- 外部ライブラリを使わずにテキストファイルを出力するための設定 ---
Private Const adSaveCreateOverWrite = 2
Private Const adTypeText = 2

' --- 設定項目 ---
' ※モジュール1で定義済みの定数を参照します。
Private Const FILE_PREFIX As String = "住民票交付業務フロー"


'===================================================================================================
' ■■■ メイン処理（LLM連携ファイル出力） ■■■
' LLM連携用のマークダウンファイルを2種類出力します。
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

    ' 1. 入力シートの存在確認とデータ読み込み
    Set wsInput = Nothing
    On Error Resume Next
    ' モジュール1の定数を参照
    Set wsInput = ThisWorkbook.Worksheets("業務リスト")
    On Error GoTo ErrorHandler

    If wsInput Is Nothing Then
        MsgBox "「業務リスト」シートが見つかりません。", vbCritical
        GoTo Cleanup
    End If

    lastRow = wsInput.Cells(wsInput.Rows.count, "A").End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "「業務リスト」シートにデータがありません。", vbExclamation
        GoTo Cleanup
    End If

    ' ヘッダーを除いたデータを配列に一括読み込み
    dataArr = wsInput.Range("A2:J" & lastRow).Value

    ' 2. 各マークダウンコンテンツの生成
    manualMarkdown = GenerateManualMarkdown(dataArr)
    analysisMarkdown = GenerateAnalysisMarkdown(dataArr)

    ' 3. ファイルの保存 (必ずデスクトップに保存するよう変更)
    saveFolder = CreateObject("WScript.Shell").SpecialFolders("Desktop")

    manualFilePath = saveFolder & "\" & FILE_PREFIX & "(マニュアル設計用).md"
    analysisFilePath = saveFolder & "\" & FILE_PREFIX & "(課題分析用).md"

    If SaveTextToFile(manualMarkdown, manualFilePath) And SaveTextToFile(analysisMarkdown, analysisFilePath) Then
        MsgBox "LLM連携用のマークダウンファイル2点をデスクトップに出力しました。" & vbCrLf & vbCrLf & _
               "・" & manualFilePath & vbCrLf & _
               "・" & analysisFilePath, vbInformation, "出力完了"
    Else
        MsgBox "ファイルの出力中にエラーが発生しました。", vbCritical, "エラー"
    End If

Cleanup:
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    MsgBox "エラーが発生しました。" & vbCrLf & "エラー番号: " & Err.Number & vbCrLf & "エラー内容: " & Err.Description, vbCritical, "エラー発生"
    Resume Cleanup
End Sub

'===================================================================================================
' 業務マニュアル作成用のマークダウンを生成する
'===================================================================================================
Private Function GenerateManualMarkdown(dataArr As Variant) As String
    'StringBuilderを使わず、通常の文字列結合に変更
    Dim markdownContent As String
    Dim i As Long

    ' --- プロンプト部分 ---
    markdownContent = markdownContent & "# 指示" & vbCrLf
    markdownContent = markdownContent & "あなたは、プロのテクニカルライターです。以下の業務プロセス情報に基づいて、新人職員や利用者が読んでも理解できるような、丁寧で分かりやすい業務マニュアルを作成してください。" & vbCrLf & vbCrLf
    markdownContent = markdownContent & "## 作成するマニュアルの要件" & vbCrLf
    markdownContent = markdownContent & "- 業務の全体像が最初にわかるように「はじめに」と「業務の流れ」を記述してください。" & vbCrLf
    markdownContent = markdownContent & "- 「詳細な手順」では、各手順を具体的に説明してください。" & vbCrLf
    markdownContent = markdownContent & "- 担当者が「利用者」と「職員」に分かれていることを明確にし、それぞれの視点で行動が分かるように記述してください。" & vbCrLf
    markdownContent = markdownContent & "- 専門用語は避け、平易な言葉で解説してください。" & vbCrLf & vbCrLf

    ' --- データ部分 ---
    markdownContent = markdownContent & "# 業務プロセス情報" & vbCrLf
    markdownContent = markdownContent & "| 手順番号 | 担当者 | 作業や判断の内容 | 補足説明 |" & vbCrLf
    markdownContent = markdownContent & "|:---|:---|:---|:---|" & vbCrLf

    For i = 1 To UBound(dataArr, 1)
        markdownContent = markdownContent & "| " & CStr(dataArr(i, 1)) '手順番号
        markdownContent = markdownContent & " | " & CStr(dataArr(i, 2)) '担当者
        markdownContent = markdownContent & " | " & CStr(dataArr(i, 3)) '作業や判断の内容
        markdownContent = markdownContent & " | " & CStr(dataArr(i, 7)) '補足説明
        markdownContent = markdownContent & " |" & vbCrLf
    Next i

    GenerateManualMarkdown = markdownContent
End Function

'===================================================================================================
' 課題分析用のマークダウンを生成する
'===================================================================================================
Private Function GenerateAnalysisMarkdown(dataArr As Variant) As String
    'StringBuilderを使わず、通常の文字列結合に変更
    Dim markdownContent As String
    Dim i As Long

    ' --- プロンプト部分 ---
    markdownContent = markdownContent & "# 指示" & vbCrLf
    markdownContent = markdownContent & "あなたは、経験豊富な業務改善コンサルタントです。以下の業務プロセス情報と、各手順で報告されている「困りごと・課題」を分析し、具体的な改善提案をしてください。" & vbCrLf & vbCrLf
    markdownContent = markdownContent & "## 分析と提案の要件" & vbCrLf
    markdownContent = markdownContent & "- まず、報告されている課題を要約し、根本原因がどこにあるかを分析してください。" & vbCrLf
    markdownContent = markdownContent & "- 「デジタル化」「プロセスの簡略化」「職員の負担軽減」「利用者の利便性向上」の観点から、具体的な改善アクションを提案してください。" & vbCrLf
    markdownContent = markdownContent & "- 提案は、短期的に実現可能なものと、中長期的に取り組むべきものに分けて提示してください。" & vbCrLf
    markdownContent = markdownContent & "- 提案によってどのような効果（時間短縮、コスト削減、満足度向上など）が期待できるかを記述してください。" & vbCrLf & vbCrLf

    ' --- データ部分 ---
    markdownContent = markdownContent & "# 業務プロセスと課題" & vbCrLf
    markdownContent = markdownContent & "| 手順番号 | 担当者 | 作業や判断の内容 | 困りごと・課題 | 時間や件数 |" & vbCrLf
    markdownContent = markdownContent & "|:---|:---|:---|:---|:---|" & vbCrLf

    For i = 1 To UBound(dataArr, 1)
        ' 「困りごと・課題」が入力されている行のみを抽出
        If Trim(CStr(dataArr(i, 9))) <> "" Then
            markdownContent = markdownContent & "| " & CStr(dataArr(i, 1))  '手順番号
            markdownContent = markdownContent & " | " & CStr(dataArr(i, 2))  '担当者
            markdownContent = markdownContent & " | " & CStr(dataArr(i, 3))  '作業や判断の内容
            markdownContent = markdownContent & " | " & CStr(dataArr(i, 9))  '困りごと・課題
            markdownContent = markdownContent & " | " & CStr(dataArr(i, 8))  '時間や件数
            markdownContent = markdownContent & " |" & vbCrLf
        End If
    Next i

    GenerateAnalysisMarkdown = markdownContent
End Function

'===================================================================================================
' テキストコンテンツをUTF-8でファイルに保存する
'===================================================================================================
Private Function SaveTextToFile(content As String, filePath As String) As Boolean
    Dim adoStream As Object
    On Error GoTo SaveError

    ' ADODB.Streamオブジェクトを使ってUTF-8で保存
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

