Attribute VB_Name = "Module1"
Option Explicit

'=====================================================================
' 【1. 基本設定】ファイル名・シート名・拡張子などの定数
'=====================================================================
Public Const MBO_TARGET_FILE_KEYWORD As String = "MBOシート_人事評価シート" ' 対象となるMBOファイル名のキーワード
Public Const MBO_TARGET_FILE_EXTS As String = "xlsx,xlsm,xlsb,xls" ' 対象拡張子（カンマ区切り文字列）
Public Const MBO_PREV_FOLDER_NAME As String = "前期MBO" ' 前期MBOフォルダ名
Public Const MBO_TARGET_SHEET_NAME As String = "年間総合評価" ' 評価対象シート名
Public Const MBO_INITIAL_SHEET_PREFIX As String = "MBO初期計画一覧_"

'=====================================================================
' 【2. 一覧シートの列インデックス（出力先）】
'=====================================================================
Public Enum MboListColumn
    MboColEmployeeNo = 1            '社員No.
    MboColEmployeeName = 2          '氏名
    MboColDepartment = 3            '所属
    MboColJoinDate = 4              '入社日
    MboColServiceYears = 5          '勤続年数
    MboColTitle = 6                 '役職
    MboColGrade = 7                 '等級
    MboColCareer = 8                '目指すキャリア
    MboColCommittee = 9             '所属委員会
    MboColKbn = 10                  '区分
    MboColPortableSkill = 11        'ポータブルスキル
    MboColRequiredAction = 12       '求められる行動
    MboColGoalNo = 13               '目標No.
    MboColGoalA = 14                '目標設定A
    MboColCriteria = 15             '達成基準
    MboColMeasureB = 16             '目標達成施策手段B
    MboColParentGoalNo = 17         '親目標No.
    MboColTerm = 18                 '次期
    MboColDateStart = 19            '期日開始
    MboColDateEnd = 20              '期日終了
    MboColSelfWeight = 21           '本人ウェイト
    MboColSelfDifficulty = 22       '本人難易度
    MboColBossWeight = 23           '上司ウェイト
    MboColBossDifficulty = 24       '上司難易度
    MboColGoalScore = 25            '目標点数
End Enum

'=====================================================================
' 【3. 一覧シートのヘッダー文字列定数（出力先）】
'=====================================================================
Public Const MBO_HEADER_EMPLOYEE_NO As String = "社員No."
Public Const MBO_HEADER_EMPLOYEE_NAME As String = "氏名"
Public Const MBO_HEADER_DEPARTMENT As String = "所属"
Public Const MBO_HEADER_JOIN_DATE As String = "入社日"
Public Const MBO_HEADER_SERVICE_YEARS As String = "勤続年数"
Public Const MBO_HEADER_TITLE As String = "役職"
Public Const MBO_HEADER_GRADE As String = "等級"
Public Const MBO_HEADER_CAREER As String = "目指すキャリア"
Public Const MBO_HEADER_COMMITTEE As String = "所属委員会"
Public Const MBO_HEADER_KBN As String = "区分"
Public Const MBO_HEADER_PORTABLE_SKILL As String = "ポータブルスキル"
Public Const MBO_HEADER_REQUIRED_ACTION As String = "求められる行動"
Public Const MBO_HEADER_GOAL_NO As String = "目標No."
Public Const MBO_HEADER_GOAL_A As String = "目標設定A"
Public Const MBO_HEADER_CRITERIA As String = "達成基準"
Public Const MBO_HEADER_MEASURE_B As String = "目標達成施策手段B"
Public Const MBO_HEADER_PARENT_GOAL_NO As String = "親目標No."
Public Const MBO_HEADER_TERM As String = "次期"
Public Const MBO_HEADER_DATE_START As String = "期日開始"
Public Const MBO_HEADER_DATE_END As String = "期日終了"
Public Const MBO_HEADER_SELF_WEIGHT As String = "本人ウェイト"
Public Const MBO_HEADER_SELF_DIFFICULTY As String = "本人難易度"
Public Const MBO_HEADER_BOSS_WEIGHT As String = "上司ウェイト"
Public Const MBO_HEADER_BOSS_DIFFICULTY As String = "上司難易度"
Public Const MBO_HEADER_GOAL_SCORE As String = "目標点数"

'=====================================================================
' 【4. 前期MBO評価値取得用：元シートのセル位置・列番号】
'=====================================================================
Public Enum MboPrevSrcRowCol
    MboPrevSrcNameRow = 5      ' 氏名セルの行番号（J5）
    MboPrevSrcNameCol = 10     ' 氏名セルの列番号（J列=10）
    MboPrevSrcMboRow = 10      ' MBO評価値の行番号
    MboPrevSrcCmRow = 13       ' CM評価値の行番号
    MboPrevSrcTotalRow = 17    ' 総合評価値の行番号
    MboPrevSrcRankCol = 3      ' 評価ランクの列番号（C列=3）
    MboPrevSrcPointCol = 5     ' 合計点の列番号（E列=5）
End Enum

'=====================================================================
' 【5. 前期MBO評価値取得用：出力先シートの列番号】
'=====================================================================
Public Enum MboPrevOutCol
    MboPrevOutColMboRank = 26      ' Z列：前期MBOランク
    MboPrevOutColMboPoint = 27     ' AA列：前期MBO点数
    MboPrevOutColCmRank = 28       ' AB列：前期CMランク
    MboPrevOutColCmPoint = 29      ' AC列：前期CM点数
    MboPrevOutColTotalRank = 30    ' AD列：前期総合ランク
    MboPrevOutColTotalPoint = 31   ' AE列：前期総合点数
End Enum

'=====================================================================
' 【6. 前期MBO評価値取得用：出力先シートのヘッダー文字列】
'=====================================================================
Public Const MBO_HEADER_PREV_MBO_RANK As String = "前期 評価ランク（MBOシート 1年間評価）"
Public Const MBO_HEADER_PREV_MBO_POINT As String = "前期 合計点（MBOシート 1年間評価）"
Public Const MBO_HEADER_PREV_CM_RANK As String = "前期 評価ランク（CMシート 最終評価）"
Public Const MBO_HEADER_PREV_CM_POINT As String = "前期 合計点（CMシート 最終評価）"
Public Const MBO_HEADER_PREV_TOTAL_RANK As String = "前期 評価ランク（総合評価 MBO+CM）"
Public Const MBO_HEADER_PREV_TOTAL_POINT As String = "前期 合計点（総合評価 MBO+CM）"

'=====================================================================
' 【7. データ取得用：元シートの行・列インデックス】
'=====================================================================
Public Enum MboGetRow
    MboGetRowEmployeeNo = 7         '社員No.
    MboGetRowEmployeeName = 7       '氏名
    MboGetRowDepartment = 7         '所属
    MboGetRowJoinDate = 7           '入社日
    MboGetRowServiceYears = 7       '勤続年数
    MboGetRowTitle = 7              '役職
    MboGetRowGrade = 7              '等級
    MboGetRowCareer = 9             '目指すキャリア
    MboGetRowCommittee = 11         '所属委員会
End Enum

Public Enum MboGetColumn
    '--- 上部ヘッダ系
    MboGetColEmployeeNo = 3         '社員No.
    MboGetColEmployeeName = 5       '氏名
    MboGetColDepartment = 6         '所属
    MboGetColJoinDate = 7           '入社日
    MboGetColServiceYears = 10      '勤続年数
    MboGetColTitle = 12             '役職
    MboGetColGrade = 14             '等級
    MboGetColCareer = 22            '目指すキャリア (V列)
    MboGetColCommittee = 22         '所属委員会 (V列)

    '--- 目標行(18～24行)の列
    MboGetColKbn = 1                '区分
    MboGetColPortableSkill = 2      'ポータブルスキル
    MboGetColRequiredAction = 3     '求められる行動
    MboGetColGoalNo = 4             '目標No.
    MboGetColGoalA = 5              '目標設定A
    MboGetColMeasureB = 6           '目標達成施策手段B
    MboGetColParentGoalNo = 7       '親目標No.
    MboGetColTerm = 8               '次期
    MboGetColDateStart = 9          '期日開始
    MboGetColDateEnd = 10           '期日終了
    MboGetColSelfWeight = 11        '本人ウェイト
    MboGetColSelfDifficulty = 12    '本人難易度
    MboGetColBossWeight = 13        '上司ウェイト
    MboGetColBossDifficulty = 14    '上司難易度
    MboGetColGoalScore = 15         '目標点数
End Enum

'=====================================================================
' 【8. 固定値（シート名やデータ範囲）】
'=====================================================================
Public Const MBO_SOURCE_SHEET_NAME As String = "目標設定シート" ' データ取得元シート名
Public Const MBO_SOURCE_FIRST_DATA_ROW As Long = 18            ' データ開始行
Public Const MBO_SOURCE_LAST_DATA_ROW As Long = 24             ' データ終了行

Public Const MBO_UPPER_SOURCE_SHEET_NAME As String = "目標管理シート(上期)"
Public Const MBO_UPPER_SHEET_PREFIX As String = "MBO上期評価一覧_"
Public Const MBO_UPPER_HEADER_ROW As Long = 6
Public Const MBO_UPPER_HEADER_VALUE_ROW As Long = 7
Public Const MBO_UPPER_CAREER_ROW As Long = 9
Public Const MBO_UPPER_COMMITTEE_ROW As Long = 11
Public Const MBO_UPPER_DETAIL_HEADER_ROW1 As Long = 16
Public Const MBO_UPPER_DETAIL_HEADER_ROW2 As Long = 17
Public Const MBO_UPPER_DETAIL_FIRST_ROW As Long = 18
Public Const MBO_UPPER_DETAIL_LAST_ROW As Long = 24
Public Const MBO_UPPER_TOTAL_ROW As Long = 25
Public Const MBO_UPPER_PERIOD_COL As Long = 7
Public Const MBO_UPPER_PERIOD_VALUE As String = "上期"
Public Const MBO_UPPER_CAREER_HEADER_COL As Long = 19
Public Const MBO_UPPER_CAREER_VALUE_COL As Long = 22
Public Const MBO_UPPER_EVALUATOR_COL As Long = 22
Public Const MBO_UPPER_EVALUATION_DATE_COL As Long = 24
Public Const MBO_UPPER_TOTAL_COL_N As Long = 14
Public Const MBO_UPPER_TOTAL_COL_U As Long = 21
Public Const MBO_UPPER_TOTAL_COL_Z As Long = 26
Public Const MBO_UPPER_RANK_COL_P As Long = 16
Public Const MBO_UPPER_RANK_COL_S As Long = 19
Public Const MBO_UPPER_RANK_COL_V As Long = 22
Public Const MBO_UPPER_RANK_COL_X As Long = 24
Public Const MBO_UPPER_OUTPUT_TEXT_COLUMN_WIDTH As Double = 50 '約350px

Public Const MBO_FOLDER_NAME As String = "MBO"
Public Const MBO_OPEN_UPDATE_LINKS_NEVER As Long = 0
Public Const SHEET_TAB_COLOR_LOG As Long = 15132390 'RGB(230,230,230)
Public Const MERGED_FONT_SIZE_DELTA As Double = 2

'=====================================================================
' 【9. 共通MsgBox関連定数】
'=====================================================================
Public Const MSG_SAVE_REQUIRED As String = "このブックを一度保存してから実行してください。"
Public Const MSG_FOLDER_MBO_NOT_FOUND As String = "MBO フォルダが見つかりません。"
Public Const MSG_FOLDER_PREV_MBO_NOT_FOUND As String = "前期MBO フォルダが見つかりません。"
Public Const MSG_LIST_DATA_ROW_NOT_FOUND As String = "一覧シートにデータ行がありません。"

Public Const MSG_DONE_CREATE_LIST As String = "一覧作成が完了しました。"
Public Const MSG_DONE_CREATE_UPPER_LIST As String = "上期一覧作成が完了しました。"
Public Const MSG_DONE_UPDATE_PREV As String = "前期MBO の追記が完了しました。"

Public Const MSG_LABEL_SHEET_NAME As String = "シート名: "
Public Const MSG_LABEL_LOG_SHEET As String = "ログシート: "

'=====================================================================
' 【10. 共通LOG関連定数】
'=====================================================================
Public Const LOG_FOLDER_SUMMARY As String = "(SUMMARY)"
Public Const LOG_FOLDER_INFO As String = "(INFO)"

Public Const LOG_STATUS_INFO001 As String = "INFO001"
Public Const LOG_STATUS_INFO002 As String = "INFO002"
Public Const LOG_STATUS_INFO003 As String = "INFO003"
Public Const LOG_STATUS_WORN001 As String = "WORN001"
Public Const LOG_STATUS_WORN002 As String = "WORN002"
Public Const LOG_STATUS_WORN003 As String = "WORN003"
Public Const LOG_STATUS_WORN004 As String = "WORN004"
Public Const LOG_STATUS_ERROR001 As String = "ERROR001"

Public Const LOG_MATCH_TYPE_NONE As String = ""
Public Const LOG_MATCH_TYPE_EMPLOYEE_NO As String = "EmployeeNo"
Public Const LOG_MATCH_TYPE_NAME As String = "Name"

Public Const LOG_MSG_ROW_OUTPUT_PREFIX As String = "一覧行 "
Public Const LOG_MSG_ROW_OUTPUT_SUFFIX As String = " に出力"
Public Const LOG_MSG_ROW_REFLECT_SUFFIX As String = " に反映"
Public Const LOG_MSG_NODATA_TARGET_ROWS As String = "対象行に出力データなし"
Public Const LOG_MSG_NODATA_UPPER As String = "上期データがありません"
Public Const LOG_MSG_SKIP_NON_TARGET As String = "対象ファイル条件に不一致"
Public Const LOG_MSG_SHEET_MISSING_SUFFIX As String = " シートが見つかりません"
Public Const LOG_MSG_OPEN_ERROR_PREFIX As String = "ブックを開けませんでした: "
Public Const LOG_MSG_NO_MATCH_EMP_NO_NAME As String = "社員No/氏名とも一致なし"
Public Const LOG_MSG_CREATED_NEW_LOG_SHEET As String = "既存LOGが無いため新規作成しました。"

Public Const LOG_MSG_SUM_FOLDERS_SCANNED As String = "走査フォルダ数: "
Public Const LOG_MSG_SUM_FOLDERS_TARGET As String = "対象フォルダ数: "
Public Const LOG_MSG_SUM_FILES_SCANNED As String = "走査ファイル数: "
Public Const LOG_MSG_SUM_FILES_TARGET As String = "対象ファイル数: "
Public Const LOG_MSG_SUM_FILES_OPENED As String = "オープン成功数: "
Public Const LOG_MSG_SUM_ROWS_OUTPUT As String = "出力行数: "
Public Const LOG_MSG_SUM_NODATA As String = "NoData件数: "
Public Const LOG_MSG_SUM_SHEET_MISSING As String = "元シートなし件数: "
Public Const LOG_MSG_SUM_OPEN_ERROR As String = "オープン失敗数: "
Public Const LOG_MSG_SUM_MATCH_EMP_NO As String = "社員No一致数: "
Public Const LOG_MSG_SUM_MATCH_NAME As String = "氏名一致数: "
Public Const LOG_MSG_SUM_NO_MATCH As String = "不一致数: "
Public Const LOG_MSG_SUM_EVAL_SHEET_MISSING As String = "評価シートなし: "

Public Const LOG_HEADER_TIME As String = "Time"
Public Const LOG_HEADER_FOLDER As String = "Folder"
Public Const LOG_HEADER_FILE As String = "File"
Public Const LOG_HEADER_STATUS As String = "Status"
Public Const LOG_HEADER_MATCH_TYPE As String = "MatchType"
Public Const LOG_HEADER_EMPLOYEE_NO As String = "EmployeeNo"
Public Const LOG_HEADER_NAME As String = "Name"
Public Const LOG_HEADER_MESSAGE As String = "Message"
Public Const LOG_HEADER_PATH As String = "Path"

Public Const MBO_UPPER_HEADER_SELF_RANK As String = "自己評価ランク"
Public Const MBO_UPPER_HEADER_BOSS_RANK As String = "上司評価ランク"
Public Const MBO_UPPER_HEADER_SELF_TOTAL As String = "自己評価合計"
Public Const MBO_UPPER_HEADER_BOSS_TOTAL As String = "上司評価合計"
Public Const MBO_UPPER_HEADER_TOTAL_SUFFIX As String = "合計"

'=====================================================================
' 【11. 達成基準シート関連の定数】
'=====================================================================
Public Const MBO_CRITERIA_SHEET_NAME As String = "達成基準"      ' 達成基準シート名
Public Const MBO_CRITERIA_FIRST_ROW As Long = 6                 ' 達成基準データ開始行
Public Const MBO_CRITERIA_LAST_ROW As Long = 12                 ' 達成基準データ終了行
Public Const MBO_CRITERIA_COL As Long = 5                       ' 達成基準データ列（E列=5）

'==========================
' メイン処理
'==========================
Public Sub CreateMboList()

    Dim workbookTarget As Workbook
    Dim worksheetSummary As Worksheet
    Dim fileSystem As Object
    Dim folderRoot As Object
    Dim folderPerson As Object
    Dim fileMbo As Object
    
    Dim pathMbo As String
    Dim sheetNameNew As String
    
    Dim workbookSource As Workbook
    Dim worksheetSource As Worksheet
    Dim worksheetCriteria As Worksheet
    Dim worksheetLog As Worksheet
    
    Dim valueEmployeeNo As Variant
    Dim valueEmployeeName As Variant
    Dim valueDepartment As Variant
    Dim valueJoinDate As Variant
    Dim valueServiceYears As Variant
    Dim valueTitle As Variant
    Dim valueGrade As Variant
    Dim valueCareer As Variant
    Dim valueCommittee As Variant
    Dim valueCriteria As Variant
    
    Dim rowSummary As Long
    Dim rowData As Long
    Dim rowCriteria As Long
    Dim rowLog As Long
    
    Dim fileIndex As Long
    Dim rowFileStart As Long
    Dim rowFileEnd As Long
    
    Dim lastRow As Long
    Dim rangeAll As Range
    Dim defaultRowHeight As Double
    Dim i As Long
    Dim leftCols As Variant
    Dim folderHasTargetFile As Boolean
    Dim worksheetCaller As Worksheet

    Dim countFoldersScanned As Long
    Dim countFoldersTarget As Long
    Dim countFilesScanned As Long
    Dim countFilesTarget As Long
    Dim countFilesOpened As Long
    Dim countRowsOutput As Long
    Dim countNoDataRows As Long
    Dim countSheetMissing As Long
    Dim countOpenError As Long
    
    Set workbookTarget = ThisWorkbook
    Set worksheetCaller = GetCallerWorksheet()
    
    If workbookTarget.Path = vbNullString Then
        MsgBox MSG_SAVE_REQUIRED, vbExclamation
        Exit Sub
    End If
    
    pathMbo = workbookTarget.Path & Application.PathSeparator & "MBO"
    If Dir$(pathMbo, vbDirectory) = vbNullString Then
        MsgBox MSG_FOLDER_MBO_NOT_FOUND & vbCrLf & pathMbo, vbExclamation
        Exit Sub
    End If
    
    '--- 一覧シートを新規作成
    sheetNameNew = MBO_INITIAL_SHEET_PREFIX & Format(Now, "yyyymmdd_hhnnss")
    Set worksheetSummary = workbookTarget.Worksheets.Add(After:=workbookTarget.Worksheets(workbookTarget.Worksheets.Count))
    On Error Resume Next
    worksheetSummary.Name = sheetNameNew
    On Error GoTo 0
    
    '--- 見出し行を設定（ヘッダ色設定含む）
    Call SetSummaryHeader(worksheetSummary)
    rowSummary = 2

    Set worksheetLog = CreateLogSheet(workbookTarget, worksheetSummary.Name)
    SetLogHeader worksheetLog
    ApplySheetTabColors worksheetCaller, worksheetSummary, worksheetLog
    rowLog = 2
    
    '--- FSO 初期化
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Set folderRoot = fileSystem.GetFolder(pathMbo)
    
    fileIndex = 0
    
    '--- MBO\配下の各フォルダ＆ファイルを走査
    For Each folderPerson In folderRoot.SubFolders
        countFoldersScanned = countFoldersScanned + 1
        folderHasTargetFile = False

        For Each fileMbo In folderPerson.Files
            countFilesScanned = countFilesScanned + 1

            '「MBOシート人事評価シート」ファイルのみ対象
            If IsTargetMboExcelFile(fileMbo.Path) Then
                countFilesTarget = countFilesTarget + 1
                folderHasTargetFile = True

                rowFileStart = rowSummary   'このファイルの開始行

                '--- 対象ブックを開く
                Set workbookSource = Nothing
                Set worksheetSource = Nothing
                Set worksheetCriteria = Nothing

                On Error Resume Next
                Set workbookSource = OpenWorkbookReadOnlyNoUpdate(fileMbo.Path)
                If Err.Number <> 0 Or workbookSource Is Nothing Then
                    countOpenError = countOpenError + 1
                    AppendLogRow worksheetLog, rowLog, folderPerson.Name, fileMbo.Name, LOG_STATUS_ERROR001, LOG_MATCH_TYPE_NONE, "", "", LOG_MSG_OPEN_ERROR_PREFIX & Err.Description
                    Err.Clear
                    On Error GoTo 0
                    GoTo NextMainFile
                End If
                On Error GoTo 0
                countFilesOpened = countFilesOpened + 1

                On Error Resume Next
                Set worksheetSource = workbookSource.Worksheets(MBO_SOURCE_SHEET_NAME)
                Set worksheetCriteria = workbookSource.Worksheets(MBO_CRITERIA_SHEET_NAME)
                On Error GoTo 0

                If Not worksheetSource Is Nothing Then

                    '----- 上部ヘッダ部を一度だけ読み取り
                    valueEmployeeNo = worksheetSource.Cells(MboGetRowEmployeeNo, MboGetColEmployeeNo).value
                    valueEmployeeName = NormalizeName(worksheetSource.Cells(MboGetRowEmployeeName, MboGetColEmployeeName).value)
                    valueDepartment = NormalizeName(worksheetSource.Cells(MboGetRowDepartment, MboGetColDepartment).value)
                    valueJoinDate = worksheetSource.Cells(MboGetRowJoinDate, MboGetColJoinDate).value
                    valueServiceYears = worksheetSource.Cells(MboGetRowServiceYears, MboGetColServiceYears).value
                    valueTitle = worksheetSource.Cells(MboGetRowTitle, MboGetColTitle).value
                    valueGrade = worksheetSource.Cells(MboGetRowGrade, MboGetColGrade).value
                    valueCareer = worksheetSource.Cells(MboGetRowCareer, MboGetColCareer).value
                    valueCommittee = worksheetSource.Cells(MboGetRowCommittee, MboGetColCommittee).value

                    '----- データ行(18～24行)をループ
                    For rowData = MBO_SOURCE_FIRST_DATA_ROW To MBO_SOURCE_LAST_DATA_ROW

                        '★ 18～24行目の B,C,E列が全て空ならスキップ
                        If IsTargetRowEmpty(worksheetSource, rowData) Then
                            GoTo NextDataRow
                        End If

                        '--- 達成基準シートとの行対応（18→6, 19→7, ...）
                        valueCriteria = vbNullString
                        If Not worksheetCriteria Is Nothing Then
                            rowCriteria = MBO_CRITERIA_FIRST_ROW + (rowData - MBO_SOURCE_FIRST_DATA_ROW)
                            If rowCriteria >= MBO_CRITERIA_FIRST_ROW And rowCriteria <= MBO_CRITERIA_LAST_ROW Then
                                valueCriteria = worksheetCriteria.Cells(rowCriteria, MBO_CRITERIA_COL).value
                            End If
                        End If

                        With worksheetSummary
                            .Cells(rowSummary, MboColEmployeeNo).value = valueEmployeeNo
                            .Cells(rowSummary, MboColEmployeeName).value = valueEmployeeName
                            .Cells(rowSummary, MboColDepartment).value = valueDepartment
                            .Cells(rowSummary, MboColJoinDate).value = valueJoinDate
                            .Cells(rowSummary, MboColServiceYears).value = valueServiceYears
                            .Cells(rowSummary, MboColTitle).value = valueTitle
                            .Cells(rowSummary, MboColGrade).value = valueGrade
                            .Cells(rowSummary, MboColCareer).value = valueCareer
                            .Cells(rowSummary, MboColCommittee).value = valueCommittee

                            .Cells(rowSummary, MboColKbn).value = worksheetSource.Cells(rowData, MboGetColKbn).value
                            .Cells(rowSummary, MboColPortableSkill).value = worksheetSource.Cells(rowData, MboGetColPortableSkill).value
                            .Cells(rowSummary, MboColRequiredAction).value = worksheetSource.Cells(rowData, MboGetColRequiredAction).value
                            .Cells(rowSummary, MboColGoalNo).value = worksheetSource.Cells(rowData, MboGetColGoalNo).value
                            .Cells(rowSummary, MboColGoalA).value = worksheetSource.Cells(rowData, MboGetColGoalA).value
                            .Cells(rowSummary, MboColCriteria).value = valueCriteria
                            .Cells(rowSummary, MboColMeasureB).value = worksheetSource.Cells(rowData, MboGetColMeasureB).value
                            .Cells(rowSummary, MboColParentGoalNo).value = worksheetSource.Cells(rowData, MboGetColParentGoalNo).value
                            .Cells(rowSummary, MboColTerm).value = worksheetSource.Cells(rowData, MboGetColTerm).value
                            .Cells(rowSummary, MboColDateStart).value = worksheetSource.Cells(rowData, MboGetColDateStart).value
                            .Cells(rowSummary, MboColDateEnd).value = worksheetSource.Cells(rowData, MboGetColDateEnd).value
                            .Cells(rowSummary, MboColSelfWeight).value = worksheetSource.Cells(rowData, MboGetColSelfWeight).value
                            .Cells(rowSummary, MboColSelfDifficulty).value = worksheetSource.Cells(rowData, MboGetColSelfDifficulty).value
                            .Cells(rowSummary, MboColBossWeight).value = worksheetSource.Cells(rowData, MboGetColBossWeight).value
                            .Cells(rowSummary, MboColBossDifficulty).value = worksheetSource.Cells(rowData, MboGetColBossDifficulty).value
                            .Cells(rowSummary, MboColGoalScore).value = worksheetSource.Cells(rowData, MboGetColGoalScore).value

                            '==============================
                            ' 数値書式の指定
                            '==============================

                            '--- 本人／上司ウェイト：パーセンテージ表記
                            .Columns(MboColSelfWeight).NumberFormat = "0%"
                            .Columns(MboColBossWeight).NumberFormat = "0%"

                            '--- 本人難易度：0.0 表記（1 → 1.0）
                            .Columns(MboColSelfDifficulty).NumberFormat = "0.0"
                            .Columns(MboColBossDifficulty).NumberFormat = "0.0"

                            Dim lastSummaryRow As Long
                            lastSummaryRow = rowSummary - 1   '一覧シートの最終データ行

                            Dim rngSelfDifficulty As Range
                            Dim rngBossDifficulty As Range

                            Set rngSelfDifficulty = .Range( _
                                .Cells(2, MboColSelfDifficulty), _
                                .Cells(lastSummaryRow, MboColSelfDifficulty))

                            Set rngBossDifficulty = .Range( _
                                .Cells(2, MboColBossDifficulty), _
                                .Cells(lastSummaryRow, MboColBossDifficulty))

                            '========================
                            ' 既存の条件付き書式をクリア
                            '========================
                            rngSelfDifficulty.FormatConditions.Delete
                            rngBossDifficulty.FormatConditions.Delete

                            '========================
                            ' 0.9 以下（空白除外）→ 薄いグレー
                            '========================
                            With rngSelfDifficulty.FormatConditions.Add( _
                                    Type:=xlExpression, _
                                    Formula1:="=AND(NOT(ISBLANK(" & rngSelfDifficulty.Cells(1, 1).Address(False, False) & ")), " & _
                                                rngSelfDifficulty.Cells(1, 1).Address(False, False) & "<=0.9)")
                                .Interior.Color = RGB(217, 217, 217)
                            End With

                            With rngBossDifficulty.FormatConditions.Add( _
                                    Type:=xlExpression, _
                                    Formula1:="=AND(NOT(ISBLANK(" & rngBossDifficulty.Cells(1, 1).Address(False, False) & ")), " & _
                                                rngBossDifficulty.Cells(1, 1).Address(False, False) & "<=0.9)")
                                .Interior.Color = RGB(217, 217, 217)
                            End With

                            '========================
                            ' 1.1 → 薄い黄色
                            '========================
                            With rngSelfDifficulty.FormatConditions.Add( _
                                    Type:=xlCellValue, Operator:=xlEqual, Formula1:="1.1")
                                .Interior.Color = RGB(255, 242, 204)
                            End With

                            With rngBossDifficulty.FormatConditions.Add( _
                                    Type:=xlCellValue, Operator:=xlEqual, Formula1:="1.1")
                                .Interior.Color = RGB(255, 242, 204)
                            End With

                            '========================
                            ' 1.2 → 薄い赤
                            '========================
                            With rngSelfDifficulty.FormatConditions.Add( _
                                    Type:=xlCellValue, Operator:=xlEqual, Formula1:="1.2")
                                .Interior.Color = RGB(248, 203, 173)
                            End With

                            With rngBossDifficulty.FormatConditions.Add( _
                                    Type:=xlCellValue, Operator:=xlEqual, Formula1:="1.2")
                                .Interior.Color = RGB(248, 203, 173)
                            End With

                        End With

                        rowSummary = rowSummary + 1
NextDataRow:
                    Next rowData

                    'このファイル分の行終端
                    rowFileEnd = rowSummary - 1

                    '--- ファイルごとの縞々（白／薄い青）
                    If rowFileEnd >= rowFileStart Then
                        fileIndex = fileIndex + 1
                        countRowsOutput = countRowsOutput + (rowFileEnd - rowFileStart + 1)
                        Dim rangeBlock As Range
                        Set rangeBlock = worksheetSummary.Range( _
                            worksheetSummary.Cells(rowFileStart, MboColEmployeeNo), _
                            worksheetSummary.Cells(rowFileEnd, MboColGoalScore))

                        If fileIndex Mod 2 = 0 Then
                            rangeBlock.Interior.Color = RGB(221, 235, 247)   '薄い青
                        Else
                            rangeBlock.Interior.Color = vbWhite              '白
                        End If

                        AppendLogRow worksheetLog, rowLog, folderPerson.Name, fileMbo.Name, LOG_STATUS_INFO001, LOG_MATCH_TYPE_NONE, NormalizeEmployeeNo(valueEmployeeNo), valueEmployeeName, LOG_MSG_ROW_OUTPUT_PREFIX & CStr(rowFileStart) & "-" & CStr(rowFileEnd) & LOG_MSG_ROW_OUTPUT_SUFFIX
                    Else
                        countNoDataRows = countNoDataRows + 1
                        AppendLogRow worksheetLog, rowLog, folderPerson.Name, fileMbo.Name, LOG_STATUS_WORN001, LOG_MATCH_TYPE_NONE, NormalizeEmployeeNo(valueEmployeeNo), valueEmployeeName, LOG_MSG_NODATA_TARGET_ROWS
                    End If
                Else
                    countSheetMissing = countSheetMissing + 1
                    AppendLogRow worksheetLog, rowLog, folderPerson.Name, fileMbo.Name, LOG_STATUS_WORN002, LOG_MATCH_TYPE_NONE, "", "", MBO_SOURCE_SHEET_NAME & LOG_MSG_SHEET_MISSING_SUFFIX
                End If

NextMainFile:
                If Not workbookSource Is Nothing Then
                    workbookSource.Close SaveChanges:=False
                End If
                Set workbookSource = Nothing
                Set worksheetSource = Nothing
                Set worksheetCriteria = Nothing
            Else
                AppendLogRow worksheetLog, rowLog, folderPerson.Name, fileMbo.Name, LOG_STATUS_WORN003, LOG_MATCH_TYPE_NONE, "", "", LOG_MSG_SKIP_NON_TARGET
            End If
        Next fileMbo

        If folderHasTargetFile Then
            countFoldersTarget = countFoldersTarget + 1
        End If
    Next folderPerson
    
    '==============================
    ' レイアウト系（フィルタ・幅・寄せ）
    '==============================
    lastRow = rowSummary - 1
    If lastRow < 1 Then lastRow = 1
    
    With worksheetSummary
        
        '--- 全体範囲
        Set rangeAll = .Range(.Cells(1, MboColEmployeeNo), .Cells(lastRow, MboColGoalScore))
        
        '--- フィルタ設定
        rangeAll.AutoFilter
        
        '--- 列幅は AutoFit
        rangeAll.EntireColumn.AutoFit
        
        ' 長文列を固定幅
        .Columns(MboColRequiredAction).ColumnWidth = 50
        .Columns(MboColGoalA).ColumnWidth = 100
        .Columns(MboColCriteria).ColumnWidth = 100
        .Columns(MboColMeasureB).ColumnWidth = 100
        
        '--- 行の高さを 1行目と同じにそろえる
        defaultRowHeight = .Rows(1).RowHeight
        .Rows("1:" & lastRow).RowHeight = defaultRowHeight
        
        '--- 文字寄せ（数字だけ／日付以外は左寄せ）
        '    数値/日付扱い：
        '       入社日, 期日開始, 期日終了,
        '       本人ウェイト, 本人難易度,
        '       上司ウェイト, 上司難易度,
        '       目標点数
        leftCols = Array( _
            MboColEmployeeNo, _
            MboColEmployeeName, _
            MboColDepartment, _
            MboColServiceYears, _
            MboColTitle, _
            MboColGrade, _
            MboColCareer, _
            MboColCommittee, _
            MboColKbn, _
            MboColPortableSkill, _
            MboColRequiredAction, _
            MboColGoalNo, _
            MboColGoalA, _
            MboColCriteria, _
            MboColMeasureB, _
            MboColParentGoalNo, _
            MboColTerm)
        
        For i = LBound(leftCols) To UBound(leftCols)
            .Columns(leftCols(i)).HorizontalAlignment = xlLeft
        Next i
        
    End With

        AppendLogRow worksheetLog, rowLog, LOG_FOLDER_SUMMARY, "", LOG_STATUS_INFO002, LOG_MATCH_TYPE_NONE, "", "", LOG_MSG_SUM_FOLDERS_SCANNED & CStr(countFoldersScanned)
        AppendLogRow worksheetLog, rowLog, LOG_FOLDER_SUMMARY, "", LOG_STATUS_INFO002, LOG_MATCH_TYPE_NONE, "", "", LOG_MSG_SUM_FOLDERS_TARGET & CStr(countFoldersTarget)
        AppendLogRow worksheetLog, rowLog, LOG_FOLDER_SUMMARY, "", LOG_STATUS_INFO002, LOG_MATCH_TYPE_NONE, "", "", LOG_MSG_SUM_FILES_SCANNED & CStr(countFilesScanned)
        AppendLogRow worksheetLog, rowLog, LOG_FOLDER_SUMMARY, "", LOG_STATUS_INFO002, LOG_MATCH_TYPE_NONE, "", "", LOG_MSG_SUM_FILES_TARGET & CStr(countFilesTarget)
        AppendLogRow worksheetLog, rowLog, LOG_FOLDER_SUMMARY, "", LOG_STATUS_INFO002, LOG_MATCH_TYPE_NONE, "", "", LOG_MSG_SUM_FILES_OPENED & CStr(countFilesOpened)
        AppendLogRow worksheetLog, rowLog, LOG_FOLDER_SUMMARY, "", LOG_STATUS_INFO002, LOG_MATCH_TYPE_NONE, "", "", LOG_MSG_SUM_ROWS_OUTPUT & CStr(countRowsOutput)
        AppendLogRow worksheetLog, rowLog, LOG_FOLDER_SUMMARY, "", LOG_STATUS_INFO002, LOG_MATCH_TYPE_NONE, "", "", LOG_MSG_SUM_NODATA & CStr(countNoDataRows)
        AppendLogRow worksheetLog, rowLog, LOG_FOLDER_SUMMARY, "", LOG_STATUS_INFO002, LOG_MATCH_TYPE_NONE, "", "", LOG_MSG_SUM_SHEET_MISSING & CStr(countSheetMissing)
        AppendLogRow worksheetLog, rowLog, LOG_FOLDER_SUMMARY, "", LOG_STATUS_INFO002, LOG_MATCH_TYPE_NONE, "", "", LOG_MSG_SUM_OPEN_ERROR & CStr(countOpenError)

        worksheetLog.Columns("A:I").EntireColumn.AutoFit
    
            MsgBox MSG_DONE_CREATE_LIST & vbCrLf & _
                MSG_LABEL_SHEET_NAME & worksheetSummary.Name & vbCrLf & _
                MSG_LABEL_LOG_SHEET & worksheetLog.Name & vbCrLf & _
            LOG_MSG_SUM_FOLDERS_SCANNED & countFoldersScanned & vbCrLf & _
            LOG_MSG_SUM_FILES_TARGET & countFilesTarget, vbInformation
    
End Sub

'==========================================================
' 指定パスのファイルが
'  - ファイル名に「MBOシート_人事評価シート」を含む（部分一致）
'  - Excel拡張子（xlsx/xlsm/xlsb/xls）
'  - かつ「隠し/システム」ではない
'  - かつ Office一時/ロック（~$）ではない
' なら True
'==========================================================
Private Function IsTargetMboExcelFile(ByVal filePath As String) As Boolean

    Dim fileNameOnly As String
    Dim lowerName As String
    Dim ext As String
    Dim dotPos As Long
    Dim attr As Long

    On Error GoTo ErrHandler

    'ファイル名だけ取り出し
    fileNameOnly = Mid$(filePath, InStrRev(filePath, Application.PathSeparator) + 1)
    lowerName = LCase$(fileNameOnly)

    'Officeロック/一時ファイルは除外（最優先）
    If Left$(fileNameOnly, 2) = "~$" Then
        IsTargetMboExcelFile = False
        Exit Function
    End If

    '隠し/システム属性は除外
    On Error Resume Next
    attr = GetAttr(filePath)
    If Err.Number = 0 Then
        If (attr And vbHidden) <> 0 Or (attr And vbSystem) <> 0 Then
            IsTargetMboExcelFile = False
            Exit Function
        End If
    End If
    On Error GoTo 0

    '部分一致（大文字小文字無視）
    If InStr(1, lowerName, LCase$(MBO_TARGET_FILE_KEYWORD), vbTextCompare) = 0 Then
        IsTargetMboExcelFile = False
        Exit Function
    End If

    '拡張子チェック（ドットなしも安全に弾く）
    dotPos = InStrRev(lowerName, ".")
    If dotPos = 0 Then
        IsTargetMboExcelFile = False
        Exit Function
    End If

    ext = Mid$(lowerName, dotPos + 1)
    Dim arrExts As Variant
    Dim i As Integer
    IsTargetMboExcelFile = False
    arrExts = Split(MBO_TARGET_FILE_EXTS, ",")
    For i = LBound(arrExts) To UBound(arrExts)
        If ext = Trim$(arrExts(i)) Then
            IsTargetMboExcelFile = True
            Exit For
        End If
    Next i

    Exit Function

ErrHandler:
    '属性取得できない等（権限/パス不正）は対象外扱い
    IsTargetMboExcelFile = False
End Function

Private Function OpenWorkbookReadOnlyNoUpdate(ByVal filePath As String) As Workbook

    Dim previousAskToUpdateLinks As Boolean
    Dim previousDisplayAlerts As Boolean

    previousAskToUpdateLinks = Application.AskToUpdateLinks
    previousDisplayAlerts = Application.DisplayAlerts

    On Error GoTo OpenError

    Application.AskToUpdateLinks = False
    Application.DisplayAlerts = False

    Set OpenWorkbookReadOnlyNoUpdate = Workbooks.Open( _
        Filename:=filePath, _
        UpdateLinks:=MBO_OPEN_UPDATE_LINKS_NEVER, _
        ReadOnly:=True, _
        IgnoreReadOnlyRecommended:=True, _
        Notify:=False)

RestoreState:
    Application.DisplayAlerts = previousDisplayAlerts
    Application.AskToUpdateLinks = previousAskToUpdateLinks
    Exit Function

OpenError:
    Set OpenWorkbookReadOnlyNoUpdate = Nothing
    Resume RestoreState

End Function

'==========================
' 見出し行の設定＋ヘッダ装飾
'==========================
Private Sub SetSummaryHeader(ByVal worksheetSummary As Worksheet)

    With worksheetSummary
        .Cells(1, MboColEmployeeNo).value = MBO_HEADER_EMPLOYEE_NO
        .Cells(1, MboColEmployeeName).value = MBO_HEADER_EMPLOYEE_NAME
        .Cells(1, MboColDepartment).value = MBO_HEADER_DEPARTMENT
        .Cells(1, MboColJoinDate).value = MBO_HEADER_JOIN_DATE
        .Cells(1, MboColServiceYears).value = MBO_HEADER_SERVICE_YEARS
        .Cells(1, MboColTitle).value = MBO_HEADER_TITLE
        .Cells(1, MboColGrade).value = MBO_HEADER_GRADE
        .Cells(1, MboColCareer).value = MBO_HEADER_CAREER
        .Cells(1, MboColCommittee).value = MBO_HEADER_COMMITTEE
        .Cells(1, MboColKbn).value = MBO_HEADER_KBN
        .Cells(1, MboColPortableSkill).value = MBO_HEADER_PORTABLE_SKILL
        .Cells(1, MboColRequiredAction).value = MBO_HEADER_REQUIRED_ACTION
        .Cells(1, MboColGoalNo).value = MBO_HEADER_GOAL_NO
        .Cells(1, MboColGoalA).value = MBO_HEADER_GOAL_A
        .Cells(1, MboColCriteria).value = MBO_HEADER_CRITERIA
        .Cells(1, MboColMeasureB).value = MBO_HEADER_MEASURE_B
        .Cells(1, MboColParentGoalNo).value = MBO_HEADER_PARENT_GOAL_NO
        .Cells(1, MboColTerm).value = MBO_HEADER_TERM
        .Cells(1, MboColDateStart).value = MBO_HEADER_DATE_START
        .Cells(1, MboColDateEnd).value = MBO_HEADER_DATE_END
        .Cells(1, MboColSelfWeight).value = MBO_HEADER_SELF_WEIGHT
        .Cells(1, MboColSelfDifficulty).value = MBO_HEADER_SELF_DIFFICULTY
        .Cells(1, MboColBossWeight).value = MBO_HEADER_BOSS_WEIGHT
        .Cells(1, MboColBossDifficulty).value = MBO_HEADER_BOSS_DIFFICULTY
        .Cells(1, MboColGoalScore).value = MBO_HEADER_GOAL_SCORE
        
        '--- ヘッダ装飾（濃い青＋白文字）
        With .Range(.Cells(1, MboColEmployeeNo), .Cells(1, MboColGoalScore))
            .Interior.Color = RGB(0, 102, 204) '濃いめの青
            .Font.Color = vbWhite
            .Font.Bold = True
        End With
    End With

End Sub

'==========================
' 行スキップ判定
'   目標設定シートの「ポータブルスキル・求められる行動・目標設定A」が全て空なら True
'==========================
Private Function IsTargetRowEmpty(ByVal worksheetSource As Worksheet, _
                                  ByVal rowIndex As Long) As Boolean
    Dim vPortable As Variant
    Dim vRequired As Variant
    Dim vGoalA As Variant
    
    'Enum を使用して取得
    vPortable = worksheetSource.Cells(rowIndex, MboGetColPortableSkill).value
    vRequired = worksheetSource.Cells(rowIndex, MboGetColRequiredAction).value
    vGoalA = worksheetSource.Cells(rowIndex, MboGetColGoalA).value
    
    If Trim(CStr(vPortable)) = "" And _
       Trim(CStr(vRequired)) = "" And _
       Trim(CStr(vGoalA)) = "" Then
        IsTargetRowEmpty = True
    Else
        IsTargetRowEmpty = False
    End If
End Function

'==========================
' 上期評価一覧作成
'==========================
Public Sub CreateMboUpperHalfList()

    Dim workbookTarget As Workbook
    Dim worksheetSummary As Worksheet
    Dim worksheetLog As Worksheet
    Dim fileSystem As Object
    Dim folderRoot As Object
    Dim folderPerson As Object
    Dim fileMbo As Object

    Dim workbookSource As Workbook
    Dim worksheetSource As Worksheet

    Dim pathMbo As String
    Dim sheetNameNew As String

    Dim rowSummary As Long
    Dim rowData As Long
    Dim rowLog As Long
    Dim rowFileStart As Long
    Dim rowFileEnd As Long
    Dim fileIndex As Long
    Dim lastRow As Long
    Dim lastCol As Long
    Dim rangeAll As Range
    Dim defaultRowHeight As Double
    Dim valueEmployeeNo As String
    Dim valueEmployeeName As String
    Dim hasHeader As Boolean
    Dim folderHasTargetFile As Boolean
    Dim worksheetCaller As Worksheet

    Dim countFoldersScanned As Long
    Dim countFoldersTarget As Long
    Dim countFilesScanned As Long
    Dim countFilesTarget As Long
    Dim countFilesOpened As Long
    Dim countRowsOutput As Long
    Dim countNoDataRows As Long
    Dim countSheetMissing As Long
    Dim countOpenError As Long

    Set workbookTarget = ThisWorkbook
    Set worksheetCaller = GetCallerWorksheet()

    If workbookTarget.Path = vbNullString Then
        MsgBox MSG_SAVE_REQUIRED, vbExclamation
        Exit Sub
    End If

    pathMbo = workbookTarget.Path & Application.PathSeparator & MBO_FOLDER_NAME
    If Dir$(pathMbo, vbDirectory) = vbNullString Then
        MsgBox MSG_FOLDER_MBO_NOT_FOUND & vbCrLf & pathMbo, vbExclamation
        Exit Sub
    End If

    sheetNameNew = MBO_UPPER_SHEET_PREFIX & Format(Now, "yyyymmdd_hhnnss")
    Set worksheetSummary = workbookTarget.Worksheets.Add(After:=workbookTarget.Worksheets(workbookTarget.Worksheets.Count))
    On Error Resume Next
    worksheetSummary.Name = sheetNameNew
    On Error GoTo 0

    rowSummary = 2
    hasHeader = False
    lastCol = 1

    Set worksheetLog = CreateLogSheet(workbookTarget, worksheetSummary.Name)
    SetLogHeader worksheetLog
    ApplySheetTabColors worksheetCaller, worksheetSummary, worksheetLog
    rowLog = 2

    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Set folderRoot = fileSystem.GetFolder(pathMbo)

    For Each folderPerson In folderRoot.SubFolders
        countFoldersScanned = countFoldersScanned + 1
        folderHasTargetFile = False

        For Each fileMbo In folderPerson.Files
            countFilesScanned = countFilesScanned + 1

            If IsTargetMboExcelFile(fileMbo.Path) Then
                countFilesTarget = countFilesTarget + 1
                folderHasTargetFile = True

                Set workbookSource = Nothing
                Set worksheetSource = Nothing

                On Error Resume Next
                Set workbookSource = OpenWorkbookReadOnlyNoUpdate(fileMbo.Path)
                If Err.Number <> 0 Or workbookSource Is Nothing Then
                    countOpenError = countOpenError + 1
                    AppendLogRow worksheetLog, rowLog, folderPerson.Name, fileMbo.Name, LOG_STATUS_ERROR001, LOG_MATCH_TYPE_NONE, "", "", LOG_MSG_OPEN_ERROR_PREFIX & Err.Description
                    Err.Clear
                    On Error GoTo 0
                    GoTo NextUpperMainFile
                End If
                On Error GoTo 0
                countFilesOpened = countFilesOpened + 1

                On Error Resume Next
                Set worksheetSource = workbookSource.Worksheets(MBO_UPPER_SOURCE_SHEET_NAME)
                On Error GoTo 0

                If Not worksheetSource Is Nothing Then
                    If Not hasHeader Then
                        lastCol = SetUpperHalfSummaryHeader(worksheetSummary, worksheetSource)
                        hasHeader = True
                    End If

                    valueEmployeeNo = NormalizeEmployeeNo(worksheetSource.Cells(MBO_UPPER_HEADER_VALUE_ROW, 2).value)
                    valueEmployeeName = NormalizeName(CStr(worksheetSource.Cells(MBO_UPPER_HEADER_VALUE_ROW, 5).value))

                    rowFileStart = rowSummary

                    For rowData = MBO_UPPER_DETAIL_FIRST_ROW To MBO_UPPER_DETAIL_LAST_ROW
                        If IsUpperHalfPeriodRow(worksheetSource, rowData) Then
                            WriteUpperHalfSummaryRow worksheetSummary, rowSummary, worksheetSource, rowData
                            rowSummary = rowSummary + 1
                        End If
                    Next rowData

                    rowFileEnd = rowSummary - 1

                    If rowFileEnd >= rowFileStart Then
                        fileIndex = fileIndex + 1
                        countRowsOutput = countRowsOutput + (rowFileEnd - rowFileStart + 1)

                        Dim rangeBlock As Range
                        Set rangeBlock = worksheetSummary.Range( _
                            worksheetSummary.Cells(rowFileStart, 1), _
                            worksheetSummary.Cells(rowFileEnd, lastCol))

                        If fileIndex Mod 2 = 0 Then
                            rangeBlock.Interior.Color = RGB(221, 235, 247)
                        Else
                            rangeBlock.Interior.Color = vbWhite
                        End If

                        MergeUpperRepeatedSummaryCells worksheetSummary, rowFileStart, rowFileEnd, lastCol

                        AppendLogRow worksheetLog, rowLog, folderPerson.Name, fileMbo.Name, LOG_STATUS_INFO001, LOG_MATCH_TYPE_NONE, valueEmployeeNo, valueEmployeeName, LOG_MSG_ROW_OUTPUT_PREFIX & CStr(rowFileStart) & "-" & CStr(rowFileEnd) & LOG_MSG_ROW_OUTPUT_SUFFIX
                    Else
                        countNoDataRows = countNoDataRows + 1
                        AppendLogRow worksheetLog, rowLog, folderPerson.Name, fileMbo.Name, LOG_STATUS_WORN001, LOG_MATCH_TYPE_NONE, valueEmployeeNo, valueEmployeeName, LOG_MSG_NODATA_UPPER
                    End If
                Else
                    countSheetMissing = countSheetMissing + 1
                    AppendLogRow worksheetLog, rowLog, folderPerson.Name, fileMbo.Name, LOG_STATUS_WORN002, LOG_MATCH_TYPE_NONE, "", "", MBO_UPPER_SOURCE_SHEET_NAME & LOG_MSG_SHEET_MISSING_SUFFIX
                End If

NextUpperMainFile:
                If Not workbookSource Is Nothing Then
                    workbookSource.Close SaveChanges:=False
                End If
                Set workbookSource = Nothing
                Set worksheetSource = Nothing
            Else
                AppendLogRow worksheetLog, rowLog, folderPerson.Name, fileMbo.Name, LOG_STATUS_WORN003, LOG_MATCH_TYPE_NONE, "", "", LOG_MSG_SKIP_NON_TARGET
            End If
        Next fileMbo

        If folderHasTargetFile Then
            countFoldersTarget = countFoldersTarget + 1
        End If
    Next folderPerson

    If hasHeader Then
        lastRow = rowSummary - 1
        If lastRow < 1 Then lastRow = 1

        With worksheetSummary
            Set rangeAll = .Range(.Cells(1, 1), .Cells(lastRow, lastCol))
            rangeAll.AutoFilter
            rangeAll.EntireColumn.AutoFit

            defaultRowHeight = .Rows(1).RowHeight
            .Rows("1:" & lastRow).RowHeight = defaultRowHeight

            .Range(.Cells(2, 1), .Cells(lastRow, lastCol)).HorizontalAlignment = xlLeft
        End With

        ApplyUpperHalfNarrowColumnWidths worksheetSummary, lastCol

        ApplyUpperHalfConditionalFormats worksheetSummary, lastRow, lastCol
    End If

    AppendLogRow worksheetLog, rowLog, LOG_FOLDER_SUMMARY, "", LOG_STATUS_INFO002, LOG_MATCH_TYPE_NONE, "", "", LOG_MSG_SUM_FOLDERS_SCANNED & CStr(countFoldersScanned)
    AppendLogRow worksheetLog, rowLog, LOG_FOLDER_SUMMARY, "", LOG_STATUS_INFO002, LOG_MATCH_TYPE_NONE, "", "", LOG_MSG_SUM_FOLDERS_TARGET & CStr(countFoldersTarget)
    AppendLogRow worksheetLog, rowLog, LOG_FOLDER_SUMMARY, "", LOG_STATUS_INFO002, LOG_MATCH_TYPE_NONE, "", "", LOG_MSG_SUM_FILES_SCANNED & CStr(countFilesScanned)
    AppendLogRow worksheetLog, rowLog, LOG_FOLDER_SUMMARY, "", LOG_STATUS_INFO002, LOG_MATCH_TYPE_NONE, "", "", LOG_MSG_SUM_FILES_TARGET & CStr(countFilesTarget)
    AppendLogRow worksheetLog, rowLog, LOG_FOLDER_SUMMARY, "", LOG_STATUS_INFO002, LOG_MATCH_TYPE_NONE, "", "", LOG_MSG_SUM_FILES_OPENED & CStr(countFilesOpened)
    AppendLogRow worksheetLog, rowLog, LOG_FOLDER_SUMMARY, "", LOG_STATUS_INFO002, LOG_MATCH_TYPE_NONE, "", "", LOG_MSG_SUM_ROWS_OUTPUT & CStr(countRowsOutput)
    AppendLogRow worksheetLog, rowLog, LOG_FOLDER_SUMMARY, "", LOG_STATUS_INFO002, LOG_MATCH_TYPE_NONE, "", "", LOG_MSG_SUM_NODATA & CStr(countNoDataRows)
    AppendLogRow worksheetLog, rowLog, LOG_FOLDER_SUMMARY, "", LOG_STATUS_INFO002, LOG_MATCH_TYPE_NONE, "", "", LOG_MSG_SUM_SHEET_MISSING & CStr(countSheetMissing)
    AppendLogRow worksheetLog, rowLog, LOG_FOLDER_SUMMARY, "", LOG_STATUS_INFO002, LOG_MATCH_TYPE_NONE, "", "", LOG_MSG_SUM_OPEN_ERROR & CStr(countOpenError)

    worksheetLog.Columns("A:I").EntireColumn.AutoFit

        MsgBox MSG_DONE_CREATE_UPPER_LIST & vbCrLf & _
            MSG_LABEL_SHEET_NAME & worksheetSummary.Name & vbCrLf & _
            MSG_LABEL_LOG_SHEET & worksheetLog.Name & vbCrLf & _
            LOG_MSG_SUM_FOLDERS_SCANNED & countFoldersScanned & vbCrLf & _
            LOG_MSG_SUM_FILES_TARGET & countFilesTarget, vbInformation

End Sub

Private Function SetUpperHalfSummaryHeader(ByVal worksheetSummary As Worksheet, ByVal worksheetSource As Worksheet) As Long

    Dim outCol As Long
    Dim srcCol As Variant
    Dim headerText As String
    Dim topCols As Variant
    Dim detailCols As Variant
    Dim scoreCols As Variant

    outCol = 1

    topCols = GetUpperTopSourceColumns()
    detailCols = GetUpperDetailSourceColumns()
    scoreCols = GetUpperScoreSourceColumns()

    For Each srcCol In topCols
        worksheetSummary.Cells(1, outCol).value = GetUpperHeaderText(worksheetSource, MBO_UPPER_HEADER_ROW, srcCol)
        outCol = outCol + 1
    Next

    worksheetSummary.Cells(1, outCol).value = GetUpperHeaderText(worksheetSource, MBO_UPPER_HEADER_ROW, MBO_UPPER_EVALUATOR_COL)
    outCol = outCol + 1
    worksheetSummary.Cells(1, outCol).value = GetUpperHeaderText(worksheetSource, MBO_UPPER_HEADER_ROW, MBO_UPPER_EVALUATION_DATE_COL)
    outCol = outCol + 1

    worksheetSummary.Cells(1, outCol).value = GetUpperHeaderText(worksheetSource, MBO_UPPER_CAREER_ROW, MBO_UPPER_CAREER_HEADER_COL)
    outCol = outCol + 1
    worksheetSummary.Cells(1, outCol).value = GetUpperHeaderText(worksheetSource, MBO_UPPER_COMMITTEE_ROW, MBO_UPPER_CAREER_HEADER_COL)
    outCol = outCol + 1

    For Each srcCol In detailCols
        worksheetSummary.Cells(1, outCol).value = GetUpperMergedHeaderText(worksheetSource, srcCol)
        outCol = outCol + 1
    Next

    For Each srcCol In scoreCols
        worksheetSummary.Cells(1, outCol).value = GetUpperMergedHeaderText(worksheetSource, srcCol)
        outCol = outCol + 1

        If srcCol = MBO_UPPER_TOTAL_COL_U Then
            worksheetSummary.Cells(1, outCol).value = MBO_UPPER_HEADER_SELF_TOTAL
            outCol = outCol + 1
            worksheetSummary.Cells(1, outCol).value = MBO_UPPER_HEADER_SELF_RANK
            outCol = outCol + 1
        End If

        If srcCol = MBO_UPPER_TOTAL_COL_Z Then
            worksheetSummary.Cells(1, outCol).value = MBO_UPPER_HEADER_BOSS_TOTAL
            outCol = outCol + 1
            worksheetSummary.Cells(1, outCol).value = MBO_UPPER_HEADER_BOSS_RANK
            outCol = outCol + 1
        End If
    Next

    SetUpperHalfSummaryHeader = outCol - 1

    With worksheetSummary.Range(worksheetSummary.Cells(1, 1), worksheetSummary.Cells(1, SetUpperHalfSummaryHeader))
        .Interior.Color = RGB(0, 102, 204)
        .Font.Color = vbWhite
        .Font.Bold = True
    End With

End Function

Private Sub WriteUpperHalfSummaryRow(ByVal worksheetSummary As Worksheet, _
                                     ByVal outRow As Long, _
                                     ByVal worksheetSource As Worksheet, _
                                     ByVal sourceRow As Long)

    Dim outCol As Long
    Dim srcCol As Variant
    Dim topCols As Variant
    Dim detailCols As Variant
    Dim scoreCols As Variant

    outCol = 1

    topCols = GetUpperTopSourceColumns()
    detailCols = GetUpperDetailSourceColumns()
    scoreCols = GetUpperScoreSourceColumns()

    For Each srcCol In topCols
        worksheetSummary.Cells(outRow, outCol).value = worksheetSource.Cells(MBO_UPPER_HEADER_VALUE_ROW, srcCol).value
        outCol = outCol + 1
    Next

    worksheetSummary.Cells(outRow, outCol).value = worksheetSource.Cells(MBO_UPPER_HEADER_VALUE_ROW, MBO_UPPER_EVALUATOR_COL).value
    outCol = outCol + 1
    worksheetSummary.Cells(outRow, outCol).value = worksheetSource.Cells(MBO_UPPER_HEADER_VALUE_ROW, MBO_UPPER_EVALUATION_DATE_COL).value
    outCol = outCol + 1

    worksheetSummary.Cells(outRow, outCol).value = worksheetSource.Cells(MBO_UPPER_CAREER_ROW, MBO_UPPER_CAREER_VALUE_COL).value
    outCol = outCol + 1
    worksheetSummary.Cells(outRow, outCol).value = worksheetSource.Cells(MBO_UPPER_COMMITTEE_ROW, MBO_UPPER_CAREER_VALUE_COL).value
    outCol = outCol + 1

    For Each srcCol In detailCols
        worksheetSummary.Cells(outRow, outCol).value = worksheetSource.Cells(sourceRow, srcCol).value
        outCol = outCol + 1
    Next

    For Each srcCol In scoreCols
        worksheetSummary.Cells(outRow, outCol).value = worksheetSource.Cells(sourceRow, srcCol).value
        outCol = outCol + 1

        If srcCol = MBO_UPPER_TOTAL_COL_U Then
            worksheetSummary.Cells(outRow, outCol).value = worksheetSource.Cells(MBO_UPPER_TOTAL_ROW, MBO_UPPER_TOTAL_COL_U).value
            outCol = outCol + 1
            worksheetSummary.Cells(outRow, outCol).value = worksheetSource.Cells(MBO_UPPER_TOTAL_ROW, MBO_UPPER_RANK_COL_S).value
            outCol = outCol + 1
        End If

        If srcCol = MBO_UPPER_TOTAL_COL_Z Then
            worksheetSummary.Cells(outRow, outCol).value = worksheetSource.Cells(MBO_UPPER_TOTAL_ROW, MBO_UPPER_TOTAL_COL_Z).value
            outCol = outCol + 1
            worksheetSummary.Cells(outRow, outCol).value = worksheetSource.Cells(MBO_UPPER_TOTAL_ROW, MBO_UPPER_RANK_COL_X).value
            outCol = outCol + 1
        End If
    Next

End Sub

Private Function GetUpperTopSourceColumns() As Variant
    GetUpperTopSourceColumns = Array(2, 4, 5, 6, 9, 11, 13)
End Function

Private Function GetUpperDetailSourceColumns() As Variant
    GetUpperDetailSourceColumns = Array(3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14)
End Function

Private Function GetUpperScoreSourceColumns() As Variant
    GetUpperScoreSourceColumns = Array(16, 20, 21, 22, 25, 26)
End Function

Private Function IsUpperHalfPeriodRow(ByVal worksheetSource As Worksheet, ByVal rowData As Long) As Boolean

    Dim valuePeriod As String

    valuePeriod = Trim$(CStr(worksheetSource.Cells(rowData, MBO_UPPER_PERIOD_COL).value))
    IsUpperHalfPeriodRow = (StrComp(valuePeriod, MBO_UPPER_PERIOD_VALUE, vbTextCompare) = 0)

End Function

Private Sub ApplyUpperHalfConditionalFormats(ByVal worksheetSummary As Worksheet, _
                                             ByVal lastRow As Long, _
                                             ByVal lastCol As Long)

    Dim colIndex As Long
    Dim headerText As String

    For colIndex = 1 To lastCol
        headerText = CStr(worksheetSummary.Cells(1, colIndex).value)

        If IsUpperDifficultyHeader(headerText) Or IsUpperAchievementHeader(headerText) Or IsUpperTotalScoreHeader(headerText) Then
            ApplyDifficultyFormatByColumn worksheetSummary, colIndex, lastRow
        End If
    Next colIndex

    ApplyRankFormatByHeader worksheetSummary, lastRow, 1, lastCol

End Sub

Private Sub ApplyUpperHalfNarrowColumnWidths(ByVal worksheetSummary As Worksheet, _
                                             ByVal lastCol As Long)

    Dim colIndex As Long
    Dim headerText As String

    For colIndex = 1 To lastCol
        headerText = CStr(worksheetSummary.Cells(1, colIndex).value)
        If ShouldNarrowUpperColumn(headerText) Then
            worksheetSummary.Columns(colIndex).ColumnWidth = MBO_UPPER_OUTPUT_TEXT_COLUMN_WIDTH
        End If
    Next colIndex

End Sub

Private Function ShouldNarrowUpperColumn(ByVal headerText As String) As Boolean

    If InStr(1, headerText, "目標設定", vbTextCompare) > 0 And _
       InStr(1, headerText, "達成基準", vbTextCompare) > 0 Then
        ShouldNarrowUpperColumn = True
        Exit Function
    End If

    If InStr(1, headerText, "施策", vbTextCompare) > 0 And _
       InStr(1, headerText, "手段", vbTextCompare) > 0 Then
        ShouldNarrowUpperColumn = True
        Exit Function
    End If

    If InStr(1, headerText, "自己評価達成状況", vbTextCompare) > 0 And _
       InStr(1, headerText, "振り返り", vbTextCompare) > 0 Then
        ShouldNarrowUpperColumn = True
        Exit Function
    End If

    If InStr(1, headerText, "上司評価達成状況", vbTextCompare) > 0 And _
       InStr(1, headerText, "振り返り", vbTextCompare) > 0 Then
        ShouldNarrowUpperColumn = True
        Exit Function
    End If

    ShouldNarrowUpperColumn = False

End Function

Private Sub MergeUpperRepeatedSummaryCells(ByVal worksheetSummary As Worksheet, _
                                           ByVal startRow As Long, _
                                           ByVal endRow As Long, _
                                           ByVal lastCol As Long)

    Dim colIndex As Long
    Dim headerText As String
    Dim mergeRange As Range

    If endRow <= startRow Then Exit Sub

    For colIndex = 1 To lastCol
        headerText = CStr(worksheetSummary.Cells(1, colIndex).value)
        If ShouldMergeUpperSummaryColumn(headerText) Then
            Set mergeRange = worksheetSummary.Range( _
                worksheetSummary.Cells(startRow, colIndex), _
                worksheetSummary.Cells(endRow, colIndex))

            If mergeRange.MergeCells Then mergeRange.UnMerge
            If mergeRange.Rows.Count > 1 Then
                mergeRange.Offset(1, 0).Resize(mergeRange.Rows.Count - 1, 1).ClearContents
            End If
            mergeRange.Merge
            mergeRange.VerticalAlignment = xlCenter
            mergeRange.Font.Size = Application.StandardFontSize + MERGED_FONT_SIZE_DELTA
        End If
    Next colIndex

End Sub

Private Function ShouldMergeUpperSummaryColumn(ByVal headerText As String) As Boolean

    If headerText = MBO_UPPER_HEADER_SELF_TOTAL Then
        ShouldMergeUpperSummaryColumn = True
        Exit Function
    End If

    If headerText = MBO_UPPER_HEADER_SELF_RANK Then
        ShouldMergeUpperSummaryColumn = True
        Exit Function
    End If

    If headerText = MBO_UPPER_HEADER_BOSS_TOTAL Then
        ShouldMergeUpperSummaryColumn = True
        Exit Function
    End If

    If headerText = MBO_UPPER_HEADER_BOSS_RANK Then
        ShouldMergeUpperSummaryColumn = True
        Exit Function
    End If

    If InStr(1, headerText, "目標点数", vbTextCompare) > 0 And _
       InStr(1, headerText, "II", vbTextCompare) > 0 And _
       InStr(1, headerText, "合計", vbTextCompare) > 0 Then
        ShouldMergeUpperSummaryColumn = True
        Exit Function
    End If

    ShouldMergeUpperSummaryColumn = False

End Function

Private Function IsUpperDifficultyHeader(ByVal headerText As String) As Boolean

    IsUpperDifficultyHeader = (InStr(1, headerText, "難易度", vbTextCompare) > 0 And _
                              (InStr(1, headerText, "F", vbTextCompare) > 0 Or _
                               InStr(1, headerText, "H", vbTextCompare) > 0))

End Function

Private Function IsUpperAchievementHeader(ByVal headerText As String) As Boolean

    IsUpperAchievementHeader = (InStr(1, headerText, "達成度", vbTextCompare) > 0)

End Function

Private Function IsUpperTotalScoreHeader(ByVal headerText As String) As Boolean

    IsUpperTotalScoreHeader = (InStr(1, headerText, "達成点", vbTextCompare) > 0 And _
                              InStr(1, headerText, "合計", vbTextCompare) > 0) Or _
                             (InStr(1, headerText, "評価合計", vbTextCompare) > 0)

End Function

Private Sub ApplyDifficultyFormatByColumn(ByVal worksheetSummary As Worksheet, _
                                          ByVal targetCol As Long, _
                                          ByVal lastRow As Long)

    Dim targetRange As Range
    Dim firstCellAddress As String

    If lastRow < 2 Then Exit Sub

    Set targetRange = worksheetSummary.Range( _
        worksheetSummary.Cells(2, targetCol), _
        worksheetSummary.Cells(lastRow, targetCol))

    firstCellAddress = targetRange.Cells(1, 1).Address(False, False)

    targetRange.FormatConditions.Delete

    With targetRange.FormatConditions.Add( _
            Type:=xlExpression, _
            Formula1:="=AND(NOT(ISBLANK(" & firstCellAddress & ")), " & firstCellAddress & "<=0.9)")
        .Interior.Color = RGB(217, 217, 217)
    End With

    With targetRange.FormatConditions.Add( _
            Type:=xlCellValue, Operator:=xlEqual, Formula1:="1.1")
        .Interior.Color = RGB(255, 242, 204)
    End With

    With targetRange.FormatConditions.Add( _
            Type:=xlCellValue, Operator:=xlEqual, Formula1:="1.2")
        .Interior.Color = RGB(248, 203, 173)
    End With

End Sub

Private Sub ApplyRankFormatByHeader(ByVal worksheetSummary As Worksheet, _
                                    ByVal lastRow As Long, _
                                    ByVal firstCol As Long, _
                                    ByVal lastCol As Long)

    Dim colIndex As Long
    Dim headerText As String

    For colIndex = firstCol To lastCol
        headerText = CStr(worksheetSummary.Cells(1, colIndex).value)
        If InStr(1, headerText, "ランク", vbTextCompare) > 0 Then
            ApplyRankFormatByColumn worksheetSummary, colIndex, lastRow
        End If
    Next colIndex

End Sub

Private Sub ApplyRankFormatByColumn(ByVal worksheetSummary As Worksheet, _
                                    ByVal targetCol As Long, _
                                    ByVal lastRow As Long)

    Dim targetRange As Range
    Dim firstCellAddress As String

    If lastRow < 2 Then Exit Sub

    Set targetRange = worksheetSummary.Range( _
        worksheetSummary.Cells(2, targetCol), _
        worksheetSummary.Cells(lastRow, targetCol))

    firstCellAddress = targetRange.Cells(1, 1).Address(False, False)

    targetRange.FormatConditions.Delete

    With targetRange.FormatConditions.Add( _
            Type:=xlExpression, _
            Formula1:="=OR(" & firstCellAddress & "=""SS""," & firstCellAddress & "=""S"")")
        .Font.Color = RGB(255, 0, 0)
        .Font.Bold = True
    End With

    With targetRange.FormatConditions.Add( _
            Type:=xlExpression, _
            Formula1:="=" & firstCellAddress & "=""A""")
        .Font.Color = RGB(255, 0, 0)
        .Font.Bold = False
    End With

    With targetRange.FormatConditions.Add( _
            Type:=xlExpression, _
            Formula1:="=" & firstCellAddress & "=""E""")
        .Font.Color = RGB(192, 0, 0)
        .Font.Bold = True
    End With

    With targetRange.FormatConditions.Add( _
            Type:=xlExpression, _
            Formula1:="=" & firstCellAddress & "=""D""")
        .Font.Color = RGB(192, 0, 0)
        .Font.Bold = False
    End With

End Sub

Private Function GetUpperHeaderText(ByVal worksheetSource As Worksheet, ByVal rowIndex As Long, ByVal colIndex As Long) As String

    Dim s As String

    s = Replace$(CStr(worksheetSource.Cells(rowIndex, colIndex).value), vbCrLf, "")
    s = Replace$(s, vbLf, "")
    s = Trim$(s)

    If Len(s) = 0 Then
        s = "列" & ColumnLetter(colIndex)
    End If

    GetUpperHeaderText = s

End Function

Private Function GetUpperMergedHeaderText(ByVal worksheetSource As Worksheet, ByVal colIndex As Long) As String

    Dim h1 As String
    Dim h2 As String

    h1 = Replace$(CStr(worksheetSource.Cells(MBO_UPPER_DETAIL_HEADER_ROW1, colIndex).value), vbCrLf, "")
    h1 = Replace$(h1, vbLf, "")
    h1 = Trim$(h1)

    h2 = Replace$(CStr(worksheetSource.Cells(MBO_UPPER_DETAIL_HEADER_ROW2, colIndex).value), vbCrLf, "")
    h2 = Replace$(h2, vbLf, "")
    h2 = Trim$(h2)

    Select Case True
        Case Len(h1) > 0 And Len(h2) > 0
            GetUpperMergedHeaderText = h1 & h2
        Case Len(h1) > 0
            GetUpperMergedHeaderText = h1
        Case Len(h2) > 0
            GetUpperMergedHeaderText = h2
        Case Else
            GetUpperMergedHeaderText = "列" & ColumnLetter(colIndex)
    End Select

End Function

Private Function ColumnLetter(ByVal colIndex As Long) As String
    ColumnLetter = Split(Cells(1, colIndex).Address(False, False), "$", 2)(0)
End Function


'==========================
' 前期MBO 取り込み（追記）
'  - 前期MBO\配下(個人フォルダ)\MBOシート_人事評価シート*.xlsx を走査
'  - 「年間総合評価」シートから値を取得
'  - J5の氏名 と 一覧(B列)の氏名 が一致する行へ Z列以降に書き込み
'  - 同一氏名が複数行ある場合、Z～AE列を縦結合（開始～終了行）
'==========================
Public Sub UpdatePrevTermEvaluation(ByVal worksheetSummary As Worksheet)

    '（定数はファイル先頭で一括定義）

    Dim workbookTarget As Workbook
    Dim pathPrevRoot As String

    Dim fileSystem As Object
    Dim folderRoot As Object
    Dim folderPerson As Object
    Dim fileMbo As Object

    Dim workbookSource As Workbook
    Dim worksheetEval As Worksheet
    Dim worksheetSource As Worksheet
    Dim worksheetLog As Worksheet

    Dim dictRangeByEmployeeNo As Object 'Scripting.Dictionary
    Dim dictRangeByName As Object 'Scripting.Dictionary（フォールバック）
    Dim lastRow As Long
    Dim r As Long
    Dim rowLog As Long

    Dim keyEmployeeNo As String
    Dim keyName As String
    Dim startRow As Long
    Dim endRow As Long

    Dim valueMboRank As Variant
    Dim valueMboPoint As Variant
    Dim valueCmRank As Variant
    Dim valueCmPoint As Variant
    Dim valueTotalRank As Variant
    Dim valueTotalPoint As Variant

    Dim rngMerge As Range
    Dim outFirstCol As Long
    Dim outLastCol As Long
    Dim folderHasTargetFile As Boolean

    Dim countFoldersScanned As Long
    Dim countFoldersTarget As Long
    Dim countFilesScanned As Long
    Dim countFilesTarget As Long
    Dim countFilesOpened As Long
    Dim countMatchByEmployeeNo As Long
    Dim countMatchByName As Long
    Dim countNoMatch As Long
    Dim countSheetMissing As Long
    Dim countOpenError As Long

    Set workbookTarget = ThisWorkbook

    If workbookTarget.Path = vbNullString Then
        MsgBox MSG_SAVE_REQUIRED, vbExclamation
        Exit Sub
    End If

    pathPrevRoot = workbookTarget.Path & Application.PathSeparator & MBO_PREV_FOLDER_NAME
    If Dir$(pathPrevRoot, vbDirectory) = vbNullString Then
        MsgBox MSG_FOLDER_PREV_MBO_NOT_FOUND & vbCrLf & pathPrevRoot, vbExclamation
        Exit Sub
    End If

    '------------------------------
    ' 一覧の「開始行～終了行」インデックス作成
    '  - 主キー：社員No.
    '  - 補助キー：氏名（社員No.が取れない場合のフォールバック）
    '------------------------------
    lastRow = worksheetSummary.Cells(worksheetSummary.Rows.Count, MboColEmployeeName).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox MSG_LIST_DATA_ROW_NOT_FOUND, vbExclamation
        Exit Sub
    End If

    Set dictRangeByEmployeeNo = CreateObject("Scripting.Dictionary")
    dictRangeByEmployeeNo.CompareMode = 1 'vbTextCompare

    Set dictRangeByName = CreateObject("Scripting.Dictionary")
    dictRangeByName.CompareMode = 1 'vbTextCompare

    Set worksheetLog = GetLogSheetForSummary(workbookTarget, worksheetSummary.Name)
    If worksheetLog Is Nothing Then
        Set worksheetLog = CreateLogSheet(workbookTarget, worksheetSummary.Name)
        SetLogHeader worksheetLog
        ApplyLogSheetTabColor worksheetLog
        rowLog = 2
        AppendLogRow worksheetLog, rowLog, LOG_FOLDER_INFO, "", LOG_STATUS_INFO003, LOG_MATCH_TYPE_NONE, "", "", LOG_MSG_CREATED_NEW_LOG_SHEET
    Else
        ApplyLogSheetTabColor worksheetLog
        rowLog = worksheetLog.Cells(worksheetLog.Rows.Count, 1).End(xlUp).Row + 1
        If rowLog < 2 Then rowLog = 2
    End If

    '前提：CreateMboListの出力は同一氏名が連続して並ぶ
    r = 2
    Do While r <= lastRow

        keyEmployeeNo = NormalizeEmployeeNo(worksheetSummary.Cells(r, MboColEmployeeNo).value)
        keyName = NormalizeName(CStr(worksheetSummary.Cells(r, MboColEmployeeName).value))

        If Len(keyEmployeeNo) = 0 And Len(keyName) = 0 Then
            r = r + 1
        Else
            startRow = r
            endRow = r

            Do While endRow + 1 <= lastRow
                If Len(keyEmployeeNo) > 0 Then
                    If NormalizeEmployeeNo(worksheetSummary.Cells(endRow + 1, MboColEmployeeNo).value) = keyEmployeeNo Then
                        endRow = endRow + 1
                    Else
                        Exit Do
                    End If
                Else
                    If NormalizeName(CStr(worksheetSummary.Cells(endRow + 1, MboColEmployeeName).value)) = keyName Then
                        endRow = endRow + 1
                    Else
                        Exit Do
                    End If
                End If
            Loop

            If Len(keyEmployeeNo) > 0 Then
                If Not dictRangeByEmployeeNo.Exists(keyEmployeeNo) Then
                    dictRangeByEmployeeNo.Add keyEmployeeNo, Array(startRow, endRow)
                End If
            End If

            If Len(keyName) > 0 Then
                If Not dictRangeByName.Exists(keyName) Then
                    'Valueに配列（開始行, 終了行）
                    dictRangeByName.Add keyName, Array(startRow, endRow)
                End If
            End If

            r = endRow + 1
        End If
    Loop

    '------------------------------
    ' ヘッダ（Z列以降）を設定
    '------------------------------
    With worksheetSummary
        .Cells(1, MboPrevOutColMboRank).value = MBO_HEADER_PREV_MBO_RANK
        .Cells(1, MboPrevOutColMboPoint).value = MBO_HEADER_PREV_MBO_POINT
        .Cells(1, MboPrevOutColCmRank).value = MBO_HEADER_PREV_CM_RANK
        .Cells(1, MboPrevOutColCmPoint).value = MBO_HEADER_PREV_CM_POINT
        .Cells(1, MboPrevOutColTotalRank).value = MBO_HEADER_PREV_TOTAL_RANK
        .Cells(1, MboPrevOutColTotalPoint).value = MBO_HEADER_PREV_TOTAL_POINT

        'ヘッダ装飾（既存と同じ濃青＋白文字）
        With .Range(.Cells(1, MboPrevOutColMboRank), .Cells(1, MboPrevOutColTotalPoint))
            .Interior.Color = RGB(0, 102, 204)
            .Font.Color = vbWhite
            .Font.Bold = True
        End With
    End With

    outFirstCol = MboPrevOutColMboRank
    outLastCol = MboPrevOutColTotalPoint

    '------------------------------
    ' 前期MBO 配下の走査
    '------------------------------
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Set folderRoot = fileSystem.GetFolder(pathPrevRoot)

    For Each folderPerson In folderRoot.SubFolders
        countFoldersScanned = countFoldersScanned + 1
        folderHasTargetFile = False

        For Each fileMbo In folderPerson.Files
            countFilesScanned = countFilesScanned + 1

            '「MBOシート人事評価シート」ファイルのみ対象
            If IsTargetMboExcelFile(fileMbo.Path) Then
                countFilesTarget = countFilesTarget + 1
                folderHasTargetFile = True

                Set workbookSource = Nothing
                On Error Resume Next
                Set workbookSource = OpenWorkbookReadOnlyNoUpdate(fileMbo.Path)
                If Err.Number <> 0 Or workbookSource Is Nothing Then
                    countOpenError = countOpenError + 1
                    AppendLogRow worksheetLog, rowLog, folderPerson.Name, fileMbo.Name, LOG_STATUS_ERROR001, LOG_MATCH_TYPE_NONE, "", "", LOG_MSG_OPEN_ERROR_PREFIX & Err.Description
                    Err.Clear
                    On Error GoTo 0
                    GoTo NextPrevFile
                End If
                On Error GoTo 0
                countFilesOpened = countFilesOpened + 1

                On Error Resume Next
                Set worksheetEval = workbookSource.Worksheets(MBO_TARGET_SHEET_NAME)
                Set worksheetSource = workbookSource.Worksheets(MBO_SOURCE_SHEET_NAME)
                On Error GoTo 0

                If Not worksheetEval Is Nothing Then

                    keyEmployeeNo = vbNullString
                    If Not worksheetSource Is Nothing Then
                        keyEmployeeNo = NormalizeEmployeeNo(worksheetSource.Cells(MboGetRowEmployeeNo, MboGetColEmployeeNo).value)
                    End If

                    If Len(keyEmployeeNo) > 0 And dictRangeByEmployeeNo.Exists(keyEmployeeNo) Then
                        startRow = CLng(dictRangeByEmployeeNo(keyEmployeeNo)(0))
                        endRow = CLng(dictRangeByEmployeeNo(keyEmployeeNo)(1))
                        countMatchByEmployeeNo = countMatchByEmployeeNo + 1
                        keyName = NormalizeName(CStr(worksheetEval.Cells(MboPrevSrcNameRow, MboPrevSrcNameCol).value))
                    Else
                        keyName = NormalizeName(CStr(worksheetEval.Cells(MboPrevSrcNameRow, MboPrevSrcNameCol).value))
                        If Len(keyName) > 0 And dictRangeByName.Exists(keyName) Then
                            startRow = CLng(dictRangeByName(keyName)(0))
                            endRow = CLng(dictRangeByName(keyName)(1))
                            countMatchByName = countMatchByName + 1
                        Else
                            countNoMatch = countNoMatch + 1
                            AppendLogRow worksheetLog, rowLog, folderPerson.Name, fileMbo.Name, LOG_STATUS_WORN004, LOG_MATCH_TYPE_NONE, keyEmployeeNo, keyName, LOG_MSG_NO_MATCH_EMP_NO_NAME
                            GoTo NextPrevFile
                        End If
                    End If

                        '値を取得
                        valueMboRank = worksheetEval.Cells(MboPrevSrcMboRow, MboPrevSrcRankCol).value
                        valueMboPoint = worksheetEval.Cells(MboPrevSrcMboRow, MboPrevSrcPointCol).value
                        valueCmRank = worksheetEval.Cells(MboPrevSrcCmRow, MboPrevSrcRankCol).value
                        valueCmPoint = worksheetEval.Cells(MboPrevSrcCmRow, MboPrevSrcPointCol).value
                        valueTotalRank = worksheetEval.Cells(MboPrevSrcTotalRow, MboPrevSrcRankCol).value
                        valueTotalPoint = worksheetEval.Cells(MboPrevSrcTotalRow, MboPrevSrcPointCol).value

                        With worksheetSummary
                            'まず開始行に書く
                            .Cells(startRow, MboPrevOutColMboRank).value = valueMboRank
                            .Cells(startRow, MboPrevOutColMboPoint).value = valueMboPoint
                            .Cells(startRow, MboPrevOutColCmRank).value = valueCmRank
                            .Cells(startRow, MboPrevOutColCmPoint).value = valueCmPoint
                            .Cells(startRow, MboPrevOutColTotalRank).value = valueTotalRank
                            .Cells(startRow, MboPrevOutColTotalPoint).value = valueTotalPoint

                            '複数行なら縦結合（Z～AE を各列ごとに）
                            If endRow > startRow Then
                                Dim c As Long
                                For c = outFirstCol To outLastCol
                                    Set rngMerge = .Range(.Cells(startRow, c), .Cells(endRow, c))
                                    If rngMerge.MergeCells Then rngMerge.UnMerge
                                    If rngMerge.Rows.Count > 1 Then
                                        rngMerge.Offset(1, 0).Resize(rngMerge.Rows.Count - 1, 1).ClearContents
                                    End If
                                    rngMerge.Merge
                                    rngMerge.VerticalAlignment = xlCenter
                                    rngMerge.Font.Size = Application.StandardFontSize + MERGED_FONT_SIZE_DELTA
                                Next c
                            End If
                        End With

                    If Len(keyEmployeeNo) > 0 And dictRangeByEmployeeNo.Exists(keyEmployeeNo) Then
                        AppendLogRow worksheetLog, rowLog, folderPerson.Name, fileMbo.Name, LOG_STATUS_INFO001, LOG_MATCH_TYPE_EMPLOYEE_NO, keyEmployeeNo, keyName, LOG_MSG_ROW_OUTPUT_PREFIX & CStr(startRow) & "-" & CStr(endRow) & LOG_MSG_ROW_REFLECT_SUFFIX
                    Else
                        AppendLogRow worksheetLog, rowLog, folderPerson.Name, fileMbo.Name, LOG_STATUS_INFO001, LOG_MATCH_TYPE_NAME, keyEmployeeNo, keyName, LOG_MSG_ROW_OUTPUT_PREFIX & CStr(startRow) & "-" & CStr(endRow) & LOG_MSG_ROW_REFLECT_SUFFIX
                    End If
                Else
                    countSheetMissing = countSheetMissing + 1
                    AppendLogRow worksheetLog, rowLog, folderPerson.Name, fileMbo.Name, LOG_STATUS_WORN002, LOG_MATCH_TYPE_NONE, "", "", MBO_TARGET_SHEET_NAME & LOG_MSG_SHEET_MISSING_SUFFIX
                End If

NextPrevFile:
                If Not workbookSource Is Nothing Then
                    workbookSource.Close SaveChanges:=False
                End If
                Set worksheetEval = Nothing
                Set worksheetSource = Nothing
                Set workbookSource = Nothing
            Else
                AppendLogRow worksheetLog, rowLog, folderPerson.Name, fileMbo.Name, LOG_STATUS_WORN003, LOG_MATCH_TYPE_NONE, "", "", LOG_MSG_SKIP_NON_TARGET
            End If
        Next fileMbo

        If folderHasTargetFile Then
            countFoldersTarget = countFoldersTarget + 1
        End If
    Next folderPerson

    '------------------------------
    ' 仕上げ：列幅（Z～AEは50pxほどの幅）、フィルタ掛け直し
    '------------------------------
    With worksheetSummary
    
        'Z～AE は固定
        Dim colIndex As Long
        For colIndex = outFirstCol To outLastCol
            .Columns(colIndex).ColumnWidth = 5
        Next colIndex
    
        'Z～AE は中央寄せ
        .Range(.Cells(1, outFirstCol), .Cells(lastRow, outLastCol)).HorizontalAlignment = xlCenter
    
        'Z～AE 以外は AutoFit
        .Range(.Cells(1, MboColEmployeeNo), .Cells(lastRow, outFirstCol - 1)).EntireColumn.AutoFit
        
        'フィルタを掛け直し（ヘッダ1行目）
        If .AutoFilterMode Then .AutoFilterMode = False
        .Range(.Cells(1, MboColEmployeeNo), .Cells(lastRow, outLastCol)).AutoFilter
    
        '--- Z～AE も縞々（CreateMboList側の行色を踏襲）
        Dim rowIndex As Long
        Dim stripeColor As Long
        
        For rowIndex = 2 To lastRow
            'A列(社員No.)の背景色を拾って、Z～AEへ適用
            stripeColor = worksheetSummary.Cells(rowIndex, MboColEmployeeNo).Interior.Color
            worksheetSummary.Range( _
                worksheetSummary.Cells(rowIndex, outFirstCol), _
                worksheetSummary.Cells(rowIndex, outLastCol) _
            ).Interior.Color = stripeColor
        Next rowIndex
    
    End With

    ApplyRankFormatByHeader worksheetSummary, lastRow, outFirstCol, outLastCol

    AppendLogRow worksheetLog, rowLog, LOG_FOLDER_SUMMARY, "", LOG_STATUS_INFO002, LOG_MATCH_TYPE_NONE, "", "", LOG_MSG_SUM_FOLDERS_SCANNED & CStr(countFoldersScanned)
    AppendLogRow worksheetLog, rowLog, LOG_FOLDER_SUMMARY, "", LOG_STATUS_INFO002, LOG_MATCH_TYPE_NONE, "", "", LOG_MSG_SUM_FOLDERS_TARGET & CStr(countFoldersTarget)
    AppendLogRow worksheetLog, rowLog, LOG_FOLDER_SUMMARY, "", LOG_STATUS_INFO002, LOG_MATCH_TYPE_NONE, "", "", LOG_MSG_SUM_FILES_SCANNED & CStr(countFilesScanned)
    AppendLogRow worksheetLog, rowLog, LOG_FOLDER_SUMMARY, "", LOG_STATUS_INFO002, LOG_MATCH_TYPE_NONE, "", "", LOG_MSG_SUM_FILES_TARGET & CStr(countFilesTarget)
    AppendLogRow worksheetLog, rowLog, LOG_FOLDER_SUMMARY, "", LOG_STATUS_INFO002, LOG_MATCH_TYPE_NONE, "", "", LOG_MSG_SUM_FILES_OPENED & CStr(countFilesOpened)
    AppendLogRow worksheetLog, rowLog, LOG_FOLDER_SUMMARY, "", LOG_STATUS_INFO002, LOG_MATCH_TYPE_NONE, "", "", LOG_MSG_SUM_MATCH_EMP_NO & CStr(countMatchByEmployeeNo)
    AppendLogRow worksheetLog, rowLog, LOG_FOLDER_SUMMARY, "", LOG_STATUS_INFO002, LOG_MATCH_TYPE_NONE, "", "", LOG_MSG_SUM_MATCH_NAME & CStr(countMatchByName)
    AppendLogRow worksheetLog, rowLog, LOG_FOLDER_SUMMARY, "", LOG_STATUS_INFO002, LOG_MATCH_TYPE_NONE, "", "", LOG_MSG_SUM_NO_MATCH & CStr(countNoMatch)
    AppendLogRow worksheetLog, rowLog, LOG_FOLDER_SUMMARY, "", LOG_STATUS_INFO002, LOG_MATCH_TYPE_NONE, "", "", LOG_MSG_SUM_EVAL_SHEET_MISSING & CStr(countSheetMissing)
    AppendLogRow worksheetLog, rowLog, LOG_FOLDER_SUMMARY, "", LOG_STATUS_INFO002, LOG_MATCH_TYPE_NONE, "", "", LOG_MSG_SUM_OPEN_ERROR & CStr(countOpenError)

    worksheetLog.Columns("A:I").EntireColumn.AutoFit

            MsgBox MSG_DONE_UPDATE_PREV & vbCrLf & _
                MSG_LABEL_LOG_SHEET & worksheetLog.Name & vbCrLf & _
            LOG_MSG_SUM_FOLDERS_SCANNED & countFoldersScanned & vbCrLf & _
            LOG_MSG_SUM_FOLDERS_TARGET & countFoldersTarget & vbCrLf & _
            LOG_MSG_SUM_FILES_SCANNED & countFilesScanned & vbCrLf & _
            LOG_MSG_SUM_FILES_TARGET & countFilesTarget, vbInformation

End Sub

'Private Function NormalizeName(ByVal valueName As String) As String
'    Dim s As String
'    s = valueName
'    s = Replace$(s, ChrW(&H3000), " ") '全角スペース→半角
'    s = Trim$(s)
'    Do While InStr(s, "  ") > 0
'        s = Replace$(s, "  ", " ")
'    Loop
'    NormalizeName = s
'End Function

Private Function GetCallerWorksheet() As Worksheet

    On Error Resume Next
    If TypeName(ActiveSheet) = "Worksheet" Then
        Set GetCallerWorksheet = ActiveSheet
    Else
        Set GetCallerWorksheet = Nothing
    End If
    On Error GoTo 0

End Function

Private Sub ApplySheetTabColors(ByVal worksheetCaller As Worksheet, _
                                ByVal worksheetSummary As Worksheet, _
                                ByVal worksheetLog As Worksheet)

    If Not worksheetSummary Is Nothing Then
        If Not worksheetCaller Is Nothing Then
            On Error Resume Next
            worksheetSummary.Tab.Color = worksheetCaller.Tab.Color
            On Error GoTo 0
        End If
    End If

    ApplyLogSheetTabColor worksheetLog

End Sub

Private Sub ApplyLogSheetTabColor(ByVal worksheetLog As Worksheet)

    If worksheetLog Is Nothing Then Exit Sub

    On Error Resume Next
    worksheetLog.Tab.Color = SHEET_TAB_COLOR_LOG
    On Error GoTo 0

End Sub

Private Function CreateLogSheet(ByVal wb As Workbook, ByVal summarySheetName As String) As Worksheet

    Dim baseName As String
    Dim candidateName As String
    Dim n As Long

    baseName = GetLogSheetNameFromSummaryName(summarySheetName)

    If Len(baseName) > 31 Then
        baseName = Left$(baseName, 31)
    End If

    candidateName = baseName
    n = 1
    Do While WorksheetExists(wb, candidateName)
        candidateName = Left$(baseName, 28) & "_" & CStr(n)
        n = n + 1
    Loop

    Set CreateLogSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    CreateLogSheet.Name = candidateName

End Function

Private Function GetLogSheetForSummary(ByVal wb As Workbook, ByVal summarySheetName As String) As Worksheet

    Dim logSheetName As String
    Dim ws As Worksheet

    logSheetName = GetLogSheetNameFromSummaryName(summarySheetName)

    On Error Resume Next
    Set ws = wb.Worksheets(logSheetName)
    On Error GoTo 0

    Set GetLogSheetForSummary = ws

End Function

Private Function GetLogSheetNameFromSummaryName(ByVal summarySheetName As String) As String

    Dim baseName As String
    Dim suffix As String

    If Left$(summarySheetName, Len(MBO_INITIAL_SHEET_PREFIX)) = MBO_INITIAL_SHEET_PREFIX Then
        suffix = Mid$(summarySheetName, Len(MBO_INITIAL_SHEET_PREFIX) + 1)
        baseName = "LOG_" & suffix
    ElseIf Left$(summarySheetName, Len(MBO_UPPER_SHEET_PREFIX)) = MBO_UPPER_SHEET_PREFIX Then
        suffix = Mid$(summarySheetName, Len(MBO_UPPER_SHEET_PREFIX) + 1)
        baseName = "LOG_" & suffix
    Else
        baseName = "LOG_" & Format(Now, "yyyymmdd_hhnnss")
    End If

    If Len(baseName) > 31 Then
        baseName = Left$(baseName, 31)
    End If

    GetLogSheetNameFromSummaryName = baseName

End Function

Private Function WorksheetExists(ByVal wb As Workbook, ByVal sheetName As String) As Boolean

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    WorksheetExists = Not ws Is Nothing
    Set ws = Nothing
    On Error GoTo 0

End Function

Private Sub SetLogHeader(ByVal worksheetLog As Worksheet)

    With worksheetLog
        .Cells(1, 1).value = LOG_HEADER_TIME
        .Cells(1, 2).value = LOG_HEADER_FOLDER
        .Cells(1, 3).value = LOG_HEADER_FILE
        .Cells(1, 4).value = LOG_HEADER_STATUS
        .Cells(1, 5).value = LOG_HEADER_MATCH_TYPE
        .Cells(1, 6).value = LOG_HEADER_EMPLOYEE_NO
        .Cells(1, 7).value = LOG_HEADER_NAME
        .Cells(1, 8).value = LOG_HEADER_MESSAGE
        .Cells(1, 9).value = LOG_HEADER_PATH

        With .Range(.Cells(1, 1), .Cells(1, 9))
            .Interior.Color = RGB(0, 102, 204)
            .Font.Color = vbWhite
            .Font.Bold = True
        End With
    End With

End Sub

Private Sub AppendLogRow(ByVal worksheetLog As Worksheet, _
                         ByRef rowLog As Long, _
                         ByVal folderName As String, _
                         ByVal fileName As String, _
                         ByVal status As String, _
                         ByVal matchType As String, _
                         ByVal employeeNo As String, _
                         ByVal employeeName As String, _
                         ByVal message As String)

    With worksheetLog
        .Cells(rowLog, 1).value = Format(Now, "yyyy/mm/dd hh:nn:ss")
        .Cells(rowLog, 2).value = folderName
        .Cells(rowLog, 3).value = fileName
        .Cells(rowLog, 4).value = status
        .Cells(rowLog, 5).value = matchType
        .Cells(rowLog, 6).value = employeeNo
        .Cells(rowLog, 7).value = employeeName
        .Cells(rowLog, 8).value = message

        If Len(folderName) > 0 And Len(fileName) > 0 Then
            .Cells(rowLog, 9).value = folderName & Application.PathSeparator & fileName
        Else
            .Cells(rowLog, 9).value = ""
        End If
    End With

    rowLog = rowLog + 1

End Sub

Private Function NormalizeEmployeeNo(ByVal valueEmployeeNo As Variant) As String

    Dim s As String

    s = CStr(valueEmployeeNo)
    s = Replace$(s, ChrW(&H3000), " ")
    s = Trim$(s)
    s = Replace$(s, " ", vbNullString)

    On Error Resume Next
    s = StrConv(s, vbNarrow)
    On Error GoTo 0

    NormalizeEmployeeNo = UCase$(s)

End Function

Private Function NormalizeName(ByVal valueName As String) As String

    Dim s As String
    
    s = valueName
    '全角スペース → 半角スペース
    s = Replace$(s, ChrW(&H3000), " ")
    '前後の空白除去
    s = Trim$(s)
    'すべての半角スペースを削除（苗字・名前間の有無を吸収）
    s = Replace$(s, " ", vbNullString)

    NormalizeName = s
    
End Function

Public Function GetLatestMboSheet(ByVal wb As Workbook) As Worksheet
    Dim ws As Worksheet
    Dim latestSheet As Worksheet
    Dim latestTime As Date
    Dim sheetTime As Date
    Dim sheetName As String
    Dim suffix As String
    
    For Each ws In wb.Worksheets
        If Left$(ws.Name, Len(MBO_INITIAL_SHEET_PREFIX)) = MBO_INITIAL_SHEET_PREFIX Then
            sheetName = ws.Name
            suffix = Mid$(sheetName, Len(MBO_INITIAL_SHEET_PREFIX) + 1)
            On Error Resume Next
            sheetTime = CDate( _
                Mid(suffix, 1, 4) & "/" & _
                Mid(suffix, 5, 2) & "/" & _
                Mid(suffix, 7, 2) & " " & _
                Mid(suffix, 10, 2) & ":" & _
                Mid(suffix, 12, 2) & ":" & _
                Mid(suffix, 14, 2))
            On Error GoTo 0
            
            If latestSheet Is Nothing Or sheetTime > latestTime Then
                Set latestSheet = ws
                latestTime = sheetTime
            End If
        End If
    Next ws
    
    Set GetLatestMboSheet = latestSheet
End Function
