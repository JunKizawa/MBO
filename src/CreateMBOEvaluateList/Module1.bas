Attribute VB_Name = "Module1"
Option Explicit

'=====================================================================
' 【1. 基本設定】ファイル名・シート名・拡張子などの定数
'=====================================================================
Public Const MBO_TARGET_FILE_KEYWORD As String = "MBOシート_人事評価シート" ' 対象となるMBOファイル名のキーワード
Public Const MBO_TARGET_FILE_EXTS As String = "xlsx,xlsm,xlsb,xls" ' 対象拡張子（カンマ区切り文字列）
Public Const MBO_PREV_FOLDER_NAME As String = "前期MBO" ' 前期MBOフォルダ名
Public Const MBO_TARGET_SHEET_NAME As String = "年間総合評価" ' 評価対象シート名

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

'=====================================================================
' 【9. 達成基準シート関連の定数】
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
    
    Dim fileIndex As Long
    Dim rowFileStart As Long
    Dim rowFileEnd As Long
    
    Dim lastRow As Long
    Dim rangeAll As Range
    Dim defaultRowHeight As Double
    Dim i As Long
    Dim leftCols As Variant
    
    Set workbookTarget = ThisWorkbook
    
    If workbookTarget.Path = vbNullString Then
        MsgBox "このブックを一度保存してから実行してください。", vbExclamation
        Exit Sub
    End If
    
    pathMbo = workbookTarget.Path & Application.PathSeparator & "MBO"
    If Dir$(pathMbo, vbDirectory) = vbNullString Then
        MsgBox "MBO フォルダが見つかりません。" & vbCrLf & pathMbo, vbExclamation
        Exit Sub
    End If
    
    '--- 一覧シートを新規作成
    sheetNameNew = "MBO_" & Format(Now, "yyyymmdd_hhnnss")
    Set worksheetSummary = workbookTarget.Worksheets.Add(After:=workbookTarget.Worksheets(workbookTarget.Worksheets.Count))
    On Error Resume Next
    worksheetSummary.Name = sheetNameNew
    On Error GoTo 0
    
    '--- 見出し行を設定（ヘッダ色設定含む）
    Call SetSummaryHeader(worksheetSummary)
    rowSummary = 2
    
    '--- FSO 初期化
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Set folderRoot = fileSystem.GetFolder(pathMbo)
    
    fileIndex = 0
    
    '--- MBO\配下の各フォルダ＆ファイルを走査
    For Each folderPerson In folderRoot.SubFolders
        
        For Each fileMbo In folderPerson.Files
            
            '「MBOシート人事評価シート」ファイルのみ対象
            If IsTargetMboExcelFile(fileMbo.Path) Then
                
                rowFileStart = rowSummary   'このファイルの開始行
                
                '--- 対象ブックを開く
                Set workbookSource = Workbooks.Open(fileMbo.Path, ReadOnly:=True)
                
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
                        Dim rangeBlock As Range
                        Set rangeBlock = worksheetSummary.Range( _
                            worksheetSummary.Cells(rowFileStart, MboColEmployeeNo), _
                            worksheetSummary.Cells(rowFileEnd, MboColGoalScore))
                        
                        If fileIndex Mod 2 = 0 Then
                            rangeBlock.Interior.Color = RGB(221, 235, 247)   '薄い青
                        Else
                            rangeBlock.Interior.Color = vbWhite              '白
                        End If
                    End If
                    
                End If
                
                workbookSource.Close SaveChanges:=False
                Set worksheetSource = Nothing
                Set worksheetCriteria = Nothing
                
            End If
        Next fileMbo
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
    
    MsgBox "一覧作成が完了しました。" & vbCrLf & _
           "シート名: " & worksheetSummary.Name, vbInformation
    
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

    Dim dictRangeByName As Object 'Scripting.Dictionary
    Dim lastRow As Long
    Dim r As Long

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

    Set workbookTarget = ThisWorkbook

    If workbookTarget.Path = vbNullString Then
        MsgBox "このブックを一度保存してから実行してください。", vbExclamation
        Exit Sub
    End If

    pathPrevRoot = workbookTarget.Path & Application.PathSeparator & MBO_PREV_FOLDER_NAME
    If Dir$(pathPrevRoot, vbDirectory) = vbNullString Then
        MsgBox "前期MBO フォルダが見つかりません。" & vbCrLf & pathPrevRoot, vbExclamation
        Exit Sub
    End If

    '------------------------------
    ' 一覧(B列=氏名)の「開始行～終了行」インデックス作成
    '------------------------------
    lastRow = worksheetSummary.Cells(worksheetSummary.Rows.Count, MboColEmployeeName).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "一覧シートにデータ行がありません。", vbExclamation
        Exit Sub
    End If

    Set dictRangeByName = CreateObject("Scripting.Dictionary")
    dictRangeByName.CompareMode = 1 'vbTextCompare

    '前提：CreateMboListの出力は同一氏名が連続して並ぶ
    r = 2
    Do While r <= lastRow

        keyName = NormalizeName(CStr(worksheetSummary.Cells(r, MboColEmployeeName).value))

        If Len(keyName) = 0 Then
            r = r + 1
        Else
            startRow = r
            endRow = r

            Do While endRow + 1 <= lastRow
                If NormalizeName(CStr(worksheetSummary.Cells(endRow + 1, MboColEmployeeName).value)) = keyName Then
                    endRow = endRow + 1
                Else
                    Exit Do
                End If
            Loop

            If Not dictRangeByName.Exists(keyName) Then
                'Valueに配列（開始行, 終了行）
                dictRangeByName.Add keyName, Array(startRow, endRow)
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
        For Each fileMbo In folderPerson.Files

            '「MBOシート人事評価シート」ファイルのみ対象
            If IsTargetMboExcelFile(fileMbo.Path) Then

                Set workbookSource = Workbooks.Open(fileMbo.Path, ReadOnly:=True)

                On Error Resume Next
                Set worksheetEval = workbookSource.Worksheets(MBO_TARGET_SHEET_NAME)
                On Error GoTo 0

                If Not worksheetEval Is Nothing Then

                    keyName = NormalizeName(CStr(worksheetEval.Cells(MboPrevSrcNameRow, MboPrevSrcNameCol).value))

                    If Len(keyName) > 0 And dictRangeByName.Exists(keyName) Then

                        startRow = CLng(dictRangeByName(keyName)(0))
                        endRow = CLng(dictRangeByName(keyName)(1))

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
                                    rngMerge.Merge
                                    rngMerge.VerticalAlignment = xlCenter
                                Next c
                            End If
                        End With

                    End If
                End If

                workbookSource.Close SaveChanges:=False
                Set worksheetEval = Nothing
            End If
        Next fileMbo
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

    MsgBox "前期MBO の追記が完了しました。", vbInformation

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
    
    For Each ws In wb.Worksheets
        If ws.Name Like "MBO_########_######" Or ws.Name Like "MBO_*" Then
            sheetName = ws.Name
            On Error Resume Next
            sheetTime = CDate( _
                Mid(sheetName, 5, 4) & "/" & _
                Mid(sheetName, 9, 2) & "/" & _
                Mid(sheetName, 11, 2) & " " & _
                Mid(sheetName, 14, 2) & ":" & _
                Mid(sheetName, 16, 2) & ":" & _
                Mid(sheetName, 18, 2))
            On Error GoTo 0
            
            If latestSheet Is Nothing Or sheetTime > latestTime Then
                Set latestSheet = ws
                latestTime = sheetTime
            End If
        End If
    Next ws
    
    Set GetLatestMboSheet = latestSheet
End Function
