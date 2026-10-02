Attribute VB_Name = "OfficeReplace"
'==============================================================================
' Office 一括置換マクロ（VBA 版）
'
' ※ 未検証です。Python 版（office-replace コマンド）を正とし、Python が使えない環境向けの参考実装として置いています。
'    使う前に、必ずコピーしたファイルで動作を確かめてください。
'
' 使い方:
'   1. 新しい Excel ブック（.xlsm）にこのモジュールを取り込む（VBE → ファイル → ファイルのインポート）
'   2. シート「置換表」を作り、A 列 = 置換前、B 列 = 置換後、C 列 = 対象（文字列 / シート名 / ファイル名 / すべて）を 2 行目から書く
'      置換前には * と ? が使える（Excel の検索と同じ。Word / PowerPoint は Like 演算子で判定）
'   3. マクロ RunReplace を実行し、入力フォルダと出力フォルダを選ぶ
'   4. 結果はシート「置換ログ」に出る。元のファイルは変更しない（出力フォルダに保存する）
'
' 制約: サブフォルダは対象外 / パスワード付きのファイルは読み飛ばす / Word・PowerPoint のワイルドカードは段落単位の置換になり書式が崩れることがある
'==============================================================================
Option Explicit

Private Type ReplaceRule
    Before As String
    After As String
    Target As String   ' text / sheet / file / all
End Type

Private rules() As ReplaceRule
Private ruleCount As Long
Private logRow As Long

Public Sub RunReplace()
    Dim inDir As String, outDir As String
    inDir = PickFolder("入力フォルダを選んでください")
    If inDir = "" Then Exit Sub
    outDir = PickFolder("出力フォルダを選んでください（入力フォルダとは別の場所）")
    If outDir = "" Then Exit Sub
    If LCase$(outDir) = LCase$(inDir) Or InStr(1, LCase$(outDir), LCase$(inDir) & "\", vbTextCompare) = 1 Then
        MsgBox "出力フォルダを入力フォルダの中にはできません。", vbExclamation
        Exit Sub
    End If
    If Not LoadRules() Then Exit Sub
    PrepareLog

    Dim wd As Object, pp As Object, f As String, ext As String
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    f = Dir(inDir & "\*.*")
    Do While f <> ""
        If Left$(f, 2) <> "~$" Then
            ext = LCase$(Mid$(f, InStrRev(f, ".") + 1))
            Select Case ext
                Case "xlsx", "xlsm": ProcessExcel inDir & "\" & f, outDir & "\" & NewFileName(f)
                Case "docx", "docm"
                    If wd Is Nothing Then Set wd = CreateObject("Word.Application")
                    ProcessWord wd, inDir & "\" & f, outDir & "\" & NewFileName(f)
                Case "pptx", "pptm"
                    If pp Is Nothing Then Set pp = CreateObject("PowerPoint.Application")
                    ProcessPowerPoint pp, inDir & "\" & f, outDir & "\" & NewFileName(f)
                Case "xls", "doc", "ppt": WriteLog f, "対象外", "", "古い形式です。新しい形式で保存し直してください", ""
            End Select
        End If
        f = Dir()
    Loop
    If Not wd Is Nothing Then wd.Quit
    If Not pp Is Nothing Then pp.Quit
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "完了しました。シート「置換ログ」を確認してください。", vbInformation
End Sub

Private Function PickFolder(title As String) As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = title
        If .Show = -1 Then PickFolder = .SelectedItems(1)
    End With
End Function

Private Function LoadRules() As Boolean
    Dim ws As Worksheet, r As Long, last As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("置換表")
    On Error GoTo 0
    If ws Is Nothing Then
        MsgBox "シート「置換表」がありません。", vbExclamation
        Exit Function
    End If
    last = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    ruleCount = 0
    ReDim rules(1 To Application.Max(1, last))
    For r = 2 To last
        If Trim$(CStr(ws.Cells(r, 1).Value)) <> "" Then
            If Replace(CStr(ws.Cells(r, 1).Value), "*", "") = "" Then
                MsgBox r & " 行目: 置換前が * だけでは、すべての文字列が置換されます。", vbExclamation
                Exit Function
            End If
            ruleCount = ruleCount + 1
            rules(ruleCount).Before = CStr(ws.Cells(r, 1).Value)
            rules(ruleCount).After = CStr(ws.Cells(r, 2).Value)
            rules(ruleCount).Target = NormalizeTarget(CStr(ws.Cells(r, 3).Value))
        End If
    Next
    If ruleCount = 0 Then
        MsgBox "置換表にルールがありません。", vbExclamation
        Exit Function
    End If
    LoadRules = True
End Function

Private Function NormalizeTarget(s As String) As String
    Select Case Trim$(s)
        Case "", "文字列", "text": NormalizeTarget = "text"
        Case "シート名", "sheet": NormalizeTarget = "sheet"
        Case "ファイル名", "file": NormalizeTarget = "file"
        Case Else: NormalizeTarget = "all"
    End Select
End Function

Private Function Applies(i As Long, kind As String) As Boolean
    Applies = (rules(i).Target = kind Or rules(i).Target = "all")
End Function

' Like 演算子で「* と ? を含む置換前」に一致する部分を探して置き換える（Word / PowerPoint / シート名・ファイル名用）
Private Function ReplaceLike(ByVal s As String, pattern As String, repl As String, ByRef n As Long) As String
    Dim i As Long, j As Long, found As Boolean
    If InStr(pattern, "*") = 0 And InStr(pattern, "?") = 0 Then
        If InStr(s, pattern) > 0 Then
            n = n + (Len(s) - Len(Replace(s, pattern, ""))) \ Len(pattern)
            ReplaceLike = Replace(s, pattern, repl)
        Else
            ReplaceLike = s
        End If
        Exit Function
    End If
    i = 1
    Do While i <= Len(s)
        found = False
        For j = i To Len(s)   ' 最短一致
            If Mid$(s, i, j - i + 1) Like pattern Then
                s = Left$(s, i - 1) & repl & Mid$(s, j + 1)
                i = i + Len(repl)
                n = n + 1
                found = True
                Exit For
            End If
        Next
        If Not found Then i = i + 1
    Loop
    ReplaceLike = s
End Function

Private Function NewFileName(f As String) As String
    Dim stem As String, ext As String, i As Long, n As Long
    stem = Left$(f, InStrRev(f, ".") - 1)
    ext = Mid$(f, InStrRev(f, "."))
    For i = 1 To ruleCount
        If Applies(i, "file") Then stem = ReplaceLike(stem, rules(i).Before, rules(i).After, n)
    Next
    If n > 0 Then WriteLog f, "ファイル名", "", f, stem & ext
    NewFileName = stem & ext
End Function

Private Function CheckSheetName(nm As String, wb As Workbook) As String
    Dim ch As Variant, sh As Object
    If Len(nm) = 0 Then CheckSheetName = "空の名前": Exit Function
    If Len(nm) > 31 Then CheckSheetName = "31 文字を超えます": Exit Function
    For Each ch In Array(":", "\", "/", "?", "*", "[", "]")
        If InStr(nm, ch) > 0 Then CheckSheetName = "使えない文字 " & ch: Exit Function
    Next
    If LCase$(nm) = "history" Then CheckSheetName = "予約語": Exit Function
    For Each sh In wb.Sheets
        If LCase$(sh.Name) = LCase$(nm) Then CheckSheetName = "同じ名前のシートがあります": Exit Function
    Next
End Function

Private Sub ProcessExcel(src As String, dest As String)
    Dim wb As Workbook, ws As Worksheet, shp As Shape, i As Long, n As Long, nm As String, problem As String, t As String
    On Error GoTo Failed
    Set wb = Workbooks.Open(src, UpdateLinks:=0, ReadOnly:=True)
    For Each ws In wb.Worksheets
        For i = 1 To ruleCount
            If Applies(i, "text") Then
                ws.Cells.Replace What:=rules(i).Before, Replacement:=rules(i).After, LookAt:=xlPart, MatchCase:=True
                For Each shp In ws.Shapes
                    If shp.HasTextFrame Then
                        If shp.TextFrame2.HasText Then
                            n = 0
                            t = ReplaceLike(shp.TextFrame2.TextRange.Text, rules(i).Before, rules(i).After, n)
                            If n > 0 Then
                                WriteLog wb.Name, "図形", ws.Name & " " & shp.Name, shp.TextFrame2.TextRange.Text, t
                                shp.TextFrame2.TextRange.Text = t
                            End If
                        End If
                    End If
                Next
            End If
        Next
    Next
    ' セルの置換箇所の一覧は Range.Replace では取れないため、ログには件数を出さない（Python 版を推奨）
    For Each ws In wb.Worksheets
        nm = ws.Name
        For i = 1 To ruleCount
            If Applies(i, "sheet") Then nm = ReplaceLike(nm, rules(i).Before, rules(i).After, n)
        Next
        If nm <> ws.Name Then
            problem = CheckSheetName(nm, wb)
            If problem = "" Then
                WriteLog wb.Name, "シート名", "", ws.Name, nm
                ws.Name = nm   ' 数式の参照は Excel が自動で追従する
            Else
                WriteLog wb.Name, "エラー", "シート名", ws.Name & " → " & nm, problem
            End If
        End If
    Next
    wb.SaveCopyAs dest
    wb.Close SaveChanges:=False
    WriteLog wb.Name, "完了", "", "", dest
    Exit Sub
Failed:
    WriteLog Dir(src), "エラー", "", Err.Description, ""
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
End Sub

Private Sub ProcessWord(wd As Object, src As String, dest As String)
    Dim doc As Object, story As Object, rng As Object, i As Long
    On Error GoTo Failed
    Set doc = wd.Documents.Open(src, ReadOnly:=True, Visible:=False)
    For i = 1 To ruleCount
        If Applies(i, "text") Then
            For Each story In doc.StoryRanges   ' 本文・ヘッダー・フッター・テキストボックスなど
                Set rng = story
                Do
                    With rng.Find
                        .ClearFormatting
                        .Replacement.ClearFormatting
                        .Text = WordPattern(rules(i).Before)
                        .Replacement.Text = rules(i).After
                        .MatchWildcards = (InStr(rules(i).Before, "*") + InStr(rules(i).Before, "?") > 0)
                        .Forward = True
                        .Wrap = 0
                        .Execute Replace:=2   ' wdReplaceAll
                    End With
                    Set rng = rng.NextStoryRange
                Loop Until rng Is Nothing
            Next
        End If
    Next
    doc.SaveAs2 dest
    doc.Close SaveChanges:=False
    WriteLog Dir(src), "完了", "", "", dest
    Exit Sub
Failed:
    WriteLog Dir(src), "エラー", "", Err.Description, ""
    If Not doc Is Nothing Then doc.Close SaveChanges:=False
End Sub

' Word のワイルドカードは記号の意味が違うので、* と ? 以外の記号を \ でエスケープする
Private Function WordPattern(p As String) As String
    Dim i As Long, ch As String, out As String
    For i = 1 To Len(p)
        ch = Mid$(p, i, 1)
        If InStr("[]{}()<>@!\", ch) > 0 Then out = out & "\" & ch Else out = out & ch
    Next
    WordPattern = out
End Function

Private Sub ProcessPowerPoint(pp As Object, src As String, dest As String)
    Dim pres As Object, sld As Object, shp As Object, i As Long
    On Error GoTo Failed
    Set pres = pp.Presentations.Open(src, ReadOnly:=True, WithWindow:=False)
    For Each sld In pres.Slides
        For Each shp In sld.Shapes
            ReplaceInShape shp, sld.SlideIndex, Dir(src)
        Next
        For Each shp In sld.NotesPage.Shapes
            ReplaceInShape shp, sld.SlideIndex, Dir(src)
        Next
    Next
    pres.SaveCopyAs dest
    pres.Close
    WriteLog Dir(src), "完了", "", "", dest
    Exit Sub
Failed:
    WriteLog Dir(src), "エラー", "", Err.Description, ""
    If Not pres Is Nothing Then pres.Close
End Sub

Private Sub ReplaceInShape(shp As Object, slideNo As Long, fileName As String)
    Dim i As Long, r As Long, c As Long, n As Long, t As String, child As Object
    If shp.Type = 6 Then   ' グループ
        For Each child In shp.GroupItems
            ReplaceInShape child, slideNo, fileName
        Next
        Exit Sub
    End If
    If shp.HasTable Then
        For r = 1 To shp.Table.Rows.Count
            For c = 1 To shp.Table.Columns.Count
                ReplaceInTextRange shp.Table.Cell(r, c).Shape.TextFrame.TextRange, "スライド " & slideNo & " 表", fileName
            Next
        Next
    ElseIf shp.HasTextFrame Then
        If shp.TextFrame.HasText Then ReplaceInTextRange shp.TextFrame.TextRange, "スライド " & slideNo & " " & shp.Name, fileName
    End If
End Sub

Private Sub ReplaceInTextRange(tr As Object, where As String, fileName As String)
    Dim i As Long, n As Long, found As Object
    For i = 1 To ruleCount
        If Applies(i, "text") Then
            If InStr(rules(i).Before, "*") + InStr(rules(i).Before, "?") = 0 Then
                Set found = tr.Replace(FindWhat:=rules(i).Before, ReplaceWhat:=rules(i).After)   ' 書式を保つ
                Do While Not found Is Nothing
                    WriteLog fileName, "文字列", where, rules(i).Before, rules(i).After
                    Set found = tr.Replace(FindWhat:=rules(i).Before, ReplaceWhat:=rules(i).After)
                Loop
            Else
                n = 0
                Dim t As String
                t = ReplaceLike(tr.Text, rules(i).Before, rules(i).After, n)
                If n > 0 Then
                    WriteLog fileName, "文字列", where, tr.Text, t
                    tr.Text = t   ' ワイルドカードのときは書式が段落単位になる
                End If
            End If
        End If
    Next
End Sub

Private Sub PrepareLog()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("置換ログ")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = "置換ログ"
    End If
    ws.Cells.Clear
    ws.Range("A1:E1").Value = Array("ファイル", "種類", "場所", "置換前", "置換後")
    ws.Range("A1:E1").Font.Bold = True
    logRow = 2
End Sub

Private Sub WriteLog(f As String, kind As String, where As String, before As String, after As String)
    With ThisWorkbook.Worksheets("置換ログ")
        .Cells(logRow, 1).Value = f
        .Cells(logRow, 2).Value = kind
        .Cells(logRow, 3).Value = where
        .Cells(logRow, 4).Value = before
        .Cells(logRow, 5).Value = after
    End With
    logRow = logRow + 1
End Sub
