Attribute VB_Name = "Module1"
Option Explicit
Sub CSV読込クリア()
    With Sheets("読込CSV展開")
        Range(.Cells(1, 1), .Cells(Rows.Count, Columns.Count)).ClearContents
    End With
    With Sheets("MENU")
        .Range("読込最下行").ClearContents
        .Range("読込最右列").ClearContents
    End With
    MsgBox "「読込CSV展開」シートの内容をクリアしました"
End Sub
Sub 一覧整理クリア()
    With Sheets("一覧整理")
        Range(.Cells(1, 1), .Cells(Rows.Count, Columns.Count)).ClearContents
        Range(.Cells(1, 1), .Cells(Rows.Count, Columns.Count)).Borders.LineStyle = False
    End With
    MsgBox "「一覧整理」シートの内容をクリアしました"
End Sub
Sub CSV読込()
    Dim 始時 As Date, 終時 As Date
    Dim ファイル名
    Dim 全文 As String, 文字コード As String, 区切文字 As String, 連続読込モード As String
    Dim 見出行数 As Long, 読込行数 As Long, 読込列数 As Long, 終行 As Long
    Dim 整列()
    ファイル名 = Application.GetOpenFilename(FileFilter:="CSVファイル（*.csv）,*.csv", Title:="CSVファイルの選択")
    If ファイル名 = False Then Exit Sub
    始時 = Timer
    実行中.Show vbModeless
    実行中.Repaint
    With Sheets("MENU")
        文字コード = .Range("文字コード")
        区切文字 = .Range("区切文字")
        見出行数 = .Range("読込見出行数")
        連続読込モード = .Range("連続読込モード")
    End With
    With CreateObject("ADODB.Stream")
        .Charset = 文字コード
        .Open
        .LoadFromFile ファイル名
        全文 = .ReadText
        .Close
    End With
        
    Call CSV解析(全文, 区切文字, 見出行数, 読込行数, 読込列数, 整列) '前3つ代入→後3つ返しのイメージ
        
    With Sheets("読込CSV展開")
        終行 = .Cells(Rows.Count, 1).End(xlUp).Row
        If 終行 = 1 And .Cells(1, 1) = "" Then
            Range(.Cells(1, 1), .Cells(読込行数, 読込列数)) = 整列
            Else:
                Range(.Cells(終行 + 1, 1), .Cells(終行 + 読込行数, 読込列数)) = 整列
                読込行数 = 読込行数 + 終行
        End If
    End With
    With Sheets("MENU")
        .Range("読込最下行") = 読込行数
        .Range("読込最右列") = 読込列数
    End With
    終時 = Timer
    If 連続読込モード = "ON" Then
        If MsgBox("ファイルの読み込みが完了しました。" & vbCrLf & vbCrLf & "処理時間：" & 終時 - 始時 & vbCrLf & vbCrLf & "追加でファイルを読み込みますか？", vbYesNo) = vbYes Then
            Call CSV読込
        End If
    End If
    Unload 実行中
End Sub
Sub CSV解析(全文 As String, 区切文字 As String, 見出行数 As Long, 読込行数 As Long, 読込列数 As Long, 整列 As Variant)
    Dim 改行コード As String
    Dim 仮要素数 As Long, 添字 As Long, 字 As Long, 開始 As Long, 終了 As Long, 行 As Long, 列 As Long
    Dim 囲字カウント As Long 'ダブルクォーテーションの数→区切or改行時点で偶数なら次の項へ移って良い
    Dim 囲判定 As Long '当該項が特殊文字で囲まれているかどうか(0 or 1)
    
    改行コード = vbLf
    全文 = Replace(全文, vbCr, "") '改行コードがvbCrLfだった場合用の補正
    For 字 = 1 To Len(全文)
        Select Case Mid(全文, 字, 1)
            Case 区切文字, 改行コード: 仮要素数 = 仮要素数 + 1 '実際の項数より多くなる場合があるため「仮」
        End Select
    Next
    
    ReDim 終始(1 To 仮要素数, 1 To 3) '1列：項の最初の字の番号、2列：項中の文字数→Mid(全文,1列,2列)で項データ取出
    添字 = 1
    For 字 = 1 To Len(全文) '1文字ずつ解析→終始配列へ記録
        Select Case 字
            Case 1 '文頭処理
                Select Case Mid(全文, 字, 1) '囲判定
                    Case """"
                        囲判定 = 1
                        囲字カウント = 囲字カウント + 1
                    Case Else: 囲判定 = 0
                End Select
                終始(添字, 1) = 字 + 囲判定 '始字位置記録
            Case Else '文中処理：区切文字または改行コード単位で前項締め＆次項開始処理
                Select Case Mid(全文, 字, 1)
                    Case 区切文字, 改行コード
                        If 囲判定 = 0 Or 囲判定 = 1 And 囲字カウント Mod 2 = 0 Then '項区切りと断定可能か判定
                            終始(添字, 2) = 字 - 終始(添字, 1) - 囲判定 '文字数記録
                            If Mid(全文, 字, 1) = 改行コード Then '改行位置記録
                                終始(添字, 3) = 1
                                If 読込列数 = 0 Then 読込列数 = 添字
                            End If
                            Select Case Mid(全文, 字 + 1, 1) '次項の囲判定
                                Case """": 囲判定 = 1
                                Case Else: 囲判定 = 0
                            End Select
                            If 字 < Len(全文) Then
                                添字 = 添字 + 1 '次項へ移動
                                終始(添字, 1) = 字 + 1 + 囲判定 '次項の始字位置記録
                            End If
                        End If
                    Case """": 囲字カウント = 囲字カウント + 1
                End Select
        End Select
    Next
    
    読込行数 = 添字 / 読込列数
    終了 = 読込行数 * 読込列数
    ReDim 整列(1 To 読込行数, 1 To 読込列数) 'CSVを解析&整理した結果を記録
    開始 = 1
    If 見出行数 > 0 Then
        If MsgBox("見出行を含めて読み込みますか？", vbYesNo) = vbNo Then
            開始 = 読込列数 * 見出行数 + 1
            読込行数 = 読込行数 - 見出行数
        End If
    End If
    行 = 1
    列 = 1
    For 添字 = 開始 To 終了
        整列(行, 列) = Trim(Replace(Mid(全文, 終始(添字, 1), 終始(添字, 2)), """""", """")) '連続ダブルクォーテーション整理＆左右端の空白削除→記録
        Select Case 終始(添字, 3)
            Case 1
                行 = 行 + 1
                列 = 1
            Case Else
                列 = 列 + 1
        End Select
    Next
End Sub
Sub 一覧整理()
    Dim 始時 As Date, 終時 As Date
    Dim 読込最下行 As Long, 読込最右列 As Long, 見出行 As Long, 設定終行 As Long, 行 As Long, 列 As Long, 行数 As Long, 列数 As Long, 添字 As Long
    Dim 設定(), データ()
    With Sheets("MENU")
        読込最下行 = .Range("読込最下行")
        読込最右列 = .Range("読込最右列")
        見出行 = .Range("読込見出行数")
        設定終行 = .Cells(Rows.Count, 7).End(xlUp).Row
        Select Case True
            Case 読込最下行 < 1, 読込最右列 < 1
                MsgBox "読込データがありません"
                Exit Sub
            Case 設定終行 < 3
                MsgBox "一覧整理設定をしてください"
                Exit Sub
        End Select
        設定 = Range(.Cells(3, 6), .Cells(設定終行, 14))
        For 行 = 3 To 設定終行
            If 列数 < .Cells(行, 6) Then 列数 = .Cells(行, 6)
        Next
    End With
    
    始時 = Timer
    実行中.Show vbModeless
    実行中.Repaint
    With Sheets("読込CSV展開")
        データ = Range(.Cells(見出行 + 1, 1), .Cells(読込最下行, 読込最右列))
        行数 = 読込最下行 - 見出行
    End With
    With Sheets("一覧整理")
        ReDim 整理(1 To 行数 + 1, 1 To 列数)
        For 添字 = 1 To 設定終行 - 2
            整理(1, 設定(添字, 1)) = 設定(添字, 2)
            For 行 = 1 To 行数
                If 設定(添字, 3) <> "" Then 整理(行 + 1, 設定(添字, 1)) = データ(行, 設定(添字, 3))
                If 設定(添字, 5) <> "" Then 整理(行 + 1, 設定(添字, 1)) = Trim(整理(行 + 1, 設定(添字, 1)) & 設定(添字, 4) & データ(行, 設定(添字, 5)))
                If 設定(添字, 7) <> "" Then 整理(行 + 1, 設定(添字, 1)) = Trim(整理(行 + 1, 設定(添字, 1)) & 設定(添字, 6) & データ(行, 設定(添字, 7)))
                If 設定(添字, 8) <> "" Then 整理(行 + 1, 設定(添字, 1)) = Format(整理(行 + 1, 設定(添字, 1)), 設定(添字, 8))
            Next
            Select Case 設定(添字, 9)
                Case "": .Columns(設定(添字, 1)).NumberFormatLocal = "G/標準"
                Case Else: .Columns(設定(添字, 1)).NumberFormatLocal = 設定(添字, 9)
            End Select
        Next
        Range(.Cells(1, 1), .Cells(Rows.Count, Columns.Count)).ClearContents
        Range(.Cells(1, 1), .Cells(Rows.Count, Columns.Count)).Borders.LineStyle = False
        Range(.Cells(1, 1), .Cells(行数 + 1, 列数)) = 整理
        Range(.Cells(1, 1), .Cells(行数 + 1, 列数)).Borders.LineStyle = True
        For 列 = 1 To 列数
            .Columns(列).AutoFit
        Next
        .Activate
    End With
    終時 = Timer
    MsgBox "一覧整理が完了しました" & vbCrLf & vbCrLf & "処理時間：" & 終時 - 始時
    Unload 実行中
End Sub
