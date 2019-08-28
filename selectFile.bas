Sub Select()
    With Application.FileDialog(msoFileDialogFilePicker)                                'ファイル選択ダイアログを開いてファイルを選択する。
        .AllowMultiSelect = False                                                       'ファイル複数選択可否
        .Title = "Select object."                                                       'ダイアログに表示するタイトル。
        .InitialFileName = ThisWorkbook.Path                                            '初期ディレクトリはマクロブックのパス
        .Filters.Clear                                                                  'フィルターをクリア
        .Filters.Add "テキストファイル", "*.csv"                                        'ファイルを指定
        .Show                                                                           'ダイアログを表示
        If .SelectedItems.Count > 0 Then                                                'ファイルが選択された場合
             ThisWorkbook.Worksheets(1). _
             Range(”A1”).Value = .SelectedItems(1)                  '選択したエクセルのアドレスをボックス内に格納
             Exit Sub
        End If
    End With
End Sub
