Sub FolderSelect()                                           '
    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
        .Title = "Select directory"
        .InitialFileName = ThisWorkbook.Path
        .Show
            If .SelectedItems.Count > 0 Then
                ThisWorkbook.Worksheets(1) _
                .Range("B1").Value = .SelectedItems(1)
            End If
    End With
    
    With ThisWorkbook.Worksheets(1)
        If Right(.Range(OUTPUT_PATHCELL).Value, 1) <> "\" Then                                         'ディレクトリの末尾に「\」がなかったら付け加える。
            .Range(OUTPUT_PATHCELL).Value = .Range(OUTPUT_PATHCELL).Value & "\"
        End If
    End With
End Sub
