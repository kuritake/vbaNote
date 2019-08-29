Function makeFile()


Dim file1 As String
Dim file2 As String

Dim buf As String, cData As Variant, n As Long, j As Long
Dim wb As Workbook

file1 = ThisWorkbook.Path & "\" & "ファイル1.csv"
file2 = ThisWorkbook.Path & "\" & "ファイル2.csv"


Set wb = Workbooks.Add  '新規ワークブックを作成
wb.Worksheets.Add after:=Worksheets(Worksheets.Count)
    Open file1 For Input As #1
        Do Until EOF(1)
            Line Input #1, cData
            tmp = Split(Replace(cData, """", ""), ",") 'ダブルクォーテーション除去
            n = n + 1
            wb.Worksheets(1).Range("A" & n).Resize(1, UBound(tmp) + 1).NumberFormat = "@"
            wb.Worksheets(1).Range("A" & n).Resize(1, UBound(tmp) + 1).Value = tmp
                
        Loop
    Close #1
    
    Open file1 For Input As #2
        Do Until EOF(2)
            Line Input #2, cData
            tmp = Split(Replace(cData, """", ""), ",") 
            j = j + 1
            wb.Worksheets(2).Range("A" & j).Resize(1, UBound(tmp) + 1).NumberFormat = "@"
            wb.Worksheets(2).Range("A" & j).Resize(1, UBound(tmp) + 1).Value = tmp
        Loop
    Close #2
    
End Function

