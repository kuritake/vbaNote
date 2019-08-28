Option Explicit
Dim xlapp As Application

Sub priorConfirmation()

Set xlapp = Application
With xlapp
    .Calculation = xlCalculationManual  '自動計算停止
    .DisplayAlerts = False              ' メッセージ表示停止
    .ScreenUpdating = False             ' 画面描画停止
    .EnableEvents = False               ' イベント動作停止
    .EnableCancelKey = xlErrorHandler   ' Escキーでエラートラップする
    .Cursor = xlWait
    '.Visible = False
End With

With xlapp
        .EnableCancelKey = xlInterrupt                      ' Escキー動作を戻す
        .EnableEvents = True                                ' イベント動作再開
        .ScreenUpdating = True                              ' 画面描画再開
        .Cursor = xlDefault
        .DisplayAlerts = True                               ' メッセージ表示再
        .Application.Calculation = xlCalculationAutomatic   '自動計算再開
        .Visible = True
    End With
Exit Sub
