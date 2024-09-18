VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProcessControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'/******************************************
' * 描画コントロール用
' * 負荷軽減
' ******************************************/
'/***********************
' * 開始
' ***********************/
Sub StartProcess()
    With Application
        .ScreenUpdating = False                 '描画を停止
        .Calculation = xlCalculationManual      '計算を手動
        .EnableEvents = False                   'イベント抑止
    End With
End Sub

'/***********************
' * 終了
' ***********************/
Sub EndProcess()
    With Application
        .ScreenUpdating = True                  '描画を開始
        .Calculation = xlCalculationAutomatic   '計算を自動
        .EnableEvents = True                    'イベント開始
    End With
End Sub

