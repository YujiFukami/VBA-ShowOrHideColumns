Attribute VB_Name = "ModShowColumns"
Option Explicit

'ShowColumns       ・・・元場所：FukamiAddins3.ModCell 
'CheckArray1D      ・・・元場所：FukamiAddins3.ModArray
'CheckArray1DStart1・・・元場所：FukamiAddins3.ModArray



Public Sub ShowColumns(ColumnABCList1D, TargetSheet As Worksheet, Optional ByVal MaxColABC As String, Optional InputShow As Boolean = True)
'指定列のみ表示にする
'20210917

'引数
'ColumnABCList・・・非表示対象の列名の1次元配列 例) ("A","B","C")
'TargetSheet  ・・・対象のシート
'MaxColABC    ・・・非表示切替対象の列範囲の最大列
'InputShow    ・・・指令列を表示ならTrue,非表示ならFalse。デフォルトはTrue
                                                                 
    '引数チェック
    Call CheckArray1D(ColumnABCList1D, "ColumnABCList1D")
    Call CheckArray1DStart1(ColumnABCList1D, "ColumnABCList1D")
    
    If MaxColABC = "" Then '非表示切替対象の列範囲の最大列が指定されていない場合はシートの最終列
        MaxColABC = Split(Cells(1, Columns.Count).Address(True, False), "$")(0) '最終列番号のアルファベット取得
    End If
    
    Dim I          As Long
    Dim N          As Long
    Dim ColumnName As String    '表示対象の列名をまとめたもの
    N = UBound(ColumnABCList1D) '対象の列の個数
    ColumnName = ""             '列名まとめの初期化
    For I = 1 To N
        ColumnName = ColumnName & ColumnABCList1D(I) & ":" & ColumnABCList1D(I)
        If I < N Then '列名の最後だけ","をつけない
            ColumnName = ColumnName & ","
        End If
    Next I
    
    Dim TargetCell As Range                        '対象範囲のセルオブジェクト
    Set TargetCell = TargetSheet.Range(ColumnName) '対象範囲をセルオブジェクトで取得
                                                                                    
    Application.ScreenUpdating = False             '画面更新を解除して高速化
    
    If InputShow = True Then                                 '表示に切り替えるか、非表示に切り替えるか
        TargetSheet.Columns("A:" & MaxColABC).Hidden = True  '全体を非表示
        TargetCell.EntireColumn.Hidden = False               '指令列のみ表示する
    Else
        TargetSheet.Columns("A:" & MaxColABC).Hidden = False '全体を非表示
        TargetCell.EntireColumn.Hidden = True                '指令列のみ表示する
    End If
    
    ActiveWindow.ScrollColumn = 1     '一番左の列にスクロールして表示する
    Application.ScreenUpdating = True '画面更新解除の解除
    
End Sub

Private Sub CheckArray1D(InputArray, Optional HairetuName As String = "配列")
'入力配列が1次元配列かどうかチェックする
'20210804

    Dim Dummy As Integer
    On Error Resume Next
    Dummy = UBound(InputArray, 2)
    On Error GoTo 0
    If Dummy <> 0 Then
        MsgBox (HairetuName & "は1次元配列を入力してください")
        Stop
        Exit Sub '入力元のプロシージャを確認するために抜ける
    End If

End Sub

Private Sub CheckArray1DStart1(InputArray, Optional HairetuName As String = "配列")
'入力1次元配列の開始番号が1かどうかチェックする
'20210804

    If LBound(InputArray, 1) <> 1 Then
        MsgBox (HairetuName & "の開始要素番号は1にしてください")
        Stop
        Exit Sub '入力元のプロシージャを確認するために抜ける
    End If

End Sub


