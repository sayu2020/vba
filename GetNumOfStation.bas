Attribute VB_Name = "GetNumOfStation"
Option Explicit
Sub GetNumOfStation()
    
    'ファイル名に他〇局or計〇局という文言が含まれる場合、行数にそれらの数字を足していき局数を算出するマクロ
    
    Dim wbDst As Workbook       '結果出力先ブックのワークブック
    Dim wsDst As Worksheet      '結果出力先ブックのワークシート
    Dim wbOrg As Workbook       '抽出元ブックのワークブック
    Dim wsOrg As Worksheet      '抽出元ブックのワークシート
    Dim NameOrg As Variant      '検索値範囲（ファイル名）
    Dim StartRow As Long        'フィルター後の一番上の行
    Dim EndRow As Long          'フィルター後の一番下の行
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim Num As Long             '他〇局、計○局の丸に入る数値
    Dim Sum As Long             '局数合計
    Dim CellNum As Long         'データ個数
    
    Set wbDst = Workbooks("データ分析まとめ.xlsm")
    Set wbOrg = Workbooks(wbDst.Worksheets("マクロ").Range("D3").Value)
    Set wbOrg = wbOrg.Worksheets("CHK打刻 DB")
    
    With wsOrg
    
        'フィルター適用後の一番上の行番号取得
        StartRow = 0
    
        For i = 5 To .Cells(.Rows.Count, 5).End(xlUp).Row
            If .Cells(i, 5).EntireRow.Hidden = False Then
                If StartRow = 0 Then
                    StartRow = .Cells(i, 5).Row
                    Exit For
                End If
            End If
        Next
        
        'データ個数取得
        For j = 5 To .Cells(.Rows.Count, 5).End(xlUp).Row
            If .Cells(j, 5).EntireRow.Hidden = False Then
                CellNum = CellNum + 1
            End If
        Next
        
        '局数取得
        NameOrg = .Range(.Cells(StartRow, 5), .Cells(.Cells(.Rows.Count, 5).End(xlUp).Row, 5))
        
        Sum = 0
        
        For k = LBound(NameOrg, 1) To UBound(NameOrg, 1)                                                  'フィルター後表示されているすべての行を見ていく
            If .Cells(StartRow + k - 1, 1).EntireRow.Hidden = False Then
                If InStr(NameOrg(k, 1), "他") > 0 Then                                                    'ファイル名に"他"という文言を含んでいたら
                    If IsNumeric(Mid(NameOrg(k, 1), InStr(NameOrg(k, 1), "他") + 1, 1)) Then              '"他"の次の文字が数字なら
                        If IsNumeric(Mid(NameOrg(k, 1), InStr(NameOrg(k, 1), "他") + 1, 2)) Then          '"他"の次の次の文字が数字ならそれらの二桁の数字を合計に足す
                            Num = Mid(NameOrg(k, 1), InStr(NameOrg(k, 1), "他") + 1, 2)
                            Sum = Sum + Num
                        Else                                                                              '"他"の次の次の文字が数字でなければ"他"の次の一桁の数字を合計に足す
                            Num = Mid(NameOrg(k, 1), InStr(NameOrg(k, 1), "他") + 1, 1)
                            Sum = Sum + Num
                            Sum = Sum + Num
                        End If
                    End If
                ElseIf InStr(NameOrg(k, 1), "計") > O Then                                                'ファイル名に"計"という文言を含んでいたら
                    If IsNumeric(Mid(NameOrg(k, 1), InStr(NameOrg(k, 1), "計") + 1, 1)) Then              '"計"の次の文字が数字なら
                        If IsNumeric(Mid(NameOrg(k, 1), InStr(NameOrg(k, 1), "計") + 1, 2)) Then          '"計"の次の次の文字が数字ならそれらの二桁の数字-1を合計に足す
                            Num = Mid(NameOrg(k, 1), InStr(NameOrg(k, 1), "計") + 1, 2)
                            Sum = Sum + Num - 1
                        Else                                                                              '"計"の次の次の文字が数字でなければ"計"の次の一桁の数字-1を合計に足す
                            Num = Mid(NameOrg(k, 1), InStr(NameOrg(k, 1), "計") + 1, 1)
                            Sum = Sum + Num - 1
                        End If
                    End If
                End If
            End If
        Next
    End With
    
    wsDst.Range("M8") = CellNum + Sum
    
    MsgBox "処理完了しました。"
        
    
End Sub
