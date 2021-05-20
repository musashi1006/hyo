'SelectionChange event procedure on Excel Sheet Code:
'Added how to use below:
'１．Private～End SubまでをExcelのシートコード（例、「アンケート」シートのシートコード）に貼り付けます。
'２．１．のExcelにセルへのリンク参照を横一列に貼り付けるシートを作成しておきます（例、「横一列」シート）。
'３．２．で作成したシートの適当な横並びセル２つに値を入力しておきます（例、"A1"と"B1"セル）。
'４．３．までで下準備完了です。
'５．以下、使用例：
'６．「アンケート」シートのリンク参照を作りたいセルを順番にクリックしていきます。
'７．「横一列」シートのC1セルから右側のセルに、「アンケート」シートの指定セルへの絶対参照リンクが作成されています。
'８．作成し直したければ、「横一列」シートのC1セルから右側のセルを消去して、手順６．を繰り返します。
'９．自己責任での使用をお願いいたします。
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    'シート名などを初期設定します
    Dim sh_name_in, sh_name_out, txt_link, rng0 As String
    sh_name_in = "アンケート"
    sh_name_out = "横一列"
    txt_link = "='" & sh_name_in & "'!"
    rng0 = "A1"

    'クリックしたセルのアドレスを変数aaに格納します
    Dim aa As String
    aa = Target.Item(1).Address(RowAbsolute:=True, ColumnAbsolute:=True)

    '指定するセルへのリンク参照を貼り付けるシートを指定します
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets(sh_name_out)

    '貼り付けシートのセル位置rng0から探した右端の右隣りに指定セルへのリンク参照を作成します
    With sh
        '一応のエラー処理。完璧ではありません
        If .Range(rng0).Next.Value = "" Then
            MsgBox sh_name_out & "シートの" & rng0 & "セルとその右隣のセルに値を入力しておきましょう。"
            .Activate
            .Range(rng0).Next.Activate
        Else
            .Range(rng0).End(xlToRight).Next.Formula = txt_link & aa
        End If
    End With
End Sub
