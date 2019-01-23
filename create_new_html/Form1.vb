Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        ''INIファイルを読み込む。
        Dim dbini As IO.StreamReader
        Dim stCurrentDir As String = System.IO.Directory.GetCurrentDirectory()
        CuDr = stCurrentDir

        If IO.File.Exists(CuDr & "\config.ini") = True Then
            dbini = New IO.StreamReader(CuDr & "\config.ini", System.Text.Encoding.Default)

            For lp1 As Integer = 1 To 1
                Dim tbxn1 As String = "Cf_TextBox" & lp1.ToString

                Dim cs As Control() = Me.Controls.Find(tbxn1, True)
                If cs.Length > 0 Then
                    CType(cs(0), TextBox).Text = dbini.ReadLine
                End If
            Next

            ''メイン作業タブへ
            Me.TabControl1.SelectedTab = TabPage1

            dbini.Close()
            dbini.Dispose()
        Else
            MessageBox.Show("設定ファイルが見つからないか壊れています。", "通知")
            Me.TabControl1.SelectedTab = TabPage2

        End If
    End Sub

    Private Sub Cf_Button1_Click(sender As Object, e As EventArgs) Handles Cf_Button1.Click
        close_save()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        OpenFileDialog1.InitialDirectory = Cf_TextBox1.Text
        OpenFileDialog1.Filter = "Excelファイル(*.xlsx;*.xlsm)|*.xlsx;*.xlsm|すべてのファイル(*.*)|*.*"
        If OpenFileDialog1.ShowDialog() = DialogResult.OK Then
            ''ファイルがあった場合の処理
            opflname = OpenFileDialog1.FileName
            excel_DataRead01(opflname)

        Else
            MessageBox.Show("キャンセルされました", "通知")

        End If
    End Sub

    Private Sub exbut01_Click(sender As Object, e As EventArgs) Handles exbut01.Click
        If opflname = "" Then
            MsgBox("データが読み込まれていません", MsgBoxStyle.Critical And MsgBoxStyle.OkOnly, "警告")
        Else
            targetLabel03.Text = "楽天PC用新商品"
            ex_exbut01()
        End If
    End Sub

    Private Sub exbut02_Click(sender As Object, e As EventArgs) Handles exbut02.Click
        If opflname = "" Then
            MsgBox("データが読み込まれていません", MsgBoxStyle.Critical And MsgBoxStyle.OkOnly, "警告")
        Else
            targetLabel03.Text = "楽天スマホ用新商品"
            ex_exbut02()
        End If
    End Sub

    Private Sub exbut03_Click(sender As Object, e As EventArgs) Handles exbut03.Click
        If opflname = "" Then
            MsgBox("データが読み込まれていません", MsgBoxStyle.Critical And MsgBoxStyle.OkOnly, "警告")
        Else
            targetLabel03.Text = "Yahoo!PC用新商品"
            ex_exbut03()
        End If
    End Sub

    Private Sub exbut04_Click(sender As Object, e As EventArgs) Handles exbut04.Click
        If opflname = "" Then
            MsgBox("データが読み込まれていません", MsgBoxStyle.Critical And MsgBoxStyle.OkOnly, "警告")
        Else
            targetLabel03.Text = "Yahoo!スマホ用新商品"
            ex_exbut04()
        End If
    End Sub

    Private Sub exbut05_Click(sender As Object, e As EventArgs) Handles exbut05.Click
        If opflname = "" Then
            MsgBox("データが読み込まれていません", MsgBoxStyle.Critical And MsgBoxStyle.OkOnly, "警告")
        Else
            targetLabel03.Text = "ポンパレPC用新商品"
            ex_exbut05()
        End If
    End Sub

    Private Sub exbut06_Click(sender As Object, e As EventArgs) Handles exbut06.Click
        If opflname = "" Then
            MsgBox("データが読み込まれていません", MsgBoxStyle.Critical And MsgBoxStyle.OkOnly, "警告")
        Else
            targetLabel03.Text = "ポンパレスマホ用新商品"
            ex_exbut06()
        End If
    End Sub

    Private Sub exbut07_Click(sender As Object, e As EventArgs) Handles exbut07.Click
        If opflname = "" Then
            MsgBox("データが読み込まれていません", MsgBoxStyle.Critical And MsgBoxStyle.OkOnly, "警告")
        Else
            targetLabel03.Text = "WawmaPC用新商品"
            ex_exbut07()
        End If
    End Sub

    Private Sub exbut08_Click(sender As Object, e As EventArgs) Handles exbut08.Click
        If opflname = "" Then
            MsgBox("データが読み込まれていません", MsgBoxStyle.Critical And MsgBoxStyle.OkOnly, "警告")
        Else
            targetLabel03.Text = "生活空間(楽天)PC用新商品"
            ex_exbut08()
        End If
    End Sub

    Private Sub exbut09_Click(sender As Object, e As EventArgs) Handles exbut09.Click
        If opflname = "" Then
            MsgBox("データが読み込まれていません", MsgBoxStyle.Critical And MsgBoxStyle.OkOnly, "警告")
        Else
            targetLabel03.Text = "生活空間(楽天)スマホ用新商品"
            ex_exbut09()
        End If
    End Sub

    Private Sub exbut10_Click(sender As Object, e As EventArgs) Handles exbut10.Click
        If opflname = "" Then
            MsgBox("データが読み込まれていません", MsgBoxStyle.Critical And MsgBoxStyle.OkOnly, "警告")
        Else
            targetLabel03.Text = "生活空間(Yahoo!)PC用新商品"
            ex_exbut10()
        End If
    End Sub

    Private Sub exbut11_Click(sender As Object, e As EventArgs) Handles exbut11.Click
        If opflname = "" Then
            MsgBox("データが読み込まれていません", MsgBoxStyle.Critical And MsgBoxStyle.OkOnly, "警告")
        Else
            targetLabel03.Text = "生活空間(Yahoo!)スマホ用新商品"
            ex_exbut11()
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Clipboard.SetText(Me.TextBox1.Text)
        MsgBox("コピーしました", MsgBoxStyle.Information And MsgBoxStyle.OkOnly, "完了")
    End Sub
End Class
