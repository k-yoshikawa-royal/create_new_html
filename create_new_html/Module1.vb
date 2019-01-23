Imports System.IO
Imports System.Web
Imports MySql.Data.MySqlClient
Imports NPOI.SS.UserModel
Imports NPOI.XSSF.UserModel

Module Module1

    Public CuDr As String       ''exeの動くカレントディレクトリを格納
    Public mysqlCon As New MySqlConnection
    Public sqlCommand As New MySqlCommand
    Public opflname As String
    Public branchno As Integer


    Sub sql_st()
        ''データベースに接続

        Dim Builder = New MySqlConnectionStringBuilder()
        ' データベースに接続するために必要な情報をBuilderに与える。データベース情報はGitに乗せないこと。
        Builder.Server = ""
        Builder.Port =
        Builder.UserID = ""
        Builder.Password = ""
        Builder.Database = ""

        Dim ConStr = Builder.ToString()

        mysqlCon.ConnectionString = ConStr
        mysqlCon.Open()

    End Sub

    Sub sql_cl()
        ' データベースの切断
        mysqlCon.Close()
    End Sub

    Function sql_result_return(ByVal query As String) As DataTable
        ''データセットを返すSELECT系のSQLを処理するコード

        Dim dt As New DataTable()

        Try
            ' 4.データ取得のためのアダプタの設定
            Dim Adapter = New MySqlDataAdapter(query, mysqlCon)

            ' 5.データを取得
            Dim Ds As New DataSet
            Adapter.Fill(dt)

            Return dt
        Catch ex As Exception

            Return dt
        End Try

    End Function

    Function sql_result_no(ByVal query As String)
        ''データセットを返さない、DELETE、UPDATE、INSERT系のSQLを処理するコード

        Try
            sqlCommand.Connection = mysqlCon
            sqlCommand.CommandText = query
            sqlCommand.ExecuteNonQuery()

            Return "Complete"
        Catch ex As Exception

            Return ex.Message
        End Try

    End Function

    Function dt2unepocht(ByVal vbdate As DateTime) As Long
        ''VBで使用出来る日付を入力すると、UNIX エポック秒に変換する
        vbdate = vbdate.ToUniversalTime()

        Dim dt1 As New DateTime(1970, 1, 1, 0, 0, 0, 0)
        Dim elapsedTime As TimeSpan = vbdate - dt1

        Return CType(elapsedTime.TotalSeconds, Long)

    End Function

    Function idc01(ByVal vo As String)

        ''商品管理IDを、URL用の半角に変える変数
        idc01 = StrConv(vo, 2)

    End Function

    Function imgsamp_rakuten2(ByVal ur As String)

        ''新RMSの[商品画像URL]から、先頭の画像URLを抽出する自作関数。リニュアル後用
        '抽出データを、縮小画像用にURL変換も行う

        Dim ret01 As String = ""

        If IsDBNull(ur) = True Then
        Else
            Dim foundIndex As Integer = ur.IndexOf(" ")

            Select Case foundIndex

                Case < 1
                    ret01 = ur

                Case Is <> 0

                    ret01 = Left(ur, foundIndex)

            End Select

        End If

        Return ret01

    End Function

    Function imgsamp_ponparemall(ur)

        ''新RMSの[商品画像URL]から、先頭の画像URLを抽出する自作関数
        '抽出データを、縮小画像用にURL変換も行う


        Dim ret01 As String = ""

        If IsDBNull(ur) = True Then
        Else
            Dim foundIndex As Integer = ur.IndexOf(",")


            Select Case foundIndex

                Case < 1
                    ret01 = ur

                Case Is <> 0

                    ret01 = Left(ur, foundIndex)

            End Select

        End If

        Return ret01

    End Function

    Sub close_save()
        ''設定用ファイルの保存

        Dim dtx1 As String = ""

        For lp1 As Integer = 1 To 1
            Dim tbxn1 As String = "Cf_TextBox" & lp1.ToString

            Dim cs As Control() = Form1.Controls.Find(tbxn1, True)
            If cs.Length > 0 Then
                dtx1 &= CType(cs(0), TextBox).Text
                dtx1 &= vbCrLf
            End If
        Next


        Dim stCurrentDir As String = System.IO.Directory.GetCurrentDirectory()
        CuDr = stCurrentDir

        Dim excsv1 As IO.StreamWriter
        excsv1 = New IO.StreamWriter(CuDr & "\config.ini", False, System.Text.Encoding.GetEncoding("shift_jis"))
        excsv1.Write(dtx1)
        excsv1.Close()
        excsv1.Dispose()

    End Sub

    Sub formtxboxcler()
        ''フォームのデータをクリアする。

        For lpc01 As Integer = 0 To 19

            Dim lpc02 As String = lpc01.ToString
            lpc02 = lpc02.PadLeft(2, "0"c)

            Dim tebx1 As Control() = Form1.Controls.Find("prid" & lpc02, True)
            If tebx1.Length < 0 Then
                MsgBox("IDのボックスがありません", MsgBoxStyle.Critical And MsgBoxStyle.OkOnly, "警告")
            End If

            CType(tebx1(0), TextBox).Text = ""

            Dim tebx2 As Control() = Form1.Controls.Find("name" & lpc02, True)
            If tebx2.Length < 0 Then
                MsgBox("商品名のボックスがありません", MsgBoxStyle.Critical And MsgBoxStyle.OkOnly, "警告")
            End If

            CType(tebx2(0), TextBox).Text = ""


        Next
    End Sub

    Sub excel_DataRead01(ByVal ofn As String)
        ''シートのデータを20件表示させる

        formtxboxcler()

        Dim rfs As FileStream = File.OpenRead(ofn)
        Dim book02 As IWorkbook = New XSSFWorkbook(rfs)
        rfs.Close()
        Dim sheetNo As Integer = book02.GetSheetIndex("新商品表示リスト")

        '番号指定でシートを取得する（番号は０～）
        Dim sheet2 As ISheet = book02.GetSheetAt(sheetNo)

        Dim r1 As Integer = 3
        Dim c1 As Integer = 2

        For lpc01 As Integer = 0 To 19

            Dim lpc02 As String = lpc01.ToString
            lpc02 = lpc02.PadLeft(2, "0"c)

            Dim tebx1 As Control() = Form1.Controls.Find("prid" & lpc02, True)
            If tebx1.Length < 0 Then
                MsgBox("IDのボックスがありません", MsgBoxStyle.Critical And MsgBoxStyle.OkOnly, "警告")
            End If

            Try
                CType(tebx1(0), TextBox).Text = sheet2.GetRow(r1).GetCell(c1).ToString()
            Catch ex As System.NullReferenceException
                CType(tebx1(0), TextBox).Text = ""
            End Try

            c1 += 1

            Dim tebx2 As Control() = Form1.Controls.Find("name" & lpc02, True)
            If tebx2.Length < 0 Then
                MsgBox("商品名のボックスがありません", MsgBoxStyle.Critical And MsgBoxStyle.OkOnly, "警告")
            End If

            Try
                CType(tebx2(0), TextBox).Text = sheet2.GetRow(r1).GetCell(c1).ToString()
            Catch ex As System.NullReferenceException
                CType(tebx2(0), TextBox).Text = ""
            End Try

            c1 += 3

            r1 += 1
            c1 = 2
        Next

    End Sub


    Sub ex_exbut01()

        ''ロイヤル（楽天）用新商品ＨＴＭＬ出力・通常作業者が使用する

        ''データベースと接続
        Call sql_st()

        Dim sql1 As String
        Dim htm As String

        Dim nwtm01 As DateTime = Now()

        Dim tmst As String = nwtm01.Year.ToString.PadLeft(4, "0"c)
        tmst &= "."
        tmst &= nwtm01.Month.ToString.PadLeft(2, "0"c)
        tmst &= "."
        tmst &= nwtm01.Day.ToString.PadLeft(2, "0"c)
        tmst &= " update."

        Dim surl As String
        Dim gurl As String
        Dim snam As String

        ''差し込み用HTML作成
        htm = "                    <!--新商品ここから-->" & vbCrLf
        htm &= "                    <div class=""contents-02"">" & vbCrLf
        htm &= "                      <img src=""common/img/title-new.jpg"">" & vbCrLf
        ''2017/04/17 新商品には更新日時を入れないのでコメントアウト
        htm &= "<!--"
        htm &= "                      <p class=""date"">" & vbCrLf
        htm &= "                        <span>"
        htm &= tmst
        htm &= "</span>" & vbCrLf
        htm &= "                      </p>" & vbCrLf
        htm &= "-->"

        htm &= "                      <div class=""slide-category slide autoplay"">" & vbCrLf


        Dim j As Integer = 0
        Dim i As Integer = 1

        Do
            Dim tebx1 As Control() = Form1.Controls.Find("prid" & j.ToString.PadLeft(2, "0"c), True)
            If tebx1.Length < 0 Then
                MsgBox("IDのボックスがありません", MsgBoxStyle.Critical And MsgBoxStyle.OkOnly, "警告")
            End If

            If CType(tebx1(0), TextBox).Text = "" Then
            Else

                sql1 = "SELECT"
                sql1 &= " `nRms_item_newest`.`商品管理番号（商品URL）` , "
                sql1 &= " `product_ledger`.`商品名`,"
                sql1 &= " `nRms_item_newest`.`商品画像URL` "
                sql1 &= "FROM `nRms_item_newest` INNER JOIN `product_ledger` "
                sql1 &= "ON `nRms_item_newest`.`商品管理番号（商品URL）` = `product_ledger`.`ID` "
                sql1 &= "WHERE `nRms_item_newest`.`商品管理番号（商品URL）` = """
                sql1 &= CType(tebx1(0), TextBox).Text
                sql1 &= """"
                sql1 &= " AND "
                sql1 &= "`商品番号` = '"
                sql1 &= CType(tebx1(0), TextBox).Text
                sql1 &= "' LIMIT 1;"

                Dim dTb1 As DataTable = sql_result_return(sql1)

                If dTb1.Rows.Count = 0 Then
                    ''ロイヤル通販(楽天)にはテキストボックスの商品はない
                    j = j + 1
                    If j > 19 Then
                        Exit Do
                    End If

                Else
                    Dim n As Long = dTb1.Rows.Count
                    For Each DRow As DataRow In dTb1.Rows
                        surl = DRow.Item(0)
                        snam = DRow.Item(1)
                        gurl = DRow.Item(2)

                        htm &= "                        <!--" & i & "個目-->" & vbCrLf
                        htm &= "                        <div>" & vbCrLf
                        htm &= "                          <a href=""https://item.rakuten.co.jp/royal3000/"
                        htm &= surl
                        htm &= "/"">" & vbCrLf
                        htm &= "                            <img src="""
                        htm &= imgsamp_rakuten2(gurl)
                        htm &= """>" & vbCrLf
                        htm &= "                            <p>"
                        htm &= snam
                        htm &= "</p>" & vbCrLf
                        htm &= "                          </a>" & vbCrLf
                        htm &= "                        </div>" & vbCrLf


                        j = j + 1

                        If j > 30 Then
                            Exit Do
                        End If

                        i = i + 1
                        If i > 10 Then
                            Exit Do
                        End If

                    Next
                End If
            End If
        Loop

        htm = htm & "                      </div>" & vbLf
        htm = htm & "                      <!-- .slide-category END-->" & vbLf
        htm = htm & "                      <div class=""more"">" & vbLf
        htm = htm & "                        <a href=""https://item.rakuten.co.jp/royal3000/c/0000000189/""><img src=""common/img/button-more_out.jpg""></a>" & vbLf
        htm = htm & "                      </div>" & vbLf
        htm = htm & "                    </div>" & vbLf
        htm = htm & "                    <!-- .contents-02 END -->" & vbLf
        htm = htm & "                    <!-- 新商品ここまで -->" & vbLf



        Form1.TextBox1.Text = htm

        ''データベースを切断
        Call sql_cl()

    End Sub


    Sub ex_exbut02()

        ''楽天スマホ用新商品ＨＴＭＬ出力・通常作業者が使用する

        ''データベースと接続
        Call sql_st()

        Dim sql1 As String
        Dim htm As String

        Dim nwtm01 As DateTime = Now()

        Dim tmst As String = nwtm01.Year.ToString.PadLeft(4, "0"c)
        tmst &= "."
        tmst &= nwtm01.Month.ToString.PadLeft(2, "0"c)
        tmst &= "."
        tmst &= nwtm01.Day.ToString.PadLeft(2, "0"c)
        tmst &= " update."

        Dim surl As String
        Dim gurl As String
        Dim snam As String

        ''差し込み用HTML作成
        htm = "                <!-- 新商品ここから -->" & vbCrLf
        htm &= "                <div class=""contents-02"">" & vbCrLf
        htm &= "                    <div class=""title-img"">" & vbCrLf
        htm &= "                        <img src=""common/img/sptitle-new.jpg"">" & vbCrLf
        htm &= "                        <p>新商品をチェック</p>" & vbCrLf
        htm &= "                    </div>" & vbCrLf
        htm &= "                    <div class=""slide_list"">" & vbCrLf

        '更新日時が必要になったらコメントアウト2017/6/13
        'htm &= "					<p class=""date"">" & vbCrLf
        'htm &= "						<span>"
        'htm &= tmst
        'htm &= "</span>" & vbCrLf
        'htm &= "					</p>" & vbCrLf


        Dim j As Integer = 0
        Dim i As Integer = 1

        Do
            Dim tebx1 As Control() = Form1.Controls.Find("prid" & j.ToString.PadLeft(2, "0"c), True)
            If tebx1.Length < 0 Then
                MsgBox("IDのボックスがありません", MsgBoxStyle.Critical And MsgBoxStyle.OkOnly, "警告")
            End If

            If CType(tebx1(0), TextBox).Text = "" Then
            Else

                sql1 = "Select"
                sql1 &= " `nRms_item_newest`.`商品管理番号（商品URL）` , "
                sql1 &= " `product_ledger`.`商品名`,"
                sql1 &= " `nRms_item_newest`.`商品画像URL` "
                sql1 &= "FROM `nRms_item_newest` INNER JOIN `product_ledger` "
                sql1 &= "On `nRms_item_newest`.`商品管理番号（商品URL）` = `product_ledger`.`ID` "
                sql1 &= "WHERE `nRms_item_newest`.`商品管理番号（商品URL）` = '"
                sql1 &= CType(tebx1(0), TextBox).Text
                sql1 &= "' AND "
                sql1 &= "`nRms_item_newest`.`商品番号` = '"
                sql1 &= CType(tebx1(0), TextBox).Text
                sql1 &= "' LIMIT 1;"

                Dim dTb1 As DataTable = sql_result_return(sql1)

                If dTb1.Rows.Count = 0 Then
                    ''ロイヤル通販(楽天)にはテキストボックスの商品はない
                    j = j + 1
                    If j > 19 Then
                        Exit Do
                    End If

                Else
                    Dim n As Long = dTb1.Rows.Count
                    For Each DRow As DataRow In dTb1.Rows
                        surl = DRow.Item(0)
                        snam = DRow.Item(1)
                        gurl = DRow.Item(2)

                        htm &= "                        <!--" & i & "個目-->" & vbCrLf
                        htm &= "                        <div>" & vbCrLf
                        htm &= "                            <a href=""https://item.rakuten.co.jp/royal3000/"
                        htm &= surl
                        htm &= "/"">" & vbCrLf
                        htm &= "                                <img src="""
                        htm &= imgsamp_rakuten2(gurl)
                        htm &= """>" & vbCrLf
                        htm &= "                                <p>"
                        htm &= snam
                        htm &= "</p>" & vbCrLf
                        htm &= "                            </a>" & vbCrLf
                        htm &= "                        </div>" & vbCrLf


                        j = j + 1

                        If j > 30 Then
                            Exit Do
                        End If

                        i = i + 1
                        If i > 10 Then
                            Exit Do
                        End If

                    Next
                End If
            End If
        Loop

        htm &= "                    </div>" & vbCrLf
        htm &= "                    <!-- .slide_list END-->" & vbCrLf
        htm &= "                    <div class=""more"">" & vbCrLf
        htm &= "                        <a href=""https://item.rakuten.co.jp/royal3000/c/0000000189/"">もっと見る</a>" & vbCrLf
        htm &= "                    </div>" & vbCrLf
        htm &= "                </div>" & vbCrLf
        htm &= "                <!-- .contents-02 END -->" & vbCrLf
        htm &= "                <!-- 新商品ここまで -->" & vbCrLf



        Form1.TextBox1.Text = htm

        ''データベースを切断
        Call sql_cl()

    End Sub

    Sub ex_exbut03()

        ''ヤフーＰＣ用新商品ＨＴＭＬ出力・通常作業者が使用する

        ''データベースと接続
        Call sql_st()

        Dim sql1 As String
        Dim htm As String

        Dim nwtm01 As DateTime = Now()

        Dim tmst As String = nwtm01.Year.ToString.PadLeft(4, "0"c)
        tmst &= "."
        tmst &= nwtm01.Month.ToString.PadLeft(2, "0"c)
        tmst &= "."
        tmst &= nwtm01.Day.ToString.PadLeft(2, "0"c)
        tmst &= " update."

        Dim surl As String
        Dim snam As String

        ''差し込み用HTML作成
        htm = "                    <!-- 新商品ここから -->" & vbCrLf
        htm &= "                    <div class=""contents-02"">" & vbCrLf
        htm &= "                        <img src=""common/img/title-new.jpg"">" & vbCrLf
        htm &= "                        <div class=""slide-category slide autoplay"">" & vbCrLf

        '更新日時が必要になったらコメントアウト2017/6/13
        'htm &= "					<p class=""date"">" & vbCrLf
        'htm &= "						<span>"
        'htm &= tmst
        'htm &= "</span>" & vbCrLf
        'htm &= "					</p>" & vbCrLf


        Dim j As Integer = 0
        Dim i As Integer = 1

        Do
            Dim tebx1 As Control() = Form1.Controls.Find("prid" & j.ToString.PadLeft(2, "0"c), True)
            If tebx1.Length < 0 Then
                MsgBox("IDのボックスがありません", MsgBoxStyle.Critical And MsgBoxStyle.OkOnly, "警告")
            End If

            If CType(tebx1(0), TextBox).Text = "" Then
            Else

                sql1 = "SELECT"
                sql1 &= " `shopping_yahoo_data_newest`.`code` , "
                sql1 &= " `product_ledger`.`商品名`,"
                sql1 &= " `shopping_yahoo_data_newest`.`original-price`,"
                sql1 &= " `shopping_yahoo_data_newest`.`price`"
                sql1 &= " FROM `shopping_yahoo_data_newest` INNER JOIN `product_ledger` "
                sql1 &= "ON `shopping_yahoo_data_newest`.`code` = `product_ledger`.`ID` "
                sql1 &= " WHERE"
                sql1 &= " `code` = "
                sql1 &= "'"
                sql1 &= CType(tebx1(0), TextBox).Text
                sql1 &= "' LIMIT 1;"


                Dim dTb1 As DataTable = sql_result_return(sql1)

                If dTb1.Rows.Count = 0 Then
                    ''ロイヤル通販(楽天)にはテキストボックスの商品はない
                    j = j + 1
                    If j > 19 Then
                        Exit Do
                    End If

                Else
                    Dim n As Long = dTb1.Rows.Count
                    For Each DRow As DataRow In dTb1.Rows
                        surl = DRow.Item(0)
                        surl = idc01(surl)
                        snam = DRow.Item(1)

                        htm &= "                            <!--" & i & "個目-->" & vbCrLf
                        htm &= "                            <div>" & vbCrLf
                        htm &= "                                <a href=""https://store.shopping.yahoo.co.jp/royal3000/"
                        htm &= surl
                        htm &= ".html"">" & vbCrLf
                        htm &= "                                    <img src=""https://item-shopping.c.yimg.jp/i/j/royal3000_"
                        htm &= surl
                        htm &= """>" & vbCrLf
                        htm &= "                                    <p>"
                        htm &= snam
                        htm &= "</p>" & vbCrLf
                        htm &= "                                </a>" & vbCrLf
                        htm &= "                            </div>" & vbCrLf

                        j = j + 1

                        If j > 30 Then
                            Exit Do
                        End If

                        i = i + 1
                        If i > 10 Then
                            Exit Do
                        End If

                    Next
                End If
            End If
        Loop

        htm &= "                        </div>" & vbCrLf
        htm &= "                        <!-- .slide-category END-->" & vbCrLf
        htm &= "                        <div class=""more"">" & vbCrLf
        htm &= "                        <a href=""https://store.shopping.yahoo.co.jp/royal3000/bfb7c3e5be.html""><img src=""common/img/button-more_out.jpg""></a>" & vbCrLf
        htm &= "                        </div>" & vbCrLf
        htm &= "                    </div>" & vbCrLf
        htm &= "                    <!-- .contents-02 END -->" & vbCrLf
        htm &= "                    <!-- 新商品ここまで -->" & vbCrLf


        Form1.TextBox1.Text = htm

        ''データベースを切断
        Call sql_cl()

    End Sub

    Sub ex_exbut04()

        ''ヤフースマホ用新商品ＨＴＭＬ出力・通常作業者が使用する

        ''データベースと接続
        Call sql_st()

        Dim sql1 As String
        Dim htm As String

        Dim nwtm01 As DateTime = Now()

        Dim tmst As String = nwtm01.Year.ToString.PadLeft(4, "0"c)
        tmst &= "."
        tmst &= nwtm01.Month.ToString.PadLeft(2, "0"c)
        tmst &= "."
        tmst &= nwtm01.Day.ToString.PadLeft(2, "0"c)
        tmst &= " update."

        Dim surl As String
        Dim snam As String

        ''差し込み用HTML作成

        htm = "                <!-- 新商品ここから -->" & vbCrLf
        htm &= "                <div class=""contents-02"">" & vbCrLf
        htm &= "                    <div class=""title-img"">" & vbCrLf
        htm &= "                        <img src=""common/img/sptitle-new.jpg"">" & vbCrLf
        htm &= "                        <p>新商品をチェック</p>" & vbCrLf
        htm &= "                    </div>" & vbCrLf
        htm &= "                    <div class=""slide_list"">" & vbCrLf

        '更新日時が必要になったらコメントアウト2017/6/13
        'htm &= "					<p class=""date"">" & vbCrLf
        'htm &= "						<span>"
        'htm &= tmst
        'htm &= "</span>" & vbCrLf
        'htm &= "					</p>" & vbCrLf

        Dim j As Integer = 0
        Dim i As Integer = 1

        Do
            Dim tebx1 As Control() = Form1.Controls.Find("prid" & j.ToString.PadLeft(2, "0"c), True)
            If tebx1.Length < 0 Then
                MsgBox("IDのボックスがありません", MsgBoxStyle.Critical And MsgBoxStyle.OkOnly, "警告")
            End If

            If CType(tebx1(0), TextBox).Text = "" Then
            Else

                sql1 = "SELECT"
                sql1 &= " `shopping_yahoo_data_newest`.`code` , "
                sql1 &= " `product_ledger`.`商品名`,"
                sql1 &= " `shopping_yahoo_data_newest`.`original-price`,"
                sql1 &= " `shopping_yahoo_data_newest`.`price`"
                sql1 &= " FROM `shopping_yahoo_data_newest` INNER JOIN `product_ledger` "
                sql1 &= "ON `shopping_yahoo_data_newest`.`code` = `product_ledger`.`ID` "
                sql1 &= " WHERE"
                sql1 &= " `code` = "
                sql1 &= "'"
                sql1 &= CType(tebx1(0), TextBox).Text
                sql1 &= "' LIMIT 1;"

                Dim dTb1 As DataTable = sql_result_return(sql1)

                If dTb1.Rows.Count = 0 Then
                    ''ロイヤル通販(楽天)にはテキストボックスの商品はない
                    j = j + 1
                    If j > 19 Then
                        Exit Do
                    End If

                Else
                    Dim n As Long = dTb1.Rows.Count
                    For Each DRow As DataRow In dTb1.Rows
                        surl = DRow.Item(0)
                        surl = idc01(surl)
                        snam = DRow.Item(1)


                        htm &= "                        <!--"
                        htm &= i
                        htm &= "個目-->" & vbCrLf
                        htm &= "                        <div>" & vbCrLf
                        htm &= "                            <a href=""https://store.shopping.yahoo.co.jp/royal3000/"
                        htm &= surl
                        htm &= ".html"">" & vbCrLf
                        htm &= "                                <img src=""https://item-shopping.c.yimg.jp/i/j/royal3000_"
                        htm &= surl
                        htm &= """>" & vbCrLf
                        htm &= "                                <p>"
                        htm &= snam
                        htm &= "</p>" & vbCrLf
                        htm &= "                            </a>" & vbCrLf
                        htm &= "                        </div>" & vbCrLf

                        j = j + 1

                        If j > 30 Then
                            Exit Do
                        End If

                        i = i + 1
                        If i > 10 Then
                            Exit Do
                        End If

                    Next
                End If
            End If
        Loop

        htm &= "                    </div>" & vbCrLf
        htm &= "                    <!-- .slide_list END-->" & vbCrLf
        htm &= "                    <div class=""more"">" & vbCrLf
        htm &= "                        <a href=""https://store.shopping.yahoo.co.jp/royal3000/bfb7c3e5be.html"">もっと見る</a>" & vbCrLf
        htm &= "                    </div>" & vbCrLf
        htm &= "                </div>" & vbCrLf
        htm &= "                <!-- .contents-02 END -->" & vbCrLf
        htm &= "                <!-- 新商品ここまで -->" & vbCrLf

        Form1.TextBox1.Text = htm

        ''データベースを切断
        Call sql_cl()

    End Sub

    Sub ex_exbut05()

        ''ポンパレＰＣ新商品ＨＴＭＬ出力・通常作業者が使用する

        ''データベースと接続
        Call sql_st()

        Dim sql1 As String
        Dim htm As String

        Dim nwtm01 As DateTime = Now()

        Dim tmst As String = nwtm01.Year.ToString.PadLeft(4, "0"c)
        tmst &= "."
        tmst &= nwtm01.Month.ToString.PadLeft(2, "0"c)
        tmst &= "."
        tmst &= nwtm01.Day.ToString.PadLeft(2, "0"c)
        tmst &= " update."

        ''差し込み用HTML作成

        htm = "                    <!--新商品ここから-->" & vbCrLf
        htm &= "                    <div class=""contents-02"">" & vbCrLf
        htm &= "                        <img src=""common/img/title-new.jpg"">" & vbCrLf
        htm &= "                        <div class=""slide-category slide autoplay"">" & vbCrLf

        '更新日時が必要になったらコメントアウト2017/6/13
        'htm &= "						<p Class=""date"">" & vbCrLf
        'htm &= "							<span>"
        'htm &= tmst
        'htm &= "</span>" & vbCrLf
        'htm &= "						</p>" & vbCrLf


        Dim j As Integer = 0
        Dim i As Integer = 1

        Do
            Dim tebx1 As Control() = Form1.Controls.Find("prid" & j.ToString.PadLeft(2, "0"c), True)
            If tebx1.Length < 0 Then
                MsgBox("IDのボックスがありません", MsgBoxStyle.Critical And MsgBoxStyle.OkOnly, "警告")
            End If

            If CType(tebx1(0), TextBox).Text = "" Then
            Else

                sql1 = "SELECT"
                sql1 &= " `ponparemall_item_newest`.`商品管理ID（商品URL）` , "
                sql1 &= " `ponparemall_item_newest`.`商品ID`,"
                sql1 &= " `product_ledger`.`商品名`,"
                sql1 &= " `ponparemall_item_newest`.`販売価格`,"
                sql1 &= " `ponparemall_item_newest`.`表示価格`,"
                sql1 &= " `ponparemall_item_newest`.`商品画像URL`"
                sql1 &= " FROM `ponparemall_item_newest` INNER JOIN `product_ledger` "
                sql1 &= "ON `ponparemall_item_newest`.`商品管理ID（商品URL）` = `product_ledger`.`ID` "
                sql1 &= " WHERE `商品管理ID（商品URL）` = "
                sql1 &= "'"
                sql1 &= CType(tebx1(0), TextBox).Text
                sql1 &= "'"
                sql1 &= " AND"
                sql1 &= " `商品ID` = "
                sql1 &= "'"
                sql1 &= CType(tebx1(0), TextBox).Text
                sql1 &= "' LIMIT 1;"


                Dim dTb1 As DataTable = sql_result_return(sql1)

                If dTb1.Rows.Count = 0 Then
                    ''ロイヤル通販(楽天)にはテキストボックスの商品はない
                    j = j + 1
                    If j > 19 Then
                        Exit Do
                    End If

                Else
                    Dim n As Long = dTb1.Rows.Count
                    For Each DRow As DataRow In dTb1.Rows
                        Dim surl As String = DRow.Item(0)
                        Dim gurl As String = DRow.Item(5)
                        Dim snam As String = DRow.Item(2)


                        htm &= "                            <!--"
                        htm &= i
                        htm &= "個目-->" & vbCrLf
                        htm &= "                            <div>" & vbCrLf
                        htm &= "                                <a href=""https://store.ponparemall.com/royal3000/goods/"
                        htm &= surl
                        htm &= "/"">" & vbCrLf
                        htm &= "                                    <img src="""
                        htm &= imgsamp_ponparemall(gurl)
                        htm &= """>" & vbCrLf
                        htm &= "                                    <p>"
                        htm &= snam
                        htm &= "</p>" & vbCrLf
                        htm &= "                                </a>" & vbCrLf
                        htm &= "                            </div>" & vbCrLf

                        j = j + 1

                        If j > 30 Then
                            Exit Do
                        End If

                        i = i + 1
                        If i > 10 Then
                            Exit Do
                        End If

                    Next
                End If
            End If
        Loop

        htm &= "                        </div>" & vbCrLf
        htm &= "                        <!-- .slide-category END-->" & vbCrLf
        htm &= "                        <div class=""more"">" & vbCrLf
        htm &= "                        <a href=""https://store.ponparemall.com/royal3000/category/0000000001/""><img src=""common/img/button-more_out.jpg""></a>" & vbCrLf
        htm &= "                        </div>" & vbCrLf
        htm &= "                    </div>" & vbCrLf
        htm &= "                    <!-- .contents-02 END -->" & vbCrLf
        htm &= "                    <!-- 新商品ここまで -->" & vbCrLf



        Form1.TextBox1.Text = htm

        ''データベースを切断
        Call sql_cl()

    End Sub

    Sub ex_exbut06()

        ''ポンパレスマホ新商品ＨＴＭＬ出力・通常作業者が使用する

        ''データベースと接続
        Call sql_st()

        Dim sql1 As String
        Dim htm As String

        Dim nwtm01 As DateTime = Now()

        Dim tmst As String = nwtm01.Year.ToString.PadLeft(4, "0"c)
        tmst &= "."
        tmst &= nwtm01.Month.ToString.PadLeft(2, "0"c)
        tmst &= "."
        tmst &= nwtm01.Day.ToString.PadLeft(2, "0"c)
        tmst &= " update."

        ''差し込み用HTML作成
        htm = "                <!-- 新商品ここから -->" & vbCrLf
        htm &= "                <div class=""contents-02"">" & vbCrLf
        htm &= "                    <div class=""title-img"">" & vbCrLf
        htm &= "                        <img src=""common/img/sptitle-new.jpg"">" & vbCrLf
        htm &= "                        <p>新商品をチェック</p>" & vbCrLf
        htm &= "                    </div>" & vbCrLf
        htm &= "                    <div class=""slide_list"">" & vbCrLf

        '更新日時が必要になったらコメントアウト2017/6/13
        'htm &= "					<p class=""date"">" & vbCrLf
        'htm &= "							<span>"
        'htm &= tmst
        'htm &= "</span>" & vbCrLf
        'htm &= "					</p>" & vbCrLf



        Dim j As Integer = 0
        Dim i As Integer = 1

        Do
            Dim tebx1 As Control() = Form1.Controls.Find("prid" & j.ToString.PadLeft(2, "0"c), True)
            If tebx1.Length < 0 Then
                MsgBox("IDのボックスがありません", MsgBoxStyle.Critical And MsgBoxStyle.OkOnly, "警告")
            End If

            If CType(tebx1(0), TextBox).Text = "" Then
            Else

                sql1 = "SELECT"
                sql1 &= " `ponparemall_item_newest`.`商品管理ID（商品URL）` , "
                sql1 &= " `ponparemall_item_newest`.`商品ID`,"
                sql1 &= " `product_ledger`.`商品名`,"
                sql1 &= " `ponparemall_item_newest`.`販売価格`,"
                sql1 &= " `ponparemall_item_newest`.`表示価格`,"
                sql1 &= " `ponparemall_item_newest`.`商品画像URL`"
                sql1 &= " FROM `ponparemall_item_newest` INNER JOIN `product_ledger` "
                sql1 &= "ON `ponparemall_item_newest`.`商品管理ID（商品URL）` = `product_ledger`.`ID` "
                sql1 &= " WHERE `商品管理ID（商品URL）` = "
                sql1 &= "'"
                sql1 &= CType(tebx1(0), TextBox).Text
                sql1 &= "'"
                sql1 &= " AND"
                sql1 &= " `商品ID` = "
                sql1 &= "'"
                sql1 &= CType(tebx1(0), TextBox).Text
                sql1 &= "' LIMIT 1;"


                Dim dTb1 As DataTable = sql_result_return(sql1)

                If dTb1.Rows.Count = 0 Then
                    ''ロイヤル通販(楽天)にはテキストボックスの商品はない
                    j = j + 1
                    If j > 19 Then
                        Exit Do
                    End If

                Else
                    Dim n As Long = dTb1.Rows.Count
                    For Each DRow As DataRow In dTb1.Rows
                        Dim surl As String = DRow.Item(0)
                        Dim gurl As String = DRow.Item(5)
                        Dim snam As String = DRow.Item(2)

                        htm &= "                        <!--"
                        htm &= i
                        htm &= "個目-->" & vbCrLf
                        htm &= "                        <div>" & vbCrLf
                        htm &= "                            <a href=""https://store.ponparemall.com/royal3000/goods/"
                        htm &= surl
                        htm &= "/"">" & vbCrLf
                        htm &= "                                <img src="""
                        htm &= imgsamp_ponparemall(gurl)
                        htm &= """>" & vbCrLf
                        htm &= "                                <p>"
                        htm &= snam
                        htm &= "</p>" & vbCrLf
                        htm &= "                            </a>" & vbCrLf
                        htm &= "                        </div>" & vbCrLf

                        j = j + 1

                        If j > 30 Then
                            Exit Do
                        End If

                        i = i + 1
                        If i > 10 Then
                            Exit Do
                        End If

                    Next
                End If
            End If
        Loop

        htm &= "                    </div>" & vbCrLf
        htm &= "                    <!-- .slide_list END-->" & vbCrLf
        htm &= "                    <div class=""more"">" & vbCrLf
        htm &= "                        <a href=""https://store.ponparemall.com/royal3000/category/0000000001/"">もっと見る</a>" & vbCrLf
        htm &= "                    </div>" & vbCrLf
        htm &= "                </div>" & vbCrLf
        htm &= "                <!-- .contents-02 END -->" & vbCrLf
        htm &= "                <!-- 新商品ここまで -->" & vbCrLf



        Form1.TextBox1.Text = htm

        ''データベースを切断
        Call sql_cl()

    End Sub

    Sub ex_exbut07()

        ''WawmaＰＣ／スマホ兼用　新商品ＨＴＭＬ出力・通常作業者が使用する

        ''データベースと接続
        Call sql_st()

        Dim sql1 As String
        Dim htm As String

        Dim j As Integer
        Dim i As Integer = 1

        ''ロット番号出力用
        htm = ""

        Do
            Dim tebx1 As Control() = Form1.Controls.Find("prid" & j.ToString.PadLeft(2, "0"c), True)
            If tebx1.Length < 0 Then
                MsgBox("IDのボックスがありません", MsgBoxStyle.Critical And MsgBoxStyle.OkOnly, "警告")
            End If

            If CType(tebx1(0), TextBox).Text = "" Then
            Else

                sql1 = "SELECT"
                sql1 &= " `wawmanager_data_newest`.`lotNumber`"
                sql1 &= ",`wawmanager_data_newest`.`itemCode`"
                sql1 &= ",`product_ledger`.`商品名`"
                sql1 &= " FROM `wawmanager_data_newest` INNER JOIN `product_ledger`"
                sql1 &= " ON `wawmanager_data_newest`.`itemCode` = `product_ledger`.`ID`"
                sql1 &= " WHERE `wawmanager_data_newest`.`itemCode` = '"
                sql1 &= CType(tebx1(0), TextBox).Text
                sql1 &= "';"

                Dim dTb1 As DataTable = sql_result_return(sql1)

                If dTb1.Rows.Count = 0 Then
                    ''ロイヤル通販(楽天)にはテキストボックスの商品はない
                    j = j + 1
                    If j > 19 Then
                        Exit Do
                    End If

                Else
                    Dim n As Long = dTb1.Rows.Count
                    For Each DRow As DataRow In dTb1.Rows
                        Dim lot As String = DRow.Item(0)
                        Dim cod As String = DRow.Item(1)
                        Dim snam As String = DRow.Item(2)

                        htm &= i
                        htm &= "個目："
                        htm &= snam
                        htm &= "("
                        htm &= cod
                        htm &= ")"
                        htm &= "："
                        htm &= lot
                        htm &= vbCrLf
                        htm &= vbCrLf

                        j = j + 1

                        If j > 20 Then
                            Exit Do
                        End If

                        i = i + 1
                        If i > 10 Then
                            Exit Do
                        End If

                    Next
                End If
            End If
        Loop


        Form1.TextBox1.Text = htm

        ''データベースを切断
        Call sql_cl()

    End Sub

    Sub ex_exbut08()

        ''生活空間（楽天）ＰＣ用新商品ＨＴＭＬ出力・通常作業者が使用する

        ''データベースと接続
        Call sql_st()

        Dim sql1 As String
        Dim htm As String

        ''差し込み用HTML作成

        htm = "<!--新商品ここから-->" & vbCrLf
        htm &= "<div class=""contents-03"">" & vbCrLf
        htm &= "  <h2 class=""title-cmn"">NEW ARRIVAL</h2>" & vbCrLf
        htm &= "  <div class=""slide-category slide autoplay"">" & vbCrLf


        Dim j As Integer = 0
        Dim i As Integer = 1

        Do
            Dim tebx1 As Control() = Form1.Controls.Find("prid" & j.ToString.PadLeft(2, "0"c), True)
            If tebx1.Length < 0 Then
                MsgBox("IDのボックスがありません", MsgBoxStyle.Critical And MsgBoxStyle.OkOnly, "警告")
            End If

            If CType(tebx1(0), TextBox).Text = "" Then
            Else

                sql1 = "SELECT"
                sql1 &= " `商品管理番号（商品URL）` , "
                sql1 &= " `商品名`,"
                sql1 &= " `商品画像URL` "
                sql1 &= "FROM `nRms_seikatsukukan_item_newest` "
                sql1 &= "WHERE `商品管理番号（商品URL）` = "
                sql1 &= "'"
                sql1 &= CType(tebx1(0), TextBox).Text
                sql1 &= "'"
                sql1 &= " AND "
                sql1 &= "`商品番号` = "
                sql1 &= "'"
                sql1 &= CType(tebx1(0), TextBox).Text
                sql1 &= "'"
                sql1 &= " AND "
                sql1 &= "`倉庫指定` = 0"
                sql1 &= " LIMIT 1;"


                Dim dTb1 As DataTable = sql_result_return(sql1)

                If dTb1.Rows.Count = 0 Then
                    ''ロイヤル通販(楽天)にはテキストボックスの商品はない
                    j = j + 1
                    If j > 19 Then
                        Exit Do
                    End If

                Else
                    Dim n As Long = dTb1.Rows.Count
                    For Each DRow As DataRow In dTb1.Rows
                        Dim surl As String = DRow.Item(0)
                        Dim gurl As String = DRow.Item(2)

                        htm &= "<!--"
                        htm &= i
                        htm &= "個目-->" & vbCrLf

                        htm &= "    <div>" & vbCrLf
                        htm &= "      <div class=""inner"">" & vbCrLf
                        htm &= "        <a href=""https://item.rakuten.co.jp/seikatsukukan/"
                        htm &= surl
                        htm &= "/"">" & vbCrLf

                        htm &= "          <img src="""
                        htm &= imgsamp_rakuten2(gurl)
                        htm &= """ alt="""">" & vbCrLf
                        htm &= "        </a>" & vbCrLf
                        htm &= "      </div>" & vbCrLf
                        htm &= "    </div>" & vbCrLf

                        j = j + 1

                        If j > 19 Then
                            Exit Do
                        End If

                        i = i + 1
                        If i > 10 Then
                            Exit Do
                        End If

                    Next
                End If
            End If
        Loop

        htm &= "  </div>" & vbCrLf
        htm &= "  <!--/slide-category--> " & vbCrLf
        htm &= "</div>" & vbCrLf
        htm &= "<!--/contents-02--> " & vbCrLf
        htm &= "<!--新商品ここまで-->" & vbCrLf


        Form1.TextBox1.Text = htm

        ''データベースを切断
        Call sql_cl()

    End Sub

    Sub ex_exbut09()

        ''生活空間（楽天）スマホ用新商品ＨＴＭＬ出力・通常作業者が使用する

        ''データベースと接続
        Call sql_st()

        Dim sql1 As String
        Dim htm As String


        ''差し込み用HTML作成
        htm = "<!--新商品ここから-->" & vbCrLf
        htm &= "<section class=""section mb40"">" & vbCrLf
        htm &= "    <h3 class=""section_title"">NEW ARRIVAL</h3>" & vbCrLf
        htm &= "    <div class=""slide_list"">" & vbCrLf


        Dim j As Integer = 0
        Dim i As Integer = 1

        Do
            Dim tebx1 As Control() = Form1.Controls.Find("prid" & j.ToString.PadLeft(2, "0"c), True)
            If tebx1.Length < 0 Then
                MsgBox("IDのボックスがありません", MsgBoxStyle.Critical And MsgBoxStyle.OkOnly, "警告")
            End If

            If CType(tebx1(0), TextBox).Text = "" Then
            Else

                sql1 = "SELECT"
                sql1 &= " `商品管理番号（商品URL）` , "
                sql1 &= " `商品名`,"
                sql1 &= " `商品画像URL` "
                sql1 &= "FROM `nRms_seikatsukukan_item_newest` "
                sql1 &= "WHERE `商品管理番号（商品URL）` = "
                sql1 &= "'"
                sql1 &= CType(tebx1(0), TextBox).Text
                sql1 &= "'"
                sql1 &= " AND "
                sql1 &= "`商品番号` = "
                sql1 &= "'"
                sql1 &= CType(tebx1(0), TextBox).Text
                sql1 &= "'"
                sql1 &= " AND "
                sql1 &= "`倉庫指定` = 0"
                sql1 &= " LIMIT 1;"

                Dim dTb1 As DataTable = sql_result_return(sql1)

                If dTb1.Rows.Count = 0 Then
                    ''ロイヤル通販(楽天)にはテキストボックスの商品はない
                    j = j + 1
                    If j > 19 Then
                        Exit Do
                    End If

                Else
                    Dim n As Long = dTb1.Rows.Count
                    For Each DRow As DataRow In dTb1.Rows
                        Dim surl As String = DRow.Item(0)
                        Dim gurl As String = DRow.Item(2)


                        htm &= "<!--"
                        htm &= i
                        htm &= "個目-->" & vbCrLf

                        htm &= "        <a href=""https://item.rakuten.co.jp/seikatsukukan/"
                        htm &= surl
                        htm &= "/"" style=""background-image:url("
                        htm &= imgsamp_rakuten2(gurl)
                        htm &= ");"">"

                        htm &= "</a>" & vbCrLf

                        j = j + 1

                        If j > 19 Then
                            Exit Do
                        End If

                        i = i + 1
                        If i > 10 Then
                            Exit Do
                        End If

                    Next
                End If
            End If
        Loop

        htm &= "  </div>" & vbCrLf
        htm &= "</section>" & vbCrLf
        htm &= "<!--週間ランキングここまで-->" & vbCrLf


        Form1.TextBox1.Text = htm

        ''データベースを切断
        Call sql_cl()

    End Sub

    Sub ex_exbut10()

        ''生活空間（Yahoo!）ＰＣ用新商品ＨＴＭＬ出力・通常作業者が使用する

        ''データベースと接続
        Call sql_st()

        Dim sql1 As String
        Dim htm As String


        ''差し込み用HTML作成
        htm = "<!--新着ここから-->" & vbCrLf
        htm &= "<div class=""contents-03"">" & vbCrLf
        htm &= "  <h2 class=""title-cmn"">NEW ARRIVAL</h2>" & vbCrLf
        htm &= "  <div class=""slide-category slide autoplay"">" & vbCrLf


        Dim j As Integer = 0
        Dim i As Integer = 1

        Do
            Dim tebx1 As Control() = Form1.Controls.Find("prid" & j.ToString.PadLeft(2, "0"c), True)
            If tebx1.Length < 0 Then
                MsgBox("IDのボックスがありません", MsgBoxStyle.Critical And MsgBoxStyle.OkOnly, "警告")
            End If

            If CType(tebx1(0), TextBox).Text = "" Then
            Else

                sql1 = "SELECT"
                sql1 &= " `code` , "
                sql1 &= " `name`,"
                sql1 &= " `original-price`,"
                sql1 &= " `price`"
                sql1 &= " FROM seikatsukukan_data_newest"
                sql1 &= " WHERE"
                sql1 &= " `code` = "
                sql1 &= "'"
                sql1 &= CType(tebx1(0), TextBox).Text
                sql1 &= "'"
                sql1 &= " AND "
                sql1 &= "`display` =1"
                sql1 &= ";"

                Dim dTb1 As DataTable = sql_result_return(sql1)

                If dTb1.Rows.Count = 0 Then
                    ''ロイヤル通販(楽天)にはテキストボックスの商品はない
                    j = j + 1
                    If j > 19 Then
                        Exit Do
                    End If

                Else
                    Dim n As Long = dTb1.Rows.Count
                    For Each DRow As DataRow In dTb1.Rows
                        Dim surl As String = DRow.Item(0)
                        surl = idc01(surl)

                        htm &= "<!--"
                        htm &= i
                        htm &= "個目-->" & vbCrLf

                        htm &= "    <div>" & vbCrLf
                        htm &= "      <div class=""inner"">" & vbCrLf
                        htm &= "        <a href=""https://store.shopping.yahoo.co.jp/seikatsukukan/"
                        htm &= surl
                        htm &= ".html#ItemInfo"">" & vbCrLf

                        htm &= "          <img src=""https://item-shopping.c.yimg.jp/i/j/seikatsukukan_"
                        htm &= surl
                        htm &= """ alt="""">" & vbCrLf
                        htm &= "        </a>" & vbCrLf
                        htm &= "      </div>" & vbCrLf
                        htm &= "    </div>" & vbCrLf


                        j = j + 1

                        If j > 19 Then
                            Exit Do
                        End If

                        i = i + 1
                        If i > 10 Then
                            Exit Do
                        End If

                    Next
                End If
            End If
        Loop

        htm &= "  </div>" & vbCrLf
        htm &= "  <!--/slide-category--> " & vbCrLf
        htm &= "</div>" & vbCrLf
        htm &= "<!--/contents-02--> " & vbCrLf
        htm &= "<!--新着ここまで-->" & vbCrLf
        Form1.TextBox1.Text = htm

        ''データベースを切断
        Call sql_cl()

    End Sub

    Sub ex_exbut11()

        ''生活空間（Yahoo!）スマホ用ランキングＨＴＭＬ出力・通常作業者が使用する

        ''データベースと接続
        Call sql_st()

        Dim sql1 As String
        Dim htm As String

        Dim nwtm01 As DateTime = Now()

        ''差し込み用HTML作成

        htm = "<!--新商品ここから-->" & vbCrLf
        htm &= "<section class=""section mb40"">" & vbCrLf
        htm &= "    <h3 class=""section_title"">NEW ARRIVAL</h3>" & vbCrLf
        htm &= "    <div class=""slide_list"">" & vbCrLf

        Dim j As Integer = 0
        Dim i As Integer = 1

        Do
            Dim tebx1 As Control() = Form1.Controls.Find("prid" & j.ToString.PadLeft(2, "0"c), True)
            If tebx1.Length < 0 Then
                MsgBox("IDのボックスがありません", MsgBoxStyle.Critical And MsgBoxStyle.OkOnly, "警告")
            End If

            If CType(tebx1(0), TextBox).Text = "" Then
            Else

                sql1 = "SELECT"
                sql1 &= " `code` , "
                sql1 &= " `name`,"
                sql1 &= " `original-price`,"
                sql1 &= " `price`"
                sql1 &= " FROM seikatsukukan_data_newest"
                sql1 &= " WHERE"
                sql1 &= " `code` = "
                sql1 &= "'"
                sql1 &= CType(tebx1(0), TextBox).Text
                sql1 &= "'"
                sql1 &= " AND "
                sql1 &= "`display` =1"
                sql1 &= ";"


                Dim dTb1 As DataTable = sql_result_return(sql1)

                If dTb1.Rows.Count = 0 Then
                    ''ロイヤル通販(楽天)にはテキストボックスの商品はない
                    j = j + 1
                    If j > 19 Then
                        Exit Do
                    End If

                Else
                    Dim n As Long = dTb1.Rows.Count
                    For Each DRow As DataRow In dTb1.Rows
                        Dim surl As String = DRow.Item(0)
                        surl = idc01(surl)

                        htm &= "<!--"
                        htm &= i
                        htm &= "個目-->" & vbCrLf

                        htm &= "    <a href=""https://store.shopping.yahoo.co.jp/seikatsukukan/"
                        htm &= surl
                        htm &= ".html"" style=""background-image:url("
                        htm &= "https://item-shopping.c.yimg.jp/i/l/seikatsukukan_"
                        htm &= surl
                        htm &= ");""></a>" & vbCrLf


                        j = j + 1

                        If j > 30 Then
                            Exit Do
                        End If

                        i = i + 1
                        If i > 10 Then
                            Exit Do
                        End If

                    Next
                End If
            End If
        Loop

        htm &= "  </div>" & vbCrLf
        htm &= "</section>" & vbCrLf
        htm &= "<!--週間ランキングここまで-->" & vbCrLf

        Form1.TextBox1.Text = htm

        ''データベースを切断
        Call sql_cl()

    End Sub



End Module
