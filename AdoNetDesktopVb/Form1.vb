Imports System.Data.SqlClient
Imports System.Text.RegularExpressions
Imports Microsoft.Data.SqlClient
Public Class Form1

    Private sqlserver As FileSupportSqlSvr
    Private DbOpenType As Boolean = True
    Friend WithEvents LabelDic As New Dictionary(Of String, Label)()
    Friend WithEvents LabelNumId As Label
    Friend WithEvents LabelFoot As Label
    Friend WithEvents TextBoxNote As TextBox
    Friend WithEvents TextBoxShohinName As TextBox
    Friend WithEvents TextBoxShohinNum As TextBox
    Friend WithEvents RichTextBox1 As RichTextBox
    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents ButtonDelete As Button
    Friend WithEvents ButtonUpdate As Button
    Friend WithEvents ButtonInsert As Button
    Friend WithEvents ButtonQuery As Button
    Friend WithEvents BindingSource1 As BindingSource

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load

        Call FormDesignSetting()

    End Sub

    ''' <summary>商品を全件表示します。</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ButtonQuery_Click(sender As Object, e As EventArgs) Handles ButtonQuery.Click

        Using sqlserver = New FileSupportSqlSvr()
            DatabaseOpen(sqlserver)

            Dim list As List(Of ShohinDto) = sqlserver.DataReader(Of ShohinDto)("select * from ShohinDataDesk order by NumId asc").ToList()
            BindingSource1.DataSource = list
            DataGridView1.DataSource = BindingSource1
            Call DataGridSetting()
            Call TextBoxClear()
            RichTextBox1.AppendText("商品を全件表示しました。" & vbCrLf)
        End Using

    End Sub

    ''' <summary>テキストボックスによる内容で商品を追加します。</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ButtonInsert_Click(sender As Object, e As EventArgs) Handles ButtonInsert.Click

        Dim sqlstr As String = ""

        If Regex.IsMatch(TextBoxShohinNum.Text, "^[0-9]{1,4}$") = False Then
            MessageBox.Show("商品番号は半角数値の0～9999でなければなりません。", "メッセージ", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        Using sqlserver = New FileSupportSqlSvr()
            DatabaseOpen(sqlserver)
            sqlstr = "insert into ShohinDataDesk (ShohinNum, ShohinName, EditDate, EditTime, Note) "
            sqlstr &= "values (@ShohinNum, @ShohinName, @EditDate, @EditTime, @Note)"
            Using cmd = New SqlCommand()
                cmd.Parameters.Clear()
                cmd.Parameters.Add("@ShohinNum", SqlDbType.SmallInt)
                cmd.Parameters.Add("@ShohinName", SqlDbType.Char, 50)
                cmd.Parameters.Add("@EditDate", SqlDbType.Decimal, 8)
                cmd.Parameters.Add("@EditTime", SqlDbType.Decimal, 6)
                cmd.Parameters.Add("@Note", SqlDbType.VarChar, 255)

                cmd.Parameters("@ShohinNum").Value = Short.Parse(TextBoxShohinNum.Text)
                cmd.Parameters("@ShohinName").Value = TextBoxShohinName.Text
                cmd.Parameters("@EditDate").Value = Format(Now, "yyyyMMdd")
                cmd.Parameters("@EditTime").Value = Format(Now, "HHmmss")
                cmd.Parameters("@Note").Value = TextBoxNote.Text

                Using tran As SqlTransaction = sqlserver.GoTransaction()
                    Try
                        sqlserver.NonQuery(sqlstr, cmd, tran)
                    Catch ex As SqlException
                        sqlserver.TransactionRollback(tran)
                        Throw
                    End Try
                    sqlserver.TransactionCommit(tran)
                End Using
                RichTextBox1.AppendText("1件追加しました" & vbCrLf)
            End Using
        End Using

    End Sub

    ''' <summary>商品ID(NumId)による商品の更新を行います。</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ButtonUpdate_Click(sender As Object, e As EventArgs) Handles ButtonUpdate.Click

        Dim sqlstr As String = ""

        If DataGridView1.Rows.Count <= 0 Or LabelNumId.Text = "" Then
            MessageBox.Show("更新する商品行が選択できていません。", "商品IDなし", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        If Regex.IsMatch(TextBoxShohinNum.Text, "^[0-9]{1,4}$") = False Then
            MessageBox.Show("商品番号は半角数値の0～9999でなければなりません。", "メッセージ", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        Using sqlserver = New FileSupportSqlSvr()
            DatabaseOpen(sqlserver)
            sqlstr = "update ShohinDataDesk set ShohinNum=@ShohinNum, ShohinName=@ShohinName"
            sqlstr &= ", EditDate=@EditDate, EditTime=@EditTime, Note=@Note where NumId=@NumId"
            Using cmd = New SqlCommand()
                cmd.Parameters.Clear()
                cmd.Parameters.Add("@NumId", SqlDbType.Int)
                cmd.Parameters.Add("@ShohinNum", SqlDbType.SmallInt)
                cmd.Parameters.Add("@ShohinName", SqlDbType.Char, 50)
                cmd.Parameters.Add("@EditDate", SqlDbType.Decimal, 8)
                cmd.Parameters.Add("@EditTime", SqlDbType.Decimal, 6)
                cmd.Parameters.Add("@Note", SqlDbType.VarChar, 255)

                cmd.Parameters("@NumId").Value = Integer.Parse(LabelNumId.Text)
                cmd.Parameters("@ShohinNum").Value = Short.Parse(TextBoxShohinNum.Text)
                cmd.Parameters("@ShohinName").Value = TextBoxShohinName.Text
                cmd.Parameters("@EditDate").Value = Format(Now, "yyyyMMdd")
                cmd.Parameters("@EditTime").Value = Format(Now, "HHmmss")
                cmd.Parameters("@Note").Value = TextBoxNote.Text

                Using tran As SqlTransaction = sqlserver.GoTransaction()
                    Try
                        sqlserver.NonQuery(sqlstr, cmd, tran)
                    Catch ex As SqlException
                        sqlserver.TransactionRollback(tran)
                        Throw
                    End Try
                    sqlserver.TransactionCommit(tran)
                End Using
                RichTextBox1.AppendText("選択された商品を更新しました。" & vbCrLf)
            End Using
        End Using

    End Sub

    ''' <summary>商品ID(NumId)による商品を削除します。</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub ButtonDelete_Click(sender As Object, e As EventArgs) Handles ButtonDelete.Click

        Dim DelNum As Integer
        Dim sqlstr As String = ""

        If DataGridView1.Rows.Count <= 0 Or LabelNumId.Text = "" Then
            MessageBox.Show("削除する商品行が選択がされていません", "商品IDなし", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        DelNum = Integer.Parse(LabelNumId.Text)
        Using sqlserver = New FileSupportSqlSvr()
            DatabaseOpen(sqlserver)
            sqlstr = "delete from ShohinDataDesk where NumId = @NumId"
            Using cmd = New SqlCommand()
                cmd.Parameters.Clear()
                cmd.Parameters.Add("@NumId", SqlDbType.Int)
                cmd.Parameters("@NumId").Value = DelNum

                Using tran As SqlTransaction = sqlserver.GoTransaction()
                    Try
                        sqlserver.NonQuery(sqlstr, cmd, tran)
                    Catch ex As SqlException
                        sqlserver.TransactionRollback(tran)
                        Throw
                    End Try
                    sqlserver.TransactionCommit(tran)
                End Using
                RichTextBox1.AppendText(DelNum & "の行を削除しました" & vbCrLf)
            End Using
        End Using

    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick

        LabelNumId.Text = Integer.Parse(DataGridView1.CurrentRow.Cells("NumId").Value)
        TextBoxShohinNum.Text = DataGridView1.CurrentRow.Cells("ShohinNum").Value
        TextBoxShohinName.Text = DataGridView1.CurrentRow.Cells("ShohinName").Value
        TextBoxNote.Text = DataGridView1.CurrentRow.Cells("Note").Value

    End Sub

    Private Function LabelsSetting(name As String, txt As String, x As Integer, y As Integer, w As Integer, h As Integer) As Label

        Dim label As New Label()
        label.Name = name
        label.AutoSize = False
        label.Text = txt
        label.Location = New Point(x, y)
        label.Size = New Size(w, h)
        LabelDic.Add(label.Name, label)
        Controls.Add(label)

        Return label

    End Function

    ''' <summary>
    ''' 与えられたControlクラスオブジェクトのロケーション、サイズを設定しフォームに追加しControlオブジェクトで戻します
    ''' </summary>
    ''' <param name="ctl">System.Windows.Forms.Control</param>
    ''' <param name="name">オブジェクト名</param>
    ''' <param name="x">ロケーションX</param>
    ''' <param name="y">ロケーションY</param>
    ''' <param name="w">コントロールの横サイズ</param>
    ''' <param name="h">コントロールの縦サイズ</param>
    ''' <returns>System.Windows.Forms.Control</returns>
    Private Function ControlsSetting(ctl As Control, name As String, x As Integer, y As Integer, w As Integer, h As Integer) As Control

        ctl.Name = name
        ctl.Location = New Point(x, y)
        ctl.Size = New Size(w, h)
        Controls.Add(ctl)

        Return ctl

    End Function

    Private Sub FormDesignSetting()

        Me.components = New System.ComponentModel.Container()

        Name = "Form1"
        Text = "ADO.NET + デスクトップアプリ + SQL Server"
        Location = New Point(500, 200)
        Size = New Size(800, 600)
        MaximizeBox = False
        MinimizeBox = False
        'Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        'Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        'Me.ClientSize = New System.Drawing.Size(784, 561)
        'Me.ResumeLayout(False)
        'Me.PerformLayout()


        DataGridView1 = New DataGridView()
        DataGridView1 = CType(ControlsSetting(DataGridView1, "DataGridView1", 25, 25, 730, 200), DataGridView)
        CType(DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()

        BindingSource1 = New BindingSource(Me.components)
        CType(BindingSource1, System.ComponentModel.ISupportInitialize).EndInit()

        RichTextBox1 = New RichTextBox()
        RichTextBox1.ReadOnly = True
        RichTextBox1 = CType(ControlsSetting(RichTextBox1, "RichTextBox1", 25, 235, 350, 200), RichTextBox)

        ButtonQuery = New Button()
        ButtonQuery.Text = "抽出"
        ButtonQuery.TabIndex = 3
        ButtonQuery.UseVisualStyleBackColor = True
        ButtonQuery = CType(ControlsSetting(ButtonQuery, "ButtonQuery", 50, 470, 150, 50), Button)

        ButtonInsert = New Button()
        ButtonInsert.Text = "追加"
        ButtonInsert.TabIndex = 4
        ButtonInsert.UseVisualStyleBackColor = True
        ButtonInsert = CType(ControlsSetting(ButtonInsert, "ButtonInsert", 230, 470, 150, 50), Button)

        ButtonUpdate = New Button()
        ButtonUpdate.Text = "更新"
        ButtonUpdate.TabIndex = 5
        ButtonUpdate.UseVisualStyleBackColor = True
        ButtonUpdate = CType(ControlsSetting(ButtonUpdate, "ButtonUpdate", 410, 470, 150, 50), Button)

        ButtonDelete = New Button()
        ButtonDelete.Text = "削除"
        ButtonDelete.TabIndex = 6
        ButtonDelete.UseVisualStyleBackColor = True
        ButtonDelete = CType(ControlsSetting(ButtonDelete, "ButtonDelete", 590, 470, 150, 50), Button)

        LabelsSetting("Label1", "商品ID：", 385, 250, 75, 25)
        LabelsSetting("Label2", "商品番号：", 385, 300, 75, 25)
        LabelsSetting("Label3", "商品名：", 385, 350, 75, 25)
        LabelsSetting("Label4", "備考：", 385, 400, 60, 25)

        LabelNumId = New Label()
        LabelNumId.AutoSize = False
        LabelNumId.Text = ""
        LabelNumId.TextAlign = ContentAlignment.TopRight
        LabelNumId = CType(ControlsSetting(LabelNumId, "LabelNumId", 690, 250, 60, 19), Label)

        LabelFoot = New Label()
        LabelFoot.AutoSize = False
        LabelFoot.Text = "Copyright (c)  2021-2024  support-for-bingo"
        LabelFoot = CType(ControlsSetting(LabelFoot, "LabelFoot", 30, 535, 300, 19), Label)

        TextBoxShohinNum = New TextBox()
        TextBoxShohinNum.TabIndex = 0
        TextBoxShohinNum = CType(ControlsSetting(TextBoxShohinNum, "TextBoxShohinNum", 600, 300, 150, 19), TextBox)

        TextBoxShohinName = New TextBox()
        TextBoxShohinName.TabIndex = 1
        TextBoxShohinName = CType(ControlsSetting(TextBoxShohinName, "TextBoxShohinName", 550, 350, 200, 19), TextBox)

        TextBoxNote = New TextBox()
        TextBoxNote.TabIndex = 2
        TextBoxNote = CType(ControlsSetting(TextBoxNote, "TextBoxNote", 450, 400, 300, 19), TextBox)

    End Sub

    Private Sub TextBoxClear()

        LabelNumId.Text = ""
        TextBoxShohinNum.Text = ""
        TextBoxShohinName.Text = ""
        TextBoxNote.Text = ""

    End Sub

    Private Sub DatabaseOpen(ByVal sqlserver As FileSupportSqlSvr)

        If DbOpenType Then
            Dim path As String = Environment.CurrentDirectory & "\"
            If sqlserver.Open(path, "MsSqlServer.xml") = False Then
                MessageBox.Show("データベース設定ファイルがありません。" & vbCrLf & "アプリケーションを終了します")
                Application.Exit()
            End If
        Else
            sqlserver.Host = "lpc:(local)"
            sqlserver.Instance = "SQLEXPRESS"
            sqlserver.LoginMode = True
            sqlserver.Catalog = "AdoNetSample"
            sqlserver.ConnectTimeout = 3
            sqlserver.MultipleActiveResultSets = False
            sqlserver.Open()
        End If

    End Sub

    Private Sub DataGridSetting()

        DataGridView1.Columns("NumId").HeaderText = "商品ID"
        DataGridView1.Columns("ShohinNum").HeaderText = "商品番号"
        DataGridView1.Columns("ShohinName").HeaderText = "商品名"
        DataGridView1.Columns("EditDate").HeaderText = "編集日付"
        DataGridView1.Columns("EditTime").HeaderText = "編集時刻"
        DataGridView1.Columns("Note").HeaderText = "備考"
        DataGridView1.Columns("NumId").Width = 70
        DataGridView1.Columns("Note").Width = 250
        DataGridView1.Columns("EditDate").DefaultCellStyle.Format = "0000/00/00"
        DataGridView1.Columns("EditTime").DefaultCellStyle.Format = "00:00:00"
        DataGridView1.AllowUserToAddRows = False
        DataGridView1.RowHeadersVisible = False
        DataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect
        DataGridView1.ReadOnly = True

    End Sub

End Class