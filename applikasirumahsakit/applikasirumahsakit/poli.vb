Imports System.Data.OleDb
Public Class poli

    Dim connection As OleDbConnection
    Dim dataadapter As OleDbDataAdapter
    Dim dataset As New DataSet
    Dim command As New OleDbCommand
    Dim rd As OleDbDataReader
    Dim str As String

    Sub tampil()
        dataadapter = New OleDbDataAdapter("select * from poli ", connection)
        dataset = New DataSet
        dataset.Clear()
        dataadapter.Fill(dataset, "poli")
        DataGridView1.DataSource = (dataset.Tables("poli"))
        DataGridView1.ReadOnly = True
        DataGridView1.Columns(0).Width = 100
        DataGridView1.Columns(1).Width = 100
    End Sub

    Sub reset()
        TextBox1.Text = ""
        TextBox2.Text = ""
    End Sub

    Private Sub poli_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        str = "provider=microsoft.ACE.OLEDB.12.0;data source=didindb.accdb"
        connection = New OleDbConnection(str)
        If connection.State = ConnectionState.Closed Then
            connection.Open()
        End If
    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Then
            MsgBox("Data inputan semua wajib diisi", MsgBoxStyle.OkOnly)
            Exit Sub
        Else
            command = New OleDbCommand("select * from poli where KodeObt='" & TextBox1.Text & "'", connection)
            rd = command.ExecuteReader()
            rd.Read()
            If Not rd.HasRows Then
                Dim dbtambah As String
                dbtambah = "INSERT INTO poli values('" & TextBox1.Text & "','" & TextBox2.Text & "')"
                command = New OleDbCommand(dbtambah, connection)
                command.ExecuteNonQuery()
                tampil()
                reset()
            Else
                Dim dbedit As String
                dbedit = "UPDATE poli SET namapoli = '" & TextBox2.Text & "'"
                command = New OleDbCommand(dbedit, connection)
                command.ExecuteNonQuery()
                tampil()
                reset()

                command.ExecuteNonQuery()
                tampil()
                reset()
            End If
        End If
    End Sub
End Class