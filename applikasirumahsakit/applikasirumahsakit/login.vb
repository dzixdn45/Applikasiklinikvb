Imports System.Data.OleDb

Public Class login
    Dim connection As OleDbConnection
    Dim dataadapter As OleDbDataAdapter
    Dim command As New OleDbCommand
    Dim read As OleDbDataReader
    Dim str As String
    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        command = New OleDbCommand("select * from kodepemakai where namapemakai='" & TextBox1.Text & "' and passpemakai='" & TextBox2.Text & "'", connection)
        read = command.ExecuteReader
        If (read.Read()) Then
            pendaftaran.Show()
            pendaftaran.TextBox8.Text = TextBox1.Text
            Me.Hide()
            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox1.Focus()

        Else
            MsgBox("username dan password anda salah mohon dipebaiki lagi", MsgBoxStyle.OkOnly, -"login gagal")
            TextBox1.Text = ""
            TextBox2.Text = ""
            TextBox1.Focus()
        End If
    End Sub

    Private Sub login_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        str = "provider=microsoft.ACE.OLEDB.12.0;data source=didindb.accdb"
        connection = New OleDbConnection(str)
        If connection.State = ConnectionState.Closed Then
            connection.Open()

        End If
    End Sub
End Class