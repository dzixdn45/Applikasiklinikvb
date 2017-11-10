Imports System.Data.OleDb
Public Class pendaftaran
    Dim connection As OleDbConnection
    Dim dataadapter As OleDbDataAdapter
    Dim dataset As New DataSet
    Dim command As New OleDbCommand
    Dim read As OleDbDataReader
    Dim str As String
    Sub tampil()
        dataadapter = New OleDbDataAdapter("select * from pendaftaran", connection)
        dataset = New DataSet
        dataset.Clear()
        dataadapter.Fill(dataset, "pendaftaran")
        DataGridView1.DataSource = (dataset.Tables("pendaftaran"))
        DataGridView1.ReadOnly = True
        DataGridView1.Columns(0).Width = 100
        DataGridView1.Columns(1).Width = 100
        DataGridView1.Columns(2).Width = 100
        DataGridView1.Columns(3).Width = 100
        DataGridView1.Columns(4).Width = 100
        DataGridView1.Columns(5).Width = 100
        DataGridView1.Columns(6).Width = 100
    End Sub
    Sub tampilpoli()
        command = New OleDbCommand("select * from poli", connection)
        read = command.ExecuteReader
        ComboBox1.Items.Clear()
        Do While read.Read
            ComboBox1.Items.Add(read.Item(1))
        Loop
    End Sub
    Sub nootomatis()
        command = New OleDbCommand("SELECT * FROM pendaftaran WHERE nomor_pendaftaran in (select max(nomor_pendaftaran) from pendaftaran) order by nomor_pendaftaran desc", connection)
        read = command.ExecuteReader
        read.Read()
        If Not read.HasRows Then
            TextBox1.Text = Format(Now, "yyMMdd") + "0001"
            TextBox1.ReadOnly = True

        Else
            If Microsoft.VisualBasic.Left(read.GetString(0), 6) <> Format(Now, "yyMMdd") Then
                TextBox1.Text = Format(Now, "yyMMdd") + "0001"
                TextBox1.ReadOnly = True
            Else
                TextBox1.Text = read.GetString(0) + 1
                TextBox1.ReadOnly = True
            End If

        End If
    End Sub
    Sub tampilpasien()
        command = New OleDbCommand("select * from Pasien", connection)
        read = command.ExecuteReader
        ComboBox2.Items.Clear()
        Do While read.Read
            ComboBox2.Items.Add(read.Item(0))
        Loop
    End Sub
    Sub tampildokter()

    End Sub

    Private Sub pendaftaran_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        str = "provider=microsoft.ACE.OLEDB.12.0;data source=didindb.accdb"
        connection = New OleDbConnection(str)
        If connection.State = ConnectionState.Closed Then
            connection.Open()
        End If
        Label17.Text = Today
        tampilpoli()
        tampilpasien()
        tampil()
        nootomatis()

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        command = New OleDbCommand("select * from Dokter Where spesialis = '" & ComboBox1.Text & "'", connection)
        read = command.ExecuteReader()
        ListBox1.Items.Clear()
        Do While read.Read
            ListBox1.Items.Add(read.Item(1))
            TextBox9.Text = read.Item(0)
            TextBox8.Text = read.Item(6)
        Loop
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        command = New OleDbCommand("select * from Pasien Where KodePsn = '" & ComboBox2.Text & "'", connection)
        read = command.ExecuteReader()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox5.Clear()
        read.Read()
        If read.HasRows Then
            TextBox2.Text = read.Item("NamaPsn")
            TextBox3.Text = read.Item("NamaPsn")
            TextBox4.Text = read.Item("AlamatPsn")
            TextBox5.Text = read.Item("GenderPsn")
            TextBox6.Text = read.Item("TlpPsn")
        End If
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ListBox1.SelectedIndexChanged
        Dim nama_dkt As String = ListBox1.SelectedItem
        command = New OleDbCommand("select * from Dokter Where namadokter = '" & nama_dkt & "'", connection)
        read = command.ExecuteReader()
        read.Read()
        If read.HasRows Then
            TextBox7.Text = read.Item(5)
        End If
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        If ComboBox1.Text = "" Or ComboBox2.Text = "" Or TextBox7.Text = "" Then
            MsgBox("Data belum lengkap")
        End If
        Dim dbtambah1 As String
        dbtambah1 = "INSERT INTO pendaftaran VALUES('" & TextBox1.Text & "','" & TextBox2.Text & "','" & Label17.Text & "','" & TextBox8.Text & "','" & TextBox9.Text & "','" & ComboBox2.Text & "','" & TextBox7.Text & "',0)"
        command = New OleDbCommand(dbtambah1, connection)
        command.ExecuteNonQuery()
        tampil()
        Reset()
        nootomatis()
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click

    End Sub

    Private Sub TextBox1_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox1.TextChanged

    End Sub
End Class