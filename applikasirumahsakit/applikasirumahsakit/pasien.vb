Imports System.Data.OleDb
Public Class pasien
    Dim connection As OleDbConnection
    Dim dataadapter As OleDbDataAdapter
    Dim dataset As New DataSet
    Dim command As New OleDbCommand
    Dim read As OleDbDataReader
    Dim str As String
    Sub tampil()
        dataadapter = New OleDbDataAdapter("select * from Pasien", connection)
        dataset = New DataSet
        dataset.Clear()
        dataadapter.Fill(dataset, "Pasien")
        DGV.DataSource = (dataset.Tables("Pasien"))
        DGV.ReadOnly = True
        DGV.Columns(0).Width = 100
        DGV.Columns(1).Width = 100
        DGV.Columns(2).Width = 100
        DGV.Columns(3).Width = 100
        DGV.Columns(4).Width = 100
        DGV.Columns(5).Width = 100
    End Sub
    Sub reset()
        TextBox2.Text = ""
        TextBox3.Text = ""
        ComboBox1.Text = ""
        TextBox4.Text = ""
    End Sub
    Private Sub pasien_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        str = "provider=microsoft.ACE.OLEDB.12.0;data source=didindb.accdb"
        connection = New OleDbConnection(str)
        If connection.State = ConnectionState.Closed Then
            connection.Open()
        End If
        tampil()
        reset()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If ComboBox1.Text <> "laki-laki" Or ComboBox1.Text <> "perempuan" Then
            MessageBox.Show("inputkan dengan benar", "WARNING")
            If TextBox1.Text = "" Or TextBox2.Text = "" Or TextBox3.Text = "" Then
                MsgBox("Data inputan semua wajib diisi", MsgBoxStyle.OkOnly)
                Exit Sub
            Else
                command = New OleDbCommand("select * from Pasien where KodePsn='" & TextBox1.Text & "'", connection)
                read = command.ExecuteReader()
                read.Read()
                If Not read.HasRows Then
                    Dim dbtambah As String
                    dbtambah = "INSERT INTO Pasien values('" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & ComboBox1.Text & "','" & DateTimePicker1.Text & "','" & TextBox4.Text & "')"
                    command = New OleDbCommand(dbtambah, connection)
                    command.ExecuteNonQuery()
                    tampil()
                    reset()

                Else
                    Dim dbedit As String
                    dbedit = "UPDATE Pasien SET NamaPsn = '" & TextBox2.Text & "',AlamatPsn = '" & TextBox3.Text & "',GenderPsn = '" & ComboBox1.Text & "',TglPsn = '" & DateTimePicker1.Text & "',TlpPsn = '" & TextBox4.Text & "' WHERE KodePsn = '" & TextBox1.Text & "'"
                    command = New OleDbCommand(dbedit, connection)
                    command.ExecuteNonQuery()
                    tampil()
                    reset()
                End If
            End If
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If MsgBox("Yakin akan menghapus data?", MsgBoxStyle.YesNo, _
"Konfirmasi") = MsgBoxResult.No Then Exit Sub
        Dim dbhapus As String
        dbhapus = "DELETE FROM Pasien WHERE KodePsn = " & _
              "'" & DGV.CurrentRow.Cells(0).Value & "'"

        command = New OleDbCommand(dbhapus, connection)
        command.ExecuteNonQuery()
        tampil()
        reset()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        reset()
    End Sub

    Private Sub TextBox5_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox5.TextChanged
        dataadapter = New OleDbDataAdapter("Select * from Pasien Where KodePsn like '" & Me.TextBox5.Text & "%' or NamaPsn like '" & Me.TextBox5.Text & "%' or AlamatPsn like '" & Me.TextBox4.Text & "%' or GenderPsn like '" & Me.TextBox4.Text & "%'", connection)
        dataset = New DataSet
        dataset.Clear()
        dataadapter.Fill(dataset, "Pasien")
        DGV.DataSource = (dataset.Tables("Pasien"))
        Dim tampil1 As String
        tampil1 = "select * from Pasien where KodePsn ='" & TextBox1.Text & "'"
        command = New OleDbCommand(tampil1, connection)
        read = command.ExecuteReader

        If read.Read Then
            TextBox1.Text = read.Item("KodePsn")
            TextBox2.Text = read.Item("NamaPsn")
            TextBox3.Text = read.Item("AlamatPsn")
            ComboBox1.Text = read.Item("GenderPsn")
            DateTimePicker1.Text = read.Item("TglPsn")
            TextBox4.Text = read.Item("TlpPsn")
            TextBox1.Enabled = False

        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Close()
    End Sub

    Private Sub DGV_CellMouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DGV.CellMouseClick
        TextBox1.Text = DGV.CurrentRow.Cells(0).Value
        TextBox2.Text = DGV.CurrentRow.Cells(1).Value
        TextBox3.Text = DGV.CurrentRow.Cells(2).Value
        ComboBox1.Text = DGV.CurrentRow.Cells(3).Value
        DateTimePicker1.Text = DGV.CurrentRow.Cells(4).Value
        TextBox4.Text = DGV.CurrentRow.Cells(5).Value
    End Sub
End Class