Imports System.Data.OleDb

Public Class formobat
    Dim connection As OleDbConnection
    Dim dataadapter As OleDbDataAdapter
    Dim dataset As New DataSet
    Dim command As New OleDbCommand
    Dim rd As OleDbDataReader
    Dim str As String

    Sub tampil()
        dataadapter = New OleDbDataAdapter("select * from Obat", connection)
        dataset = New DataSet
        dataset.Clear()
        dataadapter.Fill(dataset, "Obat")
        DGV.DataSource = (dataset.Tables("Obat"))
        DGV.ReadOnly = True
        DGV.Columns(0).Width = 100
        DGV.Columns(1).Width = 100
        DGV.Columns(2).Width = 100
        DGV.Columns(3).Width = 100
        DGV.Columns(4).Width = 100
        DGV.Columns(5).Width = 100
    End Sub
    Sub reset()
        TextBox1.Text = ""
        TextBox2.Text = ""
        ComboBox2.Text = ""
        ComboBox1.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
    End Sub
    Sub tampilkategori()
        command = New OleDbCommand("select * from poli", connection)
        rd = command.ExecuteReader()
        ComboBox1.Items.Clear()
        Do While rd.Read
            ComboBox1.Items.Add(rd.Item(1))
        Loop

    End Sub
    Sub tampiljenis()
        command = New OleDbCommand("select * from jenis", connection)
        rd = command.ExecuteReader()
        ComboBox2.Items.Clear()
        Do While rd.Read
            ComboBox2.Items.Add(rd.Item(1))
        Loop

    End Sub
    Sub kodeotomatis()
        Dim strsementara As String = ""
        Dim strisi As String = ""
        command = New OleDbCommand("select * from Obat order by KodeObt desc", connection)
        rd = command.ExecuteReader()
        If rd.Read Then
            strsementara = Mid(rd.Item("KodeObt"), 2, 2)
            strisi = Val(strsementara) + 1
            TextBox1.Text = "OBT" + Mid("0", 1, 2 - strisi.Length) & strisi
        Else
            TextBox1.Text = "OBT01"
        End If
    End Sub
    Private Sub formobat_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        str = "provider=microsoft.ACE.OLEDB.12.0;data source=didindb.accdb"
        connection = New OleDbConnection(str)
        If connection.State = ConnectionState.Closed Then
            connection.Open()
        End If
        kodeotomatis()
        tampil()
        tampilkategori()
        tampiljenis()

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If TextBox1.Text = "" Or TextBox2.Text = "" Or ComboBox2.Text = "" Then
            MsgBox("Data inputan semua wajib diisi", MsgBoxStyle.OkOnly)
            Exit Sub
        Else
            command = New OleDbCommand("select * from Obat where KodeObt='" & TextBox1.Text & "'", connection)
            rd = command.ExecuteReader()
            rd.Read()
            If Not rd.HasRows Then
                Dim dbtambah As String
                dbtambah = "INSERT INTO Obat values('" & TextBox1.Text & "','" & TextBox2.Text & "','" & ComboBox2.Text & "','" & ComboBox1.Text & "','" & TextBox5.Text & "','" & TextBox6.Text & "')"
                command = New OleDbCommand(dbtambah, connection)
                command.ExecuteNonQuery()
                kodeotomatis()
                tampil()
                reset()
            Else
                Dim dbedit As String
                dbedit = "UPDATE Obat SET NamaObat = '" & TextBox2.Text & "',JenisObt = '" & ComboBox2.Text & "',Kategori = '" & ComboBox1.Text & "',HargaObt = '" & TextBox5.Text & "', JumlahObt = '" & TextBox6.Text & "' WHERE KodeObt = '" & TextBox1.Text & "'"
                command = New OleDbCommand(dbedit, connection)
                command.ExecuteNonQuery()
                kodeotomatis()
                tampil()
                reset()
            End If
        End If
    End Sub
    Private Sub DGV_CellMouseClick(sender As Object, e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DGV.CellMouseClick
        TextBox1.Text = DGV.CurrentRow.Cells(0).Value
        TextBox2.Text = DGV.CurrentRow.Cells(1).Value
        ComboBox2.Text = DGV.CurrentRow.Cells(2).Value
        ComboBox1.Text = DGV.CurrentRow.Cells(3).Value
        TextBox5.Text = DGV.CurrentRow.Cells(4).Value
        TextBox6.Text = DGV.CurrentRow.Cells(5).Value
        Button2.Enabled = False
        kodeotomatis()
    End Sub

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        If MsgBox("Yakin akan menghapus data?", MsgBoxStyle.YesNo, _
       "Konfirmasi") = MsgBoxResult.No Then Exit Sub
        Dim dbhapus As String
        dbhapus = "DELETE FROM Obat WHERE KodeObt = " & _
              "'" & DGV.CurrentRow.Cells(0).Value & "'"

        command = New OleDbCommand(dbhapus, connection)
        command.ExecuteNonQuery()
        tampil()
        kodeotomatis()
        reset()
    End Sub

    Private Sub TextBox7_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox7.TextChanged
        dataadapter = New OleDbDataAdapter("Select * from Obat Where KodeObt like '" & Me.TextBox7.Text & "%' or NamaObat like '" & Me.TextBox7.Text & "%' or JenisObt like '" & Me.TextBox7.Text & "%'", connection)
        dataset = New DataSet
        dataset.Clear()
        dataadapter.Fill(dataset, "Obat")
        DGV.DataSource = (dataset.Tables("Obat"))
        Dim tampil1 As String
        tampil1 = "select * from Obat where KodeObt ='" & TextBox1.Text & "'"
        command = New OleDbCommand(tampil1, connection)
        rd = command.ExecuteReader

        If rd.Read Then
            TextBox1.Text = rd.Item("KodeObt")
            TextBox2.Text = rd.Item("NamaObat")
            ComboBox2.Text = rd.Item("JenisObt")
            ComboBox1.Text = rd.Item("Kategori")
            TextBox5.Text = rd.Item("HargaObt")
            TextBox7.Text = rd.Item("JumlahObt")
            TextBox1.Enabled = False

        End If
        kodeotomatis()
    End Sub

    Private Sub Button6_Click(sender As System.Object, e As System.EventArgs) Handles Button6.Click
        login.Show()

    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        Close()
    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click

    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        kodeotomatis()
    End Sub
End Class
