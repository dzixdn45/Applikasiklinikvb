Imports System.Data.OleDb
Public Class resep
    Dim connection As OleDbConnection
    Dim dataadapter As OleDbDataAdapter
    Dim dataset As New DataSet
    Dim command As New OleDbCommand
    Dim read As OleDbDataReader
    Dim str As String
    Sub buatkolombaru()
        DataGridView1.Columns.Add("kode", "kode")
        DataGridView1.Columns.Add("nama", "nama obat")
        DataGridView1.Columns.Add("harga", "harga")
        DataGridView1.Columns.Add("dosis", "dosis")
        DataGridView1.Columns.Add("subtotal", "subtotal")
    End Sub

    Sub aturkolom()
        DataGridView1.Columns(0).Width = 100
        DataGridView1.Columns(1).Width = 100
        DataGridView1.Columns(2).Width = 100
        DataGridView1.Columns(3).Width = 100
        DataGridView1.Columns(4).Width = 100
    End Sub

    Sub tampiltransaksi()
        command = New OleDbCommand("SELECT * FROM Obat WHERE KodeObt = '" & ListBox1.Items(0) & "'", connection)
        read = command.ExecuteReader
        ListBox1.Items.Clear()
        Do While read.Read
        Loop
    End Sub

    Sub totalitem()
        Dim hitungitem As Integer = 0
        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            hitungitem = hitungitem + Val(DataGridView1.Rows(i).Cells(4).Value)
            Label9.Text = hitungitem
        Next
    End Sub
    Sub totalharga()
        Dim hitungharga As Integer = 0
        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            hitungharga = hitungharga + Val(DataGridView1.Rows(i).Cells(4).Value)
            TextBox8.Text = hitungharga
        Next
    End Sub

    Sub tampilobat()
        command = New OleDbCommand("SELECT * FROM Obat WHERE Kategori = '" & TextBox7.Text & "'", connection)
        read = command.ExecuteReader
        ListBox1.Items.Clear()
        Do While read.Read
            ListBox1.Items.Add(read.Item(0) & Space(5) & read.Item(1))
        Loop

    End Sub

    Sub tampildaftar()
        command = New OleDbCommand("select * from pendaftaran", connection)
        read = command.ExecuteReader
        ComboBox1.Items.Clear()
        Do While read.Read
            ComboBox1.Items.Add(read.Item(0))
        Loop
    End Sub
    Private Sub resep_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        str = "provider=microsoft.ACE.OLEDB.12.0;data source=didindb.accdb"
        connection = New OleDbConnection(str)
        If connection.State = ConnectionState.Closed Then
            connection.Open()
        End If
        tampildaftar()
        tampilobat()
        buatkolombaru()
        aturkolom()
        TextBox1.Text = Today
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        command = New OleDbCommand("select * from pendaftaran where nomor_pendaftaran = '" & ComboBox1.Text & "'", connection)
        read = command.ExecuteReader()
        ListBox1.Items.Clear()
        Do While read.Read
            TextBox2.Text = read.Item(4)
            TextBox3.Text = read.Item(5)
            TextBox4.Text = read.Item(3)
        Loop
        command = New OleDbCommand("select * from Dokter where kd_dokter = '" & TextBox2.Text & "'", connection)
        read = command.ExecuteReader()
        ListBox1.Items.Clear()
        Do While read.Read
            TextBox5.Text = read.Item(1)
        Loop
        command = New OleDbCommand("select * from Pasien where KodePsn = '" & TextBox3.Text & "'", connection)
        read = command.ExecuteReader()
        ListBox1.Items.Clear()
        Do While read.Read
            TextBox6.Text = read.Item(1)
        Loop

        command = New OleDbCommand("select * from Poli where kd_poli = '" & TextBox4.Text & "'", connection)
        read = command.ExecuteReader()
        ListBox1.Items.Clear()
        Do While read.Read
            TextBox7.Text = read.Item(1)
        Loop
        tampilobat()
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        If TextBox8.Text = "" Or TextBox9.Text = "" Or TextBox10.Text = "" Or Label9.Text = "" Then
            MsgBox("Data anda belum lengkap coeg , tidak ada transaksi atau pembayaran kosong")
            Exit Sub
        End If

        Dim insertresep As String = "insert into Resep(kodersp, tglresep, kodeobt, kodepasien, kodepoli, totalharga, Dibayar, kembalian) values " & "('" & ComboBox1.Text & "', '" & TextBox1.Text & "', '" & TextBox2.Text & "', '" & TextBox3.Text & "', '" & TextBox4.Text & "', '" & TextBox8.Text & "', '" & TextBox9.Text & "', '" & TextBox10.Text & "')"
        command = New OleDbCommand(insertresep, connection)
        command.ExecuteNonQuery()
        For baris As Integer = 0 To DataGridView1.Rows.Count - 2

            Dim sqlsave As String = "insert into Detail(NomorRsp, KodeObt, Harga, Dosis, SubTotal) values" & "('" & ComboBox1.Text & "', '" & DataGridView1.Rows(baris).Cells(0).Value & "', '" & DataGridView1.Rows(baris).Cells(2).Value & "', '" & DataGridView1.Rows(baris).Cells(3).Value & "', '" & DataGridView1.Rows(baris).Cells(4).Value & "')"
            command = New OleDbCommand(sqlsave, connection)
            command.ExecuteNonQuery()

            'kurangi stok

            command = New OleDbCommand("select * from Obat where KodeObt='" & DataGridView1.Rows(baris).Cells(0).Value & "'", connection)
            read = command.ExecuteReader
            read.Read()
            If read.HasRows Then
                Dim kuranginstok As String = "update Obat set JumlahObt= '" & read.Item(5) - DataGridView1.Rows(baris).Cells(3).Value & "' where KodeObt='" & DataGridView1.Rows(baris).Cells(0).Value & "'"
                command = New OleDbCommand(kuranginstok, connection)
                command.ExecuteNonQuery()
            End If
        Next baris

        Dim simpanbayar As String = "insert into pembayaran(kodebayar, kodepasien, tanggalbayar, jumlahbayar) values('" & ComboBox1.Text & "', '" & TextBox3.Text & "', '" & TextBox1.Text & "', '" & TextBox8.Text & "')"
        command = New OleDbCommand(simpanbayar, connection)
        command.ExecuteNonQuery()

        Dim ubahketerangan As String = "update pendaftaran set keterangan='1' where nomor_pendaftaran='" & ComboBox1.Text & "'"
        command = New OleDbCommand(ubahketerangan, connection)
        command.ExecuteNonQuery()
        DataGridView1.Columns.Clear()
        Call buatkolombaru()
        Call tampiltransaksi()
    End Sub

    Private Sub DataGridView1_CellEndEdit(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellEndEdit
        If e.ColumnIndex = 0 Then
            command = New OleDbCommand("select * from Obat where KodeObt = '" & DataGridView1.Rows(e.RowIndex).Cells(0).Value & "'", connection)
            read = command.ExecuteReader
            read.Read()
            If read.HasRows Then
                DataGridView1.Rows(e.RowIndex).Cells(1).Value = read.Item(1)
                DataGridView1.Rows(e.RowIndex).Cells(2).Value = read.Item(4)
                DataGridView1.Rows(e.RowIndex).Cells(3).Value = 1
                DataGridView1.Rows(e.RowIndex).Cells(4).Value = DataGridView1.Rows(e.RowIndex).Cells(2).Value * DataGridView1.Rows(e.RowIndex).Cells(3).Value
                Call totalitem()
                Call totalharga()

            Else
                MsgBox("Kode obat tidak terdaftar")

            End If
        End If

        If e.ColumnIndex = 3 Then
            command = New OleDbCommand("select * from Obat where KodeObt = '" & DataGridView1.Rows(e.RowIndex).Cells(0).Value & "'", connection)
            read = command.ExecuteReader
            read.Read()
            If read.HasRows Then
                If DataGridView1.Rows(e.RowIndex).Cells(3).Value > read.Item(4) Then
                    MsgBox("Stok Obat hanya ada" & read.Item(4) & "")
                    DataGridView1.Rows(e.RowIndex).Cells(4).Value = DataGridView1.Rows(e.RowIndex).Cells(3).Value
                    Call totalitem()
                    Call totalharga()

                Else
                    DataGridView1.Rows(e.RowIndex).Cells(4).Value = DataGridView1.Rows(e.RowIndex).Cells(2).Value * DataGridView1.Rows(e.RowIndex).Cells(3).Value
                    Call totalitem()
                    Call totalharga()
                End If
            End If
        End If
    End Sub

    Private Sub DataGridView1_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles DataGridView1.KeyPress
        On Error Resume Next
        If e.KeyChar = Chr(27) Then
            DataGridView1.Rows.RemoveAt(DataGridView1.CurrentCell.RowIndex)
        End If
        Call totalitem()
        Call totalharga()
        TextBox9.Text = ""
        TextBox10.Text = ""
        If Not ((e.KeyChar >= "a" And e.KeyChar <= "z") Or e.KeyChar = vbBack) Then
            e.Handled() = False
        End If
    End Sub

    Private Sub TextBox9_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles TextBox9.KeyPress
        If e.KeyChar = Chr(13) Then
            If Val(TextBox9.Text) < Val(TextBox8.Text) Then
                MsgBox("pembayaran anda kurang")
                TextBox10.Text = ""
                TextBox9.Focus()

            Else
                TextBox10.Text = Val(TextBox9.Text) - Val(TextBox8.Text)
                Button1.Focus()
            End If
        End If
    End Sub

End Class