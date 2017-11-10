Imports System.Data.OleDb
Public Class dokter
    Dim connection As OleDbConnection
    Dim dataadapter As OleDbDataAdapter
    Dim dataset As New DataSet
    Dim command As New OleDbCommand
    Dim read As OleDbDataReader
    Dim str As String

    Sub tampil()
        dataadapter = New OleDbDataAdapter("select * from Dokter", connection)
        dataset = New DataSet
        dataset.Clear()
        dataadapter.Fill(dataset, "Dokter")
        DGV.DataSource = (dataset.Tables("Dokter"))
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
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox6.Text = ""
    End Sub
    Sub kodeotomatis()
        comma()
    End Sub
    Private Sub dokter_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        str = "provider=microsoft.ACE.OLEDB.12.0;data source=didindb.accdb"
        connection = New OleDbConnection(str)
        If connection.State = ConnectionState.Closed Then
            connection.Open()
        End If
        tampil()
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Dim dbtambah As String
        If Button2.Enabled = True Then
            dbtambah = "INSERT INTO Dokter values('" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & TextBox4.Text & "','" & TextBox5.Text & "','" & TextBox6.Text & "')"
        Else
            dbtambah = "UPDATE Dokter SET namadokter = '" & TextBox2.Text & "',spesialis = '" & TextBox3.Text & "',alamatdokter = '" & TextBox4.Text & "',telepondokter = '" & TextBox5.Text & "', tarif = '" & TextBox6.Text & "',kdpoli = '" & TextBox7.Text & "' WHERE kd_dokter = '" & TextBox1.Text & "'"
        End If

        command = New OleDbCommand(dbtambah, connection)
        command.ExecuteNonQuery()
        tampil()
        reset()
    End Sub

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        If MsgBox("Yakin akan menghapus data?", MsgBoxStyle.YesNo, _
 "Konfirmasi") = MsgBoxResult.No Then Exit Sub
        Dim dbhapus As String
        dbhapus = "DELETE FROM Dokter WHERE kd_dokter = " & _
              "'" & DGV.CurrentRow.Cells(0).Value & "'"

        command = New OleDbCommand(dbhapus, connection)
        command.ExecuteNonQuery()
        tampil()
        reset()
    End Sub

    Private Sub DGV_CellMouseClick(sender As Object, e As System.Windows.Forms.DataGridViewCellMouseEventArgs) Handles DGV.CellMouseClick
        TextBox1.Text = DGV.CurrentRow.Cells(0).Value
        TextBox2.Text = DGV.CurrentRow.Cells(1).Value
        TextBox3.Text = DGV.CurrentRow.Cells(2).Value
        TextBox4.Text = DGV.CurrentRow.Cells(3).Value
        TextBox5.Text = DGV.CurrentRow.Cells(4).Value
        TextBox6.Text = DGV.CurrentRow.Cells(5).Value
        TextBox7.Text = DGV.CurrentRow.Cells(6).Value
        Button2.Enabled = False
        TextBox1.ReadOnly = True
    End Sub

    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles Button4.Click
        reset()
    End Sub

    Private Sub TextBox8_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        dataadapter = New OleDbDataAdapter("Select * from Dokter Where kd_dokter like '" & Me.TextBox8.Text & "%' or namadokter like '" & Me.TextBox8.Text & "%' or spesialis like '" & Me.TextBox8.Text & "%' or kdpoli like '" & Me.TextBox8.Text & "%'", connection)
        dataset = New DataSet
        dataset.Clear()
        dataadapter.Fill(dataset, "Dokter")
        DGV.DataSource = (dataset.Tables("Dokter"))
        Dim tampil1 As String
        tampil1 = "select * from Dokter where kd_dokter ='" & TextBox1.Text & "'"
        command = New OleDbCommand(tampil1, connection)
        read = command.ExecuteReader

        If read.Read Then
            TextBox1.Text = read.Item("kd_dokter")
            TextBox2.Text = read.Item("namadokter")
            TextBox3.Text = read.Item("spesialis")
            TextBox4.Text = read.Item("alamatdokter")
            TextBox5.Text = read.Item("telepondokter")
            TextBox6.Text = read.Item("tarif")
            TextBox7.Text = read.Item("kdpoli")
            TextBox1.Enabled = False

        End If
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub TextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub TextBox3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox3.TextChanged

    End Sub
End Class