Imports System.Data.OleDb
Imports System.Data
Imports System.Data.Odbc
Public Class Form1
    Sub DataPersediaanBarangGudang()
        LOKASI = KONEKSI
        CNN = New OleDbConnection(LOKASI)
        If CNN.State = ConnectionState.Closed Then CNN.Open()
        da = New OleDbDataAdapter("Select * From Data_Persediaan_Barang_Gudang", CNN)
        ds = New DataSet
        ds.Clear()
        da.Fill(ds, "Data_Persediaan_Barang_Gudang")
        DataGridView1.DataSource = (ds.Tables("Data_Persediaan_Barang_Gudang"))
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Call DataPersediaanBarangGudang()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Me.Close()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Form3.Show()
        Me.Hide()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim karakter As String = "0123456789"
        Dim charact As New Random
        TextBox1.Text = karakter(charact.Next(karakter.Length)) & karakter(charact.Next(karakter.Length)) & karakter(charact.Next(karakter.Length)) & karakter(charact.Next(karakter.Length)) & karakter(charact.Next(karakter.Length)) & karakter(charact.Next(karakter.Length)) & karakter(charact.Next(karakter.Length))
        CNN = New OleDbConnection(KONEKSI)
        Try
            If CNN.State <> ConnectionState.Closed Then CNN.Close()
            CNN.Open()
            OLECMD = New OleDbCommand("insert into Data_Persediaan_Barang_Gudang(Kode_barang,Kode_persediaan,Nama_barang,Jumlah_barang,Jenis_barang) values('" & TextBox1.Text & "','" & TextBox2.Text & "','" & TextBox3.Text & "','" & NumericUpDown1.Value & "','" & ComboBox1.Text & "')", CNN)
            X = OLECMD.ExecuteNonQuery
            If X = 1 Then
                MessageBox.Show("Terinput", "Input", MessageBoxButtons.OK,
                MessageBoxIcon.Information)
            Else
                MessageBox.Show("Gagal Input", "Cannot Input",
                MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
            DataGridView1.Refresh()
            Call DataPersediaanBarangGudang()
            CNN.Close()
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim hasil
        Dim CNN As New OleDbConnection()
        hasil = MsgBox("apakah anda yakin untuk mengahapus data?", vbYesNo + vbQuestion, "Konfirmasi")
        If hasil = vbNo Then
            Exit Sub
        End If
        CNN.ConnectionString = KONEKSI
        CNN.Open()
        Try
            Dim OLECMD As New OleDbCommand()
            OLECMD.Connection = CNN
            OLECMD.CommandText = "DELETE from Data_Persediaan_Barang_Gudang where Kode_barang='" & Me.DataGridView1.CurrentRow.Cells(0).Value & "'"
            OLECMD.ExecuteNonQuery()
        Finally
            CNN.Close()
        End Try
        Me.DataGridView1.Rows.Remove(Me.DataGridView1.CurrentRow)
        Call DataPersediaanBarangGudang()
    End Sub
End Class
