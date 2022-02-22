Imports System.Data.OleDb
Imports System.Data
Imports System.Data.Odbc
Public Class Form2
    Sub DataPenjualan()
        LOKASI = KONEKSI
        CNN = New OleDbConnection(LOKASI)
        If CNN.State = ConnectionState.Closed Then CNN.Open()
        da = New OleDbDataAdapter("Select * From Data_Penjualan", CNN)
        ds = New DataSet
        ds.Clear()
        da.Fill(ds, "Data_Penjualan")
        DataGridView1.DataSource = (ds.Tables("Data_Penjualan"))
    End Sub
    Sub DataPersediaanBarangToko()
        LOKASI = KONEKSI
        CNN = New OleDbConnection(LOKASI)
        If CNN.State = ConnectionState.Closed Then CNN.Open()
        da = New OleDbDataAdapter("Select * From Data_Persediaan_Barang_Toko", CNN)
        ds = New DataSet
        ds.Clear()
        da.Fill(ds, "Data_Persediaan_Barang_Toko")
        DataGridView2.DataSource =
        (ds.Tables("Data_Persediaan_Barang_Toko"))
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim A, B, Total
        A = NumericUpDown1.Value
        B = TextBox4.Text
        Total = A * B
        TextBox3.Text = Total
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Call DataPenjualan()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Call DataPersediaanBarangToko()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim karakter As String = "0123456789"
        Dim charact As New Random
        TextBox1.Text = karakter(charact.Next(karakter.Length)) & karakter(charact.Next(karakter.Length)) & karakter(charact.Next(karakter.Length)) & karakter(charact.Next(karakter.Length)) & karakter(charact.Next(karakter.Length)) & karakter(charact.Next(karakter.Length)) & karakter(charact.Next(karakter.Length))
        CNN = New OleDbConnection(KONEKSI)
        Try
            If CNN.State <> ConnectionState.Closed Then CNN.Close()
            CNN.Open()
            OLECMD = New OleDbCommand("insert into Data_Penjualan(No_transaksi,No_id,Nama_barang,Jumlah_barang,Harga,Total_Harga) values('" & TextBox1.Text & "','" & TextBox5.Text & "','" & TextBox2.Text & "','" & NumericUpDown1.Value & "','" & TextBox4.Text & "','" & TextBox3.Text & "')", CNN)
            X = OLECMD.ExecuteNonQuery
            If X = 1 Then
                MessageBox.Show("Terinput", "Input", MessageBoxButtons.OK,
                MessageBoxIcon.Information)
            Else
                MessageBox.Show("Gagal Input", "Cannot Input",
                MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
            DataGridView1.Refresh()
            Call DataPenjualan()
            CNN.Close()
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message)
        End Try
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
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
            OLECMD.CommandText = "DELETE from Data_Penjualan where No_transaksi='" & Me.DataGridView1.CurrentRow.Cells(0).Value & "'"
            OLECMD.ExecuteNonQuery()
        Finally
            CNN.Close()
        End Try
        Me.DataGridView1.Rows.Remove(Me.DataGridView1.CurrentRow)
        Call DataPenjualan()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Form3.Show()
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Me.Close()
    End Sub
End Class