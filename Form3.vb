Imports System.Data.OleDb
Imports System.Data
Imports System.Data.Odbc
Public Class Form3
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Select Case ComboBox1.Text
            Case "GUDANG"
                Call KoneksiUser()
                OLECMD = New OleDbCommand("select * from Data_Login where No_id='" & TextBox1.Text & "' and User='" & TextBox2.Text & "'and Pass='" & TextBox3.Text & "'", CNN)
                dread = OLECMD.ExecuteReader
                dread.Read()
                If dread.HasRows Then
                    Form1.Show()
                    Me.Hide()
                End If
            Case "PENJUALAN"
                Call KoneksiUser()
                OLECMD = New OleDbCommand("select * from Data_Login where No_id='" & TextBox1.Text & "' and User='" & TextBox2.Text & "'and Pass='" & TextBox3.Text & "'", CNN)
                dread = OLECMD.ExecuteReader
                dread.Read()
                If dread.HasRows Then
                    Form2.Show()
                    Me.Hide()
                End If
            Case "ADMINISTRATOR"
                Call KoneksiUser()
                OLECMD = New OleDbCommand("select * from Data_Login where No_id='" & TextBox1.Text & "' and User='" & TextBox2.Text & "'and Pass='" & TextBox3.Text & "'", CNN)
                dread = OLECMD.ExecuteReader
                dread.Read()
                If dread.HasRows Then
                    Form4.Show()
                    Me.Hide()
                End If
        End Select
    End Sub
End Class