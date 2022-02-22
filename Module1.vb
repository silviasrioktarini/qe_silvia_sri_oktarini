Imports System.Data.OleDb
Imports System.Data
Imports System.Data.Odbc
Module Module1
    Public ds As DataSet
    Public da As OleDbDataAdapter
    Public OLECMD As OleDbCommand
    Public OLERDR As OleDbDataReader
    Public CNN As OleDbConnection
    Public dread As OleDbDataReader
    Public KONEKSI As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\G_Hasna Ardhya Ningrum_122190147_TUGAS MODUL IV.accdb"
    Public X As Integer
    Public LOKASI As String
    Public Sub KoneksiUser()
        LOKASI = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Application.StartupPath & "\G_Hasna Ardhya Ningrum_122190147_TUGAS MODUL IV.accdb;"
        CNN = New OleDbConnection(LOKASI)
        If CNN.State = ConnectionState.Closed Then CNN.Open()
    End Sub
End Module
