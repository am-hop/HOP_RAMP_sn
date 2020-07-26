Imports System.Data.SQLite

Module ModSettings

    Public con As New SQLiteConnection
    Public cmd As New SQLiteCommand
    Public dr As SQLiteDataReader
    Public da As SQLiteDataAdapter
    Public dt As New DataTable
    Public ds As DataSet
    Public query As String


    Public dbFullpath As String = Application.StartupPath & "\\RAMP.db"
    Public conStr As String = String.Format("Data Source = {0}; Version = 3;", dbFullpath)

    Public Sub OpenCon()
        Try
            con = New SQLiteConnection
            If con.State = ConnectionState.Open Then con.Close()
            con.ConnectionString = conStr
            con.Open()
            'MessageBox.Show("Connection successful")
        Catch ex As Exception
            MessageBox.Show(ex.Message & "OpenCon", "Database Connection Error!", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Module
