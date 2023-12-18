Imports System.Data.OleDb

Public Class Table_Form

    Public Sub Table()
        Dim con As OleDbConnection = New OleDbConnection
        con.ConnectionString = ("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\LF Database\PICO Cycle Count_db.accdb;Jet OLEDB:Database Password=lfcycle")

        Try
            Dim Data As New DataTable
            Dim adap As New OleDbDataAdapter
            Dim dataSet As New DataSet()

            Data = New DataTable()

            adap = New OleDbDataAdapter("SELECT Material_Type, Part_Number, Weight, Final_Qty, UoM FROM Report_tb", con)

            con.Open()
            adap.Fill(Data)

            con.Close()

            DataGridView1.DataSource = Data

            DataGridView1.Columns("Material_Type").HeaderText = "Material Type"
            DataGridView1.Columns("Material_Type").HeaderCell.Style.Font = New Font(DataGridView1.Font, FontStyle.Bold)
            DataGridView1.Columns("Material_Type").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            DataGridView1.Columns("Part_Number").HeaderText = "Part Number"
            DataGridView1.Columns("Part_Number").HeaderCell.Style.Font = New Font(DataGridView1.Font, FontStyle.Bold)

            DataGridView1.Columns("Weight").HeaderText = "Weight"
            DataGridView1.Columns("Weight").HeaderCell.Style.Font = New Font(DataGridView1.Font, FontStyle.Bold)
            DataGridView1.Columns("Weight").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            DataGridView1.Columns("Final_Qty").HeaderText = "Final Qty"
            DataGridView1.Columns("Final_Qty").HeaderCell.Style.Font = New Font(DataGridView1.Font, FontStyle.Bold)
            DataGridView1.Columns("Final_Qty").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            DataGridView1.Columns("UoM").HeaderText = "UoM"
            DataGridView1.Columns("UoM").HeaderCell.Style.Font = New Font(DataGridView1.Font, FontStyle.Bold)
            DataGridView1.Columns("UoM").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter


        Catch ex As Exception
            MsgBox(ex.Message, vbCritical)

        End Try
    End Sub
    Private Sub Table_Form_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Table()
    End Sub

    Private Sub Table_Form_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        Form1.Show()
    End Sub
End Class