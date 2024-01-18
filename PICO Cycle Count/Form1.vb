Imports System.ComponentModel.Design
Imports System.Data.OleDb
Imports System.Globalization
Imports System.Web.UI.WebControls
Imports OfficeOpenXml
Imports OfficeOpenXml.Style

Public Class Form1

    Public ConversionVal As Decimal
    Public WeightVal As Decimal
    Public Eq As Decimal

    Public Spool As Decimal
    Public WTperSpool As Decimal
    Public TotalSpoolWT As Decimal
    Public Multiplier As Decimal
    Public Meter As Decimal
    Public Eq2 As Decimal

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        lblMaterialType.Text = ""
        lblValue.Text = ""

        txtSpool.Visible = False
        Label5.Visible = False

        'SerialPort1.Open()
    End Sub

    'I:\Dept_Tech_Support\LF Software\LF Database

    Public Sub Get_db()
        Dim con As OleDbConnection = New OleDbConnection
        con.ConnectionString = ("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\LF Database\PICO Cycle Count_db.accdb;Jet OLEDB:Database Password=lfcycle")
        Try
            Dim MyData As String
            Dim cmd As New OleDbCommand
            Dim Data As New DataTable
            Dim adap As New OleDbDataAdapter
            con.Open()

            MyData = "SELECT * From CycleCount_tb WHERE Part_Number = '" & txtPartNumber.Text & "'"
            cmd.Connection = con
            cmd.CommandText = MyData
            adap.SelectCommand = cmd

            adap.Fill(Data)

            If Data.Rows.Count > 0 Then

                lblMaterialType.Text = Data.Rows(0).Item("Material_Type").ToString
                lblValue.Text = Data.Rows(0).Item("Uom").ToString
                ConversionVal = Data.Rows(0).Item("Conversion_Factor").ToString
                WTperSpool = Data.Rows(0).Item("WT_Per_Spool").ToString
                Multiplier = Data.Rows(0).Item("Multiplier").ToString
                Console.WriteLine(ConversionVal)
                Console.WriteLine(WTperSpool)
                Console.WriteLine(Multiplier)
                txtWeight.Focus()
                If WTperSpool > 0 Then
                    txtSpool.Visible = True
                    Label5.Visible = True
                End If
            Else
                MsgBox("Part number does not exist in the database.", MessageBoxIcon.Error)
                txtPartNumber.Text = ""
                lblMaterialType.Text = ""
                lblValue.Text = ""
            End If
        Catch ex As Exception
            MsgBox(ex.Message, vbCritical)
        Finally
            con.Close()
        End Try
    End Sub

    Public Sub Record()
        Dim partnum, wt, finalqty, _date, _time, unit, material As String
        Dim command As String

        Dim currentDate As Date = DateTime.Now
        Dim weekNumber As Integer = CultureInfo.CurrentCulture.Calendar.GetWeekOfYear(currentDate, CalendarWeekRule.FirstDay, DayOfWeek.Sunday)

        material = lblMaterialType.Text
        partnum = txtPartNumber.Text
        wt = txtWeight.Text
        finalqty = Eq
        unit = lblValue.Text
        _time = Format(Now, "hh:mm tt")
        _date = Format(Now, "MMM dd, yyyy")
        Dim con As OleDbConnection = New OleDbConnection
        con.ConnectionString = ("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\LF Database\PICO Cycle Count_db.accdb;Jet OLEDB:Database Password=lfcycle")
        Try
            con.Open()
            command = "INSERT INTO [Report_tb] ([Material_Type],[Part_Number],[Weight],[Final_Qty],[UoM],[_time],[_date],[_weeknum]) VALUES (@material,@partnum, @wt, @finalqty, @unit, @_time, @_date, @weekNumber)"
            Using cmd3 As OleDbCommand = New OleDbCommand(command, con)
                cmd3.Parameters.AddWithValue("@material", material)
                cmd3.Parameters.AddWithValue("@partnum", partnum)
                cmd3.Parameters.AddWithValue("@title", wt)
                cmd3.Parameters.AddWithValue("@finalqty", finalqty)
                cmd3.Parameters.AddWithValue("@unit", unit)
                cmd3.Parameters.AddWithValue("@mytime", _time)
                cmd3.Parameters.AddWithValue("@_date", _date)
                cmd3.Parameters.AddWithValue("@weekNumber", weekNumber)
                cmd3.ExecuteNonQuery()
            End Using
            con.Close()
        Catch ex As Exception
            MsgBox(ex.Message, vbCritical)
        End Try
    End Sub

    Public Sub Record2()
        Dim partnum, wt, finalqty, _date, _time, unit, material As String
        Dim command As String

        Dim currentDate As Date = DateTime.Now
        Dim weekNumber As Integer = CultureInfo.CurrentCulture.Calendar.GetWeekOfYear(currentDate, CalendarWeekRule.FirstDay, DayOfWeek.Sunday)

        material = lblMaterialType.Text
        partnum = txtPartNumber.Text
        wt = txtWeight.Text
        finalqty = Eq2
        unit = lblValue.Text
        _time = Format(Now, "hh:mm tt")
        _date = Format(Now, "MMM dd, yyyy")
        Dim con As OleDbConnection = New OleDbConnection
        con.ConnectionString = ("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\LF Database\PICO Cycle Count_db.accdb;Jet OLEDB:Database Password=lfcycle")
        Try
            con.Open()
            command = "INSERT INTO [Report_tb] ([Material_Type],[Part_Number],[Weight],[Final_Qty],[UoM],[_time],[_date],[_weeknum]) VALUES (@material, @partnum, @wt, @finalqty, @unit, @_time, @_date, @weekNumber)"
            Using cmd3 As OleDbCommand = New OleDbCommand(command, con)
                cmd3.Parameters.AddWithValue("@material", material)
                cmd3.Parameters.AddWithValue("@partnum", partnum)
                cmd3.Parameters.AddWithValue("@title", wt)
                cmd3.Parameters.AddWithValue("@finalqty", finalqty)
                cmd3.Parameters.AddWithValue("@unit", unit)
                cmd3.Parameters.AddWithValue("@mytime", _time)
                cmd3.Parameters.AddWithValue("@_date", _date)
                cmd3.Parameters.AddWithValue("@weekNumber", weekNumber)
                cmd3.ExecuteNonQuery()
            End Using
            con.Close()
        Catch ex As Exception
            MsgBox(ex.Message, vbCritical)
        End Try
    End Sub

    Public Sub Get_FirstDate()
        Dim connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\LF Database\PICO Cycle Count_db.accdb;Jet OLEDB:Database Password=lfcycle"
        Dim query As String = "SELECT TOP 1 [_date] FROM Report_tb ORDER BY [_date] ASC"

        Using connection As New OleDbConnection(connectionString)
            connection.Open()
            Using command As New OleDbCommand(query, connection)
                Dim result As Object = command.ExecuteScalar()

                If result IsNot Nothing AndAlso Not DBNull.Value.Equals(result) Then
                    Dim firsdate As DateTime = Convert.ToDateTime(result)
                    Console.WriteLine("First date: " & firsdate.ToString("MMM dd, yyyy"))
                Else
                    Console.WriteLine("No data found.")
                End If
            End Using
        End Using
    End Sub

    Public Sub Get_LastDate()
        Dim connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\LF Database\PICO Cycle Count_db.accdb;Jet OLEDB:Database Password=lfcycle"
        Dim query As String = "SELECT TOP 1 [_date] FROM Report_tb ORDER BY [_date] DESC"

        Using connection As New OleDbConnection(connectionString)
            connection.Open()
            Using command As New OleDbCommand(query, connection)
                Dim result As Object = command.ExecuteScalar()

                If result IsNot Nothing AndAlso Not DBNull.Value.Equals(result) Then
                    Dim firsdate As DateTime = Convert.ToDateTime(result)
                    Console.WriteLine("First date: " & firsdate.ToString("MMM dd, yyyy"))
                Else
                    Console.WriteLine("No data found.")
                End If
            End Using
        End Using
    End Sub

    '************************************** GET The lowest work week **************************************
    Public FirstWW As String
    Public Sub Get_MinWW()
        Dim connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\LF Database\PICO Cycle Count_db.accdb;Jet OLEDB:Database Password=lfcycle"


        Using connection As New OleDbConnection(connectionString)
            Try

                connection.Open()

                Dim query As String = "SELECT MIN([_weeknum]) AS FirstWorkWeek FROM Report_tb"
                Using command As New OleDbCommand(query, connection)
                    Dim result As Object = command.ExecuteScalar()

                    If result IsNot Nothing AndAlso Not DBNull.Value.Equals(result) Then
                        Dim firstWorkWeek As Integer = Convert.ToInt32(result)
                        FirstWW = firstWorkWeek
                        'MessageBox.Show($"The first work week number is: {FirstWW}")
                    Else
                        MessageBox.Show("No data found.")
                    End If
                End Using
            Catch ex As Exception
                MessageBox.Show($"Error: {ex.Message}")
            End Try
        End Using
    End Sub

    '************************************** GET The Highest work week **************************************
    Public LastWW As String
    Public Sub Get_MaxWW()
        Dim connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\LF Database\PICO Cycle Count_db.accdb;Jet OLEDB:Database Password=lfcycle"


        Using connection As New OleDbConnection(connectionString)
            Try

                connection.Open()

                Dim query As String = "SELECT MAX([_weeknum]) AS LastWorkWeek FROM Report_tb"
                Using command As New OleDbCommand(query, connection)
                    Dim result As Object = command.ExecuteScalar()

                    If result IsNot Nothing AndAlso Not DBNull.Value.Equals(result) Then
                        Dim lastWorkWeek As Integer = Convert.ToInt32(result)

                        LastWW = lastWorkWeek - 1
                        'MessageBox.Show($"The Last work week number is: {LastWW}")
                    Else
                        MessageBox.Show("No data found.")
                    End If
                End Using
            Catch ex As Exception
                MessageBox.Show($"Error: {ex.Message}")
            End Try
        End Using
    End Sub
    Private Sub txtPartNumber_KeyUp(sender As Object, e As KeyEventArgs) Handles txtPartNumber.KeyUp
        If e.KeyCode = Keys.Enter Then
            Get_db()
        End If
    End Sub

    Private Sub btnSubmit_Click(sender As Object, e As EventArgs) Handles btnSubmit.Click
        If txtPartNumber.Text = "" Then
            MsgBox("Please enter part number!", MessageBoxIcon.Error)
            txtPartNumber.Focus()
        ElseIf txtWeight.Text = "" Then
            MsgBox("Please enter weight!", MessageBoxIcon.Error)
            txtWeight.Focus()
        Else
            If WTperSpool > 0 Then
                If txtSpool.Text = "" Then
                    MsgBox("Please enter number of spool!", MessageBoxIcon.Error)
                    txtSpool.Focus()
                Else

                    WeightVal = CDec(txtWeight.Text)
                    TotalSpoolWT = CDec(txtSpool.Text) * WTperSpool

                    Eq2 = Math.Round((WeightVal - TotalSpoolWT) * Multiplier, 4)
                    If Eq2 < 0 Then
                        MsgBox("Weight in grams cannot be lower than spool weight!", MessageBoxIcon.Error)
                        txtWeight.Text = ""
                        txtSpool.Text = ""
                        txtWeight.Focus()
                    Else

                        'Get_WW()
                        Record2()
                        MsgBox("submission completed")
                        txtSpool.Text = ""
                        txtSpool.Visible = False
                        Label5.Visible = False

                        txtPartNumber.Text = ""
                        txtWeight.Text = ""
                        lblMaterialType.Text = ""
                        lblValue.Text = ""
                    End If
                End If
            Else
                WeightVal = CDec(txtWeight.Text)
                Eq = Math.Round(WeightVal * ConversionVal, 4)
                'Get_WW()
                Record()
                MsgBox("submission completed")
                txtPartNumber.Text = ""
                txtWeight.Text = ""
                lblMaterialType.Text = ""
                lblValue.Text = ""
            End If

        End If
    End Sub

    '************************************** GET ALL PART NUMBER FROM REPORT **************************************
    'Public PartnumData As String

    'Public Sub GetPartnum()
    '    Dim connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\LF Database\PICO Cycle Count_db.accdb;Jet OLEDB:Database Password=lfcycle"
    '    Dim query As String = "SELECT Part_Number FROM Report_tb;"

    '    Using connection As New OleDbConnection(connectionString)
    '        connection.Open()

    '        Using command As New OleDbCommand(query, connection)
    '            Using reader As OleDbDataReader = command.ExecuteReader()
    '                If reader.HasRows Then
    '                    While reader.Read()
    '                        Dim partnum As Object = reader.GetValue(0)
    '                        If Not DBNull.Value.Equals(partnum) Then
    '                            PartnumData = partnum
    '                            'Console.WriteLine(PartnumData)
    '                        Else
    '                            Console.WriteLine("No data.")
    '                        End If
    '                    End While
    '                Else
    '                    Console.WriteLine("No data found.")
    '                End If
    '            End Using
    '        End Using
    '    End Using
    'End Sub


    Public PartnumData As New List(Of String)
    Public WeightData As New List(Of String)
    Public FinalqtyData As New List(Of String)
    Public UoMData As New List(Of String)
    Public MaterialData As New List(Of String)
    Public Sub GetReports()
        Dim connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\LF Database\PICO Cycle Count_db.accdb;Jet OLEDB:Database Password=lfcycle"
        Dim query As String = "SELECT Material_Type, Part_Number, Weight, Final_Qty, UoM FROM Report_tb;"

        Using connection As New OleDbConnection(connectionString)
            connection.Open()

            Using command As New OleDbCommand(query, connection)
                Using reader As OleDbDataReader = command.ExecuteReader()
                    If reader.HasRows Then
                        While reader.Read()
                            Dim mat As Object = reader.GetValue(0)
                            Dim partnum As Object = reader.GetValue(1)
                            Dim weight As Object = reader.GetValue(2)
                            Dim finalQty As Object = reader.GetValue(3)
                            Dim uomdt As Object = reader.GetValue(4)

                            If Not DBNull.Value.Equals(mat) Then
                                MaterialData.Add(mat.ToString())
                            Else
                                Console.WriteLine("No data.")
                            End If

                            If Not DBNull.Value.Equals(partnum) Then
                                PartnumData.Add(partnum.ToString())
                            Else
                                Console.WriteLine("No data.")
                            End If

                            If Not DBNull.Value.Equals(weight) Then
                                WeightData.Add(weight.ToString())
                            Else
                                Console.WriteLine("No data.")
                            End If

                            If Not DBNull.Value.Equals(finalQty) Then
                                FinalqtyData.Add(finalQty.ToString())
                            Else
                                Console.WriteLine("No data.")
                            End If

                            If Not DBNull.Value.Equals(uomdt) Then
                                UoMData.Add(uomdt.ToString())
                            Else
                                Console.WriteLine("No data.")
                            End If

                        End While
                    Else
                        Console.WriteLine("No data found.")
                    End If
                End Using
            End Using
        End Using
    End Sub


    'yyyy/MMM dd_HHmmtt 
    Public dateNtime As String '= DateTime.Now.ToString("yyyy/MMM dd")
    Public get_message As String

    '*********************** A1:I1 **************************
    Public Sub SaveReport()
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial
        dateNtime = DateTime.Now.ToString("MMM dd yyyy - HH_mmtt ")
        Dim dateNtime2 = DateTime.Now.ToString("MMMM dd yyyy")
        Dim get_FolderPath As String = "C:\PICO Cycle Count Report\PICO Cycle Count. " & dateNtime & ".xlsx"

        Try
            Using package As New ExcelPackage()
                Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets.Add("PICO Cycle Count Report")

                worksheet.View.PageLayoutView = True
                Dim reportHeader As ExcelRange = worksheet.Cells("A1:I1")
                reportHeader.Merge = True
                reportHeader.Value = "PICO Cycle Count Report"
                reportHeader.Style.Font.Bold = True
                reportHeader.Style.Font.Size = 18
                reportHeader.Style.Font.Name = "Arial"
                reportHeader.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center

                ' Merge cells for "DATE OF CYCLECOUNT:" and "PREPARED BY"
                worksheet.Cells("A2:C2").Merge = True
                worksheet.Cells("D2:E2").Merge = True
                worksheet.Cells("F2:G2").Merge = True
                worksheet.Cells("H2:I2").Merge = True

                worksheet.Cells("A2:C2").Style.Font.Bold = True
                'worksheet.Cells("A2:C2").Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                'Dim cyclecount As ExcelRange = worksheet.Cells("A2:C2")
                worksheet.Cells("A2:C2").Value = "DATE OF CYCLECOUNT:"
                worksheet.Cells("D2:E2").Value = dateNtime2

                worksheet.Cells("F2:G2").Style.Font.Bold = True
                worksheet.Cells("F2:G2").Value = "PREPARED BY"

                ' Merge cells for "COVERED WEEK:" and "Signature"
                worksheet.Cells("A3:C3").Merge = True
                worksheet.Cells("D3:E3").Merge = True
                worksheet.Cells("F3:G3").Merge = True
                worksheet.Cells("H3:I3").Merge = True

                worksheet.Cells("A3:C3").Style.Font.Bold = True
                worksheet.Cells("A3:C3").Value = "COVERED WEEK:"
                worksheet.Cells("D3:E3").Value = "WW" & LastWW  '"WW" & FirstWW & " - " & "WW" & LastWW

                worksheet.Cells("F3:G3").Style.Font.Bold = True
                worksheet.Cells("F3:G3").Value = "SIGNATURE"

                ' Merge cells for headers
                worksheet.Cells("A5:B5").Merge = True
                worksheet.Cells("C5:D5").Merge = True
                worksheet.Cells("E5:F5").Merge = True
                worksheet.Cells("G5:H5").Merge = True
                'worksheet.Cells("I5:J5").Merge = True

                worksheet.Cells("A5:J5").Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                worksheet.Cells("A5:J5").Style.Font.Bold = True

                worksheet.Cells("A5:B5").Value = "MATERIAL TYPE"

                worksheet.Cells("C5:D5").Value = "PART NUMBER"

                worksheet.Cells("E5:F5").Value = "WT IN GRAMS"

                worksheet.Cells("G5:H5").Value = "FINAL QTY"

                worksheet.Cells("I5").Value = "UoM"

                For i As Integer = 0 To PartnumData.Count - 1
                    Dim mat As String = MaterialData(i)
                    Dim partnum As String = PartnumData(i)
                    Dim weight As String = WeightData(i)
                    Dim finalQty As String = FinalqtyData(i)
                    Dim uom As String = UoMData(i)

                    ' Populate the Excel worksheet
                    Dim row = i + 6
                    worksheet.Cells(row, 1).Value = mat
                    worksheet.Cells(row, 3).Value = partnum
                    worksheet.Cells(row, 5).Value = weight
                    worksheet.Cells(row, 7).Value = finalQty
                    worksheet.Cells(row, 9).Value = uom

                    ' Merge cells for "ABC", "DE", "FGH", "IJ" data
                    worksheet.Cells(row, 1, row, 2).Merge = True
                    worksheet.Cells(row, 3, row, 4).Merge = True
                    worksheet.Cells(row, 5, row, 6).Merge = True
                    worksheet.Cells(row, 7, row, 8).Merge = True
                    'worksheet.Cells(row, 10, row, 10).Merge = True

                    worksheet.Cells(row, 1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                    worksheet.Cells(row, 3).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                    worksheet.Cells(row, 5).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                    worksheet.Cells(row, 7).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                    worksheet.Cells(row, 9).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                Next


                ' Protect workbook structure (prevent adding/removing sheets, etc.)
                worksheet.Protection.IsProtected = True
                worksheet.Protection.SetPassword("PICOprotect") ' Set a password if desired

                ' Protect workbook window (prevent resizing, moving, etc.)
                worksheet.Protection.AllowFormatColumns = False
                worksheet.Protection.AllowFormatRows = False
                worksheet.Protection.AllowInsertColumns = False
                worksheet.Protection.AllowInsertRows = False
                worksheet.Protection.AllowSort = False
                worksheet.Protection.AllowAutoFilter = False
                worksheet.Protection.AllowDeleteColumns = False
                worksheet.Protection.AllowDeleteRows = False

                ' Save the Excel package to a file
                package.SaveAs(New System.IO.FileInfo(get_FolderPath))
            End Using

            MessageBox.Show("The data is saved in " & get_FolderPath)

            PartnumData.Clear()
            WeightData.Clear()
            FinalqtyData.Clear()
            UoMData.Clear()

        Catch ex As Exception
            MsgBox(ex.Message, vbCritical)
        End Try
    End Sub

    '*********************** A1:J1 **************************
    'Public Sub SaveReport()
    '    ExcelPackage.LicenseContext = LicenseContext.NonCommercial
    '    dateNtime = DateTime.Now.ToString("MMM dd yyyy - HH_mmtt ")
    '    Dim dateNtime2 = DateTime.Now.ToString("MMMM dd yyyy")
    '    Dim get_FolderPath As String = "C:\PICO Cycle Count Report\PICO Cycle Count. " & dateNtime & ".xlsx"

    '    Try
    '        Using package As New ExcelPackage()
    '            Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets.Add("PICO Cycle Count Report")

    '            worksheet.View.PageLayoutView = True
    '            Dim reportHeader As ExcelRange = worksheet.Cells("A1:J1")
    '            reportHeader.Merge = True
    '            reportHeader.Value = "PICO Cycle Count Report"
    '            reportHeader.Style.Font.Bold = True
    '            reportHeader.Style.Font.Size = 18
    '            reportHeader.Style.Font.Name = "Arial"
    '            reportHeader.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center

    '            ' Merge cells for "DATE OF CYCLECOUNT:" and "PREPARED BY"
    '            worksheet.Cells("A2:C2").Merge = True
    '            worksheet.Cells("D2:F2").Merge = True
    '            worksheet.Cells("G2:H2").Merge = True
    '            worksheet.Cells("I2:J2").Merge = True

    '            worksheet.Cells("A2:C2").Style.Font.Bold = True
    '            'worksheet.Cells("A2:C2").Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
    '            'Dim cyclecount As ExcelRange = worksheet.Cells("A2:C2")
    '            worksheet.Cells("A2:C2").Value = "DATE OF CYCLECOUNT:"
    '            worksheet.Cells("D2:F2").Value = dateNtime2

    '            worksheet.Cells("G2:H2").Style.Font.Bold = True
    '            worksheet.Cells("G2:H2").Value = "PREPARED BY"

    '            ' Merge cells for "COVERED WEEK:" and "Signature"
    '            worksheet.Cells("A3:C3").Merge = True
    '            worksheet.Cells("D3:F3").Merge = True
    '            worksheet.Cells("G3:H3").Merge = True
    '            worksheet.Cells("I3:J3").Merge = True

    '            worksheet.Cells("A3:C3").Style.Font.Bold = True
    '            worksheet.Cells("A3:C3").Value = "COVERED WEEK:"
    '            worksheet.Cells("D3:F3").Value = "WW" & FirstWW & " - " & "WW" & LastWW

    '            worksheet.Cells("G3:H3").Style.Font.Bold = True
    '            worksheet.Cells("G3:H3").Value = "SIGNATURE"

    '            ' Merge cells for headers
    '            worksheet.Cells("A5:B5").Merge = True
    '            worksheet.Cells("C5:E5").Merge = True
    '            worksheet.Cells("F5:G5").Merge = True
    '            worksheet.Cells("H5:I5").Merge = True
    '            'worksheet.Cells("I5:J5").Merge = True

    '            worksheet.Cells("A5:J5").Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
    '            worksheet.Cells("A5:J5").Style.Font.Bold = True

    '            worksheet.Cells("A5:B5").Value = "MATERIAL TYPE"

    '            worksheet.Cells("C5:E5").Value = "PART NUMBER"

    '            worksheet.Cells("F5:G5").Value = "WT IN GRAMS"

    '            worksheet.Cells("H5:I5").Value = "FINAL QTY"

    '            worksheet.Cells("J5").Value = "UoM"

    '            For i As Integer = 0 To PartnumData.Count - 1
    '                Dim mat As String = MaterialData(i)
    '                Dim partnum As String = PartnumData(i)
    '                Dim weight As String = WeightData(i)
    '                Dim finalQty As String = FinalqtyData(i)
    '                Dim uom As String = UoMData(i)

    '                ' Populate the Excel worksheet
    '                Dim row = i + 6
    '                worksheet.Cells(row, 1).Value = mat
    '                worksheet.Cells(row, 3).Value = partnum
    '                worksheet.Cells(row, 6).Value = weight
    '                worksheet.Cells(row, 8).Value = finalQty
    '                worksheet.Cells(row, 10).Value = uom

    '                ' Merge cells for "ABC", "DE", "FGH", "IJ" data
    '                worksheet.Cells(row, 1, row, 2).Merge = True
    '                worksheet.Cells(row, 3, row, 5).Merge = True
    '                worksheet.Cells(row, 6, row, 7).Merge = True
    '                worksheet.Cells(row, 8, row, 9).Merge = True
    '                'worksheet.Cells(row, 10, row, 10).Merge = True

    '                worksheet.Cells(row, 1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
    '                worksheet.Cells(row, 3).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
    '                worksheet.Cells(row, 6).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
    '                worksheet.Cells(row, 8).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
    '                worksheet.Cells(row, 10).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
    '            Next


    '            ' Protect workbook structure (prevent adding/removing sheets, etc.)
    '            worksheet.Protection.IsProtected = True
    '            worksheet.Protection.SetPassword("PICOprotect") ' Set a password if desired

    '            ' Protect workbook window (prevent resizing, moving, etc.)
    '            worksheet.Protection.AllowFormatColumns = False
    '            worksheet.Protection.AllowFormatRows = False
    '            worksheet.Protection.AllowInsertColumns = False
    '            worksheet.Protection.AllowInsertRows = False
    '            worksheet.Protection.AllowSort = False
    '            worksheet.Protection.AllowAutoFilter = False
    '            worksheet.Protection.AllowDeleteColumns = False
    '            worksheet.Protection.AllowDeleteRows = False

    '            ' Save the Excel package to a file
    '            package.SaveAs(New System.IO.FileInfo(get_FolderPath))
    '        End Using

    '        MessageBox.Show("The data is saved in " & get_FolderPath)

    '        PartnumData.Clear()
    '        WeightData.Clear()
    '        FinalqtyData.Clear()
    '        UoMData.Clear()

    '    Catch ex As Exception
    '        MsgBox(ex.Message, vbCritical)
    '    End Try
    'End Sub

    'Public Sub SaveReport()
    '    dateNtime = DateTime.Now.ToString("MMM dd yyyy - HH_mmtt ")
    '    Dim dateNtime2 = DateTime.Now.ToString("MMMM dd yyyy")

    '    Dim get_FolderPath As String = "C:\PICO Cycle Count Report\PICO Cycle Count. " & dateNtime & ".csv"

    '    get_message = "DATE OF CYCLECOUNT:" & "," & dateNtime2 & "," & "," & "PREPARED BY" & vbCrLf &
    '    "COVERED WEEK:" & "," & FirstWW & "-" & LastWW & "," & "," & "Signature" & "," & vbCrLf & vbCrLf &
    '    "PART NUMBER" & "," & "WT IN GRAMS" & "," & "FINAL QTY" & "," & "UoM" & vbCrLf

    '    Try
    '        ' Assuming you have arrays for WeightData, FinalqtyData, and UoMData
    '        For i As Integer = 0 To PartnumData.Count - 1
    '            Dim partnum As String = "'" & PartnumData(i)
    '            Dim weight As String = WeightData(i)
    '            Dim finalQty As String = FinalqtyData(i)
    '            Dim uomdt As String = UoMData(i)

    '            ' Concatenate the values for each row
    '            get_message &= $"{partnum},{weight},{finalQty},{uomdt}{vbCrLf}"
    '        Next

    '        My.Computer.FileSystem.WriteAllText(get_FolderPath, get_message, False)
    '        MessageBox.Show("The data is saved in " & get_FolderPath)

    '        PartnumData.Clear()
    '        WeightData.Clear()
    '        FinalqtyData.Clear()
    '        UoMData.Clear()

    '    Catch ex As Exception
    '        MsgBox(ex, vbCritical)
    '    End Try
    'End Sub

    Private Sub btnGenerate_Click(sender As Object, e As EventArgs) Handles btnGenerate.Click
        Dim dialog As DialogResult
        dialog = MessageBox.Show("Do you want to generate report?", "PICO Cycle Count", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If dialog = DialogResult.No Then
            Return
        Else
            'GetPartnum()
            GetReports()
            'Get_MinWW()
            Get_MaxWW()

            SaveReport()
            DeleteReport()

        End If
    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        Dim dialog As DialogResult
        dialog = MessageBox.Show("Do you want to clear data?", "PICO Cycle Count", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If dialog = DialogResult.No Then
            Return
        Else
            If lblMaterialType.Text = "Axial Tape" Then
                txtSpool.Text = ""
                txtSpool.Visible = False
                Label5.Visible = False

                txtPartNumber.Text = ""
                txtWeight.Text = ""
                lblMaterialType.Text = ""
                lblValue.Text = ""
                txtPartNumber.Focus()
            Else
                txtPartNumber.Text = ""
                txtWeight.Text = ""
                lblMaterialType.Text = ""
                lblValue.Text = ""
                txtPartNumber.Focus()
            End If

        End If

    End Sub

    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        Dim dialog As DialogResult
        dialog = MessageBox.Show("Do you want to exit?", "PICO Cycle Count", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If dialog = DialogResult.No Then
            e.Cancel = True
        Else
            Application.ExitThread()
        End If

    End Sub

    Private Sub txtWeight_KeyUp(sender As Object, e As KeyEventArgs) Handles txtWeight.KeyUp
        If e.KeyCode = Keys.Enter Then
            If lblMaterialType.Text = "Axial Tape" Then
                txtSpool.Focus()
            End If
        End If
    End Sub

    Public Sub DeleteReport()

        Dim connectionString As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\LF Database\PICO Cycle Count_db.accdb;Jet OLEDB:Database Password=lfcycle"
        Dim connection As New OleDbConnection(connectionString)
        Dim command As String
        Try
            connection.Open()
            command = "DELETE FROM Report_tb"
            Using cmd3 As OleDbCommand = New OleDbCommand(command, connection)
                cmd3.ExecuteNonQuery()
            End Using
            connection.Close()
        Catch ex As Exception
            MsgBox(ex.Message, vbCritical)
        End Try
    End Sub

    Private Sub ReportToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ReportToolStripMenuItem.Click
        Me.Hide()
        Table_Form.ShowDialog()
    End Sub

    Private Sub txtSpool_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtSpool.KeyPress
        If Asc(e.KeyChar) <> 8 Then
            'If Asc(e.KeyChar) <> 45 Then
            If Asc(e.KeyChar) <> 46 Then
                If (Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57) Then
                    e.Handled = True
                    'MessageBox.Show("Please enter numeric value!", "Invalid Entry", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If
            End If
            'End If
        End If
    End Sub

    Private Sub txtWeight_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtWeight.KeyPress
        If Asc(e.KeyChar) <> 8 Then
            'If Asc(e.KeyChar) <> 45 Then
            If Asc(e.KeyChar) <> 46 Then
                If (Asc(e.KeyChar) < 48 Or Asc(e.KeyChar) > 57) Then
                    e.Handled = True
                    'MessageBox.Show("Please enter numeric value!", "Invalid Entry", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End If
            'End If
        End If
    End Sub

    Private Sub SerialPort1_DataReceived(sender As Object, e As IO.Ports.SerialDataReceivedEventArgs) Handles SerialPort1.DataReceived
        'Console.WriteLine(SerialPort1.ReadLine)
        'txtWeight.Text = SerialPort1.ReadLine
    End Sub
End Class
