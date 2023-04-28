Imports Excel = Microsoft.Office.Interop.Excel
Imports System.IO
Imports System.Text

Public Class Form1
    Private Sub BtnUpload_Click(sender As Object, e As EventArgs) Handles BtnUpload.Click
        Dim openFileDialog1 As New OpenFileDialog()
        openFileDialog1.Filter = "CSV Files (*.csv)|*.csv"
        openFileDialog1.Title = "Select a CSV File"
        If openFileDialog1.ShowDialog() = DialogResult.OK Then
            Dim csvData As String = File.ReadAllText(openFileDialog1.FileName)
            Dim lines() As String = csvData.Split(ControlChars.Lf)
            Dim data As New DataTable()
            Dim headers() As String = lines(0).Split(",")
            For Each header As String In headers
                data.Columns.Add(header, GetType(String))
            Next
            For i As Integer = 1 To lines.Length - 1
                Dim fields() As String = lines(i).Split(",")
                data.Rows.Add(fields)
            Next
            DataGridView1.DataSource = data
        End If
    End Sub


    Private Sub BtnBackup_Click(sender As Object, e As EventArgs) Handles BtnBackup.Click
        Dim saveFileDialog1 As New SaveFileDialog()
        saveFileDialog1.Filter = "CSV Files (*.csv)|*.csv"
        saveFileDialog1.Title = "Save CSV File Backup"
        If saveFileDialog1.ShowDialog() = DialogResult.OK Then
            Dim writer As New StreamWriter(saveFileDialog1.FileName)
            For Each row As DataGridViewRow In DataGridView1.Rows
                Dim cells As List(Of String) = (From cell As DataGridViewCell In row.Cells Select Convert.ToString(cell.Value)).ToList()
                writer.WriteLine(String.Join(",", cells))
            Next
            writer.Flush()
            writer.Close()
            MessageBox.Show("Backup successfully created!", "Backup Status", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub


    Private Sub BtnView_Click(sender As Object, e As EventArgs) Handles BtnView.Click
        Dim openFileDialog1 As New OpenFileDialog()
        openFileDialog1.Filter = "CSV Files (*.csv)|*.csv"
        openFileDialog1.Title = "Select a CSV File to View"
        If openFileDialog1.ShowDialog() = DialogResult.OK Then
            Dim excelApp As New Excel.Application()
            excelApp.Visible = True
            Dim workbook As Excel.Workbook = excelApp.Workbooks.Open(openFileDialog1.FileName)
        End If
    End Sub


    Private Sub BtnNewFile_Click(sender As Object, e As EventArgs) Handles BtnNewFile.Click
        Dim data As New DataTable()
        data.Columns.Add("Column1", GetType(String))
        data.Columns.Add("Column2", GetType(String))
        data.Columns.Add("Column3", GetType(String))
        data.Columns.Add("Column4", GetType(String))
        data.Columns.Add("Column5", GetType(String))
        data.Rows.Add("Value 1", "Value 2", "Value 3", "Value 4", "Value 5")
        data.Rows.Add("Value 6", "Value 7", "Value 8", "Value 9", "Value 10")
        data.Rows.Add("Value 11", "Value 12", "Value 13", "Value 14", "Value 15")
        data.Rows.Add("Value 16", "Value 17", "Value 18", "Value 19", "Value 20")
        data.Rows.Add("Value 21", "Value 22", "Value 23", "Value 24", "Value 25")
        DataGridView1.DataSource = data

        'Prompt the user to save the new file
        Dim saveFileDialog1 As New SaveFileDialog()
        saveFileDialog1.Filter = "CSV Files (*.csv)|*.csv"
        saveFileDialog1.Title = "Save the new CSV file"
        If saveFileDialog1.ShowDialog() = DialogResult.OK Then
            'Save the data to the specified file
            Dim sb As New StringBuilder()
            Dim headers() As String = data.Columns.Cast(Of DataColumn)().Select(Function(c) c.ColumnName).ToArray()
            sb.AppendLine(String.Join(",", headers))
            For Each row As DataRow In data.Rows
                Dim fields() As String = row.ItemArray.Select(Function(f) f.ToString()).ToArray()
                sb.AppendLine(String.Join(",", fields))
            Next
            File.WriteAllText(saveFileDialog1.FileName, sb.ToString())
        End If
    End Sub


    Private Sub BtnPrint_Click(sender As Object, e As EventArgs) Handles BtnPrint.Click
        'Create a new datatable
        Dim myDataTable As New DataTable()

        'Add columns to the datatable
        myDataTable.Columns.Add("Column1", GetType(String))
        myDataTable.Columns.Add("Column2", GetType(Integer))
        myDataTable.Columns.Add("Column3", GetType(Double))

        'Add rows to the datatable
        myDataTable.Rows.Add("Row1", 1, 1.1)
        myDataTable.Rows.Add("Row2", 2, 2.2)
        myDataTable.Rows.Add("Row3", 3, 3.3)

        'Bind the datatable to the DataGrid control
        DataGridView1.DataSource = myDataTable

        'Export the data to a CSV file
        Dim csv As New StringBuilder()

        'Add column headers to the CSV file
        Dim header As String = String.Join(",", myDataTable.Columns.Cast(Of DataColumn).Select(Function(column) column.ColumnName))
        csv.AppendLine(header)

        'Add each row to the CSV file
        For Each row As DataRow In myDataTable.Rows
            Dim line As String = String.Join(",", row.ItemArray.Select(Function(x) x.ToString()))
            csv.AppendLine(line)
        Next

        'Prompt the user to choose a save location and file name
        Dim saveFileDialog1 As New SaveFileDialog()
        saveFileDialog1.Filter = "CSV file (*.csv)|*.csv"
        saveFileDialog1.Title = "Save CSV File As..."
        saveFileDialog1.ShowDialog()

        'Save the CSV file
        If saveFileDialog1.FileName <> "" Then
            File.WriteAllText(saveFileDialog1.FileName, csv.ToString())
            MessageBox.Show("CSV file saved successfully.")
        End If
    End Sub
End Class
