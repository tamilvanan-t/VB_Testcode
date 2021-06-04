Imports System.IO
Imports System.Data
Imports System.Collections.Generic
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Spreadsheet
Public Class ReadExcel
    Public Function ImportExcel(filePath As String) As String()
        Dim outputArray() As String = New String(1) {"0", "0"}
        'Open the Excel file in Read Mode using OpenXml.
        Using doc As SpreadsheetDocument = SpreadsheetDocument.Open(filePath, False)
            'Read the first Sheet from Excel file.
            Dim sheet As Sheet = doc.WorkbookPart.Workbook.Sheets.GetFirstChild(Of Sheet)()

            'Get the Worksheet instance.
            Dim worksheet As Worksheet = TryCast(doc.WorkbookPart.GetPartById(sheet.Id.Value), WorksheetPart).Worksheet

            'Fetch all the rows present in the Worksheet.
            Dim rows As IEnumerable(Of Row) = worksheet.GetFirstChild(Of SheetData)().Descendants(Of Row)()

            Dim plannedPages As String
            Dim exportedPages As String

            plannedPages = GetValueOfCell(3, 6, rows, doc)
            exportedPages = GetValueOfCell(1, 11, rows, doc)
            outputArray(0) = plannedPages
            outputArray(1) = exportedPages

            Console.WriteLine(plannedPages)
            Console.WriteLine(exportedPages)

            'GridView1.DataSource = dt
            'GridView1.DataBind()
        End Using
        Return outputArray
    End Function

    Private Function GetValueOfCell(column As String, rowNumber As Integer, rows As IEnumerable(Of Row), doc As SpreadsheetDocument) As String
        'Loop through the Worksheet rows.
        Dim value As String
        value = ""
        Dim rowCount As Integer
        rowCount = 0
        For Each row As Row In rows
            'Use the first row to add columns to DataTable.
            If rowCount = rowNumber Then
                Console.WriteLine(rowCount)
                Dim colCount As Integer
                colCount = 0
                For Each cell As Cell In row.Descendants(Of Cell)()
                    Console.WriteLine(colCount)
                    If colCount = column Then

                        value = GetValue(doc, cell)
                    End If
                    colCount = colCount + 1
                Next
            End If
            rowCount = rowCount + 1
        Next
        Return value
    End Function

    Private Function GetValue(doc As SpreadsheetDocument, cell As Cell) As String
        Dim value As String = cell.CellValue.InnerText
        If cell.DataType IsNot Nothing AndAlso cell.DataType.Value = CellValues.SharedString Then
            Return doc.WorkbookPart.SharedStringTablePart.SharedStringTable.ChildElements.GetItem(Integer.Parse(value)).InnerText
        End If
        Return value
    End Function
End Class