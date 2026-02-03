Option Explicit On

Public Class NCUDataExtractor


    Public Function ExtractNCUData(oSheet As Microsoft.Office.Interop.Excel.Worksheet) As Dictionary(Of String, ExcelData)

        ' 1. Obtener última fila de forma segura
        Dim lastRow As Integer = GetLastRow(oSheet)

        ' Si la hoja no tiene datos suficientes (encabezado en filas 1 y 2), salimos temprano
        Dim oDic As New Dictionary(Of String, ExcelData)

        ' 2. Recorrido de filas desde la 3 hasta el final
        Console.WriteLine(">>> Extracting data from Excel...")
        For i As Integer = 3 To lastRow
            Dim cellKey = oSheet.Cells(i, 4).Text   ' Clave única: Part Number de referencia (Columna D)
            Dim key As String = If(cellKey IsNot Nothing, cellKey.ToString().Trim(), "")
            If Not String.IsNullOrWhiteSpace(key) AndAlso Not oDic.ContainsKey(key) Then    ' Validamos clave no vacía y no duplicada en el dic
                Dim oExcelData As New ExcelData With {
                    .DescriptionRef = oSheet.Cells(i, 3).Text.ToString(),
                    .Source = 2,
                    .Nomenclature = oSheet.Cells(i, 5).Text.ToString(),
                    .Definition = oSheet.Cells(i, 5).Text.ToString()
                }
                oDic.Add(key, oExcelData)
            End If
        Next
        Return oDic
    End Function

    Private Function GetLastRow(oSheet As Microsoft.Office.Interop.Excel.Worksheet) As Integer
        Try
            Dim lastCell = oSheet.Cells.Find("*", , , ,
                Microsoft.Office.Interop.Excel.XlSearchOrder.xlByRows,
               Microsoft.Office.Interop.Excel.XlSearchDirection.xlPrevious)
            If lastCell Is Nothing Then Return 0
            Return lastCell.Row
        Catch
            Return 0
        End Try
    End Function

End Class
