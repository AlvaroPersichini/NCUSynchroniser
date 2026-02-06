Public Class NCUDataInjector

    Sub InjectNCUDataToExcel(oSheet As Microsoft.Office.Interop.Excel.Worksheet, oNCUDic As Dictionary(Of String, ExcelData))


        Dim lastRow As Integer = oSheet.Cells(oSheet.Rows.Count, 1).End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row
        Dim ncuCode As String

        For row As Integer = 3 To lastRow

            ' ncuCode = Convert.ToString(oSheet.Cells(row, 4).Value)
            ncuCode = (oSheet.Cells(row, 4).Value).ToString().Trim()


            If oNCUDic.ContainsKey(ncuCode) Then
                oSheet.Cells(row, 6).Value = oNCUDic(ncuCode).DescriptionRef
                oSheet.Cells(row, 8).Value = 2
                oSheet.Cells(row, 10).Value = oNCUDic(ncuCode).Nomenclature
                oSheet.Cells(row, 11).Value = oNCUDic(ncuCode).Definition

            End If
        Next

        oSheet.Range("D:D, E:E, F:F, H:H, J:J, K:K").Columns.AutoFit()
        oSheet.Columns("E").ColumnWidth = 24

    End Sub

End Class
