Option Explicit On


Module Program

    Sub Main()
        Console.WriteLine(">>> Starting Process...")
        Console.WriteLine("------------------------------------------------")

        Dim oExcelSession As New ExcelSession()
        Const SourcePath As String = "D:\OneDrive\_CATIA\_V5R21-DLN\NCU\CATALOGO-NCU.xlsx"
        Dim oNCUSheet As Microsoft.Office.Interop.Excel.Worksheet = oExcelSession.GetNCUSheet(SourcePath)
        Dim oActiveWorkbook As Microsoft.Office.Interop.Excel.Workbook = oExcelSession.GetActiveWorkbook()

        ' 3. VALIDACIÓN CRÍTICA: ¿Está todo listo?
        If Not oExcelSession.IsReady OrElse oNCUSheet Is Nothing Then
            Console.WriteLine("!!! ABORTING: Excel Session is not ready.")
            Console.WriteLine(oExcelSession.ErrorMessage)
            Exit Sub
        End If

        ' Aquí oActiveWorkbook NO es Nothing
        Dim oActivesheet As Microsoft.Office.Interop.Excel.Worksheet = oExcelSession.ActiveSheet

        Console.WriteLine("------------------------------------------------")
        Console.WriteLine($">>> Active Workbook Name: {oActiveWorkbook.Name}")

        ' Extracción y Cierre
        Dim oNCUDataExtractor As New NCUDataExtractor()
        Dim oNCUDic As Dictionary(Of String, ExcelData) = oNCUDataExtractor.ExtractNCUData(oNCUSheet)
        oExcelSession.CloseNCU()

        Console.WriteLine($">>> NCU Data Extracted: {oNCUDic.Count} items.")

        ' Inyección
        Dim oNCUDataInjector As New NCUDataInjector()
        oNCUDataInjector.InjectNCUDataToExcel(oActivesheet, oNCUDic)

        Console.WriteLine(">>> NCU Data Injection Completed.")
    End Sub

End Module
