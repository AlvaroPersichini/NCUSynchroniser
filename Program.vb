

Module Program

    Sub Main()


        Dim xlSession As New ExcelSession()
        If Not xlSession.IsReady Then
            Console.WriteLine(xlSession.ErrorMessage)
            Return
        End If


        ' NCU Data Extraction
        Const SourcePath As String = "D:\OneDrive\_CATIA\_V5R21-DLN\NCU\CATALOGO-NCU.xlsx"
        Dim oNCUSheet As Microsoft.Office.Interop.Excel.Worksheet = xlSession.GetNCUSheet(SourcePath)
        Dim oNCUDataExtractor As New NCUDataExtractor()
        Dim oNCUDic As Dictionary(Of String, ExcelData) = oNCUDataExtractor.ExtractNCUData(oNCUSheet)
        xlSession.CloseNCU()
        Console.WriteLine($">>> NCU Data Extracted: {oNCUDic.Count} items.")




        ' Active Sheet 
        Dim oActiveSheet As Microsoft.Office.Interop.Excel.Worksheet = xlSession.GetActiveSheet()
        If oActiveSheet Is Nothing Then
            Console.WriteLine(">>> [ERROR] No hay una hoja activa en el libro de Excel.")
            Return
        End If
        Console.WriteLine($">>> Active Sheet Name: {oActiveSheet.Name}")




        ' Data Injection




    End Sub

End Module
