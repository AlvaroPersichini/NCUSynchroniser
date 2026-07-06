Option Explicit On
Imports CatiaExcelClassLibrary

Module Program

    Sub Main()


        Console.WriteLine(">>> Starting Process...")


        ' Sesion de Excel
        Dim oExcelSession As New ExcelSession
        If oExcelSession.Application Is Nothing Then
            Console.WriteLine(oExcelSession.ErrorMessage)
            Exit Sub
        End If



        ' Obtener la hoja NCU
        Const SourcePath As String = "D:\OneDrive\_CATIA\_V5R21-DLN\NCU\CATALOGO-NCU.xlsx"
        Dim oNCUSheet As Microsoft.Office.Interop.Excel.Worksheet = oExcelSession.GetNCUSheet(SourcePath)
        If oNCUSheet Is Nothing Then
            Console.WriteLine("No se pudo cargar la hoja NCU. " & oExcelSession.ErrorMessage)
            Exit Sub
        End If



        ' Extracción 
        Dim oNCUDataExtractor As New NCUDataExtractor()
        Dim oNCUDic As Dictionary(Of String, ExcelData) = oNCUDataExtractor.ExtractNCUData(oNCUSheet)
        Console.WriteLine($">>> NCU Data Extracted: {oNCUDic.Count} items.")



        ' Cierre
        oExcelSession.CloseNCU()



        '' Obtener el WorkSheet activo
        Dim oActiveSheet As Microsoft.Office.Interop.Excel.Worksheet = oExcelSession.GetActiveSheet()
        If oActiveSheet Is Nothing Then
            Console.WriteLine("ActiveWorkbook is nothing. " & oExcelSession.ErrorMessage)
            Exit Sub
        End If



        'Console.WriteLine($">>> Target Sheet: {oActiveSheet.Name}")


        ' Inyección
        Dim oNCUDataInjector As New NCUDataInjector()
        oNCUDataInjector.InjectNCUDataToExcel(oActiveSheet, oNCUDic)


        Console.WriteLine(">>> NCU Data Injection Completed.")


    End Sub

End Module
