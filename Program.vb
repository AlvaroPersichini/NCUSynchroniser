Module Program

    Sub Main()




        ' Excel
        Dim xlSession As New ExcelSession()
        If Not xlSession.IsReady Then
            Console.WriteLine(xlSession.ErrorMessage)
            Return
        End If



        ' Extraccion

        Dim oNCUDataExtractor As New NCUDataExtractor()
        Dim oDic As Dictionary(Of String, ExcelData) = oNCUDataExtractor.ExtractNCUData(xlSession.ActiveSheet)










    End Sub

End Module
