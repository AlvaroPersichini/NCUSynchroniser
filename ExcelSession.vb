Option Explicit On


Public Class ExcelSession
    ' Declaración de la API de Windows para obtener el PID
    <Runtime.InteropServices.DllImport("user32.dll")>
    Private Shared Function GetWindowThreadProcessId(ByVal hWnd As IntPtr, ByRef lpdwProcessId As Integer) As Integer
    End Function
    Public Property Application As Microsoft.Office.Interop.Excel.Application
    Public Property Workbook As Microsoft.Office.Interop.Excel.Workbook
    Public Property ActiveSheet As Microsoft.Office.Interop.Excel.Worksheet
    Public Property IsReady As Boolean = False
    Public Property ErrorMessage As String = ""
    Public Property SheetNCU As Microsoft.Office.Interop.Excel.Worksheet

    Private _ncuWorkbook As Microsoft.Office.Interop.Excel.Workbook

    Public Sub New()
        Try
            ' 1. Conexión con Excel
            Me.Application = CType(Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application"), Microsoft.Office.Interop.Excel.Application)
            Dim excelHwnd As New IntPtr(Me.Application.Hwnd)
            Dim excelPid As Integer
            GetWindowThreadProcessId(excelHwnd, excelPid)
            AppActivate(excelPid)
            SendKeys.SendWait("{ESC}")



            Me.IsReady = True

        Catch ex As Runtime.InteropServices.COMException
            Me.ErrorMessage = ">>> [ERROR] No se detectó ninguna instancia de Excel abierta."
        Catch ex As Exception
            Me.ErrorMessage = ">>> [ERROR] " & ex.Message
        End Try
    End Sub




    Public Function GetActiveSheet() As Microsoft.Office.Interop.Excel.Worksheet

        If Me.Application Is Nothing Then
            Me.ErrorMessage = ">>> [ERROR] Excel no está disponible."
            Return Nothing
        End If

        If Me.Application.ActiveWorkbook Is Nothing Then
            Me.ErrorMessage = ">>> [ERROR] Excel abierto pero sin libros activos."
            Return Nothing
        End If

        Return CType(Me.Application.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)

    End Function


    Public Function GetNCUSheet(ByVal ncuPath As String) As Microsoft.Office.Interop.Excel.Worksheet

        If Me.Application Is Nothing OrElse Not IO.File.Exists(ncuPath) Then
            Return Nothing
        End If

        Try
            ' --- EFECTO FANTASMA ---
            ' Congelamos la pantalla para que el usuario no vea el parpadeo del nuevo libro
            Me.Application.ScreenUpdating = False


            ' Abrimos como ReadOnly para evitar carteles de "Archivo en uso"
            _ncuWorkbook = Me.Application.Workbooks.Open(ncuPath, ReadOnly:=True)
            _ncuWorkbook.Windows(1).Visible = False


            ' Devolvemos la hoja
            Dim oSheet As Microsoft.Office.Interop.Excel.Worksheet = CType(_ncuWorkbook.Worksheets(1), Microsoft.Office.Interop.Excel.Worksheet)

            ' Reactivamos la actualización de pantalla
            Me.Application.ScreenUpdating = True

            Return oSheet
        Catch ex As Exception
            Me.Application.ScreenUpdating = True
            Me.ErrorMessage = ex.Message
            Return Nothing
        End Try
    End Function


    Public Sub CloseNCU()
        If _ncuWorkbook IsNot Nothing Then
            _ncuWorkbook.Close(SaveChanges:=False)
            _ncuWorkbook = Nothing
        End If
    End Sub


End Class