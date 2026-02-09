Option Explicit On
Option Strict On

Public Class ExcelSession
    ' Declaración de la API de Windows para obtener el PID
    <System.Runtime.InteropServices.DllImport("user32.dll")>
    Private Shared Function GetWindowThreadProcessId(ByVal hWnd As IntPtr, ByRef lpdwProcessId As Integer) As Integer
    End Function
    Public Property Application As Microsoft.Office.Interop.Excel.Application
    Public Property Workbooks As Microsoft.Office.Interop.Excel.Workbooks
    Public Property Workbook As Microsoft.Office.Interop.Excel.Workbook
    Public Property Sheets As Microsoft.Office.Interop.Excel.Sheets
    Public Property Worksheet As Microsoft.Office.Interop.Excel.Worksheet
    Public Property ActiveSheet As Microsoft.Office.Interop.Excel.Worksheet
    Public Property NCUWorkbook As Microsoft.Office.Interop.Excel.Workbook
    Public Property IsReady As Boolean = False
    Public Property ErrorMessage As String = ""


    Function CreateNewWorkbook() As Microsoft.Office.Interop.Excel.Workbook
        Try
            With Me
                .Application = New Microsoft.Office.Interop.Excel.Application With {
                .Visible = False,
                .ScreenUpdating = False,
                .DisplayAlerts = False
            }
                .Workbook = .Application.Workbooks.Add()
                .IsReady = True
                Return .Workbook
            End With
        Catch ex As Exception
            Me.ErrorMessage = "Error al iniciar Excel: " & ex.Message
            MsgBox(Me.ErrorMessage, MsgBoxStyle.Critical)
            Me.IsReady = False
            Return Nothing
        End Try
    End Function


    Function GetActiveWorkbook() As Microsoft.Office.Interop.Excel.Workbook
        Me.IsReady = False ' <--- 1. Resetear siempre al empezar
        Me.ErrorMessage = ""

        Try
            Me.Application = CType(Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application"), Microsoft.Office.Interop.Excel.Application)
            If Me.Application.ActiveWorkbook Is Nothing Then
                Me.ErrorMessage = ">>> [ERROR] Excel abierto pero sin libros activos."
                Return Nothing
            End If
        Catch ex As Exception
            Me.IsReady = False ' <--- 2. Asegurar el estado en el primer fallo
            Me.ErrorMessage = ">>> [ERROR] No se pudo conectar con Excel: " & ex.Message
            Return Nothing
        End Try

        ' Desbloqueo de celda en edición
        Try
            Dim pid As Integer
            GetWindowThreadProcessId(New IntPtr(Me.Application.Hwnd), pid)
            AppActivate(pid)
            SendKeys.SendWait("{ESC}")
        Catch
            ' No es crítico si esto falla, pero podrías loguearlo
        End Try

        ' Asignación de objetos internos
        Try
            With Me
                .Workbooks = .Application.Workbooks
                .Workbook = .Application.ActiveWorkbook
                .ActiveSheet = CType(.Workbook.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)
                .IsReady = True ' <--- 3. Éxito total
                .ErrorMessage = ""
            End With
            Return Me.Workbook
        Catch ex As Exception
            Me.IsReady = False
            Me.ErrorMessage = ">>> [ERROR] Error al acceder a los elementos del libro: " & ex.Message
            Return Nothing
        End Try
    End Function



    Function GetNCUSheet(ByVal ncuPath As String) As Microsoft.Office.Interop.Excel.Worksheet
        Dim NCUSheet As Microsoft.Office.Interop.Excel.Worksheet
        ' No ponemos IsReady = False aquí porque esta función suele ser previa a GetActiveWorkbook
        Try
            ' Intentar capturar la instancia
            Me.Application = CType(Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application"), Microsoft.Office.Interop.Excel.Application)

            With Me.Application
                .ScreenUpdating = False
                .DisplayAlerts = False
            End With

            Me.NCUWorkbook = Me.Application.Workbooks.Open(ncuPath, ReadOnly:=True)
            Me.NCUWorkbook.Windows(1).Visible = False
            NCUSheet = CType(Me.NCUWorkbook.Worksheets(1), Microsoft.Office.Interop.Excel.Worksheet)

            Return NCUSheet

        Catch ex As Exception
            ' Si falla la apertura del NCU, notificamos. 
            ' El Main decidirá si abortar usando el retorno Nothing.
            Me.ErrorMessage = "No se pudo abrir el archivo NCU: " & ex.Message
            Return Nothing
        Finally
            If Me.Application IsNot Nothing Then
                Me.Application.ScreenUpdating = True
                Me.Application.DisplayAlerts = True
            End If
        End Try
    End Function

    Public Sub CloseNCU()
        If Me.NCUWorkbook IsNot Nothing Then
            Me.NCUWorkbook.Close(SaveChanges:=False)
            Me.NCUWorkbook = Nothing
        End If
    End Sub


End Class