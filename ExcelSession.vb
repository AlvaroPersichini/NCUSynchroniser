Option Explicit On
Option Strict On

Public Class ExcelSession

    <Runtime.InteropServices.DllImport("user32.dll")>
    Private Shared Function GetWindowThreadProcessId(ByVal hWnd As IntPtr, ByRef lpdwProcessId As Integer) As Integer
    End Function


    Public Property IsReady As Boolean
    Public Property ActiveSheet As Microsoft.Office.Interop.Excel.Worksheet
    Public Property App As Microsoft.Office.Interop.Excel.Application
    Public Property Workbooks As Microsoft.Office.Interop.Excel.Workbooks
    Public Property Workbook As Microsoft.Office.Interop.Excel.Workbook
    Public Property Sheets As Microsoft.Office.Interop.Excel.Sheets
    Public Property Worksheet As Microsoft.Office.Interop.Excel.Worksheet
    Public Property NCUWorkbook As Microsoft.Office.Interop.Excel.Workbook
    Public Property ErrorMessage As String = ""
    Public Property NCUSheet As Microsoft.Office.Interop.Excel.Worksheet



    Public Sub New()
        Try
            App = CType(Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application"), Microsoft.Office.Interop.Excel.Application)
        Catch
            App = Nothing
            Me.ErrorMessage = "No se pudo iniciar una nueva instancia de Excel."
        End Try

    End Sub






    Function CreateNewWorkbook() As Microsoft.Office.Interop.Excel.Workbook

        If Me.App Is Nothing Then

            Try
                Me.App = New Microsoft.Office.Interop.Excel.Application()
            Catch ex As Exception
                Me.ErrorMessage = "No se pudo iniciar una nueva instancia de Excel."
                Return Nothing
            End Try

        End If

        Try
            With Me
                .Workbook = .App.Workbooks.Add()
                .IsReady = True
                Return .Workbook
            End With

        Catch ex As Exception

            Me.ErrorMessage = "Error al intentar crear un nuevo Workbook: " & ex.Message

            Return Nothing

        End Try

    End Function



    Function GetActiveSheet() As Microsoft.Office.Interop.Excel.Worksheet

        If Me.App Is Nothing Then
            Me.ErrorMessage = "No hay session de Excel activa"
            Return Nothing
        End If


        If Me.App.ActiveWorkbook Is Nothing Then
            Me.ErrorMessage = "Excel abierto pero sin libros activos."
            Return Nothing
        End If



        Try
            Dim pid As Integer
            GetWindowThreadProcessId(New IntPtr(Me.App.Hwnd), pid)
            AppActivate(pid)
            SendKeys.SendWait("{ESC}")
        Catch
            Me.ErrorMessage = "Excel en modo edicion u ocupado"
        End Try


        Try
            With Me
                .Workbooks = .App.Workbooks
                .Workbook = .App.ActiveWorkbook
                .ActiveSheet = CType(.Workbook.ActiveSheet, Microsoft.Office.Interop.Excel.Worksheet)

            End With

            Return Me.ActiveSheet

        Catch ex As Exception

            Me.ErrorMessage = "Error al acceder a los elementos del libro: " & ex.Message

            Return Nothing

        End Try


    End Function


    Function GetNCUSheet(ByVal ncuPath As String) As Microsoft.Office.Interop.Excel.Worksheet

        If Me.App Is Nothing Then
            Try
                Me.App = New Microsoft.Office.Interop.Excel.Application()
            Catch ex As Exception
                Me.ErrorMessage = "No se pudo iniciar una nueva instancia de Excel."
                Return Nothing
            End Try
        End If


        Try
            With Me.App
                .ScreenUpdating = False
                .DisplayAlerts = False
            End With
            Me.NCUWorkbook = Me.App.Workbooks.Open(ncuPath, ReadOnly:=True)
            Me.NCUWorkbook.Windows(1).Visible = False
            NCUSheet = CType(Me.NCUWorkbook.Worksheets(1), Microsoft.Office.Interop.Excel.Worksheet)
            Return NCUSheet

        Catch ex As Exception

            Return Nothing


        Finally

            If Me.App IsNot Nothing Then
                Me.App.ScreenUpdating = True
                Me.App.DisplayAlerts = True
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