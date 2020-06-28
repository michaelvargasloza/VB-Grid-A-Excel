Imports System.Data.SqlClient

Public Class frmInicio
    Dim currentTime As Date

    Private Sub btnGenerar_Click(sender As Object, e As EventArgs) Handles btnGenerar.Click
        Try
            Dim dt As New DataTable
            Dim cn As String = "Server=;Database=;User Id=;Password=;Connection Timeout=8000"
            Using adaptador As New SqlDataAdapter("query", cn)
                adaptador.Fill(dt)
            End Using
            dgvTabla.DataSource = dt

            MsgBox("Consulta generada correctamente.")
        Catch ex As Exception
            MsgBox("Error al consultar a la base de datos: " + ex.ToString)
        End Try
    End Sub

    Private Sub btnGuardar_Click(sender As Object, e As EventArgs) Handles btnGuardar.Click
        Try
            GridAExcel(dgvTabla)
        Catch ex As Exception
            MsgBox("Error al generar el archivo: " + ex.ToString)
        End Try
    End Sub

    Function GridAExcel(ByVal ElGrid As DataGridView) As Boolean
        currentTime = TimeOfDay

        Dim correcto As Boolean = True
        Dim d As Date
        d = Date.Now
        Dim completDate As String = d.ToString()
        Dim aux() As String = completDate.Split(" ")
        Dim fechaCompleta() As String = aux(0).Split("/")

        Dim guardarTabla As New SaveFileDialog
        'Destino donde se guarda el reporte
        guardarTabla.FileName = "C:\DIRECCION\Reporte " & Now.Day & "-" & Now.Month & "-" & Now.Year & ".xlsx"

        Try
            ExportarExcel(dgvTabla, guardarTabla.FileName)
        Catch
            correcto = False
        End Try
        If correcto Then
            'lbxEstado.Items.Add(currentTime + " Reporte [" + reporte + "] generado correctamente.")
        End If
        Return True
    End Function

    Private Sub SalirToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SalirToolStripMenuItem.Click
        Me.Close()
    End Sub

    Public Sub ExportarExcel(ByVal dgv As DataGridView, ByVal pth As String)
        Dim xlApp As Object = CreateObject("Excel.Application")

        'creamos una nueva hoja de calculo 
        Dim xlWB As Object = xlApp.WorkBooks.add
        Dim xlWS As Object = xlWB.WorkSheets(1)

        'exportamos los caracteres de las columnas
        For c As Integer = 0 To dgv.Columns.Count - 1
            xlWS.cells(1, c + 1).value = dgv.Columns(c).HeaderText
        Next

        'exportamos contenido de las tablas
        For r As Integer = 0 To dgv.RowCount - 1
            For c As Integer = 0 To dgv.Columns.Count - 1
                xlWS.cells(r + 2, c + 1).value = dgv.Item(c, r).Value
            Next
        Next

        'guardamos la hoja de calculo en la ruta especificada 
        xlWB.saveas(pth)
        xlWS = Nothing
        xlWB = Nothing
        xlApp.quit()
        xlApp.kill()
        xlApp = Nothing
    End Sub
End Class
