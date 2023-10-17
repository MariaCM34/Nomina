Imports System.Data.SqlClient

Public Class Prestamo

    Dim cnx As New Conexion()

#Region "Eventos"

    'Se ejcuta al cargar la página por primera vez
    Private Sub Prestamo_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DateTimePicker4.Value = Today
        DateTimePicker1.MinDate = Today
    End Sub

    'Calcular los días
    Private Sub DateTimePicker1_ValueChanged(sender As Object, e As EventArgs) Handles DateTimePicker1.ValueChanged
        Dim fechaIni As Date = Today
        Dim fechaFin As Date
        fechaFin = DateTimePicker1.Value
        Dim tiempo As Integer = (fechaFin - fechaIni.AddDays(0)).Days
        If Not (tiempo = 0) Then
                Label23.Text = tiempo.ToString()
            Else
                Label23.Text = "..."
            End If
    End Sub

    'Guardar
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox3.Text = "" Then
            MsgBox("Debe ingresar el número de documento")
        ElseIf TextBox1.Text = "" Then
            MsgBox("Debe ingresar el valor del préstamo")
        ElseIf ComboBox1.Text = "" Then
            MsgBox("Debe seleccionar un método de pago")
        Else
            Try
                Dim cmm As New SqlCommand("sp_mant_prestamo", cnx.conexion)
                cmm.CommandType = CommandType.StoredProcedure

                Dim fecha As String = convertFecha(DateTimePicker1.Text)

                'Todos los datos son obligatorios
                cmm.Parameters.AddWithValue("@Num_documento", TextBox3.Text)
                cmm.Parameters.AddWithValue("@Fecha_fin_prestamo", fecha)
                cmm.Parameters.AddWithValue("@Metodo_pago", ComboBox1.SelectedItem.ToString())
                cmm.Parameters.AddWithValue("@Valor_prestamo", Math.Round(Convert.ToSingle(TextBox1.Text), 2))

                If cnx.ejecutarSP(cmm) Then
                    MsgBox("Préstamo agregado correctamente")
                Else
                    MsgBox("Hubo un error agregando los datos")
                End If
            Catch ex As Exception
                'MsgBox(ex.Message)
            End Try
        End If
    End Sub

    'Cancelar
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Hide()
        menum.Show()
    End Sub
#End Region

#Region "Métodos"

    'Método para cambiar a formato yyyy-MM-dd la fecha y poder almacenarla en la BD
    Function convertFecha(fecha As String)
        Dim vec() As String = Split(fecha, "/")
        Dim fec As String = vec(2) & "-" & vec(1) & "-" & vec(0)

        Return fec
    End Function

    Private Function BuscarPrestamosActivos()
        Dim query As String = "Select hv.Numero_Documento as Documento, hv.Nombre, hv.Apellido, p.Valor_prestamo, p.Fecha_fin_prestamo AS Fecha_Fin from Hoja_de_vida hv
                                INNER JOIN Prestamos p ON hv.Numero_documento = p.Num_documento
                                WHERE Fecha_fin_prestamo > GETDATE()
                                ORDER BY p.Fecha_fin_prestamo"

        Return cnx.execSelectListPrestamoActivo(query)
    End Function

    Private Sub OrganizarListaPrestamo()
        ' Obtiene la lista de empleados activos
        Dim prestamosActivos As List(Of PrestamoViewModel) = BuscarPrestamosActivos()
        Try
            If (prestamosActivos.Count < 0) Then
                MsgBox("No se encuentran prestamos activos.")
                Return
            End If

            For Each prestamo In prestamosActivos
                PrestamosGrid.Rows.Add(
                    prestamo.Documento,
                    prestamo.Nombre,
                    prestamo.Apellido,
                    prestamo.FechaFin,
                    prestamo.Total
                    )
            Next
        Catch ex As Exception
            MsgBox("Hubo un Error, intentalo más tarde")
        End Try

    End Sub

    Private Sub ExportarListaPrestamos()
        Try
            Dim save As New SaveFileDialog
            Dim ruta As String
            Dim xlApp As Object = CreateObject("Excel.Application")
            Dim pth As String = ""
            Dim xlwb As Object = xlApp.WorkBooks.add
            Dim xlws As Object = xlwb.WorkSheets(1)
            For c As Integer = 0 To PrestamosGrid.Columns.Count - 1
                xlws.cells(1, c + 1).Value = PrestamosGrid.Columns(c).HeaderText
            Next

            For r As Integer = 0 To PrestamosGrid.RowCount - 1
                For c As Integer = 0 To PrestamosGrid.Columns.Count - 1
                    xlws.cells(r + 2, c + 1).value = Convert.ToString(PrestamosGrid.Item(c, r).Value)
                Next
            Next

            Dim SaveFileDialog1 As SaveFileDialog = New SaveFileDialog
            SaveFileDialog1.InitialDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            SaveFileDialog1.Filter = "Archivo Excel | *.xlsx"
            SaveFileDialog1.FilterIndex = 2
            If SaveFileDialog1.ShowDialog = DialogResult.OK Then
                ruta = SaveFileDialog1.FileName
                xlwb.saveas(ruta)
                xlws = Nothing
                xlwb = Nothing
                xlApp.quit()
                MsgBox("Se creo el archivo en " + ruta)
            End If

        Catch ex As Exception

        End Try
    End Sub


    Private Sub PrestamosGrid_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles PrestamosGrid.CellContentClick

    End Sub

    Private Sub ExportarPrestamos_Click(sender As Object, e As EventArgs) Handles ExportarPrestamos.Click
        ExportarListaPrestamos()
    End Sub

    Private Sub BuscarPrestamos_Click_1(sender As Object, e As EventArgs) Handles BuscarPrestamos.Click
        PrestamosGrid.Rows.Clear()
        OrganizarListaPrestamo()
    End Sub

#End Region

End Class