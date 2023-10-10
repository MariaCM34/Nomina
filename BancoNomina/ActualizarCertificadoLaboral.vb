﻿Imports System.Data.SqlClient
Imports System.IO
Imports AxAcroPDFLib

Public Class ActualizarCertificadoLaboral

    Dim cnx As New Conexion()

    Public documento As String
    Dim archivo As Byte()

#Region "Eventos"

    'Cambio en el combo de id para cargar datos
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        cargarDatos(ComboBox1.SelectedItem)
    End Sub

    'Actualizar datos
    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        Try
            Dim cmm As New SqlCommand("sp_mant_certificado_laboral", cnx.conexion)
            cmm.CommandType = CommandType.StoredProcedure

            'Todos los datos son obligatorios en la BD
            cmm.Parameters.AddWithValue("@vlr", 1) '1 para modificar
            cmm.Parameters.AddWithValue("@id", Convert.ToInt16(ComboBox1.SelectedItem))
            cmm.Parameters.AddWithValue("@Nume_documento", documento)

            Dim valor As Integer = validarCampos(TextBox29, CheckBox15)
            If valor = 1 Or valor = 2 Then
                cmm.Parameters.AddWithValue("@Certificado_laboral", archivo)
            ElseIf valor = 3 Then
                If Not TextBox29.Text = "" Then
                    cmm.Parameters.AddWithValue("@Certificado_laboral", pasarABinario(TextBox29.Text))
                Else
                    cmm.Parameters.AddWithValue("@Certificado_laboral", archivo)
                End If
            End If

            If (cnx.ejecutarSP(cmm)) Then
                MsgBox("Certificado laboral actualizado correctamente")
            Else
                MsgBox("Hubo un error actualizando los datos")
                AxAcroPDF4.Dispose()
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    'Eliminar datos
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Dim cmm As New SqlCommand("sp_eliminar_certificado_laboral", cnx.conexion)
            cmm.CommandType = CommandType.StoredProcedure

            cmm.Parameters.AddWithValue("@id", Convert.ToInt16(ComboBox1.SelectedItem))
            cmm.Parameters.AddWithValue("@Nume_documento", documento)

            If cnx.ejecutarSP(cmm) Then
                MsgBox("El certificado laboral fue eliminado correctamente")
            Else
                MsgBox("Hubo un error eliminando los datos")
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    'Devolverse a gestión
    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Me.Hide()
        HojaVida.Show()
        HojaVida.TabPage6.Show()
    End Sub
#End Region

#Region "Métodos"

    'Método para cargar ids de comunicados
    Function llenarComboId()
        ComboBox1.Items.Clear()
        Button10.Enabled = False
        Button1.Enabled = False
        LinkLabel15.Enabled = False
        AxAcroPDF4.LoadFile("none")

        Dim consulta As String = ("select id from Certificado_laboral where Nume_documento = " & documento)
        Dim valores As List(Of String) = cnx.execSelectVarios(consulta, False)

        If Not valores.Count() = 0 Then
            For Each item As String In valores
                ComboBox1.Items.Add(item)
            Next
        Else
            MsgBox("No hay certificados relacionados al documento " & documento)
        End If
        TextBox29.Text = "_"
    End Function

    'Método para consultar datos y cargarlos
    Private Sub cargarDatos(id As String)
        Dim consulta As String = ("select Certificado_laboral from Certificado_laboral where id = " & id)
        Dim valores As List(Of Byte()) = cnx.execSelectVarios(consulta, True)

        If Not valores.Count() = 0 Then
            recuperarPDF(valores.ElementAt(0), CheckBox15)
            Button10.Enabled = True
            Button1.Enabled = True
            LinkLabel15.Enabled = True
        End If
    End Sub

    'Método para leer archivos pdf, escribir ruta y seleccionar check
    Private Sub leerPDF(AxAcroPDF1 As AxAcroPDF, campo As TextBox, chk As CheckBox)
        'abrir el explorador de archivos desde un link 
        Try
            Dim archivo As New OpenFileDialog
            archivo.Filter = "Archivo PDF|*.pdf"
            If archivo.ShowDialog = DialogResult.OK Then
                campo.Text = archivo.FileName
                AxAcroPDF1.src = archivo.FileName
                chk.Checked = True
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    'Método que retorna un archivo en binario para poder almacenarlo
    Function pasarABinario(texto As String)
        If Not (texto = "") Then
            Dim path As String = texto
            Dim ruta As New FileStream(path, FileMode.Open, FileAccess.Read)
            Dim binario(ruta.Length) As Byte
            ruta.Read(binario, 0, ruta.Length) 'Leer archivo y pasarlo a binario
            ruta.Close()

            Return binario 'Retorna el archivo en binario
        Else
            Dim binario(0) As Byte
            Return binario 'Retorna un byte de tamaño 0 vacío, si no se cargó ningún documento
        End If
    End Function

    'Método para quitar archivos que ya estaban cargados
    Private Sub quitarArchivos(chk As CheckBox, txtbox As TextBox, Acro As AxAcroPDF)
        If Not (chk.Checked) Then
            txtbox.Text = "."
            Acro.LoadFile("none")
        End If
    End Sub

    'Método para recuperar pdf
    Private Sub recuperarPDF(byt As Byte(), chk As CheckBox)
        Try
            Dim directorioArchivo As String
            directorioArchivo = System.AppDomain.CurrentDomain.BaseDirectory() & "temp.pdf"

            Dim bytes() As Byte
            bytes = byt

            Dim res As Boolean = BytesAArchivo(bytes, directorioArchivo, chk)
            If (res) Then
                AxAcroPDF4.src = directorioArchivo
                My.Computer.FileSystem.DeleteFile(directorioArchivo)
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    'Método para convertir bytes a pdf
    Function BytesAArchivo(ByVal bytes() As Byte, ByVal Path As String, chk As CheckBox)
        Try
            Dim K As Long = UBound(bytes)
            If Not K = 0 Then
                archivo = bytes
                Dim fs As New FileStream(Path, FileMode.OpenOrCreate, FileAccess.Write)
                fs.Write(bytes, 0, K)
                fs.Close()
                chk.Checked = True
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message, ex)
        End Try
    End Function

    'Método para ejcutar SP archivos
    Function validarCampos(text As TextBox, chk As CheckBox)
        Try
            If (text.Text = ".") Then
                If Not (chk.Checked) Then
                    'Se modifica el archivo a 0
                    Return 1
                End If
            ElseIf text.Text = "_" Then
                If chk.Checked Then
                    'No se realiza ninguna acción
                    Return 2
                End If
            Else
                'Se modifica archivo por uno nuevo
                Return 3
            End If

            Return 1
        Catch ex As Exception
            Return 1
        End Try
    End Function

#End Region

#Region "Link & check"

    'Link para subir archivo
    Private Sub LinkLabel15_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel15.LinkClicked
        leerPDF(AxAcroPDF4, TextBox29, CheckBox15)
    End Sub

    'Check unchecked
    Private Sub CheckBox15_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox15.CheckedChanged
        quitarArchivos(CheckBox15, TextBox29, AxAcroPDF4)
    End Sub
#End Region

End Class