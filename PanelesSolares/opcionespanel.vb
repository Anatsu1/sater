Imports System.Data

Public Class opcionespanel

    ' vbnc (Mono) no implementa las instancias por defecto de formularios que
    ' ofrece Visual Basic ("otroForm.Show()" sin instanciar). Esta propiedad hace
    ' explicito lo mismo: una unica instancia viva por pantalla, recreada si fue
    ' cerrada. Compila igual en Visual Studio y deja la dependencia a la vista.
    Private Shared instancia As opcionespanel

    Public Shared ReadOnly Property Actual() As opcionespanel
        Get
            If instancia Is Nothing OrElse instancia.IsDisposed Then
                instancia = New opcionespanel()
            End If
            Return instancia
        End Get
    End Property


    Private Sub opcionespanel_Load_1(sender As Object, e As EventArgs) Handles MyBase.Load
        inicio.Actual.RegistrarApertura()

        ' La grilla es un selector de catalogo, no un editor: sin esto el usuario
        ' puede escribir sobre las celdas y perder lo tipeado sin ningun aviso,
        ' porque esos cambios se quedan en el DataTable y nunca llegan a la base.
        vistapanel.ReadOnly = True
        vistapanel.AllowUserToAddRows = False
        actualizardatagrid()
    End Sub

    Private Sub actualizardatagrid()
        Try
            vistapanel.DataSource = BaseDatos.Consultar("SELECT * FROM paneles")
        Catch ex As Exception
            BaseDatos.Reportar(ex, "sin conexion")
        End Try
    End Sub

    Private Sub vistapanel_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles vistapanel.CellClick
        If vistapanel.CurrentRow Is Nothing Then Exit Sub
        Dim i As Integer = vistapanel.CurrentRow.Index
        idPanel.Text = CStr(vistapanel.Item(0, i).Value)
        txtTipo.Text = CStr(vistapanel.Item(1, i).Value)
        txtWatts.Text = CStr(vistapanel.Item(2, i).Value)
        txtEficiencia.Text = CStr(vistapanel.Item(3, i).Value)
        txtMarca.Text = CStr(vistapanel.Item(4, i).Value)
        txtModelo.Text = CStr(vistapanel.Item(5, i).Value)
        txtDimensiones.Text = CStr(vistapanel.Item(6, i).Value)
    End Sub

    Private Sub IDToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles IDToolStripMenuItem.Click
        If buscarId.Text = "" Then Exit Sub
        Dim id As Integer
        If Not Integer.TryParse(buscarId.Text, id) Then
            MsgBox("El ID debe ser un numero", MsgBoxStyle.Critical, "A T E N C I O N")
            Exit Sub
        End If
        buscar("SELECT * FROM paneles WHERE idPanel = @p0", id, "NO HAY REGISTROS CON DICHO ID")
    End Sub

    Private Sub TipoPanelToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TipoPanelToolStripMenuItem.Click
        If buscarTipo.Text = "" Then Exit Sub
        ' Mismo problema de inyeccion que en el filtro por nombre de electrodomestico.
        buscar("SELECT * FROM paneles WHERE tipoPanel LIKE @p0", "%" & buscarTipo.Text.Trim() & "%", _
               "NO HAY REGISTROS CON DICHO TIPO DE PANEL")
    End Sub

    Private Sub buscar(consulta As String, valor As Object, sinResultados As String)
        Try
            Dim tabla As DataTable = BaseDatos.Consultar(consulta, valor)
            If tabla.Rows.Count = 0 Then
                MsgBox(sinResultados, MsgBoxStyle.Critical, "A T E N C I O N")
                limpiar()
                actualizardatagrid()
                Exit Sub
            End If

            vistapanel.DataSource = tabla
            Dim fila As DataRow = tabla.Rows(0)
            idPanel.Text = CStr(fila("idPanel"))
            txtTipo.Text = CStr(fila("tipoPanel"))
            txtWatts.Text = CStr(fila("watts"))
            txtEficiencia.Text = CStr(fila("eficiencia"))
            txtMarca.Text = CStr(fila("marca"))
            txtModelo.Text = CStr(fila("modelo"))
            txtDimensiones.Text = CStr(fila("dimensiones"))
        Catch ex As Exception
            BaseDatos.Reportar(ex, "A T E N C I O N")
        End Try
    End Sub

    Private Sub limpiar()
        idPanel.Clear()
        txtTipo.Clear()
        txtWatts.Clear()
        txtEficiencia.Clear()
        txtMarca.Clear()
        txtModelo.Clear()
        txtDimensiones.Clear()
    End Sub

    Private Sub modificar_Click(sender As Object, e As EventArgs) Handles modificar.Click
        If idPanel.Text = "" Then
            MsgBox("Seleccione primero un panel de la lista")
            Exit Sub
        End If
        habilitarEdicion(True)
    End Sub

    Private Sub habilitarEdicion(activo As Boolean)
        txtTipo.Enabled = activo
        txtWatts.Enabled = activo
        txtEficiencia.Enabled = activo
        txtMarca.Enabled = activo
        txtModelo.Enabled = activo
        txtDimensiones.Enabled = activo
        enviar.Enabled = activo
        buscarId.Enabled = Not activo
        buscarTipo.Enabled = Not activo
    End Sub

    Private Sub enviar_Click(sender As Object, e As EventArgs) Handles enviar.Click
        If txtTipo.Text = "" Or txtWatts.Text = "" Or txtModelo.Text = "" Or _
           txtMarca.Text = "" Or txtEficiencia.Text = "" Or txtDimensiones.Text = "" Then
            MsgBox("No puede dejar campos vacios")
            Exit Sub
        End If

        Dim id As Integer
        Dim watts As Double
        Dim eficiencia As Double
        If Not Integer.TryParse(idPanel.Text, id) Then
            MsgBox("Seleccione primero un panel de la lista")
            Exit Sub
        End If
        If Not Entradas.LeerNumero(txtWatts.Text, watts) OrElse _
           Not Entradas.LeerNumero(txtEficiencia.Text, eficiencia) Then
            MsgBox("Los watts y la eficiencia deben ser numeros validos")
            Exit Sub
        End If
        If watts <= 0 Then
            MsgBox("La potencia del panel debe ser mayor a 0 W")
            Exit Sub
        End If
        If eficiencia > 100 Or eficiencia <= 0 Then
            MsgBox("no puede existir mas de 100% de eficiencia o 0%")
            Exit Sub
        End If

        Try
            BaseDatos.Ejecutar(
                "UPDATE paneles SET tipoPanel = @p0, watts = @p1, eficiencia = @p2, " & _
                "marca = @p3, modelo = @p4, dimensiones = @p5 WHERE idPanel = @p6", _
                txtTipo.Text.Trim(), watts, eficiencia, _
                txtMarca.Text.Trim(), txtModelo.Text.Trim(), txtDimensiones.Text.Trim(), id)
            MsgBox("Actualizado correctamente", MsgBoxStyle.Information, " correcto")
            limpiar()
            actualizardatagrid()
            habilitarEdicion(False)
        Catch ex As Exception
            BaseDatos.Reportar(ex, "atencion")
        End Try
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim id As Integer
        If Not Integer.TryParse(idPanel.Text, id) Then
            MsgBox("Tiene que seleccionar un id para poder eliminar el respectivo panel!")
            Exit Sub
        End If

        If MsgBox("Esta seguro que desea eliminar...", MsgBoxStyle.YesNo, "¿eliminar?") <> MsgBoxResult.Yes Then
            MsgBox("Cancelo la eliminacion", MsgBoxStyle.Critical, "Cancelado")
            limpiar()
            Exit Sub
        End If

        Try
            BaseDatos.Ejecutar("DELETE FROM paneles WHERE idPanel = @p0", id)
            MsgBox("Eliminado correctamente", MsgBoxStyle.Information, "Correcto")
            limpiar()
            actualizardatagrid()
        Catch ex As Exception
            BaseDatos.Reportar(ex, "atencion")
        End Try
    End Sub

    Private Sub VOLVERToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles VOLVERToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        limpiar()
        actualizardatagrid()
    End Sub

    Private Sub buscarId_TextChanged(sender As Object, e As EventArgs) Handles buscarId.TextChanged
        buscarTipo.Enabled = (buscarId.Text = "")
    End Sub

    Private Sub buscarTipo_TextChanged(sender As Object, e As EventArgs) Handles buscarTipo.TextChanged
        buscarId.Enabled = (buscarTipo.Text = "")
    End Sub

    Private Sub txtTipo_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtTipo.KeyPress
        Entradas.SoloLetras(e)
    End Sub

    Private Sub txtWatts_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtWatts.KeyPress
        Entradas.SoloNumeros(e, sender)
    End Sub

    Private Sub txtEficiencia_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtEficiencia.KeyPress
        Entradas.SoloNumeros(e, sender)
    End Sub

    Private Sub txtMarca_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtMarca.KeyPress
        Entradas.Alfanumerico(e)
    End Sub

    Private Sub txtModelo_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtModelo.KeyPress
        Entradas.Alfanumerico(e)
    End Sub

    Private Sub txtDimensiones_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtDimensiones.KeyPress
        Entradas.SoloDimensiones(e)
    End Sub

    Private Sub AlCerrarse(ByVal sender As Object, ByVal e As FormClosedEventArgs) Handles MyBase.FormClosed
        inicio.Actual.RegistrarCierre()
    End Sub

End Class
