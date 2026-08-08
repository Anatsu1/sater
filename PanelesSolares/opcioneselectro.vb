Imports System.Data

Public Class opcioneselectro

    ' vbnc (Mono) no implementa las instancias por defecto de formularios que
    ' ofrece Visual Basic ("otroForm.Show()" sin instanciar). Esta propiedad hace
    ' explicito lo mismo: una unica instancia viva por pantalla, recreada si fue
    ' cerrada. Compila igual en Visual Studio y deja la dependencia a la vista.
    Private Shared instancia As opcioneselectro

    Public Shared ReadOnly Property Actual() As opcioneselectro
        Get
            If instancia Is Nothing OrElse instancia.IsDisposed Then
                instancia = New opcioneselectro()
            End If
            Return instancia
        End Get
    End Property


    Private Sub editarelectro_Load_1(sender As Object, e As EventArgs) Handles MyBase.Load
        inicio.Actual.RegistrarApertura()

        ' La grilla es un selector de catalogo, no un editor: sin esto el usuario
        ' puede escribir sobre las celdas y perder lo tipeado sin ningun aviso,
        ' porque esos cambios se quedan en el DataTable y nunca llegan a la base.
        vistaelectro.ReadOnly = True
        vistaelectro.AllowUserToAddRows = False
        actualizardatagrid()
    End Sub

    Private Sub actualizardatagrid()
        Try
            vistaelectro.DataSource = BaseDatos.Consultar("SELECT * FROM electro")
        Catch ex As Exception
            BaseDatos.Reportar(ex, "sin conexion")
        End Try
    End Sub

    Private Sub vistaelectro_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles vistaelectro.CellClick
        If vistaelectro.CurrentRow Is Nothing Then Exit Sub
        Dim i As Integer = vistaelectro.CurrentRow.Index
        txtidElectro.Text = CStr(vistaelectro.Item(0, i).Value)
        txtnombre.Text = CStr(vistaelectro.Item(1, i).Value)
        txtconsumohs.Text = CStr(vistaelectro.Item(2, i).Value)
        txtconsumok.Text = CStr(vistaelectro.Item(3, i).Value)
    End Sub

    Private Sub NombreToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NombreToolStripMenuItem.Click
        If buscarNombre.Text = "" Then Exit Sub
        ' El LIKE concatenaba el texto del usuario: un apostrofe rompia la consulta
        ' y "%' OR '1'='1" devolvia la tabla entera.
        buscar("SELECT * FROM electro WHERE nombre LIKE @p0", "%" & buscarNombre.Text.Trim() & "%", _
               "NO HAY REGISTROS CON DICHO NOMBRE")
    End Sub

    Private Sub IDToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles IDToolStripMenuItem.Click
        If buscarId.Text = "" Then Exit Sub
        Dim id As Integer
        If Not Integer.TryParse(buscarId.Text, id) Then
            MsgBox("El ID debe ser un numero", MsgBoxStyle.Critical, "A T E N C I O N")
            Exit Sub
        End If
        buscar("SELECT * FROM electro WHERE idElectro = @p0", id, "NO HAY REGISTROS CON DICHO ID")
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

            vistaelectro.DataSource = tabla
            Dim fila As DataRow = tabla.Rows(0)
            txtidElectro.Text = CStr(fila("idElectro"))
            txtnombre.Text = CStr(fila("nombre"))
            txtconsumohs.Text = CStr(fila("hsConsumo"))
            txtconsumok.Text = CStr(fila("conKwh"))
        Catch ex As Exception
            BaseDatos.Reportar(ex, "A T E N C I O N")
        End Try
    End Sub

    Private Sub limpiar()
        txtidElectro.Clear()
        buscarId.Clear()
        txtnombre.Clear()
        txtconsumohs.Clear()
        txtconsumok.Clear()
    End Sub

    Private Sub modificar_Click(sender As Object, e As EventArgs) Handles modificar.Click
        If txtidElectro.Text = "" Then
            MsgBox("Seleccione primero un electrodomestico de la lista")
            Exit Sub
        End If
        habilitarEdicion(True)
    End Sub

    Private Sub habilitarEdicion(activo As Boolean)
        txtnombre.Enabled = activo
        txtconsumohs.Enabled = activo
        txtconsumok.Enabled = activo
        enviar.Enabled = activo
        buscarId.Enabled = Not activo
        buscarNombre.Enabled = Not activo
    End Sub

    Private Sub enviar_Click(sender As Object, e As EventArgs) Handles enviar.Click
        If txtnombre.Text = "" Or txtconsumohs.Text = "" Or txtconsumok.Text = "" Then
            MsgBox("No puede dejar campos vacios")
            Exit Sub
        End If

        Dim id As Integer
        Dim horas As Double
        Dim kwh As Double
        If Not Integer.TryParse(txtidElectro.Text, id) Then
            MsgBox("Seleccione primero un electrodomestico de la lista")
            Exit Sub
        End If
        If Not Entradas.LeerNumero(txtconsumohs.Text, horas) OrElse _
           Not Entradas.LeerNumero(txtconsumok.Text, kwh) Then
            MsgBox("Los valores de consumo no son numeros validos")
            Exit Sub
        End If
        If horas > 24 Or horas <= 0 Then
            MsgBox("No puede ingresar mas de 24 horas ni ingresar 0 horas")
            Exit Sub
        End If
        If kwh <= 0 Then
            MsgBox("El consumo en kWh debe ser mayor a 0")
            Exit Sub
        End If

        Try
            ' El UPDATE original mandaba los numeros entre comillas simples y
            ' convertidos con CDbl: en un equipo con coma decimal grababa "0,06"
            ' como texto y el consumo se leia mal (o directamente fallaba).
            BaseDatos.Ejecutar(
                "UPDATE electro SET nombre = @p0, hsConsumo = @p1, conKwh = @p2 WHERE idElectro = @p3", _
                txtnombre.Text.Trim(), horas, kwh, id)
            MsgBox("Actualizado correctamente", MsgBoxStyle.Information, " correcto")
            limpiar()
            actualizardatagrid()
            habilitarEdicion(False)
        Catch ex As Exception
            BaseDatos.Reportar(ex, "atencion")
        End Try
    End Sub

    Private Sub borrar_Click(sender As Object, e As EventArgs) Handles borrar.Click
        Dim id As Integer
        If Not Integer.TryParse(txtidElectro.Text, id) Then
            MsgBox("Tiene que seleccionar un id para poder eliminar el respectivo electrodomestico!")
            actualizardatagrid()
            Exit Sub
        End If

        If MsgBox("Esta seguro que desea eliminar...", MsgBoxStyle.YesNo, "¿eliminar?") <> MsgBoxResult.Yes Then
            MsgBox("Cancelo la eliminacion", MsgBoxStyle.Critical, "Cancelado")
            limpiar()
            actualizardatagrid()
            Exit Sub
        End If

        Try
            BaseDatos.Ejecutar("DELETE FROM electro WHERE idElectro = @p0", id)
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
        buscarNombre.Enabled = (buscarId.Text = "")
    End Sub

    Private Sub buscarNombre_TextChanged(sender As Object, e As EventArgs) Handles buscarNombre.TextChanged
        buscarId.Enabled = (buscarNombre.Text = "")
    End Sub

    Private Sub buscarId_KeyPress(sender As Object, e As KeyPressEventArgs) Handles buscarId.KeyPress
        e.Handled = Not (Char.IsDigit(e.KeyChar) OrElse Char.IsControl(e.KeyChar))
    End Sub

    Private Sub buscarNombre_KeyPress(sender As Object, e As KeyPressEventArgs) Handles buscarNombre.KeyPress
        Entradas.SoloLetras(e)
    End Sub

    Private Sub txtnombre_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtnombre.KeyPress
        Entradas.SoloLetras(e)
    End Sub

    Private Sub txtconsumohs_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtconsumohs.KeyPress
        Entradas.SoloNumeros(e, sender)
    End Sub

    Private Sub txtconsumok_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtconsumok.KeyPress
        Entradas.SoloNumeros(e, sender)
    End Sub

    Private Sub AlCerrarse(ByVal sender As Object, ByVal e As FormClosedEventArgs) Handles MyBase.FormClosed
        inicio.Actual.RegistrarCierre()
    End Sub

End Class
