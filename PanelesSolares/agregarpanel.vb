Imports System.Data

Public Class agregarpanel

    ' vbnc (Mono) no implementa las instancias por defecto de formularios que
    ' ofrece Visual Basic ("otroForm.Show()" sin instanciar). Esta propiedad hace
    ' explicito lo mismo: una unica instancia viva por pantalla, recreada si fue
    ' cerrada. Compila igual en Visual Studio y deja la dependencia a la vista.
    Private Shared instancia As agregarpanel

    Public Shared ReadOnly Property Actual() As agregarpanel
        Get
            If instancia Is Nothing OrElse instancia.IsDisposed Then
                instancia = New agregarpanel()
            End If
            Return instancia
        End Get
    End Property


    Private Sub agregarpanel_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        inicio.Actual.RegistrarApertura()

        ' La grilla es un selector de catalogo, no un editor: sin esto el usuario
        ' puede escribir sobre las celdas y perder lo tipeado sin ningun aviso,
        ' porque esos cambios se quedan en el DataTable y nunca llegan a la base.
        dgvelectro.ReadOnly = True
        dgvelectro.AllowUserToAddRows = False
        actualizardatagrid()
    End Sub

    Private Sub actualizardatagrid()
        Try
            dgvelectro.DataSource = BaseDatos.Consultar("SELECT * FROM paneles")
        Catch ex As Exception
            BaseDatos.Reportar(ex, "sin conexion")
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If txtTipo.Text = "" Or txtWatts.Text = "" Or txtModelo.Text = "" Or _
           txtMarca.Text = "" Or txtEficiencia.Text = "" Or txtDimensiones.Text = "" Then
            MsgBox("No puede dejar campos vacios")
            Exit Sub
        End If

        Dim watts As Double
        Dim eficiencia As Double
        If Not Entradas.LeerNumero(txtWatts.Text, watts) Then
            MsgBox("Los watts no son un numero valido")
            txtWatts.Focus()
            Exit Sub
        End If
        If Not Entradas.LeerNumero(txtEficiencia.Text, eficiencia) Then
            MsgBox("La eficiencia no es un numero valido")
            txtEficiencia.Focus()
            Exit Sub
        End If

        If watts <= 0 Then
            MsgBox("La potencia del panel debe ser mayor a 0 W")
            txtWatts.Focus()
            Exit Sub
        End If
        If eficiencia > 100 Or eficiencia <= 0 Then
            MsgBox("no puede existir mas de 100% de eficiencia o 0%")
            txtEficiencia.Focus()
            Exit Sub
        End If

        Try
            ' Igual que en el alta de electrodomesticos, el INSERT original
            ' incrustaba los nombres de los controles en el SQL y nunca grababa.
            BaseDatos.Ejecutar(
                "INSERT INTO paneles (tipoPanel, watts, eficiencia, marca, modelo, dimensiones) " & _
                "VALUES (@p0, @p1, @p2, @p3, @p4, @p5)", _
                txtTipo.Text.Trim(), watts, eficiencia, _
                txtMarca.Text.Trim(), txtModelo.Text.Trim(), txtDimensiones.Text.Trim())
            MsgBox("guardado", MsgBoxStyle.Information, "correctamente")
            limpiar()
            actualizardatagrid()
        Catch ex As Exception
            BaseDatos.Reportar(ex, "atencion")
        End Try
    End Sub

    Private Sub limpiar()
        txtTipo.Clear()
        txtWatts.Clear()
        txtEficiencia.Clear()
        txtMarca.Clear()
        txtModelo.Clear()
        txtDimensiones.Clear()
        txtTipo.Focus()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        opcionespanel.Actual.Show()
        Me.Close()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
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
