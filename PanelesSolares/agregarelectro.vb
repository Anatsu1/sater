Imports System.Data

Public Class agregarelectro

    ' vbnc (Mono) no implementa las instancias por defecto de formularios que
    ' ofrece Visual Basic ("otroForm.Show()" sin instanciar). Esta propiedad hace
    ' explicito lo mismo: una unica instancia viva por pantalla, recreada si fue
    ' cerrada. Compila igual en Visual Studio y deja la dependencia a la vista.
    Private Shared instancia As agregarelectro

    Public Shared ReadOnly Property Actual() As agregarelectro
        Get
            If instancia Is Nothing OrElse instancia.IsDisposed Then
                instancia = New agregarelectro()
            End If
            Return instancia
        End Get
    End Property


    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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
            dgvelectro.DataSource = BaseDatos.Consultar("SELECT * FROM electro")
        Catch ex As Exception
            BaseDatos.Reportar(ex, "sin conexion")
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If txtnombre.Text = "" Or txtconsumohs.Text = "" Or txtconsumok.Text = "" Then
            MsgBox("No puede dejar campos vacios")
            Exit Sub
        End If

        Dim horas As Double
        Dim kwh As Double
        If Not Entradas.LeerNumero(txtconsumohs.Text, horas) Then
            MsgBox("Las horas de consumo no son un numero valido")
            txtconsumohs.Focus()
            Exit Sub
        End If
        If Not Entradas.LeerNumero(txtconsumok.Text, kwh) Then
            MsgBox("El consumo en kWh no es un numero valido")
            txtconsumok.Focus()
            Exit Sub
        End If

        If horas > 24 Or horas <= 0 Then
            MsgBox("No puede ingresar mas de 24 horas ni ingresar 0 horas")
            txtconsumohs.Focus()
            Exit Sub
        End If
        If kwh <= 0 Then
            MsgBox("El consumo en kWh debe ser mayor a 0")
            txtconsumok.Focus()
            Exit Sub
        End If

        Try
            ' El INSERT original incrustaba los nombres de los TextBox dentro del
            ' SQL ("VALUES (txtnombre, Cdbl(txtconsumohs), ...)"), por lo que el
            ' alta fallaba siempre. Ahora van como parametros reales.
            BaseDatos.Ejecutar(
                "INSERT INTO electro (nombre, hsConsumo, conKwh) VALUES (@p0, @p1, @p2)", _
                txtnombre.Text.Trim(), horas, kwh)
            MsgBox("guardado", MsgBoxStyle.Information, "correctamente")
            limpiar()
            actualizardatagrid()
        Catch ex As Exception
            BaseDatos.Reportar(ex, "atencion")
        End Try
    End Sub

    Private Sub limpiar()
        txtnombre.Clear()
        txtconsumohs.Clear()
        txtconsumok.Clear()
        txtnombre.Focus()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        opcioneselectro.Actual.Show()
        Me.Close()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Close()
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

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click

    End Sub

    Private Sub AlCerrarse(ByVal sender As Object, ByVal e As FormClosedEventArgs) Handles MyBase.FormClosed
        inicio.Actual.RegistrarCierre()
    End Sub

End Class
