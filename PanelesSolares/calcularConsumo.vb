Imports System.Data

Public Class calcularConsumo

    ' vbnc (Mono) no implementa las instancias por defecto de formularios que
    ' ofrece Visual Basic ("otroForm.Show()" sin instanciar). Esta propiedad hace
    ' explicito lo mismo: una unica instancia viva por pantalla, recreada si fue
    ' cerrada. Compila igual en Visual Studio y deja la dependencia a la vista.
    Private Shared instancia As calcularConsumo

    Public Shared ReadOnly Property Actual() As calcularConsumo
        Get
            If instancia Is Nothing OrElse instancia.IsDisposed Then
                instancia = New calcularConsumo()
            End If
            Return instancia
        End Get
    End Property


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
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

    Private Sub agregar_Click(sender As Object, e As EventArgs) Handles agregar.Click
        If cantidad.Text = "" Or txtnombre.Text = "" Or txtconsumohs.Text = "" Or txtconsumok.Text = "" Then
            MsgBox("asegurese de seleccionar un electrodomestico y completar la cantidad")
            Exit Sub
        End If

        Dim unidades As Integer
        Dim horas As Double
        Dim kwh As Double
        If Not Integer.TryParse(cantidad.Text, unidades) OrElse unidades <= 0 Then
            MsgBox("La cantidad debe ser un numero entero mayor a 0")
            cantidad.Focus()
            Exit Sub
        End If
        If Not Entradas.LeerNumero(txtconsumohs.Text, horas) OrElse _
           Not Entradas.LeerNumero(txtconsumok.Text, kwh) Then
            MsgBox("El electrodomestico seleccionado tiene valores de consumo invalidos")
            Exit Sub
        End If

        ' consumo diario = unidades x horas de uso x kWh por hora.
        ' Antes "potencia" era Integer, asi que una heladera de 0,06 kWh x 24 h
        ' (1,44 kWh/dia) se redondeaba a 1 y el total del hogar salia mal.
        Dim consumoDiario As Double = unidades * horas * kwh

        listaElectro.Items.Add(txtnombre.Text)
        listaCantidad.Items.Add(unidades)
        listaPotencia.Items.Add(Math.Round(consumoDiario, 3))
    End Sub

    Private Sub eliminar_Click(sender As Object, e As EventArgs) Handles eliminar.Click
        Dim indice As Integer
        If Not Integer.TryParse(indiceNum.Text, indice) Then
            MsgBox("Seleccione un elemento de la lista")
            Exit Sub
        End If

        ' El indice quedaba obsoleto despues de borrar y RemoveAt tiraba excepcion.
        If indice < 0 Or indice >= listaElectro.Items.Count Then
            MsgBox("Seleccione un elemento de la lista")
            indiceNum.Clear()
            Exit Sub
        End If

        listaElectro.Items.RemoveAt(indice)
        listaCantidad.Items.RemoveAt(indice)
        listaPotencia.Items.RemoveAt(indice)
        indiceNum.Clear()
    End Sub

    Private Sub listaElectro_SelectedIndexChanged(sender As Object, e As EventArgs) Handles listaElectro.SelectedIndexChanged
        indiceNum.Text = CStr(listaElectro.SelectedIndex)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If listaElectro.Items.Count <= 0 Then
            MsgBox("Por favor ingrese un elemento en la lista...")
            Exit Sub
        End If

        ' SelectedItem = "" no detectaba la falta de seleccion de forma confiable.
        If hsSol.SelectedIndex < 0 Then
            MsgBox("Por favor ingrese la cantidad de horas promedio. . .")
            hsSol.Focus()
            Exit Sub
        End If

        Dim horasSol As Double
        If Not Entradas.LeerNumero(CStr(hsSol.SelectedItem), horasSol) OrElse horasSol <= 0 Then
            MsgBox("Las horas de sol seleccionadas no son validas")
            Exit Sub
        End If

        Dim total As Double = 0
        For i As Integer = 0 To listaPotencia.Items.Count - 1
            total += CDbl(listaPotencia.Items(i))
        Next

        consumodia.Text = total.ToString("N2")
        consumoanio.Text = (total * 365).ToString("N2")

        ' Los datos se pasan explicitamente en vez de que "recomendacion" lea los
        ' controles de este formulario (y despues lo cierre desde su propio Load).
        recomendacion.Actual.ConsumoDiario = total
        recomendacion.Actual.HorasSol = horasSol
        recomendacion.Actual.Show()
        Me.Close()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.Close()
    End Sub

    Private Sub cantidad_KeyPress(sender As Object, e As KeyPressEventArgs) Handles cantidad.KeyPress
        e.Handled = Not (Char.IsDigit(e.KeyChar) OrElse Char.IsControl(e.KeyChar))
    End Sub

    Private Sub AlCerrarse(ByVal sender As Object, ByVal e As FormClosedEventArgs) Handles MyBase.FormClosed
        inicio.Actual.RegistrarCierre()
    End Sub

End Class
