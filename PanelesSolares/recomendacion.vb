Imports System.Data
Imports System.Globalization

Public Class recomendacion

    ' vbnc (Mono) no implementa las instancias por defecto de formularios que
    ' ofrece Visual Basic ("otroForm.Show()" sin instanciar). Esta propiedad hace
    ' explicito lo mismo: una unica instancia viva por pantalla, recreada si fue
    ' cerrada. Compila igual en Visual Studio y deja la dependencia a la vista.
    Private Shared instancia As recomendacion

    Public Shared ReadOnly Property Actual() As recomendacion
        Get
            If instancia Is Nothing OrElse instancia.IsDisposed Then
                instancia = New recomendacion()
            End If
            Return instancia
        End Get
    End Property


    Private consumoDiarioHogar As Double
    Private horasDeSol As Double

    ''' <summary>Consumo total del hogar en kWh/dia. Lo carga calcularConsumo.</summary>
    Public Property ConsumoDiario() As Double
        Get
            Return consumoDiarioHogar
        End Get
        Set(ByVal value As Double)
            consumoDiarioHogar = value
        End Set
    End Property

    ''' <summary>Horas de sol utiles promedio elegidas por el usuario.</summary>
    Public Property HorasSol() As Double
        Get
            Return horasDeSol
        End Get
        Set(ByVal value As Double)
            horasDeSol = value
        End Set
    End Property

    Private Sub recomendacion_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        inicio.Actual.RegistrarApertura()

        ' La grilla es un selector de catalogo, no un editor: sin esto el usuario
        ' puede escribir sobre las celdas y perder lo tipeado sin ningun aviso,
        ' porque esos cambios se quedan en el DataTable y nunca llegan a la base.
        vistapanel.ReadOnly = True
        vistapanel.AllowUserToAddRows = False
        ' Antes estos valores se leian con "Dim hsSol = calcularConsumo.hsSol.SelectedItem"
        ' en el constructor, quedaban congelados entre ejecuciones y valian Nothing
        ' si el formulario se abria antes de tiempo.
        If HorasSol <= 0 Then HorasSol = 1

        consumodia.Text = ConsumoDiario.ToString("N2")
        consumoanio.Text = (ConsumoDiario * 365).ToString("N2")

        ' Potencia media que hay que sostener durante el dia (W).
        wattsDia.Text = ((ConsumoDiario * 1000) / 24).ToString("N2")

        cargarPaneles("SELECT * FROM paneles")
    End Sub

    Private Sub cargarPaneles(consulta As String, ParamArray valores() As Object)
        Try
            vistapanel.DataSource = BaseDatos.Consultar(consulta, valores)
        Catch ex As Exception
            BaseDatos.Reportar(ex, "sin conexion")
        End Try
    End Sub

    Private Sub vistapanel_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles vistapanel.CellClick
        If vistapanel.CurrentRow Is Nothing Then Exit Sub
        Dim i As Integer = vistapanel.CurrentRow.Index

        ' La consulta es "SELECT *", asi que la columna 0 es idPanel. El codigo
        ' original arrancaba en 0 y corria todos los campos un lugar: mostraba el
        ' id como tipo de panel, el tipo como watts, y asi hasta perder dimensiones.
        tipoPanel.Text = CStr(vistapanel.Item(1, i).Value)
        watts.Text = CStr(vistapanel.Item(2, i).Value)
        eficiencia.Text = CStr(vistapanel.Item(3, i).Value)
        marca.Text = CStr(vistapanel.Item(4, i).Value)
        modelo.Text = CStr(vistapanel.Item(5, i).Value)
        dimensiones.Text = CStr(vistapanel.Item(6, i).Value)
    End Sub

    Private Sub generar_Click(sender As Object, e As EventArgs) Handles generar.Click
        Dim cantidadPaneles As Integer = 1
        If cantidadPanel.SelectedIndex >= 0 Then
            Integer.TryParse(CStr(cantidadPanel.SelectedItem), cantidadPaneles)
        End If
        If cantidadPaneles <= 0 Then cantidadPaneles = 1

        ' Wp minimos por panel = (kWh/dia x 1000) / horas de sol / cantidad de paneles.
        ' HorasSol ya no puede ser 0 ni Nothing, asi que no hay division por cero.
        Dim wattsPorPanel As Double = ((ConsumoDiario * 1000) / HorasSol) / cantidadPaneles

        ' El umbral va como parametro y con formato invariante: concatenado, en un
        ' equipo con coma decimal generaba "watts >= 1234,5" y rompia la consulta.
        cargarPaneles("SELECT * FROM paneles WHERE watts >= @p0 ORDER BY watts", _
                      Math.Round(wattsPorPanel, 2))

        If vistapanel.Rows.Count = 0 Then
            MsgBox("Ningun panel cargado alcanza los " & wattsPorPanel.ToString("N0") & _
                   " W necesarios. Prueba con mas paneles.", MsgBoxStyle.Information, "sin resultados")
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Close()
    End Sub

    Private Sub AlCerrarse(ByVal sender As Object, ByVal e As FormClosedEventArgs) Handles MyBase.FormClosed
        inicio.Actual.RegistrarCierre()
    End Sub

End Class
