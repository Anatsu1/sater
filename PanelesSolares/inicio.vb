Public Class inicio

    ' vbnc (Mono) no implementa las instancias por defecto de formularios que
    ' ofrece Visual Basic ("otroForm.Show()" sin instanciar). Esta propiedad hace
    ' explicito lo mismo: una unica instancia viva por pantalla, recreada si fue
    ' cerrada. Compila igual en Visual Studio y deja la dependencia a la vista.
    Private Shared instancia As inicio

    Public Shared ReadOnly Property Actual() As inicio
        Get
            If instancia Is Nothing OrElse instancia.IsDisposed Then
                instancia = New inicio()
            End If
            Return instancia
        End Get
    End Property


    ' Cantidad de formularios hijos abiertos. El menu principal se oculta mientras
    ' haya alguno y vuelve a mostrarse cuando se cierra el ultimo.
    '
    ' Antes cada opcion hacia "otroForm.Show() : Me.Hide()" y solo los botones
    ' VOLVER devolvian el foco al menu. Si el usuario cerraba la ventana hija con
    ' la X, "inicio" quedaba oculto pero abierto: como el ShutdownMode espera a
    ' que se cierre el formulario principal, la aplicacion seguia viva en memoria
    ' sin ninguna ventana visible.
    '
    ' Cada formulario hijo avisa al abrirse y al cerrarse; el menu solo lleva la
    ' cuenta, asi que agregar pantallas nuevas no obliga a tocar esta clase.
    Private abiertos As Integer

    ''' <summary>Lo llama cada formulario hijo desde su evento Load.</summary>
    Public Sub RegistrarApertura()
        abiertos += 1
        Me.Hide()
    End Sub

    ''' <summary>Lo llama cada formulario hijo desde su evento FormClosed.</summary>
    Public Sub RegistrarCierre()
        abiertos -= 1
        If abiertos <= 0 Then
            abiertos = 0
            Me.Show()
            Me.BringToFront()
        End If
    End Sub

    Private Sub inicio_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
        ' El formulario principal lo instancia MyApplication al arrancar, no la
        ' propiedad Actual: se registra a si mismo para que ambos coincidan.
        instancia = Me
    End Sub

    Private Sub empezar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles empezar.Click
        calcularConsumo.Actual.Show()
    End Sub

    Private Sub NUEVOToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles NUEVOToolStripMenuItem.Click
        agregarelectro.Actual.Show()
    End Sub

    Private Sub MODIFICARToolStripMenuItem_Click(ByVal sender As Object, ByVal e As EventArgs) Handles MODIFICARToolStripMenuItem.Click
        opcioneselectro.Actual.Show()
    End Sub

    Private Sub NUEVOToolStripMenuItem1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles NUEVOToolStripMenuItem1.Click
        agregarpanel.Actual.Show()
    End Sub

    Private Sub MODIFICARToolStripMenuItem1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles MODIFICARToolStripMenuItem1.Click
        opcionespanel.Actual.Show()
    End Sub
End Class
