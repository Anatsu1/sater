'------------------------------------------------------------------------------
' Entradas.vb - Filtros de teclado compartidos por los formularios.
'
' Antes cada formulario repetia su propia version de "solonumeros" y de los
' KeyPress, con criterios distintos: algunos aceptaban solo coma, otros solo
' punto, y ninguno impedia escribir dos separadores decimales seguidos. Ademas
' los TextChanged usaban "If Not IsNumeric(x) And x.Contains(",")", condicion
' que casi nunca se cumple y por lo tanto nunca validaba nada.
'
' Aca queda una sola implementacion, que respeta el separador decimal del
' sistema y acepta tambien el otro (para no pelearse con el teclado numerico).
'------------------------------------------------------------------------------
Imports System.Globalization
Imports System.Windows.Forms

Module Entradas

    Private ReadOnly Property Separador() As String
        Get
            Return CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator
        End Get
    End Property

    ''' <summary>Letras, espacios y teclas de control.</summary>
    Public Sub SoloLetras(e As KeyPressEventArgs)
        e.Handled = Not (Char.IsLetter(e.KeyChar) OrElse Char.IsControl(e.KeyChar) OrElse e.KeyChar = " "c)
    End Sub

    ''' <summary>Digitos, teclas de control y un unico separador decimal.</summary>
    Public Sub SoloNumeros(e As KeyPressEventArgs, campo As TextBox)
        If Char.IsDigit(e.KeyChar) OrElse Char.IsControl(e.KeyChar) Then
            e.Handled = False
            Exit Sub
        End If

        If e.KeyChar = ","c OrElse e.KeyChar = "."c Then
            ' Un solo separador por campo, y normalizado al del sistema.
            If campo Is Nothing OrElse campo.Text.Contains(",") OrElse campo.Text.Contains(".") Then
                e.Handled = True
            Else
                e.Handled = True
                Dim posicion As Integer = campo.SelectionStart
                campo.Text = campo.Text.Insert(posicion, Separador())
                campo.SelectionStart = posicion + Separador().Length
            End If
            Exit Sub
        End If

        e.Handled = True
    End Sub

    ''' <summary>Digitos, la "x" separadora y el separador decimal (ej: 2064x1024x40).</summary>
    Public Sub SoloDimensiones(e As KeyPressEventArgs)
        e.Handled = Not (Char.IsDigit(e.KeyChar) OrElse Char.IsControl(e.KeyChar) _
                         OrElse e.KeyChar = "x"c OrElse e.KeyChar = "X"c _
                         OrElse e.KeyChar = " "c OrElse e.KeyChar = ","c OrElse e.KeyChar = "."c)
    End Sub

    ''' <summary>Letras, digitos, espacios y guiones (marcas y modelos comerciales).</summary>
    Public Sub Alfanumerico(e As KeyPressEventArgs)
        e.Handled = Not (Char.IsLetterOrDigit(e.KeyChar) OrElse Char.IsControl(e.KeyChar) _
                         OrElse e.KeyChar = " "c OrElse e.KeyChar = "-"c OrElse e.KeyChar = "."c)
    End Sub

    ''' <summary>Lee un TextBox como Double aceptando coma o punto.</summary>
    Public Function LeerNumero(texto As String, ByRef valor As Double) As Boolean
        If texto Is Nothing Then Return False
        Dim normalizado As String = texto.Trim().Replace(",", Separador()).Replace(".", Separador())
        Return Double.TryParse(normalizado, NumberStyles.Any, CultureInfo.CurrentCulture, valor)
    End Function

End Module
