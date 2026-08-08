'------------------------------------------------------------------------------
' BaseDatos.vb - Capa unica de acceso a datos.
'
' Centraliza la conexion, las consultas y los comandos del sistema. Antes cada
' formulario abria su propia OleDbConnection contra una ruta relativa y armaba
' el SQL concatenando texto; eso rompia los INSERT, dejaba conexiones abiertas y
' habilitaba inyeccion SQL desde los campos de busqueda.
'
' El proveedor se elige en App.config (clave "proveedor"):
'   auto   -> oledb en Windows, sqlite en Linux/macOS           (por defecto)
'   oledb  -> Microsoft.Jet.OLEDB.4.0 sobre paneleSolares.mdb   (Windows)
'   sqlite -> Mono.Data.Sqlite sobre paneleSolares.db           (Linux / demo)
'
' El SQL se escribe siempre con marcadores @p0, @p1, ... en orden ascendente.
' Para OleDb -que solo entiende parametros posicionales- se traducen a "?".
'------------------------------------------------------------------------------
Imports System.Configuration
Imports System.Data
Imports System.Data.Common
Imports System.IO
Imports System.Text.RegularExpressions

Module BaseDatos

    Private ReadOnly proveedorActual As String = ResolverProveedor()
    Private ReadOnly archivoBase As String = Ajuste("archivoBase", ArchivoPorDefecto())
    Private ReadOnly marcadores As New Regex("@p\d+", RegexOptions.Compiled)
    Private fabricaCache As DbProviderFactory

    Private Function Ajuste(clave As String, porDefecto As String) As String
        Dim valor As String = ConfigurationManager.AppSettings(clave)
        If String.IsNullOrEmpty(valor) Then Return porDefecto
        Return valor
    End Function

    Private Function ArchivoPorDefecto() As String
        If proveedorActual = "sqlite" Then
            Return "paneleSolares.db"
        End If
        Return "paneleSolares.mdb"
    End Function

    ' Jet OLEDB solo existe en Windows: fuera de Windows se usa SQLite.
    Private Function ResolverProveedor() As String
        Dim elegido As String = Ajuste("proveedor", "auto").Trim().ToLowerInvariant()
        If elegido <> "auto" Then Return elegido

        Select Case Environment.OSVersion.Platform
            Case PlatformID.Unix, PlatformID.MacOSX
                Return "sqlite"
            Case Else
                Return "oledb"
        End Select
    End Function

    ''' <summary>Ruta absoluta al archivo de base, junto al ejecutable.</summary>
    ''' <remarks>
    ''' Antes era "Data Source=paneleSolares.mdb" (relativa): la app solo encontraba
    ''' la base si el directorio de trabajo coincidia con el del .exe.
    ''' </remarks>
    Public ReadOnly Property Ruta() As String
        Get
            Return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, archivoBase)
        End Get
    End Property

    Public ReadOnly Property UsaSqlite() As Boolean
        Get
            Return proveedorActual = "sqlite"
        End Get
    End Property

    Private Function Fabrica() As DbProviderFactory
        If fabricaCache IsNot Nothing Then Return fabricaCache

        If UsaSqlite() Then
            fabricaCache = FabricaSqlite()
        Else
            fabricaCache = DbProviderFactories.GetFactory("System.Data.OleDb")
        End If
        Return fabricaCache
    End Function

    ' Mono.Data.Sqlite no siempre esta registrado en machine.config, asi que se
    ' intenta primero por el registro estandar y despues por reflexion.
    Private Function FabricaSqlite() As DbProviderFactory
        Dim candidatos() As String = {"Mono.Data.Sqlite", "System.Data.SQLite"}

        For Each nombre As String In candidatos
            Try
                Return DbProviderFactories.GetFactory(nombre)
            Catch
                ' se prueba el siguiente
            End Try
        Next

        For Each nombre As String In candidatos
            Try
                Dim ensamblado As Reflection.Assembly = Reflection.Assembly.Load(nombre)
                Dim sufijo As String = ".SQLiteFactory"
                If nombre.StartsWith("Mono") Then sufijo = ".SqliteFactory"
                Dim tipo As Type = ensamblado.GetType(nombre & sufijo)
                If tipo IsNot Nothing Then
                    Dim campo As Reflection.FieldInfo = tipo.GetField("Instance")
                    If campo IsNot Nothing Then Return DirectCast(campo.GetValue(Nothing), DbProviderFactory)
                End If
            Catch
                ' se prueba el siguiente
            End Try
        Next

        Throw New InvalidOperationException("No se encontro un proveedor SQLite (Mono.Data.Sqlite o System.Data.SQLite).")
    End Function

    Private Function CadenaConexion() As String
        If UsaSqlite() Then
            Return "Data Source=" & Ruta() & ";Version=3;"
        End If
        Return "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Ruta() & ";Persist Security Info=False"
    End Function

    ''' <summary>Abre una conexion nueva. El llamador la cierra (usar Using).</summary>
    Public Function Abrir() As DbConnection
        Dim conexion As DbConnection = Fabrica().CreateConnection()
        conexion.ConnectionString = CadenaConexion()
        conexion.Open()
        Return conexion
    End Function

    ''' <summary>Ejecuta un SELECT y devuelve el resultado como DataTable.</summary>
    Public Function Consultar(sql As String, ParamArray valores() As Object) As DataTable
        Using conexion As DbConnection = Abrir()
            Using comando As DbCommand = CrearComando(conexion, sql, valores)
                Dim tabla As New DataTable()
                Using lector As DbDataReader = comando.ExecuteReader()
                    tabla.Load(lector)
                End Using
                Return tabla
            End Using
        End Using
    End Function

    ''' <summary>Ejecuta INSERT / UPDATE / DELETE y devuelve las filas afectadas.</summary>
    Public Function Ejecutar(sql As String, ParamArray valores() As Object) As Integer
        Using conexion As DbConnection = Abrir()
            Using comando As DbCommand = CrearComando(conexion, sql, valores)
                Return comando.ExecuteNonQuery()
            End Using
        End Using
    End Function

    Private Function CrearComando(conexion As DbConnection, sql As String, valores() As Object) As DbCommand
        Dim comando As DbCommand = conexion.CreateCommand()

        ' OleDb ignora los nombres y liga los parametros por posicion: los
        ' marcadores @pN se reemplazan por "?" respetando el orden de aparicion.
        If UsaSqlite() Then
            comando.CommandText = sql
        Else
            comando.CommandText = marcadores.Replace(sql, "?")
        End If

        If valores IsNot Nothing Then
            For i As Integer = 0 To valores.Length - 1
                Dim parametro As DbParameter = comando.CreateParameter()
                parametro.ParameterName = "@p" & i
                If valores(i) Is Nothing Then
                    parametro.Value = DBNull.Value
                Else
                    parametro.Value = valores(i)
                End If
                comando.Parameters.Add(parametro)
            Next
        End If

        Return comando
    End Function

    ''' <summary>
    ''' Mensaje uniforme para fallos de base. Muestra la causa real en vez del
    ''' "error" generico que enmascaraba los bugs de INSERT.
    ''' </summary>
    Public Sub Reportar(ex As Exception, titulo As String)
        MsgBox(ex.Message, MsgBoxStyle.Critical, titulo)
    End Sub

End Module
