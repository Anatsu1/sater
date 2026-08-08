'------------------------------------------------------------------------------
' Arranque.vb - Punto de entrada alternativo para compilar con Mono / vbnc.
'
' NO forma parte de PanelesSolares.vbproj: Visual Studio lo ignora por completo.
'
' El proyecto arranca con el Application Framework de Visual Basic
' (My.MyApplication, generado a partir de My Project/Application.myapp), que el
' compilador vbnc de Mono no implementa. Este modulo hace exactamente lo mismo
' que hacia ese framework segun Application.myapp: habilita estilos visuales y
' ejecuta el formulario principal.
'------------------------------------------------------------------------------
Module Arranque

    <STAThread()> Public Sub Main()
        System.Windows.Forms.Application.EnableVisualStyles()
        System.Windows.Forms.Application.SetCompatibleTextRenderingDefault(False)
        System.Windows.Forms.Application.Run(New inicio())
    End Sub

End Module
