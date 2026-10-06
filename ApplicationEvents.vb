Namespace My
    ' Los siguientes eventos están disponibles para MyApplication:
    ' Inicio: Se genera cuando se inicia la aplicación, antes de que se cree el formulario de inicio.
    ' Apagado: Se genera después de haberse cerrado todos los formularios de aplicación.  Este evento no se genera si la aplicación termina de forma anómala.
    ' UnhandledException: Se genera si la aplicación encuentra una excepción no controlada.
    ' StartupNextInstance: Se genera cuando se inicia una aplicación de instancia única y dicha aplicación está ya activa.
    ' NetworkAvailabilityChanged: Se genera cuando se conecta o desconecta la conexión de red.
    Partial Friend Class MyApplication

        ' Modo headless (pedido 2026-10-06): antes, la generacion automatica diaria dependia por
        ' completo de que alguien dejara esta ventana abierta para que Timer1_Tick (Genrador.vb)
        ' disparara a las 7am -- si el exe no estaba corriendo a esa hora, no se generaba nada y
        ' nadie se enteraba (asi paso el 2026-09-30/10-01, 663 contratos sin estado de cuenta).
        ' Con "AppEstadosCuenta.exe auto" (Task Scheduler real, sin ventana) se corre la misma
        ' logica (EjecutarGeneracionAutomatica en Genrador.vb) y el proceso termina solo -- el modo
        ' normal (doble clic, sin argumentos) sigue mostrando la ventana exactamente igual que antes.
        Private Sub MyApplication_Startup(sender As Object, e As Microsoft.VisualBasic.ApplicationServices.StartupEventArgs) Handles Me.Startup
            If e.CommandLine.Count > 0 AndAlso e.CommandLine(0).Trim().ToLower() = "auto" Then
                e.Cancel = True
                Dim frm As New Generador()
                frm.EjecutarGeneracionAutomatica()
                frm.Dispose()
            End If
        End Sub
    End Class
End Namespace
