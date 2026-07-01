Imports System.Diagnostics

Module Program
    <STAThread>
    Sub Main()
        Process.Start(New ProcessStartInfo("https://app.microsaledashboard.net/auth/login") With {
            .UseShellExecute = True
        })
    End Sub
End Module
