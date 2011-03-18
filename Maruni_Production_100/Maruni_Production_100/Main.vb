
Option Explicit On
Option Strict Off

Module Main

    Public Sub Main()

        Dim SBOConn As SBOConnection
        Dim ClsName As String
        ClsName = "Main"

        Try

            SBOConn = New SBOConnection
            'MsgBox("Connected! hi!")

            Dim SOToMFG As SOToMFG
            SOToMFG = New SOToMFG

            System.Windows.Forms.Application.Run()

        Catch ex As Exception
            MsgBox(ClsName & " Addon failed! SBO Not Running!")
        End Try

    End Sub

End Module
