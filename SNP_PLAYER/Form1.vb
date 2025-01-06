Imports System.ComponentModel

Public Class Form1

    Private c As New DSHOW_PLAYER

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Try
            c.FileName = "f:\vs\prueba.mp3"
            If c.Open() = False Then

            End If

            c.Play()
        Catch ex As InvalidOperationException
            MessageBox.Show("problema al abrir el archivo")
            Return
        End Try


    End Sub

    Private Sub Form1_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        Try
            c.Stop()
            c.Dispose()
        Catch ex As Exception
            Return
        End Try
    End Sub
End Class
