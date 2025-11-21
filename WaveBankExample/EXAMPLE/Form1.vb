Imports System.IO

Public Class Form1
    Private player As Object
    Private playerType As Type

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        InitializeWBKPlayer()
    End Sub

    Private Sub InitializeWBKPlayer()
        Try
            Dim dllPath As String = Application.StartupPath & "\WAVEBANK.dll"
            Dim assembly = System.Reflection.Assembly.LoadFrom(dllPath)
            playerType = assembly.GetType("WAVEBANK.WBKPlayerLib.Player")
            player = Activator.CreateInstance(playerType)


            playerType.GetMethod("LoadWBK").Invoke(player, New Object() {"DEMOBANK.wbk"})

        Catch ex As Exception

        End Try
    End Sub


    Public Sub PlaySound(waveID As String)
        Try
            If player IsNot Nothing AndAlso playerType IsNot Nothing Then
                playerType.GetMethod("PlayWave").Invoke(player, New Object() {waveID})
            End If
        Catch ex As Exception

        End Try
    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        PlaySound("001")
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        PlaySound("002")
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        PlaySound("003")
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        PlaySound("004")
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        PlaySound("005")
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        PlaySound("006")
    End Sub
End Class