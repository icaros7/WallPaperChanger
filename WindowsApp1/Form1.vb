Imports System.IO
Public Class Form1
    Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Integer, ByVal uParam As Integer, ByVal lpvParam As String, ByVal fuWinIni As Integer) As Integer
    Private Const SETDESKWALLPAPER = 20
    Private Const UPDATEINIFILE = &H1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If MsgBox("Wallpaper Changer v1 for Windows 7" + vbCrLf + "Powered By hominlab@gmail.com" + vbCrLf + "" + vbCrLf + "Change the Wallpaper to VSMac?", vbQuestion + MsgBoxStyle.YesNo, "Change?") = vbYes Then
            If Not File.Exists("C:\VSMac_1080p.png") Then
                File.Copy(Application.StartupPath + "\VSMac_1080p.png", "C:\VSMac_1080p.png")
            End If
            Try
                SystemParametersInfo(SETDESKWALLPAPER, 0, "C:\VSMac_1080p.png", UPDATEINIFILE)
                MsgBox("Done!", vbInformation, "output")
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        Else
            End
        End If
    End Sub
End Class
