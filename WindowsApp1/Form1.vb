Imports System.IO
Public Class Form1
    Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Integer, ByVal uParam As Integer, ByVal lpvParam As String, ByVal fuWinIni As Integer) As Integer
    Private Const SETDESKWALLPAPER = 20
    Private Const UPDATEINIFILE = &H1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If MsgBox("Wallpaper Changer v1 for Windows 7" + vbCrLf + "Powered By hominlab@gmail.com" + vbCrLf + "" + vbCrLf + "Continue?", vbInformation + MsgBoxStyle.YesNo, "Change?") = vbYes Then
            If File.Exists("C:\Windows\Web\Wallpaper\Windows\img0.jpg_bak") Then
                '배경 이미 바뀜
                If MsgBox("Already Changed wallpaper by me. Restore?", vbQuestion + MsgBoxStyle.YesNo, "Queston") = vbYes Then
                    File.Delete("C:\Windows\Web\Wallpaper\Windows\img0.jpg")
                    File.Move("C:\Windows\Web\Wallpaper\Windows\img0.jpg_bak", "C:\Windows\Web\Wallpaper\Windows\img0.jpg")
                    MsgBox("Done!", vbInformation, "yes!")
                    End
                Else
                    End
                End If
            Else
                '배경 바뀌지 않음
                If MsgBox("Change wallpaper to VSMac_1080p.jpg?", vbQuestion + MsgBoxStyle.YesNo, "Question") = vbYes Then
                    File.Move("C:\Windows\Web\Wallpaper\Windows\img0.jpg", "C:\Windows\Web\Wallpaper\Windows\img0.jpg_bak")
                    File.Copy(Application.StartupPath + "\VSMac_1080p.jpg", "C:\Windows\Web\Wallpaper\Windows\img0.jpg")
                    MsgBox("Done!", vbInformation, "yes!")
                    End
                Else
                    End
                End If
            End If
        Else
            End
        End If
    End Sub
End Class
