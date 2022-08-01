'Imports System.IO
Public Class Form1

    Private Function TakeScreenShot() As Bitmap

        'Get positions of screenshot on screen

        Dim screenPositionX = Me.PointToScreen(PictureBoxScreenshot.Location).X
        Dim screenPositionY = Me.PointToScreen(PictureBoxScreenshot.Location).Y
        Dim screenWidth = PictureBoxScreenshot.Width
        Dim screenHeight = PictureBoxScreenshot.Height

        Dim ScreenLeft = screenPositionX
        Dim ScreenRight = screenPositionX + screenWidth
        Dim ScreenTop = screenPositionY
        Dim ScreenBottom = screenPositionY + screenHeight

        'Define Rectangles for cropping

        Dim src_rect As New Rectangle(0, 0, My.Computer.Screen.Bounds.Width, My.Computer.Screen.Bounds.Height)
        Dim dst_rect As New Rectangle(screenPositionX, screenPositionY, My.Computer.Screen.Bounds.Width, My.Computer.Screen.Bounds.Height)
        Dim screenSize As Size = New Size(screenWidth, screenHeight)

        'take initial screenshot of entire screen

        Dim bmp_src As New Bitmap(My.Computer.Screen.Bounds.Width, My.Computer.Screen.Bounds.Height)
        Dim bmp_dst As New Bitmap(screenWidth, screenHeight)
        Dim g As Graphics = Graphics.FromImage(bmp_src)
        g.CopyFromScreen(New Point(ScreenLeft, ScreenTop), New Point(ScreenLeft, ScreenTop), screenSize)

        'crop bitmap to just selection

        Dim gr_dest As Graphics = Graphics.FromImage(bmp_dst)
        gr_dest.DrawImage(bmp_src, src_rect, dst_rect, GraphicsUnit.Pixel)

        'return cropped image

        Return bmp_dst

    End Function

    Private Sub ButtonScreenshot_Click_1(sender As Object, e As EventArgs) Handles ButtonScreenshot.Click

        Dim fileNameData
        Dim inputData

        fileNameData = DateTime.Now.ToString("yyyyMMddHHmmss")
        fileNameData = fileNameData.Insert(8, "_")

        Dim bmp As Bitmap = TakeScreenShot()

        If RadioButtonBMP.Checked = True Then

            inputData = InputBox("Type Title", "Title Prompt", " ")

            If inputData = "" Then
                Exit Sub
            End If

            fileNameData = Trim(fileNameData & " " & inputData)
            fileNameData = fileNameData.Replace(" ", "_")
            'My.Computer.Clipboard.GetText()
            'fileNameData = fileNameData.Split({vbLf}, StringSplitOptions.TrimEntries)(0)

            bmp.Save("images/" & fileNameData & ".bmp")
            bmp.SetResolution(220, 220)

        Else
            Clipboard.SetImage(bmp)

        End If

    End Sub

    Private Sub PictureBoxScreenshot_Click(sender As Object, e As EventArgs) Handles PictureBoxScreenshot.Click
        'Dim path As String = Directory.GetCurrentDirectory()
        Process.Start("explorer.exe", "images")
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        RadioButtonClipboard.Checked = True
    End Sub

End Class
