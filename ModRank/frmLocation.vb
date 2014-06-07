Public Class frmLocation
    Public X As Integer = 0
    Public Y As Integer = 0
    Public W As Integer = 0
    Public H As Integer = 0
    Public TabName As String = ""
    Public ItemName As String = ""

    Private Sub PictureBox1_Paint(sender As Object, e As PaintEventArgs) Handles PictureBox1.Paint
        Dim x As Integer, y As Integer
        Dim img As Image = New Bitmap(240, 240)
        Dim graphics As Graphics = graphics.FromImage(img)
        Dim rect As Rectangle = New Rectangle(Me.X * 20, Me.Y * 20, Me.W * 20, Me.H * 20)
        graphics.FillRectangle(Brushes.LightSkyBlue, rect)
        For x = 0 To 240 Step 20
            graphics.DrawLine(Pens.Black, x, 0, x, 240)
        Next
        For y = 0 To 240 Step 20
            graphics.DrawLine(Pens.Black, 0, y, 240, y)
        Next
        PictureBox1.BackgroundImage = img
    End Sub

    Private Sub frmLocation_Load(sender As Object, e As EventArgs) Handles Me.Load
        lblTabName.Text = Me.TabName
        Me.Text = "Item Location: " & Me.ItemName
    End Sub

    Private Sub lblTabName_Paint(sender As Object, e As PaintEventArgs) Handles lblTabName.Paint
        With e.Graphics
            .DrawRectangle(New Pen(Color.Black, 1.0!), 0I, 0I, lblTabName.Width - 1I, lblTabName.Height - 1I)
        End With
    End Sub
End Class