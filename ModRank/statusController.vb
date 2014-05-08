Imports System.Windows.Controls
Imports System.Windows.Documents
Imports System.Windows.Media

Namespace ModRank.frmMain
    Friend Class StatusController
        Private brush As Brush
        Private Property statusBox() As RichTextBox
            Get
                Return m_statusBox
            End Get
            Set(value As RichTextBox)
                m_statusBox = value
            End Set
        End Property
        Private m_statusBox As RichTextBox

        Public Sub New(statusBox As RichTextBox)
            Me.statusBox = statusBox
            Me.brush = statusBox.Foreground
        End Sub

        Public Sub Ok()
            CheckAccessAndInvoke(Sub() displayResult("OK"))
        End Sub

        Public Sub NotOK()
            CheckAccessAndInvoke(Sub() displayResult("ER"))
        End Sub

        Public Sub DisplayMessage(message As String)
            CheckAccessAndInvoke(DirectCast(Sub()
                                                Dim text As New Run(message)

                                                text.Foreground = brush
                                                text.Text = vbCr & getPaddedString(text.Text)
                                                DirectCast(statusBox.Document.Blocks.LastBlock, Paragraph).Inlines.Add(text)

                                                statusBox.ScrollToEnd()

                                            End Sub, Action))
        End Sub

        Public Sub HandleError([error] As String, toggleControls As Action)
            CheckAccessAndInvoke(DirectCast(Sub()
                                                Dim text As New Run()

                                                text.Foreground = Brushes.White
                                                text.Text = vbCr & vbCr & "["
                                                DirectCast(statusBox.Document.Blocks.LastBlock, Paragraph).Inlines.Add(text)

                                                text = New Run()
                                                text.Foreground = Brushes.Red
                                                text.Text = "Error"
                                                DirectCast(statusBox.Document.Blocks.LastBlock, Paragraph).Inlines.Add(text)

                                                text = New Run()
                                                text.Foreground = Brushes.White
                                                text.Text = "] "
                                                DirectCast(statusBox.Document.Blocks.LastBlock, Paragraph).Inlines.Add(text)

                                                text = New Run([error] & vbCr & vbCr)
                                                text.Foreground = Brushes.White
                                                DirectCast(statusBox.Document.Blocks.LastBlock, Paragraph).Inlines.Add(text)

                                                statusBox.ScrollToEnd()
                                                toggleControls()

                                            End Sub, Action))
        End Sub

        Private Sub displayResult(message As String)
            Dim text As New Run()

            text.Foreground = Brushes.White
            text.Text = "["
            DirectCast(statusBox.Document.Blocks.LastBlock, Paragraph).Inlines.Add(text)

            text = New Run()
            text.Text = message
            DirectCast(statusBox.Document.Blocks.LastBlock, Paragraph).Inlines.Add(text)

            text = New Run()
            text.Foreground = Brushes.White
            text.Text = "]"

            DirectCast(statusBox.Document.Blocks.LastBlock, Paragraph).Inlines.Add(text)

            statusBox.ScrollToEnd()
        End Sub

        Private Function getPaddedString(text As String) As String
            Return text.PadRight(90, " "c)
        End Function

        Private Sub CheckAccessAndInvoke(a As Action)
            If statusBox.Dispatcher.CheckAccess() Then
                a()
                Return
            End If

            statusBox.Dispatcher.Invoke(a)
        End Sub
    End Class
End Namespace