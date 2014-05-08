Public Class frmWeights
    Dim dtTemp As DataTable

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    Private Sub frmWeights_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            Me.Icon = GetEmbeddedIcon("ModRank.PoE.ico")
            dtTemp = frmMain.dtWeights.Clone
            dtTemp.Rows.Clear()
            For Each row In frmMain.dtWeights.Rows
                dtTemp.ImportRow(DeepCopyDataRow(row))
            Next
            dgWeights.DataSource = dtTemp
            dgWeights.Columns(0).ReadOnly = True
            dgWeights.Columns(0).Width = 210
            dgWeights.Columns(1).Width = 50
            dgWeights.Columns(2).ReadOnly = True
            dgWeights.Columns(2).Width = 210
            dgWeights.Columns(3).ReadOnly = True
            dgWeights.Columns(3).Width = 200

            dgWeights.Focus()
        Catch ex As Exception
            ErrorHandler(System.Reflection.MethodBase.GetCurrentMethod.Name, ex)
        End Try
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Try
            Dim sfd As New SaveFileDialog()
            sfd.InitialDirectory = Application.StartupPath
            sfd.CreatePrompt = False
            sfd.OverwritePrompt = True
            sfd.FileName = "weights-" & frmMain.cmbWeight.Text & ".csv"
            sfd.Filter = "Text File|*.csv"
            sfd.Title = "Save Weights File"
            If sfd.ShowDialog() = Windows.Forms.DialogResult.Cancel Then Exit Sub

            Dim bytResult As Byte = 0
            If sfd.FileName = "" Then Exit Sub
            If System.IO.Path.GetDirectoryName(sfd.FileName) <> Application.StartupPath Then
                bytResult = MsgBox("Only weights files located in the application folder will be used. Are you sure you want to save to another folder?", MsgBoxStyle.YesNo, "Save to Another Folder?")
                If bytResult = vbNo Then Exit Sub
            End If
            If System.IO.Path.GetFileNameWithoutExtension(sfd.FileName).ToLower.Substring(0, Math.Min(System.IO.Path.GetFileNameWithoutExtension(sfd.FileName).Length, 8)) <> "weights-" Then
                bytResult = MsgBox("A weights file must have 'weights-' in the start of the filename. Do you wish to add this text to the filename?", MsgBoxStyle.YesNoCancel, "Add Weights Text to Filename?")
                If bytResult = vbCancel Then Exit Sub
                If bytResult = vbYes Then sfd.FileName = System.IO.Path.GetDirectoryName(sfd.FileName) & "\weights-" & System.IO.Path.GetFileName(sfd.FileName)
            End If

            TableToCSV(dtTemp, sfd.FileName, True)

            'frmMain.dtWeights.Clear()
            'frmMain.dtWeights = LoadCSVtoDataTable(sfd.FileName)

            Dim strWeight As String = frmMain.cmbWeight.Text
            frmMain.cmbWeight.Items.Clear()
            frmMain.PopulateWeightsComboBox(strWeight)

            frmMain.blRepopulated = True

            Me.Close()
        Catch ex As Exception
            ErrorHandler(System.Reflection.MethodBase.GetCurrentMethod.Name, ex)
        End Try
    End Sub

    Private Sub dgWeights_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles dgWeights.CellEndEdit
        If Not IsNumeric(dgWeights.Rows(e.RowIndex).Cells(e.ColumnIndex).Value) Then
            MsgBox("You must enter a numeric value for a weight.", MsgBoxStyle.Critical, "Invalid Input")
            dgWeights.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = 0
        End If
    End Sub
End Class