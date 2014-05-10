Public Class frmResults
    Private _MyData As CloneableList(Of String)
    Public Property MyData() As CloneableList(Of String)
        Get
            Return _MyData
        End Get
        Set(value As CloneableList(Of String))
            _MyData = value
            RunQuery()
        End Set
    End Property

    Private Sub RunQuery()
        Try
            Dim query = From myItem In frmMain.dtOverflow Where myItem("ID") = MyData(0) And myItem("Name") = MyData(1)
            Dim dt As DataTable = query.CopyToDataTable()
            ' Add the row that is selected on the main form
            Dim query2 = From myItem In frmMain.dtRank Where myItem("ID") = MyData(0) And myItem("Name") = MyData(1)
            Dim dt2 As DataTable = query2.CopyToDataTable
            dt.ImportRow(dt2.Rows(0))
            Me.DataGridView1.DataSource = dt
            Me.DataGridView1.Columns("ID").Visible = False
            Me.DataGridView1.Columns("Index").Visible = False
            Me.DataGridView1.Columns("%").DefaultCellStyle.Format = "n1"
            frmMain.SetDataGridViewWidths(DataGridView1)
        Catch ex As Exception
            ErrorHandler(System.Reflection.MethodBase.GetCurrentMethod.Name, ex)
        End Try
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.RowIndex = -1 Then Exit Sub
        If e.ColumnIndex = DataGridView1.Columns("Rank").Index Then
            Dim sb As New System.Text.StringBuilder
            If DataGridView1.CurrentRow.Cells("*").Value = "*" Then
                If frmMain.FullInventory(DataGridView1.Rows(e.RowIndex).Cells("Index").Value).LevelGem = True Then frmMain.AddGemWarning(sb)
                sb.Append(frmMain.RankExplanation(DataGridView1.CurrentRow.Cells("ID").Value & DataGridView1.CurrentRow.Cells("Name").Value))
            Else
                If frmMain.TempInventory(DataGridView1.Rows(e.RowIndex).Cells("Index").Value).LevelGem = True Then frmMain.AddGemWarning(sb)
                sb.Append(frmMain.RankExplanation(DataGridView1.CurrentRow.Cells("ID").Value & DataGridView1.CurrentRow.Cells("Name").Value & DataGridView1.CurrentRow.Cells("Index").Value))
            End If
            MsgBox(sb.ToString, , "Item Mod Rank Explanation - " & DataGridView1.CurrentRow.Cells("Name").Value)
        ElseIf DataGridView1.Columns(e.ColumnIndex).Name.ToLower.Contains("prefix") AndAlso NotNull(DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value, "") <> "" AndAlso DataGridView1.Rows(e.RowIndex).Cells("*").Value = "*" Then
            frmMain.ShowModInfo(DataGridView1, frmMain.FullInventory(DataGridView1.Rows(e.RowIndex).Cells("Index").Value), frmMain.FullInventory(DataGridView1.Rows(e.RowIndex).Cells("Index").Value).ExplicitPrefixMods, GetNumeric(DataGridView1.Columns(e.ColumnIndex).Name) - 1, e)
        ElseIf DataGridView1.Columns(e.ColumnIndex).Name.ToLower.Contains("prefix") AndAlso NotNull(DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value, "") <> "" AndAlso DataGridView1.Rows(e.RowIndex).Cells("*").Value = "" Then
            frmMain.ShowModInfo(DataGridView1, frmMain.TempInventory(DataGridView1.Rows(e.RowIndex).Cells("Index").Value), frmMain.TempInventory(DataGridView1.Rows(e.RowIndex).Cells("Index").Value).ExplicitPrefixMods, GetNumeric(DataGridView1.Columns(e.ColumnIndex).Name) - 1, e)
        ElseIf DataGridView1.Columns(e.ColumnIndex).Name.ToLower.Contains("suffix") AndAlso NotNull(DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value, "") <> "" AndAlso DataGridView1.Rows(e.RowIndex).Cells("*").Value = "*" Then
            frmMain.ShowModInfo(DataGridView1, frmMain.FullInventory(DataGridView1.Rows(e.RowIndex).Cells("Index").Value), frmMain.FullInventory(DataGridView1.Rows(e.RowIndex).Cells("Index").Value).ExplicitSuffixMods, GetNumeric(DataGridView1.Columns(e.ColumnIndex).Name) - 1, e)
        ElseIf DataGridView1.Columns(e.ColumnIndex).Name.ToLower.Contains("suffix") AndAlso NotNull(DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value, "") <> "" AndAlso DataGridView1.Rows(e.RowIndex).Cells("*").Value = "" Then
            frmMain.ShowModInfo(DataGridView1, frmMain.TempInventory(DataGridView1.Rows(e.RowIndex).Cells("Index").Value), frmMain.TempInventory(DataGridView1.Rows(e.RowIndex).Cells("Index").Value).ExplicitSuffixMods, GetNumeric(DataGridView1.Columns(e.ColumnIndex).Name) - 1, e)
        End If
    End Sub

    Private Sub DataGridView1_CellMouseEnter(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellMouseEnter
        If frmMain.IsValidCellAddress(DataGridView1, e.RowIndex, e.ColumnIndex) AndAlso _
            (e.ColumnIndex = 0 Or DataGridView1.Columns(e.ColumnIndex).Name.Contains("fix")) Then DataGridView1.Cursor = Cursors.Hand
    End Sub

    Private Sub DataGridView1_CellMouseLeave(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellMouseLeave
        If frmMain.IsValidCellAddress(DataGridView1, e.RowIndex, e.ColumnIndex) AndAlso _
            (e.ColumnIndex = 0 Or DataGridView1.Columns(e.ColumnIndex).Name.Contains("fix")) Then DataGridView1.Cursor = Cursors.Default
    End Sub

    Private Sub DataGridView1_CellPainting(sender As Object, e As DataGridViewCellPaintingEventArgs) Handles DataGridView1.CellPainting
        If e.RowIndex = -1 Then
            frmMain.DataGridViewCellPaintingHeaderFormat(sender, e)
            Exit Sub
        End If
        Dim strName As String = DataGridView1.Columns(e.ColumnIndex).Name.ToLower
        Dim lngIndex As Long = DataGridView1.Rows(e.RowIndex).Cells("Index").Value
        If strName.Contains("prefix") AndAlso NotNull(DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value, "") <> "" AndAlso DataGridView1.Rows(e.RowIndex).Cells("*").Value = "*" Then
            frmMain.DataGridViewAddLevelBar(DataGridView1, frmMain.FullInventory(lngIndex).ExplicitPrefixMods, strName, sender, e)
        ElseIf strName.Contains("prefix") AndAlso NotNull(DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value, "") <> "" AndAlso DataGridView1.Rows(e.RowIndex).Cells("*").Value = "" Then
            frmMain.DataGridViewAddLevelBar(DataGridView1, frmMain.TempInventory(lngIndex).ExplicitPrefixMods, strName, sender, e)
        ElseIf strName.Contains("suffix") AndAlso NotNull(DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value, "") <> "" AndAlso DataGridView1.Rows(e.RowIndex).Cells("*").Value = "*" Then
            frmMain.DataGridViewAddLevelBar(DataGridView1, frmMain.FullInventory(lngIndex).ExplicitSuffixMods, strName, sender, e)
        ElseIf strName.Contains("suffix") AndAlso NotNull(DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value, "") <> "" AndAlso DataGridView1.Rows(e.RowIndex).Cells("*").Value = "" Then
            frmMain.DataGridViewAddLevelBar(DataGridView1, frmMain.TempInventory(lngIndex).ExplicitSuffixMods, strName, sender, e)
        End If
    End Sub

    Private Sub DataGridView1_RowPostPaint(sender As Object, e As DataGridViewRowPostPaintEventArgs) Handles DataGridView1.RowPostPaint
        If e.RowIndex = -1 Then Exit Sub
        If DataGridView1.Rows(e.RowIndex).Cells("*").Value = "*" Then
            frmMain.DataGridViewRowPostPaint(DataGridView1, frmMain.FullInventory, sender, e)
        Else
            frmMain.DataGridViewRowPostPaint(DataGridView1, frmMain.TempInventory, sender, e)
        End If
    End Sub

    Private Sub DataGridView1_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView1.SelectionChanged
        If DataGridView1.CurrentCell Is Nothing Then Exit Sub
        If DataGridView1.Columns(DataGridView1.CurrentCell.ColumnIndex).Name.ToLower.Contains("prefix") Or DataGridView1.Columns(DataGridView1.CurrentCell.ColumnIndex).Name.ToLower.Contains("suffix") Then DataGridView1.ClearSelection()
    End Sub

    Private Sub frmResults_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.Icon = GetEmbeddedIcon("ModRank.PoE.ico")
        Me.Left = 0
        Me.Width = Screen.PrimaryScreen.Bounds.Width
    End Sub
End Class