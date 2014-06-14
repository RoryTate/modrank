Imports ModRank.Extenders.Strings.StringExtender

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

    Private FullInv As List(Of FullItem), TempInv As List(Of FullItem)
    Public blStore As Boolean

    Private Sub RunQuery()
        Try
            Dim dt As DataTable
            If blStore = True Then
                Dim query = From myItem In dtStoreOverflow Where myItem("ID").ToString = MyData(0) And myItem("Name").ToString = MyData(1)
                dt = query.CopyToDataTable()
                ' Add the row that is selected on the main form
                Dim query2 = From myItem In dtStore Where myItem("ID").ToString = MyData(0) And myItem("Name").ToString = MyData(1)
                Dim dt2 As DataTable = query2.CopyToDataTable
                dt.ImportRow(dt2.Rows(0))
            Else
                Dim query = From myItem In dtOverflow Where myItem("ID").ToString = MyData(0) And myItem("Name").ToString = MyData(1)
                dt = query.CopyToDataTable()
                ' Add the row that is selected on the main form
                Dim query2 = From myItem In dtRank Where myItem("ID").ToString = MyData(0) And myItem("Name").ToString = MyData(1)
                Dim dt2 As DataTable = query2.CopyToDataTable
                dt.ImportRow(dt2.Rows(0))
            End If
            Me.DataGridView1.DataSource = dt
            Me.Invoke(New MyDualControlDelegate(AddressOf HideColumns), New Object() {Me, DataGridView1})
            Me.DataGridView1.Columns("%").DefaultCellStyle.Format = "n1"
            SetDataGridViewWidths(DataGridView1)
            ' To make room for new Price column take away from SubType and Implicit columns
            Dim intWidth As Integer = Math.Max(CType(Me.Invoke(New RCPD(AddressOf ReadControlProperty), New Object() {DataGridView1.Columns("SubType"), "Width"}), Integer) - 15, 0)
            Me.Invoke(New UCPD(AddressOf SetControlProperty), New Object() {DataGridView1.Columns("SubType"), "Width", intWidth})
            intWidth = Math.Max(CType(Me.Invoke(New RCPD(AddressOf ReadControlProperty), New Object() {DataGridView1.Columns("Implicit"), "Width"}), Integer) - 35, 0)
            Me.Invoke(New UCPD(AddressOf SetControlProperty), New Object() {DataGridView1.Columns("Implicit"), "Width", intWidth})
        Catch ex As Exception
            ErrorHandler(System.Reflection.MethodBase.GetCurrentMethod.Name, ex)
        End Try
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.RowIndex = -1 Then Exit Sub
        If e.ColumnIndex = DataGridView1.Columns("Rank").Index Then
            Dim sb As New System.Text.StringBuilder
            If DataGridView1.CurrentRow.Cells("*").Value.ToString = "*" Then
                If FullInv(CInt(DataGridView1.Rows(e.RowIndex).Cells("Index").Value)).LevelGem = True Then frmMain.AddGemWarning(sb)
                sb.Append(frmMain.RankExplanation(DataGridView1.CurrentRow.Cells("ID").Value.ToString & DataGridView1.CurrentRow.Cells("Name").Value.ToString))
            Else
                If TempInv(CInt(DataGridView1.Rows(e.RowIndex).Cells("Index").Value)).LevelGem = True Then frmMain.AddGemWarning(sb)
                sb.Append(frmMain.RankExplanation(DataGridView1.CurrentRow.Cells("ID").Value.ToString & DataGridView1.CurrentRow.Cells("Name").Value.ToString & DataGridView1.CurrentRow.Cells("Index").Value.ToString))
            End If
            MsgBox(sb.ToString, , "Item Mod Rank Explanation - " & DataGridView1.CurrentRow.Cells("Name").Value.ToString)
        ElseIf DataGridView1.Columns(e.ColumnIndex).Name.ToLower.Contains("prefix") AndAlso NotNull(DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value, "").ToString <> "" AndAlso DataGridView1.Rows(e.RowIndex).Cells("*").Value.ToString = "*" Then
            frmMain.ShowModInfo(DataGridView1, FullInv(CInt(DataGridView1.Rows(e.RowIndex).Cells("Index").Value)), FullInv(CInt(DataGridView1.Rows(e.RowIndex).Cells("Index").Value)).ExplicitPrefixMods, CInt(GetNumeric(DataGridView1.Columns(e.ColumnIndex).Name) - 1), e)
        ElseIf DataGridView1.Columns(e.ColumnIndex).Name.ToLower.Contains("prefix") AndAlso NotNull(DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value, "").ToString <> "" AndAlso DataGridView1.Rows(e.RowIndex).Cells("*").Value.ToString = "" Then
            frmMain.ShowModInfo(DataGridView1, TempInv(CInt(DataGridView1.Rows(e.RowIndex).Cells("Index").Value)), TempInv(CInt(DataGridView1.Rows(e.RowIndex).Cells("Index").Value)).ExplicitPrefixMods, CInt(GetNumeric(DataGridView1.Columns(e.ColumnIndex).Name) - 1), e)
        ElseIf DataGridView1.Columns(e.ColumnIndex).Name.ToLower.Contains("suffix") AndAlso NotNull(DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value, "").ToString <> "" AndAlso DataGridView1.Rows(e.RowIndex).Cells("*").Value.ToString = "*" Then
            frmMain.ShowModInfo(DataGridView1, FullInv(CInt(DataGridView1.Rows(e.RowIndex).Cells("Index").Value)), FullInv(CInt(DataGridView1.Rows(e.RowIndex).Cells("Index").Value)).ExplicitSuffixMods, CInt(GetNumeric(DataGridView1.Columns(e.ColumnIndex).Name) - 1), e)
        ElseIf DataGridView1.Columns(e.ColumnIndex).Name.ToLower.Contains("suffix") AndAlso NotNull(DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value, "").ToString <> "" AndAlso DataGridView1.Rows(e.RowIndex).Cells("*").Value.ToString = "" Then
            frmMain.ShowModInfo(DataGridView1, TempInv(CInt(DataGridView1.Rows(e.RowIndex).Cells("Index").Value)), TempInv(CInt(DataGridView1.Rows(e.RowIndex).Cells("Index").Value)).ExplicitSuffixMods, CInt(GetNumeric(DataGridView1.Columns(e.ColumnIndex).Name) - 1), e)
        ElseIf DataGridView1.Columns(e.ColumnIndex).Name.ToLower.Contains("prefix") AndAlso NotNull(DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value.ToString, "") = "" AndAlso DataGridView1.Rows(e.RowIndex).Cells("*").Value.ToString = "*" Then
            frmMain.ShowAllPossibleMods(DataGridView1, FullInv(CInt(DataGridView1.Rows(e.RowIndex).Cells("Index").Value)), FullInv(CInt(DataGridView1.Rows(e.RowIndex).Cells("Index").Value)).ExplicitPrefixMods, "Prefix")
        ElseIf DataGridView1.Columns(e.ColumnIndex).Name.ToLower.Contains("prefix") AndAlso NotNull(DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value.ToString, "") = "" AndAlso DataGridView1.Rows(e.RowIndex).Cells("*").Value.ToString = "" Then
            frmMain.ShowAllPossibleMods(DataGridView1, TempInv(CInt(DataGridView1.Rows(e.RowIndex).Cells("Index").Value)), TempInv(CInt(DataGridView1.Rows(e.RowIndex).Cells("Index").Value)).ExplicitPrefixMods, "Prefix")
        ElseIf DataGridView1.Columns(e.ColumnIndex).Name.ToLower.Contains("suffix") AndAlso NotNull(DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value.ToString, "") = "" AndAlso DataGridView1.Rows(e.RowIndex).Cells("*").Value.ToString = "*" Then
            frmMain.ShowAllPossibleMods(DataGridView1, FullInv(CInt(DataGridView1.Rows(e.RowIndex).Cells("Index").Value)), FullInv(CInt(DataGridView1.Rows(e.RowIndex).Cells("Index").Value)).ExplicitSuffixMods, "Suffix")
        ElseIf DataGridView1.Columns(e.ColumnIndex).Name.ToLower.Contains("suffix") AndAlso NotNull(DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value.ToString, "") = "" AndAlso DataGridView1.Rows(e.RowIndex).Cells("*").Value.ToString = "" Then
            frmMain.ShowAllPossibleMods(DataGridView1, TempInv(CInt(DataGridView1.Rows(e.RowIndex).Cells("Index").Value)), TempInv(CInt(DataGridView1.Rows(e.RowIndex).Cells("Index").Value)).ExplicitSuffixMods, "Suffix")
        ElseIf e.ColumnIndex = DataGridView1.Columns("Location").Index Then
            If blStore = True Then
                Dim strURL As String = "http://www.pathofexile.com/forum/view-thread/" & DataGridView1.Rows(e.RowIndex).Cells("ThreadID").Value
                Process.Start(strURL) ' Take the user to the GGG web page
            Else
                Dim TempList As New List(Of FullItem)
                If DataGridView1.CurrentRow.Cells("*").Value = "*" Then
                    TempList = FullInv
                Else
                    TempList = TempInv
                End If
                frmLocation.X = TempList(CInt(DataGridView1.Rows(e.RowIndex).Cells("Index").Value)).X
                frmLocation.Y = TempList(CInt(DataGridView1.Rows(e.RowIndex).Cells("Index").Value)).Y
                frmLocation.H = TempList(CInt(DataGridView1.Rows(e.RowIndex).Cells("Index").Value)).H
                frmLocation.W = TempList(CInt(DataGridView1.Rows(e.RowIndex).Cells("Index").Value)).W
                frmLocation.TabName = TempList(CInt(DataGridView1.Rows(e.RowIndex).Cells("Index").Value)).Location
                frmLocation.ItemName = TempList(CInt(DataGridView1.Rows(e.RowIndex).Cells("Index").Value)).Name
                If frmLocation.TabName.EndsWith(" Tab") = False Then
                    MessageBox.Show("This item is in a character's inventory and the location cannot be shown.", "Item in Character Inventory", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Else
                    frmLocation.ShowDialog(Me)
                End If
            End If
        ElseIf DataGridView1.Columns(e.ColumnIndex).Name.CompareMultiple(StringComparison.Ordinal, "Sokt", "Link") Then
            MessageBox.Show(DataGridView1.CurrentRow.Cells("SktClrs").Value, "Socket/Link Info", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    Private Sub DataGridView1_CellMouseEnter(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellMouseEnter
        If frmMain.IsValidCellAddress(DataGridView1, e.RowIndex, e.ColumnIndex) AndAlso _
            (DataGridView1.Columns(e.ColumnIndex).Name.Contains("fix") Or _
             DataGridView1.Columns(e.ColumnIndex).Name.CompareMultiple(StringComparison.Ordinal, "Sokt", "Link", "Rank", "Location")) _
         Then DataGridView1.Cursor = Cursors.Hand
    End Sub

    Private Sub DataGridView1_CellMouseLeave(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellMouseLeave
        If frmMain.IsValidCellAddress(DataGridView1, e.RowIndex, e.ColumnIndex) AndAlso _
            (DataGridView1.Columns(e.ColumnIndex).Name.Contains("fix") Or _
                DataGridView1.Columns(e.ColumnIndex).Name.CompareMultiple(StringComparison.Ordinal, "Sokt", "Link", "Rank", "Location")) _
            Then DataGridView1.Cursor = Cursors.Default
    End Sub

    Private Sub DataGridView1_CellPainting(sender As Object, e As DataGridViewCellPaintingEventArgs) Handles DataGridView1.CellPainting
        If e.RowIndex = -1 Then
            frmMain.DataGridViewCellPaintingHeaderFormat(sender, e)
            Exit Sub
        End If
        Dim strName As String = DataGridView1.Columns(e.ColumnIndex).Name.ToLower
        Dim intIndex As Integer = CInt(DataGridView1.Rows(e.RowIndex).Cells("Index").Value)
        If strName.Contains("prefix") AndAlso NotNull(DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value.ToString, "") <> "" AndAlso DataGridView1.Rows(e.RowIndex).Cells("*").Value.ToString = "*" Then
            frmMain.DataGridViewAddLevelBar(DataGridView1, FullInv(intIndex).ExplicitPrefixMods, strName, sender, e)
        ElseIf strName.Contains("prefix") AndAlso NotNull(DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value.ToString, "") <> "" AndAlso DataGridView1.Rows(e.RowIndex).Cells("*").Value.ToString = "" Then
            frmMain.DataGridViewAddLevelBar(DataGridView1, TempInv(intIndex).ExplicitPrefixMods, strName, sender, e)
        ElseIf strName.Contains("suffix") AndAlso NotNull(DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value.ToString, "") <> "" AndAlso DataGridView1.Rows(e.RowIndex).Cells("*").Value.ToString = "*" Then
            frmMain.DataGridViewAddLevelBar(DataGridView1, FullInv(intIndex).ExplicitSuffixMods, strName, sender, e)
        ElseIf strName.Contains("suffix") AndAlso NotNull(DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value.ToString, "") <> "" AndAlso DataGridView1.Rows(e.RowIndex).Cells("*").Value.ToString = "" Then
            frmMain.DataGridViewAddLevelBar(DataGridView1, TempInv(intIndex).ExplicitSuffixMods, strName, sender, e)
        End If
    End Sub

    Private Sub DataGridView1_RowPostPaint(sender As Object, e As DataGridViewRowPostPaintEventArgs) Handles DataGridView1.RowPostPaint
        If e.RowIndex = -1 Then Exit Sub
        If DataGridView1.Rows(e.RowIndex).Cells("*").Value.ToString = "*" Then
            frmMain.DataGridViewRowPostPaint(DataGridView1, FullInv, sender, e)
        Else
            frmMain.DataGridViewRowPostPaint(DataGridView1, TempInv, sender, e)
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
        If frmMain.TabControl1.SelectedTab.Name = "TabPage2" Then
            FullInv = FullStoreInventory
            TempInv = TempStoreInventory
        Else
            FullInv = FullInventory
            TempInv = TempInventory
        End If
    End Sub
End Class