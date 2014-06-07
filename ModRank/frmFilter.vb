Imports ModRank.Extenders.Strings.StringExtender

Public Class frmFilter
    Public intIndex As Integer = 0, intHeight As Integer = 25
    Public Delegate Sub ClickDelegate(ByVal sender As Object, ByVal e As EventArgs)
    Public Delegate Sub FillComboBox(ByVal strCombo As String, ByVal ds As DataSet)
    Public blFinishedLoading As Boolean = False
    Public blStore As Boolean
    Public intMaxWidth As Integer   ' The max width of the cmbText0 combo box on load, since we may have to change its width depending on user selection
    Public intLeft As Integer   ' The left position of the cmbText0 combo box on load
    Public lstDistinctPre As New CloneableList(Of String)
    Public lstDistinctSuf As New CloneableList(Of String)
    Public lstType As New CloneableList(Of String)
    Public lstAllSingleMods As New CloneableList(Of String)

    Private Sub frmFilter_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            blFinishedLoading = False
            Me.Icon = GetEmbeddedIcon("ModRank.PoE.ico")
            If cmbOperator0.Items.Count = 0 Then cmbOperator0.DataSource = New String() {"=", "<>", "LIKE", "<=", ">="} : cmbOperator0.SelectedIndex = 0
            If (cmbField0.Items.Count = 0 And cmbText0.Items.Count = 0) Then
                cmbAscDesc.SelectedIndex = 1        ' DESC is the more common choice
                cmbAndOr0.SelectedIndex = 0
                Dim blAddedSpecialFields As Boolean = False
                ' Loop through the table and place the field names into the cmbfield box
                For Each col As DataColumn In dtRank.Columns
                    If col.ColumnName.CompareMultiple(StringComparison.Ordinal, "Prefix 1", "Prefix 2", "Prefix 3", "Suffix 1", "Suffix 2", "Suffix 3", "Implicit") Or _
                        col.Ordinal >= bytDisplay Then
                        If blAddedSpecialFields = False Then
                            cmbField0.Items.Add("Prefix Type")
                            cmbField0.Items.Add("Number of Prefixes")
                            cmbField0.Items.Add("Suffix Type")
                            cmbField0.Items.Add("Number of Suffixes")
                            cmbField0.Items.Add("Implicit Type")
                            cmbField0.Items.Add("Number of Implicits")
                            cmbField0.Items.Add("Mod Total Value")
                            cmbField0.Items.Add("Total Number of Explicits")
                            If blStore = True Then cmbField0.Items.Add("Price") : cmbField0.Items.Add("ThreadID")
                            blAddedSpecialFields = True
                        End If
                        Continue For
                    ElseIf col.ColumnName = "Sokt" Then
                        cmbField0.Items.Add("Sokt")
                        cmbField0.Items.Add("Socket Colour and Links")
                        Continue For
                    End If
                    cmbField0.Items.Add(col.ColumnName)
                    cmbOrderField.Items.Add(col.ColumnName)
                Next
                cmbOrderField.Items.Add("pcount")
                cmbOrderField.Items.Add("scount")
                cmbOrderField.Items.Add("icount")
                cmbOrderField.Items.Add("ecount")
                For i = 1 To 6 : cmbOrderField.Items.Add("Tot" & i) : Next
                ' Used to populate cmbText0 with the entries from the weights/mods file
                For Each row As DataRow In dtWeights.Rows
                    Dim strName As String = row("ExportField").ToString & IIf(row("ExportField2").ToString <> "", "/" & row("ExportField2").ToString, "").ToString
                    If dtMods.Select("Description='" & row("Description").ToString & "'")(0)("Prefix/Suffix").ToString = "Prefix" Then
                        If lstDistinctPre.IndexOf(strName) = -1 Then lstDistinctPre.Add(strName)
                    Else
                        If lstDistinctSuf.IndexOf(strName) = -1 Then lstDistinctSuf.Add(strName)
                    End If
                    If row("Description").ToString = "Base Item Found Rarity +%" Then lstDistinctSuf.Add("% increased Rarity of Items found")
                    If strName.IndexOf("/") = -1 Then
                        If lstAllSingleMods.IndexOf(strName) = -1 Then lstAllSingleMods.Add(strName)
                    End If
                Next
                lstDistinctPre.Sort() : lstDistinctSuf.Sort() : lstAllSingleMods.Sort()
                ' Used to populate cmbText0 with the list of types
                Dim items As Array = System.Enum.GetValues(GetType(POEApi.Model.GearType))
                For Each intType As Integer In items
                    If items(intType).ToString.CompareMultiple(StringComparison.Ordinal, "Unknown", "Flask", "Map", "QuestItem") = False Then
                        If items(intType).ToString.CompareMultiple(StringComparison.Ordinal, "Sword", "Axe", "Mace") = True Then
                            lstType.Add(items(intType).ToString & " (1h)")
                            lstType.Add(items(intType).ToString & " (2h)")
                        Else
                            lstType.Add(items(intType).ToString)
                        End If
                    End If
                Next
                lstType.Sort()
                cmbField0.SelectedIndex = 0
                cmbOrderField.SelectedIndex = 0
                intMaxWidth = cmbText0.Width : intLeft = cmbText0.Left
            End If
            blFinishedLoading = True
        Catch ex As Exception
            ErrorHandler(System.Reflection.MethodBase.GetCurrentMethod.Name, ex)
        End Try
    End Sub

    Private Sub cmdCancel_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        For intCounter As Integer = intIndex To 0 Step -1
            Dim btnTemp As Button = CType(Me.Controls("cmdMinus" & intIndex), Button)
            btnTemp.PerformClick()
        Next
        If blStore = True Then
            txtOrderBy.Text = ""
            If strStoreFilter.Length <> 0 Then frmMain.SetStoreFilter("")
        Else
            txtOrderBy.Text = ""
            If strFilter.Length <> 0 Then frmMain.SetFilter("")
        End If
        Me.Close()
    End Sub

    Private Sub cmdActivate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdActivate.Click
        Try
            strRawFilter = CreateFilter()
            If strRawFilter <> "" Then
                If blStore = True Then
                    frmMain.SetStoreFilter(strRawFilter)
                Else
                    frmMain.SetFilter(strRawFilter)
                End If
                Me.Close()
            End If
        Catch ex As Exception
            ErrorHandler(System.Reflection.MethodBase.GetCurrentMethod.Name, ex)
        End Try
    End Sub

    Private Function CreateFilter() As String
        Dim intBrak As Integer = 0
        Dim sb As New System.Text.StringBuilder
        For intCounter As Integer = 0 To intIndex
            Dim strLeftBrak As String = Me.Controls("txtLeftBrak" & intCounter).Text.ToString
            Dim cmbFld As ComboBox = CType(Me.Controls("cmbField" & intCounter), ComboBox)
            Dim strOp As String = Me.Controls("cmbOperator" & intCounter).Text.ToString
            Dim txtText As TextBox = CType(Me.Controls("txtText" & intCounter), TextBox)
            Dim txtValue As TextBox = CType(Me.Controls("txtValue" & intCounter), TextBox)
            Dim cmbText As ComboBox = CType(Me.Controls("cmbText" & intCounter), ComboBox)
            Dim strRightBrak As String = Me.Controls("txtRightBrak" & intCounter).Text.ToString
            Dim strAndOr As String = IIf(intCounter <> intIndex, " " & Me.Controls("cmbAndOr" & intCounter).Text.ToString & " ", "").ToString

            If cmbFld.SelectedIndex <> -1 Or (cmbText.Visible And (cmbText.SelectedIndex <> -1 Or cmbText.Text.Length <> 0)) Then
                If strOp.Equals("LIKE") And ((cmbText.Visible = False And (txtText.Text.IndexOf("*") <> -1 Or txtText.Text.IndexOf("%") <> -1)) Or _
                                              (cmbText.Visible = True And (cmbText.Text.IndexOf("*") <> -1 Or cmbText.Text.IndexOf("%") <> -1))) Then
                    MessageBox.Show("Do not use LIKE for prefix/suffix type searches with a % type mod. Also, do not enter wildcards (*, %) for LIKE searches, they will be added automatically at the start and end of the text entered.", "LIKE Wildcard Syntax Error", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1)
                    Return ""
                End If
                Dim intTest As Integer, sngTest As Single
                Dim strTemp As String = "", strInner As String = ""
                If cmbFld.Text.CompareMultiple(StringComparison.Ordinal, "Prefix Type", "Suffix Type", "Number of Prefixes", "Number of Suffixes", "Implicit Type", _
                                               "Number of Implicits", "Mod Total Value", "Total Number of Explicits", "Socket Number/Colour", "Socket Colour and Links", "Price") = False Then
                    If dtRank.Columns(cmbFld.Text).DataType = System.Type.GetType("System.String") Then
                        strTemp = "'" & IIf(strOp.Equals("LIKE"), "%", "").ToString & IIf(cmbText.Visible, cmbText.Text, txtText.Text).ToString.Trim & IIf(strOp.Equals("LIKE"), "%", "").ToString & "'"
                    ElseIf cmbFld.Text.CompareMultiple(StringComparison.Ordinal, "Rank", "Level", "Sokt", "Link", "%") = True Then
                        If strOp.Equals("LIKE") Then
                            MessageBox.Show("Unable to apply/save the filter. LIKE comparison cannot be used with numeric values.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1)
                            Return ""
                        End If
                        If Integer.TryParse(txtText.Text.Trim, intTest) = False And Single.TryParse(txtText.Text.Trim, sngTest) = False Then  ' Test for numeric values
                            MessageBox.Show("Unable to apply/save the filter. You must enter a valid number for numeric field '" & cmbFld.Text & "'.", "Invalid Number", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1)
                            Return ""
                        End If
                        strTemp = IIf(cmbText.Visible, cmbText.Text, txtText.Text).ToString.Trim
                    Else
                        strTemp = IIf(cmbText.Visible, cmbText.Text, txtText.Text).ToString.Trim
                    End If
                ElseIf cmbFld.Text.CompareMultiple(StringComparison.Ordinal, "Prefix Type", "Suffix Type", "Implicit Type", "Socket Colour and Links") = True Then
                    If strOp.Equals("LIKE") And txtText.Text.Trim.Contains(" ") Then
                        MessageBox.Show("Unable to apply/save the filter. LIKE comparison cannot be used with multiple socket colour string searches. Try adding each colour string as its own condition with an AND clause.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1)
                        Return ""
                    End If
                    strTemp = "'" & IIf(strOp.Equals("LIKE"), "%", "").ToString & IIf(cmbText.Visible, cmbText.Text, txtText.Text).ToString.Trim & IIf(strOp.Equals("LIKE"), "%", "").ToString & "'"
                ElseIf cmbFld.Text.CompareMultiple(StringComparison.Ordinal, "Number of Prefixes", "Number of Suffixes", "Number of Implicits", "Total Number of Explicits") = True Then
                    If strOp.Equals("LIKE") Then
                        MessageBox.Show("Unable to apply/save the filter. LIKE comparison cannot be used with numeric values.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1)
                        Return ""
                    End If
                    If Integer.TryParse(txtText.Text.Trim, intTest) = False And Single.TryParse(txtText.Text.Trim, sngTest) = False Then
                        MessageBox.Show("Unable to apply/save the filter. You must enter a valid number for numeric field '" & cmbFld.Text & "'.", "Invalid Number", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1)
                        Return ""
                    End If
                    strTemp = IIf(cmbText.Visible, cmbText.Text, txtText.Text).ToString.Trim
                ElseIf cmbFld.Text = "Price" Then
                    If strOp.Equals("LIKE") Then
                        MessageBox.Show("Unable to apply/save the filter. LIKE comparison cannot be used with numeric values.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1)
                        Return ""
                    End If
                    If Integer.TryParse(txtValue.Text.Trim, intTest) = False And Single.TryParse(txtValue.Text.Trim, sngTest) = False And txtValue.Text.Contains("/") = False Then
                        MessageBox.Show("Unable to apply/save the filter. You must enter a valid number for the value field.", "Invalid Number", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1)
                        txtValue.SelectAll() : txtValue.Focus()
                        Return ""
                    End If
                    If cmbText.SelectedIndex = -1 Then
                        MessageBox.Show("Unable to apply/save the filter. You must select an orb type.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1)
                        cmbText.DroppedDown = True
                        Return ""
                    End If
                    strInner = "[" & cmbFld.Text & "]" & strD & strOp & strD & txtValue.Text.Trim & strD & "'" & cmbText.Text.Trim & "'"
                    sb.Append(strLeftBrak & strD & strInner & strD & strRightBrak & strD & strAndOr & vbCrLf)
                    Continue For
                ElseIf cmbFld.Text = "Mod Total Value" Then
                    If strOp.Equals("LIKE") Then
                        MessageBox.Show("Unable to apply/save the filter. LIKE comparison cannot be used with numeric values.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1)
                        Return ""
                    End If
                    If Integer.TryParse(txtValue.Text.Trim, intTest) = False And Single.TryParse(txtValue.Text.Trim, sngTest) = False And txtValue.Text.Contains("/") = False Then
                        MessageBox.Show("Unable to apply/save the filter. You must enter a valid number for the value field.", "Invalid Number", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1)
                        txtValue.SelectAll() : txtValue.Focus()
                        Return ""
                    End If
                    If cmbText.SelectedIndex = -1 Then
                        MessageBox.Show("Unable to apply/save the filter. You must select a mod type from which the value total can be calculated.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1)
                        cmbText.DroppedDown = True
                        Return ""
                    End If
                    strInner = "[" & cmbFld.Text & "]" & strD & strOp & strD & txtValue.Text.Trim & strD & "'" & cmbText.Text.Trim & "'"
                    sb.Append(strLeftBrak & strD & strInner & strD & strRightBrak & strD & strAndOr & vbCrLf)
                    Continue For
                End If
                strInner = "[" & cmbFld.Text & "]" & strD & strOp & strD & If(strOp = " LIKE ", "'%", "") & strTemp & If(strOp = " LIKE ", "%' ", "")
                intBrak += strLeftBrak.Count(Function(c As Char) c = "(")
                intBrak -= strRightBrak.Count(Function(c As Char) c = ")")
                sb.Append(strLeftBrak & strD & strInner & strD & strRightBrak & strD & strAndOr & vbCrLf)
            End If
        Next
        If intBrak <> 0 Then
            MessageBox.Show("Unable to apply/save the filter. The filter is invalid because it is missing " & IIf(intBrak < 0, Math.Abs(intBrak) & " opening", intBrak & " closing").ToString & " bracket(s). Please fix this and try activating again.", _
                            "Invalid Filter: Missing Brackets", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
            Return ""
        End If
        If sb.Length = 0 Then
            MessageBox.Show("Unable to apply/save the filter. You must create at least one valid filter before activating.", "No Valid Filters Found", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
            Return ""
        End If
        If txtOrderBy.Text.Trim.Length <> 0 Then
            sb.Append("ORDER BY" & strD & txtOrderBy.Text)
        Else
            sb.Append("ORDER BY" & strD & "[Rank] DESC")
        End If
        Return sb.ToString
    End Function

    Private Sub cmdPlus0_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdPlus0.Click
        Try
            Dim intCurrentIndex As Integer = GetNumeric(sender.name.ToString)
            Dim intTab = DirectCast(sender, Button).TabIndex
            If intCurrentIndex <> intIndex Then
                For intCounter As Integer = intIndex To intCurrentIndex + 1 Step -1
                    For Each strName In {"txtLeftBrak", "cmbField", "cmbOperator", "txtText", "txtValue", "cmbText", "txtRightBrak", "cmbAndOr", "cmdMinus", "cmdPlus"}
                        Me.Controls(strName & intCounter).Top += intHeight
                        Me.Controls(strName & intCounter).Name = strName & intCounter + 1
                    Next
                Next
            End If
            Me.Height = Me.Height + intHeight
            Dim txtLeftBrak As New TextBox, cmbField As New ComboBox, cmbOperator As New ComboBox, txtText As New TextBox, txtValue As New TextBox, cmbText As New ComboBox, txtRightBrak As New TextBox, cmbAndOr As New ComboBox, cmdMinus As New Button, cmdPlus As New Button
            Dim arr(,) As Control = {{txtLeftBrak, txtLeftBrak0}, {cmbField, cmbField0}, {cmbOperator, cmbOperator0}, {txtText, txtText0}, {txtValue, txtValue0}, {cmbText, cmbText0}, {txtRightBrak, txtRightBrak0}, {cmbAndOr, cmbAndOr0}, {cmdMinus, cmdMinus0}, {cmdPlus, cmdPlus0}}
            For i = 0 To UBound(arr, 1)
                arr(i, 0).Name = GetChars(arr(i, 1).Name) & intCurrentIndex + 1
                arr(i, 0).Width = CInt(IIf(arr(i, 0).Name.Contains("cmbText"), intMaxWidth, arr(i, 1).Width))
                arr(i, 0).Left = CInt(IIf(arr(i, 0).Name.Contains("cmbText"), intLeft, arr(i, 1).Left))
                arr(i, 0).Height = arr(i, 1).Height
                arr(i, 0).Top = Me.Controls(GetChars(arr(i, 1).Name) & intCurrentIndex).Top + intHeight
                arr(i, 0).TabIndex = 11 + (intIndex * 10) + i
                ToolTip1.SetToolTip(arr(i, 0), ToolTip1.GetToolTip(arr(i, 1)))
                If arr(i, 1).Name.Contains("cmb") Then
                    Dim cmbTemp As ComboBox = DirectCast(arr(i, 0), ComboBox) : Dim cmbTemp2 As ComboBox = DirectCast(arr(i, 1), ComboBox)
                    cmbTemp.DropDownStyle = cmbTemp2.DropDownStyle
                    cmbTemp.DropDownWidth = cmbTemp2.DropDownWidth
                    'cmbTemp.AutoCompleteMode = cmbTemp2.AutoCompleteMode
                    'cmbTemp.AutoCompleteSource = cmbTemp2.AutoCompleteSource
                    cmbTemp.MaxDropDownItems = 32
                End If
                If arr(i, 1).Name.Contains("cmbText") Then arr(i, 0).Visible = False
                If arr(i, 1).Name.Contains("txtValue") Then arr(i, 0).Visible = False
                If arr(i, 1).Name.Contains("cmdMinus") Then arr(i, 0).Text = "-"
                If arr(i, 1).Name.Contains("cmdPlus") Then arr(i, 0).Text = "+"
            Next
            cmdMinus.TextAlign = ContentAlignment.MiddleRight
            'cmdPlus.Top += intHeight
            cmbOrderField.Top += intHeight : cmbAscDesc.Top += intHeight : btnAddOrdering.Top += intHeight
            lblOrderBy.Top += intHeight : txtOrderBy.Top += intHeight
            cmdActivate.Top += intHeight : cmdCancel.Top += intHeight
            ' We'll be adding new items in, so increment the tabindex of the items "below" them on the form
            For Each ctl As Control In {cmbOrderField, cmbAscDesc, btnAddOrdering, lblOrderBy, txtOrderBy, cmdActivate, cmdCancel}
                ctl.TabIndex += 11
            Next
            Me.Controls.AddRange(New Control() {txtLeftBrak, cmbField, cmbOperator, txtText, txtValue, cmbText, txtRightBrak, cmbAndOr, cmdMinus, cmdPlus})
            Dim obj(cmbField0.Items.Count - 1) As Object
            cmbField0.Items.CopyTo(obj, 0)
            cmbField.Items.AddRange(obj) : cmbField.SelectedIndex = 0
            Dim obj2(cmbOperator0.Items.Count - 1) As Object
            cmbOperator0.Items.CopyTo(obj2, 0)
            cmbOperator.Items.AddRange(obj2) : cmbOperator.SelectedIndex = 0
            Dim obj3(cmbText0.Items.Count - 1) As Object
            cmbText0.Items.CopyTo(obj3, 0)
            cmbText.Items.AddRange(obj3) : cmbText.SelectedIndex = -1
            Dim obj4(cmbAndOr0.Items.Count - 1) As Object
            cmbAndOr0.Items.CopyTo(obj4, 0)
            cmbAndOr.Items.AddRange(obj4) : cmbAndOr.SelectedIndex = 0
            AddHandler txtLeftBrak.TextChanged, AddressOf OnlyAllowBrackets
            AddHandler cmbField.SelectedIndexChanged, AddressOf cmbFieldSelectedIndexChanged
            AddHandler txtRightBrak.TextChanged, AddressOf OnlyAllowBrackets
            AddHandler cmdMinus.Click, AddressOf MinusButtonClick
            AddHandler cmdPlus.Click, AddressOf cmdPlus0_Click
            intIndex += 1
        Catch ex As Exception
            ErrorHandler(System.Reflection.MethodBase.GetCurrentMethod.Name, ex)
        End Try
    End Sub

    Private Sub cmdMinus0_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdMinus0.Click
        txtLeftBrak0.Text = "" : cmbField0.SelectedIndex = -1 : cmbOperator0.SelectedIndex = 0 : txtText0.Text = ""
        cmbText0.SelectedIndex = -1 : txtText0.Visible = True : txtRightBrak0.Text = "" : cmbText0.Visible = False
        txtValue0.Text = "" : txtValue0.Visible = False
    End Sub

    Private Sub MinusButtonClick(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim intCurrentIndex As Integer = GetNumeric(sender.name.ToString)
        If intCurrentIndex = 0 Then cmdMinus0_Click(cmdMinus0, EventArgs.Empty) : Exit Sub
        'cmdPlus0.Top -= intHeight
        cmbOrderField.Top -= intHeight : cmbAscDesc.Top -= intHeight : btnAddOrdering.Top -= intHeight
        lblOrderBy.Top -= intHeight : txtOrderBy.Top -= intHeight
        cmdActivate.Top -= intHeight : cmdCancel.Top -= intHeight
        Dim txtLeftBrak As TextBox = CType(Me.Controls("txtLeftBrak" & intCurrentIndex), TextBox)
        Dim cmbField As ComboBox = CType(Me.Controls("cmbField" & intCurrentIndex), ComboBox)
        Dim cmbText As ComboBox = CType(Me.Controls("cmbText" & intCurrentIndex), ComboBox)
        Dim txtRightBrak As TextBox = CType(Me.Controls("txtRightBrak" & intCurrentIndex), TextBox)
        Dim cmbAndOr As ComboBox = CType(Me.Controls("cmbAndOr" & intCurrentIndex), ComboBox)
        RemoveHandler cmbField.SelectedIndexChanged, AddressOf cmbFieldSelectedIndexChanged
        RemoveHandler Me.Controls("cmdMinus" & intCurrentIndex).Click, AddressOf MinusButtonClick
        RemoveHandler txtLeftBrak.TextChanged, AddressOf OnlyAllowBrackets
        RemoveHandler txtRightBrak.TextChanged, AddressOf OnlyAllowBrackets
        Me.Controls.Remove(txtLeftBrak)
        Me.Controls.Remove(cmbField)
        Me.Controls.Remove(Me.Controls("cmbOperator" & intCurrentIndex))
        Me.Controls.Remove(Me.Controls("txtText" & intCurrentIndex))
        Me.Controls.Remove(Me.Controls("txtValue" & intCurrentIndex))
        Me.Controls.Remove(cmbText)
        Me.Controls.Remove(txtRightBrak)
        Me.Controls.Remove(cmbAndOr)
        Me.Controls.Remove(Me.Controls("cmdMinus" & intCurrentIndex))
        Me.Controls.Remove(Me.Controls("cmdPlus" & intCurrentIndex))
        For intCounter As Integer = intIndex To intCurrentIndex + 1 Step -1
            For Each strName In {"txtLeftBrak", "cmbField", "cmbOperator", "txtText", "txtValue", "cmbText", "txtRightBrak", "cmbAndOr", "cmdMinus", "cmdPlus"}
                Me.Controls(strName & intCounter).Top -= intHeight
                Me.Controls(strName & intCounter).Name = strName & intCounter - 1
            Next
        Next
        For Each ctl As Control In {cmbOrderField, cmbAscDesc, btnAddOrdering, lblOrderBy, txtOrderBy, cmdActivate, cmdCancel}
            ctl.TabIndex -= 11
        Next
        intIndex -= 1
        Me.Height -= intHeight
    End Sub

    Private Sub cmbField0_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbField0.SelectedIndexChanged
        cmbFieldSelectedIndexChanged(sender, e)
    End Sub

    Private Sub cmbFieldSelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim cmbFld As ComboBox = DirectCast(sender, ComboBox)
        Dim intCurrentIndex As Integer = GetNumeric(sender.name.ToString)
        If cmbFld.SelectedIndex = -1 Or cmbFld.Text.Length = 0 Then Exit Sub
        Me.Cursor = Cursors.WaitCursor
        Application.DoEvents()
        Dim cmbText As ComboBox = CType(Me.Controls("cmbText" & intCurrentIndex), ComboBox)
        Me.Controls("txtValue" & intCurrentIndex).Visible = False
        If cmbFld.Text = "Prefix Type" Then
            cmbText.Left = intLeft : cmbText.Width = intMaxWidth
            cmbText.Visible = True : Me.Controls("txtText" & intCurrentIndex).Visible = False
            cmbText.DataSource = lstDistinctPre.Clone
            If blFinishedLoading Then cmbText.DroppedDown = True : cmbText.Focus()
        ElseIf cmbFld.Text = "Suffix Type" Then
            cmbText.Left = intLeft : cmbText.Width = intMaxWidth
            cmbText.Visible = True : Me.Controls("txtText" & intCurrentIndex).Visible = False
            cmbText.DataSource = lstDistinctSuf.Clone
            If blFinishedLoading Then cmbText.DroppedDown = True : cmbText.Focus()
        ElseIf cmbFld.Text = "Type" Then
            cmbText.Left = intLeft : cmbText.Width = intMaxWidth
            cmbText.Visible = True : Me.Controls("txtText" & intCurrentIndex).Visible = False
            cmbText.DataSource = lstType.Clone
            If blFinishedLoading Then cmbText.DroppedDown = True : cmbText.Focus()
        ElseIf cmbFld.Text.CompareMultiple(StringComparison.Ordinal, "Gem", "Crpt") = True Then
            cmbText.Left = intLeft : cmbText.Width = intMaxWidth
            cmbText.Visible = True : Me.Controls("txtText" & intCurrentIndex).Visible = False
            cmbText.DataSource = New String() {"True", "False"}
            If blFinishedLoading Then cmbText.DroppedDown = True : cmbText.Focus()
        ElseIf cmbFld.Text = "Mod Total Value" Then
            cmbText.Left = intLeft + 44 : cmbText.Width = intMaxWidth - 44
            cmbText.Visible = True : Me.Controls("txtText" & intCurrentIndex).Visible = False
            cmbText.DataSource = lstAllSingleMods.Clone : cmbText.SelectedIndex = 0
            Dim cmbOperator As ComboBox = CType(Me.Controls("cmbOperator" & intCurrentIndex), ComboBox)
            cmbOperator.SelectedIndex = cmbOperator.FindStringExact(">=")     ' >= is the most common comparison for mod total value searches
            Me.Controls("txtValue" & intCurrentIndex).Visible = True : Me.Controls("txtValue" & intCurrentIndex).Focus()
        ElseIf cmbFld.Text = "Price" Then
            cmbText.Left = intLeft + 44 : cmbText.Width = intMaxWidth - 44
            cmbText.Visible = True : Me.Controls("txtText" & intCurrentIndex).Visible = False
            cmbText.DataSource = New String() {"Exa", "Chaos", "Alch", "GCP"} : cmbText.SelectedIndex = 0
            Dim cmbOperator As ComboBox = CType(Me.Controls("cmbOperator" & intCurrentIndex), ComboBox)
            cmbOperator.SelectedIndex = cmbOperator.FindStringExact("<=")     ' >= is the most common comparison for price searches
            Me.Controls("txtValue" & intCurrentIndex).Visible = True : Me.Controls("txtValue" & intCurrentIndex).Focus()
        Else
            If cmbText.Visible = True Then cmbText.Visible = False : cmbText.DroppedDown = False
            If Me.Controls("txtValue" & intCurrentIndex).Visible = True Then Me.Controls("txtValue" & intCurrentIndex).Visible = False
            Me.Controls("txtText" & intCurrentIndex).Visible = True : Me.Controls("txtText" & intCurrentIndex).Focus()
        End If
        Me.Cursor = Cursors.Default
    End Sub

    Private Sub txtLeftBrak0_TextChanged(sender As Object, e As EventArgs) Handles txtLeftBrak0.TextChanged
        OnlyAllowBrackets(sender, e)
    End Sub

    Private Sub txtRightBrak0_TextChanged(sender As Object, e As EventArgs) Handles txtRightBrak0.TextChanged
        OnlyAllowBrackets(sender, e)
    End Sub

    Public Sub OnlyAllowBrackets(sender As Object, e As EventArgs)
        Dim txtBox As TextBox = DirectCast(sender, TextBox)
        Dim blLeftBracket As Boolean
        If txtBox.Name.ToLower.Contains("txtleft") Then blLeftBracket = True Else blLeftBracket = False
        Dim theText As String = txtBox.Text
        Dim Letter As String
        Dim SelectionIndex As Integer = txtBox.SelectionStart
        Dim Change As Integer

        For x As Integer = 0 To txtBox.Text.Length - 1
            Letter = txtBox.Text.Substring(x, 1)
            If Letter <> IIf(blLeftBracket, "(", ")").ToString Then
                theText = theText.Replace(Letter, String.Empty)
                Change = 1
            End If
        Next

        txtBox.Text = theText
        txtBox.Select(SelectionIndex - Change, 0)
    End Sub

    Private Sub btnAddOrdering_Click(sender As Object, e As EventArgs) Handles btnAddOrdering.Click
        If txtOrderBy.Text.Trim.Length = 0 Then
            txtOrderBy.Text = "[" & cmbOrderField.Text & "] " & cmbAscDesc.Text
        Else
            txtOrderBy.Text = txtOrderBy.Text & ", [" & cmbOrderField.Text & "] " & cmbAscDesc.Text
        End If
    End Sub

    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Try
            Dim sfd As New SaveFileDialog()
            If System.IO.Directory.Exists(Application.StartupPath & "\Filters") = False Then System.IO.Directory.CreateDirectory(Application.StartupPath & "\Filters")
            If blStore = True Then
                sfd.InitialDirectory = Application.StartupPath & "\Store\Filters"
            Else
                sfd.InitialDirectory = Application.StartupPath & "\Filters"
            End If
            sfd.CreatePrompt = False
            sfd.OverwritePrompt = True
            sfd.FileName = "myfilter"
            sfd.Filter = "Text File|*.txt"
            sfd.Title = "Save Search/Filter"
            If sfd.ShowDialog(Me) = Windows.Forms.DialogResult.Cancel Then Exit Sub

            If sfd.FileName = "" Then Exit Sub

            Dim strTemp As String = CreateFilter()
            If strTemp <> "" Then System.IO.File.WriteAllText(sfd.FileName, strTemp)
        Catch ex As Exception
            ErrorHandler(System.Reflection.MethodBase.GetCurrentMethod.Name, ex)
        End Try
    End Sub

    Private Sub btnLoad_Click(sender As Object, e As EventArgs) Handles btnLoad.Click
        Try
            Dim ofd As New OpenFileDialog
            If System.IO.Directory.Exists(Application.StartupPath & "\Filters") = False Then System.IO.Directory.CreateDirectory(Application.StartupPath & "\Filters")
            If blSTore = True Then
                ofd.InitialDirectory = Application.StartupPath & "\Store\Filters"
            Else
                ofd.InitialDirectory = Application.StartupPath & "\Filters"
            End If
            ofd.Filter = "Text File|*.txt"
            ofd.Title = "Open Search/Filter"
            Dim res As DialogResult = ofd.ShowDialog(Me)
            If res = Windows.Forms.DialogResult.Cancel Then Exit Sub

            For intCounter As Integer = 0 To intIndex
                Dim btnTemp As Button = CType(Me.Controls("cmdMinus" & intIndex), Button)
                btnTemp.PerformClick()
            Next
            txtOrderBy.Text = ""

            Dim intTemp As Integer = 0
            Dim strFilter As String = System.IO.File.ReadAllText(ofd.FileName)
            Dim strLines() As String = strFilter.Split(CChar(vbCrLf))
            blFinishedLoading = False
            For Each strLine In strLines
                Dim strElement() As String = strLine.Split(CChar(strD))
                If intTemp <> 0 And strElement(0).Trim <> "ORDER BY" Then cmdPlus0_Click(Me.Controls("cmdPlus" & intIndex), EventArgs.Empty)
                If strElement(0).Trim = "ORDER BY" Then
                    For Each strE In strElement
                        If strE.Trim = "ORDER BY" Then Continue For
                        txtOrderBy.Text = IIf(txtOrderBy.Text = "", txtOrderBy.Text, txtOrderBy.Text & " ").ToString & strE.Trim
                    Next
                    Continue For
                End If
                Me.Controls("txtLeftBrak" & intIndex).Text = strElement(0).Trim
                Me.Controls("cmbField" & intIndex).Text = strElement(1).Substring(1, strElement(1).Length - 2)
                Me.Controls("cmbOperator" & intIndex).Text = strElement(2)
                Dim cmbText As ComboBox = CType(Me.Controls("cmbText" & intIndex), ComboBox)
                If strElement(1).CompareMultiple(StringComparison.Ordinal, "[Prefix Type]", "[Suffix Type]", "[Type]", "[Gem]", "[Crpt]") Then
                    Dim intFound As Integer = 0
                    If strElement(2) = "LIKE" Then
                        intFound = cmbText.FindStringExact(strElement(3).Substring(2, strElement(3).Length - 4))
                        If intFound = -1 Then
                            cmbText.SelectedIndex = -1 : cmbText.Text = strElement(3).Substring(2, strElement(3).Length - 4)
                        Else
                            cmbText.SelectedIndex = intFound
                        End If
                    Else
                        If strElement(1) = "[Gem]" Or strElement(1) = "[Crpt]" Then
                            cmbText.Text = strElement(3)
                        Else
                            intFound = cmbText.FindStringExact(strElement(3).Substring(1, strElement(3).Length - 2))
                            If intFound = -1 Then
                                cmbText.SelectedIndex = -1 : cmbText.Text = strElement(3).Substring(1, strElement(3).Length - 2)
                            Else
                                cmbText.SelectedIndex = intFound
                            End If
                        End If
                    End If
                ElseIf strElement(1) = "[Mod Total Value]" Or strElement(1) = "[Price]" Then
                    Me.Controls("txtValue" & intIndex).Text = strElement(3)
                    cmbText.SelectedIndex = cmbText.FindStringExact(strElement(4).Substring(1, strElement(4).Length - 2))
                    Me.Controls("txtRightBrak" & intIndex).Text = strElement(5).Trim
                    Me.Controls("cmbAndOr" & intIndex).Text = strElement(6).Trim
                    intTemp += 1
                    Continue For
                ElseIf strElement(1).CompareMultiple(StringComparison.Ordinal, "[Number of Prefixes]", "[Number of Suffixes]", "[Number of Implicits]", "[Total Number of Explicits]") Then
                    Me.Controls("txtText" & intIndex).Text = strElement(3)
                ElseIf strElement(1).CompareMultiple(StringComparison.Ordinal, "[Implicit Type]", "[Socket Count and Links]") Then
                    Me.Controls("txtText" & intIndex).Text = strElement(3).Substring(1, strElement(3).Length - 2)
                ElseIf dtRank.Columns.Contains(strElement(1).Substring(1, strElement(1).Length - 2)) = True AndAlso _
                    dtRank.Columns(strElement(1).Substring(1, strElement(1).Length - 2)).DataType = System.Type.GetType("System.String") Then
                    If strElement(2) = "LIKE" Then
                        Me.Controls("txtText" & intIndex).Text = strElement(3).Substring(2, strElement(3).Length - 4)
                    Else
                        Me.Controls("txtText" & intIndex).Text = strElement(3).Substring(1, strElement(3).Length - 2)
                    End If
                ElseIf dtRank.Columns.Contains(strElement(1).Substring(1, strElement(1).Length - 2)) = False Then
                    MessageBox.Show("Warning: Could not find column named " & strElement(1).Substring(1, strElement(1).Length - 2) & " in dtRank datatable. Please edit the filter manually to fix the column name.", "Column Not Found", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Continue For
                Else
                    Me.Controls("txtText" & intIndex).Text = strElement(3)
                End If
                Me.Controls("txtRightBrak" & intIndex).Text = strElement(4).Trim
                Me.Controls("cmbAndOr" & intIndex).Text = strElement(5).Trim
                intTemp += 1
            Next
            ' Position the form vertically in the center of the screen...use max so that we don't go negative and lose the top control box
            Me.Top = CInt(Math.Max((Screen.PrimaryScreen.Bounds.Height - Me.Height) / 2, 0))
            blFinishedLoading = True

        Catch ex As Exception
            While CheckIfControlExists(Me, "cmdMinus" & intIndex) = False And intIndex <> 0
                intIndex -= 1
            End While
            ErrorHandler(System.Reflection.MethodBase.GetCurrentMethod.Name, ex)
        End Try
    End Sub
End Class