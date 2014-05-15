<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmFilter
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdActivate = New System.Windows.Forms.Button()
        Me.cmbField0 = New System.Windows.Forms.ComboBox()
        Me.cmbOperator0 = New System.Windows.Forms.ComboBox()
        Me.txtText0 = New System.Windows.Forms.TextBox()
        Me.cmdMinus0 = New System.Windows.Forms.Button()
        Me.cmdPlus0 = New System.Windows.Forms.Button()
        Me.cmbText0 = New System.Windows.Forms.ComboBox()
        Me.txtLeftBrak0 = New System.Windows.Forms.TextBox()
        Me.txtRightBrak0 = New System.Windows.Forms.TextBox()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.btnLoad = New System.Windows.Forms.Button()
        Me.cmbOrderField = New System.Windows.Forms.ComboBox()
        Me.btnAddOrdering = New System.Windows.Forms.Button()
        Me.cmbAscDesc = New System.Windows.Forms.ComboBox()
        Me.txtOrderBy = New System.Windows.Forms.TextBox()
        Me.lblOrderBy = New System.Windows.Forms.Label()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.cmbAndOr0 = New System.Windows.Forms.ComboBox()
        Me.txtValue0 = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'cmdCancel
        '
        Me.cmdCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.cmdCancel.Location = New System.Drawing.Point(263, 128)
        Me.cmdCancel.Margin = New System.Windows.Forms.Padding(2)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(97, 21)
        Me.cmdCancel.TabIndex = 4
        Me.cmdCancel.Text = "&Deactivate/Exit"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'cmdActivate
        '
        Me.cmdActivate.Location = New System.Drawing.Point(151, 128)
        Me.cmdActivate.Margin = New System.Windows.Forms.Padding(2)
        Me.cmdActivate.Name = "cmdActivate"
        Me.cmdActivate.Size = New System.Drawing.Size(97, 21)
        Me.cmdActivate.TabIndex = 5
        Me.cmdActivate.Text = "&Activate"
        Me.cmdActivate.UseVisualStyleBackColor = True
        '
        'cmbField0
        '
        Me.cmbField0.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest
        Me.cmbField0.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.cmbField0.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbField0.DropDownWidth = 150
        Me.cmbField0.FormattingEnabled = True
        Me.cmbField0.Location = New System.Drawing.Point(33, 43)
        Me.cmbField0.Margin = New System.Windows.Forms.Padding(2)
        Me.cmbField0.MaxDropDownItems = 32
        Me.cmbField0.Name = "cmbField0"
        Me.cmbField0.Size = New System.Drawing.Size(107, 21)
        Me.cmbField0.TabIndex = 6
        '
        'cmbOperator0
        '
        Me.cmbOperator0.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest
        Me.cmbOperator0.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.cmbOperator0.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbOperator0.FormattingEnabled = True
        Me.cmbOperator0.Location = New System.Drawing.Point(144, 43)
        Me.cmbOperator0.Margin = New System.Windows.Forms.Padding(2)
        Me.cmbOperator0.MaxDropDownItems = 32
        Me.cmbOperator0.Name = "cmbOperator0"
        Me.cmbOperator0.Size = New System.Drawing.Size(46, 21)
        Me.cmbOperator0.TabIndex = 7
        '
        'txtText0
        '
        Me.txtText0.Location = New System.Drawing.Point(194, 43)
        Me.txtText0.Margin = New System.Windows.Forms.Padding(2)
        Me.txtText0.Name = "txtText0"
        Me.txtText0.Size = New System.Drawing.Size(197, 22)
        Me.txtText0.TabIndex = 8
        '
        'cmdMinus0
        '
        Me.cmdMinus0.Location = New System.Drawing.Point(467, 43)
        Me.cmdMinus0.Margin = New System.Windows.Forms.Padding(2)
        Me.cmdMinus0.Name = "cmdMinus0"
        Me.cmdMinus0.Size = New System.Drawing.Size(16, 22)
        Me.cmdMinus0.TabIndex = 9
        Me.cmdMinus0.Text = "-"
        Me.cmdMinus0.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.cmdMinus0.UseVisualStyleBackColor = True
        '
        'cmdPlus0
        '
        Me.cmdPlus0.Location = New System.Drawing.Point(485, 43)
        Me.cmdPlus0.Margin = New System.Windows.Forms.Padding(2)
        Me.cmdPlus0.Name = "cmdPlus0"
        Me.cmdPlus0.Size = New System.Drawing.Size(16, 22)
        Me.cmdPlus0.TabIndex = 10
        Me.cmdPlus0.Text = "+"
        Me.cmdPlus0.UseVisualStyleBackColor = True
        '
        'cmbText0
        '
        Me.cmbText0.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest
        Me.cmbText0.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.cmbText0.DropDownWidth = 400
        Me.cmbText0.FormattingEnabled = True
        Me.cmbText0.Location = New System.Drawing.Point(194, 43)
        Me.cmbText0.Margin = New System.Windows.Forms.Padding(2)
        Me.cmbText0.MaxDropDownItems = 32
        Me.cmbText0.Name = "cmbText0"
        Me.cmbText0.Size = New System.Drawing.Size(197, 21)
        Me.cmbText0.TabIndex = 11
        Me.cmbText0.Visible = False
        '
        'txtLeftBrak0
        '
        Me.txtLeftBrak0.Location = New System.Drawing.Point(11, 43)
        Me.txtLeftBrak0.Margin = New System.Windows.Forms.Padding(2)
        Me.txtLeftBrak0.Name = "txtLeftBrak0"
        Me.txtLeftBrak0.Size = New System.Drawing.Size(18, 22)
        Me.txtLeftBrak0.TabIndex = 12
        Me.ToolTip1.SetToolTip(Me.txtLeftBrak0, "Place opening brackets here if order of precendence is required.")
        '
        'txtRightBrak0
        '
        Me.txtRightBrak0.Location = New System.Drawing.Point(394, 43)
        Me.txtRightBrak0.Margin = New System.Windows.Forms.Padding(2)
        Me.txtRightBrak0.Name = "txtRightBrak0"
        Me.txtRightBrak0.Size = New System.Drawing.Size(18, 22)
        Me.txtRightBrak0.TabIndex = 13
        Me.ToolTip1.SetToolTip(Me.txtRightBrak0, "Place closing brackets here if order of precendence is required.")
        '
        'btnSave
        '
        Me.btnSave.Location = New System.Drawing.Point(123, 13)
        Me.btnSave.Margin = New System.Windows.Forms.Padding(2)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(125, 21)
        Me.btnSave.TabIndex = 14
        Me.btnSave.Text = "&Save Search/Filter"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'btnLoad
        '
        Me.btnLoad.Location = New System.Drawing.Point(263, 13)
        Me.btnLoad.Margin = New System.Windows.Forms.Padding(2)
        Me.btnLoad.Name = "btnLoad"
        Me.btnLoad.Size = New System.Drawing.Size(125, 21)
        Me.btnLoad.TabIndex = 15
        Me.btnLoad.Text = "&Load Search/Filter"
        Me.btnLoad.UseVisualStyleBackColor = True
        '
        'cmbOrderField
        '
        Me.cmbOrderField.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest
        Me.cmbOrderField.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.cmbOrderField.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbOrderField.FormattingEnabled = True
        Me.cmbOrderField.Location = New System.Drawing.Point(15, 78)
        Me.cmbOrderField.Margin = New System.Windows.Forms.Padding(2)
        Me.cmbOrderField.MaxDropDownItems = 32
        Me.cmbOrderField.Name = "cmbOrderField"
        Me.cmbOrderField.Size = New System.Drawing.Size(147, 21)
        Me.cmbOrderField.TabIndex = 16
        '
        'btnAddOrdering
        '
        Me.btnAddOrdering.Location = New System.Drawing.Point(254, 77)
        Me.btnAddOrdering.Margin = New System.Windows.Forms.Padding(2)
        Me.btnAddOrdering.Name = "btnAddOrdering"
        Me.btnAddOrdering.Size = New System.Drawing.Size(97, 21)
        Me.btnAddOrdering.TabIndex = 17
        Me.btnAddOrdering.Text = "Add &Ordering"
        Me.btnAddOrdering.UseVisualStyleBackColor = True
        '
        'cmbAscDesc
        '
        Me.cmbAscDesc.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest
        Me.cmbAscDesc.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.cmbAscDesc.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbAscDesc.FormattingEnabled = True
        Me.cmbAscDesc.Items.AddRange(New Object() {"ASC", "DESC"})
        Me.cmbAscDesc.Location = New System.Drawing.Point(166, 78)
        Me.cmbAscDesc.Margin = New System.Windows.Forms.Padding(2)
        Me.cmbAscDesc.MaxDropDownItems = 32
        Me.cmbAscDesc.Name = "cmbAscDesc"
        Me.cmbAscDesc.Size = New System.Drawing.Size(68, 21)
        Me.cmbAscDesc.TabIndex = 18
        '
        'txtOrderBy
        '
        Me.txtOrderBy.Location = New System.Drawing.Point(72, 102)
        Me.txtOrderBy.Margin = New System.Windows.Forms.Padding(2)
        Me.txtOrderBy.Name = "txtOrderBy"
        Me.txtOrderBy.Size = New System.Drawing.Size(429, 22)
        Me.txtOrderBy.TabIndex = 19
        '
        'lblOrderBy
        '
        Me.lblOrderBy.AutoSize = True
        Me.lblOrderBy.Location = New System.Drawing.Point(12, 108)
        Me.lblOrderBy.Name = "lblOrderBy"
        Me.lblOrderBy.Size = New System.Drawing.Size(55, 13)
        Me.lblOrderBy.TabIndex = 20
        Me.lblOrderBy.Text = "Order By:"
        '
        'ToolTip1
        '
        Me.ToolTip1.AutoPopDelay = 5000
        Me.ToolTip1.InitialDelay = 200
        Me.ToolTip1.ReshowDelay = 100
        '
        'cmbAndOr0
        '
        Me.cmbAndOr0.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest
        Me.cmbAndOr0.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.cmbAndOr0.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbAndOr0.FormattingEnabled = True
        Me.cmbAndOr0.Items.AddRange(New Object() {"AND", "OR"})
        Me.cmbAndOr0.Location = New System.Drawing.Point(416, 43)
        Me.cmbAndOr0.Margin = New System.Windows.Forms.Padding(2)
        Me.cmbAndOr0.MaxDropDownItems = 32
        Me.cmbAndOr0.Name = "cmbAndOr0"
        Me.cmbAndOr0.Size = New System.Drawing.Size(47, 21)
        Me.cmbAndOr0.TabIndex = 21
        '
        'txtValue0
        '
        Me.txtValue0.Location = New System.Drawing.Point(194, 43)
        Me.txtValue0.Margin = New System.Windows.Forms.Padding(2)
        Me.txtValue0.Name = "txtValue0"
        Me.txtValue0.Size = New System.Drawing.Size(40, 22)
        Me.txtValue0.TabIndex = 22
        Me.txtValue0.Visible = False
        '
        'frmFilter
        '
        Me.AcceptButton = Me.cmdActivate
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.cmdCancel
        Me.ClientSize = New System.Drawing.Size(511, 160)
        Me.Controls.Add(Me.txtValue0)
        Me.Controls.Add(Me.cmbAndOr0)
        Me.Controls.Add(Me.lblOrderBy)
        Me.Controls.Add(Me.txtOrderBy)
        Me.Controls.Add(Me.cmbAscDesc)
        Me.Controls.Add(Me.btnAddOrdering)
        Me.Controls.Add(Me.cmbOrderField)
        Me.Controls.Add(Me.btnLoad)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.txtRightBrak0)
        Me.Controls.Add(Me.txtLeftBrak0)
        Me.Controls.Add(Me.cmdPlus0)
        Me.Controls.Add(Me.cmdMinus0)
        Me.Controls.Add(Me.txtText0)
        Me.Controls.Add(Me.cmbOperator0)
        Me.Controls.Add(Me.cmbField0)
        Me.Controls.Add(Me.cmdActivate)
        Me.Controls.Add(Me.cmdCancel)
        Me.Controls.Add(Me.cmbText0)
        Me.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmFilter"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Activate Search/Filter"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents cmdActivate As System.Windows.Forms.Button
    Friend WithEvents cmbField0 As System.Windows.Forms.ComboBox
    Friend WithEvents cmbOperator0 As System.Windows.Forms.ComboBox
    Friend WithEvents txtText0 As System.Windows.Forms.TextBox
    Friend WithEvents cmdMinus0 As System.Windows.Forms.Button
    Friend WithEvents cmdPlus0 As System.Windows.Forms.Button
    Friend WithEvents cmbText0 As System.Windows.Forms.ComboBox
    Friend WithEvents txtLeftBrak0 As System.Windows.Forms.TextBox
    Friend WithEvents txtRightBrak0 As System.Windows.Forms.TextBox
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents btnLoad As System.Windows.Forms.Button
    Friend WithEvents cmbOrderField As System.Windows.Forms.ComboBox
    Friend WithEvents btnAddOrdering As System.Windows.Forms.Button
    Friend WithEvents cmbAscDesc As System.Windows.Forms.ComboBox
    Friend WithEvents txtOrderBy As System.Windows.Forms.TextBox
    Friend WithEvents lblOrderBy As System.Windows.Forms.Label
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents cmbAndOr0 As System.Windows.Forms.ComboBox
    Friend WithEvents txtValue0 As System.Windows.Forms.TextBox
End Class
