<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
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
        Dim DataGridViewCellStyle7 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle8 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle9 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.btnLoad = New System.Windows.Forms.Button()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.txtEmail = New System.Windows.Forms.TextBox()
        Me.lblEmail = New System.Windows.Forms.Label()
        Me.lblPassword = New System.Windows.Forms.Label()
        Me.ElementHost1 = New System.Windows.Forms.Integration.ElementHost()
        Me.btnOffline = New System.Windows.Forms.Button()
        Me.ElementHost2 = New System.Windows.Forms.Integration.ElementHost()
        Me.cmbWeight = New System.Windows.Forms.ComboBox()
        Me.lblWeights = New System.Windows.Forms.Label()
        Me.chkSession = New System.Windows.Forms.CheckBox()
        Me.gpLegend = New System.Windows.Forms.GroupBox()
        Me.lblDoubleClick = New System.Windows.Forms.Label()
        Me.lblFixedValue = New System.Windows.Forms.Label()
        Me.pic3 = New System.Windows.Forms.PictureBox()
        Me.lblHighWeightMod = New System.Windows.Forms.Label()
        Me.lblLowWeightMod = New System.Windows.Forms.Label()
        Me.lblMax = New System.Windows.Forms.Label()
        Me.lblMin = New System.Windows.Forms.Label()
        Me.lblOtherSolutions = New System.Windows.Forms.Label()
        Me.lblUnknownValues = New System.Windows.Forms.Label()
        Me.lblHasGem = New System.Windows.Forms.Label()
        Me.pic2 = New System.Windows.Forms.PictureBox()
        Me.pic1 = New System.Windows.Forms.PictureBox()
        Me.pic0 = New System.Windows.Forms.PictureBox()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.btnEditWeights = New System.Windows.Forms.Button()
        Me.lblpb = New System.Windows.Forms.Label()
        Me.pb = New System.Windows.Forms.ProgressBar()
        Me.grpProgress = New System.Windows.Forms.GroupBox()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.gpLegend.SuspendLayout()
        CType(Me.pic3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pic2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pic1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pic0, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpProgress.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnLoad
        '
        Me.btnLoad.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnLoad.Location = New System.Drawing.Point(13, 13)
        Me.btnLoad.Name = "btnLoad"
        Me.btnLoad.Size = New System.Drawing.Size(75, 23)
        Me.btnLoad.TabIndex = 0
        Me.btnLoad.Text = "&Login"
        Me.btnLoad.UseVisualStyleBackColor = True
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToDeleteRows = False
        Me.DataGridView1.AllowUserToOrderColumns = True
        Me.DataGridView1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGridView1.BackgroundColor = System.Drawing.Color.White
        Me.DataGridView1.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None
        DataGridViewCellStyle7.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle7.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle7.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle7.ForeColor = System.Drawing.Color.White
        DataGridViewCellStyle7.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle7.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle7.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridView1.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle7
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle8.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle8.BackColor = System.Drawing.Color.White
        DataGridViewCellStyle8.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle8.ForeColor = System.Drawing.Color.Black
        DataGridViewCellStyle8.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle8.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle8.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.DataGridView1.DefaultCellStyle = DataGridViewCellStyle8
        Me.DataGridView1.EnableHeadersVisualStyles = False
        Me.DataGridView1.Location = New System.Drawing.Point(0, 104)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        Me.DataGridView1.RowHeadersVisible = False
        DataGridViewCellStyle9.BackColor = System.Drawing.Color.White
        Me.DataGridView1.RowsDefaultCellStyle = DataGridViewCellStyle9
        Me.DataGridView1.Size = New System.Drawing.Size(500, 564)
        Me.DataGridView1.TabIndex = 10
        Me.DataGridView1.Visible = False
        '
        'txtEmail
        '
        Me.txtEmail.BackColor = System.Drawing.SystemColors.Window
        Me.txtEmail.Location = New System.Drawing.Point(74, 42)
        Me.txtEmail.Name = "txtEmail"
        Me.txtEmail.Size = New System.Drawing.Size(202, 22)
        Me.txtEmail.TabIndex = 3
        '
        'lblEmail
        '
        Me.lblEmail.AutoSize = True
        Me.lblEmail.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblEmail.Location = New System.Drawing.Point(10, 45)
        Me.lblEmail.Name = "lblEmail"
        Me.lblEmail.Size = New System.Drawing.Size(37, 13)
        Me.lblEmail.TabIndex = 2
        Me.lblEmail.Text = "Email:"
        '
        'lblPassword
        '
        Me.lblPassword.AutoSize = True
        Me.lblPassword.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblPassword.Location = New System.Drawing.Point(9, 72)
        Me.lblPassword.Name = "lblPassword"
        Me.lblPassword.Size = New System.Drawing.Size(59, 13)
        Me.lblPassword.TabIndex = 4
        Me.lblPassword.Text = "Password:"
        '
        'ElementHost1
        '
        Me.ElementHost1.BackColor = System.Drawing.SystemColors.Window
        Me.ElementHost1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.ElementHost1.Location = New System.Drawing.Point(74, 68)
        Me.ElementHost1.Name = "ElementHost1"
        Me.ElementHost1.Size = New System.Drawing.Size(202, 22)
        Me.ElementHost1.TabIndex = 5
        Me.ElementHost1.Text = "ElementHost1"
        Me.ElementHost1.Child = Nothing
        '
        'btnOffline
        '
        Me.btnOffline.Location = New System.Drawing.Point(110, 13)
        Me.btnOffline.Name = "btnOffline"
        Me.btnOffline.Size = New System.Drawing.Size(75, 23)
        Me.btnOffline.TabIndex = 1
        Me.btnOffline.Text = "&Offline"
        Me.btnOffline.UseVisualStyleBackColor = True
        '
        'ElementHost2
        '
        Me.ElementHost2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ElementHost2.BackColor = System.Drawing.SystemColors.Window
        Me.ElementHost2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.ElementHost2.Location = New System.Drawing.Point(12, 104)
        Me.ElementHost2.Name = "ElementHost2"
        Me.ElementHost2.Size = New System.Drawing.Size(477, 554)
        Me.ElementHost2.TabIndex = 9
        Me.ElementHost2.Text = "ElementHost2"
        Me.ElementHost2.Child = Nothing
        '
        'cmbWeight
        '
        Me.cmbWeight.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend
        Me.cmbWeight.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.cmbWeight.BackColor = System.Drawing.SystemColors.Window
        Me.cmbWeight.FormattingEnabled = True
        Me.cmbWeight.Location = New System.Drawing.Point(368, 68)
        Me.cmbWeight.Name = "cmbWeight"
        Me.cmbWeight.Size = New System.Drawing.Size(121, 21)
        Me.cmbWeight.TabIndex = 8
        '
        'lblWeights
        '
        Me.lblWeights.AutoSize = True
        Me.lblWeights.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblWeights.Location = New System.Drawing.Point(293, 72)
        Me.lblWeights.Name = "lblWeights"
        Me.lblWeights.Size = New System.Drawing.Size(73, 13)
        Me.lblWeights.TabIndex = 6
        Me.lblWeights.Text = "Weights List:"
        '
        'chkSession
        '
        Me.chkSession.AutoSize = True
        Me.chkSession.BackColor = System.Drawing.SystemColors.Control
        Me.chkSession.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkSession.Location = New System.Drawing.Point(205, 19)
        Me.chkSession.Name = "chkSession"
        Me.chkSession.Size = New System.Drawing.Size(98, 17)
        Me.chkSession.TabIndex = 2
        Me.chkSession.Text = "&Use SessionID"
        Me.chkSession.UseVisualStyleBackColor = False
        '
        'gpLegend
        '
        Me.gpLegend.Controls.Add(Me.lblDoubleClick)
        Me.gpLegend.Controls.Add(Me.lblFixedValue)
        Me.gpLegend.Controls.Add(Me.pic3)
        Me.gpLegend.Controls.Add(Me.lblHighWeightMod)
        Me.gpLegend.Controls.Add(Me.lblLowWeightMod)
        Me.gpLegend.Controls.Add(Me.lblMax)
        Me.gpLegend.Controls.Add(Me.lblMin)
        Me.gpLegend.Controls.Add(Me.lblOtherSolutions)
        Me.gpLegend.Controls.Add(Me.lblUnknownValues)
        Me.gpLegend.Controls.Add(Me.lblHasGem)
        Me.gpLegend.Controls.Add(Me.pic2)
        Me.gpLegend.Controls.Add(Me.pic1)
        Me.gpLegend.Controls.Add(Me.pic0)
        Me.gpLegend.ForeColor = System.Drawing.SystemColors.ControlText
        Me.gpLegend.Location = New System.Drawing.Point(13, 13)
        Me.gpLegend.Name = "gpLegend"
        Me.gpLegend.Size = New System.Drawing.Size(835, 75)
        Me.gpLegend.TabIndex = 10
        Me.gpLegend.TabStop = False
        Me.gpLegend.Text = "Legend (hover over each entry for more details)"
        Me.gpLegend.Visible = False
        '
        'lblDoubleClick
        '
        Me.lblDoubleClick.AutoSize = True
        Me.lblDoubleClick.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDoubleClick.Location = New System.Drawing.Point(510, 23)
        Me.lblDoubleClick.Name = "lblDoubleClick"
        Me.lblDoubleClick.Size = New System.Drawing.Size(302, 13)
        Me.lblDoubleClick.TabIndex = 19
        Me.lblDoubleClick.Text = "Double-Click Enabled Columns: Rank, Prefix n, Suffix n, *"
        '
        'lblFixedValue
        '
        Me.lblFixedValue.AutoSize = True
        Me.lblFixedValue.Location = New System.Drawing.Point(172, 49)
        Me.lblFixedValue.Name = "lblFixedValue"
        Me.lblFixedValue.Size = New System.Drawing.Size(93, 13)
        Me.lblFixedValue.TabIndex = 18
        Me.lblFixedValue.Text = "Fixed Value Mod"
        '
        'pic3
        '
        Me.pic3.Location = New System.Drawing.Point(148, 46)
        Me.pic3.Name = "pic3"
        Me.pic3.Size = New System.Drawing.Size(20, 20)
        Me.pic3.TabIndex = 17
        Me.pic3.TabStop = False
        '
        'lblHighWeightMod
        '
        Me.lblHighWeightMod.AutoSize = True
        Me.lblHighWeightMod.Location = New System.Drawing.Point(392, 23)
        Me.lblHighWeightMod.Name = "lblHighWeightMod"
        Me.lblHighWeightMod.Size = New System.Drawing.Size(100, 13)
        Me.lblHighWeightMod.TabIndex = 16
        Me.lblHighWeightMod.Text = "High Weight Mod"
        '
        'lblLowWeightMod
        '
        Me.lblLowWeightMod.AutoSize = True
        Me.lblLowWeightMod.Location = New System.Drawing.Point(392, 49)
        Me.lblLowWeightMod.Name = "lblLowWeightMod"
        Me.lblLowWeightMod.Size = New System.Drawing.Size(96, 13)
        Me.lblLowWeightMod.TabIndex = 15
        Me.lblLowWeightMod.Text = "Low Weight Mod"
        '
        'lblMax
        '
        Me.lblMax.AutoSize = True
        Me.lblMax.Location = New System.Drawing.Point(317, 23)
        Me.lblMax.Name = "lblMax"
        Me.lblMax.Size = New System.Drawing.Size(71, 13)
        Me.lblMax.TabIndex = 14
        Me.lblMax.Text = "Example text"
        '
        'lblMin
        '
        Me.lblMin.AutoSize = True
        Me.lblMin.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblMin.Location = New System.Drawing.Point(318, 49)
        Me.lblMin.Name = "lblMin"
        Me.lblMin.Size = New System.Drawing.Size(67, 13)
        Me.lblMin.TabIndex = 13
        Me.lblMin.Text = "Example text"
        '
        'lblOtherSolutions
        '
        Me.lblOtherSolutions.AutoSize = True
        Me.lblOtherSolutions.Location = New System.Drawing.Point(172, 23)
        Me.lblOtherSolutions.Name = "lblOtherSolutions"
        Me.lblOtherSolutions.Size = New System.Drawing.Size(126, 13)
        Me.lblOtherSolutions.TabIndex = 12
        Me.lblOtherSolutions.Text = "Alt Mod Solutions Exist"
        '
        'lblUnknownValues
        '
        Me.lblUnknownValues.AutoSize = True
        Me.lblUnknownValues.Location = New System.Drawing.Point(37, 49)
        Me.lblUnknownValues.Name = "lblUnknownValues"
        Me.lblUnknownValues.Size = New System.Drawing.Size(100, 13)
        Me.lblUnknownValues.TabIndex = 11
        Me.lblUnknownValues.Text = "Legacy Mod Value"
        '
        'lblHasGem
        '
        Me.lblHasGem.AutoSize = True
        Me.lblHasGem.Location = New System.Drawing.Point(37, 23)
        Me.lblHasGem.Name = "lblHasGem"
        Me.lblHasGem.Size = New System.Drawing.Size(85, 13)
        Me.lblHasGem.TabIndex = 3
        Me.lblHasGem.Text = "Gem Level Reqt"
        '
        'pic2
        '
        Me.pic2.Location = New System.Drawing.Point(148, 20)
        Me.pic2.Name = "pic2"
        Me.pic2.Size = New System.Drawing.Size(20, 20)
        Me.pic2.TabIndex = 2
        Me.pic2.TabStop = False
        '
        'pic1
        '
        Me.pic1.Location = New System.Drawing.Point(13, 46)
        Me.pic1.Name = "pic1"
        Me.pic1.Size = New System.Drawing.Size(20, 20)
        Me.pic1.TabIndex = 1
        Me.pic1.TabStop = False
        '
        'pic0
        '
        Me.pic0.Location = New System.Drawing.Point(13, 20)
        Me.pic0.Name = "pic0"
        Me.pic0.Size = New System.Drawing.Size(20, 20)
        Me.pic0.TabIndex = 0
        Me.pic0.TabStop = False
        '
        'ToolTip1
        '
        Me.ToolTip1.AutoPopDelay = 10000
        Me.ToolTip1.InitialDelay = 250
        Me.ToolTip1.ReshowDelay = 100
        '
        'btnEditWeights
        '
        Me.btnEditWeights.Location = New System.Drawing.Point(368, 42)
        Me.btnEditWeights.Name = "btnEditWeights"
        Me.btnEditWeights.Size = New System.Drawing.Size(121, 23)
        Me.btnEditWeights.TabIndex = 7
        Me.btnEditWeights.Text = "&Edit Weights"
        Me.btnEditWeights.UseVisualStyleBackColor = True
        '
        'lblpb
        '
        Me.lblpb.AutoSize = True
        Me.lblpb.Location = New System.Drawing.Point(66, 58)
        Me.lblpb.Name = "lblpb"
        Me.lblpb.Size = New System.Drawing.Size(145, 13)
        Me.lblpb.TabIndex = 12
        Me.lblpb.Text = "Please wait, recalculating..."
        Me.lblpb.UseWaitCursor = True
        Me.lblpb.Visible = False
        '
        'pb
        '
        Me.pb.Location = New System.Drawing.Point(9, 21)
        Me.pb.Name = "pb"
        Me.pb.Size = New System.Drawing.Size(259, 23)
        Me.pb.TabIndex = 11
        Me.pb.UseWaitCursor = True
        Me.pb.Visible = False
        '
        'grpProgress
        '
        Me.grpProgress.Controls.Add(Me.pb)
        Me.grpProgress.Controls.Add(Me.lblpb)
        Me.grpProgress.Location = New System.Drawing.Point(110, 208)
        Me.grpProgress.Name = "grpProgress"
        Me.grpProgress.Size = New System.Drawing.Size(277, 84)
        Me.grpProgress.TabIndex = 13
        Me.grpProgress.TabStop = False
        Me.grpProgress.Visible = False
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(501, 670)
        Me.Controls.Add(Me.grpProgress)
        Me.Controls.Add(Me.btnEditWeights)
        Me.Controls.Add(Me.chkSession)
        Me.Controls.Add(Me.lblWeights)
        Me.Controls.Add(Me.cmbWeight)
        Me.Controls.Add(Me.ElementHost2)
        Me.Controls.Add(Me.btnOffline)
        Me.Controls.Add(Me.ElementHost1)
        Me.Controls.Add(Me.lblPassword)
        Me.Controls.Add(Me.lblEmail)
        Me.Controls.Add(Me.txtEmail)
        Me.Controls.Add(Me.DataGridView1)
        Me.Controls.Add(Me.btnLoad)
        Me.Controls.Add(Me.gpLegend)
        Me.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "frmMain"
        Me.Text = "ModRank"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.gpLegend.ResumeLayout(False)
        Me.gpLegend.PerformLayout()
        CType(Me.pic3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pic2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pic1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pic0, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpProgress.ResumeLayout(False)
        Me.grpProgress.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnLoad As System.Windows.Forms.Button
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents txtEmail As System.Windows.Forms.TextBox
    Friend WithEvents lblEmail As System.Windows.Forms.Label
    Friend WithEvents lblPassword As System.Windows.Forms.Label
    Friend WithEvents ElementHost1 As System.Windows.Forms.Integration.ElementHost
    Friend WithEvents btnOffline As System.Windows.Forms.Button
    Friend WithEvents ElementHost2 As System.Windows.Forms.Integration.ElementHost
    Friend WithEvents cmbWeight As System.Windows.Forms.ComboBox
    Friend WithEvents lblWeights As System.Windows.Forms.Label
    Friend WithEvents chkSession As System.Windows.Forms.CheckBox
    Friend WithEvents gpLegend As System.Windows.Forms.GroupBox
    Friend WithEvents pic2 As System.Windows.Forms.PictureBox
    Friend WithEvents pic1 As System.Windows.Forms.PictureBox
    Friend WithEvents pic0 As System.Windows.Forms.PictureBox
    Friend WithEvents lblUnknownValues As System.Windows.Forms.Label
    Friend WithEvents lblHasGem As System.Windows.Forms.Label
    Friend WithEvents lblOtherSolutions As System.Windows.Forms.Label
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents lblHighWeightMod As System.Windows.Forms.Label
    Friend WithEvents lblLowWeightMod As System.Windows.Forms.Label
    Friend WithEvents lblMax As System.Windows.Forms.Label
    Friend WithEvents lblMin As System.Windows.Forms.Label
    Friend WithEvents lblFixedValue As System.Windows.Forms.Label
    Friend WithEvents pic3 As System.Windows.Forms.PictureBox
    Friend WithEvents btnEditWeights As System.Windows.Forms.Button
    Friend WithEvents lblDoubleClick As System.Windows.Forms.Label
    Friend WithEvents lblpb As System.Windows.Forms.Label
    Public WithEvents pb As System.Windows.Forms.ProgressBar
    Friend WithEvents grpProgress As System.Windows.Forms.GroupBox

End Class
