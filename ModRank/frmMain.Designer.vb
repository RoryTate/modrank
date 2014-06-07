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
        Dim DataGridViewCellStyle10 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle11 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle12 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.btnDownloadJSON = New System.Windows.Forms.Button()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.lblRecordCount = New System.Windows.Forms.Label()
        Me.btnSearch = New System.Windows.Forms.Button()
        Me.chkSession = New System.Windows.Forms.CheckBox()
        Me.ElementHost2 = New System.Windows.Forms.Integration.ElementHost()
        Me.btnOffline = New System.Windows.Forms.Button()
        Me.ElementHost1 = New System.Windows.Forms.Integration.ElementHost()
        Me.lblPassword = New System.Windows.Forms.Label()
        Me.lblEmail = New System.Windows.Forms.Label()
        Me.txtEmail = New System.Windows.Forms.TextBox()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.btnLoad = New System.Windows.Forms.Button()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.btnRefresh = New System.Windows.Forms.Button()
        Me.btnShopSearch = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnStoreOffline = New System.Windows.Forms.Button()
        Me.lblRecordCount2 = New System.Windows.Forms.Label()
        Me.DataGridView2 = New System.Windows.Forms.DataGridView()
        Me.txtThread = New System.Windows.Forms.TextBox()
        Me.lblpb = New System.Windows.Forms.Label()
        Me.pb = New System.Windows.Forms.ProgressBar()
        Me.grpProgress = New System.Windows.Forms.GroupBox()
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
        Me.btnEditWeights = New System.Windows.Forms.Button()
        Me.lblWeights = New System.Windows.Forms.Label()
        Me.cmbWeight = New System.Windows.Forms.ComboBox()
        Me.TabControl1.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage2.SuspendLayout()
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpProgress.SuspendLayout()
        Me.gpLegend.SuspendLayout()
        CType(Me.pic3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pic2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pic1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pic0, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'ToolTip1
        '
        Me.ToolTip1.AutoPopDelay = 10000
        Me.ToolTip1.InitialDelay = 250
        Me.ToolTip1.ReshowDelay = 100
        '
        'btnDownloadJSON
        '
        Me.btnDownloadJSON.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnDownloadJSON.Location = New System.Drawing.Point(8, 36)
        Me.btnDownloadJSON.Name = "btnDownloadJSON"
        Me.btnDownloadJSON.Size = New System.Drawing.Size(73, 23)
        Me.btnDownloadJSON.TabIndex = 1
        Me.btnDownloadJSON.Text = "&Download"
        Me.btnDownloadJSON.UseVisualStyleBackColor = True
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage2)
        Me.TabControl1.DrawMode = System.Windows.Forms.TabDrawMode.OwnerDrawFixed
        Me.TabControl1.Location = New System.Drawing.Point(0, 12)
        Me.TabControl1.Margin = New System.Windows.Forms.Padding(0, 3, 0, 3)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(506, 660)
        Me.TabControl1.TabIndex = 17
        '
        'TabPage1
        '
        Me.TabPage1.BackColor = System.Drawing.SystemColors.Control
        Me.TabPage1.Controls.Add(Me.lblRecordCount)
        Me.TabPage1.Controls.Add(Me.btnSearch)
        Me.TabPage1.Controls.Add(Me.chkSession)
        Me.TabPage1.Controls.Add(Me.ElementHost2)
        Me.TabPage1.Controls.Add(Me.btnOffline)
        Me.TabPage1.Controls.Add(Me.ElementHost1)
        Me.TabPage1.Controls.Add(Me.lblPassword)
        Me.TabPage1.Controls.Add(Me.lblEmail)
        Me.TabPage1.Controls.Add(Me.txtEmail)
        Me.TabPage1.Controls.Add(Me.DataGridView1)
        Me.TabPage1.Controls.Add(Me.btnLoad)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Margin = New System.Windows.Forms.Padding(0, 3, 0, 3)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Padding = New System.Windows.Forms.Padding(0, 3, 0, 3)
        Me.TabPage1.Size = New System.Drawing.Size(498, 634)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "Stash"
        '
        'lblRecordCount
        '
        Me.lblRecordCount.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblRecordCount.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblRecordCount.Location = New System.Drawing.Point(288, 79)
        Me.lblRecordCount.Name = "lblRecordCount"
        Me.lblRecordCount.Size = New System.Drawing.Size(207, 13)
        Me.lblRecordCount.TabIndex = 30
        Me.lblRecordCount.Text = "Number of Rows:"
        Me.lblRecordCount.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.lblRecordCount.Visible = False
        '
        'btnSearch
        '
        Me.btnSearch.Location = New System.Drawing.Point(336, 8)
        Me.btnSearch.Name = "btnSearch"
        Me.btnSearch.Size = New System.Drawing.Size(150, 23)
        Me.btnSearch.TabIndex = 23
        Me.btnSearch.Text = "&Activate Search/Filter"
        Me.btnSearch.UseVisualStyleBackColor = True
        '
        'chkSession
        '
        Me.chkSession.AutoSize = True
        Me.chkSession.BackColor = System.Drawing.Color.Transparent
        Me.chkSession.ForeColor = System.Drawing.SystemColors.ControlText
        Me.chkSession.Location = New System.Drawing.Point(199, 14)
        Me.chkSession.Name = "chkSession"
        Me.chkSession.Size = New System.Drawing.Size(98, 17)
        Me.chkSession.TabIndex = 18
        Me.chkSession.Text = "&Use SessionID"
        Me.chkSession.UseVisualStyleBackColor = False
        '
        'ElementHost2
        '
        Me.ElementHost2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ElementHost2.BackColor = System.Drawing.SystemColors.Window
        Me.ElementHost2.ForeColor = System.Drawing.SystemColors.WindowText
        Me.ElementHost2.Location = New System.Drawing.Point(2, 99)
        Me.ElementHost2.Name = "ElementHost2"
        Me.ElementHost2.Size = New System.Drawing.Size(489, 535)
        Me.ElementHost2.TabIndex = 28
        Me.ElementHost2.Text = "ElementHost2"
        Me.ElementHost2.Child = Nothing
        '
        'btnOffline
        '
        Me.btnOffline.Location = New System.Drawing.Point(107, 8)
        Me.btnOffline.Name = "btnOffline"
        Me.btnOffline.Size = New System.Drawing.Size(75, 23)
        Me.btnOffline.TabIndex = 17
        Me.btnOffline.Text = "&Offline"
        Me.btnOffline.UseVisualStyleBackColor = True
        '
        'ElementHost1
        '
        Me.ElementHost1.BackColor = System.Drawing.SystemColors.Window
        Me.ElementHost1.ForeColor = System.Drawing.SystemColors.WindowText
        Me.ElementHost1.Location = New System.Drawing.Point(71, 63)
        Me.ElementHost1.Name = "ElementHost1"
        Me.ElementHost1.Size = New System.Drawing.Size(202, 22)
        Me.ElementHost1.TabIndex = 22
        Me.ElementHost1.Text = "ElementHost1"
        Me.ElementHost1.Child = Nothing
        '
        'lblPassword
        '
        Me.lblPassword.AutoSize = True
        Me.lblPassword.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblPassword.Location = New System.Drawing.Point(3, 67)
        Me.lblPassword.Name = "lblPassword"
        Me.lblPassword.Size = New System.Drawing.Size(59, 13)
        Me.lblPassword.TabIndex = 21
        Me.lblPassword.Text = "Password:"
        '
        'lblEmail
        '
        Me.lblEmail.AutoSize = True
        Me.lblEmail.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblEmail.Location = New System.Drawing.Point(4, 40)
        Me.lblEmail.Name = "lblEmail"
        Me.lblEmail.Size = New System.Drawing.Size(37, 13)
        Me.lblEmail.TabIndex = 19
        Me.lblEmail.Text = "Email:"
        '
        'txtEmail
        '
        Me.txtEmail.BackColor = System.Drawing.SystemColors.Window
        Me.txtEmail.Location = New System.Drawing.Point(71, 37)
        Me.txtEmail.Name = "txtEmail"
        Me.txtEmail.Size = New System.Drawing.Size(202, 22)
        Me.txtEmail.TabIndex = 20
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
        Me.DataGridView1.Location = New System.Drawing.Point(0, 99)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        Me.DataGridView1.RowHeadersVisible = False
        DataGridViewCellStyle9.BackColor = System.Drawing.Color.White
        Me.DataGridView1.RowsDefaultCellStyle = DataGridViewCellStyle9
        Me.DataGridView1.Size = New System.Drawing.Size(498, 535)
        Me.DataGridView1.TabIndex = 29
        Me.DataGridView1.Visible = False
        '
        'btnLoad
        '
        Me.btnLoad.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnLoad.Location = New System.Drawing.Point(10, 8)
        Me.btnLoad.Name = "btnLoad"
        Me.btnLoad.Size = New System.Drawing.Size(75, 23)
        Me.btnLoad.TabIndex = 16
        Me.btnLoad.Text = "&Login"
        Me.btnLoad.UseVisualStyleBackColor = True
        '
        'TabPage2
        '
        Me.TabPage2.BackColor = System.Drawing.SystemColors.Control
        Me.TabPage2.Controls.Add(Me.btnRefresh)
        Me.TabPage2.Controls.Add(Me.btnShopSearch)
        Me.TabPage2.Controls.Add(Me.Label1)
        Me.TabPage2.Controls.Add(Me.btnStoreOffline)
        Me.TabPage2.Controls.Add(Me.lblRecordCount2)
        Me.TabPage2.Controls.Add(Me.DataGridView2)
        Me.TabPage2.Controls.Add(Me.txtThread)
        Me.TabPage2.Controls.Add(Me.btnDownloadJSON)
        Me.TabPage2.Location = New System.Drawing.Point(4, 22)
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.Padding = New System.Windows.Forms.Padding(3)
        Me.TabPage2.Size = New System.Drawing.Size(498, 634)
        Me.TabPage2.TabIndex = 1
        Me.TabPage2.Text = "Store"
        '
        'btnRefresh
        '
        Me.btnRefresh.Location = New System.Drawing.Point(8, 66)
        Me.btnRefresh.Name = "btnRefresh"
        Me.btnRefresh.Size = New System.Drawing.Size(165, 23)
        Me.btnRefresh.TabIndex = 33
        Me.btnRefresh.Text = "&Refresh Local Store Cache"
        Me.btnRefresh.UseVisualStyleBackColor = True
        '
        'btnShopSearch
        '
        Me.btnShopSearch.Location = New System.Drawing.Point(336, 8)
        Me.btnShopSearch.Name = "btnShopSearch"
        Me.btnShopSearch.Size = New System.Drawing.Size(150, 23)
        Me.btnShopSearch.TabIndex = 32
        Me.btnShopSearch.Text = "&Activate Search/Filter"
        Me.btnShopSearch.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Label1.Location = New System.Drawing.Point(11, 12)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(75, 13)
        Me.Label1.TabIndex = 31
        Me.Label1.Text = "Store Thread:"
        '
        'btnStoreOffline
        '
        Me.btnStoreOffline.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnStoreOffline.Location = New System.Drawing.Point(100, 36)
        Me.btnStoreOffline.Name = "btnStoreOffline"
        Me.btnStoreOffline.Size = New System.Drawing.Size(73, 23)
        Me.btnStoreOffline.TabIndex = 4
        Me.btnStoreOffline.Text = "&Offline"
        Me.btnStoreOffline.UseVisualStyleBackColor = True
        '
        'lblRecordCount2
        '
        Me.lblRecordCount2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblRecordCount2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblRecordCount2.Location = New System.Drawing.Point(288, 79)
        Me.lblRecordCount2.Name = "lblRecordCount2"
        Me.lblRecordCount2.Size = New System.Drawing.Size(207, 13)
        Me.lblRecordCount2.TabIndex = 2
        Me.lblRecordCount2.Text = "Number of Rows:"
        Me.lblRecordCount2.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.lblRecordCount2.Visible = False
        '
        'DataGridView2
        '
        Me.DataGridView2.AllowUserToAddRows = False
        Me.DataGridView2.AllowUserToDeleteRows = False
        Me.DataGridView2.AllowUserToOrderColumns = True
        Me.DataGridView2.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGridView2.BackgroundColor = System.Drawing.Color.White
        Me.DataGridView2.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None
        DataGridViewCellStyle10.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle10.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle10.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle10.ForeColor = System.Drawing.Color.White
        DataGridViewCellStyle10.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle10.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle10.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.DataGridView2.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle10
        Me.DataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle11.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle11.BackColor = System.Drawing.Color.White
        DataGridViewCellStyle11.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle11.ForeColor = System.Drawing.Color.Black
        DataGridViewCellStyle11.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle11.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle11.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.DataGridView2.DefaultCellStyle = DataGridViewCellStyle11
        Me.DataGridView2.EnableHeadersVisualStyles = False
        Me.DataGridView2.Location = New System.Drawing.Point(0, 99)
        Me.DataGridView2.Name = "DataGridView2"
        Me.DataGridView2.ReadOnly = True
        Me.DataGridView2.RowHeadersVisible = False
        DataGridViewCellStyle12.BackColor = System.Drawing.Color.White
        Me.DataGridView2.RowsDefaultCellStyle = DataGridViewCellStyle12
        Me.DataGridView2.Size = New System.Drawing.Size(498, 535)
        Me.DataGridView2.TabIndex = 3
        Me.DataGridView2.Visible = False
        '
        'txtThread
        '
        Me.txtThread.Location = New System.Drawing.Point(89, 8)
        Me.txtThread.Name = "txtThread"
        Me.txtThread.Size = New System.Drawing.Size(215, 22)
        Me.txtThread.TabIndex = 0
        '
        'lblpb
        '
        Me.lblpb.Location = New System.Drawing.Point(5, 58)
        Me.lblpb.Name = "lblpb"
        Me.lblpb.Size = New System.Drawing.Size(297, 13)
        Me.lblpb.TabIndex = 12
        Me.lblpb.Text = "Please wait, recalculating..."
        Me.lblpb.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.lblpb.UseWaitCursor = True
        Me.lblpb.Visible = False
        '
        'pb
        '
        Me.pb.Location = New System.Drawing.Point(9, 21)
        Me.pb.Name = "pb"
        Me.pb.Size = New System.Drawing.Size(289, 23)
        Me.pb.TabIndex = 11
        Me.pb.UseWaitCursor = True
        Me.pb.Visible = False
        '
        'grpProgress
        '
        Me.grpProgress.BackColor = System.Drawing.SystemColors.Control
        Me.grpProgress.Controls.Add(Me.lblpb)
        Me.grpProgress.Controls.Add(Me.pb)
        Me.grpProgress.Location = New System.Drawing.Point(95, 208)
        Me.grpProgress.Name = "grpProgress"
        Me.grpProgress.Size = New System.Drawing.Size(307, 84)
        Me.grpProgress.TabIndex = 14
        Me.grpProgress.TabStop = False
        Me.grpProgress.Visible = False
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
        Me.gpLegend.Location = New System.Drawing.Point(9, 41)
        Me.gpLegend.Name = "gpLegend"
        Me.gpLegend.Size = New System.Drawing.Size(870, 75)
        Me.gpLegend.TabIndex = 28
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
        Me.lblDoubleClick.Size = New System.Drawing.Size(355, 13)
        Me.lblDoubleClick.TabIndex = 19
        Me.lblDoubleClick.Text = "Drill-Down Columns: Rank, Prefix n, Suffix n, *, Location, Sokt, Link"
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
        'btnEditWeights
        '
        Me.btnEditWeights.Location = New System.Drawing.Point(369, 70)
        Me.btnEditWeights.Name = "btnEditWeights"
        Me.btnEditWeights.Size = New System.Drawing.Size(121, 23)
        Me.btnEditWeights.TabIndex = 29
        Me.btnEditWeights.Text = "&Edit Weights"
        Me.btnEditWeights.UseVisualStyleBackColor = True
        '
        'lblWeights
        '
        Me.lblWeights.AutoSize = True
        Me.lblWeights.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblWeights.Location = New System.Drawing.Point(291, 100)
        Me.lblWeights.Name = "lblWeights"
        Me.lblWeights.Size = New System.Drawing.Size(73, 13)
        Me.lblWeights.TabIndex = 30
        Me.lblWeights.Text = "Weights List:"
        '
        'cmbWeight
        '
        Me.cmbWeight.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend
        Me.cmbWeight.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.cmbWeight.BackColor = System.Drawing.SystemColors.Window
        Me.cmbWeight.FormattingEnabled = True
        Me.cmbWeight.Location = New System.Drawing.Point(369, 96)
        Me.cmbWeight.Name = "cmbWeight"
        Me.cmbWeight.Size = New System.Drawing.Size(121, 21)
        Me.cmbWeight.TabIndex = 31
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(501, 670)
        Me.Controls.Add(Me.grpProgress)
        Me.Controls.Add(Me.btnEditWeights)
        Me.Controls.Add(Me.lblWeights)
        Me.Controls.Add(Me.cmbWeight)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.gpLegend)
        Me.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "frmMain"
        Me.Text = "ModRank"
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpProgress.ResumeLayout(False)
        Me.gpLegend.ResumeLayout(False)
        Me.gpLegend.PerformLayout()
        CType(Me.pic3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pic2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pic1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pic0, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents btnDownloadJSON As System.Windows.Forms.Button
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents lblRecordCount As System.Windows.Forms.Label
    Friend WithEvents btnSearch As System.Windows.Forms.Button
    Friend WithEvents chkSession As System.Windows.Forms.CheckBox
    Friend WithEvents ElementHost2 As System.Windows.Forms.Integration.ElementHost
    Friend WithEvents btnOffline As System.Windows.Forms.Button
    Friend WithEvents ElementHost1 As System.Windows.Forms.Integration.ElementHost
    Friend WithEvents lblPassword As System.Windows.Forms.Label
    Friend WithEvents lblEmail As System.Windows.Forms.Label
    Friend WithEvents txtEmail As System.Windows.Forms.TextBox
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents btnLoad As System.Windows.Forms.Button
    Friend WithEvents TabPage2 As System.Windows.Forms.TabPage
    Friend WithEvents lblpb As System.Windows.Forms.Label
    Public WithEvents pb As System.Windows.Forms.ProgressBar
    Friend WithEvents grpProgress As System.Windows.Forms.GroupBox
    Friend WithEvents txtThread As System.Windows.Forms.TextBox
    Friend WithEvents DataGridView2 As System.Windows.Forms.DataGridView
    Friend WithEvents lblRecordCount2 As System.Windows.Forms.Label
    Friend WithEvents gpLegend As System.Windows.Forms.GroupBox
    Friend WithEvents lblDoubleClick As System.Windows.Forms.Label
    Friend WithEvents lblFixedValue As System.Windows.Forms.Label
    Friend WithEvents pic3 As System.Windows.Forms.PictureBox
    Friend WithEvents lblHighWeightMod As System.Windows.Forms.Label
    Friend WithEvents lblLowWeightMod As System.Windows.Forms.Label
    Friend WithEvents lblMax As System.Windows.Forms.Label
    Friend WithEvents lblMin As System.Windows.Forms.Label
    Friend WithEvents lblOtherSolutions As System.Windows.Forms.Label
    Friend WithEvents lblUnknownValues As System.Windows.Forms.Label
    Friend WithEvents lblHasGem As System.Windows.Forms.Label
    Friend WithEvents pic2 As System.Windows.Forms.PictureBox
    Friend WithEvents pic1 As System.Windows.Forms.PictureBox
    Friend WithEvents pic0 As System.Windows.Forms.PictureBox
    Friend WithEvents btnStoreOffline As System.Windows.Forms.Button
    Friend WithEvents btnEditWeights As System.Windows.Forms.Button
    Friend WithEvents lblWeights As System.Windows.Forms.Label
    Friend WithEvents cmbWeight As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnShopSearch As System.Windows.Forms.Button
    Friend WithEvents btnRefresh As System.Windows.Forms.Button

End Class
