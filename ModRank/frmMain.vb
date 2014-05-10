Imports System
Imports System.IO
Imports POEApi.Model
Imports POEApi.Infrastructure
Imports POEApi.Transport
Imports System.Security
Imports System.Windows.Controls
Imports System.Net
Imports System.Reflection
Imports System.Drawing.Drawing2D

Imports ModRank.Extenders.Strings.StringExtender
Imports System.Drawing.Imaging
Imports POEApi.Model.Events
Imports POEApi.Infrastructure.Events
Imports System.Deployment.Application

Public Class frmMain
    Public dtRank As DataTable = New DataTable()
    Public dtOverflow As DataTable = New DataTable
    Public dtMods As DataTable = New DataTable()
    Public dtWeights As DataTable = New DataTable()
    Public RecalculateThread As Threading.Thread
    Public ProgressBarThread As Threading.Thread

    Private WPFPassword As New PasswordBox
    Private statusBox As Windows.Controls.RichTextBox = New RichTextBox
    Private statusController As ModRank.frmMain.StatusController
    Private blOffline As Boolean = False

    Public FullStash As Stash
    Public Shared FullInventory As New List(Of FullItem)
    Public Shared TempInventory As New List(Of FullItem)
    Public invLocation As New Dictionary(Of Item, String)
    Public Shared Characters As New List(Of Character)()
    Public Shared Leagues As New List(Of String)()
    Public Shared FontCollection As New Drawing.Text.PrivateFontCollection()
    Public statusCounter As Long, lngModCounter As Long
    Public blSolomonsJudgment As New Dictionary(Of String, Boolean)        ' Used to decide which mod gets both .5's in a two-way combined split
    Public blAddedOne As Boolean = False    '  Used in the dynamic mod evaluation method, to help know when a mod contains "Legacy" values
    Public blScroll As Boolean = False

    Dim oldColorDark As Color = Color.FromArgb(127, 127, 127)
    Dim oldColorLight As Color = Color.FromArgb(195, 195, 195)

    Public RankExplanation As New Dictionary(Of String, String)

    Private Delegate Sub FillCredentialsDelegate()
    Public Delegate Sub MyDelegate()
    Public Delegate Function MyDelegateFunction()
    Public Delegate Sub MyControlDelegate(myControl As Object)

    Public blRepopulated As Boolean = False

    ' RCPD = Read Control Property Delegate (used a lot so abbreviated)
    Public Delegate Function RCPD(ByVal MyControl As Object, ByVal MyProperty As Object) As String
    ' UCPD = Update Control Property Delegate
    Public Delegate Sub UCPD(ByVal MyControl As Object, ByVal MyProperty As Object, ByVal MyValue As Object)

    Public bytColumns As Byte = 0       ' The max number of columns displayed in the datagridview

    Private Shared m_currentCharacter As Character = Nothing

    Public Shared Property CurrentCharacter() As Character
        Get
            Return m_currentCharacter
        End Get
        Set(value As Character)
            m_currentCharacter = value
            RaiseEvent CharacterChanged(Model, New System.ComponentModel.PropertyChangedEventArgs("CurrentCharacter"))
        End Set
    End Property

    Public Shared Event LeagueChanged As System.ComponentModel.PropertyChangedEventHandler
    Public Shared Event CharacterChanged As System.ComponentModel.PropertyChangedEventHandler

    Private Shared m_currentLeague As String = String.Empty

    Public Shared Property CurrentLeague() As String
        Get
            Return m_currentLeague
        End Get
        Set(value As String)
            If m_currentLeague = value Then
                Return
            End If

            m_currentLeague = value
            Characters = Model.GetCharacters().Where(Function(c) c.League = value).ToList()
            CurrentCharacter = Characters.First()
            RaiseEvent LeagueChanged(Model, New System.ComponentModel.PropertyChangedEventArgs("CurrentLeague"))
        End Set
    End Property

    Private Sub frmMain_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        If blRepopulated = False Then Exit Sub
        ' Pull down the combo box, since the user will likely want to select their new weights file to use
        cmbWeight.DroppedDown = True
        blRepopulated = False
    End Sub

    Private Sub frmMain_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        Try
            strWeightFile = cmbWeight.Text
            Settings.UserSettings("WeightFile") = strWeightFile
            Settings.Save()
        Catch ex As Exception
            ErrorHandler(System.Reflection.MethodBase.GetCurrentMethod.Name, ex)
        End Try
    End Sub

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles Me.Load
        Try
            Me.Icon = GetEmbeddedIcon("ModRank.PoE.ico")
            Me.Text = "ModRank v" & My.Application.Info.Version.ToString

            ' Let's see if we can improve the display performance of the datagridview a bit...holy crap! This setting works unbelievably well!
            GetType(DataGridView).InvokeMember("DoubleBuffered", BindingFlags.NonPublic Or BindingFlags.Instance Or BindingFlags.SetProperty, Nothing, DataGridView1, New Object() {True})
            gpLegend.Visible = False
            Dim baseImage As Bitmap = GetEmbeddedBitmap("ModRank.legend.png")
            Dim Colors() As Color = {ColorHasGem, ColorUnknownValue, ColorOtherSolutions, ColorNoModVariation}
            For x As Integer = 0 To 3
                Dim img As New Bitmap(10, 10)
                Dim gr As Graphics = Graphics.FromImage(img)
                gr.Clear(Colors(x))
                Dim newColorLight As Color = getColorLight(img)

                Dim newImage1 As Bitmap = DirectCast(baseImage.Clone, Bitmap)
                Dim att As New ImageAttributes
                Dim map(1) As ColorMap
                map(0) = New ColorMap
                map(0).OldColor = oldColorDark
                map(0).NewColor = Colors(x)
                map(1) = New ColorMap
                map(1).OldColor = oldColorLight
                map(1).NewColor = newColorLight
                att.SetRemapTable(map)

                Dim newImage2 As New Bitmap(newImage1.Width, newImage1.Height)
                gr = Graphics.FromImage(newImage2)
                gr.DrawImage(newImage1, New Rectangle(0, 0, newImage1.Width, newImage1.Height), 0, 0, newImage1.Width, newImage1.Height, Drawing.GraphicsUnit.Pixel, att)
                Dim MyPic As PictureBox = Me.gpLegend.Controls("pic" & x)
                MyPic.Image = newImage2
            Next

            gpLegend.Controls("lblMin").ForeColor = ColorMin
            gpLegend.Controls("lblMax").ForeColor = ColorMax

            LoadSettings()

            If UserSettings("WeightFile").Trim = "" Then UserSettings("WeightFile") = "default"
            If File.Exists(Application.StartupPath & "\weights-" & UserSettings("WeightFile") & ".csv") = False Then
                UserSettings("WeightFile") = "default"
            End If
            ' Set locale for dtweights too...users might want to set decimal weights
            dtWeights.Locale = Globalization.CultureInfo.InvariantCulture
            dtWeights = LoadCSVtoDataTable(Application.StartupPath & "\weights-" & UserSettings("WeightFile") & ".csv")

            dtMods.Locale = Globalization.CultureInfo.InvariantCulture
            dtMods.Columns.Add("Categories", GetType(String))
            dtMods.Columns.Add("Description", GetType(String))
            dtMods.Columns.Add("Value", GetType(String))
            dtMods.Columns.Add("MinV", GetType(Single))
            dtMods.Columns.Add("MaxV", GetType(Single))
            dtMods.Columns.Add("MinV2", GetType(Single))
            dtMods.Columns.Add("MaxV2", GetType(Single))
            dtMods.Columns.Add("Name", GetType(String))
            dtMods.Columns.Add("Level", GetType(Integer))
            dtMods.Columns.Add("Prefix/Suffix", GetType(String))
            dtMods.Columns.Add("Ring", GetType(Boolean))
            dtMods.Columns.Add("Amulet", GetType(Boolean))
            dtMods.Columns.Add("Belt", GetType(Boolean))
            dtMods.Columns.Add("Helmet", GetType(Boolean))
            dtMods.Columns.Add("Gloves", GetType(Boolean))
            dtMods.Columns.Add("Boots", GetType(Boolean))
            dtMods.Columns.Add("Chest", GetType(Boolean))
            dtMods.Columns.Add("Shield", GetType(Boolean))
            dtMods.Columns.Add("Quiver", GetType(Boolean))
            dtMods.Columns.Add("Wand", GetType(Boolean))
            dtMods.Columns.Add("Dagger", GetType(Boolean))
            dtMods.Columns.Add("Claw", GetType(Boolean))
            dtMods.Columns.Add("Sceptre", GetType(Boolean))
            dtMods.Columns.Add("Staff", GetType(Boolean))
            dtMods.Columns.Add("Bow", GetType(Boolean))
            dtMods.Columns.Add("1h Swords and Axes", GetType(Boolean))
            dtMods.Columns.Add("2h Swords and Axes", GetType(Boolean))
            dtMods.Columns.Add("1h Maces", GetType(Boolean))
            dtMods.Columns.Add("2h Maces", GetType(Boolean))

            Dim dtTemp As New DataTable()
            dtTemp = LoadCSVtoDataTable(Application.StartupPath & "\mods.csv")
            ' Now we copy dtTemp into dtMods...this is done so the field structure is correct for querying/sorting/etc
            For Each dtRow As DataRow In dtTemp.Rows
                Dim newRow As DataRow = dtMods.NewRow
                newRow("Categories") = dtRow("Categories")
                newRow("Description") = dtRow("Description")
                newRow("Value") = dtRow("Value")
                newRow("MinV") = IIf(IsDBNull(dtRow("MinV")) Or dtRow("MinV") = "", 0, dtRow("MinV"))
                newRow("MaxV") = IIf(IsDBNull(dtRow("MaxV")) Or dtRow("MaxV") = "", 0, dtRow("MaxV"))
                newRow("MinV2") = IIf(IsDBNull(dtRow("MinV2")) Or dtRow("MinV2") = "", 0, dtRow("MinV2"))
                newRow("MaxV2") = IIf(IsDBNull(dtRow("MaxV2")) Or dtRow("MaxV2") = "", 0, dtRow("MaxV2"))
                newRow("Name") = dtRow("Name")
                newRow("Level") = IIf(IsDBNull(dtRow("Level")) Or dtRow("Level") = "", 0, dtRow("Level"))
                newRow("Prefix/Suffix") = dtRow("Prefix/Suffix")
                For i = 10 To 28
                    newRow(i) = IIf(dtRow(i).ToString.IndexOf("Yes", StringComparison.OrdinalIgnoreCase) > -1, True, False)
                Next
                dtMods.Rows.Add(newRow)
            Next

            ' Create dtRank datatable schema
            dtRank.Columns.Add("Rank", GetType(Integer))
            dtRank.Columns.Add("%", GetType(Single))
            dtRank.Columns.Add("Type", GetType(String))
            dtRank.Columns.Add("SubType", GetType(String))
            dtRank.Columns.Add("Leag", GetType(String))
            dtRank.Columns.Add("Location", GetType(String))
            For b As Byte = 1 To 3 : dtRank.Columns.Add("Prefix " & b, GetType(String)) : Next
            For b As Byte = 1 To 3 : dtRank.Columns.Add("Suffix " & b, GetType(String)) : Next
            dtRank.Columns.Add("Implicit", GetType(String))
            dtRank.Columns.Add("*", GetType(String))
            dtRank.Columns.Add("Name", GetType(String))
            dtRank.Columns.Add("Level", GetType(Integer))
            dtRank.Columns.Add("Gem", GetType(Boolean))
            dtRank.Columns.Add("Sokt", GetType(Byte))
            dtRank.Columns.Add("Link", GetType(Byte))
            dtRank.Columns.Add("Corrupt", GetType(Boolean))
            dtRank.Columns.Add("Index", GetType(Long))
            dtRank.Columns.Add("ID", GetType(String))
            Dim primaryKey(1) As DataColumn
            primaryKey(1) = dtRank.Columns("ID")
            dtRank.PrimaryKey = primaryKey

            dtOverflow = dtRank.Clone

            WPFPassword.BorderThickness = New Windows.Thickness(1.0)
            Dim winColor As Color = SystemColors.Window
            Dim txtColor As Color = SystemColors.WindowText
            'Dim myBrush As System.Windows.Media.Brush = converter.convert("#ABADB3")
            'WPFPassword.BorderBrush = myBrush
            WPFPassword.Background = New System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromArgb(winColor.A, winColor.R, winColor.G, winColor.B))
            WPFPassword.Foreground = New System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromArgb(txtColor.A, txtColor.R, txtColor.G, txtColor.B))

            ElementHost1.Child = WPFPassword
            statusBox.IsReadOnly = True
            statusBox.BorderThickness = New Windows.Thickness(1.0)
            'statusBox.BorderBrush = myBrush
            statusBox.Background = New System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromArgb(winColor.A, winColor.R, winColor.G, winColor.B))
            statusBox.Foreground = New System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromArgb(txtColor.A, txtColor.R, txtColor.G, txtColor.B))
            ElementHost2.Child = statusBox

            statusController = New ModRank.frmMain.StatusController(Me.statusBox)

            txtEmail.Text = UserSettings("AccountLogin")
            chkSession.Checked = UserSettings("UseSessionID")
            blFormChanged = String.IsNullOrEmpty(Settings.UserSettings("AccountPassword"))
            If blFormChanged = False Then WPFPassword.Password = String.Empty.PadLeft(8)
            AddHandler WPFPassword.PasswordChanged, AddressOf PasswordChanged
            AddHandler WPFPassword.IsEnabledChanged, AddressOf PasswordEnabledChanged

            PopulateWeightsComboBox(UserSettings("WeightFile"))

            Dim strHasGem As String = "A Rank, %, Level, and Gem field with this background color indicates that the level requirements for this item are set by a gem. The mod level ranges will likely be increased incorrectly, and the ranking and percentile scores are artifically lowered as a result." & _
                Environment.NewLine & Environment.NewLine & "It is recommended that the gem be unequipped from the item before evaluating it further using the rank and percentile scores."
            ToolTip1.SetToolTip(lblHasGem, strHasGem)
            ToolTip1.SetToolTip(pic0, strHasGem)
            Dim strUnknownValues As String = "A mod with this background color indicates that the mod value is outside of the possible ranges, meaning that this is likely a legacy mod value." & _
                Environment.NewLine & Environment.NewLine & "Click any cell with this background color to get more information on why this was considered a legacy mod value."
            ToolTip1.SetToolTip(lblUnknownValues, strUnknownValues)
            ToolTip1.SetToolTip(pic1, strUnknownValues)
            Dim strOtherSolutions As String = "A '*' column with this background color indicates that alternate solutions of prefixes/suffixes and/or values exist for this item." & _
                Environment.NewLine & Environment.NewLine & "Click any cell with this background color to view the alternate configurations of prefix/suffix mods for this item."
            ToolTip1.SetToolTip(lblOtherSolutions, strOtherSolutions)
            ToolTip1.SetToolTip(pic2, strOtherSolutions)
            Dim strFixedValue As String = "A mod with this background color has no possible variation in its values or level, so the normal 'mod level bar' would be misleading." & _
                Environment.NewLine & Environment.NewLine & "The mod is instead displayed with this background color to show that it is 'fixed' in value."
            ToolTip1.SetToolTip(lblFixedValue, strFixedValue)
            ToolTip1.SetToolTip(pic3, strFixedValue)
            Dim strHighWeightMod As String = "Mod text in this font indicates the mod is of higher value. The trigger weight for this can be set or even disabled in the settings file." & _
                Environment.NewLine & Environment.NewLine & "Look for the following entry in settings.xml: <Setting name=""WeightMax"" value=""9"" />" & _
                Environment.NewLine & Environment.NewLine & "Note: A greater than or equal to (>=) comparison is used for this value."
            ToolTip1.SetToolTip(lblHighWeightMod, strHighWeightMod)
            ToolTip1.SetToolTip(lblMax, strHighWeightMod)
            Dim strLowWeightMod As String = "Mod text in this font indicates the mod is of lower value. The trigger weight for this can be set or even disabled in the settings file." & _
                Environment.NewLine & Environment.NewLine & "Look for the following entry in settings.xml: <Setting name=""WeightMin"" value=""1"" />" & _
                Environment.NewLine & Environment.NewLine & "Note: A less than or equal to (<=) comparison is used for this value."
            ToolTip1.SetToolTip(lblLowWeightMod, strLowWeightMod)
            ToolTip1.SetToolTip(lblMin, strLowWeightMod)
            Dim strDoubleClick As String = "Click on any of the Rank, Prefix, Suffix, and * cells (where values exist) to get more detailed information on the item."
            ToolTip1.SetToolTip(lblDoubleClick, strDoubleClick)

        Catch ex As Exception
            ErrorHandler(System.Reflection.MethodBase.GetCurrentMethod.Name, ex)
        End Try
    End Sub

    Public Sub PasswordChanged()
        blFormChanged = True
    End Sub

    Public Sub PasswordEnabledChanged()
        If WPFPassword.IsEnabled = True Then Exit Sub
        Dim ctlColor As Color = SystemColors.Control
        WPFPassword.Background = New System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromArgb(ctlColor.A, ctlColor.R, ctlColor.G, ctlColor.B))
    End Sub

    Public Sub PopulateWeightsComboBox(strSelection As String)
        Try
            cmbWeight.Items.Add("Default")
            Dim dir = Application.StartupPath
            For Each file As String In System.IO.Directory.GetFiles(dir)
                Dim strTemp As String = System.IO.Path.GetFileNameWithoutExtension(file)
                If strTemp.ToLower.StartsWith("weights") And file.ToLower.Contains("csv") Then
                    If strTemp.ToLower.Equals("weights-default") = False Then cmbWeight.Items.Add(StrConv(strTemp.Substring(strTemp.IndexOf("-") + 1), VbStrConv.ProperCase))
                End If
            Next

            If cmbWeight.Items.Contains(StrConv(strSelection, vbProperCase)) Then
                cmbWeight.SelectedItem = StrConv(strSelection, vbProperCase)
            Else
                cmbWeight.SelectedIndex = 0
            End If
        Catch ex As Exception
            ErrorHandler(System.Reflection.MethodBase.GetCurrentMethod.Name, ex)
        End Try
    End Sub

    Private Function getColorLight(ByVal img As Bitmap) As Color
        Dim image_attr As New ImageAttributes
        Dim cm As ColorMatrix

        Dim bm As New Bitmap(img.Width, img.Height)
        Dim gr As Graphics = Graphics.FromImage(bm)
        Dim rect As Rectangle = _
            Rectangle.Round(img.GetBounds(GraphicsUnit.Pixel))

        cm = New ColorMatrix(New Single()() { _
            New Single() {1.0, 0.0, 0.0, 0.0, 0.0}, _
            New Single() {0.0, 1.0, 0.0, 0.0, 0.0}, _
            New Single() {0.0, 0.0, 1.0, 0.0, 0.0}, _
            New Single() {0.0, 0.0, 0.0, 0.0, 0.0}, _
            New Single() {0.0, 0.0, 0.0, 0.25, 1.0}})
        image_attr.SetColorMatrix(cm)
        gr.DrawImage(img, rect, 0, 0, img.Width, img.Height, _
            GraphicsUnit.Pixel, image_attr)

        Return bm.GetPixel(1, 1)
    End Function

    Private Sub btnLoad_Click(sender As Object, e As EventArgs) Handles btnLoad.Click
        blOffline = False
        Dim loginThread As New System.Threading.Thread(Sub() Me.login(blOffline))
        loginThread.Start()
    End Sub

    Private Sub btnOffline_Click(sender As Object, e As EventArgs) Handles btnOffline.Click
        blOffline = True
        Dim loginThread As New System.Threading.Thread(Sub() Me.login(blOffline))
        loginThread.Start()
    End Sub

    Private Sub login(blOffline As Boolean)
        Try
            If Me.txtEmail.Text.Trim = "" Then
                MsgBox("You must enter an email address/username first.", MsgBoxStyle.Critical, "Invalid Input")
                Exit Sub
            End If
            If Me.WPFPassword.SecurePassword.ToString.Trim = "" And Not blOffline Then
                MsgBox("You must enter a password/sessionid to log in.", MsgBoxStyle.Critical, "Invalid Input")
                Exit Sub
            End If
            AddHandler Model.StashLoading, AddressOf model_StashLoading
            AddHandler Model.Throttled, AddressOf model_Throttled

            Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {btnLoad, "Enabled", False})
            Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {btnOffline, "Enabled", False})
            Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {cmbWeight, "Enabled", False})
            Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {lblWeights, "Enabled", False})
            Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {txtEmail, "Enabled", False})
            Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {lblEmail, "Enabled", False})
            Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {chkSession, "Enabled", False})
            Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {btnEditWeights, "Enabled", False})
            Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {ElementHost1, "Enabled", False})
            Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {lblPassword, "Enabled", False})

            Dim password As SecureString
            If blFormChanged Then
                password = Me.WPFPassword.SecurePassword
            Else
                password = UserSettings("AccountPassword").Decrypt
            End If
            Email = Me.txtEmail.Text
            useSession = chkSession.Checked
            Model.Authenticate(Email, password, blOffline, useSession)
            saveSettings(password)
            If Not blOffline Then
                Model.ForceRefresh()
                statusController.DisplayMessage("Loading characters...")
            Else
                statusController.DisplayMessage("Loading ModRank in offline mode...")
            End If

            Dim chars As List(Of Character)
            Try
                chars = Model.GetCharacters()
            Catch wex As WebException
                Logger.Log(wex)
                statusController.NotOK()
                Throw New Exception("Failed to load characters", wex.InnerException)
            End Try
            statusController.Ok()


            Dim downloadOnlyMyLeagues As Boolean = False
            downloadOnlyMyLeagues = (Settings.UserSettings.ContainsKey("DownloadOnlyMyLeagues") AndAlso Boolean.TryParse(Settings.UserSettings("DownloadOnlyMyLeagues"), downloadOnlyMyLeagues) AndAlso downloadOnlyMyLeagues AndAlso Settings.Lists.ContainsKey("MyLeagues") AndAlso Settings.Lists("MyLeagues").Count > 0)

            statusController.DisplayMessage("Now loading and indexing all the inventory data...")

            invLocation = New Dictionary(Of Item, String)   ' Reinitialize our location list
            Dim TempInventoryAll As New List(Of Item)
            For Each character In chars
                If character.League = "Void" Then
                    Continue For
                End If

                If downloadOnlyMyLeagues AndAlso Not Settings.Lists("MyLeagues").Contains(character.League) Then
                    Continue For
                End If

                Characters.Add(character)

                If Not blOffline Then
                    statusController.DisplayMessage((String.Format("Loading {0}'s inventory...", character.Name)))
                End If
                Dim TempInventory As New List(Of Item)
                TempInventory.AddRange(Model.GetInventory(character.Name))
                TempInventoryAll.AddRange(TempInventory)
                updateStatus(True, blOffline)
                ' Store the character's name for this particular item, for looking up later
                invLocation = invLocation.Concat(TempInventory.ToDictionary(Function(x) x, Function(charname) character.Name)).ToDictionary(Function(x) x.Key, Function(x) x.Value)
                CurrentLeague = character.League
                If Leagues.Contains(character.League) = False Then Leagues.Add(character.League)
            Next

            If Not blOffline Then statusController.DisplayMessage("Done loading character inventory data!")

            ' Get only the gear types out of the temp inventory list (we don't need gems, maps, currency, etc)
            statusCounter = 0 : lngModCounter = 0
            Dim query As IEnumerable(Of Item) = TempInventoryAll.Where(Function(Item) Item.ItemType = ItemType.Gear AndAlso Item.Name.ToLower <> "" AndAlso Item.TypeLine.ToLower.Contains("map") = False)
            AddToFullInventory(query, False)

            For Each league In Leagues
                FullStash = Model.GetStash(league)
                Dim TempStash As New List(Of Item)
                TempStash = FullStash.Get(Of Item)()

                query = TempStash.Where(Function(Item) Item.ItemType = ItemType.Gear AndAlso Item.Name.ToLower <> "" AndAlso Item.TypeLine.ToLower.Contains("map") = False)
                AddToFullInventory(query, True)
            Next

            statusController.DisplayMessage("Completed indexing all " & statusCounter & " stash rare items. comprising a total of approximately " & lngModCounter & " mods.")
            statusController.DisplayMessage("Now loading the data into a table to display the results...please wait, this may take a moment.")

            AddToDataTable(FullInventory, dtRank, True, "Full Inventory Table")
            AddToDataTable(TempInventory, dtOverflow, True, "Overflow Inventory Table")

            Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {DataGridView1, "DataSource", dtRank})
            bytColumns = CByte(Me.Invoke(New RCPD(AddressOf ReadControlProperty), New Object() {DataGridView1.Columns, "Count"}).ToString)
            Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {DataGridView1.Columns(bytColumns - 1), "Visible", False})
            Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {DataGridView1.Columns(bytColumns - 2), "Visible", False})
            Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {DataGridView1.Columns("%").DefaultCellStyle, "Format", "n1"})
            Me.Invoke(New MyDelegate(AddressOf SortDataGridView))
            Me.Invoke(New MyControlDelegate(AddressOf SetDataGridViewWidths), New Object() {DataGridView1})
            Dim FirstCell As DataGridViewCell
            FirstCell = Me.Invoke(New MyDelegateFunction(AddressOf ReturnFirstCell))
            Me.Invoke(New UCPD(AddressOf SetControlProperty), New Object() {DataGridView1, "FirstDisplayedCell", FirstCell})

            Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {Me, "WindowState", FormWindowState.Maximized})
            Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {ElementHost2, "Visible", False})
            Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {DataGridView1, "Visible", True})
            Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {gpLegend, "Left", 550})
            Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {gpLegend, "Visible", True})
            Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {btnEditWeights, "Enabled", True})
            Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {cmbWeight, "Enabled", True})
            Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {lblWeights, "Enabled", True})
            Me.Invoke(New MyDelegate(AddressOf SetDataGridFocus))

            RemoveHandler Model.StashLoading, AddressOf model_StashLoading
            RemoveHandler Model.Throttled, AddressOf model_Throttled

        Catch ex As Exception
            Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {btnLoad, "Enabled", True})
            Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {btnOffline, "Enabled", True})
            Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {cmbWeight, "Enabled", True})
            Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {lblWeights, "Enabled", True})
            Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {txtEmail, "Enabled", True})
            Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {lblEmail, "Enabled", True})
            Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {chkSession, "Enabled", True})
            Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {btnEditWeights, "Enabled", True})
            Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {ElementHost1, "Enabled", True})
            Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {lblPassword, "Enabled", True})
            RemoveHandler Model.StashLoading, AddressOf model_StashLoading
            RemoveHandler Model.Throttled, AddressOf model_Throttled
            ErrorHandler(System.Reflection.MethodBase.GetCurrentMethod.Name, ex)
        End Try
    End Sub

    Public Function ReadControlProperty(ByVal MyControl As Object, ByVal MyProperty As Object) As String
        Dim p As System.Reflection.PropertyInfo = MyControl.GetType().GetProperty(DirectCast(MyProperty, String))
        ReadControlProperty = p.GetValue(MyControl, Nothing).ToString
    End Function

    Private Sub SetControlProperty(ByVal MyControl As Object, ByVal MyProperty As Object, ByVal MyValue As Object)
        Dim p As System.Reflection.PropertyInfo = MyControl.GetType().GetProperty(DirectCast(MyProperty, String))
        p.SetValue(MyControl, MyValue, Nothing)
    End Sub

    Private Sub SetDataGridFocus()
        Me.DataGridView1.Focus()
        DataGridView1.UseWaitCursor = False
    End Sub

    Private Sub DataGridRefresh()
        Me.DataGridView1.Refresh()
    End Sub

    Private Sub SortDataGridView()
        DataGridView1.Sort(DataGridView1.Columns(0), System.ComponentModel.ListSortDirection.Descending)
    End Sub

    Public Sub SetDataGridViewWidths(dg As DataGridView)
        Try
            Dim bytCounter As Byte = 0
            For Each MyWidth In Split(UserSettings("RowWidths"), ",")
                If bytCounter > dg.Columns.Count - 1 Then Exit For
                If IsNumeric(MyWidth) = False Then Continue For
                If MyWidth <= 0 Then
                    dg.Columns(bytCounter).Visible = False
                Else
                    dg.Columns(bytCounter).Width = MyWidth
                End If
                bytCounter += 1
            Next
        Catch ex As Exception
            ErrorHandler(System.Reflection.MethodBase.GetCurrentMethod.Name, ex)
        End Try
    End Sub

    Private Sub loadCharacterInventory(character As Character, offline As Boolean)
        Try
            Dim success As Boolean = False

            If Not offline Then
                statusController.DisplayMessage((String.Format("Loading {0}'s inventory...", character.Name)))
            End If

            Dim inventory As List(Of Item) = Nothing
            Try
                inventory = Model.GetInventory(character.Name)
                success = True
            Catch generatedExceptionName As WebException
                inventory = New List(Of Item)()
                success = False
            End Try

            updateStatus(success, offline)
        Catch ex As Exception
            ErrorHandler(System.Reflection.MethodBase.GetCurrentMethod.Name, ex)
        End Try
    End Sub

    Private Sub updateStatus(success As Boolean, offline As Boolean)
        If offline Then
            Return
        End If

        If success Then
            statusController.Ok()
        Else
            statusController.NotOK()
        End If
    End Sub

    Private Sub model_StashLoading(sender As POEModel, e As StashLoadedEventArgs)
        MyUpdate("Loading " + CurrentLeague & " Stash Tab " & Convert.ToString((e.StashID + 1)) & "...", e)
    End Sub

    Private Sub MyUpdate(message As String, e As POEEventArgs)
        If e.State = POEEventState.BeforeEvent Then
            statusController.DisplayMessage(message)
            Return
        End If

        statusController.Ok()
    End Sub

    Private Sub model_Throttled(sender As Object, e As ThottledEventArgs)
        If e.WaitTime.TotalSeconds > 4 Then
            MyUpdate(String.Format("GGG Server request limit hit, throttling activated. Please wait {0} seconds", e.WaitTime.Seconds), New POEEventArgs(POEEventState.BeforeEvent))
        End If
    End Sub

    Private Sub AddToFullInventory(query As IEnumerable(Of Item), blStash As Boolean)
        Try
            ' Convert the temporary inventory list into our FullInventory list of FullItem class types
            Dim myGear As Gear, lngCount As Long = 0
            For Each myGear In query
                If myGear.Rarity <> Rarity.Rare Then Continue For ' We're only interested in rares
                Dim newFullItem As New FullItem
                newFullItem.ID = myGear.UniqueIDHash
                If blStash = False Then
                    If invLocation.TryGetValue(myGear, newFullItem.Location) = True Then
                        newFullItem.Location = newFullItem.Location    ' The item is in a character's inventory
                    End If
                Else
                    newFullItem.Location = FullStash.GetTabNameByInventoryId(myGear.inventoryId) & " Tab"   ' It's in the stash on tab ?
                End If
                newFullItem.League = myGear.League
                newFullItem.ItemType = myGear.BaseType
                newFullItem.GearType = myGear.GearType.ToString
                If newFullItem.ItemType Is Nothing Or newFullItem.GearType = "Unknown" Then
                    MessageBox.Show("An item (Name: " & myGear.Name & ", Location: " & newFullItem.Location & ") does not have an entry for its type and/or subtype in the APIs. Please report this to:" & Environment.NewLine & Environment.NewLine & _
                                    "https://github.com/RoryTate/modrank/issues" & Environment.NewLine & Environment.NewLine & _
                                    "Also, please provide the actual base type (i.e. Quiver) and subtype (i.e. Light Quiver) that are missing.", "Item Type/Subtype Not Found", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Continue For
                End If
                newFullItem.TypeLine = myGear.TypeLine.ToString
                newFullItem.H = myGear.H
                newFullItem.W = myGear.W
                newFullItem.Name = myGear.Name
                newFullItem.Rarity = myGear.Rarity
                For i = 0 To myGear.Requirements.Count - 1
                    If myGear.Requirements(i).Name.ToLower = "level" Then
                        newFullItem.Level = GetNumeric(myGear.Requirements(i).Value)
                        newFullItem.LevelGem = myGear.Requirements(i).Value.IndexOf("gem", StringComparison.OrdinalIgnoreCase) > -1
                        Exit For ' We don't care about any other requirements, so exit the loop
                    End If
                Next
                newFullItem.Sockets = myGear.NumberOfSockets
                For i = 6 To 0 Step -1
                    If myGear.IsLinked(i) Then
                        newFullItem.Links = IIf(i = 1, 0, i) : Exit For
                    End If
                Next
                newFullItem.Corrupted = myGear.Corrupted
                If Not IsNothing(myGear.Implicitmods) Then
                    For i = 0 To myGear.Implicitmods.Count - 1
                        Dim newFullMod As New FullMod
                        newFullMod.FullText = myGear.Implicitmods(i)
                        If myGear.Implicitmods(i).IndexOf("-", StringComparison.OrdinalIgnoreCase) > -1 Then
                            newFullMod.Value1 = GetNumeric(myGear.Implicitmods(i), 0, myGear.Implicitmods(i).IndexOf("-", StringComparison.OrdinalIgnoreCase))
                            newFullMod.MaxValue1 = GetNumeric(myGear.Implicitmods(i), myGear.Implicitmods(i).IndexOf("-", StringComparison.OrdinalIgnoreCase), myGear.Implicitmods(i).Length)
                        Else
                            newFullMod.Value1 = GetNumeric(myGear.Implicitmods(i))
                        End If
                        newFullMod.Type1 = GetChars(myGear.Implicitmods(i))
                        newFullItem.ImplicitMods.Add(newFullMod)
                    Next
                End If
                ' Make sure that increased item rarity comes last, since it can be either a prefix or a suffix
                ReorderExplicitMods(myGear)
                blAddedOne = False
                If Not IsNothing(myGear.Explicitmods) Then
                    blSolomonsJudgment.Clear()
                    EvaluateExplicitMods(myGear, newFullItem)
                    If newFullItem.ExplicitPrefixMods.Count = 0 And newFullItem.ExplicitSuffixMods.Count = 0 Then
                        If TempInventory.Count = 0 Then
                            EvaluateExplicitMods(myGear, newFullItem, True, True)       ' We didn't find anything, so run again, allowing legacy values to be set
                        Else
                            If TempInventory(TempInventory.Count - 1).Name <> myGear.Name Then
                                EvaluateExplicitMods(myGear, newFullItem, True, True)       ' We didn't find anything, so run again, allowing legacy values to be set
                            End If
                        End If
                    End If
                    If blAddedOne = True Then
                        ' First rename the associated entry in the RankExplanation dictionary...(have to add it in with the proper name and then remove the old one)
                        RankExplanation.Add(myGear.UniqueIDHash & myGear.Name, RankExplanation(myGear.UniqueIDHash & myGear.Name & TempInventory.Count - 1))
                        RankExplanation.Remove(myGear.UniqueIDHash & myGear.Name & TempInventory.Count - 1)
                        FullInventory.Add(TempInventory(TempInventory.Count - 1).Clone)
                        TempInventory.RemoveAt(TempInventory.Count - 1)
                        If lngCount < TempInventory.Count Then
                            FullInventory(FullInventory.Count - 1).OtherSolutions = True
                        End If
                        lngCount = TempInventory.Count
                    Else
                        FullInventory.Add(newFullItem)
                    End If
                End If
                statusCounter += 1 : If statusCounter Mod 100 = 0 Then statusController.DisplayMessage("Indexed " & statusCounter & " rare items and " & lngModCounter & " mods...")
                ' Enable the following to limit the item loading to 100 items when quick development cycles are required
                'If statusCounter = 100 Then Exit Sub
            Next
        Catch ex As Exception
            ErrorHandler(System.Reflection.MethodBase.GetCurrentMethod.Name, ex)
        End Try
    End Sub

    Public Sub ReorderExplicitMods(ByRef myGear As Gear)
        Try
            Dim bytCounter As Byte = myGear.Explicitmods.Count - 1, strType As String = ""
            If Not IsNothing(myGear.Explicitmods) Then
                For i = 0 To myGear.Explicitmods.Count - 1      ' Run the loop to put increased rarity at the end first, since it can be a prefix or suffix, and is the most important (Note: must go to the end of the loop, to decrement bytCounter)
                    strType = GetChars(myGear.Explicitmods(i))
                    If strType.CompareMultiple(StringComparison.OrdinalIgnoreCase, "% increased Rarity of Items found") = True Then
                        Dim strTemp As String = myGear.Explicitmods(bytCounter)
                        myGear.Explicitmods(bytCounter) = myGear.Explicitmods(i)
                        myGear.Explicitmods(i) = strTemp
                        bytCounter -= 1
                        Exit For
                    End If
                Next
                For i = 0 To myGear.Explicitmods.Count - 1      ' Put combined mods -- and mods that get dragged into the combinations -- at the end, so that our looping search goes quicker
                    If i >= bytCounter Then Exit For
                    ' Use a do loop to reorder, since the mod swap might initially exchange a mod at the last spot that is also in this list
                    Do While GetChars(myGear.Explicitmods(i)).CompareMultiple(StringComparison.OrdinalIgnoreCase, "+ to Accuracy Rating", "% increased Block and Stun Recovery", _
                                               "% increased Accuracy Rating", "+ to maximum Mana", "% increased Light Radius", "% increased Armour", _
                                               "% increased Armour and Energy Shield", "% increased Armour and Evasion", "% increased Energy Shield", _
                                               "% increased Evasion and Energy Shield", "% increased Evasion Rating", "% increased Physical Damage", _
                                               "% increased Spell Damage") = True
                        Dim strTemp As String = myGear.Explicitmods(bytCounter)
                        myGear.Explicitmods(bytCounter) = myGear.Explicitmods(i)
                        myGear.Explicitmods(i) = strTemp
                        If i >= bytCounter Then Exit For
                        bytCounter -= 1
                    Loop
                Next
            End If
        Catch ex As Exception
            ErrorHandler(System.Reflection.MethodBase.GetCurrentMethod.Name, ex)
        End Try
    End Sub

    Public Function RunModResultQuery(newFullItem As FullItem, Optional result As DataRow = Nothing, Optional strForceDescription As String = "", Optional strAffix As String = "") As DataRow()
        Try
            Dim strGearType As String = ""
            If newFullItem.GearType.ToString = "Sword" Or newFullItem.GearType.ToString = "Axe" Then
                If newFullItem.H = 4 And newFullItem.W = 2 Then
                    strGearType = "[2h Swords and Axes]"
                Else
                    If newFullItem.TypeLine.ToString.CompareMultiple(StringComparison.OrdinalIgnoreCase, "corroded blade", "longsword", "butcher sword", "headman's sword") Then
                        strGearType = "[2h Swords and Axes]"
                    Else
                        strGearType = "[1h Swords and Axes]"
                    End If
                End If
            ElseIf newFullItem.GearType.ToString = "Mace" Then
                If newFullItem.H = 4 And newFullItem.W = 2 Then strGearType = "[2h Maces]" Else strGearType = "[1h Maces]"
            Else
                strGearType = newFullItem.GearType.ToString
            End If
            ' Note: Level required is 80% of the highest level magic modifier, but the best way to calculate is to add .49 onto the level to make sure we have the maximum,
            ' and then multiply that by 1.25 (or divide by 0.8 if you prefer), and take the rounded result
            ' For example, if the level reqt is 8, then we take 8.49 * 1.25 = 10.6125 = 11  as opposed to incorrectly trying  8 * 1.25 = 10
            Dim strLevel As String = IIf(newFullItem.Level <> 0, "AND Level<=" & Math.Round(1.25 * (newFullItem.Level + 0.49)), "")
            Dim strDescription As String = ""
            If strForceDescription <> "" Then
                strDescription = strForceDescription
            Else
                strDescription = result("Description")
            End If

            Dim strWhere As String = "Description='" & strDescription & "' AND " & strGearType & "=True " & strLevel
            If strAffix <> "" Then strWhere += " AND [Prefix/Suffix]='" & strAffix & "'"
            Dim ModResult() As DataRow = dtMods.Select(strWhere)
            Return ModResult
        Catch ex As Exception
            ErrorHandler(System.Reflection.MethodBase.GetCurrentMethod.Name, ex, "GearType: " & newFullItem.GearType & vbCrLf & "Level: " & newFullItem.Level & vbCrLf & "ForceDesc: " & strForceDescription & vbCrLf & "Affix: " & strAffix)
            Return Nothing
        End Try
    End Function

    Public Function CheckForCombinedMod(result() As DataRow, j As Integer, newFullMod As FullMod, myGear As Gear, i As Integer) As Byte
        ' This function looks to see if the mod entry selected from weights-*.csv is a combined mod, and if successful will return the 
        ' index/position of the other mod from the explicitmods list
        Dim strMod As String = ""
        If newFullMod.Type1 = result(j)("ExportField") Then
            strMod = result(j)("ExportField2")
        Else
            strMod = result(j)("ExportField")
        End If
        For k = i To myGear.Explicitmods.Count - 1  ' Look ahead at upcoming mods to find the position for its "companion" mod
            If GetChars(myGear.Explicitmods(k)) = strMod Then Return k
        Next
        Return 0
    End Function

    Public Sub EvaluateExplicitMods(mygear As Gear, ByRef newfullitem As FullItem, Optional blForceFullSearch As Boolean = False, Optional blAllowLegacy As Boolean = False)
        Try
            Dim result() As DataRow = Nothing
            Dim ModList As New List(Of DataRow), ModPos As New Dictionary(Of String, Byte)
            Dim strField As String = "", strField2 As String = "", strAffix As String = ""
            Dim blCombinedModsAdded As Boolean = False
            For i = 0 To mygear.Explicitmods.Count - 1
                strField = GetChars(mygear.Explicitmods(i))
                ModPos.Add(strField.ToLower, i)
                result = dtWeights.Select("ExportField = '" & strField & "' OR ExportField2 = '" & strField & "'")
                For j = 0 To result.Count - 1
                    If result(j)("ExportField2") <> "" Then ' Do we need to check the second mod for a combined mod?
                        strField2 = IIf(strField = result(j)("ExportField2"), result(j)("ExportField"), result(j)("ExportField2"))
                        If mygear.Explicitmods.Find(Function(x) GetChars(x) = strField2) = "" Then Continue For ' The second mod isn't part of this item...move on
                        blCombinedModsAdded = True  ' If we added a combined mod, will want to run the 'exhaustive' search for more below, otherwise skip the second search
                    End If
                    If strField = "% increased Rarity of Items found" And GetNumeric(mygear.Explicitmods(i)) > 9 Then   ' Rarity has both prefix and suffix values, so add both to the mod list
                        blCombinedModsAdded = True      ' Update: Rather than trying to solve the whole thing here, just set our search to try both and let the recursion routine do the rest
                        Dim tmpResult As DataRow = DeepCopyDataRow(result(j))   ' Cannot do a shallow reference copy, or else we are unable to create separate ModList entries
                        AddToModList(ModList, newfullitem, tmpResult, result(j)("Description") & ",Prefix")     ' We must distinguish our key names to provide the search routine a way to treat this as a "combined" mod
                        Dim tmpResult2 As DataRow = DeepCopyDataRow(result(j))  ' Do it again, as we don't want to affect dtWeights either
                        AddToModList(ModList, newfullitem, tmpResult2, result(j)("Description") & ",Suffix")
                    Else
                        AddToModList(ModList, newfullitem, result(j))
                    End If
                Next
                If result.Count = 0 Then    ' No match, that is strange...
                    MsgBox("Warning: could not find '" & strField & "' mod in weights list. Please check that your weights-*.csv file is properly configured.")
                End If
            Next
            ' Our ModList now contains all of the possible "combined" mods that could have been assigned to this item (we have exhausted all of the branches)
            Dim ModStatsList As New Dictionary(Of String, DataRow)    ' Tranform the ModsList into one that includes the full value ranges and stats for the mod
            Dim MaxIndex As New Dictionary(Of String, Byte)     ' This is used for the more complex search method
            Dim RemoveFromModList As New List(Of Byte)
            For Each row In ModList
                Dim strOverrideDescription As String = ""
                If row("Description").ToString.Contains("Base Item Found Rarity +%") Then
                    'Select Case GetNumeric(mygear.Explicitmods(ModPos(row("ExportField").ToString.ToLower)))
                    '    Case 6 To 9    ' The rarity mod is a prefix
                    '        strAffix = "Prefix"
                    '        'Case 10 To 13      ' It's either a prefix or a suffix
                    '        'Case 14 To 18      ' Now it could be a shared value of both a prefix and a suffix, providing we haven't already hit the max number of them, or it could just be one of either
                    '        'Case 19 To 24       ' For amulets, rings, and helmets, still could be just a prefix or just a suffix, but for boots/gloves it's now a shared value of both
                    '        'Case 27 To 50        ' It must be a shared value of both prefixes and suffixes no matter what geartype
                    '    Case Else
                    '        strAffix = "Both"
                    'End Select
                    If row("Description").ToString.Contains(",") Then
                        strOverrideDescription = row("Description").ToString.Split(",")(0)
                        strAffix = row("Description").ToString.Split(",")(1)
                    End If
                End If
                Dim tempModResult As DataRow() = RunModResultQuery(newfullitem, row, strOverrideDescription, strAffix)
                For j = 0 To tempModResult.Count - 1    ' The query returns all of the possible value ranges for this mod, restricted only by level (max)
                    ModStatsList.Add(row("ExportField").ToString.ToLower & row("ExportField2").ToString.ToLower & j & IIf(strAffix <> "", "," & strAffix, ""), tempModResult(j))
                Next
                If tempModResult.Count = 0 Then
                    ' We will remove this row from the modlist, since it doesn't return any results
                    RemoveFromModList.Add(ModList.IndexOf(row))
                Else
                    MaxIndex.Add(row("ExportField").ToString.ToLower & IIf(strAffix <> "", "," & strAffix, "") & row("ExportField2").ToString.ToLower, tempModResult.Count - 1)
                End If
            Next
            If RemoveFromModList.Count <> 0 Then
                For i = 0 To RemoveFromModList.Count - 1
                    ModList.RemoveAt(RemoveFromModList(i))
                Next
                RemoveFromModList.RemoveRange(0, RemoveFromModList.Count)
            End If
            Dim intRank As Integer = 0
            ' Use simple method if all the mods are independent and a 1-1 mapping exists, so we can set them in a quick, efficient and straightforward manner
            ' Also check the global boolean blForceFullSearch to see if we've called ourselves again because this method won't work
            If (blCombinedModsAdded = False And MaxIndex.Count = ModList.Count) And blForceFullSearch = False Then
                For i = 0 To mygear.Explicitmods.Count - 1
                    lngModCounter += 1
                    Dim newMod As New FullMod
                    newMod.FullText = mygear.Explicitmods(i)
                    newMod.Type1 = GetChars(mygear.Explicitmods(i))
                    If mygear.Explicitmods(i).IndexOf("-", StringComparison.OrdinalIgnoreCase) > -1 Then
                        newMod.Value1 = GetNumeric(mygear.Explicitmods(i), 0, mygear.Explicitmods(i).IndexOf("-", StringComparison.OrdinalIgnoreCase))
                        newMod.MaxValue1 = GetNumeric(mygear.Explicitmods(i), mygear.Explicitmods(i).IndexOf("-", StringComparison.OrdinalIgnoreCase), mygear.Explicitmods(i).Length)
                    Else
                        newMod.Value1 = GetNumeric(mygear.Explicitmods(i))
                    End If
                    newMod.Weight = CInt(ModList(i)("Weight").ToString)
                    'Dim temprow As New Dictionary(Of String, DataRow)
                    Dim temprow = From mymod In ModStatsList Where mymod.Key.ToLower.Contains(newMod.Type1.ToLower) And mymod.Value("MinV") <= newMod.Value1 And mymod.Value("MaxV") >= newMod.Value1
                    If newMod.MaxValue1 > 0 Then
                        temprow = From mymod In ModStatsList Where mymod.Key.ToLower.Contains(newMod.Type1.ToLower) And mymod.Value("MinV") <= newMod.Value1 And mymod.Value("MinV2") >= newMod.Value1 _
                                    And mymod.Value("MaxV2") <= newMod.MaxValue1 And mymod.Value("MaxV") >= newMod.MaxValue1
                    End If
                    If temprow.Count = 0 Then
                        newMod.UnknownValues = True ' This might be a legacy mod that's outside of the ranges in mods.csv
                        strAffix = RunModResultQuery(newfullitem, , ModList(i)("Description"))(0)("Prefix/Suffix")
                        Dim sngMaxV As Single = ModStatsList(newMod.Type1.ToLower & MaxIndex(newMod.Type1.ToLower))("MaxV")
                        RankUnknownValueMod(newMod, sngMaxV, MaxIndex(newMod.Type1.ToLower), mygear.UniqueIDHash, mygear.Name, strAffix)
                        GoTo AddMod
                    End If
                    Dim strKey As String = temprow(temprow.Count - 1).Key.ToLower     ' Choose the last row to get the highest possible level, just like the combined algorithm does
                    newMod.BaseLowerV1 = ModStatsList(strKey)("MinV").ToString
                    If newMod.MaxValue1 <> 0 And newMod.MaxValue1 <> Nothing Then ' This is a mod with range (i.e. 12-16 damage)
                        newMod.BaseLowerMaxV1 = ModStatsList(strKey)("MaxV2").ToString
                        newMod.BaseUpperV1 = ModStatsList(strKey)("MinV2").ToString
                        newMod.BaseUpperMaxV1 = ModStatsList(strKey)("MaxV").ToString
                    Else
                        newMod.BaseUpperV1 = ModStatsList(strKey)("MaxV").ToString
                    End If
                    newMod.MiniLvl = ModStatsList(strKey)("Level").ToString
                    strAffix = ModStatsList(strKey)("Prefix/Suffix").ToString
                    intRank += CalculateRank(newMod, MaxIndex(newMod.Type1.ToLower & newMod.Type2.ToLower), strKey, mygear.UniqueIDHash, mygear.Name, strAffix)
                    Dim sb As New System.Text.StringBuilder
                    sb.Append(vbCrLf & "(" & String.Format("{0:'+'0;'-'0}", intRank) & ") " & IIf(i = mygear.Explicitmods.Count - 1, "Final Rank", " (running total)" & vbCrLf))
                    AddExplanation(mygear.UniqueIDHash & mygear.Name, sb)

AddMod:
                    If newMod.Type1 = "% increased Rarity of Items found" Then  ' If it's rarity then it can be a prefix or a suffix, so see if one of the lists is full
                        If newfullitem.ExplicitPrefixMods.Count >= 4 Then strAffix = "Suffix"
                        If newfullitem.ExplicitSuffixMods.Count >= 4 Then strAffix = "Prefix"
                    End If
                    If strAffix = "Prefix" Then
                        newfullitem.ExplicitPrefixMods.Add(newMod)
                    Else
                        newfullitem.ExplicitSuffixMods.Add(newMod)
                    End If
                Next
                If (newfullitem.ExplicitPrefixMods.Count >= 4 Or newfullitem.ExplicitSuffixMods.Count >= 4) Then
                    ' We've likely tried to fit a combined mod requiring item into our simple algorithm...drop everything and try the "full" method
                    newfullitem.ExplicitPrefixMods.RemoveRange(0, newfullitem.ExplicitPrefixMods.Count) ' Remove any mod list changes we might have made
                    newfullitem.ExplicitSuffixMods.RemoveRange(0, newfullitem.ExplicitSuffixMods.Count)
                    RankExplanation.Remove(mygear.UniqueIDHash & mygear.Name)
                    EvaluateExplicitMods(mygear, newfullitem, True)   ' Call ourselves again, this time with the forcefullsearch boolean set to true
                    Exit Sub
                End If
                newfullitem.Rank = intRank
                newfullitem.Percentile = CalculatePercentile(newfullitem).ToString("0.0")
                Exit Sub
            End If
            ' Note: Combined Mod Search Method begins here!
            ' If this is a combined mod, then we'll have to use a less efficient method.
            Dim maxpos(ModList.Count - 1) As Byte, curpos(ModList.Count - 1) As Integer, bytPos As Byte = 0
            For Each mykey In MaxIndex
                maxpos(bytPos) = mykey.Value
                bytPos += 1
            Next
            DynamicMultiDimLoop(curpos, maxpos, ModList, ModStatsList, 0, mygear, MaxIndex, ModPos, newfullitem, blAllowLegacy)
        Catch ex As Exception
            ErrorHandler(System.Reflection.MethodBase.GetCurrentMethod.Name, ex, "Item Name: " & mygear.Name)
        End Try
    End Sub

    Public Sub AddToModList(ModList As List(Of DataRow), newfullitem As FullItem, MyRow As DataRow, Optional strForceDescription As String = "")
        Try
            If Not ModList.Contains(MyRow) Then ' Before adding it, see if it already is in the list
                Dim tmpResult() As DataRow = RunModResultQuery(newfullitem, MyRow)
                If tmpResult.Count > 0 Then
                    ModList.Add(MyRow)
                    If strForceDescription <> "" Then ModList(ModList.Count - 1)("Description") = strForceDescription
                End If
            End If
        Catch ex As Exception
            ErrorHandler(System.Reflection.MethodBase.GetCurrentMethod.Name, ex, "New Mod Description: " & MyRow("Description") & vbCrLf & _
                         "Current ModList Contents: " & String.Join(vbCrLf, ModList.[Select](Function(r) r("Description").ToString()).ToArray()))
        End Try
    End Sub

    Public Function CalculateRank(newMod As FullMod, bytMaxIndex As Byte, strKey As String, strID As String, strName As String, strAffix As String, Optional strOverrideKey As String = "") As Integer
        Try
            ' Now let's calculate the score/rank...
            Dim sb As New System.Text.StringBuilder
            sb.Append("------Ranked " & strAffix & " mod: " & newMod.FullText)
            ' If the weight isn't already -1 (for mods that don't help), start taking off points for:
            ' - not hitting the highest possible mod level (-11 points per level below the max)
            ' - not getting the maximum value for the mod (-10 points max)...this is calculated based on the range of values possible as follows:
            '   
            ' Note: the above calculation is done with a weight of 5 points max for mods with internal ranges (i.e. Adds Cold Damage, etc)
            ' Note: The maximum that can be lost is normalized to a max of -10, so that low ranked mods with low rng rolls do not send an item unfairly into the negatives
            Dim strKey2 As String = ""
            Dim bytModWeight1 As Byte = 5, bytModWeight2 As Byte = 5    ' The defaults for splitting the weight between mods (in internally ranged mod's, and also between mod 1 and mod 2 -- where internal ranges cannot occur)
            If newMod.Type2.Length = 0 Then
                strKey2 = newMod.Type1.ToLower
                bytModWeight1 = 10 : bytModWeight2 = 0
                If newMod.MaxValue1 <> 0 Then bytModWeight1 = 5 : bytModWeight2 = 5
            Else
                strKey2 = newMod.Type1.ToLower & newMod.Type2.ToLower
                If newMod.BaseUpperV1 - newMod.BaseLowerV1 = 0 Then
                    bytModWeight2 = 10 : bytModWeight1 = 0
                ElseIf newMod.BaseUpperV2 - newMod.BaseLowerV2 = 0 Then
                    bytModWeight1 = 10 : bytModWeight2 = 0
                Else
                    bytModWeight1 = 5 : bytModWeight2 = 5
                End If
            End If
            ' All item mods can score a max of 100 (if the weight is 10 and no other penalties are incurred)
            CalculateRank += newMod.Weight * 10
            sb.Append(vbCrLf & String.Format("{0:'+'0;'-'0}", newMod.Weight * 10) & vbTab & "(mod weight * 10) = (" & newMod.Weight & " * 10) = " & newMod.Weight * 10)

            ' Set the modlevelactual and modlevelmax for this mod
            newMod.ModLevelActual = GetNumeric(strKey)
            newMod.ModLevelMax = bytMaxIndex

            If newMod.Weight < 0 Then GoTo AddExplanationGoto

            Dim intLevelRank As Integer = ((bytMaxIndex - GetNumeric(strKey)) * 10)       ' The amount lost because of not hitting the highest possible level for the mod
            sb.Append(vbCrLf & String.Format("{0:'+'0;'-'0;'-'0}", intLevelRank * -1) & vbTab & "(max mod level - mod level) * 10 = (" & bytMaxIndex & " - " & _
                      GetNumeric(strKey) & ") * 10 = " & intLevelRank)

            Dim intValueRank1 As Integer = 0                                                    ' The amount lost for not getting the maximum value for the first mod
            If newMod.BaseUpperV1 <> newMod.BaseLowerV1 Then
                If newMod.MaxValue1 <> 0 And newMod.BaseUpperMaxV1 <> newMod.BaseLowerMaxV1 Then
                    Dim intValuePart1 As Integer = 0, intValuePart2 As Integer = 0
                    intValuePart1 = Math.Round(((newMod.BaseUpperV1 - newMod.Value1) / (newMod.BaseUpperV1 - newMod.BaseLowerV1)) * bytModWeight1)
                    sb.Append(vbCrLf & String.Format("{0:'+'0;'-'0;'-'0}", intValuePart1 * -1) & vbTab & _
                              (newMod.BaseUpperV1 - newMod.Value1).ToString("0.#") & " below lower max of " & newMod.BaseUpperV1.ToString("0.#") & _
                              " (range: " & (newMod.BaseUpperV1 - newMod.BaseLowerV1).ToString("0.#") & ", weight: " & bytModWeight1.ToString("0.#") & ")")
                    sb.Append(vbCrLf & vbTab & "weight*(max-value)/range = " & bytModWeight1.ToString("0.#") & "*(" & _
                              newMod.BaseUpperV1 & "-" & newMod.Value1.ToString("0.#") & ")/" & (newMod.BaseUpperV1 - newMod.BaseLowerV1).ToString("0.#") & _
                              "=" & intValuePart1.ToString("0.#"))
                    intValuePart2 = Math.Round(((newMod.BaseUpperMaxV1 - newMod.MaxValue1) / (newMod.BaseUpperMaxV1 - newMod.BaseLowerMaxV1)) * bytModWeight2)
                    sb.Append(vbCrLf & String.Format("{0:'+'0;'-'0;'-'0}", intValuePart2 * -1) & vbTab & _
                              newMod.BaseUpperMaxV1 - newMod.MaxValue1 & " below upper max of " & newMod.BaseUpperMaxV1.ToString("0.#") & _
                              " (range: " & newMod.BaseUpperMaxV1 - newMod.BaseLowerMaxV1 & ", weight: " & bytModWeight2 & ")")
                    sb.Append(vbCrLf & vbTab & "weight*(max-value)/range = " & bytModWeight2.ToString("0.#") & "*(" & _
                              newMod.BaseUpperMaxV1 & "-" & newMod.MaxValue1.ToString("0.#") & ")/" & (newMod.BaseUpperMaxV1 - newMod.BaseLowerMaxV1).ToString("0.#") & _
                              "=" & intValuePart2.ToString("0.#"))
                    intValueRank1 = intValuePart1 + intValuePart2
                Else
                    intValueRank1 = Math.Round(((newMod.BaseUpperV1 - newMod.Value1) / (newMod.BaseUpperV1 - newMod.BaseLowerV1)) * bytModWeight1)
                    sb.Append(vbCrLf & String.Format("{0:'+'0;'-'0;'-'0}", intValueRank1 * -1) & vbTab & (newMod.BaseUpperV1 - newMod.Value1).ToString("0.#") & " below max of " & newMod.BaseUpperV1.ToString("0.#") & " (range: " & (newMod.BaseUpperV1 - newMod.BaseLowerV1).ToString("0.#") & ", weight: " & bytModWeight1.ToString("0.#") & ")")
                    sb.Append(vbCrLf & vbTab & "weight*(max-value)/range = " & bytModWeight1.ToString("0.#") & "*(" & newMod.BaseUpperV1.ToString("0.#") & "-" & newMod.Value1.ToString("0.#") & ")/" & (newMod.BaseUpperV1 - newMod.BaseLowerV1).ToString("0.#") & "=" & intValueRank1.ToString("0.#"))
                End If
            Else
                If newMod.MaxValue1 <> 0 And newMod.BaseUpperMaxV1 <> newMod.BaseLowerMaxV1 Then
                    sb.Append(vbCrLf & String.Format("{0:'+'0;'-'0;'-'0}", 0) & vbTab & "Value at max of " & newMod.BaseUpperV1 & " (range: 0)")
                    intValueRank1 = Math.Round(((newMod.BaseUpperMaxV1 - newMod.MaxValue1) / (newMod.BaseUpperMaxV1 - newMod.BaseLowerMaxV1)) * bytModWeight2)
                    sb.Append(vbCrLf & String.Format("{0:'+'0;'-'0;'-'0}", intValueRank1 * -1) & vbTab & _
                              newMod.BaseUpperMaxV1 - newMod.MaxValue1 & " below upper max of " & newMod.BaseUpperMaxV1.ToString("0.#") & _
                              " (range: " & newMod.BaseUpperMaxV1 - newMod.BaseLowerMaxV1 & ", weight: " & bytModWeight2 & ")")
                    sb.Append(vbCrLf & vbTab & "weight*(max-value)/range = " & bytModWeight2.ToString("0.#") & "*(" & _
                              newMod.BaseUpperMaxV1 & "-" & newMod.MaxValue1.ToString("0.#") & ")/" & (newMod.BaseUpperMaxV1 - newMod.BaseLowerMaxV1).ToString("0.#") & _
                              "=" & intValueRank1.ToString("0.#"))
                ElseIf newMod.MaxValue1 = 0 Then
                    intValueRank1 = 0
                    sb.Append(vbCrLf & String.Format("{0:'+'0;'-'0;'-'0}", 0) & vbTab & "Value at max of " & newMod.BaseUpperV1 & " (range: 0)")
                ElseIf newMod.MaxValue1 <> 0 And newMod.BaseUpperMaxV1 = newMod.BaseLowerMaxV1 Then
                    intValueRank1 = 0
                    sb.Append(vbCrLf & String.Format("{0:'+'0;'-'0;'-'0}", 0) & vbTab & "Both lower and upper values at max of " & newMod.BaseUpperV1 & " and " & newMod.BaseUpperMaxV1 & " (range: 0)")
                End If
            End If
            Dim intValueRank2 As Integer = 0                                                    ' The amount lost for not getting the maximum value for the second mod
            If newMod.Type2.Length <> 0 Then
                If newMod.BaseUpperV2 - newMod.BaseLowerV2 <> 0 Then
                    intValueRank2 = Math.Round(((newMod.BaseUpperV2 - newMod.Value2) / (newMod.BaseUpperV2 - newMod.BaseLowerV2)) * bytModWeight2)
                    sb.Append(vbCrLf & String.Format("{0:'+'0;'-'0;'-'0}", intValueRank2 * -1) & vbTab & newMod.BaseUpperV2 - newMod.Value2 & " below max of " & newMod.BaseUpperV2 & " (range: " & newMod.BaseUpperV2 - newMod.BaseLowerV2 & ", weight: " & bytModWeight2 & ")")
                    sb.Append(vbCrLf & vbTab & "weight*(max-value)/range = " & bytModWeight2 & "*(" & newMod.BaseUpperV2 & "-" & newMod.Value2 & ")/" & newMod.BaseUpperV2 - newMod.BaseLowerV2 & "=" & intValueRank2)
                Else
                    intValueRank2 = 0
                    sb.Append(vbCrLf & String.Format("{0:'+'0;'-'0;'-'0}", 0) & vbTab & newMod.BaseUpperV2 - newMod.Value2 & " below max of " & newMod.BaseUpperV2)
                End If
            End If

            CalculateRank -= Math.Min(newMod.Weight * 10 + 10, intLevelRank + intValueRank1 + intValueRank2)
            If newMod.Weight * 10 + 10 < intLevelRank + intValueRank1 + intValueRank2 Then sb.Append(vbCrLf & vbCrLf & "-10 (capped)")

AddExplanationGoto:
            AddExplanation(IIf(strOverrideKey = "", strID & strName, strOverrideKey), sb)
        Catch ex As Exception
            ErrorHandler(System.Reflection.MethodBase.GetCurrentMethod.Name, ex)
            Return 0
        End Try
    End Function

    Public Function CalculatePercentile(MyItem As FullItem) As Single
        Try
            'Formula:
            'To find the percentile rank of a score, x, out of a set of n scores, where x is not included:
            ' 100 * (number of scores below x) / n = percentile rank
            'Where n = number of scores 
            Dim lngNumConfigurationsBelow As Long = 0, lngNumConfigurations As Long = 0, bytNumMods As Byte = 0, sngExtrConfigurations As Single = 0
            For Each myModList In {MyItem.ExplicitPrefixMods, MyItem.ExplicitSuffixMods}
                For Each myMod In myModList
                    Dim result() As DataRow = Nothing
                    result = dtWeights.Select("ExportField = '" & myMod.Type1 & "'" & IIf(myMod.Type2 <> "", " AND ExportField2 = '" & myMod.Type2 & "'", " AND ExportField2 = ''"))
                    If result.Count <> 0 Then
                        For Each row In result
                            For Each modRow In RunModResultQuery(MyItem, row)
                                If myMod.BaseUpperMaxV1 <> 0 Then
                                    If modRow("MinV") < myMod.Value1 Then lngNumConfigurationsBelow += modRow("MinV2") - modRow("MinV") + 1 - IIf(myMod.Value1 <= modRow("MinV2"), modRow("MinV2") - myMod.Value1 + 1, 0)
                                    lngNumConfigurations += modRow("MinV2") - modRow("MinV") + 1
                                    If modRow("MaxV2") < myMod.MaxValue1 Then lngNumConfigurationsBelow += modRow("MaxV2") - modRow("MaxV") + 1 - IIf(myMod.MaxValue1 <= modRow("MaxV2"), modRow("MaxV") - myMod.MaxValue1 + 1, 0)
                                    lngNumConfigurations += modRow("MaxV2") - modRow("MaxV") + 1
                                Else
                                    If modRow("MinV") < myMod.Value1 Then lngNumConfigurationsBelow += modRow("MaxV") - modRow("MinV") + 1 - IIf(myMod.Value1 < modRow("MaxV"), modRow("MaxV") - myMod.Value1 + 1, 0)
                                    lngNumConfigurations += modRow("MaxV") - modRow("MinV") + 1
                                End If
                                If myMod.Value2 <> 0 Then
                                    If modRow("MinV2") < myMod.Value2 Then lngNumConfigurationsBelow += modRow("MaxV2") - modRow("MinV2") + 1 - IIf(myMod.Value2 < modRow("MaxV2"), modRow("MaxV2") - myMod.Value2 + 1, 0)
                                    lngNumConfigurations += modRow("MaxV2") - modRow("MinV2") + 1
                                End If
                            Next
                        Next
                    End If
                    bytNumMods += 1
                Next
            Next
            'If bytNumMods < 6 Then sngExtrConfigurations = (6 - bytNumMods) * (lngNumConfigurations / bytNumMods) ' Use a calculation of the average number of configs for the missing mods
            CalculatePercentile = 100 * ((lngNumConfigurationsBelow)) / (lngNumConfigurations + sngExtrConfigurations)
            CalculatePercentile = Math.Max(1, CalculatePercentile)      ' By definition, a percentile cannot be less than 1%
            CalculatePercentile = Math.Min(99, CalculatePercentile)     ' By definition, a percentile cannot be greater than 99%
        Catch ex As Exception
            ErrorHandler(System.Reflection.MethodBase.GetCurrentMethod.Name, ex)
            Return 0
        End Try
    End Function

    Public Sub AddExplanation(strKey As String, sb As System.Text.StringBuilder)
        If RankExplanation.ContainsKey(strKey) = False Then
            RankExplanation.Add(strKey, sb.ToString)
        Else
            RankExplanation(strKey) += vbCrLf & sb.ToString
        End If
    End Sub

    Public Function RankUnknownValueMod(newMod As FullMod, sngMaxV As Single, bytMaxIndex As Byte, strID As String, strName As String, strAffix As String, Optional strOverridKey As String = "") As Integer
        Try
            Dim sbTemp As New System.Text.StringBuilder
            sbTemp.Append("------Ranked " & strAffix & " mod: " & newMod.FullText)
            If newMod.Weight > 0 Then
                If newMod.Value1 > sngMaxV Then
                    RankUnknownValueMod += newMod.Weight * 11      ' Our value is above the max for this level, so set speakers to 11!
                    sbTemp.Append(vbCrLf & String.Format("{0:'+'0;'-'0}", newMod.Weight * 11) & vbTab & "(mod weight * 11) = (" & newMod.Weight & " * 11) = " & newMod.Weight * 11)
                    sbTemp.Append(vbCrLf & vbTab & "Values above maximum range, so setting speakers to 11")
                Else
                    RankUnknownValueMod += newMod.Weight * 10   ' The value is below the min for the lowest level, so set weight as normal, but punish for value/level
                    sbTemp.Append(vbCrLf & String.Format("{0:'+'0;'-'0}", newMod.Weight * 10) & vbTab & "(mod weight * 10) = (" & newMod.Weight & " * 10) = " & newMod.Weight * 10)
                    Dim intLevelRank As Integer = ((bytMaxIndex + 1) * 10)       ' The amount lost because of not hitting the highest possible level for the mod
                    sbTemp.Append(vbCrLf & String.Format("{0:'+'0;'-'0;'-'0}", intLevelRank * -1) & vbTab & "(max mod level + 1) * 10 = (" & bytMaxIndex + 1 & " * 10) = " & intLevelRank)
                    RankUnknownValueMod -= Math.Min(newMod.Weight * 10 + 10, intLevelRank)
                    If newMod.Weight * 10 + 10 < intLevelRank Then sbTemp.Append(vbCrLf & vbCrLf & "-10 (capped)")
                End If
            Else
                RankUnknownValueMod += newMod.Weight * 10
                sbTemp.Append(vbCrLf & String.Format("{0:'+'0;'-'0}", newMod.Weight * 10) & vbTab & "(mod weight * 10) = (" & newMod.Weight & " * 10) = " & newMod.Weight * 10)
            End If
            sbTemp.Append(vbCrLf & vbCrLf & "(" & String.Format("{0:'+'0;'-'0}", RankUnknownValueMod) & ") (running total)" & vbCrLf)
            AddExplanation(IIf(strOverridKey = "", strID & strName, strOverridKey), sbTemp)
        Catch ex As Exception
            ErrorHandler(System.Reflection.MethodBase.GetCurrentMethod.Name, ex)
            Return 0
        End Try
    End Function

    Public Sub DynamicMultiDimLoop(ByRef curpos() As Integer, ByRef maxpos() As Byte, ByRef ModList As List(Of DataRow), ModStatsList As Dictionary(Of String, DataRow), curdim As Byte, myGear As Gear, MaxIndex As Dictionary(Of String, Byte), modPos As Dictionary(Of String, Byte), ByRef newFullItem As FullItem, blAllowLegacy As Boolean)
        Try
            For i = maxpos(curdim) To -1 Step -1    ' We go to -1, since -1 is the index where we don't use the potential mod
                curpos(curdim) = i
                Dim strStatsKey As String = ""
                If i >= 0 Then
                    strStatsKey = ModList(curdim)("ExportField").ToString.ToLower & ModList(curdim)("ExportField2").ToString.ToLower & i
                    If ModStatsList.ContainsKey(strStatsKey) = False Then   ' It must be a rarity mod that could be a shared total of both a prefix and a suffix
                        strStatsKey += "," & ModList(curdim)("Description").ToString.Split(",")(1)
                    End If
                End If
                ' If the 'min' value for this mod is above the actual value, then skip it and go down a level in order to trim our search space
                ' If the 'max' value for this mod is below the actual value, AND it is not a combined mod (meaning there is no other way to satisfy the value), then go up a level and trim some more
                ' The i>0 condition means that we will try at least once, so that we will set an "UnknownValue=True" if we're having trouble with a legacy mod/value
                If ModList(curdim)("ExportField2") = "" And i > 0 Then
                    Dim bytPos As Byte = modPos(ModList(curdim)("ExportField").ToString.ToLower)
                    If myGear.Explicitmods(bytPos).IndexOf("-", StringComparison.OrdinalIgnoreCase) > -1 Then   ' We have to check a ranged value for both minimums
                        If GetNumeric(myGear.Explicitmods(bytPos), 0, myGear.Explicitmods(bytPos).IndexOf("-", StringComparison.OrdinalIgnoreCase)) < ModStatsList(strStatsKey)("MinV") Then Continue For
                        If GetNumeric(myGear.Explicitmods(bytPos), myGear.Explicitmods(bytPos).IndexOf("-", StringComparison.OrdinalIgnoreCase), myGear.Explicitmods(bytPos).Length) < ModStatsList(strStatsKey)("MaxV2") Then Continue For
                        If GetNumeric(myGear.Explicitmods(bytPos), 0, myGear.Explicitmods(bytPos).IndexOf("-", StringComparison.OrdinalIgnoreCase)) > ModStatsList(strStatsKey)("MinV2") And blAddedOne = True Then Exit For
                        If GetNumeric(myGear.Explicitmods(bytPos), myGear.Explicitmods(bytPos).IndexOf("-", StringComparison.OrdinalIgnoreCase), myGear.Explicitmods(bytPos).Length) > ModStatsList(strStatsKey)("MaxV") And blAddedOne = True Then Exit For
                    Else
                        If GetNumeric(myGear.Explicitmods(bytPos)) < ModStatsList(strStatsKey)("MinV") Then Continue For
                        If HaveToShare(GetChars(myGear.Explicitmods(bytPos)).ToLower, ModList, curpos) = False Then
                            If GetNumeric(myGear.Explicitmods(bytPos)) > ModStatsList(strStatsKey)("MaxV") And blAddedOne = True Then Exit For
                        End If
                    End If
                ElseIf i > 0 Then
                    Dim bytPos As Byte = modPos(ModList(curdim)("ExportField").ToString.ToLower)
                    Dim bytPos2 As Byte = modPos(ModList(curdim)("ExportField2").ToString.ToLower)
                    If GetNumeric(myGear.Explicitmods(bytPos)) < ModStatsList(strStatsKey)("MinV") Then Continue For
                    If GetNumeric(myGear.Explicitmods(bytPos2)) < ModStatsList(strStatsKey)("MinV2") Then Continue For
                End If

                If curdim < UBound(maxpos) Then
                    ' Recursively call this sub again to repeat the loop for the next dimension of the list/array
                    DynamicMultiDimLoop(curpos, maxpos, ModList, ModStatsList, curdim + 1, myGear, MaxIndex, modPos, newFullItem, blAllowLegacy)
                Else
                    ' We have a possible match!
                    Dim bytPos(MaxIndex.Count - 1, 1) As Byte, blmatch As Boolean = True
                    Dim ModMinTotals(myGear.Explicitmods.Count - 1) As Single, ModMaxTotals(myGear.Explicitmods.Count - 1) As Single
                    For j = 0 To MaxIndex.Count - 1
                        If curpos(j) = -1 Then Continue For ' If position is -1, we are trying to get the proper totals without this mod
                        Dim strTempStatKey As String = ModList(j)("ExportField").ToString.ToLower & ModList(j)("ExportField2").ToString.ToLower & curpos(j)
                        If ModList(j)("Description").ToString.Contains(",") = True Then
                            strTempStatKey += "," & ModList(j)("Description").ToString.Split(",")(1)
                        End If
                        bytPos(j, 0) = modPos(ModList(j)("ExportField").ToString.ToLower)
                        ModMinTotals(bytPos(j, 0)) += ModStatsList(strTempStatKey)("MinV")
                        ModMaxTotals(bytPos(j, 0)) += ModStatsList(strTempStatKey)("MaxV")
                        If ModList(j)("ExportField2") <> "" Then
                            bytPos(j, 1) = modPos(ModList(j)("ExportField2").ToString.ToLower)
                            ModMinTotals(bytPos(j, 1)) += ModStatsList(strTempStatKey)("MinV2")
                            ModMaxTotals(bytPos(j, 1)) += ModStatsList(strTempStatKey)("MaxV2")
                        End If
                    Next
                    For j = 0 To ModMinTotals.Count - 1
                        'If curpos(j) = -1 Then Continue For ' If position is -1, we're getting the totals without this mod
                        If myGear.Explicitmods(j).IndexOf("-", StringComparison.OrdinalIgnoreCase) > -1 Then
                            If ModMinTotals(j) > GetNumeric(myGear.Explicitmods(j), 0, myGear.Explicitmods(j).IndexOf("-", StringComparison.OrdinalIgnoreCase)) Then blmatch = False : Exit For
                            If ModMaxTotals(j) < GetNumeric(myGear.Explicitmods(j), myGear.Explicitmods(j).IndexOf("-", StringComparison.OrdinalIgnoreCase), myGear.Explicitmods(j).Length) Then blmatch = False : Exit For
                        Else
                            If ModMinTotals(j) > GetNumeric(myGear.Explicitmods(j)) Then blmatch = False : Exit For
                            If ModMaxTotals(j) < GetNumeric(myGear.Explicitmods(j)) Then
                                If HaveToShare(GetChars(myGear.Explicitmods(j)).ToLower, ModList, curpos) = False Then
                                    If blAddedOne = True Or blAllowLegacy = False Then blmatch = False : Exit For
                                    ' Initially, when blAllowLegacy is false, we want to try everything until we know we have a match, at which point we can be more picky
                                    ' If blAddedOne is false and we're running again with AllowLegacy to true, we have failed to find a maximum large enough for a non-combined mod
                                    ' (likely due to a legacy value), so bypass the value testing for this mod, the blMatch will stay as true, and we'll set it as an "Unknown Value" mod
                                Else
                                    blmatch = False : Exit For
                                End If
                            End If
                        End If
                    Next
                    If blmatch = True Then
                        Dim intRank As Integer = 0
                        Dim strRankExplanationKey As String = myGear.UniqueIDHash & myGear.Name & TempInventory.Count
                        For j = 0 To ModList.Count - 1
                            If curpos(j) = -1 Then Continue For ' If position is -1, this mod position in ModList will not be used
                            Dim newMod As New FullMod, strAffix As String = ""
                            newMod.FullText = myGear.Explicitmods(bytPos(j, 0))
                            newMod.Type1 = GetChars(myGear.Explicitmods(bytPos(j, 0)))
                            newMod.Weight = CInt(ModList(j)("Weight").ToString)
                            Dim strKey As String = "", strMaxIndexKey As String = ""
                            If ModList(j)("Description").ToString.Contains(",") Then
                                strKey = ModList(j)("ExportField").tolower & ModList(j)("ExportField2").tolower & curpos(j) & "," & ModList(j)("Description").ToString.Split(",")(1)
                                strMaxIndexKey = ModList(j)("ExportField").tolower & ModList(j)("ExportField2").tolower & "," & ModList(j)("Description").ToString.Split(",")(1)
                            Else
                                strKey = ModList(j)("ExportField").tolower & ModList(j)("ExportField2").tolower & curpos(j)
                                strMaxIndexKey = ModList(j)("ExportField").tolower & ModList(j)("ExportField2").tolower
                            End If

                            If myGear.Explicitmods(bytPos(j, 0)).IndexOf("-", StringComparison.OrdinalIgnoreCase) > -1 Then
                                newMod.Value1 = GetNumeric(myGear.Explicitmods(bytPos(j, 0)), 0, myGear.Explicitmods(bytPos(j, 0)).IndexOf("-", StringComparison.OrdinalIgnoreCase))
                                newMod.MaxValue1 = GetNumeric(myGear.Explicitmods(bytPos(j, 0)), myGear.Explicitmods(bytPos(j, 0)).IndexOf("-", StringComparison.OrdinalIgnoreCase), myGear.Explicitmods(bytPos(j, 0)).Length)
                            Else
                                If HaveToShare(ModList(j)("ExportField").tolower, ModList, curpos) = True Then
                                    ' We only have a portion of this value...have to apply a weighting formula to determine how much is ours and how much the other mod(s) will chip in
                                    newMod.Value1 = DistributeValues(strKey, ModList(j)("ExportField"), GetNumeric(myGear.Explicitmods(bytPos(j, 0))), ModList, ModStatsList, curpos)
                                    ' Have to change the fulltext to reflect the new value
                                    newMod.FullText = BuildFullText(newMod.Type1, newMod.Value1)
                                Else
                                    newMod.Value1 = GetNumeric(myGear.Explicitmods(bytPos(j, 0)))
                                    If GetNumeric(myGear.Explicitmods(bytPos(j, 0))) > ModStatsList(strKey)("MaxV").ToString And blAllowLegacy = True Then
                                        newMod.UnknownValues = True ' This might be a legacy mod that's outside of the ranges in mods.csv
                                        strAffix = RunModResultQuery(newFullItem, , ModList(j)("Description"))(0)("Prefix/Suffix")
                                        Dim sngMaxV As Single = ModStatsList(newMod.Type1.ToLower & MaxIndex(newMod.Type1.ToLower))("MaxV")
                                        intRank += RankUnknownValueMod(newMod, sngMaxV, MaxIndex(strMaxIndexKey), myGear.UniqueIDHash, myGear.Name, strAffix, strRankExplanationKey)
                                        ' Don't jump ahead just yet, we might have an ExportField2 to set
                                    End If
                                End If
                            End If
                            If ModList(j)("ExportField2") <> "" Then
                                newMod.Type2 = GetChars(myGear.Explicitmods(bytPos(j, 1)))
                                If HaveToShare(ModList(j)("ExportField2").tolower, ModList, curpos) = True Then
                                    ' We only have a portion of this value...have to apply a weighting formula to determine how much is ours and how much the other mod(s) will chip in
                                    newMod.Value2 = DistributeValues(strKey, ModList(j)("ExportField2"), GetNumeric(myGear.Explicitmods(bytPos(j, 1))), ModList, ModStatsList, curpos)
                                    If newMod.FullText.IndexOf("/") = -1 Then
                                        newMod.FullText += "/" & BuildFullText(newMod.Type2, newMod.Value2)
                                    Else
                                        ' Have to rebuild the mod values, since they are now shared
                                        newMod.FullText = newMod.FullText.Substring(0, newMod.FullText.IndexOf("/")) & BuildFullText(newMod.Type2, newMod.Value2)
                                    End If
                                Else
                                    newMod.Value2 = GetNumeric(myGear.Explicitmods(bytPos(j, 1)))
                                    ' This second mod may not be part of the original mod text
                                    If newMod.FullText.IndexOf("/") = -1 Then newMod.FullText += "/" & BuildFullText(newMod.Type2, newMod.Value2)
                                End If
                            End If
                            If newMod.UnknownValues = True Then GoTo AddMod2

                            newMod.BaseLowerV1 = ModStatsList(strKey)("MinV").ToString
                            If newMod.MaxValue1 <> 0 And newMod.MaxValue1 <> Nothing Then ' This is a mod with range (i.e. 12-16 damage)
                                newMod.BaseLowerMaxV1 = ModStatsList(strKey)("MaxV2").ToString
                                newMod.BaseUpperV1 = ModStatsList(strKey)("MinV2").ToString
                                newMod.BaseUpperMaxV1 = ModStatsList(strKey)("MaxV").ToString
                            Else
                                newMod.BaseUpperV1 = ModStatsList(strKey)("MaxV").ToString
                            End If
                            If ModList(j)("ExportField2") <> "" Then
                                newMod.BaseLowerV2 = ModStatsList(strKey)("MinV2").ToString
                                newMod.BaseUpperV2 = ModStatsList(strKey)("MaxV2").ToString
                            End If
                            newMod.MiniLvl = ModStatsList(strKey)("Level").ToString
                            strAffix = ModStatsList(strKey)("Prefix/Suffix").ToString
                            intRank += CalculateRank(newMod, MaxIndex(strMaxIndexKey), strKey, myGear.UniqueIDHash, myGear.Name, strAffix, strRankExplanationKey)
                            Dim sb As New System.Text.StringBuilder
                            sb.Append(vbCrLf & "(" & String.Format("{0:'+'0;'-'0}", intRank) & ")  (running total)" & vbCrLf)
                            RankExplanation(strRankExplanationKey) += vbCrLf & sb.ToString

AddMod2:
                            If strAffix = "Prefix" Then
                                newFullItem.ExplicitPrefixMods.Add(newMod)
                            Else
                                newFullItem.ExplicitSuffixMods.Add(newMod)
                            End If
                        Next
                        If (newFullItem.ExplicitPrefixMods.Count >= 4 Or newFullItem.ExplicitSuffixMods.Count >= 4) Then
                            ' We've likely tried to add a combined mod in a way that isn't allowed because of the other mods, which means this is not a possible configuration, so remove
                            newFullItem.ExplicitPrefixMods.RemoveRange(0, newFullItem.ExplicitPrefixMods.Count) ' Remove any mod list changes we might have made
                            newFullItem.ExplicitSuffixMods.RemoveRange(0, newFullItem.ExplicitSuffixMods.Count)
                            ' We'll have added stuff to the rankexplanation dictionary, so clean the entry associated with this item
                            RankExplanation.Remove(strRankExplanationKey)
                            Continue For
                        End If
                        newFullItem.Rank = intRank
                        newFullItem.Percentile = CalculatePercentile(newFullItem).ToString("0.0")
                        RankExplanation(strRankExplanationKey) = RankExplanation(strRankExplanationKey).Substring(0, RankExplanation(strRankExplanationKey).LastIndexOf("(running total)")) & "Final Rank" & vbCrLf
                        If TempInventory.IndexOf(newFullItem) = -1 Then
                            Dim tmpItem As FullItem = newFullItem.Clone
                            TempInventory.Add(tmpItem)
                        End If
                        newFullItem.ExplicitPrefixMods.RemoveRange(0, newFullItem.ExplicitPrefixMods.Count) ' Remove any mod list changes we might have made
                        newFullItem.ExplicitSuffixMods.RemoveRange(0, newFullItem.ExplicitSuffixMods.Count)
                        blAddedOne = True
                    End If
                End If
                'If blAddedOne = True Then Exit Sub
            Next
        Catch ex As Exception
            ErrorHandler(System.Reflection.MethodBase.GetCurrentMethod.Name, ex)
        End Try
    End Sub

    Public Function HaveToShare(strMyMod As String, ModList As List(Of DataRow), curpos() As Integer) As Boolean
        ' Look in ModList to see if there are multiple entries for this mod, but first check curpos to see if the entry is even in consideration for this particular iteration
        Dim blFound As Boolean = False
        For i = 0 To ModList.Count - 1
            If curpos(i) = -1 Then Continue For
            Dim strKey As String = ModList(i)("ExportField").tolower & ModList(i)("ExportField2").tolower
            If ModList(i)("ExportField").tolower = strMyMod.ToLower Or ModList(i)("ExportField2").tolower = strMyMod.ToLower Then
                If blFound = False Then
                    blFound = True
                Else
                    Return True
                End If
            End If
        Next
        HaveToShare = False
    End Function

    Public Function DistributeValues(strMyKey As String, strMod As String, sngValue As Single, ModList As List(Of DataRow), ModStatsList As Dictionary(Of String, DataRow), curpos() As Integer) As Single
        ' Apply a weighting formula to determine how much of sngValue belongs in the strMyKey mod and how much the other mod(s) will chip in
        Dim sngRange As Single = 0, sngTotalRange As Single = 0
        Dim sngMax As Single = 0, sngTotalMax As Single = 0
        For i = 0 To ModList.Count - 1
            If curpos(i) = -1 Then Continue For
            Dim strKey As String = ""
            If ModList(i)("Description").ToString.Contains(",") Then
                strKey = ModList(i)("ExportField").tolower & ModList(i)("ExportField2").tolower & curpos(i) & "," & ModList(i)("Description").ToString.Split(",")(1)
            Else
                strKey = ModList(i)("ExportField").tolower & ModList(i)("ExportField2").tolower & curpos(i)
            End If

            If ModList(i)("ExportField").tolower = strMod.ToLower Then
                If strMyKey.ToLower <> strKey.ToLower Then
                    sngTotalRange += ModStatsList(strKey)("MaxV").ToString - ModStatsList(strKey)("MinV").ToString
                    sngTotalMax += ModStatsList(strKey)("MaxV").ToString
                Else
                    sngRange = ModStatsList(strKey)("MaxV").ToString - ModStatsList(strKey)("MinV").ToString
                    sngMax = ModStatsList(strKey)("MaxV").ToString
                    sngTotalRange += sngRange
                    sngTotalMax += sngMax
                End If
            ElseIf ModList(i)("ExportField2").tolower = strMod.ToLower Then
                If strMyKey.ToLower <> strKey.ToLower Then
                    sngTotalRange += ModStatsList(strKey)("MaxV2").ToString - ModStatsList(strKey)("MinV2").ToString
                    sngTotalMax += ModStatsList(strKey)("MaxV2").ToString()
                Else
                    sngRange = ModStatsList(strKey)("MaxV2").ToString - ModStatsList(strKey)("MinV2").ToString
                    sngMax = ModStatsList(strKey)("MaxV2").ToString
                    sngTotalRange += sngRange
                    sngTotalMax += sngMax
                End If
            End If
        Next
        Dim sngCalculatedValue As Single = sngMax - (sngRange * (sngTotalMax - sngValue) / sngTotalRange)
        If sngCalculatedValue - Math.Truncate(sngCalculatedValue) = 0.5 Then
            If blSolomonsJudgment.Count = 0 Then
                blSolomonsJudgment.Add(strMod, True)
                sngCalculatedValue = Math.Truncate(sngCalculatedValue)
            Else
                sngCalculatedValue = Math.Ceiling(sngCalculatedValue)
            End If
        Else
            sngCalculatedValue = Math.Round(sngCalculatedValue)
        End If
        DistributeValues = sngCalculatedValue
    End Function

    Public Function BuildFullText(strType As String, sngValue As Single) As String
        ' When we have to share values for combined mods, we will need to build a full mod text string ourselves, since we don't match the mod text in the original item
        ' This function is a simple way to tell where to put the value in the string depending on the mod text (i.e. after the +, or before the %, etc)
        If strType.IndexOf("+") <> -1 Then Return strType.Substring(0, strType.IndexOf("+") + 1) & sngValue & strType.Substring(strType.IndexOf("+") + 1)
        If strType.IndexOf("%") <> -1 Then Return strType.Substring(0, strType.IndexOf("%")) & sngValue & strType.Substring(strType.IndexOf("%"))
        If strType.IndexOf("Adds") <> -1 Then Return strType.Substring(0, strType.IndexOf("Adds") + 5) & sngValue & strType.Substring(strType.IndexOf("Adds") + 4)
        If strType.IndexOf("Reflects") <> -1 Then Return strType.Substring(0, strType.IndexOf("Reflects") + 9) & sngValue & strType.Substring(strType.IndexOf("Reflects") + 8)
        Return sngValue & " " & strType
    End Function

    Public Sub AddToDataTable(myList As List(Of FullItem), dt As DataTable, blShowProgress As Boolean, strDisplayText As String)
        Try
            Dim intCounter As Integer = 0
            For Each it In myList
                Dim row As DataRow = dt.NewRow()
                row("Rank") = it.Rank
                row("%") = it.Percentile
                row("Type") = If(it.GearType, DBNull.Value)
                row("SubType") = If(it.ItemType, DBNull.Value)
                row("Leag") = If(it.League, DBNull.Value)
                row("Location") = If(it.Location, DBNull.Value)
                row("Name") = If(it.Name, DBNull.Value)
                row("Level") = it.Level
                row("Gem") = it.LevelGem
                row("Sokt") = it.Sockets
                row("Link") = it.Links
                If it.ExplicitPrefixMods Is Nothing = False Then
                    For i = 0 To it.ExplicitPrefixMods.Count - 1
                        row("Prefix " & i + 1) = it.ExplicitPrefixMods(i).FullText
                    Next
                End If
                If it.ExplicitSuffixMods Is Nothing = False Then
                    For i = 0 To it.ExplicitSuffixMods.Count - 1
                        row("Suffix " & i + 1) = it.ExplicitSuffixMods(i).FullText
                    Next
                End If
                If it.ImplicitMods Is Nothing = False Then
                    For i = 0 To it.ImplicitMods.Count - 1
                        row("Implicit") = row("Implicit") & it.ImplicitMods(i).FullText & ", "
                    Next
                    row("Implicit") = row("Implicit").ToString.Substring(0, Math.Max(2, row("Implicit").ToString.Length) - 2)
                End If
                row("*") = If(it.OtherSolutions, "*", "")
                row("Corrupt") = it.Corrupted
                row("Index") = myList.IndexOf(it)
                row("ID") = it.ID
                dt.Rows.Add(row)
                intCounter += 1
                If intCounter Mod 100 = 0 And blShowProgress Then statusController.DisplayMessage(strDisplayText & ": Completed adding " & intCounter & " of " & myList.Count & " rows.")
            Next
            If blShowProgress Then statusController.DisplayMessage(strDisplayText & ": Completed adding all " & myList.Count & " rows.")
        Catch ex As Exception
            ErrorHandler(System.Reflection.MethodBase.GetCurrentMethod.Name, ex, "Datatable: " & dt.TableName)
        End Try
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.RowIndex = -1 Then Exit Sub
        If e.ColumnIndex = DataGridView1.Columns("Rank").Index Then
            Dim sb As New System.Text.StringBuilder
            If FullInventory(DataGridView1.Rows(e.RowIndex).Cells("Index").Value).LevelGem = True Then AddGemWarning(sb)
            sb.Append(RankExplanation(DataGridView1.CurrentRow.Cells("ID").Value & DataGridView1.CurrentRow.Cells("Name").Value))
            MsgBox(sb.ToString, , "Item Mod Rank Explanation - " & DataGridView1.CurrentRow.Cells("Name").Value)
        ElseIf e.ColumnIndex = DataGridView1.Columns("*").Index AndAlso DataGridView1.CurrentRow.Cells("*").Value = "*" Then
            If Application.OpenForms().OfType(Of frmResults).Any = False Then frmResults.Show(Me)
            frmResults.Text = "Possible Mod Solutions for '" & DataGridView1.CurrentRow.Cells("Name").Value & "'"
            Dim tmpList As New CloneableList(Of String)
            tmpList.Add(DataGridView1.CurrentRow.Cells("ID").Value)
            tmpList.Add(DataGridView1.CurrentRow.Cells("Name").Value)
            frmResults.MyData = tmpList
        ElseIf DataGridView1.Columns(e.ColumnIndex).Name.ToLower.Contains("prefix") AndAlso NotNull(DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value, "") <> "" Then
            ShowModInfo(DataGridView1, FullInventory(DataGridView1.Rows(e.RowIndex).Cells("Index").Value), FullInventory(DataGridView1.Rows(e.RowIndex).Cells("Index").Value).ExplicitPrefixMods, GetNumeric(DataGridView1.Columns(e.ColumnIndex).Name) - 1, e)
        ElseIf DataGridView1.Columns(e.ColumnIndex).Name.ToLower.Contains("suffix") AndAlso NotNull(DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value, "") <> "" Then
            ShowModInfo(DataGridView1, FullInventory(DataGridView1.Rows(e.RowIndex).Cells("Index").Value), FullInventory(DataGridView1.Rows(e.RowIndex).Cells("Index").Value).ExplicitSuffixMods, GetNumeric(DataGridView1.Columns(e.ColumnIndex).Name) - 1, e)
        End If
    End Sub

    Public Sub ShowModInfo(dg As DataGridView, MyItem As FullItem, MyMod As List(Of FullMod), bytIndex As Byte, e As DataGridViewCellEventArgs)
        Dim sb As New System.Text.StringBuilder
        If MyItem.LevelGem = True Then AddGemWarning(sb)
        sb.Append("Mod Text: " & MyMod(bytIndex).FullText & vbCrLf)
        sb.Append("Mod Weight: " & MyMod(bytIndex).Weight & vbCrLf)
        Dim result() As DataRow = Nothing
        result = dtWeights.Select("ExportField = '" & MyMod(bytIndex).Type1 & "'" & IIf(MyMod(bytIndex).Type2 <> "", " AND ExportField2 = '" & MyMod(bytIndex).Type2 & "'", " AND ExportField2 = ''"))
        sb.Append(vbCrLf & "Possible Mod Levels and Values:" & vbCrLf)
        Dim strAffix As String = dg.Columns(e.ColumnIndex).Name.Substring(0, 6)
        Dim blAddedTopRow As Boolean = False
        If result.Count <> 0 Then
            For Each row In result
                Dim bytModLevel As Byte = 0
                For Each ModRow In RunModResultQuery(MyItem, row, , strAffix)
                    If blAddedTopRow = False Then
                        sb.Append("----------" & vbTab & "------------" & vbCrLf)
                        sb.Append("| Level" & vbTab & "| Values" & vbCrLf)
                        sb.Append("----------" & vbTab & "------------" & vbCrLf)
                        blAddedTopRow = True
                    End If
                    If ModRow("Description") <> row("Description") Then Continue For
                    strAffix = ModRow("Prefix/Suffix")
                    sb.Append("| " & ModRow("Level") & IIf(MyMod(bytIndex).ModLevelActual = bytModLevel And MyMod(bytIndex).UnknownValues = False, "(*)", "") & vbTab & "| " & ModRow("Value") & vbCrLf)
                    bytModLevel += 1
                Next
            Next
            sb.Append("----------" & vbTab & "------------" & vbCrLf & vbCrLf)
        End If
        sb.Append(IIf(MyMod(bytIndex).UnknownValues, "Note: the value(s) for this mod are beyond the possible ranges listed...this is likely a 'legacy' mod value for which the level cannot be indicated.", "Current level is indicated by (*)."))
        MsgBox(sb.ToString, , "Mod Info for " & strAffix & " '" & MyMod(bytIndex).FullText & "'")
    End Sub

    Public Sub AddGemWarning(ByRef sb As System.Text.StringBuilder)
        sb.Append("**WARNING!!** Item level requirements set by an equipped gem...mod level ranges are likely too large, and the ranking will be artificially lowered as a result. It is recommended that the gem be removed before evaluating the item in this manner." & vbCrLf & vbCrLf)
    End Sub

    Private Sub DataGridView1_CellMouseEnter(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellMouseEnter
        If blScroll = True Then Exit Sub
        If IsValidCellAddress(DataGridView1, e.RowIndex, e.ColumnIndex) AndAlso _
            (e.ColumnIndex = 0 Or DataGridView1.Columns(e.ColumnIndex).Name.Contains("fix") Or e.ColumnIndex = 13) Then DataGridView1.Cursor = Cursors.Hand
    End Sub

    Public Function IsValidCellAddress(dg As DataGridView, rowIndex As Integer, columnIndex As Integer) As Boolean
        Return rowIndex >= 0 AndAlso rowIndex < dg.RowCount AndAlso columnIndex >= 0 AndAlso columnIndex <= dg.ColumnCount
    End Function

    Private Sub DataGridView1_CellMouseLeave(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellMouseLeave
        If blScroll = True Then Exit Sub
        If IsValidCellAddress(DataGridView1, e.RowIndex, e.ColumnIndex) AndAlso _
            (e.ColumnIndex = 0 Or DataGridView1.Columns(e.ColumnIndex).Name.Contains("fix") Or e.ColumnIndex = 13) Then DataGridView1.Cursor = Cursors.Default
    End Sub

    Private Sub DataGridView1_CellMouseMove(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseMove
        blScroll = False
    End Sub

    Private Sub DataGridView1_CellPainting(sender As Object, e As DataGridViewCellPaintingEventArgs) Handles DataGridView1.CellPainting
        If e.RowIndex = -1 Then
            DataGridViewCellPaintingHeaderFormat(sender, e)
            Exit Sub
        End If
        Dim strName As String = DataGridView1.Columns(e.ColumnIndex).Name.ToLower
        Dim lngIndex As Long = DataGridView1.Rows(e.RowIndex).Cells("Index").Value
        If strName.Contains("prefix") AndAlso NotNull(DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value, "") <> "" Then
            DataGridViewAddLevelBar(DataGridView1, FullInventory(lngIndex).ExplicitPrefixMods, strName, sender, e)
        ElseIf strName.Contains("suffix") AndAlso NotNull(DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value, "") <> "" Then
            DataGridViewAddLevelBar(DataGridView1, FullInventory(lngIndex).ExplicitSuffixMods, strName, sender, e)
        End If
    End Sub

    Public Sub DataGridViewCellPaintingHeaderFormat(sender As Object, e As DataGridViewCellPaintingEventArgs)
        Dim c1 As Color = Color.FromArgb(255, 54, 54, 54)
        Dim c2 As Color = Color.FromArgb(255, 62, 62, 62)
        Dim c3 As Color = Color.FromArgb(255, 98, 98, 98)

        Dim br As New LinearGradientBrush(e.CellBounds, c1, c3, 90, True)
        Dim cb As New ColorBlend()
        cb.Positions = New Single() {0, CSng(0.5), 1}
        cb.Colors = New Color() {c1, c2, c3}
        br.InterpolationColors = cb

        e.Graphics.FillRectangle(br, e.CellBounds)
        e.PaintContent(e.ClipBounds)
        e.Handled = True
    End Sub

    Public Sub DataGridViewAddLevelBar(dg As DataGridView, MyModList As List(Of FullMod), strName As String, sender As Object, e As DataGridViewCellPaintingEventArgs)
        Try
            Dim bytPos As Byte = GetNumeric(strName) - 1
            Dim p As Double = 0, q As Double = 0
            'Dim sngMinDen As Single = 0     ' This is the denominator unit used by the values, for use in setting our level "notches"
            With MyModList(bytPos)
                If .UnknownValues = True Then Exit Sub
                ' In the calculations below, the +1 offset in both numerator and denominator means that the shortest "bar" at each level is a little larger
                ' than visually expected, but this is needed to separate the maximum value in a lower level from the minimum value in the level above it.
                ' This was also chosen to make the largest possible bar fill the cell, which is more visually pleasing thank other options.
                ' The goal with this is to give the user a quick idea of how good/bad the RNG for the mod is...if a more detailed understanding is required, 
                ' the double-click info is always available
                Dim sngModWeight1 As Single = 0.5, sngModWeight2 As Single = 0.5
                If .BaseUpperMaxV1 <> 0 Then
                    If .BaseUpperV1 = .BaseLowerV1 And .BaseUpperMaxV1 = .BaseLowerMaxV1 And .ModLevelMax = 0 Then HighlightNoVariationMods(dg, e) : Exit Sub ' There is no variation possible, so exit sub and don't draw bar
                    If .BaseUpperV1 = .BaseLowerV1 Then sngModWeight1 = 0 : sngModWeight2 = 1
                    If .BaseUpperMaxV1 = .BaseLowerMaxV1 Then sngModWeight1 = 1 : sngModWeight2 = 0
                    q = IIf(.BaseUpperV1 = .BaseLowerV1, 1, (.Value1 - .BaseLowerV1 + 1)) / IIf(.BaseUpperV1 = .BaseLowerV1, 1, (.BaseUpperV1 - .BaseLowerV1 + 1)) * sngModWeight1 + _
                        IIf(.BaseUpperMaxV1 = .BaseLowerMaxV1, 1, (.MaxValue1 - .BaseLowerMaxV1 + 1)) / IIf(.BaseUpperMaxV1 = .BaseLowerMaxV1, 1, (.BaseUpperMaxV1 - .BaseLowerMaxV1 + 1)) * sngModWeight2
                    'sngMinDen = IIf(.BaseUpperV1 = .BaseLowerV1, 0, 1 / IIf(.BaseUpperV1 = .BaseLowerV1, 1, .BaseUpperV1 - .BaseLowerV1 + 1)) * sngModWeight1 + _
                    '    IIf(.BaseUpperMaxV1 = .BaseLowerMaxV1, 0, 1 / IIf(.BaseUpperMaxV1 = .BaseLowerMaxV1, 1, (.BaseUpperMaxV1 - .BaseLowerMaxV1 + 1))) * sngModWeight2
                Else
                    If .BaseUpperV1 = .BaseLowerV1 And .Value2 = 0 And .ModLevelMax = 0 Then HighlightNoVariationMods(dg, e) : Exit Sub ' There is no variation possible, so exit sub and don't draw bar
                    q = IIf(.BaseUpperV1 = .BaseLowerV1, 1, (.Value1 - .BaseLowerV1 + 1)) / IIf(.BaseUpperV1 = .BaseLowerV1, 1, (.BaseUpperV1 - .BaseLowerV1 + 1))
                    'sngMinDen = IIf(.BaseUpperV1 = .BaseLowerV1, 0, 1 / IIf(.BaseUpperV1 = .BaseLowerV1, 1, .BaseUpperV1 - .BaseLowerV1 + 1))
                End If
                If .Value2 <> 0 Then
                    If .BaseUpperV1 = .BaseLowerV1 And .BaseUpperV2 = .BaseLowerV2 And .ModLevelMax = 0 Then HighlightNoVariationMods(dg, e) : Exit Sub ' There is no variation possible, so exit sub and don't draw bar
                    If .BaseUpperV1 = .BaseLowerV1 Then sngModWeight1 = 0 : sngModWeight2 = 1
                    If .BaseUpperV2 = .BaseLowerV2 Then sngModWeight1 = 1 : sngModWeight2 = 0
                    q = IIf(.BaseUpperV1 = .BaseLowerV1, 1, (.Value1 - .BaseLowerV1 + 1)) / IIf(.BaseUpperV1 = .BaseLowerV1, 1, (.BaseUpperV1 - .BaseLowerV1 + 1)) * sngModWeight1 + _
                        IIf(.BaseUpperV2 = .BaseLowerV2, 1, (.Value2 - .BaseLowerV2 + 1)) / IIf(.BaseUpperV2 = .BaseLowerV2, 1, (.BaseUpperV2 - .BaseLowerV2 + 1)) * sngModWeight2
                    'sngMinDen = IIf(.BaseUpperV1 = .BaseLowerV1, 0, 1) / IIf(.BaseUpperV1 = .BaseLowerV1, 1, (.BaseUpperV1 - .BaseLowerV1 + 1)) * sngModWeight1 + _
                    '    IIf(.BaseUpperV2 = .BaseLowerV2, 0, 1) / IIf(.BaseUpperV2 = .BaseLowerV2, 1, (.BaseUpperV2 - .BaseLowerV2 + 1)) * sngModWeight2
                End If
                If .ModLevelMax = 0 Then
                    p = q
                Else
                    p = .ModLevelActual / (.ModLevelMax + 1) + (q / (.ModLevelMax + 1))
                End If
            End With
            If p = 0 Then Exit Sub
            e.Graphics.FillRectangle(Brushes.White, e.CellBounds)
            e.Paint(e.ClipBounds, DataGridViewPaintParts.All And Not DataGridViewPaintParts.Background)
            Dim r As Drawing.Rectangle
            r.X = e.CellBounds.X + 3
            r.Y = e.CellBounds.Y + 3
            r.Width = (e.CellBounds.Width - 6) * p
            r.Height = e.CellBounds.Height - 6
            Dim br2 As New Drawing.Drawing2D.LinearGradientBrush(r, Drawing.Color.White, Drawing.Color.DarkGray, Drawing.Drawing2D.LinearGradientMode.Vertical)
            e.Graphics.FillRectangle(br2, r)
            e.PaintContent(e.ClipBounds)
            '' Now draw the vertical notches that indicate to the user where the level 'breaks' are
            'For i = 0 To MyModList(bytPos).ModLevelMax
            '    Dim sngStart As Single = CSng(e.CellBounds.X + 3 + e.CellBounds.Width * i / (MyModList(bytPos).ModLevelMax + 1)) + sngMinDen * (e.CellBounds.Width - 6)
            '    e.Graphics.DrawLine(New Pen(SystemColors.ControlDark, 2), sngStart, CSng(e.CellBounds.Y), sngStart, CSng(e.CellBounds.Y + 3))
            'Next
            e.Handled = True
        Catch ex As Exception
            ErrorHandler(System.Reflection.MethodBase.GetCurrentMethod.Name, ex)
        End Try
    End Sub

    Private Sub DataGridView1_RowPostPaint(sender As Object, e As DataGridViewRowPostPaintEventArgs) Handles DataGridView1.RowPostPaint
        If e.RowIndex = -1 Then Exit Sub
        DataGridViewRowPostPaint(DataGridView1, FullInventory, sender, e)
    End Sub

    Public Sub DataGridViewRowPostPaint(dg As DataGridView, MyInventory As List(Of FullItem), sender As Object, e As DataGridViewRowPostPaintEventArgs)
        Try
            If dg.Rows(e.RowIndex).Cells("Gem").Value = True Then
                dg.Rows(e.RowIndex).Cells("Gem").Style.BackColor = ColorHasGem
                dg.Rows(e.RowIndex).Cells("Rank").Style.BackColor = ColorHasGem
                dg.Rows(e.RowIndex).Cells("Level").Style.BackColor = ColorHasGem
                dg.Rows(e.RowIndex).Cells("%").Style.BackColor = ColorHasGem
            End If
            Dim lngindex As Long = dg.Rows(e.RowIndex).Cells("Index").Value
            If lngindex <> -1 Then
                HighlightUnknownMods(MyInventory(lngindex).ExplicitPrefixMods, "Prefix ", e.RowIndex)
                HighlightUnknownMods(MyInventory(lngindex).ExplicitSuffixMods, "Suffix ", e.RowIndex)
                For i = 1 To 3
                    If i > MyInventory(lngindex).ExplicitPrefixMods.Count Then Exit For
                    If intWeightMax <= MyInventory(lngindex).ExplicitPrefixMods(i - 1).Weight Then
                        dg.Rows(e.RowIndex).Cells("Prefix " & i).Style.ForeColor = ColorMax
                        dg.Rows(e.RowIndex).Cells("Prefix " & i).Style.Font = New Font("Segoe UI", 8.25, FontStyle.Regular)
                    ElseIf intWeightMin >= MyInventory(lngindex).ExplicitPrefixMods(i - 1).Weight Then
                        dg.Rows(e.RowIndex).Cells("Prefix " & i).Style.ForeColor = ColorMin
                        dg.Rows(e.RowIndex).Cells("Prefix " & i).Style.Font = New Font("Segoe UI", 8.25, FontStyle.Italic)
                    End If
                Next
                For i = 1 To 3
                    If i > MyInventory(lngindex).ExplicitSuffixMods.Count Then Exit For
                    If intWeightMax <= MyInventory(lngindex).ExplicitSuffixMods(i - 1).Weight Then
                        dg.Rows(e.RowIndex).Cells("Suffix " & i).Style.ForeColor = ColorMax
                        dg.Rows(e.RowIndex).Cells("Suffix " & i).Style.Font = New Font("Segoe UI", 8.25, FontStyle.Regular)
                    ElseIf intWeightMin >= MyInventory(lngindex).ExplicitSuffixMods(i - 1).Weight Then
                        dg.Rows(e.RowIndex).Cells("Suffix " & i).Style.ForeColor = ColorMin
                        dg.Rows(e.RowIndex).Cells("Suffix " & i).Style.Font = New Font("Segoe UI", 8.25, FontStyle.Italic)
                    End If
                Next
            End If
            If dg.Rows(e.RowIndex).Cells("*").Value = "*" Then dg.Rows(e.RowIndex).Cells("*").Style.BackColor = ColorOtherSolutions
        Catch ex As Exception
            ErrorHandler(System.Reflection.MethodBase.GetCurrentMethod.Name, ex)
        End Try
    End Sub

    Private Sub HighlightUnknownMods(MyList As List(Of FullMod), strAffix As String, RowIndex As Integer)
        For Each MyMod In MyList
            If MyMod.UnknownValues = True Then
                DataGridView1.Rows(RowIndex).Cells(strAffix & MyList.IndexOf(MyMod) + 1).Style.BackColor = ColorUnknownValue
            End If
        Next
    End Sub

    Public Sub HighlightNoVariationMods(dg As DataGridView, e As DataGridViewCellPaintingEventArgs)
        e.Graphics.FillRectangle(Brushes.White, e.CellBounds)
        e.Paint(e.ClipBounds, DataGridViewPaintParts.All And Not DataGridViewPaintParts.Background)
        dg.Rows(e.RowIndex).Cells(e.ColumnIndex).Style.BackColor = ColorNoModVariation
    End Sub

    Private Sub cmbWeight_SelectionChangeCommitted(sender As Object, e As EventArgs) Handles cmbWeight.SelectionChangeCommitted
        Try
            ' First we load the new csv into dtWeights
            UserSettings("Weight") = Me.Invoke(New RCPD(AddressOf ReadControlProperty), New Object() {cmbWeight, "SelectedItem"}).ToString
            dtWeights = LoadCSVtoDataTable(Application.StartupPath & "\weights-" & UserSettings("Weight") & ".csv")

            If DataGridView1.Visible = True Then
                grpProgress.Left = (Me.Width - grpProgress.Width) / 2
                grpProgress.Top = (Me.Height - grpProgress.Height) / 2
                pb.Minimum = 1
                pb.Maximum = CInt((FullInventory.Count + TempInventory.Count) / 10)
                pb.Value = 1
                pb.Step = 1
                grpProgress.Visible = True
                pb.Visible = True
                lblpb.Visible = True
                ' Recalculate all the rankings based on the new weights
                RecalculateThread = New Threading.Thread(AddressOf RecalculateAllRankings)
                RecalculateThread.SetApartmentState(Threading.ApartmentState.STA)
                RecalculateThread.Start()
            End If

        Catch ex As Exception
            ErrorHandler(System.Reflection.MethodBase.GetCurrentMethod.Name, ex)
        End Try
    End Sub

    Public Sub PBPerformStep()
        pb.PerformStep()
    End Sub

    Public Sub PBClose()
        grpProgress.Visible = False
        pb.Visible = False
        lblpb.Visible = False
    End Sub

    Public Sub RecalculateAllRankings()
        Try
            ' Full speed ahead with the full recalculating!
            Me.Invoke(New UCPD(AddressOf SetControlProperty), New Object() {cmbWeight, "Enabled", False})
            Me.Invoke(New UCPD(AddressOf SetControlProperty), New Object() {lblWeights, "Enabled", False})
            Me.Invoke(New UCPD(AddressOf SetControlProperty), New Object() {btnEditWeights, "Enabled", False})
            Me.Invoke(New UCPD(AddressOf SetControlProperty), New Object() {DataGridView1, "Visible", False})
            Me.Invoke(New UCPD(AddressOf SetControlProperty), New Object() {DataGridView1, "DataSource", ""})
            Me.Invoke(New UCPD(AddressOf SetControlProperty), New Object() {Me, "UseWaitCursor", True})
            Dim lngCounter As Long = 0
            For Each mylist In {FullInventory, TempInventory}
                Dim strList As String = ""
                If mylist.Count = FullInventory.Count Then strList = "Full" Else strList = "Temp"
                For Each it In mylist
                    it.Rank = 0
                    Dim strRankKey As String = IIf(strList = "Full", it.ID & it.Name, it.ID & it.Name & mylist.IndexOf(it))
                    RankExplanation(strRankKey) = ""
                    For Each ModList In {it.ExplicitPrefixMods, it.ExplicitSuffixMods}
                        Dim strAffix As String = ""
                        If ModList.Count <> it.ExplicitPrefixMods.Count Then
                            strAffix = "Suffix"
                        Else
                            If ModList(0).FullText <> it.ExplicitPrefixMods(0).FullText Then strAffix = "Suffix" Else strAffix = "Prefix"
                        End If
                        For Each myMod In ModList
                            Dim result() As DataRow = Nothing
                            result = dtWeights.Select("ExportField = '" & myMod.Type1 & "'" & IIf(myMod.Type2 <> "", " AND ExportField2 = '" & myMod.Type2 & "'", " AND ExportField2 = ''"))
                            If result.Count <> 0 Then
                                For Each row In result
                                    If RunModResultQuery(it, row).Count <> 0 Then
                                        myMod.Weight = row("Weight")
                                        Exit For
                                    End If
                                Next
                                If myMod.UnknownValues = True Then
                                    it.Rank += RankUnknownValueMod(myMod, myMod.MaxValue1, myMod.ModLevelMax, it.ID, it.Name, strAffix, strRankKey)
                                Else
                                    it.Rank += CalculateRank(myMod, myMod.ModLevelMax, result(0)("Description") & myMod.ModLevelActual, it.ID, it.Name, strAffix, strRankKey)
                                End If
                            End If
                            RankExplanation(strRankKey) += vbCrLf & vbCrLf & "(" & String.Format("{0:'+'0;'-'0}", it.Rank) & ")  (running total)" & vbCrLf
                        Next
                    Next
                    RankExplanation(strRankKey) = RankExplanation(strRankKey).Substring(0, RankExplanation(strRankKey).LastIndexOf("(running total)")) & "Final Rank" & vbCrLf
                    lngCounter += 1
                    If lngCounter Mod 10 = 0 Then
                        Me.Invoke(New MyDelegate(AddressOf PBPerformStep))
                        'Application.DoEvents()
                    End If
                Next
            Next
            Application.DoEvents()
            dtRank.Clear() : dtOverflow.Clear()
            AddToDataTable(FullInventory, dtRank, True, "Full Inventory Table")
            AddToDataTable(TempInventory, dtOverflow, True, "Overflow Inventory Table")

            Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {DataGridView1, "DataSource", dtRank})
            bytColumns = CByte(Me.Invoke(New RCPD(AddressOf ReadControlProperty), New Object() {DataGridView1.Columns, "Count"}).ToString)
            Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {DataGridView1.Columns(bytColumns - 1), "Visible", False})
            Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {DataGridView1.Columns(bytColumns - 2), "Visible", False})
            Me.Invoke(New MyDelegate(AddressOf SortDataGridView))
            Me.Invoke(New MyControlDelegate(AddressOf SetDataGridViewWidths), New Object() {DataGridView1})
            'Me.Invoke(New UCPD(AddressOf SetControlProperty), New Object() {DataGridView1.Rows(0), "Selected", True})
            Dim FirstCell As DataGridViewCell
            FirstCell = Me.Invoke(New MyDelegateFunction(AddressOf ReturnFirstCell))
            Me.Invoke(New UCPD(AddressOf SetControlProperty), New Object() {DataGridView1, "FirstDisplayedCell", FirstCell})
            Me.Invoke(New UCPD(AddressOf SetControlProperty), New Object() {DataGridView1, "Visible", True})
            'frmPB.Close() : frmPB.Dispose()
            Me.Invoke(New MyDelegate(AddressOf PBClose))
            Me.Invoke(New UCPD(AddressOf SetControlProperty), New Object() {cmbWeight, "Enabled", True})
            Me.Invoke(New UCPD(AddressOf SetControlProperty), New Object() {lblWeights, "Enabled", True})
            Me.Invoke(New UCPD(AddressOf SetControlProperty), New Object() {btnEditWeights, "Enabled", True})
            Me.Invoke(New UCPD(AddressOf SetControlProperty), New Object() {Me, "UseWaitCursor", False})
            Application.DoEvents()
            ' Sometimes the wait cursor property doesn't get unset for the datagridview, perhaps because it is not visible for some time?
            Me.Invoke(New UCPD(AddressOf SetControlProperty), New Object() {DataGridView1, "UseWaitCursor", False})
            Me.Invoke(New UCPD(AddressOf SetControlProperty), New Object() {DataGridView1, "Cursor", Cursors.Default})
            Me.BeginInvoke(New MyDelegate(AddressOf SetDataGridFocus))

        Catch ex As Exception
            Me.Invoke(New MyDelegate(AddressOf PBClose))
            Me.Invoke(New UCPD(AddressOf SetControlProperty), New Object() {cmbWeight, "Enabled", True})
            Me.Invoke(New UCPD(AddressOf SetControlProperty), New Object() {lblWeights, "Enabled", True})
            Me.Invoke(New UCPD(AddressOf SetControlProperty), New Object() {btnEditWeights, "Enabled", True})
            Me.Invoke(New UCPD(AddressOf SetControlProperty), New Object() {Me, "UseWaitCursor", False})

            ErrorHandler(System.Reflection.MethodBase.GetCurrentMethod.Name, ex)
        End Try
    End Sub

    Private Function ReturnFirstCell() As DataGridViewCell
        Return DataGridView1.Rows(0).Cells(0)
    End Function

    Private Sub DataGridView1_Scroll(sender As Object, e As ScrollEventArgs) Handles DataGridView1.Scroll
        blScroll = True
    End Sub

    Private Sub DataGridView1_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView1.SelectionChanged
        If DataGridView1.CurrentCell Is Nothing Then Exit Sub
        If DataGridView1.Cursor <> Cursors.Default Then DataGridView1.Cursor = Cursors.Default
        If DataGridView1.Columns(DataGridView1.CurrentCell.ColumnIndex).Name.ToLower.Contains("prefix") Or DataGridView1.Columns(DataGridView1.CurrentCell.ColumnIndex).Name.ToLower.Contains("suffix") Then DataGridView1.ClearSelection()
    End Sub

    Private Sub chkSession_CheckedChanged(sender As Object, e As EventArgs) Handles chkSession.CheckedChanged
        If chkSession.Checked = True Then
            lblEmail.Text = "Username:"
            lblPassword.Text = "SessionID:"
        Else
            lblEmail.Text = "Email:"
            lblPassword.Text = "Password:"
        End If
    End Sub

    Private Sub btnEditWeights_Click(sender As Object, e As EventArgs) Handles btnEditWeights.Click
        frmWeights.ShowDialog()
    End Sub

End Class