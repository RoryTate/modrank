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
    Public RecalculateThread As Threading.Thread
    Public ProgressBarThread As Threading.Thread
    Public DownloadThread As Threading.Thread
    Public LoadCacheThread As Threading.Thread
    Public RefreshCacheThread As Threading.Thread

    Private WPFPassword As New PasswordBox
    Private statusBox As Windows.Controls.RichTextBox = New RichTextBox
    Private statusController As ModRank.frmMain.StatusController
    Private blOffline As Boolean = False

    Public FullStash As Stash
    Public invLocation As New Dictionary(Of Item, String)
    Public Shared Characters As New List(Of Character)()
    Public Shared Leagues As New List(Of String)()
    Public Shared FontCollection As New Drawing.Text.PrivateFontCollection()
    Public statusCounter As Long, lngModCounter As Long
    Public blSolomonsJudgment As New Dictionary(Of String, Boolean)        ' Used to decide which mod gets both .5's in a two-way combined split
    Public blAddedOne As Boolean = False    '  Used in the dynamic mod evaluation method, to help know when a mod contains "Legacy" values
    Public blScroll As Boolean = False

    Public lngStoreCount As Long = 0

    Dim oldColorDark As Color = Color.FromArgb(127, 127, 127)
    Dim oldColorLight As Color = Color.FromArgb(195, 195, 195)

    Public RankExplanation As New Dictionary(Of String, String)

    Private Delegate Sub FillCredentialsDelegate()
    Public Delegate Sub MyDelegate()
    Public Delegate Sub MyDGDelegate(dg As DataGridView)
    Public Delegate Function MyDelegateFunction() As Object
    Public Delegate Function DataGridCell(dg As DataGridView) As DataGridViewCell
    Public Delegate Sub pbSetDefaults(intMax As Integer, strLabel As String)
    Public Delegate Sub MyClickDelegate(sender As Object, e As EventArgs)

    Public blRepopulated As Boolean = False

    Private frmShopFilter As New frmFilter

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
        ' Pull down the combo box, since the user will likely want to select their new weights file to use, but only if we aren't recalculating
        If IsNothing(RecalculateThread) = False AndAlso RecalculateThread.IsAlive = True Then Exit Sub
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
            If System.IO.Directory.Exists(Application.StartupPath & "\Store") = False Then System.IO.Directory.CreateDirectory(Application.StartupPath & "\Store")
            If System.IO.Directory.Exists(Application.StartupPath & "\Store\Filters") = False Then System.IO.Directory.CreateDirectory(Application.StartupPath & "\Store\Filters")
            ' Let's see if we can improve the display performance of the datagridview a bit...holy crap! This setting works unbelievably well!
            GetType(DataGridView).InvokeMember("DoubleBuffered", BindingFlags.NonPublic Or BindingFlags.Instance Or BindingFlags.SetProperty, Nothing, DataGridView1, New Object() {True})
            GetType(DataGridView).InvokeMember("DoubleBuffered", BindingFlags.NonPublic Or BindingFlags.Instance Or BindingFlags.SetProperty, Nothing, DataGridView2, New Object() {True})
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
                Dim MyPic As PictureBox = CType(Me.gpLegend.Controls("pic" & x), PictureBox)
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
                newRow("MinV") = IIf(IsDBNull(dtRow("MinV")) Or dtRow("MinV") Is "", 0, dtRow("MinV"))
                newRow("MaxV") = IIf(IsDBNull(dtRow("MaxV")) Or dtRow("MaxV") Is "", 0, dtRow("MaxV"))
                newRow("MinV2") = IIf(IsDBNull(dtRow("MinV2")) Or dtRow("MinV2") Is "", 0, dtRow("MinV2"))
                newRow("MaxV2") = IIf(IsDBNull(dtRow("MaxV2")) Or dtRow("MaxV2") Is "", 0, dtRow("MaxV2"))
                newRow("Name") = dtRow("Name")
                newRow("Level") = IIf(IsDBNull(dtRow("Level")) Or dtRow("Level") Is "", 0, dtRow("Level"))
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
            dtRank.Columns.Add("Qal", GetType(Byte))
            dtRank.Columns.Add("Crpt", GetType(Boolean))
            dtRank.Columns.Add("Price", GetType(String))
            dtRank.Columns.Add("PriceNum", GetType(Integer))
            dtRank.Columns.Add("PriceOrb", GetType(String))
            dtRank.Columns.Add("Tot1", GetType(Single))
            dtRank.Columns.Add("Tot2", GetType(Single))
            dtRank.Columns.Add("Tot3", GetType(Single))
            dtRank.Columns.Add("Tot4", GetType(Single))
            dtRank.Columns.Add("Tot5", GetType(Single))
            dtRank.Columns.Add("Tot6", GetType(Single))
            dtRank.Columns.Add("SktClrs", GetType(String))
            dtRank.Columns.Add("SktClrsSearch", GetType(String))
            dtRank.Columns.Add("Arm", GetType(Integer))
            dtRank.Columns.Add("Eva", GetType(Integer))
            dtRank.Columns.Add("ES", GetType(Integer))
            dtRank.Columns.Add("pft1", GetType(String))
            dtRank.Columns.Add("pt1", GetType(String))
            dtRank.Columns.Add("pt12", GetType(String))
            dtRank.Columns.Add("pft2", GetType(String))
            dtRank.Columns.Add("pt2", GetType(String))
            dtRank.Columns.Add("pt22", GetType(String))
            dtRank.Columns.Add("pft3", GetType(String))
            dtRank.Columns.Add("pt3", GetType(String))
            dtRank.Columns.Add("pt32", GetType(String))
            dtRank.Columns.Add("pv1", GetType(Single))
            dtRank.Columns.Add("pv1m", GetType(Single))
            dtRank.Columns.Add("pv12", GetType(Single))
            dtRank.Columns.Add("pv2", GetType(Single))
            dtRank.Columns.Add("pv2m", GetType(Single))
            dtRank.Columns.Add("pv22", GetType(Single))
            dtRank.Columns.Add("pv3", GetType(Single))
            dtRank.Columns.Add("pv3m", GetType(Single))
            dtRank.Columns.Add("pv32", GetType(Single))
            dtRank.Columns.Add("pcount", GetType(Byte))
            dtRank.Columns.Add("sft1", GetType(String))
            dtRank.Columns.Add("st1", GetType(String))
            dtRank.Columns.Add("st12", GetType(String))
            dtRank.Columns.Add("sft2", GetType(String))
            dtRank.Columns.Add("st2", GetType(String))
            dtRank.Columns.Add("st22", GetType(String))
            dtRank.Columns.Add("sft3", GetType(String))
            dtRank.Columns.Add("st3", GetType(String))
            dtRank.Columns.Add("st32", GetType(String))
            dtRank.Columns.Add("sv1", GetType(Single))
            dtRank.Columns.Add("sv1m", GetType(Single))
            dtRank.Columns.Add("sv12", GetType(Single))
            dtRank.Columns.Add("sv2", GetType(Single))
            dtRank.Columns.Add("sv2m", GetType(Single))
            dtRank.Columns.Add("sv22", GetType(Single))
            dtRank.Columns.Add("sv3", GetType(Single))
            dtRank.Columns.Add("sv3m", GetType(Single))
            dtRank.Columns.Add("sv32", GetType(Single))
            dtRank.Columns.Add("scount", GetType(Byte))
            dtRank.Columns.Add("ecount", GetType(Byte))
            dtRank.Columns.Add("it1", GetType(String))
            dtRank.Columns.Add("it2", GetType(String))
            dtRank.Columns.Add("it3", GetType(String))
            dtRank.Columns.Add("iv1", GetType(Single))
            dtRank.Columns.Add("iv1m", GetType(Single))
            dtRank.Columns.Add("iv2", GetType(Single))
            dtRank.Columns.Add("iv2m", GetType(Single))
            dtRank.Columns.Add("iv3", GetType(Single))
            dtRank.Columns.Add("iv3m", GetType(Single))
            dtRank.Columns.Add("icount", GetType(Byte))
            dtRank.Columns.Add("ThreadID", GetType(String))
            dtRank.Columns.Add("Index", GetType(Long))
            dtRank.Columns.Add("ID", GetType(String))
            Dim primaryKey(1) As DataColumn
            primaryKey(1) = dtRank.Columns("ID")
            dtRank.PrimaryKey = primaryKey

            dtOverflow = dtRank.Clone
            dtRankFilter = dtRank.Clone
            dtStore = dtRank.Clone
            dtStoreOverflow = dtRank.Clone

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
            chkSession.Checked = CBool(UserSettings("UseSessionID"))
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
            Dim strDoubleClick As String = "Click on any of the Rank, Prefix, Suffix, *, Location, and Sokt cells (where values exist) to get more detailed information on the item."
            ToolTip1.SetToolTip(lblDoubleClick, strDoubleClick)

            Dim strDownload As String = "Enter poexplorer/GGG forum thread or full JSON URL in this box. For example:" & Environment.NewLine & Environment.NewLine & _
                "796106" & Environment.NewLine & Environment.NewLine & _
                "OR" & Environment.NewLine & Environment.NewLine & _
                "http://poexplorer.com/threads/796106.json"
            ToolTip1.SetToolTip(txtThread, strDownload)

        Catch ex As Exception
            ErrorHandler(System.Reflection.MethodBase.GetCurrentMethod.Name, ex)
        End Try
    End Sub

    Public Sub PasswordChanged(sender As Object, e As System.Windows.RoutedEventArgs)
        blFormChanged = True
    End Sub

    Public Sub PasswordEnabledChanged(sender As Object, e As System.Windows.DependencyPropertyChangedEventArgs)
        If WPFPassword.IsEnabled = True Then Exit Sub
        Dim ctlColor As Color = SystemColors.Control
        WPFPassword.Background = New System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromArgb(ctlColor.A, ctlColor.R, ctlColor.G, ctlColor.B))
    End Sub

    Public Sub PopulateWeightsComboBox(strSelection As String, Optional blReload As Boolean = False)
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

            If blReload Then
                dtWeights.Clear()
                dtWeights = LoadCSVtoDataTable(Application.StartupPath & "\weights-" & strSelection & ".csv")
                If dtRank.Rows.Count <> 0 Or dtStore.Rows.Count <> 0 Then
                    ' Recalculate all the rankings based on the new weights
                    RecalculateThread = New Threading.Thread(AddressOf RecalculateAllRankings)
                    RecalculateThread.SetApartmentState(Threading.ApartmentState.STA)
                    RecalculateThread.Start()
                End If
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

            EnableDisableControls(False, New List(Of String)(New String() {"ElementHost2"}))

            Dim password As SecureString
            If blFormChanged Then
                password = Me.WPFPassword.SecurePassword
            Else
                password = UserSettings("AccountPassword").Decrypt
            End If
            Email = Me.txtEmail.Text
            useSession = chkSession.Checked
            Model.Authenticate(Email, password, blOffline, useSession)
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

            FullInventoryCache.Clear() : TempInventoryCache.Clear()
            Dim strCache As String = Application.StartupPath & "\fsinv.cache"
            If blOffline And File.Exists(strCache) Then LoadCache(FullInventory, strCache) Else LoadCache(FullInventoryCache, strCache)
            strCache = Application.StartupPath & "\tsinv.cache"
            If blOffline And File.Exists(strCache) Then LoadCache(TempInventory, strCache) Else LoadCache(TempInventoryCache, strCache)
            
            If blOffline And File.Exists(Application.StartupPath & "\fsinv.cache") Then GoTo DataTableLoad

            Dim lngCount As Long = 0    ' Counter to keep track of where we are in temporary inventory additions (used by complex search algorithm)
            ' Get only the gear types out of the temp inventory list (we don't need gems, maps, currency, etc)
            statusCounter = 0 : lngModCounter = 0
            Dim query As IEnumerable(Of Item) = TempInventoryAll.Where(Function(Item) Item.ItemType = ItemType.Gear AndAlso Item.Name.ToLower <> "" _
                                                                           AndAlso Item.TypeLine.ToLower.Contains("map") = False AndAlso Item.TypeLine.ToLower.Contains("peninsula") = False)
            AddToFullInventory(query, False, lngCount)

            For Each league In Leagues
                CurrentLeague = league
                FullStash = Model.GetStash(league)
                Dim TempStash As New List(Of Item)
                TempStash = FullStash.Get(Of Item)()

                query = TempStash.Where(Function(Item) Item.ItemType = ItemType.Gear AndAlso Item.Name.ToLower <> "" _
                                            AndAlso Item.TypeLine.ToLower.Contains("map") = False AndAlso Item.TypeLine.ToLower.Contains("peninsula") = False)
                AddToFullInventory(query, True, lngCount)
            Next

DataTableLoad:
            statusController.DisplayMessage("Calculating rank scores...")
            RecalculateAllRankings(True, False, True)
            statusController.DisplayMessage("Completed indexing all " & statusCounter & " stash rare items. comprising a total of approximately " & lngModCounter & " mods.")
            statusController.DisplayMessage("Now loading the data into a table to display the results...please wait, this may take a moment.")

            AddToDataTable(FullInventory, dtRank, True, "Full Inventory Table")
            strCache = Application.StartupPath & "\fsinv.cache"
            If blOffline = False Or File.Exists(strCache) = False Then
                WriteCache(FullInventory, strCache)
            End If

            AddToDataTable(TempInventory, dtOverflow, True, "Overflow Inventory Table")
            strCache = Application.StartupPath & "\tsinv.cache"
            If blOffline = False Or File.Exists(strCache) = False Then
                WriteCache(TempInventory, strCache)
            End If

            Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {DataGridView1, "DataSource", dtRank})
            Me.Invoke(New MyDualControlDelegate(AddressOf HideColumns), New Object() {Me, DataGridView1})
            Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {DataGridView1.Columns("%").DefaultCellStyle, "Format", "n1"})
            Me.Invoke(New MyDGDelegate(AddressOf SortDataGridView), DataGridView1)
            'Me.Invoke(New MyDelegate(AddressOf SortDataGridView))
            Me.Invoke(New MyControlDelegate(AddressOf SetDataGridViewWidths), New Object() {DataGridView1})
            Dim FirstCell As DataGridViewCell
            FirstCell = Me.Invoke(New DataGridCell(AddressOf ReturnFirstCell), New Object() {DataGridView1})
            Me.Invoke(New UCPD(AddressOf SetControlProperty), New Object() {DataGridView1, "FirstDisplayedCell", FirstCell})

            Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {Me, "WindowState", FormWindowState.Maximized})
            EnableDisableControls(True, New List(Of String)(New String() {"ElementHost2", "txtEmail", "lblEmail", "ElementHost1", "lblPassword", "btnLoad", "btnOffline", "chkSession"}))
            Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {ElementHost2, "Visible", False})
            Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {DataGridView1, "Visible", True})
            Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {gpLegend, "Left", 550})
            Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {gpLegend, "Visible", True})
            Me.Invoke(New MyDelegate(AddressOf BringLegendToFront))
            Dim intRows As Integer = CInt(Me.Invoke(New RCPD(AddressOf ReadControlProperty), New Object() {DataGridView1.Rows, "Count"}).ToString)
            Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {lblRecordCount, "Text", "Number of Rows: " & intRows})
            Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {lblRecordCount, "Visible", True})
            Me.BeginInvoke(New MyDGDelegate(AddressOf SetDataGridFocus), DataGridView1)
            'Me.Invoke(New MyDelegate(AddressOf SetDataGridFocus))
            RemoveHandler Model.StashLoading, AddressOf model_StashLoading
            RemoveHandler Model.Throttled, AddressOf model_Throttled
            saveSettings(password)
            ' Enable the following to see all of the columns to debug and troubleshoot
            'For i = bytDisplay To bytColumns - 1
            '    Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {DataGridView1.Columns(i), "Visible", True})
            'Next

            If strFilter.Trim <> "" Then
                Dim FilterThread As New Threading.Thread(Sub() ApplyFilter(dtRank, dtOverflow, DataGridView1, lblRecordCount, strOrderBy, strFilter))
                FilterThread.SetApartmentState(Threading.ApartmentState.STA)
                FilterThread.Start()
            End If

        Catch ex As Exception
            EnableDisableControls(True, New List(Of String)(New String() {"ElementHost2"}))
            RemoveHandler Model.StashLoading, AddressOf model_StashLoading
            RemoveHandler Model.Throttled, AddressOf model_Throttled
            ErrorHandler(System.Reflection.MethodBase.GetCurrentMethod.Name, ex)
        End Try
    End Sub

    Public Sub EnableDisableControls(blEnable As Boolean, Optional lstIgnore As List(Of String) = Nothing)
        For Each ctl As System.Windows.Forms.Control In Me.Controls
            If lstIgnore Is Nothing = False AndAlso lstIgnore.Contains(ctl.Name) = True Then Continue For
            If ctl.HasChildren Then
                For Each ctlChild As System.Windows.Forms.Control In ctl.Controls
                    If lstIgnore Is Nothing = False AndAlso lstIgnore.Contains(ctlChild.Name) = True Then Continue For
                    If ctlChild.HasChildren Then
                        For Each ctlGrandChild As System.Windows.Forms.Control In ctlChild.Controls
                            If lstIgnore Is Nothing = False AndAlso lstIgnore.Contains(ctlGrandChild.Name) = True Then Continue For
                            Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {ctlGrandChild, "Enabled", blEnable})
                        Next
                    Else
                        Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {ctlChild, "Enabled", blEnable})
                    End If
                Next
            Else
                Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {ctl, "Enabled", blEnable})
            End If
        Next
    End Sub

    Private Sub SetDataGridFocus(dg As DataGridView)
        dg.Focus()
        dg.UseWaitCursor = False
    End Sub

    Private Sub DataGridRefresh()
        Me.DataGridView1.Refresh()
    End Sub

    Private Sub DataGridClearSelection(dg As DataGridView)
        dg.ClearSelection()
    End Sub

    Private Sub SortDataGridView(dg As DataGridView)
        dg.Sort(dg.Columns(0), System.ComponentModel.ListSortDirection.Descending)
    End Sub

    Private Sub BringLegendToFront()
        gpLegend.BringToFront()
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

    Private Sub AddToFullInventory(query As IEnumerable(Of Item), blStash As Boolean, ByRef lngCount As Long)
        Try
            ' Convert the temporary inventory list into our FullInventory list of FullItem class types
            Dim myGear As Gear
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
                If myGear.GearType.ToString.CompareMultiple(StringComparison.Ordinal, "Sword", "Mace", "Axe") = True Then
                    If myGear.GearType.ToString = "Sword" Or myGear.GearType.ToString = "Axe" Then
                        If myGear.H = 4 And myGear.W = 2 Then
                            newFullItem.GearType = myGear.GearType.ToString & " (2h)"
                        Else
                            If myGear.TypeLine.ToString.CompareMultiple(StringComparison.OrdinalIgnoreCase, "corroded blade", "longsword", "butcher sword", "headman's sword") Then
                                newFullItem.GearType = myGear.GearType.ToString & " (2h)"
                            Else
                                newFullItem.GearType = myGear.GearType.ToString & " (1h)"
                            End If
                        End If
                    ElseIf myGear.GearType.ToString = "Mace" Then
                        If myGear.H = 4 And myGear.W = 2 Then newFullItem.GearType = myGear.GearType.ToString & " (2h)" Else newFullItem.GearType = myGear.GearType.ToString & " (1h)"
                    End If
                Else
                    newFullItem.GearType = myGear.GearType.ToString
                End If
                If newFullItem.ItemType Is Nothing Or newFullItem.GearType = "Unknown" Then
                    MessageBox.Show("An item (Name: " & myGear.Name & ", Location: " & newFullItem.Location & ") does not have an entry for its type and/or subtype in the APIs. Please report this to:" & Environment.NewLine & Environment.NewLine & _
                                    "https://github.com/RoryTate/modrank/issues" & Environment.NewLine & Environment.NewLine & _
                                    "Also, please provide the actual base type (i.e. Quiver) and subtype (i.e. Light Quiver) that are missing.", "Item Type/Subtype Not Found", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Continue For
                End If
                newFullItem.TypeLine = myGear.TypeLine.ToString
                newFullItem.H = CByte(myGear.H)
                newFullItem.W = CByte(myGear.W)
                newFullItem.X = myGear.X
                newFullItem.Y = myGear.Y
                newFullItem.Name = myGear.Name
                newFullItem.Rarity = myGear.Rarity
                For i = 0 To myGear.Requirements.Count - 1
                    If myGear.Requirements(i).Name.ToLower = "level" Then
                        newFullItem.Level = CByte(GetNumeric(myGear.Requirements(i).Value))
                        newFullItem.LevelGem = myGear.Requirements(i).Value.IndexOf("gem", StringComparison.OrdinalIgnoreCase) > -1
                        Exit For ' We don't care about any other requirements, so exit the loop
                    End If
                Next
                If Not IsNothing(myGear.Properties) Then
                    For i = 0 To myGear.Properties.Count - 1
                        If myGear.Properties(i).Name.ToLower = "armour" Then
                            newFullItem.Arm = CInt(myGear.Properties(i).Values(0).Item1)
                        ElseIf myGear.Properties(i).Name.ToLower = "evasion rating" Then
                            newFullItem.Eva = CInt(myGear.Properties(i).Values(0).Item1)
                        ElseIf myGear.Properties(i).Name.ToLower = "energy shield" Then
                            newFullItem.ES = CInt(myGear.Properties(i).Values(0).Item1)
                        End If
                    Next
                End If
                newFullItem.Sockets = CByte(myGear.NumberOfSockets)
                Dim sb As New System.Text.StringBuilder(11), intGroup As Integer = 0, blFirst As Boolean = True
                For i = 0 To newFullItem.Sockets - 1
                    If intGroup = myGear.Sockets(i).Group And blFirst = False Then
                        sb.Append("-")
                    ElseIf blFirst = False Then
                        sb.Append(" ")
                    End If
                    blFirst = False
                    sb.Append(myGear.Sockets(i).Attribute)
                    intGroup = myGear.Sockets(i).Group
                Next
                sb.Replace("I", "b") : sb.Replace("S", "r") : sb.Replace("D", "g")
                newFullItem.Colours = sb.ToString
                For i = 6 To 0 Step -1
                    If myGear.IsLinked(i) Then
                        newFullItem.Links = CByte(IIf(i = 1, 0, i)) : Exit For
                    End If
                Next
                If myGear.IsQuality Then newFullItem.Quality = CByte(myGear.Quality)
                newFullItem.Corrupted = myGear.Corrupted
                Dim pri As New Price : pri.Exa = 0 : pri.Alch = 0 : pri.Chaos = 0 : pri.GCP = 0
                newFullItem.Price = pri
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
                Dim blIndexed As Boolean = False
                If FullInventoryCache.Count <> 0 Then
                    Dim q As IEnumerable(Of FullItem) = FullInventoryCache.Where(Function(s) s.ID = myGear.UniqueIDHash And s.Name = myGear.Name)
                    If q.Count <> 0 Then
                        Dim testMod As New CloneableList(Of FullMod)
                        For Each m As CloneableList(Of FullMod) In {q(0).ExplicitPrefixMods, q(0).ExplicitSuffixMods}
                            For Each myMod As FullMod In m
                                Dim newMod As New FullMod
                                Dim qry As IEnumerable(Of FullMod) = testMod.Where(Function(x) x.Type1 = myMod.Type1)
                                If qry.Count = 0 Then
                                    newMod.Type1 = myMod.Type1
                                    newMod.Value1 = myMod.Value1
                                    newMod.MaxValue1 = myMod.MaxValue1
                                    testMod.Add(newMod)
                                Else
                                    testMod(testMod.IndexOf(qry(0))).Value1 += myMod.Value1
                                End If
                                If myMod.Type2 <> "" Then
                                    Dim newMod2 As New FullMod
                                    qry = testMod.Where(Function(x) x.Type1 = myMod.Type2)
                                    If qry.Count = 0 Then
                                        newMod2.Type1 = myMod.Type2
                                        newMod2.Value1 = myMod.Value2
                                        testMod.Add(newMod2)
                                    Else
                                        testMod(testMod.IndexOf(qry(0))).Value1 += myMod.Value2
                                    End If
                                End If
                            Next
                        Next
                        For Each m As String In myGear.Explicitmods
                            Dim strName As String = GetChars(m)
                            Dim sngValue As Single = GetNumeric(m)
                            Dim qry As IEnumerable(Of FullMod) = testMod.Where(Function(x) x.Type1 = strName)
                            If qry.Count = 0 Then
                                GoTo IndexMod
                            Else
                                If testMod(testMod.IndexOf(qry(0))).Value1 <> sngValue Then
                                    If m.Contains("-") Then
                                        If testMod(testMod.IndexOf(qry(0))).Value1 <> GetNumeric(m, 0, m.IndexOf("-", StringComparison.OrdinalIgnoreCase)) Then GoTo IndexMod
                                        If testMod(testMod.IndexOf(qry(0))).MaxValue1 <> GetNumeric(m, m.IndexOf("-", StringComparison.OrdinalIgnoreCase), m.Length) Then GoTo IndexMod
                                    Else
                                        GoTo IndexMod
                                    End If
                                End If
                            End If
                        Next
                        newFullItem.Rank = q(0).Rank
                        newFullItem.Percentile = q(0).Percentile
                        newFullItem.ExplicitPrefixMods = q(0).ExplicitPrefixMods.Clone
                        newFullItem.ExplicitSuffixMods = q(0).ExplicitSuffixMods.Clone
                        newFullItem.OtherSolutions = q(0).OtherSolutions
                        FullInventory.Add(newFullItem)
                        If newFullItem.OtherSolutions = True Then
                            Dim q2 As IEnumerable(Of FullItem) = TempInventoryCache.Where(Function(s) s.ID = myGear.UniqueIDHash And s.Name = myGear.Name)
                            For Each qItem In q2
                                Dim tmpItem As FullItem = CType(qItem.Clone, FullItem)
                                TempInventory.Add(tmpItem)
                            Next
                            lngCount = TempInventory.Count
                        End If
                        blIndexed = True
                    End If
                End If
                If blIndexed = True Then GoTo IncrementProgressBar
IndexMod:
                ' Make sure that increased item rarity comes last, since it can be either a prefix or a suffix
                ReorderExplicitMods(myGear.Explicitmods, myGear.Explicitmods.Count)
                blAddedOne = False
                If Not IsNothing(myGear.Explicitmods) Then
                    blSolomonsJudgment.Clear()
                    EvaluateExplicitMods(myGear.Explicitmods, myGear.Explicitmods.Count, myGear.UniqueIDHash.ToString, myGear.Name, newFullItem, TempInventory)
                    If newFullItem.ExplicitPrefixMods.Count = 0 And newFullItem.ExplicitSuffixMods.Count = 0 Then
                        If TempInventory.Count = 0 Then
                            EvaluateExplicitMods(myGear.Explicitmods, myGear.Explicitmods.Count, myGear.UniqueIDHash.ToString, myGear.Name, newFullItem, TempInventory, True, True)       ' We didn't find anything, so run again, allowing legacy values to be set
                        Else
                            If TempInventory(TempInventory.Count - 1).Name <> myGear.Name Then
                                EvaluateExplicitMods(myGear.Explicitmods, myGear.Explicitmods.Count, myGear.UniqueIDHash.ToString, myGear.Name, newFullItem, TempInventory, True, True)       ' We didn't find anything, so run again, allowing legacy values to be set
                            End If
                        End If
                    End If
                    If blAddedOne = True Then
                        ' First rename the associated entry in the RankExplanation dictionary...(have to add it in with the proper name and then remove the old one)
                        'RankExplanation.Add(myGear.UniqueIDHash & myGear.Name, RankExplanation(myGear.UniqueIDHash & myGear.Name & TempInventory.Count - 1))
                        'RankExplanation.Remove(myGear.UniqueIDHash & myGear.Name & TempInventory.Count - 1)
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
IncrementProgressBar:
                statusCounter += 1 : If statusCounter Mod 100 = 0 Then
                    If FullInventoryCache.Count = 0 Then
                        statusController.DisplayMessage("Indexed " & statusCounter & " rare items and " & lngModCounter & " mods...")
                    Else
                        If lngModCounter <> 0 Then
                            statusController.DisplayMessage("Indexed " & statusCounter & " rare items using local cache, plus indexed " & lngModCounter & " mods for new/updated items...")
                        Else
                            statusController.DisplayMessage("Indexed " & statusCounter & " rare items using local cache...")
                        End If
                    End If
                End If
                ' Enable the following to limit the item loading to 100 items when quick development cycles are required
                'If statusCounter = 100 Then Exit Sub
            Next
        Catch ex As Exception
            ErrorHandler(System.Reflection.MethodBase.GetCurrentMethod.Name, ex)
        End Try
    End Sub

    Public Sub ReorderExplicitMods(lstMods As List(Of String), intModCount As Integer)
        Try
            Dim intCounter As Integer = intModCount - 1, strType As String = ""
            If Not IsNothing(lstMods) Then
                For i = 0 To intModCount - 1      ' Run the loop to put increased rarity at the end first, since it can be a prefix or suffix, and is the most important (Note: must go to the end of the loop, to decrement bytCounter)
                    strType = GetChars(lstMods(i))
                    If strType.CompareMultiple(StringComparison.OrdinalIgnoreCase, "% increased Rarity of Items found") = True Then
                        Dim strTemp As String = lstMods(intCounter)
                        lstMods(intCounter) = lstMods(i)
                        lstMods(i) = strTemp
                        intCounter -= 1
                        Exit For
                    End If
                Next
                For i = 0 To intModCount - 1      ' Put combined mods -- and mods that get dragged into the combinations -- at the end, so that our looping search goes quicker
                    If i >= intCounter Then Exit For
                    ' Use a do loop to reorder, since the mod swap might initially exchange a mod at the last spot that is also in this list
                    Do While GetChars(lstMods(i)).CompareMultiple(StringComparison.OrdinalIgnoreCase, "+ to Accuracy Rating", "% increased Block and Stun Recovery", _
                                               "% increased Accuracy Rating", "+ to maximum Mana", "% increased Light Radius", "% increased Armour", _
                                               "% increased Armour and Energy Shield", "% increased Armour and Evasion", "% increased Energy Shield", _
                                               "% increased Evasion and Energy Shield", "% increased Evasion Rating", "% increased Physical Damage", _
                                               "% increased Spell Damage") = True
                        Dim strTemp As String = lstMods(intCounter)
                        lstMods(intCounter) = lstMods(i)
                        lstMods(i) = strTemp
                        If i >= intCounter Then Exit For
                        intCounter -= 1
                    Loop
                Next
            End If
        Catch ex As Exception
            ErrorHandler(System.Reflection.MethodBase.GetCurrentMethod.Name, ex)
        End Try
    End Sub

    'Public Function CheckForCombinedMod(result() As DataRow, j As Integer, newFullMod As FullMod, myGear As Gear, i As Integer) As Byte
    '    ' This function looks to see if the mod entry selected from weights-*.csv is a combined mod, and if successful will return the 
    '    ' index/position of the other mod from the explicitmods list
    '    Dim strMod As String = ""
    '    If newFullMod.Type1 Is result(j)("ExportField") Then
    '        strMod = result(j)("ExportField2").ToString
    '    Else
    '        strMod = result(j)("ExportField").ToString
    '    End If
    '    For k = i To myGear.Explicitmods.Count - 1  ' Look ahead at upcoming mods to find the position for its "companion" mod
    '        If GetChars(myGear.Explicitmods(k)) = strMod Then Return CByte(k)
    '    Next
    '    Return 0
    'End Function

    Public Sub EvaluateExplicitMods(lstMods As List(Of String), intModCount As Integer, strID As String, strName As String, ByRef newfullitem As FullItem, ByRef lstTempInventory As List(Of FullItem), Optional blForceFullSearch As Boolean = False, Optional blAllowLegacy As Boolean = False)
        Try
            Dim result() As DataRow = Nothing
            Dim ModList As New List(Of DataRow), ModPos As New Dictionary(Of String, Integer)
            Dim strField As String = "", strField2 As String = "", strAffix As String = ""
            Dim blCombinedModsAdded As Boolean = False
            For i = 0 To intModCount - 1
                strField = GetChars(lstMods(i))
                ModPos.Add(strField.ToLower, i)
                result = dtWeights.Select("ExportField = '" & strField & "' OR ExportField2 = '" & strField & "'")
                For j = 0 To result.Count - 1
                    If result(j)("ExportField2").ToString <> "" Then ' Do we need to check the second mod for a combined mod?
                        strField2 = IIf(strField = result(j)("ExportField2").ToString, result(j)("ExportField"), result(j)("ExportField2")).ToString
                        If lstMods.Find(Function(x) GetChars(x) = strField2) = "" Then Continue For ' The second mod isn't part of this item...move on
                        blCombinedModsAdded = True  ' If we added a combined mod, will want to run the 'exhaustive' search for more below, otherwise skip the second search
                    End If
                    Dim blResult As Boolean
                    If strField = "% increased Rarity of Items found" And GetNumeric(lstMods(i)) > 9 Then   ' Rarity has both prefix and suffix values, so add both to the mod list
                        blCombinedModsAdded = True      ' Update: Rather than trying to solve the whole thing here, just set our search to try both and let the recursion routine do the rest
                        Dim tmpResult As DataRow = DeepCopyDataRow(result(j))   ' Cannot do a shallow reference copy, or else we are unable to create separate ModList entries
                        blResult = AddToModList(ModList, newfullitem, tmpResult, result(j)("Description").ToString & ",Prefix")     ' We must distinguish our key names to provide the search routine a way to treat this as a "combined" mod
                        Dim tmpResult2 As DataRow = DeepCopyDataRow(result(j))  ' Do it again, as we don't want to affect dtWeights either
                        blResult = AddToModList(ModList, newfullitem, tmpResult2, result(j)("Description").ToString & ",Suffix")
                    Else
                        blResult = AddToModList(ModList, newfullitem, result(j))
                    End If
                    If blResult = False Then Exit Sub
                Next
                If result.Count = 0 Then    ' No match, that is strange...
                    MsgBox("Warning: could not find '" & strField & "' mod in weights list. Please check that your weights-*.csv file is properly configured.")
                    Exit Sub
                End If
            Next
            ' Our ModList now contains all of the possible "combined" mods that could have been assigned to this item (we have exhausted all of the branches)
            Dim ModStatsList As New Dictionary(Of String, DataRow)    ' Tranform the ModsList into one that includes the full value ranges and stats for the mod
            Dim MaxIndex As New Dictionary(Of String, Integer)     ' This is used for the more complex search method
            Dim RemoveFromModList As New List(Of Integer)
            For Each row In ModList
                Dim strOverrideDescription As String = ""
                If row("Description").ToString.Contains("Base Item Found Rarity +%") Then
                    'Select Case GetNumeric(lstMods(ModPos(row("ExportField").ToString.ToLower)))
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
                        strOverrideDescription = row("Description").ToString.Split(CChar(","))(0)
                        strAffix = row("Description").ToString.Split(CChar(","))(1)
                    End If
                End If
                Dim tempModResult As DataRow() = RunModResultQuery(newfullitem, row, strOverrideDescription, strAffix)
                For j = 0 To tempModResult.Count - 1    ' The query returns all of the possible value ranges for this mod, restricted only by level (max)
                    ModStatsList.Add(row("ExportField").ToString.ToLower & row("ExportField2").ToString.ToLower & j & IIf(strAffix <> "", "," & strAffix, "").ToString, tempModResult(j))
                Next
                If tempModResult.Count = 0 Then
                    ' We will remove this row from the modlist, since it doesn't return any results
                    RemoveFromModList.Add(ModList.IndexOf(row))
                Else
                    MaxIndex.Add(row("ExportField").ToString.ToLower & IIf(strAffix <> "", "," & strAffix, "").ToString & row("ExportField2").ToString.ToLower, tempModResult.Count - 1)
                End If
            Next
            If RemoveFromModList.Count <> 0 Then
                For i = 0 To RemoveFromModList.Count - 1
                    ModList.RemoveAt(RemoveFromModList(i))
                Next
                RemoveFromModList.RemoveRange(0, RemoveFromModList.Count)
            End If
            'Dim sngRank As Single = 0
            ' Use simple method if all the mods are independent and a 1-1 mapping exists, so we can set them in a quick, efficient and straightforward manner
            ' Also check the global boolean blForceFullSearch to see if we've called ourselves again because this method won't work
            If (blCombinedModsAdded = False And MaxIndex.Count = ModList.Count) And blForceFullSearch = False Then
                For i = 0 To intModCount - 1
                    lngModCounter += 1
                    Dim newMod As New FullMod
                    newMod.FullText = lstMods(i)
                    newMod.Type1 = GetChars(lstMods(i))
                    If lstMods(i).IndexOf("-", StringComparison.OrdinalIgnoreCase) > -1 Then
                        newMod.Value1 = GetNumeric(lstMods(i), 0, lstMods(i).IndexOf("-", StringComparison.OrdinalIgnoreCase))
                        newMod.MaxValue1 = GetNumeric(lstMods(i), lstMods(i).IndexOf("-", StringComparison.OrdinalIgnoreCase), lstMods(i).Length)
                    Else
                        newMod.Value1 = GetNumeric(lstMods(i))
                    End If
                    newMod.Weight = CInt(ModList(i)("Weight").ToString)
                    'Dim temprow As New Dictionary(Of String, DataRow)
                    Dim temprow = From mymod In ModStatsList Where mymod.Key.ToLower.Contains(newMod.Type1.ToLower) And CSng(mymod.Value("MinV")) <= newMod.Value1 And CSng(mymod.Value("MaxV")) >= newMod.Value1
                    If newMod.MaxValue1 > 0 Then
                        temprow = From mymod In ModStatsList Where mymod.Key.ToLower.Contains(newMod.Type1.ToLower) And CSng(mymod.Value("MinV")) <= newMod.Value1 And CSng(mymod.Value("MinV2")) >= newMod.Value1 _
                                    And CSng(mymod.Value("MaxV2")) <= newMod.MaxValue1 And CSng(mymod.Value("MaxV")) >= newMod.MaxValue1
                    End If
                    If temprow.Count = 0 Then
                        newMod.UnknownValues = True ' This might be a legacy mod that's outside of the ranges in mods.csv
                        strAffix = RunModResultQuery(newfullitem, , ModList(i)("Description").ToString)(0)("Prefix/Suffix").ToString
                        'Dim sngMaxV As Single = CSng(ModStatsList(newMod.Type1.ToLower & MaxIndex(newMod.Type1.ToLower))("MaxV"))
                        'RankUnknownValueMod(newMod, sngMaxV, MaxIndex(newMod.Type1.ToLower), strID, strName, strAffix)
                        GoTo AddMod
                    End If
                    Dim strKey As String = temprow(temprow.Count - 1).Key.ToLower     ' Choose the last row to get the highest possible level, just like the combined algorithm does
                    newMod.BaseLowerV1 = CSng(ModStatsList(strKey)("MinV").ToString)
                    If newMod.MaxValue1 <> 0 And newMod.MaxValue1 <> Nothing Then ' This is a mod with range (i.e. 12-16 damage)
                        newMod.BaseLowerMaxV1 = CSng(ModStatsList(strKey)("MaxV2").ToString)
                        newMod.BaseUpperV1 = CSng(ModStatsList(strKey)("MinV2").ToString)
                        newMod.BaseUpperMaxV1 = CSng(ModStatsList(strKey)("MaxV").ToString)
                    Else
                        newMod.BaseUpperV1 = CSng(ModStatsList(strKey)("MaxV").ToString)
                    End If
                    newMod.MiniLvl = CSng(ModStatsList(strKey)("Level").ToString)
                    strAffix = ModStatsList(strKey)("Prefix/Suffix").ToString
                    newMod.ModLevelActual = CInt(GetNumeric(strKey))
                    newMod.ModLevelMax = MaxIndex(newMod.Type1.ToLower & newMod.Type2.ToLower)
                    'sngRank += CalculateRank(newMod, MaxIndex(newMod.Type1.ToLower & newMod.Type2.ToLower), strKey, strID, strName, strAffix)
                    'Dim sb As New System.Text.StringBuilder
                    'sb.Append(vbCrLf & "(" & String.Format("{0:'+'0;'-'0}", sngRank) & ") " & IIf(i = intModCount - 1, "Final Rank", " (running total)" & vbCrLf).ToString)
                    'AddExplanation(strID & strName, sb)

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
                    'RankExplanation.Remove(strID & strName)
                    EvaluateExplicitMods(lstMods, intModCount, strID, strName, newfullitem, lstTempInventory, True)   ' Call ourselves again, this time with the forcefullsearch boolean set to true
                    Exit Sub
                End If
                'newfullitem.Rank = sngRank
                newfullitem.Percentile = CSng(CalculatePercentile(newfullitem).ToString("0.0"))
                Exit Sub
            End If
            ' Note: Combined Mod Search Method begins here!
            ' If this is a combined mod, then we'll have to use a less efficient method.
            Dim maxpos(ModList.Count - 1) As Integer, curpos(ModList.Count - 1) As Integer, intPos As Integer = 0
            For Each mykey In MaxIndex
                maxpos(intPos) = mykey.Value
                intPos += 1
            Next
            DynamicMultiDimLoop(curpos, maxpos, ModList, ModStatsList, 0, lstMods, intModCount, strID, strName, MaxIndex, ModPos, newfullitem, blAllowLegacy, lstTempInventory)
        Catch ex As Exception
            ErrorHandler(System.Reflection.MethodBase.GetCurrentMethod.Name, ex, "Item Name: " & strName)
        End Try
    End Sub

    Public Function AddToModList(ModList As List(Of DataRow), newfullitem As FullItem, MyRow As DataRow, Optional strForceDescription As String = "") As Boolean
        Try
            If Not ModList.Contains(MyRow) Then ' Before adding it, see if it already is in the list
                Dim tmpResult() As DataRow = RunModResultQuery(newfullitem, MyRow)
                If tmpResult.Count > 0 Then
                    ModList.Add(MyRow)
                    If strForceDescription <> "" Then ModList(ModList.Count - 1)("Description") = strForceDescription
                End If
            End If
            Return True
        Catch ex As Exception
            ErrorHandler(System.Reflection.MethodBase.GetCurrentMethod.Name, ex, "New Mod Description: " & MyRow("Description").ToString & vbCrLf & _
                         "Current ModList Contents: " & String.Join(vbCrLf, ModList.[Select](Function(r) r("Description").ToString()).ToArray()))
            Return False
        End Try
    End Function

    Public Function CalculateRank(newMod As FullMod, intMaxIndex As Integer, strKey As String, strID As String, strName As String, strAffix As String, Optional strOverrideKey As String = "") As Single
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
            'newMod.ModLevelActual = CInt(GetNumeric(strKey))
            'newMod.ModLevelMax = intMaxIndex

            If newMod.Weight < 0 Then GoTo AddExplanationGoto

            Dim intLevelRank As Integer = CInt((intMaxIndex - GetNumeric(strKey)) * 10)       ' The amount lost because of not hitting the highest possible level for the mod
            sb.Append(vbCrLf & String.Format("{0:'+'0;'-'0;'-'0}", intLevelRank * -1) & vbTab & "(max mod level - mod level) * 10 = (" & intMaxIndex & " - " & _
                      GetNumeric(strKey) & ") * 10 = " & intLevelRank)

            Dim intValueRank1 As Integer = 0                                                    ' The amount lost for not getting the maximum value for the first mod
            If newMod.BaseUpperV1 <> newMod.BaseLowerV1 Then
                If newMod.MaxValue1 <> 0 And newMod.BaseUpperMaxV1 <> newMod.BaseLowerMaxV1 Then
                    Dim intValuePart1 As Integer = 0, intValuePart2 As Integer = 0
                    intValuePart1 = CInt(Math.Round(((newMod.BaseUpperV1 - newMod.Value1) / (newMod.BaseUpperV1 - newMod.BaseLowerV1)) * bytModWeight1))
                    sb.Append(vbCrLf & String.Format("{0:'+'0;'-'0;'-'0}", intValuePart1 * -1) & vbTab & _
                              (newMod.BaseUpperV1 - newMod.Value1).ToString("0.#") & " below lower max of " & newMod.BaseUpperV1.ToString("0.#") & _
                              " (range: " & (newMod.BaseUpperV1 - newMod.BaseLowerV1).ToString("0.#") & ", weight: " & bytModWeight1.ToString("0.#") & ")")
                    sb.Append(vbCrLf & vbTab & "weight*(max-value)/range = " & bytModWeight1.ToString("0.#") & "*(" & _
                              newMod.BaseUpperV1 & "-" & newMod.Value1.ToString("0.#") & ")/" & (newMod.BaseUpperV1 - newMod.BaseLowerV1).ToString("0.#") & _
                              "=" & intValuePart1.ToString("0.#"))
                    intValuePart2 = CInt(Math.Round(((newMod.BaseUpperMaxV1 - newMod.MaxValue1) / (newMod.BaseUpperMaxV1 - newMod.BaseLowerMaxV1)) * bytModWeight2))
                    sb.Append(vbCrLf & String.Format("{0:'+'0;'-'0;'-'0}", intValuePart2 * -1) & vbTab & _
                              newMod.BaseUpperMaxV1 - newMod.MaxValue1 & " below upper max of " & newMod.BaseUpperMaxV1.ToString("0.#") & _
                              " (range: " & newMod.BaseUpperMaxV1 - newMod.BaseLowerMaxV1 & ", weight: " & bytModWeight2 & ")")
                    sb.Append(vbCrLf & vbTab & "weight*(max-value)/range = " & bytModWeight2.ToString("0.#") & "*(" & _
                              newMod.BaseUpperMaxV1 & "-" & newMod.MaxValue1.ToString("0.#") & ")/" & (newMod.BaseUpperMaxV1 - newMod.BaseLowerMaxV1).ToString("0.#") & _
                              "=" & intValuePart2.ToString("0.#"))
                    intValueRank1 = intValuePart1 + intValuePart2
                Else
                    intValueRank1 = CInt(Math.Round(((newMod.BaseUpperV1 - newMod.Value1) / (newMod.BaseUpperV1 - newMod.BaseLowerV1)) * bytModWeight1))
                    sb.Append(vbCrLf & String.Format("{0:'+'0;'-'0;'-'0}", intValueRank1 * -1) & vbTab & (newMod.BaseUpperV1 - newMod.Value1).ToString("0.#") & " below max of " & newMod.BaseUpperV1.ToString("0.#") & " (range: " & (newMod.BaseUpperV1 - newMod.BaseLowerV1).ToString("0.#") & ", weight: " & bytModWeight1.ToString("0.#") & ")")
                    sb.Append(vbCrLf & vbTab & "weight*(max-value)/range = " & bytModWeight1.ToString("0.#") & "*(" & newMod.BaseUpperV1.ToString("0.#") & "-" & newMod.Value1.ToString("0.#") & ")/" & (newMod.BaseUpperV1 - newMod.BaseLowerV1).ToString("0.#") & "=" & intValueRank1.ToString("0.#"))
                End If
            Else
                If newMod.MaxValue1 <> 0 And newMod.BaseUpperMaxV1 <> newMod.BaseLowerMaxV1 Then
                    sb.Append(vbCrLf & String.Format("{0:'+'0;'-'0;'-'0}", 0) & vbTab & "Value at max of " & newMod.BaseUpperV1 & " (range: 0)")
                    intValueRank1 = CInt(Math.Round(((newMod.BaseUpperMaxV1 - newMod.MaxValue1) / (newMod.BaseUpperMaxV1 - newMod.BaseLowerMaxV1)) * bytModWeight2))
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
                    intValueRank2 = CInt(Math.Round(((newMod.BaseUpperV2 - newMod.Value2) / (newMod.BaseUpperV2 - newMod.BaseLowerV2)) * bytModWeight2))
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
            AddExplanation(IIf(strOverrideKey = "", strID & strName, strOverrideKey).ToString, sb)
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
            Dim lngNumConfigurationsBelow As Long = 0, lngNumConfigurations As Long = 0, intNumMods As Integer = 0, sngExtrConfigurations As Single = 0
            For Each myModList In {MyItem.ExplicitPrefixMods, MyItem.ExplicitSuffixMods}
                For Each myMod In myModList
                    Dim result() As DataRow = Nothing
                    result = dtWeights.Select("ExportField = '" & myMod.Type1 & "'" & IIf(myMod.Type2 <> "", " AND ExportField2 = '" & myMod.Type2 & "'", " AND ExportField2 = ''").ToString)
                    If result.Count <> 0 Then
                        For Each row In result
                            For Each modRow In RunModResultQuery(MyItem, row)
                                If myMod.BaseUpperMaxV1 <> 0 Then
                                    If CSng(modRow("MinV")) < myMod.Value1 Then lngNumConfigurationsBelow += CLng(CSng(modRow("MinV2")) - CSng(modRow("MinV")) + 1 - CSng(IIf(myMod.Value1 <= CSng(modRow("MinV2")), CSng(modRow("MinV2")) - myMod.Value1 + 1, 0)))
                                    lngNumConfigurations += CLng(CSng(modRow("MinV2")) - CSng(modRow("MinV")) + 1)
                                    If CSng(modRow("MaxV2")) < myMod.MaxValue1 Then lngNumConfigurationsBelow += CLng(CSng(modRow("MaxV2")) - CSng(modRow("MaxV")) + 1 - CSng(IIf(myMod.MaxValue1 <= CSng(modRow("MaxV2")), CSng(modRow("MaxV")) - myMod.MaxValue1 + 1, 0)))
                                    lngNumConfigurations += CLng(CSng(modRow("MaxV2")) - CSng(modRow("MaxV")) + 1)
                                Else
                                    If CSng(modRow("MinV")) < myMod.Value1 Then lngNumConfigurationsBelow += CLng(CSng(modRow("MaxV")) - CSng(modRow("MinV")) + 1 - CSng(IIf(myMod.Value1 < CSng(modRow("MaxV")), CSng(modRow("MaxV")) - myMod.Value1 + 1, 0)))
                                    lngNumConfigurations += CLng(CSng(modRow("MaxV")) - CSng(modRow("MinV")) + 1)
                                End If
                                If myMod.Value2 <> 0 Then
                                    If CSng(modRow("MinV2")) < myMod.Value2 Then lngNumConfigurationsBelow += CLng(CSng(modRow("MaxV2")) - CSng(modRow("MinV2")) + 1 - CSng(IIf(myMod.Value2 < CSng(modRow("MaxV2")), CSng(modRow("MaxV2")) - myMod.Value2 + 1, 0)))
                                    lngNumConfigurations += CLng(CSng(modRow("MaxV2")) - CSng(modRow("MinV2")) + 1)
                                End If
                            Next
                        Next
                    End If
                    intNumMods += 1
                Next
            Next
            'If bytNumMods < 6 Then sngExtrConfigurations = (6 - bytNumMods) * (lngNumConfigurations / intNumMods) ' Use a calculation of the average number of configs for the missing mods
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

    Public Function RankUnknownValueMod(newMod As FullMod, sngMaxV As Single, intMaxIndex As Integer, strID As String, strName As String, strAffix As String, Optional strOverridKey As String = "") As Single
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
                    Dim intLevelRank As Integer = ((intMaxIndex + 1) * 10)       ' The amount lost because of not hitting the highest possible level for the mod
                    sbTemp.Append(vbCrLf & String.Format("{0:'+'0;'-'0;'-'0}", intLevelRank * -1) & vbTab & "(max mod level + 1) * 10 = (" & intMaxIndex + 1 & " * 10) = " & intLevelRank)
                    RankUnknownValueMod -= Math.Min(newMod.Weight * 10 + 10, intLevelRank)
                    If newMod.Weight * 10 + 10 < intLevelRank Then sbTemp.Append(vbCrLf & vbCrLf & "-10 (capped)")
                End If
            Else
                RankUnknownValueMod += newMod.Weight * 10
                sbTemp.Append(vbCrLf & String.Format("{0:'+'0;'-'0}", newMod.Weight * 10) & vbTab & "(mod weight * 10) = (" & newMod.Weight & " * 10) = " & newMod.Weight * 10)
            End If
            sbTemp.Append(vbCrLf & vbCrLf & "(" & String.Format("{0:'+'0;'-'0}", RankUnknownValueMod) & ") (running total)" & vbCrLf)
            AddExplanation(IIf(strOverridKey = "", strID & strName, strOverridKey).ToString, sbTemp)
        Catch ex As Exception
            ErrorHandler(System.Reflection.MethodBase.GetCurrentMethod.Name, ex)
            Return 0
        End Try
    End Function

    Public Sub DynamicMultiDimLoop(ByRef curpos() As Integer, ByRef maxpos() As Integer, ByRef ModList As List(Of DataRow), ModStatsList As Dictionary(Of String, DataRow), curdim As Integer, lstMods As List(Of String), intModCount As Integer, strID As String, strName As String, MaxIndex As Dictionary(Of String, Integer), modPos As Dictionary(Of String, Integer), ByRef newFullItem As FullItem, blAllowLegacy As Boolean, lstTempInventory As List(Of FullItem))
        Try
            For i = maxpos(curdim) To -1 Step -1    ' We go to -1, since -1 is the index where we don't use the potential mod
                curpos(curdim) = i
                Dim strStatsKey As String = "", strMod As String = "", strMod2 As String = ""
                If i >= 0 Then
                    strStatsKey = ModList(curdim)("ExportField").ToString.ToLower & ModList(curdim)("ExportField2").ToString.ToLower & i
                    If ModStatsList.ContainsKey(strStatsKey) = False Then   ' It must be a rarity mod that could be a shared total of both a prefix and a suffix
                        strStatsKey += "," & ModList(curdim)("Description").ToString.Split(CChar(","))(1)
                    End If
                End If
                ' If the 'min' value for this mod is above the actual value, then skip it and go down a level in order to trim our search space
                ' If the 'max' value for this mod is below the actual value, AND it is not a combined mod (meaning there is no other way to satisfy the value), then go up a level and trim some more
                ' The i>0 condition means that we will try at least once, so that we will set an "UnknownValue=True" if we're having trouble with a legacy mod/value
                If ModList(curdim)("ExportField2").ToString = "" And i > 0 Then
                    Dim intPos As Integer = modPos(ModList(curdim)("ExportField").ToString.ToLower)
                    strMod = lstMods(intPos)
                    If strMod.IndexOf("-", StringComparison.OrdinalIgnoreCase) > -1 Then   ' We have to check a ranged value for both minimums
                        If GetNumeric(strMod, 0, strMod.IndexOf("-", StringComparison.OrdinalIgnoreCase)) < CSng(ModStatsList(strStatsKey)("MinV")) Then Continue For
                        If GetNumeric(strMod, strMod.IndexOf("-", StringComparison.OrdinalIgnoreCase), strMod.Length) < CSng(ModStatsList(strStatsKey)("MaxV2")) Then Continue For
                        If GetNumeric(strMod, 0, strMod.IndexOf("-", StringComparison.OrdinalIgnoreCase)) > CSng(ModStatsList(strStatsKey)("MinV2")) And blAddedOne = True Then Exit For
                        If GetNumeric(strMod, strMod.IndexOf("-", StringComparison.OrdinalIgnoreCase), strMod.Length) > CSng(ModStatsList(strStatsKey)("MaxV")) And blAddedOne = True Then Exit For
                    Else
                        If GetNumeric(strMod) < CSng(ModStatsList(strStatsKey)("MinV")) Then Continue For
                        If HaveToShare(GetChars(strMod).ToLower, ModList, curpos) = False Then
                            If GetNumeric(strMod) > CSng(ModStatsList(strStatsKey)("MaxV")) And blAddedOne = True Then Exit For
                        End If
                    End If
                ElseIf i > 0 Then
                    Dim intPos As Integer = modPos(ModList(curdim)("ExportField").ToString.ToLower)
                    strMod = lstMods(intPos)
                    Dim intPos2 As Integer = modPos(ModList(curdim)("ExportField2").ToString.ToLower)
                    strMod2 = lstMods(intPos2)
                    If GetNumeric(strMod) < CSng(ModStatsList(strStatsKey)("MinV")) Then Continue For
                    If GetNumeric(strMod2) < CSng(ModStatsList(strStatsKey)("MinV2")) Then Continue For
                End If

                If curdim < UBound(maxpos) Then
                    ' Recursively call this sub again to repeat the loop for the next dimension of the list/array
                    DynamicMultiDimLoop(curpos, maxpos, ModList, ModStatsList, curdim + 1, lstMods, intModCount, strID, strName, MaxIndex, modPos, newFullItem, blAllowLegacy, lstTempInventory)
                Else
                    ' We have a possible match!
                    Dim intPos(MaxIndex.Count - 1, 1) As Integer, blmatch As Boolean = True
                    Dim ModMinTotals(intModCount - 1) As Single, ModMaxTotals(intModCount - 1) As Single
                    For j = 0 To MaxIndex.Count - 1
                        If curpos(j) = -1 Then Continue For ' If position is -1, we are trying to get the proper totals without this mod
                        Dim strTempStatKey As String = ModList(j)("ExportField").ToString.ToLower & ModList(j)("ExportField2").ToString.ToLower & curpos(j)
                        If ModList(j)("Description").ToString.Contains(",") = True Then
                            strTempStatKey += "," & ModList(j)("Description").ToString.Split(CChar(","))(1)
                        End If
                        intPos(j, 0) = modPos(ModList(j)("ExportField").ToString.ToLower)
                        ModMinTotals(intPos(j, 0)) += CSng(ModStatsList(strTempStatKey)("MinV"))
                        ModMaxTotals(intPos(j, 0)) += CSng(ModStatsList(strTempStatKey)("MaxV"))
                        If ModList(j)("ExportField2").ToString <> "" Then
                            intPos(j, 1) = modPos(ModList(j)("ExportField2").ToString.ToLower)
                            ModMinTotals(intPos(j, 1)) += CSng(ModStatsList(strTempStatKey)("MinV2"))
                            ModMaxTotals(intPos(j, 1)) += CSng(ModStatsList(strTempStatKey)("MaxV2"))
                        End If
                    Next
                    For j = 0 To ModMinTotals.Count - 1
                        'If curpos(j) = -1 Then Continue For ' If position is -1, we're getting the totals without this mod
                        strMod = lstMods(j)
                        If strMod.IndexOf("-", StringComparison.OrdinalIgnoreCase) > -1 Then
                            If ModMinTotals(j) > GetNumeric(strMod, 0, strMod.IndexOf("-", StringComparison.OrdinalIgnoreCase)) Then blmatch = False : Exit For
                            If ModMaxTotals(j) < GetNumeric(strMod, strMod.IndexOf("-", StringComparison.OrdinalIgnoreCase), strMod.Length) Then blmatch = False : Exit For
                        Else
                            If ModMinTotals(j) > GetNumeric(strMod) Then blmatch = False : Exit For
                            If ModMaxTotals(j) < GetNumeric(strMod) Then
                                If HaveToShare(GetChars(strMod).ToLower, ModList, curpos) = False Then
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
                        'Dim sngRank As Single = 0
                        'Dim strRankExplanationKey As String = strID & strName & lstTempInventory.Count
                        For j = 0 To ModList.Count - 1
                            If curpos(j) = -1 Then Continue For ' If position is -1, this mod position in ModList will not be used
                            Dim newMod As New FullMod, strAffix As String = ""
                            strMod = lstMods(intPos(j, 0))
                            newMod.FullText = strMod
                            newMod.Type1 = GetChars(strMod)
                            newMod.Weight = CInt(ModList(j)("Weight").ToString)
                            Dim strKey As String = "", strMaxIndexKey As String = ""
                            If ModList(j)("Description").ToString.Contains(",") Then
                                strKey = ModList(j)("ExportField").ToString.ToLower & ModList(j)("ExportField2").ToString.ToLower & curpos(j) & "," & ModList(j)("Description").ToString.Split(CChar(","))(1)
                                strMaxIndexKey = ModList(j)("ExportField").ToString.ToLower & ModList(j)("ExportField2").ToString.ToLower & "," & ModList(j)("Description").ToString.Split(CChar(","))(1)
                            Else
                                strKey = ModList(j)("ExportField").ToString.ToLower & ModList(j)("ExportField2").ToString.ToLower & curpos(j)
                                strMaxIndexKey = ModList(j)("ExportField").ToString.ToLower & ModList(j)("ExportField2").ToString.ToLower
                            End If

                            If strMod.IndexOf("-", StringComparison.OrdinalIgnoreCase) > -1 Then
                                newMod.Value1 = GetNumeric(strMod, 0, strMod.IndexOf("-", StringComparison.OrdinalIgnoreCase))
                                newMod.MaxValue1 = GetNumeric(strMod, strMod.IndexOf("-", StringComparison.OrdinalIgnoreCase), strMod.Length)
                            Else
                                If HaveToShare(ModList(j)("ExportField").ToString.ToLower, ModList, curpos) = True Then
                                    ' We only have a portion of this value...have to apply a weighting formula to determine how much is ours and how much the other mod(s) will chip in
                                    newMod.Value1 = DistributeValues(strKey, ModList(j)("ExportField").ToString, GetNumeric(strMod), ModList, ModStatsList, curpos)
                                    ' Have to change the fulltext to reflect the new value
                                    newMod.FullText = BuildFullText(newMod.Type1, newMod.Value1)
                                Else
                                    newMod.Value1 = GetNumeric(strMod)
                                    If GetNumeric(strMod) > CSng(ModStatsList(strKey)("MaxV").ToString) And blAllowLegacy = True Then
                                        newMod.UnknownValues = True ' This might be a legacy mod that's outside of the ranges in mods.csv
                                        If GetChars(strMod).ToLower = "% increased rarity of items found" Then ' Can't use the description field to run ModResultQuery, set affix manually - fix bug 11 
                                            strAffix = ModList(j)("Description").ToString.Split(CChar(","))(1)
                                        Else
                                            strAffix = RunModResultQuery(newFullItem, , ModList(j)("Description").ToString)(0)("Prefix/Suffix").ToString
                                        End If
                                        'Dim sngMaxV As Single = CSng(ModStatsList(newMod.Type1.ToLower & MaxIndex(newMod.Type1.ToLower & ModList(j)("ExportField2").ToString.ToLower))("MaxV"))
                                        'sngRank += RankUnknownValueMod(newMod, sngMaxV, MaxIndex(strMaxIndexKey), strID, strName, strAffix, strRankExplanationKey)
                                        ' Don't jump ahead just yet, we might have an ExportField2 to set
                                    End If
                                End If
                            End If
                            If ModList(j)("ExportField2").ToString <> "" Then
                                strMod2 = lstMods(intPos(j, 1))
                                newMod.Type2 = GetChars(strMod2)
                                If HaveToShare(ModList(j)("ExportField2").ToString.ToLower, ModList, curpos) = True Then
                                    ' We only have a portion of this value...have to apply a weighting formula to determine how much is ours and how much the other mod(s) will chip in
                                    newMod.Value2 = DistributeValues(strKey, ModList(j)("ExportField2").ToString, GetNumeric(strMod2), ModList, ModStatsList, curpos)
                                    If newMod.FullText.IndexOf("/") = -1 Then
                                        newMod.FullText += "/" & BuildFullText(newMod.Type2, newMod.Value2)
                                    Else
                                        ' Have to rebuild the mod values, since they are now shared
                                        newMod.FullText = newMod.FullText.Substring(0, newMod.FullText.IndexOf("/")) & BuildFullText(newMod.Type2, newMod.Value2)
                                    End If
                                Else
                                    newMod.Value2 = GetNumeric(strMod2)
                                    ' This second mod may not be part of the original mod text
                                    If newMod.FullText.IndexOf("/") = -1 Then newMod.FullText += "/" & BuildFullText(newMod.Type2, newMod.Value2)
                                End If
                            End If
                            If newMod.UnknownValues = True Then GoTo AddMod2

                            newMod.BaseLowerV1 = CSng(ModStatsList(strKey)("MinV").ToString)
                            If newMod.MaxValue1 <> 0 And newMod.MaxValue1 <> Nothing Then ' This is a mod with range (i.e. 12-16 damage)
                                newMod.BaseLowerMaxV1 = CSng(ModStatsList(strKey)("MaxV2").ToString)
                                newMod.BaseUpperV1 = CSng(ModStatsList(strKey)("MinV2").ToString)
                                newMod.BaseUpperMaxV1 = CSng(ModStatsList(strKey)("MaxV").ToString)
                            Else
                                newMod.BaseUpperV1 = CSng(ModStatsList(strKey)("MaxV").ToString)
                            End If
                            If ModList(j)("ExportField2").ToString <> "" Then
                                newMod.BaseLowerV2 = CSng(ModStatsList(strKey)("MinV2").ToString)
                                newMod.BaseUpperV2 = CSng(ModStatsList(strKey)("MaxV2").ToString)
                            End If
                            newMod.MiniLvl = CSng(ModStatsList(strKey)("Level").ToString)
                            strAffix = ModStatsList(strKey)("Prefix/Suffix").ToString
                            newMod.ModLevelActual = CInt(GetNumeric(strKey))
                            newMod.ModLevelMax = MaxIndex(strMaxIndexKey)
                            'sngRank += CalculateRank(newMod, MaxIndex(strMaxIndexKey), strKey, strID, strName, strAffix, strRankExplanationKey)
                            'Dim sb As New System.Text.StringBuilder
                            'sb.Append(vbCrLf & "(" & String.Format("{0:'+'0;'-'0}", sngRank) & ")  (running total)" & vbCrLf)
                            'RankExplanation(strRankExplanationKey) += vbCrLf & sb.ToString

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
                            'RankExplanation.Remove(strRankExplanationKey)
                            Continue For
                        End If
                        'newFullItem.Rank = sngRank
                        newFullItem.Percentile = CSng(CalculatePercentile(newFullItem).ToString("0.0"))
                        'RankExplanation(strRankExplanationKey) = RankExplanation(strRankExplanationKey).Substring(0, RankExplanation(strRankExplanationKey).LastIndexOf("(running total)")) & "Final Rank" & vbCrLf
                        If lstTempInventory.IndexOf(newFullItem) = -1 Then
                            Dim tmpItem As FullItem = CType(newFullItem.Clone, FullItem)
                            lstTempInventory.Add(tmpItem)
                        End If
                        newFullItem.ExplicitPrefixMods.RemoveRange(0, newFullItem.ExplicitPrefixMods.Count) ' Remove any mod list changes we might have made
                        newFullItem.ExplicitSuffixMods.RemoveRange(0, newFullItem.ExplicitSuffixMods.Count)
                        blAddedOne = True
                    End If
                End If
                'If blAddedOne = True Then Exit Sub
            Next
        Catch ex As Exception
            ErrorHandler(System.Reflection.MethodBase.GetCurrentMethod.Name, ex, "Item Name: " & strName)
        End Try
    End Sub

    Public Function HaveToShare(strMyMod As String, ModList As List(Of DataRow), curpos() As Integer) As Boolean
        ' Look in ModList to see if there are multiple entries for this mod, but first check curpos to see if the entry is even in consideration for this particular iteration
        Dim blFound As Boolean = False
        For i = 0 To ModList.Count - 1
            If curpos(i) = -1 Then Continue For
            Dim strKey As String = ModList(i)("ExportField").ToString.ToLower & ModList(i)("ExportField2").ToString.ToLower
            If ModList(i)("ExportField").ToString.ToLower = strMyMod.ToLower Or ModList(i)("ExportField2").ToString.ToLower = strMyMod.ToLower Then
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
                strKey = ModList(i)("ExportField").ToString.ToLower & ModList(i)("ExportField2").ToString.ToLower & curpos(i) & "," & ModList(i)("Description").ToString.Split(CChar(","))(1)
            Else
                strKey = ModList(i)("ExportField").ToString.ToLower & ModList(i)("ExportField2").ToString.ToLower & curpos(i)
            End If

            If ModList(i)("ExportField").ToString.ToLower = strMod.ToLower Then
                If strMyKey.ToLower <> strKey.ToLower Then
                    sngTotalRange += CSng(ModStatsList(strKey)("MaxV").ToString) - CSng(ModStatsList(strKey)("MinV").ToString)
                    sngTotalMax += CSng(ModStatsList(strKey)("MaxV").ToString)
                Else
                    sngRange = CSng(ModStatsList(strKey)("MaxV").ToString) - CSng(ModStatsList(strKey)("MinV").ToString)
                    sngMax = CSng(ModStatsList(strKey)("MaxV").ToString)
                    sngTotalRange += sngRange
                    sngTotalMax += sngMax
                End If
            ElseIf ModList(i)("ExportField2").ToString.ToLower = strMod.ToLower Then
                If strMyKey.ToLower <> strKey.ToLower Then
                    sngTotalRange += CSng(ModStatsList(strKey)("MaxV2").ToString) - CSng(ModStatsList(strKey)("MinV2").ToString)
                    sngTotalMax += CSng(ModStatsList(strKey)("MaxV2").ToString())
                Else
                    sngRange = CSng(ModStatsList(strKey)("MaxV2").ToString) - CSng(ModStatsList(strKey)("MinV2").ToString)
                    sngMax = CSng(ModStatsList(strKey)("MaxV2").ToString)
                    sngTotalRange += sngRange
                    sngTotalMax += sngMax
                End If
            End If
        Next
        Dim sngCalculatedValue As Single = sngMax - (sngRange * (sngTotalMax - sngValue) / sngTotalRange)
        If sngCalculatedValue - Math.Truncate(sngCalculatedValue) = 0.5 Then
            If blSolomonsJudgment.Count = 0 Then
                blSolomonsJudgment.Add(strMod, True)
                sngCalculatedValue = CSng(Math.Truncate(sngCalculatedValue))
            Else
                sngCalculatedValue = CSng(Math.Ceiling(sngCalculatedValue))
            End If
        Else
            sngCalculatedValue = CSng(Math.Round(sngCalculatedValue))
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
                Dim intPCount As Integer = 0, intSCount As Integer = 0, intICount As Integer = 0
                Dim row As DataRow = dt.NewRow()
                row("Rank") = it.Rank
                row("%") = it.Percentile
                row("Type") = If(it.GearType, "")
                row("SubType") = If(it.ItemType, "")
                row("Leag") = If(it.League, "")
                row("Location") = If(it.Location, "")
                row("Name") = If(it.Name, "")
                row("Level") = it.Level
                row("Gem") = it.LevelGem
                row("Sokt") = it.Sockets
                row("SktClrs") = it.Colours
                If it.Colours.Count <> 0 Then
                    Dim tempList As List(Of String) = it.Colours.Split(" ").ToList
                    Dim newList As New List(Of String)
                    For Each s As String In tempList
                        s = s.Replace("-", "")
                        Dim c() As Char = s.ToCharArray
                        Array.Sort(c)
                        newList.Add(New String(c))
                    Next
                    row("SktClrsSearch") = String.Join(" ", newList.ToArray())
                Else
                    row("SktClrsSearch") = ""
                End If
                row("Link") = it.Links
                If it.ExplicitPrefixMods Is Nothing = False Then
                    For i = 0 To it.ExplicitPrefixMods.Count - 1
                        row("Prefix " & i + 1) = it.ExplicitPrefixMods(i).FullText
                        row("pft" & i + 1) = it.ExplicitPrefixMods(i).Type1 & IIf(it.ExplicitPrefixMods(i).Type2 <> "", "/" & it.ExplicitPrefixMods(i).Type2, "").ToString
                        row("pt" & i + 1) = it.ExplicitPrefixMods(i).Type1
                        row("pt" & i + 1 & "2") = it.ExplicitPrefixMods(i).Type2
                        row("pv" & i + 1) = it.ExplicitPrefixMods(i).Value1
                        row("pv" & i + 1 & "m") = it.ExplicitPrefixMods(i).MaxValue1
                        row("pv" & i + 1 & "2") = it.ExplicitPrefixMods(i).Value2
                        intPCount += 1
                    Next
                    For i = it.ExplicitPrefixMods.Count To 2    ' Make the other fields non-null, so that searches/filters can do proper comparisons
                        row("Prefix " & i + 1) = ""
                        row("pft" & i + 1) = ""
                        row("pt" & i + 1) = ""
                        row("pt" & i + 1 & "2") = ""
                        row("pv" & i + 1) = 0
                        row("pv" & i + 1 & "m") = 0
                        row("pv" & i + 1 & "2") = 0
                    Next
                End If
                row("pcount") = intPCount
                If it.ExplicitSuffixMods Is Nothing = False Then
                    For i = 0 To it.ExplicitSuffixMods.Count - 1
                        row("Suffix " & i + 1) = it.ExplicitSuffixMods(i).FullText
                        row("sft" & i + 1) = it.ExplicitSuffixMods(i).Type1 & IIf(it.ExplicitSuffixMods(i).Type2 <> "", "/" & it.ExplicitSuffixMods(i).Type2, "").ToString
                        row("st" & i + 1) = it.ExplicitSuffixMods(i).Type1
                        row("st" & i + 1 & "2") = it.ExplicitSuffixMods(i).Type2
                        row("sv" & i + 1) = it.ExplicitSuffixMods(i).Value1
                        row("sv" & i + 1 & "m") = it.ExplicitSuffixMods(i).MaxValue1
                        row("sv" & i + 1 & "2") = it.ExplicitSuffixMods(i).Value2
                        intSCount += 1
                    Next
                    For i = it.ExplicitSuffixMods.Count To 2    ' Make the other fields non-null, so that searches/filters can do proper comparisons
                        row("Suffix " & i + 1) = ""
                        row("sft" & i + 1) = ""
                        row("st" & i + 1) = ""
                        row("st" & i + 1 & "2") = ""
                        row("sv" & i + 1) = 0
                        row("sv" & i + 1 & "m") = 0
                        row("sv" & i + 1 & "2") = 0
                    Next
                End If
                row("scount") = intSCount
                row("ecount") = it.ExplicitPrefixMods.Count + it.ExplicitSuffixMods.Count
                If it.ImplicitMods Is Nothing = False Then
                    For i = 0 To it.ImplicitMods.Count - 1
                        row("Implicit") = row("Implicit").ToString & it.ImplicitMods(i).FullText & ", "
                        row("it" & i + 1) = it.ImplicitMods(i).Type1
                        row("iv" & i + 1) = it.ImplicitMods(i).Value1
                        row("iv" & i + 1 & "m") = it.ImplicitMods(i).MaxValue1
                        intICount += 1
                    Next
                    For i = it.ImplicitMods.Count To 2    ' Make the other fields non-null, so that searches/filters can do proper comparisons
                        row("it" & i + 1) = ""
                        row("iv" & i + 1) = 0
                        row("iv" & i + 1 & "m") = 0
                    Next
                    row("Implicit") = row("Implicit").ToString.Substring(0, Math.Max(2, row("Implicit").ToString.Length) - 2)
                End If
                row("icount") = intICount
                row("Qal") = it.Quality
                row("*") = If(it.OtherSolutions, "*", "")
                row("Crpt") = it.Corrupted
                If IsNothing(it.Price) Then
                    row("Price") = ""
                    row("PriceNum") = 0
                    row("PriceOrb") = ""
                Else
                    If IsNothing(it.Price.Exa) = False Then
                        row("Price") = it.Price.Exa & " " & "Exa"
                        row("PriceNum") = it.Price.Exa
                        row("PriceOrb") = "Exa"
                    ElseIf IsNothing(it.Price.Chaos) = False Then
                        row("Price") = it.Price.Chaos & " " & "Chaos"
                        row("PriceNum") = it.Price.Chaos
                        row("PriceOrb") = "Chaos"
                    ElseIf IsNothing(it.Price.Alch) = False Then
                        row("Price") = it.Price.Alch & " " & "Alch"
                        row("PriceNum") = it.Price.Alch
                        row("PriceOrb") = "Alch"
                    ElseIf IsNothing(it.Price.GCP) = False Then
                        row("Price") = it.Price.GCP & " " & "GCP"
                        row("PriceNum") = it.Price.GCP
                        row("PriceOrb") = "GCP"
                    Else
                        row("Price") = ""
                        row("PriceNum") = 0
                        row("PriceOrb") = ""
                    End If
                End If
                row("Arm") = it.Arm
                row("Eva") = it.Eva
                row("ES") = it.ES
                row("ThreadID") = If(IsNothing(it.ThreadID), "", it.ThreadID)
                row("Index") = myList.IndexOf(it)
                row("ID") = it.ID.ToString
                dt.Rows.Add(row)
                intCounter += 1
                If intCounter Mod 100 = 0 And blShowProgress Then statusController.DisplayMessage(strDisplayText & ": Completed adding " & intCounter & " of " & myList.Count & " rows.")
                If intCounter Mod 10 = 0 And blShowProgress = False Then Me.Invoke(New MyDelegate(AddressOf PBPerformStep))
            Next
            If blShowProgress Then statusController.DisplayMessage(strDisplayText & ": Completed adding all " & myList.Count & " rows.")
        Catch ex As Exception
            ErrorHandler(System.Reflection.MethodBase.GetCurrentMethod.Name, ex)
        End Try
    End Sub

    Private Sub DataGridView1_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If e.RowIndex = -1 Then Exit Sub
        Dim strGearType As String = ""
        If e.ColumnIndex = DataGridView1.Columns("Rank").Index Then
            Dim sb As New System.Text.StringBuilder
            If FullInventory(CInt(DataGridView1.Rows(e.RowIndex).Cells("Index").Value)).LevelGem = True Then AddGemWarning(sb)
            Dim strKey As String = DataGridView1.CurrentRow.Cells("ID").Value.ToString & DataGridView1.CurrentRow.Cells("Name").Value.ToString
            If RankExplanation.ContainsKey(strKey) = False Then
                MessageBox.Show("Could not find explanation for this rank. Please report this to:" & Environment.NewLine & Environment.NewLine & _
                                    "https://github.com/RoryTate/modrank/issues", "Key Not Found", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Exit Sub
            End If
            sb.Append(RankExplanation(DataGridView1.CurrentRow.Cells("ID").Value.ToString & DataGridView1.CurrentRow.Cells("Name").Value.ToString))
            MessageBox.Show(sb.ToString, "Item Mod Rank Explanation - " & DataGridView1.CurrentRow.Cells("Name").Value.ToString, MessageBoxButtons.OK, MessageBoxIcon.Information)
        ElseIf e.ColumnIndex = DataGridView1.Columns("*").Index AndAlso DataGridView1.CurrentRow.Cells("*").Value.ToString = "*" Then
            If Application.OpenForms().OfType(Of frmResults).Any = False Then frmResults.Show(Me)
            frmResults.Text = "Possible Mod Solutions for '" & DataGridView1.CurrentRow.Cells("Name").Value.ToString & "'"
            frmResults.blStore = False
            Dim tmpList As New CloneableList(Of String)
            tmpList.Add(DataGridView1.CurrentRow.Cells("ID").Value.ToString)
            tmpList.Add(DataGridView1.CurrentRow.Cells("Name").Value.ToString)
            frmResults.MyData = tmpList
        ElseIf DataGridView1.Columns(e.ColumnIndex).Name.ToLower.Contains("prefix") AndAlso NotNull(DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value.ToString, "") <> "" Then
            ShowModInfo(DataGridView1, FullInventory(CInt(DataGridView1.Rows(e.RowIndex).Cells("Index").Value)), FullInventory(CInt(DataGridView1.Rows(e.RowIndex).Cells("Index").Value)).ExplicitPrefixMods, CInt(GetNumeric(DataGridView1.Columns(e.ColumnIndex).Name)) - 1, e)
        ElseIf DataGridView1.Columns(e.ColumnIndex).Name.ToLower.Contains("suffix") AndAlso NotNull(DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value.ToString, "") <> "" Then
            ShowModInfo(DataGridView1, FullInventory(CInt(DataGridView1.Rows(e.RowIndex).Cells("Index").Value)), FullInventory(CInt(DataGridView1.Rows(e.RowIndex).Cells("Index").Value)).ExplicitSuffixMods, CInt(GetNumeric(DataGridView1.Columns(e.ColumnIndex).Name)) - 1, e)
        ElseIf DataGridView1.Columns(e.ColumnIndex).Name.ToLower.Contains("prefix") AndAlso NotNull(DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value.ToString, "") = "" Then
            ShowAllPossibleMods(DataGridView1, FullInventory(CInt(DataGridView1.Rows(e.RowIndex).Cells("Index").Value)), FullInventory(CInt(DataGridView1.Rows(e.RowIndex).Cells("Index").Value)).ExplicitPrefixMods, "Prefix")
        ElseIf DataGridView1.Columns(e.ColumnIndex).Name.ToLower.Contains("suffix") AndAlso NotNull(DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value.ToString, "") = "" Then
            ShowAllPossibleMods(DataGridView1, FullInventory(CInt(DataGridView1.Rows(e.RowIndex).Cells("Index").Value)), FullInventory(CInt(DataGridView1.Rows(e.RowIndex).Cells("Index").Value)).ExplicitSuffixMods, "Suffix")
        ElseIf e.ColumnIndex = DataGridView1.Columns("Location").Index Then
            frmLocation.X = FullInventory(CInt(DataGridView1.Rows(e.RowIndex).Cells("Index").Value)).X
            frmLocation.Y = FullInventory(CInt(DataGridView1.Rows(e.RowIndex).Cells("Index").Value)).Y
            frmLocation.H = FullInventory(CInt(DataGridView1.Rows(e.RowIndex).Cells("Index").Value)).H
            frmLocation.W = FullInventory(CInt(DataGridView1.Rows(e.RowIndex).Cells("Index").Value)).W
            frmLocation.TabName = FullInventory(CInt(DataGridView1.Rows(e.RowIndex).Cells("Index").Value)).Location
            frmLocation.ItemName = FullInventory(CInt(DataGridView1.Rows(e.RowIndex).Cells("Index").Value)).Name
            If frmLocation.TabName.EndsWith(" Tab") = False Then
                MessageBox.Show("This item is in a character's inventory and the location cannot be shown.", "Item in Character Inventory", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Else
                frmLocation.ShowDialog(Me)
            End If
        ElseIf DataGridView1.Columns(e.ColumnIndex).Name.CompareMultiple(StringComparison.Ordinal, "Sokt", "Link") Then
            MessageBox.Show(DataGridView1.CurrentRow.Cells("SktClrs").Value, "Socket/Link Info", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    Public Sub ShowModInfo(dg As DataGridView, MyItem As FullItem, MyMod As List(Of FullMod), intIndex As Integer, e As DataGridViewCellEventArgs)
        Dim sb As New System.Text.StringBuilder
        If MyItem.LevelGem = True Then AddGemWarning(sb)
        sb.Append("Mod Text: " & MyMod(intIndex).FullText & vbCrLf)
        sb.Append("Mod Weight: " & MyMod(intIndex).Weight & vbCrLf)
        Dim result() As DataRow = Nothing
        result = dtWeights.Select("ExportField = '" & MyMod(intIndex).Type1 & "'" & IIf(MyMod(intIndex).Type2 <> "", " AND ExportField2 = '" & MyMod(intIndex).Type2 & "'", " AND ExportField2 = ''").ToString)
        sb.Append(vbCrLf & "Possible Mod Levels and Values:" & vbCrLf)
        Dim strAffix As String = dg.Columns(e.ColumnIndex).Name.Substring(0, 6)
        Dim blAddedTopRow As Boolean = False
        If result.Count <> 0 Then
            For Each row In result
                Dim intModLevel As Integer = 0
                For Each ModRow In RunModResultQuery(MyItem, row, , strAffix)
                    If blAddedTopRow = False Then
                        sb.Append("----------" & vbTab & "------------" & vbCrLf)
                        sb.Append("| Level" & vbTab & "| Values" & vbCrLf)
                        sb.Append("----------" & vbTab & "------------" & vbCrLf)
                        blAddedTopRow = True
                    End If
                    If ModRow("Description").ToString <> row("Description").ToString Then Continue For
                    strAffix = ModRow("Prefix/Suffix").ToString
                    sb.Append("| " & ModRow("Level").ToString & IIf(MyMod(intIndex).ModLevelActual = intModLevel And MyMod(intIndex).UnknownValues = False, "(*)", "").ToString & vbTab & "| " & ModRow("Value").ToString & vbCrLf)
                    intModLevel += 1
                Next
            Next
            sb.Append("----------" & vbTab & "------------" & vbCrLf & vbCrLf)
        End If
        If MyMod(intIndex).FullText.Contains("Life Regenerated") Then
            sb.Append("Note: 'Life Regenerated...' values are shown as whole numbers without proper decimal values on GGG's web site, which is likely a bug. Please note that the actual value for this mod may be higher or lower than indicated." & vbCrLf & vbCrLf)
        End If
        sb.Append(IIf(MyMod(intIndex).UnknownValues, "Note: the value(s) for this mod are beyond the possible ranges listed...this is likely a 'legacy' mod value for which the level cannot be indicated.", "Current level is indicated by (*)."))
        MsgBox(sb.ToString, , "Mod Info for " & strAffix & " '" & MyMod(intIndex).FullText & "'")
    End Sub

    Public Sub ShowAllPossibleMods(dg As DataGridView, myItem As FullItem, myMod As List(Of FullMod), strAffix As String)
        Dim strGearType As String = ""
        If myItem.GearType.ToString.CompareMultiple(StringComparison.Ordinal, "Sword (2h)", "Axe (2h)") Then
            strGearType = "[2h Swords and Axes]"
        ElseIf myItem.GearType.ToString.CompareMultiple(StringComparison.Ordinal, "Sword (1h)", "Axe (1h)") Then
            strGearType = "[1h Swords and Axes]"
        ElseIf myItem.GearType.ToString = "Mace (2h)" Then
            strGearType = "[2h Maces]"
        ElseIf myItem.GearType.ToString = "Mace (1h)" Then
            strGearType = "[1h Maces]"
        Else
            strGearType = myItem.GearType.ToString
        End If
        Dim sb As New System.Text.StringBuilder, lstMods As New List(Of String)
        For Each row As DataRow In dtWeights.Rows
            Dim strDesc As String = row("Description")
            Dim dtRow() As DataRow = dtMods.Select("[Description]='" & strDesc & "' AND " & strGearType & "=True " & _
                                                   "AND Level<=" & Math.Round(1.25 * (myItem.Level + 0.49)))
            If dtRow.Count = 0 Then Continue For
            If dtRow(0)("Prefix/Suffix") = strAffix Then
                If (strGearType = "Chest" Or strGearType = "Shield" Or strGearType = "Boots") And _
                    dtRow(0)("Categories").ToString.CompareMultiple(StringComparison.Ordinal, "Evasion", "Evasion and stun recovery", "Armor", "Armor and stun recovery", _
                                                                   "Energy shield", "Energy shield and stun recovery", "Dexterity increase", "Strenght increase", "Intelligence increase") Then
                    If (myItem.Eva > 0 And ((myItem.Arm = 0 And myItem.ES = 0 And dtRow(0)("Categories").Contains("Evasion")) Or dtRow(0)("Categories").ToString = "Dexterity increase")) Or _
                        (myItem.Arm > 0 And ((myItem.Eva = 0 And myItem.ES = 0 And dtRow(0)("Categories").Contains("Armor")) Or dtRow(0)("Categories").ToString = "Strenght increase")) Or _
                        (myItem.ES > 0 And ((myItem.Eva = 0 And myItem.Arm = 0 And dtRow(0)("Categories").Contains("Energy shield")) Or dtRow(0)("Categories").ToString = "Intelligence increase")) Then
                        If lstMods.Contains(row("ExportField").ToString & IIf(row("ExportField2").ToString = "", "", "/" & row("ExportField2").ToString)) = False Then
                            lstMods.Add(row("ExportField").ToString & IIf(row("ExportField2").ToString = "", "", "/" & row("ExportField2").ToString))
                        End If
                    End If
                ElseIf (strGearType = "Chest" Or strGearType = "Shield" Or strGearType = "Boots" Or strGearType = "Helmet" Or strGearType = "Gloves") And _
                    dtRow(0)("Categories").ToString.CompareMultiple(StringComparison.Ordinal, "Hybrid defense", "Hybrid defense and stun recovery") Then
                    If (myItem.Eva > 0 And myItem.Arm > 0 And dtRow(0)("Description").ToString.Contains("Armour And Evasion")) Or _
                        (myItem.ES > 0 And myItem.Arm > 0 And dtRow(0)("Description").ToString.Contains("Armour And Energy Shield")) Or _
                        (myItem.Eva > 0 And myItem.ES > 0 And dtRow(0)("Description").ToString.Contains("Evasion And Energy Shield")) Or _
                        (myItem.Arm > 0 And dtRow(0)("Description").ToString.Contains("Armour") And dtRow(0)("Description").ToString.Contains("Evasion") = False And dtRow(0)("Description").ToString.Contains("Energy Shield") = False) Or _
                        (myItem.Eva > 0 And dtRow(0)("Description").ToString.Contains("Evasion") And dtRow(0)("Description").ToString.Contains("Armour") = False And dtRow(0)("Description").ToString.Contains("Energy Shield") = False) Or _
                        (myItem.ES > 0 And dtRow(0)("Description").ToString.Contains("Energy Shield") And dtRow(0)("Description").ToString.Contains("Evasion") = False And dtRow(0)("Description").ToString.Contains("Armour") = False) Then
                        If lstMods.Contains(row("ExportField").ToString & IIf(row("ExportField2").ToString = "", "", "/" & row("ExportField2").ToString)) = False Then
                            lstMods.Add(row("ExportField").ToString & IIf(row("ExportField2").ToString = "", "", "/" & row("ExportField2").ToString))
                        End If
                    End If
                ElseIf strGearType = "Helmet" And dtRow(0)("Categories").ToString.CompareMultiple(StringComparison.Ordinal, "Evasion", "Evasion and stun recovery", "Armor", "Armor and stun recovery", _
                                                                   "Energy shield", "Energy shield and stun recovery", "Dexterity increase", "Strenght increase") Then
                    If (myItem.Eva > 0 And ((myItem.Arm = 0 And myItem.ES = 0 And dtRow(0)("Categories").Contains("Evasion")) Or dtRow(0)("Categories").ToString = "Dexterity increase")) Or _
                        (myItem.Arm > 0 And ((myItem.Eva = 0 And myItem.ES = 0 And dtRow(0)("Categories").Contains("Armor")) Or dtRow(0)("Categories").ToString = "Strenght increase")) Or _
                        (myItem.ES > 0 And myItem.Eva = 0 And myItem.Arm = 0 And dtRow(0)("Categories").Contains("Energy shield")) Then
                        If lstMods.Contains(row("ExportField").ToString & IIf(row("ExportField2").ToString = "", "", "/" & row("ExportField2").ToString)) = False Then
                            lstMods.Add(row("ExportField").ToString & IIf(row("ExportField2").ToString = "", "", "/" & row("ExportField2").ToString))
                        End If
                    End If
                ElseIf strGearType = "Gloves" And dtRow(0)("Categories").ToString.CompareMultiple(StringComparison.Ordinal, "Evasion", "Evasion and stun recovery", "Armor", "Armor and stun recovery", _
                                                                   "Energy shield", "Energy shield and stun recovery", "Strenght increase", "Intelligence increase") Then
                    If (myItem.Arm > 0 And ((myItem.Eva = 0 And myItem.ES = 0 And dtRow(0)("Categories").Contains("Armor")) Or dtRow(0)("Categories").ToString = "Strenght increase")) Or _
                        (myItem.ES > 0 And ((myItem.Eva = 0 And myItem.Arm = 0 And dtRow(0)("Categories").Contains("Energy shield")) Or dtRow(0)("Categories").ToString = "Intelligence increase")) Then
                        If lstMods.Contains(row("ExportField").ToString & IIf(row("ExportField2").ToString = "", "", "/" & row("ExportField2").ToString)) = False Then
                            lstMods.Add(row("ExportField").ToString & IIf(row("ExportField2").ToString = "", "", "/" & row("ExportField2").ToString))
                        End If
                    End If
                ElseIf strGearType = "Shield" And dtRow(0)("Description").ToString.CompareMultiple(StringComparison.Ordinal, "Local Socketed Cold Gem Level +", "Local Socketed Fire Gem Level +", _
                                                  "Local Socketed Lightning Gem Level +", "Spell Dmg +%", "Spell Critical Strike Chance +%", "Mana Regeneration Rate +%") Then
                    If myItem.ES > 0 And myItem.Arm = 0 And myItem.Eva = 0 Then
                        If lstMods.Contains(row("ExportField").ToString & IIf(row("ExportField2").ToString = "", "", "/" & row("ExportField2").ToString)) = False Then
                            lstMods.Add(row("ExportField").ToString & IIf(row("ExportField2").ToString = "", "", "/" & row("ExportField2").ToString))
                        End If
                    End If
                ElseIf strGearType = "Shield" And dtRow(0)("Description").ToString.CompareMultiple(StringComparison.Ordinal, "Local Socketed Melee Gem Level +") Then
                    If myItem.Eva > 0 Or myItem.Arm > 0 Then
                        If lstMods.Contains(row("ExportField").ToString & IIf(row("ExportField2").ToString = "", "", "/" & row("ExportField2").ToString)) = False Then
                            lstMods.Add(row("ExportField").ToString & IIf(row("ExportField2").ToString = "", "", "/" & row("ExportField2").ToString))
                        End If
                    End If
                ElseIf (strGearType = "Chest" Or strGearType = "Shield" Or strGearType = "Boots" Or strGearType = "Helmet" Or strGearType = "Gloves") And _
                    dtRow(0)("Categories").ToString.CompareMultiple(StringComparison.Ordinal, "Mana") Then
                    If myItem.ES > 0 Then
                        If lstMods.Contains(row("ExportField").ToString & IIf(row("ExportField2").ToString = "", "", "/" & row("ExportField2").ToString)) = False Then
                            lstMods.Add(row("ExportField").ToString & IIf(row("ExportField2").ToString = "", "", "/" & row("ExportField2").ToString))
                        End If
                    End If
                Else
                    If lstMods.Contains(row("ExportField").ToString & IIf(row("ExportField2").ToString = "", "", "/" & row("ExportField2").ToString)) = False Then
                        lstMods.Add(row("ExportField").ToString & IIf(row("ExportField2").ToString = "", "", "/" & row("ExportField2").ToString))
                    End If
                End If
            ElseIf strDesc = "Base Item Found Rarity +%" Then   ' Rarity can be both a suffix and a prefix
                For Each r As DataRow In dtRow
                    If r("Prefix/Suffix") <> strAffix Then
                        If lstMods.Contains("% increased Rarity of Items found") = False Then lstMods.Add("% increased Rarity of Items found")
                        Exit For
                    End If
                Next
            End If
        Next
        lstMods.Sort()
        For Each strMod In lstMods
            For Each m As FullMod In myMod
                If (m.Type1 & IIf(m.Type2 = "", "", "/" & m.Type2)).ToString.ToLower = strMod.ToLower Then
                    sb.Append("(*) ")
                    Exit For
                End If
            Next
            sb.Append(strMod & Environment.NewLine)
        Next
        MessageBox.Show(sb.ToString, "List of Possible " & strAffix & " Mods for Type: " & myItem.GearType, MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Public Sub AddGemWarning(ByRef sb As System.Text.StringBuilder)
        sb.Append("**WARNING!!** Item level requirements set by an equipped gem...mod level ranges are likely too large, and the ranking will be artificially lowered as a result. It is recommended that the gem be removed before evaluating the item in this manner." & vbCrLf & vbCrLf)
    End Sub

    Private Sub DataGridView1_CellMouseEnter(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellMouseEnter
        If blScroll = True Then Exit Sub
        If IsValidCellAddress(DataGridView1, e.RowIndex, e.ColumnIndex) AndAlso _
            (DataGridView1.Columns(e.ColumnIndex).Name.Contains("fix") Or _
             DataGridView1.Columns(e.ColumnIndex).Name.CompareMultiple(StringComparison.Ordinal, "Sokt", "Link", "Rank", "Location", "*")) _
         Then DataGridView1.Cursor = Cursors.Hand
    End Sub

    Public Function IsValidCellAddress(dg As DataGridView, rowIndex As Integer, columnIndex As Integer) As Boolean
        Return rowIndex >= 0 AndAlso rowIndex < dg.RowCount AndAlso columnIndex >= 0 AndAlso columnIndex <= dg.ColumnCount
    End Function

    Private Sub DataGridView1_CellMouseLeave(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellMouseLeave
        If blScroll = True Then Exit Sub
        If IsValidCellAddress(DataGridView1, e.RowIndex, e.ColumnIndex) AndAlso _
            (DataGridView1.Columns(e.ColumnIndex).Name.Contains("fix") Or _
             DataGridView1.Columns(e.ColumnIndex).Name.CompareMultiple(StringComparison.Ordinal, "Sokt", "Link", "Rank", "Location", "*")) _
         Then DataGridView1.Cursor = Cursors.Default
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
        Dim intIndex As Integer = CInt(DataGridView1.Rows(e.RowIndex).Cells("Index").Value)
        If strName.Contains("prefix") AndAlso NotNull(DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value, "").ToString <> "" Then
            DataGridViewAddLevelBar(DataGridView1, FullInventory(intIndex).ExplicitPrefixMods, strName, sender, e)
        ElseIf strName.Contains("suffix") AndAlso NotNull(DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex).Value, "").ToString <> "" Then
            DataGridViewAddLevelBar(DataGridView1, FullInventory(intIndex).ExplicitSuffixMods, strName, sender, e)
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
            Dim intPos As Integer = CInt(GetNumeric(strName) - 1)
            Dim p As Double = 0, q As Double = 0
            'Dim sngMinDen As Single = 0     ' This is the denominator unit used by the values, for use in setting our level "notches"
            With MyModList(intPos)
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
                    q = CSng(IIf(.BaseUpperV1 = .BaseLowerV1, 1, (.Value1 - .BaseLowerV1 + 1))) / CSng(IIf(.BaseUpperV1 = .BaseLowerV1, 1, (.BaseUpperV1 - .BaseLowerV1 + 1))) * sngModWeight1 + _
                        CSng(IIf(.BaseUpperMaxV1 = .BaseLowerMaxV1, 1, (.MaxValue1 - .BaseLowerMaxV1 + 1))) / CSng(IIf(.BaseUpperMaxV1 = .BaseLowerMaxV1, 1, (.BaseUpperMaxV1 - .BaseLowerMaxV1 + 1))) * sngModWeight2
                    'sngMinDen = IIf(.BaseUpperV1 = .BaseLowerV1, 0, 1 / IIf(.BaseUpperV1 = .BaseLowerV1, 1, .BaseUpperV1 - .BaseLowerV1 + 1)) * sngModWeight1 + _
                    '    IIf(.BaseUpperMaxV1 = .BaseLowerMaxV1, 0, 1 / IIf(.BaseUpperMaxV1 = .BaseLowerMaxV1, 1, (.BaseUpperMaxV1 - .BaseLowerMaxV1 + 1))) * sngModWeight2
                Else
                    If .BaseUpperV1 = .BaseLowerV1 And .Value2 = 0 And .ModLevelMax = 0 Then HighlightNoVariationMods(dg, e) : Exit Sub ' There is no variation possible, so exit sub and don't draw bar
                    q = CSng(IIf(.BaseUpperV1 = .BaseLowerV1, 1, (.Value1 - .BaseLowerV1 + 1))) / CSng(IIf(.BaseUpperV1 = .BaseLowerV1, 1, (.BaseUpperV1 - .BaseLowerV1 + 1)))
                    'sngMinDen = IIf(.BaseUpperV1 = .BaseLowerV1, 0, 1 / IIf(.BaseUpperV1 = .BaseLowerV1, 1, .BaseUpperV1 - .BaseLowerV1 + 1))
                End If
                If .Value2 <> 0 Then
                    If .BaseUpperV1 = .BaseLowerV1 And .BaseUpperV2 = .BaseLowerV2 And .ModLevelMax = 0 Then HighlightNoVariationMods(dg, e) : Exit Sub ' There is no variation possible, so exit sub and don't draw bar
                    If .BaseUpperV1 = .BaseLowerV1 Then sngModWeight1 = 0 : sngModWeight2 = 1
                    If .BaseUpperV2 = .BaseLowerV2 Then sngModWeight1 = 1 : sngModWeight2 = 0
                    q = CSng(IIf(.BaseUpperV1 = .BaseLowerV1, 1, (.Value1 - .BaseLowerV1 + 1))) / CSng(IIf(.BaseUpperV1 = .BaseLowerV1, 1, (.BaseUpperV1 - .BaseLowerV1 + 1))) * sngModWeight1 + _
                        CSng(IIf(.BaseUpperV2 = .BaseLowerV2, 1, (.Value2 - .BaseLowerV2 + 1))) / CSng(IIf(.BaseUpperV2 = .BaseLowerV2, 1, (.BaseUpperV2 - .BaseLowerV2 + 1))) * sngModWeight2
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
            r.Width = CInt((e.CellBounds.Width - 6) * p)
            r.Height = e.CellBounds.Height - 6
            Dim br2 As New Drawing.Drawing2D.LinearGradientBrush(r, Drawing.Color.White, Drawing.Color.DarkGray, Drawing.Drawing2D.LinearGradientMode.Vertical)
            e.Graphics.FillRectangle(br2, r)
            e.PaintContent(e.ClipBounds)
            '' Now draw the vertical notches that indicate to the user where the level 'breaks' are
            'For i = 0 To MyModList(intPos).ModLevelMax
            '    Dim sngStart As Single = CSng(e.CellBounds.X + 3 + e.CellBounds.Width * i / (MyModList(intPos).ModLevelMax + 1)) + sngMinDen * (e.CellBounds.Width - 6)
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
            If CBool(dg.Rows(e.RowIndex).Cells("Gem").Value) = True Then
                dg.Rows(e.RowIndex).Cells("Gem").Style.BackColor = ColorHasGem
                dg.Rows(e.RowIndex).Cells("Rank").Style.BackColor = ColorHasGem
                dg.Rows(e.RowIndex).Cells("Level").Style.BackColor = ColorHasGem
                dg.Rows(e.RowIndex).Cells("%").Style.BackColor = ColorHasGem
            End If
            Dim intIndex As Integer = CInt(dg.Rows(e.RowIndex).Cells("Index").Value)
            If intIndex <> -1 Then
                HighlightUnknownMods(dg, MyInventory(intIndex).ExplicitPrefixMods, "Prefix ", e.RowIndex)
                HighlightUnknownMods(dg, MyInventory(intIndex).ExplicitSuffixMods, "Suffix ", e.RowIndex)
                For i = 1 To 3
                    If i > MyInventory(intIndex).ExplicitPrefixMods.Count Then Exit For
                    If intWeightMax <= MyInventory(intIndex).ExplicitPrefixMods(i - 1).Weight Then
                        dg.Rows(e.RowIndex).Cells("Prefix " & i).Style.ForeColor = CType(ColorMax, Color)
                        dg.Rows(e.RowIndex).Cells("Prefix " & i).Style.Font = New Font("Segoe UI", 8.25, FontStyle.Regular)
                    ElseIf intWeightMin >= MyInventory(intIndex).ExplicitPrefixMods(i - 1).Weight Then
                        dg.Rows(e.RowIndex).Cells("Prefix " & i).Style.ForeColor = CType(ColorMin, Color)
                        dg.Rows(e.RowIndex).Cells("Prefix " & i).Style.Font = New Font("Segoe UI", 8.25, FontStyle.Italic)
                    End If
                Next
                For i = 1 To 3
                    If i > MyInventory(intIndex).ExplicitSuffixMods.Count Then Exit For
                    If intWeightMax <= MyInventory(intIndex).ExplicitSuffixMods(i - 1).Weight Then
                        dg.Rows(e.RowIndex).Cells("Suffix " & i).Style.ForeColor = CType(ColorMax, Color)
                        dg.Rows(e.RowIndex).Cells("Suffix " & i).Style.Font = New Font("Segoe UI", 8.25, FontStyle.Regular)
                    ElseIf intWeightMin >= MyInventory(intIndex).ExplicitSuffixMods(i - 1).Weight Then
                        dg.Rows(e.RowIndex).Cells("Suffix " & i).Style.ForeColor = CType(ColorMin, Color)
                        dg.Rows(e.RowIndex).Cells("Suffix " & i).Style.Font = New Font("Segoe UI", 8.25, FontStyle.Italic)
                    End If
                Next
            End If
            If dg.Rows(e.RowIndex).Cells("*").Value.ToString = "*" Then dg.Rows(e.RowIndex).Cells("*").Style.BackColor = ColorOtherSolutions
        Catch ex As Exception
            ErrorHandler(System.Reflection.MethodBase.GetCurrentMethod.Name, ex, "Item Name: " & dg.Rows(e.RowIndex).Cells("Name").Value)
        End Try
    End Sub

    Private Sub HighlightUnknownMods(dg As DataGridView, MyList As List(Of FullMod), strAffix As String, RowIndex As Integer)
        For Each MyMod In MyList
            If MyMod.UnknownValues = True Then
                dg.Rows(RowIndex).Cells(strAffix & MyList.IndexOf(MyMod) + 1).Style.BackColor = ColorUnknownValue
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

            If DataGridView1.Visible = True Or DataGridView2.Visible = True Then
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
        pb.PerformStep() : pb.Refresh()
    End Sub

    Public Sub PBClose()
        grpProgress.Visible = False
        pb.Visible = False
        lblpb.Visible = False
    End Sub

    Public Sub SetPBDefaults(intMax As Integer, strLabel As String)
        grpProgress.Left = CInt((Me.Width - grpProgress.Width) / 2)
        grpProgress.Top = CInt((Me.Height - grpProgress.Height) / 2)
        lblpb.Text = strLabel
        If strLabel.ToLower.Contains("poexplorer") Then
            pb.Maximum = 100
            pb.MarqueeAnimationSpeed = 100
            pb.Style = ProgressBarStyle.Marquee
        Else
            pb.Minimum = 0
            pb.Maximum = CInt((intMax) / 10)
            pb.Value = 0
            pb.Step = 1
            pb.Style = ProgressBarStyle.Blocks
        End If
        grpProgress.Visible = True
        pb.Visible = True
        lblpb.Visible = True
    End Sub

    Public Sub RecalculateAllRankings(Optional blPartialRefresh As Boolean = False, Optional blShowPB As Boolean = True, Optional blSkipStore As Boolean = False, Optional blSkipStash As Boolean = False)
        Try
            ' Full speed ahead with the full recalculating!
            Dim intTotal As Integer = IIf(blSkipStash = False, FullInventory.Count + TempInventory.Count, 0) + IIf(blSkipStore = False, FullStoreInventory.Count + TempStoreInventory.Count, 0)
            If blShowPB = True Then Me.Invoke(New pbSetDefaults(AddressOf SetPBDefaults), New Object() {intTotal, "Please wait, calculating rankings..."})
            If blPartialRefresh = False Then EnableDisableControls(False, New List(Of String)(New String() {"pb", "lblpb", "grpProgress"}))
            'Me.Invoke(New UCPD(AddressOf SetControlProperty), New Object() {cmbWeight, "Enabled", False})
            'Me.Invoke(New UCPD(AddressOf SetControlProperty), New Object() {lblWeights, "Enabled", False})
            'Me.Invoke(New UCPD(AddressOf SetControlProperty), New Object() {btnEditWeights, "Enabled", False})
            Dim blFilter As Boolean = strFilter <> ""
            If blPartialRefresh = False Then
                Me.Invoke(New UCPD(AddressOf SetControlProperty), New Object() {DataGridView1, "Visible", False})
                Me.Invoke(New UCPD(AddressOf SetControlProperty), New Object() {DataGridView1, "DataSource", ""})
                Me.Invoke(New UCPD(AddressOf SetControlProperty), New Object() {DataGridView2, "Visible", False})
                Me.Invoke(New UCPD(AddressOf SetControlProperty), New Object() {DataGridView2, "DataSource", ""})
                Me.Invoke(New UCPD(AddressOf SetControlProperty), New Object() {Me, "UseWaitCursor", True})
                Application.DoEvents()
            End If

            Dim lngCounter As Long = 0
            For Each mylist In {FullInventory, TempInventory, FullStoreInventory, TempStoreInventory}
                Dim strList As String = ""
                If mylist.Equals(FullInventory) Or mylist.Equals(FullStoreInventory) Then strList = "Full" Else strList = "Temp"
                If blSkipStash = True AndAlso (mylist.Equals(FullInventory) Or mylist.Equals(TempInventory)) Then Continue For
                If blSkipStore = True AndAlso (mylist.Equals(FullStoreInventory) Or mylist.Equals(TempStoreInventory)) Then Continue For
                For Each it In mylist
                    it.Rank = 0
                    Dim strRankKey As String = IIf(strList = "Full", it.ID & it.Name, it.ID & it.Name & mylist.IndexOf(it)).ToString
                    RankExplanation(strRankKey) = ""
                    For Each ModList In {it.ExplicitPrefixMods, it.ExplicitSuffixMods}
                        Dim strAffix As String = ""
                        If ModList.Count <> it.ExplicitPrefixMods.Count Or ModList.Count = 0 Then
                            strAffix = "Suffix"
                        Else
                            If ModList(0).FullText <> it.ExplicitPrefixMods(0).FullText Then strAffix = "Suffix" Else strAffix = "Prefix"
                        End If
                        For Each myMod In ModList
                            Dim result() As DataRow = Nothing
                            result = dtWeights.Select("ExportField = '" & myMod.Type1 & "'" & IIf(myMod.Type2 <> "", " AND ExportField2 = '" & myMod.Type2 & "'", " AND ExportField2 = ''").ToString)
                            If result.Count <> 0 Then
                                For Each row In result
                                    If RunModResultQuery(it, row).Count <> 0 Then
                                        myMod.Weight = CSng(row("Weight"))
                                        Exit For
                                    End If
                                Next
                                If myMod.UnknownValues = True Then
                                    it.Rank += RankUnknownValueMod(myMod, myMod.MaxValue1, myMod.ModLevelMax, it.ID.ToString, it.Name, strAffix, strRankKey)
                                Else
                                    it.Rank += CalculateRank(myMod, myMod.ModLevelMax, result(0)("Description").ToString & myMod.ModLevelActual, it.ID.ToString, it.Name, strAffix, strRankKey)
                                End If
                            End If
                            RankExplanation(strRankKey) += vbCrLf & vbCrLf & "(" & String.Format("{0:'+'0;'-'0}", it.Rank) & ")  (running total)" & vbCrLf
                        Next
                    Next
                    If RankExplanation(strRankKey).LastIndexOf("(running total)") <> -1 Then _
                        RankExplanation(strRankKey) = RankExplanation(strRankKey).Substring(0, RankExplanation(strRankKey).LastIndexOf("(running total)")) & "Final Rank" & vbCrLf
                    lngCounter += 1
                    If lngCounter Mod 10 = 0 And blShowPB Then
                        Me.Invoke(New MyDelegate(AddressOf PBPerformStep))
                    ElseIf lngCounter Mod 100 = 0 And blShowPB = False Then
                        statusController.DisplayMessage("Ranked " & lngCounter & " item entries...")
                    End If
                Next
            Next
            Application.DoEvents()
            If blPartialRefresh Then Exit Sub
            dtRank.Clear() : dtOverflow.Clear() : dtStore.Clear() : dtStoreOverflow.Clear()

            If FullInventory.Count <> 0 Then
                AddToDataTable(FullInventory, dtRank, False, "")
                AddToDataTable(TempInventory, dtOverflow, False, "")

                Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {DataGridView1, "DataSource", dtRank})
                Me.Invoke(New MyDualControlDelegate(AddressOf HideColumns), New Object() {Me, DataGridView1})
                Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {DataGridView1.Columns("%").DefaultCellStyle, "Format", "n1"})
                Me.Invoke(New MyDGDelegate(AddressOf SortDataGridView), DataGridView1)
                Me.Invoke(New MyControlDelegate(AddressOf SetDataGridViewWidths), New Object() {DataGridView1})
                If blFilter Then ApplyFilter(dtRank, dtOverflow, DataGridView1, lblRecordCount, strOrderBy, strFilter)
                Dim FirstCell As DataGridViewCell
                FirstCell = Me.Invoke(New DataGridCell(AddressOf ReturnFirstCell), New Object() {DataGridView1})
                Me.Invoke(New UCPD(AddressOf SetControlProperty), New Object() {DataGridView1, "FirstDisplayedCell", FirstCell})
                Me.Invoke(New UCPD(AddressOf SetControlProperty), New Object() {DataGridView1, "Visible", True})
            End If

            If FullStoreInventory.Count <> 0 Then
                AddToDataTable(FullStoreInventory, dtStore, False, "")
                AddToDataTable(TempStoreInventory, dtStoreOverflow, False, "")

                Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {DataGridView2, "DataSource", dtStore})
                Me.Invoke(New MyDualControlDelegate(AddressOf HideColumns), New Object() {Me, DataGridView2})
                Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {DataGridView2.Columns("%").DefaultCellStyle, "Format", "n1"})
                Me.Invoke(New MyDGDelegate(AddressOf SortDataGridView), DataGridView2)
                Me.Invoke(New MyControlDelegate(AddressOf SetDataGridViewWidths), New Object() {DataGridView2})
                ' To make room for new Price column take away from SubType and Implicit columns
                Dim intWidth As Integer = Math.Max(CType(Me.Invoke(New RCPD(AddressOf ReadControlProperty), New Object() {DataGridView2.Columns("SubType"), "Width"}), Integer) - 15, 0)
                Me.Invoke(New UCPD(AddressOf SetControlProperty), New Object() {DataGridView2.Columns("SubType"), "Width", intWidth})
                intWidth = Math.Max(CType(Me.Invoke(New RCPD(AddressOf ReadControlProperty), New Object() {DataGridView2.Columns("Implicit"), "Width"}), Integer) - 35, 0)
                Me.Invoke(New UCPD(AddressOf SetControlProperty), New Object() {DataGridView2.Columns("Implicit"), "Width", intWidth})
                If blFilter Then ApplyFilter(dtStore, dtStoreOverflow, DataGridView2, lblRecordCount2, strStoreOrderBy, strStoreFilter)
                Dim FirstCell As DataGridViewCell
                FirstCell = Me.Invoke(New DataGridCell(AddressOf ReturnFirstCell), New Object() {DataGridView2})
                Me.Invoke(New UCPD(AddressOf SetControlProperty), New Object() {DataGridView2, "FirstDisplayedCell", FirstCell})
                Me.Invoke(New UCPD(AddressOf SetControlProperty), New Object() {DataGridView2, "Visible", True})
            End If

            'frmPB.Close() : frmPB.Dispose()
            Me.Invoke(New MyDelegate(AddressOf PBClose))
            If FullInventory.Count <> 0 Then
                EnableDisableControls(True, New List(Of String)(New String() {"ElementHost2", "txtEmail", "lblEmail", "ElementHost1", "lblPassword", "btnLoad", "btnOffline", "chkSession"}))
            Else
                EnableDisableControls(True)
            End If

            'Me.Invoke(New UCPD(AddressOf SetControlProperty), New Object() {cmbWeight, "Enabled", True})
            'Me.Invoke(New UCPD(AddressOf SetControlProperty), New Object() {lblWeights, "Enabled", True})
            'Me.Invoke(New UCPD(AddressOf SetControlProperty), New Object() {btnEditWeights, "Enabled", True})
            Me.Invoke(New UCPD(AddressOf SetControlProperty), New Object() {Me, "UseWaitCursor", False})
            Application.DoEvents()
            ' Sometimes the wait cursor property doesn't get unset for the datagridview, perhaps because it is not visible for some time?
            Me.Invoke(New UCPD(AddressOf SetControlProperty), New Object() {DataGridView1, "UseWaitCursor", False})
            Me.Invoke(New UCPD(AddressOf SetControlProperty), New Object() {DataGridView1, "Cursor", Cursors.Default})
            Me.Invoke(New UCPD(AddressOf SetControlProperty), New Object() {DataGridView2, "UseWaitCursor", False})
            Me.Invoke(New UCPD(AddressOf SetControlProperty), New Object() {DataGridView2, "Cursor", Cursors.Default})
            Dim strTabName As String = CType(Me.Invoke(New MyDelegateFunction(AddressOf CurrentTabName)), String)
            If strTabName = "tabpage2" Then
                Me.BeginInvoke(New MyDGDelegate(AddressOf SetDataGridFocus), DataGridView2)
            Else
                Me.BeginInvoke(New MyDGDelegate(AddressOf SetDataGridFocus), DataGridView1)
            End If
            'Me.BeginInvoke(New MyDelegate(AddressOf SetDataGridFocus))

        Catch ex As Exception
            Me.Invoke(New MyDelegate(AddressOf PBClose))
            If FullInventory.Count <> 0 Then
                EnableDisableControls(True, New List(Of String)(New String() {"ElementHost2", "txtEmail", "lblEmail", "ElementHost1", "lblPassword", "btnLoad", "btnOffline", "chkSession"}))
            Else
                EnableDisableControls(True)
            End If
            'Me.Invoke(New UCPD(AddressOf SetControlProperty), New Object() {cmbWeight, "Enabled", True})
            'Me.Invoke(New UCPD(AddressOf SetControlProperty), New Object() {lblWeights, "Enabled", True})
            'Me.Invoke(New UCPD(AddressOf SetControlProperty), New Object() {btnEditWeights, "Enabled", True})
            Me.Invoke(New UCPD(AddressOf SetControlProperty), New Object() {Me, "UseWaitCursor", False})

            ErrorHandler(System.Reflection.MethodBase.GetCurrentMethod.Name, ex)
        End Try
    End Sub

    Private Function ReturnFirstCell(dg As DataGridView) As DataGridViewCell
        Return dg.Rows(0).Cells(0)
    End Function

    Private Function CurrentTabName() As String
        Return TabControl1.SelectedTab.Name.ToLower
    End Function

    Private Sub DataGridView1_Scroll(sender As Object, e As ScrollEventArgs) Handles DataGridView1.Scroll
        blScroll = True
    End Sub

    Private Sub DataGridView1_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView1.SelectionChanged
        If DataGridView1.CurrentCell Is Nothing Then Exit Sub
        If DataGridView1.Cursor <> Cursors.Default Then Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {DataGridView1, "Cursor", Cursors.Default})
        If DataGridView1.Columns(DataGridView1.CurrentCell.ColumnIndex).Name.ToLower.Contains("prefix") Or _
            DataGridView1.Columns(DataGridView1.CurrentCell.ColumnIndex).Name.ToLower.Contains("suffix") Then Me.BeginInvoke(New MyDGDelegate(AddressOf DataGridClearSelection), DataGridView1)
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

    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        frmFilter.blStore = False
        frmFilter.ShowDialog(Me)
    End Sub

    Public Sub SetFilter(strTemp As String)
        If strTemp = "" Then
            strFilter = ""
            strOrderBy = "[RANK] DESC"
        Else
            Dim strTestFilter As String = ConvertFilterToSQL(strTemp)
            If strTestFilter = strFilter Then Exit Sub
            strFilter = strTestFilter.Substring(0, strTestFilter.IndexOf("ORDER BY"))
            strOrderBy = strTestFilter.Substring(strTestFilter.IndexOf("ORDER BY") + 9)
        End If
        If dtRank.Rows.Count = 0 Then Exit Sub ' Don't apply the filter if we haven't loaded data yet
        Dim FilterThread As New Threading.Thread(Sub() Me.ApplyFilter(dtRank, dtOverflow, DataGridView1, lblRecordCount, strOrderBy, strFilter))
        FilterThread.SetApartmentState(Threading.ApartmentState.STA)
        FilterThread.Start()
    End Sub

    Public Sub SetStoreFilter(strTemp As String)
        If strTemp = "" Then
            strStoreFilter = ""
            strStoreOrderBy = "[RANK] DESC"
        Else
            Dim strTestFilter As String = ConvertFilterToSQL(strTemp)
            If strTestFilter = strStoreFilter Then Exit Sub
            strStoreFilter = strTestFilter.Substring(0, strTestFilter.IndexOf("ORDER BY"))
            strStoreOrderBy = strTestFilter.Substring(strTestFilter.IndexOf("ORDER BY") + 9)
        End If
        If dtStore.Rows.Count = 0 Then Exit Sub ' Don't apply the filter if we haven't loaded data yet
        Dim FilterThread As New Threading.Thread(Sub() Me.ApplyFilter(dtStore, dtStoreOverflow, DataGridView2, lblRecordCount2, strStoreOrderBy, strStoreFilter))
        FilterThread.SetApartmentState(Threading.ApartmentState.STA)
        FilterThread.Start()
    End Sub

    Public Sub ApplyFilter(dt As DataTable, dtOver As DataTable, dg As DataGridView, lblRec As System.Windows.Forms.Label, strOrder As String, strF As String)
        Try
            Dim FirstCell As DataGridViewCell
            Dim intRows As Integer = 0
            If strF.Trim.Length = 0 Then
                Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {dg, "DataSource", dt})
                intRows = CInt(Me.Invoke(New RCPD(AddressOf ReadControlProperty), New Object() {dg.Rows, "Count"}).ToString)
                Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {lblRec, "Text", "Number of Rows: " & intRows})
                FirstCell = Me.Invoke(New DataGridCell(AddressOf ReturnFirstCell), New Object() {dg})
                Me.Invoke(New UCPD(AddressOf SetControlProperty), New Object() {dg, "FirstDisplayedCell", FirstCell})
                Exit Sub
            End If
            dtRankFilter = dt.Select(strF).CopyToDataTable
            If strF.Contains("[*]") = False Then   ' Don't add from dtrOverflow if the filter is already dealing with membership in that table (or not) via the "*" field
                ' Apply the filter to the dtOverflow and add in any items that satisfy the conditions but aren't already selected
                Dim drRows As DataRow() = dtOver.Select(strF)
                For Each row As DataRow In drRows
                    Dim strId As String = row("ID").ToString
                    Dim strName As String = row("Name").ToString
                    Dim dr As DataRow() = dtRankFilter.Select("[ID]='" & strId & "' AND [Name]='" & strName & "'")
                    If dr.Count = 0 Then
                        ' Get the associated datarow from dtRank and force it into dtRankFilter
                        Dim dr2 As DataRow() = dt.Select("[ID]='" & strId & "' AND [Name]='" & strName & "'")
                        If dr2.Count <> 0 Then
                            dtRankFilter.ImportRow(dr2(0))
                        End If
                    End If
                Next
            End If
            If strRawFilter.Contains("[Mod Total Value]") Then     ' This has a [Mod Total Value] component
                For i = 0 To Math.Min(lstTotalTypes.Count - 1, 5)
                    BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {dg.Columns("Tot" & i + 1), "Visible", True})
                    BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {dg.Columns("Tot" & i + 1), "Width", 33})
                    For Each row As DataRow In dtRankFilter.Rows
                        Dim intTotal As Integer = 0
                        Dim strType As String = lstTotalTypes(i)
                        For Each strField In {"pt1", "pt12", "pt2", "pt22", "pt3", "pt32", "st1", "st12", "st2", "st22", "st3", "st32", "it1", "it2", "it3"}
                            intTotal += CInt(IIf(row(strField).ToString = strType, row(strField.Replace("t", "v")).ToString, 0))
                        Next
                        row("Tot" & i + 1) = intTotal
                    Next
                Next
            Else
                For i = 1 To 6
                    BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {dg.Columns("Tot" & i), "Visible", False})
                Next
            End If
            If strOrder.Contains("Tot") = True Then     ' Reapply the ordering now that the total fields are populated
                dtRankFilter = dtRankFilter.Select(strF).CopyToDataTable
            End If
            Dim dv As DataView = New DataView(dtRankFilter)
            dv.Sort = strOrder
            Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {dg, "DataSource", dv})
            intRows = CInt(Me.Invoke(New RCPD(AddressOf ReadControlProperty), New Object() {dg.Rows, "Count"}).ToString)
            Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {lblRec, "Text", "Number of Rows: " & intRows & " (filter active)"})
            'FirstCell = CType(Me.Invoke(New MyDelegateFunction(AddressOf ReturnFirstCell)), DataGridViewCell)
            'Me.Invoke(New UCPD(AddressOf SetControlProperty), New Object() {dg, "FirstDisplayedCell", FirstCell})
            Me.BeginInvoke(New MyDGDelegate(AddressOf SetDataGridFocus), dg)
            'Me.Invoke(New MyDelegate(AddressOf SetDataGridFocus))

        Catch ex As Exception
            If ex.Message = "The source contains no DataRows." Then
                MessageBox.Show("The search returned 0 records." & Environment.NewLine & Environment.NewLine & _
                                "Query: " & Environment.NewLine & strF, "Search Returned No Results", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Dim strTabName As String = CType(Me.Invoke(New MyDelegateFunction(AddressOf CurrentTabName)), String)
                If strTabName.ToLower = "tabpage2" Then
                    Me.Invoke(New MyClickDelegate(AddressOf btnShopSearch_Click), New Object() {btnShopSearch, EventArgs.Empty})
                Else
                    Me.Invoke(New MyClickDelegate(AddressOf btnSearch_Click), New Object() {btnSearch, EventArgs.Empty})
                End If

            Else
                ErrorHandler(System.Reflection.MethodBase.GetCurrentMethod.Name, ex, "Filter: " & strF & Environment.NewLine & "Order: " & strOrder)
            End If
        End Try
    End Sub

    Private Sub btnDownloadJSON_Click(sender As Object, e As EventArgs) Handles btnDownloadJSON.Click
        If txtThread.Text.Trim = "" Then
            MessageBox.Show("You must enter a URL/thread.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        Dim strThreadURL As String = "", strThread As String = ""
        If txtThread.Text.Trim.StartsWith("http") Then
            strThreadURL = txtThread.Text.Trim
            strThread = txtThread.Text.Trim.Substring(txtThread.Text.Trim.IndexOf("threads/") + 8, _
                                                      txtThread.Text.Trim.IndexOf(".json") - (txtThread.Text.Trim.IndexOf("threads/") + 8))
        Else
            strThreadURL = "http://poexplorer.com/threads/" & txtThread.Text.Trim & ".json"
            strThread = txtThread.Text.Trim
        End If
        DownloadThread = New Threading.Thread(Sub() DownloadJSON(strThreadURL, strThread, True))
        DownloadThread.SetApartmentState(Threading.ApartmentState.STA)
        DownloadThread.Start()
    End Sub

    Public Sub DownloadJSON(strThreadURL As String, strThread As String, blLoad As Boolean)
        Dim request As HttpWebRequest
        Dim response As HttpWebResponse = Nothing
        Dim sr As StreamReader
        Dim blSuccess As Boolean = False
        Try
            'Me.Invoke(New pbSetDefaults(AddressOf SetPBDefaults), New Object() {1, "Downloading store JSON..."})
            EnableDisableControls(False, New List(Of String)(New String() {"pb", "lblpb", "grpProgress"}))
            Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {DataGridView2, "Visible", False})
            Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {lblRecordCount2, "Visible", False})
            Dim storeMerge As New List(Of JSON_Store)
            RemoveAllStoreJSONs(strThread)
            Me.Invoke(New pbSetDefaults(AddressOf SetPBDefaults), New Object() {1, "Contacting poexplorer.com..."})
            For i = 1 To 10000
                Dim strThreadURLPage As String = strThreadURL & "?page=" & i
                request = DirectCast(WebRequest.Create(strThreadURLPage), HttpWebRequest)
                response = DirectCast(request.GetResponse(), HttpWebResponse)
                sr = New StreamReader(response.GetResponseStream())
                Dim strLines As String = sr.ReadToEnd()
                sr.Close()
                If strLines = "[]" Then Exit For
                Dim jss As New System.Web.Script.Serialization.JavaScriptSerializer()
                jss.MaxJsonLength = Int32.MaxValue
                Dim store As List(Of JSON_Store) = jss.Deserialize(Of List(Of JSON_Store))(strLines)
                If store.Count = 0 Then Exit For
                storeMerge.AddRange(store)
                Me.Invoke(New UCPD(AddressOf SetControlProperty), New Object() {lblpb, "Text", "Downloading " & strThread & " store JSON (page " & i & ")..."})
                Dim swOut As New StreamWriter(Application.StartupPath & "\Store\" & strThread & "-" & i & ".json")
                swOut.WriteLine(strLines)
                swOut.Close()
                blSuccess = True
            Next
            If blLoad = False Then Exit Sub
            If blSuccess = False Then
                If FullInventory.Count <> 0 Then
                    EnableDisableControls(True, New List(Of String)(New String() {"ElementHost2", "txtEmail", "lblEmail", "ElementHost1", "lblPassword", "btnLoad", "btnOffline", "chkSession"}))
                Else
                    EnableDisableControls(True)
                End If
                Me.Invoke(New MyDelegate(AddressOf PBClose))
                If FullStoreInventory.Count <> 0 Then
                    Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {DataGridView2, "Visible", True})
                    Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {lblRecordCount2, "Visible", True})
                End If
                MessageBox.Show("Could not find store or the store is empty." & Environment.NewLine & Environment.NewLine & _
                                "URL: " & strThreadURL, "No Items Returned", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Exit Sub
            End If
            Dim storeQuery As IEnumerable(Of JSON_Store) = storeMerge.Where(Function(Item) Item.Rarity_Name = "Rare" AndAlso Item.Name.ToLower <> "" _
                                                AndAlso Item.Item_Type.ToLower.Contains("map") = False AndAlso Item.Item_Type.ToLower.Contains("peninsula") = False And Item.Identified = True)
            If storeQuery.Count = 0 Then
                If FullInventory.Count <> 0 Then
                    EnableDisableControls(True, New List(Of String)(New String() {"ElementHost2", "txtEmail", "lblEmail", "ElementHost1", "lblPassword", "btnLoad", "btnOffline", "chkSession"}))
                Else
                    EnableDisableControls(True)
                End If
                Me.Invoke(New MyDelegate(AddressOf PBClose))
                If FullStoreInventory.Count <> 0 Then
                    Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {DataGridView2, "Visible", True})
                    Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {lblRecordCount2, "Visible", True})
                End If
                MessageBox.Show("Could not find any rare items in the store." & Environment.NewLine & Environment.NewLine & _
                                "URL: " & strThreadURL, "No Rare Items Found", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Else
                Me.Invoke(New pbSetDefaults(AddressOf SetPBDefaults), New Object() {storeQuery.Count + 10, "Download complete, now indexing all " & storeQuery.Count & " items..."})
                IndexStoreJSON(storeQuery, False)
            End If
        Catch ex As Exception
            Me.Invoke(New MyDelegate(AddressOf PBClose))
            If FullInventory.Count <> 0 Then
                EnableDisableControls(True, New List(Of String)(New String() {"ElementHost2", "txtEmail", "lblEmail", "ElementHost1", "lblPassword", "btnLoad", "btnOffline", "chkSession"}))
            Else
                EnableDisableControls(True)
            End If
            ErrorHandler(System.Reflection.MethodBase.GetCurrentMethod.Name, ex, "Thread URL: " & strThreadURL)
        End Try
    End Sub

    Private Sub DataGridView2_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellClick
        If e.RowIndex = -1 Then Exit Sub
        If e.ColumnIndex = DataGridView2.Columns("Rank").Index Then
            Dim sb As New System.Text.StringBuilder
            If FullStoreInventory(CInt(DataGridView2.Rows(e.RowIndex).Cells("Index").Value)).LevelGem = True Then AddGemWarning(sb)
            Dim strKey As String = DataGridView2.CurrentRow.Cells("ID").Value.ToString & DataGridView2.CurrentRow.Cells("Name").Value.ToString
            If RankExplanation.ContainsKey(strKey) = False Then
                MessageBox.Show("Could not find explanation for this rank. Please report this to:" & Environment.NewLine & Environment.NewLine & _
                                    "https://github.com/RoryTate/modrank/issues", "Key Not Found", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Exit Sub
            End If
            sb.Append(RankExplanation(strKey))
            MessageBox.Show(sb.ToString, "Item Mod Rank Explanation - " & DataGridView2.CurrentRow.Cells("Name").Value.ToString, MessageBoxButtons.OK, MessageBoxIcon.Information)
        ElseIf e.ColumnIndex = DataGridView2.Columns("*").Index AndAlso DataGridView2.CurrentRow.Cells("*").Value.ToString = "*" Then
            If Application.OpenForms().OfType(Of frmResults).Any = False Then frmResults.Show(Me)
            frmResults.Text = "Possible Mod Solutions for '" & DataGridView2.CurrentRow.Cells("Name").Value.ToString & "'"
            frmResults.blStore = True
            Dim tmpList As New CloneableList(Of String)
            tmpList.Add(DataGridView2.CurrentRow.Cells("ID").Value.ToString)
            tmpList.Add(DataGridView2.CurrentRow.Cells("Name").Value.ToString)
            frmResults.MyData = tmpList
        ElseIf DataGridView2.Columns(e.ColumnIndex).Name.ToLower.Contains("prefix") AndAlso NotNull(DataGridView2.Rows(e.RowIndex).Cells(e.ColumnIndex).Value.ToString, "") <> "" Then
            ShowModInfo(DataGridView2, FullStoreInventory(CInt(DataGridView2.Rows(e.RowIndex).Cells("Index").Value)), FullStoreInventory(CInt(DataGridView2.Rows(e.RowIndex).Cells("Index").Value)).ExplicitPrefixMods, CInt(GetNumeric(DataGridView2.Columns(e.ColumnIndex).Name)) - 1, e)
        ElseIf DataGridView2.Columns(e.ColumnIndex).Name.ToLower.Contains("suffix") AndAlso NotNull(DataGridView2.Rows(e.RowIndex).Cells(e.ColumnIndex).Value.ToString, "") <> "" Then
            ShowModInfo(DataGridView2, FullStoreInventory(CInt(DataGridView2.Rows(e.RowIndex).Cells("Index").Value)), FullStoreInventory(CInt(DataGridView2.Rows(e.RowIndex).Cells("Index").Value)).ExplicitSuffixMods, CInt(GetNumeric(DataGridView2.Columns(e.ColumnIndex).Name)) - 1, e)
        ElseIf DataGridView2.Columns(e.ColumnIndex).Name.ToLower.Contains("prefix") AndAlso NotNull(DataGridView2.Rows(e.RowIndex).Cells(e.ColumnIndex).Value.ToString, "") = "" Then
            ShowAllPossibleMods(DataGridView2, FullStoreInventory(CInt(DataGridView2.Rows(e.RowIndex).Cells("Index").Value)), FullStoreInventory(CInt(DataGridView2.Rows(e.RowIndex).Cells("Index").Value)).ExplicitPrefixMods, "Prefix")
        ElseIf DataGridView2.Columns(e.ColumnIndex).Name.ToLower.Contains("suffix") AndAlso NotNull(DataGridView2.Rows(e.RowIndex).Cells(e.ColumnIndex).Value.ToString, "") = "" Then
            ShowAllPossibleMods(DataGridView2, FullStoreInventory(CInt(DataGridView2.Rows(e.RowIndex).Cells("Index").Value)), FullStoreInventory(CInt(DataGridView2.Rows(e.RowIndex).Cells("Index").Value)).ExplicitSuffixMods, "Suffix")
        ElseIf e.ColumnIndex = DataGridView2.Columns("Location").Index Then
            ' Take the user to the GGG web page associated with the ThreadID
            Dim strURL As String = "http://www.pathofexile.com/forum/view-thread/" & DataGridView2.Rows(e.RowIndex).Cells("ThreadID").Value
            Process.Start(strURL)
        ElseIf DataGridView2.Columns(e.ColumnIndex).Name.CompareMultiple(StringComparison.Ordinal, "Sokt", "Link") Then
            MessageBox.Show(DataGridView2.CurrentRow.Cells("SktClrs").Value, "Socket/Link Info", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    Private Sub DataGridView2_CellMouseEnter(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellMouseEnter
        If blScroll = True Then Exit Sub
        If IsValidCellAddress(DataGridView2, e.RowIndex, e.ColumnIndex) AndAlso _
            (DataGridView2.Columns(e.ColumnIndex).Name.Contains("fix") Or _
             DataGridView2.Columns(e.ColumnIndex).Name.CompareMultiple(StringComparison.Ordinal, "Sokt", "Link", "Rank", "Location", "*")) _
         Then DataGridView2.Cursor = Cursors.Hand
    End Sub

    Private Sub DataGridView2_CellMouseLeave(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellMouseLeave
        If blScroll = True Then Exit Sub
        If IsValidCellAddress(DataGridView2, e.RowIndex, e.ColumnIndex) AndAlso _
            (DataGridView2.Columns(e.ColumnIndex).Name.Contains("fix") Or _
             DataGridView2.Columns(e.ColumnIndex).Name.CompareMultiple(StringComparison.Ordinal, "Sokt", "Link", "Rank", "Location", "*")) _
         Then DataGridView2.Cursor = Cursors.Default
    End Sub

    Private Sub DataGridView2_CellMouseMove(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView2.CellMouseMove
        blScroll = False
    End Sub

    Private Sub DataGridView2_CellPainting(sender As Object, e As DataGridViewCellPaintingEventArgs) Handles DataGridView2.CellPainting
        If e.RowIndex = -1 Then
            DataGridViewCellPaintingHeaderFormat(sender, e)
            Exit Sub
        End If
        Dim strName As String = DataGridView2.Columns(e.ColumnIndex).Name.ToLower
        Dim intIndex As Integer = CInt(DataGridView2.Rows(e.RowIndex).Cells("Index").Value)
        If strName.Contains("prefix") AndAlso NotNull(DataGridView2.Rows(e.RowIndex).Cells(e.ColumnIndex).Value, "").ToString <> "" Then
            DataGridViewAddLevelBar(DataGridView2, FullStoreInventory(intIndex).ExplicitPrefixMods, strName, sender, e)
        ElseIf strName.Contains("suffix") AndAlso NotNull(DataGridView2.Rows(e.RowIndex).Cells(e.ColumnIndex).Value, "").ToString <> "" Then
            DataGridViewAddLevelBar(DataGridView2, FullStoreInventory(intIndex).ExplicitSuffixMods, strName, sender, e)
        End If
    End Sub

    Private Sub DataGridView2_RowPostPaint(sender As Object, e As DataGridViewRowPostPaintEventArgs) Handles DataGridView2.RowPostPaint
        If e.RowIndex = -1 Then Exit Sub
        DataGridViewRowPostPaint(DataGridView2, FullStoreInventory, sender, e)
    End Sub

    Private Sub DataGridView2_Scroll(sender As Object, e As ScrollEventArgs) Handles DataGridView2.Scroll
        blScroll = True
    End Sub

    Private Sub DataGridView2_SelectionChanged(sender As Object, e As EventArgs) Handles DataGridView2.SelectionChanged
        If DataGridView2.CurrentCell Is Nothing Then Exit Sub
        If DataGridView2.Cursor <> Cursors.Default Then Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {DataGridView2, "Cursor", Cursors.Default})
        If DataGridView2.Columns(DataGridView2.CurrentCell.ColumnIndex).Name.ToLower.Contains("prefix") Or _
            DataGridView2.Columns(DataGridView2.CurrentCell.ColumnIndex).Name.ToLower.Contains("suffix") Then Me.BeginInvoke(New MyDGDelegate(AddressOf DataGridClearSelection), DataGridView2)
    End Sub

    Private Sub TabControl1_DrawItem(ByVal sender As Object, ByVal e As System.Windows.Forms.DrawItemEventArgs) Handles TabControl1.DrawItem

        'Firstly we'll define some parameters.
        Dim CurrentTab As TabPage = TabControl1.TabPages(e.Index)
        Dim ItemRect As Rectangle = TabControl1.GetTabRect(e.Index)
        Dim ctlColor As Color = SystemColors.Control
        Dim txtColor As Color = SystemColors.WindowText
        Dim FillBrush As New SolidBrush(ctlColor)
        Dim TextBrush As New SolidBrush(txtColor)
        Dim sf As New StringFormat
        sf.Alignment = StringAlignment.Center
        sf.LineAlignment = StringAlignment.Center

        'If we are currently painting the Selected TabItem we'll 
        'change the brush colors and inflate the rectangle.
        If CBool(e.State And DrawItemState.Selected) Then
            FillBrush.Color = ctlColor
            TextBrush.Color = txtColor
            ItemRect.Inflate(2, 2)
        End If

        'Set up rotation for left and right aligned tabs
        If TabControl1.Alignment = TabAlignment.Left Or TabControl1.Alignment = TabAlignment.Right Then
            Dim RotateAngle As Single = 90
            If TabControl1.Alignment = TabAlignment.Left Then RotateAngle = 270
            Dim cp As New PointF(ItemRect.Left + (ItemRect.Width \ 2), ItemRect.Top + (ItemRect.Height \ 2))
            e.Graphics.TranslateTransform(cp.X, cp.Y)
            e.Graphics.RotateTransform(RotateAngle)
            ItemRect = New Rectangle(-(ItemRect.Height \ 2), -(ItemRect.Width \ 2), ItemRect.Height, ItemRect.Width)
        End If

        'Next we'll paint the TabItem with our Fill Brush
        e.Graphics.FillRectangle(FillBrush, ItemRect)

        'Now draw the text.
        e.Graphics.DrawString(CurrentTab.Text, e.Font, TextBrush, RectangleF.op_Implicit(ItemRect), sf)

        'Reset any Graphics rotation
        e.Graphics.ResetTransform()

        'Finally, we should Dispose of our brushes.
        FillBrush.Dispose()
        TextBrush.Dispose()

    End Sub

    Private Sub txtThread_KeyDown(sender As Object, e As KeyEventArgs) Handles txtThread.KeyDown
        If e.KeyCode = Keys.Enter Then
            btnDownloadJSON_Click(sender, New EventArgs())
        End If
    End Sub

    Private Sub btnStoreOffline_Click(sender As Object, e As EventArgs) Handles btnStoreOffline.Click
        LoadCacheThread = New Threading.Thread(AddressOf LoadStoreCache)
        LoadCacheThread.SetApartmentState(Threading.ApartmentState.STA)
        LoadCacheThread.Start()
    End Sub

    Private Sub LoadStoreCache()
        Dim sr As StreamReader
        Dim storeMerge As New List(Of JSON_Store)
        Dim dir = Application.StartupPath & "\Store"
        Try
            EnableDisableControls(False, New List(Of String)(New String() {"pb", "lblpb", "grpProgress"}))
            Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {DataGridView2, "Visible", False})
            Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {lblRecordCount2, "Visible", False})
            Me.Invoke(New pbSetDefaults(AddressOf SetPBDefaults), New Object() {1, "Loading local poexplorer JSONs..."})
            For Each file As String In System.IO.Directory.GetFiles(dir)
                If System.IO.Path.GetExtension(file).ToLower = ".json" Then
                    sr = New StreamReader(file)
                    Dim strLines As String = sr.ReadToEnd()
                    sr.Close()
                    If strLines = "[]" Then Exit For
                    Dim jss As New System.Web.Script.Serialization.JavaScriptSerializer()
                    jss.MaxJsonLength = Int32.MaxValue
                    Dim store As List(Of JSON_Store) = jss.Deserialize(Of List(Of JSON_Store))(strLines)
                    If store.Count = 0 Then Exit For
                    storeMerge.AddRange(store)
                End If
            Next
            Dim storeQuery As IEnumerable(Of JSON_Store) = storeMerge.Where(Function(Item) Item.Rarity_Name = "Rare" AndAlso Item.Name.ToLower <> "" _
                                                   AndAlso Item.Item_Type.ToLower.Contains("map") = False AndAlso Item.Item_Type.ToLower.Contains("peninsula") = False And Item.Identified = True)
            If storeQuery.Count = 0 Then
                Me.Invoke(New MyDelegate(AddressOf PBClose))
                MessageBox.Show("No items found", "No Items Found", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                If FullInventory.Count <> 0 Then
                    EnableDisableControls(True, New List(Of String)(New String() {"ElementHost2", "txtEmail", "lblEmail", "ElementHost1", "lblPassword", "btnLoad", "btnOffline", "chkSession"}))
                Else
                    EnableDisableControls(True)
                End If
                If FullStoreInventory.Count <> 0 Then
                    Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {DataGridView2, "Visible", True})
                    Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {lblRecordCount2, "Visible", True})
                End If
            Else
                Me.Invoke(New pbSetDefaults(AddressOf SetPBDefaults), New Object() {storeQuery.Count + 10, "Finished loading cache, now indexing all " & storeQuery.Count & " items..."})
                IndexStoreJSON(storeQuery, True)
            End If
        Catch ex As Exception
            Me.Invoke(New MyDelegate(AddressOf PBClose))
            If FullInventory.Count <> 0 Then
                EnableDisableControls(True, New List(Of String)(New String() {"ElementHost2", "txtEmail", "lblEmail", "ElementHost1", "lblPassword", "btnLoad", "btnOffline", "chkSession"}))
            Else
                EnableDisableControls(True)
            End If
            ErrorHandler(System.Reflection.MethodBase.GetCurrentMethod.Name, ex)
        End Try
    End Sub

    Private Sub IndexStoreJSON(storeQuery As IEnumerable(Of JSON_Store), blOffline As Boolean)
        Try
            If blOffline Then FullStoreInventory.Clear() : TempStoreInventory.Clear() : dtStore.Clear() : dtStoreOverflow.Clear()
            Dim strThreadID As String = storeQuery(0).Thread_ID
            If blOffline = False And FullStoreInventory.Count > 0 Then
                For i = FullStoreInventory.Count - 1 To 0 Step -1
                    If FullStoreInventory(i).ThreadID = strThreadID Then FullStoreInventory.RemoveAt(i)
                Next
                For i = TempStoreInventory.Count - 1 To 0 Step -1
                    If TempStoreInventory(i).ThreadID = strThreadID Then TempStoreInventory.RemoveAt(i)
                Next
            End If
            Dim lngCounter As Long = 0

            FullStoreInventoryCache.Clear() : TempStoreInventoryCache.Clear()
            Dim strCache As String = Application.StartupPath & "\Store\fsinv.cache"
            If blOffline And File.Exists(strCache) Then LoadCache(FullStoreInventory, strCache) Else LoadCache(FullStoreInventoryCache, strCache)
            strCache = Application.StartupPath & "\Store\tsinv.cache"
            If blOffline And File.Exists(strCache) Then LoadCache(TempStoreInventory, strCache) Else LoadCache(TempStoreInventoryCache, strCache)

            If blOffline And File.Exists(Application.StartupPath & "\Store\fsinv.cache") Then GoTo DataTableLoad

            'Me.Invoke(New MyDelegate(AddressOf PBPerformStep))
            For Each storeItem As JSON_Store In storeQuery
                'MessageBox.Show("Name: " & storeItem.Name & Environment.NewLine & "Mod Count: " & storeItem.Stats.Count)
                Dim storeFullItem As New FullItem
                storeFullItem.ID = storeItem.ID
                storeFullItem.Name = storeItem.Name
                storeFullItem.W = storeItem.W
                storeFullItem.H = storeItem.H
                storeFullItem.Quality = storeItem.Quality
                Select Case storeItem.Item_Type
                    Case "Glove"
                        storeFullItem.GearType = "Gloves"
                    Case "Boot"
                        storeFullItem.GearType = "Boots"
                    Case "BodyArmour"
                        storeFullItem.GearType = "Chest"
                    Case "OneHandAxe"
                        storeFullItem.GearType = "Axe (1h)"
                    Case "OneHandMace"
                        storeFullItem.GearType = "Mace (1h)"
                    Case "OneHandSword", "ThrustingOneHandSword"
                        storeFullItem.GearType = "Sword (1h)"
                    Case "TwoHandSword"
                        storeFullItem.GearType = "Sword (2h)"
                    Case "TwoHandAxe"
                        storeFullItem.GearType = "Axe (2h)"
                    Case "TwoHandMace"
                        storeFullItem.GearType = "Mace (2h)"
                    Case Else
                        storeFullItem.GearType = storeItem.Item_Type
                End Select
                storeFullItem.ItemType = storeItem.Base_Name
                storeFullItem.TypeLine = storeItem.Base_Name
                storeFullItem.Sockets = IIf(storeItem.Socket_Count.HasValue, storeItem.Socket_Count, 0)
                Dim sb As New System.Text.StringBuilder(11), intGroup As Integer = 0, blFirst As Boolean = True
                For i = 0 To storeFullItem.Sockets - 1
                    If intGroup = storeItem.Sockets(i).Group And blFirst = False Then
                        sb.Append("-")
                    ElseIf blFirst = False Then
                        sb.Append(" ")
                    End If
                    blFirst = False
                    sb.Append(storeItem.Sockets(i).Attr)
                    intGroup = storeItem.Sockets(i).Group
                Next
                sb.Replace("I", "b") : sb.Replace("S", "r") : sb.Replace("D", "g")
                storeFullItem.Colours = sb.ToString
                If storeItem.Linked_Socket_Count.HasValue Then
                    storeFullItem.Links = IIf(storeItem.Linked_Socket_Count = 1, 0, storeItem.Linked_Socket_Count)
                Else
                    storeFullItem.Links = 0
                End If
                storeFullItem.Level = CByte(IIf(storeItem.Level.HasValue, storeItem.Level, 1))
                storeFullItem.Arm = IIf(IsNothing(storeItem.Armour), 0, storeItem.Armour)
                storeFullItem.Eva = IIf(IsNothing(storeItem.Evasion), 0, storeItem.Evasion)
                storeFullItem.ES = IIf(IsNothing(storeItem.Energy_Shield), 0, storeItem.Energy_Shield)
                storeFullItem.League = StrConv(storeItem.League_Name, vbProperCase)
                storeFullItem.Location = storeItem.Account
                storeFullItem.Rarity = Rarity.Rare
                storeFullItem.Corrupted = storeItem.Corrupted
                storeFullItem.Price = storeItem.Price
                storeFullItem.ThreadID = storeItem.Thread_ID
                Dim queryImp As IEnumerable(Of Stats) = storeItem.Stats.Where(Function(s) s.Hidden = False And s.Implicit = True)
                For Each s As Stats In queryImp
                    Dim newImpMod As New FullMod
                    newImpMod.FullText = s.Name
                    If s.Name.IndexOf("-", StringComparison.OrdinalIgnoreCase) > -1 Then
                        newImpMod.Value1 = GetNumeric(s.Name, 0, s.Name.IndexOf("-", StringComparison.OrdinalIgnoreCase))
                        newImpMod.MaxValue1 = GetNumeric(s.Name, s.Name.IndexOf("-", StringComparison.OrdinalIgnoreCase), s.Name.Length)
                    Else
                        newImpMod.Value1 = GetNumeric(s.Name)
                    End If
                    newImpMod.Type1 = GetChars(s.Name)
                    storeFullItem.ImplicitMods.Add(newImpMod)
                Next

                Dim query As IEnumerable(Of Stats) = storeItem.Stats.Where(Function(s) s.Hidden = False And s.Implicit = False)
                Dim lstStr As New List(Of String)
                For Each Stat In query
                    lstStr.Add(Stat.Name)
                Next
                Dim blIndexed As Boolean = False
                If FullStoreInventoryCache.Count <> 0 Then
                    Dim q As IEnumerable(Of FullItem) = FullStoreInventoryCache.Where(Function(s) s.ID = storeItem.ID And s.ThreadID = storeItem.Thread_ID And s.Name = storeItem.Name)
                    If q.Count <> 0 Then
                        Dim testMod As New CloneableList(Of FullMod)
                        For Each m As CloneableList(Of FullMod) In {q(0).ExplicitPrefixMods, q(0).ExplicitSuffixMods}
                            For Each myMod As FullMod In m
                                Dim newMod As New FullMod
                                Dim qry As IEnumerable(Of FullMod) = testMod.Where(Function(x) x.Type1 = myMod.Type1)
                                If qry.Count = 0 Then
                                    newMod.Type1 = myMod.Type1
                                    newMod.Value1 = myMod.Value1
                                    newMod.MaxValue1 = myMod.MaxValue1
                                    testMod.Add(newMod)
                                Else
                                    testMod(testMod.IndexOf(qry(0))).Value1 += myMod.Value1
                                End If
                                If myMod.Type2 <> "" Then
                                    Dim newMod2 As New FullMod
                                    qry = testMod.Where(Function(x) x.Type1 = myMod.Type2)
                                    If qry.Count = 0 Then
                                        newMod2.Type1 = myMod.Type2
                                        newMod2.Value1 = myMod.Value2
                                        testMod.Add(newMod2)
                                    Else
                                        testMod(testMod.IndexOf(qry(0))).Value1 += myMod.Value2
                                    End If
                                End If
                            Next
                        Next
                        For Each m As Stats In storeItem.Stats.Where(Function(s) s.Hidden = False And s.Implicit = False)
                            Dim strName As String = GetChars(m.Name)
                            Dim sngValue As Single = GetNumeric(m.Name)
                            Dim qry As IEnumerable(Of FullMod) = testMod.Where(Function(x) x.Type1 = strName)
                            If qry.Count = 0 Then
                                GoTo IndexMod
                            Else
                                If testMod(testMod.IndexOf(qry(0))).Value1 <> sngValue Then
                                    If m.Name.Contains("-") Then
                                        If testMod(testMod.IndexOf(qry(0))).Value1 <> GetNumeric(m.Name, 0, m.Name.IndexOf("-", StringComparison.OrdinalIgnoreCase)) Then GoTo IndexMod
                                        If testMod(testMod.IndexOf(qry(0))).MaxValue1 <> GetNumeric(m.Name, m.Name.IndexOf("-", StringComparison.OrdinalIgnoreCase), m.Name.Length) Then GoTo IndexMod
                                    Else
                                        GoTo IndexMod
                                    End If
                                End If
                            End If
                        Next
                        storeFullItem.Rank = q(0).Rank
                        storeFullItem.Percentile = q(0).Percentile
                        storeFullItem.ExplicitPrefixMods = q(0).ExplicitPrefixMods.Clone
                        storeFullItem.ExplicitSuffixMods = q(0).ExplicitSuffixMods.Clone
                        storeFullItem.OtherSolutions = q(0).OtherSolutions
                        FullStoreInventory.Add(storeFullItem)
                        If storeFullItem.OtherSolutions = True Then
                            Dim q2 As IEnumerable(Of FullItem) = TempStoreInventoryCache.Where(Function(s) s.ID = storeItem.ID And s.ThreadID = storeItem.Thread_ID And s.Name = storeItem.Name)
                            For Each qItem In q2
                                Dim tmpItem As FullItem = CType(qItem.Clone, FullItem)
                                TempStoreInventory.Add(tmpItem)
                            Next
                            lngStoreCount = TempStoreInventory.Count
                        End If
                        blIndexed = True
                    End If
                End If
                If blIndexed = True Then GoTo IncrementProgressBar
IndexMod:
                ' Make sure that increased item rarity comes last, since it can be either a prefix or a suffix
                ReorderExplicitMods(lstStr, lstStr.Count)
                blAddedOne = False
                If Not IsNothing(lstStr) Then
                    blSolomonsJudgment.Clear()
                    EvaluateExplicitMods(lstStr, lstStr.Count, storeItem.ID.ToString, storeItem.Name, storeFullItem, TempStoreInventory)
                    If storeFullItem.ExplicitPrefixMods.Count = 0 And storeFullItem.ExplicitSuffixMods.Count = 0 Then
                        If TempStoreInventory.Count = 0 Then
                            EvaluateExplicitMods(lstStr, lstStr.Count, storeItem.ID.ToString, storeItem.Name, storeFullItem, TempStoreInventory, True, True)       ' We didn't find anything, so run again, allowing legacy values to be set
                        Else
                            If TempStoreInventory(TempStoreInventory.Count - 1).Name <> storeItem.Name Then
                                EvaluateExplicitMods(lstStr, lstStr.Count, storeItem.ID.ToString, storeItem.Name, storeFullItem, TempStoreInventory, True, True)       ' We didn't find anything, so run again, allowing legacy values to be set
                            End If
                        End If
                    End If
                    If blAddedOne = True Then
                        ' Check if we've already added this rankexplanation key to the dictionary
                        'If RankExplanation.ContainsKey(storeItem.ID.ToString & storeItem.Name) = True Then
                        '    RankExplanation(storeItem.ID.ToString & storeItem.Name) = RankExplanation(storeItem.ID.ToString & storeItem.Name & TempStoreInventory.Count - 1)
                        'Else
                        '    ' First rename the associated entry in the RankExplanation dictionary...(have to add it in with the proper name and then remove the old one)
                        '    RankExplanation.Add(storeItem.ID.ToString & storeItem.Name, RankExplanation(storeItem.ID.ToString & storeItem.Name & TempStoreInventory.Count - 1))
                        'End If
                        'RankExplanation.Remove(storeItem.ID.ToString & storeItem.Name & TempStoreInventory.Count - 1)
                        FullStoreInventory.Add(TempStoreInventory(TempStoreInventory.Count - 1).Clone)
                        TempStoreInventory.RemoveAt(TempStoreInventory.Count - 1)
                        If lngStoreCount < TempStoreInventory.Count Then
                            FullStoreInventory(FullStoreInventory.Count - 1).OtherSolutions = True
                        End If
                        lngStoreCount = TempStoreInventory.Count
                    Else
                        FullStoreInventory.Add(storeFullItem)
                    End If
                End If
IncrementProgressBar:
                lngCounter += 1
                If lngCounter Mod 10 = 0 Then
                    Me.Invoke(New MyDelegate(AddressOf PBPerformStep))
                End If
            Next

DataTableLoad:
            RecalculateAllRankings(True, , , True)
            If dtStore.Rows.Count > 0 Then dtStore.Clear()
            If dtStoreOverflow.Rows.Count > 0 Then dtStoreOverflow.Clear()
            Me.Invoke(New pbSetDefaults(AddressOf SetPBDefaults), New Object() {FullStoreInventory.Count, "Adding Store inventories to datatable"})
            AddToDataTable(FullStoreInventory, dtStore, False, "")
            strCache = Application.StartupPath & "\Store\fsinv.cache"
            If blOffline = False Or File.Exists(strCache) = False Then
                WriteCache(FullStoreInventory, strCache)
            End If

            Me.Invoke(New pbSetDefaults(AddressOf SetPBDefaults), New Object() {TempStoreInventory.Count, "Adding Store overflow inventories to datatable"})
            AddToDataTable(TempStoreInventory, dtStoreOverflow, False, "")
            ' Write the TempStoreInventory into the tsinv.cache file as a serialized JSON
            strCache = Application.StartupPath & "\Store\tsinv.cache"
            If blOffline = False Or File.Exists(strCache) = False Then
                WriteCache(TempStoreInventory, strCache)
            End If

            Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {DataGridView2, "DataSource", dtStore})
            Me.Invoke(New MyDualControlDelegate(AddressOf HideColumns), New Object() {Me, DataGridView2})
            Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {DataGridView2.Columns("%").DefaultCellStyle, "Format", "n1"})
            Me.Invoke(New MyDGDelegate(AddressOf SortDataGridView), DataGridView2)
            'Me.Invoke(New MyDelegate(AddressOf SortDataGridView))
            Me.Invoke(New MyControlDelegate(AddressOf SetDataGridViewWidths), New Object() {DataGridView2})
            ' To make room for new Price column take away from SubType and Implicit columns
            Dim intWidth As Integer = Math.Max(CType(Me.Invoke(New RCPD(AddressOf ReadControlProperty), New Object() {DataGridView2.Columns("SubType"), "Width"}), Integer) - 15, 0)
            Me.Invoke(New UCPD(AddressOf SetControlProperty), New Object() {DataGridView2.Columns("SubType"), "Width", intWidth})
            intWidth = Math.Max(CType(Me.Invoke(New RCPD(AddressOf ReadControlProperty), New Object() {DataGridView2.Columns("Implicit"), "Width"}), Integer) - 35, 0)
            Me.Invoke(New UCPD(AddressOf SetControlProperty), New Object() {DataGridView2.Columns("Implicit"), "Width", intWidth})
            Dim FirstCell As DataGridViewCell
            FirstCell = Me.Invoke(New DataGridCell(AddressOf ReturnFirstCell), New Object() {DataGridView2})
            Me.Invoke(New UCPD(AddressOf SetControlProperty), New Object() {DataGridView2, "FirstDisplayedCell", FirstCell})

            Me.Invoke(New MyDelegate(AddressOf PBClose))
            Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {Me, "WindowState", FormWindowState.Maximized})

            Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {DataGridView2, "Visible", True})
            Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {gpLegend, "Left", 550})
            Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {gpLegend, "Visible", True})
            Me.Invoke(New MyDelegate(AddressOf BringLegendToFront))
            Dim intRows As Integer = CInt(Me.Invoke(New RCPD(AddressOf ReadControlProperty), New Object() {DataGridView2.Rows, "Count"}).ToString)
            Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {lblRecordCount2, "Text", "Number of Rows: " & intRows})
            Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {lblRecordCount2, "Visible", True})
            Me.BeginInvoke(New MyDGDelegate(AddressOf SetDataGridFocus), DataGridView2)

            If strStoreFilter.Trim <> "" Then ApplyFilter(dtStore, dtStoreOverflow, DataGridView2, lblRecordCount2, strStoreOrderBy, strStoreFilter)
            If FullInventory.Count <> 0 Then
                EnableDisableControls(True, New List(Of String)(New String() {"ElementHost2", "txtEmail", "lblEmail", "ElementHost1", "lblPassword", "btnLoad", "btnOffline", "chkSession"}))
            Else
                EnableDisableControls(True)
            End If

        Catch ex As Exception
            Me.Invoke(New MyDelegate(AddressOf PBClose))
            If FullInventory.Count <> 0 Then
                EnableDisableControls(True, New List(Of String)(New String() {"ElementHost2", "txtEmail", "lblEmail", "ElementHost1", "lblPassword", "btnLoad", "btnOffline", "chkSession"}))
            Else
                EnableDisableControls(True)
            End If
            ErrorHandler(System.Reflection.MethodBase.GetCurrentMethod.Name, ex)
        End Try
    End Sub

    Private Sub btnShopSearch_Click(sender As Object, e As EventArgs) Handles btnShopSearch.Click
        frmShopFilter.blStore = True
        frmShopFilter.ShowDialog(Me)
    End Sub

    Private Sub TabControl1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TabControl1.SelectedIndexChanged
        If DataGridView1.Visible = True Then DataGridView1.PerformLayout() : DataGridView1.Focus()
        If DataGridView2.Visible = True Then DataGridView2.PerformLayout() : DataGridView2.Focus()
    End Sub

    Private Sub btnRefresh_Click(sender As Object, e As EventArgs) Handles btnRefresh.Click
        Dim dlgResult As DialogResult
        dlgResult = MessageBox.Show("Refreshing all the downloaded store caches may take some time. Are you sure?", "Please Confirm", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If dlgResult = Windows.Forms.DialogResult.No Then Exit Sub
        RefreshCacheThread = New Threading.Thread(AddressOf RefreshStoreCache)
        RefreshCacheThread.SetApartmentState(Threading.ApartmentState.STA)
        RefreshCacheThread.Start()
    End Sub

    Public Sub RefreshStoreCache()
        Try
            Dim dir = Application.StartupPath & "\Store"
            EnableDisableControls(False, New List(Of String)(New String() {"pb", "lblpb", "grpProgress"}))
            Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {DataGridView2, "Visible", False})
            Me.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {lblRecordCount2, "Visible", False})
            Me.Invoke(New pbSetDefaults(AddressOf SetPBDefaults), New Object() {1, "Indexing local poexplorer JSONs..."})
            Dim lstThread As New List(Of String)
            For Each file As String In System.IO.Directory.GetFiles(dir)
                Dim strTemp As String = System.IO.Path.GetFileNameWithoutExtension(file).Substring(0, System.IO.Path.GetFileNameWithoutExtension(file).IndexOf("-"))
                If lstThread.Contains(strTemp) = False Then lstThread.Add(strTemp)
            Next
            For Each strThread In lstThread
                RemoveAllStoreJSONs(strThread)
                DownloadJSON("http://poexplorer.com/threads/" & strThread & ".json", strThread, False)
            Next
            btnStoreOffline_Click(btnStoreOffline, EventArgs.Empty)
        Catch ex As Exception
            ErrorHandler(System.Reflection.MethodBase.GetCurrentMethod.Name, ex)
        End Try
    End Sub

    Public Sub RemoveAllStoreJSONs(strStore As String)
        Try
            Dim dir = Application.StartupPath & "\Store"
            For i = 1 To 1000
                If File.Exists(dir & "\" & strStore & "-" & i & ".json") = True Then
                    File.Delete(dir & "\" & strStore & "-" & i & ".json")
                Else
                    Exit Sub
                End If
            Next
        Catch ex As Exception
            ErrorHandler(System.Reflection.MethodBase.GetCurrentMethod.Name, ex)
        End Try
    End Sub

    Public Sub LoadCache(ByRef lstInv As List(Of FullItem), strCache As String)
        If File.Exists(strCache) Then   ' If we have a cache, load it to speed up mod indexing
            Dim sr As StreamReader = New StreamReader(strCache)
            Dim strLines As String = sr.ReadToEnd()
            sr.Close()
            Dim jssTemp As New System.Web.Script.Serialization.JavaScriptSerializer()
            jssTemp.MaxJsonLength = Int32.MaxValue
            lstInv = jssTemp.Deserialize(Of List(Of FullItem))(strLines)
        End If
    End Sub

    Public Sub WriteCache(ByRef lstInv As List(Of FullItem), strCache As String)
        ' Write the FullStoreInventory into the fsinv.cache file as a serialized JSON
        If File.Exists(strCache) Then File.Delete(strCache)
        Dim jss As New System.Web.Script.Serialization.JavaScriptSerializer()
        jss.MaxJsonLength = Int32.MaxValue
        System.IO.File.WriteAllText(strCache, jss.Serialize(lstInv))
    End Sub
End Class