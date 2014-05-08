﻿Imports POEApi.Infrastructure
Imports POEApi.Model
Imports System
Imports System.Collections.Generic
Imports System.Globalization
Imports System.Linq
Imports System.Xml.Linq
Imports System.Security

Module mdlSettings
    Public originalDoc As XElement
    Public strLocation As String = "settings.xml"
    Public blFormChanged As Boolean = False
    Public Email As String = ""
    Public useSession As Boolean = False
    Public strWeightFile As String = ""

    Public Model As New POEApi.Model.POEModel

    Public ColorHasGem As Color = Color.LightGreen
    Public ColorUnknownValue As Color = Color.Teal
    Public ColorOtherSolutions As Color = Color.LightSalmon
    Public ColorNoModVariation As Color = Color.LightBlue
    Public ColorMax = Color.Blue
    Public ColorMin = Color.Red
    Public intWeightMax As Integer = 9
    Public intWeightMin As Integer = 1

    Public Property UserSettings() As Dictionary(Of String, String)
        Get
            Return m_UserSettings
        End Get
        Private Set(value As Dictionary(Of String, String))
            m_UserSettings = value
        End Set
    End Property

    Private m_UserSettings As Dictionary(Of String, String)

    Public Property ProxySettings() As Dictionary(Of String, String)
        Get
            Return m_ProxySettings
        End Get
        Private Set(value As Dictionary(Of String, String))
            m_ProxySettings = value
        End Set
    End Property

    Private m_ProxySettings As Dictionary(Of String, String)

    Public Sub LoadSettings()
        Try
            originalDoc = XElement.Load(strLocation)
            UserSettings = originalDoc.Elements("UserSettings").Descendants().ToDictionary(Function(setting) setting.Attribute("name").Value, Function(setting) setting.Attribute("value").Value)
            ProxySettings = originalDoc.Elements("ProxySettings").Descendants().ToDictionary(Function(setting) setting.Attribute("name").Value, Function(setting) setting.Attribute("value").Value)
        Catch ex As Exception

        End Try
    End Sub

    Public Sub saveSettings(password As SecureString)
        Settings.UserSettings("AccountLogin") = Email
        Settings.UserSettings("AccountPassword") = password.Encrypt()
        Settings.UserSettings("UseSessionID") = useSession.ToString()
        Settings.UserSettings("WeightFile") = strWeightFile
        Settings.Save()
    End Sub

    Function GetEmbeddedIcon(ByVal strName As String) As Icon
        Return New Icon(System.Reflection.Assembly.GetExecutingAssembly.GetManifestResourceStream(strName))
    End Function

    Function GetEmbeddedBitmap(ByVal strName As String) As Bitmap
        Return New Bitmap(System.Reflection.Assembly.GetExecutingAssembly.GetManifestResourceStream(strName))
    End Function

    Public Sub ErrorHandler(ByVal strName As String, ByVal ex As Exception, Optional ByVal strExtraText As String = "")
        MessageBox.Show("Error in " & strName & ": " & Environment.NewLine & ex.Message & IIf(strExtraText <> "", Environment.NewLine & Environment.NewLine & strExtraText, ""), _
                        "ModRank", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly)
    End Sub
End Module
