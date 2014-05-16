Imports POEApi.Model
Imports ModRank.Extenders.Strings

Module mdlMain
    Public dtRank As DataTable = New DataTable()
    Public dtRankFilter As DataTable = New DataTable()
    Public dtOverflow As DataTable = New DataTable
    Public bytDisplay As Byte = 21     ' Don't display column index's past this in the datagrid...also used by the search filter when populating combobox's
    Public dtMods As DataTable = New DataTable()
    Public dtWeights As DataTable = New DataTable()
    Public bytColumns As Byte = 0       ' The max number of columns displayed in the datagridview
    Public strD As String = ";"     ' The delimiter used in the filter...use one that will not be likely to appear in SQL
    Public strFilter As String = ""
    Public strOrderBy As String = ""
    Public strRawFilter As String = ""
    Public lstTotalTypes As New List(Of String)

    ' RCPD = Read Control Property Delegate (used a lot so abbreviated)
    Public Delegate Function RCPD(ByVal MyControl As Object, ByVal MyProperty As Object) As String
    ' UCPD = Update Control Property Delegate
    Public Delegate Sub UCPD(ByVal MyControl As Object, ByVal MyProperty As Object, ByVal MyValue As Object)
    Public Delegate Sub MyControlDelegate(myControl As Object)
    Public Delegate Sub MyDualControlDelegate(myControl As Object, myControl2 As Object)

    Public Sub SetDataGridViewWidths(dg As DataGridView)
        Try
            Dim intCounter As Integer = 0
            For Each MyWidth In Split(UserSettings("RowWidths"), ",")
                If intCounter > dg.Columns.Count - 1 Then Exit For
                If IsNumeric(MyWidth) = False Then Continue For
                If CInt(MyWidth) <= 0 Then
                    dg.Columns(intCounter).Visible = False
                Else
                    dg.Columns(intCounter).Width = CInt(MyWidth)
                End If
                intCounter += 1
            Next
            If intCounter < bytDisplay Then     ' We're missing some new columns...the settings.xml needs to be updated for this field with new defaults
                dg.Columns(bytDisplay - 1).Width = 30
                dg.Columns(bytDisplay - 2).Width = 25
                Settings.UserSettings("RowWidths") = UserSettings("RowWidths").Substring(0, UserSettings("RowWidths").LastIndexOf(",")) & ",25,30"
            End If
        Catch ex As Exception
            ErrorHandler(System.Reflection.MethodBase.GetCurrentMethod.Name, ex)
        End Try
    End Sub

    Public Sub HideColumns(frm As Form, dg As DataGridView)
        bytColumns = CByte(frm.Invoke(New RCPD(AddressOf ReadControlProperty), New Object() {dg.Columns, "Count"}).ToString)
        For i = bytDisplay To bytColumns - 1
            frm.BeginInvoke(New UCPD(AddressOf SetControlProperty), New Object() {dg.Columns(i), "Visible", False})
        Next
    End Sub

    Public Function ReadControlProperty(ByVal MyControl As Object, ByVal MyProperty As Object) As String
        Dim p As System.Reflection.PropertyInfo = MyControl.GetType().GetProperty(DirectCast(MyProperty, String))
        ReadControlProperty = p.GetValue(MyControl, Nothing).ToString
    End Function

    Public Sub SetControlProperty(ByVal MyControl As Object, ByVal MyProperty As Object, ByVal MyValue As Object)
        Dim p As System.Reflection.PropertyInfo = MyControl.GetType().GetProperty(DirectCast(MyProperty, String))
        p.SetValue(MyControl, MyValue, Nothing)
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
            Dim strLevel As String = IIf(newFullItem.Level <> 0, "AND Level<=" & Math.Round(1.25 * (newFullItem.Level + 0.49)), "").ToString
            Dim strDescription As String = ""
            If strForceDescription <> "" Then
                strDescription = strForceDescription
            Else
                strDescription = result("Description").ToString
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

    Public Function ConvertFilterToSQL(strRawFilter As String) As String
        Dim sb As New System.Text.StringBuilder
        Dim strLines() As String = strRawFilter.Split(CChar(vbCrLf))
        lstTotalTypes.Clear()
        For Each strLine In strLines
            Dim strElement() As String = strLine.Split(CChar(strD))
            If strElement(1).CompareMultiple(StringComparison.Ordinal, "[Prefix Type]", "[Suffix Type]", "[Number of Prefixes]", _
                                             "[Number of Suffixes]", "[Implicit Type]", "[Number of Implicits]", "[Mod Total Value]", "[Total Number of Explicits]") = True Then
                Dim strAndOr As String = IIf(strElement(2) = "<>", " AND ", " OR ").ToString ' For <>, the multiple field OR conditions must become an AND condition to function correctly
                Select Case strElement(1)   ' Element at second position is the datatable field
                    Case "[Number of Prefixes]"
                        strElement(1) = "[pcount]"
                    Case "[Number of Suffixes]"
                        strElement(1) = "[scount]"
                    Case "[Number of Implicits]"
                        strElement(1) = "[icount]"
                    Case "[Total Number of Explicits]"
                        strElement(1) = "[pcount]+[scount]"
                    Case "[Prefix Type]"
                        sb.Append(strElement(0) & "([pft1]" & strElement(2) & strElement(3) & strAndOr & _
                                                "[pft2]" & strElement(2) & strElement(3) & strAndOr & _
                                                "[pft3]" & strElement(2) & strElement(3) & ")" & strElement(4) & strElement(5))
                        Continue For
                    Case "[Suffix Type]"
                        sb.Append(strElement(0) & "([sft1]" & strElement(2) & strElement(3) & strAndOr & _
                                                "[sft2]" & strElement(2) & strElement(3) & strAndOr & _
                                                "[sft3]" & strElement(2) & strElement(3) & ")" & strElement(4) & strElement(5))
                        Continue For
                    Case "[Implicit Type]"
                        sb.Append(strElement(0) & "([it1]" & strElement(2) & strElement(3) & strAndOr & _
                                                "[it2]" & strElement(2) & strElement(3) & strAndOr & _
                                                "[it3]" & strElement(2) & strElement(3) & ")" & strElement(4) & strElement(5))
                        Continue For
                    Case "[Mod Total Value]"
                        ' Check to see if there is a "/" in the value...if so, have to look at value1 and maxvalue1 for type1
                        If strElement(3).Contains("/") = True Then
                            Dim intLow As Integer = 0, intHigh As Integer = 0, blLow As Boolean = True, blHigh As Boolean = True
                            If IsNumeric(strElement(3).Substring(0, strElement(3).IndexOf("/"))) Then intLow = CInt(strElement(3).Substring(0, strElement(3).IndexOf("/"))) Else blLow = False
                            If IsNumeric(strElement(3).Substring(strElement(3).IndexOf("/") + 1, strElement(3).Length - strElement(3).IndexOf("/") - 1)) Then _
                                intHigh = CInt(strElement(3).Substring(strElement(3).IndexOf("/") + 1, strElement(3).Length - strElement(3).IndexOf("/") - 1)) Else blHigh = False
                            If blLow Then sb.Append(strElement(0) & "IIF([pt1]=" & strElement(4) & ",[pv1],0) + " & _
                              "IIF([pt2]=" & strElement(4) & ",[pv2],0) + " & _
                              "IIF([pt3]=" & strElement(4) & ",[pv3],0) + " & _
                              "IIF([st1]=" & strElement(4) & ",[sv1],0) + " & _
                              "IIF([st2]=" & strElement(4) & ",[sv2],0) + " & _
                              "IIF([st3]=" & strElement(4) & ",[sv3],0) + " & _
                              "IIF([it1]=" & strElement(4) & ",[iv1],0) + " & _
                              "IIF([it2]=" & strElement(4) & ",[iv2],0) + " & _
                              "IIF([it3]=" & strElement(4) & ",[iv3],0)" & _
                              strElement(2) & intLow & strElement(5) & IIf(blHigh, " AND ", strElement(6)).ToString)
                            If blHigh Then sb.Append(strElement(0) & "IIF([pt1]=" & strElement(4) & ",[pv1m],0) + " & _
                              "IIF([pt2]=" & strElement(4) & ",[pv2m],0) + " & _
                              "IIF([pt3]=" & strElement(4) & ",[pv3m],0) + " & _
                              "IIF([st1]=" & strElement(4) & ",[sv1m],0) + " & _
                              "IIF([st2]=" & strElement(4) & ",[sv2m],0) + " & _
                              "IIF([st3]=" & strElement(4) & ",[sv3m],0) + " & _
                              "IIF([it1]=" & strElement(4) & ",[iv1m],0) + " & _
                              "IIF([it2]=" & strElement(4) & ",[iv2m],0) + " & _
                              "IIF([it3]=" & strElement(4) & ",[iv3m],0)" & _
                              strElement(2) & intHigh & strElement(5) & strElement(6))
                        Else
                            sb.Append(strElement(0) & "IIF([pt1]=" & strElement(4) & ",[pv1],0) + " & _
                                  "IIF([pt12]=" & strElement(4) & ",[pv12],0) + " & _
                                  "IIF([pt2]=" & strElement(4) & ",[pv2],0) + " & _
                                  "IIF([pt22]=" & strElement(4) & ",[pv22],0) + " & _
                                  "IIF([pt3]=" & strElement(4) & ",[pv3],0) + " & _
                                  "IIF([pt32]=" & strElement(4) & ",[pv32],0) + " & _
                                  "IIF([st1]=" & strElement(4) & ",[sv1],0) + " & _
                                  "IIF([st12]=" & strElement(4) & ",[sv12],0) + " & _
                                  "IIF([st2]=" & strElement(4) & ",[sv2],0) + " & _
                                  "IIF([st22]=" & strElement(4) & ",[sv22],0) + " & _
                                  "IIF([st3]=" & strElement(4) & ",[sv3],0) + " & _
                                  "IIF([st32]=" & strElement(4) & ",[sv32],0) + " & _
                                  "IIF([it1]=" & strElement(4) & ",[iv1],0) + " & _
                                  "IIF([it2]=" & strElement(4) & ",[iv2],0) + " & _
                                  "IIF([it3]=" & strElement(4) & ",[iv3],0)" & _
                                  strElement(2) & strElement(3) & strElement(5) & strElement(6))
                        End If
                        lstTotalTypes.Add(strElement(4).Substring(1, strElement(4).Length - 2))
                        Continue For
                End Select
            End If
            sb.Append(String.Join(" ", strElement) & " ")
        Next
        ConvertFilterToSQL = sb.ToString
    End Function
End Module
