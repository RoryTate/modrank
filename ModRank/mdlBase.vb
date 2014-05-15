Module mdlBase
    Public Function GetNumeric(str As String, Optional intStart As Integer = 0, Optional intEnd As Integer = 0) As Single
        If IsNumeric(str) Then Return Long.Parse(str.ToString) : Exit Function
        If intStart < 0 Then intStart = 0
        If intEnd < 0 Then intEnd = 0
        Dim blLimit As Boolean = IIf(intStart > 0 Or intEnd > 0, True, False)
        Dim sb As New System.Text.StringBuilder(str.Length)
        Dim intCounter As Integer = 0
        For Each ch As Char In str
            If blLimit = False Or (blLimit And intCounter > intStart And intCounter < intEnd) Then
                If Char.IsDigit(ch) Or ch.ToString = "." Then sb.Append(ch)
            End If
            intCounter += 1
        Next
        Return Single.Parse(IIf(sb.ToString = "", 0, sb.ToString))
    End Function

    Public Function GetChars(str As String) As String
        Dim sb As New System.Text.StringBuilder(str.Length)
        For Each ch As Char In str
            If Char.IsDigit(ch) = False Then sb.Append(ch)
        Next
        ' Remove dash (-) and double space
        sb.Replace("-", "")
        While sb.ToString.IndexOf("  ", StringComparison.OrdinalIgnoreCase) > -1
            sb.Replace("  ", " ")
        End While
        Return Trim(sb.ToString)
    End Function

    Public Function NotNull(Of T)(ByVal Value As T, ByVal DefaultValue As T) As T
        If Value Is Nothing OrElse IsDBNull(Value) Then
            Return DefaultValue
        Else
            Return Value
        End If
    End Function

    Public Function DeepCopyDataRow(ByRef dr As DataRow) As DataRow
        Dim dt As DataTable = dr.Table.Clone
        dt.Rows.Clear()
        dt.ImportRow(dr)
        Return dt.Rows(0)
    End Function

    Public Function LoadCSVtoDataTable(ByVal strLocation As String) As DataTable
        Try
            Dim result As New DataTable
            If System.IO.File.Exists(strLocation) Then
                Dim parser As New Microsoft.VisualBasic.FileIO.TextFieldParser(strLocation)
                parser.Delimiters = New String() {","}
                parser.HasFieldsEnclosedInQuotes = True 'use if data may contain delimiters 
                parser.TextFieldType = FileIO.FieldType.Delimited
                parser.TrimWhiteSpace = True
                Dim HeaderFlag As Boolean = True
                While Not parser.EndOfData
                    If AddValuesToTable(parser.ReadFields, result, HeaderFlag) Then
                        HeaderFlag = False
                    Else
                        parser.Close()
                        Return Nothing
                    End If
                End While
                parser.Close()
                Return result
            Else : Return Nothing
            End If
        Catch ex As Exception
            Console.WriteLine(ex.ToString())
            Return Nothing
        End Try
    End Function

    Public Function AddValuesToTable(ByRef source() As String, ByVal destination As DataTable, Optional ByVal HeaderFlag As Boolean = False) As Boolean
        'Ensures a datatable can hold an array of values and then adds a new row 
        Try
            Dim existing As Integer = destination.Columns.Count
            If HeaderFlag Then
                Resolve_Duplicate_Names(source)
                For i As Integer = 0 To source.Length - existing - 1
                    destination.Columns.Add(source(i).ToString, GetType(String))
                Next i
                Return True
            End If
            For i As Integer = 0 To source.Length - existing - 1
                destination.Columns.Add("Column" & (existing + 1 + i).ToString, GetType(String))
            Next
            destination.Rows.Add(source)
            For i = 0 To destination.Columns.Count - 1
                If IsDBNull(destination(destination.Rows.Count - 1).Item(i)) Then
                    destination(destination.Rows.Count - 1).Item(i) = ""
                End If
            Next
            Return True
        Catch ex As Exception
            Console.WriteLine(ex.ToString())
            Return False
        End Try
    End Function

    Public Sub Resolve_Duplicate_Names(ByRef source() As String)
        Try
            ' Resolves the possibility of duplicated names by appending "Duplicate Name" and a number at the end of any duplicates
            Dim i, n, dnum As Integer
            dnum = 1
            For n = 0 To source.Length - 1
                For i = n + 1 To source.Length - 1
                    If source(i) = source(n) Then
                        source(i) = source(i) & "Duplicate Name " & dnum
                        dnum += 1
                    End If
                Next
            Next
            Return
        Catch ex As Exception
            Console.WriteLine(ex.ToString())
        End Try
    End Sub

    Public Function TableToCSV(ByVal sourceTable As DataTable, ByVal filePathName As String, Optional ByVal HasHeader As Boolean = False) As Boolean
        'Writes a datatable back into a csv 
        Try
            Dim sb As New System.Text.StringBuilder
            If HasHeader Then
                Dim nameArray(200) As Object
                Dim i As Integer = 0
                For Each col As DataColumn In sourceTable.Columns
                    nameArray(i) = CType(col.ColumnName, Object)
                    i += 1
                Next col
                ReDim Preserve nameArray(i - 1)
                sb.AppendLine(String.Join(",", Array.ConvertAll(Of Object, String)(nameArray, _
                                Function(o As Object) If(o.ToString.Contains(","), _
                                ControlChars.Quote & o.ToString & ControlChars.Quote, o))))
            End If
            For Each dr As DataRow In sourceTable.Rows
                sb.AppendLine(String.Join(",", Array.ConvertAll(Of Object, String)(dr.ItemArray, _
                                Function(o As Object) If(o.ToString.Contains(","), _
                                ControlChars.Quote & o.ToString & ControlChars.Quote, o.ToString))))
            Next
            System.IO.File.WriteAllText(filePathName, sb.ToString)
            Return True
        Catch ex As Exception
            Console.WriteLine(ex.ToString())
            Return False
        End Try
    End Function

    Public Function CheckIfControlExists(frm As Form, strName As String) As Boolean
        For Each ctl As Control In frm.Controls
            If ctl.Name = strName Then
                Return True
            End If
        Next
        Return False
    End Function
End Module
