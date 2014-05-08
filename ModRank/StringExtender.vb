Imports System.Runtime.CompilerServices

Namespace Extenders.Strings

    Public Module StringExtender

        <Extension()> _
        Public Function CompareMultiple(ByVal this As String, compareType As StringComparison, ParamArray compareValues As String()) As Boolean

            Dim s As String

            For Each s In compareValues
                If (this.Equals(s, compareType)) Then
                    Return True
                End If
            Next s

            Return False
        End Function

    End Module

End Namespace
