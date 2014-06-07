Public Class JSON_Store
    Public ID As Integer
    Public Name As String
    Public Base_Name As String
    Public Account As String
    Public Quality As Integer
    Public Armour As Integer?
    Public Evasion As Integer?
    Public Energy_Shield As Integer?
    Public Level As Single?
    Public Price As Price
    Public Thread_ID As String
    Public Verified As Boolean
    Public Identified As Boolean
    Public Indexed_At As Date
    Public Thread_Updated_At As Date
    Public League_Name As String
    Public Rarity_Name As String
    Public Socket_Combination As String
    Public Corrupted As Boolean
    Public Item_Type As String
    Public Socket_Count As Integer?
    Public Linked_Socket_Count As Integer?
    Public Sockets As List(Of Socket)
    Public W As Integer
    Public H As Integer
    Public Stats As List(Of Stats)
End Class

Public Class Price
    Public Exa As Single?
    Public Chaos As Single?
    Public GCP As Single?
    Public Alch As Single?
End Class

Public Class Stats
    Public Name As String
    Public Mod_ID As Integer
    Public Value As Single?
    Public Hidden As Boolean
    Public Implicit As Boolean
End Class

Public Class Socket
    Public Attr As String
    Public Group As Integer
End Class
