Public Class JSON_Store
    Inherits CloneableObject

    Public Property ID As Integer
    Public Property Name As String
    Public Property Base_Name As String
    Public Property Account As String
    Public Property Quality As Integer
    Public Property Armour As Integer?
    Public Property Evasion As Integer?
    Public Property Energy_Shield As Integer?
    Public Property Level As Single?
    Public Property Price As Price
    Public Property Thread_ID As String
    Public Property Verified As Boolean
    Public Property Identified As Boolean
    Public Property Indexed_At As Date
    Public Property Thread_Updated_At As Date
    Public Property League_Name As String
    Public Property Rarity_Name As String
    Public Property Socket_Combination As String
    Public Property Corrupted As Boolean
    Public Property Item_Type As String
    Public Property Socket_Count As Integer?
    Public Property Linked_Socket_Count As Integer?
    Public Property Sockets As List(Of Socket)
    Public Property W As Integer
    Public Property H As Integer
    Public Property Stats As List(Of Stats)
End Class

Public Class Price
    Inherits CloneableObject

    Public Property Exa As Single?
    Public Property Chaos As Single?
    Public Property GCP As Single?
    Public Property Alch As Single?
End Class

Public Class Stats
    Inherits CloneableObject

    Public Property Name As String
    Public Property Mod_ID As Integer
    Public Property Value As Single?
    Public Property Hidden As Boolean
    Public Property Implicit As Boolean
End Class

Public Class Socket
    Inherits CloneableObject

    Public Property Attr As String
    Public Property Group As Integer
End Class