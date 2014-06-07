Imports POEApi.Model
Imports System.Reflection

Public Class FullItem
    Inherits CloneableObject

    Public Property ID As Integer
    Public Property League As String
    Public Property Location As String
    Public Property Rarity As Rarity
    Public Property Name As String
    Public Property GearType As String      ' Ring, Amulet, Shield, etc (Note: Sword, Axe, Mace are split into Sword (1h), Sword (2h), etc)
    Public Property ItemType As String
    Public Property TypeLine As String
    Public Property H As Byte
    Public Property W As Byte
    Public Property X As Integer
    Public Property Y As Integer
    Public Property Sockets As Byte
    Public Property Colours As String
    Public Property Links As Byte
    Public Property Level As Byte = 1
    Public Property LevelGem As Boolean
    Public Property ExplicitPrefixMods As New CloneableList(Of FullMod)
    Public Property ExplicitSuffixMods As New CloneableList(Of FullMod)
    Public Property OtherSolutions As Boolean = False       ' True if there are other solutions to the mod distribution
    Public Property ImplicitMods As New CloneableList(Of FullMod)
    Public Property Quality As Byte = 0
    Public Property Corrupted As Boolean
    Public Property Rank As Single
    Public Property Percentile As Single
    Public Property Price As Price
    Public Property ThreadID As String = ""
End Class

Public Class FullMod
    Inherits CloneableObject

    Public Property FullText As String          ' The full text of the mod before indexing (i.e. "+10% to Fire Resistance" or "Adds 5-8 Cold Damage")
    Public Property Type1 As String             ' The name/type of the mod (i.e. "+% to Fire Resistance" or "Adds Cold Damage")
    Public Property Value1 As Single            ' The value of the mod (i.e. 10 for "+10% to Fire Resistance"), or the lower value for a mod with a range (i.e. 5 for "Adds 5-8 Cold Damage")
    Public Property BaseLowerV1 As Single       ' The possible lowest value for the mod (i.e. 8 for an 8-12 range), or the possible lowest value for the range minimum (i.e. 5 in 5-6 to 8-10)
    Public Property BaseUpperV1 As Single       ' The possible highest value for the mod (i.e. 12 for an 8-12 range), or the possible highest value for the range minimum (i.e. 6 in 5-6 to 8-10)
    Public Property MaxValue1 As Single         ' 0 for single valued mods (which are the majority), or the upper value for a mod with a range (i.e. 8 for "Adds 5-8 Cold Damage")
    Public Property BaseLowerMaxV1 As Single    ' 0 for single valued mods (which are the majority), or the possible lowest value for the range maximum (i.e. 8 in 5-6 to 8-10)
    Public Property BaseUpperMaxV1 As Single    ' 0 for single valued mods (which are the majority), or the possible highest value for the range maximum (i.e. 10 in 5-6 to 8-10)
    Public Property Type2 As String = ""        ' The name/type of the second mod in a combined mod (i.e. "+ to Accuracy Rating" for "% increased Physical Damage / + to Accuracy Rating"
    Public Property Value2 As Single            ' The value of the second mod in a combined mod (i.e. 14 for +14 to Accuracy Rating) (Note: there are no ranged values for combined mods)
    Public Property BaseLowerV2 As Single       ' The possible lowest value for the second mod (i.e. 11 for an 11-18 range)
    Public Property BaseUpperV2 As Single       ' The possible highest value for the second mod (i.e. 18 for an 11-18 range)
    Public Property Weight As Single            ' The weight/rank given to the mod in the weights-*.csv file
    Public Property MiniLvl As Single           ' The item level associated with the mod (note: the actual max item level requirement is 80% of the largest mod level, rounded *down*)
    Public Property ModLevelActual As Integer            ' The actual level of the mod (0-based) ("+10% to Fire Resistance" is the first entry possible, so would be level 0)
    Public Property ModLevelMax As Integer                 ' The maximum level of the mod (0-based)
    Public Property UnknownValues As Boolean = False  ' Some legacy items may have value ranges not listed in the mods datatable, use this to handle them
End Class

Public Class CloneableList(Of T)
    Inherits List(Of T)
    Implements System.ICloneable

    Public Function Clone() As Object Implements System.ICloneable.Clone
        Dim NewList As New CloneableList(Of T)
        If Me.Count > 0 Then
            Dim ICloneType As Type = Me(0).GetType.GetInterface("ICloneable", True)
            If Not (ICloneType Is Nothing) Then
                For Each Value As T In Me
                    NewList.Add(CType(CType(Value, ICloneable).Clone, T))
                Next
            Else
                Dim MethodsList() As MethodInfo = Me(0).GetType.GetMethods
                For Each Value As T In Me
                    NewList.Add(Value)
                Next
            End If
            Return NewList
        Else
            Return NewList
        End If
    End Function
End Class

Public MustInherit Class CloneableObject
    Implements System.ICloneable

    Private Function Clone(ByVal vObj As Object) As Object
        If Not vObj Is Nothing Then
            If vObj.GetType.IsValueType OrElse vObj.GetType Is Type.GetType("System.String") Then
                Return vObj
            Else
                Dim newObject As Object = Activator.CreateInstance(vObj.GetType)
                If Not newObject.GetType.GetInterface("IEnumerable", True) Is Nothing AndAlso Not newObject.GetType.GetInterface("ICloneable", True) Is Nothing Then
                    'This is a cloneable enumeration object so just clone it
                    newObject = CType(vObj, ICloneable).Clone
                    Return newObject
                Else
                    For Each Item As PropertyInfo In newObject.GetType.GetProperties
                        'If a property has the ICloneable interface, then call the interface clone method
                        If Not (Item.PropertyType.GetInterface("ICloneable") Is Nothing) Then
                            If Item.CanWrite Then
                                Dim IClone As ICloneable = CType(Item.GetValue(vObj, Nothing), ICloneable)
                                Item.SetValue(newObject, IClone.Clone, Nothing)
                            End If
                        Else
                            'Otherwise just set the value
                            If Item.CanWrite Then
                                Item.SetValue(newObject, Clone(Item.GetValue(vObj, Nothing)), Nothing)
                            End If
                        End If
                    Next
                    Return newObject
                End If
            End If
        Else
            Return Nothing
        End If
    End Function

    Public Function Clone() As Object Implements System.ICloneable.Clone
        Return Clone(Me)
    End Function
End Class