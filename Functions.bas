Attribute VB_Name = "Functions"
Public Function GetFreeEntity()
    ' Find a place in the World() array, that isn't being used
    ' by anything...
    For n = 2 To TotalEntities
        If World(n).Used = False Then GetFreeEntity = n: Exit Function
    Next n
End Function



