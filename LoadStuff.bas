Attribute VB_Name = "LoadStuff"
Public AllGunTypes As Byte

Public Sub LoadWeapons()
    ' Load weapons. This bit doesnt really work yet, either.
    ' In the last Tank game, when you fired, the bullet came out from
    ' the center of the tank, and was generally a crappy way of doing
    ' it. What would be better would be to define the location of guns
    ' in the 3D model, and then trigger the 3D point to fire when you
    ' want it to. Even more, you should be able to define weapons from
    ' an external file. This is that file...
    Open App.Path + "\Weapons.txt" For Input As #1
    Line Input #1, temp
    Line Input #1, temp
    Line Input #1, temp
    AllGunTypes = 0
    Do:
    AllGunTypes = AllGunTypes + 1
    Input #1, Weapon(AllGunTypes).Name
    Input #1, Weapon(AllGunTypes).Type
    Input #1, Weapon(AllGunTypes).Model
    Input #1, Weapon(AllGunTypes).Speed
    Input #1, Weapon(AllGunTypes).Delay
    Input #1, Weapon(AllGunTypes).Range
    Loop While EOF(1) <> True
    Close
    For n = 0 To 10
        For m = 0 To Model(n).TotalGuns
            For a = 0 To AllGunTypes
                If Model(n).Weapon(m).Weapon = Weapon(a).Name Then
                    Model(n).Weapon(m).Weapon = a
                End If
            Next a
        Next m
    Next n
End Sub


Public Sub LoadMap(FileName)
    ' This loads the map. Its not really that important, cos once you've
    ' killed the original tanks, move tanks appear at random.
    Dim Command As String, Value As String
    Open FileName For Input As #1
    Input #1, TotalModels
    For n = 1 To TotalModels: Line Input #1, temp: Next n
    LightCount = 1
    Light(1).Used = True
    Light(1).Orgin.y = 3
    Do:
        Input #1, Command
        Command = LCase(Command)
        If Command <> "" Then
            Line Input #1, Value
            If Command = "worldspace" Then Add = Str(Value): World(Add).Used = True
            If Command = "age" Then World(Add).Age = Str(Value)
            If Command = "model" Then World(Add).Model = Str(Value)
            If Command = "object" Then World(Add).Entity = Str(Value)
            If Command = "origin.x" Then World(Add).Origin.x = Str(Value)
            If Command = "origin.y" Then World(Add).Origin.y = -Str(Value)
            If Command = "origin.z" Then World(Add).Origin.z = Str(Value)
            If Command = "angle.x" Then World(Add).Angle.x = Str(Value)
            If Command = "angle.y" Then World(Add).Angle.y = Str(Value)
            If Command = "angle.z" Then World(Add).Angle.z = Str(Value)
            If Command = "target" Then World(Add).Target = Str(Value)
            World(Add).Scale.x = 1: World(Add).Scale.y = 1: World(Add).Scale.z = 1
        End If
    Loop While EOF(1) <> True
    Close
End Sub


Public Sub LoadModel(MapFileName)
    ' Loads a model from a file. I'm not going to explain it, cos you
    ' probebly will save 3D models differently, or not at all...
    Open MapFileName For Input As #2
    Input #2, TotalModels
    For ModelNumber = 1 To TotalModels
        Input #2, FileName
        Open App.Path & FileName For Input As #1
        Input #1, Model(ModelNumber).TotalPoints
        Model(ModelNumber).TotalPoints = Model(ModelNumber).TotalPoints - 1
        Input #1, Model(ModelNumber).TotalEdges, temp
        Input #1, Model(ModelNumber).TotalGuns
        Model(ModelNumber).TotalPoints = Model(ModelNumber).TotalPoints + 1
        For n = 1 To Model(ModelNumber).TotalPoints
            Input #1, x
            Input #1, z
            Input #1, y
            Model(ModelNumber).Points(n).x = x / 20
            Model(ModelNumber).Points(n).z = z / 20
            Model(ModelNumber).Points(n).y = -y / 18
        Next n
        For n = 1 To Model(ModelNumber).TotalEdges
            Input #1, Model(ModelNumber).EdgeCount(n)
            For m = 1 To Model(ModelNumber).EdgeCount(n)
                Input #1, Model(ModelNumber).Edge(n, m)
            Next m
        Next n
        For n = 0 To Model(ModelNumber).TotalGuns - 1
            Input #1, Model(ModelNumber).Weapon(n).Name
            Input #1, Model(ModelNumber).Weapon(n).Weapon
            Input #1, x
            Model(ModelNumber).Weapon(n).Active = True
            If x = 1 Then Model(ModelNumber).Weapon(n).Active = False
            Input #1, x: Model(ModelNumber).Weapon(n).Origin.x = x / 20
            Input #1, y: Model(ModelNumber).Weapon(n).Origin.y = y / 18
            Input #1, z: Model(ModelNumber).Weapon(n).Origin.z = z / 20
            Input #1, Model(ModelNumber).Weapon(n).Angle
        Next n
        Close #1
    Next ModelNumber
    Close
End Sub
