Attribute VB_Name = "DrawModel"
Private Const Zeye = 800
Sub DrawMe(TiltGun)
    ' To draw the model, it first rotates the barral of the gun up,
    ' then the whole tank around its axis, and then the whole tank
    ' around the camara angle.
    ' To rotate, it takes the 3D co-ordanates in the 3D model, rotates
    ' these points and stored them in a seperate array named 'Rotated'
    ' It then uses this set of 3D co-ordanates to draw the model
    Dim nn As Byte
    Dim MTD As Byte
    MTD = Entity.Model
    Call RotateGunUp(TiltGun)
    Call Rotate(Entity.Angle.x, Entity.Angle.y - Counter, Entity.Angle.z)
    Call RotateAgain(Camara.Angle.y)
    DrawShadows
    For eachface = Model(MTD).TotalEdges To 1 Step -1
        ThisFace = eachface
        Edges = Model(MTD).EdgeCount(eachface)
        For m = 1 To Edges
           x = (Rotated(Model(MTD).Edge(ThisFace, m) + 1).x * Entity.Scale.x) - Entity.Origin.x
           y = (Rotated(Model(MTD).Edge(ThisFace, m) + 1).y * Entity.Scale.y) - Entity.Origin.y
           z = (Rotated(Model(MTD).Edge(ThisFace, m) + 1).z * Entity.Scale.z) - Entity.Origin.z
           ner(m).x = x
           ner(m).y = y
           ner(m).z = z
        Next m
        r = eachface / Model(MTD).TotalEdges
        Form1.CreateTriangle Edges, r, r, r
    Next eachface
End Sub

Private Sub DrawShadows()
    ' The first tank game had shadows projected from a positionable
    ' light source. However, the same code transfered to this program
    ' came up with some wierd problems, most likey to do with
    ' the scale - In the first tank game, the tanks were 200 units
    ' long. Here, they are about 1 unit long. Therefore, I got rid of
    ' the movable light bit. I'll but that in later. Now, the shadow
    ' is directly below the model.
    MTD = Entity.Model
    For Face = 1 To Model(MTD).TotalEdges
        For Edge = 1 To Model(MTD).EdgeCount(Face)
            temp = -Rotated(Model(MTD).Edge(Face, Edge) + 1).y
            lX = Light(1).Orgin.x + Entity.Origin.x
            lY = Light(1).Orgin.y
            Lz = Light(1).Orgin.z + Entity.Origin.z
            x = (Rotated(Model(MTD).Edge(Face, Edge) + 1).x * Entity.Scale.x) - Entity.Origin.x
            y = -Camara.Origin.y
            z = (Rotated(Model(MTD).Edge(Face, Edge) + 1).z * Entity.Scale.x) - Entity.Origin.z
            ner(Edge).x = x
            ner(Edge).y = y
            ner(Edge).z = z
        Next Edge
        Form1.CreateTriangle Edge - 1, 0, 0, 0
    Next Face
End Sub

