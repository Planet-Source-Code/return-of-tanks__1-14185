Attribute VB_Name = "DrawWorld"
Public Sub DrawView(WorldAngle)
    ' This section draws the cannon abgle and the compas in those
    ' little picture boxes at the top of the screen
    Form1.AngPic.Cls
    y = -Sine(CannonAngle) * 40
    x = Cosine(CannonAngle) * 40
    yy = Form1.AngPic.ScaleHeight / 1.7
    XX = 100
    Form1.AngPic.Line (XX + x, yy + y)-(XX, yy)
    Form1.AngPic.Print CannonAngle
    Form1.Compass.Cls
    x = Form1.Compass.ScaleWidth / 2
    y = Form1.Compass.ScaleHeight / 2
    Form1.Compass.Circle (x, y), x * 0.9
    For n = 0 To 359 Step 90
        An = (n + WorldAngle) Mod 360
        XX = Sine(An) * x / 2
        yy = Cosine(An) * y / 2
        Form1.Compass.Line (x + XX, y + yy)-(x + XX + XX, y + yy + yy)
    Next n
    ' Next, it rotates the locations of the objects around the
    ' camara, and stores these values in an seperate array
    RotateTheWorld (WorldAngle)
    ' Then it goes throught the array, and draws the model
    For m = TotalEntities To 1 Step -1
        Cu = m
        If World(Cu).Used = True Then
            Entity.Entity = World(Cu).Entity
            If Entity.Entity <> 0 Then
                Entity.Model = World(Cu).Model
                Entity.Origin.x = RotateWorld(Cu).x
                Entity.Origin.y = RotateWorld(Cu).y
                Entity.Origin.z = RotateWorld(Cu).z
                Entity.Angle.x = World(Cu).Angle.x
                Entity.Angle.y = World(Cu).Angle.y + 90
                Entity.Angle.z = World(Cu).Angle.z
                Entity.Scale.x = World(Cu).Scale.x
                Entity.Scale.y = World(Cu).Scale.y
                Entity.Scale.z = World(Cu).Scale.z
                If m = 1 Then
                        DrawMe (CannonAngle)
                    Else
                        DrawMe (World(m).GunAngle)
                End If
            End If
        End If
    Next m
End Sub
