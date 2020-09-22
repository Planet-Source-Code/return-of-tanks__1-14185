Attribute VB_Name = "AI"
Public Sub DoAI(Unit)
    ' This makes stuff move around. The most complrx bit of this is the
    ' way the tanks move towards you. It does this by working out the
    ' angle between it the position it wants to get to (You). Then, if
    ' the angle if greater than 0, it turns to the left, if the angle
    ' is less than 0 it turns right.
    ' It then uses Pythag' to work out how far away the target is. So
    ' long as the target isn't to close, the tank moves in the direction
    ' its facing. Depending on the distance to the target, it raises and
    ' lowers its gun, and fires at random, provided its in range.
    
    If World(Unit).Entity = 1 Then
        If Unit <> 1 Then
            piz = (22 / 7) * 18.2
            TXX = World(World(Unit).Target).Origin.x - World(Unit).Origin.x
            Tyy = World(World(Unit).Target).Origin.z - World(Unit).Origin.z
            PAngle = (-World(Unit).Angle.y / piz)
            Angle = FindAngle(TXX, Tyy, PAngle)
            If Angle > 0 Then World(Unit).Angle.y = World(Unit).Angle.y - 2
            If Angle < 0 Then World(Unit).Angle.y = World(Unit).Angle.y + 2
            x1 = World(World(Unit).Target).Origin.x
            x2 = World(World(Unit).Target).Origin.z
            x3 = World(Unit).Origin.x
            x4 = World(Unit).Origin.z
            dist = Distance(x1, x2, x3, x4)
            If dist > 6 Then World(Unit).GunAngle = World(Unit).GunAngle + 3: If World(Unit).GunAngle > 45 Then World(Unit).GunAngle = 45
            If dist < 5 Then World(Unit).GunAngle = World(Unit).GunAngle - 3: If World(Unit).GunAngle < 5 Then World(Unit).GunAngle = 5
            If dist < 9 And Rnd > 1 - (dist / 100) Then FireMe Unit, "Cannon"
            If dist > 3 Then
                Anga = -World(Unit).Angle.y Mod 360
                x = Sine(Anga)
                y = Cosine(Anga)
                x = x / 5
                y = y / 5
                World(Unit).Origin.x = World(Unit).Origin.x - x
                World(Unit).Origin.z = World(Unit).Origin.z - y
            End If
       End If
       If World(Unit).Health < 0 Then World(Unit).Age = World(Unit).Age + 1
       ' If this tank has been hit, then its age goes up, once per
       ' frame, to create a delay before it explodes..
       ' Three frames after it gets hit...BOOM!
       If World(Unit).Age = 3 Then
           ' Create 20 peices of scrap that make the tank look like
           ' its flying apart...
           For n = 1 To 20
               Add = GetFreeEntity
               World(Add).Used = True
               World(Add).Model = 2 + Int(Rnd * 2)
               World(Add).Entity = 2
               ang = Int(Rnd * 360)
               VAng = Int(Rnd * 90)
               World(Add).Origin.x = World(Unit).Origin.x
               World(Add).Origin.y = World(Unit).Origin.y
               World(Add).Origin.z = World(Unit).Origin.z
               World(Add).ScrapSpeed.x = (Sine(ang) * Cosine(VAng)) * 0.25
               World(Add).ScrapSpeed.z = (Cosine(ang) * Cosine(VAng)) * 0.25
               World(Add).ScrapSpeed.y = (Sine(VAng) * Rnd) * 0.5
               World(Add).ScrapAngle.x = (Rnd * 40) - 20
               World(Add).ScrapAngle.z = (Rnd * 40) - 20
               World(Add).ScrapAngle.y = (Rnd * 40) - 20
               World(Add).Scale.x = 1: World(Add).Scale.y = 1: World(Add).Scale.z = 1
               World(Add).Age = 80 + (Rnd * 30)
           Next n
           ' Add a new tank at a random location, so the fun never ends...
           Add = GetFreeEntity
           World(Add).Model = 1
           World(Add).Entity = 1
           World(Add).Scale.x = 1
           World(Add).Scale.y = 1
           World(Add).Scale.z = 1
           World(Add).Angle.x = 1
           World(Add).Angle.y = 1
           World(Add).Angle.z = 1
           World(Add).Origin.x = (Rnd * 24) - 12
           World(Add).Origin.z = (Rnd * 24) - 12
           World(Add).Origin.y = 0
           World(Add).Target = 1
           World(Add).Health = 0
           World(Add).Used = True
        End If
        ' 5 frames after being hit, the actual tank is removed, so only
        ' peices of flying metal remain... This makes the effect look better
        ' than removing the tank straight away.
        If World(Unit).Age = 5 Then
           World(Unit).Age = -20
           World(Unit).Used = False
        End If
    End If
    
 
    ' This bit does the bullets. It goes throught all the tanks, and
    ' finds wether the bullet is within 1 unit of the tank along each
    ' axis. Its really not very acurate, but I don't really know much
    ' about collision detection. Still, it works here...
    If World(Unit).Entity = 4 Then
        x2 = World(Unit).Origin.x
        y2 = World(Unit).Origin.y
        z2 = World(Unit).Origin.z
        For n = 1 To TotalEntities
            If World(n).Entity = 1 And World(Unit).Owner <> n And World(Unit).Used = True Then
                x1 = World(n).Origin.x
                y1 = World(n).Origin.y
                z1 = World(n).Origin.z
                If x1 - 1 < x2 And x1 + 1 > x2 Then
                    If z1 - 1 < z2 And z1 + 1 > z2 Then
                        If y1 - 1 < y2 And y1 + 1 > y2 Then
                            ' If the bullet has hit a tank, then reduce the
                            ' health on that tank...
                            World(n).Health = World(n).Health - 1
                            World(Unit).Used = False
                            If World(Unit).Owner = 1 Then Form1.Frag = Form1.Frag + 0.5
                        End If
                    End If
                End If
            End If
        Next n
    End If

 
    ' This moves the scrap metal, and the bullets around.
    If World(Unit).Entity = 2 Or World(Unit).Entity = 4 Or World(Unit).Entity = 3 Then
        World(Unit).Age = World(Unit).Age - 1
        If World(Unit).Age < 9 And World(Unit).Entity = 2 Then
            World(Unit).Scale.x = World(Unit).Age / 10
            World(Unit).Scale.y = World(Unit).Age / 10
            World(Unit).Scale.z = World(Unit).Age / 10
        End If
        If World(Unit).Age = 0 Then World(Unit).Used = False
        World(Unit).Origin.x = World(Unit).Origin.x + World(Unit).ScrapSpeed.x
        World(Unit).Origin.y = World(Unit).Origin.y + World(Unit).ScrapSpeed.y
        World(Unit).Origin.z = World(Unit).Origin.z + World(Unit).ScrapSpeed.z
        World(Unit).Angle.x = World(Unit).Angle.x + World(Unit).ScrapAngle.x
        World(Unit).Angle.y = World(Unit).Angle.y + World(Unit).ScrapAngle.y
        World(Unit).Angle.z = World(Unit).Angle.z + World(Unit).ScrapAngle.z
        World(Unit).ScrapSpeed.y = World(Unit).ScrapSpeed.y - 0.025
        If World(Unit).Origin.y <= -0.4 Then
            If World(Unit).Entity = 4 Then
                ' If a bullet his the ground, it makes 2 little peices of
                ' scrap, to show dirt flying up...
                For n = 1 To 2
                    Add = GetFreeEntity
                    World(Add).Used = True
                    World(Add).Model = 2
                    World(Add).Entity = 3
                    ang = Int(Rnd * 360)
                    VAng = Int(Rnd * 90)
                    World(Add).Origin.x = World(Unit).Origin.x
                    World(Add).Origin.y = World(Unit).Origin.y
                    World(Add).Origin.z = World(Unit).Origin.z
                    World(Add).ScrapSpeed.x = (Sine(ang) * Cosine(VAng)) * 0.5
                    World(Add).ScrapSpeed.z = (Cosine(ang) * Cosine(VAng)) * 0.5
                    World(Add).ScrapSpeed.y = (Sine(VAng) * Rnd) * 0.25
                    World(Add).ScrapAngle.x = (Rnd * 40) - 20
                    World(Add).ScrapAngle.z = (Rnd * 40) - 20
                    World(Add).ScrapAngle.y = (Rnd * 40) - 20
                    World(Add).Scale.x = 0.3: World(Add).Scale.y = 0.3: World(Add).Scale.z = 0.3
                    World(Add).Age = 20 + (Rnd * 10)
                Next n
                World(Unit).Used = False
            Else
                ' When scrap hits the ground, it slides a little bit on the
                ' ground, and it lies flat on the ground (So the flat peices
                ' look better)
                World(Unit).Origin.y = -0.4
                World(Unit).ScrapSpeed.x = World(Unit).ScrapSpeed.x * 0.8
                World(Unit).ScrapSpeed.y = World(Unit).ScrapSpeed.y * 0.8
                World(Unit).ScrapSpeed.z = World(Unit).ScrapSpeed.z * 0.8
                World(Unit).Angle.x = 0
                World(Unit).Angle.z = 0
                World(Unit).ScrapAngle.x = 0
                World(Unit).ScrapAngle.y = 0
                World(Unit).ScrapAngle.z = 0
            End If
        End If
    End If
End Sub
