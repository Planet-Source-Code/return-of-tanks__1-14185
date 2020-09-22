Attribute VB_Name = "Fire"
Public Sub FireMe(Unit, WeaponName)
    ' This is my great idea to make firing alot easier. I had it working
    ' in a different program, but it didn't want to work in this program
    ' cos alot of stuff had to be scaled around, and inverted, and loads
    ' of other crap.
    ' Still, I'll get it working next version, if I get a ZBuffer working
    
    
    rg = Rnd * 0.3
    For n = 0 To Model(World(Unit).Model).TotalGuns - 1
        If Model(World(Unit).Model).Weapon(n).Name = WeaponName Then
            WeapNum = Val(Model(World(Unit).Model).Weapon(n).Weapon)
            If Weapon(WeapNum).Type = 1 Then GoSub ProjectileWeapon
        End If
    Next n
Exit Sub

ProjectileWeapon:
    Add = GetFreeEntity
    World(Add).Used = True
    World(Add).Model = 2
    World(Add).Entity = 4
    ang = -World(Unit).Angle.y
    ang = ang Mod 360
    VAng = World(Unit).GunAngle
    If Unit = 1 Then VAng = CannonAngle
    GoSub FindStart
    World(Add).Origin.x = World(Unit).Origin.x + x
    World(Add).Origin.y = World(Unit).Origin.y + y
    World(Add).Origin.z = World(Unit).Origin.z + z
    World(Add).ScrapSpeed.x = -(Sine(ang) * Cosine(VAng)) * 0.5
    World(Add).ScrapSpeed.z = -(Cosine(ang) * Cosine(VAng)) * 0.5
    World(Add).ScrapSpeed.y = Sine(VAng) * (0.7 + rg * 0.3)
    World(Add).ScrapAngle.x = (Rnd * 140) - 70
    World(Add).ScrapAngle.z = (Rnd * 140) - 70
    World(Add).ScrapAngle.y = (Rnd * 140) - 70
    World(Add).Scale.x = 1
    World(Add).Scale.y = 1
    World(Add).Scale.z = 1
    World(Add).Age = 70 + (rg * 30)
    World(Add).Owner = Unit
Return




BeamWeapon:
    GoSub FindStart
    Angle2 = Angle2 + Model(World(Unit).Model).Weapon(n).Angle
    Angle2 = Abs(Angle2 Mod 360)
    For m = 0 To Weapon(WeapNum).Range
        Add = FindFree
        World(Add).Used = True
        World(Add).Entity = 3
        World(Add).Model = Weapon(WeapNum).Model
        World(Add).Age = Int(m / 6) + Weapon(WeapNum).Delay
        zz = Sine(Angle2) * (m * 100)
        XX = Cosine(Angle2) * (m * 100)
        World(Add).Origin.x = World(Unit).Origin.x - x - XX
        World(Add).Origin.y = World(Unit).Origin.y - y
        World(Add).Origin.z = World(Unit).Origin.z - z - zz
        World(Add).Angle.x = World(Unit).Angle.x
        World(Add).Angle.y = Angle2
        World(Add).Angle.z = World(Unit).Angle.z
    Next m
Return

FindStart:
    Angle1 = World(Unit).Angle.x Mod 360
    Angle2 = World(Unit).Angle.y Mod 360
    Angle3 = World(Unit).Angle.z Mod 360
    x = -0.3
    If n = 1 Then x = 0
    y = 0.85
    z = -0.2
    Xrotated = x: Yrotated = Cosine(Angle1) * y - Sine(Angle1) * z: Zrotated = Sine(Angle1) * y + Cosine(Angle1) * z: x = Xrotated: y = Yrotated: z = Zrotated
    Xrotated = Cosine(Angle2) * x - Sine(Angle2) * z:   Yrotated = y:   Zrotated = Sine(Angle2) * x + Cosine(Angle2) * z
    x = Xrotated: y = Yrotated: z = Zrotated: Xrotated = Cosine(Angle3) * x - Sine(Angle3) * y:   Yrotated = Sine(Angle3) * x + Cosine(Angle3) * y:   Zrotated = z
    x = Xrotated: y = Yrotated:   z = Zrotated
Return

End Sub
