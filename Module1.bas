Attribute VB_Name = "Main"
Type Coord
    x As Single
    y As Single
    z As Single
End Type

Type LightDis
    Orgin As Coord
    Used As Boolean
End Type

Type CamaraDis
    Origin As Coord
    Angle As Coord
End Type

Type WorldDis
    Owner As Byte
    GunAngle As Integer
    Target As Byte
    Model As Byte
    Health As Integer
    Entity As Byte
    Origin As Coord
    Angle As Coord
    Scale As Coord
    Used As Boolean
    Age As Integer
    ScrapSpeed As Coord
    ScrapAngle As Coord
End Type

Type Guns
    Origin As Coord
    Name As String
    Weapon As String
    Active As Boolean
    Angle As Integer
End Type

Type Model
    Weapon(12) As Guns
    TotalPoints As Integer
    TotalEdges As Integer
    TotalGuns As Byte
    Points(200) As Coord
    EdgeUsed(200) As Boolean
    EdgeCount(200) As Byte
    Edge(200, 12) As Integer
End Type

Public Type Weapons
    Name As String
    Type As String
    Model As Byte
    Speed As Byte
    Delay As Integer
    Range As Integer
End Type

Public Const TotalEntities = 400
Public CannonAngle As Byte
Public Light(10) As LightDis
Public ner(14) As Coord
Public Rotated(200) As Coord
Public Weapon(10) As Weapons
Public Model(10) As Model
Public Xof As Integer, Yof As Integer
Public World(TotalEntities) As WorldDis
Public RotateWorld(TotalEntities) As Coord
Public Entity As WorldDis
Public Const PI = 3.14159265358979
Public Sine(-361 To 361) As Double
Public Cosine(-361 To 361) As Double
Public Camara As CamaraDis
Private Const Skale = 0.25

Public Sub RotateTheWorld(viewangle2)
    ' Rotate the positions of the models around the camara
    viewangle2 = viewangle2 Mod 360
    For Cu = 0 To TotalEntities
        If World(Cu).Used = True Then
            m = 1
            x = World(Cu).Origin.x - Camara.Origin.x
            y = -World(Cu).Origin.y + Camara.Origin.y
            z = World(Cu).Origin.z - Camara.Origin.z
            Xrotated = Cosine(viewangle2) * x - Sine(viewangle2) * z
            Yrotated = y
            Zrotated = Sine(viewangle2) * x + Cosine(viewangle2) * z
            RotateWorld(Cu).x = Xrotated
            RotateWorld(Cu).y = Yrotated
            RotateWorld(Cu).z = Zrotated
        End If
    Next Cu
End Sub

Sub RotateAgain(Angle2)
    ' This rotates models around the Y axis. This is needed so that
    ' when you rotate the model, you have to seperatly rotate it to
    ' ajust for the view angle. Take this sub out if you want to see
    ' the problem.
    Dim Cu As Integer
    Dim MTD As Byte
    MTD = Entity.Model
    Angle1 = Angle1 Mod 360
    Angle2 = Angle2 Mod 360
    Angle3 = Angle3 Mod 360
    For Cu = 1 To Model(MTD).TotalPoints
        x = Rotated(Cu).x
        y = Rotated(Cu).y
        z = Rotated(Cu).z
        Xrotated = Cosine(Angle2) * x - Sine(Angle2) * z
        Yrotated = y
        Zrotated = Sine(Angle2) * x + Cosine(Angle2) * z
        Rotated(Cu).x = Xrotated * Skale
        Rotated(Cu).y = Yrotated * Skale
        Rotated(Cu).z = Zrotated * Skale
    Next Cu
End Sub

Sub RotateGunUp(TiltAngle)
    ' This sub rotates the last 16 points of the 'Tank' model. This
    ' 16 points define the gun barrals.
    Dim Cu As Integer
    Dim MTD As Byte
    MTD = Entity.Model
    Angle = TiltAngle
    For Cu = 48 - 15 To 48
        x = Model(1).Points(Cu).x - 3
        y = Model(1).Points(Cu).y - 3
        z = Model(1).Points(Cu).z
        Xrotated = Cosine(Angle) * x - Sine(Angle) * y
        Yrotated = Sine(Angle) * x + Cosine(Angle) * y
        Zrotated = z
        Rotated(Cu).x = Xrotated + 3
        Rotated(Cu).y = Yrotated + 3
        Rotated(Cu).z = Zrotated
    Next Cu
End Sub


Sub Rotate(Angle1, Angle2, Angle3)
    ' This uses 3 sets of maths functions to rotate 3D points around
    ' the center (0,0,0). This function rotates the models around all
    ' three axis
    Dim Cu As Integer
    Dim MTD As Byte
    MTD = Entity.Model
    Angle1 = Angle1 Mod 360
    Angle2 = Angle2 Mod 360
    Angle3 = Angle3 Mod 360
    For Cu = 1 To Model(MTD).TotalPoints
        If Cu < 48 - 15 Then
            x = Model(MTD).Points(Cu).x
            y = Model(MTD).Points(Cu).y
            z = Model(MTD).Points(Cu).z
        Else
            x = Rotated(Cu).x
            y = Rotated(Cu).y
            z = Rotated(Cu).z
        End If
        Xrotated = x
        Yrotated = Cosine(Angle1) * y - Sine(Angle1) * z
        Zrotated = Sine(Angle1) * y + Cosine(Angle1) * z
        x = Xrotated: y = Yrotated: z = Zrotated
        Xrotated = Cosine(Angle2) * x - Sine(Angle2) * z
        Yrotated = y
        Zrotated = Sine(Angle2) * x + Cosine(Angle2) * z
        x = Xrotated: y = Yrotated: z = Zrotated
        Xrotated = Cosine(Angle3) * x - Sine(Angle3) * y
        Yrotated = Sine(Angle3) * x + Cosine(Angle3) * y
        Zrotated = z
        Rotated(Cu).x = Xrotated
        Rotated(Cu).y = Yrotated
        Rotated(Cu).z = Zrotated
    Next Cu
End Sub

Sub Make_LookUp()
    ' Floating point maths is slow, this is faster...
    ' Don't work out Sin and Cos all the time, when you can work it
    ' out once, store it in an array, and then look in that array
    ' for values that you want.
    For i = -361 To 361
        Sine(i) = Sin(i / 180 * PI)
        Cosine(i) = Cos(i / 180 * PI)
    Next
End Sub
