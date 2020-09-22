VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5145
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6150
   ForeColor       =   &H000000FF&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   2  'Cross
   ScaleHeight     =   5145
   ScaleWidth      =   6150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.VScrollBar CamDis 
      Height          =   1455
      Left            =   3840
      Max             =   -1
      Min             =   -60
      TabIndex        =   9
      Top             =   0
      Value           =   -1
      Width           =   375
   End
   Begin VB.PictureBox Picture1 
      Height          =   1455
      Left            =   4200
      ScaleHeight     =   1395
      ScaleWidth      =   1275
      TabIndex        =   6
      Top             =   0
      Width           =   1335
      Begin VB.Label Label3 
         Caption         =   "Score"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   1335
      End
      Begin VB.Label Frag 
         Alignment       =   2  'Center
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   0
         TabIndex        =   7
         Top             =   600
         Width           =   1215
      End
   End
   Begin MSComctlLib.ProgressBar FirePause 
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   1080
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Max             =   10
      Scrolling       =   1
   End
   Begin VB.PictureBox AngPic 
      AutoRedraw      =   -1  'True
      DrawWidth       =   4
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   0
      ScaleHeight     =   69
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   157
      TabIndex        =   1
      Top             =   0
      Width           =   2415
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Cannon Angle"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1560
      Top             =   1560
   End
   Begin VB.PictureBox Compass 
      Height          =   1455
      Left            =   2400
      ScaleHeight     =   1395
      ScaleWidth      =   1395
      TabIndex        =   4
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label lblDiedMessage 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "You Be Dead"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   0
      TabIndex        =   5
      Top             =   3600
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   0
      TabIndex        =   0
      Top             =   1440
      Visible         =   0   'False
      Width           =   1515
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DX As New DirectX7
Dim DDRAW As DirectDraw7
Dim SurfDesc As DDSURFACEDESC2
Dim Primary As DirectDrawSurface7
Dim Backbuffer As DirectDrawSurface7
Dim Clipper As DirectDrawClipper
Dim DestRect As RECT
Dim SrcRect As RECT
Dim D3D As Direct3D7
Dim D3Ddevice As Direct3DDevice7
Dim Viewport(0) As D3DRECT
Dim VPdesc As D3DVIEWPORT7
Dim Vertex(0 To 5) As D3DVERTEX
Dim Material As D3DMATERIAL7
Dim matWorld As D3DMATRIX
Dim matView  As D3DMATRIX
Dim matProj As D3DMATRIX
Dim matSpin As D3DMATRIX
Dim Frames As Integer
Dim CloseProg As Boolean

Function DirectDrawInit() As Long
    ' Sets up DirectX stuff. I didn't actually think of all this stuff myself. This bit is from PSC
    Set DDRAW = DX.DirectDrawCreate("")
    DDRAW.SetCooperativeLevel hWnd, DDSCL_NORMAL
    SurfDesc.lFlags = DDSD_CAPS
    SurfDesc.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
    Set Primary = DDRAW.CreateSurface(SurfDesc)
    SurfDesc.lFlags = DDSD_HEIGHT Or DDSD_WIDTH Or DDSD_CAPS
    SurfDesc.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN Or DDSCAPS_3DDEVICE
    DX.GetWindowRect hWnd, DestRect
    SurfDesc.lWidth = DestRect.Right - DestRect.Left
    SurfDesc.lHeight = DestRect.Bottom - DestRect.Top
    Set Backbuffer = DDRAW.CreateSurface(SurfDesc)
    With SrcRect
            .Left = 0: .Top = 0
            .Bottom = SurfDesc.lHeight
            .Right = SurfDesc.lWidth
    End With
    Set Clipper = DDRAW.CreateClipper(0)
    Clipper.SetHWnd hWnd
    Primary.SetClipper Clipper
    DirectDrawInit = Err.Number
End Function

Function Direct3DInit() As Long
    ' Sets up DirectX stuff. I didn't actually think of all this stuff myself. This bit is from PSC
    Set D3D = DDRAW.GetDirect3D
    Set D3Ddevice = D3D.CreateDevice("IID_IDirect3DRGBDevice", Backbuffer)
    VPdesc.lWidth = DestRect.Right - DestRect.Left
    VPdesc.lHeight = DestRect.Bottom - DestRect.Top
    VPdesc.minz = 0
    VPdesc.maxz = 1
    D3Ddevice.SetViewport VPdesc
    With Viewport(0)
        .x1 = 0: .y1 = 0
        .x2 = VPdesc.lWidth
        .y2 = VPdesc.lHeight
    End With
    D3Ddevice.SetRenderState D3DRENDERSTATE_AMBIENT, DX.CreateColorRGBA(1, 1, 1, 1)
    D3Ddevice.SetRenderState D3DRENDERSTATE_CULLMODE, D3DCULL_CCW
    Material.Ambient.r = 0.6
    Material.Ambient.g = 0.2
    Material.Ambient.b = 0
    D3Ddevice.SetMaterial Material
    DX.IdentityMatrix matWorld
    D3Ddevice.SetTransform D3DTRANSFORMSTATE_WORLD, matWorld
    DX.IdentityMatrix matView
    DX.ViewMatrix matView, MakeVector(0, 0, -0.1), MakeVector(0, 0, 0), MakeVector(0, 1, 0), 0
    D3Ddevice.SetTransform D3DTRANSFORMSTATE_VIEW, matView
    DX.IdentityMatrix matProj
    DX.ProjectionMatrix matProj, 1, 90, 3.14 / 2.1
    D3Ddevice.SetTransform D3DTRANSFORMSTATE_PROJECTION, matProj
    Dim x As D3DCLIPSTATUS
    Direct3DInit = Err.Number
End Function

Function AttachZbuffer(Surf As DirectDrawSurface7, BPP As Byte, Guid As String, Width As Integer, Height As Integer, Optional UseVideoMem As Boolean = False) As Boolean
    ' This is the bit thats really pissing my off.. Setting up a ZBuffer. This game would probebly
    ' look miles better if the ZBuffer would work. I hanv't really got a clue how to do it.
    ' If you know how, DO IT FOR ME DO IT FOR ME DO IT FOR ME DO IT FOR ME DO IT FOR ME
    'On Error Resume Next
    Dim ddpfZBuffer As DDPIXELFORMAT
    Dim d3dEnumPFs As Direct3DEnumPixelFormats
    Dim i As Long
    Dim SurfDesc As DDSURFACEDESC2
    Set d3dEnumPFs = D3D.GetEnumZBufferFormats(Guid)
    For i = 1 To d3dEnumPFs.GetCount()
        d3dEnumPFs.GetItem i, ddpfZBuffer
        If ddpfZBuffer.lFlags = DDPF_ZBUFFER Then Exit For
    Next i
    If Err.Number <> DD_OK Then Exit Function
    
    SurfDesc.lFlags = DDSD_CAPS Or DDSD_WIDTH Or DDSD_HEIGHT Or DDSD_PIXELFORMAT
    If UseVideoMem = False Then
        SurfDesc.ddsCaps.lCaps = DDSCAPS_ZBUFFER Or DDSCAPS_SYSTEMMEMORY
    Else
        SurfDesc.ddsCaps.lCaps = DDSCAPS_ZBUFFER Or DDSCAPS_VIDEOMEMORY
    End If
    SurfDesc.lWidth = Width
    SurfDesc.lHeight = Height
    SurfDesc.ddpfPixelFormat = ddpfZBuffer
    Set Zbuff = DDRAW.CreateSurface(SurfDesc)
    If Err.Number <> DD_OK Then Exit Function
    Surf.AddAttachedSurface Zbuff
    If Err.Number = DD_OK Then AttachZbuffer = True
End Function

Function MakeVector(x As Single, y As Single, z As Single) As D3DVECTOR
    With MakeVector
        .x = x
        .y = y
        .z = z
    End With
End Function

Private Sub AngPic_Click()
    PopupMenu StupidBillGatesCantMakeProgramingLanguages.MainMenu
End Sub

Private Sub CamDis_Scroll()
    ' Move the distance between the camara and 0
    DX.ViewMatrix matView, MakeVector(0, 0, CamDis / 10), MakeVector(0, 0, 0), MakeVector(0, 1, 0), 0
    D3Ddevice.SetTransform D3DTRANSFORMSTATE_VIEW, matView
End Sub

Private Sub Compass_Click()
    PopupMenu StupidBillGatesCantMakeProgramingLanguages.MainMenu
End Sub

Private Sub Form_DblClick()
    End
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    ' Move tank when you press stuff. Again, if I knew somthing about DirectX, I could use
    ' DirectInput or somthing, so key presses are smother, and you can press loads of keys at
    ' the same time
    If World(1).Used = True Then
        i = UCase(Chr(KeyCode))
        VAng = Camara.Angle.y Mod 360
        rDis = Sine(VAng) / 2
        vDis = Cosine(VAng) / 2
        If i = "I" Then
            Camara.Origin.z = Camara.Origin.z - vDis
            Camara.Origin.x = Camara.Origin.x - rDis
        End If
        If i = "K" Then
            Camara.Origin.z = Camara.Origin.z + vDis
            Camara.Origin.x = Camara.Origin.x + rDis
        End If
        If i = "J" Then Camara.Angle.y = Camara.Angle.y - 5
        If i = "L" Then Camara.Angle.y = Camara.Angle.y + 5
        If i = " " And FirePause = 10 Then FireMe 1, "Cannon": FirePause = 0
        If i = "Y" Then CannonAngle = CannonAngle + 5
        If i = "H" Then CannonAngle = CannonAngle - 5
        If CannonAngle > 90 Then CannonAngle = 90
        If CannonAngle < 5 Then CannonAngle = 5
    Else
        lblDiedMessage.Top = Me.ScaleHeight / 2
        lblDiedMessage.Width = Me.ScaleWidth
        lblDiedMessage.Visible = True
    End If
End Sub

Private Sub Form_Load()
   CannonAngle = 5
   Me.Show
   If DirectDrawInit() <> DD_OK Then Unload Me
   If Direct3DInit() <> DD_OK Then Unload Me
   LoadModel App.Path + "\Level1.txt"
   LoadMap App.Path + "\Level1.txt"
   LoadWeapons
   Make_LookUp
   Me.Show
   MainLoop
End Sub

Public Sub CreateTriangle(Edge, r, g, b)
    ' This dosn't just draw triangles, it draws any shape you want. However, it does draw shapes
    ' by making lots of triangles. If you ask it to draw a cube, it will draw 2 triangles instead.
    ' A hexogon is made from 6 triiangles and so on...
    Material.Ambient.r = r
    Material.Ambient.g = g
    Material.Ambient.b = b
    D3Ddevice.SetMaterial Material
    For n = 2 To Edge - 1
        x = ner(1).x: y = ner(1).y - 0.5: z = ner(1).z
        DX.CreateD3DVertex x, y, z, 0, 0, 0, 0, 0, Vertex(0)
        x = ner(n).x: y = ner(n).y - 0.5: z = ner(n).z
        DX.CreateD3DVertex x, y, z, 0, 0, 0, 0, 0, Vertex(1)
        x = ner(n + 1).x: y = ner(n + 1).y - 0.5: z = ner(n + 1).z
        DX.CreateD3DVertex x, y, z, 0, 0, 0, 0, 0, Vertex(2)
        D3Ddevice.DrawPrimitive D3DPT_TRIANGLEFAN, D3DFVF_VERTEX, Vertex(0), 3, D3DDP_DEFAULT
    Next n
End Sub

Private Sub MainLoop()
    ' You really shouldn't use a timer for animation. Appart from anything, it has a maximum
    ' speed of about 26 calls per second (or somthing like that) so that sevirly limits your
    ' frame rate. Instead, use a Do Loop, with a DoEvents it the middle. This will cause a loop
    ' that dosn't block up the system, but runs as fast as possible..
    Dim n As Integer
    Camara.Origin.y = 1
    Do
        DoEvents
        If FirePause <> FirePause.Max Then FirePause = FirePause + 1
        World(1).Origin.x = Camara.Origin.x
        World(1).Origin.z = Camara.Origin.z
        World(1).Angle.y = -Camara.Angle.y
        For n = 1 To TotalEntities
            If World(n).Used = True Then DoAI (n)
        Next n
        Frames = Frames + 1
        D3Ddevice.Clear 1, Viewport(), D3DCLEAR_TARGET, RGB(8, 8, 8), 0, 0
        D3Ddevice.BeginScene
        DrawView (WorldAngle + Camara.Angle.y)
        D3Ddevice.EndScene
        DX.RotateYMatrix matSpin, Counter / 360
        D3Ddevice.SetTransform D3DTRANSFORMSTATE_WORLD, matSpin
        DX.GetWindowRect hWnd, DestRect
        Primary.Blt DestRect, Backbuffer, SrcRect, DDBLT_WAIT
        DoEvents
        Label1.Refresh
    Loop Until CloseProg = True ' This ends the loop when the form is unloaded, else you
                                ' have to do Ctrl Alt Del to close the program
    End
End Sub

' ###########  Boring stuff, Don't bother looking at them  ###########

Private Sub Form_Unload(Cancel As Integer)
    CloseProg = True
End Sub

Private Sub Label1_Click()
    If Label1.Visible = True Then Label1.Visible = False Else Label1.Visible = True
End Sub

Private Sub Label2_Click()
    PopupMenu StupidBillGatesCantMakeProgramingLanguages.MainMenu
End Sub

Private Sub Timer1_Timer()
    Label1 = Frames: Frames = 0
End Sub

Private Sub CamDis_Change()
    DX.ViewMatrix matView, MakeVector(0, 0, CamDis / 10), MakeVector(0, 0, 0), MakeVector(0, 1, 0), 0
    D3Ddevice.SetTransform D3DTRANSFORMSTATE_VIEW, matView
End Sub
    
    
    
    
    
    
