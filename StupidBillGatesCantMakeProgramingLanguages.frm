VERSION 5.00
Begin VB.Form StupidBillGatesCantMakeProgramingLanguages 
   Caption         =   "Form2"
   ClientHeight    =   1875
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   2430
   LinkTopic       =   "Form2"
   ScaleHeight     =   1875
   ScaleWidth      =   2430
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu MainMenu 
      Caption         =   "MainMenu"
      Begin VB.Menu Menu 
         Caption         =   "Help"
         Index           =   1
      End
      Begin VB.Menu Menu 
         Caption         =   "Frame Rate"
         Index           =   2
      End
      Begin VB.Menu Menu 
         Caption         =   "Exit"
         Index           =   3
      End
   End
End
Attribute VB_Name = "StupidBillGatesCantMakeProgramingLanguages"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Menu_Click(Index As Integer)
    ' VB is wierd. You cant have a form with menus unless you have
    ' a border on that form. BUT you can have a form without a border
    ' and then popup a menu from another form. Clever , eh?
    Select Case Index
        Case 1
            Msxf = Msxf & "The keys are :-" & vbNewLine
            Msxf = Msxf & "'J' , 'L' Turn Left and Right" & vbNewLine
            Msxf = Msxf & "'i' , 'k' Forward and Backward" & vbNewLine
            Msxf = Msxf & "'Y' , 'H' Rasie and lower cannon angle" & vbNewLine
            Msxf = Msxf & "Space = FIRE" & vbNewLine & vbNewLine
            Msxf = Msxf & "Double click main screen to quit" & vbNewLine & vbNewLine
            Msxf = Msxf & "Go out and kill things!!!" & vbNewLine
            MsgBox Msxf, , "Help"
        Case 2
            If Menu(2).Checked = True Then
                Menu(2).Checked = False
                Form1.Label1.Visible = False
            Else
                Menu(2).Checked = True
                Form1.Label1.Visible = True
            End If
    Case 3
            End
    End Select
End Sub
