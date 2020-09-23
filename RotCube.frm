VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form RotCube 
   Caption         =   "RotCube"
   ClientHeight    =   3120
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4620
   Icon            =   "RotCube.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3120
   ScaleWidth      =   4620
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar PB 
      Height          =   1815
      Left            =   4200
      TabIndex        =   3
      Top             =   1080
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   3201
      _Version        =   393216
      Appearance      =   1
      Max             =   360
      Orientation     =   1
      Scrolling       =   1
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Quit"
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   600
      Width           =   1335
   End
   Begin VB.Timer Painter 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3240
      Top             =   1560
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.PictureBox Pic 
      BackColor       =   &H00000000&
      DrawWidth       =   2
      ForeColor       =   &H000000FF&
      Height          =   3135
      Left            =   0
      ScaleHeight     =   3075
      ScaleWidth      =   3075
      TabIndex        =   0
      Top             =   0
      Width           =   3135
   End
   Begin VB.Label Angle 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "360"
      Height          =   255
      Left            =   4200
      TabIndex        =   4
      Top             =   2880
      Width           =   375
   End
End
Attribute VB_Name = "RotCube"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'I tried to make this code fit on most peoples screen but if your screen
'is too small, to read this your gonna have to keep scrolling
'left and right--sorry bout that.


'Hello, thank you for checking out my latest project, It only took
'me an hour or so but the code looks pretty complex. The general setup
'is the same for every line, here it is:
'
'   Pic.Line ((Radius * 2 * Cos((Theta + 0) * PiOver180) + Pic.Width / 2),
'   (Radius * Sin((Theta + 0) * PiOver180) + Pic.Height / 3))-((Radius
'   * 2 * Cos((Theta + 90) * PiOver180) + Pic.Width / 2), (Radius *
'   Sin((Theta + 90) * PiOver180) + Pic.Height / 3))
'
'This looks like pretty serious math but I just used simple math from my
'Algebra 2 class and adapted it to VB. But for those of you that dont fully
'understand this, here is the base formula for a circle in VB
'
'   for Theta = 0 to 360
'   Picture1.PSet (   Radius * Cos(Theta)   ,   Radius * Sin(Theta)   )
'   next
'
'The formula was just adapted to draw a rotateing cube. All (well I think all)
'3D games use this formula when turning the character, just as a note for ya.
'- If you have any questions or comments please put them in my mailbox at
'spinflip@graalmail.com      BTW: consider this code your own, take it,
'steal it, sell it... have fun but if you enjoy/steal it, please vote :)

Dim Theta As Double
Const Radius = 500
Const PiOver180 = (3.1415 / 180)


Private Sub Command1_Click()
If Command1.Caption = "Start" Then
Command1.Caption = "Stop"
Else
Command1.Caption = "Start"
End If

Painter.Enabled = Not Painter.Enabled
Pic.DrawWidth = 2              'Pretty cool to change
End Sub

Private Sub Command2_Click()
Painter.Enabled = Not Painter.Enabled
Unload Me
End Sub

Private Sub Painter_Timer()
Theta = Theta + 1
If Theta > 360 Then Theta = 0 'This really isnt needed since
                              'sin(0) equals sin(360) but
                              'I put it in here for the progressbar
PB.Value = Theta
Angle.Caption = Theta
Pic.Cls
'*************
'Top Square
'*************
Pic.Line ((Radius * 2 * Cos((Theta + 0) * PiOver180) + Pic.Width / 2), (Radius * Sin((Theta + 0) * _
PiOver180) + Pic.Height / 3))-((Radius * 2 * Cos((Theta + 90) * PiOver180) + Pic.Width / 2), (Radius * _
Sin((Theta + 90) * PiOver180) + Pic.Height / 3))

Pic.Line ((Radius * 2 * Cos((Theta + 90) * PiOver180) + Pic.Width / 2), (Radius * Sin((Theta + 90) * _
PiOver180) + Pic.Height / 3))-((Radius * 2 * Cos((Theta + 180) * PiOver180) + Pic.Width / 2), (Radius * _
Sin((Theta + 180) * PiOver180) + Pic.Height / 3))

Pic.Line ((Radius * 2 * Cos((Theta + 180) * PiOver180) + Pic.Width / 2), (Radius * Sin((Theta + 180) * _
PiOver180) + Pic.Height / 3))-((Radius * 2 * Cos((Theta + 270) * PiOver180) + Pic.Width / 2), (Radius * _
Sin((Theta + 270) * PiOver180) + Pic.Height / 3))

Pic.Line ((Radius * 2 * Cos((Theta + 270) * PiOver180) + Pic.Width / 2), (Radius * Sin((Theta + 270) * _
PiOver180) + Pic.Height / 3))-((Radius * 2 * Cos((Theta + 0) * PiOver180) + Pic.Width / 2), (Radius * _
Sin((Theta + 0) * PiOver180) + Pic.Height / 3))

'*************
'Bottom Square
'*************
Pic.Line ((Radius * 2 * Cos((Theta + 0) * PiOver180) + Pic.Width / 2), (Radius * Sin((Theta + 0) * _
PiOver180) + (Pic.Height / 3) * 2))-((Radius * 2 * Cos((Theta + 90) * PiOver180) + Pic.Width / 2), (Radius * _
Sin((Theta + 90) * PiOver180) + (Pic.Height / 3) * 2))

Pic.Line ((Radius * 2 * Cos((Theta + 90) * PiOver180) + Pic.Width / 2), (Radius * Sin((Theta + 90) * _
PiOver180) + (Pic.Height / 3) * 2))-((Radius * 2 * Cos((Theta + 180) * PiOver180) + Pic.Width / 2), (Radius * _
Sin((Theta + 180) * PiOver180) + (Pic.Height / 3) * 2))

Pic.Line ((Radius * 2 * Cos((Theta + 180) * PiOver180) + Pic.Width / 2), (Radius * Sin((Theta + 180) * _
PiOver180) + (Pic.Height / 3) * 2))-((Radius * 2 * Cos((Theta + 270) * PiOver180) + Pic.Width / 2), (Radius * _
Sin((Theta + 270) * PiOver180) + (Pic.Height / 3) * 2))

Pic.Line ((Radius * 2 * Cos((Theta + 270) * PiOver180) + Pic.Width / 2), (Radius * Sin((Theta + 270) * _
PiOver180) + (Pic.Height / 3) * 2))-((Radius * 2 * Cos((Theta + 0) * PiOver180) + Pic.Width / 2), (Radius * _
Sin((Theta + 0) * PiOver180) + (Pic.Height / 3) * 2))

'*************
'Lines InBetween
'*************
Pic.Line ((Radius * 2 * Cos((Theta + 0) * PiOver180) + Pic.Width / 2), (Radius * Sin((Theta + 0) * _
PiOver180) + Pic.Height / 3))-((Radius * 2 * Cos((Theta + 0) * PiOver180) + Pic.Width / 2), (Radius * _
Sin((Theta + 0) * PiOver180) + (Pic.Height / 3) * 2))

Pic.Line ((Radius * 2 * Cos((Theta + 90) * PiOver180) + Pic.Width / 2), (Radius * Sin((Theta + 90) * _
PiOver180) + Pic.Height / 3))-((Radius * 2 * Cos((Theta + 90) * PiOver180) + Pic.Width / 2), (Radius * _
Sin((Theta + 90) * PiOver180) + (Pic.Height / 3) * 2))

Pic.Line ((Radius * 2 * Cos((Theta + 180) * PiOver180) + Pic.Width / 2), (Radius * Sin((Theta + 180) * _
PiOver180) + Pic.Height / 3))-((Radius * 2 * Cos((Theta + 180) * PiOver180) + Pic.Width / 2), (Radius * _
Sin((Theta + 180) * PiOver180) + (Pic.Height / 3) * 2))

Pic.Line ((Radius * 2 * Cos((Theta + 270) * PiOver180) + Pic.Width / 2), (Radius * Sin((Theta + 270) * _
PiOver180) + Pic.Height / 3))-((Radius * 2 * Cos((Theta + 270) * PiOver180) + Pic.Width / 2), (Radius * _
Sin((Theta + 270) * PiOver180) + (Pic.Height / 3) * 2))
End Sub
