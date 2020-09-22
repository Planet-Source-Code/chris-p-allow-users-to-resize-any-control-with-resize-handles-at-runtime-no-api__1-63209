VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4455
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7530
   LinkTopic       =   "Form1"
   ScaleHeight     =   4455
   ScaleWidth      =   7530
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   180
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "Form1.frx":0000
      Top             =   840
      Width           =   2355
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command 1"
      Height          =   555
      Left            =   180
      TabIndex        =   3
      Top             =   180
      Width           =   2355
   End
   Begin VB.PictureBox test1 
      AutoRedraw      =   -1  'True
      FillColor       =   &H8000000F&
      FillStyle       =   0  'Solid
      Height          =   3015
      Left            =   2760
      ScaleHeight     =   2955
      ScaleWidth      =   4095
      TabIndex        =   1
      Top             =   540
      Width           =   4155
      Begin VB.CheckBox Check1 
         Caption         =   "right - click me to resize me"
         Height          =   255
         Left            =   1080
         TabIndex        =   5
         Top             =   1320
         Width           =   2355
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   1200
         ScaleHeight     =   705
         ScaleWidth      =   1845
         TabIndex        =   2
         Top             =   480
         Width           =   1875
      End
   End
   Begin VB.PictureBox handle 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   90
      Index           =   0
      Left            =   60
      MousePointer    =   8  'Size NW SE
      ScaleHeight     =   6
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   6
      TabIndex        =   0
      Top             =   4260
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "click here to hide the resize handles"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2820
      TabIndex        =   7
      Top             =   120
      Width           =   4155
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000010&
      BackStyle       =   0  'Transparent
      Caption         =   "Right-click any control to resize/move it. Left-click it to remove the gripper handles"
      Height          =   555
      Left            =   2880
      TabIndex        =   6
      Top             =   3720
      Width           =   3915
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then allowresize Check1, Me
If Button = 1 Then handles_hide Me
End Sub

Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then allowresize Command1, Me
If Button = 1 Then handles_hide Me
End Sub

Private Sub handle_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    handle_press X, Y
End Sub

Private Sub handle_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    handle_move Index, Button, Shift, X, Y, Me
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then allowresize Label1, Me
If Button = 1 Then handles_hide Me
End Sub

Private Sub Label2_Click()
handles_hide Me
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then allowresize Picture1, Me
If Button = 1 Then handles_hide Me
End Sub

Private Sub test1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then allowresize test1, Me
If Button = 1 Then handles_hide Me
End Sub

Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then allowresize Text1, Me
If Button = 1 Then handles_hide Me
End Sub
