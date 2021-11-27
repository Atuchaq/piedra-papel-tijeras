VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00008080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Piedra Papel Tijeras"
   ClientHeight    =   7980
   ClientLeft      =   6795
   ClientTop       =   3990
   ClientWidth     =   13635
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7980
   ScaleWidth      =   13635
   Begin VB.TextBox Text2 
      BackColor       =   &H00008080&
      BeginProperty Font 
         Name            =   "Russo One"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9960
      TabIndex        =   4
      Top             =   1320
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00008080&
      BeginProperty Font 
         Name            =   "Russo One"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9960
      TabIndex        =   2
      Top             =   480
      Width           =   3375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "RESET"
      BeginProperty Font 
         Name            =   "Russo One"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   10800
      TabIndex        =   1
      Top             =   6840
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Russo One"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   6840
      Width           =   2655
   End
   Begin VB.Image Image5 
      Height          =   615
      Left            =   0
      Top             =   0
      Width           =   615
   End
   Begin VB.Label Label1 
      BackColor       =   &H00008080&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Russo One"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4080
      TabIndex        =   3
      Top             =   3120
      Width           =   5655
   End
   Begin VB.Image Image4 
      Height          =   2295
      Left            =   5400
      Top             =   120
      Width           =   2895
   End
   Begin VB.Image Image3 
      Height          =   3000
      Left            =   8640
      Picture         =   "Form1.frx":0000
      Top             =   3600
      Width           =   3000
   End
   Begin VB.Image Image2 
      Height          =   3000
      Left            =   2400
      Picture         =   "Form1.frx":35E4
      Top             =   3720
      Width           =   3000
   End
   Begin VB.Image Image1 
      Height          =   3000
      Left            =   5760
      Picture         =   "Form1.frx":79BC
      Top             =   3600
      Width           =   3000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Command1_Click()
End
End Sub

Private Sub Command2_Click()
    Player = 0
    Maquina = 0
    Image4.Picture = Image5.Picture
    Text1.Text = "PLAYER: " & Player
    Text2.Text = " MAQUINA: " & Maquina
    Label1.Caption = "REINICIAMOS!!!"
End Sub

Private Sub Form_Load()
    Text1.Text = "PLAYER: " & Player
    Text2.Text = " MAQUINA: " & Maquina
End Sub

Private Sub Image1_Click()
    Call ganador(2)
End Sub

Private Sub Image2_Click()
    Call ganador(1)
End Sub

Private Sub Image3_Click()
    Call ganador(3)
End Sub

