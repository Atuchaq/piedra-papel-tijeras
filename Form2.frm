VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   0  'None
   Caption         =   "Fin"
   ClientHeight    =   5310
   ClientLeft      =   9900
   ClientTop       =   5400
   ClientWidth     =   6270
   LinkTopic       =   "Form2"
   Picture         =   "Form2.frx":0000
   ScaleHeight     =   5310
   ScaleWidth      =   6270
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "aceptar"
      BeginProperty Font 
         Name            =   "Russo One"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1920
      TabIndex        =   0
      Top             =   4560
      Width           =   2295
   End
   Begin VB.Label Label1 
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
      Height          =   855
      Left            =   1440
      TabIndex        =   1
      Top             =   240
      Width           =   4815
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Form1.Enabled = True
    Player = 0
    Maquina = 0
    Form1.Image4.Picture = Form1.Image5.Picture
    Form1.Label1.Caption = ""
    Form1.Text1.Text = "PLAYER: " & Player
    Form1.Text2.Text = " MAQUINA: " & Maquina
    Form1.Show
    Unload Form2
End Sub


Private Sub Form_Load()
    If (Maquina = 10) Then
                Label1.Caption = "GANO LA MAQUINA!!!"
            Else
                Label1.Caption = "GANO EL PLAYER!!!"
    End If
End Sub
