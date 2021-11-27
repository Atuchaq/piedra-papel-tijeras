Attribute VB_Name = "Module1"
Option Explicit
Public m As Byte
Public num As Byte
Public Maquina As Byte
Public Player As Byte

Public Function papelero() As Byte

    num = Int((3 * Rnd) + 1)
    papelero = num
End Function
Public Sub ganador(num)
    
    m = papelero()
    mostrar (m)
    If Not (num = m) Then
    
    If (num > m) Then
                If (num = 3 And m = 1) Then
                                Form1.Label1.Caption = "GANA LA MAQUINA!!!"
                                Maquina = Maquina + 1
                    Else
                        Form1.Label1.Caption = "GANA EL PLAYER!!!"
                        Player = Player + 1
                End If
        Else
            If (num = 1 And m = 3) Then
                            Form1.Label1.Caption = "GANA EL PLAYER!!!"
                        Player = Player + 1
                    Else
                        Form1.Label1.Caption = "GANA LA MAQUINA!!!"
                                Maquina = Maquina + 1
            End If
    
    End If
        Else
            Form1.Label1.Caption = "EMPATEEE!!!"
    End If
    Form1.Text1.Text = "PLAYER: " & Player
    Form1.Text2.Text = " MAQUINA: " & Maquina
    If (Maquina = 10 Or Player = 10) Then
            Form1.Enabled = False
            Form2.Show
    End If
    End Sub

Public Sub mostrar(num)
    Select Case num
Case 1: Form1.Image4.Picture = Form1.Image2.Picture
Case 2: Form1.Image4.Picture = Form1.Image1.Picture
Case 3: Form1.Image4.Picture = Form1.Image3.Picture
End Select
    
End Sub
