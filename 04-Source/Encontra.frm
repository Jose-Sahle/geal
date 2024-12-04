VERSION 5.00
Begin VB.Form frmEncontra 
   ClientHeight    =   1575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6480
   Icon            =   "Encontra.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   6480
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   345
      Left            =   1800
      TabIndex        =   1
      Top             =   1110
      Width           =   2895
   End
   Begin VB.Frame frDepartamento 
      Height          =   1035
      Left            =   0
      TabIndex        =   0
      Top             =   -60
      Width           =   6465
      Begin VB.ComboBox cmbEncontra 
         Height          =   315
         Left            =   90
         TabIndex        =   2
         Top             =   390
         Width           =   6285
      End
   End
End
Attribute VB_Name = "frmEncontra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TBLEncontra As Table
Dim EncontraAberto As Boolean
Dim IndiceEncontraAtivo$

Public Código
Public BancoDeDados
Public Tabela
Public Inicio%
Public Fim%
Private Sub cmbEncontra_LostFocus()
    Dim Cont%, Encontrou As Boolean
    
    Encontrou = False
    
    For Cont = 0 To cmbEncontra.ListCount - 1
        If InStr(cmbEncontra.List(Cont), UCase(cmbEncontra.Text)) = 1 Then
            Encontrou = True
            cmbEncontra.ListIndex = Cont
            Exit For
        End If
    Next
    
    If Not Encontrou Then
        cmbEncontra.ListIndex = 0
    End If
End Sub
Private Sub cmdOK_Click()
    Código = Mid(cmbEncontra.Text, Inicio, Fim)
    Unload Me
End Sub
Private Sub Form_Load()

    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    
    EncontraAberto = AbreTabela(Dicionário, BancoDeDados, Tabela, DBCadastro, TBLEncontra, TBLTabela, dbOpenTable)
    
    If EncontraAberto Then
        IndiceEncontraAtivo = Tabela & "1"
        TBLEncontra.Index = IndiceEncontraAtivo
    Else
        MsgBox "Não consegui abrir a tabela " & "'" & BancoDeDados & "'" & "!", vbCritical, "Erro"
        Exit Sub
    End If
    
    TBLEncontra.MoveFirst
    
    Do While Not TBLEncontra.EOF
        cmbEncontra.AddItem TBLEncontra("CÓDIGO") & "-" & TBLEncontra("DESCRIÇÃO")
        TBLEncontra.MoveNext
    Loop
    
    cmbEncontra.ListIndex = 0
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Código = Mid(cmbEncontra.Text, Inicio, Fim)
End Sub
