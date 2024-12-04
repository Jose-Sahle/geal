VERSION 5.00
Begin VB.Form frmEncontraCliente 
   Caption         =   "Cliente"
   ClientHeight    =   3150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6630
   LinkTopic       =   "Form1"
   ScaleHeight     =   3150
   ScaleWidth      =   6630
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   345
      Left            =   1770
      TabIndex        =   4
      Top             =   2700
      Width           =   2835
   End
   Begin VB.Frame frFornecedor 
      Height          =   2565
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      Begin VB.ListBox lstCliente 
         Height          =   1815
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   6345
      End
      Begin VB.TextBox txtNomeRazãoSocial 
         Height          =   285
         Left            =   1230
         TabIndex        =   1
         Top             =   240
         Width           =   5235
      End
      Begin VB.Label lblNomeRazãoSocial 
         Caption         =   "Razão Social"
         Height          =   195
         Left            =   150
         TabIndex        =   2
         Top             =   270
         Width           =   1065
      End
   End
End
Attribute VB_Name = "frmEncontraCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TBLEncontra As Table
Dim EncontraAberto As Boolean
Dim IndiceEncontraAtivo$
Dim CGCCliente()

Dim Pula As Boolean

Public CGC
Public BancoDeDados
Public Tabela
Public Inicio%
Public Fim%
Private Sub cmdOk_Click()
    If lstCliente.ListIndex = -1 Then
        MsgBox "Nenhum Cliente foi escolhido!", , "Aviso"
        Exit Sub
    End If
    Unload Me
End Sub
Private Sub Form_Load()
    Dim Elemento%
    
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    
    EncontraAberto = AbreTabela(Dicionário, BancoDeDados, Tabela, DBCadastro, TBLEncontra, TBLTabela, dbOpenTable)
    
    If EncontraAberto Then
        IndiceEncontraAtivo = Tabela & "2"
        TBLEncontra.Index = IndiceEncontraAtivo
    Else
        MsgBox "Não consegui abrir a tabela " & "'" & BancoDeDados & "'" & "!", vbCritical, "Erro"
        Exit Sub
    End If
    
    TBLEncontra.MoveFirst
    ReDim CGCCliente(0 To (TBLEncontra.RecordCount - 1))
    Elemento = 0
    
    Do While Not TBLEncontra.EOF
        CGCCliente(Elemento) = TBLEncontra("CGC - CPF")
        lstCliente.AddItem TBLEncontra("RAZÃO SOCIAL")
        TBLEncontra.MoveNext
        Elemento = Elemento + 1
    Loop
    
    lstCliente.ListIndex = -1
    txtNomeRazãoSocial = Empty
    
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If lstCliente.ListIndex = -1 Then
        MsgBox "Nenhum Cliente foi escolhido!", , "Aviso"
        Cancel = 1
        Exit Sub
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If lstCliente.ListIndex = -1 Then
        MsgBox "Nenhum Cliente foi escolhido!", , "Aviso"
        Cancel = 1
        Exit Sub
    End If
    CGC = CGCCliente(lstCliente.ListIndex)
End Sub
Private Sub lstCliente_Click()
    If Pula Then
        Exit Sub
    End If
    Pula = True
    txtNomeRazãoSocial = lstCliente.List(lstCliente.ListIndex)
    Pula = False
End Sub
Private Sub lstCliente_DblClick()
    cmdOk_Click
End Sub
Private Sub txtNomeRazãoSocial_Change()
    Dim Cont%, Encontrou As Boolean
    
    If Pula Then
        Exit Sub
    End If
    
    Pula = True
    
    Encontrou = False
    
    For Cont = 0 To lstCliente.ListCount - 1
        If InStr(UCase(lstCliente.List(Cont)), UCase(txtNomeRazãoSocial)) = 1 Then
            Encontrou = True
            lstCliente.ListIndex = Cont
            Exit For
        End If
    Next
    
    If Not Encontrou Then
        lstCliente.ListIndex = -1
    End If
    
    Pula = False
End Sub
