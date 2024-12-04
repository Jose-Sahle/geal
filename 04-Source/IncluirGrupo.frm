VERSION 5.00
Begin VB.Form frmIncluirGrupo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Incluir Grupo"
   ClientHeight    =   3630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6240
   Icon            =   "IncluirGrupo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3630
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdIncluir 
      Caption         =   "&Incluir"
      Height          =   345
      Left            =   3720
      TabIndex        =   2
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   4980
      TabIndex        =   1
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Frame frGrupo 
      Caption         =   "Grupo"
      Height          =   3195
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6225
      Begin VB.ListBox lstGrupo 
         Height          =   2790
         Left            =   60
         TabIndex        =   3
         Top             =   240
         Width           =   6075
      End
   End
End
Attribute VB_Name = "frmIncluirGrupo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TBLGrupos As Table
Dim GruposAberto As Boolean
Dim IndiceGruposAtivo$

Dim ArrayGrupos() As String

Dim lFechar As Boolean

Public GrupoEscolhido As String
Public GrupoCódigo As Long
Private Sub cmdCancelar_Click()
    GrupoEscolhido = Empty
    Unload Me
End Sub
Private Sub cmdIncluir_Click()
    If lstGrupo.ListIndex < 0 Then
        MsgBox "Nenhum grupo foi selecionado !", vbInformation, "Aviso"
        Exit Sub
    End If
    
    GrupoEscolhido = lstGrupo.List(lstGrupo.ListIndex)
    TBLGrupos.Seek "=", GrupoEscolhido
    GrupoCódigo = TBLGrupos("CÓDIGO")
    Unload Me
End Sub
Private Sub Form_Activate()
    If lFechar Then
        Unload Me
    End If
End Sub
Private Sub Form_Load()
    Dim Dimensão As Integer
    Dim lAchou As Boolean
    Dim Cont As Integer, Cont1 As Integer
    
    lFechar = False
    
    GrupoEscolhido = Empty
    
    GruposAberto = AbreTabela(Dicionário, "USUÁRIO", "GRUPO", DBUsuário, TBLGrupos, TBLTabela, dbOpenTable)
    
    If GruposAberto Then
        IndiceGruposAtivo = "GRUPO2"
        TBLGrupos.Index = IndiceGruposAtivo
    Else
        MsgBox "Não consegui abrir a tabela 'GRUPO' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    If TBLGrupos.RecordCount = 0 Then
        MsgBox "A tabela 'GRUPOS' está vazia!", vbInformation, "Aviso"
        lFechar = True
        Exit Sub
    End If
    
    Dimensão = frmUsuários.GetCountGrupo
    If Dimensão = 0 Then
        Dimensão = 1
    End If
    
    ReDim ArrayGrupos(1 To Dimensão)
    
    For Cont = 1 To frmUsuários.GetCountGrupo
        ArrayGrupos(Cont) = frmUsuários.GetGrupo(Cont)
    Next
    
    lstGrupo.Clear
    TBLGrupos.MoveFirst
    
    For Cont = 1 To TBLGrupos.RecordCount
        lAchou = False
        For Cont1 = 1 To Dimensão
            If ArrayGrupos(Cont1) = TBLGrupos("DESCRIÇÃO") Then
                lAchou = True
            End If
        Next
        If Not lAchou Then
            lstGrupo.AddItem TBLGrupos("DESCRIÇÃO")
        End If
        TBLGrupos.MoveNext
    Next
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If GruposAberto Then
        TBLGrupos.Close
    End If
End Sub
