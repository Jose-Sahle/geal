VERSION 5.00
Begin VB.Form frmConsultaSQL 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta SQL"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5640
   Icon            =   "ConsultaSQL.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4995
   ScaleWidth      =   5640
   Begin VB.CommandButton cmdRemover 
      Caption         =   "&Remover"
      Height          =   345
      Left            =   1050
      TabIndex        =   6
      Top             =   4620
      Width           =   1245
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "E&ditar"
      Height          =   345
      Left            =   1050
      TabIndex        =   5
      Top             =   4260
      Width           =   1245
   End
   Begin VB.CommandButton cmdNovo 
      Caption         =   "&Novo"
      Height          =   345
      Left            =   1050
      TabIndex        =   4
      Top             =   3900
      Width           =   1245
   End
   Begin VB.CommandButton cmdFechar 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   345
      Left            =   4350
      TabIndex        =   3
      Top             =   4620
      Width           =   1245
   End
   Begin VB.CommandButton cmdVisualizar 
      Caption         =   "Visualizar"
      Height          =   1065
      Left            =   30
      Picture         =   "ConsultaSQL.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3900
      Width           =   1005
   End
   Begin VB.Frame frConsultas 
      Caption         =   "Consultas"
      Height          =   3825
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5625
      Begin VB.ListBox lstConsultaSQL 
         Height          =   3375
         Left            =   120
         TabIndex        =   1
         Top             =   270
         Width           =   5325
      End
   End
End
Attribute VB_Name = "frmConsultaSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TBLConsultaSQL As Recordset
Dim IndiceConsultaSQLAtivo As String
Dim ConsultaSQLAberto As Boolean

Dim Item    As Integer
Private Sub Editar()
    frmEscreverConsultaSQL.Nome = lstConsultaSQL.List(lstConsultaSQL.ListIndex)
    frmEscreverConsultaSQL.Tipo = vbAlterar
    frmEscreverConsultaSQL.Show 1
    Item = lstConsultaSQL.ListIndex
    PreencheConsulta
    PosItem
End Sub
Public Sub Excluir()
    Dim Confirmação As Integer, Msg1$, Msg2$

    If lstConsultaSQL.ListIndex < 0 Then
        Exit Sub
    End If
    
    Msg1 = "Você está preste a apagar uma consulta !"
    Msg2 = "Tem certeza?"
    Msg2 = String(((Len(Msg1) - Len(Msg2)) / 2), " ") + Msg2
    Confirmação = MsgBox(Msg1 + vbCr + Msg2, vbYesNo + vbQuestion + vbDefaultButton2, "Confirmação")
    
    If Confirmação = vbNo Then
        Exit Sub
    End If
    
    If Not PosRecords Then
        Exit Sub
    End If
    
    TBLConsultaSQL.Delete
    
    Item = 0
    PreencheConsulta
    PosItem
End Sub
Public Sub Incluir()
    frmEscreverConsultaSQL.Tipo = vbIncluir
    frmEscreverConsultaSQL.Show 1
    PreencheConsulta
    Item = TBLConsultaSQL.RecordCount - 1
    PosItem
End Sub
Private Sub PosItem()
    If lstConsultaSQL.ListCount > 0 Then
        lstConsultaSQL.ListIndex = Item
    Else
        lstConsultaSQL.ListIndex = -1
    End If
End Sub
Private Function PosRecords() As Boolean
    TBLConsultaSQL.Seek "=", lstConsultaSQL.List(lstConsultaSQL.ListIndex)
    
    If Not TBLConsultaSQL.NoMatch Then
        PosRecords = True
    Else
        MsgBox "Não foi possível apagar a consulta " & lstConsultaSQL.List(lstConsultaSQL.ListIndex), vbExclamation, "Aviso"
        PosRecords = False
        Exit Function
    End If
End Function
Private Sub PreencheConsulta()
    lstConsultaSQL.Clear
    
    If TBLConsultaSQL.RecordCount = 0 Then
        Exit Sub
    End If
    
    TBLConsultaSQL.MoveFirst
    
    Do While Not TBLConsultaSQL.EOF
        lstConsultaSQL.AddItem TBLConsultaSQL("DESCRIÇÃO")
        TBLConsultaSQL.MoveNext
    Loop
End Sub
Private Sub cmdEditar_Click()
    If lstConsultaSQL.ListIndex < 0 Then
        Exit Sub
    End If
    Editar
End Sub
Private Sub cmdFechar_Click()
    Unload Me
End Sub
Private Sub cmdNovo_Click()
    Incluir
End Sub
Private Sub cmdRemover_Click()
    If lstConsultaSQL.ListIndex < 0 Then
        Exit Sub
    End If
    Excluir
End Sub
Private Sub cmdVisualizar_Click()
    If lstConsultaSQL.ListIndex < 0 Then
        Exit Sub
    End If
    If Not PosRecords Then
        Exit Sub
    End If
    
    frmVisualizarConsultaSQL.BancoDeDados = DBCadastro.Name
    frmVisualizarConsultaSQL.SQL = TBLConsultaSQL("CONSULTA")
    frmVisualizarConsultaSQL.Intervalo = TBLConsultaSQL("INTERVALO")
    frmVisualizarConsultaSQL.Nome = lstConsultaSQL.List(lstConsultaSQL.ListIndex)
    frmVisualizarConsultaSQL.Show 1
End Sub
Private Sub Form_Activate()
    If Not ConsultaSQLAberto Then
        Unload Me
    End If
    
    AllBotões False
    
    BotãoIncluir True
    
    If TBLConsultaSQL.RecordCount > 0 Then
        BotãoExcluir True
    End If

    NavegaçãoInferior False
    NavegaçãoSuperior False
End Sub
Private Sub Form_Load()

    Item = 0
    lstConsultaSQL.Clear
    
    ConsultaSQLAberto = AbreTabela(Dicionário, "UTILITÁRIO", "CONSULTA", DBUtilitário, TBLConsultaSQL, TBLTabela, dbOpenTable)
    
    If ConsultaSQLAberto Then
        IndiceConsultaSQLAtivo = "CONSULTA1"
        TBLConsultaSQL.Index = IndiceConsultaSQLAtivo
    Else
        MsgBox "Não consegui abrir a tabela 'Consulta' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    PreencheConsulta
    PosItem
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If ConsultaSQLAberto Then
        TBLConsultaSQL.Close
    End If
    If Forms.Count = 2 Then
        AllBotões False
    End If
    Set frmConsultaSQL = Nothing
End Sub
