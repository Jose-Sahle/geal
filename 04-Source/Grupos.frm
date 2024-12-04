VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGrupos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Grupos"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6330
   Icon            =   "Grupos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5535
   ScaleWidth      =   6330
   Begin VB.CommandButton cmdGravar 
      Caption         =   "&Gravar"
      Height          =   345
      Left            =   3720
      TabIndex        =   3
      Top             =   5160
      Width           =   1245
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   5040
      TabIndex        =   4
      Top             =   5160
      Width           =   1245
   End
   Begin VB.Frame frPermissões 
      Caption         =   "Permissões"
      Height          =   3795
      Left            =   0
      TabIndex        =   8
      Top             =   1290
      Width           =   6315
      Begin MSComctlLib.TreeView tvwPermissões 
         Height          =   3555
         Left            =   60
         TabIndex        =   2
         Top             =   180
         Width           =   6165
         _ExtentX        =   10874
         _ExtentY        =   6271
         _Version        =   393217
         Style           =   7
         Checkboxes      =   -1  'True
         BorderStyle     =   1
         Appearance      =   1
      End
   End
   Begin VB.Frame frGrupos 
      Caption         =   "Grupo"
      Height          =   1275
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   6315
      Begin VB.TextBox txtCódigo 
         Height          =   285
         Left            =   1080
         TabIndex        =   0
         Top             =   300
         Width           =   750
      End
      Begin VB.TextBox txtDescrição 
         Height          =   285
         Left            =   1080
         TabIndex        =   1
         Top             =   750
         Width           =   5000
      End
      Begin VB.Label lblCódigo 
         Caption         =   "Código"
         Height          =   200
         Left            =   150
         TabIndex        =   7
         Top             =   330
         Width           =   855
      End
      Begin VB.Label lblDescrição 
         Caption         =   "Descrição"
         Height          =   195
         Left            =   150
         TabIndex        =   6
         Top             =   780
         Width           =   885
      End
   End
End
Attribute VB_Name = "frmGrupos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TBLGrupos As Table
Dim GruposAberto As Boolean
Dim IndiceGruposAtivo$

Dim lPula As Boolean
Dim lInserir As Boolean
Dim lAlterar As Boolean
Dim mFechar As Boolean

Dim lAllowInsert  As Boolean
Dim lAllowEdit    As Boolean
Dim lAllowDelete  As Boolean
Dim lAllowConsult As Boolean

Dim StatusBarAviso$

Dim DataBaseName(1 To 1) As String
Public Relatório$
Public TotalDatabaseName%

Public lAtualizar As Boolean
Private Sub AtivaCampos()
    BotãoImprimir True
    frGrupos.Enabled = True
    frPermissões.Enabled = True
    BotãoGravar (lInserir Or lAllowEdit)
    cmdCancelar.Enabled = (lInserir Or lAllowEdit)
    cmdGravar.Enabled = (lInserir Or lAllowEdit)
End Sub
Private Function BuscaNode(ByVal Node As Node, ByVal Texto As String) As Boolean
    Dim NextNode As Node, LastNode As Node
    Dim Cont As Byte
    
    Set NextNode = Node.FirstSibling
    Set LastNode = Node.LastSibling
    
    If NextNode.Text = Texto Then
        BuscaNode = True
        Exit Function
    End If
    
    Do While NextNode <> LastNode
        Set NextNode = NextNode.Next
        If NextNode.Text = Texto Then
            BuscaNode = True
            Exit Function
        End If
    Loop
    
    BuscaNode = False
End Function
Private Function Cancelamento()
    Dim Confirmação%, Espaços%, Msg1$, Msg2$
    
    Msg1 = "Você está preste a cancelar a operação que esta realizando !"
    Msg2 = "Tem certeza?"
    Espaços = ((Len(Msg1) - Len(Msg2)) / 2) + 4
    Msg2 = String(Espaços, " ") + Msg2
    Confirmação = MsgBox(Msg1 + vbCr + Msg2, vbYesNo + vbQuestion + vbDefaultButton2, "Confirmação")
    
    If Confirmação = vbNo Then
        Cancelamento = False
        Exit Function
    End If
    
    If lInserir Then
        StatusBarAviso = "Inclusão cancelada"
    End If
    If lAlterar Then
        StatusBarAviso = "Alteração cancelada"
    End If
    BarraDeStatus StatusBarAviso
    
    lInserir = False
    lAlterar = False
    BotãoIncluir True
    
    If TBLGrupos.RecordCount = 0 Then
        NavegaçãoInferior False
        NavegaçãoSuperior False
        BotãoGravar False
        cmdGravar.Enabled = False
        cmdCancelar.Enabled = False
        DesativaCampos
        ZeraCampos
        Cancelamento = True
        Exit Function
    End If
    
    Cancelamento = True
    
    TestaInferior TBLGrupos, True, True, True
    TestaSuperior TBLGrupos, True, True, True
    
    GetRecords
End Function
Private Sub DesativaCampos()
    BotãoImprimir False
    frGrupos.Enabled = False
    frPermissões.Enabled = False
    cmdGravar.Enabled = False
    cmdCancelar.Enabled = False
    BotãoGravar False
End Sub
Public Sub Encontrar()
    Set frmEncontrar.DBBancoDeDados = DBUsuário
    frmEncontrar.NomeDaJanela = "Grupos"
    frmEncontrar.LabelDescription = "Descrição"
    frmEncontrar.Mensagem = "Nenhuma GRUPO foi selecionada!"
    frmEncontrar.BancoDeDados = "USUÁRIO"
    frmEncontrar.Tabela = "GRUPO"
    frmEncontrar.Indice = "2"
    frmEncontrar.CampoChave = "CÓDIGO"
    frmEncontrar.CampoPreencheLista = "DESCRIÇÃO"
    frmEncontrar.Show vbModal
    lPula = True
    txtCódigo = frmEncontrar.Chave
    lPula = False
    PosRecords
End Sub
Private Sub EstruturaDasPermissões()
    'Cria TreeView para permissões
    Dim No As Node, intIndexI As Integer, intIndexII As Integer, intIndexIII As Integer
    
    ' ** Arquivo
    Set No = tvwPermissões.Nodes.Add
        No.Text = "Arquivo"
        No.Key = "Arquivo"
        intIndexI = No.Index
        
        'Agência
        Set No = tvwPermissões.Nodes.Add(intIndexI, tvwChild)
            No.Text = "Agência"
            No.Key = UCase(No.Text)
            intIndexII = No.Index
            'Inclusão
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Incluir"
                No.Tag = "I"
            'Alteração
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Alterar"
                No.Tag = "A"
            'Exclusão
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Excluir"
                No.Tag = "E"
            'Consulta
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Consultar"
                No.Tag = "C"

        'Banco
        Set No = tvwPermissões.Nodes.Add(intIndexI, tvwChild)
            No.Text = "Banco"
            No.Key = UCase(No.Text)
            intIndexII = No.Index
            'Inclusão
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Incluir"
                No.Tag = "I"
            'Alteração
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Alterar"
                No.Tag = "A"
            'Exclusão
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Excluir"
                No.Tag = "E"
            'Consulta
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Consultar"
                No.Tag = "C"

        'Cliente
        Set No = tvwPermissões.Nodes.Add(intIndexI, tvwChild)
            No.Text = "Cliente"
            No.Key = UCase(No.Text)
            intIndexII = No.Index
            'Inclusão
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Incluir"
                No.Tag = "I"
            'Alteração
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Alterar"
                No.Tag = "A"
            'Exclusão
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Excluir"
                No.Tag = "E"
            'Consulta
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Consultar"
                No.Tag = "C"
    
        'Conta Corrente
        Set No = tvwPermissões.Nodes.Add(intIndexI, tvwChild)
            No.Text = "Conta Corrente"
            No.Key = UCase(No.Text)
            intIndexII = No.Index
            'Inclusão
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Incluir"
                No.Tag = "I"
            'Alteração
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Alterar"
                No.Tag = "A"
            'Exclusão
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Excluir"
                No.Tag = "E"
            'Consulta
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Consultar"
                No.Tag = "C"
        
        'Fornecedor
        Set No = tvwPermissões.Nodes.Add(intIndexI, tvwChild)
            No.Text = "Fornecedor"
            No.Key = UCase(No.Text)
            intIndexII = No.Index
            'Inclusão
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Incluir"
                No.Tag = "I"
            'Alteração
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Alterar"
                No.Tag = "A"
            'Exclusão
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Excluir"
                No.Tag = "E"
            'Consulta
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Consultar"
                No.Tag = "C"
            
        'Funcionário
        Set No = tvwPermissões.Nodes.Add(intIndexI, tvwChild)
            No.Text = "Funcionário"
            No.Key = UCase(No.Text)
            intIndexII = No.Index
            'Inclusão
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Incluir"
                No.Tag = "I"
            'Alteração
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Alterar"
                No.Tag = "A"
            'Exclusão
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Excluir"
                No.Tag = "E"
            'Consulta
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Consultar"
                No.Tag = "C"
            
        'Produto
        Set No = tvwPermissões.Nodes.Add(intIndexI, tvwChild)
            No.Text = "Produto"
            No.Key = UCase(No.Text)
            intIndexII = No.Index
            'Inclusão
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Incluir"
                No.Tag = "I"
            'Alteração
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Alterar"
                No.Tag = "A"
            'Exclusão
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Excluir"
                No.Tag = "E"
            'Consulta
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Consultar"
                No.Tag = "C"
            'Quantidade
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Alterar Quantidade"
                No.Tag = "Q"
            'Alterar preço
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Alterar preço"
                No.Tag = "P"
            'Alterar código
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Alterar código"
                No.Tag = "G"
        
        'Despesas
        Set No = tvwPermissões.Nodes.Add(intIndexI, tvwChild)
            No.Text = "Despesas"
            No.Key = UCase(No.Text)
            intIndexII = No.Index
            'Inclusão
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Incluir"
                No.Tag = "I"
            'Alteração
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Alterar"
                No.Tag = "A"
            'Exclusão
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Excluir"
                No.Tag = "E"
            'Consulta
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Consultar"
                No.Tag = "C"
            
    ' ** Movimento
    Set No = tvwPermissões.Nodes.Add
        No.Text = "Movimento"
        No.Key = UCase(No.Text)
        intIndexI = No.Index
        
        'Compra
        Set No = tvwPermissões.Nodes.Add(intIndexI, tvwChild)
            No.Text = "Compra"
            No.Key = UCase(No.Text)
            intIndexII = No.Index
            'Inclusão
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Incluir"
                No.Tag = "I"
            'Alteração
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Alterar"
                No.Tag = "A"
            'Exclusão
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Excluir"
                No.Tag = "E"
            'Consulta
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Consultar"
                No.Tag = "C"
        
        'Devolução/Troca (Compra)
        Set No = tvwPermissões.Nodes.Add(intIndexI, tvwChild)
            No.Text = "Devolução/Troca (Compra)"
            No.Key = UCase(No.Text)
            intIndexII = No.Index
            'Inclusão
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Incluir"
                No.Tag = "I"
            'Alteração
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Alterar"
                No.Tag = "A"
            'Exclusão
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Excluir"
                No.Tag = "E"
            'Consulta
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Consultar"
                No.Tag = "C"
            
        'Venda
        Set No = tvwPermissões.Nodes.Add(intIndexI, tvwChild)
            No.Text = "Venda"
            No.Key = UCase(No.Text)
            intIndexII = No.Index
            'Inclusão
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Incluir"
                No.Tag = "I"
            'Alteração
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Alterar"
                No.Tag = "A"
            'Exclusão
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Excluir"
                No.Tag = "E"
            'Consulta
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Consultar"
                No.Tag = "C"
            'Entrar automaticamente em inclusão
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Entrar automaticamente em inclusão"
                No.Tag = "U"
                
            'Autorizar valor além do desconto máximo
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Permite autorizar desconto"
                No.Tag = "D"
            
        'Devolução/Troca (Venda)
        Set No = tvwPermissões.Nodes.Add(intIndexI, tvwChild)
            No.Text = "Devolução/Troca (Venda)"
            No.Key = UCase(No.Text)
            intIndexII = No.Index
            'Inclusão
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Incluir"
                No.Tag = "I"
            'Alteração
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Alterar"
                No.Tag = "A"
            'Exclusão
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Excluir"
                No.Tag = "E"
            'Consulta
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Consultar"
                No.Tag = "C"
            
        'Movimento Diário
        Set No = tvwPermissões.Nodes.Add(intIndexI, tvwChild)
            No.Text = "Movimento Diário"
            No.Key = UCase(No.Text)
            intIndexII = No.Index
            'Inclusão
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Incluir"
                No.Tag = "I"
            'Alteração
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Alterar"
                No.Tag = "A"
            'Exclusão
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Excluir"
                No.Tag = "E"
            'Consulta
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Consultar"
                No.Tag = "C"
        
        'Conta Corrente (Movimento)
        Set No = tvwPermissões.Nodes.Add(intIndexI, tvwChild)
            No.Text = "Conta Corrente (Movimento)"
            No.Key = UCase(No.Text)
            intIndexII = No.Index
            'Inclusão
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Incluir"
                No.Tag = "I"
            'Alteração
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Alterar"
                No.Tag = "A"
            'Exclusão
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Excluir"
                No.Tag = "E"
            'Consulta
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Consultar"
                No.Tag = "C"
        
        'Caixa
        Set No = tvwPermissões.Nodes.Add(intIndexI, tvwChild)
            No.Text = "Caixa"
            No.Key = UCase(No.Text)
            intIndexII = No.Index
            'Operar
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Operar"
                No.Tag = "O"
            'Entrada
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Entrada"
                No.Tag = "E"
            'Sangria
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Sangria"
                No.Tag = "S"
            'Cancelar último item
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Cancelar último item"
                No.Tag = "U"
            'Cancelar
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Cancelar item"
                No.Tag = "C"
                
        'Caixa Fácil
        Set No = tvwPermissões.Nodes.Add(intIndexI, tvwChild)
            No.Text = "Caixa Fácil"
            No.Key = UCase(No.Text)
            intIndexII = No.Index
            'Operar
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Operar"
                No.Tag = "O"
        
        'PDV
        Set No = tvwPermissões.Nodes.Add(intIndexI, tvwChild)
            No.Text = "PDV"
            No.Key = UCase(No.Text)
            intIndexII = No.Index
            'Acesso
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Acesso"
                No.Tag = "A"
            
        'Abertura do Caixa
        Set No = tvwPermissões.Nodes.Add(intIndexI, tvwChild)
            No.Text = "Abertura do Caixa"
            No.Key = UCase(No.Text)
            intIndexII = No.Index
            'Acesso
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Acesso"
                No.Tag = "A"
                
        'Fechamento do Caixa
        Set No = tvwPermissões.Nodes.Add(intIndexI, tvwChild)
            No.Text = "Fechamento do Caixa"
            No.Key = UCase(No.Text)
            intIndexII = No.Index
            'Acesso
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Acesso"
                No.Tag = "A"
                
    ' ** Parâmetros
    Set No = tvwPermissões.Nodes.Add
        No.Text = "Parâmetros"
        No.Key = UCase(No.Text)
        intIndexI = No.Index
    
        'Departamento
        Set No = tvwPermissões.Nodes.Add(intIndexI, tvwChild)
            No.Text = "Departamento"
            No.Key = UCase(No.Text)
            intIndexII = No.Index
            'Inclusão
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Incluir"
                No.Tag = "I"
            'Alteração
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Alterar"
                No.Tag = "A"
            'Exclusão
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Excluir"
                No.Tag = "E"
            'Consulta
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Consultar"
                No.Tag = "C"
        
        'Seção
        Set No = tvwPermissões.Nodes.Add(intIndexI, tvwChild)
            No.Text = "Seção"
            No.Key = UCase(No.Text)
            intIndexII = No.Index
            'Inclusão
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Incluir"
                No.Tag = "I"
            'Alteração
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Alterar"
                No.Tag = "A"
            'Exclusão
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Excluir"
                No.Tag = "E"
            'Consulta
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Consultar"
                No.Tag = "C"
        
        'Departamento - Seção
        Set No = tvwPermissões.Nodes.Add(intIndexI, tvwChild)
            No.Text = "Departamento - Seção"
            No.Key = UCase(No.Text)
            intIndexII = No.Index
            'Inclusão
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Incluir"
                No.Tag = "I"
            'Alteração
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Alterar"
                No.Tag = "A"
            'Exclusão
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Excluir"
                No.Tag = "E"
            'Consulta
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Consultar"
                No.Tag = "C"
    
        'Tipo de ICM
        Set No = tvwPermissões.Nodes.Add(intIndexI, tvwChild)
            No.Text = "Tipo de ICM"
            No.Key = UCase(No.Text)
            intIndexII = No.Index
            'Inclusão
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Incluir"
                No.Tag = "I"
            'Alteração
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Alterar"
                No.Tag = "A"
            'Exclusão
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Excluir"
                No.Tag = "E"
            'Consulta
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Consultar"
                No.Tag = "C"
            
        'Tipo de Embalagem
        Set No = tvwPermissões.Nodes.Add(intIndexI, tvwChild)
            No.Text = "Tipo de Embalagem"
            No.Key = UCase(No.Text)
            intIndexII = No.Index
            'Inclusão
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Incluir"
                No.Tag = "I"
            'Alteração
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Alterar"
                No.Tag = "A"
            'Exclusão
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Excluir"
                No.Tag = "E"
            'Consulta
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Consultar"
                No.Tag = "C"
            
        'Unidades
        Set No = tvwPermissões.Nodes.Add(intIndexI, tvwChild)
            No.Text = "Unidades"
            No.Key = UCase(No.Text)
            intIndexII = No.Index
            'Inclusão
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Incluir"
                No.Tag = "I"
            'Alteração
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Alterar"
                No.Tag = "A"
            'Exclusão
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Excluir"
                No.Tag = "E"
            'Consulta
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Consultar"
                No.Tag = "C"
            
        'Localidade de Estoque
        Set No = tvwPermissões.Nodes.Add(intIndexI, tvwChild)
            No.Text = "Localidade de Estoque"
            No.Key = UCase(No.Text)
            intIndexII = No.Index
            'Inclusão
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Incluir"
                No.Tag = "I"
            'Alteração
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Alterar"
                No.Tag = "A"
            'Exclusão
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Excluir"
                No.Tag = "E"
            'Consulta
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Consultar"
                No.Tag = "C"
        
        'Plano de Pagamento
        Set No = tvwPermissões.Nodes.Add(intIndexI, tvwChild)
            No.Text = "Plano de Pagamento"
            No.Key = UCase(No.Text)
            intIndexII = No.Index
            'Inclusão
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Incluir"
                No.Tag = "I"
            'Alteração
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Alterar"
                No.Tag = "A"
            'Exclusão
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Excluir"
                No.Tag = "E"
            'Consulta
            Set No = tvwPermissões.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Consultar"
                No.Tag = "C"
End Sub
Public Sub Excluir()
    Dim Confirmação As Integer, Msg1$, Msg2$

    If lAlterar Then
       If Not Cancelamento Then
            Exit Sub
        End If
    End If

    StatusBarAviso = "Exclusão"
    BarraDeStatus StatusBarAviso
    
    Msg1 = "Você está preste a apagar um registro !"
    Msg2 = "Tem certeza?"
    Msg2 = String(((Len(Msg1) - Len(Msg2)) / 2), " ") + Msg2
    Confirmação = MsgBox(Msg1 + vbCr + Msg2, vbYesNo + vbQuestion + vbDefaultButton2, "Confirmação")
    
    If Confirmação = vbNo Then
        Exit Sub
    End If
    
    WS.BeginTrans
    
    TBLGrupos.Delete
    
    If Err <> 0 Then
        GeraMensagemDeErro "Grupos - Excluir", True
        StatusBarAviso = "Falha na exclusão"
        BarraDeStatus StatusBarAviso
        Exit Sub
    End If
    
    WS.CommitTrans
    
    StatusBarAviso = "Exclusão bem sucedida"
    BarraDeStatus StatusBarAviso
    
    If TBLGrupos.RecordCount = 0 Then
        NavegaçãoInferior False
        NavegaçãoSuperior False
        BotãoExcluir False
        BotãoGravar False
        cmdGravar.Enabled = False
        cmdCancelar.Enabled = False
        DesativaCampos
        ZeraCampos
        Exit Sub
    End If
    
    If TBLGrupos.BOF Then
        TBLGrupos.MoveFirst
    ElseIf TBLGrupos.EOF Then
        TBLGrupos.MoveLast
    Else
        TBLGrupos.MovePrevious
        If TBLGrupos.BOF Then
            TBLGrupos.MoveNext
        End If
    End If
    
    GetRecords
    
    TestaInferior TBLGrupos, True, True, True
    TestaSuperior TBLGrupos, True, True, True
End Sub
Public Sub Gravar()
    If lInserir Then
        If SetRecords Then
            PosRecords
            lInserir = False
            StatusBarAviso = "Inclusão bem sucedida"
        Else
            StatusBarAviso = "Falha na inclusão"
        End If
    Else
        If TBLGrupos.RecordCount > 0 And Not TBLGrupos.BOF And Not TBLGrupos.EOF Then
            If SetRecords Then
                PosRecords
                lAlterar = False
                StatusBarAviso = "Alteração bem sucedida"
            Else
                StatusBarAviso = "Falha na alteração"
            End If
        End If
    End If
    
    BarraDeStatus StatusBarAviso
    
    TestaInferior TBLGrupos, True, True, True
    TestaSuperior TBLGrupos, True, True, True
    
    If TBLGrupos.RecordCount = 0 Then
        If Not lInserir And Not lAlterar Then
            BotãoExcluir False
            BotãoGravar False
            cmdGravar.Enabled = False
            cmdCancelar.Enabled = False
        End If
    Else
        BotãoExcluir True
    End If
    
    BotãoIncluir True
    
    If txtCódigo.Enabled Then
        txtCódigo.SetFocus
    End If
End Sub
Public Sub Incluir()
    If lAlterar Then
       If Not Cancelamento Then
            Exit Sub
        End If
    End If
    
    lInserir = True
    
    ZeraCampos
    AtivaCampos
    
    BotãoGravar (lInserir Or lAllowEdit)
    BotãoIncluir False
    cmdGravar.Enabled = (lInserir Or lAllowEdit)
    cmdCancelar.Enabled = (lInserir Or lAllowEdit)
    
    NavegaçãoInferior False
    NavegaçãoSuperior False
    
    StatusBarAviso = "Inclusão"
    BarraDeStatus StatusBarAviso
    
    MontarMarcas
    txtCódigo.SetFocus
End Sub
Private Sub MontarMarcas()
    Dim Cont As Byte
    Dim No As Node
    Dim LastNo As String
    Dim lSair As Boolean
    Dim strResultado As String
    
    Set No = tvwPermissões.Nodes.Item(1).Root
    
    LastNo = No.LastSibling
    lSair = False
    
    Do
        If No.Text = LastNo Then
            lSair = True
        End If
        
        strResultado = VerificaNode(No)
        
        If InStr(strResultado, "1") <> 0 Then
            No.Checked = True
            If InStr(strResultado, "T") <> 0 Then
                No.ForeColor = &H80000008
            Else
                No.ForeColor = &H80000002
            End If
        Else
            No.Checked = False
            No.ForeColor = &H80000008
        End If
        
        Set No = No.Next
        
    Loop While Not lSair
    
End Sub
Public Sub MoveFirst()
    If lAlterar Then
       If Not Cancelamento Then
            Exit Sub
        End If
    End If
    
    StatusBarAviso = "Pronto"
    BarraDeStatus StatusBarAviso
    
    TBLGrupos.MoveFirst
    
    NavegaçãoInferior False
    NavegaçãoSuperior True
    
    GetRecords
End Sub
Public Sub MoveLast()
    If lAlterar Then
       If Not Cancelamento Then
            Exit Sub
        End If
    End If
    
    StatusBarAviso = "Pronto"
    BarraDeStatus StatusBarAviso
    
    TBLGrupos.MoveLast
    
    NavegaçãoInferior True
    NavegaçãoSuperior False
    
    GetRecords
End Sub
Public Sub MoveNext()
    If lAlterar Then
       If Not Cancelamento Then
            Exit Sub
        End If
    End If
    
    TBLGrupos.MoveNext
    If TBLGrupos.EOF Then
        TBLGrupos.MovePrevious
        Exit Sub
    End If
    
    StatusBarAviso = "Pronto"
    BarraDeStatus StatusBarAviso
    
    NavegaçãoInferior True
    TestaSuperior TBLGrupos, True, True, True
    
    GetRecords
End Sub
Public Sub MovePrevious()
    If lAlterar Then
       If Not Cancelamento Then
            Exit Sub
        End If
    End If
    
    TBLGrupos.MovePrevious
    If TBLGrupos.BOF Then
        TBLGrupos.MoveNext
        Exit Sub
    End If
    
    StatusBarAviso = "Pronto"
    BarraDeStatus StatusBarAviso
    
    NavegaçãoSuperior True
    TestaInferior TBLGrupos, True, True, True
    
    GetRecords
End Sub
Public Sub PosRecords()
    If TBLGrupos.RecordCount = 0 Then
        Exit Sub
    End If
    
    TBLGrupos.Seek "=", txtCódigo
    If TBLGrupos.NoMatch Then
        MsgBox "Não consegui encontrar " + txtCódigo, vbExclamation, "Erro"
        TBLGrupos.MoveFirst
        NavegaçãoInferior False
        NavegaçãoInferior True
    Else
        TestaInferior TBLGrupos, True, True, True
        TestaSuperior TBLGrupos, True, True, True
    End If
    GetRecords
End Sub
Private Function GetRecords() As Boolean
    On Error GoTo Erro
    
    Dim Cont As Byte, Cont1 As Byte
    Dim No As Node, NoFilho As Node
    
    lPula = True
    txtCódigo = TBLGrupos("CÓDIGO")
    txtDescrição = TBLGrupos("DESCRIÇÃO")
    
    For Cont = 2 To TBLGrupos.Fields.Count - 1
        Set No = tvwPermissões.Nodes.Item(TBLGrupos(Cont).Name)
        Set NoFilho = No.Child
        For Cont1 = 1 To No.Children
            If InStr(TBLGrupos(Cont), NoFilho.Tag) <> 0 Then
                NoFilho.Checked = True
            Else
                NoFilho.Checked = False
            End If
            Set NoFilho = NoFilho.Next
        Next
    Next
    
    MontarMarcas
    lPula = False
    GetRecords = True
    
    Exit Function
    
Erro:
    GeraMensagemDeErro "GetRecords - " & TBLGrupos(Cont).Name
    GetRecords = False
End Function
Private Function SetRecords()
    On Error GoTo Erro
    
    Dim Msg$
    Dim Confirmação As Integer, Msg1$, Msg2$, Cont As Byte, Cont1 As Byte
    Dim No As Node, NoFilho As Node, LinhaDePermissão As String
    
    WS.BeginTrans 'Inicia uma Transação
    
    If lInserir Then
        TBLGrupos.AddNew
    Else
        TBLGrupos.Edit
    End If
    
    TBLGrupos("CÓDIGO") = txtCódigo
    TBLGrupos("DESCRIÇÃO") = txtDescrição
    
    For Cont = 2 To TBLGrupos.Fields.Count - 1
        Set No = tvwPermissões.Nodes.Item(TBLGrupos(Cont).Name)
        LinhaDePermissão = Empty
        Set NoFilho = No.Child
        For Cont1 = 1 To No.Children
            If NoFilho.Checked Then
                LinhaDePermissão = LinhaDePermissão + NoFilho.Tag
            End If
            Set NoFilho = NoFilho.Next
        Next
        TBLGrupos(Cont) = LinhaDePermissão
    Next
    
    TBLGrupos.Update
        
Erro:
    If Err <> 0 Then
        TBLGrupos.CancelUpdate
        GeraMensagemDeErro "Grupos - SetRecords", True
        SetRecords = False
        Exit Function
    End If

    WS.CommitTrans 'Grava as alterações ou inclusões se não houverem erros
    
    SetRecords = True
End Function
Private Sub VerificaChildren(ByVal Node As Node, ByVal TotalNode, ByVal lValor As Boolean)
    On Error Resume Next
        
    Dim Cont As Byte
    
    For Cont = 1 To TotalNode
        Node.Checked = lValor
        If Node.Children <> 0 Then
            VerificaChildren Node.Child, Node.Children, lValor
        End If
        Set Node = Node.Next
    Next
End Sub
Private Function VerificaNode(ByVal No As Node)
    Dim Cont As Byte, Cont1 As Byte
    Dim NoFilho As Node
    Dim strResultado As String
    Dim lTodosMarcados As Boolean
    Dim lPeloMenosUmMarcado As Boolean
        
    Set NoFilho = No.Child
    
    lPeloMenosUmMarcado = False
    lTodosMarcados = True
    
    For Cont = 1 To No.Children
        If NoFilho.Children = 0 Then
            If NoFilho.Checked And NoFilho.ForeColor = &H80000008 Then
                lPeloMenosUmMarcado = True
            Else
                lTodosMarcados = False
            End If
        Else
            strResultado = VerificaNode(NoFilho)
            If InStr(strResultado, "1") <> 0 Then
                lPeloMenosUmMarcado = True
                NoFilho.Checked = True
                If InStr(strResultado, "T") <> 0 Then
                    NoFilho.ForeColor = &H80000008
                Else
                    lTodosMarcados = False
                    NoFilho.ForeColor = &H80000002
                End If
            Else
                lTodosMarcados = False
                NoFilho.Checked = False
                NoFilho.ForeColor = &H80000008
            End If
        End If
        Set NoFilho = NoFilho.Next
    Next
    
    strResultado = Empty
    If lPeloMenosUmMarcado Then
        strResultado = strResultado & "1"
    End If
    If lTodosMarcados Then
        strResultado = strResultado & "T"
    End If
    
    VerificaNode = strResultado
End Function
Private Sub ZeraCampos()
    Dim Cont As Byte, Cont1 As Byte
    Dim No As Node, NoFilho As Node
    Dim LastNo As String
    Dim lSair As Boolean
    
    txtCódigo = Empty
    txtDescrição = Empty
    
    Set No = tvwPermissões.Nodes.Item(1).Root
    LastNo = No.LastSibling
    lSair = False
    Do
        If No.Text = LastNo Then
            lSair = True
        End If
        No.Checked = False
        No.Expanded = False
        Set NoFilho = No.Child
        For Cont1 = 1 To No.Children
            ZeraFilho NoFilho
            Set NoFilho = NoFilho.Next
        Next
        Set No = No.Next
    Loop While Not lSair
End Sub
Private Sub ZeraFilho(ByVal No As Node)
    Dim Cont As Byte
    Dim NoFilho As Node
    
    No.Checked = False
    No.Expanded = False
    
    If No.Children = 0 Then
        Exit Sub
    End If
    
    Set NoFilho = No.Child
    For Cont = 1 To No.Children
        ZeraFilho NoFilho
        Set NoFilho = NoFilho.Next
    Next
End Sub
Private Sub cmdCancelar_Click()
    Cancelamento
End Sub
Private Sub cmdGravar_Click()
    Gravar
End Sub
Private Sub Form_Activate()
    If mFechar Then
        Unload Me
        Exit Sub
    End If
    If Not GruposAberto Then
        Unload Me
        Exit Sub
    End If
    
    TestaInferior TBLGrupos, True, True, True
    TestaSuperior TBLGrupos, True, True, True
    
    If TBLGrupos.RecordCount = 0 Then
        BotãoGravar False
        cmdGravar.Enabled = False
        cmdCancelar.Enabled = False
        BotãoImprimir False
    Else
        BotãoGravar (lInserir Or lAllowEdit)
        cmdGravar.Enabled = (lInserir Or lAllowEdit)
        cmdCancelar.Enabled = (lInserir Or lAllowEdit)
        BotãoImprimir True
    End If
    
    If lInserir Then
        BotãoGravar (lInserir Or lAllowEdit)
        cmdGravar.Enabled = (lInserir Or lAllowEdit)
        cmdCancelar.Enabled = (lInserir Or lAllowEdit)
        NavegaçãoInferior False
        NavegaçãoSuperior False
        BotãoExcluir False
        BotãoIncluir False
    ElseIf lAlterar Then
        BotãoIncluir True
    Else
        BotãoIncluir True
        StatusBarAviso = "Pronto"
    End If
    
    If lAtualizar Then
        BotãoAtualizar True
    Else
        BotãoAtualizar False
    End If
    
    BarraDeStatus StatusBarAviso
End Sub
Private Sub Form_Deactivate()
    cmdGravar.Enabled = False
    cmdCancelar.Enabled = False
    BotãoImprimir False
End Sub
Private Sub Form_Load()
    On Error GoTo Erro
    
    lAllowInsert = True
    lAllowEdit = True
    lAllowDelete = True
    lAllowConsult = True
    
    EstruturaDasPermissões
    
    ZeraCampos

    lPula = False
    lInserir = False
    lAlterar = False
    
    GruposAberto = AbreTabela(Dicionário, "USUÁRIO", "GRUPO", DBUsuário, TBLGrupos, TBLTabela, dbOpenTable)
    
    If GruposAberto Then
        IndiceGruposAtivo = "GRUPO1"
        TBLGrupos.Index = IndiceGruposAtivo
    Else
        MsgBox "Não consegui abrir a tabela 'GRUPO' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    BotãoIncluir True
 
    If TBLGrupos.RecordCount = 0 Then
        DesativaCampos
        BotãoExcluir False
        BotãoGravar False
    Else
        AtivaCampos
        BotãoExcluir True
        BotãoGravar (lInserir Or lAllowEdit)
        If Not GetRecords Then
            mFechar = True
            Exit Sub
        End If
    End If
    
    NavegaçãoInferior False
        
    If TBLGrupos.RecordCount = 0 Or TBLGrupos.RecordCount = 1 Then
        NavegaçãoSuperior False
    Else
        NavegaçãoInferior True
    End If
                        
    StatusBarAviso = "Pronto"
    Relatório = AddPath(AplicaçãoPath, "REPORT\Grupos.RPT")
    TotalDatabaseName = 1
    DataBaseName(1) = AddPath(AplicaçãoPath, "DATABASE\USUÁRIO.MDB")
    mFechar = False
    
    Exit Sub
    
Erro:
    GeraMensagemDeErro "Grupos - Load"
    mFechar = True
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If lInserir Then
        MsgBox "Você está em uma inclusão!", vbExclamation, Caption
        StatusBarAviso = "Finalize a inclusão"
        BarraDeStatus StatusBarAviso
        Cancel = 1
        SetaFocus Me
        mdiGeal.Mostrar
        Exit Sub
    End If
    If lAlterar Then
        MsgBox "Você está em uma alteração!", vbExclamation, Caption
        StatusBarAviso = "Finalize a alteração"
        BarraDeStatus StatusBarAviso
        Cancel = 1
        SetaFocus Me
        mdiGeal.Mostrar
        Exit Sub
    End If
    
    mdiGeal.StatusBar.Panels("Posição").Visible = False
    ResizeStatusBar
    
    Set frmGrupos = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If GruposAberto Then
        TBLGrupos.Close
    End If
    If Forms.Count = 2 Then
        AllBotões False
    End If
End Sub
Private Sub tvwPermissões_NodeCheck(ByVal Node As MSComctlLib.Node)
    Dim NextNo As Node, Cont As Byte
        
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Alteração"
        BarraDeStatus StatusBarAviso
    End If
    
    'Verifica se há filho
    If Node.Children <> 0 Then
        VerificaChildren Node.Child, Node.Children, Node.Checked
    End If
    
    MontarMarcas
End Sub
Private Sub txtCódigo_Change()
    If Not lPula Then
        FormatMask "9999", txtCódigo
    End If
End Sub
Private Sub txtCódigo_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Alteração"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub txtCódigo_LostFocus()
    LeftBlank txtCódigo
End Sub
Private Sub txtDescrição_Change()
    If Not lPula Then
        FormatMask "@!S30", txtDescrição
    End If
End Sub
Private Sub txtDescrição_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Alteração"
        BarraDeStatus StatusBarAviso
    End If
End Sub

