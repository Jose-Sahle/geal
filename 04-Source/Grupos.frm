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
   Begin VB.Frame frPermiss�es 
      Caption         =   "Permiss�es"
      Height          =   3795
      Left            =   0
      TabIndex        =   8
      Top             =   1290
      Width           =   6315
      Begin MSComctlLib.TreeView tvwPermiss�es 
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
      Begin VB.TextBox txtC�digo 
         Height          =   285
         Left            =   1080
         TabIndex        =   0
         Top             =   300
         Width           =   750
      End
      Begin VB.TextBox txtDescri��o 
         Height          =   285
         Left            =   1080
         TabIndex        =   1
         Top             =   750
         Width           =   5000
      End
      Begin VB.Label lblC�digo 
         Caption         =   "C�digo"
         Height          =   200
         Left            =   150
         TabIndex        =   7
         Top             =   330
         Width           =   855
      End
      Begin VB.Label lblDescri��o 
         Caption         =   "Descri��o"
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
Public Relat�rio$
Public TotalDatabaseName%

Public lAtualizar As Boolean
Private Sub AtivaCampos()
    Bot�oImprimir True
    frGrupos.Enabled = True
    frPermiss�es.Enabled = True
    Bot�oGravar (lInserir Or lAllowEdit)
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
    Dim Confirma��o%, Espa�os%, Msg1$, Msg2$
    
    Msg1 = "Voc� est� preste a cancelar a opera��o que esta realizando !"
    Msg2 = "Tem certeza?"
    Espa�os = ((Len(Msg1) - Len(Msg2)) / 2) + 4
    Msg2 = String(Espa�os, " ") + Msg2
    Confirma��o = MsgBox(Msg1 + vbCr + Msg2, vbYesNo + vbQuestion + vbDefaultButton2, "Confirma��o")
    
    If Confirma��o = vbNo Then
        Cancelamento = False
        Exit Function
    End If
    
    If lInserir Then
        StatusBarAviso = "Inclus�o cancelada"
    End If
    If lAlterar Then
        StatusBarAviso = "Altera��o cancelada"
    End If
    BarraDeStatus StatusBarAviso
    
    lInserir = False
    lAlterar = False
    Bot�oIncluir True
    
    If TBLGrupos.RecordCount = 0 Then
        Navega��oInferior False
        Navega��oSuperior False
        Bot�oGravar False
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
    Bot�oImprimir False
    frGrupos.Enabled = False
    frPermiss�es.Enabled = False
    cmdGravar.Enabled = False
    cmdCancelar.Enabled = False
    Bot�oGravar False
End Sub
Public Sub Encontrar()
    Set frmEncontrar.DBBancoDeDados = DBUsu�rio
    frmEncontrar.NomeDaJanela = "Grupos"
    frmEncontrar.LabelDescription = "Descri��o"
    frmEncontrar.Mensagem = "Nenhuma GRUPO foi selecionada!"
    frmEncontrar.BancoDeDados = "USU�RIO"
    frmEncontrar.Tabela = "GRUPO"
    frmEncontrar.Indice = "2"
    frmEncontrar.CampoChave = "C�DIGO"
    frmEncontrar.CampoPreencheLista = "DESCRI��O"
    frmEncontrar.Show vbModal
    lPula = True
    txtC�digo = frmEncontrar.Chave
    lPula = False
    PosRecords
End Sub
Private Sub EstruturaDasPermiss�es()
    'Cria TreeView para permiss�es
    Dim No As Node, intIndexI As Integer, intIndexII As Integer, intIndexIII As Integer
    
    ' ** Arquivo
    Set No = tvwPermiss�es.Nodes.Add
        No.Text = "Arquivo"
        No.Key = "Arquivo"
        intIndexI = No.Index
        
        'Ag�ncia
        Set No = tvwPermiss�es.Nodes.Add(intIndexI, tvwChild)
            No.Text = "Ag�ncia"
            No.Key = UCase(No.Text)
            intIndexII = No.Index
            'Inclus�o
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Incluir"
                No.Tag = "I"
            'Altera��o
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Alterar"
                No.Tag = "A"
            'Exclus�o
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Excluir"
                No.Tag = "E"
            'Consulta
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Consultar"
                No.Tag = "C"

        'Banco
        Set No = tvwPermiss�es.Nodes.Add(intIndexI, tvwChild)
            No.Text = "Banco"
            No.Key = UCase(No.Text)
            intIndexII = No.Index
            'Inclus�o
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Incluir"
                No.Tag = "I"
            'Altera��o
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Alterar"
                No.Tag = "A"
            'Exclus�o
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Excluir"
                No.Tag = "E"
            'Consulta
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Consultar"
                No.Tag = "C"

        'Cliente
        Set No = tvwPermiss�es.Nodes.Add(intIndexI, tvwChild)
            No.Text = "Cliente"
            No.Key = UCase(No.Text)
            intIndexII = No.Index
            'Inclus�o
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Incluir"
                No.Tag = "I"
            'Altera��o
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Alterar"
                No.Tag = "A"
            'Exclus�o
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Excluir"
                No.Tag = "E"
            'Consulta
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Consultar"
                No.Tag = "C"
    
        'Conta Corrente
        Set No = tvwPermiss�es.Nodes.Add(intIndexI, tvwChild)
            No.Text = "Conta Corrente"
            No.Key = UCase(No.Text)
            intIndexII = No.Index
            'Inclus�o
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Incluir"
                No.Tag = "I"
            'Altera��o
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Alterar"
                No.Tag = "A"
            'Exclus�o
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Excluir"
                No.Tag = "E"
            'Consulta
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Consultar"
                No.Tag = "C"
        
        'Fornecedor
        Set No = tvwPermiss�es.Nodes.Add(intIndexI, tvwChild)
            No.Text = "Fornecedor"
            No.Key = UCase(No.Text)
            intIndexII = No.Index
            'Inclus�o
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Incluir"
                No.Tag = "I"
            'Altera��o
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Alterar"
                No.Tag = "A"
            'Exclus�o
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Excluir"
                No.Tag = "E"
            'Consulta
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Consultar"
                No.Tag = "C"
            
        'Funcion�rio
        Set No = tvwPermiss�es.Nodes.Add(intIndexI, tvwChild)
            No.Text = "Funcion�rio"
            No.Key = UCase(No.Text)
            intIndexII = No.Index
            'Inclus�o
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Incluir"
                No.Tag = "I"
            'Altera��o
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Alterar"
                No.Tag = "A"
            'Exclus�o
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Excluir"
                No.Tag = "E"
            'Consulta
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Consultar"
                No.Tag = "C"
            
        'Produto
        Set No = tvwPermiss�es.Nodes.Add(intIndexI, tvwChild)
            No.Text = "Produto"
            No.Key = UCase(No.Text)
            intIndexII = No.Index
            'Inclus�o
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Incluir"
                No.Tag = "I"
            'Altera��o
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Alterar"
                No.Tag = "A"
            'Exclus�o
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Excluir"
                No.Tag = "E"
            'Consulta
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Consultar"
                No.Tag = "C"
            'Quantidade
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Alterar Quantidade"
                No.Tag = "Q"
            'Alterar pre�o
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Alterar pre�o"
                No.Tag = "P"
            'Alterar c�digo
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Alterar c�digo"
                No.Tag = "G"
        
        'Despesas
        Set No = tvwPermiss�es.Nodes.Add(intIndexI, tvwChild)
            No.Text = "Despesas"
            No.Key = UCase(No.Text)
            intIndexII = No.Index
            'Inclus�o
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Incluir"
                No.Tag = "I"
            'Altera��o
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Alterar"
                No.Tag = "A"
            'Exclus�o
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Excluir"
                No.Tag = "E"
            'Consulta
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Consultar"
                No.Tag = "C"
            
    ' ** Movimento
    Set No = tvwPermiss�es.Nodes.Add
        No.Text = "Movimento"
        No.Key = UCase(No.Text)
        intIndexI = No.Index
        
        'Compra
        Set No = tvwPermiss�es.Nodes.Add(intIndexI, tvwChild)
            No.Text = "Compra"
            No.Key = UCase(No.Text)
            intIndexII = No.Index
            'Inclus�o
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Incluir"
                No.Tag = "I"
            'Altera��o
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Alterar"
                No.Tag = "A"
            'Exclus�o
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Excluir"
                No.Tag = "E"
            'Consulta
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Consultar"
                No.Tag = "C"
        
        'Devolu��o/Troca (Compra)
        Set No = tvwPermiss�es.Nodes.Add(intIndexI, tvwChild)
            No.Text = "Devolu��o/Troca (Compra)"
            No.Key = UCase(No.Text)
            intIndexII = No.Index
            'Inclus�o
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Incluir"
                No.Tag = "I"
            'Altera��o
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Alterar"
                No.Tag = "A"
            'Exclus�o
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Excluir"
                No.Tag = "E"
            'Consulta
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Consultar"
                No.Tag = "C"
            
        'Venda
        Set No = tvwPermiss�es.Nodes.Add(intIndexI, tvwChild)
            No.Text = "Venda"
            No.Key = UCase(No.Text)
            intIndexII = No.Index
            'Inclus�o
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Incluir"
                No.Tag = "I"
            'Altera��o
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Alterar"
                No.Tag = "A"
            'Exclus�o
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Excluir"
                No.Tag = "E"
            'Consulta
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Consultar"
                No.Tag = "C"
            'Entrar automaticamente em inclus�o
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Entrar automaticamente em inclus�o"
                No.Tag = "U"
                
            'Autorizar valor al�m do desconto m�ximo
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Permite autorizar desconto"
                No.Tag = "D"
            
        'Devolu��o/Troca (Venda)
        Set No = tvwPermiss�es.Nodes.Add(intIndexI, tvwChild)
            No.Text = "Devolu��o/Troca (Venda)"
            No.Key = UCase(No.Text)
            intIndexII = No.Index
            'Inclus�o
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Incluir"
                No.Tag = "I"
            'Altera��o
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Alterar"
                No.Tag = "A"
            'Exclus�o
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Excluir"
                No.Tag = "E"
            'Consulta
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Consultar"
                No.Tag = "C"
            
        'Movimento Di�rio
        Set No = tvwPermiss�es.Nodes.Add(intIndexI, tvwChild)
            No.Text = "Movimento Di�rio"
            No.Key = UCase(No.Text)
            intIndexII = No.Index
            'Inclus�o
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Incluir"
                No.Tag = "I"
            'Altera��o
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Alterar"
                No.Tag = "A"
            'Exclus�o
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Excluir"
                No.Tag = "E"
            'Consulta
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Consultar"
                No.Tag = "C"
        
        'Conta Corrente (Movimento)
        Set No = tvwPermiss�es.Nodes.Add(intIndexI, tvwChild)
            No.Text = "Conta Corrente (Movimento)"
            No.Key = UCase(No.Text)
            intIndexII = No.Index
            'Inclus�o
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Incluir"
                No.Tag = "I"
            'Altera��o
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Alterar"
                No.Tag = "A"
            'Exclus�o
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Excluir"
                No.Tag = "E"
            'Consulta
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Consultar"
                No.Tag = "C"
        
        'Caixa
        Set No = tvwPermiss�es.Nodes.Add(intIndexI, tvwChild)
            No.Text = "Caixa"
            No.Key = UCase(No.Text)
            intIndexII = No.Index
            'Operar
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Operar"
                No.Tag = "O"
            'Entrada
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Entrada"
                No.Tag = "E"
            'Sangria
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Sangria"
                No.Tag = "S"
            'Cancelar �ltimo item
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Cancelar �ltimo item"
                No.Tag = "U"
            'Cancelar
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Cancelar item"
                No.Tag = "C"
                
        'Caixa F�cil
        Set No = tvwPermiss�es.Nodes.Add(intIndexI, tvwChild)
            No.Text = "Caixa F�cil"
            No.Key = UCase(No.Text)
            intIndexII = No.Index
            'Operar
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Operar"
                No.Tag = "O"
        
        'PDV
        Set No = tvwPermiss�es.Nodes.Add(intIndexI, tvwChild)
            No.Text = "PDV"
            No.Key = UCase(No.Text)
            intIndexII = No.Index
            'Acesso
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Acesso"
                No.Tag = "A"
            
        'Abertura do Caixa
        Set No = tvwPermiss�es.Nodes.Add(intIndexI, tvwChild)
            No.Text = "Abertura do Caixa"
            No.Key = UCase(No.Text)
            intIndexII = No.Index
            'Acesso
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Acesso"
                No.Tag = "A"
                
        'Fechamento do Caixa
        Set No = tvwPermiss�es.Nodes.Add(intIndexI, tvwChild)
            No.Text = "Fechamento do Caixa"
            No.Key = UCase(No.Text)
            intIndexII = No.Index
            'Acesso
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Acesso"
                No.Tag = "A"
                
    ' ** Par�metros
    Set No = tvwPermiss�es.Nodes.Add
        No.Text = "Par�metros"
        No.Key = UCase(No.Text)
        intIndexI = No.Index
    
        'Departamento
        Set No = tvwPermiss�es.Nodes.Add(intIndexI, tvwChild)
            No.Text = "Departamento"
            No.Key = UCase(No.Text)
            intIndexII = No.Index
            'Inclus�o
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Incluir"
                No.Tag = "I"
            'Altera��o
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Alterar"
                No.Tag = "A"
            'Exclus�o
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Excluir"
                No.Tag = "E"
            'Consulta
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Consultar"
                No.Tag = "C"
        
        'Se��o
        Set No = tvwPermiss�es.Nodes.Add(intIndexI, tvwChild)
            No.Text = "Se��o"
            No.Key = UCase(No.Text)
            intIndexII = No.Index
            'Inclus�o
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Incluir"
                No.Tag = "I"
            'Altera��o
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Alterar"
                No.Tag = "A"
            'Exclus�o
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Excluir"
                No.Tag = "E"
            'Consulta
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Consultar"
                No.Tag = "C"
        
        'Departamento - Se��o
        Set No = tvwPermiss�es.Nodes.Add(intIndexI, tvwChild)
            No.Text = "Departamento - Se��o"
            No.Key = UCase(No.Text)
            intIndexII = No.Index
            'Inclus�o
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Incluir"
                No.Tag = "I"
            'Altera��o
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Alterar"
                No.Tag = "A"
            'Exclus�o
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Excluir"
                No.Tag = "E"
            'Consulta
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Consultar"
                No.Tag = "C"
    
        'Tipo de ICM
        Set No = tvwPermiss�es.Nodes.Add(intIndexI, tvwChild)
            No.Text = "Tipo de ICM"
            No.Key = UCase(No.Text)
            intIndexII = No.Index
            'Inclus�o
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Incluir"
                No.Tag = "I"
            'Altera��o
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Alterar"
                No.Tag = "A"
            'Exclus�o
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Excluir"
                No.Tag = "E"
            'Consulta
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Consultar"
                No.Tag = "C"
            
        'Tipo de Embalagem
        Set No = tvwPermiss�es.Nodes.Add(intIndexI, tvwChild)
            No.Text = "Tipo de Embalagem"
            No.Key = UCase(No.Text)
            intIndexII = No.Index
            'Inclus�o
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Incluir"
                No.Tag = "I"
            'Altera��o
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Alterar"
                No.Tag = "A"
            'Exclus�o
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Excluir"
                No.Tag = "E"
            'Consulta
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Consultar"
                No.Tag = "C"
            
        'Unidades
        Set No = tvwPermiss�es.Nodes.Add(intIndexI, tvwChild)
            No.Text = "Unidades"
            No.Key = UCase(No.Text)
            intIndexII = No.Index
            'Inclus�o
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Incluir"
                No.Tag = "I"
            'Altera��o
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Alterar"
                No.Tag = "A"
            'Exclus�o
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Excluir"
                No.Tag = "E"
            'Consulta
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Consultar"
                No.Tag = "C"
            
        'Localidade de Estoque
        Set No = tvwPermiss�es.Nodes.Add(intIndexI, tvwChild)
            No.Text = "Localidade de Estoque"
            No.Key = UCase(No.Text)
            intIndexII = No.Index
            'Inclus�o
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Incluir"
                No.Tag = "I"
            'Altera��o
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Alterar"
                No.Tag = "A"
            'Exclus�o
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Excluir"
                No.Tag = "E"
            'Consulta
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Consultar"
                No.Tag = "C"
        
        'Plano de Pagamento
        Set No = tvwPermiss�es.Nodes.Add(intIndexI, tvwChild)
            No.Text = "Plano de Pagamento"
            No.Key = UCase(No.Text)
            intIndexII = No.Index
            'Inclus�o
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Incluir"
                No.Tag = "I"
            'Altera��o
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Alterar"
                No.Tag = "A"
            'Exclus�o
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Excluir"
                No.Tag = "E"
            'Consulta
            Set No = tvwPermiss�es.Nodes.Add(intIndexII, tvwChild)
                No.Text = "Consultar"
                No.Tag = "C"
End Sub
Public Sub Excluir()
    Dim Confirma��o As Integer, Msg1$, Msg2$

    If lAlterar Then
       If Not Cancelamento Then
            Exit Sub
        End If
    End If

    StatusBarAviso = "Exclus�o"
    BarraDeStatus StatusBarAviso
    
    Msg1 = "Voc� est� preste a apagar um registro !"
    Msg2 = "Tem certeza?"
    Msg2 = String(((Len(Msg1) - Len(Msg2)) / 2), " ") + Msg2
    Confirma��o = MsgBox(Msg1 + vbCr + Msg2, vbYesNo + vbQuestion + vbDefaultButton2, "Confirma��o")
    
    If Confirma��o = vbNo Then
        Exit Sub
    End If
    
    WS.BeginTrans
    
    TBLGrupos.Delete
    
    If Err <> 0 Then
        GeraMensagemDeErro "Grupos - Excluir", True
        StatusBarAviso = "Falha na exclus�o"
        BarraDeStatus StatusBarAviso
        Exit Sub
    End If
    
    WS.CommitTrans
    
    StatusBarAviso = "Exclus�o bem sucedida"
    BarraDeStatus StatusBarAviso
    
    If TBLGrupos.RecordCount = 0 Then
        Navega��oInferior False
        Navega��oSuperior False
        Bot�oExcluir False
        Bot�oGravar False
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
            StatusBarAviso = "Inclus�o bem sucedida"
        Else
            StatusBarAviso = "Falha na inclus�o"
        End If
    Else
        If TBLGrupos.RecordCount > 0 And Not TBLGrupos.BOF And Not TBLGrupos.EOF Then
            If SetRecords Then
                PosRecords
                lAlterar = False
                StatusBarAviso = "Altera��o bem sucedida"
            Else
                StatusBarAviso = "Falha na altera��o"
            End If
        End If
    End If
    
    BarraDeStatus StatusBarAviso
    
    TestaInferior TBLGrupos, True, True, True
    TestaSuperior TBLGrupos, True, True, True
    
    If TBLGrupos.RecordCount = 0 Then
        If Not lInserir And Not lAlterar Then
            Bot�oExcluir False
            Bot�oGravar False
            cmdGravar.Enabled = False
            cmdCancelar.Enabled = False
        End If
    Else
        Bot�oExcluir True
    End If
    
    Bot�oIncluir True
    
    If txtC�digo.Enabled Then
        txtC�digo.SetFocus
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
    
    Bot�oGravar (lInserir Or lAllowEdit)
    Bot�oIncluir False
    cmdGravar.Enabled = (lInserir Or lAllowEdit)
    cmdCancelar.Enabled = (lInserir Or lAllowEdit)
    
    Navega��oInferior False
    Navega��oSuperior False
    
    StatusBarAviso = "Inclus�o"
    BarraDeStatus StatusBarAviso
    
    MontarMarcas
    txtC�digo.SetFocus
End Sub
Private Sub MontarMarcas()
    Dim Cont As Byte
    Dim No As Node
    Dim LastNo As String
    Dim lSair As Boolean
    Dim strResultado As String
    
    Set No = tvwPermiss�es.Nodes.Item(1).Root
    
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
    
    Navega��oInferior False
    Navega��oSuperior True
    
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
    
    Navega��oInferior True
    Navega��oSuperior False
    
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
    
    Navega��oInferior True
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
    
    Navega��oSuperior True
    TestaInferior TBLGrupos, True, True, True
    
    GetRecords
End Sub
Public Sub PosRecords()
    If TBLGrupos.RecordCount = 0 Then
        Exit Sub
    End If
    
    TBLGrupos.Seek "=", txtC�digo
    If TBLGrupos.NoMatch Then
        MsgBox "N�o consegui encontrar " + txtC�digo, vbExclamation, "Erro"
        TBLGrupos.MoveFirst
        Navega��oInferior False
        Navega��oInferior True
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
    txtC�digo = TBLGrupos("C�DIGO")
    txtDescri��o = TBLGrupos("DESCRI��O")
    
    For Cont = 2 To TBLGrupos.Fields.Count - 1
        Set No = tvwPermiss�es.Nodes.Item(TBLGrupos(Cont).Name)
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
    Dim Confirma��o As Integer, Msg1$, Msg2$, Cont As Byte, Cont1 As Byte
    Dim No As Node, NoFilho As Node, LinhaDePermiss�o As String
    
    WS.BeginTrans 'Inicia uma Transa��o
    
    If lInserir Then
        TBLGrupos.AddNew
    Else
        TBLGrupos.Edit
    End If
    
    TBLGrupos("C�DIGO") = txtC�digo
    TBLGrupos("DESCRI��O") = txtDescri��o
    
    For Cont = 2 To TBLGrupos.Fields.Count - 1
        Set No = tvwPermiss�es.Nodes.Item(TBLGrupos(Cont).Name)
        LinhaDePermiss�o = Empty
        Set NoFilho = No.Child
        For Cont1 = 1 To No.Children
            If NoFilho.Checked Then
                LinhaDePermiss�o = LinhaDePermiss�o + NoFilho.Tag
            End If
            Set NoFilho = NoFilho.Next
        Next
        TBLGrupos(Cont) = LinhaDePermiss�o
    Next
    
    TBLGrupos.Update
        
Erro:
    If Err <> 0 Then
        TBLGrupos.CancelUpdate
        GeraMensagemDeErro "Grupos - SetRecords", True
        SetRecords = False
        Exit Function
    End If

    WS.CommitTrans 'Grava as altera��es ou inclus�es se n�o houverem erros
    
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
    
    txtC�digo = Empty
    txtDescri��o = Empty
    
    Set No = tvwPermiss�es.Nodes.Item(1).Root
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
        Bot�oGravar False
        cmdGravar.Enabled = False
        cmdCancelar.Enabled = False
        Bot�oImprimir False
    Else
        Bot�oGravar (lInserir Or lAllowEdit)
        cmdGravar.Enabled = (lInserir Or lAllowEdit)
        cmdCancelar.Enabled = (lInserir Or lAllowEdit)
        Bot�oImprimir True
    End If
    
    If lInserir Then
        Bot�oGravar (lInserir Or lAllowEdit)
        cmdGravar.Enabled = (lInserir Or lAllowEdit)
        cmdCancelar.Enabled = (lInserir Or lAllowEdit)
        Navega��oInferior False
        Navega��oSuperior False
        Bot�oExcluir False
        Bot�oIncluir False
    ElseIf lAlterar Then
        Bot�oIncluir True
    Else
        Bot�oIncluir True
        StatusBarAviso = "Pronto"
    End If
    
    If lAtualizar Then
        Bot�oAtualizar True
    Else
        Bot�oAtualizar False
    End If
    
    BarraDeStatus StatusBarAviso
End Sub
Private Sub Form_Deactivate()
    cmdGravar.Enabled = False
    cmdCancelar.Enabled = False
    Bot�oImprimir False
End Sub
Private Sub Form_Load()
    On Error GoTo Erro
    
    lAllowInsert = True
    lAllowEdit = True
    lAllowDelete = True
    lAllowConsult = True
    
    EstruturaDasPermiss�es
    
    ZeraCampos

    lPula = False
    lInserir = False
    lAlterar = False
    
    GruposAberto = AbreTabela(Dicion�rio, "USU�RIO", "GRUPO", DBUsu�rio, TBLGrupos, TBLTabela, dbOpenTable)
    
    If GruposAberto Then
        IndiceGruposAtivo = "GRUPO1"
        TBLGrupos.Index = IndiceGruposAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'GRUPO' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    Bot�oIncluir True
 
    If TBLGrupos.RecordCount = 0 Then
        DesativaCampos
        Bot�oExcluir False
        Bot�oGravar False
    Else
        AtivaCampos
        Bot�oExcluir True
        Bot�oGravar (lInserir Or lAllowEdit)
        If Not GetRecords Then
            mFechar = True
            Exit Sub
        End If
    End If
    
    Navega��oInferior False
        
    If TBLGrupos.RecordCount = 0 Or TBLGrupos.RecordCount = 1 Then
        Navega��oSuperior False
    Else
        Navega��oInferior True
    End If
                        
    StatusBarAviso = "Pronto"
    Relat�rio = AddPath(Aplica��oPath, "REPORT\Grupos.RPT")
    TotalDatabaseName = 1
    DataBaseName(1) = AddPath(Aplica��oPath, "DATABASE\USU�RIO.MDB")
    mFechar = False
    
    Exit Sub
    
Erro:
    GeraMensagemDeErro "Grupos - Load"
    mFechar = True
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If lInserir Then
        MsgBox "Voc� est� em uma inclus�o!", vbExclamation, Caption
        StatusBarAviso = "Finalize a inclus�o"
        BarraDeStatus StatusBarAviso
        Cancel = 1
        SetaFocus Me
        mdiGeal.Mostrar
        Exit Sub
    End If
    If lAlterar Then
        MsgBox "Voc� est� em uma altera��o!", vbExclamation, Caption
        StatusBarAviso = "Finalize a altera��o"
        BarraDeStatus StatusBarAviso
        Cancel = 1
        SetaFocus Me
        mdiGeal.Mostrar
        Exit Sub
    End If
    
    mdiGeal.StatusBar.Panels("Posi��o").Visible = False
    ResizeStatusBar
    
    Set frmGrupos = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If GruposAberto Then
        TBLGrupos.Close
    End If
    If Forms.Count = 2 Then
        AllBot�es False
    End If
End Sub
Private Sub tvwPermiss�es_NodeCheck(ByVal Node As MSComctlLib.Node)
    Dim NextNo As Node, Cont As Byte
        
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Altera��o"
        BarraDeStatus StatusBarAviso
    End If
    
    'Verifica se h� filho
    If Node.Children <> 0 Then
        VerificaChildren Node.Child, Node.Children, Node.Checked
    End If
    
    MontarMarcas
End Sub
Private Sub txtC�digo_Change()
    If Not lPula Then
        FormatMask "9999", txtC�digo
    End If
End Sub
Private Sub txtC�digo_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Altera��o"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub txtC�digo_LostFocus()
    LeftBlank txtC�digo
End Sub
Private Sub txtDescri��o_Change()
    If Not lPula Then
        FormatMask "@!S30", txtDescri��o
    End If
End Sub
Private Sub txtDescri��o_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Altera��o"
        BarraDeStatus StatusBarAviso
    End If
End Sub

