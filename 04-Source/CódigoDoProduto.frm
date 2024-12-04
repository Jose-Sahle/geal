VERSION 5.00
Begin VB.Form frmCódigoDoProduto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Código do Produto"
   ClientHeight    =   1680
   ClientLeft      =   2430
   ClientTop       =   1530
   ClientWidth     =   4800
   ForeColor       =   &H8000000D&
   Icon            =   "CódigoDoProduto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1680
   ScaleWidth      =   4800
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   3510
      TabIndex        =   3
      Top             =   1275
      Width           =   1245
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "&Gravar"
      Default         =   -1  'True
      Height          =   345
      Left            =   2190
      TabIndex        =   2
      Top             =   1275
      Width           =   1245
   End
   Begin VB.Frame frCódigoDoProduto 
      Height          =   1230
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   4785
      Begin VB.TextBox txtCódigo 
         Height          =   285
         Left            =   1200
         TabIndex        =   1
         Top             =   690
         Width           =   3405
      End
      Begin VB.TextBox txtFornecedor 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1200
         TabIndex        =   0
         Text            =   "  .   .   /    -"
         Top             =   300
         Width           =   2250
      End
      Begin VB.Label lblCódigo 
         Caption         =   "Código"
         Height          =   195
         Left            =   150
         TabIndex        =   6
         Top             =   720
         Width           =   765
      End
      Begin VB.Label lblFornecedor 
         Caption         =   "Fornecedor"
         ForeColor       =   &H80000017&
         Height          =   330
         Left            =   150
         TabIndex        =   5
         Top             =   330
         Width           =   900
      End
   End
End
Attribute VB_Name = "frmCódigoDoProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lPula As Boolean

Dim Elemento%

Dim StatusBar$

Dim lInserir As Boolean
Dim lAlterar As Boolean
Dim lAceitar As Boolean
Public lAlteração As Boolean

Dim lAllowInsert  As Boolean
Dim lAllowEdit    As Boolean
Dim lAllowDelete  As Boolean
Dim lAllowConsult As Boolean

Dim ArrayProdutoTotal%
Dim ArrayProdutoCódigo() As Variant
Dim ArrayProdutoFornecedor() As Variant

Public lAtualizar As Boolean
Private Sub AtivaCampos()
    frCódigoDoProduto.Enabled = True
    BotãoGravar (lInserir Or lAllowEdit)
    cmdCancelar.Enabled = (lInserir Or lAllowEdit)
    cmdGravar.Enabled = (lInserir Or lAllowEdit)
End Sub
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
        BarraDeStatus "Inclusão cancelada"
    End If
    If lAlterar Then
        BarraDeStatus "Alteração cancelada"
    End If
    
    BotãoIncluir lAllowInsert
    
    If ArrayProdutoTotal = 0 Then
        NavegaçãoInferior False
        NavegaçãoSuperior False
        BotãoGravar False
        cmdGravar.Enabled = False
        cmdCancelar.Enabled = False
        DesativaCampos
        ZeraCampos
        Cancelamento = True
        lInserir = False
        lAlterar = False
        Exit Function
    End If
    
    lInserir = False
    lAlterar = False
    
    Cancelamento = True
    
    TestaInferiorArray Elemento, ArrayProdutoCódigo(), lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperiorArray Elemento, ArrayProdutoCódigo(), lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Function
Private Sub ClearArray()
    ArrayProdutoTotal = 0
    ReDim ArrayProdutoCódigo(1 To 1)
    ReDim ArrayProdutoFornecedor(1 To 1)
End Sub
Private Sub DesativaCampos()
    frCódigoDoProduto.Enabled = False
    cmdGravar.Enabled = False
    cmdCancelar.Enabled = False
    BotãoGravar False
End Sub
Public Sub Excluir()
    Dim Confirmação As Integer, Msg1$, Msg2$
    Dim SQL As String
    
    If lAlterar Then
       If Not Cancelamento Then
            Exit Sub
        End If
    End If
    
    Msg1 = "Você está preste a apagar um registro !"
    Msg2 = "Tem certeza?"
    Msg2 = String(((Len(Msg1) - Len(Msg2)) / 2), " ") + Msg2
    Confirmação = MsgBox(Msg1 + vbCr + Msg2, vbYesNo + vbQuestion + vbDefaultButton2, "Confirmação")
    
    If Confirmação = vbNo Then
        Exit Sub
    End If
    
    BarraDeStatus "Exclusão de Código do Produto"

    Adel Elemento, ArrayProdutoCódigo()
    Adel Elemento, ArrayProdutoFornecedor()
    
    ArrayProdutoTotal = ArrayProdutoTotal - 1
    
    If Elemento > ArrayProdutoTotal Then
        Elemento = ArrayProdutoTotal
    End If
    
    lAlteração = True
    
    If ArrayProdutoTotal = 0 Then
        ClearArray
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
    
    SizeArray ArrayProdutoTotal
    
    Log gUsuário, "Exclusão - Código do Produto: " & txtCódigo
    
    GetRecords
    
    TestaInferiorArray Elemento, ArrayProdutoCódigo(), lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperiorArray Elemento, ArrayProdutoCódigo(), lAllowEdit, lAllowDelete, lAllowConsult
End Sub
Private Sub FillArray()
    Dim Cont%
    
    SizeArray frmProduto.ArrayProdutoTotal
    ArrayProdutoTotal = frmProduto.ArrayProdutoTotal
    
    For Cont = 1 To frmProduto.ArrayProdutoTotal
        ArrayProdutoCódigo(Cont) = frmProduto.GetArrayProduto("Código", Cont)
        ArrayProdutoFornecedor(Cont) = frmProduto.GetArrayProduto("Fornecedor", Cont)
    Next
End Sub
Public Sub Gravar()
    If lInserir Then
        ArrayProdutoTotal = ArrayProdutoTotal + 1
        Elemento = ArrayProdutoTotal
        SizeArray ArrayProdutoTotal
        SetRecords
        PosRecords
        BarraDeStatus "Inclusão bem sucedida"
        lInserir = False
        lAlteração = True
    ElseIf lAlterar Then
        If ArrayProdutoTotal > 0 Then
            SetRecords
            PosRecords
            lAlterar = False
            lAlteração = True
            BarraDeStatus "Alteração bem sucedida"
        End If
    Else
        Exit Sub
    End If
    
    TestaInferiorArray Elemento, ArrayProdutoCódigo(), lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperiorArray Elemento, ArrayProdutoCódigo(), lAllowEdit, lAllowDelete, lAllowConsult
    
    If ArrayProdutoTotal = 0 Then
        If Not lInserir And Not lAlterar Then
            BotãoExcluir False
            BotãoGravar False
            cmdGravar.Enabled = False
            cmdCancelar.Enabled = False
        End If
    Else
        BotãoExcluir lAllowDelete
    End If
    
    BotãoIncluir lAllowInsert
    StatusBar = frmProduto.StatusBarAviso
    
    If txtFornecedor.Enabled Then
        txtFornecedor.SetFocus
    End If
End Sub
Public Sub Incluir()
    If lAlterar Then
       If Not Cancelamento Then
            Exit Sub
        End If
    End If
    
    lInserir = True
        
    StatusBar = "Inclusão de Lote do Produto"
    BarraDeStatus StatusBar
    
    ZeraCampos
    AtivaCampos
    
    BotãoGravar (lInserir Or lAllowEdit)
    BotãoIncluir False
    cmdGravar.Enabled = (lInserir Or lAllowEdit)
    cmdCancelar.Enabled = (lInserir Or lAllowEdit)
    
    NavegaçãoInferior False
    NavegaçãoSuperior False
    
    txtFornecedor.SetFocus
End Sub
Public Sub MoveFirst()
    If lAlterar Then
       If Not Cancelamento Then
            Exit Sub
        End If
    End If
        
    StatusBar = frmProduto.StatusBarAviso
    BarraDeStatus StatusBar
    
    Elemento = 1
    
    NavegaçãoInferior False
    NavegaçãoSuperior lAllowConsult
    
    GetRecords
End Sub
Public Sub MoveLast()
    If lAlterar Then
       If Not Cancelamento Then
            Exit Sub
        End If
    End If
    
    StatusBar = frmProduto.StatusBarAviso
    BarraDeStatus StatusBar
    
    Elemento = ArrayProdutoTotal
    
    NavegaçãoInferior lAllowConsult
    NavegaçãoSuperior False
    
    GetRecords
End Sub
Public Sub MoveNext()
    If lAlterar Then
       If Not Cancelamento Then
            Exit Sub
        End If
    End If
    
    StatusBar = frmProduto.StatusBarAviso
    BarraDeStatus StatusBar
    
    Elemento = Elemento + 1
    
    If Elemento > ArrayProdutoTotal Then
        Elemento = ArrayProdutoTotal
        Exit Sub
    End If
    
    NavegaçãoInferior lAllowConsult
    TestaSuperiorArray Elemento, ArrayProdutoCódigo(), lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Sub
Public Sub MovePrevious()
    If lAlterar Then
       If Not Cancelamento Then
            Exit Sub
        End If
    End If
    
    StatusBar = frmProduto.StatusBarAviso
    BarraDeStatus StatusBar
    
    Elemento = Elemento - 1
    
    If Elemento < 1 Then
        Elemento = 1
        Exit Sub
    End If
    
    NavegaçãoSuperior lAllowConsult
    TestaInferiorArray Elemento, ArrayProdutoCódigo(), lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Sub
Public Sub PosRecords()

End Sub
Private Sub GetRecords()
    On Error GoTo Erro
    
    If Not lAllowConsult Then
        ZeraCampos
        DesativaCampos
        Exit Sub
    End If
    txtCódigo = ArrayProdutoCódigo(Elemento)
    txtFornecedor = ArrayProdutoFornecedor(Elemento)
    If Not lAllowEdit Then
        DesativaCampos
    End If
    
    Exit Sub
    
Erro:
    GeraMensagemDeErro "Código Do Produto - GetRecords "
    Resume Next
End Sub
Private Function SetRecords()
    ArrayProdutoCódigo(Elemento) = txtCódigo
    ArrayProdutoFornecedor(Elemento) = txtFornecedor
    Log gUsuário, "Código do Produto " & txtCódigo
End Function
Private Sub SaveArray()
    Dim Cont%
    
    frmProduto.SizeArrayProduto ArrayProdutoTotal
    frmProduto.ArrayProdutoTotal = ArrayProdutoTotal
    
    For Cont = 1 To frmProduto.ArrayProdutoTotal
        frmProduto.SetArrayProduto "Código", ArrayProdutoCódigo(Cont), Cont
        frmProduto.SetArrayProduto "Fornecedor", ArrayProdutoFornecedor(Cont), Cont
    Next
End Sub
Private Sub SizeArray(ByVal Tamanho As Integer)
    ASize Tamanho, ArrayProdutoCódigo()
    ASize Tamanho, ArrayProdutoFornecedor()
End Sub
Private Sub ZeraCampos()
    txtCódigo = Empty
    txtFornecedor = Empty
End Sub
Private Sub cmdCancelar_Click()
    Cancelamento
End Sub
Private Sub cmdGravar_Click()
    Gravar
End Sub
Private Sub Form_Activate()
    TestaInferiorArray Elemento, ArrayProdutoCódigo(), lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperiorArray Elemento, ArrayProdutoCódigo(), lAllowEdit, lAllowDelete, lAllowConsult
    
    BarraDeStatus StatusBar
    
    If ArrayProdutoTotal = 0 Then
        BotãoGravar False
        cmdGravar.Enabled = False
        cmdCancelar.Enabled = False
    Else
        BotãoGravar (lInserir Or lAllowEdit)
        cmdGravar.Enabled = (lInserir Or lAllowEdit)
        cmdCancelar.Enabled = (lInserir Or lAllowEdit)
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
        BotãoIncluir lAllowInsert
    Else
        BotãoIncluir lAllowInsert
    End If
    
    If lAtualizar Then
        BotãoAtualizar True
    Else
        BotãoAtualizar False
    End If
End Sub
Private Sub Form_Deactivate()
    Beep
    frmCódigoDoProduto.SetFocus
End Sub
Private Sub Form_Load()
    lPula = False
    
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    
    lAllowInsert = Allow("PRODUTO", "I")
    lAllowEdit = Allow("PRODUTO", "G")
    lAllowDelete = Allow("PRODUTO", "E")
    lAllowConsult = Allow("PRODUTO", "C")
    
    FillArray
        
    BotãoIncluir lAllowInsert
    
    StatusBar = frmProduto.StatusBarAviso
 
    If ArrayProdutoTotal = 0 Then
        DesativaCampos
        BotãoExcluir False
        BotãoGravar False
    Else
        Elemento = 1
        AtivaCampos
        BotãoExcluir lAllowDelete
        BotãoGravar (lInserir Or lAllowEdit)
        GetRecords
    End If
    
    NavegaçãoInferior False
        
    If ArrayProdutoTotal = 0 Or ArrayProdutoTotal = 1 Then
        NavegaçãoSuperior False
    Else
        NavegaçãoInferior lAllowConsult
    End If
        
    lInserir = False
    lAlterar = False
    Exit Sub
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim Confirmação
    
    If lInserir Then
        MsgBox "Você está em uma inclusão!", vbExclamation, Caption
        BarraDeStatus "Finalize a inclusão"
        Cancel = 1
        SetaFocus Me
        mdiGeal.Mostrar
        Exit Sub
    End If
    If lAlterar Then
        MsgBox "Você está em uma alteração!", vbExclamation, Caption
        BarraDeStatus "Finalize a alteração"
        Cancel = 1
        SetaFocus Me
        mdiGeal.Mostrar
        Exit Sub
    End If
    
    If lAlteração Then
        Confirmação = MsgBox("       Aceita as alterações ?", vbQuestion + vbDefaultButton1 + vbYesNoCancel, "Confirmação")
        
        If Confirmação = vbYes Then
            SaveArray
            If Not frmProduto.lAlterarArrayProduto Then
                frmProduto.lAlterarArrayProduto = True
                If Not frmProduto.lInserir Then
                    frmProduto.lAlterar = True
                End If
            End If
        ElseIf Confirmação = vbCancel Then
            Cancel = 1
        ElseIf Confirmação = vbNo Then
        End If
    End If
    
    Set frmCódigoDoProduto = Nothing
End Sub
Private Sub txtCódigo_Change()
    FormatMask "@!S13", txtCódigo
End Sub
Private Sub txtCódigo_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBar = "Alteração de Código do Produto"
        BarraDeStatus StatusBar
    End If
End Sub
Private Sub txtFornecedor_Change()
    FormatMask "99.999.999/9999-99", txtFornecedor
End Sub
Private Sub txtFornecedor_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBar = "Alteração de Código do Produto"
        BarraDeStatus StatusBar
    End If
End Sub
Private Sub txtFornecedor_LostFocus()
    If Not IsCorrectFornecedor(txtFornecedor) Then
        MsgBox "Fornecedor não cadastrado!", vbInformation, "Aviso"
        Set frmEncontrar.DBBancoDeDados = DBCadastro
        frmEncontrar.Mensagem = "Nenhum fornecedor foi selecionado!"
        frmEncontrar.BancoDeDados = "CADASTRO"
        frmEncontrar.Tabela = "FORNECEDOR"
        frmEncontrar.Indice = "2"
        frmEncontrar.CampoChave = "CGC - CPF"
        frmEncontrar.CampoPreencheLista = "RAZÃO SOCIAL"
        frmEncontrar.Show vbModal
        txtFornecedor = frmEncontrar.Chave
    End If
End Sub
