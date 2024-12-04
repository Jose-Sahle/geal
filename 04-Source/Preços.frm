VERSION 5.00
Begin VB.Form frmPreços 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Preços"
   ClientHeight    =   2460
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   4800
   Icon            =   "Preços.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2460
   ScaleWidth      =   4800
   Begin VB.Frame frPreços 
      Height          =   2010
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   4785
      Begin VB.VScrollBar vscrMargemDeLucro 
         Height          =   345
         Left            =   2295
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1455
         Width           =   210
      End
      Begin VB.TextBox txtMargemDeLucro 
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
         Left            =   1470
         TabIndex        =   3
         Text            =   "  0,00"
         Top             =   1470
         Width           =   825
      End
      Begin VB.TextBox txtPreçoDeVenda 
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
         Left            =   1470
         TabIndex        =   2
         Text            =   "         0,00"
         Top             =   1080
         Width           =   1680
      End
      Begin VB.TextBox txtPreçoDeCusto 
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
         Left            =   1470
         TabIndex        =   1
         Text            =   "         0,00"
         Top             =   690
         Width           =   1680
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
         Left            =   1470
         TabIndex        =   0
         Text            =   "  .   .   /    -"
         Top             =   300
         Width           =   2280
      End
      Begin VB.Label lblPorcentagem 
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2580
         TabIndex        =   12
         Top             =   1470
         Width           =   285
      End
      Begin VB.Label lblMargemDeLucro 
         Caption         =   "Margem de Lucro"
         Height          =   240
         Left            =   135
         TabIndex        =   10
         Top             =   1500
         Width           =   1305
      End
      Begin VB.Label lblPreçoDeVenda 
         Caption         =   "Preço de Venda"
         Height          =   195
         Left            =   150
         TabIndex        =   9
         Top             =   1110
         Width           =   1215
      End
      Begin VB.Label lblPreçoDeCusto 
         Caption         =   "Preço de Custo"
         Height          =   195
         Left            =   150
         TabIndex        =   8
         Top             =   720
         Width           =   1170
      End
      Begin VB.Label lblFornecedor 
         Caption         =   "Fornecedor"
         Height          =   330
         Left            =   150
         TabIndex        =   7
         Top             =   330
         Width           =   900
      End
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "&Gravar"
      Default         =   -1  'True
      Height          =   345
      Left            =   2190
      TabIndex        =   4
      Top             =   2070
      Width           =   1245
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   3510
      TabIndex        =   5
      Top             =   2070
      Width           =   1245
   End
End
Attribute VB_Name = "frmPreços"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim lPula As Boolean

Dim Elemento%

Dim StatusBar$

Dim lAllowInsert  As Boolean
Dim lAllowEdit    As Boolean
Dim lAllowDelete  As Boolean
Dim lAllowConsult As Boolean

Dim lInserir As Boolean
Dim lAlterar As Boolean
Dim lAceitar As Boolean
Public lAlteração As Boolean

Dim ArrayPreçoTotal%
Dim ArrayPreçoFornecedor() As Variant
Dim ArrayPreçoCusto() As Variant
Dim ArrayPreçoVenda() As Variant
Dim ArrayPreçoLucro() As Variant

Public lAtualizar As Boolean
Private Sub AtivaCampos()
    frPreços.Enabled = True
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
    
    If ArrayPreçoTotal = 0 Then
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
    
    TestaInferiorArray Elemento, ArrayPreçoFornecedor(), lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperiorArray Elemento, ArrayPreçoFornecedor(), lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Function
Private Sub ClearArray()
    ArrayPreçoTotal = 0
    ReDim ArrayPreçoFornecedor(1 To 1)
    ReDim ArrayPreçoCusto(1 To 1)
    ReDim ArrayPreçoVenda(1 To 1)
    ReDim ArrayPreçoLucro(1 To 1)
End Sub
Private Sub DesativaCampos()
    frPreços.Enabled = False
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
    
    BarraDeStatus "Exclusão de Preço do Produto"

    Adel Elemento, ArrayPreçoFornecedor()
    Adel Elemento, ArrayPreçoCusto()
    Adel Elemento, ArrayPreçoVenda()
    Adel Elemento, ArrayPreçoLucro()
    
    ArrayPreçoTotal = ArrayPreçoTotal - 1
    
    If Elemento > ArrayPreçoTotal Then
        Elemento = ArrayPreçoTotal
    End If
    
    lAlteração = True
    
    If ArrayPreçoTotal = 0 Then
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
    
    SizeArray ArrayPreçoTotal
    
    GetRecords
    
    TestaInferiorArray Elemento, ArrayPreçoFornecedor(), lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperiorArray Elemento, ArrayPreçoFornecedor(), lAllowEdit, lAllowDelete, lAllowConsult
End Sub
Private Sub FillArray()
    Dim Cont%
    
    SizeArray frmProduto.ArrayPreçoTotal
    ArrayPreçoTotal = frmProduto.ArrayPreçoTotal
    
    For Cont = 1 To frmProduto.ArrayPreçoTotal
        ArrayPreçoFornecedor(Cont) = frmProduto.GetArrayPreço("Fornecedor", Cont)
        ArrayPreçoCusto(Cont) = frmProduto.GetArrayPreço("Custo", Cont)
        ArrayPreçoVenda(Cont) = frmProduto.GetArrayPreço("Venda", Cont)
        ArrayPreçoLucro(Cont) = frmProduto.GetArrayPreço("Lucro", Cont)
    Next
End Sub
Public Sub Gravar()
    If lInserir Then
        ArrayPreçoTotal = ArrayPreçoTotal + 1
        Elemento = ArrayPreçoTotal
        SizeArray ArrayPreçoTotal
        SetRecords
        PosRecords
        BarraDeStatus "Inclusão bem sucedida"
        lInserir = False
        lAlteração = True
    ElseIf lAlterar Then
        If ArrayPreçoTotal > 0 Then
            SetRecords
            PosRecords
            lAlterar = False
            lAlteração = True
            BarraDeStatus "Alteração bem sucedida"
        End If
    Else
        Exit Sub
    End If
    
    TestaInferiorArray Elemento, ArrayPreçoFornecedor(), lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperiorArray Elemento, ArrayPreçoFornecedor(), lAllowEdit, lAllowDelete, lAllowConsult
    
    If ArrayPreçoTotal = 0 Then
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
        
    StatusBar = "Inclusão de Preço do Produto"
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
    
    Elemento = ArrayPreçoTotal
    
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
    
    If Elemento > ArrayPreçoTotal Then
        Elemento = ArrayPreçoTotal
        Exit Sub
    End If
    
    NavegaçãoInferior lAllowConsult
    TestaSuperiorArray Elemento, ArrayPreçoFornecedor(), lAllowEdit, lAllowDelete, lAllowConsult
    
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
    TestaInferiorArray Elemento, ArrayPreçoFornecedor(), lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Sub
Public Sub PosRecords()

End Sub
Private Sub GetRecords()
    On Error GoTo Erro
    
    lPula = True
    If Not lAllowConsult Then
        ZeraCampos
        DesativaCampos
        lPula = False
        Exit Sub
    End If
    txtFornecedor = ArrayPreçoFornecedor(Elemento)
    txtPreçoDeCusto = ArrayPreçoCusto(Elemento)
    txtPreçoDeVenda = ArrayPreçoVenda(Elemento)
    txtMargemDeLucro = ArrayPreçoLucro(Elemento)
    lPula = False
    If Not lAllowEdit Then
        DesativaCampos
    End If
    
    Exit Sub
    
Erro:
    GeraMensagemDeErro "Preços - GetRecords "
    Resume Next
End Sub
Private Function SetRecords()
    ArrayPreçoFornecedor(Elemento) = txtFornecedor
    ArrayPreçoCusto(Elemento) = txtPreçoDeCusto
    ArrayPreçoVenda(Elemento) = txtPreçoDeVenda
    ArrayPreçoLucro(Elemento) = txtMargemDeLucro
    Log gUsuário, "Preço do  Produto"
End Function
Private Sub SaveArray()
    Dim Cont%
    
    frmProduto.SizeArrayPreço ArrayPreçoTotal
    frmProduto.ArrayPreçoTotal = ArrayPreçoTotal
    
    For Cont = 1 To frmProduto.ArrayPreçoTotal
        frmProduto.SetArrayPreço "Fornecedor", ArrayPreçoFornecedor(Cont), Cont
        frmProduto.SetArrayPreço "Custo", ArrayPreçoCusto(Cont), Cont
        frmProduto.SetArrayPreço "Venda", ArrayPreçoVenda(Cont), Cont
        frmProduto.SetArrayPreço "Lucro", ArrayPreçoLucro(Cont), Cont
    Next
End Sub
Private Sub SizeArray(ByVal Tamanho As Integer)
    ASize Tamanho, ArrayPreçoFornecedor()
    ASize Tamanho, ArrayPreçoCusto()
    ASize Tamanho, ArrayPreçoVenda()
    ASize Tamanho, ArrayPreçoLucro()
End Sub
Private Sub ZeraCampos()
    txtFornecedor = Empty
    txtPreçoDeCusto = Empty
    txtPreçoDeVenda = Empty
    txtMargemDeLucro = Empty
End Sub
Private Sub cmdCancelar_Click()
    Cancelamento
End Sub
Private Sub cmdGravar_Click()
    Gravar
End Sub
Private Sub Form_Activate()
    TestaInferiorArray Elemento, ArrayPreçoFornecedor(), lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperiorArray Elemento, ArrayPreçoFornecedor(), lAllowEdit, lAllowDelete, lAllowConsult
    
    BarraDeStatus StatusBar
    
    If ArrayPreçoTotal = 0 Then
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
    frmPreços.SetFocus
End Sub
Private Sub Form_Load()
    lPula = False
    
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    
    lAllowInsert = Allow("PRODUTO", "I")
    lAllowEdit = Allow("PRODUTO", "P")
    lAllowDelete = Allow("PRODUTO", "E")
    lAllowConsult = Allow("PRODUTO", "C")
    
    FillArray
        
    BotãoIncluir lAllowInsert
    
    StatusBar = frmProduto.StatusBarAviso
 
    If ArrayPreçoTotal = 0 Then
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
        
    If ArrayPreçoTotal = 0 Or ArrayPreçoTotal = 1 Then
        NavegaçãoSuperior False
    Else
        NavegaçãoInferior lAllowConsult
    End If
    
    lInserir = False
    lAlterar = False
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
            If Not frmProduto.lAlterarArrayPreço Then
                frmProduto.lAlterarArrayPreço = True
                If Not frmProduto.lInserir Then
                    frmProduto.lAlterar = True
                End If
            End If
        ElseIf Confirmação = vbCancel Then
            Cancel = 1
        ElseIf Confirmação = vbNo Then
        End If
    End If
    
    Set frmPreços = Nothing
End Sub
Private Sub txtFornecedor_Change()
    FormatMask "99.999.999/9999-99", txtFornecedor
End Sub
Private Sub txtFornecedor_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBar = "Alteração de Preço do Produto"
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
Private Sub txtMargemDeLucro_Change()
    If Not lPula Then
        FormatMask "@K 999,99", txtMargemDeLucro
        lPula = True
        vscrMargemDeLucro.Value = Int(Val(txtMargemDeLucro))
        lPula = False
    End If
End Sub
Private Sub txtMargemDeLucro_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBar = "Alteração de Preço do Produto"
        BarraDeStatus StatusBar
    End If
End Sub
Private Sub txtMargemDeLucro_LostFocus()
    Dim Valor As Double, Aux$, AuxCusto$
    lPula = True
    FormatMask "@V ##0,00", txtMargemDeLucro
    AuxCusto = StrTran(txtPreçoDeCusto, ".", "")
    AuxCusto = StrTran(AuxCusto, ",", ".")
    Valor = Val(AuxCusto) * (1 + (Val(StrTran(txtMargemDeLucro, ",", ".")) / 100))
    Aux = Format(Valor, "##,###,##0.00")
    txtPreçoDeVenda = Aux
    FormatMask "@V ##.###.##0,00", txtPreçoDeVenda
    lPula = False
End Sub
Private Sub txtPreçoDeCusto_Change()
    If Not lPula Then
        FormatMask "@K 99.999.999,99", txtPreçoDeCusto
    End If
End Sub
Private Sub txtPreçoDeCusto_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBar = "Alteração de Preço do Produto"
        BarraDeStatus StatusBar
    End If
End Sub
Private Sub txtPreçoDeCusto_LostFocus()
    Dim Valor As Double, AuxVenda$, AuxCusto$
    
    If lPula Then
        Exit Sub
    End If
    
    lPula = True
    FormatMask "@V ##.###.##0,00", txtPreçoDeCusto
    If Val(StrTran(txtPreçoDeCusto, ",", ".")) > 0 And Val(StrTran(txtPreçoDeVenda, ",", ".")) > 0 Then
        AuxVenda = StrTran(txtPreçoDeVenda, ".", "")
        AuxVenda = StrTran(AuxVenda, ",", ".")
        AuxCusto = StrTran(txtPreçoDeCusto, ".", "")
        AuxCusto = StrTran(AuxCusto, ",", ".")
        '((Val(StrTran(txtPreçoDeVenda, ",", ".")) / Val(StrTran(txtPreçoDeCusto, ",", "."))) - 1) * 100
        Valor = ((Val(AuxVenda) / Val(AuxCusto)) - 1) * 100
        txtMargemDeLucro = StrTran(Format(Valor, "##0.00"), ".", ",")
        txtMargemDeLucro_LostFocus
        lPula = True
        If Int(Valor) >= 0 Then
            vscrMargemDeLucro.Value = Int(Valor)
        Else
            MsgBox "Valor inválido! " & vbCr & "Verifique se os valores estão corretos.", vbOKOnly, "Erro"
            vscrMargemDeLucro.Value = 0
        End If
    End If
    lPula = False
End Sub
Private Sub txtPreçoDeVenda_Change()
    If Not lPula Then
        lPula = True
        FormatMask "@K 99.999.999,99", txtPreçoDeVenda
        lPula = False
    End If
End Sub
Private Sub txtPreçoDeVenda_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBar = "Alteração de Preço do Produto"
        BarraDeStatus StatusBar
    End If
End Sub
Private Sub txtPreçoDeVenda_LostFocus()
    Dim Valor As Double, AuxVenda$, AuxCusto$
    
    If lPula Then
        Exit Sub
    End If
    
    lPula = True
    FormatMask "@V ##.###.##0,00", txtPreçoDeVenda
    If Val(StrTran(txtPreçoDeCusto, ",", ".")) > 0 And Val(StrTran(txtPreçoDeVenda, ",", ".")) > 0 Then
        AuxVenda = StrTran(txtPreçoDeVenda, ".", "")
        AuxVenda = StrTran(AuxVenda, ",", ".")
        AuxCusto = StrTran(txtPreçoDeCusto, ".", "")
        AuxCusto = StrTran(AuxCusto, ",", ".")
        '((Val(StrTran(txtPreçoDeVenda, ",", ".")) / Val(StrTran(txtPreçoDeCusto, ",", "."))) - 1) * 100
        Valor = ((Val(AuxVenda) / Val(AuxCusto)) - 1) * 100
        txtMargemDeLucro = StrTran(Format(Valor, "##0.00"), ".", ",")
        txtMargemDeLucro_LostFocus
        lPula = True
        If Int(Valor) >= 0 Then
            vscrMargemDeLucro.Value = Int(Valor)
        Else
            MsgBox "Valor inválido! " & vbCr & "Verifique se os valores estão corretos.", vbOKOnly, "Erro"
            vscrMargemDeLucro.Value = 0
        End If
    End If
    lPula = False
End Sub
Private Sub vscrMargemDeLucro_Change()
    Dim Valor As Double, Aux$, AuxCusto$
    
    If Not lPula Then
        lPula = True
        txtMargemDeLucro = Trim(Str(vscrMargemDeLucro.Value)) + "00"
        FormatMask "@N ##0,00", txtMargemDeLucro
        AuxCusto = StrTran(txtPreçoDeCusto, ".", "")
        AuxCusto = StrTran(AuxCusto, ",", ".")
        Valor = Val(AuxCusto) * (1 + (Val(StrTran(txtMargemDeLucro, ",", ".")) / 100))
        Aux = Format(Valor, "##,###,##0.00")
        txtPreçoDeVenda = Aux
        FormatMask "@V ##.###.##0,00", txtPreçoDeVenda
        lPula = False
        If Not lInserir Then
            lAlterar = True
            StatusBar = "Alteração de Preço do Produto"
            BarraDeStatus StatusBar
        End If
    End If
End Sub
