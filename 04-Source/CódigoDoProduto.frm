VERSION 5.00
Begin VB.Form frmC�digoDoProduto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "C�digo do Produto"
   ClientHeight    =   1680
   ClientLeft      =   2430
   ClientTop       =   1530
   ClientWidth     =   4800
   ForeColor       =   &H8000000D&
   Icon            =   "C�digoDoProduto.frx":0000
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
   Begin VB.Frame frC�digoDoProduto 
      Height          =   1230
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   4785
      Begin VB.TextBox txtC�digo 
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
      Begin VB.Label lblC�digo 
         Caption         =   "C�digo"
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
Attribute VB_Name = "frmC�digoDoProduto"
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
Public lAltera��o As Boolean

Dim lAllowInsert  As Boolean
Dim lAllowEdit    As Boolean
Dim lAllowDelete  As Boolean
Dim lAllowConsult As Boolean

Dim ArrayProdutoTotal%
Dim ArrayProdutoC�digo() As Variant
Dim ArrayProdutoFornecedor() As Variant

Public lAtualizar As Boolean
Private Sub AtivaCampos()
    frC�digoDoProduto.Enabled = True
    Bot�oGravar (lInserir Or lAllowEdit)
    cmdCancelar.Enabled = (lInserir Or lAllowEdit)
    cmdGravar.Enabled = (lInserir Or lAllowEdit)
End Sub
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
        BarraDeStatus "Inclus�o cancelada"
    End If
    If lAlterar Then
        BarraDeStatus "Altera��o cancelada"
    End If
    
    Bot�oIncluir lAllowInsert
    
    If ArrayProdutoTotal = 0 Then
        Navega��oInferior False
        Navega��oSuperior False
        Bot�oGravar False
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
    
    TestaInferiorArray Elemento, ArrayProdutoC�digo(), lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperiorArray Elemento, ArrayProdutoC�digo(), lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Function
Private Sub ClearArray()
    ArrayProdutoTotal = 0
    ReDim ArrayProdutoC�digo(1 To 1)
    ReDim ArrayProdutoFornecedor(1 To 1)
End Sub
Private Sub DesativaCampos()
    frC�digoDoProduto.Enabled = False
    cmdGravar.Enabled = False
    cmdCancelar.Enabled = False
    Bot�oGravar False
End Sub
Public Sub Excluir()
    Dim Confirma��o As Integer, Msg1$, Msg2$
    Dim SQL As String
    
    If lAlterar Then
       If Not Cancelamento Then
            Exit Sub
        End If
    End If
    
    Msg1 = "Voc� est� preste a apagar um registro !"
    Msg2 = "Tem certeza?"
    Msg2 = String(((Len(Msg1) - Len(Msg2)) / 2), " ") + Msg2
    Confirma��o = MsgBox(Msg1 + vbCr + Msg2, vbYesNo + vbQuestion + vbDefaultButton2, "Confirma��o")
    
    If Confirma��o = vbNo Then
        Exit Sub
    End If
    
    BarraDeStatus "Exclus�o de C�digo do Produto"

    Adel Elemento, ArrayProdutoC�digo()
    Adel Elemento, ArrayProdutoFornecedor()
    
    ArrayProdutoTotal = ArrayProdutoTotal - 1
    
    If Elemento > ArrayProdutoTotal Then
        Elemento = ArrayProdutoTotal
    End If
    
    lAltera��o = True
    
    If ArrayProdutoTotal = 0 Then
        ClearArray
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
    
    SizeArray ArrayProdutoTotal
    
    Log gUsu�rio, "Exclus�o - C�digo do Produto: " & txtC�digo
    
    GetRecords
    
    TestaInferiorArray Elemento, ArrayProdutoC�digo(), lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperiorArray Elemento, ArrayProdutoC�digo(), lAllowEdit, lAllowDelete, lAllowConsult
End Sub
Private Sub FillArray()
    Dim Cont%
    
    SizeArray frmProduto.ArrayProdutoTotal
    ArrayProdutoTotal = frmProduto.ArrayProdutoTotal
    
    For Cont = 1 To frmProduto.ArrayProdutoTotal
        ArrayProdutoC�digo(Cont) = frmProduto.GetArrayProduto("C�digo", Cont)
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
        BarraDeStatus "Inclus�o bem sucedida"
        lInserir = False
        lAltera��o = True
    ElseIf lAlterar Then
        If ArrayProdutoTotal > 0 Then
            SetRecords
            PosRecords
            lAlterar = False
            lAltera��o = True
            BarraDeStatus "Altera��o bem sucedida"
        End If
    Else
        Exit Sub
    End If
    
    TestaInferiorArray Elemento, ArrayProdutoC�digo(), lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperiorArray Elemento, ArrayProdutoC�digo(), lAllowEdit, lAllowDelete, lAllowConsult
    
    If ArrayProdutoTotal = 0 Then
        If Not lInserir And Not lAlterar Then
            Bot�oExcluir False
            Bot�oGravar False
            cmdGravar.Enabled = False
            cmdCancelar.Enabled = False
        End If
    Else
        Bot�oExcluir lAllowDelete
    End If
    
    Bot�oIncluir lAllowInsert
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
        
    StatusBar = "Inclus�o de Lote do Produto"
    BarraDeStatus StatusBar
    
    ZeraCampos
    AtivaCampos
    
    Bot�oGravar (lInserir Or lAllowEdit)
    Bot�oIncluir False
    cmdGravar.Enabled = (lInserir Or lAllowEdit)
    cmdCancelar.Enabled = (lInserir Or lAllowEdit)
    
    Navega��oInferior False
    Navega��oSuperior False
    
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
    
    Navega��oInferior False
    Navega��oSuperior lAllowConsult
    
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
    
    Navega��oInferior lAllowConsult
    Navega��oSuperior False
    
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
    
    Navega��oInferior lAllowConsult
    TestaSuperiorArray Elemento, ArrayProdutoC�digo(), lAllowEdit, lAllowDelete, lAllowConsult
    
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
    
    Navega��oSuperior lAllowConsult
    TestaInferiorArray Elemento, ArrayProdutoC�digo(), lAllowEdit, lAllowDelete, lAllowConsult
    
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
    txtC�digo = ArrayProdutoC�digo(Elemento)
    txtFornecedor = ArrayProdutoFornecedor(Elemento)
    If Not lAllowEdit Then
        DesativaCampos
    End If
    
    Exit Sub
    
Erro:
    GeraMensagemDeErro "C�digo Do Produto - GetRecords "
    Resume Next
End Sub
Private Function SetRecords()
    ArrayProdutoC�digo(Elemento) = txtC�digo
    ArrayProdutoFornecedor(Elemento) = txtFornecedor
    Log gUsu�rio, "C�digo do Produto " & txtC�digo
End Function
Private Sub SaveArray()
    Dim Cont%
    
    frmProduto.SizeArrayProduto ArrayProdutoTotal
    frmProduto.ArrayProdutoTotal = ArrayProdutoTotal
    
    For Cont = 1 To frmProduto.ArrayProdutoTotal
        frmProduto.SetArrayProduto "C�digo", ArrayProdutoC�digo(Cont), Cont
        frmProduto.SetArrayProduto "Fornecedor", ArrayProdutoFornecedor(Cont), Cont
    Next
End Sub
Private Sub SizeArray(ByVal Tamanho As Integer)
    ASize Tamanho, ArrayProdutoC�digo()
    ASize Tamanho, ArrayProdutoFornecedor()
End Sub
Private Sub ZeraCampos()
    txtC�digo = Empty
    txtFornecedor = Empty
End Sub
Private Sub cmdCancelar_Click()
    Cancelamento
End Sub
Private Sub cmdGravar_Click()
    Gravar
End Sub
Private Sub Form_Activate()
    TestaInferiorArray Elemento, ArrayProdutoC�digo(), lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperiorArray Elemento, ArrayProdutoC�digo(), lAllowEdit, lAllowDelete, lAllowConsult
    
    BarraDeStatus StatusBar
    
    If ArrayProdutoTotal = 0 Then
        Bot�oGravar False
        cmdGravar.Enabled = False
        cmdCancelar.Enabled = False
    Else
        Bot�oGravar (lInserir Or lAllowEdit)
        cmdGravar.Enabled = (lInserir Or lAllowEdit)
        cmdCancelar.Enabled = (lInserir Or lAllowEdit)
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
        Bot�oIncluir lAllowInsert
    Else
        Bot�oIncluir lAllowInsert
    End If
    
    If lAtualizar Then
        Bot�oAtualizar True
    Else
        Bot�oAtualizar False
    End If
End Sub
Private Sub Form_Deactivate()
    Beep
    frmC�digoDoProduto.SetFocus
End Sub
Private Sub Form_Load()
    lPula = False
    
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    
    lAllowInsert = Allow("PRODUTO", "I")
    lAllowEdit = Allow("PRODUTO", "G")
    lAllowDelete = Allow("PRODUTO", "E")
    lAllowConsult = Allow("PRODUTO", "C")
    
    FillArray
        
    Bot�oIncluir lAllowInsert
    
    StatusBar = frmProduto.StatusBarAviso
 
    If ArrayProdutoTotal = 0 Then
        DesativaCampos
        Bot�oExcluir False
        Bot�oGravar False
    Else
        Elemento = 1
        AtivaCampos
        Bot�oExcluir lAllowDelete
        Bot�oGravar (lInserir Or lAllowEdit)
        GetRecords
    End If
    
    Navega��oInferior False
        
    If ArrayProdutoTotal = 0 Or ArrayProdutoTotal = 1 Then
        Navega��oSuperior False
    Else
        Navega��oInferior lAllowConsult
    End If
        
    lInserir = False
    lAlterar = False
    Exit Sub
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim Confirma��o
    
    If lInserir Then
        MsgBox "Voc� est� em uma inclus�o!", vbExclamation, Caption
        BarraDeStatus "Finalize a inclus�o"
        Cancel = 1
        SetaFocus Me
        mdiGeal.Mostrar
        Exit Sub
    End If
    If lAlterar Then
        MsgBox "Voc� est� em uma altera��o!", vbExclamation, Caption
        BarraDeStatus "Finalize a altera��o"
        Cancel = 1
        SetaFocus Me
        mdiGeal.Mostrar
        Exit Sub
    End If
    
    If lAltera��o Then
        Confirma��o = MsgBox("       Aceita as altera��es ?", vbQuestion + vbDefaultButton1 + vbYesNoCancel, "Confirma��o")
        
        If Confirma��o = vbYes Then
            SaveArray
            If Not frmProduto.lAlterarArrayProduto Then
                frmProduto.lAlterarArrayProduto = True
                If Not frmProduto.lInserir Then
                    frmProduto.lAlterar = True
                End If
            End If
        ElseIf Confirma��o = vbCancel Then
            Cancel = 1
        ElseIf Confirma��o = vbNo Then
        End If
    End If
    
    Set frmC�digoDoProduto = Nothing
End Sub
Private Sub txtC�digo_Change()
    FormatMask "@!S13", txtC�digo
End Sub
Private Sub txtC�digo_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBar = "Altera��o de C�digo do Produto"
        BarraDeStatus StatusBar
    End If
End Sub
Private Sub txtFornecedor_Change()
    FormatMask "99.999.999/9999-99", txtFornecedor
End Sub
Private Sub txtFornecedor_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBar = "Altera��o de C�digo do Produto"
        BarraDeStatus StatusBar
    End If
End Sub
Private Sub txtFornecedor_LostFocus()
    If Not IsCorrectFornecedor(txtFornecedor) Then
        MsgBox "Fornecedor n�o cadastrado!", vbInformation, "Aviso"
        Set frmEncontrar.DBBancoDeDados = DBCadastro
        frmEncontrar.Mensagem = "Nenhum fornecedor foi selecionado!"
        frmEncontrar.BancoDeDados = "CADASTRO"
        frmEncontrar.Tabela = "FORNECEDOR"
        frmEncontrar.Indice = "2"
        frmEncontrar.CampoChave = "CGC - CPF"
        frmEncontrar.CampoPreencheLista = "RAZ�O SOCIAL"
        frmEncontrar.Show vbModal
        txtFornecedor = frmEncontrar.Chave
    End If
End Sub
