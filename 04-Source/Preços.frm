VERSION 5.00
Begin VB.Form frmPre�os 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pre�os"
   ClientHeight    =   2460
   ClientLeft      =   1140
   ClientTop       =   1515
   ClientWidth     =   4800
   Icon            =   "Pre�os.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2460
   ScaleWidth      =   4800
   Begin VB.Frame frPre�os 
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
      Begin VB.TextBox txtPre�oDeVenda 
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
      Begin VB.TextBox txtPre�oDeCusto 
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
      Begin VB.Label lblPre�oDeVenda 
         Caption         =   "Pre�o de Venda"
         Height          =   195
         Left            =   150
         TabIndex        =   9
         Top             =   1110
         Width           =   1215
      End
      Begin VB.Label lblPre�oDeCusto 
         Caption         =   "Pre�o de Custo"
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
Attribute VB_Name = "frmPre�os"
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
Public lAltera��o As Boolean

Dim ArrayPre�oTotal%
Dim ArrayPre�oFornecedor() As Variant
Dim ArrayPre�oCusto() As Variant
Dim ArrayPre�oVenda() As Variant
Dim ArrayPre�oLucro() As Variant

Public lAtualizar As Boolean
Private Sub AtivaCampos()
    frPre�os.Enabled = True
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
    
    If ArrayPre�oTotal = 0 Then
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
    
    TestaInferiorArray Elemento, ArrayPre�oFornecedor(), lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperiorArray Elemento, ArrayPre�oFornecedor(), lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Function
Private Sub ClearArray()
    ArrayPre�oTotal = 0
    ReDim ArrayPre�oFornecedor(1 To 1)
    ReDim ArrayPre�oCusto(1 To 1)
    ReDim ArrayPre�oVenda(1 To 1)
    ReDim ArrayPre�oLucro(1 To 1)
End Sub
Private Sub DesativaCampos()
    frPre�os.Enabled = False
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
    
    BarraDeStatus "Exclus�o de Pre�o do Produto"

    Adel Elemento, ArrayPre�oFornecedor()
    Adel Elemento, ArrayPre�oCusto()
    Adel Elemento, ArrayPre�oVenda()
    Adel Elemento, ArrayPre�oLucro()
    
    ArrayPre�oTotal = ArrayPre�oTotal - 1
    
    If Elemento > ArrayPre�oTotal Then
        Elemento = ArrayPre�oTotal
    End If
    
    lAltera��o = True
    
    If ArrayPre�oTotal = 0 Then
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
    
    SizeArray ArrayPre�oTotal
    
    GetRecords
    
    TestaInferiorArray Elemento, ArrayPre�oFornecedor(), lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperiorArray Elemento, ArrayPre�oFornecedor(), lAllowEdit, lAllowDelete, lAllowConsult
End Sub
Private Sub FillArray()
    Dim Cont%
    
    SizeArray frmProduto.ArrayPre�oTotal
    ArrayPre�oTotal = frmProduto.ArrayPre�oTotal
    
    For Cont = 1 To frmProduto.ArrayPre�oTotal
        ArrayPre�oFornecedor(Cont) = frmProduto.GetArrayPre�o("Fornecedor", Cont)
        ArrayPre�oCusto(Cont) = frmProduto.GetArrayPre�o("Custo", Cont)
        ArrayPre�oVenda(Cont) = frmProduto.GetArrayPre�o("Venda", Cont)
        ArrayPre�oLucro(Cont) = frmProduto.GetArrayPre�o("Lucro", Cont)
    Next
End Sub
Public Sub Gravar()
    If lInserir Then
        ArrayPre�oTotal = ArrayPre�oTotal + 1
        Elemento = ArrayPre�oTotal
        SizeArray ArrayPre�oTotal
        SetRecords
        PosRecords
        BarraDeStatus "Inclus�o bem sucedida"
        lInserir = False
        lAltera��o = True
    ElseIf lAlterar Then
        If ArrayPre�oTotal > 0 Then
            SetRecords
            PosRecords
            lAlterar = False
            lAltera��o = True
            BarraDeStatus "Altera��o bem sucedida"
        End If
    Else
        Exit Sub
    End If
    
    TestaInferiorArray Elemento, ArrayPre�oFornecedor(), lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperiorArray Elemento, ArrayPre�oFornecedor(), lAllowEdit, lAllowDelete, lAllowConsult
    
    If ArrayPre�oTotal = 0 Then
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
        
    StatusBar = "Inclus�o de Pre�o do Produto"
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
    
    Elemento = ArrayPre�oTotal
    
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
    
    If Elemento > ArrayPre�oTotal Then
        Elemento = ArrayPre�oTotal
        Exit Sub
    End If
    
    Navega��oInferior lAllowConsult
    TestaSuperiorArray Elemento, ArrayPre�oFornecedor(), lAllowEdit, lAllowDelete, lAllowConsult
    
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
    TestaInferiorArray Elemento, ArrayPre�oFornecedor(), lAllowEdit, lAllowDelete, lAllowConsult
    
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
    txtFornecedor = ArrayPre�oFornecedor(Elemento)
    txtPre�oDeCusto = ArrayPre�oCusto(Elemento)
    txtPre�oDeVenda = ArrayPre�oVenda(Elemento)
    txtMargemDeLucro = ArrayPre�oLucro(Elemento)
    lPula = False
    If Not lAllowEdit Then
        DesativaCampos
    End If
    
    Exit Sub
    
Erro:
    GeraMensagemDeErro "Pre�os - GetRecords "
    Resume Next
End Sub
Private Function SetRecords()
    ArrayPre�oFornecedor(Elemento) = txtFornecedor
    ArrayPre�oCusto(Elemento) = txtPre�oDeCusto
    ArrayPre�oVenda(Elemento) = txtPre�oDeVenda
    ArrayPre�oLucro(Elemento) = txtMargemDeLucro
    Log gUsu�rio, "Pre�o do  Produto"
End Function
Private Sub SaveArray()
    Dim Cont%
    
    frmProduto.SizeArrayPre�o ArrayPre�oTotal
    frmProduto.ArrayPre�oTotal = ArrayPre�oTotal
    
    For Cont = 1 To frmProduto.ArrayPre�oTotal
        frmProduto.SetArrayPre�o "Fornecedor", ArrayPre�oFornecedor(Cont), Cont
        frmProduto.SetArrayPre�o "Custo", ArrayPre�oCusto(Cont), Cont
        frmProduto.SetArrayPre�o "Venda", ArrayPre�oVenda(Cont), Cont
        frmProduto.SetArrayPre�o "Lucro", ArrayPre�oLucro(Cont), Cont
    Next
End Sub
Private Sub SizeArray(ByVal Tamanho As Integer)
    ASize Tamanho, ArrayPre�oFornecedor()
    ASize Tamanho, ArrayPre�oCusto()
    ASize Tamanho, ArrayPre�oVenda()
    ASize Tamanho, ArrayPre�oLucro()
End Sub
Private Sub ZeraCampos()
    txtFornecedor = Empty
    txtPre�oDeCusto = Empty
    txtPre�oDeVenda = Empty
    txtMargemDeLucro = Empty
End Sub
Private Sub cmdCancelar_Click()
    Cancelamento
End Sub
Private Sub cmdGravar_Click()
    Gravar
End Sub
Private Sub Form_Activate()
    TestaInferiorArray Elemento, ArrayPre�oFornecedor(), lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperiorArray Elemento, ArrayPre�oFornecedor(), lAllowEdit, lAllowDelete, lAllowConsult
    
    BarraDeStatus StatusBar
    
    If ArrayPre�oTotal = 0 Then
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
    frmPre�os.SetFocus
End Sub
Private Sub Form_Load()
    lPula = False
    
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    
    lAllowInsert = Allow("PRODUTO", "I")
    lAllowEdit = Allow("PRODUTO", "P")
    lAllowDelete = Allow("PRODUTO", "E")
    lAllowConsult = Allow("PRODUTO", "C")
    
    FillArray
        
    Bot�oIncluir lAllowInsert
    
    StatusBar = frmProduto.StatusBarAviso
 
    If ArrayPre�oTotal = 0 Then
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
        
    If ArrayPre�oTotal = 0 Or ArrayPre�oTotal = 1 Then
        Navega��oSuperior False
    Else
        Navega��oInferior lAllowConsult
    End If
    
    lInserir = False
    lAlterar = False
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
            If Not frmProduto.lAlterarArrayPre�o Then
                frmProduto.lAlterarArrayPre�o = True
                If Not frmProduto.lInserir Then
                    frmProduto.lAlterar = True
                End If
            End If
        ElseIf Confirma��o = vbCancel Then
            Cancel = 1
        ElseIf Confirma��o = vbNo Then
        End If
    End If
    
    Set frmPre�os = Nothing
End Sub
Private Sub txtFornecedor_Change()
    FormatMask "99.999.999/9999-99", txtFornecedor
End Sub
Private Sub txtFornecedor_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBar = "Altera��o de Pre�o do Produto"
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
        StatusBar = "Altera��o de Pre�o do Produto"
        BarraDeStatus StatusBar
    End If
End Sub
Private Sub txtMargemDeLucro_LostFocus()
    Dim Valor As Double, Aux$, AuxCusto$
    lPula = True
    FormatMask "@V ##0,00", txtMargemDeLucro
    AuxCusto = StrTran(txtPre�oDeCusto, ".", "")
    AuxCusto = StrTran(AuxCusto, ",", ".")
    Valor = Val(AuxCusto) * (1 + (Val(StrTran(txtMargemDeLucro, ",", ".")) / 100))
    Aux = Format(Valor, "##,###,##0.00")
    txtPre�oDeVenda = Aux
    FormatMask "@V ##.###.##0,00", txtPre�oDeVenda
    lPula = False
End Sub
Private Sub txtPre�oDeCusto_Change()
    If Not lPula Then
        FormatMask "@K 99.999.999,99", txtPre�oDeCusto
    End If
End Sub
Private Sub txtPre�oDeCusto_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBar = "Altera��o de Pre�o do Produto"
        BarraDeStatus StatusBar
    End If
End Sub
Private Sub txtPre�oDeCusto_LostFocus()
    Dim Valor As Double, AuxVenda$, AuxCusto$
    
    If lPula Then
        Exit Sub
    End If
    
    lPula = True
    FormatMask "@V ##.###.##0,00", txtPre�oDeCusto
    If Val(StrTran(txtPre�oDeCusto, ",", ".")) > 0 And Val(StrTran(txtPre�oDeVenda, ",", ".")) > 0 Then
        AuxVenda = StrTran(txtPre�oDeVenda, ".", "")
        AuxVenda = StrTran(AuxVenda, ",", ".")
        AuxCusto = StrTran(txtPre�oDeCusto, ".", "")
        AuxCusto = StrTran(AuxCusto, ",", ".")
        '((Val(StrTran(txtPre�oDeVenda, ",", ".")) / Val(StrTran(txtPre�oDeCusto, ",", "."))) - 1) * 100
        Valor = ((Val(AuxVenda) / Val(AuxCusto)) - 1) * 100
        txtMargemDeLucro = StrTran(Format(Valor, "##0.00"), ".", ",")
        txtMargemDeLucro_LostFocus
        lPula = True
        If Int(Valor) >= 0 Then
            vscrMargemDeLucro.Value = Int(Valor)
        Else
            MsgBox "Valor inv�lido! " & vbCr & "Verifique se os valores est�o corretos.", vbOKOnly, "Erro"
            vscrMargemDeLucro.Value = 0
        End If
    End If
    lPula = False
End Sub
Private Sub txtPre�oDeVenda_Change()
    If Not lPula Then
        lPula = True
        FormatMask "@K 99.999.999,99", txtPre�oDeVenda
        lPula = False
    End If
End Sub
Private Sub txtPre�oDeVenda_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBar = "Altera��o de Pre�o do Produto"
        BarraDeStatus StatusBar
    End If
End Sub
Private Sub txtPre�oDeVenda_LostFocus()
    Dim Valor As Double, AuxVenda$, AuxCusto$
    
    If lPula Then
        Exit Sub
    End If
    
    lPula = True
    FormatMask "@V ##.###.##0,00", txtPre�oDeVenda
    If Val(StrTran(txtPre�oDeCusto, ",", ".")) > 0 And Val(StrTran(txtPre�oDeVenda, ",", ".")) > 0 Then
        AuxVenda = StrTran(txtPre�oDeVenda, ".", "")
        AuxVenda = StrTran(AuxVenda, ",", ".")
        AuxCusto = StrTran(txtPre�oDeCusto, ".", "")
        AuxCusto = StrTran(AuxCusto, ",", ".")
        '((Val(StrTran(txtPre�oDeVenda, ",", ".")) / Val(StrTran(txtPre�oDeCusto, ",", "."))) - 1) * 100
        Valor = ((Val(AuxVenda) / Val(AuxCusto)) - 1) * 100
        txtMargemDeLucro = StrTran(Format(Valor, "##0.00"), ".", ",")
        txtMargemDeLucro_LostFocus
        lPula = True
        If Int(Valor) >= 0 Then
            vscrMargemDeLucro.Value = Int(Valor)
        Else
            MsgBox "Valor inv�lido! " & vbCr & "Verifique se os valores est�o corretos.", vbOKOnly, "Erro"
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
        AuxCusto = StrTran(txtPre�oDeCusto, ".", "")
        AuxCusto = StrTran(AuxCusto, ",", ".")
        Valor = Val(AuxCusto) * (1 + (Val(StrTran(txtMargemDeLucro, ",", ".")) / 100))
        Aux = Format(Valor, "##,###,##0.00")
        txtPre�oDeVenda = Aux
        FormatMask "@V ##.###.##0,00", txtPre�oDeVenda
        lPula = False
        If Not lInserir Then
            lAlterar = True
            StatusBar = "Altera��o de Pre�o do Produto"
            BarraDeStatus StatusBar
        End If
    End If
End Sub
