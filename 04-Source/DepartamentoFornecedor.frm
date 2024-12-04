VERSION 5.00
Begin VB.Form frmDepartamentoFornecedor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Departamento do Fornecedor"
   ClientHeight    =   3945
   ClientLeft      =   1560
   ClientTop       =   1515
   ClientWidth     =   6540
   Icon            =   "DepartamentoFornecedor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3945
   ScaleWidth      =   6540
   Begin VB.CommandButton cmdGravar 
      Caption         =   "&Gravar"
      Default         =   -1  'True
      Height          =   345
      Left            =   3960
      TabIndex        =   9
      Top             =   3570
      Width           =   1245
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   5280
      TabIndex        =   10
      Top             =   3570
      Width           =   1245
   End
   Begin VB.Frame frDadosCadastrais 
      Height          =   3510
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   6525
      Begin VB.TextBox txtEMail 
         Height          =   285
         Left            =   1200
         TabIndex        =   8
         Top             =   3030
         Width           =   5235
      End
      Begin VB.TextBox txtFax 
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
         Left            =   4380
         TabIndex        =   7
         Text            =   "(    )    -    "
         Top             =   2580
         Width           =   1900
      End
      Begin VB.TextBox txtTelefone 
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
         TabIndex        =   6
         Text            =   "(    )    -    "
         Top             =   2550
         Width           =   1900
      End
      Begin VB.TextBox txtCep 
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
         Left            =   4380
         TabIndex        =   5
         Text            =   "  .   -   "
         Top             =   2100
         Width           =   1300
      End
      Begin VB.TextBox txtUF 
         Height          =   285
         Left            =   1200
         TabIndex        =   4
         Top             =   2130
         Width           =   435
      End
      Begin VB.TextBox txtCidade 
         Height          =   285
         Left            =   1200
         TabIndex        =   3
         Top             =   1680
         Width           =   5235
      End
      Begin VB.TextBox txtBairro 
         Height          =   285
         Left            =   1200
         TabIndex        =   2
         Top             =   1230
         Width           =   5235
      End
      Begin VB.TextBox txtEndereço 
         Height          =   285
         Left            =   1200
         TabIndex        =   1
         Top             =   780
         Width           =   5235
      End
      Begin VB.TextBox txtContato 
         Height          =   285
         Left            =   1200
         TabIndex        =   0
         Top             =   330
         Width           =   5235
      End
      Begin VB.Label lblEMail 
         Caption         =   "E-Mail"
         Height          =   240
         Left            =   150
         TabIndex        =   20
         Top             =   3060
         Width           =   510
      End
      Begin VB.Label lblFax 
         Caption         =   "Fax"
         Height          =   195
         Left            =   3930
         TabIndex        =   19
         Top             =   2610
         Width           =   345
      End
      Begin VB.Label lblTelefone 
         Caption         =   "Telefone"
         Height          =   210
         Left            =   150
         TabIndex        =   18
         Top             =   2610
         Width           =   660
      End
      Begin VB.Label lblCep 
         Caption         =   "CEP"
         Height          =   195
         Left            =   3930
         TabIndex        =   17
         Top             =   2160
         Width           =   375
      End
      Begin VB.Label lblUF 
         Caption         =   "U. F."
         Height          =   225
         Left            =   150
         TabIndex        =   16
         Top             =   2160
         Width           =   405
      End
      Begin VB.Label lblCidade 
         Caption         =   "Cidade"
         Height          =   225
         Left            =   150
         TabIndex        =   15
         Top             =   1710
         Width           =   945
      End
      Begin VB.Label lblBairro 
         Caption         =   "Bairro"
         Height          =   225
         Left            =   150
         TabIndex        =   14
         Top             =   1260
         Width           =   945
      End
      Begin VB.Label lblEndereço 
         Caption         =   "Endereço"
         Height          =   195
         Left            =   150
         TabIndex        =   13
         Top             =   810
         Width           =   975
      End
      Begin VB.Label lblContato 
         Caption         =   "Contato"
         Height          =   195
         Left            =   150
         TabIndex        =   12
         Top             =   360
         Width           =   1065
      End
   End
End
Attribute VB_Name = "frmDepartamentoFornecedor"
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

Dim ArrayTotal%
Dim ArrayContato() As Variant
Dim ArrayEndereço() As Variant
Dim ArrayBairro() As Variant
Dim ArrayCidade() As Variant
Dim ArrayUF() As Variant
Dim ArrayCEP() As Variant
Dim ArrayTelefone() As Variant
Dim ArrayFax() As Variant
Dim ArrayEMail() As Variant

Public lAtualizar As Boolean
Private Sub AtivaCampos()
    frDadosCadastrais.Enabled = True
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
    
    If ArrayTotal = 0 Then
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
    
    TestaInferiorArray Elemento, ArrayContato(), lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperiorArray Elemento, ArrayContato(), lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Function
Private Sub ClearArray()
    ArrayTotal = 0
    ReDim ArrayContato(1 To 1)
    ReDim ArrayEndereço(1 To 1)
    ReDim ArrayBairro(1 To 1)
    ReDim ArrayCidade(1 To 1)
    ReDim ArrayUF(1 To 1)
    ReDim ArrayCEP(1 To 1)
    ReDim ArrayTelefone(1 To 1)
    ReDim ArrayFax(1 To 1)
    ReDim ArrayEMail(1 To 1)
End Sub
Private Sub DesativaCampos()
    frDadosCadastrais.Enabled = False
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
    
    BarraDeStatus "Exclusão de Departamento do Fornecedor"

    Adel Elemento, ArrayContato()
    Adel Elemento, ArrayEndereço()
    Adel Elemento, ArrayBairro()
    Adel Elemento, ArrayCidade()
    Adel Elemento, ArrayUF()
    Adel Elemento, ArrayCEP()
    Adel Elemento, ArrayTelefone()
    Adel Elemento, ArrayFax()
    Adel Elemento, ArrayEMail()
    
    ArrayTotal = ArrayTotal - 1
    
    If Elemento > ArrayTotal Then
        Elemento = ArrayTotal
    End If
    
    lAlteração = True
    
    If ArrayTotal = 0 Then
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
    
    SizeArray ArrayTotal
    
    GetRecords
    
    TestaInferiorArray Elemento, ArrayContato(), lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperiorArray Elemento, ArrayContato(), lAllowEdit, lAllowDelete, lAllowConsult
End Sub
Private Sub FillArray()
    Dim Cont%
    
    SizeArray frmFornecedor.ArrayTotal
    ArrayTotal = frmFornecedor.ArrayTotal
    
    For Cont = 1 To frmFornecedor.ArrayTotal
        ArrayContato(Cont) = frmFornecedor.GetArray("Contato", Cont)
        ArrayEndereço(Cont) = frmFornecedor.GetArray("Endereço", Cont)
        ArrayBairro(Cont) = frmFornecedor.GetArray("Bairro", Cont)
        ArrayCidade(Cont) = frmFornecedor.GetArray("Cidade", Cont)
        ArrayUF(Cont) = frmFornecedor.GetArray("UF", Cont)
        ArrayCEP(Cont) = frmFornecedor.GetArray("CEP", Cont)
        ArrayTelefone(Cont) = frmFornecedor.GetArray("Telefone", Cont)
        ArrayFax(Cont) = frmFornecedor.GetArray("Fax", Cont)
        ArrayEMail(Cont) = frmFornecedor.GetArray("EMail", Cont)
    Next
End Sub
Public Sub Gravar()
    If lInserir Then
        ArrayTotal = ArrayTotal + 1
        Elemento = ArrayTotal
        SizeArray ArrayTotal
        SetRecords
        PosRecords
        BarraDeStatus "Inclusão bem sucedida"
        lInserir = False
        lAlteração = True
    ElseIf lAlterar Then
        If ArrayTotal > 0 Then
            SetRecords
            PosRecords
            lAlterar = False
            lAlteração = True
            BarraDeStatus "Alteração bem sucedida"
        End If
    Else
        Exit Sub
    End If
    
    TestaInferiorArray Elemento, ArrayContato(), lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperiorArray Elemento, ArrayContato(), lAllowEdit, lAllowDelete, lAllowConsult
    
    If ArrayTotal = 0 Then
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
    StatusBar = frmFornecedor.StatusBarAviso
    
    If txtContato.Enabled Then
        txtContato.SetFocus
    End If
End Sub
Public Sub Incluir()
    If lAlterar Then
       If Not Cancelamento Then
            Exit Sub
        End If
    End If
    
    lInserir = True
        
    StatusBar = "Inclusão de Departamento do Fornecedor"
    BarraDeStatus StatusBar
    
    ZeraCampos
    AtivaCampos
    
    BotãoGravar (lInserir Or lAllowEdit)
    BotãoIncluir False
    cmdGravar.Enabled = (lInserir Or lAllowEdit)
    cmdCancelar.Enabled = (lInserir Or lAllowEdit)
    
    NavegaçãoInferior False
    NavegaçãoSuperior False
    
    txtContato.SetFocus
End Sub
Public Sub MoveFirst()
    If lAlterar Then
       If Not Cancelamento Then
            Exit Sub
        End If
    End If
        
    StatusBar = frmFornecedor.StatusBarAviso
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
    
    StatusBar = frmFornecedor.StatusBarAviso
    BarraDeStatus StatusBar
    
    Elemento = ArrayTotal
    
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
    
    StatusBar = frmFornecedor.StatusBarAviso
    BarraDeStatus StatusBar
    
    Elemento = Elemento + 1
    
    If Elemento > ArrayTotal Then
        Elemento = ArrayTotal
        Exit Sub
    End If
    
    NavegaçãoInferior lAllowConsult
    TestaSuperiorArray Elemento, ArrayContato(), lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Sub
Public Sub MovePrevious()
    If lAlterar Then
       If Not Cancelamento Then
            Exit Sub
        End If
    End If
    
    StatusBar = frmFornecedor.StatusBarAviso
    BarraDeStatus StatusBar
    
    Elemento = Elemento - 1
    
    If Elemento < 1 Then
        Elemento = 1
        Exit Sub
    End If
    
    NavegaçãoSuperior lAllowConsult
    TestaInferiorArray Elemento, ArrayContato(), lAllowEdit, lAllowDelete, lAllowConsult
    
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
    txtContato = ArrayContato(Elemento)
    txtEndereço = ArrayEndereço(Elemento)
    txtBairro = ArrayBairro(Elemento)
    txtCidade = ArrayCidade(Elemento)
    txtUF = ArrayUF(Elemento)
    txtCep = ArrayCEP(Elemento)
    txtTelefone = ArrayTelefone(Elemento)
    txtFax = ArrayFax(Elemento)
    txtEMail = ArrayEMail(Elemento)
    
    Exit Sub
    
Erro:
    GeraMensagemDeErro "Departamento Fornecedor - GetRecords "
    Resume Next
End Sub
Private Sub SetRecords()
    ArrayContato(Elemento) = txtContato
    ArrayEndereço(Elemento) = txtEndereço
    ArrayBairro(Elemento) = txtBairro
    ArrayCidade(Elemento) = txtCidade
    ArrayUF(Elemento) = txtUF
    ArrayCEP(Elemento) = txtCep
    ArrayTelefone(Elemento) = txtTelefone
    ArrayFax(Elemento) = txtFax
    ArrayEMail(Elemento) = txtEMail
End Sub
Private Sub SaveArray()
    Dim Cont%
    
    frmFornecedor.SizeArray ArrayTotal
    frmFornecedor.ArrayTotal = ArrayTotal
    
    For Cont = 1 To frmFornecedor.ArrayTotal
        frmFornecedor.SetArray "Contato", ArrayContato(Cont), Cont
        frmFornecedor.SetArray "Endereço", ArrayEndereço(Cont), Cont
        frmFornecedor.SetArray "Bairro", ArrayBairro(Cont), Cont
        frmFornecedor.SetArray "Cidade", ArrayCidade(Cont), Cont
        frmFornecedor.SetArray "UF", ArrayUF(Cont), Cont
        frmFornecedor.SetArray "CEP", ArrayCEP(Cont), Cont
        frmFornecedor.SetArray "Telefone", ArrayTelefone(Cont), Cont
        frmFornecedor.SetArray "Fax", ArrayFax(Cont), Cont
        frmFornecedor.SetArray "EMail", ArrayEMail(Cont), Cont
    Next
End Sub
Private Sub SizeArray(ByVal Tamanho As Integer)
    ASize Tamanho, ArrayContato()
    ASize Tamanho, ArrayEndereço()
    ASize Tamanho, ArrayBairro()
    ASize Tamanho, ArrayCidade()
    ASize Tamanho, ArrayUF()
    ASize Tamanho, ArrayCEP()
    ASize Tamanho, ArrayTelefone()
    ASize Tamanho, ArrayFax()
    ASize Tamanho, ArrayEMail()
End Sub
Private Sub ZeraCampos()
    txtContato = Empty
    txtEndereço = Empty
    txtBairro = Empty
    txtCidade = Empty
    txtUF = Empty
    txtCep = Empty
    txtTelefone = Empty
    txtFax = Empty
    txtEMail = Empty
End Sub
Private Sub cmdCancelar_Click()
    Cancelamento
End Sub
Private Sub cmdGravar_Click()
    Gravar
End Sub
Private Sub Form_Activate()
    TestaInferiorArray Elemento, ArrayContato(), lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperiorArray Elemento, ArrayContato(), lAllowEdit, lAllowDelete, lAllowConsult
    
    BarraDeStatus StatusBar
    
    If ArrayTotal = 0 Then
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
    frmDepartamentoFornecedor.SetFocus
End Sub
Private Sub Form_Load()
    lPula = False
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    
    lAllowInsert = Allow("DEPARTAMENTO", "I")
    lAllowEdit = Allow("DEPARTAMENTO", "A")
    lAllowDelete = Allow("DEPARTAMENTO", "E")
    lAllowConsult = Allow("DEPARTAMENTO", "C")
    
    FillArray
        
    BotãoIncluir lAllowInsert
    
    StatusBar = frmFornecedor.StatusBarAviso
 
    If ArrayTotal = 0 Then
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
        
    If ArrayTotal = 0 Or ArrayTotal = 1 Then
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
            If Not frmFornecedor.lAlterarArray Then
                frmFornecedor.lAlterarArray = True
                If Not frmFornecedor.lInserir Then
                    frmFornecedor.lAlterar = True
                End If
            End If
        ElseIf Confirmação = vbCancel Then
            Cancel = 1
        ElseIf Confirmação = vbNo Then
        End If
    End If
    
    Set frmDepartamentoFornecedor = Nothing
End Sub
Private Sub txtBairro_Change()
    FormatMask "@!S30", txtBairro
End Sub
Private Sub txtBairro_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBar = "Alteração de Departamento do Fornecedor"
        BarraDeStatus StatusBar
    End If
End Sub
Private Sub txtCep_Change()
    If Not lPula Then
        lPula = True
        NumericSpaceOnly txtCep
        FormatMask "99.999-999", txtCep
        lPula = False
    End If
End Sub
Private Sub txtCep_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBar = "Alteração de Departamento do Fornecedor"
        BarraDeStatus StatusBar
    End If
End Sub
Private Sub txtCidade_Change()
    FormatMask "@!S30", txtCidade
End Sub
Private Sub txtCidade_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBar = "Alteração de Departamento do Fornecedor"
        BarraDeStatus StatusBar
    End If
End Sub
Private Sub txtContato_Change()
    FormatMask "@!S40", txtContato
End Sub
Private Sub txtContato_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBar = "Alteração de Departamento do Fornecedor"
        BarraDeStatus StatusBar
    End If
End Sub
Private Sub txtEMail_Change()
    FormatMask "@S40", txtEMail
End Sub
Private Sub txtEMail_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBar = "Alteração de Departamento do Fornecedor"
        BarraDeStatus StatusBar
    End If
End Sub
Private Sub txtEndereço_Change()
    FormatMask "@S40", txtEndereço
End Sub
Private Sub txtEndereço_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBar = "Alteração de Departamento do Fornecedor"
        BarraDeStatus StatusBar
    End If
End Sub
Private Sub txtFax_Change()
    If Not lPula Then
        lPula = True
        NumericSpaceOnly txtFax
        FormatMask "(####)####-####", txtFax
        lPula = False
    End If
End Sub
Private Sub txtFax_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBar = "Alteração de Departamento do Fornecedor"
        BarraDeStatus StatusBar
    End If
End Sub
Private Sub txtTelefone_Change()
    If Not lPula Then
        lPula = True
        NumericSpaceOnly txtTelefone
        FormatMask "(####)####-####", txtTelefone
        lPula = False
    End If
End Sub
Private Sub txtTelefone_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBar = "Alteração de Departamento do Fornecedor"
        BarraDeStatus StatusBar
    End If
End Sub
Private Sub txtUF_Change()
    FormatMask "@! AA", txtUF
End Sub
Private Sub txtUF_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBar = "Alteração de Departamento do Fornecedor"
        BarraDeStatus StatusBar
    End If
End Sub
