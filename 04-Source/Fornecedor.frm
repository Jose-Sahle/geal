VERSION 5.00
Begin VB.Form frmFornecedor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fornecedor"
   ClientHeight    =   4080
   ClientLeft      =   2970
   ClientTop       =   1500
   ClientWidth     =   6540
   Icon            =   "Fornecedor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4080
   ScaleWidth      =   6540
   Begin VB.CommandButton cmdGravar 
      Caption         =   "&Gravar"
      Default         =   -1  'True
      Height          =   345
      Left            =   3945
      TabIndex        =   8
      Top             =   3735
      Width           =   1245
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   5265
      TabIndex        =   9
      Top             =   3735
      Width           =   1245
   End
   Begin VB.Frame frDadosDeInscrições 
      Caption         =   "Inscrições"
      Height          =   1125
      Left            =   0
      TabIndex        =   18
      Top             =   2565
      Width           =   6525
      Begin VB.TextBox txtCgcCpf 
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
         TabIndex        =   7
         Text            =   "*"
         Top             =   690
         Width           =   2310
      End
      Begin VB.TextBox txtInscrEst 
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
         Text            =   "*"
         Top             =   240
         Width           =   2310
      End
      Begin VB.Label lblCgc 
         Caption         =   "C. G. C."
         Height          =   195
         Left            =   150
         TabIndex        =   20
         Top             =   720
         Width           =   825
      End
      Begin VB.Label lblInscrEst 
         Caption         =   "Inscr. Est."
         Height          =   225
         Left            =   150
         TabIndex        =   19
         Top             =   270
         Width           =   885
      End
   End
   Begin VB.Frame frDadosCadastrais 
      Caption         =   " Dados Cadastrais "
      Height          =   2565
      Left            =   0
      TabIndex        =   11
      Top             =   -15
      Width           =   6525
      Begin VB.CommandButton cmdDepartamento 
         Caption         =   "&Departamento..."
         Height          =   345
         Left            =   5115
         TabIndex        =   10
         Top             =   2070
         Width           =   1335
      End
      Begin VB.TextBox txtNomeRazãoSocial 
         Height          =   285
         Left            =   1200
         TabIndex        =   0
         Text            =   "*"
         Top             =   330
         Width           =   5235
      End
      Begin VB.TextBox txtEndereço 
         Height          =   285
         Left            =   1200
         TabIndex        =   1
         Text            =   "*"
         Top             =   780
         Width           =   5235
      End
      Begin VB.TextBox txtBairro 
         Height          =   285
         Left            =   1200
         TabIndex        =   2
         Text            =   "*"
         Top             =   1230
         Width           =   5235
      End
      Begin VB.TextBox txtCidade 
         Height          =   285
         Left            =   1200
         TabIndex        =   3
         Text            =   "*"
         Top             =   1680
         Width           =   5235
      End
      Begin VB.TextBox txtUF 
         Height          =   285
         Left            =   1200
         TabIndex        =   4
         Text            =   "*"
         Top             =   2130
         Width           =   435
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
         Left            =   2700
         TabIndex        =   5
         Text            =   "*"
         Top             =   2100
         Width           =   1305
      End
      Begin VB.Label lblNomeRazãoSocial 
         Caption         =   "Razão Social"
         Height          =   195
         Left            =   150
         TabIndex        =   17
         Top             =   360
         Width           =   1065
      End
      Begin VB.Label lblEndereço 
         Caption         =   "Endereço"
         Height          =   195
         Left            =   150
         TabIndex        =   16
         Top             =   810
         Width           =   975
      End
      Begin VB.Label lblBairro 
         Caption         =   "Bairro"
         Height          =   225
         Left            =   150
         TabIndex        =   15
         Top             =   1260
         Width           =   945
      End
      Begin VB.Label lblCidade 
         Caption         =   "Cidade"
         Height          =   225
         Left            =   150
         TabIndex        =   14
         Top             =   1710
         Width           =   945
      End
      Begin VB.Label lblUF 
         Caption         =   "U. F."
         Height          =   225
         Left            =   150
         TabIndex        =   13
         Top             =   2160
         Width           =   465
      End
      Begin VB.Label lblCep 
         Caption         =   "CEP"
         Height          =   195
         Left            =   2250
         TabIndex        =   12
         Top             =   2160
         Width           =   315
      End
   End
End
Attribute VB_Name = "frmFornecedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TBLFornecedor As Table
Dim FornecedorAberto As Boolean
Dim IndiceFornecedorAtivo$
Dim txtCGCCPFFornecedorAnterior As String

Dim TBLDepartamentoFornecedor As Table
Dim DepartamentoFornecedorAberto As Boolean
Dim IndiceDepartamentoFornecedorAtivo$

Dim lAllowInsert  As Boolean
Dim lAllowEdit    As Boolean
Dim lAllowDelete  As Boolean
Dim lAllowConsult As Boolean

Public ArrayTotal%
Dim ArrayContato() As Variant
Dim ArrayEndereço() As Variant
Dim ArrayBairro() As Variant
Dim ArrayCidade() As Variant
Dim ArrayUF() As Variant
Dim ArrayCEP() As Variant
Dim ArrayTelefone() As Variant
Dim ArrayFax() As Variant
Dim ArrayEMail() As Variant
Dim ArrayRecno() As Variant

Public lInserir As Boolean
Public lAlterar As Boolean
Public lAlterarArray As Boolean

Dim mFechar As Boolean

Public StatusBarAviso$

Dim lPula As Boolean

Dim DataBaseName(1 To 1) As String
Public Relatório$
Public TotalDatabaseName%

Public lAtualizar As Boolean
Private Sub AtivaCampos()
    BotãoImprimir True
    frDadosCadastrais.Enabled = True
    frDadosDeInscrições.Enabled = True
    BotãoGravar (lInserir Or lAllowEdit)
    cmdGravar.Enabled = (lInserir Or lAllowEdit)
    cmdCancelar.Enabled = (lInserir Or lAllowEdit)
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
        StatusBarAviso = "Inclusão cancelada"
    End If
    If lAlterar Then
        StatusBarAviso = "Alteração cancelada"
    End If
    BarraDeStatus StatusBarAviso
    
    BotãoIncluir lAllowInsert
    ClearArray
    
    If TBLFornecedor.RecordCount = 0 Then
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
    lAlterarArray = False
    
    Cancelamento = True
    
    TestaInferior TBLFornecedor, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLFornecedor, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Function
Public Sub ClearArray()
    ArrayTotal = 0
    lAlterarArray = False
    frmDepartamentoFornecedor.lAlteração = False
    ReDim ArrayRecno(0)
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
    BotãoImprimir False
    frDadosCadastrais.Enabled = False
    frDadosDeInscrições.Enabled = False
    cmdGravar.Enabled = False
    cmdCancelar.Enabled = False
    BotãoGravar False
End Sub
Public Sub Encontrar()
    If Not lAllowConsult Then
        Exit Sub
    End If
    Set frmEncontrar.DBBancoDeDados = DBCadastro
    frmEncontrar.NomeDaJanela = "Fornecedor"
    frmEncontrar.Mensagem = "Nenhum fornecedor foi selecionado!"
    frmEncontrar.BancoDeDados = "CADASTRO"
    frmEncontrar.Tabela = "FORNECEDOR"
    frmEncontrar.Indice = "2"
    frmEncontrar.CampoChave = "CGC - CPF"
    frmEncontrar.CampoPreencheLista = "RAZÃO SOCIAL"
    frmEncontrar.Show vbModal
    lPula = True
    txtCgcCpf = frmEncontrar.Chave
    lPula = False
    PosRecords
End Sub
Public Sub Excluir()
    Dim Confirmação As Integer, Msg1$, Msg2$
    Dim SQL As String

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
    
    TBLFornecedor.Delete
    
    SQL = "Delete * From [DEPARTAMENTO - FORNECEDOR] Where [CGC - CPF]= '" + txtCgcCpf + "'"
    DBCadastro.Execute SQL
    
    If Err <> 0 Then
        GeraMensagemDeErro "Fornecedor - Excluir", True
        StatusBarAviso = "Falha na exclusão"
        BarraDeStatus StatusBarAviso
        Exit Sub
    End If
    
    WS.CommitTrans
        
    Log gUsuário, "Exclusão - Fornecedor: " & txtNomeRazãoSocial
    
    ClearArray
    
    StatusBarAviso = "Exclusão bem sucedida"
    BarraDeStatus StatusBarAviso
    
    If TBLFornecedor.RecordCount = 0 Then
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
    
    If TBLFornecedor.BOF Then
        TBLFornecedor.MoveFirst
    ElseIf TBLFornecedor.EOF Then
        TBLFornecedor.MoveLast
    Else
        TBLFornecedor.MovePrevious
        If TBLFornecedor.BOF Then
            TBLFornecedor.MoveNext
        End If
    End If
    
    GetRecords
    
    TestaInferior TBLFornecedor, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLFornecedor, lAllowEdit, lAllowDelete, lAllowConsult
End Sub
Private Sub FillArray()
    TBLDepartamentoFornecedor.Seek "=", txtCgcCpf
    
    If TBLDepartamentoFornecedor.NoMatch Then
        Exit Sub
    End If
    
    ArrayTotal = 0
    
    ReDim ArrayRecno(1 To 1)
    
    Do While Not TBLDepartamentoFornecedor.EOF
        If TBLDepartamentoFornecedor("CGC - CPF") <> txtCgcCpf Then
            Exit Do
        End If
        
        ArrayTotal = ArrayTotal + 1
        
        SizeArray (ArrayTotal)
        
        ReDim Preserve ArrayRecno(1 To ArrayTotal)
        ArrayRecno(ArrayTotal) = TBLDepartamentoFornecedor.Bookmark
            
        ArrayContato(ArrayTotal) = TBLDepartamentoFornecedor("CONTATO")
        ArrayEndereço(ArrayTotal) = TBLDepartamentoFornecedor("ENDEREÇO")
        ArrayBairro(ArrayTotal) = TBLDepartamentoFornecedor("BAIRRO")
        ArrayCidade(ArrayTotal) = TBLDepartamentoFornecedor("CIDADE")
        ArrayUF(ArrayTotal) = TBLDepartamentoFornecedor("UF")
        ArrayCEP(ArrayTotal) = TBLDepartamentoFornecedor("CEP")
        ArrayTelefone(ArrayTotal) = TBLDepartamentoFornecedor("TELEFONE")
        ArrayFax(ArrayTotal) = TBLDepartamentoFornecedor("FAX")
        ArrayEMail(ArrayTotal) = TBLDepartamentoFornecedor("E-MAIL")
        
        TBLDepartamentoFornecedor.MoveNext
    Loop
End Sub
Public Function GetArray(ByVal Nome As String, ByVal Elemento As Integer) As String
    If Nome = "Contato" Then
        GetArray = ArrayContato(Elemento)
    ElseIf Nome = "Endereço" Then
        GetArray = ArrayEndereço(Elemento)
    ElseIf Nome = "Bairro" Then
        GetArray = ArrayBairro(Elemento)
    ElseIf Nome = "Cidade" Then
        GetArray = ArrayCidade(Elemento)
    ElseIf Nome = "UF" Then
        GetArray = ArrayUF(Elemento)
    ElseIf Nome = "CEP" Then
        GetArray = ArrayCEP(Elemento)
    ElseIf Nome = "Telefone" Then
        GetArray = ArrayTelefone(Elemento)
    ElseIf Nome = "Fax" Then
        GetArray = ArrayFax(Elemento)
    ElseIf Nome = "EMail" Then
        GetArray = ArrayEMail(Elemento)
    End If
End Function
Public Sub Gravar()
    If lInserir Then
        If SetRecords Then
            PosRecords
            lInserir = False
            ClearArray
            StatusBarAviso = "Inclusão bem sucedida"
        Else
            StatusBarAviso = "Falha na inclusão"
        End If
    ElseIf lAlterar Then
        If TBLFornecedor.RecordCount > 0 And Not TBLFornecedor.BOF And Not TBLFornecedor.EOF Then
            If SetRecords Then
                PosRecords
                lAlterar = False
                ClearArray
                StatusBarAviso = "Alteração bem sucedida"
            Else
                StatusBarAviso = "Falha na alteração"
            End If
        End If
    End If
    
    BarraDeStatus StatusBarAviso
    
    TestaInferior TBLFornecedor, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLFornecedor, lAllowEdit, lAllowDelete, lAllowConsult
    
    If TBLFornecedor.RecordCount = 0 Then
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
    
    If txtNomeRazãoSocial.Enabled Then
        txtNomeRazãoSocial.SetFocus
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
    
    txtNomeRazãoSocial.SetFocus
End Sub
Public Sub MoveFirst()
    If lAlterar Then
       If Not Cancelamento Then
            Exit Sub
        End If
    End If
    
    StatusBarAviso = "Pronto"
    BarraDeStatus StatusBarAviso
    
    TBLFornecedor.MoveFirst
    ClearArray
    
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
    
    StatusBarAviso = "Pronto"
    BarraDeStatus StatusBarAviso
    
    TBLFornecedor.MoveLast
    ClearArray
    
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
    
    TBLFornecedor.MoveNext
    
    If TBLFornecedor.EOF Then
        TBLFornecedor.MovePrevious
        Exit Sub
    End If
    
    ClearArray
    
    StatusBarAviso = "Pronto"
    BarraDeStatus StatusBarAviso
    
    NavegaçãoInferior lAllowConsult
    TestaSuperior TBLFornecedor, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Sub
Public Sub MovePrevious()
    If lAlterar Then
       If Not Cancelamento Then
            Exit Sub
        End If
    End If
    
    TBLFornecedor.MovePrevious
    
    If TBLFornecedor.BOF Then
        TBLFornecedor.MoveNext
        Exit Sub
    End If
    
    ClearArray
    
    StatusBarAviso = "Pronto"
    BarraDeStatus StatusBarAviso
    
    NavegaçãoSuperior lAllowConsult
    TestaInferior TBLFornecedor, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Sub
Public Sub PosRecords()
    If TBLFornecedor.RecordCount = 0 Then
        Exit Sub
    End If
    
    TBLFornecedor.Seek "=", txtCgcCpf
    If TBLFornecedor.NoMatch Then
        MsgBox "Não consegui encontrar " + txtCgcCpf, vbExclamation, "Erro"
        TBLFornecedor.MoveFirst
        NavegaçãoInferior False
        NavegaçãoInferior lAllowConsult
    Else
        TestaInferior TBLFornecedor, lAllowEdit, lAllowDelete, lAllowConsult
        TestaSuperior TBLFornecedor, lAllowEdit, lAllowDelete, lAllowConsult
    End If
    GetRecords
End Sub
Public Function PushDataBaseName(ByVal Posição As Integer) As String
    PushDataBaseName = DataBaseName(Posição)
End Function
Public Sub GetRecords()
    On Error GoTo Erro
    
    lPula = True
    If Not lAllowConsult Then
        ZeraCampos
        DesativaCampos
        lPula = False
        Exit Sub
    End If
    txtNomeRazãoSocial = TBLFornecedor("RAZÃO SOCIAL")
    txtEndereço = TBLFornecedor("ENDEREÇO")
    txtBairro = TBLFornecedor("BAIRRO")
    txtCidade = TBLFornecedor("CIDADE")
    txtUF = TBLFornecedor("UF")
    txtCep = TBLFornecedor("CEP")
    txtInscrEst = TBLFornecedor("INSCRIÇÃO ESTADUAL")
    txtCgcCpf = TBLFornecedor("CGC - CPF")
    lPula = False
    If Not lAllowEdit Then
        DesativaCampos
    End If
    
    Exit Sub
    
Erro:
    GeraMensagemDeErro "Fornecedor - GetRecords "
    Resume Next
End Sub
Private Function SetRecords()
    On Error GoTo Erro
    
    Dim Msg$
    Dim Confirmação As Integer, Msg1$, Msg2$
    Dim SQL As String
    Dim Cont%, Recno%
    
    WS.BeginTrans 'Inicia uma Transação
        
    If lInserir Then
        TBLFornecedor.AddNew
    Else
        TBLFornecedor.Edit
    End If
    
    TBLFornecedor("RAZÃO SOCIAL") = txtNomeRazãoSocial
    TBLFornecedor("ENDEREÇO") = txtEndereço
    TBLFornecedor("BAIRRO") = txtBairro
    TBLFornecedor("CIDADE") = txtCidade
    TBLFornecedor("UF") = txtUF
    TBLFornecedor("CEP") = Mid(txtCep, 1, 10)
    TBLFornecedor("INSCRIÇÃO ESTADUAL") = Mid(txtInscrEst, 1, 14)
    TBLFornecedor("CGC - CPF") = Mid(txtCgcCpf, 1, 18)
    If lInserir Then
        TBLFornecedor("USERNAME - CRIA") = gUsuário
        TBLFornecedor("DATA - CRIA") = Date
        TBLFornecedor("HORA - CRIA") = Time
        TBLFornecedor("USERNAME - ALTERA") = "VAZIO"
        TBLFornecedor("DATA - ALTERA") = vbNull
        TBLFornecedor("HORA - ALTERA") = vbNull
    End If
    If lAlterar Then
        TBLFornecedor("USERNAME - ALTERA") = gUsuário
        TBLFornecedor("DATA - ALTERA") = Date
        TBLFornecedor("HORA - ALTERA") = Time
    End If
    TBLFornecedor.Update
    
Erro:
    If Err <> 0 Then
        TBLFornecedor.CancelUpdate
        GeraMensagemDeErro "Fornecedor - SetRecords", True
        SetRecords = False
        Exit Function
    End If
    
    Recno = 0
    
    If Err = 0 Then
        On Error Resume Next
        Recno = UBound(ArrayRecno)
        Err = 0
    End If
    
    On Error GoTo ErroDep
    
    If lAlterarArray Then
        For Cont = 1 To ArrayTotal
            If Cont <= Recno Then
                TBLDepartamentoFornecedor.Bookmark = ArrayRecno(Cont)
                TBLDepartamentoFornecedor.Edit
            Else
                TBLDepartamentoFornecedor.AddNew
                lAlterar = False
                lInserir = True
            End If
            TBLDepartamentoFornecedor("CGC - CPF") = txtCgcCpf
            TBLDepartamentoFornecedor("CONTATO") = ArrayContato(Cont)
            TBLDepartamentoFornecedor("ENDEREÇO") = ArrayEndereço(Cont)
            TBLDepartamentoFornecedor("BAIRRO") = ArrayBairro(Cont)
            TBLDepartamentoFornecedor("CIDADE") = ArrayCidade(Cont)
            TBLDepartamentoFornecedor("UF") = ArrayUF(Cont)
            TBLDepartamentoFornecedor("CEP") = ArrayCEP(Cont)
            TBLDepartamentoFornecedor("TELEFONE") = ArrayTelefone(Cont)
            TBLDepartamentoFornecedor("FAX") = ArrayFax(Cont)
            TBLDepartamentoFornecedor("E-MAIL") = ArrayEMail(Cont)
            If lInserir Then
                TBLDepartamentoFornecedor("USERNAME - CRIA") = gUsuário
                TBLDepartamentoFornecedor("DATA - CRIA") = Date
                TBLDepartamentoFornecedor("HORA - CRIA") = Time
                TBLDepartamentoFornecedor("USERNAME - ALTERA") = "VAZIO"
                TBLDepartamentoFornecedor("DATA - ALTERA") = vbNull
                TBLDepartamentoFornecedor("HORA - ALTERA") = vbNull
            End If
            If lAlterar Then
                TBLDepartamentoFornecedor("USERNAME - ALTERA") = gUsuário
                TBLDepartamentoFornecedor("DATA - ALTERA") = Date
                TBLDepartamentoFornecedor("HORA - ALTERA") = Time
            End If
            TBLDepartamentoFornecedor.Update
        Next
        If Cont <= Recno Then
            ArrayTotal = Cont
            For Cont = ArrayTotal To Recno
                TBLDepartamentoFornecedor.Bookmark = ArrayRecno(Cont)
                TBLDepartamentoFornecedor.Delete
            Next
        End If
    End If
    
ErroDep:
    If Err <> 0 Then
        TBLDepartamentoFornecedor.CancelUpdate
        GeraMensagemDeErro "Fornecedor - SetRecords", True
        SetRecords = False
        Exit Function
    End If

    WS.CommitTrans 'Grava as alterações ou inclusões se não houverem erros
        
    If lInserir Then
        Log gUsuário, "Inclusão - Fornecedor: " & txtNomeRazãoSocial
    Else
        Log gUsuário, "Alteração - Fornecedor: " & txtNomeRazãoSocial
    End If
    
    lAlterar = False
    lInserir = False
    ClearArray
    
    SetRecords = True
End Function
Public Sub SetArray(ByVal Nome As String, ByVal Valor As String, ByVal Elemento As Integer)
    If Nome = "Contato" Then
        ArrayContato(Elemento) = Valor
    ElseIf Nome = "Endereço" Then
        ArrayEndereço(Elemento) = Valor
    ElseIf Nome = "Bairro" Then
        ArrayBairro(Elemento) = Valor
    ElseIf Nome = "Cidade" Then
        ArrayCidade(Elemento) = Valor
    ElseIf Nome = "UF" Then
        ArrayUF(Elemento) = Valor
    ElseIf Nome = "CEP" Then
        ArrayCEP(Elemento) = Valor
    ElseIf Nome = "Telefone" Then
        ArrayTelefone(Elemento) = Valor
    ElseIf Nome = "Fax" Then
        ArrayFax(Elemento) = Valor
    ElseIf Nome = "EMail" Then
        ArrayEMail(Elemento) = Valor
    End If
End Sub
Public Sub SizeArray(ByVal Tamanho As Integer)
    ArrayTotal = Tamanho
    ASize Tamanho, ArrayContato
    ASize Tamanho, ArrayEndereço
    ASize Tamanho, ArrayBairro
    ASize Tamanho, ArrayCidade
    ASize Tamanho, ArrayUF
    ASize Tamanho, ArrayCEP
    ASize Tamanho, ArrayTelefone
    ASize Tamanho, ArrayFax
    ASize Tamanho, ArrayEMail
End Sub
Private Sub ZeraCampos()
    txtNomeRazãoSocial = Empty
    txtEndereço = Empty
    txtBairro = Empty
    txtCidade = Empty
    txtUF = Empty
    txtCep = Empty
    txtInscrEst = Empty
    txtCgcCpf = Empty
End Sub
Private Sub cmdCancelar_Click()
    Cancelamento
End Sub
Private Sub cmdDepartamento_Click()
    If Not lInserir Then
        StatusBarAviso = "Alteração"
        BarraDeStatus StatusBarAviso
    End If
    If Not lAlterarArray Then
        FillArray
    End If
    frmDepartamentoFornecedor.Show 0
End Sub
Private Sub cmdGravar_Click()
    Gravar
End Sub
Private Sub Form_Activate()
    If mFechar Then
        Unload Me
        Exit Sub
    End If
    If Not FornecedorAberto Then
        Unload Me
        Exit Sub
    End If
    If Not DepartamentoFornecedorAberto Then
        Unload Me
        Exit Sub
    End If
    
    TestaInferior TBLFornecedor, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLFornecedor, lAllowEdit, lAllowDelete, lAllowConsult
    If TBLFornecedor.RecordCount = 0 Then
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
        BotãoIncluir lAllowInsert
    Else
        BotãoIncluir lAllowInsert
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
    
    lAllowInsert = Allow("FORNECEDOR", "I")
    lAllowEdit = Allow("FORNECEDOR", "A")
    lAllowDelete = Allow("FORNECEDOR", "E")
    lAllowConsult = Allow("FORNECEDOR", "C")
    
    ZeraCampos
    
    lInserir = False
    lAlterar = False
    lPula = False
    
    FornecedorAberto = AbreTabela(Dicionário, "CADASTRO", "FORNECEDOR", DBCadastro, TBLFornecedor, TBLTabela, dbOpenTable)
    
    If FornecedorAberto Then
        IndiceFornecedorAtivo = "FORNECEDOR1"
        TBLFornecedor.Index = IndiceFornecedorAtivo
    Else
        MsgBox "Não consegui abrir a tabela 'Fornecedor' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    DepartamentoFornecedorAberto = AbreTabela(Dicionário, "CADASTRO", "DEPARTAMENTO - FORNECEDOR", DBCadastro, TBLDepartamentoFornecedor, TBLTabela, dbOpenTable)
    
    If DepartamentoFornecedorAberto Then
        IndiceDepartamentoFornecedorAtivo = "DEPARTAMENTOFORNECEDOR1"
        TBLDepartamentoFornecedor.Index = IndiceDepartamentoFornecedorAtivo
    Else
        MsgBox "Não consegui abrir a tabela 'Departamento do Fornecedor' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    BotãoIncluir lAllowInsert
 
    If TBLFornecedor.RecordCount = 0 Then
        DesativaCampos
        BotãoExcluir False
        BotãoGravar False
    Else
        AtivaCampos
        BotãoExcluir lAllowDelete
        BotãoGravar (lInserir Or lAllowEdit)
        GetRecords
    End If
    
    NavegaçãoInferior False
        
    If TBLFornecedor.RecordCount = 0 Or TBLFornecedor.RecordCount = 1 Then
        NavegaçãoSuperior False
    Else
        NavegaçãoInferior lAllowConsult
    End If
    
    lInserir = False
    lAlterar = False
    StatusBarAviso = "Pronto"
    Relatório = AddPath(AplicaçãoPath, "REPORT\FORNECEDOR.RPT")
    TotalDatabaseName = 1
    DataBaseName(1) = AddPath(AplicaçãoPath, "DATABASE\CADASTRO.MDB")
    mFechar = False
    Exit Sub
    
Erro:
    GeraMensagemDeErro "Fornecedor - Load"
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
    
    Set frmFornecedor = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If FornecedorAberto Then
        TBLFornecedor.Close
    End If
    If DepartamentoFornecedorAberto Then
        TBLDepartamentoFornecedor.Close
    End If
    If Forms.Count = 2 Then
        AllBotões False
    End If
End Sub
Private Sub txtBairro_Change()
    If lPula Then Exit Sub
    FormatMask "@!S30", txtBairro
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Alteração"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub txtCep_Change()
    If Not lPula Then
        lPula = True
        NumericSpaceOnly txtCep
        FormatMask "99.999-999", txtCep
        If Not lInserir Then
            lAlterar = True
            StatusBarAviso = "Alteração"
            BarraDeStatus StatusBarAviso
        End If
        lPula = False
    End If
End Sub
Private Sub txtCgcCpf_Change()
    If Not lPula Then
        lPula = True
        NumericOnly txtCgcCpf
        FormatMask "99.999.999/9999-99", txtCgcCpf
        If Not lInserir Then
            lAlterar = True
            StatusBarAviso = "Alteração"
            BarraDeStatus StatusBarAviso
        End If
        lPula = False
    End If
End Sub
Private Sub txtCgcCpf_LostFocus()
    If Not IsCorrectCGC(txtCgcCpf) Then
        MsgBox "C. G. C. Inválido!", vbCritical, "Error"
        txtCgcCpf.SetFocus
    End If
End Sub
Private Sub txtCidade_Change()
    If lPula Then Exit Sub
    FormatMask "@!S30", txtCidade
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Alteração"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub txtEndereço_Change()
    If lPula Then Exit Sub
    FormatMask "@S40", txtEndereço
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Alteração"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub txtInscrEst_Change()
    If lPula Then Exit Sub
    FormatMask "@S14", txtInscrEst
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Alteração"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub txtNomeRazãoSocial_Change()
    If lPula Then Exit Sub
    
    FormatMask "@!S40", txtNomeRazãoSocial
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Alteração"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub txtUF_Change()
    If lPula Then Exit Sub
    
    FormatMask "@! AA", txtUF
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Alteração"
        BarraDeStatus StatusBarAviso
    End If
End Sub
