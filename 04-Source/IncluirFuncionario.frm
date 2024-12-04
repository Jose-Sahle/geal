VERSION 5.00
Begin VB.Form frmUsu�rioCadastro 
   Caption         =   "Cadastro de Usu�rio"
   ClientHeight    =   1785
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6540
   Icon            =   "IncluirFuncionario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   6540
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5250
      TabIndex        =   4
      Top             =   1380
      Width           =   1245
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      Height          =   375
      Left            =   3960
      TabIndex        =   3
      Top             =   1380
      Width           =   1245
   End
   Begin VB.Frame frUserName 
      Height          =   1305
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   6525
      Begin VB.TextBox txtNomeDoFuncion�rio 
         Height          =   285
         Left            =   2580
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   330
         Width           =   3825
      End
      Begin VB.TextBox txtC�digoDoFuncion�rio 
         Height          =   285
         Left            =   1950
         TabIndex        =   0
         Top             =   330
         Width           =   585
      End
      Begin VB.TextBox txtUserName 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1110
         TabIndex        =   2
         Top             =   780
         Width           =   1425
      End
      Begin VB.Label lblC�digoDeFuncion�rio 
         Caption         =   "C�digo do Funcion�rio"
         Height          =   195
         Left            =   180
         TabIndex        =   7
         Top             =   360
         Width           =   1635
      End
      Begin VB.Label lblUserName 
         Caption         =   "Usu�rio"
         Height          =   225
         Left            =   180
         TabIndex        =   6
         Top             =   810
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmUsu�rioCadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TBLUsu�rio As Table
Dim Usu�rioAberto As Boolean
Dim IndiceAtivoUsu�rio As String

Dim lInserir As Boolean
Dim lAlterar As Boolean
Dim lPush As Boolean
Dim mFechar As Boolean
Dim mFirst As Boolean
Dim mPula As Boolean

Dim StatusBarAviso$

Public TipoOpera��o As Integer
Public CampoChave As String
Public Cancel As Boolean

Public lAtualizar As Boolean
Public Sub Gravar()
    If lInserir Then
        If SetRecords Then
            StatusBarAviso = "Inclus�o bem sucedida"
            Unload Me
            GoTo Fim
            Exit Sub
        Else
            StatusBarAviso = "Falha na inclus�o"
        End If
    Else
        If SetRecords Then
            StatusBarAviso = "Altera��o bem sucedida"
            Unload Me
            GoTo Fim
            Exit Sub
        Else
            StatusBarAviso = "Falha na altera��o"
        End If
    End If
    
    If txtC�digoDoFuncion�rio.Enabled Then
        txtC�digoDoFuncion�rio.SetFocus
    End If
Fim:
    BarraDeStatus StatusBarAviso
End Sub
Private Function PosRecords(ByVal Chave$)
    TBLUsu�rio.Seek "=", Chave
    If TBLUsu�rio.NoMatch Then
        MsgBox "N�o consegui encontrar o UserName " + Chave, vbExclamation, "Erro"
        PosRecords = False
    Else
        PosRecords = True
    End If
End Function
Private Sub GetRecords()
    lPush = True
    mPula = True
    txtC�digoDoFuncion�rio = TBLUsu�rio("C�DIGO DE FUNCION�RIO")
    txtNomeDoFuncion�rio = BuscaFuncion�rio(TBLUsu�rio("C�DIGO DE FUNCION�RIO"))
    txtUserName = CampoChave
    lPush = False
    mPula = False
End Sub
Private Function SetRecords() As Boolean
    On Error GoTo Erro

    WS.BeginTrans
    
    If lInserir Then
        TBLUsu�rio.AddNew
    Else
        TBLUsu�rio.Edit
    End If
    
    If lInserir Then
        TBLUsu�rio("C�DIGO DE FUNCION�RIO") = txtC�digoDoFuncion�rio
        TBLUsu�rio("SENHA") = ValidaSenha("GEAL")
    End If
    TBLUsu�rio("USERNAME") = txtUserName
    CampoChave = txtUserName
    
    If lInserir Then
        TBLUsu�rio("USERNAME - CRIA") = gUsu�rio
        TBLUsu�rio("DATA - CRIA") = Date
        TBLUsu�rio("HORA - CRIA") = Time
        TBLUsu�rio("USERNAME - ALTERA") = "VAZIO"
        TBLUsu�rio("DATA - ALTERA") = vbNull
        TBLUsu�rio("HORA - ALTERA") = vbNull
    End If
    If lAlterar Then
        TBLUsu�rio("USERNAME - ALTERA") = gUsu�rio
        TBLUsu�rio("DATA - ALTERA") = Date
        TBLUsu�rio("HORA - ALTERA") = Time
    End If
    TBLUsu�rio.Update
    
Erro:
    If Err <> 0 Then
        TBLUsu�rio.CancelUpdate
        GeraMensagemDeErro "Usu�rioCadastro - SetRecords - " & txtUserName, True
        SetRecords = False
        Exit Function
    End If
    
    WS.CommitTrans 'Grava as altera��es ou inclus�es se n�o houverem erros
    
    SetRecords = True
End Function
Private Sub ZeraCampos()
    lPush = True
    mPula = True
    txtC�digoDoFuncion�rio = Empty
    txtNomeDoFuncion�rio = Empty
    txtUserName = ""
    mPula = False
    lPush = False
End Sub
Private Sub cmdCancelar_Click()
    Cancel = True
    Unload Me
End Sub
Private Sub cmdOK_Click()
    Gravar
    Cancel = False
End Sub
Private Sub Form_Activate()
    If Not Usu�rioAberto And Not mFechar Then
        Unload Me
        Exit Sub
    End If
    
    If mFirst And TipoOpera��o = vbIncluir Then
        txtC�digoDoFuncion�rio.SetFocus
    Else
        txtUserName.SetFocus
    End If
    
    mFirst = False
    
    If lAtualizar Then
        Bot�oAtualizar True
    Else
        Bot�oAtualizar False
    End If
End Sub
Private Sub Form_Load()

    mFechar = False
    mPula = False
    
    Usu�rioAberto = AbreTabela(Dicion�rio, "USU�RIO", "USU�RIO", DBUsu�rio, TBLUsu�rio, TBLTabela, dbOpenTable)
        
    If Usu�rioAberto Then
        IndiceAtivoUsu�rio = "USU�RIO1"
        TBLUsu�rio.Index = IndiceAtivoUsu�rio
    Else
        MsgBox "N�o consegui abrir a tabela 'Usu�rio' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    If TipoOpera��o = vbIncluir Then
        ZeraCampos
        lInserir = True
        lAlterar = False
    ElseIf TipoOpera��o = vbAlterar Then
        txtC�digoDoFuncion�rio.Enabled = False
        If PosRecords(CampoChave) Then
            GetRecords
            lInserir = False
            lAlterar = True
        Else
            mFechar = True
            Exit Sub
        End If
    End If
    mFirst = True
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If Usu�rioAberto Then
        TBLUsu�rio.Close
    End If
    
    Set frmUsu�rioCadastro = Nothing
End Sub
Private Sub txtC�digoDoFuncion�rio_Change()
    FormatMask "99", txtC�digoDoFuncion�rio
End Sub
Private Sub txtC�digoDoFuncion�rio_LostFocus()
    If txtC�digoDoFuncion�rio = Empty Then
        Exit Sub
    End If
    If Not IsCorrectFuncion�rio(txtC�digoDoFuncion�rio) Then
        MsgBox "Funcion�rio n�o cadastrado!", vbInformation, "Aviso"
        Set frmEncontrar.DBBancoDeDados = DBUsu�rio
        frmEncontrar.LabelDescription = "Nome"
        frmEncontrar.NomeDaJanela = "Funcion�rio"
        frmEncontrar.Mensagem = "Nenhum funcion�rio foi selecionado!"
        frmEncontrar.BancoDeDados = "USU�RIO"
        frmEncontrar.Tabela = "FUNCION�RIO"
        frmEncontrar.Indice = "1"
        frmEncontrar.CampoChave = "C�DIGO"
        frmEncontrar.CampoPreencheLista = "NOME"
        frmEncontrar.Show vbModal
        txtC�digoDoFuncion�rio = frmEncontrar.Chave
        txtNomeDoFuncion�rio = frmEncontrar.Nome
    Else
        txtNomeDoFuncion�rio = BuscaFuncion�rio(Val(txtC�digoDoFuncion�rio))
    End If
End Sub
Private Sub txtUserName_Change()
    FormatMask "@! AAAAAA", txtUserName
End Sub
Private Sub txtUserName_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        If Not lPush Then
            lAlterar = True
            StatusBarAviso = "Altera��o"
            BarraDeStatus StatusBarAviso
        End If
    End If
End Sub
