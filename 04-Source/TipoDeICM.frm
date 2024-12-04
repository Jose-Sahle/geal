VERSION 5.00
Begin VB.Form frmTipoDeICM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipo de ICM"
   ClientHeight    =   2145
   ClientLeft      =   1605
   ClientTop       =   2790
   ClientWidth     =   6480
   Icon            =   "TipoDeICM.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2145
   ScaleWidth      =   6480
   Begin VB.Frame frICM 
      Height          =   1635
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   6465
      Begin VB.TextBox txtC�digoDoPDV 
         Height          =   285
         Left            =   5820
         TabIndex        =   3
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox txtICM 
         Height          =   285
         Left            =   1200
         TabIndex        =   2
         Text            =   " 0,00"
         Top             =   1200
         Width           =   555
      End
      Begin VB.TextBox txtC�digo 
         Height          =   285
         Left            =   1200
         TabIndex        =   0
         Top             =   300
         Width           =   750
      End
      Begin VB.TextBox txtDescri��o 
         Height          =   285
         Left            =   1200
         TabIndex        =   1
         Top             =   750
         Width           =   5000
      End
      Begin VB.Label C�digoDoPDV 
         Caption         =   "C�digo do PDV"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4350
         TabIndex        =   10
         Top             =   1260
         Width           =   1605
      End
      Begin VB.Label lblICM 
         Caption         =   "ICM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   180
         TabIndex        =   9
         Top             =   1260
         Width           =   420
      End
      Begin VB.Label lblC�digo 
         Caption         =   "C�digo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   200
         Left            =   150
         TabIndex        =   8
         Top             =   330
         Width           =   855
      End
      Begin VB.Label lblDescri��o 
         Caption         =   "Descri��o"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   7
         Top             =   780
         Width           =   885
      End
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "&Gravar"
      Height          =   345
      Left            =   3885
      TabIndex        =   4
      Top             =   1740
      Width           =   1245
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   5205
      TabIndex        =   5
      Top             =   1740
      Width           =   1245
   End
End
Attribute VB_Name = "frmTipoDeICM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TBLTipoDeICM As Table
Dim TipoDeICMAberto As Boolean
Dim IndiceTipoDeICMAtivo$

Dim lAllowInsert  As Boolean
Dim lAllowEdit    As Boolean
Dim lAllowDelete  As Boolean
Dim lAllowConsult As Boolean

Dim lPula As Boolean
Dim lInserir As Boolean
Dim lAlterar As Boolean
Dim mFechar As Boolean

Dim StatusBarAviso$

Dim DataBaseName(1 To 1) As String
Public Relat�rio$
Public TotalDatabaseName%

Public lAtualizar As Boolean
Private Sub AtivaCampos()
    Bot�oImprimir True
    frICM.Enabled = True
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
        StatusBarAviso = "Inclus�o cancelada"
    End If
    If lAlterar Then
        StatusBarAviso = "Altera��o cancelada"
    End If
    BarraDeStatus StatusBarAviso
    
    lInserir = False
    lAlterar = False
    Bot�oIncluir lAllowInsert
    
    If TBLTipoDeICM.RecordCount = 0 Then
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
    
    TestaInferior TBLTipoDeICM, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLTipoDeICM, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Function
Private Sub DesativaCampos()
    Bot�oImprimir False
    frICM.Enabled = False
    cmdGravar.Enabled = False
    cmdCancelar.Enabled = False
    Bot�oGravar False
End Sub
Public Sub Encontrar()
    If Not lAllowConsult Then
        Exit Sub
    End If
    Set frmEncontrar.DBBancoDeDados = DBCadastro
    frmEncontrar.NomeDaJanela = "Tipo de ICM"
    frmEncontrar.LabelDescription = "Descri��o"
    frmEncontrar.Mensagem = "Nenhum Tipo de ICM foi selecionado!"
    frmEncontrar.BancoDeDados = "CADASTRO"
    frmEncontrar.Tabela = "TIPO DE ICM"
    frmEncontrar.Indice = "2"
    frmEncontrar.CampoChave = "C�DIGO"
    frmEncontrar.CampoPreencheLista = "DESCRI��O"
    frmEncontrar.Show vbModal
    lPula = True
    txtC�digo = frmEncontrar.Chave
    lPula = False
    PosRecords
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
    
    TBLTipoDeICM.Delete
    
    If Err <> 0 Then
        GeraMensagemDeErro "TipoDeICM - Excluir", True
        StatusBarAviso = "Falha na exclus�o"
        BarraDeStatus StatusBarAviso
        Exit Sub
    End If
    
    WS.CommitTrans
    
    Log gUsu�rio, "Exclus�o - Tipo de ICM: " & txtC�digo & " - " & txtDescri��o
    
    StatusBarAviso = "Exclus�o bem sucedida"
    BarraDeStatus StatusBarAviso
    
    If TBLTipoDeICM.RecordCount = 0 Then
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
    
    If TBLTipoDeICM.BOF Then
        TBLTipoDeICM.MoveFirst
    ElseIf TBLTipoDeICM.EOF Then
        TBLTipoDeICM.MoveLast
    Else
        TBLTipoDeICM.MovePrevious
        If TBLTipoDeICM.BOF Then
            TBLTipoDeICM.MoveNext
        End If
    End If
    
    GetRecords
    
    TestaInferior TBLTipoDeICM, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLTipoDeICM, lAllowEdit, lAllowDelete, lAllowConsult
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
        If TBLTipoDeICM.RecordCount > 0 And Not TBLTipoDeICM.BOF And Not TBLTipoDeICM.EOF Then
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
    
    TestaInferior TBLTipoDeICM, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLTipoDeICM, lAllowEdit, lAllowDelete, lAllowConsult
    
    If TBLTipoDeICM.RecordCount = 0 Then
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
    
    txtC�digo.SetFocus
End Sub
Public Sub MoveFirst()
    If lAlterar Then
       If Not Cancelamento Then
            Exit Sub
        End If
    End If
    
    StatusBarAviso = "Pronto"
    BarraDeStatus StatusBarAviso
    
    TBLTipoDeICM.MoveFirst
    
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
    
    StatusBarAviso = "Pronto"
    BarraDeStatus StatusBarAviso
    
    TBLTipoDeICM.MoveLast
    
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
    
    TBLTipoDeICM.MoveNext
    If TBLTipoDeICM.EOF Then
        TBLTipoDeICM.MovePrevious
        Exit Sub
    End If
    
    StatusBarAviso = "Pronto"
    BarraDeStatus StatusBarAviso
    
    Navega��oInferior lAllowConsult
    TestaSuperior TBLTipoDeICM, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Sub
Public Sub MovePrevious()
    If lAlterar Then
       If Not Cancelamento Then
            Exit Sub
        End If
    End If
    
    TBLTipoDeICM.MovePrevious
    If TBLTipoDeICM.BOF Then
        TBLTipoDeICM.MoveNext
        Exit Sub
    End If
    
    StatusBarAviso = "Pronto"
    BarraDeStatus StatusBarAviso
    
    Navega��oSuperior lAllowConsult
    TestaInferior TBLTipoDeICM, lAllowEdit, lAllowDelete, lAllowConsult
    
    GetRecords
End Sub
Public Sub PosRecords()
    If TBLTipoDeICM.RecordCount = 0 Then
        Exit Sub
    End If
    
    TBLTipoDeICM.Seek "=", txtC�digo
    If TBLTipoDeICM.NoMatch Then
        MsgBox "N�o consegui encontrar " + txtC�digo, vbExclamation, "Erro"
        TBLTipoDeICM.MoveFirst
        Navega��oInferior False
        Navega��oInferior lAllowConsult
    Else
        TestaInferior TBLTipoDeICM, lAllowEdit, lAllowDelete, lAllowConsult
        TestaSuperior TBLTipoDeICM, lAllowEdit, lAllowDelete, lAllowConsult
    End If
    GetRecords
End Sub
Private Sub GetRecords()
    On Error GoTo Erro
    
    If Not lAllowConsult Then
        ZeraCampos
        DesativaCampos
        Exit Sub
    End If
    
    txtC�digo = TBLTipoDeICM("C�DIGO")
    txtDescri��o = TBLTipoDeICM("DESCRI��O")
    txtICM = StrVal(TBLTipoDeICM("ICM"))
    txtICM_LostFocus
    txtC�digoDoPDV = TBLTipoDeICM("C�DIGO DO PDV")
    
    If Not lAllowEdit Then
        DesativaCampos
    End If
    
    Exit Sub
    
Erro:
    GeraMensagemDeErro "Tipo de ICM - GetRecords"
    Resume Next
End Sub
Private Function SetRecords()
    On Error GoTo Erro
    
    Dim Msg$
    Dim Confirma��o As Integer, Msg1$, Msg2$
    
    WS.BeginTrans 'Inicia uma Transa��o
    
    If lInserir Then
        TBLTipoDeICM.AddNew
    Else
        TBLTipoDeICM.Edit
    End If
    
    TBLTipoDeICM("C�DIGO") = txtC�digo
    TBLTipoDeICM("DESCRI��O") = txtDescri��o
    TBLTipoDeICM("ICM") = ValStr(txtICM)
    TBLTipoDeICM("C�DIGO DO PDV") = txtC�digoDoPDV
    If lInserir Then
        TBLTipoDeICM("USERNAME - CRIA") = gUsu�rio
        TBLTipoDeICM("DATA - CRIA") = Date
        TBLTipoDeICM("HORA - CRIA") = Time
        TBLTipoDeICM("USERNAME - ALTERA") = "VAZIO"
        TBLTipoDeICM("DATA - ALTERA") = vbNull
        TBLTipoDeICM("HORA - ALTERA") = vbNull
    End If
    If lAlterar Then
        TBLTipoDeICM("USERNAME - ALTERA") = gUsu�rio
        TBLTipoDeICM("DATA - ALTERA") = Date
        TBLTipoDeICM("HORA - ALTERA") = Time
    End If
    TBLTipoDeICM.Update
        
Erro:
    If Err <> 0 Then
        TBLTipoDeICM.CancelUpdate
        GeraMensagemDeErro "TipoDeICM - SetRecords", True
        SetRecords = False
        Exit Function
    End If

    WS.CommitTrans 'Grava as altera��es ou inclus�es se n�o houverem erros
    
    If lInserir Then
        Log gUsu�rio, "Inclus�o - Tipo de ICM: " & txtC�digo & " - " & txtDescri��o
    Else
        Log gUsu�rio, "Altera��o - Tipo de ICM: " & txtC�digo & " - " & txtDescri��o
    End If
    
    SetRecords = True
End Function
Private Sub ZeraCampos()
    txtC�digo = Empty
    txtDescri��o = Empty
    txtICM = "0,00"
    txtICM_LostFocus
    txtC�digoDoPDV = Empty
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
    If Not TipoDeICMAberto Then
        Unload Me
        Exit Sub
    End If
    
    TestaInferior TBLTipoDeICM, lAllowEdit, lAllowDelete, lAllowConsult
    TestaSuperior TBLTipoDeICM, lAllowEdit, lAllowDelete, lAllowConsult
    
    If TBLTipoDeICM.RecordCount = 0 Then
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
        Bot�oIncluir lAllowInsert
    Else
        Bot�oIncluir lAllowInsert
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
    
    lAllowInsert = Allow("TIPO DE ICM", "I")
    lAllowEdit = Allow("TIPO DE ICM", "A")
    lAllowDelete = Allow("TIPO DE ICM", "E")
    lAllowConsult = Allow("TIPO DE ICM", "C")
    
    ZeraCampos
    
    lPula = False
    lInserir = False
    lAlterar = False
    
    TipoDeICMAberto = AbreTabela(Dicion�rio, "CADASTRO", "TIPO DE ICM", DBCadastro, TBLTipoDeICM, TBLTabela, dbOpenTable)
    
    If TipoDeICMAberto Then
        IndiceTipoDeICMAtivo = "TIPODEICM1"
        TBLTipoDeICM.Index = IndiceTipoDeICMAtivo
    Else
        MsgBox "N�o consegui abrir a tabela 'Tipo De Embalagem' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    Bot�oIncluir lAllowInsert
 
    If TBLTipoDeICM.RecordCount = 0 Then
        DesativaCampos
        Bot�oExcluir False
        Bot�oGravar False
    Else
        AtivaCampos
        Bot�oExcluir lAllowDelete
        Bot�oGravar (lInserir Or lAllowEdit)
        GetRecords
    End If
    
    Navega��oInferior False
        
    If TBLTipoDeICM.RecordCount = 0 Or TBLTipoDeICM.RecordCount = 1 Then
        Navega��oSuperior False
    Else
        Navega��oInferior lAllowConsult
    End If
    
    StatusBarAviso = "Pronto"
    Relat�rio = AddPath(Aplica��oPath, "REPORT\TIPO DE ICM.RPT")
    TotalDatabaseName = 1
    DataBaseName(1) = AddPath(Aplica��oPath, "DATABASE\CADASTRO.MDB")
    mFechar = False
    Exit Sub
    
Erro:
    GeraMensagemDeErro "TipoDeICM - Load"
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
    
    Set frmTipoDeICM = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If TipoDeICMAberto Then
        TBLTipoDeICM.Close
    End If
    If Forms.Count = 2 Then
        AllBot�es False
    End If
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
Private Sub txtICM_Change()
    FormatMask "@K 99,99", txtICM
End Sub
Private Sub txtICM_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBarAviso = "Altera��o"
        BarraDeStatus StatusBarAviso
    End If
End Sub
Private Sub txtICM_LostFocus()
    FormatMask "@V #0,00", txtICM
End Sub

