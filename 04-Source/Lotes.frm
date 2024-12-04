VERSION 5.00
Begin VB.Form frmLotes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lotes"
   ClientHeight    =   1980
   ClientLeft      =   1965
   ClientTop       =   1560
   ClientWidth     =   3480
   Icon            =   "Lotes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   1980
   ScaleWidth      =   3480
   Begin VB.Frame frLotes 
      Height          =   1560
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   3405
      Begin VB.TextBox txtMúltiplo 
         Height          =   285
         Left            =   1200
         TabIndex        =   2
         Text            =   "1"
         Top             =   645
         Width           =   765
      End
      Begin VB.VScrollBar vscrMúltiplo 
         Height          =   315
         Left            =   1965
         Max             =   9999
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   630
         Value           =   1
         Width           =   210
      End
      Begin VB.TextBox txtDígito 
         Height          =   285
         Left            =   2670
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   270
         Width           =   285
      End
      Begin VB.VScrollBar vscrQuantidade 
         Height          =   315
         Left            =   1965
         Max             =   9999
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1015
         Value           =   1
         Width           =   210
      End
      Begin VB.TextBox txtCódigo 
         Height          =   285
         Left            =   1200
         TabIndex        =   0
         Top             =   270
         Width           =   1455
      End
      Begin VB.TextBox txtQuantidade 
         Height          =   285
         Left            =   1200
         TabIndex        =   4
         Text            =   "1"
         Top             =   1030
         Width           =   765
      End
      Begin VB.Label lblQuantidade 
         Caption         =   "Quantidade"
         Height          =   165
         Left            =   150
         TabIndex        =   13
         Top             =   1090
         Width           =   855
      End
      Begin VB.Label lblQuadrados 
         Caption         =   "2"
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
         Left            =   2550
         TabIndex        =   12
         Top             =   660
         Width           =   225
      End
      Begin VB.Label lblMetros 
         Caption         =   "m"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2310
         TabIndex        =   11
         Top             =   600
         Width           =   285
      End
      Begin VB.Label lblCódigo 
         Caption         =   "Código"
         Height          =   225
         Left            =   150
         TabIndex        =   10
         Top             =   330
         Width           =   900
      End
      Begin VB.Label lblMúltiplo 
         Caption         =   "Múltiplo"
         Height          =   195
         Left            =   150
         TabIndex        =   9
         Top             =   705
         Width           =   840
      End
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "&Gravar"
      Default         =   -1  'True
      Height          =   345
      Left            =   870
      TabIndex        =   6
      Top             =   1605
      Width           =   1245
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   2190
      TabIndex        =   7
      Top             =   1605
      Width           =   1245
   End
End
Attribute VB_Name = "frmLotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const MAXCOL = 4

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

Dim ArrayLoteTotal%
Dim ArrayLote() As Variant

Public lAtualizar As Boolean
Public mÚltimoDígito As Byte
Public mJanela As Form
Private Sub AdelLote(ByVal Elemento%)
    Dim Cont As Integer, Cont1 As Byte
    
    For Cont = Elemento To UBound(ArrayLote, 2) - 1
        For Cont1 = 1 To MAXCOL
            ArrayLote(Cont1, Cont) = ArrayLote(Cont1, Cont + 1)
        Next
    Next
    
    For Cont = 1 To MAXCOL
        ArrayLote(Cont, UBound(ArrayLote, 2)) = Empty
    Next
End Sub
Private Sub AtivaCampos()
    frLotes.Enabled = True
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
    
    If ArrayLoteTotal = 0 Then
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
    
    TestaInferiorArray Elemento, ArrayLote(), lAllowEdit, lAllowDelete, lAllowConsult, 2
    TestaSuperiorArray Elemento, ArrayLote(), lAllowEdit, lAllowDelete, lAllowConsult, 2
    
    GetRecords
End Function
Private Sub ClearArray()
    ArrayLoteTotal = 0
    ReDim ArrayLote(MAXCOL, 1)
End Sub
Private Sub DesativaCampos()
    frLotes.Enabled = False
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
    
    BarraDeStatus "Exclusão de Lote do Produto"

    AdelLote Elemento
    
    ArrayLoteTotal = ArrayLoteTotal - 1
    
    If Elemento > ArrayLoteTotal Then
        Elemento = ArrayLoteTotal
    End If
    
    lAlteração = True
    
    If ArrayLoteTotal = 0 Then
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
    
    SizeArray ArrayLoteTotal
    
    Log gUsuário, "Exclusão - Lote do Produto: " & txtCódigo
    
    GetRecords
    
    TestaInferiorArray Elemento, ArrayLote(), lAllowEdit, lAllowDelete, lAllowConsult, 2
    TestaSuperiorArray Elemento, ArrayLote(), lAllowEdit, lAllowDelete, lAllowConsult, 2
End Sub
Private Sub FillArray()
    Dim Cont%
    
    SizeArray mJanela.ArrayLoteTotal
    ArrayLoteTotal = mJanela.ArrayLoteTotal
    
    For Cont = 1 To mJanela.ArrayLoteTotal
        ArrayLote(1, Cont) = mJanela.GetArrayLote(1, Cont)
        ArrayLote(2, Cont) = mJanela.GetArrayLote(2, Cont)
        ArrayLote(3, Cont) = mJanela.GetArrayLote(3, Cont)
        ArrayLote(4, Cont) = mJanela.GetArrayLote(4, Cont)
    Next
End Sub
Public Sub Gravar()
    If lInserir Then
        ArrayLoteTotal = ArrayLoteTotal + 1
        Elemento = ArrayLoteTotal
        SizeArray ArrayLoteTotal
        SetRecords
        PosRecords
        BarraDeStatus "Inclusão bem sucedida"
        lInserir = False
        lAlteração = True
    ElseIf lAlterar Then
        If ArrayLoteTotal > 0 Then
            SetRecords
            PosRecords
            lAlterar = False
            lAlteração = True
            BarraDeStatus "Alteração bem sucedida"
        End If
    Else
        Exit Sub
    End If
    
    TestaInferiorArray Elemento, ArrayLote(), lAllowEdit, lAllowDelete, lAllowConsult, 2
    TestaSuperiorArray Elemento, ArrayLote(), lAllowEdit, lAllowDelete, lAllowConsult, 2
    
    If ArrayLoteTotal = 0 Then
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
    StatusBar = mJanela.StatusBarAviso
    
    If txtMúltiplo.Enabled Then
        txtMúltiplo.SetFocus
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
    
    txtCódigo = Format(Date, "ddmmyy")
    
    txtDígito = mÚltimoDígito + 1
    
    BotãoGravar (lInserir Or lAllowEdit)
    BotãoIncluir False
    cmdGravar.Enabled = (lInserir Or lAllowEdit)
    cmdCancelar.Enabled = (lInserir Or lAllowEdit)
    
    NavegaçãoInferior False
    NavegaçãoSuperior False
    
    txtMúltiplo.SetFocus
End Sub
Public Sub MoveFirst()
    If lAlterar Then
       If Not Cancelamento Then
            Exit Sub
        End If
    End If
        
    StatusBar = mJanela.StatusBarAviso
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
    
    StatusBar = mJanela.StatusBarAviso
    BarraDeStatus StatusBar
    
    Elemento = ArrayLoteTotal
    
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
    
    StatusBar = mJanela.StatusBarAviso
    BarraDeStatus StatusBar
    
    Elemento = Elemento + 1
    
    If Elemento > ArrayLoteTotal Then
        Elemento = ArrayLoteTotal
        Exit Sub
    End If
    
    NavegaçãoInferior lAllowConsult
    TestaSuperiorArray Elemento, ArrayLote(), lAllowEdit, lAllowDelete, lAllowConsult, 2
    
    GetRecords
End Sub
Public Sub MovePrevious()
    If lAlterar Then
       If Not Cancelamento Then
            Exit Sub
        End If
    End If
    
    StatusBar = mJanela.StatusBarAviso
    BarraDeStatus StatusBar
    
    Elemento = Elemento - 1
    
    If Elemento < 1 Then
        Elemento = 1
        Exit Sub
    End If
    
    NavegaçãoSuperior lAllowConsult
    TestaInferiorArray Elemento, ArrayLote(), lAllowEdit, lAllowDelete, lAllowConsult, 2
    
    GetRecords
End Sub
Public Sub PosRecords()
    txtMúltiplo.SetFocus
End Sub
Private Sub GetRecords()
    On Error GoTo Erro
    
    If Not lAllowConsult Then
        ZeraCampos
        DesativaCampos
        Exit Sub
    End If
    txtCódigo = ArrayLote(1, Elemento)
    txtDígito = ArrayLote(2, Elemento)
    txtMúltiplo = FormatStringMask("@V ###0,00", ArrayLote(3, Elemento))
    txtQuantidade = FormatStringMask("@V ###0,00", ArrayLote(4, Elemento))
    If Not lAllowEdit Then
        DesativaCampos
    End If
    
    Exit Sub
    
Erro:
    GeraMensagemDeErro "Lotes - GetRecords "
    Resume Next
End Sub
Private Function SetRecords()
    ArrayLote(1, Elemento) = txtCódigo
    ArrayLote(2, Elemento) = txtDígito
    ArrayLote(3, Elemento) = txtMúltiplo
    ArrayLote(4, Elemento) = txtQuantidade
    
    Log gUsuário, "Lote do Produto " & txtCódigo
    If lInserir Then
        mÚltimoDígito = mÚltimoDígito + 1
    End If
End Function
Private Sub SaveArray()
    Dim Cont%, Cont1%
    
    mJanela.SizeArrayLote ArrayLoteTotal
    mJanela.ArrayLoteTotal = ArrayLoteTotal
    
    For Cont = 1 To mJanela.ArrayLoteTotal
        For Cont1 = 1 To MAXCOL
            mJanela.SetArrayLote Cont1, ArrayLote(Cont1, Cont), Cont
        Next
    Next
End Sub
Private Sub SizeArray(ByVal Tamanho As Integer)
    If Tamanho > 0 Then
        ReDim Preserve ArrayLote(MAXCOL, Tamanho)
    End If
End Sub
Private Sub ZeraCampos()
    txtCódigo = Empty
    txtDígito = Empty
    txtMúltiplo = FormatStringMask("@V ###0,00", "1")
    txtQuantidade = FormatStringMask("@V ###0,00", "1")
End Sub
Private Sub cmdCancelar_Click()
    Cancelamento
End Sub
Private Sub cmdGravar_Click()
    Gravar
End Sub
Private Sub Form_Activate()
    TestaInferiorArray Elemento, ArrayLote(), lAllowEdit, lAllowDelete, lAllowConsult, 2
    TestaSuperiorArray Elemento, ArrayLote(), lAllowEdit, lAllowDelete, lAllowConsult, 2
    
    BarraDeStatus StatusBar
    
    If ArrayLoteTotal = 0 Then
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
    frmLotes.SetFocus
End Sub
Private Sub Form_Load()
    lPula = False
    
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    
    lAllowInsert = Allow("PRODUTO", "I")
    lAllowEdit = Allow("PRODUTO", "A")
    lAllowDelete = Allow("PRODUTO", "E")
    lAllowConsult = Allow("PRODUTO", "C")
    
    FillArray
        
    BotãoIncluir lAllowInsert
    
    StatusBar = mJanela.StatusBarAviso
 
    If ArrayLoteTotal = 0 Then
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
        
    If ArrayLoteTotal = 0 Or ArrayLoteTotal = 1 Then
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
            mJanela.mÚltimoDígito = mÚltimoDígito
            If Not mJanela.lAlterarArrayLote Then
                mJanela.lAlterarArrayLote = True
                If Not mJanela.lInserir Then
                    mJanela.lAlterar = True
                End If
            End If
        ElseIf Confirmação = vbCancel Then
            Cancel = 1
        ElseIf Confirmação = vbNo Then
        End If
    End If
    
    mJanela.AtualizaQuantidade
    
    Set frmLotes = Nothing
End Sub
Private Sub txtCódigo_Change()
    If Not lPula Then
        FormatMask "@K 99999999", txtCódigo
    End If
End Sub
Private Sub txtDígito_Change()
    If Not lPula Then
        FormatMask "@K 99", txtDígito
    End If
End Sub
Private Sub txtMúltiplo_Change()
    If Not lPula Then
        FormatMask "@K 9999,99", txtMúltiplo
        lPula = True
        vscrMúltiplo.Value = Int(ValStr(txtMúltiplo))
        lPula = False
    End If
End Sub
Private Sub txtMúltiplo_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBar = "Alteração de Lote do Produto"
        BarraDeStatus StatusBar
    End If
End Sub
Private Sub txtMúltiplo_LostFocus()
    FormatMask "@V ###0,00", txtMúltiplo
End Sub
Private Sub txtQuantidade_Change()
    If Not lPula Then
        FormatMask "@K 9999,99", txtQuantidade
        lPula = True
        vscrQuantidade.Value = Int(ValStr(txtQuantidade))
        lPula = False
    End If
End Sub
Private Sub txtQuantidade_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBar = "Alteração de Lote do Produto"
        BarraDeStatus StatusBar
    End If
End Sub
Private Sub txtQuantidade_LostFocus()
    FormatMask "@V ###0,00", txtQuantidade
End Sub
Private Sub vscrMúltiplo_Change()
    Dim NumeroDecimal As Single
    
    If Not lPula Then
        NumeroDecimal = ValStr(txtMúltiplo) - Int(ValStr(txtMúltiplo))
        txtMúltiplo = FormatStringMask("@V ###0,00", StrVal((vscrMúltiplo.Value + NumeroDecimal)))
        If Not lInserir Then
            lAlterar = True
            StatusBar = "Alteração de Lote do Produto"
            BarraDeStatus StatusBar
        End If
    End If
End Sub
Private Sub vscrMúltiplo_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBar = "Alteração de Lote do Produto"
        BarraDeStatus StatusBar
    End If
End Sub
Private Sub vscrQuantidade_Change()
    Dim NumeroDecimal As Single
    
    If Not lPula Then
        NumeroDecimal = ValStr(txtQuantidade) - Int(ValStr(txtQuantidade))
        txtQuantidade = FormatStringMask("@V ###0,00", StrVal((vscrQuantidade.Value + NumeroDecimal)))
        If Not lInserir Then
            lAlterar = True
            StatusBar = "Alteração de Lote do Produto"
            BarraDeStatus StatusBar
        End If
    End If
End Sub
Private Sub vscrQuantidade_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBar = "Alteração de Lote do Produto"
        BarraDeStatus StatusBar
    End If
End Sub
