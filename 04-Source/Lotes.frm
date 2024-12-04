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
      Begin VB.TextBox txtM�ltiplo 
         Height          =   285
         Left            =   1200
         TabIndex        =   2
         Text            =   "1"
         Top             =   645
         Width           =   765
      End
      Begin VB.VScrollBar vscrM�ltiplo 
         Height          =   315
         Left            =   1965
         Max             =   9999
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   630
         Value           =   1
         Width           =   210
      End
      Begin VB.TextBox txtD�gito 
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
      Begin VB.TextBox txtC�digo 
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
      Begin VB.Label lblC�digo 
         Caption         =   "C�digo"
         Height          =   225
         Left            =   150
         TabIndex        =   10
         Top             =   330
         Width           =   900
      End
      Begin VB.Label lblM�ltiplo 
         Caption         =   "M�ltiplo"
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
Public lAltera��o As Boolean

Dim ArrayLoteTotal%
Dim ArrayLote() As Variant

Public lAtualizar As Boolean
Public m�ltimoD�gito As Byte
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
    
    If ArrayLoteTotal = 0 Then
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
    
    BarraDeStatus "Exclus�o de Lote do Produto"

    AdelLote Elemento
    
    ArrayLoteTotal = ArrayLoteTotal - 1
    
    If Elemento > ArrayLoteTotal Then
        Elemento = ArrayLoteTotal
    End If
    
    lAltera��o = True
    
    If ArrayLoteTotal = 0 Then
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
    
    SizeArray ArrayLoteTotal
    
    Log gUsu�rio, "Exclus�o - Lote do Produto: " & txtC�digo
    
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
        BarraDeStatus "Inclus�o bem sucedida"
        lInserir = False
        lAltera��o = True
    ElseIf lAlterar Then
        If ArrayLoteTotal > 0 Then
            SetRecords
            PosRecords
            lAlterar = False
            lAltera��o = True
            BarraDeStatus "Altera��o bem sucedida"
        End If
    Else
        Exit Sub
    End If
    
    TestaInferiorArray Elemento, ArrayLote(), lAllowEdit, lAllowDelete, lAllowConsult, 2
    TestaSuperiorArray Elemento, ArrayLote(), lAllowEdit, lAllowDelete, lAllowConsult, 2
    
    If ArrayLoteTotal = 0 Then
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
    StatusBar = mJanela.StatusBarAviso
    
    If txtM�ltiplo.Enabled Then
        txtM�ltiplo.SetFocus
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
    
    txtC�digo = Format(Date, "ddmmyy")
    
    txtD�gito = m�ltimoD�gito + 1
    
    Bot�oGravar (lInserir Or lAllowEdit)
    Bot�oIncluir False
    cmdGravar.Enabled = (lInserir Or lAllowEdit)
    cmdCancelar.Enabled = (lInserir Or lAllowEdit)
    
    Navega��oInferior False
    Navega��oSuperior False
    
    txtM�ltiplo.SetFocus
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
    
    StatusBar = mJanela.StatusBarAviso
    BarraDeStatus StatusBar
    
    Elemento = ArrayLoteTotal
    
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
    
    StatusBar = mJanela.StatusBarAviso
    BarraDeStatus StatusBar
    
    Elemento = Elemento + 1
    
    If Elemento > ArrayLoteTotal Then
        Elemento = ArrayLoteTotal
        Exit Sub
    End If
    
    Navega��oInferior lAllowConsult
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
    
    Navega��oSuperior lAllowConsult
    TestaInferiorArray Elemento, ArrayLote(), lAllowEdit, lAllowDelete, lAllowConsult, 2
    
    GetRecords
End Sub
Public Sub PosRecords()
    txtM�ltiplo.SetFocus
End Sub
Private Sub GetRecords()
    On Error GoTo Erro
    
    If Not lAllowConsult Then
        ZeraCampos
        DesativaCampos
        Exit Sub
    End If
    txtC�digo = ArrayLote(1, Elemento)
    txtD�gito = ArrayLote(2, Elemento)
    txtM�ltiplo = FormatStringMask("@V ###0,00", ArrayLote(3, Elemento))
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
    ArrayLote(1, Elemento) = txtC�digo
    ArrayLote(2, Elemento) = txtD�gito
    ArrayLote(3, Elemento) = txtM�ltiplo
    ArrayLote(4, Elemento) = txtQuantidade
    
    Log gUsu�rio, "Lote do Produto " & txtC�digo
    If lInserir Then
        m�ltimoD�gito = m�ltimoD�gito + 1
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
    txtC�digo = Empty
    txtD�gito = Empty
    txtM�ltiplo = FormatStringMask("@V ###0,00", "1")
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
        
    Bot�oIncluir lAllowInsert
    
    StatusBar = mJanela.StatusBarAviso
 
    If ArrayLoteTotal = 0 Then
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
        
    If ArrayLoteTotal = 0 Or ArrayLoteTotal = 1 Then
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
            mJanela.m�ltimoD�gito = m�ltimoD�gito
            If Not mJanela.lAlterarArrayLote Then
                mJanela.lAlterarArrayLote = True
                If Not mJanela.lInserir Then
                    mJanela.lAlterar = True
                End If
            End If
        ElseIf Confirma��o = vbCancel Then
            Cancel = 1
        ElseIf Confirma��o = vbNo Then
        End If
    End If
    
    mJanela.AtualizaQuantidade
    
    Set frmLotes = Nothing
End Sub
Private Sub txtC�digo_Change()
    If Not lPula Then
        FormatMask "@K 99999999", txtC�digo
    End If
End Sub
Private Sub txtD�gito_Change()
    If Not lPula Then
        FormatMask "@K 99", txtD�gito
    End If
End Sub
Private Sub txtM�ltiplo_Change()
    If Not lPula Then
        FormatMask "@K 9999,99", txtM�ltiplo
        lPula = True
        vscrM�ltiplo.Value = Int(ValStr(txtM�ltiplo))
        lPula = False
    End If
End Sub
Private Sub txtM�ltiplo_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBar = "Altera��o de Lote do Produto"
        BarraDeStatus StatusBar
    End If
End Sub
Private Sub txtM�ltiplo_LostFocus()
    FormatMask "@V ###0,00", txtM�ltiplo
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
        StatusBar = "Altera��o de Lote do Produto"
        BarraDeStatus StatusBar
    End If
End Sub
Private Sub txtQuantidade_LostFocus()
    FormatMask "@V ###0,00", txtQuantidade
End Sub
Private Sub vscrM�ltiplo_Change()
    Dim NumeroDecimal As Single
    
    If Not lPula Then
        NumeroDecimal = ValStr(txtM�ltiplo) - Int(ValStr(txtM�ltiplo))
        txtM�ltiplo = FormatStringMask("@V ###0,00", StrVal((vscrM�ltiplo.Value + NumeroDecimal)))
        If Not lInserir Then
            lAlterar = True
            StatusBar = "Altera��o de Lote do Produto"
            BarraDeStatus StatusBar
        End If
    End If
End Sub
Private Sub vscrM�ltiplo_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBar = "Altera��o de Lote do Produto"
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
            StatusBar = "Altera��o de Lote do Produto"
            BarraDeStatus StatusBar
        End If
    End If
End Sub
Private Sub vscrQuantidade_KeyPress(KeyAscii As Integer)
    If Not lInserir Then
        lAlterar = True
        StatusBar = "Altera��o de Lote do Produto"
        BarraDeStatus StatusBar
    End If
End Sub
