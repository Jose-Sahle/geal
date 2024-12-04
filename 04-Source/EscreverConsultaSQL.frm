VERSION 5.00
Begin VB.Form frmEscreverConsultaSQL 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7365
   Icon            =   "EscreverConsultaSQL.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   7365
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   6090
      TabIndex        =   8
      Top             =   5730
      Width           =   1245
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "&Gravar"
      Height          =   345
      Left            =   4800
      TabIndex        =   7
      Top             =   5730
      Width           =   1245
   End
   Begin VB.Frame frSQL 
      Caption         =   "Seqüência SQL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   0
      TabIndex        =   6
      Top             =   1470
      Width           =   7365
      Begin VB.TextBox txtSQL 
         BackColor       =   &H00C0C0C0&
         Height          =   3885
         Left            =   90
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   7185
      End
   End
   Begin VB.Frame frEscreverConsultaSQL 
      Height          =   1455
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   7365
      Begin VB.TextBox txtIntervalo 
         BackColor       =   &H00C0C0C0&
         Height          =   345
         Left            =   1140
         TabIndex        =   1
         Top             =   840
         Width           =   465
      End
      Begin VB.TextBox txtDescrição 
         BackColor       =   &H00C0C0C0&
         Height          =   345
         Left            =   1140
         TabIndex        =   0
         Top             =   300
         Width           =   5745
      End
      Begin VB.Label lblSegundos 
         Caption         =   "segundos"
         Height          =   225
         Left            =   1710
         TabIndex        =   9
         Top             =   960
         Width           =   765
      End
      Begin VB.Label lblIntervalo 
         Caption         =   "Intervalo"
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
         Left            =   180
         TabIndex        =   5
         Top             =   930
         Width           =   825
      End
      Begin VB.Label lblDescrição 
         Caption         =   "Descrição"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   4
         Top             =   360
         Width           =   1035
      End
   End
End
Attribute VB_Name = "frmEscreverConsultaSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TBLConsultaSQL As Recordset
Dim IndiceConsultaSQLAtivo As String
Dim ConsultaSQLAberto As Boolean

Public Nome As String
Public Tipo As Byte
Private Sub Grava()
    If SetRecords Then
        Unload Me
    End If
End Sub
Private Sub GetRecords()
    TBLConsultaSQL.Seek "=", Nome
    
    txtDescrição = TBLConsultaSQL("DESCRIÇÃO")
    txtIntervalo = StrVal((TBLConsultaSQL("INTERVALO") / 1000))
    txtSQL = TBLConsultaSQL("CONSULTA")
End Sub
Private Function SetRecords()
    On Error GoTo Erro
    
    
    If Tipo = vbIncluir Then
        TBLConsultaSQL.AddNew
    Else
        TBLConsultaSQL.Edit
    End If
    
    TBLConsultaSQL("DESCRIÇÃO") = txtDescrição
    TBLConsultaSQL("INTERVALO") = ValStr(txtIntervalo * 1000)
    TBLConsultaSQL("CONSULTA") = txtSQL
    
    TBLConsultaSQL.Update
    
    SetRecords = True
    
    Exit Function
    
Erro:
    TBLConsultaSQL.CancelUpdate
    GeraMensagemDeErro "EscreveConsultaSQL - SetRecords"
    SetRecords = False
End Function
Private Sub ZeraCampos()
    txtDescrição = Empty
    txtIntervalo = Empty
    txtSQL = Empty
End Sub
Private Sub cmdCancelar_Click()
    Unload Me
End Sub
Private Sub cmdGravar_Click()
    Grava
End Sub
Private Sub Form_Activate()
    If Not ConsultaSQLAberto Then
        Unload Me
    End If
    
    AllBotões False
    
    NavegaçãoInferior False
    NavegaçãoSuperior False
    
    txtDescrição.SetFocus
End Sub
Private Sub Form_Load()
    ConsultaSQLAberto = AbreTabela(Dicionário, "UTILITÁRIO", "CONSULTA", DBUtilitário, TBLConsultaSQL, TBLTabela, dbOpenTable)
    
    If ConsultaSQLAberto Then
        IndiceConsultaSQLAtivo = "CONSULTA1"
        TBLConsultaSQL.Index = IndiceConsultaSQLAtivo
    Else
        MsgBox "Não consegui abrir a tabela 'Consulta' !", vbCritical, "Erro"
        Exit Sub
    End If
    
    If Tipo = vbIncluir Then
        Caption = "Incluisão de Consulta SQL"
        ZeraCampos
    Else
        Caption = "Edição de Consulta SQL - " & Nome
        GetRecords
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmEscreverConsultaSQL = Nothing
End Sub
Private Sub txtDescrição_Change()
    FormatMask "@S30", txtDescrição
End Sub
Private Sub txtIntervalo_Change()
    FormatMask "99", txtIntervalo
End Sub
Private Sub txtIntervalo_LostFocus()
    FormatMask "@N #0", txtIntervalo
End Sub
Private Sub txtSQL_Change()
    FormatMask "@!", txtSQL
End Sub
