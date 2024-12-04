VERSION 5.00
Begin VB.Form frmImpressaoDeCheque 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impressão de Cheques"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6780
   Icon            =   "frmImpressaoDeCheque.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   6780
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdFechar 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   345
      Left            =   5460
      TabIndex        =   6
      Top             =   1620
      Width           =   1245
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   345
      Left            =   4140
      TabIndex        =   5
      Top             =   1620
      Width           =   1245
   End
   Begin VB.TextBox txtDadosDoVerso2 
      Height          =   315
      Left            =   2460
      TabIndex        =   4
      Top             =   1200
      Width           =   4245
   End
   Begin VB.TextBox txtDadosDoVerso1 
      Height          =   315
      Left            =   2460
      TabIndex        =   3
      Top             =   840
      Width           =   4245
   End
   Begin VB.TextBox txtData 
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
      Left            =   5370
      TabIndex        =   2
      Text            =   "99/99/9999"
      Top             =   450
      Width           =   1335
   End
   Begin VB.TextBox txtValor 
      Height          =   315
      Left            =   2460
      TabIndex        =   1
      Top             =   450
      Width           =   2175
   End
   Begin VB.TextBox txtBanco 
      Height          =   315
      Left            =   2460
      TabIndex        =   0
      Top             =   90
      Width           =   645
   End
   Begin VB.Label lblDadosDoVerso 
      Caption         =   "Dados do Verso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   555
      Left            =   1590
      TabIndex        =   10
      Top             =   870
      Width           =   795
   End
   Begin VB.Label lblData 
      Caption         =   "Data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   225
      Left            =   4770
      TabIndex        =   9
      Top             =   510
      Width           =   525
   End
   Begin VB.Label lblValor 
      Caption         =   "Valor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   225
      Left            =   1590
      TabIndex        =   8
      Top             =   510
      Width           =   615
   End
   Begin VB.Label lblBanco 
      Caption         =   "Banco"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   225
      Left            =   1590
      TabIndex        =   7
      Top             =   150
      Width           =   615
   End
   Begin VB.Image imgCheque 
      BorderStyle     =   1  'Fixed Single
      Height          =   1335
      Left            =   90
      Picture         =   "frmImpressaoDeCheque.frx":030A
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1410
   End
End
Attribute VB_Name = "frmImpressaoDeCheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lPula As Boolean
Private Sub cmdFechar_Click()
    Unload Me
End Sub
Private Sub cmdImprimir_Click()
    Dim Banco As String
    Dim Valor As String
    Dim Data As String
    Dim DadosAdicionais As String
    
    Banco = LeftBlankString(Trim(txtBanco), 3)
    
    Valor = StrTran(txtValor, ".")
    Valor = StrTran(Valor, ",")
    Valor = LeftBlankString(Trim(Valor), 12)
    
    Data = Trim(StrTran(txtData, "/"))
    
    DadosAdicionais = RightBlankString(Trim(txtDadosDoVerso1), 60) & RightBlankString(Trim(txtDadosDoVerso2), 60)
    
    If Not ImpressaoDeCheque(Banco, Valor, Data, DadosAdicionais) Then
        MsgBox "Falha na impressão do cheque", vbCritical, "Erro"
    End If
End Sub
Private Sub Form_Load()
    lPula = True
    txtBanco.Text = Empty
    txtValor.Text = Empty
    txtData.Text = "  /  /    "
    txtDadosDoVerso1.Text = Empty
    txtDadosDoVerso2.Text = Empty
    lPula = False
End Sub
Private Sub txtBanco_Change()
    If Not lPula Then
        FormatMask "999", txtBanco
    End If
End Sub
Private Sub txtDadosDoVerso1_Change()
    FormatMask "@!", txtDadosDoVerso1
End Sub
Private Sub txtDadosDoVerso2_Change()
    FormatMask "@!", txtDadosDoVerso2
End Sub
Private Sub txtData_Change()
    If Not lPula Then
        lPula = True
        FormatMask DataMask, txtData
        lPula = False
    End If
End Sub
Private Sub txtData_LostFocus()
    If StrTran(txtData.Text, "/") <> Space(8) Then
        lPula = True
        CorrigeData DataMask, txtData, Date
        lPula = False
        If Not FormatMask(CheckDataMask, txtData) Then
            Beep
            MsgBox "Data inválida !", vbCritical, "Erro"
            txtData.SelStart = 0
            txtData.SetFocus
        End If
    End If
End Sub
Private Sub txtValor_Change()
    If Not lPula Then
        FormatMask "@K 99.999.999,99", txtValor
    End If
End Sub
Private Sub txtValor_LostFocus()
    lPula = True
    FormatMask "@V ##.###.##0,00", txtValor
    lPula = False
End Sub
