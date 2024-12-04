VERSION 5.00
Begin VB.Form frmImpressão 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Impresão"
   ClientHeight    =   2295
   ClientLeft      =   4290
   ClientTop       =   3375
   ClientWidth     =   2640
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2295
   ScaleWidth      =   2640
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   345
      Left            =   30
      TabIndex        =   5
      Top             =   1845
      Width           =   1245
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   1350
      TabIndex        =   4
      Top             =   1845
      Width           =   1245
   End
   Begin VB.Frame frImpressão 
      Height          =   1740
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2640
      Begin VB.OptionButton optArquivo 
         Caption         =   "Arquivo"
         Height          =   195
         Left            =   390
         TabIndex        =   3
         Top             =   1240
         Width           =   1125
      End
      Begin VB.OptionButton optJanela 
         Caption         =   "Janela"
         Height          =   315
         Left            =   390
         TabIndex        =   2
         Top             =   740
         Width           =   1215
      End
      Begin VB.OptionButton optImpressora 
         Caption         =   "Impressora"
         Height          =   255
         Left            =   390
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmImpressão"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public vbOk As Boolean

Public Action%

Private Sub cmdCancelar_Click()
    vbOk = False
    Unload frmImpressão
End Sub


Private Sub cmdOk_Click()
    vbOk = True
    If optJanela Then
        Action = 0
    ElseIf optImpressora Then
        Action = 1
    ElseIf optArquivo Then
        Action = 2
    End If
    Unload frmImpressão
End Sub


Private Sub Form_Load()
    Move (Screen.Width - Width) / 2, (Screen.Height - Height) / 2
    optImpressora.Value = True
    vbOk = False
End Sub


