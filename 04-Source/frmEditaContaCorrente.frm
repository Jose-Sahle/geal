VERSION 5.00
Begin VB.Form frmEditaContaCorrente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edita Conta Corrente"
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   8400
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   7170
      TabIndex        =   2
      Top             =   2550
      Width           =   1245
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "&Gravar"
      Height          =   345
      Left            =   5850
      TabIndex        =   1
      Top             =   2550
      Width           =   1245
   End
   Begin VB.Frame frEditaContaCorrente 
      Height          =   2505
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8385
      Begin VB.TextBox txtValor 
         Height          =   315
         Left            =   5610
         TabIndex        =   9
         Top             =   2070
         Width           =   2415
      End
      Begin VB.ComboBox cmbDébitoCrédito 
         Height          =   315
         ItemData        =   "frmEditaContaCorrente.frx":0000
         Left            =   2070
         List            =   "frmEditaContaCorrente.frx":000A
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "Crédito"
         Top             =   210
         Width           =   1545
      End
      Begin VB.TextBox txtHistórico 
         Height          =   1095
         Left            =   690
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   870
         Width           =   7335
      End
      Begin VB.TextBox txtData 
         Height          =   315
         Left            =   690
         TabIndex        =   4
         Text            =   "  /  /"
         Top             =   210
         Width           =   915
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
         Height          =   195
         Left            =   4980
         TabIndex        =   8
         Top             =   2130
         Width           =   585
      End
      Begin VB.Label lblHistórico 
         Caption         =   "Histórico"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   210
         TabIndex        =   5
         Top             =   630
         Width           =   855
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
         Height          =   255
         Left            =   180
         TabIndex        =   3
         Top             =   240
         Width           =   465
      End
   End
End
Attribute VB_Name = "frmEditaContaCorrente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lPula As Boolean

Public lCancelado As Boolean
Public Data As String
Public Histórico As String
Public DébitoCrédito As String
Public Valor As Currency
Private Sub cmbDébitoCrédito_Click()
    cmbDébitoCrédito.Locked = True
End Sub
Private Sub cmbDébitoCrédito_DropDown()
    cmbDébitoCrédito.Locked = False
End Sub
Private Sub cmdCancelar_Click()
    lCancelado = True
    
    Unload Me
End Sub
Private Sub cmdGravar_Click()
    lCancelado = False
    Data = txtData
    Histórico = txtHistórico
    DébitoCrédito = Mid(cmbDébitoCrédito.Text, 1, 1)
    Valor = ValStr(txtValor)
    
    Unload Me
End Sub
Private Sub txtData_Change()
    FormatMask DataMask, txtData
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
    FormatMask "@K 99.999.999,99", txtValor
End Sub
Private Sub txtValor_LostFocus()
    FormatMask "@V ##.###.##0,00", txtValor
End Sub
