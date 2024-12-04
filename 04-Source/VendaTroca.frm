VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmVendaTroca 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Venda - Troca"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9540
   Icon            =   "VendaTroca.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6555
   ScaleWidth      =   9540
   Begin VB.Frame frDadosCadastrais 
      Height          =   1140
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   9540
      Begin VB.TextBox txtCliente 
         Height          =   300
         Left            =   1200
         TabIndex        =   18
         Top             =   690
         Width           =   5235
      End
      Begin VB.TextBox txtData 
         Height          =   285
         Left            =   8250
         TabIndex        =   17
         Text            =   "  /  /"
         Top             =   690
         Width           =   990
      End
      Begin VB.TextBox txtOrçamento 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   16
         Top             =   300
         Width           =   765
      End
      Begin VB.CommandButton cmdTabelaCliente 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6480
         TabIndex        =   15
         Top             =   660
         Width           =   375
      End
      Begin VB.Label lblCliente 
         Caption         =   "Cliente"
         Height          =   180
         Left            =   150
         TabIndex        =   21
         Top             =   720
         Width           =   645
      End
      Begin VB.Label lblData 
         Caption         =   "Data"
         Height          =   210
         Left            =   7680
         TabIndex        =   20
         Top             =   720
         Width           =   465
      End
      Begin VB.Label lblOrçamento 
         Caption         =   "Orçamento"
         Height          =   180
         Left            =   150
         TabIndex        =   19
         Top             =   330
         Width           =   825
      End
   End
   Begin VB.Frame frItens 
      Caption         =   " Itens "
      Height          =   3255
      Left            =   0
      TabIndex        =   12
      Top             =   1140
      Width           =   9540
      Begin MSDBGrid.DBGrid dbgrdItens 
         Height          =   2925
         Left            =   60
         OleObjectBlob   =   "VendaTroca.frx":030A
         TabIndex        =   13
         Top             =   210
         Width           =   9405
      End
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "&Gravar"
      Height          =   345
      Left            =   6990
      TabIndex        =   11
      Top             =   6165
      Width           =   1245
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   8310
      TabIndex        =   10
      Top             =   6165
      Width           =   1245
   End
   Begin VB.Frame frTotais 
      Caption         =   "Totais "
      Height          =   1695
      Left            =   0
      TabIndex        =   0
      Top             =   4410
      Width           =   9525
      Begin VB.TextBox txtValor 
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
         Left            =   7740
         TabIndex        =   5
         Top             =   150
         Width           =   1665
      End
      Begin VB.TextBox txtDesconto 
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
         Left            =   7740
         TabIndex        =   4
         Top             =   540
         Width           =   1665
      End
      Begin VB.TextBox txtValorTotal 
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
         Left            =   6750
         TabIndex        =   3
         Text            =   "R$"
         Top             =   1320
         Width           =   2655
      End
      Begin VB.TextBox txtValorBonus 
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
         Left            =   7740
         TabIndex        =   2
         Text            =   "         0,00"
         Top             =   930
         Width           =   1665
      End
      Begin VB.TextBox txtPorcentagemBonus 
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
         Left            =   6750
         TabIndex        =   1
         Text            =   "  0,00"
         Top             =   930
         Width           =   855
      End
      Begin VB.Label lblBonus 
         Caption         =   "Bonus"
         Height          =   195
         Left            =   6180
         TabIndex        =   9
         Top             =   990
         Width           =   495
      End
      Begin VB.Label lblDesconto 
         Caption         =   "Desconto"
         Height          =   255
         Left            =   6930
         TabIndex        =   8
         Top             =   630
         Width           =   1065
      End
      Begin VB.Label lblTotalGeral 
         Caption         =   "Total do Orçamento"
         Height          =   225
         Left            =   5280
         TabIndex        =   7
         Top             =   1350
         Width           =   1425
      End
      Begin VB.Label lblSubTotal 
         Caption         =   "Sub Total"
         Height          =   255
         Left            =   6930
         TabIndex        =   6
         Top             =   240
         Width           =   705
      End
   End
End
Attribute VB_Name = "frmVendaTroca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public lAtualizar As Boolean
Public Sub Refaz()

End Sub
