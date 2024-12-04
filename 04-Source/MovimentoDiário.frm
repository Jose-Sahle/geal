VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmMovimentoDiário 
   Caption         =   "Movimento Diário"
   ClientHeight    =   6165
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8355
   Icon            =   "MovimentoDiário.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   8355
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   7110
      TabIndex        =   2
      Top             =   5790
      Width           =   1245
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "&Gravar"
      Height          =   345
      Left            =   5790
      TabIndex        =   1
      Top             =   5790
      Width           =   1245
   End
   Begin VB.Frame fr 
      Height          =   5715
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8355
      Begin VB.Frame frBanco 
         Caption         =   "Banco"
         Height          =   795
         Left            =   2370
         TabIndex        =   5
         Top             =   4830
         Width           =   2205
      End
      Begin VB.Frame frCaixa 
         Caption         =   "Caixa"
         Height          =   795
         Left            =   90
         TabIndex        =   4
         Top             =   4830
         Width           =   2205
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Height          =   4665
         Left            =   60
         OleObjectBlob   =   "MovimentoDiário.frx":030A
         TabIndex        =   3
         Top             =   150
         Width           =   8235
      End
   End
End
Attribute VB_Name = "frmMovimentoDiário"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
