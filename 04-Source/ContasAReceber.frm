VERSION 5.00
Begin VB.Form frmContasAReceber 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contas a receber"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8415
   Icon            =   "ContasAReceber.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4440
   ScaleWidth      =   8415
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   7170
      TabIndex        =   3
      Top             =   4080
      Width           =   1245
   End
   Begin VB.CommandButton cmdGravar 
      Caption         =   "&Gravar"
      Height          =   345
      Left            =   5850
      TabIndex        =   2
      Top             =   4080
      Width           =   1245
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   2925
      Left            =   0
      TabIndex        =   1
      Top             =   1110
      Width           =   8415
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8415
   End
End
Attribute VB_Name = "frmContasAReceber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public lAtualizar As Boolean
Private Sub Form_Activate()
    mdiGeal.Toolbar.Buttons("Calendario").Enabled = True
End Sub
Private Sub Form_Deactivate()
    mdiGeal.Toolbar.Buttons("Calendario").Enabled = False
End Sub
Private Sub Form_Load()
    mdiGeal.Toolbar.Buttons("Calendario").Visible = True
    mdiGeal.Toolbar.Buttons("Separador").Visible = True
End Sub
Private Sub Form_Unload(Cancel As Integer)
    mdiGeal.Toolbar.Buttons("Calendario").Visible = False
    mdiGeal.Toolbar.Buttons("Separador").Visible = False
End Sub
