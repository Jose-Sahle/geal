VERSION 5.00
Begin VB.Form frmImprimirEntrega 
   Caption         =   "Impressão de Entrega"
   ClientHeight    =   1080
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6165
   Icon            =   "frmImprimirEntrega.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1080
   ScaleWidth      =   6165
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frImpressora 
      Caption         =   "Impressora"
      Height          =   615
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6135
      Begin VB.ComboBox cmbImpressora 
         Height          =   315
         Left            =   150
         TabIndex        =   3
         Top             =   210
         Width           =   5865
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   345
      Left            =   3600
      TabIndex        =   1
      Top             =   690
      Width           =   1245
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   4890
      TabIndex        =   0
      Top             =   690
      Width           =   1245
   End
End
Attribute VB_Name = "frmImprimirEntrega"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim AllImpressoras() As Printer

Public Impressora As Printer
Private Sub cmdCancelar_Click()
    Set Impressora = Nothing
    Unload Me
End Sub
Private Sub cmdOK_Click()
    Set Impressora = AllImpressoras((cmbImpressora.ListIndex + 1))
    Unload Me
End Sub
Private Sub Form_Load()
    On Error GoTo Erro
    
    Dim Cont       As Byte
    Dim Impressora As Printer

    ReDim Preserve AllImpressoras(1 To Printers.Count)
    
    Cont = 0
    For Each Impressora In Printers
        Cont = Cont + 1
        cmbImpressora.AddItem Impressora.DeviceName
        Set AllImpressoras(Cont) = Impressora
    Next
    
    Exit Sub
    
Erro:
    GeraMensagemDeErro "Nota Fiscal - Load"
End Sub
