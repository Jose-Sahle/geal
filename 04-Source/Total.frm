VERSION 5.00
Begin VB.Form frmTotal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Total"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "Total.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   345
      Left            =   3390
      TabIndex        =   3
      Top             =   3030
      Width           =   1245
   End
   Begin VB.Frame frTroco 
      Caption         =   "Troco"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   0
      TabIndex        =   5
      Top             =   2130
      Width           =   4665
      Begin VB.Label lblTroco 
         Alignment       =   2  'Center
         Caption         =   "R$124,30"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   405
         Left            =   180
         TabIndex        =   9
         Top             =   240
         Width           =   4245
      End
   End
   Begin VB.Frame frRecebimento 
      Caption         =   "Recebimento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1185
      Left            =   0
      TabIndex        =   4
      Top             =   930
      Width           =   4665
      Begin VB.TextBox txtValor2 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1350
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   690
         Width           =   2895
      End
      Begin VB.TextBox txtValor1 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1350
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   270
         Width           =   2895
      End
      Begin VB.Label Label1 
         Caption         =   "Valor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   8
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lblValor1 
         Caption         =   "Valor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   7
         Top             =   300
         Width           =   615
      End
   End
   Begin VB.Frame frTotalAPagar 
      Caption         =   "Total a pagar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   4665
      Begin VB.Label lblTotalAPagar 
         Alignment       =   2  'Center
         Caption         =   "R$124,30"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   405
         Left            =   210
         TabIndex        =   6
         Top             =   300
         Width           =   4245
      End
   End
End
Attribute VB_Name = "frmTotal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lPula As Boolean

Public Total As String
Public ValorAPagar As Currency
Private Sub Troco()
    Dim Valor1 As Currency, Valor2 As Currency, Resultado As Currency
    
    Valor1 = ValStr(txtValor1)
    Valor2 = ValStr(txtValor2)
    
    Resultado = ValorAPagar - Valor1 - Valor2
    
    If Resultado < 0 Then
        Resultado = Resultado * (-1)
        lblTroco = "R$ " & Trim(FormatStringMask("@V #.###.###.##0,00", StrVal(Resultado)))
    End If
End Sub
Private Sub cmdOK_Click()
    Dim Valor1 As Currency, Valor2 As Currency
    
    Valor1 = ValStr(txtValor1)
    Valor2 = ValStr(txtValor2)
    
    If ValorAPagar > Valor1 + Valor2 Then
        MsgBox "Valores incorretos!", vbCritical, "Erro"
        txtValor1.SetFocus
        Exit Sub
    End If
    
    Total = StrTran(FormatStringMask("@V 0000000000,00", StrVal(Valor1 + Valor2)), ",")
    Unload Me
End Sub
Private Sub Form_Load()
    lPula = True
    lblTotalAPagar = "R$ " & Trim(FormatStringMask("@V #.###.###.##0,00", StrVal(ValorAPagar)))
    txtValor1 = FormatStringMask("@V #.###.###.##0,00", StrVal(0))
    txtValor2 = FormatStringMask("@V #.###.###.##0,00", StrVal(0))
    lblTroco = "R$ 0"
    lPula = False
End Sub
Private Sub txtValor1_Change()
    If Not lPula Then
        FormatMask "@K 9.999.999.999,99", txtValor1
    End If
End Sub
Private Sub txtValor1_LostFocus()
    lPula = True
    FormatMask "@V #.###.###.##0,00", txtValor1
    Troco
    lPula = False
End Sub
Private Sub txtValor2_Change()
    If Not lPula Then
        FormatMask "@K 9.999.999.999,99", txtValor2
    End If
End Sub
Private Sub txtValor2_LostFocus()
    lPula = True
    FormatMask "@V #.###.###.##0,00", txtValor2
    Troco
    lPula = False
End Sub

