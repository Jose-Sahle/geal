VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmVisualizarConsultaSQL 
   Caption         =   "Visualização de Consulta SQL"
   ClientHeight    =   6300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7590
   Icon            =   "VisualizarConsulta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   7590
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdFechar 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   345
      Left            =   6300
      TabIndex        =   2
      Top             =   5910
      Width           =   1245
   End
   Begin VB.Data dtConsultaSQL 
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   510
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5880
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Timer Timer 
      Left            =   60
      Top             =   5910
   End
   Begin VB.Frame frConsultaSQL 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5865
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7575
      Begin MSDBGrid.DBGrid dbgrdConsultaSQL 
         Bindings        =   "VisualizarConsulta.frx":030A
         Height          =   5385
         Left            =   60
         OleObjectBlob   =   "VisualizarConsulta.frx":0326
         TabIndex        =   1
         Top             =   420
         Width           =   7425
      End
   End
End
Attribute VB_Name = "frmVisualizarConsultaSQL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lFechar As Boolean

Public Nome As String
Public Intervalo As Integer
Public SQL As String
Public BancoDeDados As String
Private Sub cmdFechar_Click()
    Unload Me
End Sub
Private Sub Form_Activate()
    If lFechar Then
        Unload Me
    End If
End Sub
Private Sub Form_Load()
    On Error GoTo Erro
    
    lFechar = False
    
    dtConsultaSQL.DataBaseName = BancoDeDados
    dtConsultaSQL.RecordSource = SQL
    dtConsultaSQL.Refresh
    
    dbgrdConsultaSQL.Refresh
    
    frConsultaSQL.Caption = Nome
    
    Timer.Interval = Intervalo
    If Intervalo > 0 Then
        Timer.Enabled = True
    End If
    
    Exit Sub
    
Erro:
    GeraMensagemDeErro "Visualizar Consulta"
    lFechar = True
End Sub
Private Sub Form_Resize()
    If WindowState = 1 Then
        Exit Sub
    End If
    
    cmdFechar.Top = Height - 795
    cmdFechar.Left = Width - 1410
    
    frConsultaSQL.Height = cmdFechar.Top - 45
    frConsultaSQL.Width = cmdFechar.Left + 1275
    
    dbgrdConsultaSQL.Height = cmdFechar.Top - 525
    dbgrdConsultaSQL.Width = cmdFechar.Left + 1125
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmVisualizarConsultaSQL = Nothing
End Sub
Private Sub Timer_Timer()
    dtConsultaSQL.Refresh
End Sub
