VERSION 5.00
Begin VB.Form frmEncontrar 
   Caption         =   "Fornecedor"
   ClientHeight    =   3150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6630
   Icon            =   "Encontrar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   6630
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   345
      Left            =   1770
      TabIndex        =   4
      Top             =   2700
      Width           =   2835
   End
   Begin VB.Frame frFornecedor 
      Height          =   2565
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6615
      Begin VB.ListBox lstEncontrar 
         Height          =   1815
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   6345
      End
      Begin VB.TextBox txtNomeRazãoSocial 
         Height          =   285
         Left            =   1230
         TabIndex        =   1
         Top             =   240
         Width           =   5235
      End
      Begin VB.Label lblNomeRazãoSocial 
         Caption         =   "Razão Social"
         Height          =   195
         Left            =   150
         TabIndex        =   2
         Top             =   270
         Width           =   1065
      End
   End
End
Attribute VB_Name = "frmEncontrar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TBLEncontra As Table
Dim EncontraAberto As Boolean
Dim IndiceEncontraAtivo$
Dim ChaveEncontrar()

Dim ArrayCampoChave()
Dim TamanhoCampoChave As Byte

Dim ArrayCampoPreencheLista()
Dim TamanhoCampoPreencheLista As Byte

Dim lPula As Boolean
Dim mFechar As Boolean

Public DBBancoDeDados As Database
Public LabelDescription$
Public NomeDaJanela$
Public CampoChave$
Public CampoPreencheLista$
Public CampoPrincipal$
Public Indice$
Public Mensagem$
Public Chave$
Public BancoDeDados$
Public Nome$
Public Tabela$
Public Inicio%
Public Fim%
Private Sub Finalizar()
    If lstEncontrar.ListIndex = -1 Then
        MsgBox Mensagem, , "Aviso"
        Exit Sub
    End If
    Unload Me
End Sub
Private Sub cmdOK_Click()
    Finalizar
End Sub
Private Sub Form_Activate()
    If Not EncontraAberto Or mFechar Then
        Unload Me
    End If
End Sub
Private Sub Form_Load()
    On Error Resume Next
    
    Dim Elemento%
    Dim Aux$, Cont As Byte
    
    mFechar = False
    
    If NomeDaJanela <> "" Then
        Caption = NomeDaJanela
    End If
    
    If LabelDescription <> Empty Then
        lblNomeRazãoSocial = LabelDescription
    End If
        
    EncontraAberto = AbreTabela(Dicionário, BancoDeDados, Tabela, DBBancoDeDados, TBLEncontra, TBLTabela, dbOpenTable)
    
    If EncontraAberto Then
        IndiceEncontraAtivo = StrTran(Tabela, " ") & Indice
        TBLEncontra.Index = IndiceEncontraAtivo
    Else
        MsgBox "Não consegui abrir a tabela " & "'" & BancoDeDados & "'" & "!", vbCritical, "Erro"
        Exit Sub
    End If
    
    If TBLEncontra.EOF Or TBLEncontra.BOF Then
        MsgBox "Não existe nenhum registro cadastrado nesta tabela!", vbExclamation, "Aviso"
        mFechar = True
        Exit Sub
    End If
    
    'Campo chave
    TamanhoCampoChave = 1
    Aux = GetWordSeparatedBy(CampoChave, TamanhoCampoChave, ",")
    Do While Aux <> ""
        ReDim Preserve ArrayCampoChave(1 To TamanhoCampoChave)
        ArrayCampoChave(TamanhoCampoChave) = Aux
        TamanhoCampoChave = TamanhoCampoChave + 1
        Aux = GetWordSeparatedBy(CampoChave, TamanhoCampoChave, ",")
    Loop
    TamanhoCampoChave = TamanhoCampoChave - 1
    
    'Campo que preenche o List Box
    TamanhoCampoPreencheLista = 1
    Aux = GetWordSeparatedBy(CampoPreencheLista, TamanhoCampoPreencheLista, ",")
    Do While Aux <> ""
        ReDim Preserve ArrayCampoPreencheLista(1 To TamanhoCampoPreencheLista)
        ArrayCampoPreencheLista(TamanhoCampoPreencheLista) = Aux
        TamanhoCampoPreencheLista = TamanhoCampoPreencheLista + 1
        Aux = GetWordSeparatedBy(CampoPreencheLista, TamanhoCampoPreencheLista, ",")
    Loop
    TamanhoCampoPreencheLista = TamanhoCampoPreencheLista - 1
    
    TBLEncontra.MoveFirst
    Elemento = 0
    
    Do While Not TBLEncontra.EOF
        ReDim Preserve ChaveEncontrar(0 To Elemento)
        
        Aux = Empty
        For Cont = 1 To TamanhoCampoChave
             Aux = Aux & TBLEncontra(ArrayCampoChave(Cont)) & ","
        Next
        Aux = Left(Aux, Len(Aux) - 1)
        ChaveEncontrar(Elemento) = Aux
        
        Aux = Empty
        For Cont = 1 To TamanhoCampoPreencheLista
             Aux = Aux & TBLEncontra(ArrayCampoPreencheLista(Cont)) & " - "
        Next
        Aux = Left(Aux, Len(Aux) - 3)
        lstEncontrar.AddItem Aux
        
        TBLEncontra.MoveNext
        Elemento = Elemento + 1
    Loop
    
    If CampoPrincipal <> Empty Then
        txtNomeRazãoSocial.Text = CampoPrincipal
    Else
        lstEncontrar.ListIndex = -1
        txtNomeRazãoSocial = Empty
    End If
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not mFechar Then
        If lstEncontrar.ListIndex = -1 Then
            MsgBox Mensagem, , "Aviso"
            Cancel = 1
            Exit Sub
        End If
    End If
    If EncontraAberto Then
        TBLEncontra.Close
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If mFechar Then Exit Sub
    If lstEncontrar.ListIndex = -1 Then
        MsgBox Mensagem, , "Aviso"
        Cancel = 1
        Exit Sub
    End If
    Chave = ChaveEncontrar(lstEncontrar.ListIndex)
    Nome = lstEncontrar.List(lstEncontrar.ListIndex)
End Sub
Private Sub lstEncontrar_Click()
    If lPula Then
        Exit Sub
    End If
    lPula = True
    txtNomeRazãoSocial = lstEncontrar.List(lstEncontrar.ListIndex)
    lPula = False
End Sub
Private Sub txtNomeRazãoSocial_Change()
    Dim Cont%, Encontrou As Boolean
    
    If lPula Then
        Exit Sub
    End If
    
    lPula = True
    
    Encontrou = False
    
    For Cont = 0 To lstEncontrar.ListCount - 1
        If InStr(UCase(lstEncontrar.List(Cont)), UCase(txtNomeRazãoSocial)) = 1 Then
            Encontrou = True
            lstEncontrar.ListIndex = Cont
            Exit For
        End If
    Next
    
    If Not Encontrou Then
        lstEncontrar.ListIndex = -1
    End If
    
    lPula = False
End Sub
