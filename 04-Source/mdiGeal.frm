VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm mdiGeal 
   BackColor       =   &H8000000C&
   Caption         =   "Geal - Gerenciador de Estoque-Administrativo"
   ClientHeight    =   2340
   ClientLeft      =   1410
   ClientTop       =   2925
   ClientWidth     =   9480
   Icon            =   "mdiGeal.frx":0000
   LinkTopic       =   "MDIForm1"
   NegotiateToolbars=   0   'False
   ScrollBars      =   0   'False
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   1
      Top             =   1995
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   609
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   8
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   3528
            MinWidth        =   3528
            Key             =   "Aviso"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
            Key             =   "Posi��o"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Alignment       =   1
            Object.Width           =   706
            MinWidth        =   706
            TextSave        =   "INS"
            Key             =   "INS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1057
            MinWidth        =   1057
            TextSave        =   "CAPS"
            Key             =   "CAPS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   4
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1057
            MinWidth        =   1057
            TextSave        =   "SCRL"
            Key             =   "SCRL"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            Object.Width           =   1057
            MinWidth        =   1057
            TextSave        =   "NUM"
            Key             =   "NUM"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "30/12/99"
            Key             =   "Data"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Bevel           =   2
            Object.Width           =   1235
            MinWidth        =   1235
            TextSave        =   "13:21"
            Key             =   "Hora"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageListInativo 
      Left            =   1830
      Top             =   510
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiGeal.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiGeal.frx":075E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiGeal.frx":0BB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiGeal.frx":1006
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiGeal.frx":145A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiGeal.frx":18AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiGeal.frx":1D02
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiGeal.frx":2156
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiGeal.frx":25AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiGeal.frx":29FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiGeal.frx":2E52
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiGeal.frx":32A6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageListDesabilitado 
      Left            =   2400
      Top             =   510
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiGeal.frx":3442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiGeal.frx":3896
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiGeal.frx":3CEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiGeal.frx":413E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiGeal.frx":4592
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiGeal.frx":49E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiGeal.frx":4E3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiGeal.frx":528E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiGeal.frx":56E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiGeal.frx":5B36
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiGeal.frx":5F8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiGeal.frx":63DE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageListInativo"
      DisabledImageList=   "ImageListDesabilitado"
      HotImageList    =   "ImageListAtivo"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   19
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Incluir"
            Object.ToolTipText     =   "Incluir novo registro"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Gravar"
            Object.ToolTipText     =   "Gravar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Excluir"
            Object.ToolTipText     =   "Apaga o registro corrente"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Imprimir"
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cortar"
            Object.ToolTipText     =   "Cortar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copiar"
            Object.ToolTipText     =   "Copiar"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Colar"
            Object.ToolTipText     =   "Colar"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "MoveFirst"
            Object.ToolTipText     =   "Primeiro registro"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "MovePrevious"
            Object.ToolTipText     =   "Pr�ximo registro"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "MoveNext"
            Object.ToolTipText     =   "Registro anterior"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "MoveLast"
            Object.ToolTipText     =   "�ltimo registro"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Calendario"
            Object.ToolTipText     =   "Faixa de data para selecionar registros"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Separador"
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageListAtivo 
      Left            =   1260
      Top             =   510
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiGeal.frx":657A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiGeal.frx":69CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiGeal.frx":6E22
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiGeal.frx":7282
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiGeal.frx":76E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiGeal.frx":7B42
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiGeal.frx":7F96
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiGeal.frx":83F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiGeal.frx":8856
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiGeal.frx":8CB6
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiGeal.frx":9116
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "mdiGeal.frx":9576
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   780
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuArquivo 
      Caption         =   "&Arquivo"
      Begin VB.Menu mnuArquivoAbrir 
         Caption         =   "&Abrir"
         Begin VB.Menu mnuArquivoAbrirAgencia 
            Caption         =   "&Ag�ncia..."
         End
         Begin VB.Menu mnuArquivoAbrirBanco 
            Caption         =   "&Banco..."
         End
         Begin VB.Menu mnuArquivoAbrirCliente 
            Caption         =   "&Cliente..."
         End
         Begin VB.Menu mnuArquivoAbrirContaCorrente 
            Caption         =   "Conta Co&rrente..."
         End
         Begin VB.Menu mnuArquivoAbrirFornecedor 
            Caption         =   "&Fornecedor..."
         End
         Begin VB.Menu mnuArquivoAbrirFuncion�rio 
            Caption         =   "F&uncion�rio..."
         End
         Begin VB.Menu mnuArquivoAbrirProduto 
            Caption         =   "&Produto..."
         End
         Begin VB.Menu mnuArquivoAbrirDespesas 
            Caption         =   "&Despesas..."
         End
      End
      Begin VB.Menu mnuArquivoFechar 
         Caption         =   "&Fechar"
      End
      Begin VB.Menu mnuArquivoFecharTudo 
         Caption         =   "Fechar &Tudo"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuArquivoSalvar 
         Caption         =   "&Salvar"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuArquivoConfigurarImpressora 
         Caption         =   "&Configurar Impressora..."
      End
      Begin VB.Menu mnuArquivoImprimir 
         Caption         =   "&Imprimir..."
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuArquivoBloquearSistema 
         Caption         =   "B&loquear Sistema"
      End
      Begin VB.Menu mnuArquivoConectarOutroUsu�rio 
         Caption         =   "Co&nectar a outro usu�rio..."
      End
      Begin VB.Menu mnuArquivoMudan��DeSenha 
         Caption         =   "&Mudan�a de Senha..."
      End
      Begin VB.Menu mnuArquivoSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuArquivoSair 
         Caption         =   "Sa&ir"
      End
   End
   Begin VB.Menu mnuEditar 
      Caption         =   "&Editar"
      Begin VB.Menu mnuEditarCortar 
         Caption         =   "&Cortar"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditarCopiar 
         Caption         =   "Co&piar"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditarColar 
         Caption         =   "Co&lar"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuSep5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditarEncontrar 
         Caption         =   "&Encontrar..."
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuSep6 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditarIncluir 
         Caption         =   "&Incluir"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuEditarExcluir 
         Caption         =   "E&xcluir"
      End
      Begin VB.Menu mnuEditarSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditarAtualizar 
         Caption         =   "A&tualizar"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuExibir 
      Caption         =   "E&xibir"
      Begin VB.Menu mnuExibirBarradeFerramentas 
         Caption         =   "Barra de &Ferramentas"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuExibirBarradeStatus 
         Caption         =   "Barra de &Status"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSep7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExibirData 
         Caption         =   "&Data"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuExibirHora 
         Caption         =   "&Hora"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu mnuNavega��o 
      Caption         =   "&Navega��o"
      Begin VB.Menu mnuNavega��oPrimeiroRegistro 
         Caption         =   "&Primeiro registro"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuNavega��oRegistroAnterior 
         Caption         =   "Registro &anterior"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuNavega��oPr�ximoRegistro 
         Caption         =   "Pr�ximo &registro"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuNavega��o�ltimoRegistro 
         Caption         =   "�ltimo registro"
         Shortcut        =   ^L
      End
   End
   Begin VB.Menu mnuMovimento 
      Caption         =   "&Movimento"
      Begin VB.Menu mnuMovimentoEntrada 
         Caption         =   "&Entrada de Produto"
         Begin VB.Menu mnuMovimentoEntradaCompra 
            Caption         =   "&Compra..."
         End
         Begin VB.Menu mnuMovimentoEntradaDevolu��oTroca 
            Caption         =   "&Devolu��o/Troca..."
         End
      End
      Begin VB.Menu mnuMovimentoSa�da 
         Caption         =   "&Sa�da de Produto"
         Begin VB.Menu mnuMovimentoSa�daVenda 
            Caption         =   "&Venda..."
         End
         Begin VB.Menu mnuMovimentoSa�daDevolu��oTroca 
            Caption         =   "&Devolu��o/Troca..."
         End
      End
      Begin VB.Menu mnuSep8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMovimentoDespesas 
         Caption         =   "Despesas..."
      End
      Begin VB.Menu mnuSep20 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMovimentoMovimentoDi�rio 
         Caption         =   "Movimento Di�rio..."
      End
      Begin VB.Menu mnuMovimentoContaCorrente 
         Caption         =   "&Conta Corrente..."
      End
      Begin VB.Menu mnuSep9 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMovimentoCaixa 
         Caption         =   "Cai&xa..."
      End
      Begin VB.Menu mnuMovimentoCaixaF�cil 
         Caption         =   "Caixa &F�cil..."
      End
      Begin VB.Menu mnuSep19 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMovimentoAberturaDoCaixa 
         Caption         =   "Abertura do Caixa..."
      End
      Begin VB.Menu mnuMovimentoFechamentoDoCaixa 
         Caption         =   "Fechamento do Caixa..."
      End
      Begin VB.Menu mnuPDV 
         Caption         =   "PDV..."
      End
   End
   Begin VB.Menu mnuPar�metros 
      Caption         =   "&Par�metros"
      Begin VB.Menu mnuPar�metrosDepartamento 
         Caption         =   "&Departamento..."
      End
      Begin VB.Menu mnuPar�metrosSe��o 
         Caption         =   "&Se��o..."
      End
      Begin VB.Menu mnuPar�metrosDepartamentoSe��o 
         Caption         =   "Depar&tamento - Se��o..."
      End
      Begin VB.Menu mnuSep11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPar�metrosTipodeICM 
         Caption         =   "Tipo de &ICM..."
      End
      Begin VB.Menu mnuPar�metrosTipodeEmbalagem 
         Caption         =   "Tipo de &Embalagem..."
      End
      Begin VB.Menu mnuSep15 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPar�metrosUnidades 
         Caption         =   "U&nidades..."
      End
      Begin VB.Menu mnuSep12 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPar�metrosLocalidadeDeEstoque 
         Caption         =   "&Localidade de Estoque..."
      End
      Begin VB.Menu mnuSep16 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPar�metrosPlanoDePagamento 
         Caption         =   "Pl&ano de Pagamento..."
      End
      Begin VB.Menu mnuPar�metroSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPar�metrosUsu�rios 
         Caption         =   "&Usu�rios..."
      End
      Begin VB.Menu mnuPar�metrosGrupos 
         Caption         =   "&Grupos..."
      End
      Begin VB.Menu mnuPar�metroSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPar�metrosCaixa 
         Caption         =   "&Caixa..."
      End
      Begin VB.Menu mnuPar�metroSep4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPar�metrosSenhaDoSistema 
         Caption         =   "Senha do Sistema..."
      End
   End
   Begin VB.Menu mnuUtilit�rios 
      Caption         =   "&Utilit�rios"
      Begin VB.Menu mnuUtilit�riosEnviarMensagem 
         Caption         =   "Enviar &Mensagem..."
      End
      Begin VB.Menu mnuUtilit�riosCaixadeEntrada 
         Caption         =   "Caixa de &Entrada..."
      End
      Begin VB.Menu mnuUtilit�riosCaixadeSa�da 
         Caption         =   "Caixa de &Sa�da..."
      End
      Begin VB.Menu mnuSep13 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAtualiza��oInternet 
         Caption         =   "&Atualiza��o via Internet..."
      End
      Begin VB.Menu mnuUtilitarioSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUtilit�riosAgendadeCompromisso 
         Caption         =   "Agenda de &Compromisso..."
      End
      Begin VB.Menu mnuAgendadeTelefone 
         Caption         =   "Agenda de &Telefone..."
      End
      Begin VB.Menu mnuSep18 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUtilit�riosConsultaSQL 
         Caption         =   "Consultas Gerenciais..."
      End
   End
   Begin VB.Menu mnuEntregas 
      Caption         =   "En&tregas"
      Begin VB.Menu mnuEntregasDefinir 
         Caption         =   "&Definir..."
      End
      Begin VB.Menu mnuEntregasManuten��o 
         Caption         =   "&Manuten��o..."
      End
   End
   Begin VB.Menu mnuJanela 
      Caption         =   "&Janela"
      WindowList      =   -1  'True
      Begin VB.Menu mnuJanelaEmCascata 
         Caption         =   "&Em Cascata"
      End
   End
   Begin VB.Menu mnuAjuda 
      Caption         =   "A&juda"
      Begin VB.Menu mnuAjudaConte�do 
         Caption         =   "Conte�do..."
      End
      Begin VB.Menu mnuAjudaProcurarPor 
         Caption         =   "&Procurar por..."
      End
      Begin VB.Menu mnuSep14 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAjudaSobreoGeal 
         Caption         =   "Sobre Geal..."
      End
   End
End
Attribute VB_Name = "mdiGeal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub MDIForm_Load()
    On Error GoTo Erro
        
    Caption = Caption & " - " & gUsu�rio
        
    gAPPNAME = "Geal"
    
    AllBot�es False
    
    glPortaAberta = False
        
    Exit Sub
    
Erro:
    GeraMensagemDeErro "mdiGeal - Load"
    Unload Me
End Sub
Private Sub MDIForm_Resize()
    If WindowState <> 1 Then
        gEstado = WindowState
    End If
    ResizeStatusBar
End Sub
Private Sub MDIForm_Unload(Cancel As Integer)
    If Dicion�rioAberto Then
        TBLArquivo.Close
        TBLTabela.Close
        TBLCampo.Close
        TBLIndice.Close
        DBDGeal.Close
    End If
        
    FechaBaseDeDados
            
    SetRegistryString "Geal", "Geral", "Usu�rio", gUsu�rio
    
    Set PDV = Nothing
    
    End
End Sub
Private Sub mnuAgendadeTelefone_Click()
    MsgBox "Em desenvolvimento!", vbInformation, "Aviso"
End Sub
Private Sub mnuAjudaSobreoGeal_Click()
    frmAbout.Show vbModal
End Sub
Private Sub mnuArquivoAbrirAgencia_Click()
    frmAg�ncia.Show 0
End Sub
Private Sub mnuArquivoAbrirBanco_Click()
    frmBanco.Show 0
End Sub
Private Sub mnuArquivoAbrirCliente_Click()
    frmCliente.Show 0
End Sub
Private Sub mnuArquivoAbrirContaCorrente_Click()
    frmContaCorrente.Show 0
End Sub
Private Sub mnuArquivoAbrirFornecedor_Click()
    frmFornecedor.Show 0
End Sub
Private Sub mnuArquivoAbrirFuncion�rio_Click()
    frmFuncion�rio.Show 0
End Sub
Private Sub mnuArquivoAbrirProduto_Click()
    frmProduto.Show 0
End Sub
Private Sub mnuArquivoAbrirDespesas_Click()
    frmDespesas.Show 0
End Sub
Private Sub mnuArquivoBloquearSistema_Click()
    On Error GoTo Erro
    
    'Valida Usu�rio
    Dim Cont As Byte, OldUsu�rio As String
    
    OldUsu�rio = gUsu�rio
    
    For Cont = 2 - 1 To Forms.Count - 1
        Forms(Cont).Hide
    Next
    
ValidaUsu�rio:
    frmValidaUsu�rio.GravaUsu�rio = True
    frmValidaUsu�rio.Bloquear = True
    frmValidaUsu�rio.WindowTop = 0
    frmValidaUsu�rio.WindowHeight = 0
    frmValidaUsu�rio.Show 1
    
    If frmValidaUsu�rio.Usu�rio <> Empty Then
        If frmValidaUsu�rio.Usu�rio = OldUsu�rio Then
            For Cont = 2 - 1 To Forms.Count - 1
                Forms(Cont).Show
            Next
        Else
            MsgBox "Esta esta��o somente pode ser desbloqueada pelo usu�rio " & OldUsu�rio, vbInformation, "Aviso"
            GoTo ValidaUsu�rio
        End If
    Else
        MsgBox "Esta esta��o somente pode ser desbloqueada pelo usu�rio " & OldUsu�rio, vbInformation, "Aviso"
        GoTo ValidaUsu�rio
    End If
    
    Set frmValidaUsu�rio = Nothing
    
    Exit Sub
    
Erro:
    MsgBox "Esta esta��o somente pode ser desbloqueada pelo usu�rio " & OldUsu�rio, vbInformation, "Aviso"
    GoTo ValidaUsu�rio
End Sub
Private Sub mnuArquivoConectarOutroUsu�rio_Click()
    On Error GoTo Erro
    
    'Valida Usu�rio
    Dim Cont As Byte, OldUsu�rio As String
    
    OldUsu�rio = gUsu�rio
    
    For Cont = 2 - 1 To Forms.Count - 1
        Forms(Cont).Hide
    Next
    
ValidaUsu�rio:
    frmValidaUsu�rio.GravaUsu�rio = True
    frmValidaUsu�rio.WindowTop = 0
    frmValidaUsu�rio.WindowHeight = 0
    frmValidaUsu�rio.Show 1
    
    If frmValidaUsu�rio.Usu�rio <> Empty Then
        If frmValidaUsu�rio.Usu�rio = OldUsu�rio Then
            For Cont = 2 - 1 To Forms.Count - 1
                Forms(Cont).Show
            Next
        Else
            mnuArquivoFecharTudo_Click
            If Err.Number <> 0 Then
                MsgBox "Esta esta��o somente pode ser desbloqueada pelo usu�rio " & OldUsu�rio, vbInformation, "Aviso"
                GoTo ValidaUsu�rio
            End If
            gUsu�rio = frmValidaUsu�rio.Usu�rio
            Caption = "Geal - Gerenciador de Estoque-Administrativo - " & gUsu�rio
            ChamaConfigura��es gUsu�rio
        End If
    Else
        If MsgBox("Deseja fechar o Sistema Geal?", vbInformation + vbYesNo, "Confirma��o") = vbYes Then
            mnuArquivoSair_Click
            If Err.Number <> 0 Then
                MsgBox "Esta esta��o somente pode ser desbloqueada pelo usu�rio " & OldUsu�rio, vbInformation, "Aviso"
                GoTo ValidaUsu�rio
            End If
        Else
            GoTo ValidaUsu�rio
        End If
    End If
    
    SetRegistryString "Geal", "Geral", "Usu�rio", gUsu�rio
    
    Set frmValidaUsu�rio = Nothing
    
    Exit Sub
    
Erro:
    MsgBox "Esta esta��o somente pode ser desbloqueada pelo usu�rio " & OldUsu�rio, vbInformation, "Aviso"
    GoTo ValidaUsu�rio
End Sub
Private Sub mnuArquivoConfigurarImpressora_Click()
    CommonDialog.Flags = cdlPDPrintSetup
    CommonDialog.Action = 5
End Sub
Private Sub mnuArquivoFechar_Click()
    Dim FormAtivo As Form
    
    If Forms.Count <= 1 Then Exit Sub
    
    Set FormAtivo = ActiveForm
    
    Unload FormAtivo
End Sub
Private Sub mnuArquivoFecharTudo_Click()
    On Error Resume Next
    
    Dim NomeDaJanela As String
    
    If Forms.Count <= 1 Then Exit Sub
    
    Do While Forms.Count > 1
        NomeDaJanela = Forms(1).Name
        Unload Forms(1)
        If Forms.Count <= 1 Then
            Exit Do
        Else
            If NomeDaJanela = Forms(1).Name Or Err.Number <> 0 Then
                Exit Do
            End If
        End If
        DoEvents
    Loop
End Sub
Private Sub mnuArquivoImprimir_Click()
    mdiGeal.ActiveForm.Imprimir
End Sub
Private Sub mnuArquivoMudan��DeSenha_Click()
    frmMudan�aDeSenha.Usu�rio = gUsu�rio
    frmMudan�aDeSenha.Show 1
End Sub
Private Sub mnuArquivoSair_Click()
    Unload Me
End Sub
Private Sub mnuArquivoSalvar_Click()
    If Forms.Count > 1 Then
        mdiGeal.ActiveForm.Gravar
    End If
End Sub
Private Sub mnuAtualiza��oInternet_Click()
    MsgBox "Em desenvolvimento!", vbInformation, "Aviso"
End Sub
Private Sub mnuEditarAtualizar_Click()
    If ActiveForm.lAtualizar Then
        ActiveForm.Atualizar
    End If
End Sub
Private Sub mnuEditarColar_Click()
    On Error GoTo Sair
    Dim Texto$, Cntl As Control, FormAtivo As Form
    
    Set FormAtivo = mdiGeal.ActiveForm
    Set Cntl = mdiGeal.ActiveForm.ActiveControl
    
    If TypeOf Cntl Is TextBox Then
        Texto = Cntl.Text
    End If
    
    Fun��oColar Cntl
    
    If TypeOf Cntl Is TextBox Then
        If Texto <> Cntl.Text Then
            FormAtivo.lAlterar = True
            FormAtivo.StatusBarAviso = "Altera��o"
            BarraDeStatus FormAtivo.StatusBarAviso
        End If
    End If
Sair:
End Sub
Private Sub mnuEditarCopiar_Click()
    On Error GoTo Sair
    Dim Texto$, Cntl As Control, FormAtivo As Form
    
    Set FormAtivo = mdiGeal.ActiveForm
    Set Cntl = mdiGeal.ActiveForm.ActiveControl
    
    If TypeOf Cntl Is TextBox Then
        Texto = Cntl.Text
    End If
    
    Fun��oCopiar Cntl
    
    If TypeOf Cntl Is TextBox Then
        If Texto <> Cntl.Text Then
            FormAtivo.lAlterar = True
            FormAtivo.StatusBarAviso = "Altera��o"
            BarraDeStatus FormAtivo.StatusBarAviso
        End If
    End If
Sair:
End Sub
Private Sub mnuEditarCortar_Click()
    On Error GoTo Sair
    Dim Texto$, Cntl As Control, FormAtivo As Form
    
    Set FormAtivo = mdiGeal.ActiveForm
    Set Cntl = mdiGeal.ActiveForm.ActiveControl
    
    If TypeOf Cntl Is TextBox Then
        Texto = Cntl.Text
    End If
    
    Fun��oCortar Cntl
    
    If TypeOf Cntl Is TextBox Then
        If Texto <> Cntl.Text Then
            FormAtivo.lAlterar = True
            FormAtivo.StatusBarAviso = "Altera��o"
            BarraDeStatus FormAtivo.StatusBarAviso
        End If
    End If
Sair:
End Sub
Private Sub mnuEditarEncontrar_Click()
    On Error Resume Next
    
    ActiveForm.Encontrar
End Sub
Private Sub mnuEditarExcluir_Click()
    If Forms.Count > 1 Then
        mdiGeal.ActiveForm.Excluir
    End If
End Sub
Private Sub mnuEditarIncluir_Click()
    If Forms.Count > 1 Then
        mdiGeal.ActiveForm.Incluir
    End If
End Sub
Private Sub mnuEntregasDefinir_Click()
    frmEntrega.Show 0
End Sub
Private Sub mnuExibirBarradeFerramentas_Click()
    If mnuExibirBarradeFerramentas.Checked = True Then
        mnuExibirBarradeFerramentas.Checked = False
        Toolbar.Visible = False
    Else
        mnuExibirBarradeFerramentas.Checked = True
        Toolbar.Visible = True
    End If
End Sub
Private Sub mnuExibirBarradeStatus_Click()
    If mnuExibirBarradeStatus.Checked = True Then
        mnuExibirBarradeStatus.Checked = False
        StatusBar.Visible = False
    Else
        mnuExibirBarradeStatus.Checked = True
        StatusBar.Visible = True
    End If
End Sub
Private Sub mnuExibirData_Click()
    If mnuExibirData.Checked = True Then
        mnuExibirData.Checked = False
        StatusBar.Panels("Data").Visible = False
    Else
        mnuExibirData.Checked = True
        StatusBar.Panels("Data").Visible = True
    End If
    ResizeStatusBar
End Sub
Private Sub mnuExibirHora_Click()
    If mnuExibirHora.Checked = True Then
        mnuExibirHora.Checked = False
        StatusBar.Panels("Hora").Visible = False
    Else
        mnuExibirHora.Checked = True
        StatusBar.Panels("Hora").Visible = True
    End If
    ResizeStatusBar
End Sub
Private Sub mnuJanelaEmCascata_Click()
    mdiGeal.Arrange vbCascade
End Sub
Private Sub mnuMovimentoAberturaDoCaixa_Click()
    frmAberturaDoCaixa.Show 1
End Sub
Private Sub mnuMovimentoCaixa_Click()
    frmCaixa.Show 1
End Sub
Private Sub mnuMovimentoCaixaF�cil_Click()
    frmCaixaF�cil.Show 0
End Sub
Private Sub mnuMovimentoContaCorrente_Click()
    frmMovimentoDeContaCorrente.Show 0
End Sub
Private Sub mnuMovimentoEntradaCompra_Click()
    frmCompra.Show 0
End Sub
Private Sub mnuMovimentoEntradaDevolu��oTroca_Click()
    frmCompraDevolu��oTroca.Show 0
End Sub
Private Sub mnuMovimentoFechamentoDoCaixa_Click()
    frmFechamentoDoCaixa.Show 1
End Sub
Private Sub mnuMovimentoSa�daDevolu��oTroca_Click()
    frmVendaDevolu��oTroca.Show 0
End Sub
Private Sub mnuMovimentoSa�daVenda_Click()
    frmVenda.Show 0
End Sub
Private Sub mnuMovimentoDespesas_Click()
    frmApontamentoDeDespesas.Show 0
End Sub
Private Sub mnuNavega��oPrimeiroRegistro_Click()
    If Forms.Count > 1 Then
        mdiGeal.ActiveForm.MoveFirst
    End If
End Sub
Private Sub mnuNavega��oPr�ximoRegistro_Click()
    If Forms.Count > 1 Then
        mdiGeal.ActiveForm.MoveNext
    End If
End Sub
Private Sub mnuNavega��oRegistroAnterior_Click()
    If Forms.Count > 1 Then
        mdiGeal.ActiveForm.MovePrevious
    End If
End Sub
Private Sub mnuNavega��o�ltimoRegistro_Click()
    If Forms.Count > 1 Then
        mdiGeal.ActiveForm.MoveLast
    End If
End Sub
Private Sub mnuPar�metrosCaixa_Click()
    frmPar�metrosCaixa.Show 1
End Sub
Private Sub mnuPar�metrosDepartamento_Click()
    frmDepartamento.Show 0
End Sub
Private Sub mnuPar�metrosDepartamentoSe��o_Click()
    frmDepartamentoSe��o.Show 0
End Sub
Private Sub mnuPar�metrosGrupos_Click()
    frmGrupos.Show 0
End Sub
Private Sub mnuPar�metrosLocalidadeDeEstoque_Click()
    frmLocal.Show 0
End Sub
Private Sub mnuPar�metrosPlanoDePagamento_Click()
    frmPlanoDePagamento.Show 0
End Sub
Private Sub mnuPar�metrosSe��o_Click()
    frmSe��o.Show 0
End Sub
Private Sub mnuPar�metrosSenhaDoSistema_Click()
    frmSenhaDoSistema.Show 1
End Sub
Private Sub mnuPar�metrosTipodeEmbalagem_Click()
    frmTipoDeEmbalagem.Show 0
End Sub
Private Sub mnuPar�metrosTipodeICM_Click()
    frmTipoDeICM.Show 0
End Sub
Private Sub mnuPar�metrosUnidades_Click()
    frmUnidades.Show 0
End Sub
Private Sub mnuPar�metrosUsu�rios_Click()
    If gUsu�rio = "ADMIN" Then
        frmUsu�rios.Show 1
    Else
        MsgBox "Acesso n�o permitido!", vbCritical, "Aviso"
    End If
End Sub
Private Sub mnuPDV_Click()
    frmPDV.Show 1
End Sub
Private Sub mnuUtilit�riosAgendadeCompromisso_Click()
    MsgBox "Em desenvolvimento!", vbInformation, "Aviso"
End Sub
Private Sub mnuUtilit�riosCaixadeEntrada_Click()
    MsgBox "Em desenvolvimento!", vbInformation, "Aviso"
End Sub
Private Sub mnuUtilit�riosCaixadeSa�da_Click()
    MsgBox "Em desenvolvimento!", vbInformation, "Aviso"
End Sub
Private Sub mnuUtilit�riosConsultaSQL_Click()
    frmConsultaSQL.Show 0
End Sub
Private Sub mnuUtilit�riosEnviarMensagem_Click()
    MsgBox "Em desenvolvimento!", vbInformation, "Aviso"
End Sub
Public Sub Mostrar()
    mdiGeal.WindowState = gEstado
End Sub
Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Key = "Incluir" Then
        mnuEditarIncluir_Click
    ElseIf Button.Key = "Excluir" Then
        mnuEditarExcluir_Click
    ElseIf Button.Key = "Gravar" Then
        mnuArquivoSalvar_Click
    ElseIf Button.Key = "Imprimir" Then
        mnuArquivoImprimir_Click
    ElseIf Button.Key = "Cortar" Then
        mnuEditarCortar_Click
    ElseIf Button.Key = "Copiar" Then
        mnuEditarCopiar_Click
    ElseIf Button.Key = "Colar" Then
        mnuEditarColar_Click
    ElseIf Button.Key = "MoveFirst" Then
        mnuNavega��oPrimeiroRegistro_Click
    ElseIf Button.Key = "MovePrevious" Then
        mnuNavega��oRegistroAnterior_Click
    ElseIf Button.Key = "MoveNext" Then
        mnuNavega��oPr�ximoRegistro_Click
    ElseIf Button.Key = "MoveLast" Then
        mnuNavega��o�ltimoRegistro_Click
    ElseIf Button.Key = "Internet" Then
        mnuAtualiza��oInternet_Click
    End If
    On Error Resume Next
    mdiGeal.ActiveForm.ActiveControl.SetFocus
    On Error GoTo 0
End Sub
