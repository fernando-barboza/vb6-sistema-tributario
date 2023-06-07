VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmCadEmpresa 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Entidade"
   ClientHeight    =   6285
   ClientLeft      =   2730
   ClientTop       =   1680
   ClientWidth     =   8370
   HelpContextID   =   25
   Icon            =   "CadEmpresa.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   8370
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab tab_3dEmpresa 
      Height          =   6195
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   8265
      _ExtentX        =   14579
      _ExtentY        =   10927
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      ForeColor       =   -2147483641
      TabCaption(0)   =   " Entidade "
      TabPicture(0)   =   "CadEmpresa.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblstrNomeFantasia"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblstrNome"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblstrInscricaoEstadual"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblstrCGC"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fra_Endereco"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtstrInscricaoEstadual"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtstrCGC"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtstrNome"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtstrNomeFantasia"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "fra_Brasao"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "fra_LogoTipo"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "fra_FaixaCEP"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "Parâmetros"
      TabPicture(1)   =   "CadEmpresa.frx":105E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "frm_Diretorios"
      Tab(1).Control(1)=   "Fra_Febraban"
      Tab(1).Control(2)=   "fra_Certidoes"
      Tab(1).Control(3)=   "fra_Alvaras"
      Tab(1).Control(4)=   "Frame1"
      Tab(1).ControlCount=   5
      Begin VB.Frame fra_FaixaCEP 
         Caption         =   "Faixa de CEP"
         Height          =   735
         Left            =   120
         TabIndex        =   35
         Top             =   5160
         Width           =   3825
         Begin VB.TextBox txtintcepfinal 
            Height          =   285
            Left            =   2550
            TabIndex        =   39
            Top             =   300
            Width           =   945
         End
         Begin VB.TextBox txtintcepinicial 
            Height          =   285
            Left            =   690
            MaxLength       =   9
            TabIndex        =   37
            Top             =   300
            Width           =   945
         End
         Begin VB.Label lblfinal 
            AutoSize        =   -1  'True
            Caption         =   "Final"
            Height          =   195
            Left            =   2130
            TabIndex        =   38
            Top             =   330
            Width           =   330
         End
         Begin VB.Label lblinicial 
            AutoSize        =   -1  'True
            Caption         =   "Inicial"
            Height          =   195
            Left            =   180
            TabIndex        =   36
            Top             =   330
            Width           =   405
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Cadastro de Logradouro"
         Height          =   675
         Left            =   -74790
         TabIndex        =   61
         Top             =   4350
         Width           =   6705
         Begin VB.CheckBox chk_bytLogradouro 
            Caption         =   "Obrigatoriedade: Tipo da Lei / Nº do Processo / Lei de aprovação / Exercicío"
            Height          =   195
            Left            =   210
            TabIndex        =   62
            Top             =   330
            Width           =   6345
         End
      End
      Begin VB.Frame fra_Alvaras 
         Caption         =   "Alvarás"
         Height          =   675
         Left            =   -74790
         TabIndex        =   58
         Top             =   3660
         Width           =   6705
         Begin VB.TextBox txtintnumeroalvarafuncionamento 
            Height          =   285
            Left            =   2700
            MaxLength       =   4
            TabIndex        =   60
            Top             =   270
            Width           =   855
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Numero Alvará de Funcionamento:"
            Height          =   195
            Left            =   195
            TabIndex        =   59
            Top             =   345
            Width           =   2490
         End
      End
      Begin VB.Frame fra_Certidoes 
         Caption         =   "Certidões"
         Height          =   1005
         Left            =   -74790
         TabIndex        =   51
         Top             =   2580
         Width           =   6705
         Begin VB.TextBox txtintcertidaomobiliario 
            Height          =   285
            Left            =   2355
            MaxLength       =   4
            TabIndex        =   55
            Top             =   600
            Width           =   855
         End
         Begin VB.TextBox txtINTNUMEROGUIAPOSITIVA 
            Height          =   285
            Left            =   5655
            MaxLength       =   4
            TabIndex        =   57
            Top             =   300
            Width           =   855
         End
         Begin VB.TextBox txtINTNUMEROGUIANEGATIVA 
            Height          =   285
            Left            =   2370
            MaxLength       =   4
            TabIndex        =   53
            Top             =   270
            Width           =   855
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Número do Cadastro Mobiliario"
            Height          =   195
            Left            =   150
            TabIndex        =   54
            Top             =   675
            Width           =   2160
         End
         Begin VB.Label lblIntCertidadaoPositiva 
            AutoSize        =   -1  'True
            Caption         =   "Número Certidão Positiva"
            Height          =   195
            Left            =   3780
            TabIndex        =   56
            Top             =   345
            Width           =   1785
         End
         Begin VB.Label lblIntCertidadaoBaixa 
            AutoSize        =   -1  'True
            Caption         =   "Número Certidão Negativa"
            Height          =   195
            Left            =   195
            TabIndex        =   52
            Top             =   345
            Width           =   1875
         End
      End
      Begin VB.Frame Fra_Febraban 
         Caption         =   "Febraban"
         Height          =   765
         Left            =   -74790
         TabIndex        =   40
         Top             =   390
         Width           =   1575
         Begin VB.TextBox txtintFebraban 
            Height          =   285
            Left            =   840
            MaxLength       =   4
            TabIndex        =   42
            Top             =   300
            Width           =   495
         End
         Begin VB.Label lblfebraban 
            AutoSize        =   -1  'True
            Caption         =   "Código"
            Height          =   195
            Left            =   195
            TabIndex        =   41
            Top             =   345
            Width           =   495
         End
      End
      Begin VB.Frame fra_LogoTipo 
         Caption         =   " Logotipo "
         Height          =   1365
         Left            =   5715
         TabIndex        =   10
         Top             =   1560
         Width           =   1125
         Begin VB.TextBox txtintLogotipo 
            Height          =   285
            Left            =   30
            TabIndex        =   65
            Top             =   600
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.Image img_LogoTipo 
            BorderStyle     =   1  'Fixed Single
            Height          =   1110
            Left            =   0
            MouseIcon       =   "CadEmpresa.frx":107A
            MousePointer    =   99  'Custom
            Stretch         =   -1  'True
            Top             =   240
            Width           =   1110
         End
      End
      Begin VB.Frame fra_Brasao 
         Caption         =   " Brasão "
         Height          =   1365
         Left            =   1545
         TabIndex        =   9
         Top             =   1560
         Width           =   1125
         Begin VB.TextBox txtintBrasao 
            Height          =   285
            Left            =   30
            TabIndex        =   64
            Top             =   600
            Visible         =   0   'False
            Width           =   1065
         End
         Begin VB.Image img_Brasao 
            BorderStyle     =   1  'Fixed Single
            Height          =   1110
            Left            =   0
            MouseIcon       =   "CadEmpresa.frx":1384
            MousePointer    =   99  'Custom
            Stretch         =   -1  'True
            Top             =   240
            Width           =   1110
         End
      End
      Begin VB.TextBox txtstrNomeFantasia 
         Height          =   285
         Left            =   1545
         MaxLength       =   60
         TabIndex        =   4
         Top             =   840
         Width           =   6555
      End
      Begin VB.TextBox txtstrNome 
         Height          =   285
         Left            =   1545
         MaxLength       =   60
         TabIndex        =   2
         Top             =   480
         Width           =   6555
      End
      Begin VB.TextBox txtstrCGC 
         Height          =   285
         Left            =   1545
         MaxLength       =   18
         TabIndex        =   6
         Top             =   1200
         Width           =   2370
      End
      Begin VB.TextBox txtstrInscricaoEstadual 
         Height          =   285
         Left            =   5730
         MaxLength       =   15
         TabIndex        =   8
         Top             =   1230
         Width           =   2370
      End
      Begin VB.Frame fra_Endereco 
         Caption         =   " Endereço "
         ForeColor       =   &H80000007&
         Height          =   2100
         Left            =   120
         TabIndex        =   11
         Top             =   2925
         Width           =   7980
         Begin VB.CommandButton cmd_UF 
            Height          =   300
            Left            =   7500
            Picture         =   "CadEmpresa.frx":168E
            Style           =   1  'Graphical
            TabIndex        =   28
            TabStop         =   0   'False
            Tag             =   "1276"
            ToolTipText     =   "Ativa Cadastro de Unidades Federativas"
            Top             =   990
            Width           =   330
         End
         Begin VB.CommandButton cmd_Cidade 
            Height          =   300
            Left            =   5460
            Picture         =   "CadEmpresa.frx":1A18
            Style           =   1  'Graphical
            TabIndex        =   25
            TabStop         =   0   'False
            Tag             =   "53"
            ToolTipText     =   "Ativa Cadastro de Cidade"
            Top             =   990
            Width           =   330
         End
         Begin VB.CommandButton cmd_Bairro 
            Height          =   300
            Left            =   5460
            Picture         =   "CadEmpresa.frx":1DA2
            Style           =   1  'Graphical
            TabIndex        =   20
            TabStop         =   0   'False
            Tag             =   "581"
            ToolTipText     =   "Ativa Cadastro de Bairro"
            Top             =   630
            Width           =   330
         End
         Begin VB.CommandButton cmd_Logradouro 
            Height          =   300
            Left            =   6600
            Picture         =   "CadEmpresa.frx":212C
            Style           =   1  'Graphical
            TabIndex        =   16
            TabStop         =   0   'False
            Tag             =   "584"
            ToolTipText     =   "Ativa Cadastro de Logradouro"
            Top             =   240
            Width           =   330
         End
         Begin VB.CommandButton cmd_Tipo 
            Height          =   300
            Left            =   1740
            Picture         =   "CadEmpresa.frx":24B6
            Style           =   1  'Graphical
            TabIndex        =   14
            TabStop         =   0   'False
            Tag             =   "582"
            ToolTipText     =   "Ativa Cadastro de Logradouro"
            Top             =   240
            Width           =   330
         End
         Begin MSDataListLib.DataCombo dbcintCidade 
            Height          =   315
            Left            =   1005
            TabIndex        =   24
            Top             =   990
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            Text            =   ""
         End
         Begin VB.TextBox txtstrEmail 
            Height          =   285
            Left            =   1005
            MaxLength       =   45
            TabIndex        =   30
            Top             =   1350
            Width           =   4770
         End
         Begin VB.TextBox txtstrTelefone 
            Height          =   285
            Left            =   1005
            MaxLength       =   17
            TabIndex        =   32
            Top             =   1680
            Width           =   1605
         End
         Begin VB.TextBox txtintCep 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   6900
            MaxLength       =   9
            MultiLine       =   -1  'True
            TabIndex        =   22
            Top             =   630
            Width           =   945
         End
         Begin VB.TextBox txtstrNumero 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   6900
            MaxLength       =   6
            MultiLine       =   -1  'True
            TabIndex        =   17
            Top             =   240
            Width           =   945
         End
         Begin MSDataListLib.DataCombo dbcintUF 
            Height          =   315
            Left            =   6840
            TabIndex        =   27
            Top             =   990
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            Text            =   ""
         End
         Begin VB.TextBox txtstrFax 
            Height          =   285
            Left            =   4170
            MaxLength       =   17
            TabIndex        =   34
            Top             =   1695
            Width           =   1605
         End
         Begin MSDataListLib.DataCombo dbcintLogradouro 
            Height          =   315
            Left            =   2040
            TabIndex        =   15
            Top             =   240
            Width           =   4575
            _ExtentX        =   8070
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbcintBairro 
            Height          =   315
            Left            =   1005
            TabIndex        =   19
            Top             =   630
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dbcintTipoLogradouro 
            Height          =   315
            Left            =   1005
            TabIndex        =   13
            Top             =   240
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            Text            =   ""
         End
         Begin VB.Label lblintCodTipoLogradouro 
            AutoSize        =   -1  'True
            Caption         =   "Logradouro"
            Height          =   195
            Left            =   120
            TabIndex        =   12
            Top             =   315
            Width           =   810
         End
         Begin VB.Label lblintCodBairro 
            AutoSize        =   -1  'True
            Caption         =   "Bairro"
            Height          =   195
            Left            =   525
            TabIndex        =   18
            Top             =   671
            Width           =   405
         End
         Begin VB.Label lblintCodCidade 
            AutoSize        =   -1  'True
            Caption         =   "Cidade"
            Height          =   195
            Left            =   435
            TabIndex        =   23
            Top             =   1027
            Width           =   495
         End
         Begin VB.Label lblstrEmail 
            AutoSize        =   -1  'True
            Caption         =   "e-mail"
            Height          =   195
            Left            =   525
            TabIndex        =   29
            Top             =   1383
            Width           =   405
         End
         Begin VB.Label lblstrTelefone 
            AutoSize        =   -1  'True
            Caption         =   "Telefone"
            Height          =   195
            Left            =   300
            TabIndex        =   31
            Top             =   1740
            Width           =   630
         End
         Begin VB.Label lblintCEP 
            AutoSize        =   -1  'True
            Caption         =   "CEP"
            Height          =   195
            Left            =   6480
            TabIndex        =   21
            Top             =   690
            Width           =   315
         End
         Begin VB.Label lblstrFax 
            AutoSize        =   -1  'True
            Caption         =   "Fax"
            Height          =   195
            Left            =   3750
            TabIndex        =   33
            Top             =   1740
            Width           =   255
         End
         Begin VB.Label lblintUF 
            AutoSize        =   -1  'True
            Caption         =   "UF"
            Height          =   195
            Left            =   6540
            TabIndex        =   26
            Top             =   1050
            Width           =   210
         End
      End
      Begin VB.Frame frm_Diretorios 
         Caption         =   "Diretórios"
         Height          =   1335
         Left            =   -74790
         TabIndex        =   43
         Top             =   1200
         Width           =   7815
         Begin VB.CheckBox chk_bytatualizaterminais 
            Caption         =   "Atualizar terminais"
            Height          =   195
            Left            =   1080
            TabIndex        =   50
            Top             =   1020
            Width           =   1785
         End
         Begin VB.CommandButton cmdAbrir 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "System"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   1
            Left            =   7320
            TabIndex        =   49
            Top             =   660
            Width           =   345
         End
         Begin VB.CommandButton cmdAbrir 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "System"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Index           =   0
            Left            =   7320
            TabIndex        =   46
            Top             =   270
            Width           =   345
         End
         Begin VB.TextBox txtstrDocumentos 
            Height          =   285
            Left            =   1080
            MaxLength       =   60
            TabIndex        =   45
            Top             =   300
            Width           =   6165
         End
         Begin VB.TextBox txtstrAtualizacao 
            Height          =   285
            Left            =   1080
            MaxLength       =   60
            TabIndex        =   48
            Top             =   690
            Width           =   6165
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Atualização"
            Height          =   195
            Left            =   180
            TabIndex        =   47
            Top             =   720
            Width           =   825
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Documentos"
            Height          =   195
            Left            =   105
            TabIndex        =   44
            Top             =   345
            Width           =   900
         End
      End
      Begin VB.Label lblstrCGC 
         AutoSize        =   -1  'True
         Caption         =   "CNPJ"
         Height          =   195
         Left            =   1035
         TabIndex        =   5
         Top             =   1230
         Width           =   405
      End
      Begin VB.Label lblstrInscricaoEstadual 
         AutoSize        =   -1  'True
         Caption         =   "Inscrição Estadual"
         Height          =   195
         Left            =   4350
         TabIndex        =   7
         Top             =   1230
         Width           =   1305
      End
      Begin VB.Label lblstrNome 
         AutoSize        =   -1  'True
         Caption         =   "Nome"
         Height          =   195
         Left            =   1020
         TabIndex        =   1
         Top             =   525
         Width           =   420
      End
      Begin VB.Label lblstrNomeFantasia 
         AutoSize        =   -1  'True
         Caption         =   "Nome Fantasia"
         Height          =   195
         Left            =   375
         TabIndex        =   3
         Top             =   877
         Width           =   1065
      End
   End
   Begin VB.TextBox txtPKId 
      Height          =   270
      Left            =   1110
      TabIndex        =   63
      Text            =   "1"
      Top             =   120
      Visible         =   0   'False
      Width           =   645
   End
End
Attribute VB_Name = "frmCadEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const MAX_PATH = 260

Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Dim mblnAlterando                As Boolean

Private Sub cmdAbrir_Click(Index As Integer)
    Dim intNull             As Integer
    Dim lngIDList           As Long
    Dim strPath             As String
    Dim udtBI As BrowseInfo

    With udtBI
        .hWndOwner = Me.Hwnd
        Select Case Index
            Case 0
                .lpszTitle = lstrcat("Diretório Documentos", "")
            Case 1
                .lpszTitle = lstrcat("Diretório atualização", "")
        End Select
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With
    
    lngIDList = SHBrowseForFolder(udtBI)
    If lngIDList Then
        strPath = String$(MAX_PATH, 0)
        SHGetPathFromIDList lngIDList, strPath
        CoTaskMemFree lngIDList
        intNull = InStr(strPath, vbNullChar)
        If intNull Then
            strPath = Left$(strPath, intNull - 1)
        End If
    End If

    Select Case Index
        Case 0
            txtstrDocumentos.Text = strPath
        Case 1
            txtstrAtualizacao.Text = strPath
    End Select
End Sub

Private Sub dbcintBairro_Click(Area As Integer)
   DropDownDataCombo dbcintBairro, Me, Area
End Sub

Private Sub dbcintBairro_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbcintBairro, Me, , KeyCode, Shift
End Sub

Private Sub dbcintBairro_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub dbcintCidade_Click(Area As Integer)
   DropDownDataCombo dbcintCidade, Me, Area
End Sub

Private Sub dbcintCidade_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbcintCidade, Me, , KeyCode, Shift
End Sub

Private Sub dbcintCidade_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub dbcintLogradouro_Click(Area As Integer)
   DropDownDataCombo dbcintLogradouro, Me, Area
End Sub

Private Sub dbcintLogradouro_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbcintLogradouro, Me, , KeyCode, Shift
End Sub

Private Sub dbcintLogradouro_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub dbcintTipoLogradouro_Click(Area As Integer)
   DropDownDataCombo dbcintTipoLogradouro, Me, Area
End Sub

Private Sub dbcintTipoLogradouro_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbcintTipoLogradouro, Me, , KeyCode, Shift
End Sub

Private Sub dbcintTipoLogradouro_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub dbcintUf_Click(Area As Integer)
   DropDownDataCombo dbcintUF, Me, Area
End Sub

Private Sub dbcintUf_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbcintUF, Me, , KeyCode, Shift
End Sub

Private Sub dbcintUF_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "U", dbcintUF
End Sub

Private Sub cmd_Bairro_Click()
    CarregaForm frmCadBairro, dbcintBairro
End Sub

Private Sub cmd_Cidade_Click()
    CarregaForm frmCadCidade, dbcintCidade
End Sub

Private Sub cmd_Logradouro_Click()
    CarregaForm frmCadLogradouro, dbcintLogradouro
End Sub

Private Sub cmd_Tipo_Click()
    CarregaForm frmCadTipoLogradouro, dbcintTipoLogradouro
End Sub

Private Sub cmd_UF_Click()
    CarregaForm frmCadUF, dbcintUF
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 573
    VirificaGradeListView Me
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrAplicar, gstrDeletar, gstrSalvar, gstrBrasao, gstrLogotipo
End Sub

Private Sub Form_Deactivate()
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrAplicar, gstrDeletar
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = vbKeyF1 Then
        Call_HtmlHelp Me.HelpContextID
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrAplicar, gstrDeletar, gstrBrasao, gstrLogotipo
End Sub

Private Sub img_Brasao_DblClick()
    MantemForm gstrBrasao
End Sub

Private Sub img_LogoTipo_DblClick()
    MantemForm gstrLogotipo
End Sub


Private Sub txtintCepFinal_GotFocus()
    MarcaCampo txtintcepfinal
End Sub

Private Sub txtintCepFinal_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "E", txtintcepfinal
End Sub

Private Sub txtintCepInicial_GotFocus()
    MarcaCampo txtintcepinicial
End Sub

Private Sub txtintCepInicial_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "E", txtintcepinicial
End Sub

Private Sub txtintcertidaomobiliario_GotFocus()
    MarcaCampo txtintcertidaomobiliario
End Sub

Private Sub txtintcertidaomobiliario_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintcertidaomobiliario
End Sub

Private Sub txtintCEP_GotFocus()
    MarcaCampo txtintCep
End Sub

Private Sub txtintCEP_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "E", txtintCep
End Sub



Private Sub txtintnumeroalvarafuncionamento_GotFocus()
    MarcaCampo txtintnumeroalvarafuncionamento
End Sub

Private Sub txtintnumeroalvarafuncionamento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintnumeroalvarafuncionamento
End Sub

Private Sub txtINTNUMEROGUIANEGATIVA_GotFocus()
    MarcaCampo txtINTNUMEROGUIANEGATIVA
End Sub

Private Sub txtINTNUMEROGUIANEGATIVA_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtINTNUMEROGUIANEGATIVA
End Sub

Private Sub txtINTNUMEROGUIAPOSITIVA_GotFocus()
    MarcaCampo txtINTNUMEROGUIAPOSITIVA
End Sub

Private Sub txtINTNUMEROGUIAPOSITIVA_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtINTNUMEROGUIAPOSITIVA
End Sub

Private Sub txtstrCGC_GotFocus()
    'Retira o caracteres não numérico do CGC
    txtstrCGC = gstrValorSemMascara(txtstrCGC)
    'Marca o campo para sobrepor a nova digitação
    MarcaCampo txtstrCGC
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
    Select Case UCase(strModoOperacao)
    Case UCase(gstrBrasao)
        frmCadImagem.CadastraFoto img_Brasao, txtintBrasao
        frmCadImagem.Caption = "Brasão"
    Case UCase(gstrLogotipo)
        frmCadImagem.CadastraFoto img_LogoTipo, txtintLogotipo
        frmCadImagem.Caption = "Logotipo"
    Case gstrSalvar
        GravaEmpresa
    End Select
End Sub

Private Sub txtstrCGC_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtstrCGC
End Sub

Private Sub txtstrCGC_LostFocus()
    txtstrCGC = gstrCGCCPFFormatado(txtstrCGC)
End Sub

Private Sub txtstrEmail_GotFocus()
    MarcaCampo txtstrEmail
End Sub

Private Sub txtstrEmail_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub txtstrFax_GotFocus()
    MarcaCampo txtstrFax
End Sub

Private Sub txtstrFax_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtstrFax
End Sub

Private Sub txtstrInscricaoEstadual_GotFocus()
    MarcaCampo txtstrInscricaoEstadual
End Sub

Private Sub txtstrInscricaoEstadual_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub txtstrNome_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub txtstrNomeFantasia_GotFocus()
    MarcaCampo txtstrNomeFantasia
End Sub

Private Sub txtstrNomeFantasia_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub txtstrNumero_GotFocus()
    MarcaCampo txtstrNumero
End Sub

Private Sub txtstrNumero_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtstrNumero
End Sub

Private Sub txtstrNome_GotFocus()
    MarcaCampo txtstrNome
End Sub

Private Sub txtstrTelefone_GotFocus()
    MarcaCampo txtstrTelefone
End Sub

Private Sub txtstrTelefone_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtstrTelefone
End Sub

Private Sub Form_Load()
    mblnAlterando = False
    LeDaTabelaParaObj "", dbcintUF, gstrQueryUF
    LeDaTabelaParaObj gstrBairro, dbcintBairro
    LeDaTabelaParaObj gstrLogradouro, dbcintLogradouro
    LeDaTabelaParaObj gstrTipoLogradouro, dbcintTipoLogradouro, gstrQueryTipoLogradouro
    LeDaTabelaParaObj gstrCidade, dbcintCidade, gstrQueryCidade
    LeImagemLogotipo img_Brasao, img_LogoTipo
    LeEmpresa
End Sub

Private Function blnDadosOk() As Boolean
    If Trim(txtstrCGC) <> "" And gblnCGCOk(txtstrCGC) = False Then
        ExibeMensagem "O CNPJ deve ser informado corretamente."
        txtstrCGC.SetFocus
    ElseIf Trim(txtintFebraban) = 0 Then
        ExibeMensagem "O código de Febraban não pode ser '0'."
        txtintFebraban.SetFocus
    ElseIf Trim(txtintcertidaomobiliario) = "" Then
        ExibeMensagem "O Número da Certidão Mobiliária não pode ser nulo"
        txtintcertidaomobiliario.SetFocus
    ElseIf Trim(txtintnumeroalvarafuncionamento) = "" Then
        ExibeMensagem "O Número do Alvará de Funcionamento não pode ser nulo"
        txtintnumeroalvarafuncionamento.SetFocus
    Else
        blnDadosOk = True
    End If
End Function

Private Sub GravaEmpresa()
    Dim strSql    As String
    If blnDadosOk Then
        If gblnExclusaoGravacaoOk(IIf(mblnAlterando, "A", "I"), " de " & Trim(txtstrNome)) Then
            If mblnAlterando Then
                strSql = strQueryAlteraEmpresa
            Else
                strSql = strQueryIncluiEmpresa
            End If
            Set gobjBanco = New clsBanco
            If gobjBanco.Execute(strSql) Then
                Unload Me
            End If
        End If
    End If
End Sub

Private Function strQueryIncluiEmpresa() As String
    Dim strSql  As String
    Dim STRNOME As String
    
    STRNOME = gstrStringCripitografada((txtstrNome), True, True)
    If Trim(STRNOME) = "" Then
        STRNOME = "NULL"
    Else
        STRNOME = "'" & STRNOME & "'"
    End If
    
    strSql = ""
    strSql = strSql & "INSERT INTO " & gstrEmpresa & " ("
    strSql = strSql & "strNome, strNomeFantasia, strCGC, strInscricaoEstadual, "
    strSql = strSql & "intTipoLogradouro, intLogradouro, strNumero, intBairro, "
    strSql = strSql & "intCep, intCidade, intUF, strEmail, "
    strSql = strSql & "strTelefone, strFax,intcepinicial,intcepfinal,intLogotipo, intBrasao, intSerie, strDocumentos, strAtualizacao, IntFebraban, "
    strSql = strSql & " INTNUMEROGUIANEGATIVA, INTNUMEROGUIAPOSITIVA, INTNUMEROCERTIDAOCADMOBILIARIO,INTNUMEROALVARAFUNCIONAMENTO ,bytobrigatoriologradouro, dtmDtAtualizacao, lngCodUsr, bytatualizaterminais"
    strSql = strSql & ") VALUES ("
    strSql = strSql & STRNOME & ", '"
    strSql = strSql & Trim(txtstrNomeFantasia) & "', '"
    strSql = strSql & Trim(txtstrCGC) & "', '"
    strSql = strSql & Trim(txtstrInscricaoEstadual) & "', "
    strSql = strSql & gstrItemData(dbcintTipoLogradouro, True) & ", "
    strSql = strSql & gstrItemData(dbcintLogradouro, True) & ", "
    strSql = strSql & Val(txtstrNumero) & ", "
    strSql = strSql & gstrItemData(dbcintBairro, True) & ", "
    strSql = strSql & Val(gstrValorSemMascara(txtintCep)) & ", "
    strSql = strSql & gstrItemData(dbcintCidade, True) & ", "
    strSql = strSql & gstrItemData(dbcintUF, True) & ", '"
    strSql = strSql & Trim(txtstrEmail) & "', '"
    strSql = strSql & Trim(txtstrTelefone) & "', '"
    strSql = strSql & Trim(txtstrFax) & "', "
    strSql = strSql & Val(gstrValorSemMascara(txtintcepinicial)) & ", "
    strSql = strSql & Val(gstrValorSemMascara(txtintcepfinal)) & ", "
    strSql = strSql & gstrConvVrParaSql(txtintLogotipo) & ", "
    strSql = strSql & gstrConvVrParaSql(txtintBrasao) & ", "
    strSql = strSql & glngNumeroDeSerie(CStr(txtstrNome)) & ", '"
    strSql = strSql & Trim(txtstrDocumentos.Text) & "', '"
    strSql = strSql & Trim(txtstrAtualizacao.Text) & "', "
    strSql = strSql & gstrENulo(Trim(txtintFebraban.Text), False, True) & ", "
    strSql = strSql & gstrENulo(Trim(txtINTNUMEROGUIANEGATIVA.Text), False, True) & ", "
    strSql = strSql & gstrENulo(Trim(txtINTNUMEROGUIAPOSITIVA.Text), False, True) & ", "
    strSql = strSql & gstrENulo(Trim(txtintcertidaomobiliario.Text), False, True) & ", "
    strSql = strSql & gstrENulo(Trim(txtintnumeroalvarafuncionamento.Text), False, True) & ", "
    strSql = strSql & gstrENulo(chk_bytLogradouro, False, True) & ", "
    strSql = strSql & gstrConvDtParaSql(gstrDataDoSistema()) & ", "
    strSql = strSql & glngCodUsr & ", "
    strSql = strSql & gstrENulo(chk_bytatualizaterminais, False, True) & ")"
    
    strQueryIncluiEmpresa = strSql
    
End Function

Private Function strQueryAlteraEmpresa() As String
    Dim strSql As String
    
    strSql = ""
    strSql = strSql & "UPDATE " & gstrEmpresa & " SET "
    strSql = strSql & "strNomeFantasia = '" & Trim(txtstrNomeFantasia) & "', "
    strSql = strSql & "strCGC = '" & Trim(txtstrCGC) & "', "
    strSql = strSql & "strInscricaoEstadual = '" & Trim(txtstrInscricaoEstadual) & "', "
    strSql = strSql & "intTipoLogradouro = " & gstrItemData(dbcintTipoLogradouro, True) & ", "
    strSql = strSql & "intLogradouro = " & gstrItemData(dbcintLogradouro, True) & ", "
    strSql = strSql & "strNumero = '" & Trim(txtstrNumero) & "', "
    strSql = strSql & "intBairro = " & gstrItemData(dbcintBairro, True) & ", "
    strSql = strSql & "intCEP = " & Val(gstrValorSemMascara(txtintCep)) & ", "
    strSql = strSql & "intCidade = " & gstrItemData(dbcintCidade, True) & ", "
    strSql = strSql & "intUF = " & gstrItemData(dbcintUF, True) & ", "
    strSql = strSql & "strEmail = '" & Trim(txtstrEmail) & "', "
    strSql = strSql & "strTelefone = '" & Trim(txtstrTelefone) & "', "
    strSql = strSql & "strFax = '" & Trim(txtstrFax) & "', "
    strSql = strSql & "intcepinicial = " & Val(gstrValorSemMascara(txtintcepinicial)) & ", "
    strSql = strSql & "intcepfinal = " & Val(gstrValorSemMascara(txtintcepfinal)) & ", "
    strSql = strSql & "intLogotipo = " & gstrConvVrParaSql(txtintLogotipo) & ", "
    strSql = strSql & "intBrasao = " & gstrConvVrParaSql(txtintBrasao) & ", "
    strSql = strSql & "strDocumentos = '" & Trim(txtstrDocumentos.Text) & "', "
    strSql = strSql & "strAtualizacao = '" & Trim(txtstrAtualizacao.Text) & "', "
    strSql = strSql & "intFebraban = " & gstrENulo(txtintFebraban, False, True) & ", "
    strSql = strSql & "INTNUMEROGUIANEGATIVA = " & gstrENulo(txtINTNUMEROGUIANEGATIVA, False, True) & ", "
    strSql = strSql & "INTNUMEROGUIAPOSITIVA = " & gstrENulo(txtINTNUMEROGUIAPOSITIVA, False, True) & ", "
    strSql = strSql & "INTNUMEROCERTIDAOCADMOBILIARIO = " & gstrENulo(txtintcertidaomobiliario, False, True) & ", "
    strSql = strSql & "INTNUMEROALVARAFUNCIONAMENTO = " & gstrENulo(txtintnumeroalvarafuncionamento, False, True) & ", "
    strSql = strSql & "bytobrigatoriologradouro = " & gstrENulo(chk_bytLogradouro, False, True) & ", "
    strSql = strSql & "dtmDtAtualizacao = "
    strSql = strSql & gstrConvDtParaSql(gstrDataDoSistema()) & ", "
    strSql = strSql & "bytatualizaterminais = " & gstrENulo(chk_bytatualizaterminais, False, True) & ", "
    strSql = strSql & "lngCodUsr = " & glngCodUsr
    strQueryAlteraEmpresa = strSql
End Function

Private Sub LeEmpresa()
    Dim strSql          As String
    Dim adoResultado    As ADODB.Recordset
    strSql = ""
    strSql = strSql & "SELECT strNome, strNomeFantasia, strCGC, "
    strSql = strSql & "strInscricaoEstadual, intTipoLogradouro, intLogradouro, INTNUMEROGUIANEGATIVA, INTNUMEROGUIAPOSITIVA, INTNUMEROCERTIDAOCADMOBILIARIO, "
    strSql = strSql & "strNumero, intBairro, intCep, intCidade, intUf,  "
    strSql = strSql & "strEmail, strTelefone, strFax,intcepinicial,intcepfinal, intLogotipo, intBrasao, strDocumentos, strAtualizacao, IntFebraban,intnumeroalvarafuncionamento, bytobrigatoriologradouro, bytatualizaterminais "
    strSql = strSql & "FROM " & gstrEmpresa
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        With adoResultado
            If .EOF = False Then
                txtstrNome = gstrStringCripitografada((!STRNOME), , True)
                txtstrNomeFantasia = Trim(!strNomeFantasia)
                txtstrCGC = gstrCGCCPFFormatado(!strCGC)
                txtstrInscricaoEstadual = gstrVerificaCampoNulo(!strInscricaoEstadual)
                dbcintTipoLogradouro.BoundText = gstrVerificaCampoNulo(!intTipoLogradouro)
                dbcintLogradouro.BoundText = gstrVerificaCampoNulo(!intLogradouro)
                txtstrNumero = gstrVerificaCampoNulo(!strNumero)
                dbcintBairro.BoundText = gstrVerificaCampoNulo(!intBairro)
                txtintCep = gstrCEPFormatado(!INTCEP)
                dbcintCidade.BoundText = gstrVerificaCampoNulo(!intCidade)
                dbcintUF.BoundText = gstrVerificaCampoNulo(!intUf)
                txtstrEmail = Trim(!strEmail)
                txtstrTelefone = Trim(!strTelefone)
                txtstrFax = Trim(!strFax)
                txtintcepinicial = gstrCEPFormatado(!intCepInicial)
                txtintcepfinal = gstrCEPFormatado(!intCepFinal)
                txtintLogotipo = gstrENulo(!intLogotipo)
                txtintBrasao = gstrENulo(!intBrasao)
                txtstrDocumentos.Text = gstrENulo(!strDocumentos)
                txtstrAtualizacao.Text = gstrENulo(!strAtualizacao)
                txtintFebraban.Text = gstrENulo(!intFebraban)
                txtINTNUMEROGUIANEGATIVA = gstrENulo(!INTNUMEROGUIANEGATIVA)
                txtINTNUMEROGUIAPOSITIVA = gstrENulo(!INTNUMEROGUIAPOSITIVA)
                txtintcertidaomobiliario = gstrENulo(!INTNUMEROCERTIDAOCADMOBILIARIO)
                txtintnumeroalvarafuncionamento = gstrENulo(!INTNUMEROALVARAFUNCIONAMENTO)
                chk_bytLogradouro.Value = Val(gstrENulo(!bytobrigatoriologradouro))
                chk_bytatualizaterminais.Value = Val(gstrENulo(!bytatualizaterminais))
                mblnAlterando = True
                txtstrNome.Enabled = False
            End If
            .Close
        End With
    End If
End Sub
