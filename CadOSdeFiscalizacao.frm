VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Begin VB.Form frmCadOSdeFiscalizacao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "O.S. de Fiscalização"
   ClientHeight    =   6780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8580
   Icon            =   "CadOSdeFiscalizacao.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   8580
   Begin VB.TextBox txtPKId 
      Height          =   285
      Left            =   7440
      TabIndex        =   38
      Top             =   60
      Visible         =   0   'False
      Width           =   945
   End
   Begin TabDlg.SSTab tab_3DPasta 
      Height          =   5055
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   8325
      _ExtentX        =   14684
      _ExtentY        =   8916
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "O.S. de Fiscalização"
      TabPicture(0)   =   "CadOSdeFiscalizacao.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl_strIncricaoCadastral"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "dbc_strInscricaoCadastral"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fra_bytOrigem"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fra_Proprietario"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Fra_Razao"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Fiscais"
      TabPicture(1)   =   "CadOSdeFiscalizacao.frx":105E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "tdb_Fiscal"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Fra_Razao 
         Caption         =   "Razão"
         Height          =   1275
         Left            =   5370
         TabIndex        =   34
         Top             =   540
         Width           =   2805
         Begin VB.TextBox txtstrRazao 
            Height          =   945
            Left            =   90
            TabIndex        =   35
            Top             =   270
            Width           =   2625
         End
      End
      Begin VB.Frame fra_Proprietario 
         Height          =   3150
         Left            =   120
         TabIndex        =   7
         Top             =   1800
         Width           =   8070
         Begin VB.TextBox txt_intContribuinte 
            BackColor       =   &H80000004&
            Height          =   285
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txt_strCNPJCPFP 
            BackColor       =   &H80000004&
            Height          =   285
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   600
            Width           =   1845
         End
         Begin VB.TextBox txt_strNome 
            BackColor       =   &H80000004&
            Height          =   285
            Left            =   2310
            Locked          =   -1  'True
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   240
            Width           =   5595
         End
         Begin VB.TextBox txt_strAtividadeBasica 
            BackColor       =   &H80000004&
            Height          =   285
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   26
            TabStop         =   0   'False
            Top             =   960
            Width           =   4110
         End
         Begin VB.TextBox txt_strAtividadePrincipal 
            BackColor       =   &H80000004&
            Height          =   285
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   1290
            Width           =   4110
         End
         Begin VB.Frame fra_Endereco 
            Caption         =   "Endereço"
            Height          =   1395
            Left            =   120
            TabIndex        =   8
            Top             =   1620
            Width           =   7845
            Begin VB.TextBox txt_UF 
               BackColor       =   &H80000004&
               Height          =   285
               Left            =   5070
               Locked          =   -1  'True
               TabIndex        =   16
               Top             =   990
               Width           =   510
            End
            Begin VB.TextBox txt_Municipio 
               BackColor       =   &H80000004&
               Height          =   285
               Left            =   5070
               Locked          =   -1  'True
               TabIndex        =   15
               Top             =   630
               Width           =   2625
            End
            Begin VB.TextBox txt_Cep 
               BackColor       =   &H80000004&
               Height          =   285
               Left            =   6615
               Locked          =   -1  'True
               TabIndex        =   14
               Top             =   990
               Width           =   1080
            End
            Begin VB.TextBox txt_Complemento 
               BackColor       =   &H80000004&
               Height          =   285
               Left            =   6810
               Locked          =   -1  'True
               TabIndex        =   13
               Top             =   270
               Width           =   870
            End
            Begin VB.TextBox txt_Numero 
               BackColor       =   &H80000004&
               Height          =   285
               Left            =   5400
               Locked          =   -1  'True
               TabIndex        =   12
               Top             =   270
               Width           =   795
            End
            Begin VB.TextBox txt_Logradouro 
               BackColor       =   &H80000004&
               Height          =   285
               Left            =   1020
               Locked          =   -1  'True
               TabIndex        =   11
               Top             =   270
               Width           =   4005
            End
            Begin VB.TextBox txt_Bairro 
               BackColor       =   &H80000004&
               Height          =   285
               Left            =   1020
               Locked          =   -1  'True
               TabIndex        =   10
               Top             =   630
               Width           =   3105
            End
            Begin VB.TextBox txt_Distrito 
               BackColor       =   &H80000004&
               Height          =   285
               Left            =   1020
               Locked          =   -1  'True
               TabIndex        =   9
               Top             =   990
               Width           =   3525
            End
            Begin VB.Label lbl_strDistritoC 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Distrito"
               Height          =   195
               Left            =   450
               TabIndex        =   24
               Top             =   1050
               Width           =   480
            End
            Begin VB.Label lbl_intMunicipioC 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Município"
               Height          =   195
               Left            =   4290
               TabIndex        =   23
               Top             =   690
               Width           =   705
            End
            Begin VB.Label lbl_intBairroC 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Bairro"
               Height          =   195
               Left            =   525
               TabIndex        =   22
               Top             =   690
               Width           =   405
            End
            Begin VB.Label lbl_intLogradouroC 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Logradouro"
               Height          =   195
               Left            =   120
               TabIndex        =   21
               Top             =   330
               Width           =   810
            End
            Begin VB.Label lbl_intNumeroC 
               AutoSize        =   -1  'True
               Caption         =   "Nº"
               Height          =   195
               Left            =   5130
               TabIndex        =   20
               Top             =   330
               Width           =   180
            End
            Begin VB.Label lbl_strComplementoC 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Compl."
               Height          =   195
               Left            =   6300
               TabIndex        =   19
               Top             =   330
               Width           =   480
            End
            Begin VB.Label lbl_intUFC 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "UF"
               Height          =   195
               Left            =   4770
               TabIndex        =   18
               Top             =   1065
               Width           =   210
            End
            Begin VB.Label lbl_intCepC 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Cep"
               Height          =   195
               Left            =   6240
               TabIndex        =   17
               Top             =   1050
               Width           =   285
            End
         End
         Begin VB.Label lbl_strCNPJCPFP 
            AutoSize        =   -1  'True
            Caption         =   "CPF/CNPJ"
            Height          =   195
            Left            =   690
            TabIndex        =   33
            Top             =   645
            Width           =   780
         End
         Begin VB.Label lbl_strNome 
            AutoSize        =   -1  'True
            Caption         =   "Proprietário"
            Height          =   195
            Left            =   675
            TabIndex        =   32
            Top             =   285
            Width           =   795
         End
         Begin VB.Label lbl_strAtividadeBasica 
            AutoSize        =   -1  'True
            Caption         =   "Atividade básica"
            Height          =   195
            Left            =   300
            TabIndex        =   31
            Top             =   1005
            Width           =   1170
         End
         Begin VB.Label lbl_strAtividadePrincipal 
            AutoSize        =   -1  'True
            Caption         =   "Atividade principal"
            Height          =   195
            Left            =   180
            TabIndex        =   30
            Top             =   1335
            Width           =   1290
         End
      End
      Begin VB.Frame fra_bytOrigem 
         Caption         =   " Origem "
         Height          =   705
         Left            =   120
         TabIndex        =   1
         Top             =   540
         Width           =   5145
         Begin VB.OptionButton optbytOrigem 
            Caption         =   "Imobiliário Rural"
            Height          =   195
            Index           =   2
            Left            =   3630
            TabIndex        =   4
            Top             =   300
            Width           =   1425
         End
         Begin VB.OptionButton optbytOrigem 
            Caption         =   "Imobiliário Urbano"
            Height          =   195
            Index           =   1
            Left            =   1710
            TabIndex        =   3
            Top             =   300
            Width           =   1695
         End
         Begin VB.OptionButton optbytOrigem 
            Caption         =   "Econômico"
            Height          =   195
            Index           =   0
            Left            =   330
            TabIndex        =   2
            Top             =   300
            Width           =   1155
         End
      End
      Begin MSDataListLib.DataCombo dbc_strInscricaoCadastral 
         Height          =   315
         Left            =   1710
         TabIndex        =   6
         Top             =   1470
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   556
         _Version        =   393216
         IntegralHeight  =   0   'False
         Text            =   ""
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_Fiscal 
         Height          =   4125
         Left            =   -74850
         TabIndex        =   36
         Top             =   810
         Width           =   7965
         _ExtentX        =   14049
         _ExtentY        =   7276
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "PKId"
         Columns(0).DataField=   "PKId"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   68
         Columns(1)._MaxComboItems=   20
         Columns(1).ValueItems(0)._DefaultItem=   0
         Columns(1).ValueItems(0).Value=   ""
         Columns(1).ValueItems(0).Value.vt=   8
         Columns(1).ValueItems(0).DisplayValue=   ""
         Columns(1).ValueItems(0).DisplayValue.vt=   8
         Columns(1).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
         Columns(1).ValueItems.Count=   1
         Columns(1).DataField=   "Boolean"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Fiscal"
         Columns(2).DataField=   "strNome"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Horário"
         Columns(3).DataField=   "dtmDtHorario"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Data"
         Columns(4).DataField=   ""
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   5
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).ScrollBars=   2
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=5"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0).AllowSizing=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=450"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=370"
         Splits(0)._ColumnProps(10)=   "Column(1).AllowSizing=0"
         Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(12)=   "Column(2).Width=9287"
         Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=9208"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(16)=   "Column(3).Width=2037"
         Splits(0)._ColumnProps(17)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(18)=   "Column(3)._WidthInPix=1958"
         Splits(0)._ColumnProps(19)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(20)=   "Column(4).Width=1826"
         Splits(0)._ColumnProps(21)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(22)=   "Column(4)._WidthInPix=1746"
         Splits(0)._ColumnProps(23)=   "Column(4).Order=5"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         DataMode        =   4
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         MultipleLines   =   0
         CellTipsWidth   =   0
         DeadAreaBackColor=   12632256
         RowDividerColor =   12632256
         RowSubDividerColor=   12632256
         DirectionAfterEnter=   1
         MaxRows         =   250000
         ViewColumnCaptionWidth=   0
         ViewColumnWidth =   0
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
         _StyleDefs(11)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
         _StyleDefs(12)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(13)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(14)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(15)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(16)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(17)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(18)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(19)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(20)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(21)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(23)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(24)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(25)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(26)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(27)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(28)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(29)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(30)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(31)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=46,.parent=13"
         _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
         _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
         _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
         _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
         _StyleDefs(49)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
         _StyleDefs(52)  =   "Named:id=33:Normal"
         _StyleDefs(53)  =   ":id=33,.parent=0"
         _StyleDefs(54)  =   "Named:id=34:Heading"
         _StyleDefs(55)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(56)  =   ":id=34,.wraptext=-1"
         _StyleDefs(57)  =   "Named:id=35:Footing"
         _StyleDefs(58)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(59)  =   "Named:id=36:Selected"
         _StyleDefs(60)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(61)  =   "Named:id=37:Caption"
         _StyleDefs(62)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(63)  =   "Named:id=38:HighlightRow"
         _StyleDefs(64)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(65)  =   "Named:id=39:EvenRow"
         _StyleDefs(66)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(67)  =   "Named:id=40:OddRow"
         _StyleDefs(68)  =   ":id=40,.parent=33"
         _StyleDefs(69)  =   "Named:id=41:RecordSelector"
         _StyleDefs(70)  =   ":id=41,.parent=34"
         _StyleDefs(71)  =   "Named:id=42:FilterBar"
         _StyleDefs(72)  =   ":id=42,.parent=33"
      End
      Begin VB.Label lbl_strIncricaoCadastral 
         Caption         =   "Inscrição Cadastral"
         Height          =   255
         Left            =   180
         TabIndex        =   5
         Top             =   1500
         Width           =   1425
      End
   End
   Begin TrueOleDBGrid70.TDBGrid tdb_OrdemServico 
      Height          =   1485
      Left            =   120
      TabIndex        =   37
      Top             =   5220
      Width           =   8325
      _ExtentX        =   14684
      _ExtentY        =   2619
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "PKID"
      Columns(0).DataField=   "PKID"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Inscrição Cadastral"
      Columns(1).DataField=   "Inscrição Cadastral"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Proprietário"
      Columns(2).DataField=   "strNome"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   3
      Splits(0)._UserFlags=   0
      Splits(0).MarqueeStyle=   3
      Splits(0).RecordSelectors=   0   'False
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).ScrollBars=   2
      Splits(0).DividerColor=   12632256
      Splits(0).FilterBar=   -1  'True
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=3"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
      Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(8)=   "Column(1).Width=3440"
      Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=3360"
      Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(13)=   "Column(2).Width=10742"
      Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=10663"
      Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTips        =   1
      CellTipsWidth   =   0
      DeadAreaBackColor=   12632256
      RowDividerColor =   12632256
      RowSubDividerColor=   12632256
      DirectionAfterEnter=   1
      MaxRows         =   250000
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=101,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
      _StyleDefs(11)  =   "FooterStyle:id=3,.parent=1,.namedParent=35"
      _StyleDefs(12)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(13)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(14)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(15)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(16)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(17)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(18)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(19)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(20)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(21)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(22)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(23)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(24)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(25)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(26)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(27)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(28)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(29)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(30)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(31)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(32)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(33)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(34)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(35)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(36)  =   "Splits(0).Columns(1).Style:id=50,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(1).HeadingStyle:id=47,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(1).FooterStyle:id=48,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(1).EditorStyle:id=49,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
      _StyleDefs(44)  =   "Named:id=33:Normal"
      _StyleDefs(45)  =   ":id=33,.parent=0"
      _StyleDefs(46)  =   "Named:id=34:Heading"
      _StyleDefs(47)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(48)  =   ":id=34,.wraptext=-1"
      _StyleDefs(49)  =   "Named:id=35:Footing"
      _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(51)  =   "Named:id=36:Selected"
      _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(53)  =   "Named:id=37:Caption"
      _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(55)  =   "Named:id=38:HighlightRow"
      _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
      _StyleDefs(57)  =   "Named:id=39:EvenRow"
      _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(59)  =   "Named:id=40:OddRow"
      _StyleDefs(60)  =   ":id=40,.parent=33"
      _StyleDefs(61)  =   "Named:id=41:RecordSelector"
      _StyleDefs(62)  =   ":id=41,.parent=34"
      _StyleDefs(63)  =   "Named:id=42:FilterBar"
      _StyleDefs(64)  =   ":id=42,.parent=33"
   End
End
Attribute VB_Name = "frmCadOSdeFiscalizacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mblnAlterando           As Boolean
Dim mobjAux                 As Object
Dim mblnSelecionou          As Boolean
Dim mblnPrimeiraVez         As Boolean
Dim opt                     As Integer
Dim MatFiscais              As XArrayDB

Private Sub dbc_strInscricaoCadastral_Click(Area As Integer)
    DropDownDataCombo dbc_strInscricaoCadastral, Me, Area
    If Area = 2 And dbc_strInscricaoCadastral.MatchedWithList Then
        If opt = 0 Then
            ProprietarioEconomico
            GridFiscalInscricaoCadastral opt
        Else
            ProprietarioImobiliario opt
            GridFiscalInscricaoCadastral 1
        End If
    End If
End Sub

Private Sub dbc_strInscricaoCadastral_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_strInscricaoCadastral, Me, , KeyCode, Shift
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 649
    VirificaGradeListView Me
    If mblnSelecionou Then
       HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
    Else
       HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
    End If
    If mobjAux Is Nothing Then
       HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
    Else
       HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
    End If
End Sub

Private Sub Form_Deactivate()
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
End Sub

Private Sub Form_Load()
    opt = 3
    mblnAlterando = False
    VerificaObjParaAplicar mobjAux
    dbc_strInscricaoCadastral.Enabled = False
    TrocaCorObjeto dbc_strInscricaoCadastral, True
    txtstrRazao.Enabled = False
    TrocaCorObjeto txtstrRazao, True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    mblnSelecionou = False
    mblnPrimeiraVez = False
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)

'******************************************************************************************
' Data: 08/05/2003
' Alteração: - Substituição da chamada direta à stored procedure pela função
'            gstrStoredProcedure.
' Responsável: Everton Bianchini
'******************************************************************************************

Dim varBookMark As Variant
Dim strSql As String

If UCase(strModoOperacao) = gstrPreencherLista Then
    PreencherListaDeOpcoes Me.ActiveControl
    Exit Sub
End If

tdb_Fiscal.Update
If UCase(strModoOperacao) = "SALVAR" Or UCase(strModoOperacao) = "DELETAR" Then
    mblnPrimeiraVez = False
End If

If UCase(strModoOperacao) = "SALVAR" Then
    If DadosOk Then
        If mblnAlterando Then
            If MsgBox("Confirma Alteração ?", vbYesNo + vbQuestion) = vbYes Then
'                strSQL = " sp_OrdemServico '" & strPKId & "','" & strDTM & "'," & txtPKId & ",'" & txtstrRazao
'                strSQL = strSQL & "'," & opt & "," & dbc_strInscricaoCadastral.BoundText & "," & glngCodUsr & ",'ALTERAR'"
                strSql = gstrStoredProcedure("sp_OrdemServico", "'" & strPKId & "','" & strDTM & "'," & txtPKId & ",'" & txtstrRazao & _
                            "'," & opt & "," & dbc_strInscricaoCadastral.BoundText & "," & glngCodUsr & ",'ALTERAR'", True)
            Else
                Exit Sub
            End If
        Else
            If MsgBox("Confirma Inclusão ?", vbYesNo + vbQuestion) = vbYes Then
'                strSQL = " sp_OrdemServico '" & strPKId & "','" & strDTM & "',0,'" & txtstrRazao
'                strSQL = strSQL & "'," & opt & "," & dbc_strInscricaoCadastral.BoundText & "," & glngCodUsr & ",'INCLUIR'"
                strSql = gstrStoredProcedure("sp_OrdemServico", "'" & strPKId & "','" & strDTM & "',0,'" & txtstrRazao & _
                            "'," & opt & "," & dbc_strInscricaoCadastral.BoundText & "," & glngCodUsr & ",'INCLUIR'", True)
            Else
                Exit Sub
            End If
        End If
        Set gobjBanco = New clsBanco
        gobjBanco.ExecutaBeginTrans
        If Not gobjBanco.Execute(strSql, False) Then
            gobjBanco.ExecutaRollbackTrans
            Exit Sub
        Else
            gobjBanco.ExecutaCommitTrans
            QueryNovo
            optbytOrigem_Click opt
        End If
    End If
End If
If UCase(strModoOperacao) = "NOVO" Then
    QueryNovo
End If
If UCase(strModoOperacao) = "DELETAR" Then
    If txtPKId <> "" Then
        If MsgBox("Confirma Exclusão ?", vbYesNo + vbQuestion) = vbYes Then
'            strSQL = " sp_OrdemServico '',''," & txtPKId & ",'',3,0,0,'DELETAR'"
            strSql = gstrStoredProcedure("sp_OrdemServico", "'',''," & txtPKId & ",'',3,0,0,'DELETAR'", True)
            Set gobjBanco = New clsBanco
            gobjBanco.ExecutaBeginTrans
            If Not gobjBanco.Execute(strSql, False) Then
                gobjBanco.ExecutaRollbackTrans
                Exit Sub
            Else
                gobjBanco.ExecutaCommitTrans
                QueryNovo
                optbytOrigem_Click opt
            End If
        End If
    Else
        ExibeMensagem "Deve ser Selecionado algum Registro"
    End If
End If
If UCase(strModoOperacao) = "IMPRIMIR" Then
    If opt = 3 Then
        ExibeMensagem "Deve ser selecionada alguma Origem"
        Exit Sub
    End If
'    strSQL = " sp_OrdemServico '','',0,''," & opt & ",0,0,'IMPRIMIR'"
    strSql = gstrStoredProcedure("sp_OrdemServico", "'','',0,''," & opt & ",0,0,'IMPRIMIR'", True)
    ImprimeRelatorio rptOSdeFiscalizacao, strSql
End If
HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
End Sub

Private Sub optbytOrigem_Click(Index As Integer)
dbc_strInscricaoCadastral.Enabled = True
dbc_strInscricaoCadastral.BoundText = ""
Set dbc_strInscricaoCadastral.RowSource = Nothing
TrocaCorObjeto dbc_strInscricaoCadastral, False
txtstrRazao.Enabled = True
TrocaCorObjeto txtstrRazao, False
LimpaEndereco
opt = Index
mblnAlterando = False

If Index = 0 Then 'Economico
    QueryNovo
    tdb_OrdemServico.Columns(1).DataField = "strInscricaoCadastral"
    dbc_strInscricaoCadastral.Tag = "SELECT PKId, strInscricaoCadastral FROM " & gstrEconomico & " ;strInscricaoCadastral"
    LeDaTabelaParaObj "", tdb_OrdemServico, tdb_OS_Economico
    GridFiscal 0
Else
    QueryNovo
    tdb_OrdemServico.Columns(1).DataField = "strInscricaoAnterior"
    GridFiscal 1
    If Index = 1 Then 'Imobiliario Urbano
        dbc_strInscricaoCadastral.Tag = "SELECT PKId, strInscricaoAnterior FROM " & gstrImobiliario & " ;strInscricaoAnterior"
        LeDaTabelaParaObj "", tdb_OrdemServico, tdb_OS_ImobiliarioUrbano
    Else 'Imobiliario Rural Principal
        dbc_strInscricaoCadastral.Tag = "SELECT PKId, strInscricaoAnterior FROM " & gstrImobiliarioRural & " ;strInscricaoAnterior"
        LeDaTabelaParaObj "", tdb_OrdemServico, tdb_OS_ImobiliarioRural
    End If
End If
End Sub

Private Sub tdb_Fiscal_AfterColUpdate(ByVal ColIndex As Integer)
tdb_Fiscal.Update
End Sub

Private Sub tdb_Fiscal_AfterUpdate()
tdb_Fiscal.Update
End Sub

Private Sub tdb_Fiscal_KeyPress(KeyAscii As Integer)
    Select Case tdb_Fiscal.Col
        Case 3
            CaracterValido KeyAscii, "H", tdb_Fiscal
        Case 4
            CaracterValido KeyAscii, "D", tdb_Fiscal
    End Select
End Sub

Private Sub TDB_OrdemServico_Click()
    mblnPrimeiraVez = True
    With tdb_OrdemServico
        If Not .EOF And Not .BOF Then
            If .Bookmark = 1 Then
                tdb_OrdemServico_RowColChange 0, 0
            End If
        End If
    End With
End Sub

Sub tdb_OrdemServico_DblClick()
    MantemForm gstrAplicar
End Sub

Private Sub tdb_OrdemServico_FilterChange()
    gblnFilraCampos tdb_OrdemServico
End Sub

Private Sub tdb_OrdemServico_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    With tdb_OrdemServico
        If Not .EOF And Not .BOF Then
            txtPKId.Text = .Columns("PKID").Value

            If mblnPrimeiraVez Then
                LeDadosDoTDB
                If opt = 0 Then
                    ProprietarioEconomico
                    GridFiscalInscricaoCadastral opt
                Else
                    ProprietarioImobiliario opt
                    GridFiscalInscricaoCadastral 1
                End If
                
                gCorLinhaSelecionada tdb_OrdemServico

                HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar

                If mobjAux Is Nothing Then
                    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
                Else
                    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
                End If
                mblnSelecionou = True
                mblnAlterando = True
            End If

        End If
    End With
End Sub

Private Function tdb_OS_Economico() As String
    Dim strSql As String
    strSql = strSql & " SELECT OS.PKId, EC.strInscricaoCadastral, CC.strNome "
    strSql = strSql & " FROM " & gstrOrdemServico & " OS,"
    strSql = strSql & gstrEconomico & " EC,"
    strSql = strSql & gstrContribuinte & " CC "
    strSql = strSql & " WHERE OS.bytOrigem = 0 "
    strSql = strSql & " AND OS.intInscricaoCadastral = EC.PKId "
    strSql = strSql & " AND EC.intContribuinte = CC.PKId "
    tdb_OS_Economico = strSql
End Function

Private Function tdb_OS_ImobiliarioUrbano() As String
    Dim strSql As String
    strSql = strSql & " SELECT OS.PKId, IU.strInscricaoAnterior, CC.strNome "
    strSql = strSql & " FROM " & gstrOrdemServico & " OS,"
    strSql = strSql & gstrImobiliario & " IU,"
    strSql = strSql & gstrContribuinte & " CC "
    strSql = strSql & " WHERE OS.bytOrigem = 1 "
    strSql = strSql & " AND OS.intInscricaoCadastral = IU.PKId "
    strSql = strSql & " AND IU.intContribuinte = CC.PKId "
    tdb_OS_ImobiliarioUrbano = strSql
End Function

Private Function tdb_OS_ImobiliarioRural() As String
    Dim strSql As String
    strSql = strSql & " SELECT OS.PKId, IR.strInscricaoAnterior, CC.strNome "
    strSql = strSql & " FROM " & gstrOrdemServico & " OS,"
    strSql = strSql & gstrImobiliarioRural & " IR,"
    strSql = strSql & gstrContribuinte & " CC "
    strSql = strSql & " WHERE OS.bytOrigem = 2 "
    strSql = strSql & " AND OS.intInscricaoCadastral = IR.PKId "
    strSql = strSql & " AND IR.intContribuinte = CC.PKId "
    tdb_OS_ImobiliarioRural = strSql
End Function

Private Function ProprietarioEconomico()

'******************************************************************************************
' Data: 02/05/2003
' Alteração: - Substituição do comando nativo CONVERT do SQL Server pela função gstrCONVERT
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 05/05/2003
' Alteração: - Substituição do comando nativo de concatenação de strings (+) do SQL Server
'            pela variável strCONCAT.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 05/05/2003
' Alteração: - Retirado comando "AS" utilizado para dar apelidos às tabelas, o qual não é
'            permitido pelo Oracle, da cláusula FROM.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Dim strSql          As String
    Dim adoResultado    As ADODB.Recordset
    
    Set gobjBanco = New clsBanco

    strSql = ""
    strSql = strSql & " SELECT CO.PKId AS CodigoContribuinte, CO.strNome, LG.strDescricao AS Logradouro, A.intNumero, "
    strSql = strSql & " A.strComplemento , A.intCep, BA.strDescricao AS Bairro, UF.strSigla AS UF, CO.strCNPJCPF, "
    strSql = strSql & " CO.strDistritoC, CI.strDescricao AS Municipio "
    strSql = strSql & " FROM "
    
'    strSQL = strSQL & gstrEconomico & " AS A, "
    strSql = strSql & gstrEconomico & " A, "
    
'    strSQL = strSQL & gstrContribuinte & " AS CO, "
    strSql = strSql & gstrContribuinte & " CO, "
'    strSQL = strSQL & gstrComposicaoDaReceita & " AS CR, "
    strSql = strSql & gstrComposicaoDaReceita & " CR, "
'    strSQL = strSQL & gstrLogradouro & " AS LG, "
    strSql = strSql & gstrLogradouro & " LG, "
'    strSQL = strSQL & gstrCidade & " AS CI, "
    strSql = strSql & gstrCidade & " CI, "
'    strSQL = strSQL & gstrBairro & " AS BA, "
    strSql = strSql & gstrBairro & " BA, "
'    strSQL = strSQL & gstrUF & " AS UF "
    strSql = strSql & gstrUF & " UF "
    strSql = strSql & " WHERE "
    strSql = strSql & " CO.PKId = A.intContribuinte "
    strSql = strSql & " AND CI.PKId = CO.intMunicipioC "
    strSql = strSql & " AND LG.PKId = A.intLogradouro "
    strSql = strSql & " AND BA.PKId = A.intBairro "
    strSql = strSql & " AND UF.PKId = CO.intUF "
    strSql = strSql & " AND A.PKID = " & dbc_strInscricaoCadastral.BoundText
    
    LimpaEndereco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            With adoResultado
                txt_strCNPJCPFP = gstrCGCCPFFormatado(!strCNPJCPF)
                txt_strNome.Text = !STRNOME
                txt_intContribuinte.Text = !CodigoContribuinte
                txt_Logradouro.Text = !Logradouro
                txt_Numero.Text = !intNumero
                txt_Complemento.Text = gstrENulo(!strComplemento)
                txt_Bairro.Text = gstrENulo(!Bairro)
                txt_Municipio.Text = gstrENulo(!Municipio)
                txt_Distrito.Text = gstrENulo(!strDistritoC)
                txt_UF.Text = gstrENulo(!UF)
                txt_Cep.Text = gstrCEPFormatado(!intCep)
            End With
        End If
    End If
    
    
    Set gobjBanco = New clsBanco

    strSql = ""
'    strSql = strSql & " SELECT CONVERT(VARCHAR, AEC.intCodigo) + ' - ' + AEC.strDescricao AS strAtividadePrincipal "
    strSql = strSql & " SELECT " & gstrCONVERT(CDT_VARCHAR, "AEC.intCodigo") & strCONCAT & " ' - ' " & strCONCAT & " AEC.strDescricao AS strAtividadePrincipal "
    strSql = strSql & " FROM "
'    strSQL = strSQL & gstrEconomico & " AS A, "
    strSql = strSql & gstrEconomico & " A, "
'    strSQL = strSQL & gstrAtividadeDaEmpresa & " AS AE, "
    strSql = strSql & gstrAtividadeDaEmpresa & " AE, "
'    strSQL = strSQL & gstrAtividadeEC & " AS AEC "
    strSql = strSql & gstrAtividadeEC & " AEC "
    strSql = strSql & " WHERE "
    strSql = strSql & " AE.intEconomico = A.PKId AND AE.blnPrincipal = 1 "
    strSql = strSql & " AND AEC.PKId = AE.intAtividade "
    strSql = strSql & " AND A.PKID = " & dbc_strInscricaoCadastral.BoundText
    
    txt_strAtividadePrincipal.Text = ""
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            With adoResultado
                txt_strAtividadePrincipal.Text = !strAtividadePrincipal
            End With
        End If
    End If
    
    Set gobjBanco = New clsBanco

    strSql = ""
    strSql = strSql & " SELECT AB.strDescricao AS strAtividadeBasica "
    strSql = strSql & " FROM "
'    strSQL = strSQL & gstrEconomico & " AS A, "
    strSql = strSql & gstrEconomico & " A, "
'    strSQL = strSQL & gstrAtividadeBasica & " AS AB "
    strSql = strSql & gstrAtividadeBasica & " AB "
    strSql = strSql & " WHERE AB.PKId = A.intAtividadeBasica "
    strSql = strSql & " AND A.PKID = " & dbc_strInscricaoCadastral.BoundText
    
    txt_strAtividadeBasica.Text = ""
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            With adoResultado
                txt_strAtividadeBasica.Text = !strAtividadeBasica
            End With
        End If
    End If
End Function

Private Function ProprietarioImobiliario(Index As Integer)

'******************************************************************************************
' Data: 05/05/2003
' Alteração: - Retirado comando "AS" utilizado para dar apelidos às tabelas, o qual não é
'            permitido pelo Oracle, da cláusula FROM.
' Responsável: Everton Bianchini
'******************************************************************************************

    Dim strSql          As String
    Dim adoResultado    As ADODB.Recordset
    
    Set gobjBanco = New clsBanco

    strSql = ""
    strSql = strSql & " SELECT CO.PKId AS CodigoContribuinte, CO.strNome, LG.strDescricao AS Logradouro, A.intNumero, "
    strSql = strSql & " A.strComplemento , A.intCep, BA.strDescricao AS Bairro, UF.strSigla AS UF, A.strCNPJCPF, "
    strSql = strSql & " CO.strDistritoC, CI.strDescricao AS Municipio, A.dblValorITBI "
    strSql = strSql & IIf(opt = 1, ", CR.strDescricao AS Composicao, A.intComposicao ", "")
    strSql = strSql & " FROM "
    If opt = 2 Then
'        strSQL = strSQL & gstrImobiliarioRural & " AS A, "
        strSql = strSql & gstrImobiliarioRural & " A, "
    Else
'        strSQL = strSQL & gstrImobiliario & " AS A, "
        strSql = strSql & gstrImobiliario & " A, "
    End If
'    strSQL = strSQL & gstrContribuinte & " AS CO, "
    strSql = strSql & gstrContribuinte & " CO, "
'    strSQL = strSQL & IIf(opt = 1, gstrComposicaoDaReceita & " AS CR, ", "")
    strSql = strSql & IIf(opt = 1, gstrComposicaoDaReceita & " CR, ", "")
'    strSQL = strSQL & gstrLogradouro & " AS LG, "
    strSql = strSql & gstrLogradouro & " LG, "
'    strSQL = strSQL & gstrCidade & " AS CI, "
    strSql = strSql & gstrCidade & " CI, "
'    strSQL = strSQL & gstrBairro & " AS BA, "
    strSql = strSql & gstrBairro & " BA, "
'    strSQL = strSQL & gstrUF & " AS UF "
    strSql = strSql & gstrUF & " UF "
    strSql = strSql & " WHERE "
    strSql = strSql & " CO.PKId = A.intContribuinte "
    strSql = strSql & " AND CI.PKId = CO.intMunicipioC "
    strSql = strSql & IIf(opt = 1, " AND CR.PKId = A.intComposicao ", "")
    strSql = strSql & " AND LG.PKId = A.intLogradouro "
    strSql = strSql & " AND BA.PKId = A.intBairro "
    strSql = strSql & " AND UF.PKId = A.intUF "
    strSql = strSql & " AND A.PKID = " & dbc_strInscricaoCadastral.BoundText
    
    LimpaEndereco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            With adoResultado
                txt_strCNPJCPFP = gstrCGCCPFFormatado(gstrENulo(!strCNPJCPF))
                txt_strNome.Text = gstrENulo(!STRNOME)
                txt_intContribuinte.Text = gstrENulo(!CodigoContribuinte)
                txt_Logradouro.Text = gstrENulo(!Logradouro)
                txt_Numero.Text = gstrENulo(!intNumero)
                txt_Complemento.Text = gstrENulo(!strComplemento)
                txt_Bairro.Text = gstrENulo(!Bairro)
                txt_Municipio.Text = gstrENulo(!Municipio)
                txt_Distrito.Text = gstrENulo(!strDistritoC)
                txt_UF.Text = gstrENulo(!UF)
                txt_Cep.Text = gstrCEPFormatado(gstrENulo(!intCep))
            End With
        End If
    End If
    
End Function

Private Function LimpaEndereco()
    txt_strCNPJCPFP = ""
    txt_strNome.Text = ""
    txt_intContribuinte.Text = ""
    txt_Logradouro.Text = ""
    txt_Numero.Text = ""
    txt_Complemento.Text = ""
    txt_Bairro.Text = ""
    txt_Municipio.Text = ""
    txt_Distrito.Text = ""
    txt_UF.Text = ""
    txt_Cep.Text = ""
    txt_strAtividadeBasica.Text = ""
    txt_strAtividadePrincipal.Text = ""
End Function

Private Sub GridFiscal(bytTipo As Integer)
Dim strSql As String
Dim adoRec As ADODB.Recordset
Dim varAux As String

On Error GoTo Err_Handle

Set MatFiscais = New XArrayDB
MatFiscais.Clear

MatFiscais.ReDim 0, 0, 0, 4

strSql = ""
strSql = strSql & " SELECT FC.PKId, CC.strNome "
strSql = strSql & " FROM " & gstrFiscais & " FC,"
strSql = strSql & gstrContribuinte & " CC"
strSql = strSql & " WHERE CC.PKId = FC.intContribuinte "
strSql = strSql & " AND FC.bytTipoFiscal = " & bytTipo
strSql = strSql & " Order By strNome"

Set gobjBanco = New clsBanco

If gobjBanco.CriaADO(strSql, 10, adoRec) Then
    With adoRec
        If Not .EOF Then
            MatFiscais.ReDim 0, .RecordCount - 1, 0, 5
            Do While Not .EOF
                varAux = !Pkid
                MatFiscais(.AbsolutePosition - 1, 0) = varAux
                
                MatFiscais(.AbsolutePosition - 1, 1) = False
                
                varAux = !STRNOME
                MatFiscais(.AbsolutePosition - 1, 2) = varAux
            
                MatFiscais(.AbsolutePosition - 1, 3) = ""
                
                MatFiscais(.AbsolutePosition - 1, 4) = ""
                
                .MoveNext
            Loop
        End If
    End With
End If

Set tdb_Fiscal.Array = MatFiscais
tdb_Fiscal.ReBind
tdb_Fiscal.Refresh

Exit Sub
Err_Handle:
End Sub

Private Function DadosOk() As Boolean
    Dim blnSelecionouTDB As Boolean
    Dim i As Integer
    If opt = 3 Then
        ExibeMensagem "Deve ser selecionada alguma origem "
        DadosOk = False
        Exit Function
    End If
    If dbc_strInscricaoCadastral.BoundText = "" Then
        ExibeMensagem "Deve ser selecionada alguma Inscrição Cadastral "
        DadosOk = False
        Exit Function
    End If
    If txtstrRazao.Text = "" Then
        ExibeMensagem "Deve ser digitada alguma Razão para a Fiscalização "
        DadosOk = False
        Exit Function
    End If
    For i = 0 To MatFiscais.Count(1) - 1
        If MatFiscais(i, 1) = -1 Then
            blnSelecionouTDB = True
            If MatFiscais(i, 3) = "" Then
                ExibeMensagem "Para o funcionário " & MatFiscais(i, 2) & " deve ser digitado um Horário "
                DadosOk = False
                Exit Function
            End If
            If MatFiscais(i, 4) = "" Then
                ExibeMensagem "Para o funcionário " & MatFiscais(i, 2) & " deve ser digitado uma Data "
                DadosOk = False
                Exit Function
            End If
        End If
    Next
    If Not blnSelecionouTDB Then
        ExibeMensagem "Deve ser Selecionado Algum Fiscal"
        DadosOk = False
        Exit Function
    End If
    DadosOk = True
End Function

Private Function strPKId() As String
    Dim strSql As String
    Dim i As Integer
    strSql = ""
    For i = 0 To MatFiscais.Count(1) - 1
        If MatFiscais(i, 1) = -1 Then
            If strSql <> "" Then
               strSql = strSql & ","
            End If
            strSql = strSql & MatFiscais(i, 0)
        End If
    Next

    strPKId = strSql
End Function

Private Function strDTM() As String

'******************************************************************************************
' Data: 12/05/2003
' Alteração: - Adaptada concatenação do campo data para o Oracle.
' Responsável: Everton Bianchini
'******************************************************************************************

    Dim strSql As String
    Dim i As Integer
    strSql = ""
    For i = 0 To MatFiscais.Count(1) - 1
        If MatFiscais(i, 1) = -1 Then
            If strSql <> "" Then
               strSql = strSql & ","
            End If
            
            If (bytDBType = EDatabases.SQLServer) Then
                strSql = strSql & Mid(gstrConvDtParaSql(MatFiscais(i, 4)), 2, 10) & " " & MatFiscais(i, 3) & ":00"
            
            ElseIf (bytDBType = EDatabases.Oracle) Then
                strSql = strSql & Format(MatFiscais(i, 4), "yyyy/mm/dd") & " " & MatFiscais(i, 3) & ":00"
            
            End If
        End If
    Next
    strDTM = strSql
End Function

Private Sub QueryNovo()
dbc_strInscricaoCadastral.BoundText = ""
LimpaEndereco
txtstrRazao = ""
If opt <> 3 Then
    GridFiscal opt
    mblnAlterando = False
End If
End Sub

Private Function LeDadosDoTDB()
Dim strSql As String
Dim adoResultado As ADODB.Recordset

strSql = strSql & " SELECT intInscricaoCadastral, strRazaoDaFiscalizacao "
strSql = strSql & " FROM " & gstrOrdemServico
strSql = strSql & " WHERE PKId = " & txtPKId

Set gobjBanco = New clsBanco

If gobjBanco.CriaADO(strSql, 10, adoResultado) Then
    With adoResultado
        If Not .EOF Then
            PreencherListaDeOpcoes dbc_strInscricaoCadastral, !intInscricaoCadastral
            dbc_strInscricaoCadastral.BoundText = (!intInscricaoCadastral)
            txtstrRazao = (!strRazaoDaFiscalizacao)
        End If
    End With
End If
Set gobjBanco = Nothing
Set adoResultado = Nothing
End Function


Private Sub GridFiscalInscricaoCadastral(bytTipo As Integer)

'******************************************************************************************
' Data: 12/05/2003
' Alteração: - Incluída chamada à função gstrConvert a fim de permitir a execução do
'            comando SELECT devido à incompatibilidade do comando UNION entre o SQL Server
'            e o Oracle.
' Responsável: Everton Bianchini
'******************************************************************************************

Dim strSql As String
Dim adoRec As ADODB.Recordset
Dim varAux As String

On Error GoTo Err_Handle

Set MatFiscais = New XArrayDB
MatFiscais.Clear

MatFiscais.ReDim 0, 0, 0, 4

strSql = " SELECT FC.PKId, 1 as Boolean, " & _
        " CC.strNome as strNome, OSF.dtmDtHorario As dtmDtHorario" & _
        " FROM " & gstrOrdemServico & " OS," & _
        gstrOrdemServicoFiscal & " OSF," & _
        gstrFiscais & " FC," & _
        gstrContribuinte & " CC" & _
        " WHERE CC.PKId = FC.intContribuinte " & _
        " AND FC.PKId = OSF.intFiscal " & _
        " AND OS.PKId = OSF.intOrdemServico " & _
        " AND OS.intInscricaoCadastral = " & dbc_strInscricaoCadastral.BoundText & _
        " AND OS.bytOrigem = " & opt & _
        " UNION "

'strSql = strSql & " SELECT FC.PKId, 0 as Boolean, CC.strNome , NULL As dtmDtHorario "
strSql = strSql & " SELECT FC.PKId, 0 as Boolean, CC.strNome , " & gstrCONVERT(CDT_DATETIME, "NULL") & " As dtmDtHorario " & _
        " FROM " & gstrFiscais & " FC," & _
        gstrContribuinte & " CC" & _
        " WHERE CC.PKId = FC.intContribuinte " & _
        " AND FC.PKId NOT IN("

strSql = strSql & " SELECT FC.PKId " & _
        " FROM " & gstrOrdemServico & " OS," & _
        gstrOrdemServicoFiscal & " OSF," & _
        gstrFiscais & " FC " & _
        " WHERE FC.PKId = OSF.intFiscal " & _
        " AND OS.PKId = OSF.intOrdemServico " & _
        " AND OS.intInscricaoCadastral = " & dbc_strInscricaoCadastral.BoundText & _
        " AND OS.bytOrigem = " & opt & ")" & _
        " AND FC.bytTipoFiscal = " & bytTipo & _
        " Order By strNome "

Set gobjBanco = New clsBanco
If gobjBanco.CriaADO(strSql, 10, adoRec) Then
    With adoRec
        If Not .EOF Then
            MatFiscais.ReDim 0, .RecordCount - 1, 0, 5
            Do While Not .EOF

                varAux = !Pkid
                MatFiscais(.AbsolutePosition - 1, 0) = varAux
            
                MatFiscais(.AbsolutePosition - 1, 1) = CBool(!Boolean)

                varAux = !STRNOME
                MatFiscais(.AbsolutePosition - 1, 2) = varAux

                varAux = IIf(IsNull(!dtmDtHorario), "", Format(!dtmDtHorario, "hh:mm"))
                MatFiscais(.AbsolutePosition - 1, 3) = varAux

                varAux = IIf(IsNull(!dtmDtHorario), "", Format(!dtmDtHorario, "dd/mm/yyyy"))
                MatFiscais(.AbsolutePosition - 1, 4) = varAux
                
                .MoveNext
            Loop
        End If
    End With
End If

Set tdb_Fiscal.Array = MatFiscais
tdb_Fiscal.ReBind
tdb_Fiscal.Refresh

Exit Sub
Err_Handle:
End Sub

Private Sub txtstrRazao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrRazao
End Sub
