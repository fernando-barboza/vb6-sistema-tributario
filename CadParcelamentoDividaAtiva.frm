VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmCadParcelamentoDividaAtiva 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parcelamento da Dívida Ativa"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8550
   Icon            =   "CadParcelamentoDividaAtiva.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   8550
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   5775
      Left            =   105
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   120
      Width           =   8355
      _ExtentX        =   14737
      _ExtentY        =   10186
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Parcelamento da Dívida Ativa"
      TabPicture(0)   =   "CadParcelamentoDividaAtiva.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl_strInscricaoCadastral"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl_intOcorrencia"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "dbc_intOcorrencia"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "dbc_strInscricaoCadastral"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fra_Frame"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "fra_Inscricao"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Emissão de Guias de Arrecadação "
      TabPicture(1)   =   "CadParcelamentoDividaAtiva.frx":105E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fra_EmissaoDeGuias"
      Tab(1).ControlCount=   1
      Begin VB.Frame fra_EmissaoDeGuias 
         Height          =   5145
         Left            =   -74805
         TabIndex        =   33
         Top             =   405
         Width           =   7935
         Begin MSDataListLib.DataCombo dbc_strInscricaoFinal 
            Height          =   315
            Left            =   2220
            TabIndex        =   13
            Top             =   780
            Width           =   5235
            _ExtentX        =   9234
            _ExtentY        =   556
            _Version        =   393216
            IntegralHeight  =   0   'False
            Text            =   ""
         End
         Begin VB.TextBox txt_DataDeVencimento 
            Height          =   285
            Left            =   6420
            MaxLength       =   15
            TabIndex        =   15
            Top             =   1140
            Width           =   1035
         End
         Begin VB.TextBox txt_intExercicio 
            Height          =   285
            Left            =   2220
            MaxLength       =   4
            TabIndex        =   14
            Top             =   1140
            Width           =   525
         End
         Begin VB.Frame fra_Mensagem1 
            Caption         =   "Mensagem 1                                "
            Height          =   1560
            Left            =   510
            TabIndex        =   36
            Top             =   1560
            Width           =   6945
            Begin VB.CheckBox chk_EmBranco1 
               Caption         =   "Em branco"
               Height          =   195
               Left            =   1350
               TabIndex        =   16
               Top             =   0
               Width           =   1095
            End
            Begin VB.TextBox txt_Mensagem1 
               Height          =   795
               Left            =   120
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   18
               TabStop         =   0   'False
               Top             =   645
               Width           =   6675
            End
            Begin MSDataListLib.DataCombo dbc_intMensagem1 
               Height          =   315
               Left            =   1080
               TabIndex        =   17
               Top             =   270
               Width           =   5715
               _ExtentX        =   10081
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               Text            =   ""
            End
            Begin VB.Label lbl_Label 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Mensagem"
               Height          =   195
               Left            =   120
               TabIndex        =   37
               Top             =   390
               Width           =   780
            End
         End
         Begin VB.Frame fra_Mensagem2 
            Caption         =   "Mensagem 2                                "
            Height          =   1590
            Left            =   510
            TabIndex        =   34
            Top             =   3270
            Width           =   6945
            Begin VB.CheckBox chk_EmBranco2 
               Caption         =   "Em branco"
               Height          =   195
               Left            =   1350
               TabIndex        =   19
               Top             =   0
               Width           =   1095
            End
            Begin VB.TextBox txt_Mensagem2 
               Height          =   795
               Left            =   120
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   21
               TabStop         =   0   'False
               Top             =   660
               Width           =   6675
            End
            Begin MSDataListLib.DataCombo dbc_intMensagem2 
               Height          =   315
               Left            =   1080
               TabIndex        =   20
               Top             =   270
               Width           =   5715
               _ExtentX        =   10081
               _ExtentY        =   556
               _Version        =   393216
               Style           =   2
               Text            =   ""
            End
            Begin VB.Label lbl_Label2 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Mensagem"
               Height          =   195
               Left            =   120
               TabIndex        =   35
               Top             =   390
               Width           =   780
            End
         End
         Begin MSDataListLib.DataCombo dbc_strInscricaoInicial 
            Height          =   315
            Left            =   2220
            TabIndex        =   12
            Top             =   420
            Width           =   5235
            _ExtentX        =   9234
            _ExtentY        =   556
            _Version        =   393216
            IntegralHeight  =   0   'False
            Text            =   ""
         End
         Begin VB.Label lbl_alabel 
            AutoSize        =   -1  'True
            Caption         =   "Data de Vencimento"
            Height          =   195
            Left            =   4860
            TabIndex        =   41
            Top             =   1200
            Width           =   1455
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Exercício"
            Height          =   195
            Left            =   1455
            TabIndex        =   40
            Top             =   1215
            Width           =   675
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Inscrição Cadastral Inicial"
            Height          =   195
            Left            =   330
            TabIndex        =   39
            Top             =   510
            Width           =   1800
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Inscrição Cadastral Final"
            Height          =   195
            Left            =   405
            TabIndex        =   38
            Top             =   870
            Width           =   1725
         End
      End
      Begin VB.Frame fra_Inscricao 
         Height          =   645
         Left            =   105
         TabIndex        =   30
         Top             =   465
         Width           =   8085
         Begin VB.OptionButton optbitTipoDeInscricao 
            Caption         =   "Imobiliário Urbano"
            Height          =   195
            Index           =   0
            Left            =   60
            TabIndex        =   0
            Top             =   270
            Value           =   -1  'True
            Width           =   1605
         End
         Begin VB.OptionButton optbitTipoDeInscricao 
            Caption         =   "Imobiliário Rural"
            CausesValidation=   0   'False
            Height          =   195
            Index           =   1
            Left            =   1650
            TabIndex        =   1
            Top             =   270
            Width           =   1425
         End
         Begin VB.OptionButton optbitTipoDeInscricao 
            Caption         =   "Econômico"
            CausesValidation=   0   'False
            Height          =   195
            Index           =   2
            Left            =   3105
            TabIndex        =   2
            Top             =   270
            Width           =   1155
         End
         Begin VB.OptionButton optbitTipoDeInscricao 
            Caption         =   "Contribuição de Melhorias"
            CausesValidation=   0   'False
            Height          =   195
            Index           =   3
            Left            =   4245
            TabIndex        =   3
            Top             =   270
            Width           =   2205
         End
         Begin VB.OptionButton optbitTipoDeInscricao 
            Caption         =   "Receitas Diversas"
            CausesValidation=   0   'False
            Height          =   195
            Index           =   4
            Left            =   6420
            TabIndex        =   4
            Top             =   270
            Width           =   1605
         End
      End
      Begin VB.Frame fra_Frame 
         Height          =   3060
         Left            =   150
         TabIndex        =   23
         Top             =   2235
         Width           =   7995
         Begin VB.TextBox txt_dtmDataVencimento 
            Height          =   285
            Left            =   6285
            MaxLength       =   15
            TabIndex        =   8
            Top             =   225
            Width           =   1005
         End
         Begin VB.TextBox txt_intIntervalo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   6285
            MaxLength       =   15
            TabIndex        =   10
            Top             =   585
            Width           =   1005
         End
         Begin VB.TextBox txt_dblDesconto 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1770
            MaxLength       =   3
            TabIndex        =   9
            Top             =   585
            Width           =   1035
         End
         Begin VB.TextBox txt_intNumeroParcelas 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1770
            MaxLength       =   3
            TabIndex        =   7
            Top             =   225
            Width           =   1035
         End
         Begin TrueOleDBGrid70.TDBGrid tdb_Exercicios 
            Height          =   1905
            Left            =   1770
            TabIndex        =   11
            Top             =   945
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   3360
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Exercício"
            Columns(0).DataField=   "intExercicio"
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
            Columns(1).DataField=   ""
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   2
            Splits(0)._UserFlags=   1
            Splits(0).MarqueeStyle=   5
            Splits(0).RecordSelectors=   0   'False
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).ScrollBars=   2
            Splits(0).DividerColor=   12632256
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=2"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=1588"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1508"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(6)=   "Column(1).Width=450"
            Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=370"
            Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(10)=   "Column(1).AllowSizing=0"
            Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
            Splits.Count    =   1
            PrintInfos(0)._StateFlags=   0
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
            DirectionAfterEnter=   0
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
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33"
            _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34"
            _StyleDefs(9)   =   "FooterStyle:id=3,.parent=1,.namedParent=35"
            _StyleDefs(10)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(11)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
            _StyleDefs(12)  =   "EditorStyle:id=7,.parent=1"
            _StyleDefs(13)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
            _StyleDefs(14)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
            _StyleDefs(15)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
            _StyleDefs(16)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
            _StyleDefs(17)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
            _StyleDefs(18)  =   "Splits(0).Style:id=13,.parent=1"
            _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
            _StyleDefs(20)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
            _StyleDefs(21)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(22)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(23)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.bgcolor=&H0&"
            _StyleDefs(24)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(25)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
            _StyleDefs(26)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(27)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=32,.parent=13"
            _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
            _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
            _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
            _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
            _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
            _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
            _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
            _StyleDefs(38)  =   "Named:id=33:Normal"
            _StyleDefs(39)  =   ":id=33,.parent=0"
            _StyleDefs(40)  =   "Named:id=34:Heading"
            _StyleDefs(41)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(42)  =   ":id=34,.wraptext=-1"
            _StyleDefs(43)  =   "Named:id=35:Footing"
            _StyleDefs(44)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(45)  =   "Named:id=36:Selected"
            _StyleDefs(46)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(47)  =   "Named:id=37:Caption"
            _StyleDefs(48)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(49)  =   "Named:id=38:HighlightRow"
            _StyleDefs(50)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(51)  =   "Named:id=39:EvenRow"
            _StyleDefs(52)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(53)  =   "Named:id=40:OddRow"
            _StyleDefs(54)  =   ":id=40,.parent=33"
            _StyleDefs(55)  =   "Named:id=41:RecordSelector"
            _StyleDefs(56)  =   ":id=41,.parent=34"
            _StyleDefs(57)  =   "Named:id=42:FilterBar"
            _StyleDefs(58)  =   ":id=42,.parent=33"
         End
         Begin VB.Label lbl_dias 
            AutoSize        =   -1  'True
            Caption         =   "dias."
            Height          =   195
            Left            =   7350
            TabIndex        =   29
            Top             =   660
            Width           =   330
         End
         Begin VB.Label lbl_intIntervalo 
            AutoSize        =   -1  'True
            Caption         =   "Intervalo entre Parcelas"
            Height          =   195
            Left            =   4455
            TabIndex        =   28
            Top             =   660
            Width           =   1680
         End
         Begin VB.Label lbl_dblAliquota 
            AutoSize        =   -1  'True
            Caption         =   "Desconto"
            Height          =   195
            Left            =   975
            TabIndex        =   27
            Top             =   660
            Width           =   690
         End
         Begin VB.Label lbl_intNumeroParcelas 
            AutoSize        =   -1  'True
            Caption         =   "Número de Parcelas"
            Height          =   195
            Left            =   225
            TabIndex        =   26
            Top             =   300
            Width           =   1440
         End
         Begin VB.Label lbl_dtmDataVencimento 
            AutoSize        =   -1  'True
            Caption         =   "Data de Vencimento"
            Height          =   195
            Left            =   4680
            TabIndex        =   25
            Top             =   285
            Width           =   1455
         End
         Begin VB.Label lbl_p1 
            AutoSize        =   -1  'True
            Caption         =   "%"
            Height          =   195
            Left            =   2880
            TabIndex        =   24
            Top             =   675
            Width           =   120
         End
      End
      Begin MSDataListLib.DataCombo dbc_strInscricaoCadastral 
         Height          =   315
         Left            =   1680
         TabIndex        =   5
         Top             =   1320
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   556
         _Version        =   393216
         IntegralHeight  =   0   'False
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbc_intOcorrencia 
         Height          =   315
         Left            =   1680
         TabIndex        =   6
         Top             =   1695
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label lbl_intOcorrencia 
         AutoSize        =   -1  'True
         Caption         =   "Ocorrência"
         Height          =   195
         Left            =   750
         TabIndex        =   32
         Top             =   1800
         Width           =   780
      End
      Begin VB.Label lbl_strInscricaoCadastral 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Inscrição Cadastral"
         Height          =   195
         Left            =   180
         TabIndex        =   31
         Top             =   1410
         Width           =   1350
      End
   End
End
Attribute VB_Name = "frmCadParcelamentoDividaAtiva"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim xExercicio                  As XArrayDB

Private Sub dbc_intMensagem1_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_intMensagem1, Me, , KeyCode, Shift
End Sub

Private Sub dbc_intMensagem2_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_intMensagem2, Me, , KeyCode, Shift
End Sub

Private Sub dbc_intOcorrencia_Click(Area As Integer)
    DropDownDataCombo dbc_intOcorrencia, Me, Area
End Sub

Private Sub dbc_intOcorrencia_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_intOcorrencia, Me, , KeyCode, Shift
End Sub

Private Sub dbc_strInscricaoCadastral_Click(Area As Integer)
    DropDownDataCombo dbc_strInscricaoCadastral, Me, Area
End Sub

Private Sub dbc_strInscricaoCadastral_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_strInscricaoCadastral, Me, , KeyCode, Shift
End Sub

Private Sub dbc_strInscricaoFinal_Click(Area As Integer)
    DropDownDataCombo dbc_strInscricaoFinal, Me, Area
End Sub

Private Sub dbc_strInscricaoFinal_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_strInscricaoFinal, Me, , KeyCode, Shift
End Sub

Private Sub dbc_strInscricaoInicial_Click(Area As Integer)
    DropDownDataCombo dbc_strInscricaoInicial, Me, Area
End Sub

Private Sub dbc_strInscricaoInicial_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_strInscricaoInicial, Me, , KeyCode, Shift
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 643
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrNovo, gstrSalvar, gstrDeletar
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrCalcularReajuste
End Sub

Private Sub MontaGridExercicio()
    Dim intAno As Integer
    
    Set xExercicio = New XArrayDB
    xExercicio.Clear
    xExercicio.ReDim 1, 10, 0, 1
    For intAno = 1 To 10
        xExercicio(intAno, 0) = Year(Date) - intAno
    Next
    Set tdb_Exercicios.Array = xExercicio
End Sub

Private Sub Form_Load()
    tab_3dPasta.Tab = 0
    optbitTipoDeInscricao_Click (0)
    LeDaTabelaParaObj gstrOcorrencia, dbc_intOcorrencia, strQuerryOcorrencia
    MontaGridExercicio
    
    '''GUIA
    dbc_strInscricaoInicial.Tag = strQueryInscricaoGuia & ";strInscricaoCadastral"
    dbc_strInscricaoFinal.Tag = strQueryInscricaoGuia & ";strInscricaoCadastral"
    LeDaTabelaParaObj gstrMensagem, dbc_intMensagem1, strQueryMensagem
    LeDaTabelaParaObj gstrMensagem, dbc_intMensagem2, strQueryMensagem

End Sub

Private Function strQuerryOcorrencia() As String
    Dim strSQL As String
    strSQL = ""
    strSQL = strSQL & " SELECT PKId, strDescricao "
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrOcorrencia
    strSQL = strSQL & " WHERE "
    strSQL = strSQL & " intUtilizacaoDaOcorrencia = 1 "
    strSQL = strSQL & " ORDER BY strDescricao "
    strQuerryOcorrencia = strSQL
End Function

Private Function strQueryInscricaoGuia() As String

'******************************************************************************************
' Data: 02/05/2003
' Alteração: - Substituição do comando nativo CONVERT do SQL Server pela função gstrCONVERT
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Dim strSQL As String

    strSQL = ""
    strSQL = strSQL & " SELECT PKId, strInscricaoCadastral "
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrEconomico
    strSQL = strSQL & " WHERE "
    strSQL = strSQL & " dtmDataBaixa IS NULL " 'Verifica se existe data de baixa
    strSQL = strSQL & " ORDER BY "
'    strSql = strSql & " CONVERT(NUMERIC,strInscricaoCadastral) "
    strSQL = strSQL & gstrCONVERT(CDT_NUMERIC, "strInscricaoCadastral")
    strSQL = strSQL

    strQueryInscricaoGuia = strSQL
End Function

Private Function strQueryMensagem() As String

'******************************************************************************************
' Data: 05/05/2003
' Alteração: - Substituição do comando nativo de concatenação de strings (+) do SQL Server
'            pela variável strCONCAT.
' Responsável: Everton Bianchini
'******************************************************************************************

    Dim strSQL As String

    strSQL = ""
'    strSQL = strSQL & "SELECT PKId, ltrim(rtrim(PKId)) + ' - ' + ltrim(rtrim(strDescricao)) as Descricao "
    strSQL = strSQL & "SELECT PKId, ltrim(rtrim(PKId)) " & strCONCAT & " ' - ' " & strCONCAT & " ltrim(rtrim(strDescricao)) as Descricao "
    strSQL = strSQL & " FROM " & gstrMensagem
    strSQL = strSQL & " ORDER BY PKId "

    strQueryMensagem = strSQL
End Function

Private Sub Form_Unload(Cancel As Integer)
    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrSalvar, gstrNovo
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrCalcularReajuste
End Sub

Private Function blnValidaDados() As Boolean
    Dim intIndice As Integer
    
    If Not dbc_strInscricaoCadastral.MatchedWithList Then
        ExibeMensagem "A Inscrição Cadastral tem que ser selecionada."
        dbc_strInscricaoCadastral.SetFocus
        Exit Function
    End If
    If Not dbc_intOcorrencia.MatchedWithList Then
        ExibeMensagem "A Inscrição Final tem que ser selecionada."
        dbc_intOcorrencia.SetFocus
        Exit Function
    End If
    
    If Trim(txt_intNumeroParcelas.Text) = "" Then
        ExibeMensagem "O Nº de parcelas tem que ser digitado."
        txt_intNumeroParcelas.SetFocus
        Exit Function
    End If
    
    If Trim(txt_dtmDataVencimento.Text) = "" Then
        ExibeMensagem " A Data do Vencimento da 1ª parcela tem que ser digitada."
        txt_dtmDataVencimento.SetFocus
        Exit Function
    ElseIf Not gblnDataValida(txt_dtmDataVencimento.Text, True) Then
            txt_dtmDataVencimento.SetFocus
            Exit Function
    End If
    
    If Trim(txt_intIntervalo.Text) = "" Then
        ExibeMensagem "O intervalo tem que ser digitado."
        txt_intIntervalo.SetFocus
        Exit Function
    End If
    
    tdb_Exercicios.Update
    For intIndice = 1 To xExercicio.Count(1)
        If xExercicio(intIndice, 1) = -1 Then
            blnValidaDados = True
            Exit Function
        End If
    Next
    If Not blnValidaDados Then
        ExibeMensagem "Selecione um exercício para efetuar o cálculo!"
        tdb_Exercicios.SetFocus
    End If
    
End Function

Private Function BuscaPKIdReceita(intComposicaoDaReceita As Integer) As String
    Dim strSQL As String
    Dim strPKId As String
    Dim adoRec As ADODB.Recordset
    
    strSQL = strSQL & " SELECT A.PKId FROM "
    strSQL = strSQL & gstrReceita & " A,"
    strSQL = strSQL & gstrValorCompoRec & " B"
    strSQL = strSQL & " WHERE A.PKId = B.intReceita "
    strSQL = strSQL & " AND B.intComposicaoDaReceita = " & intComposicaoDaReceita
    strPKId = ""
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 10, adoRec) Then
        With adoRec
            If Not (.BOF And .EOF) Then
                .MoveFirst
                While Not .EOF
                    If strPKId <> "" Then
                        strPKId = strPKId & ", "
                    End If
                    strPKId = strPKId & !PKId
                    .MoveNext
                Wend
            End If
        End With
    End If
    BuscaPKIdReceita = strPKId
End Function

Private Function BuscaInscricaoCadastral() As String
    BuscaInscricaoCadastral = Trim(Left(dbc_strInscricaoCadastral.Text, InStr(1, dbc_strInscricaoCadastral.Text, "-", vbTextCompare) - 1))
End Function

Private Sub ReParcelaDividaAtiva()

'******************************************************************************************
' Data: 02/05/2003
' Alteração: - Substituição do comando nativo GETDATE() do SQL Server pela variável
'            strGETDATE.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 02/05/2003
' Alteração: - Substituição do comando nativo ISNULL() do SQL Server pela função gstrISNULL
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 08/05/2003
' Alteração: - Incluídos comandos de inicialização e finalização de bloco PL/SQL permitindo
'            , assim, a execução de múltiplos comandos SQL de uma única vez.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    
    Dim strSQL                                        As String
    Dim strSqlSequencia                       As String
    Dim strSqlDividaAtiva                      As String
    Dim strMsg                                      As String
    Dim adoRecParcelaReceita           As ADODB.Recordset
    Dim adoRecSequencia                   As ADODB.Recordset
    Dim adoRecDetalheDividaAtiva      As ADODB.Recordset
    Dim intIndiceParcelaReceita           As Integer
    Dim intIndice                                    As Integer
    Dim strExercicios                            As String
    Dim dblValorTotal                            As Double
    Dim dblValorDesconto                    As Double
    Dim dblValorDescontoParcela       As Double
    Dim dblValorParcelaReceita           As Double
    Dim lngSequencia                           As Long
    Dim datDataVencimentoReceita    As Date
    Dim dblValorResto                          As Double
    Dim dblValorRestoDesconto          As Double
    Dim intParcelas                              As Integer
    Dim intContaPagina                        As Integer
    Dim intDividaAtiva                           As Integer
    Dim intNumeroPagina                     As Integer
    Dim intNumeroInscricao                 As Integer
    Dim datDataInscricao                     As Date
    Dim intNumeroLivroInscricao         As Integer
    
    If Val(txt_intNumeroParcelas.Text) = 0 Then
        intParcelas = 1
    Else
        intParcelas = Val(txt_intNumeroParcelas.Text)
    End If
    
    strMsg = "Confirma o parcelamento da dívida ativa do contribuinte " & Trim(Mid(dbc_strInscricaoCadastral.Text, InStr(1, dbc_strInscricaoCadastral.Text, "-", vbTextCompare) + 1)) & Chr(10)
    If intParcelas > 1 Then
        strMsg = strMsg & "em " & intParcelas & " parcela(s) ?"
    Else
        strMsg = strMsg & "em uma  parcela ?"
    End If
    
    If gblnExclusaoGravacaoOk("", strMsg, True) Then
        
        Screen.MousePointer = 11
        strExercicios = ""
        For intIndice = 1 To xExercicio.Count(1)
            If xExercicio(intIndice, 1) = -1 Then
                If strExercicios <> "" Then
                    strExercicios = strExercicios & ", "
                End If
                strExercicios = strExercicios & xExercicio(intIndice, 0)
            End If
        Next
            
        strSQL = "SELECT SUM(PAR.dblValorParcela) AS dblValorTotal, PAR.intComposicaoDaReceita, LAN.intExercicio, LAN.dtmLancamento, LAN.bitUtilizacaoDebito, LAN.bytOrigem "
        strSQL = strSQL & "FROM " & gstrParcelaReceita & " PAR, " & gstrLancamentoCalculo & " LAN "
'        strSql = strSql & "WHERE PAR.bytAtiva = 1 AND PAR.dtmDataVencimento < GETDATE() AND LAN.intContribuinte = " & dbc_strInscricaoCadastral.BoundText
        strSQL = strSQL & "WHERE PAR.bytAtiva = 1 AND PAR.dtmDataVencimento < " & strGETDATE & " AND LAN.intContribuinte = " & dbc_strInscricaoCadastral.BoundText
        strSQL = strSQL & " AND LAN.strInscricaoCadastral = '" & BuscaInscricaoCadastral & "' AND LAN.intExercicio IN (" & strExercicios & ") "
        strSQL = strSQL & "AND LAN.PKId = PAR.intLancamentoCalculo "
        strSQL = strSQL & "GROUP BY PAR.intComposicaoDaReceita, LAN.intExercicio, LAN.dtmLancamento, LAN.bitUtilizacaoDebito, LAN.bytOrigem"
        
        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strSQL, 10, adoRecParcelaReceita) Then
            With adoRecParcelaReceita
                If Not (.BOF And .EOF) Then
                    
                    .MoveFirst
                    
                    strSQL = ""
                    
                    If (bytDBType = EDatabases.Oracle) Then
                        strSQL = "DECLARE "
                        strSQL = strSQL & "TYPE tp_csr IS REF CURSOR; "
                        strSQL = strSQL & "csr tp_csr; "
                        strSQL = strSQL & "numDESCONTO NUMBER := 0; "
                        strSQL = strSQL & "BEGIN "
                    End If
                    
                    'DELEÇÃO NA TABELA Parcela Taxa
                    strSQL = strSQL & "DELETE FROM " & gstrParcelaTaxa & " WHERE intLancamentoCalculo IN (SELECT DISTINCT PAR.intLancamentoCalculo "
                    strSQL = strSQL & "FROM " & gstrParcelaReceita & " PAR, " & gstrLancamentoCalculo & " LAN "
'                    strSql = strSql & "WHERE PAR.bytAtiva = 1 AND PAR.dtmDataVencimento < GETDATE() AND LAN.intContribuinte = " & dbc_strInscricaoCadastral.BoundText
                    strSQL = strSQL & "WHERE PAR.bytAtiva = 1 AND PAR.dtmDataVencimento < " & strGETDATE & " AND LAN.intContribuinte = " & dbc_strInscricaoCadastral.BoundText
                    strSQL = strSQL & " AND LAN.strInscricaoCadastral = '" & BuscaInscricaoCadastral & "'"
                    strSQL = strSQL & " AND LAN.intExercicio IN (" & strExercicios & ") "
                    strSQL = strSQL & "AND LAN.PKId = PAR.intLancamentoCalculo)"
                    
                    strSQL = strSQL & IIf((bytDBType = EDatabases.Oracle), "; ", " ")
                    
                    'DELEÇÃO NA TABELA Parcela Receita
                    strSQL = strSQL & "DELETE FROM " & gstrParcelaReceita & " WHERE PKId IN (SELECT PAR.PKId "
                    strSQL = strSQL & "FROM " & gstrParcelaReceita & " PAR, " & gstrLancamentoCalculo & " LAN "
'                    strSql = strSql & "WHERE PAR.bytAtiva = 1 AND PAR.dtmDataVencimento < GETDATE() AND LAN.intContribuinte = " & dbc_strInscricaoCadastral.BoundText
                    strSQL = strSQL & "WHERE PAR.bytAtiva = 1 AND PAR.dtmDataVencimento < " & strGETDATE & " AND LAN.intContribuinte = " & dbc_strInscricaoCadastral.BoundText
                    strSQL = strSQL & " AND LAN.strInscricaoCadastral = '" & BuscaInscricaoCadastral & "'"
                    strSQL = strSQL & " AND LAN.intExercicio IN (" & strExercicios & ") "
                    strSQL = strSQL & "AND LAN.PKId = PAR.intLancamentoCalculo)"
                    
                    strSQL = strSQL & IIf((bytDBType = EDatabases.Oracle), "; ", " ")
                    
                    'DELEÇÃO NA TABELA Detalhe Dívida Ativa
                    strSQL = strSQL & "DELETE FROM " & gstrDetalheDividaAtiva & " WHERE strInscricaoCadastral = '" & BuscaInscricaoCadastral & "' AND intDividaAtiva IN (SELECT PKId FROM " & gstrDividaAtiva & " WHERE intContribuinte = " & dbc_strInscricaoCadastral.BoundText & ")"
                    
                    strSQL = strSQL & IIf((bytDBType = EDatabases.Oracle), "; ", " ")
                    
                    While Not .EOF
                                           
                        'Pesquisa a sequência da composição da receita
                        strSqlSequencia = ""
'                        strSqlSequencia = strSqlSequencia & " SELECT ISNULL(MAX(strSequencia),0) + 1 AS Maximo FROM " & gstrLancamentoCalculo
                        strSqlSequencia = strSqlSequencia & " SELECT " & gstrISNULL("MAX(strSequencia)", "0") & " + 1 AS Maximo FROM " & gstrLancamentoCalculo
                        strSqlSequencia = strSqlSequencia & " WHERE intComposicaoReceita = " & (!intComposicaoDaReceita)
                        strSqlSequencia = strSqlSequencia & " AND intContribuinte = " & dbc_strInscricaoCadastral.BoundText
                        strSqlSequencia = strSqlSequencia & " AND intExercicio = " & gintExercicio
                
                        Set gobjBanco = New clsBanco
                        If gobjBanco.CriaADO(strSqlSequencia, 10, adoRecSequencia) Then
                            lngSequencia = adoRecSequencia!Maximo
                        End If
                        
                        'INSERE LANÇAMENTO CALCULO
                        strSQL = strSQL & " INSERT INTO " & gstrLancamentoCalculo
                        strSQL = strSQL & " (intExercicio, intContribuinte, intComposicaoReceita, intMensagem, strInscricaoCadastral, "
                        strSQL = strSQL & " dtmLancamento, dtmVencimento, intNumeroDeParcelas, intIntervaloEntreParcelas, "
                        strSQL = strSQL & " bitUtilizacaoDebito, intOcorrencia, bytOrigem, strSequencia, dtmDtAtualizacao, lngCodUsr ) VALUES ( "
                        strSQL = strSQL & gintExercicio
                        strSQL = strSQL & ", " & Val(dbc_strInscricaoCadastral.BoundText)
                        strSQL = strSQL & ", " & (!intComposicaoDaReceita)
                        strSQL = strSQL & ", NULL" 'Mensagem - pode conter null
                        strSQL = strSQL & ", '" & BuscaInscricaoCadastral  'Inscrição cadastral (Para receitas diversas - código do contribuinte)
                        strSQL = strSQL & "', " & gstrConvDtParaSql((!dtmLancamento))
                        strSQL = strSQL & ", " & gstrConvDtParaSql(txt_dtmDataVencimento.Text)
                        strSQL = strSQL & ", " & intParcelas
                        strSQL = strSQL & ", " & Val(txt_intIntervalo.Text)
                        strSQL = strSQL & ", " & (!bitUtilizacaoDebito)
                        strSQL = strSQL & ", " & Val(dbc_intOcorrencia.BoundText)   'Ocorrência
                        strSQL = strSQL & ", " & (!bytOrigem)
                        strSQL = strSQL & "," & CStr(lngSequencia)
'                        strSql = strSql & ", GETDATE()"
                        strSQL = strSQL & ", " & strGETDATE
                        strSQL = strSQL & ", " & glngCodUsr
                        strSQL = strSQL & " )"
                        
                        strSQL = strSQL & IIf((bytDBType = EDatabases.Oracle), "; ", " ")
                        
                          'Gravar as Parcelas Taxas
                        strSQL = strSQL & IIf((bytDBType = EDatabases.Oracle), "pe_EfetuaCalculo.", " EXECUTE ")
                        
                        strSQL = strSQL & "sp_EfetuaCalculo" & IIf((bytDBType = EDatabases.Oracle), "(", " ") & _
                            "'" & BuscaPKIdReceita((!intComposicaoDaReceita)) & "'," & (!intComposicaoDaReceita) & ",21,"
                        strSQL = strSQL & txt_intNumeroParcelas.Text & "," & gstrConvDtParaSql(txt_dtmDataVencimento.Text) & "," & txt_intIntervalo.Text
'                        strSql = strSql & ",0,0," & glngCodUsr
                        strSQL = strSQL & ",0," & IIf((bytDBType = EDatabases.Oracle), " numDesconto", " 0") & "," & glngCodUsr
                        
                        strSQL = strSQL & IIf((bytDBType = EDatabases.Oracle), ", csr); ", " ")
                        
                        'Fim Gravar
              
                        dblValorDesconto = 0
                        If Val(txt_dblDesconto.Text) <> 0 Then
                            dblValorDesconto = gstrConvVrDoSql((!dblValorTotal) * (txt_dblDesconto.Text / 100))
                        End If
                        
                        dblValorTotal = !dblValorTotal - dblValorDesconto
                        dblValorParcelaReceita = gstrConvVrDoSql(dblValorTotal / Val(intParcelas))
                        dblValorDescontoParcela = gstrConvVrDoSql(dblValorDesconto / Val(intParcelas))
                        datDataVencimentoReceita = txt_dtmDataVencimento.Text
                                            
                        strSqlDividaAtiva = ""
                        strSqlDividaAtiva = strSqlDividaAtiva & " SELECT DIV.PKId, DET.intNumeroPaginaInscricao, "
                        strSqlDividaAtiva = strSqlDividaAtiva & "DET.intNumeroInscricao, DET.dtmInscricao, "
                        strSqlDividaAtiva = strSqlDividaAtiva & "DET.intNumeroLivroInscricao "
                        strSqlDividaAtiva = strSqlDividaAtiva & "FROM " & gstrDetalheDividaAtiva & " DET, "
                        strSqlDividaAtiva = strSqlDividaAtiva & gstrDividaAtiva & " DIV "
                        strSqlDividaAtiva = strSqlDividaAtiva & " WHERE DIV.intContribuinte = " & dbc_strInscricaoCadastral.BoundText
                        strSqlDividaAtiva = strSqlDividaAtiva & " AND DET.strInscricaoCadastral = '" & BuscaInscricaoCadastral
                        strSqlDividaAtiva = strSqlDividaAtiva & "' AND DET.intComposicaoReceita = " & (!intComposicaoDaReceita)
                        strSqlDividaAtiva = strSqlDividaAtiva & " AND DET.intExercicio = " & (!intExercicio)
                        strSqlDividaAtiva = strSqlDividaAtiva & " AND DET.intDividaAtiva = DIV.PKId "
                        strSqlDividaAtiva = strSqlDividaAtiva & " ORDER BY DET.intNumeroParcela "
                
                        Set gobjBanco = New clsBanco
                        If gobjBanco.CriaADO(strSqlDividaAtiva, 10, adoRecDetalheDividaAtiva) Then
                            With adoRecDetalheDividaAtiva
                                If Not .EOF Then
                                    intDividaAtiva = Val((!PKId))
                                    intNumeroPagina = Val((!intNumeroPaginaInscricao))
                                    intNumeroInscricao = Val((!intNumeroInscricao))
                                    datDataInscricao = (!dtmInscricao)
                                    intNumeroLivroInscricao = Val((!intNumeroLivroInscricao))
                                End If
                            End With
                        End If
                        intContaPagina = 1
                        
                        'Loop para gravar a parcela receita
                        
                        For intIndiceParcelaReceita = 1 To intParcelas
                        
                            If intIndiceParcelaReceita = intParcelas Then
                                dblValorParcelaReceita = (dblValorParcelaReceita * intParcelas) - dblValorResto
                                dblValorDescontoParcela = (dblValorDescontoParcela * intParcelas) - dblValorRestoDesconto
                            Else
                                dblValorResto = gstrConvVrDoSql(dblValorResto + dblValorParcelaReceita)
                                dblValorRestoDesconto = gstrConvVrDoSql(dblValorRestoDesconto + dblValorDescontoParcela)
                            End If
                            
                            strSQL = strSQL & " INSERT INTO " & gstrParcelaReceita
                            strSQL = strSQL & " (intLancamentoCalculo, intComposicaoDaReceita, intNumeroParcela, dtmDataVencimento, "
                            strSQL = strSQL & " dblValorParcela, dblValorDesconto, bytDividaAjuizada, bytSimulado, bytPrescrita, "
                            strSQL = strSQL & " bytCancelada, bytAtiva, bytSuspensaoDeExigencia, dtmDtAtualizacao, lngCodUsr) "
                            strSQL = strSQL & " (SELECT MAX(PKId) "
                            strSQL = strSQL & ", " & (!intComposicaoDaReceita)
                            strSQL = strSQL & ", " & intIndiceParcelaReceita
                            strSQL = strSQL & ", " & gstrConvDtParaSql(datDataVencimentoReceita)
                            strSQL = strSQL & ", " & gstrConvVrParaSql(gstrConvVrDoSql(dblValorParcelaReceita))
                            strSQL = strSQL & ", " & gstrConvVrParaSql(gstrConvVrDoSql(dblValorDescontoParcela))
                            strSQL = strSQL & ", 0" 'Dívida Ajuizada
                            strSQL = strSQL & ", 0" 'Simulado
                            strSQL = strSQL & ", 0" 'Prescrita
                            strSQL = strSQL & ", 0" 'Cancelada
                            strSQL = strSQL & ", 1" 'Divida Ativa
'                            strSql = strSql & ",0, GETDATE()"
                            strSQL = strSQL & ",0, " & strGETDATE
                            strSQL = strSQL & ", " & glngCodUsr
                            strSQL = strSQL & " FROM " & gstrLancamentoCalculo & ")"
                            
                            strSQL = strSQL & IIf((bytDBType = EDatabases.Oracle), "; ", " ")
                            
                            If intContaPagina > 60 Then
                                intNumeroPagina = intNumeroPagina + 1
                                intContaPagina = 1
                            End If
                            
                            strSQL = strSQL & " INSERT INTO " & gstrDetalheDividaAtiva
                            strSQL = strSQL & " (intDividaAtiva, strInscricaoCadastral, intExercicio, dtmVencimento, "
                            strSQL = strSQL & " intNumeroParcela, dtmInscricao, intComposicaoReceita, intOcorrencia, "
                            strSQL = strSQL & " dblValorOriginal, dblValorAtual, bytOrigem, bytDebitoGeradoManualmente, "
                            strSQL = strSQL & " bytSituacao, intNumeroLivroInscricao, intNumeroPaginaInscricao, intNumeroInscricao, "
                            strSQL = strSQL & " dtmDtAtualizacao, lngCodUsr ) "
                            strSQL = strSQL & " VALUES (" & intDividaAtiva
                            strSQL = strSQL & ", '" & BuscaInscricaoCadastral
                            strSQL = strSQL & "', " & gintExercicio
                            strSQL = strSQL & ", " & gstrConvDtParaSql(datDataVencimentoReceita)
                            datDataVencimentoReceita = datDataVencimentoReceita + Val(txt_intIntervalo.Text)
                            strSQL = strSQL & ", " & intIndiceParcelaReceita
                            strSQL = strSQL & ", " & gstrConvDtParaSql(datDataInscricao)
                            strSQL = strSQL & ", " & (!intComposicaoDaReceita)
                            strSQL = strSQL & ", " & dbc_intOcorrencia.BoundText
                            strSQL = strSQL & ", " & gstrConvVrParaSql(gstrConvVrDoSql(dblValorParcelaReceita))
                            strSQL = strSQL & ", " & gstrConvVrParaSql(gstrConvVrDoSql(dblValorParcelaReceita))
                            strSQL = strSQL & ", " & (!bytOrigem)
                            strSQL = strSQL & ", 0 " 'zero é débito gerado pelo sistema - 1 é débito gerado manualmente
                            strSQL = strSQL & ", 2 " 'situacao "Em Aberto"
                            strSQL = strSQL & ", " & intNumeroInscricao
                            strSQL = strSQL & ", " & intNumeroPagina
                            strSQL = strSQL & ", " & intNumeroInscricao
'                            strSql = strSql & ", GETDATE()"
                            strSQL = strSQL & ", " & strGETDATE
                            strSQL = strSQL & ", " & glngCodUsr
                            strSQL = strSQL & ")"
                            
                            strSQL = strSQL & IIf((bytDBType = EDatabases.Oracle), "; ", " ")
                            
                            intContaPagina = intContaPagina + 1
                            intNumeroInscricao = intNumeroInscricao + 1
                        Next
                        .MoveNext
                    Wend
                    
                    strSQL = strSQL & IIf((bytDBType = EDatabases.Oracle), "END; ", "")
                    
                    Set gobjBanco = New clsBanco
                    gobjBanco.ExecutaBeginTrans
                    If gobjBanco.Execute(strSQL, False) Then
                        gobjBanco.ExecutaCommitTrans
                        ExibeMensagem "Parcelamento efetuado com sucesso!"
                    Else
                        gobjBanco.ExecutaRollbackTrans
                    End If
                Else
                    ExibeMensagem "Não há nenhuma dívida ativa para o contribuinte " & Chr(10) & Trim(Mid(dbc_strInscricaoCadastral.Text, InStr(1, dbc_strInscricaoCadastral.Text, "-", vbTextCompare) + 1))
                End If
            End With
            Screen.MousePointer = vbDefault
        End If
    End If
End Sub

Private Sub LimpaControlesDoTAB1()
    Dim intIndice As Integer
    
    optbitTipoDeInscricao(0).Value = True
    dbc_strInscricaoCadastral.BoundText = 0
    dbc_intOcorrencia.BoundText = 0
    tdb_Exercicios.Update
    For intIndice = 1 To xExercicio.Count(1)
        xExercicio(intIndice, 1) = 0
    Next
    txt_intNumeroParcelas.Text = ""
    txt_dtmDataVencimento.Text = ""
    txt_dblDesconto.Text = ""
    txt_intIntervalo.Text = ""
End Sub

Private Sub LimpaControlesDoTAB2()
    dbc_strInscricaoInicial.BoundText = ""
    dbc_strInscricaoFinal.BoundText = ""
    txt_intExercicio.Text = ""
    txt_DataDeVencimento.Text = ""
    dbc_intMensagem1.BoundText = ""
    dbc_intMensagem2.BoundText = ""
    txt_Mensagem1 = ""
    txt_Mensagem2 = ""
    dbc_strInscricaoInicial.SetFocus
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)

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
' Data: 08/05/2003
' Alteração: - Substituição da chamada direta à stored procedure pela função
'            gstrStoredProcedure.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Dim strSQL                As String
    Dim adoResultado As ADODB.Recordset
    Dim intParcelas       As Integer
    Dim strMsg              As String
    Dim strExercicios    As String
    Dim intIndice            As Integer
    Dim i As Integer
    Dim j As Integer
    
On Error Resume Next
    If UCase(strModoOperacao) = UCase(gstrCalcularReajuste) Then
        If blnValidaDados Then
            'ReParcelaDividaAtiva

           'Por storage procedure
            If Val(txt_intNumeroParcelas.Text) = 0 Then
                intParcelas = 1
            Else
                intParcelas = Val(txt_intNumeroParcelas.Text)
            End If

            strMsg = "Confirma o parcelamento da dívida ativa do contribuinte " & Trim(Mid(dbc_strInscricaoCadastral.Text, InStr(1, dbc_strInscricaoCadastral.Text, "-", vbTextCompare) + 1)) & Chr(10)
            If intParcelas > 1 Then
                strMsg = strMsg & "em " & intParcelas & " parcela(s) ?"
            Else
                strMsg = strMsg & "em uma  parcela ?"
            End If

            If gblnExclusaoGravacaoOk("", strMsg, True) Then
                
                Screen.MousePointer = 11
                strExercicios = ""
                For intIndice = 1 To xExercicio.Count(1)
                    If xExercicio(intIndice, 1) = -1 Then
                        If strExercicios <> "" Then
                            strExercicios = strExercicios & ", "
                        End If
                        strExercicios = strExercicios & xExercicio(intIndice, 0)
                    End If
                Next

'                strSQL = "sp_ParcelamentoDividaAtiva " & dbc_strInscricaoCadastral.BoundText & ", '" & BuscaInscricaoCadastral & "', " & dbc_intOcorrencia.BoundText & ", "
'                strSQL = strSQL & intParcelas & ", " & gstrConvDtParaSql(txt_dtmDataVencimento.Text) & ", "
'                strSQL = strSQL & txt_intIntervalo.Text & ", " & Val(txt_dblDesconto.Text) & ", '"
'                strSQL = strSQL & strExercicios & "', " & glngCodUsr
                strSQL = gstrStoredProcedure("sp_ParcelamentoDividaAtiva", dbc_strInscricaoCadastral.BoundText & ", '" & BuscaInscricaoCadastral & "', " & dbc_intOcorrencia.BoundText & ", " & _
                        intParcelas & ", " & gstrConvDtParaSql(txt_dtmDataVencimento.Text) & ", " & _
                        txt_intIntervalo.Text & ", " & Val(txt_dblDesconto.Text) & ", '" & _
                        strExercicios & "', " & glngCodUsr, True)
                
                Set gobjBanco = New clsBanco
                If gobjBanco.CriaADO(strSQL, 10, adoResultado) Then
                    With adoResultado
                        If Not (.BOF And .EOF) Then
                            If (.Fields(0).Value) Then
                                ExibeMensagem "Parcelamento efetuado com sucesso!"
                            Else
                                ExibeMensagem "Não há nenhuma dívida ativa para o contribuinte " & Chr(10) & Trim(Mid(dbc_strInscricaoCadastral.Text, InStr(1, dbc_strInscricaoCadastral.Text, "-", vbTextCompare) + 1))
                            End If
                        End If
                    End With
                End If
                
            End If
            Screen.MousePointer = vbDefault
        End If
    End If
    
    If UCase(strModoOperacao) = UCase(gstrNovo) Then
        If tab_3dPasta.Tab = 0 Then
            LimpaControlesDoTAB1
        Else
            LimpaControlesDoTAB2
        End If
    End If
    
    If UCase(strModoOperacao) = UCase(gstrFechar) Then
        Unload Me
    End If
    
    For i = 0 To 4
        If optbitTipoDeInscricao(i).Value Then
            j = i
            Exit For
        End If
    Next i
    
    If strModoOperacao = gstrPreencherLista Then
        Select Case j
            
            Case 4
                strSQL = ""
'                strSql = "SELECT DISTINCT REC.intContribuinte, CONVERT(NVARCHAR, REC.intContribuinte) + ' - ' + CON.strNome AS Descricao FROM " & gstrReceitaDiversa & " REC, " & gstrContribuinte & " CON WHERE CON.PKId = REC.intContribuinte ORDER BY REC.intContribuinte "
                strSQL = "SELECT DISTINCT REC.intContribuinte, " & gstrCONVERT(CDT_NVARCHAR, "REC.intContribuinte") & strCONCAT & " ' - ' " & strCONCAT & " CON.strNome AS Descricao FROM " & gstrReceitaDiversa & " REC, " & gstrContribuinte & " CON WHERE CON.PKId = REC.intContribuinte ORDER BY REC.intContribuinte "

            Case Else
                strSQL = strQueryInscricao(j)
        End Select
        dbc_strInscricaoCadastral.Tag = strSQL & ";strNome"
        PreencherListaDeOpcoes Me.ActiveControl
    End If
    
End Sub

Private Sub optbitTipoDeInscricao_Click(Index As Integer)
    Dim strSQL As String
    Dim intIndice As Integer

    optbitTipoDeInscricao(Index).CausesValidation = True

    For intIndice = 0 To 4
        If intIndice <> Index Then
            optbitTipoDeInscricao(intIndice).CausesValidation = False
        End If
    Next

    Set dbc_strInscricaoCadastral.RowSource = Nothing
    dbc_strInscricaoCadastral.Text = ""
    If Index = 4 Then
        lbl_strInscricaoCadastral.Caption = "Contribuinte"
    Else
        lbl_strInscricaoCadastral.Caption = "Inscrição Cadastral"
    End If
End Sub

Private Function strQueryInscricao(Index As Integer) As String

'******************************************************************************************
' Data: 02/05/2003
' Alteração: - Substituição do comando nativo CONVERT do SQL Server pela função gstrCONVERT
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 05/05/2003
' Alteração: - Substituição do comando nativo de concatenação de strings (+) do SQL Server
'            pela variável strCONCAT.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Dim strSQL As String
    
    strSQL = ""
    If Index = 0 Or Index = 1 Then
'        strSQL = strSQL & " SELECT B.PKId, LTRIM(RTRIM(A.strInscricaoAnterior)) + ' - ' +  LTRIM(RTRIM(B.strNome)) AS Descricao " 'A.strInscricaoAnterior
        strSQL = strSQL & " SELECT B.PKId, LTRIM(RTRIM(A.strInscricaoAnterior)) " & strCONCAT & " ' - ' " & strCONCAT & "  LTRIM(RTRIM(B.strNome)) AS Descricao " 'A.strInscricaoAnterior
    ElseIf Index = 2 Then
'        strSQL = strSQL & " SELECT B.PKId, LTRIM(RTRIM(A.strInscricaoCadastral)) + ' - ' +  LTRIM(RTRIM(B.strNome)) AS Descricao " 'A.strInscricaoCadastral
        strSQL = strSQL & " SELECT B.PKId, LTRIM(RTRIM(A.strInscricaoCadastral)) " & strCONCAT & " ' - ' " & strCONCAT & "  LTRIM(RTRIM(B.strNome)) AS Descricao " 'A.strInscricaoCadastral
    ElseIf Index = 3 Then
'        strSQL = strSQL & " SELECT C.PKId, LTRIM(RTRIM(A.strInscricaoAnterior)) + ' - ' +  LTRIM(RTRIM(C.strNome)) AS Descricao "
        strSQL = strSQL & " SELECT C.PKId, LTRIM(RTRIM(A.strInscricaoAnterior)) " & strCONCAT & " ' - ' " & strCONCAT & "  LTRIM(RTRIM(C.strNome)) AS Descricao "
    End If
    
    strSQL = strSQL & " FROM "
    
    If Index = 0 Then
        strSQL = strSQL & gstrImobiliario & " A, "
        strSQL = strSQL & gstrContribuinte & " B "
    ElseIf Index = 1 Then
        strSQL = strSQL & gstrImobiliarioRural & " A, "
        strSQL = strSQL & gstrContribuinte & " B "
    ElseIf Index = 2 Then
        strSQL = strSQL & gstrEconomico & " A, "
        strSQL = strSQL & gstrContribuinte & " B "
    ElseIf Index = 3 Then
        strSQL = strSQL & gstrImobiliario & " A, "
        strSQL = strSQL & gstrContribuicaoMelhoria & " B, "
        strSQL = strSQL & gstrContribuinte & " C "
    End If
    strSQL = strSQL & " WHERE "
    If Index = 0 Or Index = 1 Then
        strSQL = strSQL & " A.intContribuinte = B.PKId "
'        strSql = strSql & " ORDER BY convert(numeric,strInscricaoAnterior) "
        strSQL = strSQL & " ORDER BY " & gstrCONVERT(CDT_NUMERIC, "strInscricaoAnterior")
    ElseIf Index = 2 Then
        strSQL = strSQL & " A.intContribuinte = B.PKId "
'        strSql = strSql & " ORDER BY convert(numeric,strInscricaoCadastral) "
        strSQL = strSQL & " ORDER BY " & gstrCONVERT(CDT_NUMERIC, "strInscricaoCadastral")
    ElseIf Index = 3 Then
        strSQL = strSQL & " B.intImobiliario = A.PKId "
        strSQL = strSQL & " AND A.intContribuinte = C.PKId "
'        strSql = strSql & " ORDER BY convert(numeric,strInscricaoAnterior) "
        strSQL = strSQL & " ORDER BY " & gstrCONVERT(CDT_NUMERIC, "strInscricaoAnterior")
    End If
    
strQueryInscricao = strSQL
End Function

Private Function blnDadosGuiaOK() As Boolean

    If dbc_strInscricaoInicial.BoundText = "" Then
        ExibeMensagem "Selecione uma Inscrição Cadastral Inicial para gerar a Guia de Arrecadação."
        dbc_strInscricaoInicial.SetFocus
        Exit Function
    End If

    If dbc_strInscricaoFinal.BoundText = "" Then
        ExibeMensagem "Selecione uma Inscrição Cadastral Final para gerar a Guia de Arrecadação."
        dbc_strInscricaoFinal.SetFocus
        Exit Function
    End If

    If txt_intExercicio.Text = "" Then
        ExibeMensagem "O Exercício deve ser Digitado."
        txt_intExercicio.SetFocus
        Exit Function
    End If

    If txt_dtmDataVencimento.Text = "" Then
        ExibeMensagem "A data de vencimento deve ser digitada."
        txt_dtmDataVencimento.SetFocus
        Exit Function
    ElseIf gblnDataValida(txt_dtmDataVencimento.Text) = False Then
        ExibeMensagem "Data de vencimento inválida."
        txt_dtmDataVencimento.SetFocus
        Exit Function
    End If

    If chk_EmBranco1.Value = 0 Then
        If txt_Mensagem1.Text = "" Then
            ExibeMensagem "A mensagem 1 tem que ser selecionada."
            Exit Function
        End If
    End If

    If chk_EmBranco2.Value = 0 Then
        If txt_Mensagem2.Text = "" Then
            ExibeMensagem "A mensagem 2 tem que ser selecionada."
            Exit Function
        End If
    End If

    blnDadosGuiaOK = True
End Function

Private Sub dbc_intMensagem1_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_intMensagem1
End Sub

Private Sub dbc_intMensagem2_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_intMensagem2
End Sub

Private Sub txt_DataDeVencimento_LostFocus()
    txt_DataDeVencimento.Text = gstrDataFormatada(txt_DataDeVencimento.Text)
End Sub

Private Sub txt_dblDesconto_GotFocus()
    MarcaCampo txt_dblDesconto
End Sub

Private Sub txt_dblDesconto_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_dblDesconto
End Sub

Private Sub txt_dtmDataVencimento_GotFocus()
    MarcaCampo txt_dtmDataVencimento
End Sub

Private Sub txt_dtmDataVencimento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txt_dtmDataVencimento
End Sub

Private Sub txt_dtmDataVencimento_LostFocus()
    txt_dtmDataVencimento.Text = gstrDataFormatada(txt_dtmDataVencimento.Text)
End Sub

Private Sub tab_3dPasta_Click(PreviousTab As Integer)
    Select Case tab_3dPasta.Tab
        Case 0
            HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrImprimir
            HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrCalcularReajuste
        Case 1
            HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrImprimir, gstrNovo
            HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrCalcularReajuste
    End Select
End Sub

Private Sub txt_intIntervalo_GotFocus()
    MarcaCampo txt_intIntervalo
End Sub

Private Sub txt_intIntervalo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_intIntervalo
End Sub

Private Sub txt_intNumeroParcelas_GotFocus()
    MarcaCampo txt_intNumeroParcelas
End Sub

Private Sub txt_intNumeroParcelas_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_intNumeroParcelas
End Sub

Private Sub chk_EmBranco1_Click()
    If chk_EmBranco1.Value = 1 Then
        dbc_intMensagem1.BoundText = ""
        dbc_intMensagem1.Enabled = False
        TrocaCorObjeto dbc_intMensagem1, True
        txt_Mensagem1.Text = ""
        txt_Mensagem1.Enabled = False
        TrocaCorObjeto txt_Mensagem1, True
    Else
        dbc_intMensagem1.Enabled = True
        TrocaCorObjeto dbc_intMensagem1, False
        txt_Mensagem1.Enabled = True
        TrocaCorObjeto txt_Mensagem1, False
    End If
End Sub

Private Sub chk_EmBranco2_Click()
    If chk_EmBranco2.Value = 1 Then
        dbc_intMensagem2.BoundText = ""
        dbc_intMensagem2.Enabled = False
        TrocaCorObjeto dbc_intMensagem2, True
        txt_Mensagem2.Text = ""
        txt_Mensagem2.Enabled = False
        TrocaCorObjeto txt_Mensagem2, True
    Else
        dbc_intMensagem2.Enabled = True
        TrocaCorObjeto dbc_intMensagem2, False
        txt_Mensagem2.Enabled = True
        TrocaCorObjeto txt_Mensagem2, False
    End If
End Sub

Private Sub dbc_intMensagem1_Click(Area As Integer)
    DropDownDataCombo dbc_intMensagem1, Me, Area
    If Area = 2 Then
        LeDoComboParaTXT1
    End If
End Sub

Private Sub dbc_intMensagem2_Click(Area As Integer)
    DropDownDataCombo dbc_intMensagem2, Me, Area
    If Area = 2 Then
        LeDoComboParaTXT2
    End If
End Sub

Private Function LeDoComboParaTXT1()
Dim strSQL As String
Dim adoResultado As ADODB.Recordset

    strSQL = ""
    strSQL = strSQL & " SELECT strMensagem "
    strSQL = strSQL & " FROM " & gstrMensagem
    strSQL = strSQL & " WHERE PKId = " & Val(dbc_intMensagem1.BoundText)
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        If adoResultado.EOF = False Then
            txt_Mensagem1.Text = adoResultado!strMensagem
            adoResultado.MoveNext
        Else
            txt_Mensagem1.Text = ""
        End If
    End If
End Function

Private Function LeDoComboParaTXT2()
Dim strSQL As String
Dim adoResultado As ADODB.Recordset

    strSQL = ""
    strSQL = strSQL & " SELECT strMensagem "
    strSQL = strSQL & " FROM " & gstrMensagem
    strSQL = strSQL & " WHERE PKId = " & Val(dbc_intMensagem2.BoundText)
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        If adoResultado.EOF = False Then
            txt_Mensagem2.Text = adoResultado!strMensagem
            adoResultado.MoveNext
        Else
            txt_Mensagem2.Text = ""
        End If
    End If
End Function

Private Sub txt_DataDeVencimento_GotFocus()
    MarcaCampo txt_DataDeVencimento
End Sub

Private Sub txt_DataDeVencimento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txt_DataDeVencimento
End Sub

Private Sub txt_intExercicio_GotFocus()
    MarcaCampo txt_intExercicio
End Sub

Private Sub txt_intExercicio_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_intExercicio
End Sub

Private Sub dbc_strInscricaoInicial_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_strInscricaoInicial
End Sub

Private Sub dbc_strInscricaoFinal_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_strInscricaoFinal
End Sub

