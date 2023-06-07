VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Begin VB.Form frmCadCalculoISSQNFixoAnual 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lançamento de ISSQN Fixo ou Anual"
   ClientHeight    =   6120
   ClientLeft      =   2160
   ClientTop       =   1860
   ClientWidth     =   8475
   Icon            =   "CadCalculoISSQNFixoAnual.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6120
   ScaleWidth      =   8475
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   6015
      Left            =   60
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   30
      Width           =   8355
      _ExtentX        =   14737
      _ExtentY        =   10610
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Lançamento de ISSQN Fixo ou Anual"
      TabPicture(0)   =   "CadCalculoISSQNFixoAnual.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl_strComposicaoReceita"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl_dblDesconto"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lbl_p2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lbl_dtmDataPagamento"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lbl_Exercicio"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lbl_strInscricaoCadastralInicial"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lbl_strInscricaoCadastralFinal"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lbl_intOcorrencia"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lbl_ParcelaFinal"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lbl_ParcelaIncial"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lbl_dblValorInformado"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "dbc_intOcorrencia"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "dbc_strInscricaoCadastralInicial"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "dbc_intComposicaoDaReceita"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "tdb_VencimentoParcelas"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "dbc_strInscricaoCadastralFinal"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txt_dblDesconto"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txt_dtmDataPagamento"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txt_intExercicio"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "chk_Selecionar"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txt_intParcelaFinal"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txt_intParcelaInicial"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txt_dblValorInformado"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).ControlCount=   23
      TabCaption(1)   =   "Composição da Receita"
      TabPicture(1)   =   "CadCalculoISSQNFixoAnual.frx":105E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "tdb_Atividades"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Emissão de Guias de Arrecadação"
      TabPicture(2)   =   "CadCalculoISSQNFixoAnual.frx":107A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fra_EmissaoDeGuias"
      Tab(2).ControlCount=   1
      Begin VB.TextBox txt_dblValorInformado 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4770
         MaxLength       =   12
         TabIndex        =   10
         Top             =   2835
         Width           =   1515
      End
      Begin VB.TextBox txt_intParcelaInicial 
         Height          =   285
         Left            =   2310
         MaxLength       =   15
         TabIndex        =   8
         Top             =   2835
         Width           =   480
      End
      Begin VB.TextBox txt_intParcelaFinal 
         Height          =   285
         Left            =   3330
         MaxLength       =   15
         TabIndex        =   9
         Top             =   2835
         Width           =   480
      End
      Begin VB.Frame fra_EmissaoDeGuias 
         Height          =   5220
         Left            =   -74670
         TabIndex        =   28
         Top             =   480
         Width           =   7665
         Begin VB.Frame fra_Mensagem2 
            Caption         =   "Mensagem 2                                "
            Height          =   1695
            Left            =   420
            TabIndex        =   31
            Top             =   3015
            Width           =   6945
            Begin VB.TextBox txt_Mensagem2 
               Height          =   885
               Left            =   120
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   18
               TabStop         =   0   'False
               Top             =   660
               Width           =   6675
            End
            Begin VB.CheckBox chk_EmBranco2 
               Caption         =   "Em branco"
               Height          =   195
               Left            =   1350
               TabIndex        =   16
               Top             =   0
               Width           =   1095
            End
            Begin MSDataListLib.DataCombo dbc_intMensagem2 
               Height          =   315
               Left            =   1080
               TabIndex        =   17
               Top             =   270
               Width           =   5715
               _ExtentX        =   10081
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
            End
            Begin VB.Label lbl_Label2 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Mensagem"
               Height          =   195
               Left            =   120
               TabIndex        =   32
               Top             =   390
               Width           =   780
            End
         End
         Begin VB.Frame fra_Mensagem1 
            Caption         =   "Mensagem 1                                "
            Height          =   1695
            Left            =   420
            TabIndex        =   29
            Top             =   675
            Width           =   6945
            Begin VB.TextBox txt_Mensagem1 
               Height          =   885
               Left            =   120
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   15
               TabStop         =   0   'False
               Top             =   660
               Width           =   6675
            End
            Begin VB.CheckBox chk_EmBranco1 
               Caption         =   "Em branco"
               Height          =   195
               Left            =   1350
               TabIndex        =   13
               Top             =   0
               Width           =   1095
            End
            Begin MSDataListLib.DataCombo dbc_intMensagem1 
               Height          =   315
               Left            =   1080
               TabIndex        =   14
               Top             =   270
               Width           =   5715
               _ExtentX        =   10081
               _ExtentY        =   556
               _Version        =   393216
               Text            =   ""
            End
            Begin VB.Label lbl_Label 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Mensagem"
               Height          =   195
               Left            =   120
               TabIndex        =   30
               Top             =   390
               Width           =   780
            End
         End
      End
      Begin VB.CheckBox chk_Selecionar 
         Caption         =   "Selecionar todas as Inscrições"
         Height          =   255
         Left            =   2310
         TabIndex        =   2
         Top             =   1350
         Width           =   2835
      End
      Begin VB.TextBox txt_intExercicio 
         Height          =   285
         Left            =   2310
         MaxLength       =   4
         TabIndex        =   3
         Top             =   1695
         Width           =   525
      End
      Begin VB.TextBox txt_dtmDataPagamento 
         Height          =   285
         Left            =   7140
         MaxLength       =   15
         TabIndex        =   5
         Top             =   1695
         Width           =   1035
      End
      Begin VB.TextBox txt_dblDesconto 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4050
         MaxLength       =   15
         TabIndex        =   4
         Top             =   1710
         Width           =   855
      End
      Begin MSDataListLib.DataCombo dbc_strInscricaoCadastralFinal 
         Height          =   315
         Left            =   2310
         TabIndex        =   1
         Top             =   960
         Width           =   5865
         _ExtentX        =   10345
         _ExtentY        =   556
         _Version        =   393216
         IntegralHeight  =   0   'False
         Text            =   ""
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_VencimentoParcelas 
         Height          =   2610
         Left            =   150
         TabIndex        =   11
         Top             =   3195
         Width           =   8025
         _ExtentX        =   14155
         _ExtentY        =   4604
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).DataField=   "PKID"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Nº Parcela"
         Columns(1).DataField=   "intNumero"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Data Vencimento"
         Columns(2).DataField=   "dtmDataDaParcela"
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
         Splits(0)._ColumnProps(1)=   "Column(0).Width=1640"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1561"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=3043"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=2963"
         Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(12)=   "Column(2).Width=3545"
         Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=3466"
         Splits(0)._ColumnProps(15)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
         _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
         _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
         _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
         _StyleDefs(48)  =   "Named:id=33:Normal"
         _StyleDefs(49)  =   ":id=33,.parent=0"
         _StyleDefs(50)  =   "Named:id=34:Heading"
         _StyleDefs(51)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(52)  =   ":id=34,.wraptext=-1"
         _StyleDefs(53)  =   "Named:id=35:Footing"
         _StyleDefs(54)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(55)  =   "Named:id=36:Selected"
         _StyleDefs(56)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(57)  =   "Named:id=37:Caption"
         _StyleDefs(58)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(59)  =   "Named:id=38:HighlightRow"
         _StyleDefs(60)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
         _StyleDefs(61)  =   "Named:id=39:EvenRow"
         _StyleDefs(62)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(63)  =   "Named:id=40:OddRow"
         _StyleDefs(64)  =   ":id=40,.parent=33"
         _StyleDefs(65)  =   "Named:id=41:RecordSelector"
         _StyleDefs(66)  =   ":id=41,.parent=34"
         _StyleDefs(67)  =   "Named:id=42:FilterBar"
         _StyleDefs(68)  =   ":id=42,.parent=33"
      End
      Begin MSDataListLib.DataCombo dbc_intComposicaoDaReceita 
         Height          =   315
         Left            =   2310
         TabIndex        =   6
         Top             =   2070
         Width           =   5865
         _ExtentX        =   10345
         _ExtentY        =   556
         _Version        =   393216
         IntegralHeight  =   0   'False
         Text            =   ""
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_Atividades 
         Height          =   4995
         Left            =   -74850
         TabIndex        =   12
         Top             =   690
         Width           =   8025
         _ExtentX        =   14155
         _ExtentY        =   8811
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
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Descrição"
         Columns(2).DataField=   "strDescricao"
         Columns(2).DropDown=   "tdd_Atividades"
         Columns(2).DropDown.vt=   8
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   3
         Splits(0)._UserFlags=   1
         Splits(0).MarqueeStyle=   5
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).ScrollBars=   2
         Splits(0).DividerColor=   12632256
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
         Splits(0)._ColumnProps(8)=   "Column(1).Width=450"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=370"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1).AllowSizing=0"
         Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(14)=   "Column(2).Width=12065"
         Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=11986"
         Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(19)=   "Column(2).AutoDropDown=1"
         Splits(0)._ColumnProps(20)=   "Column(2).DropDownList=1"
         Splits(0)._ColumnProps(21)=   "Column(2).AutoCompletion=1"
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
         _StyleDefs(40)  =   "Splits(0).Columns(2).Style:id=50,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
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
         _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(57)  =   "Named:id=39:EvenRow"
         _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(59)  =   "Named:id=40:OddRow"
         _StyleDefs(60)  =   ":id=40,.parent=33"
         _StyleDefs(61)  =   "Named:id=41:RecordSelector"
         _StyleDefs(62)  =   ":id=41,.parent=34"
         _StyleDefs(63)  =   "Named:id=42:FilterBar"
         _StyleDefs(64)  =   ":id=42,.parent=33"
      End
      Begin MSDataListLib.DataCombo dbc_strInscricaoCadastralInicial 
         Height          =   315
         Left            =   2310
         TabIndex        =   0
         Top             =   570
         Width           =   5865
         _ExtentX        =   10345
         _ExtentY        =   556
         _Version        =   393216
         IntegralHeight  =   0   'False
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbc_intOcorrencia 
         Height          =   315
         Left            =   2310
         TabIndex        =   7
         Top             =   2460
         Width           =   5865
         _ExtentX        =   10345
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin VB.Label lbl_dblValorInformado 
         AutoSize        =   -1  'True
         Caption         =   "Valor"
         Height          =   195
         Left            =   4110
         TabIndex        =   35
         Top             =   2880
         Width           =   360
      End
      Begin VB.Label lbl_ParcelaIncial 
         AutoSize        =   -1  'True
         Caption         =   "Parcela"
         Height          =   195
         Left            =   1650
         TabIndex        =   34
         Top             =   2880
         Width           =   540
      End
      Begin VB.Label lbl_ParcelaFinal 
         AutoSize        =   -1  'True
         Caption         =   "até"
         Height          =   195
         Left            =   2925
         TabIndex        =   33
         Top             =   2880
         Width           =   225
      End
      Begin VB.Label lbl_intOcorrencia 
         AutoSize        =   -1  'True
         Caption         =   "Ocorrência"
         Height          =   195
         Left            =   1410
         TabIndex        =   27
         Top             =   2550
         Width           =   780
      End
      Begin VB.Label lbl_strInscricaoCadastralFinal 
         AutoSize        =   -1  'True
         Caption         =   "Inscrição Cadastral Final"
         Height          =   195
         Left            =   465
         TabIndex        =   26
         Top             =   1050
         Width           =   1725
      End
      Begin VB.Label lbl_strInscricaoCadastralInicial 
         AutoSize        =   -1  'True
         Caption         =   "Inscrição Cadastral Inicial"
         Height          =   195
         Left            =   390
         TabIndex        =   25
         Top             =   660
         Width           =   1800
      End
      Begin VB.Label lbl_Exercicio 
         AutoSize        =   -1  'True
         Caption         =   "Exercício"
         Height          =   195
         Left            =   1515
         TabIndex        =   24
         Top             =   1770
         Width           =   675
      End
      Begin VB.Label lbl_dtmDataPagamento 
         AutoSize        =   -1  'True
         Caption         =   "Data de Lançamento"
         Height          =   195
         Left            =   5490
         TabIndex        =   23
         Top             =   1785
         Width           =   1500
      End
      Begin VB.Label lbl_p2 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   195
         Left            =   4980
         TabIndex        =   22
         Top             =   1770
         Width           =   120
      End
      Begin VB.Label lbl_dblDesconto 
         AutoSize        =   -1  'True
         Caption         =   "Desconto"
         Height          =   195
         Left            =   3210
         TabIndex        =   21
         Top             =   1785
         Width           =   690
      End
      Begin VB.Label lbl_strComposicaoReceita 
         AutoSize        =   -1  'True
         Caption         =   "Composição da Receita"
         Height          =   195
         Left            =   495
         TabIndex        =   20
         Top             =   2145
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmCadCalculoISSQNFixoAnual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mblnAlterando                   As Boolean
Dim mobjAux                         As Object
Dim mblnSelecionou                  As Boolean
Dim mblnPrimeiraVez                 As Boolean
Dim xarReceita                      As XArrayDB
Dim DataVencimento()
Dim adoRecDataVencimentoParcelaZero As ADODB.Recordset
Dim adoResultado                    As ADODB.Recordset

Private Sub chk_Selecionar_Click()
    If chk_Selecionar.Value = 1 Then
        dbc_strInscricaoCadastralInicial.BoundText = ""
        dbc_strInscricaoCadastralFinal.BoundText = ""
        dbc_strInscricaoCadastralInicial.Enabled = False
        TrocaCorObjeto dbc_strInscricaoCadastralInicial, True
        dbc_strInscricaoCadastralFinal.Enabled = False
        TrocaCorObjeto dbc_strInscricaoCadastralFinal, True
        txt_intExercicio.SetFocus
    Else
        dbc_strInscricaoCadastralInicial.Enabled = True
        TrocaCorObjeto dbc_strInscricaoCadastralInicial, False
        dbc_strInscricaoCadastralFinal.Enabled = True
        TrocaCorObjeto dbc_strInscricaoCadastralFinal, False
        dbc_strInscricaoCadastralInicial.SetFocus
    End If
End Sub

Private Sub dbc_intComposicaoDaReceita_Click(Area As Integer)
    If Area = 2 And dbc_intComposicaoDaReceita.MatchedWithList Then
        MontaAtividade dbc_intComposicaoDaReceita.BoundText
        If txt_intExercicio <> "" And dbc_intComposicaoDaReceita.BoundText <> "" Then
            MontaDataVencimento
            gobjBanco.CriaADO montaDataVencimentoParcelaZero, 5, adoRecDataVencimentoParcelaZero
            If Not (adoRecDataVencimentoParcelaZero.BOF And adoRecDataVencimentoParcelaZero.EOF) Then
                LeDaTabelaParaObj gstrVencimentosDasParcelas, tdb_VencimentoParcelas, gQueryTDB_VencimentoParcelasReceita(3, Val(txt_intExercicio.Text))
                tdb_VencimentoParcelas.MoveFirst
                txt_intParcelaInicial.Text = gstrENulo(tdb_VencimentoParcelas.Columns("intNumero").Value)
                tdb_VencimentoParcelas.MoveLast
                txt_intParcelaFinal.Text = gstrENulo(tdb_VencimentoParcelas.Columns("intNumero").Value)
                tdb_VencimentoParcelas.MoveFirst
            End If
        End If
    End If
End Sub

Private Sub dbc_intMensagem1_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbc_intMensagem1, Me, , KeyCode, Shift
End Sub

Private Sub dbc_intMensagem2_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbc_intMensagem2, Me, , KeyCode, Shift
End Sub

Private Sub dbc_strInscricaoCadastralFinal_Click(Area As Integer)
   DropDownDataCombo dbc_strInscricaoCadastralFinal, Me, Area
End Sub

Private Sub dbc_strInscricaoCadastralFinal_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbc_strInscricaoCadastralFinal, Me, , KeyCode, Shift
End Sub

Private Sub dbc_strInscricaoCadastralInicial_Click(Area As Integer)
   DropDownDataCombo dbc_strInscricaoCadastralInicial, Me, Area
End Sub

Private Sub dbc_strInscricaoCadastralInicial_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbc_strInscricaoCadastralInicial, Me, , KeyCode, Shift
End Sub

Private Sub tdb_Atividades_AfterColUpdate(ByVal ColIndex As Integer)
    tdb_Atividades.Update
End Sub

Private Sub txt_dblValorInformado_LostFocus()
    txt_dblValorInformado = gstrConvVrDoSql(txt_dblValorInformado)
End Sub

Private Sub txt_intExercicio_LostFocus()
    If txt_intExercicio <> "" And dbc_intComposicaoDaReceita.BoundText <> "" Then
        LeDaTabelaParaObj gstrVencimentosDasParcelas, tdb_VencimentoParcelas, gQueryTDB_VencimentoParcelasReceita(3, Val(txt_intExercicio.Text))
    End If
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 665
    VirificaGradeListView Me
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrCalcularReajuste
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrImprimir
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrDeletar
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
    mblnAlterando = False
    'LeDaTabelaParaObj gstrEconomico, dbc_strInscricaoCadastralInicial, strQueryInscricao
    'LeDaTabelaParaObj gstrEconomico, dbc_strInscricaoCadastralFinal, strQueryInscricao
    dbc_strInscricaoCadastralInicial.Tag = strQueryInscricao & ";E.strInscricaoCadastral"
    dbc_strInscricaoCadastralFinal.Tag = strQueryInscricao & ";E.strInscricaoCadastral"
    
    LeDaTabelaParaObj gstrComposicaoDaReceita, dbc_intComposicaoDaReceita, QueryComposicao
    LeDaTabelaParaObj gstrOcorrencia, dbc_intOcorrencia, strQuerryOcorrencia

'''GUIA
    LeDaTabelaParaObj gstrMensagem, dbc_intMensagem1, strQueryMensagem
    LeDaTabelaParaObj gstrMensagem, dbc_intMensagem2, strQueryMensagem
    VerificaObjParaAplicar mobjAux
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    mblnSelecionou = False
    mblnPrimeiraVez = False
End Sub

Private Sub tab_3dPasta_Click(PreviousTab As Integer)
    If tab_3dPasta.Tab = 1 Then
        tdb_VencimentoParcelas.ReOpen
        If tdb_VencimentoParcelas.Columns(1) = "" Then
            tab_3dPasta.Tab = 0
            ExibeMensagem " Para este Exercício e/ou Receita, não foram encontrados " & Chr(13) _
                        & " respectivas  datas  de  vencimentos, para se calcular o " & Chr(13) _
                        & " imposto; ou a data da Parcela  Zero não foi cadastrada. "
            dbc_intComposicaoDaReceita.Text = ""
            txt_intExercicio.SetFocus
            Exit Sub
        End If
    End If
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
    Dim strSQL              As String
    Dim blnExisteLancamento As Boolean
    Dim strInscricaoInicial As String
    Dim strInscricaoFinal   As String
    
    If strModoOperacao = gstrImprimir Then
        If blnDadosGuiaOK Then
            If chk_Selecionar.Value = 0 Then
                strInscricaoInicial = Trim(Left(dbc_strInscricaoCadastralInicial.Text, InStr(1, dbc_strInscricaoCadastralInicial.Text, "-", vbTextCompare) - 1))
                strInscricaoFinal = Trim(Left(dbc_strInscricaoCadastralFinal.Text, InStr(1, dbc_strInscricaoCadastralFinal.Text, "-", vbTextCompare) - 1))
            End If
            strSQL = gstrQueryRelatorioGuiaDeArrecadacao(blnExisteLancamento, strInscricaoInicial, strInscricaoFinal, txt_intExercicio.Text, dbc_intComposicaoDaReceita.BoundText, IIf(chk_Selecionar.Value <> 0, True, False), , Val(txt_intParcelaInicial.Text), Val(txt_intParcelaFinal.Text))
            If blnExisteLancamento Then
                Set gfrmFormularioQueEstaImprimindoGuia = Me
                rptGuiaDeArrecadacaoMunicipal.strImposto = dbc_intComposicaoDaReceita.Text
                ImprimeRelatorio rptGuiaDeArrecadacaoMunicipal, strSQL
            End If
        End If
    End If
    If strModoOperacao = gstrNovo Then
        LimpaControlesDoFormulario
    End If
    If strModoOperacao = gstrFechar Then
        Unload Me
    End If
    
    If strModoOperacao = gstrCalcularReajuste Then
        If blnDadosOk Then
            CalculoLancamento
        End If
    End If
    
    If strModoOperacao = gstrPreencherLista Then
        PreencherListaDeOpcoes Me.ActiveControl
    End If
End Sub
Private Sub CalculoLancamento()
    Dim strSQL              As String
    Dim adoResultado        As ADODB.Recordset
    Dim intNumeroParcelas   As Integer
    Dim dblValorComDesconto As String
    Dim strInscricaoInicial As String
    Dim strInscricaoFinal   As String
    
    Screen.MousePointer = vbHourglass
    If blnDadosOk And blnContaParcelas(intNumeroParcelas) Then
        Screen.MousePointer = vbNormal
        If MsgBox("Deseja Efetuar o Cálculo do ISSQN Fixo Anual ?", vbYesNo, "Tributário") = vbYes Then
            Screen.MousePointer = vbHourglass
            Set gobjBanco = New clsBanco
            gobjBanco.ExecutaBeginTrans
            If chk_Selecionar.Value = 0 Then
                strInscricaoInicial = Trim(Left(dbc_strInscricaoCadastralInicial.Text, InStr(1, dbc_strInscricaoCadastralInicial.Text, "-", vbTextCompare) - 1))
                strInscricaoFinal = Trim(Left(dbc_strInscricaoCadastralFinal.Text, InStr(1, dbc_strInscricaoCadastralFinal.Text, "-", vbTextCompare) - 1))
            End If
            If Not gBlnVerificaLancamentos(txt_intExercicio, dbc_intComposicaoDaReceita.BoundText, _
                                        dbc_intComposicaoDaReceita.Text, Val(txt_intParcelaFinal.Text) - Val(txt_intParcelaInicial.Text) + 1, _
                                        gstrConvDtParaSql(txt_dtmDataPagamento), chk_Selecionar.Value, _
                                        strInscricaoInicial, strInscricaoFinal) Then
                gobjBanco.ExecutaRollbackTrans
                Screen.MousePointer = vbNormal
                Exit Sub
            End If
            strSQL = strLancamentos("-1", "-1")
            'Executa o Lançamento
            If gobjBanco.Execute(strSQL, False) Then
                gobjBanco.ExecutaCommitTrans
                Screen.MousePointer = vbNormal
                ExibeMensagem "Cálculo efetuado com sucesso!"
            Else
                gobjBanco.ExecutaRollbackTrans
            End If
        End If
    End If
    Screen.MousePointer = vbNormal
End Sub

Private Function blnContaParcelas(ByRef intNumeroParcelas As Integer) As Boolean
    Dim adoNumeroParcelas   As ADODB.Recordset
    Dim strSQL              As String
    strSQL = gQueryTDB_VencimentoParcelasReceita(3, Val(txt_intExercicio.Text))
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoNumeroParcelas) Then
        If adoNumeroParcelas.RecordCount <= 0 Then
            ExibeMensagem "Data das Parcelas não cadastradas. "
            blnContaParcelas = False
            Set adoNumeroParcelas = Nothing
        Else
            intNumeroParcelas = adoNumeroParcelas.RecordCount
            blnContaParcelas = True
            Set adoNumeroParcelas = Nothing
        End If
    End If
End Function

Private Function strContribuintes() As String

'******************************************************************************************
' Data: 05/05/2003
' Alteração: - Retirado comando "AS" utilizado para dar apelidos às tabelas, o qual não é
'            permitido pelo Oracle, da cláusula FROM.
' Responsável: Everton Bianchini
'******************************************************************************************

    Dim strSQL          As String
    strSQL = " SELECT EC.intContribuinte, EC.strInscricaoCadastral "
'            " FROM tblEconomico AS EC, "
    strSQL = strSQL & " FROM tblEconomico EC, "
'            " tblAtividadeDaEmpresa AS AE, "
    strSQL = strSQL & " tblAtividadeDaEmpresa AE, "
'            " tblAtividadeEC AS AEC "
    strSQL = strSQL & " tblAtividadeEC AEC " & _
            " WHERE EC.PKID = AE.intEconomico " & _
            " AND EC.intAtividadeBasica = 5" & _
            " AND AE.intEconomico = EC.PKId " & _
            " AND AEC.PKId = AE.intAtividade " & _
            " AND AE.blnPrincipal = 1 "
    If chk_Selecionar.Value <> 1 Then
        strSQL = strSQL & " AND EC.PKId Between '" & dbc_strInscricaoCadastralInicial.BoundText & "' AND '" & _
                        dbc_strInscricaoCadastralFinal.BoundText & "'"
    End If
    strSQL = strSQL & " AND dtmDataBaixa IS NULL " & _
            " AND EC.intComposicao = 3 "
    strContribuintes = strSQL
End Function

Private Function strLancamentos(dblValorAparcelar As String, dblValorNaoParcelado As String) As String

    Dim strSQL As String
    Dim strContribuinte As String
    strContribuinte = strContribuintes
    strContribuinte = Replace(strContribuinte, "'", Chr(34))
    MontaDataVencimento
    strSQL = gstrStoredProcedure("sp_CalculoLancamentoReceitas", "3, " & gstrConvVrParaSql(dblValorAparcelar) & " , " & gstrConvVrParaSql(dblValorNaoParcelado) & ",'" & _
            strPKId & "','" & strContribuinte & "'," & Val(txt_intExercicio.Text) & _
            "," & dbc_intComposicaoDaReceita.BoundText & ", NULL, " & _
            gstrConvDtParaSql(txt_dtmDataPagamento.Text) & _
            "," & gstrConvDtParaSql(DataVencimento(1, 1)) & "," & txt_intParcelaInicial.Text & "," & txt_intParcelaFinal.Text & ",0,2," & _
            Val(dbc_intOcorrencia.BoundText) & ",2," & glngCodUsr & _
            ",0," & Abs(Val(txt_dblDesconto.Text)) & ",-1,'" & gstrConvVrParaSql(txt_dblValorInformado.Text) & "'")
    strLancamentos = strSQL
End Function
'====================

Private Function blnDadosOk() As Boolean
    Dim i As Integer
    Dim adoContribuinte As ADODB.Recordset
    Dim strSQL As String
    strSQL = strContribuintes
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoContribuinte) Then
        If adoContribuinte.RecordCount <= 0 Then
            ExibeMensagem "Não foram encontrados contribuintes para calcular."
            Set adoContribuinte = Nothing
            Exit Function
        End If
    End If
    Set adoContribuinte = Nothing
    If chk_Selecionar.Value <> 1 Then
        If dbc_strInscricaoCadastralInicial.BoundText = "" Then
            tab_3dPasta.Tab = 0
            dbc_strInscricaoCadastralInicial.SetFocus
            ExibeMensagem "Selecione uma Inscrição Cadastral Inicial para efetuar o cálculo."
            blnDadosOk = False
            Exit Function
        End If
        If dbc_intOcorrencia.MatchedWithList = False Then
            ExibeMensagem "O campo Ocorrência não pode ser Nulo."
            tab_3dPasta.Tab = 0
            dbc_intOcorrencia.SetFocus
            Exit Function
        End If
        If dbc_strInscricaoCadastralFinal.BoundText = "" Then
            tab_3dPasta.Tab = 0
            dbc_strInscricaoCadastralFinal.SetFocus
            ExibeMensagem "Selecione uma Inscrição Cadastral Final para efetuar o cálculo."
            blnDadosOk = False
            Exit Function
        End If
    End If
    If txt_intExercicio.Text = "" Then
        tab_3dPasta.Tab = 0
        txt_intExercicio.SetFocus
        ExibeMensagem "O Exercício deve ser Digitado."
        blnDadosOk = False
        Exit Function
    End If
    
    If txt_dtmDataPagamento.Text = "" Then
        tab_3dPasta.Tab = 0
        txt_dtmDataPagamento.SetFocus
        ExibeMensagem "A data de pagamento deve ser digitada."
        blnDadosOk = False
        Exit Function
    Else
        If gblnDataValida(txt_dtmDataPagamento.Text) = False Then
            ExibeMensagem "A data de pagamento não é válida."
            tab_3dPasta.Tab = 0
            txt_dtmDataPagamento.SetFocus
            Exit Function
        End If
    End If
    If Year(txt_dtmDataPagamento.Text) <> Val(txt_intExercicio.Text) Then
        ExibeMensagem "O ano do exercício deve ser igual ao ano da data de lançamento."
        tab_3dPasta.Tab = 0
        txt_intExercicio.SetFocus
        Exit Function
    End If
    If dbc_intComposicaoDaReceita.BoundText = "" Then
        tab_3dPasta.Tab = 0
        dbc_intComposicaoDaReceita.SetFocus
        ExibeMensagem "Selecione uma Composição da Receita para efetuar o cálculo."
        blnDadosOk = False
        Exit Function
    End If
    
    If txt_dblDesconto.Text = "" Then
        txt_dblDesconto.Text = "0"
    End If
    If txt_dblValorInformado.Text = "" Then
        ExibeMensagem "O campo Valor deve ser informado "
        tab_3dPasta.Tab = 0
        txt_dblValorInformado.SetFocus
        blnDadosOk = False
        Exit Function
    End If
    If (txt_intParcelaInicial.Text = "") Or (txt_intParcelaFinal.Text = "") Then
        ExibeMensagem "O intervalo de parcelas deve ser definido "
        If txt_intParcelaInicial.Text = "" Then
            txt_intParcelaInicial.SetFocus
        Else
            txt_intParcelaFinal.SetFocus
        End If
        Exit Function
    End If
    If Val(txt_intParcelaInicial.Text) > Val(txt_intParcelaFinal) Then
        ExibeMensagem "O número da parcela final deve ser maior que o número da parcela inicial"
        txt_intParcelaFinal.SetFocus
        Exit Function
    End If
    For i = 0 To xarReceita.Count(1) - 1
        If xarReceita(i, 2) = -1 Then
            blnDadosOk = True
            Exit Function
        End If
    Next
    ExibeMensagem "Selecione uma receita para efetuar o cálculo!"
    tab_3dPasta.Tab = 1
End Function

Private Function strPKId() As String
    Dim strSQL As String
    Dim i As Integer
    strSQL = ""
    For i = 0 To xarReceita.Count(1) - 1
        If xarReceita(i, 2) = -1 Then
            If strSQL <> "" Then
               strSQL = strSQL & ","
            End If
        strSQL = strSQL & xarReceita(i, 0)
        End If
    Next
    strPKId = strSQL
End Function

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

Private Function QueryComposicao() As String
    Dim strSQL As String
    strSQL = ""
    strSQL = strSQL & "SELECT PKId, strDescricao FROM " & gstrComposicaoDaReceita
    strSQL = strSQL & " WHERE intUtilizacao = 2 "
    strSQL = strSQL & " ORDER BY strDescricao "
    QueryComposicao = strSQL
End Function

Private Function strQueryInscricao() As String

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
    
'    strSQL = " SELECT E.PKId AS PKId, (E.strInscricaoCadastral + ' - ' + C.strNome ) AS Inscricao "
    strSQL = " SELECT E.PKId AS PKId, (E.strInscricaoCadastral " & strCONCAT & " ' - ' " & strCONCAT & " C.strNome ) AS Inscricao " & _
            " FROM " & gstrContribuinte & " C, " & _
            gstrEconomico & " E " & _
            " WHERE C.PKId = E.intContribuinte " & _
            " AND E.intAtividadeBasica = 5 " & _
            " AND E.dtmDataBaixa IS NULL " & _
            " AND E.intComposicao = 3 "
'            " ORDER BY CONVERT(NUMERIC, E.strInscricaoCadastral) "
    strSQL = strSQL & " ORDER BY " & gstrCONVERT(cdt_numeric, "E.strInscricaoCadastral")
    
    strQueryInscricao = strSQL
End Function

Private Function montaDataVencimentoParcelaZero() As String

'******************************************************************************************
' Data: 05/05/2003
' Alteração: - Substituição do comando nativo YEAR() do SQL Server pela função gstrDATEPART
' Responsável: Everton Bianchini
'******************************************************************************************

Dim strSQL As String

strSQL = ""
strSQL = strSQL & "SELECT VP.PKId, VP.intNumero, VP.dtmDataDaParcela "
strSQL = strSQL & " FROM " & gstrVencimentosDasParcelas & " VP,"
strSQL = strSQL & gstrVencimentos & " VC "
strSQL = strSQL & " WHERE VP.intNumero = 0 AND VC.PKId = VP.intVencimento "
strSQL = strSQL & " AND VC.intTributo = 3 "
'strSql = strSql & " AND YEAR(dtmDataDaParcela) = " & txt_intExercicio.Text
strSQL = strSQL & " AND " & gstrDATEPART(strYEAR, "dtmDataDaParcela") & " = " & txt_intExercicio.Text
strSQL = strSQL & " ORDER BY intNumero "
montaDataVencimentoParcelaZero = strSQL

End Function

'Private Sub BuscaAtividadePrincipal()
'Dim strSql As String
'
'strSql = ""
'strSql = strSql & " SELECT distinct byttipodovalor, dblvalor "
'strSql = strSql & " FROM tblAtividadeEC AEC, tblAtividade ATV, tblUtilizacaoDaTabelaDeValor UTV, tblTabelaDeValor TV "
'strSql = strSql & " WHERE AEC.PKId = ATV.intAtividade "
'strSql = strSql & " AND ATV.intUtilizacao = UTV.PKId "
'strSql = strSql & " AND UTV.PKId = TV.intCodigoDaUtilizacao "
'strSql = strSql & " AND AEC.PKId = " & adoRecContribuinte!CodAtividadePrincipal
'strSql = strSql
'
'Set gobjBanco = New clsBanco
'gobjBanco.CriaADO strSql, 5, adoRecAtividadePrincipal
'
'End Sub

'Private Sub Contribuintes()
'Dim strSql          As String
'
'strSql = ""
'strSql = strSql & " SELECT EC.intContribuinte, EC.strInscricaoCadastral, AEC.PKID AS CodAtividadePrincipal, AEC.strDescricao AS AtividadePrincipal "
'strSql = strSql & " FROM tblEconomico AS EC, "
'strSql = strSql & " tblTributoEmpresa AS TE, "
'strSql = strSql & " tblAtividadeDaEmpresa AS AE, "
'strSql = strSql & " tblAtividadeEC AS AEC "
'strSql = strSql & " WHERE EC.PKId = TE.intEconomico "
'strSql = strSql & " AND EC.PKID = AE.intEconomico "
'strSql = strSql & " AND AE.intEconomico = EC.PKId "
'strSql = strSql & " AND AEC.PKId = AE.intAtividade "
'strSql = strSql & " AND AE.blnPrincipal = 1 "
'If chk_Selecionar.Value <> 1 Then
'    strSql = strSql & " AND EC.PKId Between '" & dbc_strInscricaoCadastralInicial.BoundText & "' AND '" & dbc_strInscricaoCadastralFinal.BoundText & "'"
'End If
'strSql = strSql & " AND dtmDataBaixa IS NULL " 'Verifica se existe data de baixa e aborta o cálculo.
'strSql = strSql & " AND TE.intTributo = 3 " 'Verifica em Tributos e faixas se está cadastrado ISSQN Fixo
'
'
'Set gobjBanco = New clsBanco
'gobjBanco.CriaADO strSql, 5, adoRecContribuinte
'
'End Sub

Private Sub MontaAtividade(intComposicaoReceita As Integer)
    Dim strSQL As String
    Dim adoRec As ADODB.Recordset
    Dim varAux As String
    
    On Error GoTo Err_Handle
    
    Set xarReceita = New XArrayDB
    xarReceita.Clear
    
    xarReceita.ReDim 0, 0, 0, 2
    
    strSQL = ""
    strSQL = strSQL & " SELECT A.PKId, A.strDescricao FROM "
    strSQL = strSQL & gstrReceita & " A,"
    strSQL = strSQL & gstrValorCompoRec & " B"
    strSQL = strSQL & " WHERE A.PKId = B.intReceita "
    strSQL = strSQL & " AND B.intComposicaoDaReceita = " & intComposicaoReceita
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSQL, 10, adoRec) Then
        With adoRec
            If Not .EOF Then
                xarReceita.ReDim 0, .RecordCount - 1, 0, 2
                Do While Not .EOF
                    varAux = !Pkid
                    xarReceita(.AbsolutePosition - 1, 0) = varAux
                    
                    varAux = False
                    xarReceita(.AbsolutePosition - 1, 1) = varAux
                
                    varAux = !strDescricao
                    xarReceita(.AbsolutePosition - 1, 2) = varAux
                    
                    .MoveNext
                Loop
            End If
        End With
    End If
    
    Set tdb_Atividades.Array = xarReceita
    tdb_Atividades.Rebind
    tdb_Atividades.Refresh
    
    Exit Sub
Err_Handle:

End Sub

Private Sub MontaDataVencimento()
Dim strSQL As String
Dim adoRec As ADODB.Recordset

On Error GoTo Err_Handle

Set adoRec = tdb_VencimentoParcelas.DataSource
    
    With adoRec
        If Not .EOF Then
            ReDim DataVencimento(1 To adoRec.RecordCount, 1 To 1)
            Do While Not .EOF
                DataVencimento(.AbsolutePosition, 1) = !dtmDataDaParcela
                .MoveNext
            Loop
        End If
    End With

Exit Sub
Err_Handle:

End Sub

'''''######################## caracter valido e marca campo ###########################''''

Private Sub txt_dblValorInformado_GotFocus()
    MarcaCampo txt_dblValorInformado
End Sub

Private Sub txt_dblValorInformado_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txt_dblValorInformado
End Sub

Private Sub txt_dtmDataPagamento_GotFocus()
    MarcaCampo txt_dtmDataPagamento
End Sub

Private Sub txt_dtmDataPagamento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txt_dtmDataPagamento
End Sub

Private Sub txt_dtmDataPagamento_LostFocus()
    txt_dtmDataPagamento.Text = gstrDataFormatada(txt_dtmDataPagamento.Text)
End Sub

Private Sub txt_dblDesconto_GotFocus()
    MarcaCampo txt_dblDesconto
End Sub

Private Sub txt_dblDesconto_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txt_dblDesconto
End Sub


'''>>>>>>>>>>>>>>>>>>GUIA DE ARRECADAÇÃO

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

Private Function blnDadosGuiaOK() As Boolean
    
    If chk_Selecionar.Value = 0 Then
        If dbc_strInscricaoCadastralInicial.BoundText = "" Then
            tab_3dPasta.Tab = 0
            ExibeMensagem "Selecione uma Inscrição Cadastral Inicial para gerar a Guia de Arrecadação."
            dbc_strInscricaoCadastralInicial.SetFocus
            Exit Function
        End If
        
        If dbc_strInscricaoCadastralFinal.BoundText = "" Then
            tab_3dPasta.Tab = 0
            ExibeMensagem "Selecione uma Inscrição Cadastral Final para gerar a  Guia de Arrecadação."
            dbc_strInscricaoCadastralFinal.SetFocus
            Exit Function
        End If
    End If
    If Trim(txt_intExercicio.Text) = "" Then
        tab_3dPasta.Tab = 0
        ExibeMensagem "O Exercício deve ser Digitado."
        txt_intExercicio.SetFocus
        Exit Function
    End If
    If dbc_intComposicaoDaReceita.BoundText = "" Then
        tab_3dPasta.Tab = 0
        ExibeMensagem "A Composição da Receita deve ser selecionada."
        dbc_intComposicaoDaReceita.SetFocus
        Exit Function
    End If
    If Trim(txt_intParcelaInicial.Text) = "" Then
        tab_3dPasta.Tab = 0
        ExibeMensagem "A Parcela Inicial deve ser digitada."
        txt_intParcelaInicial.SetFocus
        Exit Function
    ElseIf Trim(txt_intParcelaFinal.Text) = "" Then
        tab_3dPasta.Tab = 0
        ExibeMensagem "A Parcela Final deve ser digitada."
        txt_intParcelaFinal.SetFocus
        Exit Function
    ElseIf Val(txt_intParcelaInicial.Text) > Val(txt_intParcelaFinal.Text) Then
        tab_3dPasta.Tab = 0
        ExibeMensagem "A Parcela Inicial deve ser anterior a Parcela Final."
        txt_intParcelaInicial.SetFocus
        Exit Function
    Else
        tdb_VencimentoParcelas.MoveFirst
        If Val(txt_intParcelaInicial.Text) < Val(tdb_VencimentoParcelas.Columns("intNumero").Value) Then
            tab_3dPasta.Tab = 0
            ExibeMensagem "A Parcela Incial não ser pode menor que o número da primeira parcela."
            txt_intParcelaInicial.SetFocus
            Exit Function
        Else
            tdb_VencimentoParcelas.MoveLast
            If Val(txt_intParcelaFinal.Text) > Val(tdb_VencimentoParcelas.Columns("intNumero").Value) Then
                tab_3dPasta.Tab = 0
                ExibeMensagem "A Parcela Final não ser pode maior que o número de última parcela."
                tdb_VencimentoParcelas.MoveFirst
                txt_intParcelaFinal.SetFocus
                Exit Function
            End If
        End If
    End If
        
    If chk_EmBranco1.Value = 0 Then
        If Trim(txt_Mensagem1.Text) = "" Then
            tab_3dPasta.Tab = 2
            ExibeMensagem "A mensagem 1 tem que ser selecionada."
            dbc_intMensagem1.SetFocus
            Exit Function
        End If
    End If
    If chk_EmBranco2.Value = 0 Then
        If Trim(txt_Mensagem2.Text) = "" Then
            tab_3dPasta.Tab = 2
            ExibeMensagem "A mensagem 2 tem que ser selecionada."
            dbc_intMensagem2.SetFocus
            Exit Function
        End If
    End If
    blnDadosGuiaOK = True
End Function

Private Sub LimpaControlesDoFormulario()
    dbc_strInscricaoCadastralInicial.BoundText = ""
    dbc_strInscricaoCadastralFinal.BoundText = ""
    chk_Selecionar.Value = 0
    txt_intExercicio.Text = ""
    txt_dblDesconto.Text = ""
    dbc_intComposicaoDaReceita.BoundText = ""
    Set tdb_VencimentoParcelas.DataSource = Nothing
    dbc_intOcorrencia.BoundText = ""
    txt_intParcelaInicial.Text = ""
    txt_intParcelaFinal.Text = ""
    
    Set xarReceita = New XArrayDB
    xarReceita.Clear
    xarReceita.ReDim 0, 0, 0, 2
    Set tdb_Atividades.Array = xarReceita
    tdb_Atividades.Rebind
    tdb_Atividades.Refresh
    
    chk_EmBranco1.Value = 0
    chk_EmBranco2.Value = 0
    dbc_intMensagem1.BoundText = ""
    dbc_intMensagem2.BoundText = ""
    txt_Mensagem1 = ""
    txt_Mensagem2 = ""
    tab_3dPasta.Tab = 0
    dbc_strInscricaoCadastralInicial.SetFocus
End Sub

Private Sub dbc_intMensagem1_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_intMensagem1
End Sub

Private Sub dbc_intMensagem2_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_intMensagem2
End Sub

Private Sub txt_intExercicio_GotFocus()
    MarcaCampo txt_intExercicio
End Sub

Private Sub txt_intExercicio_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_intExercicio
End Sub

Private Sub dbc_strInscricaoCadastralInicial_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_strInscricaoCadastralInicial
End Sub

Private Sub dbc_strInscricaoCadastralfinal_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_strInscricaoCadastralFinal
End Sub

Private Sub txt_intParcelaInicial_GotFocus()
    tab_3dPasta.Tab = 0
    MarcaCampo txt_intParcelaInicial
End Sub

Private Sub txt_intParcelaInicial_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_intParcelaInicial
End Sub

Private Sub txt_intParcelaFinal_GotFocus()
    tab_3dPasta.Tab = 0
    MarcaCampo txt_intParcelaFinal
End Sub

Private Sub txt_intParcelaFinal_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_intParcelaFinal
End Sub

Private Sub txt_Mensagem1_GotFocus()
    tab_3dPasta.Tab = 2
    MarcaCampo txt_Mensagem1
End Sub

Private Sub txt_Mensagem1_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_Mensagem1
End Sub

Private Sub txt_Mensagem2_GotFocus()
    tab_3dPasta.Tab = 2
    MarcaCampo txt_Mensagem2
End Sub

Private Sub txt_Mensagem2_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_Mensagem2
End Sub

Private Sub tdb_Atividades_AfterUpdate()
    tdb_Atividades.Update
End Sub
