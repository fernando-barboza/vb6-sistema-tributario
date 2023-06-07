VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Begin VB.Form frmCadCalculoISSQNHomologadoMensal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lançamento de ISSQN Homologado / Mensal"
   ClientHeight    =   5820
   ClientLeft      =   1890
   ClientTop       =   2265
   ClientWidth     =   8655
   Icon            =   "CadCalculoISSQNHomologadoMensal.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   8655
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   5625
      Left            =   180
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   150
      Width           =   8355
      _ExtentX        =   14737
      _ExtentY        =   9922
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "Lançamento de ISSQN Homologado/Mensal"
      TabPicture(0)   =   "CadCalculoISSQNHomologadoMensal.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl_Exercicio"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl_strInscricaoCadastralInicial"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lbl_intComposicaoDaReceita"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lbl_strInscricaoCadastralFinal"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lbl_dtmDataPagamento"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lbl_intOcorrencia"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lbl_ParcelaFinal"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lbl_ParcelaIncial"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "dbc_intOcorrencia"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "tdb_VencimentoParcelas"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "dbc_strInscricaoCadastralInicial"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "dbc_intComposicaoDaReceita"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "dbc_strInscricaoCadastralFinal"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txt_Exercicio"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txt_dtmDataPagamento"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "chk_Selecionar"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txt_intParcelaFinal"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txt_intParcelaInicial"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).ControlCount=   18
      TabCaption(1)   =   "Composição da Receita"
      TabPicture(1)   =   "CadCalculoISSQNHomologadoMensal.frx":105E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "tdb_Atividades"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Emissão de Guias de Arrecadação"
      TabPicture(2)   =   "CadCalculoISSQNHomologadoMensal.frx":107A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fra_EmissaoDeGuias"
      Tab(2).ControlCount=   1
      Begin VB.TextBox txt_intParcelaInicial 
         Height          =   285
         Left            =   3840
         MaxLength       =   15
         TabIndex        =   4
         Top             =   1725
         Width           =   540
      End
      Begin VB.TextBox txt_intParcelaFinal 
         Height          =   285
         Left            =   4830
         MaxLength       =   15
         TabIndex        =   5
         Top             =   1725
         Width           =   540
      End
      Begin VB.CheckBox chk_Selecionar 
         Caption         =   "Selecionar todas as Inscrições"
         Height          =   255
         Left            =   2370
         TabIndex        =   2
         Top             =   1410
         Width           =   2835
      End
      Begin VB.Frame fra_EmissaoDeGuias 
         Height          =   5055
         Left            =   -74880
         TabIndex        =   23
         Top             =   450
         Width           =   8055
         Begin VB.Frame fra_Mensagem1 
            Caption         =   "Mensagem 1                                "
            Height          =   1695
            Left            =   420
            TabIndex        =   26
            Top             =   750
            Width           =   6945
            Begin VB.CheckBox chk_EmBranco1 
               Caption         =   "Em branco"
               Height          =   195
               Left            =   1350
               TabIndex        =   11
               Top             =   0
               Width           =   1095
            End
            Begin VB.TextBox txt_Mensagem1 
               Height          =   885
               Left            =   120
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   13
               TabStop         =   0   'False
               Top             =   660
               Width           =   6675
            End
            Begin MSDataListLib.DataCombo dbc_intMensagem1 
               Height          =   315
               Left            =   1080
               TabIndex        =   12
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
               TabIndex        =   27
               Top             =   390
               Width           =   780
            End
         End
         Begin VB.Frame fra_Mensagem2 
            Caption         =   "Mensagem 2                                "
            Height          =   1695
            Left            =   420
            TabIndex        =   24
            Top             =   2580
            Width           =   6945
            Begin VB.CheckBox chk_EmBranco2 
               Caption         =   "Em branco"
               Height          =   195
               Left            =   1350
               TabIndex        =   14
               Top             =   0
               Width           =   1095
            End
            Begin VB.TextBox txt_Mensagem2 
               Height          =   885
               Left            =   120
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   16
               TabStop         =   0   'False
               Top             =   660
               Width           =   6675
            End
            Begin MSDataListLib.DataCombo dbc_intMensagem2 
               Height          =   315
               Left            =   1080
               TabIndex        =   15
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
               TabIndex        =   25
               Top             =   390
               Width           =   780
            End
         End
      End
      Begin VB.TextBox txt_dtmDataPagamento 
         Height          =   285
         Left            =   7140
         MaxLength       =   15
         TabIndex        =   6
         Top             =   1710
         Width           =   1035
      End
      Begin VB.TextBox txt_Exercicio 
         Height          =   285
         Left            =   2370
         MaxLength       =   4
         TabIndex        =   3
         Top             =   1725
         Width           =   525
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_Atividades 
         Height          =   4695
         Left            =   -74850
         TabIndex        =   10
         Top             =   690
         Width           =   8025
         _ExtentX        =   14155
         _ExtentY        =   8281
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
         Splits(0)._ColumnProps(14)=   "Column(2).Width=12039"
         Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=11959"
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=46,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=50,.parent=13"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
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
         _StyleDefs(60)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(61)  =   "Named:id=39:EvenRow"
         _StyleDefs(62)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(63)  =   "Named:id=40:OddRow"
         _StyleDefs(64)  =   ":id=40,.parent=33"
         _StyleDefs(65)  =   "Named:id=41:RecordSelector"
         _StyleDefs(66)  =   ":id=41,.parent=34"
         _StyleDefs(67)  =   "Named:id=42:FilterBar"
         _StyleDefs(68)  =   ":id=42,.parent=33"
      End
      Begin MSDataListLib.DataCombo dbc_strInscricaoCadastralFinal 
         Height          =   315
         Left            =   2370
         TabIndex        =   1
         Top             =   1020
         Width           =   5805
         _ExtentX        =   10239
         _ExtentY        =   556
         _Version        =   393216
         IntegralHeight  =   0   'False
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbc_intComposicaoDaReceita 
         Height          =   315
         Left            =   2370
         TabIndex        =   7
         Top             =   2070
         Width           =   5805
         _ExtentX        =   10239
         _ExtentY        =   556
         _Version        =   393216
         IntegralHeight  =   0   'False
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbc_strInscricaoCadastralInicial 
         Height          =   315
         Left            =   2370
         TabIndex        =   0
         Top             =   630
         Width           =   5805
         _ExtentX        =   10239
         _ExtentY        =   556
         _Version        =   393216
         IntegralHeight  =   0   'False
         Text            =   ""
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_VencimentoParcelas 
         Height          =   2475
         Left            =   150
         TabIndex        =   9
         Top             =   2970
         Width           =   8025
         _ExtentX        =   14155
         _ExtentY        =   4366
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
      Begin MSDataListLib.DataCombo dbc_intOcorrencia 
         Height          =   315
         Left            =   2370
         TabIndex        =   8
         Top             =   2460
         Width           =   5805
         _ExtentX        =   10239
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin VB.Label lbl_ParcelaIncial 
         AutoSize        =   -1  'True
         Caption         =   "Parcela"
         Height          =   195
         Left            =   3150
         TabIndex        =   30
         Top             =   1770
         Width           =   540
      End
      Begin VB.Label lbl_ParcelaFinal 
         AutoSize        =   -1  'True
         Caption         =   "até"
         Height          =   195
         Left            =   4500
         TabIndex        =   29
         Top             =   1770
         Width           =   225
      End
      Begin VB.Label lbl_intOcorrencia 
         AutoSize        =   -1  'True
         Caption         =   "Ocorrência"
         Height          =   195
         Left            =   1440
         TabIndex        =   28
         Top             =   2550
         Width           =   780
      End
      Begin VB.Label lbl_dtmDataPagamento 
         AutoSize        =   -1  'True
         Caption         =   "Data de Lançamento"
         Height          =   195
         Left            =   5520
         TabIndex        =   22
         Top             =   1785
         Width           =   1500
      End
      Begin VB.Label lbl_strInscricaoCadastralFinal 
         AutoSize        =   -1  'True
         Caption         =   "Inscrição Cadastral Final"
         Height          =   195
         Left            =   495
         TabIndex        =   21
         Top             =   1110
         Width           =   1725
      End
      Begin VB.Label lbl_intComposicaoDaReceita 
         AutoSize        =   -1  'True
         Caption         =   "Composição da Receita"
         Height          =   195
         Left            =   525
         TabIndex        =   20
         Top             =   2175
         Width           =   1695
      End
      Begin VB.Label lbl_strInscricaoCadastralInicial 
         AutoSize        =   -1  'True
         Caption         =   "Inscrição Cadastral Inicial"
         Height          =   195
         Left            =   420
         TabIndex        =   18
         Top             =   720
         Width           =   1800
      End
      Begin VB.Label lbl_Exercicio 
         AutoSize        =   -1  'True
         Caption         =   "Exercício"
         Height          =   195
         Left            =   1545
         TabIndex        =   17
         Top             =   1830
         Width           =   675
      End
   End
End
Attribute VB_Name = "frmCadCalculoISSQNHomologadoMensal"
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

Private Sub chk_Selecionar_Click()
    If chk_Selecionar.Value = 1 Then
        dbc_strInscricaoCadastralInicial.BoundText = ""
        dbc_strInscricaoCadastralFinal.BoundText = ""
        dbc_strInscricaoCadastralInicial.Enabled = False
        TrocaCorObjeto dbc_strInscricaoCadastralInicial, True
        dbc_strInscricaoCadastralFinal.Enabled = False
        TrocaCorObjeto dbc_strInscricaoCadastralFinal, True
        txt_Exercicio.SetFocus
    Else
        dbc_strInscricaoCadastralInicial.Enabled = True
        TrocaCorObjeto dbc_strInscricaoCadastralInicial, False
        dbc_strInscricaoCadastralFinal.Enabled = True
        TrocaCorObjeto dbc_strInscricaoCadastralFinal, False
        dbc_strInscricaoCadastralInicial.SetFocus
    End If

End Sub

Private Sub dbc_intComposicaoDaReceita_Click(Area As Integer)
    DropDownDataCombo dbc_intComposicaoDaReceita, Me, Area
    If Area = 2 And dbc_intComposicaoDaReceita.MatchedWithList Then
        MontaAtividade dbc_intComposicaoDaReceita.BoundText
        If txt_Exercicio <> "" And dbc_intComposicaoDaReceita.BoundText <> "" Then
            LeDaTabelaParaObj gstrVencimentosDasParcelas, tdb_VencimentoParcelas, gQueryTDB_VencimentoParcelasReceita(4, Val(txt_Exercicio.Text))
            tdb_VencimentoParcelas.MoveFirst
            txt_intParcelaInicial.Text = gstrENulo(tdb_VencimentoParcelas.Columns("intNumero").Value)
            tdb_VencimentoParcelas.MoveLast
            txt_intParcelaFinal.Text = gstrENulo(tdb_VencimentoParcelas.Columns("intNumero").Value)
            tdb_VencimentoParcelas.MoveFirst
        End If
    End If
End Sub

Private Sub dbc_intComposicaoDaReceita_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbc_intComposicaoDaReceita, Me, , KeyCode, Shift
End Sub

Private Sub dbc_intComposicaoDaReceita_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_intComposicaoDaReceita
End Sub

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

Private Sub dbc_strInscricaoCadastralFinal_Click(Area As Integer)
   DropDownDataCombo dbc_strInscricaoCadastralFinal, Me, Area
End Sub

Private Sub dbc_strInscricaoCadastralFinal_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbc_strInscricaoCadastralFinal, Me, , KeyCode, Shift
End Sub

Private Sub dbc_strInscricaoCadastralfinal_KeyPress(KeyAscii As Integer)
   CaracterValido KeyAscii, "A", dbc_strInscricaoCadastralFinal
End Sub

Private Sub dbc_strInscricaoCadastralInicial_Click(Area As Integer)
   DropDownDataCombo dbc_strInscricaoCadastralInicial, Me, Area
End Sub

Private Sub dbc_strInscricaoCadastralInicial_KeyDown(KeyCode As Integer, Shift As Integer)
   DropDownDataCombo dbc_strInscricaoCadastralInicial, Me, , KeyCode, Shift
End Sub

Private Sub dbc_strInscricaoCadastralInicial_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_strInscricaoCadastralInicial
End Sub

Private Sub tdb_Atividades_AfterColUpdate(ByVal ColIndex As Integer)
tdb_Atividades.Update
End Sub

Private Sub txt_Exercicio_LostFocus()
    If txt_Exercicio <> "" And dbc_intComposicaoDaReceita.BoundText <> "" Then
        LeDaTabelaParaObj gstrVencimentosDasParcelas, tdb_VencimentoParcelas, gQueryTDB_VencimentoParcelasReceita(4, Val(txt_Exercicio.Text))
    End If
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 667
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
    dbc_strInscricaoCadastralInicial.Tag = strQueryInscricao & ";strInscricaoCadastral"
    dbc_strInscricaoCadastralFinal.Tag = strQueryInscricao & ";strInscricaoCadastral"
    
    
    LeDaTabelaParaObj gstrComposicaoDaReceita, dbc_intComposicaoDaReceita, QueryComposicao
    VerificaObjParaAplicar mobjAux
    LeDaTabelaParaObj gstrOcorrencia, dbc_intOcorrencia, strQuerryOcorrencia
    LeDaTabelaParaObj gstrMensagem, dbc_intMensagem1, strQueryMensagem
    LeDaTabelaParaObj gstrMensagem, dbc_intMensagem2, strQueryMensagem
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    mblnSelecionou = False
    mblnPrimeiraVez = False
End Sub

Private Sub tab_3dPasta_Click(PreviousTab As Integer)
    Select Case tab_3dPasta.Tab
        Case 1
           tdb_VencimentoParcelas.ReOpen
            If tdb_VencimentoParcelas.Columns(1) = "" Then
                ExibeMensagem "Para este exercício não foram encontrados respectivas " & Chr(13) _
                                & " datas de vencimentos, para se calcular o imposto."
                tab_3dPasta.Tab = 0
                txt_Exercicio.SetFocus
                Exit Sub
            End If
    End Select
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
    Dim strSql              As String
    Dim blnExisteLancamento As Boolean
    Dim strInscricaoInicial As String
    Dim strInscricaoFinal   As String
    
    If tab_3dPasta.Tab = 2 Then
        If UCase(strModoOperacao) = UCase(gstrNovo) Then
            LimpaObjetos
        End If
        If UCase(strModoOperacao) = UCase(gstrFechar) Then
            Unload Me
        End If
    End If
    If UCase(strModoOperacao) = UCase(gstrImprimir) Then
        If blnDadosGuiaOK Then
            If chk_Selecionar.Value = 0 Then
                strInscricaoInicial = Trim(Left(dbc_strInscricaoCadastralInicial.Text, InStr(1, dbc_strInscricaoCadastralInicial.Text, "-", vbTextCompare) - 1))
                strInscricaoFinal = Trim(Left(dbc_strInscricaoCadastralFinal.Text, InStr(1, dbc_strInscricaoCadastralFinal.Text, "-", vbTextCompare) - 1))
            End If
            strSql = gstrQueryRelatorioGuiaDeArrecadacao(blnExisteLancamento, strInscricaoInicial, strInscricaoFinal, txt_Exercicio.Text, dbc_intComposicaoDaReceita.BoundText, IIf(chk_Selecionar.Value <> 0, True, False), , Val(txt_intParcelaInicial.Text), Val(txt_intParcelaFinal.Text))
            If blnExisteLancamento Then
                Set gfrmFormularioQueEstaImprimindoGuia = Me
                rptGuiaDeArrecadacaoMunicipal.strImposto = dbc_intComposicaoDaReceita.Text
                ImprimeRelatorio rptGuiaDeArrecadacaoMunicipal, strSql
            End If
        End If
    Else
        If UCase(strModoOperacao) = gstrCalcularReajuste Then
            If blnDadosOk Then
                CalculoLancamento
            End If
        End If
    End If
    If strModoOperacao = gstrPreencherLista Then
        PreencherListaDeOpcoes Me.ActiveControl
    End If
End Sub
'====================

Private Sub CalculoLancamento()
    Dim strSql              As String
    Dim adoResultado        As ADODB.Recordset
    Dim strInscricaoInicial As String
    Dim strInscricaoFinal   As String
    Screen.MousePointer = vbHourglass
    If blnDadosOk And blnContaParcelas Then
        If MsgBox("Deseja Efetuar o Cálculo do ISSQN Homologado Mensal ?", vbYesNo, "Tributário") = vbYes Then
            Screen.MousePointer = vbHourglass
            Set gobjBanco = New clsBanco
            gobjBanco.ExecutaBeginTrans
            If chk_Selecionar.Value = 0 Then
                strInscricaoInicial = Trim(Left(dbc_strInscricaoCadastralInicial.Text, InStr(1, dbc_strInscricaoCadastralInicial.Text, "-", vbTextCompare) - 1))
                strInscricaoFinal = Trim(Left(dbc_strInscricaoCadastralFinal.Text, InStr(1, dbc_strInscricaoCadastralFinal.Text, "-", vbTextCompare) - 1))
            End If
            If Not gBlnVerificaLancamentos(txt_Exercicio, dbc_intComposicaoDaReceita.BoundText, _
                                        dbc_intComposicaoDaReceita.Text, Val(txt_intParcelaFinal.Text) - Val(txt_intParcelaInicial.Text) + 1, _
                                        gstrConvDtParaSql(txt_dtmDataPagamento), chk_Selecionar.Value, _
                                        strInscricaoInicial, strInscricaoFinal) Then
                gobjBanco.ExecutaRollbackTrans
                Screen.MousePointer = vbNormal
                Exit Sub
            End If
            strSql = strLancamentos("-1", "-1")
            'Executa o Lançamento
            If gobjBanco.Execute(strSql, False) Then
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

Private Function blnContaParcelas() As Boolean
    Dim adoNumeroParcelas   As ADODB.Recordset
    Dim strSql              As String
    strSql = gQueryTDB_VencimentoParcelasReceita(4, Val(txt_Exercicio.Text))
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoNumeroParcelas) Then
        If adoNumeroParcelas.RecordCount <= 0 Then
            ExibeMensagem "Data das Parcelas não cadastradas. "
            Set adoNumeroParcelas = Nothing
        Else
            If Val(txt_intParcelaFinal.Text) <= adoNumeroParcelas.RecordCount Then
                blnContaParcelas = True
            Else
                ExibeMensagem "O Intervalo definido de parcelas não foi cadastrado"
            End If
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

    Dim strSql          As String
    strSql = " SELECT intContribuinte, strInscricaoCadastral FROM "
'            gstrEconomico & " AS EC "
    strSql = strSql & gstrEconomico & " EC " & _
            " WHERE DtmDataBaixa IS NULL "
    If chk_Selecionar.Value <> 1 Then
        strSql = strSql & " AND PKId BETWEEN " & dbc_strInscricaoCadastralInicial.BoundText & " AND " & dbc_strInscricaoCadastralFinal.BoundText
    End If
    strContribuintes = strSql
End Function

Private Function strLancamentos(dblValorAparcelar As String, dblValorNaoParcelado As String) As String

'******************************************************************************************
' Data: 08/05/2003
' Alteração: - Substituição da chamada direta à stored procedure pela função
'            gstrStoredProcedure.
' Responsável: Everton Bianchini
'******************************************************************************************

    Dim strSql As String
    MontaDataVencimento
'    strSql = " sp_CalculoLancamentoReceitas 4, " & gstrConvVrParaSql(dblValorAparcelar) & " , " & gstrConvVrParaSql(dblValorNaoParcelado) & ",'"
    strSql = gstrStoredProcedure("sp_CalculoLancamentoReceitas", "4, " & gstrConvVrParaSql(dblValorAparcelar) & " , " & gstrConvVrParaSql(dblValorNaoParcelado) & ",'" & _
            strPKId & "','" & strContribuintes & "'," & Val(txt_Exercicio.Text) & _
            "," & dbc_intComposicaoDaReceita.BoundText & ", NULL, " & _
            gstrConvDtParaSql(txt_dtmDataPagamento.Text) & _
            "," & gstrConvDtParaSql(DataVencimento(1, 1)) & "," & txt_intParcelaInicial.Text & "," & txt_intParcelaFinal.Text & ",0,2," & _
            Val(dbc_intOcorrencia.BoundText) & ",2," & glngCodUsr)
    strLancamentos = strSql
End Function
'Fim da Nova Versão
'====================

Private Function blnDadosOk() As Boolean
blnDadosOk = False
    Dim i As Integer
    If chk_Selecionar.Value <> 1 Then
        If dbc_strInscricaoCadastralInicial.BoundText = "" Then
            dbc_strInscricaoCadastralInicial.SetFocus
            ExibeMensagem "Selecione uma Inscrição Cadastral Inicial para efetuar o cálculo."
            dbc_strInscricaoCadastralInicial.SetFocus
            Exit Function
        End If
        If dbc_strInscricaoCadastralFinal.BoundText = "" Then
            dbc_strInscricaoCadastralFinal.SetFocus
            ExibeMensagem "Selecione uma Inscrição Cadastral Final para efetuar o cálculo."
            Exit Function
        End If
    End If
    If dbc_intOcorrencia.MatchedWithList = False Then
        ExibeMensagem "O campo Ocorrência não pode ser Nulo."
        dbc_intOcorrencia.SetFocus
        Exit Function
    End If
    If txt_Exercicio.Text = "" Then
        ExibeMensagem "O exercício deve ser digitado."
        txt_Exercicio.SetFocus
        Exit Function
    End If
    
    If txt_dtmDataPagamento.Text = "" Then
        ExibeMensagem "A data de lançamento deve ser digitada."
        txt_dtmDataPagamento.SetFocus
        Exit Function
    Else
        If gblnDataValida(txt_dtmDataPagamento.Text) = False Then
            ExibeMensagem "A data de lagamento não é válida."
            txt_dtmDataPagamento.SetFocus
            Exit Function
        End If
    End If
    If Year(txt_dtmDataPagamento.Text) <> Val(txt_Exercicio.Text) Then
        ExibeMensagem "O ano do exercício deve ser igual ao ano da data de lançamento."
        txt_Exercicio.SetFocus
        Exit Function
    End If
    If dbc_intComposicaoDaReceita.BoundText = "" Then
        dbc_intComposicaoDaReceita.SetFocus
        ExibeMensagem "Selecione uma Composição da Receita para efetuar o cálculo."
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
    ExibeMensagem "Selecione uma receita para efetuar o cálculo."
End Function

Private Function strPKId() As String
    Dim strSql As String
    Dim i As Integer
    strSql = ""
    For i = 0 To xarReceita.Count(1) - 1
        If xarReceita(i, 2) = -1 Then
            If strSql <> "" Then
               strSql = strSql & ","
            End If
        strSql = strSql & xarReceita(i, 0)
        End If
    Next

    strPKId = strSql
End Function

Private Function strQuerryOcorrencia() As String
Dim strSql As String
    strSql = ""
    strSql = strSql & " SELECT PKId, strDescricao "
    strSql = strSql & " FROM "
    strSql = strSql & gstrOcorrencia
    strSql = strSql & " WHERE "
    strSql = strSql & " intUtilizacaoDaOcorrencia = 1 "
    strSql = strSql & " ORDER BY strDescricao "
strQuerryOcorrencia = strSql
End Function

Private Function QueryComposicao() As String
    Dim strSql As String
    strSql = ""
    strSql = strSql & "SELECT PKId, strDescricao FROM " & gstrComposicaoDaReceita
    strSql = strSql & " WHERE intUtilizacao = 2 "
    strSql = strSql & " ORDER BY strDescricao "
    QueryComposicao = strSql
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
    
    Dim strSql As String

'    strSQL = " SELECT PKId AS PKIds, (strInscricaoCadastral + ' - ' + "
    strSql = " SELECT PKId AS PKIds, (strInscricaoCadastral " & strCONCAT & " ' - ' " & strCONCAT & _
            "(SELECT strNome FROM tblContribuinte " & _
            "WHERE PKId = intContribuinte)) AS Inscricao " & _
            " FROM " & gstrEconomico & _
            " WHERE  dtmDataBaixa IS NULL "
'            " ORDER BY CONVERT(NUMERIC, strInscricaoCadastral) "
    strSql = strSql & " ORDER BY " & gstrCONVERT(CDT_NUMERIC, "strInscricaoCadastral")

    strQueryInscricao = strSql
End Function

Private Sub MontaAtividade(intComposicaoReceita As Integer)
Dim strSql As String
Dim adoRec As ADODB.Recordset
Dim varAux As String

On Error GoTo Err_Handle

Set xarReceita = New XArrayDB
xarReceita.Clear

xarReceita.ReDim 0, 0, 0, 2

strSql = ""
strSql = strSql & " SELECT A.PKId, A.strDescricao FROM "
strSql = strSql & gstrReceita & " A,"
strSql = strSql & gstrValorCompoRec & " B"
strSql = strSql & " WHERE A.PKId = B.intReceita "
strSql = strSql & " AND B.intComposicaoDaReceita = " & intComposicaoReceita

Set gobjBanco = New clsBanco

If gobjBanco.CriaADO(strSql, 10, adoRec) Then
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
Dim strSql As String
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

Private Sub txt_dtmDataPagamento_GotFocus()
    MarcaCampo txt_dtmDataPagamento
End Sub

Private Sub txt_dtmDataPagamento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txt_dtmDataPagamento
End Sub

Private Sub txt_dtmDataPagamento_LostFocus()
    txt_dtmDataPagamento = gstrDataFormatada(txt_dtmDataPagamento)
End Sub

Private Sub txt_Exercicio_GotFocus()
    MarcaCampo txt_Exercicio
End Sub

Private Sub txt_Exercicio_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_Exercicio
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
    Dim strSql       As String
    Dim adoResultado As ADODB.Recordset
    strSql = " SELECT strMensagem " & _
            " FROM " & gstrMensagem & _
            " WHERE PKId = " & Val(dbc_intMensagem1.BoundText)
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        If adoResultado.EOF = False Then
            txt_Mensagem1.Text = adoResultado!strMensagem
            adoResultado.MoveNext
        Else
            txt_Mensagem1.Text = ""
        End If
    End If
End Function

Private Function LeDoComboParaTXT2()
    Dim strSql       As String
    Dim adoResultado As ADODB.Recordset
    strSql = " SELECT strMensagem " & _
             " FROM " & gstrMensagem & _
             " WHERE PKId = " & Val(dbc_intMensagem2.BoundText)
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
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

Dim strSql As String

    strSql = ""
'    strSQL = strSQL & "SELECT PKId, ltrim(rtrim(PKId)) + ' - ' + ltrim(rtrim(strDescricao)) as Descricao "
    strSql = strSql & "SELECT PKId, ltrim(rtrim(PKId)) " & strCONCAT & " ' - ' " & strCONCAT & " ltrim(rtrim(strDescricao)) as Descricao "
    strSql = strSql & " FROM " & gstrMensagem
    strSql = strSql & " ORDER BY PKId "

strQueryMensagem = strSql
End Function

Private Function blnDadosGuiaOK() As Boolean

    If dbc_intComposicaoDaReceita.BoundText = "" Then
        tab_3dPasta.Tab = 0
        ExibeMensagem "A Composição da Receita deve ser selecionada."
        dbc_intComposicaoDaReceita.SetFocus
        Exit Function
    End If
    If chk_Selecionar.Value = 0 Then
        If dbc_strInscricaoCadastralInicial.Text = "" Then
            ExibeMensagem "Selecione uma Inscrição Cadastral Inicial para gerar a Guia de Arrecadação."
            tab_3dPasta.Tab = 0
            dbc_strInscricaoCadastralInicial.SetFocus
            Exit Function
        End If
    
        If dbc_strInscricaoCadastralFinal.Text = "" Then
            ExibeMensagem "Selecione uma Inscrição Cadastral Final para gerar a Guia de Arrecadação."
            tab_3dPasta.Tab = 0
            dbc_strInscricaoCadastralFinal.SetFocus
            Exit Function
        End If
    End If
    If txt_Exercicio.Text = "" Then
        ExibeMensagem "O Exercício deve ser Digitado."
        tab_3dPasta.Tab = 0
        txt_Exercicio.SetFocus
        Exit Function
    End If
    If (txt_intParcelaInicial.Text = "") Or (txt_intParcelaFinal.Text = "") Then
        ExibeMensagem "O intervalo de parcelas deve ser definido "
        tab_3dPasta.Tab = 0
        If txt_intParcelaInicial.Text = "" Then
            txt_intParcelaInicial.SetFocus
        Else
            txt_intParcelaFinal.SetFocus
        End If
        Exit Function
    End If
    If Val(txt_intParcelaInicial.Text) > Val(txt_intParcelaFinal) Then
        ExibeMensagem "O número da parcela final deve ser maior que o número da parcela inicial"
        tab_3dPasta.Tab = 0
        txt_intParcelaFinal.SetFocus
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

Private Sub LimpaObjetos()
    dbc_intMensagem1.BoundText = ""
    dbc_intMensagem2.BoundText = ""
    txt_Mensagem1 = ""
    txt_Mensagem2 = ""
End Sub

Private Sub dbc_intMensagem1_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_intMensagem1
End Sub

Private Sub dbc_intMensagem2_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_intMensagem2
End Sub

Private Sub tdb_Atividades_AfterUpdate()
    tdb_Atividades.Update
End Sub
