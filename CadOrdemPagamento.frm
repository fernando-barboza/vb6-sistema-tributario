VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmCadOrdemPagamento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ordens de Pagamento"
   ClientHeight    =   7560
   ClientLeft      =   840
   ClientTop       =   1605
   ClientWidth     =   9930
   HelpContextID   =   5
   Icon            =   "CadOrdemPagamento.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   9930
   Begin VB.TextBox txt_tmp 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   5280
      MaxLength       =   25
      OLEDragMode     =   1  'Automatic
      TabIndex        =   59
      Top             =   30
      Visible         =   0   'False
      Width           =   1065
   End
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   7485
      Left            =   60
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   30
      Width           =   9825
      _ExtentX        =   17330
      _ExtentY        =   13203
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Ordem"
      TabPicture(0)   =   "CadOrdemPagamento.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblstrProcesso"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblTotalAPagar"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lbl_TotalResto"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lbl_TotalEmpenho"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lbl_Fornecedor"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lbl_DataOrdem"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lbl_DataVencimento"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lbl_Processo"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "dbcintCredor"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtPKId"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cbo_HistoricoLiquidacao"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cmd_HistoricoLiquidacao"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "fra_HistoricoLiquidacao"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "tab_3DPastaEmpenho"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "tdb_Lista"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtTotalAPagar"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtTotalDespesa"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtTotalResto"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtTotalEmpenho"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "chkblnPago"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "fra_bytTipo"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txt_intNContribuinte"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "cmd_Credor"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtdtmData"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txtdtmDataVencimento"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "chkblnCancelado"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txtintProcesso"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "txtintExercicio"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "txtTotalAnulacao"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "txtbitDigitoProcesso"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "txtintExercicioProcesso"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "txtstrCodigoProcesso"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "txt_CodHistorico"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).ControlCount=   35
      Begin VB.TextBox txt_CodHistorico 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3840
         MaxLength       =   8
         OLEDragMode     =   1  'Automatic
         TabIndex        =   25
         Top             =   1830
         Width           =   855
      End
      Begin VB.TextBox txtstrCodigoProcesso 
         CausesValidation=   0   'False
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
         Height          =   285
         HideSelection   =   0   'False
         Left            =   4950
         MaxLength       =   15
         MultiLine       =   -1  'True
         TabIndex        =   15
         Top             =   780
         Width           =   825
      End
      Begin VB.TextBox txtintExercicioProcesso 
         CausesValidation=   0   'False
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   5790
         MaxLength       =   4
         TabIndex        =   16
         Top             =   780
         Width           =   465
      End
      Begin VB.TextBox txtbitDigitoProcesso 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   6270
         MaxLength       =   2
         TabIndex        =   17
         Top             =   780
         Width           =   285
      End
      Begin VB.TextBox txtTotalAnulacao 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000000&
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   285
         Left            =   6720
         MultiLine       =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   70
         Top             =   4890
         Width           =   1215
      End
      Begin VB.TextBox txtintExercicio 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1800
         TabIndex        =   3
         Top             =   390
         Width           =   675
      End
      Begin VB.TextBox txtintProcesso 
         Height          =   285
         Left            =   720
         MaxLength       =   25
         OLEDragMode     =   1  'Automatic
         TabIndex        =   2
         Top             =   390
         Width           =   1065
      End
      Begin VB.CheckBox chkblnCancelado 
         Caption         =   "Cancelado"
         Enabled         =   0   'False
         Height          =   195
         Left            =   3300
         TabIndex        =   5
         Top             =   450
         Width           =   1065
      End
      Begin VB.TextBox txtdtmDataVencimento 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2685
         OLEDragMode     =   1  'Automatic
         TabIndex        =   13
         Top             =   780
         Width           =   1065
      End
      Begin VB.TextBox txtdtmData 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   720
         OLEDragMode     =   1  'Automatic
         TabIndex        =   11
         Top             =   780
         Width           =   1065
      End
      Begin VB.CommandButton cmd_Credor 
         Height          =   300
         Left            =   9390
         Picture         =   "CadOrdemPagamento.frx":105E
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         Tag             =   "15"
         ToolTipText     =   "Clique para cadastar contribuinte"
         Top             =   375
         Width           =   330
      End
      Begin VB.TextBox txt_intNContribuinte 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4935
         TabIndex        =   7
         Top             =   375
         Width           =   705
      End
      Begin VB.Frame fra_bytTipo 
         Caption         =   " Tipo"
         Height          =   1035
         Left            =   120
         TabIndex        =   18
         Top             =   1110
         Width           =   3615
         Begin VB.OptionButton optbytTipo 
            Caption         =   "Anulação de Receita"
            Height          =   195
            Index           =   3
            Left            =   1740
            TabIndex        =   22
            Top             =   630
            Width           =   1800
         End
         Begin VB.OptionButton optbytTipo 
            Caption         =   "Extra-Orçamentaria"
            Height          =   195
            Index           =   2
            Left            =   60
            TabIndex        =   21
            Top             =   645
            Width           =   1650
         End
         Begin VB.OptionButton optbytTipo 
            Caption         =   "Empenho"
            Height          =   195
            Index           =   0
            Left            =   60
            TabIndex        =   19
            Top             =   285
            Width           =   1080
         End
         Begin VB.OptionButton optbytTipo 
            Caption         =   "Resto a Pagar"
            Height          =   195
            Index           =   1
            Left            =   1740
            TabIndex        =   20
            Top             =   285
            Width           =   1320
         End
      End
      Begin VB.CheckBox chkblnPago 
         Caption         =   "Pago"
         Enabled         =   0   'False
         Height          =   195
         Left            =   2520
         TabIndex        =   4
         Top             =   450
         Width           =   855
      End
      Begin VB.TextBox txtTotalEmpenho 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000000&
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   285
         Left            =   810
         MultiLine       =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   50
         Top             =   4890
         Width           =   1215
      End
      Begin VB.TextBox txtTotalResto 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000000&
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   285
         Left            =   2595
         MultiLine       =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   49
         Top             =   4890
         Width           =   1215
      End
      Begin VB.TextBox txtTotalDespesa 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000000&
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   285
         Left            =   4650
         MultiLine       =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   48
         Top             =   4890
         Width           =   1215
      End
      Begin VB.TextBox txtTotalAPagar 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000000&
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   285
         Left            =   8355
         MultiLine       =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   47
         Top             =   4890
         Width           =   1215
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_Lista 
         Height          =   1830
         Left            =   90
         TabIndex        =   37
         Top             =   5265
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   3228
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "PKId"
         Columns(0).DataField=   "PKId"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Número"
         Columns(1).DataField=   "intNumero"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Exercício"
         Columns(2).DataField=   "intExercicio"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Empenho"
         Columns(3).DataField=   "dblTotalEmpenho"
         Columns(3).NumberFormat=   "Standard"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Resto a Pagar"
         Columns(4).DataField=   "dblTotalResto"
         Columns(4).NumberFormat=   "Standard"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Despesa"
         Columns(5).DataField=   "dblTotalDespesaExtra"
         Columns(5).NumberFormat=   "Standard"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Anulação Receita"
         Columns(6).DataField=   "dblTotalAnulacaoReceita"
         Columns(6).NumberFormat=   "FormatText Event"
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "Total"
         Columns(7).DataField=   "dbltotal"
         Columns(7).NumberFormat=   "Standard"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "Pago"
         Columns(8).DataField=   "blnPago"
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "Artigo Caixa"
         Columns(9).DataField=   "ArtigoCaixa"
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(10)._VlistStyle=   0
         Columns(10)._MaxComboItems=   5
         Columns(10).Caption=   "Dt. Pagto"
         Columns(10).DataField=   "dtmPagamento"
         Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(11)._VlistStyle=   0
         Columns(11)._MaxComboItems=   5
         Columns(11).Caption=   "Histórico"
         Columns(11).DataField=   "typHistorico"
         Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(12)._VlistStyle=   0
         Columns(12)._MaxComboItems=   5
         Columns(12).Caption=   "byttipo"
         Columns(12).DataField=   "byttipo"
         Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(13)._VlistStyle=   0
         Columns(13)._MaxComboItems=   5
         Columns(13).Caption=   "intContribuinte"
         Columns(13).DataField=   "intContribuinte"
         Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(14)._VlistStyle=   0
         Columns(14)._MaxComboItems=   5
         Columns(14).Caption=   "dtmData"
         Columns(14).DataField=   "dtmData"
         Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(15)._VlistStyle=   0
         Columns(15)._MaxComboItems=   5
         Columns(15).Caption=   "dtmDataVencimento"
         Columns(15).DataField=   "dtmDataVencimento"
         Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(16)._VlistStyle=   0
         Columns(16)._MaxComboItems=   5
         Columns(16).Caption=   "bytCancelado"
         Columns(16).DataField=   "bytCancelado"
         Columns(16)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(17)._VlistStyle=   0
         Columns(17)._MaxComboItems=   5
         Columns(17).Caption=   "Dt. Canc."
         Columns(17).DataField=   "DtmCancelamento"
         Columns(17)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   18
         Splits(0)._UserFlags=   0
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).ScrollBars=   3
         Splits(0).DividerColor=   13160664
         Splits(0).FilterBar=   -1  'True
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=18"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(7)=   "Column(0).AllowFocus=0"
         Splits(0)._ColumnProps(8)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(9)=   "Column(1).Width=1244"
         Splits(0)._ColumnProps(10)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(11)=   "Column(1)._WidthInPix=1164"
         Splits(0)._ColumnProps(12)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(14)=   "Column(2).Width=1138"
         Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=1058"
         Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(18)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(19)=   "Column(3).Width=2461"
         Splits(0)._ColumnProps(20)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(21)=   "Column(3)._WidthInPix=2381"
         Splits(0)._ColumnProps(22)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(23)=   "Column(3)._ColStyle=514"
         Splits(0)._ColumnProps(24)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(25)=   "Column(4).Width=2461"
         Splits(0)._ColumnProps(26)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(27)=   "Column(4)._WidthInPix=2381"
         Splits(0)._ColumnProps(28)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(29)=   "Column(4)._ColStyle=514"
         Splits(0)._ColumnProps(30)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(31)=   "Column(5).Width=2461"
         Splits(0)._ColumnProps(32)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(33)=   "Column(5)._WidthInPix=2381"
         Splits(0)._ColumnProps(34)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(35)=   "Column(5)._ColStyle=514"
         Splits(0)._ColumnProps(36)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(37)=   "Column(6).Width=2461"
         Splits(0)._ColumnProps(38)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(39)=   "Column(6)._WidthInPix=2381"
         Splits(0)._ColumnProps(40)=   "Column(6)._EditAlways=0"
         Splits(0)._ColumnProps(41)=   "Column(6)._ColStyle=2"
         Splits(0)._ColumnProps(42)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(43)=   "Column(7).Width=2223"
         Splits(0)._ColumnProps(44)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(45)=   "Column(7)._WidthInPix=2143"
         Splits(0)._ColumnProps(46)=   "Column(7)._EditAlways=0"
         Splits(0)._ColumnProps(47)=   "Column(7)._ColStyle=514"
         Splits(0)._ColumnProps(48)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(49)=   "Column(8).Width=873"
         Splits(0)._ColumnProps(50)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(51)=   "Column(8)._WidthInPix=794"
         Splits(0)._ColumnProps(52)=   "Column(8)._EditAlways=0"
         Splits(0)._ColumnProps(53)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(54)=   "Column(9).Width=2725"
         Splits(0)._ColumnProps(55)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(56)=   "Column(9)._WidthInPix=2646"
         Splits(0)._ColumnProps(57)=   "Column(9)._EditAlways=0"
         Splits(0)._ColumnProps(58)=   "Column(9).Order=10"
         Splits(0)._ColumnProps(59)=   "Column(10).Width=1508"
         Splits(0)._ColumnProps(60)=   "Column(10).DividerColor=0"
         Splits(0)._ColumnProps(61)=   "Column(10)._WidthInPix=1429"
         Splits(0)._ColumnProps(62)=   "Column(10)._EditAlways=0"
         Splits(0)._ColumnProps(63)=   "Column(10).Order=11"
         Splits(0)._ColumnProps(64)=   "Column(11).Width=3969"
         Splits(0)._ColumnProps(65)=   "Column(11).DividerColor=0"
         Splits(0)._ColumnProps(66)=   "Column(11)._WidthInPix=3889"
         Splits(0)._ColumnProps(67)=   "Column(11)._EditAlways=0"
         Splits(0)._ColumnProps(68)=   "Column(11).Order=12"
         Splits(0)._ColumnProps(69)=   "Column(12).Width=2725"
         Splits(0)._ColumnProps(70)=   "Column(12).DividerColor=0"
         Splits(0)._ColumnProps(71)=   "Column(12)._WidthInPix=2646"
         Splits(0)._ColumnProps(72)=   "Column(12)._EditAlways=0"
         Splits(0)._ColumnProps(73)=   "Column(12).Visible=0"
         Splits(0)._ColumnProps(74)=   "Column(12).Order=13"
         Splits(0)._ColumnProps(75)=   "Column(13).Width=2725"
         Splits(0)._ColumnProps(76)=   "Column(13).DividerColor=0"
         Splits(0)._ColumnProps(77)=   "Column(13)._WidthInPix=2646"
         Splits(0)._ColumnProps(78)=   "Column(13)._EditAlways=0"
         Splits(0)._ColumnProps(79)=   "Column(13).Visible=0"
         Splits(0)._ColumnProps(80)=   "Column(13).Order=14"
         Splits(0)._ColumnProps(81)=   "Column(14).Width=2725"
         Splits(0)._ColumnProps(82)=   "Column(14).DividerColor=0"
         Splits(0)._ColumnProps(83)=   "Column(14)._WidthInPix=2646"
         Splits(0)._ColumnProps(84)=   "Column(14)._EditAlways=0"
         Splits(0)._ColumnProps(85)=   "Column(14).Visible=0"
         Splits(0)._ColumnProps(86)=   "Column(14).Order=15"
         Splits(0)._ColumnProps(87)=   "Column(15).Width=2725"
         Splits(0)._ColumnProps(88)=   "Column(15).DividerColor=0"
         Splits(0)._ColumnProps(89)=   "Column(15)._WidthInPix=2646"
         Splits(0)._ColumnProps(90)=   "Column(15)._EditAlways=0"
         Splits(0)._ColumnProps(91)=   "Column(15).Visible=0"
         Splits(0)._ColumnProps(92)=   "Column(15).Order=16"
         Splits(0)._ColumnProps(93)=   "Column(16).Width=2725"
         Splits(0)._ColumnProps(94)=   "Column(16).DividerColor=0"
         Splits(0)._ColumnProps(95)=   "Column(16)._WidthInPix=2646"
         Splits(0)._ColumnProps(96)=   "Column(16)._EditAlways=0"
         Splits(0)._ColumnProps(97)=   "Column(16).Visible=0"
         Splits(0)._ColumnProps(98)=   "Column(16).Order=17"
         Splits(0)._ColumnProps(99)=   "Column(17).Width=1773"
         Splits(0)._ColumnProps(100)=   "Column(17).DividerColor=0"
         Splits(0)._ColumnProps(101)=   "Column(17)._WidthInPix=1693"
         Splits(0)._ColumnProps(102)=   "Column(17)._EditAlways=0"
         Splits(0)._ColumnProps(103)=   "Column(17).Order=18"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=9.75,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=System"
         PrintInfos(0).PageFooterFont=   "Size=9.75,Charset=0,Weight=700,Underline=0,Italic=0,Strikethrough=0,Name=System"
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
         RowDividerColor =   13160664
         RowSubDividerColor=   13160664
         DirectionAfterEnter=   1
         MaxRows         =   250000
         ViewColumnCaptionWidth=   0
         ViewColumnWidth =   0
         _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
         _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
         _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
         _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=64,.bold=0,.fontsize=825,.italic=0"
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
         _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38,.bgcolor=&H80000014&"
         _StyleDefs(20)  =   ":id=8,.fgcolor=&H80000012&"
         _StyleDefs(21)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(22)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(23)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(24)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(25)  =   "Splits(0).Style:id=13,.parent=1"
         _StyleDefs(26)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
         _StyleDefs(27)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
         _StyleDefs(28)  =   "Splits(0).FooterStyle:id=15,.parent=3"
         _StyleDefs(29)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
         _StyleDefs(30)  =   "Splits(0).SelectedStyle:id=18,.parent=6,.namedParent=38"
         _StyleDefs(31)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(32)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(33)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(34)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(35)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(36)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(37)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(38)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(39)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(40)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
         _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
         _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
         _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
         _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=90,.parent=13"
         _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=87,.parent=14"
         _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=88,.parent=15"
         _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=89,.parent=17"
         _StyleDefs(49)  =   "Splits(0).Columns(3).Style:id=50,.parent=13,.alignment=1"
         _StyleDefs(50)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14,.alignment=2"
         _StyleDefs(51)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(52)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(53)  =   "Splits(0).Columns(4).Style:id=54,.parent=13,.alignment=1"
         _StyleDefs(54)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14,.alignment=2"
         _StyleDefs(55)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
         _StyleDefs(56)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
         _StyleDefs(57)  =   "Splits(0).Columns(5).Style:id=58,.parent=13,.alignment=1"
         _StyleDefs(58)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14,.alignment=2"
         _StyleDefs(59)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
         _StyleDefs(60)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
         _StyleDefs(61)  =   "Splits(0).Columns(6).Style:id=94,.parent=13,.alignment=1"
         _StyleDefs(62)  =   "Splits(0).Columns(6).HeadingStyle:id=91,.parent=14"
         _StyleDefs(63)  =   "Splits(0).Columns(6).FooterStyle:id=92,.parent=15"
         _StyleDefs(64)  =   "Splits(0).Columns(6).EditorStyle:id=93,.parent=17"
         _StyleDefs(65)  =   "Splits(0).Columns(7).Style:id=62,.parent=13,.alignment=1"
         _StyleDefs(66)  =   "Splits(0).Columns(7).HeadingStyle:id=59,.parent=14,.alignment=2"
         _StyleDefs(67)  =   "Splits(0).Columns(7).FooterStyle:id=60,.parent=15"
         _StyleDefs(68)  =   "Splits(0).Columns(7).EditorStyle:id=61,.parent=17"
         _StyleDefs(69)  =   "Splits(0).Columns(8).Style:id=66,.parent=13"
         _StyleDefs(70)  =   "Splits(0).Columns(8).HeadingStyle:id=63,.parent=14"
         _StyleDefs(71)  =   "Splits(0).Columns(8).FooterStyle:id=64,.parent=15"
         _StyleDefs(72)  =   "Splits(0).Columns(8).EditorStyle:id=65,.parent=17"
         _StyleDefs(73)  =   "Splits(0).Columns(9).Style:id=106,.parent=13"
         _StyleDefs(74)  =   "Splits(0).Columns(9).HeadingStyle:id=103,.parent=14"
         _StyleDefs(75)  =   "Splits(0).Columns(9).FooterStyle:id=104,.parent=15"
         _StyleDefs(76)  =   "Splits(0).Columns(9).EditorStyle:id=105,.parent=17"
         _StyleDefs(77)  =   "Splits(0).Columns(10).Style:id=98,.parent=13"
         _StyleDefs(78)  =   "Splits(0).Columns(10).HeadingStyle:id=95,.parent=14"
         _StyleDefs(79)  =   "Splits(0).Columns(10).FooterStyle:id=96,.parent=15"
         _StyleDefs(80)  =   "Splits(0).Columns(10).EditorStyle:id=97,.parent=17"
         _StyleDefs(81)  =   "Splits(0).Columns(11).Style:id=46,.parent=13"
         _StyleDefs(82)  =   "Splits(0).Columns(11).HeadingStyle:id=43,.parent=14"
         _StyleDefs(83)  =   "Splits(0).Columns(11).FooterStyle:id=44,.parent=15"
         _StyleDefs(84)  =   "Splits(0).Columns(11).EditorStyle:id=45,.parent=17"
         _StyleDefs(85)  =   "Splits(0).Columns(12).Style:id=70,.parent=13"
         _StyleDefs(86)  =   "Splits(0).Columns(12).HeadingStyle:id=67,.parent=14"
         _StyleDefs(87)  =   "Splits(0).Columns(12).FooterStyle:id=68,.parent=15"
         _StyleDefs(88)  =   "Splits(0).Columns(12).EditorStyle:id=69,.parent=17"
         _StyleDefs(89)  =   "Splits(0).Columns(13).Style:id=74,.parent=13"
         _StyleDefs(90)  =   "Splits(0).Columns(13).HeadingStyle:id=71,.parent=14"
         _StyleDefs(91)  =   "Splits(0).Columns(13).FooterStyle:id=72,.parent=15"
         _StyleDefs(92)  =   "Splits(0).Columns(13).EditorStyle:id=73,.parent=17"
         _StyleDefs(93)  =   "Splits(0).Columns(14).Style:id=78,.parent=13"
         _StyleDefs(94)  =   "Splits(0).Columns(14).HeadingStyle:id=75,.parent=14"
         _StyleDefs(95)  =   "Splits(0).Columns(14).FooterStyle:id=76,.parent=15"
         _StyleDefs(96)  =   "Splits(0).Columns(14).EditorStyle:id=77,.parent=17"
         _StyleDefs(97)  =   "Splits(0).Columns(15).Style:id=82,.parent=13"
         _StyleDefs(98)  =   "Splits(0).Columns(15).HeadingStyle:id=79,.parent=14"
         _StyleDefs(99)  =   "Splits(0).Columns(15).FooterStyle:id=80,.parent=15"
         _StyleDefs(100) =   "Splits(0).Columns(15).EditorStyle:id=81,.parent=17"
         _StyleDefs(101) =   "Splits(0).Columns(16).Style:id=86,.parent=13"
         _StyleDefs(102) =   "Splits(0).Columns(16).HeadingStyle:id=83,.parent=14"
         _StyleDefs(103) =   "Splits(0).Columns(16).FooterStyle:id=84,.parent=15"
         _StyleDefs(104) =   "Splits(0).Columns(16).EditorStyle:id=85,.parent=17"
         _StyleDefs(105) =   "Splits(0).Columns(17).Style:id=102,.parent=13"
         _StyleDefs(106) =   "Splits(0).Columns(17).HeadingStyle:id=99,.parent=14"
         _StyleDefs(107) =   "Splits(0).Columns(17).FooterStyle:id=100,.parent=15"
         _StyleDefs(108) =   "Splits(0).Columns(17).EditorStyle:id=101,.parent=17"
         _StyleDefs(109) =   "Named:id=33:Normal"
         _StyleDefs(110) =   ":id=33,.parent=0"
         _StyleDefs(111) =   "Named:id=34:Heading"
         _StyleDefs(112) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(113) =   ":id=34,.wraptext=-1"
         _StyleDefs(114) =   "Named:id=35:Footing"
         _StyleDefs(115) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(116) =   "Named:id=36:Selected"
         _StyleDefs(117) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(118) =   "Named:id=37:Caption"
         _StyleDefs(119) =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(120) =   "Named:id=38:HighlightRow"
         _StyleDefs(121) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(122) =   "Named:id=39:EvenRow"
         _StyleDefs(123) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(124) =   "Named:id=40:OddRow"
         _StyleDefs(125) =   ":id=40,.parent=33"
         _StyleDefs(126) =   "Named:id=41:RecordSelector"
         _StyleDefs(127) =   ":id=41,.parent=34"
         _StyleDefs(128) =   "Named:id=42:FilterBar"
         _StyleDefs(129) =   ":id=42,.parent=33"
      End
      Begin TabDlg.SSTab tab_3DPastaEmpenho 
         Height          =   2565
         Left            =   90
         TabIndex        =   41
         Top             =   2250
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   4524
         _Version        =   393216
         Style           =   1
         Tabs            =   4
         TabsPerRow      =   4
         TabHeight       =   520
         TabCaption(0)   =   "Empenho"
         TabPicture(0)   =   "CadOrdemPagamento.frx":13E8
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lbl_Empenho"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "lbl_Parcela"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "lbl_FonteDeRecursoEmpenho"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "lvw_Empenho"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "dcbEmpenho"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "cmd_Empenho"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "txtFonteDeRecursoEmpenho"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "dcbParcela"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).ControlCount=   8
         TabCaption(1)   =   "Resto a pagar"
         TabPicture(1)   =   "CadOrdemPagamento.frx":1404
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "dcbParcelaResto"
         Tab(1).Control(1)=   "txtFonteDeRecursoResto"
         Tab(1).Control(2)=   "cmd_Resto"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "lvw_Resto"
         Tab(1).Control(4)=   "dcbResto"
         Tab(1).Control(5)=   "lbl_FonteDeRecursoResto"
         Tab(1).Control(6)=   "Label4"
         Tab(1).Control(7)=   "Label3"
         Tab(1).ControlCount=   8
         TabCaption(2)   =   "Extra-orçamentária"
         TabPicture(2)   =   "CadOrdemPagamento.frx":1420
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "cmd_Despesa"
         Tab(2).Control(0).Enabled=   0   'False
         Tab(2).Control(1)=   "dcbDespesa"
         Tab(2).Control(2)=   "lvw_Despesa"
         Tab(2).Control(3)=   "Label6"
         Tab(2).ControlCount=   4
         TabCaption(3)   =   "Anulação de Receita"
         TabPicture(3)   =   "CadOrdemPagamento.frx":143C
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "fra_Processo"
         Tab(3).Control(1)=   "txt_dblValor"
         Tab(3).Control(2)=   "txt_Descricao"
         Tab(3).Control(3)=   "lvw_AnulacaoReceita"
         Tab(3).Control(4)=   "lbl_Valor"
         Tab(3).Control(5)=   "lbl_Descricao"
         Tab(3).ControlCount=   6
         Begin VB.Frame fra_Processo 
            Caption         =   " Processo "
            Height          =   675
            Left            =   -67440
            TabIndex        =   69
            Top             =   270
            Width           =   1905
            Begin VB.TextBox txt_strCodigo 
               CausesValidation=   0   'False
               BeginProperty DataFormat 
                  Type            =   0
                  Format          =   "#.##0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1046
                  SubFormatType   =   0
               EndProperty
               Height          =   285
               HideSelection   =   0   'False
               Left            =   150
               MaxLength       =   15
               MultiLine       =   -1  'True
               TabIndex        =   64
               Top             =   270
               Width           =   825
            End
            Begin VB.TextBox txt_intExercicioProcesso 
               CausesValidation=   0   'False
               BeginProperty DataFormat 
                  Type            =   0
                  Format          =   "#.##0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1046
                  SubFormatType   =   0
               EndProperty
               Height          =   285
               Left            =   990
               MaxLength       =   4
               TabIndex        =   65
               Top             =   270
               Width           =   465
            End
            Begin VB.TextBox txt_bitDigito 
               BeginProperty DataFormat 
                  Type            =   0
                  Format          =   "#.##0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1046
                  SubFormatType   =   0
               EndProperty
               Height          =   285
               Left            =   1470
               MaxLength       =   2
               TabIndex        =   66
               Top             =   270
               Width           =   285
            End
         End
         Begin VB.TextBox txt_dblValor 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   -69030
            MaxLength       =   25
            MultiLine       =   -1  'True
            OLEDragMode     =   1  'Automatic
            TabIndex        =   63
            Top             =   480
            Width           =   1425
         End
         Begin VB.TextBox txt_Descricao 
            Height          =   315
            Left            =   -74010
            MaxLength       =   50
            OLEDragMode     =   1  'Automatic
            TabIndex        =   61
            Top             =   480
            Width           =   4365
         End
         Begin VB.ComboBox dcbParcelaResto 
            Height          =   315
            Left            =   -72000
            TabIndex        =   33
            Top             =   510
            Width           =   1185
         End
         Begin VB.ComboBox dcbParcela 
            Height          =   315
            Left            =   3000
            TabIndex        =   30
            Top             =   510
            Width           =   1005
         End
         Begin VB.TextBox txtFonteDeRecursoResto 
            BackColor       =   &H80000000&
            CausesValidation=   0   'False
            Enabled         =   0   'False
            Height          =   285
            Left            =   -69360
            MultiLine       =   -1  'True
            OLEDragMode     =   1  'Automatic
            OLEDropMode     =   1  'Manual
            TabIndex        =   57
            Top             =   540
            Width           =   3825
         End
         Begin VB.TextBox txtFonteDeRecursoEmpenho 
            BackColor       =   &H80000000&
            CausesValidation=   0   'False
            Enabled         =   0   'False
            Height          =   285
            Left            =   5640
            MultiLine       =   -1  'True
            OLEDragMode     =   1  'Automatic
            OLEDropMode     =   1  'Manual
            TabIndex        =   55
            Top             =   540
            Width           =   3825
         End
         Begin VB.CommandButton cmd_Despesa 
            Height          =   300
            Left            =   -73080
            Picture         =   "CadOrdemPagamento.frx":1458
            Style           =   1  'Graphical
            TabIndex        =   35
            TabStop         =   0   'False
            ToolTipText     =   "Clique para cadastar despesa"
            Top             =   510
            Width           =   330
         End
         Begin VB.CommandButton cmd_Resto 
            Height          =   300
            Left            =   -73080
            Picture         =   "CadOrdemPagamento.frx":17E2
            Style           =   1  'Graphical
            TabIndex        =   32
            TabStop         =   0   'False
            Tag             =   "246"
            ToolTipText     =   "Clique para cadastar resto a pagar"
            Top             =   510
            Width           =   330
         End
         Begin VB.CommandButton cmd_Empenho 
            Height          =   300
            Left            =   1920
            Picture         =   "CadOrdemPagamento.frx":1B6C
            Style           =   1  'Graphical
            TabIndex        =   29
            TabStop         =   0   'False
            Tag             =   "241"
            ToolTipText     =   "Clique para cadastar empenho"
            Top             =   510
            Width           =   330
         End
         Begin MSDataListLib.DataCombo dcbEmpenho 
            Height          =   315
            Left            =   720
            TabIndex        =   28
            Tag             =   "1"
            Top             =   510
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSComctlLib.ListView lvw_Empenho 
            Height          =   1545
            Left            =   90
            TabIndex        =   38
            Top             =   900
            Width           =   9405
            _ExtentX        =   16589
            _ExtentY        =   2725
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   10
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "EmpenhoParcela"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Número"
               Object.Width           =   1808
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Parcela"
               Object.Width           =   1808
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Previsão"
               Object.Width           =   1940
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Liquidação"
               Object.Width           =   1940
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   5
               Text            =   "Valor"
               Object.Width           =   2558
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   6
               Text            =   "Desconto"
               Object.Width           =   2558
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   7
               Text            =   "Líquido"
               Object.Width           =   2558
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Text            =   "Processo"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   9
               Text            =   "bytAdiatamento"
               Object.Width           =   0
            EndProperty
         End
         Begin MSComctlLib.ListView lvw_Resto 
            Height          =   1545
            Left            =   -74910
            TabIndex        =   39
            Top             =   900
            Width           =   9405
            _ExtentX        =   16589
            _ExtentY        =   2725
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   10
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Exercício"
               Object.Width           =   1808
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Resto"
               Object.Width           =   1808
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Parcela"
               Object.Width           =   1940
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Previsão"
               Object.Width           =   1940
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   5
               Text            =   "Valor"
               Object.Width           =   2558
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   6
               Text            =   "Desconto"
               Object.Width           =   2558
            EndProperty
            BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   7
               Text            =   "Líquido"
               Object.Width           =   2558
            EndProperty
            BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   8
               Text            =   "processo"
               Object.Width           =   0
            EndProperty
            BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   9
               Text            =   "bytAdiatamento"
               Object.Width           =   0
            EndProperty
         End
         Begin MSDataListLib.DataCombo dcbResto 
            Height          =   315
            Left            =   -74280
            TabIndex        =   31
            Tag             =   "1"
            Top             =   510
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo dcbDespesa 
            Height          =   315
            Left            =   -74280
            TabIndex        =   34
            Tag             =   "1"
            Top             =   510
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSComctlLib.ListView lvw_Despesa 
            Height          =   1545
            Left            =   -74910
            TabIndex        =   36
            Top             =   900
            Width           =   9405
            _ExtentX        =   16589
            _ExtentY        =   2725
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   7
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Número"
               Object.Width           =   1808
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Previsão"
               Object.Width           =   1940
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   2
               Text            =   "Valor"
               Object.Width           =   2558
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "Desconto"
               Object.Width           =   2558
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   4
               Text            =   "Liquido"
               Object.Width           =   2558
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Credor"
               Object.Width           =   8819
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Processo"
               Object.Width           =   0
            EndProperty
         End
         Begin MSComctlLib.ListView lvw_AnulacaoReceita 
            Height          =   1545
            Left            =   -74910
            TabIndex        =   67
            Top             =   900
            Width           =   9405
            _ExtentX        =   16589
            _ExtentY        =   2725
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   3
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Descrição"
               Object.Width           =   8863
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   1
               Text            =   "Valor"
               Object.Width           =   2558
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   2
               Text            =   "Processo"
               Object.Width           =   3598
            EndProperty
         End
         Begin VB.Label lbl_Valor 
            Caption         =   "Valor :"
            Height          =   285
            Left            =   -69540
            TabIndex        =   62
            Top             =   510
            Width           =   465
         End
         Begin VB.Label lbl_Descricao 
            Caption         =   "Descrição :"
            Height          =   285
            Left            =   -74910
            TabIndex        =   60
            Top             =   510
            Width           =   825
         End
         Begin VB.Label lbl_FonteDeRecursoResto 
            AutoSize        =   -1  'True
            Caption         =   "F.de Recurso"
            Height          =   195
            Left            =   -70440
            TabIndex        =   58
            Top             =   570
            Width           =   960
         End
         Begin VB.Label lbl_FonteDeRecursoEmpenho 
            AutoSize        =   -1  'True
            Caption         =   "F.de Recurso"
            Height          =   195
            Left            =   4605
            TabIndex        =   56
            Top             =   570
            Width           =   960
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Número"
            Height          =   195
            Left            =   -74910
            TabIndex        =   46
            Top             =   570
            Width           =   555
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Número"
            Height          =   195
            Left            =   -74910
            TabIndex        =   45
            Top             =   570
            Width           =   555
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Parcela"
            Height          =   195
            Left            =   -72600
            TabIndex        =   44
            Top             =   570
            Width           =   540
         End
         Begin VB.Label lbl_Parcela 
            AutoSize        =   -1  'True
            Caption         =   "Parcela"
            Height          =   195
            Left            =   2400
            TabIndex        =   43
            Top             =   570
            Width           =   540
         End
         Begin VB.Label lbl_Empenho 
            AutoSize        =   -1  'True
            Caption         =   "Número"
            Height          =   195
            Left            =   90
            TabIndex        =   42
            Top             =   570
            Width           =   555
         End
      End
      Begin VB.Frame fra_HistoricoLiquidacao 
         Caption         =   " Histórico "
         Height          =   720
         Left            =   3840
         TabIndex        =   23
         Top             =   1065
         Width           =   5895
         Begin VB.TextBox txtHistorico 
            Height          =   540
            Left            =   0
            MaxLength       =   4000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   24
            Top             =   180
            Width           =   5895
         End
      End
      Begin VB.CommandButton cmd_HistoricoLiquidacao 
         Height          =   300
         Left            =   9405
         Picture         =   "CadOrdemPagamento.frx":1EF6
         Style           =   1  'Graphical
         TabIndex        =   27
         TabStop         =   0   'False
         Tag             =   "248"
         ToolTipText     =   "Clique para cadastar histórico"
         Top             =   1845
         Width           =   330
      End
      Begin VB.ComboBox cbo_HistoricoLiquidacao 
         Height          =   315
         Left            =   4740
         OLEDragMode     =   1  'Automatic
         Sorted          =   -1  'True
         TabIndex        =   26
         ToolTipText     =   "Histórico padrão"
         Top             =   1830
         Width           =   4575
      End
      Begin VB.TextBox txtPKId 
         Enabled         =   0   'False
         Height          =   285
         Left            =   6390
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   40
         Top             =   0
         Visible         =   0   'False
         Width           =   555
      End
      Begin MSDataListLib.DataCombo dbcintCredor 
         Height          =   315
         Left            =   5640
         TabIndex        =   8
         Top             =   375
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Label lbl_Processo 
         AutoSize        =   -1  'True
         Caption         =   "Processo"
         Height          =   195
         Left            =   4245
         TabIndex        =   14
         Top             =   825
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Anulação"
         Height          =   195
         Left            =   5970
         TabIndex        =   68
         Top             =   4950
         Width           =   675
      End
      Begin VB.Label lbl_DataVencimento 
         AutoSize        =   -1  'True
         Caption         =   "Dt. Venc."
         Height          =   195
         Left            =   1920
         TabIndex        =   12
         Top             =   825
         Width           =   675
      End
      Begin VB.Label lbl_DataOrdem 
         AutoSize        =   -1  'True
         Caption         =   "Data"
         Height          =   195
         Left            =   330
         TabIndex        =   10
         Top             =   825
         Width           =   345
      End
      Begin VB.Label lbl_Fornecedor 
         AutoSize        =   -1  'True
         Caption         =   "Credor"
         Height          =   195
         Left            =   4440
         TabIndex        =   6
         Top             =   450
         Width           =   465
      End
      Begin VB.Label lbl_TotalEmpenho 
         AutoSize        =   -1  'True
         Caption         =   "Empenho"
         Height          =   195
         Left            =   90
         TabIndex        =   54
         Top             =   4950
         Width           =   675
      End
      Begin VB.Label lbl_TotalResto 
         AutoSize        =   -1  'True
         Caption         =   "Resto"
         Height          =   195
         Left            =   2130
         TabIndex        =   53
         Top             =   4950
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Despesa"
         Height          =   195
         Left            =   3960
         TabIndex        =   52
         Top             =   4950
         Width           =   630
      End
      Begin VB.Label lblTotalAPagar 
         AutoSize        =   -1  'True
         Caption         =   "Total"
         Height          =   195
         Left            =   7935
         TabIndex        =   51
         Top             =   4950
         Width           =   360
      End
      Begin VB.Label lblstrProcesso 
         AutoSize        =   -1  'True
         Caption         =   "Número"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   450
         Width           =   555
      End
   End
End
Attribute VB_Name = "frmCadOrdemPagamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
    Dim mblnAlterando           As Boolean
    Dim mblnselecionou          As Boolean
    Dim mblnAlterandoEmpenho    As Boolean
    Dim mblnAlterandoResto      As Boolean
    Dim mblnAlterandoConta      As Boolean
    Dim mblnInclirConta         As Boolean
    Dim mblnAlterandoDespesa    As Boolean
    Dim mobjLista               As Object
    Dim mobjAux                 As Object
    Dim mblnClickOk             As Boolean
    Dim intPkid                 As Long
    Dim Codigo                  As String
    Dim Digito                  As String
    Dim Exercicio               As String
    Dim itemAnterior            As String
    Dim mblnTelaCredor          As Boolean
    Public mblnTelaEmpenho      As Boolean
    Public strpkidParcela       As String
    Dim intPKIDEmpenho          As Long
    
Private Function gstrValorLiquidadoExtra() As String
    
Dim i             As Integer
Dim valorAcumulado As Double
    
    For i = 1 To lvw_Despesa.ListItems.Count
        valorAcumulado = Val(gstrConvVrParaSql(valorAcumulado)) + Val(gstrConvVrParaSql(lvw_Despesa.ListItems(i).SubItems(2)))
    Next
    gstrValorLiquidadoExtra = gstrConvVrParaSql(valorAcumulado)
    
End Function
    
Private Sub ImprimeBordereaux(lngProcesso As Long)

Dim strSQL  As String
      
      rptOrdemDePagamento.intOrigem = 1
      If Not optbytTipo(2).Value Then
         
         If optbytTipo(3).Value = True Then
            strSQL = "SELECT "
            strSQL = strSQL & "1 intOrigem,"
            strSQL = strSQL & "OP.PKID, "
            strSQL = strSQL & "OP.bytTipo, "
            strSQL = strSQL & "PAR.strDescricao strDescricao, "
            strSQL = strSQL & "OP.intNumero intOrdem, "
            strSQL = strSQL & "OP.typHistorico, "
            strSQL = strSQL & "OP.dtmData, "
            strSQL = strSQL & "OP.dtmDataVencimento, "
            strSQL = strSQL & "OP.intExercicio IntExercicioOP,"
            strSQL = strSQL & "PAR.strCodigo, "
            strSQL = strSQL & "PAR.intExercicioProcesso, "
            strSQL = strSQL & "PAR.bitDigito, "
            strSQL = strSQL & "CT.strNome, "
            strSQL = strSQL & "CT.CDC intContribuinte,"
            strSQL = strSQL & "CT.strCNPJCPF, "
            strSQL = strSQL & "CT.strLogradouroC strEndereco, "
            strSQL = strSQL & "CT.intNumero, "
            strSQL = strSQL & "CT.strComplemento strComplemento, "
            strSQL = strSQL & "CT.bytNaturezaJuridica,"
            strSQL = strSQL & "MP.strDescricao strMunicipio, "
            strSQL = strSQL & "UF.strSigla strUF, "
            strSQL = strSQL & "BR.strDescricao strBairro, "
            strSQL = strSQL & "CP.intCEP, "
            strSQL = strSQL & "CT.Strlogradouroc,"
            strSQL = strSQL & "CT.intNumeroC,"
            strSQL = strSQL & "CT.strComplementoC,"
            strSQL = strSQL & "MPC.strDescricao strMunicipioC,"
            strSQL = strSQL & "UFC.strSigla strUFC,"
            strSQL = strSQL & "CT.strBairroC,"
            strSQL = strSQL & "CT.intCEPC,"
            strSQL = strSQL & " PAR.DblValor dblLiquidadoTotal,"
            strSQL = strSQL & gstrConvVrParaSql(txtTotalAPagar) & " dblLiquidototal, "
            strSQL = strSQL & gstrConvVrParaSql(txtTotalAPagar) & " dblValortotal, "
            strSQL = strSQL & " 0 dblDesconto "
            strSQL = strSQL & "FROM "
            strSQL = strSQL & gstrOrdemPagamento & " OP, "
            strSQL = strSQL & gstrOrdemPagamentoAnulacaoReceita & " PAR, "
            strSQL = strSQL & gstrContribuinte & " CT, "
            strSQL = strSQL & gstrCidade & " MP, "
            strSQL = strSQL & gstrUF & " UF, "
            strSQL = strSQL & gstrCidade & " MPC, "
            strSQL = strSQL & gstrUF & " UFC, "
            strSQL = strSQL & gstrBairro & " BR, "
            strSQL = strSQL & gstrCeps & " CP "
            strSQL = strSQL & "WHERE "
            strSQL = strSQL & "OP.bytTipo = 3 AND "
            strSQL = strSQL & "OP.PKID = PAR.intOrdemPagamento AND "
            strSQL = strSQL & "OP.intContribuinte = CT.PKID AND "
            strSQL = strSQL & "MP.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intMunicipio And "
            strSQL = strSQL & "MPC.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intMunicipioC And "
            strSQL = strSQL & "CP.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intCep AND "
            strSQL = strSQL & "BR.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intBairro  AND "
            strSQL = strSQL & "UF.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intUF AND "
            strSQL = strSQL & "UFC.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intUFC AND "
            strSQL = strSQL & "OP.PKID = " & lngProcesso & " "
            
            rptOrdemDePagamentoExtra.blnAnulacaoReceita = True
            ImprimeRelatorio rptOrdemDePagamentoExtra, strSQL, "Ordem de Pagamento Extra - Orcamentaria"
         Else
         
            strSQL = ""
            strSQL = strSQL & "SELECT "
            strSQL = strSQL & "1 intOrigem,"
            strSQL = strSQL & "OP.PKID,"
            strSQL = strSQL & "EP.Pkid PkidEmpenho,"
            strSQL = strSQL & "OP.bytTipo,"
            strSQL = strSQL & "OP.intNumero intOrdem,"
            strSQL = strSQL & "OP.typHistorico,"
            strSQL = strSQL & "OP.dtmData,"
            strSQL = strSQL & "OP.dtmDataVencimento,"
            strSQL = strSQL & "OP.intExercicio IntExercicioOP,"
            strSQL = strSQL & "OPE.intParcela intParcelaOP,"
            strSQL = strSQL & "CT.strNome,"
            strSQL = strSQL & "CT.CDC intContribuinte,"
            strSQL = strSQL & "CT.strCNPJCPF,"
            strSQL = strSQL & "CT.strLogradouroC strEndereco,"
            strSQL = strSQL & "CT.intNumero,"
            strSQL = strSQL & "CT.strComplemento strComplemento,"
            strSQL = strSQL & "CT.bytNaturezaJuridica,"
            
            'As 3 proximas linhas só aparecem no relatorio quando o RPX estiver sendo executado, ao invés do form do projeto.
            strSQL = strSQL & "CB.strBanco " & strCONCAT & "' - '" & strCONCAT & " cb.strBancoDescricao strBanco, "
            strSQL = strSQL & "CB.strAgencia, "
            strSQL = strSQL & "CB.strContaCorrente, "
            
            
            strSQL = strSQL & "MP.strDescricao strMunicipio,"
            strSQL = strSQL & "UF.strSigla strUF,"
            strSQL = strSQL & "BR.strDescricao strBairro,"
            strSQL = strSQL & "CP.intCEP,"
            strSQL = strSQL & "CT.Strlogradouroc,"
            strSQL = strSQL & "CT.intNumeroC,"
            strSQL = strSQL & "CT.strComplementoC,"
            strSQL = strSQL & "MPC.strDescricao strMunicipioC,"
            strSQL = strSQL & "UFC.strSigla strUFC,"
            strSQL = strSQL & "CT.strBairroC,"
            strSQL = strSQL & "CT.intCEPC,"
            strSQL = strSQL & "EP.intNumero intEmpenho,"
            strSQL = strSQL & "EP.dblValor  dblValorEmpenho,"
            strSQL = strSQL & "Case When " & gstrISNULL("OP.Strcodigoprocesso", "' '") & " = ' ' Then EP.strCodigo else OP.Strcodigoprocesso End strCodigo, "
            strSQL = strSQL & "Case When " & gstrISNULL("OP.Strcodigoprocesso", "' '") & " = ' ' Then EP.intExercicio else OP.Intexercicioprocesso End intExercicioProcesso, "
            strSQL = strSQL & "Case When " & gstrISNULL("OP.Strcodigoprocesso", "' '") & " = ' ' Then EP.bitDigito else OP.Bitdigitoprocesso End bitDigito, "
            
            'strSql = strSql & "EP.strCodigo, "
            'strSql = strSql & "EP.intExercicio intExercicioProcesso, "
            'strSql = strSql & "EP.bitDigito, "
            strSQL = strSQL & "SEP.intNumero intParcela,"
            strSQL = strSQL & "SEP.dblValor  dblValorParcela,"
            strSQL = strSQL & gstrConvVrParaSql(txtTotalAPagar) & " dblLiquidoTotal, "
            strSQL = strSQL & gstrConvVrParaSql(dblValorTotalDescontoEmpenho) & "dblDesconto, "
            strSQL = strSQL & "(SELECT " & gstrISNULL("SUM(SEP.dblValor)", "0")
            strSQL = strSQL & "FROM "
            strSQL = strSQL & gstrSubempenho & " SEP "
            strSQL = strSQL & "WHERE "
            strSQL = strSQL & "EP.PKId = SEP.intEmpenho AND "
            strSQL = strSQL & "OPE.intParcela = SEP.PKId) - (SELECT " & gstrISNULL("SUM(SL.dblValor)", "0")
            strSQL = strSQL & "FROM " & gstrSubempenhoLiquidado & " SL "
            strSQL = strSQL & "WHERE "
            strSQL = strSQL & "SL.intParcela " & strOUTJOracle & " =" & strOUTJSQLServer & " SEP.PKid) dblLiquido, "
            strSQL = strSQL & "(SELECT " & gstrISNULL("SUM(SEP.dblValor)", "0") & " FROM " & gstrSubempenho & " SEP WHERE EP.PKId = SEP.intEmpenho AND SEP.intNumero = 0 AND SEP.Bytsituacao = 4 AND SEP.dtmData <= OP.dtmData)  DblValorAnulado ,"
            strSQL = strSQL & "FR.strDescricao strFonteRecurso, "
            strSQL = strSQL & "PT.intCodigoReduzido, "
            strSQL = strSQL & "PT.intExercicio, "
            strSQL = strSQL & "PT.strCodigo strDotacao "
            strSQL = strSQL & "FROM "
            strSQL = strSQL & gstrOrdemPagamento & " OP, "
            strSQL = strSQL & gstrOrdemPagamentoEmpenho & " OPE, "
            strSQL = strSQL & gstrContribuinte & " CT, "
            strSQL = strSQL & gstrCidade & " MP, "
            strSQL = strSQL & gstrUF & " UF, "
            strSQL = strSQL & gstrCidade & " MPC, "
            strSQL = strSQL & gstrUF & " UFC, "
            strSQL = strSQL & gstrBairro & " BR, "
            strSQL = strSQL & gstrCeps & " CP, "
            strSQL = strSQL & gstrEmpenho & " EP, "
            strSQL = strSQL & gstrSubempenho & " SEP, "
            strSQL = strSQL & "tblCredorBanco CB, "
            strSQL = strSQL & gstrFonteRecurso & " FR, "
            strSQL = strSQL & gstrProgramaDeTrabalho & " PT "
            strSQL = strSQL & "WHERE "
            strSQL = strSQL & "OP.bytTipo = 0 AND "
            strSQL = strSQL & "OP.PKID = OPE.intOrdemPagamento AND "
            strSQL = strSQL & "OP.intContribuinte = CT.PKID AND "
            strSQL = strSQL & "OPE.intParcela = SEP.PKID  AND "
            strSQL = strSQL & "SEP.intEmpenho = EP.PKID   AND "
            strSQL = strSQL & "EP.intProgramaTrabalho = PT.PKID AND "
            strSQL = strSQL & "PT.intFonteRecurso = FR.PKID AND "
            strSQL = strSQL & "MP.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intMunicipio AND "
            strSQL = strSQL & "CP.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intCep AND "
            strSQL = strSQL & "MPC.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intMunicipioC AND "
            strSQL = strSQL & "BR.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intBairro  AND "
            strSQL = strSQL & "UF.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intUF AND "
            strSQL = strSQL & "UFC.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intUFC AND "
            strSQL = strSQL & "CT.pkid " & strOUTJSQLServer & "= CB.intContribuinte " & strOUTJOracle & " AND "
            strSQL = strSQL & "OP.PKID = " & lngProcesso & " "
            
            strSQL = strSQL & "UNION SELECT "
            strSQL = strSQL & "1 intOrigem,"
            strSQL = strSQL & "OP.PKID, "
            strSQL = strSQL & "EP.Pkid PkidEmpenho,"
            strSQL = strSQL & "OP.bytTipo, "
            strSQL = strSQL & "OP.intNumero intOrdem, "
            strSQL = strSQL & "OP.typHistorico, "
            strSQL = strSQL & "OP.dtmData, "
            strSQL = strSQL & "OP.dtmDataVencimento, "
            strSQL = strSQL & "OP.intExercicio IntExercicioOP,"
            strSQL = strSQL & "OPR.intParcela intParcelaOP, "
            
            strSQL = strSQL & "CT.strNome, "
            strSQL = strSQL & "CT.CDC intContribuinte,"
            strSQL = strSQL & "CT.strCNPJCPF, "
            strSQL = strSQL & "CT.strLogradouroC strEndereco, "
            strSQL = strSQL & "CT.intNumero, "
            strSQL = strSQL & "CT.strComplemento strComplemento, "
            strSQL = strSQL & "CT.bytNaturezaJuridica,"
            
            'As 3 proximas linhas só aparecem no relatorio quando o RPX estiver sendo executado, ao invés do form do projeto.
            strSQL = strSQL & "CB.strBanco " & strCONCAT & "' - '" & strCONCAT & " cb.strBancoDescricao strBanco, "
            strSQL = strSQL & "CB.strAgencia, "
            strSQL = strSQL & "CB.strContaCorrente, "
                        
            strSQL = strSQL & "MP.strDescricao strMunicipio, "
            strSQL = strSQL & "UF.strSigla strUF, "
            strSQL = strSQL & "BR.strDescricao strBairro, "
            strSQL = strSQL & "CP.intCEP, "
            strSQL = strSQL & "CT.Strlogradouroc,"
            strSQL = strSQL & "CT.intNumeroC,"
            strSQL = strSQL & "CT.strComplementoC,"
            strSQL = strSQL & "MPC.strDescricao strMunicipioC,"
            strSQL = strSQL & "UFC.strSigla strUFC,"
            strSQL = strSQL & "CT.strBairroC,"
            strSQL = strSQL & "CT.intCEPC,"
            strSQL = strSQL & "EP.intNumero intEmpenho, "
            strSQL = strSQL & "EP.dblValor  dblValorEmpenho, "
            strSQL = strSQL & "EP.strCodigo, "
            strSQL = strSQL & "EP.intExercicio intExercicioProcesso, "
            strSQL = strSQL & "EP.bitDigito, "
            strSQL = strSQL & "SEP.intNumero intParcela, "
            strSQL = strSQL & "SEP.dblValor  dblValorParcela, "
            strSQL = strSQL & gstrConvVrParaSql(txtTotalAPagar) & " dblLiquidoTotal,"
            strSQL = strSQL & gstrConvVrParaSql(dblValorTotalDescontoRestos) & "dblDesconto, "
            strSQL = strSQL & "(SELECT " & gstrISNULL("SUM(SEP.dblValor)", "0")
            strSQL = strSQL & " FROM "
            strSQL = strSQL & " tblSubEmpenho SEP "
            strSQL = strSQL & "WHERE "
            strSQL = strSQL & "EP.PKId = SEP.intEmpenho AND "
            strSQL = strSQL & "OPR.intParcela = SEP.PKId) - (SELECT " & gstrISNULL("SUM(SL.dblValor)", "0")
            strSQL = strSQL & "FROM tblSubempenholiquidado SL "
            strSQL = strSQL & "WHERE "
            strSQL = strSQL & "SL.intParcela " & strOUTJOracle & " =" & strOUTJSQLServer & " SEP.PKid) dblLiquido, "
            strSQL = strSQL & "(SELECT " & gstrISNULL("SUM(SEP.dblValor)", "0") & " FROM " & gstrSubempenho & " SEP WHERE EP.PKId = SEP.intEmpenho AND SEP.intNumero = 0  AND SEP.Bytsituacao = 4 AND SEP.dtmData <= OP.dtmData)  DblValorAnulado ,"
            strSQL = strSQL & "FR.strDescricao strFonteRecurso, "
            strSQL = strSQL & "PT.intCodigoReduzido, "
            strSQL = strSQL & "PT.intExercicio, "
            strSQL = strSQL & "PT.strCodigo strDotacao "
            strSQL = strSQL & "FROM "
            strSQL = strSQL & gstrOrdemPagamento & " OP, "
            strSQL = strSQL & gstrOrdemPagamentoResto & " OPR, "
            strSQL = strSQL & gstrContribuinte & " CT, "
            strSQL = strSQL & gstrCidade & " MP, "
            strSQL = strSQL & gstrUF & " UF, "
            strSQL = strSQL & gstrCidade & " MPC, "
            strSQL = strSQL & gstrUF & " UFC, "
            strSQL = strSQL & gstrBairro & " BR, "
            strSQL = strSQL & gstrCeps & " CP, "
            strSQL = strSQL & gstrEmpenho & " EP, "
            strSQL = strSQL & gstrSubempenho & " SEP, "
            strSQL = strSQL & "tblCredorBanco CB, "
            strSQL = strSQL & gstrFonteRecurso & " FR, "
            strSQL = strSQL & gstrProgramaDeTrabalho & " PT "
            strSQL = strSQL & "WHERE "
            strSQL = strSQL & "OP.bytTipo = 1 AND "
            strSQL = strSQL & "OP.PKID = OPR.intOrdemPagamento AND "
            strSQL = strSQL & "OP.intContribuinte = CT.PKID AND "
            strSQL = strSQL & "OPR.intParcela = SEP.PKID  AND "
            strSQL = strSQL & "SEP.intEmpenho = EP.PKID     AND "
            strSQL = strSQL & "EP.intExercicioRP IS NOT NULL AND "
            strSQL = strSQL & "EP.intProgramaTrabalho = PT.PKID AND "
            strSQL = strSQL & "PT.intFonteRecurso = FR.PKID AND "
            strSQL = strSQL & "MP.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intMunicipio AND "
            strSQL = strSQL & "MPC.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intMunicipioC AND "
            strSQL = strSQL & "CP.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intCep AND "
            strSQL = strSQL & "BR.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intBairro  AND "
            strSQL = strSQL & "UF.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intUF AND "
            strSQL = strSQL & "UFC.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intUFC AND "
            strSQL = strSQL & "CT.pkid " & strOUTJSQLServer & "= CB.intContribuinte " & strOUTJOracle & " AND "
            strSQL = strSQL & "OP.PKID = " & lngProcesso & " "
            strSQL = strSQL & "ORDER BY intEmpenho, intParcela, IntExercicioOP"
         
            rptOrdemDePagamento.PageSettings.LeftMargin = 1190.551
            rptOrdemDePagamento.PageSettings.RightMargin = 283.4646
            rptOrdemDePagamento.PageSettings.TopMargin = 1303.937
            rptOrdemDePagamento.PageSettings.BottomMargin = 226.7717
            
            ImprimeRelatorio rptOrdemDePagamento, strSQL, "Ordem de Pagamento"
         End If
      Else
         'para OPs Extra-Orçamentárias
         strSQL = "SELECT "
         strSQL = strSQL & "1 intOrigem,"
         strSQL = strSQL & "OP.PKID, "
         strSQL = strSQL & "OP.bytTipo, "
         strSQL = strSQL & "OP.intNumero intOrdem, "
         strSQL = strSQL & "OP.typHistorico, "
         strSQL = strSQL & "OP.dtmData, "
         strSQL = strSQL & "OP.dtmDataVencimento, "
         strSQL = strSQL & "OP.intExercicio IntExercicioOP,"
         strSQL = strSQL & "PP.strCodigo, "
         strSQL = strSQL & "PP.intExercicio intExercicioProcesso, "
         strSQL = strSQL & "PP.bitDigito, "
         strSQL = strSQL & "CT.strNome, "
         strSQL = strSQL & "CT.CDC intContribuinte,"
         strSQL = strSQL & "CT.strCNPJCPF, "
         strSQL = strSQL & "CT.strLogradouroC strEndereco, "
         strSQL = strSQL & "CT.intNumero, "
         strSQL = strSQL & "CT.strComplemento strComplemento, "
         strSQL = strSQL & "CT.bytNaturezaJuridica,"
         
         'As 3 proximas linhas só aparecem no relatorio quando o RPX estiver sendo executado, ao invés do form do projeto.
         strSQL = strSQL & "CB.strBanco " & strCONCAT & "' - '" & strCONCAT & " cb.strBancoDescricao strBanco, "
         strSQL = strSQL & "CB.strAgencia, "
         strSQL = strSQL & "CB.strContaCorrente, "
         
         strSQL = strSQL & "MP.strDescricao strMunicipio, "
         strSQL = strSQL & "UF.strSigla strUF, "
         strSQL = strSQL & "BR.strDescricao strBairro, "
         strSQL = strSQL & "CP.intCEP, "
         strSQL = strSQL & "CT.Strlogradouroc,"
         strSQL = strSQL & "CT.intNumeroC,"
         strSQL = strSQL & "CT.strComplementoC,"
         strSQL = strSQL & "MPC.strDescricao strMunicipioC,"
         strSQL = strSQL & "UFC.strSigla strUFC,"
         strSQL = strSQL & "CT.strBairroC,"
         strSQL = strSQL & "CT.intCEPC,"
         strSQL = strSQL & gstrConvVrParaSql(txtTotalAPagar) & " dblLiquidoTotal,"
         strSQL = strSQL & gstrValorLiquidadoExtra & " dblLiquidadoTotal , "
         strSQL = strSQL & gstrConvVrParaSql(dblValorTotalDescontoDoExtra) & "dblDesconto, "
         strSQL = strSQL & " DEX.PKID PKIDDespExtra "
         strSQL = strSQL & "FROM "
         strSQL = strSQL & gstrOrdemPagamento & " OP, "
         strSQL = strSQL & gstrProtocolizacaoProcesso & " PP, "
         strSQL = strSQL & gstrDespesaExtraOrcamentaria & " DEX, "
         strSQL = strSQL & gstrOrdemPagamentoDespesaExtra & " OPEX, "
         strSQL = strSQL & gstrContribuinte & " CT, "
         strSQL = strSQL & gstrCidade & " MP, "
         strSQL = strSQL & gstrUF & " UF, "
         strSQL = strSQL & gstrCidade & " MPC, "
         strSQL = strSQL & gstrUF & " UFC, "
         strSQL = strSQL & gstrBairro & " BR, "
         strSQL = strSQL & "tblCredorBanco CB, "
         strSQL = strSQL & gstrCeps & " CP "
         
         strSQL = strSQL & "WHERE "
         strSQL = strSQL & "OP.bytTipo = 2 AND "
         strSQL = strSQL & "OP.PKID = OPEX.intOrdemPagamento AND "
         strSQL = strSQL & "OPEX.intDespesaExtraOrcamentaria = DEX.PKID AND "
         strSQL = strSQL & "PP.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " DEX.intProtocolizacaoProcesso AND "
         strSQL = strSQL & "OP.intContribuinte = CT.PKID AND "
         strSQL = strSQL & "MP.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intMunicipio AND "
         strSQL = strSQL & "MPC.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intMunicipioC AND "
         strSQL = strSQL & "CP.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intCep AND "
         strSQL = strSQL & "BR.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intBairro  AND "
         strSQL = strSQL & "UF.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intUF AND "
         strSQL = strSQL & "UFC.PKID " & strOUTJOracle & " =" & strOUTJSQLServer & " CT.intUFC AND "
         strSQL = strSQL & "CT.pkid " & strOUTJSQLServer & "= CB.intContribuinte " & strOUTJOracle & " AND "
         strSQL = strSQL & "OP.PKID = " & lngProcesso & " "

         ImprimeRelatorio rptOrdemDePagamentoExtra, strSQL, "Ordem de Pagamento Extra - Orcamentaria"
      End If
End Sub


Private Sub CriaViewBordereaux()

'******************************************************************************************
' Data: 11/06/2003
' Alteração: - Incluída instrução IF fazendo com que, caso o banco de dados corrente seja o
'            Oracle, a função nada faça.
' Responsável: Everton Bianchini
'******************************************************************************************

    Dim strSQL     As String
    
    If (bytDBType = EDatabases.Oracle) Then Exit Sub
    
    Set gobjBanco = New clsBanco
    strSQL = "CREATE VIEW vw_Contribuinte AS SELECT * FROM " & gstrContribuinte
    gobjBanco.Execute strSQL
    strSQL = "CREATE VIEW vw_Banco AS SELECT * FROM " & gstrBanco
    gobjBanco.Execute strSQL
    strSQL = "CREATE VIEW vw_Agencia AS SELECT * FROM " & gstrAgencia
    gobjBanco.Execute strSQL
    strSQL = "CREATE VIEW vw_ContaBancaria AS SELECT * FROM " & gstrContaBancaria
    gobjBanco.Execute strSQL
End Sub

Private Sub LePagamento()
    With tdb_Lista
        LeEmpenho
        txtintProcesso = .Columns("intNumero")
    End With
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrCancelar
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrSalvar
End Sub

Private Sub LeEmpenho()

'******************************************************************************************
' Data: 09/06/2003
' Alteração: - Substituição da chamada direta à stored procedure pela função
'            gstrStoredProcedure.
' Responsável: Everton Bianchini
'******************************************************************************************
    Dim i                    As Integer
    Dim strSQL               As String
    Dim adoResultado         As ADODB.Recordset
    strSQL = ""
'    strSql = strSql & "sp_SubempenhoOrdem " & Val(txtPKId)
    strSQL = strSQL & gstrStoredProcedure("sp_SubempenhoOrdem", CStr(Val(txtPKId)), True)
    LeDaTabelaParaObj "", lvw_Empenho, strSQL
    
    For i = 1 To lvw_Empenho.ListItems.Count
        lvw_Empenho.ListItems(i).Text = lvw_Empenho.ListItems(i).SubItems(1) _
        & lvw_Empenho.ListItems(i).SubItems(2)
    Next
    
End Sub

Private Sub LeResto()
    
'******************************************************************************************
' Data: 09/06/2003
' Alteração: - Substituição da chamada direta à stored procedure pela função
'            gstrStoredProcedure.
' Responsável: Everton Bianchini
'******************************************************************************************
    Dim i                    As Integer
    Dim strSQL               As String
    Dim adoResultado         As ADODB.Recordset
    strSQL = ""
'    strSql = strSql & "sp_RestoOrdem " & Val(txtPKId)
    strSQL = strSQL & gstrStoredProcedure("sp_RestoOrdem", CStr(Val(txtPKId)), True)
    LeDaTabelaParaObj "", lvw_Resto, strSQL
    
    For i = 1 To lvw_Resto.ListItems.Count
        lvw_Resto.ListItems(i).Text = lvw_Resto.ListItems(i).SubItems(2) _
        & lvw_Resto.ListItems(i).SubItems(3)
    Next
End Sub

Private Sub LeDespesa()

'******************************************************************************************
' Data: 09/06/2003
' Alteração: - Substituição da chamada direta à stored procedure pela função
'            gstrStoredProcedure.
' Responsável: Everton Bianchini
'******************************************************************************************

    Dim strSQL               As String
    Dim adoResultado         As ADODB.Recordset
    strSQL = ""
'    strSql = strSql & "sp_DespesaOrdem " & Val(txtPKId)
    strSQL = strSQL & gstrStoredProcedure("sp_DespesaOrdem", CStr(Val(txtPKId)), True)
    LeDaTabelaParaObj "", lvw_Despesa, strSQL
End Sub
Private Sub LeAnulacao()

'******************************************************************************************
' Data: 09/06/2003
' Alteração: - Substituição da chamada direta à stored procedure pela função
'            gstrStoredProcedure.
' Responsável: Everton Bianchini
'******************************************************************************************

    Dim strSQL               As String
    Dim adoResultado         As ADODB.Recordset
    Dim intInd               As Integer
    Dim strProcesso          As String
    
    strSQL = ""
'    strSql = strSql & "sp_DespesaOrdem " & Val(txtPKId)
    'strSQL = strSQL & gstrStoredProcedure("sp_AnulacaoReceita", CStr(Val(txtPKId)), True)
    
    strSQL = " SELECT pkid, StrDescricao, dblValor, "
    strSQL = strSQL & " strCodigo " & strCONCAT & "'/'" & strCONCAT & gstrCONVERT(CDT_VARCHAR, " intExercicioProcesso ") & strCONCAT & "'-'" & strCONCAT & gstrCONVERT(CDT_VARCHAR, " bitDigito ")
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrOrdemPagamentoAnulacaoReceita
    strSQL = strSQL & " WHERE intOrdemPagamento = " & txtPKId
    strSQL = strSQL & " ORDER BY strDescricao"
    
    LeDaTabelaParaObj "", lvw_AnulacaoReceita, strSQL
    
    For intInd = 1 To lvw_AnulacaoReceita.ListItems.Count
       If Trim(lvw_AnulacaoReceita.ListItems(intInd).ListSubItems(2)) = "/-" Then
          lvw_AnulacaoReceita.ListItems(intInd).ListSubItems(2) = ""
       Else
          If InStr(1, lvw_AnulacaoReceita.ListItems(intInd).ListSubItems(2), "/") = 0 Then
            strProcesso = lvw_AnulacaoReceita.ListItems(intInd).ListSubItems(2)
             lvw_AnulacaoReceita.ListItems(intInd).ListSubItems(2) = Mid(strProcesso, 1, Len(strProcesso) - 5) & "/" & Mid(strProcesso, Len(strProcesso) - 4, 4) & "-" & Mid(strProcesso, Len(strProcesso), 1)
          End If
       End If
    Next
    
End Sub

Private Sub LimpaTelaPagamento(Optional blnSetaData As Boolean)
    intPKIDEmpenho = 0
    txtintProcesso = ""
    TrocaCorObjeto txtintProcesso, False
    cbo_HistoricoLiquidacao = ""
    txtTotalAnulacao = ""
    txtTotalDespesa = ""
    txtTotalResto = ""
    txtTotalEmpenho = ""
    txtintExercicio = gintExercicio
    lvw_Empenho.ListItems.Clear
    lvw_Resto.ListItems.Clear
    lvw_Despesa.ListItems.Clear
    lvw_AnulacaoReceita.ListItems.Clear
    LimpaDadosEmpenho True
    LimpaDadosResto True
    LimpaDadosDespesa True
    'AtualizaListas
    dcbEmpenho.Text = ""
    dcbResto.Text = ""
    chkblnPago.Value = 0
    dcbDespesa.Text = ""
    txtHistorico.Text = ""
    txtPKId.Text = ""
    'TrocaCorObjeto fra_bytTipo, False
    
    'Alteração feita por hugo 28/07/2005
    txtstrCodigoProcesso.Text = ""
    txtintExercicioProcesso.Text = ""
    txtbitDigitoProcesso.Text = ""

    mblnAlterando = False
    DesabilitaPago False
    habilitaGuias 4
    optbytTipo(0).Value = False
    optbytTipo(1).Value = False
    optbytTipo(2).Value = False
    optbytTipo(3).Value = False
    TrocaCorObjeto optbytTipo(0), False
    TrocaCorObjeto optbytTipo(1), False
    TrocaCorObjeto optbytTipo(2), False
    TrocaCorObjeto optbytTipo(3), False
    tab_3DPastaEmpenho.Tab = 0
    dcbEmpenho.Enabled = False
    cmd_Empenho.Enabled = False
    dcbParcela.Enabled = False
    txtdtmData.Text = ""
    txtdtmDataVencimento.Text = ""
    txt_intNContribuinte.Text = ""
    dbcintCredor.Text = ""
    txt_Descricao = "Anulação de Receita"
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrCancelar
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrSalvar
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrDeletar
    
    'M4R
    DataOpPadrao
    
    On Error Resume Next
    txtintProcesso.SetFocus
End Sub

Private Sub AtualizaListas()
    LeDaTabelaParaObj "", dcbEmpenho, strQueryEmpenho
    PreencheDados
    LeDaTabelaParaObj "", dcbResto, strQueryResto
    LeDaTabelaParaObj "", dcbDespesa, strQueryDespesa
End Sub

Private Function blnDadosOk() As Boolean

   Dim strSQL              As String
   Dim strTabelaVerificar  As String
   Dim strCampoVerificar   As String
   Dim intInd              As Integer
   Dim listaPKID           As String
   Dim adoResultado        As ADODB.Recordset
   Dim mstrCodigo          As String
   
   
   If Trim(txtintProcesso) = "" Then

        
            mstrCodigo = proximoCodigoOP
            
            If mstrCodigo = "" Then
                ExibeMensagem "É necessário informar o número para a Ordem de Pagamento."
                If txtintProcesso.Enabled Then txtintProcesso.SetFocus
                Exit Function
            ElseIf MsgBox("O número da Ordem de Pagamento está vazio. Deseja usar o número " & mstrCodigo & "?", vbYesNo + vbQuestion) = vbNo Then
                If txtintProcesso.Enabled Then txtintProcesso.SetFocus
                Exit Function
            Else
                txtintProcesso.Text = mstrCodigo
            End If

       
   End If
   
   If Trim(txtdtmData.Text) = "" Then
       ExibeMensagem "A Data da Ordem de Pagamento tem que ser informada."
       If txtdtmData.Enabled Then txtdtmData.SetFocus
       Exit Function
   End If
   
   If Val(txtintExercicio) <> gintExercicio Then
      ExibeMensagem "O Exercício da O.P. não pode ser diferente do Exercício atual."
      txtintExercicio.SetFocus
      MarcaCampo txtintExercicio
      Exit Function
   End If
   
   If DatePart("yyyy", txtdtmData.Text) <> gintExercicio Then
       ExibeMensagem "A Data da Ordem de Pagamento não pode estar em outro Exercício."
       If txtdtmData.Enabled Then txtdtmData.SetFocus
       Exit Function
   End If
   
   
    'Orc677
    If txtdtmDataVencimento <> "" Then
        If Right(txtdtmDataVencimento, 4) < gintExercicio Then
            ExibeMensagem "O ano da data de vencimento não pode ser menor que " & gintExercicio & "."
            If txtdtmDataVencimento.Enabled Then txtdtmDataVencimento.SetFocus
            Exit Function
        End If
    End If
    
    If blnValidarProcesso Then
       If Len(Trim(txtstrCodigoProcesso)) > 0 Or Len(Trim(txtbitDigitoProcesso)) > 0 Or Len(Trim(txtintExercicioProcesso)) > 0 Then
            If gblnExisteCodigo(2, gstrProtocolizacaoProcesso, "strCodigo", "'" & Trim(txtstrCodigoProcesso.Text) & "'", _
               "intExercicio", Trim(txtintExercicioProcesso.Text), "bitDigito", Trim(txtbitDigitoProcesso.Text)) = False Then
               ExibeMensagem "O Processo de informado não existe."
               txtstrCodigoProcesso.SetFocus
               Exit Function
            End If
       End If
    End If
   
   
   If lvw_Despesa.ListItems.Count = 0 And lvw_Resto.ListItems.Count = 0 And lvw_Empenho.ListItems.Count = 0 And lvw_AnulacaoReceita.ListItems.Count = 0 Then
       ExibeMensagem "É necessário no mínimo um lancamento na Ordem de Pagamento"
       Exit Function
   End If
   
   If blnVerificaDataDaOrdem = True Then
       ExibeMensagem "Existem empenhos com datas maiores que a data da Ordem de Pagamento"
       Exit Function
   End If
   
   If mblnAlterando Or lvw_AnulacaoReceita.ListItems.Count > 0 Then
        blnDadosOk = True
        Exit Function
   End If
   
    
   'Verifica se algum lancamento já esta presente em alguma OP ou Pagamento
   If optbytTipo(0).Value Then
      
      For intInd = 1 To lvw_Empenho.ListItems.Count
          listaPKID = listaPKID & lvw_Empenho.ListItems(intInd).Tag & ","
      Next
      If Len(listaPKID) > 0 Then listaPKID = Mid(listaPKID, 1, Len(listaPKID) - 1)
      
      'Verifica integridade das parcelas no banco
      If lvw_Empenho.ListItems.Count > 0 Then
           strSQL = "SELECT SE.Pkid,SE.dtmLiquidacao FROM " & gstrSubempenho & " SE WHERE SE.PKID IN (" & listaPKID & ")"
           Set gobjBanco = New clsBanco
           gobjBanco.CriaADO strSQL, 5, adoResultado
           For intInd = 1 To lvw_Empenho.ListItems.Count
               adoResultado.MoveFirst
               While Not adoResultado.EOF
                   If lvw_Empenho.ListItems(intInd).Tag = adoResultado!Pkid Then
                       If lvw_Empenho.ListItems(intInd).SubItems(4) <> Format(adoResultado!dtmLiquidacao, "dd/mm/yyyy") Then
                           ExibeMensagem "A parcela número " _
                           & lvw_Empenho.ListItems(intInd).SubItems(2) & " do empenho " _
                           & lvw_Empenho.ListItems(intInd).SubItems(1) & " foi alterada durante o processo de inclusão da Ordem de Pagamento em questão verifique esta parcela na tela de empenho."
                           Exit Function
                       End If
                   End If
                   adoResultado.MoveNext
               Wend
           Next
      End If
      
      strSQL = ""
      strSQL = strSQL & "SELECT OP.intNumero ORDEM , " & gstrCONVERT(CDT_VARCHAR, "E.intNumero ") & strCONCAT & "' \ '" & strCONCAT & gstrCONVERT(CDT_VARCHAR, " SE.intnumero ") & " lancamento"
      strSQL = strSQL & " FROM "
      strSQL = strSQL & gstrSubempenho & " SE,"
      strSQL = strSQL & gstrOrdemPagamentoEmpenho & ","
      strSQL = strSQL & gstrEmpenho & " E,"
      strSQL = strSQL & gstrOrdemPagamento & " OP"
      strSQL = strSQL & " WHERE "
      strSQL = strSQL & " intParcela in(" & listaPKID & ")"
      strSQL = strSQL & " AND OP.PKID = intordempagamento"
      strSQL = strSQL & " AND (OP.Bytcancelado = 0 OR OP.Bytcancelado IS NULL)"
      strSQL = strSQL & " AND SE.PKID = intParcela"
      strSQL = strSQL & " AND E.PKID = SE.intEmpenho"
   
   
      strSQL = strSQL & " UNION"

      strSQL = strSQL & " SELECT PP.intNumero ORDEM , " & gstrCONVERT(CDT_VARCHAR, "E.intNumero ") & strCONCAT & "' \P '" & strCONCAT & gstrCONVERT(CDT_VARCHAR, " SE.intnumero ") & " lancamento"
      strSQL = strSQL & " FROM "
      strSQL = strSQL & gstrSubempenho & " SE,"
      strSQL = strSQL & gstrProcessoPagamento & " PP,"
      strSQL = strSQL & gstrEmpenho & " E"
      strSQL = strSQL & " WHERE "
      strSQL = strSQL & " SE.PKID in(" & listaPKID & ") AND PP.PKID = SE.intProcesso"
      strSQL = strSQL & " AND SE.intEmpenho = E.PKID"
   
   ElseIf optbytTipo(1).Value Then
      
      For intInd = 1 To lvw_Resto.ListItems.Count
          listaPKID = listaPKID & lvw_Resto.ListItems(intInd).Tag & ","
      Next
      If Len(listaPKID) > 0 Then listaPKID = Mid(listaPKID, 1, Len(listaPKID) - 1)
   
      strSQL = ""
      strSQL = strSQL & "SELECT OP.intNumero ORDEM , " & gstrCONVERT(CDT_VARCHAR, "E.intNumero ") & strCONCAT & "' \P '" & strCONCAT & gstrCONVERT(CDT_VARCHAR, " SE.intnumero ") & " lancamento"
      strSQL = strSQL & " FROM "
      strSQL = strSQL & gstrSubempenho & " SE,"
      strSQL = strSQL & gstrOrdemPagamentoResto & ","
      strSQL = strSQL & gstrEmpenho & " E,"
      strSQL = strSQL & gstrOrdemPagamento & " OP"
      strSQL = strSQL & " WHERE "
      strSQL = strSQL & " intParcela in(" & listaPKID & ")"
      strSQL = strSQL & " AND OP.PKID = intordempagamento"
      strSQL = strSQL & " AND (OP.Bytcancelado = 0 OR OP.Bytcancelado IS NULL)"
      strSQL = strSQL & " AND SE.PKID = intParcela"
      strSQL = strSQL & " AND E.PKID = SE.intEmpenho"
   
      strSQL = strSQL & " UNION"

      strSQL = strSQL & " SELECT PP.intNumero ORDEM , " & gstrCONVERT(CDT_VARCHAR, "E.intNumero ") & strCONCAT & "' \P '" & strCONCAT & gstrCONVERT(CDT_VARCHAR, " SE.intnumero ") & " lancamento"
      strSQL = strSQL & " FROM "
      strSQL = strSQL & gstrSubempenho & " SE,"
      strSQL = strSQL & gstrProcessoPagamento & " PP,"
      strSQL = strSQL & gstrEmpenho & " E"
      strSQL = strSQL & " WHERE "
      strSQL = strSQL & " SE.PKID in(" & listaPKID & ") AND PP.PKID = SE.intProcesso"
      strSQL = strSQL & " AND SE.intEmpenho = E.PKID"
   
   ElseIf optbytTipo(2).Value Then
      
      For intInd = 1 To lvw_Despesa.ListItems.Count
          listaPKID = listaPKID & lvw_Despesa.ListItems(intInd).Tag & ","
      Next
      If Len(listaPKID) > 0 Then listaPKID = Mid(listaPKID, 1, Len(listaPKID) - 1)
      
      strSQL = ""
      strSQL = strSQL & "SELECT OP.intNumero ORDEM , '' " & strCONCAT & gstrCONVERT(CDT_VARCHAR, " DE.intnumero ") & " Lancamento"
      strSQL = strSQL & " FROM "
      strSQL = strSQL & gstrDespesaExtraOrcamentaria & " DE,"
      strSQL = strSQL & gstrOrdemPagamentoDespesaExtra & ","
      strSQL = strSQL & gstrOrdemPagamento & " OP"
      strSQL = strSQL & " WHERE "
      strSQL = strSQL & " intdespesaextraorcamentaria in(" & listaPKID & ")"
      strSQL = strSQL & " AND (OP.Bytcancelado = 0 OR OP.Bytcancelado IS NULL)"
      strSQL = strSQL & " AND OP.PKID = intordempagamento"
      strSQL = strSQL & " AND DE.PKID = intdespesaextraorcamentaria"
      
      strSQL = strSQL & " UNION"

      strSQL = strSQL & " SELECT PP.intNumero ORDEM , 'P' " & strCONCAT & gstrCONVERT(CDT_VARCHAR, " DE.intnumero ") & " lancamento"
      strSQL = strSQL & " FROM "
      strSQL = strSQL & gstrDespesaExtraOrcamentaria & " DE,"
      strSQL = strSQL & gstrProcessoPagamento & " PP"
      strSQL = strSQL & " WHERE "
      strSQL = strSQL & " DE.PKID in(" & listaPKID & ") AND PP.PKID = DE.intProcesso"
   End If
       

   Set gobjBanco = New clsBanco
   If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
      With adoResultado
          If .EOF = False Then
              If InStr(!Lancamento, "\P") <> 0 Then
                ExibeMensagem "O lançamento N°" & Replace(!Lancamento, "\P", "\") & " já está presente no Pagamento N°" & !Ordem
              ElseIf InStr(!Lancamento, "P") <> 0 Then
                ExibeMensagem "O lançamento N°" & Replace(!Lancamento, "P", "") & " já está presente no Pagamento N°" & !Ordem
              Else
                ExibeMensagem "O lançamento N°" & !Lancamento & " já está presente na Ordem de Pagamento N°" & !Ordem
              End If
              Exit Function
          End If
      End With
   End If
   
   blnDadosOk = True
End Function

Private Sub GravaOrdemPagamento()

'******************************************************************************************
' Data: 09/06/2003
' Alteração: - Substituição da chamada direta à stored procedure pela função
'            gstrStoredProcedure.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 11/06/2003
' Alteração: - Incluídos comandos de inicialização e finalização de bloco PL/SQL
'            permitindo, assim, a execução de múltiplos comandos SQL de uma única vez.
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 11/06/2003
' Alteração: - Incluídos os nomes das colunas no comando INSERT.
' Responsável: Everton Bianchini
'******************************************************************************************
    Dim strGravacao     As String
    Dim strSQL          As String
    Dim intInd          As Integer
    Dim adoResultado    As ADODB.Recordset
    Dim lngOrdem        As Long
    Dim mbytTipo        As Byte
    Dim mstrCodigo      As String
    Dim strProcesso     As String
    Dim strExercicioProcesso As String
    Dim strDigito       As String
    Dim intPosBarra     As Integer
    Dim intPosSinal     As Integer
    
    If blnDadosOk Then
        If mblnAlterando Then
            strGravacao = "Confirma Alteração do registro?"
            If Not dbcintCredor.Enabled Then
                strGravacao = "Confirma Alteração do histórico?"
            Else
                strGravacao = "Confirma Alteração do registro?"
            End If
            
            If Trim(txtPKId) = "" Then
                ExibeMensagem "Não há nenhum registro selecionado para a alteração."
            End If
        Else
            strGravacao = "Confirma Inclusão do registro?"
        End If
        
        If gblnExclusaoGravacaoOk("I", strGravacao, True) Then
              
            If optbytTipo(0).Value Then
                mbytTipo = 0
            ElseIf optbytTipo(1).Value Then
                mbytTipo = 1
            ElseIf optbytTipo(2).Value Then
                mbytTipo = 2
            ElseIf optbytTipo(3).Value Then
                mbytTipo = 3
            End If
                  
            Set gobjBanco = New clsBanco
            gobjBanco.ExecutaBeginTrans
            strSQL = ""
               
            If mblnAlterando Then
                strSQL = strSQL & "Update " & gstrOrdemPagamento & " "
                strSQL = strSQL & "Set typHistorico = '" & Trim(txtHistorico) & "' "
                strSQL = strSQL & ", dtmDtAtualizacao= " & gstrConvDtParaSql(gstrDataDoSistema)
                strSQL = strSQL & ", lngCodUsr= " & glngCodUsr
                
                'Alteração feita por Hugo 28/07/2005
                strSQL = strSQL & ", strcodigoprocesso= '" & Trim(txtstrCodigoProcesso) & "'"
                strSQL = strSQL & ", bitdigitoprocesso= " & gstrENulo(Trim(txtbitDigitoProcesso), , True)
                strSQL = strSQL & ", intexercicioprocesso= " & gstrENulo(Trim(txtintExercicioProcesso), , True)
                
                If dbcintCredor.Enabled Then
                    strSQL = strSQL & ", bytTipo= " & mbytTipo
                    strSQL = strSQL & ", intContribuinte= " & gstrItemData(dbcintCredor)
                End If
                
                strSQL = strSQL & ", dtmData= " & gstrConvDtParaSql(txtdtmData.Text)
                
                If txtdtmDataVencimento.Text <> "" Then
                     strSQL = strSQL & ", dtmDataVencimento= " & gstrConvDtParaSql(txtdtmDataVencimento.Text)
                End If
                
                strSQL = strSQL & " WHERE PKID = " & txtPKId.Text
                
                Set gobjBanco = New clsBanco
                gobjBanco.CriaADO strSQL, 5, adoResultado
                lngOrdem = txtPKId.Text
                
                If Not dbcintCredor.Enabled Then
                    gobjBanco.ExecutaCommitTrans
                    LimpaTelaPagamento True
                    LeDaTabelaParaObj "", tdb_Lista, strQuery
                    txtintProcesso = proximoCodigoOP
                    If Not tdb_Lista.EOF Then tdb_Lista.MoveLast
                    Exit Sub
                End If
            Else
            
ProximoCodigo:
                If gblnExisteCodigo(2, gstrOrdemPagamento, "intNumero", txtintProcesso, "intExercicio", "'" & txtintExercicio & "'") Then
                    mstrCodigo = proximoCodigoOP
                    If MsgBox("O número da Ordem de Pagamento informado já se encontra cadastrado. Deseja usar o número " & mstrCodigo & "?", vbYesNo + vbQuestion) = vbNo Then
                        Set gobjBanco = New clsBanco
                        gobjBanco.ExecutaRollbackTrans
                        If txtintProcesso.Enabled Then txtintProcesso.SetFocus
                        Exit Sub
                    Else
                        txtintProcesso.Text = mstrCodigo
                        GoTo ProximoCodigo
                    End If
                End If
            

                strSQL = gstrStoredProcedure("sp_GravaOrdemPagamento", _
                    txtintProcesso & ",'" & Trim(txtHistorico) & "', " & _
                    glngCodUsr & "," & CStr(chkblnPago.Value) & _
                    "," & CStr(mbytTipo) & "," & gstrItemData(dbcintCredor) & _
                    "," & gstrConvDtParaSql(txtdtmData.Text) & _
                    "," & IIf(txtdtmDataVencimento.Text = "", "NULL", gstrConvDtParaSql(txtdtmDataVencimento.Text)) & ", " & Trim(txtintExercicio) & ", '" & Trim(txtstrCodigoProcesso) & "', " & gstrENulo(Trim(txtbitDigitoProcesso), , True) & ", " & gstrENulo(Trim(txtintExercicioProcesso), , True), True)

                lngOrdem = RetornaInsertPkid(strSQL)
                If lngOrdem = -1 Then
                    Set gobjBanco = New clsBanco
                    gobjBanco.ExecutaRollbackTrans
                    Exit Sub
                 ElseIf lngOrdem = 0 Then
                    GoTo ProximoCodigo
                End If
                    
            End If
            
            strSQL = ""
            
            strSQL = strSQL & IIf((bytDBType = EDatabases.Oracle), "BEGIN ", "")
            
            strSQL = strSQL & "DELETE " & gstrOrdemPagamentoEmpenho & " "
            strSQL = strSQL & "WHERE intOrdemPagamento = " & lngOrdem & ";"
            For intInd = 1 To lvw_Empenho.ListItems.Count
                strSQL = strSQL & "INSERT INTO " & gstrOrdemPagamentoEmpenho & " "
                
                strSQL = strSQL & "(intOrdemPagamento, intParcela, dblValor, dblValorDesconto, dtmDtAtualizacao, lngCodUsr) "
                
                strSQL = strSQL & "VALUES ("
                strSQL = strSQL & lngOrdem & ", "
                strSQL = strSQL & lvw_Empenho.ListItems(intInd).Tag & ", "
                strSQL = strSQL & gstrConvVrParaSql(lvw_Empenho.ListItems(intInd).ListSubItems(5)) & ", "
                strSQL = strSQL & gstrConvVrParaSql(lvw_Empenho.ListItems(intInd).ListSubItems(6)) & ", "
                strSQL = strSQL & gstrConvDtParaSql(gstrDataDoSistema) & ", "
                strSQL = strSQL & glngCodUsr & ");"
            Next
            
            strSQL = strSQL & "DELETE " & gstrOrdemPagamentoResto & " "
            strSQL = strSQL & "WHERE intOrdemPagamento = " & lngOrdem & ";"
            For intInd = 1 To lvw_Resto.ListItems.Count
                strSQL = strSQL & "INSERT INTO " & gstrOrdemPagamentoResto & " "
                
                strSQL = strSQL & "(intOrdemPagamento, intParcela, dblValor, dblValorDesconto, dtmDtAtualizacao, lngCodUsr) "
                
                strSQL = strSQL & "VALUES ("
                strSQL = strSQL & lngOrdem & ", "
                strSQL = strSQL & lvw_Resto.ListItems(intInd).Tag & ", "
                strSQL = strSQL & gstrConvVrParaSql(lvw_Resto.ListItems(intInd).ListSubItems(5)) & ", "
                strSQL = strSQL & gstrConvVrParaSql(lvw_Resto.ListItems(intInd).ListSubItems(6)) & ", "
                strSQL = strSQL & gstrConvDtParaSql(gstrDataDoSistema) & ", "
                strSQL = strSQL & glngCodUsr & ");"
            Next
            strSQL = strSQL & "DELETE " & gstrOrdemPagamentoDespesaExtra & " "
            strSQL = strSQL & "WHERE intOrdemPagamento = " & lngOrdem & ";"
            For intInd = 1 To lvw_Despesa.ListItems.Count
                strSQL = strSQL & "INSERT INTO "
                strSQL = strSQL & gstrOrdemPagamentoDespesaExtra & " "
                
                strSQL = strSQL & "(intOrdemPagamento, intDespesaExtraOrcamentaria, dblValor, dblValorDesconto, dtmDtAtualizacao, lngCodUsr) "
                
                strSQL = strSQL & "VALUES ("
                strSQL = strSQL & lngOrdem & ", "
                strSQL = strSQL & lvw_Despesa.ListItems(intInd).Tag & ", "
                strSQL = strSQL & gstrConvVrParaSql(lvw_Despesa.ListItems(intInd).ListSubItems(2)) & ", "
                strSQL = strSQL & gstrConvVrParaSql(lvw_Despesa.ListItems(intInd).ListSubItems(3)) & ", "
                strSQL = strSQL & gstrConvDtParaSql(gstrDataDoSistema) & ", "
                strSQL = strSQL & glngCodUsr & ");"
            Next
            
            strSQL = strSQL & "DELETE " & gstrOrdemPagamentoAnulacaoReceita & " "
            strSQL = strSQL & "WHERE intOrdemPagamento = " & lngOrdem & ";"
            For intInd = 1 To lvw_AnulacaoReceita.ListItems.Count
                
                'rotina que desagrupa o processo do formato 00/0000.0
                If lvw_AnulacaoReceita.ListItems(intInd).ListSubItems.Count = 2 Then
                    If Trim(lvw_AnulacaoReceita.ListItems(intInd).ListSubItems(2)) <> "" Then
                        intPosBarra = InStr(1, lvw_AnulacaoReceita.ListItems(intInd).ListSubItems(2), "/", vbTextCompare)
                        intPosSinal = InStr(1, lvw_AnulacaoReceita.ListItems(intInd).ListSubItems(2), "-", vbTextCompare)
                        If intPosBarra > 1 Then
                            strProcesso = Mid(lvw_AnulacaoReceita.ListItems(intInd).ListSubItems(2), 1, intPosBarra - 1)
                        End If
                        
                        If intPosBarra > 1 And intPosSinal > 1 Then
                            strExercicioProcesso = Mid(lvw_AnulacaoReceita.ListItems(intInd).ListSubItems(2), intPosBarra + 1, intPosSinal - intPosBarra - 1)
                            strDigito = Mid(lvw_AnulacaoReceita.ListItems(intInd).ListSubItems(2), intPosSinal + 1)
                        End If
                    End If
                End If
                                
                strSQL = strSQL & "INSERT INTO "
                strSQL = strSQL & gstrOrdemPagamentoAnulacaoReceita & " "
                
                strSQL = strSQL & "(intOrdemPagamento, strDescricao, dblValor, dtmDtAtualizacao, lngCodUsr, strCodigo, intExercicioProcesso, bitDigito) "
                
                strSQL = strSQL & "VALUES ("
                strSQL = strSQL & lngOrdem & ", '"
                strSQL = strSQL & lvw_AnulacaoReceita.ListItems(intInd).Text & "', "
                strSQL = strSQL & gstrConvVrParaSql(lvw_AnulacaoReceita.ListItems(intInd).ListSubItems(1)) & ", "
                strSQL = strSQL & gstrConvDtParaSql(gstrDataDoSistema) & ", "
                strSQL = strSQL & glngCodUsr & ", "
                strSQL = strSQL & "'" & strProcesso & "', "
                strSQL = strSQL & IIf(Trim(strExercicioProcesso) <> "", Trim(strExercicioProcesso), "NULL") & ", "
                strSQL = strSQL & IIf(Trim(strDigito) <> "", Trim(strDigito), "NULL") & ");"
            Next
            
            strSQL = strSQL & IIf((bytDBType = EDatabases.Oracle), " END; ", "")
            
            Set gobjBanco = New clsBanco
            If gobjBanco.Execute(strSQL) Then
                Set gobjBanco = New clsBanco
                gobjBanco.ExecutaCommitTrans
                LimpaTelaPagamento True
                LeDaTabelaParaObj "", tdb_Lista, strQuery
                
                'estas lista mudam de status durante a gravação, por isso são atualizadas

                LeDaTabelaParaObj "", dcbEmpenho, strQueryEmpenho
                PreencheDados
                LeDaTabelaParaObj "", dcbResto, strQueryResto
                LeDaTabelaParaObj "", dcbDespesa, strQueryDespesa
                txtintProcesso = proximoCodigoOP
                If Not tdb_Lista.EOF Then tdb_Lista.MoveLast
            
            Else
                Set gobjBanco = New clsBanco
                gobjBanco.ExecutaRollbackTrans
            End If
        End If
    End If
    
    Exit Sub
    
End Sub

Private Function RetornaInsertPkid(ByVal strSQL As String) As Long
    Dim gcmdADOCmdConMain As New ADODB.Command
    Dim adorsResultado  As New ADODB.Recordset
    On Error GoTo Trata_Erro

    With gcmdADOCmdConMain
        .ActiveConnection = ModGeral.gcncADOMain
        .CommandText = strSQL
        .CommandTimeout = 5
        adorsResultado.LockType = adLockReadOnly
        adorsResultado.CursorLocation = adUseClient
        adorsResultado.CursorType = adOpenStatic
        adorsResultado.Open gcmdADOCmdConMain, , , adCmdText
    End With
    
    With adorsResultado
        If .EOF = False Then
            RetornaInsertPkid = !intOrdem
        Else
            RetornaInsertPkid = -1
        End If
    End With
    
    Exit Function
Trata_Erro:
    
    Set adorsResultado = Nothing
    Set gcmdADOCmdConMain = Nothing
    If Err.Number = -2147217873 Then
        Err.Clear
        RetornaInsertPkid = 0
    Else
        ExibeDetalheErro gstrMsgErroADO(Err, strSQL), strSQL
        RetornaInsertPkid = -1
        Err.Clear
    End If
End Function

Private Sub LimpaDadosEmpenho(Optional blnNaoSetaFoco As Boolean)
    If blnNaoSetaFoco = False Then
        dcbParcela.SetFocus
    End If
    dcbParcela.ListIndex = -1
    mblnAlterandoEmpenho = False
End Sub

Private Sub LimpaDadosResto(Optional blnNaoSetaFoco As Boolean)
    If blnNaoSetaFoco = False Then
        dcbParcelaResto.SetFocus
    End If
    dcbParcelaResto.ListIndex = -1
    mblnAlterandoResto = False
End Sub

Private Sub LimpaDadosDespesa(Optional blnNaoSetaFoco As Boolean)
    If blnNaoSetaFoco = False Then
        dcbDespesa.SetFocus
    End If
    dcbDespesa = ""
    mblnAlterandoDespesa = False
End Sub
Private Sub LimpaDadosAnulacao()
    txt_Descricao = ""
    txt_dblValor = ""
    txt_Descricao = "Anulação de Receita"
    txt_strCodigo = ""
    txt_intExercicioProcesso = ""
    txt_bitDigito = ""
End Sub

Private Sub SomaTotalAPagar()
    Dim dblValor As Double
    dblValor = Val(gstrConvVrParaSql(txtTotalEmpenho)) + _
               Val(gstrConvVrParaSql(txtTotalResto)) + _
               Val(gstrConvVrParaSql(txtTotalDespesa)) + _
               Val(gstrConvVrParaSql(txtTotalAnulacao))
    txtTotalAPagar = gstrConvVrDoSql(dblValor)
End Sub

Private Sub Totaliza(lvw_Lista As ListView, txtTotal As TextBox)
    Dim intInd      As Integer
    Dim dblTotal    As Double
    With lvw_Lista
        For intInd = 1 To .ListItems.Count
            Select Case UCase(lvw_Lista.Name)
            Case "LVW_CONTA"
                dblTotal = dblTotal + Val(gstrConvVrParaSql(.ListItems(intInd).SubItems(2)))
            Case "LVW_DESPESA"
                dblTotal = dblTotal + Val(gstrConvVrParaSql(.ListItems(intInd).SubItems(4)))
            Case "LVW_ANULACAORECEITA"
                dblTotal = dblTotal + Val(gstrConvVrParaSql(.ListItems(intInd).SubItems(1)))
            Case Else
                dblTotal = dblTotal + Val(gstrConvVrParaSql(.ListItems(intInd).SubItems(7)))
            End Select
        Next
    End With
    txtTotal = gstrConvVrDoSql(dblTotal)
End Sub

Private Sub ProcuraParcelaEmpenho()
    Dim itmEncontrado    As ListItem
    Set itmEncontrado = lvw_Empenho.FindItem(Trim$(dcbEmpenho) & Trim$(dcbParcela), 0, , 0)
    If gblnEncontroItemNoListView(lvw_Empenho, Trim$(dcbEmpenho) & Trim$(dcbParcela), lvwText) Then
        mblnAlterandoEmpenho = True
    Else
        mblnAlterandoEmpenho = False
    End If
End Sub

Private Sub ProcuraParcelaResto()
'    If gblnEncontroItemNoListView(lvw_Resto, gstrItemData(dcbResto) & gstrItemData(dcbParcelaResto), lvwText) Then
 If gblnEncontroItemNoListView(lvw_Resto, gstrItemData(dcbResto) & gstrItemData(dcbParcelaResto), lvwText) Then
        mblnAlterandoResto = True
    Else
        mblnAlterandoResto = False
    End If
End Sub

Private Sub ProcuraDespesa()
    If gblnEncontroItemNoListView(lvw_Despesa, dcbDespesa.BoundText, lvwTag) Then
        mblnAlterandoDespesa = True
    Else
        mblnAlterandoDespesa = False
    End If
End Sub

Private Sub VerificaLista()
    Select Case tab_3DPastaEmpenho.Tab
    Case 0
        IncluiAlteraListaEmpenho
        LeDaTabelaParaObj "", dcbEmpenho, strQueryEmpenho
        PreencheDados
    Case 1
        IncluiAlteraListaResto
        LeDaTabelaParaObj "", dcbResto, strQueryResto
    Case 2
        IncluiAlteraListaDespesa
        LeDaTabelaParaObj "", dcbDespesa, strQueryDespesa
    Case 3
        IncluiAlteraListaAnulacao
    End Select
End Sub

Private Sub ExcluiItemLista(lvw_Lista As ListView, txtTotal As TextBox)
    With lvw_Lista
        If .ListItems.Count > 0 Then
            .ListItems.Remove .SelectedItem.Index
        End If
        If .ListItems.Count > 1 Then
            intPKIDEmpenho = 0
        End If
        Totaliza lvw_Lista, txtTotal
    End With
    
End Sub

Private Sub VerificaListaExcluir()
    Select Case tab_3DPastaEmpenho.Tab
    Case 0
        ExcluiItemLista lvw_Empenho, txtTotalEmpenho
        LimpaDadosEmpenho
        LeDaTabelaParaObj "", dcbEmpenho, strQueryEmpenho
        PreencheDados
    Case 1
        ExcluiItemLista lvw_Resto, txtTotalResto
        LimpaDadosResto
        LeDaTabelaParaObj "", dcbResto, strQueryResto
    Case 2
        ExcluiItemLista lvw_Despesa, txtTotalDespesa
        LimpaDadosDespesa
        LeDaTabelaParaObj "", dcbDespesa, strQueryDespesa
    Case 3
        ExcluiItemLista lvw_AnulacaoReceita, txtTotalAnulacao
        LimpaDadosAnulacao
    End Select
End Sub

Private Function blnDadosEmpenhoOK() As Boolean
    Dim strSQL        As String
    Dim adoResultado  As New ADODB.Recordset
    Dim strTipoEmpenhoAtual As String
    
    If dcbEmpenho.MatchedWithList = False Then
        ExibeMensagem "O número do empenho tem que ser informado corretamente ou não existe(m) parcela(s) liquidadas para este empenho."
        dcbEmpenho.SetFocus
        Exit Function
    ElseIf dcbParcela.ListIndex = -1 Then
        ExibeMensagem "O número da parcela tem que ser informado corretamente."
        dcbParcela.SetFocus
        Exit Function
    ElseIf blnProcuraParcela = True Then
        ExibeMensagem "Número de parcela informado já se encontra cadastrado."
        Exit Function
'    ElseIf VerificaEmpenhoProcesso = False Then
'        ExibeMensagem "Este Empenho não pertence ao mesmo Processo."
'        Exit Function
    End If
    
    
    If lvw_Empenho.ListItems.Count > 0 Then
        strTipoEmpenhoAtual = Val(lvw_Empenho.ListItems(1).SubItems(9))
    Else
        blnDadosEmpenhoOK = True
        Exit Function
    End If
    
    
    strSQL = ""
    strSQL = strSQL & "SELECT E.intNumero , TE.bytAdiantamento FROM "
    strSQL = strSQL & gstrEmpenho & " E,"
    strSQL = strSQL & gstrTipoEmpenho & " TE"
    strSQL = strSQL & " WHERE E.Pkid = " & gstrItemData(dcbEmpenho)
    strSQL = strSQL & " AND TE.PKID = E.IntTipo"
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 60, adoResultado) Then
      With adoResultado
         If Not .EOF Then
            If adoResultado!bytAdiantamento <> CInt(strTipoEmpenhoAtual) Then
                ExibeMensagem "Não é possivel misturar mais de um tipo de Empenho em uma  mesma Ordem de Pagamento."
                blnDadosEmpenhoOK = False
                Exit Function
            Else
                blnDadosEmpenhoOK = True
            End If
         Else
            blnDadosEmpenhoOK = True
         End If
      End With
    End If
    
    'Incluido orc1571...
    If mblnAlterando Then
        strSQL = ""
        strSQL = strSQL & "SELECT intNumero, blnPago FROM "
        strSQL = strSQL & gstrOrdemPagamento & " OP"
        strSQL = strSQL & " WHERE Pkid = " & Me.txtPKId.Text
    
        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strSQL, 60, adoResultado) Then
            With adoResultado
            If Not .EOF Then
                If Val(gstrENulo(adoResultado.Fields("blnPago"))) = 1 Then
                    ExibeMensagem "Esta Ordem Pagamento já está Paga !"
                    blnDadosEmpenhoOK = False
                    LeDaTabelaParaObj "", tdb_Lista, strQuery
                    mblnClickOk = True
                    tdb_Lista_Click
                    Exit Function
                Else
                    blnDadosEmpenhoOK = True
                End If
            Else
                blnDadosEmpenhoOK = True
            End If
            End With
        End If
    End If
End Function

Private Function blnDadosDespesaOK() As Boolean
    Dim strSQL        As String
    Dim adoResultado  As New ADODB.Recordset
    
    If dcbDespesa.MatchedWithList = False Then
        ExibeMensagem "O número da despesa tem que ser informado corretamente."
        If dcbDespesa.Enabled Then dcbDespesa.SetFocus
        Exit Function
    End If
    
    
    If lvw_Despesa.ListItems.Count > 0 Then
    
        strSQL = ""
        strSQL = strSQL & "Select intContaContabil intContaContabilGrid ,"
        strSQL = strSQL & "(Select intContaContabil FROM " & gstrDespesaExtraOrcamentaria & " WHERE PKID = " & dcbDespesa.BoundText & ") intContaContabilCombo "
        strSQL = strSQL & " FROM " & gstrDespesaExtraOrcamentaria & " WHERE PKID = " & lvw_Despesa.ListItems(1).Tag

        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strSQL, 60, adoResultado) Then
           With adoResultado
              If Not .EOF Then
                 If !intContaContabilGrid <> !intContaContabilCombo Then
                     ExibeMensagem "Não é possivel inserir itens de Despesa Extra-Orçamentária com contas diferentes."
                     If dcbDespesa.Enabled Then dcbDespesa.SetFocus
                     Exit Function
                 End If
              End If
           End With
        End If
    End If

    blnDadosDespesaOK = True
    
End Function
Private Function blnDadosAnulacaoOk() As Boolean
    If Len(Trim(txt_Descricao)) = 0 Then
        ExibeMensagem "Favor informar descricao da Anulação de Receita."
        txt_Descricao.SetFocus
    ElseIf Len(Trim(txt_dblValor)) = 0 Then
        ExibeMensagem "Favor informar valor da Anulação de Receita."
        txt_Descricao.SetFocus
    ElseIf CDbl(Trim(txt_dblValor)) = 0 Then
        ExibeMensagem "Anulação de Receita não pode ter valor zero."
        txt_Descricao.SetFocus
    ElseIf Len(Trim(txt_strCodigo)) > 0 Or Len(Trim(txt_bitDigito)) > 0 Or Len(Trim(txt_intExercicioProcesso)) > 0 Then
       If Not VerificaAnulacaoReceitaProcesso Then
          ExibeMensagem "Processo não localizado."
          Exit Function
       End If
       blnDadosAnulacaoOk = True
    Else
        blnDadosAnulacaoOk = True
    End If
End Function

Private Function blnDadosRestoOk() As Boolean

    Dim strSQL        As String
    Dim adoResultado  As New ADODB.Recordset
    Dim strTipoEmpenhoAtual As String

    If dcbResto.MatchedWithList = False Then
        ExibeMensagem "O número do resto tem que ser informado corretamente."
        dcbResto.SetFocus
    ElseIf dcbParcelaResto.ListIndex = -1 Then
        ExibeMensagem "O número da parcela tem que ser informado corretamente."
        dcbParcelaResto.SetFocus
    Else
        blnDadosRestoOk = True
    End If
    
    
    If lvw_Resto.ListItems.Count > 0 Then
        strTipoEmpenhoAtual = Val(lvw_Resto.ListItems(1).SubItems(9))
    Else
        blnDadosRestoOk = True
        Exit Function
    End If

    strSQL = ""
    strSQL = strSQL & "SELECT E.intNumero , TE.bytAdiantamento FROM "
    strSQL = strSQL & gstrEmpenho & " E,"
    strSQL = strSQL & gstrTipoEmpenho & " TE"
    strSQL = strSQL & " WHERE E.Pkid = " & gstrItemData(dcbResto)
    strSQL = strSQL & " AND TE.PKID = E.IntTipo"

   Set gobjBanco = New clsBanco
   If gobjBanco.CriaADO(strSQL, 60, adoResultado) Then
      With adoResultado
         If Not .EOF Then
            If adoResultado!bytAdiantamento <> CInt(strTipoEmpenhoAtual) Then
                blnDadosRestoOk = False
                ExibeMensagem "Não é possivel misturar mais de um tipo de Resto a Pagar em uma  mesma Ordem de Pagamento."
            Else
                blnDadosRestoOk = True
            End If
         Else
            blnDadosRestoOk = True
         End If
      End With
   End If
End Function

Private Sub IncluiAlteraListaEmpenho()

'******************************************************************************************
' Data: 09/06/2003
' Alteração: - Substituição da chamada direta à stored procedure pela função
'            gstrStoredProcedure.
' Responsável: Everton Bianchini
'******************************************************************************************

Dim strSQL          As String
Dim adoResultado    As ADODB.Recordset
    If blnDadosEmpenhoOK() Then
'        strSql = "sp_EmpenhoParaPagar " & Val(dcbParcela.BoundText)
        strSQL = gstrStoredProcedure("sp_EmpenhoParaPagar", CStr(Val(gstrItemData(dcbParcela))), True)
        ProcuraParcelaEmpenho
        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
            With adoResultado
                If .EOF = False Then
                    If mblnAlterandoEmpenho Then
                        lvw_Empenho.ListItems(lvw_Empenho.SelectedItem.Index).Text = Trim$(!INTEMPENHO) & Trim$(!INTNUMERO)
                        lvw_Empenho.SelectedItem.SubItems(1) = !INTEMPENHO
                        lvw_Empenho.SelectedItem.SubItems(2) = !INTNUMERO
                        lvw_Empenho.SelectedItem.SubItems(3) = gstrDataFormatada(!dtmPrevisao)
                        lvw_Empenho.SelectedItem.SubItems(4) = gstrDataFormatada(!dtmLiquidacao)
                        lvw_Empenho.SelectedItem.SubItems(5) = gstrConvVrDoSql(!dblValor)
                        lvw_Empenho.SelectedItem.SubItems(6) = gstrConvVrDoSql(!dblDesconto)
                        lvw_Empenho.SelectedItem.SubItems(7) = gstrConvVrDoSql(!dblLiquido)
                        lvw_Empenho.SelectedItem.SubItems(9) = !bytAdiantamento
                        If Trim(txtHistorico.Text) = "" Then txtHistorico.Text = gstrENulo(!STRHISTORICO)
                    Else
                        Set mobjLista = lvw_Empenho.ListItems.Add(, , Trim$(!INTEMPENHO) & Trim$(!INTNUMERO))
                        mobjLista.SubItems(1) = !INTEMPENHO
                        mobjLista.SubItems(2) = !INTNUMERO
                        mobjLista.SubItems(3) = gstrDataFormatada(!dtmPrevisao)
                        mobjLista.SubItems(4) = gstrDataFormatada(!dtmLiquidacao)
                        mobjLista.SubItems(5) = gstrConvVrDoSql(!dblValor)
                        mobjLista.SubItems(6) = gstrConvVrDoSql(!dblDesconto)
                        mobjLista.SubItems(7) = gstrConvVrDoSql(!dblLiquido)
                        mobjLista.SubItems(9) = gstrENulo(!bytAdiantamento)
                        mobjLista.Tag = !Pkid
                        If Trim(txtHistorico.Text) = "" Then txtHistorico.Text = gstrENulo(!STRHISTORICO)
                    End If
                End If
                LimpaDadosEmpenho
                Totaliza lvw_Empenho, txtTotalEmpenho
            End With
        End If
    End If
End Sub

Private Sub IncluiAlteraListaResto()

'******************************************************************************************
' Data: 09/06/2003
' Alteração: - Substituição da chamada direta à stored procedure pela função
'            gstrStoredProcedure.
' Responsável: Everton Bianchini
'******************************************************************************************

    Dim strSQL          As String
    Dim adoResultado    As ADODB.Recordset
    If blnDadosRestoOk() Then
        'strSql = "sp_RestoParaPagar " & Val(dcbParcelaResto.BoundText)
        'strSQL = gstrStoredProcedure("sp_RestoParaPagar", CStr(Val(dcbParcelaResto.BoundText)), True)
        
        strSQL = ""
        strSQL = "SELECT SE.Pkid Pkid, E.intExercicioRP IntExercicio,"
        strSQL = strSQL & "E.IntNumero intResto, "
        strSQL = strSQL & "SE.IntNumero intNumero, "
        strSQL = strSQL & "SE.dblvalor dblValor, "
        strSQL = strSQL & "SE.dtmData dtmPrevisao, "
        strSQL = strSQL & "E.strHistorico, "
        strSQL = strSQL & gstrISNULL("SE.dblDesconto", "0") & " + " & gstrISNULL("SUM(SL.dblvalor)", "0") & " dblDesconto, "
        strSQL = strSQL & "(" & gstrISNULL("SE.dblvalor", "0") & " - "
        strSQL = strSQL & "(" & gstrISNULL("SE.dblDesconto", "0") & " + " & gstrISNULL("SUM(Sl.dblValor)", "0") & ")) dblLiquido, "
        strSQL = strSQL & gstrISNULL("TE.bytAdiantamento", "0") & " bytAdiantamento "
        strSQL = strSQL & " FROM " & gstrEmpenho & " E ,"
        strSQL = strSQL & gstrSubempenho & " SE, " & gstrSubempenhoLiquidado & " SL ,"
        strSQL = strSQL & gstrTipoEmpenho & " TE"
        strSQL = strSQL & " WHERE "
        strSQL = strSQL & "SE.Pkid =" & gstrItemData(dcbParcelaResto) & " AND "
        strSQL = strSQL & "E.Pkid = SE.intEmpenho AND "
        strSQL = strSQL & "E.intTipo = TE.PKID AND "
        strSQL = strSQL & "SE.Pkid " & strOUTJSQLServer & "= SL.intParcela" & strOUTJOracle & " And "
        
        '(bytDBType = EDatabases.SQLServer)
        'If EDatabases.Oracle Then
        '    strSQL = strSQL & "SL.intParcela = SE.Pkid " & strOUTJSQLServer & " And "
        'ElseIf EDatabases.SQLServer Then
        '    strSQL = strSQL & "SL.intParcela =" & strOUTJSQLServer & " SE.Pkid AND "
        'End If
        strSQL = strSQL & "E.intExercicioRP =" & CStr(gintExercicio)
        strSQL = strSQL & " GROUP BY SE.Pkid, E.intExercicioRP, E.IntNumero,E.strHistorico,"
        strSQL = strSQL & " SE.IntNumero, SE.dblvalor,SE.dtmData, TE.bytAdiantamento ,SE.dblDesconto"
        
        ProcuraParcelaResto
        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
            With adoResultado
                If .EOF = False Then
                    If mblnAlterandoResto Then
'                        lvw_Resto.ListItems(lvw_Resto.SelectedItem.Index).Text = _
                                            gstrItemData(dcbResto) & gstrItemData(dcbParcelaResto)
                        lvw_Resto.ListItems(lvw_Resto.SelectedItem.Index).Text = _
                                            Trim$(!intResto) & Trim$(!INTNUMERO)
                        lvw_Resto.SelectedItem.SubItems(1) = !intExercicio
                        lvw_Resto.SelectedItem.SubItems(2) = !intResto
                        lvw_Resto.SelectedItem.SubItems(3) = !INTNUMERO
                        lvw_Resto.SelectedItem.SubItems(4) = gstrDataFormatada(!dtmPrevisao)
                        lvw_Resto.SelectedItem.SubItems(5) = gstrConvVrDoSql(!dblValor)
                        lvw_Resto.SelectedItem.SubItems(6) = gstrConvVrDoSql(!dblDesconto)
                        lvw_Resto.SelectedItem.SubItems(7) = gstrConvVrDoSql(!dblLiquido)
                        lvw_Resto.SelectedItem.SubItems(9) = !bytAdiantamento
                        lvw_Resto.SelectedItem.Tag = !Pkid
                        If Trim(txtHistorico.Text) = "" Then txtHistorico.Text = gstrENulo(!STRHISTORICO)
                    Else
                        Set mobjLista = lvw_Resto.ListItems.Add(, , gstrItemData(dcbResto) & _
                                                                    gstrItemData(dcbParcelaResto))
                        mobjLista.SubItems(1) = !intExercicio
                        mobjLista.SubItems(2) = !intResto
                        mobjLista.SubItems(3) = !INTNUMERO
                        mobjLista.SubItems(4) = gstrDataFormatada(!dtmPrevisao)
                        mobjLista.SubItems(5) = gstrConvVrDoSql(!dblValor)
                        mobjLista.SubItems(6) = gstrConvVrDoSql(!dblDesconto)
                        mobjLista.SubItems(7) = gstrConvVrDoSql(!dblLiquido)
                        mobjLista.SubItems(9) = !bytAdiantamento
                        mobjLista.Tag = !Pkid
                        If Trim(txtHistorico.Text) = "" Then txtHistorico.Text = gstrENulo(!STRHISTORICO)
                    End If
                End If
                LimpaDadosResto
                Totaliza lvw_Resto, txtTotalResto
            End With
        End If
    End If
End Sub
Private Sub IncluiAlteraListaDespesa()
    Dim strSQL          As String
    Dim adoResultado    As ADODB.Recordset
    If blnDadosDespesaOK() Then
        strSQL = ""
        strSQL = strSQL & "SELECT DP.PKId, DP.intNumero, DP.dtmData, DP.strHistorico, "
        strSQL = strSQL & "DP.dblValor, DP.dblDesconto,(DP.dblValor - " & gstrISNULL("DP.dblDesconto", "0") & ") dblLiquido, CT.strNome FROM "
        strSQL = strSQL & gstrDespesaExtraOrcamentaria & " DP, "
        strSQL = strSQL & gstrContribuinte & " CT "
        strSQL = strSQL & "WHERE DP.intContribuinte = CT.PKId "
        strSQL = strSQL & "AND  DP.PKId = " & Val(dcbDespesa.BoundText) & " "
        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
            With adoResultado
                If .EOF = False Then
                    If mblnAlterandoDespesa Then
                        lvw_Despesa.ListItems(lvw_Despesa.SelectedItem.Index).Text = Trim$(!INTNUMERO)
                        lvw_Despesa.SelectedItem.SubItems(1) = gstrDataFormatada(!DTMDATA)
                        lvw_Despesa.SelectedItem.SubItems(2) = gstrConvVrDoSql(!dblValor)
                        lvw_Despesa.SelectedItem.SubItems(3) = gstrConvVrDoSql(gstrENulo(!dblDesconto))
                        lvw_Despesa.SelectedItem.SubItems(4) = gstrConvVrDoSql(gstrENulo(!dblLiquido))
                        lvw_Despesa.SelectedItem.SubItems(5) = gstrENulo(!STRNOME)
                        lvw_Despesa.SelectedItem.Tag = !Pkid
                        If Trim(txtHistorico.Text) = "" Then txtHistorico.Text = gstrENulo(!STRHISTORICO)
                    Else
                        Set mobjLista = lvw_Despesa.ListItems.Add(, , !INTNUMERO)
                        mobjLista.SubItems(1) = gstrDataFormatada(!DTMDATA)
                        mobjLista.SubItems(2) = gstrConvVrDoSql(!dblValor)
                        mobjLista.SubItems(3) = gstrConvVrDoSql(gstrENulo(!dblDesconto))
                        mobjLista.SubItems(4) = gstrConvVrDoSql(gstrENulo(!dblLiquido))
                        mobjLista.SubItems(5) = gstrENulo(!STRNOME)
                        If Trim(txtHistorico.Text) = "" Then txtHistorico.Text = gstrENulo(!STRHISTORICO)
                        mobjLista.Tag = !Pkid
                    End If
                End If
                Totaliza lvw_Despesa, txtTotalDespesa
                LimpaDadosDespesa
            End With
        End If
    End If
End Sub
Private Sub IncluiAlteraListaAnulacao()
    If blnDadosAnulacaoOk() Then
       Set mobjLista = lvw_AnulacaoReceita.ListItems.Add(, , txt_Descricao)
       mobjLista.SubItems(1) = gstrConvVrDoSql(txt_dblValor)
       
       If Len(Trim(txt_strCodigo)) > 0 Then
          mobjLista.SubItems(2) = txt_strCodigo & "/" & txt_intExercicioProcesso & "-" & txt_bitDigito
       End If
       
       Totaliza lvw_AnulacaoReceita, txtTotalAnulacao
       LimpaDadosAnulacao
    End If
End Sub

Private Function strSqlParcelaEmpenho() As String
    Dim strSQL  As String
    strSQL = ""
    strSQL = strSQL & "SELECT PKId, intNumero "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrSubempenho & " "
    strSQL = strSQL & "WHERE bytSituacao = 2 "
    strSQL = strSQL & "AND intEmpenho = " & Val(dcbEmpenho.BoundText)
    'strSql = strSql & "AND NOT PKID IN (SELECT IntParcela FROM " & gstrOrdemPagamentoEmpenho & " ) "
    
    strSQL = strSQL & " AND NOT PKID IN "
    
    strSQL = strSQL & "(SELECT intParcela FROM "
    strSQL = strSQL & gstrOrdemPagamentoEmpenho & " OPE, "
    strSQL = strSQL & gstrOrdemPagamento & " OP "
    strSQL = strSQL & "WHERE OPE.intOrdemPagamento = OP.Pkid "
    strSQL = strSQL & "AND (OP.Bytcancelado = 0 or OP.Bytcancelado is null) "
    strSQL = strSQL & "GROUP BY intParcela) "
    
    If Len(gstrPkidSubEmpnoGrid) > 0 Then
        strSQL = strSQL & " AND NOT PKID IN "
        strSQL = strSQL & "(" & gstrPkidSubEmpnoGrid & ") "
    End If
    
    
    
    strSQL = strSQL & " ORDER BY intNumero"
    strSqlParcelaEmpenho = strSQL
End Function

Private Function strSqlParcelaResto() As String
    Dim strSQL  As String
    
    strSQL = ""
    strSQL = strSQL & "SELECT PKId, intNumero "
    strSQL = strSQL & "FROM " & gstrSubempenho & " "
    strSQL = strSQL & "WHERE intEmpenho = " & CStr(gstrItemData(dcbResto, True))
    strSQL = strSQL & " AND bytSituacao = 2"
    'strSql = strSql & " AND NOT PKID IN (SELECT IntParcela FROM " & gstrOrdemPagamentoResto & " ) "
    
    strSQL = strSQL & " AND NOT PKID IN "
    
    strSQL = strSQL & "(SELECT intParcela FROM "
    strSQL = strSQL & gstrOrdemPagamentoResto & " OPE, "
    strSQL = strSQL & gstrOrdemPagamento & " OP "
    strSQL = strSQL & "WHERE OPE.intOrdemPagamento = OP.Pkid "
    strSQL = strSQL & "AND (OP.Bytcancelado = 0 or OP.Bytcancelado is null) "
    strSQL = strSQL & "GROUP BY intParcela) "
    
    If Len(gstrPkidRestonoGrid) > 0 Then
        strSQL = strSQL & " AND NOT PKID IN "
        strSQL = strSQL & "(" & gstrPkidRestonoGrid & ") "
    End If
    
    
    strSQL = strSQL & " ORDER BY intNumero"
    
    strSqlParcelaResto = strSQL
End Function

Private Function strQueryLancamento() As String
    Dim strSQL  As String
    strSQL = ""
    strSQL = strSQL & "SELECT PC.PKId, PC.strContaContabil, PC.strDescricao, "
    strSQL = strSQL & "LC.dblValor, LC.strDocumento FROM "
    strSQL = strSQL & gstrLancamentoContabil & " LC, "
    strSQL = strSQL & gstrPlanoConta & " PC "
    strSQL = strSQL & "WHERE LC.intConta = PC.PKId "
    strSQL = strSQL & "AND LC.intProcesso = " & Val(txtPKId) & " "
    strSQL = strSQL & "ORDER BY PC.strContaContabil"
    strQueryLancamento = strSQL
End Function


Private Function strQueryEmpenho() As String
    Dim strSQL  As String
    strSQL = ""
    strSQL = strSQL & "SELECT DISTINCT EP.PKId, EP.intNumero, EP.strcodigo, EP.bitdigito, EP.intexercicio, EP.strHistorico "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrEmpenho & " EP, "
    strSQL = strSQL & gstrSubempenho & " SE, "
    strSQL = strSQL & gstrProgramaDeTrabalho & " PT "
    strSQL = strSQL & "WHERE EP.PKId = SE.intEmpenho "
    strSQL = strSQL & "AND SE.bytSituacao = 2 "
    strSQL = strSQL & "AND " & gstrISNULL("EP.intExercicioRP", "0") & " = 0 "
    
    strSQL = strSQL & "AND NOT SE.PKID IN "
    
    strSQL = strSQL & "(SELECT intParcela FROM "
    strSQL = strSQL & gstrOrdemPagamentoEmpenho & " OPE, "
    strSQL = strSQL & gstrOrdemPagamento & " OP "
    strSQL = strSQL & "WHERE OPE.intOrdemPagamento = OP.Pkid "
    strSQL = strSQL & "AND (OP.Bytcancelado = 0 or OP.Bytcancelado is null) "
    strSQL = strSQL & "GROUP BY intParcela) "
    
    If Len(gstrPkidSubEmpnoGrid) > 0 Then
        strSQL = strSQL & " AND NOT SE.PKID IN "
        strSQL = strSQL & "(" & gstrPkidSubEmpnoGrid & ") "
    End If
    
    strSQL = strSQL & "AND PT.PKID = EP.intProgramaTrabalho "
    
    If gstrItemData(dbcintCredor) > 0 Then
       strSQL = strSQL & "AND EP.intCredor = " & gstrItemData(dbcintCredor)
    End If
    
    strSQL = strSQL & " AND PT.intExercicio = " & CStr(gintExercicio)
    
    strSQL = strSQL & " ORDER BY EP.intNumero"
    strQueryEmpenho = strSQL
End Function

Private Function gstrPkidSubEmpnoGrid() As String
    Dim i As Integer
    For i = 1 To lvw_Empenho.ListItems.Count
        gstrPkidSubEmpnoGrid = gstrPkidSubEmpnoGrid & lvw_Empenho.ListItems(i).Tag & ","
    Next
    If Len(gstrPkidSubEmpnoGrid) > 0 Then
        gstrPkidSubEmpnoGrid = Mid(gstrPkidSubEmpnoGrid, 1, Len(gstrPkidSubEmpnoGrid) - 1)
    End If
    
End Function


Private Function strQueryDespesa() As String
    Dim strSQL          As String
    strSQL = ""
    strSQL = strSQL & "SELECT DP.PKId, DP.intNumero "
    strSQL = strSQL & "FROM " & gstrDespesaExtraOrcamentaria & " DP "
    strSQL = strSQL & "WHERE bytSituacao = 0 " 'Programada
    strSQL = strSQL & "AND NOT DP.PKID IN "
    
    strSQL = strSQL & "(SELECT intDespesaExtraOrcamentaria FROM "
    strSQL = strSQL & gstrOrdemPagamentoDespesaExtra & " OPD, "
    strSQL = strSQL & gstrOrdemPagamento & " OP "
    strSQL = strSQL & "WHERE OPD.intOrdemPagamento = OP.Pkid "
    strSQL = strSQL & "AND (OP.Bytcancelado = 0 or OP.Bytcancelado is null) "
    strSQL = strSQL & "GROUP BY intDespesaExtraOrcamentaria) "
    
    If gstrItemData(dbcintCredor) > 0 Then
       strSQL = strSQL & "AND DP.intContribuinte = " & gstrItemData(dbcintCredor)
    End If
    
    If Len(gstrPkidDespesaExtranoGrid) > 0 Then
        strSQL = strSQL & " AND NOT DP.PKID IN "
        strSQL = strSQL & "(" & gstrPkidDespesaExtranoGrid & ") "
    End If
    
    strSQL = strSQL & " ORDER BY DP.intNumero"
    strQueryDespesa = strSQL
End Function

Private Function gstrPkidDespesaExtranoGrid() As String
    Dim i As Integer
    For i = 1 To lvw_Despesa.ListItems.Count
        gstrPkidDespesaExtranoGrid = gstrPkidDespesaExtranoGrid & lvw_Despesa.ListItems(i).Tag & ","
    Next
    If Len(gstrPkidDespesaExtranoGrid) > 0 Then
        gstrPkidDespesaExtranoGrid = Mid(gstrPkidDespesaExtranoGrid, 1, Len(gstrPkidDespesaExtranoGrid) - 1)
    End If
    
End Function

Private Function strQueryResto()
    Dim strSQL  As String
    strSQL = ""
    strSQL = strSQL & "SELECT DISTINCT "
    If bytDBType = Oracle Then
        strSQL = strSQL & " EP.PKId, EP.intNumero ||'/'|| to_char(EP.dtmdata,'yy') as intnumero, "
    Else
        strSQL = strSQL & " EP.PKId, " & gstrCONVERT(CDT_VARCHAR, "EP.intNumero") & " + '/' + right(year(EP.dtmdata),2) as intnumero, "
    End If
    strSQL = strSQL & "EP.dtmData,EP.dblValor ,PT.strCodigo "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrEmpenho & " EP, "
    strSQL = strSQL & gstrSubempenho & " SE, "
    strSQL = strSQL & gstrProgramaDeTrabalho & " PT "
    strSQL = strSQL & "WHERE EP.intProgramaTrabalho = PT.PKId "
    strSQL = strSQL & "AND EP.PKId = SE.intEmpenho "
    strSQL = strSQL & "AND SE.bytSituacao = 2"
    strSQL = strSQL & " AND intExercicioRP = " & gintExercicio
    
    If gstrItemData(dbcintCredor) > 0 Then
       strSQL = strSQL & " AND EP.intCredor = " & gstrItemData(dbcintCredor)
    End If
    
    strSQL = strSQL & " AND NOT SE.PKID IN "
    
    strSQL = strSQL & "(SELECT intParcela FROM "
    strSQL = strSQL & gstrOrdemPagamentoResto & " OPE, "
    strSQL = strSQL & gstrOrdemPagamento & " OP "
    strSQL = strSQL & "WHERE OPE.intOrdemPagamento = OP.Pkid "
    strSQL = strSQL & "AND (OP.Bytcancelado = 0 or OP.Bytcancelado is null) "
    strSQL = strSQL & "GROUP BY intParcela) "
    
    If Len(gstrPkidRestonoGrid) > 0 Then
        strSQL = strSQL & " AND NOT SE.PKID IN "
        strSQL = strSQL & "(" & gstrPkidRestonoGrid & ") "
    End If
    
    strSQL = strSQL & " ORDER BY intNumero, EP.dtmData, PT.strCodigo "
    strQueryResto = strSQL
End Function

Private Function gstrPkidRestonoGrid() As String
    Dim i As Integer
    For i = 1 To lvw_Resto.ListItems.Count
        gstrPkidRestonoGrid = gstrPkidSubEmpnoGrid & lvw_Resto.ListItems(i).Tag & ","
    Next
    If Len(gstrPkidRestonoGrid) > 0 Then
        gstrPkidRestonoGrid = Mid(gstrPkidRestonoGrid, 1, Len(gstrPkidRestonoGrid) - 1)
    End If
    
End Function


Private Function strQueryPlanoConta() As String
    Dim strSQL  As String
    strSQL = ""
    strSQL = strSQL & "SELECT PKId, strDescricao "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrPlanoConta & " "
    strSQL = strSQL & "WHERE ABS(blnFinanceira) = 1 "
    strSQL = strSQL & "AND ABS(blnAnalitica) = 1 "
    strQueryPlanoConta = strSQL
End Function

Private Sub cbo_HistoricoLiquidacao_Change()
    txtHistorico = Trim(cbo_HistoricoLiquidacao)
End Sub

Private Sub cbo_HistoricoLiquidacao_Click()
    Dim adoResultado As ADODB.Recordset
    cbo_HistoricoLiquidacao_Change
    Set gobjBanco = New clsBanco
    gobjBanco.CriaADO "SELECT h.strcodigo FROM tblhistorico H WHERE H.STRdescricao = '" & Me.cbo_HistoricoLiquidacao.Text & "'", 10, adoResultado
    With adoResultado
        If Not .EOF Then
            Me.txt_CodHistorico.Text = gstrENulo(!strCodigo)
            Me.txtHistorico.Text = Me.cbo_HistoricoLiquidacao.Text
        Else
            Me.txtHistorico.Text = ""
            Me.cbo_HistoricoLiquidacao.Text = ""
        End If
    End With
End Sub

Private Sub cmd_Credor_Click()
    mblnTelaCredor = True
    CarregaForm frmCadContribuinte, dbcintCredor
    frmCadContribuinte.Caption = "Cadastro de Credores"
    frmCadContribuinte.Tag = "Credor"

End Sub

Private Sub cmd_Despesa_Click()
    CarregaForm frmCadDespesaExtraOrcamentaria, dcbDespesa, strQueryDespesa
    With frmCadDespesaExtraOrcamentaria
        .txtintNumero = dcbDespesa.Text
        .MantemForm gstrLocalizar
        .SelecionaDespesaExtra
    End With
End Sub

Private Sub cmd_Empenho_Click()

    frmCadEmpenho.mblnRestosAPagar = False
    CarregaForm frmCadEmpenho, dcbEmpenho, strQueryEmpenho
            
        If Trim(dcbEmpenho.Text) <> "" Then
            With frmCadEmpenho
                .txtintNumero = dcbEmpenho.Text
                .MantemForm gstrLocalizar
                .SelecionaLiquidacao
                .txt_DataLiuidacao = txtdtmData.Text
            End With
        Else
            frmCadEmpenho.tab_3dPasta.Tab = 0
        End If
End Sub

Private Function strQueryRestosPagar()
   Dim strSQL          As String
   Dim strCombo       As String
   
   strCombo = "0"
   If InStr(dcbResto.Text, "/") <> 0 Then
        strCombo = Mid(dcbResto.Text, InStr(dcbResto.Text, "/") + 1)
        If IsNumeric(strCombo) Then
            strCombo = Mid(dcbResto.Text, 1, InStr(dcbResto.Text, "/")) & Year("01/01/" & strCombo)
        End If
   End If
   
    strSQL = ""
    strSQL = strSQL & "SELECT EP.PKId, EP.intNumero , EP.dtmData "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrEmpenho & " EP "
    strSQL = strSQL & "WHERE " & gstrCONVERT(CDT_NVARCHAR, "EP.intnumero") & strCONCAT & "'/'" & strCONCAT & gstrCONVERT(CDT_NVARCHAR, gstrDATEPART(strYEAR, "EP.dtmdata")) & "  = '" & strCombo & "'"
   
   strQueryRestosPagar = strSQL
   
End Function


Private Sub cmd_HistoricoLiquidacao_Click()
    CarregaForm frmCadHistorico, cbo_HistoricoLiquidacao
End Sub

Private Sub cmd_Resto_Click()
Dim adoResultado    As ADODB.Recordset
    
   frmCadEmpenho.mblnRestosAPagar = True
   CarregaForm frmCadEmpenho, dcbResto, strQueryRestosPagar
        
   Set gobjBanco = New clsBanco
   gobjBanco.CriaADO strQueryRestosPagar, 5, adoResultado
        
   If Trim(dcbResto.Text) <> "" And adoResultado.RecordCount > 0 Then
       With frmCadEmpenho
           .txtintNumero = adoResultado!INTNUMERO
           .txtdtmData = adoResultado!DTMDATA
           .MantemForm gstrLocalizar
           .SelecionaLiquidacao
           'orc1376
           'TrocaCorObjeto .txt_intExercicioRP, True
           TrocaCorObjeto .txtintExercicioEmpenho, True
       End With
   Else
       frmCadEmpenho.tab_3dPasta.Tab = 0
   End If

End Sub

Private Sub dbcintCredor_Change()
    If mblnTelaCredor Then
        mblnTelaCredor = False
        dbcintCredor_Click 0
    End If
End Sub

Private Sub dbcintCredor_Click(Area As Integer)
    
        DropDownDataCombo dbcintCredor, Me, Area
    
        If dbcintCredor.MatchedWithList Then
            
            txt_intNContribuinte = LeCDCCredor(dbcintCredor.BoundText)
       
            If itemAnterior = dbcintCredor.BoundText Then Exit Sub
       
            If optbytTipo(0).Value Then
                dcbEmpenho.Text = ""
                lvw_Empenho.ListItems.Clear
            ElseIf optbytTipo(1).Value Then
                dcbResto.Text = ""
                lvw_Resto.ListItems.Clear
            ElseIf optbytTipo(2).Value Then
                dcbDespesa.Text = ""
                lvw_Despesa.ListItems.Clear
            ElseIf optbytTipo(3).Value Then
                'lvw_AnulacaoReceita.ListItems.Clear
            End If
            If optbytTipo(0).Value Then
                LeDaTabelaParaObj "", dcbEmpenho, strQueryEmpenho
                PreencheDados
            ElseIf optbytTipo(1).Value Then
                LeDaTabelaParaObj "", dcbResto, strQueryResto
            ElseIf optbytTipo(2).Value Then
                LeDaTabelaParaObj "", dcbDespesa, strQueryDespesa
            End If
    
            itemAnterior = dbcintCredor.BoundText
       
       End If
   
End Sub


Private Sub dbcintCredor_GotFocus()
    If dbcintCredor.BoundText <> txt_intNContribuinte And dbcintCredor.BoundText <> "" Then
         'itemAnterior = ""
         dbcintCredor_Click 0
    End If
End Sub

Private Sub dbcintCredor_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintCredor, Me, , KeyCode, Shift
End Sub

Private Sub dbcintCredor_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Sub dcbEmpenho_Change()
   If gstrItemData(dcbEmpenho) <> 0 Then
        txtFonteDeRecursoEmpenho.Text = leFontedeRecurso(gstrItemData(dcbEmpenho))
   Else
        txtFonteDeRecursoEmpenho.Text = ""
   End If
End Sub

Private Sub dcbEmpenho_LostFocus()
dcbEmpenho_Click (2)
End Sub

Private Sub dcbResto_Change()
   If gstrItemData(dcbResto) <> 0 Then
        txtFonteDeRecursoResto.Text = leFontedeRecurso(gstrItemData(dcbResto))
   Else
        txtFonteDeRecursoResto.Text = ""
   End If
End Sub

Private Sub dcbResto_LostFocus()
    dcbResto_Click (2)
End Sub

Private Sub tdb_Lista_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
   If ColIndex = 6 Then
      Value = gstrConvVrDoSql(Value)
   End If
End Sub

Private Sub tdb_Lista_HeadClick(ByVal ColIndex As Integer)

gOrdenaGrid tdb_Lista, ColIndex

End Sub

Private Sub txt_CodHistorico_GotFocus()
    MarcaCampo txt_CodHistorico
End Sub

Private Sub txt_CodHistorico_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_CodHistorico
End Sub

Private Sub txt_CodHistorico_LostFocus()
    Dim adoResultado As ADODB.Recordset
    Set gobjBanco = New clsBanco
    gobjBanco.CriaADO "SELECT h.StrDescricao FROM tblhistorico H WHERE H.STRCODIGO = '" & Me.txt_CodHistorico.Text & "'", 10, adoResultado
    With adoResultado
        If Not .EOF Then
            Me.cbo_HistoricoLiquidacao.Text = gstrENulo(!strDescricao)
            Me.txtHistorico.Text = gstrENulo(!strDescricao)
        Else
            Me.txtHistorico.Text = ""
            Me.cbo_HistoricoLiquidacao.Text = ""
        End If
    End With
End Sub

Private Sub txt_dblValor_GotFocus()
    MarcaCampo txt_dblValor
End Sub

Private Sub txt_DBLVALOR_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txt_dblValor
End Sub

Private Sub txt_DBLVALOR_LostFocus()
    txt_dblValor = gstrConvVrDoSql(txt_dblValor)
End Sub

Private Sub txt_Descricao_KeyPress(KeyAscii As Integer)
   CaracterValido KeyAscii
End Sub

Private Sub txt_intNContribuinte_KeyPress(KeyAscii As Integer)
   CaracterValido KeyAscii, "N", txt_intNContribuinte
End Sub

Private Sub txt_intNContribuinte_LostFocus()
'Cláudio
Dim strPKId As String

    strPKId = LeCDCCredor(, txt_intNContribuinte)
   
   If strPKId = "" Then
        dbcintCredor.BoundText = ""
        Exit Sub
   End If

   
    If Len(Trim(txt_intNContribuinte)) > 0 Then
        If dbcintCredor.Enabled Then dbcintCredor.SetFocus
        
        Filtrar_dbcintCredor strPKId
        
    End If
  
End Sub
Private Sub dcbDespesa_Change()
    ProcuraDespesa
End Sub

Private Sub dcbDespesa_Click(Area As Integer)
    Dim strPKId As String
    If dcbDespesa.BoundText = "" Then Exit Sub
       strPKId = dcbDespesa.BoundText
    If dcbDespesa.Text <> "" Then
        ProcuraDespesa
        If gstrItemData(dbcintCredor) = 0 Then
           PreencherListaDeOpcoes dbcintCredor, dblQueryCredorDespesa(Val(strPKId))
           txt_intNContribuinte = LeCDCCredor(dbcintCredor.BoundText)
        Else
           LeDaTabelaParaObj "", dcbDespesa, strQueryDespesa
           dcbDespesa.BoundText = strPKId
        End If
    End If
End Sub

Private Sub dcbDespesa_GotFocus()
    
    AtivaPastaDeObjeto tab_3DPastaEmpenho, 2
    
    LeDaTabelaParaObj "", dcbDespesa, strQueryDespesa
    
End Sub

Private Sub dcbEmpenho_Click(Area As Integer)
   Dim strPKId    As String
   Dim strTexto   As String
    
    If dcbEmpenho.BoundText = "" Then Exit Sub
       strPKId = dcbEmpenho.BoundText
    If dcbEmpenho.Text <> "" Then
        strTexto = dcbEmpenho.Text
        LeDaTabelaParaObj "", dcbParcela, strSqlParcelaEmpenho
        PreencheProcesso Val(strPKId), False
        If gstrItemData(dbcintCredor) = 0 Then
           PreencherListaDeOpcoes dbcintCredor, dblQueryCredor
           txt_intNContribuinte = LeCDCCredor(dbcintCredor.BoundText)
        Else
           LeDaTabelaParaObj "", dcbEmpenho, strQueryEmpenho
           dcbEmpenho.BoundText = strPKId
        End If
    End If
    If Area = 2 Then
        If IsNumeric(dcbEmpenho.BoundText) And Val(dcbEmpenho.BoundText) <> 0 Then
           intPkid = dcbEmpenho.BoundText
         Else
           dcbEmpenho.Text = strTexto
        End If
    End If
        
    If dcbParcela.ListCount = 1 Then
       dcbParcela.ListIndex = 0
    End If
    
End Sub

Private Function leFontedeRecurso(strEmpenhoPKID As String) As String
   Dim adoResultado    As ADODB.Recordset
   Dim strSQL As String
   
   strSQL = ""
      
   strSQL = strSQL & "SELECT FR.strCodigo, FR.strDescricao FROM "
   strSQL = strSQL & gstrEmpenho & " E,"
   strSQL = strSQL & gstrProgramaDeTrabalho & " PT,"
   strSQL = strSQL & gstrFonteRecurso & " FR"
   strSQL = strSQL & " WHERE E.Pkid = " & strEmpenhoPKID
   strSQL = strSQL & " AND PT.pkid = E.INTPROGRAMATRABALHO"
   strSQL = strSQL & " AND FR.pkid = PT.INTFONTERECURSO"
      
   Set gobjBanco = New clsBanco
   If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
       With adoResultado
           If .EOF = False Then
               leFontedeRecurso = Trim(!strCodigo)
               leFontedeRecurso = leFontedeRecurso & " - " & Trim(!strDescricao)
           End If
       End With
   End If
   
End Function


Private Sub dcbEmpenho_GotFocus()
   Dim intIndex As Integer
   
    AtivaPastaDeObjeto tab_3DPastaEmpenho, 0
    If Not mblnTelaEmpenho Then
       
       LeDaTabelaParaObj "", dcbEmpenho, strQueryEmpenho
       
    Else
       LeDaTabelaParaObj "", dcbParcela, strSqlParcelaEmpenho
       For intIndex = 0 To dcbParcela.ListCount - 1
          If dcbParcela.ItemData(intIndex) = Val(strpkidParcela) Then
             Exit For
          End If
       Next
       
       dcbParcela.ListIndex = intIndex
       VerificaLista
       mblnTelaEmpenho = False
       
    End If
    
End Sub

Private Sub txtbitDigitoProcesso_GotFocus()
    MarcaCampo txtbitDigitoProcesso
End Sub

Private Sub txtbitDigitoProcesso_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtbitDigitoProcesso
End Sub

Private Sub txtintExercicio_GotFocus()
    MarcaCampo txtintExercicio
End Sub

Private Sub txtintExercicio_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintExercicio
End Sub

Private Sub dcbParcela_Change()
    ProcuraParcelaEmpenho
End Sub

Private Sub dcbParcela_Click()
    ProcuraParcelaEmpenho
End Sub

Private Sub dcbParcela_GotFocus()
    AtivaPastaDeObjeto tab_3DPastaEmpenho, 0
End Sub

Private Sub dcbParcelaResto_Change()
    ProcuraParcelaResto
End Sub

Private Sub Resto_Click(Area As Integer)
    ProcuraParcelaResto
End Sub

Private Sub dcbParcelaResto_GotFocus()
    AtivaPastaDeObjeto tab_3DPastaEmpenho, 1
End Sub

Private Sub dcbResto_Click(Area As Integer)
   Dim strPKId    As String
   Dim strTexto   As String

    If dcbResto.BoundText = "" Then Exit Sub
    strPKId = dcbResto.BoundText
    If dcbResto.Text <> "" Then
        strTexto = dcbResto.Text
        LeDaTabelaParaObj "", dcbParcelaResto, strSqlParcelaResto
        PreencheProcesso Val(strPKId), False
        If gstrItemData(dbcintCredor) = 0 Then
           PreencherListaDeOpcoes dbcintCredor, dblQueryCredorResto
           txt_intNContribuinte = LeCDCCredor(dbcintCredor.BoundText)
        Else
           LeDaTabelaParaObj "", dcbResto, strQueryResto
           dcbResto.BoundText = strPKId
        End If
    End If
    
    If Area = 2 Then
        If IsNumeric(dcbResto.BoundText) And Val(dcbResto.BoundText) <> 0 Then
           intPkid = dcbResto.BoundText
         Else
           dcbResto.Text = strTexto
        End If
    End If
    
    If dcbParcelaResto.ListCount = 1 Then
       dcbParcelaResto.ListIndex = 0
    End If
    
End Sub

Private Sub dcbResto_GotFocus()
   Dim intIndex As Integer
    AtivaPastaDeObjeto tab_3DPastaEmpenho, 1
    
    If Not mblnTelaEmpenho Then
       
       LeDaTabelaParaObj "", dcbResto, strQueryResto
       
    Else
       
       LeDaTabelaParaObj "", dcbParcelaResto, strSqlParcelaResto
       
       For intIndex = 1 To dcbParcelaResto.ListCount
          If dcbParcelaResto.ItemData(intIndex) = Val(strpkidParcela) Then
             Exit For
          End If
       Next
       
       dcbParcelaResto.ListIndex = intIndex
       
       mblnTelaEmpenho = False
       
    End If
    
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 899
    VirificaGradeListView Me
    
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrIncluirItem, _
                             gstrExcluirItem, gstrSalvar
                             
    If (chkblnPago.Value = 1 Or chkblnCancelado.Value = 1) And mblnAlterando Then
        HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem, gstrCancelar
    End If
                             
    If chkblnPago.Value = 0 And mblnAlterando And chkblnCancelado.Value = 0 Then
        HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrSalvar, gstrCancelar
    End If
    
    If mblnselecionou Then
        HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar, gstrImprimir
    Else
        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar, gstrImprimir
    End If
    If mobjAux Is Nothing Then
        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
    Else
        HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
    End If
End Sub

Private Sub Form_Deactivate()
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrSalvar
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
  If KeyCode = vbKeyF1 Then
     Call_HtmlHelp Me.HelpContextID
  End If
End Sub

Private Sub Form_Load()
'******************************************************************************************
' Data: 09/06/2003
' Alteração: - Substituição da chamada direta à stored procedure pela função
'            gstrStoredProcedure.
' Responsável: Everton Bianchini
'******************************************************************************************

    mblnAlterando = False
        
    'VerificaListaAutomatica gstrOrdemPagamento, tdb_Lista, strQuery
        
    VerificaObjParaAplicar mobjAux
    LimpaTelaPagamento
    tab_3DPastaEmpenho.Tab = 0
    dbcintCredor.Tag = "SELECT PKID, strNome FROM " & gstrContribuinte & " ORDER BY strNome;strNome"
    txtintExercicio = gintExercicio
    chkblnCancelado.Enabled = False
    mblnClickOk = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 True, gstrMnuArquivo
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form_Deactivate
End Sub

Private Sub lvw_Despesa_GotFocus()
    AtivaPastaDeObjeto tab_3DPastaEmpenho, 2
End Sub

Private Sub lvw_Despesa_ItemClick(ByVal Item As MSComctlLib.ListItem)
    mblnAlterandoDespesa = True
    With lvw_Despesa
        dcbDespesa = .ListItems(.SelectedItem.Index).Text
    End With
End Sub

Private Sub lvw_Empenho_GotFocus()
    AtivaPastaDeObjeto tab_3DPastaEmpenho, 0
End Sub

Private Sub lvw_Empenho_ItemClick(ByVal Item As MSComctlLib.ListItem)
    mblnAlterandoEmpenho = True
    With lvw_Empenho
        dcbEmpenho = .ListItems(.SelectedItem.Index).SubItems(1)
        dcbParcela = .ListItems(.SelectedItem.Index).SubItems(2)
    End With
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)

    Select Case UCase(strModoOperacao)
        Case gstrNovo
            LimpaTelaPagamento True
            TrocaCorObjeto txtintExercicio, False
        Case gstrSalvar
            GravaOrdemPagamento
        Case UCase(gstrIncluirItem)
            VerificaLista
        Case UCase(gstrDeletar)
            If chkblnPago.Value = 1 Then
                ExibeMensagem "Esta Ordem de Pagamento não pode ser excluída por que está paga."
            End If
            If gblnExclusaoGravacaoOk("E", "Confirma Exclusão?", True) Then
                ExcluiOrdemPagamento
            End If
        Case UCase(gstrCancelar)
            CancelaPagamento
        Case UCase(gstrExcluirItem)
            VerificaListaExcluir
        Case UCase(gstrImprimir)
            If Val(txtPKId) = 0 Then
                ExibeMensagem "Não há registro selecionado para impressão"
            Else
                ImprimeBordereaux Val(txtPKId)
            End If
        Case UCase(gstrLocalizar)
            LeDaTabelaParaObj "", tdb_Lista, strQuery
        Case UCase(gstrPreencherLista)
        
            If ActiveControl.Name = cbo_HistoricoLiquidacao.Name Then
                LeDaTabelaParaObj gstrHistorico, cbo_HistoricoLiquidacao
            End If
            
            If ActiveControl.Name = txt_intNContribuinte.Name Or ActiveControl.Name = dbcintCredor.Name Then
               'LeDaTabelaParaObj gstrContribuinte, dbcintCredor, dbcintCredor.Tag
               PreencherListaDeOpcoes dbcintCredor
            End If
            
            If ActiveControl.Name = dcbEmpenho.Name Then
               'If gstrItemData(dbcintCredor) = 0 Then
               '   ExibeMensagem "É necessário informar o Credor para preencher esta lista"
               'Else
                  LeDaTabelaParaObj "", dcbEmpenho, strQueryEmpenho
                  PreencheDados
               'End If
            End If
            
            If ActiveControl.Name = dcbResto.Name Then
               'If gstrItemData(dbcintCredor) = 0 Then
               '   ExibeMensagem "É necessário informar o Credor para preencher esta lista"
               'Else
                  LeDaTabelaParaObj "", dcbResto, strQueryResto
               'End If
            End If
            
            If ActiveControl.Name = dcbDespesa.Name Then
               'If gstrItemData(dbcintCredor) = 0 Then
               '   ExibeMensagem "É necessário informar o Credor para preencher esta lista"
               'Else
                  LeDaTabelaParaObj "", dcbDespesa, strQueryDespesa
               'End If
            End If
            
        Case UCase(gstrAplicar)
            
            ToolBarGeral strModoOperacao, gstrOrdemPagamento, mblnAlterando, _
                 tdb_Lista, Me, mobjAux
        
        Case UCase(gstrRefresh)
            
            ToolBarGeral strModoOperacao, gstrOrdemPagamento, mblnAlterando, _
                 tdb_Lista, Me, mobjAux, strQuery
    End Select
    
End Sub

Private Sub lvw_Resto_GotFocus()
    AtivaPastaDeObjeto tab_3DPastaEmpenho, 1
End Sub

Private Sub lvw_Resto_ItemClick(ByVal Item As MSComctlLib.ListItem)
    With lvw_Resto
        dcbResto = .ListItems(.SelectedItem.Index).SubItems(2)
        dcbParcelaResto = .ListItems(.SelectedItem.Index).SubItems(3)
    End With
End Sub

Private Sub optbytTipo_Click(Index As Integer)
    habilitaGuias Index
End Sub

Private Sub tab_3DPastaEmpenho_Click(PreviousTab As Integer)
    If tab_3DPastaEmpenho.Tab = 4 Then
        TrocaCorObjeto txtintProcesso, True
        TrocaCorObjeto txtHistorico, True
        TrocaCorObjeto cbo_HistoricoLiquidacao, True
        TrocaCorObjeto cmd_HistoricoLiquidacao, True
    Else
        TrocaCorObjeto txtintProcesso, False
        TrocaCorObjeto txtHistorico, False
        TrocaCorObjeto cbo_HistoricoLiquidacao, False
        TrocaCorObjeto cmd_HistoricoLiquidacao, False
    End If
End Sub

Private Sub tdb_Lista_Click()
    If glngQtdLinhaTDBGrid(tdb_Lista) = 1 Then
        tdb_Lista_RowColChange 0, 0
    End If
 mblnClickOk = True
End Sub

Private Sub tdb_Lista_DblClick()
    MantemForm gstrAplicar
End Sub

Private Sub tdb_Lista_FilterChange()
    mblnClickOk = False
    gblnFilraCampos tdb_Lista
End Sub

Private Sub tdb_Lista_KeyDown(KeyCode As Integer, Shift As Integer)
'    mblnClickOk = True
End Sub

Private Sub tdb_Lista_KeyPress(KeyAscii As Integer)
    If tdb_Lista.Col = 2 Then
        CaracterValido KeyAscii, "N", tdb_Lista
    End If
End Sub

Private Sub tdb_Lista_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   'mblnClickOk = True
End Sub

Private Sub tdb_Lista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    With tdb_Lista
        If (Not .EOF And Not .BOF) And mblnClickOk Then
        
            mblnClickOk = False
            txtPKId = .Columns("PKId")
            gCorLinhaSelecionada tdb_Lista
            mblnAlterando = True
            
            If cbo_HistoricoLiquidacao.ListCount = 0 Then
                LeDaTabelaParaObj gstrHistorico, cbo_HistoricoLiquidacao
            End If
            
            If dcbEmpenho.MatchedWithList = False Then
                LeDaTabelaParaObj "", dcbEmpenho, strQueryEmpenho
                PreencheDados
            End If
            
            PreencheProcesso txtPKId, True
                        
            If dcbResto.MatchedWithList = False Then
                LeDaTabelaParaObj "", dcbResto, strQueryResto
            End If
            
            If dcbDespesa.MatchedWithList = False Then
                LeDaTabelaParaObj "", dcbDespesa, strQueryDespesa
            End If
            
            If dbcintCredor.MatchedWithList = False Then
                'LeDaTabelaParaObj "", dbcintCredor, dbcintCredor.Tag
                PreencherListaDeOpcoes dbcintCredor
            End If
            
            LePagamento
            
            txtTotalAnulacao = .Columns("dblTotalAnulacaoReceita")
            txtTotalEmpenho = .Columns("dblTotalEmpenho")
            txtTotalResto = .Columns("dblTotalResto")
            txtTotalDespesa = .Columns("dblTotalDespesaExtra")
            txtHistorico = .Columns("typHistorico")
            
            Filtrar_dbcintCredor .Columns("intContribuinte")
            
            dbcintCredor.BoundText = .Columns("intContribuinte")
            txt_intNContribuinte = LeCDCCredor(dbcintCredor.BoundText)
            
            itemAnterior = dbcintCredor.BoundText
            txtdtmData.Text = .Columns("dtmData")
            txtintExercicio = .Columns("intExercicio")
            txtdtmDataVencimento.Text = .Columns("dtmDataVencimento")
            chkblnCancelado.Value = IIf(.Columns("bytCancelado") = "", 0, .Columns("bytCancelado"))
                        
            
            LimpaCombos
            LeEmpenho
            LeResto
            LeDespesa
            LeAnulacao
            
            HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar, gstrImprimir
            AjustaFormatacaoResto
            AjustaFormatacaoEmpenho
            AjustaFormatacaoDespesa
            AjustaFormatacaoAnulacao
            
            If .Columns("blnPago") = "Sim" Then
                chkblnPago.Value = 1
                DesabilitaPago True
            Else
                chkblnPago.Value = 0
                If .Columns("bytCancelado") = 1 Then
                    DesabilitaPago True
                Else
                    DesabilitaPago False
                    VerificaPagamento
                End If
                
            End If
            TrocaCorObjeto txtintProcesso, True
            TrocaCorObjeto txtintExercicio, True
        End If
    End With
End Sub

Private Sub VerificaPagamento()
    Dim adoResultado    As ADODB.Recordset
    Dim strSQL          As String
    Dim listaPKID       As String
    Dim strTabela       As String
    Dim strCampo        As String
    
    Dim i As Integer
    For i = 1 To lvw_Empenho.ListItems.Count
        If Trim(lvw_Empenho.ListItems(i).SubItems(8)) <> "" Then
            DesabilitaPago True
            Exit Sub
        Else
            listaPKID = listaPKID & lvw_Empenho.ListItems(i).Tag & ","
        End If
        strTabela = gstrSubempenhoPagtoAnulado
        strCampo = "intSubempenho"
    Next
    
    For i = 1 To lvw_Resto.ListItems.Count
        If Trim(lvw_Resto.ListItems(i).SubItems(8)) <> "" Then
            DesabilitaPago True
            Exit Sub
        Else
            listaPKID = listaPKID & lvw_Resto.ListItems(i).Tag & ","
        End If
        strTabela = gstrParcelaRestoPagtoAnulado
        strCampo = "intSubempenho"
    Next

    For i = 1 To lvw_Despesa.ListItems.Count
        If Trim(lvw_Despesa.ListItems(i).SubItems(6)) <> "" Then
            DesabilitaPago True
            Exit Sub
        Else
            listaPKID = listaPKID & lvw_Despesa.ListItems(i).Tag & ","
        End If
        strTabela = gstrDespesaExtraOrcamPagtoAnulado
        strCampo = "intDespesa"
    Next

    If lvw_AnulacaoReceita.ListItems.Count > 0 Then Exit Sub



' FOI DESABILITADO EM 02/04/2004 POIS A VERIFICAÇÃO JÁ É FEITA
' ANTERIRORMENTE PELO CAMPO blnPago

'    If Len(listaPKID) > 0 Then listaPKID = Mid(listaPKID, 1, Len(listaPKID) - 1)
'
'    strSql = ""
'    strSql = strSql & " SELECT intProcesso FROM "
'    strSql = strSql & strTabela & " WHERE " & strCampo & " in (" & listaPKID & ")"
'
'    Set gobjBanco = New clsBanco
'    'M4R VERIFICAR TIAGO
'    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
'        With adoResultado
'            If .EOF = False Then
'                DesabilitaPago True
'                Exit Sub
'            End If
'        End With
'    End If

    
End Sub

Private Sub txtdtmData_GotFocus()
    MarcaCampo txtdtmData
End Sub

Private Sub txtdtmData_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtdtmData
End Sub

Private Sub txtdtmData_LostFocus()

    txtdtmData = gstrDataFormatada(txtdtmData)
    
    'ORC677
    If IsDate(txtdtmData) Then
        If Year(CDate(txtdtmData)) <> CInt(gintExercicio) Then
            ExibeMensagem "A data tem que estar no exercício de " & gintExercicio & "."
            If txtdtmData.Enabled Then txtdtmData.SetFocus
            Exit Sub
        End If
    End If
    
    If txtdtmData <> "" Then txtdtmDataVencimento.Text = txtdtmData.Text
    
End Sub

Private Sub txtdtmDataVencimento_GotFocus()
    MarcaCampo txtdtmDataVencimento
End Sub

Private Sub txtdtmDataVencimento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtdtmDataVencimento
End Sub

Private Sub txtdtmDataVencimento_LostFocus()

    txtdtmDataVencimento = gstrDataFormatada(txtdtmDataVencimento)
    
    'ORC677
    If IsDate(txtdtmDataVencimento) Then
        If Year(CDate(txtdtmDataVencimento)) < CInt(gintExercicio) Then
            ExibeMensagem "O ano da data de vencimento não pode ser menor que " & gintExercicio & "."
            If txtdtmDataVencimento.Enabled Then txtdtmDataVencimento.SetFocus
            Exit Sub
        End If
    End If
    
End Sub

Private Sub txtintExercicio_LostFocus()
 '  If Val(txtintExercicio) = gintExercicio Then
 '     buscaOPbyNumero txtintProcesso
 '   Else
 '       buscaOPbyExercicio
 '  End If
End Sub
Private Sub buscaOPbyExercicio()
    Dim strSQL  As String
    Dim adoResultado    As ADODB.Recordset
    'M4R REPARO NA CLAUSULA WHERE
    If tdb_Lista.Columns("Número") <> "" Then
        strSQL = "SELECT OP.PKID,OP.STRCHEQUE,OP.TYPHISTORICO,OP.DTMDTATUALIZACAO,OP.LNGCODUSR,OP.BLNPAGO,OP.BYTTIPO,OP.DTMDATA, "
        strSQL = strSQL & "OP.DTMDATAVENCIMENTO,OP.BYTCANCELADO,CT.STRNOME,CT.CDC PKIDCONTRIBUINTE FROM " & gstrOrdemPagamento & " OP, " & gstrContribuinte & " CT "
        strSQL = strSQL & "WHERE OP.INTNUMERO = " & tdb_Lista.Columns("Número")
        strSQL = strSQL & " AND OP.intcontribuinte = CT.PKID"
        strSQL = strSQL & " AND OP.INTEXERCICIO = " & tdb_Lista.Columns("Exercício")
        
        
        Set gobjBanco = New clsBanco
         
        If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
            If adoResultado.EOF = False Then
                    With adoResultado
                        txtPKId = !Pkid
                        chkblnPago.Value = IIf(!blnPago = 0, 0, 1)
                        chkblnCancelado.Value = IIf(IsNull(!bytCancelado) = True Or !bytCancelado = 0, 0, 1)
                        txtdtmData = IIf(IsNull(!DTMDATA) = True, "", Format(!DTMDATA, "DD/MM/YYYY"))
                        txtdtmDataVencimento = IIf(IsNull(!dtmDataVencimento) = True, "", Format(!dtmDataVencimento, "DD/MM/YYYY"))
                        txtHistorico = IIf(IsNull(!typHistorico) = True, "", !typHistorico)
                        If !bytTipo = 0 Then optbytTipo(0).Value = True
                        If !bytTipo = 1 Then optbytTipo(1).Value = True
                        If !bytTipo = 2 Then optbytTipo(2).Value = True
                        If !bytTipo = 3 Then optbytTipo(3).Value = True
                        txt_intNContribuinte = !PKIDCONTRIBUINTE
                        dbcintCredor.Text = IIf(IsNull(!STRNOME) = True, "", !STRNOME)
                        TrocaCorObjeto txt_intNContribuinte, True
                        TrocaCorObjeto dbcintCredor, True
                        TrocaCorObjeto txtdtmData, True
                        TrocaCorObjeto txtdtmDataVencimento, True
                        TrocaCorObjeto cbo_HistoricoLiquidacao, True
                        DoEvents
                        LeResto
                        LeDespesa
                        AjustaFormatacaoResto
                        AjustaFormatacaoDespesa
                        Totaliza lvw_Resto, txtTotalResto
                        Totaliza lvw_Empenho, txtTotalEmpenho
                        Totaliza lvw_Despesa, txtTotalDespesa
                        SomaTotalAPagar
                        
                    End With
                Exit Sub
             End If
        End If
        GoTo LimpaTela
     End If

LimpaTela:
    MsgBox "Não existe Ordem de Pagamento para esse exercício", vbInformation
    LimpaTelaPagamento

    
End Sub

Private Sub txtintExercicioProcesso_GotFocus()
    MarcaCampo txtintExercicioProcesso
End Sub

Private Sub txtintExercicioProcesso_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintExercicioProcesso
End Sub

Private Sub txtintProcesso_GotFocus()
    If Trim(txtintProcesso) = "" Then
       txtintProcesso = proximoCodigoOP
       txtintExercicio = gintExercicio
    End If
    MarcaCampo txtintProcesso
End Sub

Private Sub txtintProcesso_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintProcesso
End Sub

Private Sub txtintProcesso_LostFocus()
   'buscaOPbyNumero txtintProcesso
End Sub

Private Sub txtPKId_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub


Private Sub buscaOPbyNumero(ByVal strNumero As String)
    Dim intCount As Integer
    DoEvents
    intCount = 0
    If Len(Trim(strNumero)) > 0 Then
       'If gblnExisteValorNaTabela(gstrOrdemPagamento, "intNumero", strNumero) Then
       If gblnExisteCodigo(2, gstrOrdemPagamento, "intNumero", strNumero, "intExercicio", "'" & txtintExercicio & "'") Then
            If Not tdb_Lista.EOF Then
                tdb_Lista.MoveFirst
                Do While Not tdb_Lista.EOF
                    If tdb_Lista.Columns("intNumero") = strNumero Then
                        gCorLinhaSelecionada tdb_Lista
                        mblnClickOk = True
                        tdb_Lista_RowColChange 0, 0
                        Exit Do
                    Else
                        intCount = intCount + 1
                        tdb_Lista.MoveNext
                        If intCount > tdb_Lista.ApproxCount Then Exit Do
                    End If
                Loop
            End If
       End If
    End If
End Sub

Public Function strQueryRelatorio()

'******************************************************************************************
' Data: 09/06/2003
' Alteração: - Substituição do comando nativo de outer join (*= ou =*) do SQL Server pela
'            variável strOUTJSQLServer e inclusão do comando de outer join ((+)) do Oracle,
'            representado pela variável strOUTJOracle.
' Responsável: Everton Bianchini
'******************************************************************************************

    Dim strSQL  As String
    strSQL = ""
    strSQL = strSQL & "SELECT PR.intCodigoreduzido AS CodigoReduzido, CO.strDescricao AS CodigoOrcamentario, FR.strDescricao AS FonteRecurso, "
    strSQL = strSQL & "PR.blnEducacao AS Educacao, PR.blnFundef AS Fundef, PR.blnSaude AS Saude, PR.blnPessoal AS Pessoal, PR.dblValor  "
    strSQL = strSQL & "FROM "
    strSQL = strSQL & gstrPrevisaoDaReceita & " PR, "
    strSQL = strSQL & gstrCodigoOrcamentario & " CO, "
    strSQL = strSQL & gstrFonteRecurso & " FR "
'    strSql = strSql & "WHERE PR.intCodigoOrcamentario *= CO.PKId "
    strSQL = strSQL & "WHERE PR.intCodigoOrcamentario " & strOUTJSQLServer & "= CO.PKId " & strOUTJOracle
'    strSql = strSql & "AND PR.intFonteRecurso *= FR.PKId "
    strSQL = strSQL & "AND PR.intFonteRecurso " & strOUTJSQLServer & "= FR.PKId " & strOUTJOracle
    strSQL = strSQL & "ORDER BY PR.intCodigoreduzido, CO.strDescricao, FR.strDescricao, "
    strSQL = strSQL & "PR.blnEducacao, PR.blnFundef, PR.blnSaude, PR.blnPessoal "
    strQueryRelatorio = strSQL
End Function

Private Sub txt_strCodigo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_strCodigo
End Sub
Private Sub txt_intExercicioProcesso_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_intExercicioProcesso
End Sub
Private Sub txt_bitDigito_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_bitDigito
End Sub

Private Sub txtstrCodigoProcesso_GotFocus()
    MarcaCampo txtstrCodigoProcesso
End Sub

Private Sub txtstrCodigoProcesso_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtstrCodigoProcesso
End Sub

Private Sub txtTotalDespesa_Change()
    SomaTotalAPagar
End Sub
Private Sub txtTotalAnulacao_Change()
    SomaTotalAPagar
End Sub

Private Sub txtTotalEmpenho_Change()
    SomaTotalAPagar
End Sub

Private Sub txtTotalResto_Change()
    SomaTotalAPagar
End Sub

Private Sub AjustaFormatacaoDespesa()

    Dim lstItem As ListItem
    If (bytDBType = EDatabases.SQLServer) Then Exit Sub
    For Each lstItem In lvw_Despesa.ListItems
        lstItem.ListSubItems(2).Text = gstrConvVrDoSql(lstItem.ListSubItems(2).Text)
        lstItem.ListSubItems(3).Text = gstrConvVrDoSql(lstItem.ListSubItems(3).Text)
        lstItem.ListSubItems(4).Text = gstrConvVrDoSql(lstItem.ListSubItems(4).Text)
    Next
End Sub
Private Sub AjustaFormatacaoAnulacao()

    Dim lstItem As ListItem
    If (bytDBType = EDatabases.SQLServer) Then Exit Sub
    For Each lstItem In lvw_AnulacaoReceita.ListItems
        lstItem.ListSubItems(1).Text = gstrConvVrDoSql(lstItem.ListSubItems(1).Text)
    Next
End Sub

Private Sub AjustaFormatacaoEmpenho()

    Dim lstItem As ListItem
    If (bytDBType = EDatabases.SQLServer) Then Exit Sub
    For Each lstItem In lvw_Empenho.ListItems
        lstItem.ListSubItems(1).Text = Replace(lstItem.ListSubItems(1).Text, ".", "")
        lstItem.ListSubItems(5).Text = gstrConvVrDoSql(lstItem.ListSubItems(5).Text)
        lstItem.ListSubItems(6).Text = gstrConvVrDoSql(lstItem.ListSubItems(6).Text)
        lstItem.ListSubItems(7).Text = gstrConvVrDoSql(lstItem.ListSubItems(7).Text)
    Next
End Sub


Private Sub AjustaFormatacaoResto()
    Dim lstItem As ListItem
    If (bytDBType = EDatabases.SQLServer) Then Exit Sub
    For Each lstItem In lvw_Resto.ListItems
        lstItem.ListSubItems(2).Text = Replace(lstItem.ListSubItems(2).Text, ".", "")
        lstItem.ListSubItems(5).Text = gstrConvVrDoSql(lstItem.ListSubItems(5).Text)
        lstItem.ListSubItems(6).Text = gstrConvVrDoSql(lstItem.ListSubItems(6).Text)
        lstItem.ListSubItems(7).Text = gstrConvVrDoSql(lstItem.ListSubItems(7).Text)
    Next
End Sub


Private Sub ExcluiOrdemPagamento()
    Dim adoResultado    As ADODB.Recordset
    Dim strSQL As String
    
    Set gobjBanco = New clsBanco
    gobjBanco.ExecutaBeginTrans
    strSQL = ""
    
    If Trim(txtPKId.Text) = "" Then
        ExibeMensagem "Não há nenhum registro selecionado para a exclusão."
    End If


    strSQL = strSQL & IIf((bytDBType = EDatabases.Oracle), "BEGIN ", "")
    strSQL = strSQL & ""

    strSQL = strSQL & "DELETE tblOrdemPagamentoDespesaExtra WHERE intOrdemPagamento =" & txtPKId.Text & "; "
    strSQL = strSQL & "DELETE tblOrdemPagamentoResto WHERE intOrdemPagamento =" & txtPKId.Text & "; "
    strSQL = strSQL & "DELETE tblOrdemPagamentoEmpenho WHERE intOrdemPagamento =" & txtPKId.Text & "; "
    strSQL = strSQL & "DELETE tblOrdemPagamento Where PKID =" & txtPKId.Text & "; "

    
    
    

    strSQL = strSQL & IIf((bytDBType = EDatabases.Oracle), " END; ", "")

    Set gobjBanco = New clsBanco
    If gobjBanco.Execute(strSQL) Then
        Set gobjBanco = New clsBanco
        gobjBanco.ExecutaCommitTrans
        LimpaTelaPagamento True
        LeDaTabelaParaObj "", tdb_Lista, strQuery
    Else
        Set gobjBanco = New clsBanco
        gobjBanco.ExecutaRollbackTrans
    End If
    

End Sub


Private Sub DesabilitaPago(ByVal blnHabilita As Boolean)
    
    chkblnPago.Enabled = False
    HabilitaDesabilitaBotao1 Not blnHabilita, gstrBtnArquivo, gstrIncluirItem, gstrExcluirItem
    If chkblnCancelado.Value Then
        HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrCancelar
    Else
        HabilitaDesabilitaBotao1 Not blnHabilita, gstrBtnArquivo, gstrCancelar
    End If
    HabilitaDesabilitaBotao1 Not blnHabilita, gstrBtnArquivo, gstrDeletar
    

    If (Not tdb_Lista.EOF) And (tdb_Lista.Columns("bytTipo") <> "") Then
        TrocaCorObjeto optbytTipo(CInt(tdb_Lista.Columns("bytTipo"))), False
        optbytTipo(CInt(tdb_Lista.Columns("bytTipo"))).Value = True
        habilitaGuias CInt(tdb_Lista.Columns("bytTipo"))
    End If
    
    TrocaCorObjeto optbytTipo(0), blnHabilita
    TrocaCorObjeto optbytTipo(1), blnHabilita
    TrocaCorObjeto optbytTipo(2), blnHabilita
    TrocaCorObjeto optbytTipo(3), blnHabilita
    
    TrocaCorObjeto dcbEmpenho, blnHabilita
    TrocaCorObjeto dcbParcela, blnHabilita
    TrocaCorObjeto dcbResto, blnHabilita
    TrocaCorObjeto dcbParcelaResto, blnHabilita
    TrocaCorObjeto dcbDespesa, blnHabilita
    
    TrocaCorObjeto cmd_Empenho, blnHabilita
    TrocaCorObjeto cmd_Resto, blnHabilita
    TrocaCorObjeto cmd_Despesa, blnHabilita
    
    TrocaCorObjeto cmd_Credor, blnHabilita
    TrocaCorObjeto dbcintCredor, blnHabilita
    TrocaCorObjeto txt_intNContribuinte, blnHabilita
    
    TrocaCorObjeto txtdtmData, blnHabilita
    TrocaCorObjeto txtdtmDataVencimento, blnHabilita
    

    
End Sub

Private Sub CancelaPagamento()

    Dim strSQL              As String
    Dim adoResultado        As ADODB.Recordset
    Dim strDtmCancelamento  As String

    'If gblnExclusaoGravacaoOk("I", "Deseja cancelar o esta Ordem de Pagamento?", True) Then
    strDtmCancelamento = DataPrompt("Insira a data do cancelamento.", CDate(txtdtmData), "EF", "a Ordem de Pagamento")

    
    'ORC677
    If IsDate(strDtmCancelamento) Then
        If Year(CDate(strDtmCancelamento)) <> CInt(gintExercicio) Then
            ExibeMensagem "A data de cancelamento tem que estar no exercício de " & gintExercicio & "."
            Exit Sub
        End If
    End If
    

    If Not Trim(strDtmCancelamento) = "" Then
    
        Set gobjBanco = New clsBanco
        gobjBanco.ExecutaBeginTrans
        
        strSQL = ""
        strSQL = strSQL & "Update " & gstrOrdemPagamento & " "
        strSQL = strSQL & "Set Bytcancelado = 1 ,DtmCancelamento = " & gstrConvDtParaSql(strDtmCancelamento)
        strSQL = strSQL & ", dtmDtAtualizacao= " & gstrConvDtParaSql(gstrDataDoSistema)
        strSQL = strSQL & ", lngCodUsr= " & glngCodUsr
        strSQL = strSQL & " WHERE PKID = " & txtPKId.Text


        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
            gobjBanco.ExecutaCommitTrans
            LimpaTelaPagamento True
            LeDaTabelaParaObj "", tdb_Lista, strQuery
        Else
            gobjBanco.ExecutaRollbackTrans
        End If
    End If
End Sub

Private Sub LimpaCombos()
    dcbEmpenho.Text = ""
    dcbParcela.ListIndex = -1
    dcbResto.Text = ""
    dcbParcelaResto.ListIndex = -1
    dcbDespesa.Text = ""
End Sub

Private Sub habilitaGuias(Optional ByVal intGuia As Integer)
    tab_3DPastaEmpenho.TabEnabled(0) = False
    tab_3DPastaEmpenho.TabEnabled(1) = False
    tab_3DPastaEmpenho.TabEnabled(2) = False
    tab_3DPastaEmpenho.TabEnabled(3) = False
    
    dcbEmpenho.Enabled = True
    cmd_Empenho.Enabled = True
    dcbParcela.Enabled = True
     
    txtTotalAnulacao.Text = "0,00"
    txtTotalEmpenho.Text = "0,00"
    txtTotalResto.Text = "0,00"
    txtTotalDespesa.Text = "0,00"
    
    If intGuia <> 4 Then
        tab_3DPastaEmpenho.TabEnabled(intGuia) = True
        tab_3DPastaEmpenho.Tab = intGuia
    End If
    
    
    If intGuia = 0 Then
        Totaliza lvw_Empenho, txtTotalEmpenho
        lvw_AnulacaoReceita.ListItems.Clear
        lvw_Resto.ListItems.Clear
        lvw_Despesa.ListItems.Clear
    End If
    
    If intGuia = 1 Then
        Totaliza lvw_Resto, txtTotalResto
        lvw_AnulacaoReceita.ListItems.Clear
        lvw_Empenho.ListItems.Clear
        lvw_Despesa.ListItems.Clear
    End If
    
    If intGuia = 2 Then
        Totaliza lvw_Despesa, txtTotalDespesa
        lvw_AnulacaoReceita.ListItems.Clear
        lvw_Empenho.ListItems.Clear
        lvw_Resto.ListItems.Clear
    End If
    
    If intGuia = 3 Then
        Totaliza lvw_AnulacaoReceita, txtTotalAnulacao
        lvw_Empenho.ListItems.Clear
        lvw_Despesa.ListItems.Clear
        lvw_Resto.ListItems.Clear
    End If
End Sub

Private Function VerificaEmpenhoProcesso() As Boolean
    Dim strSQL As String
    Dim adoTemp As ADODB.Recordset
    
    If lvw_Empenho.ListItems.Count < 1 Then
        VerificaEmpenhoProcesso = True
        intPKIDEmpenho = dcbEmpenho.BoundText
        Exit Function
    End If
    
    
'    strSQL = ""
'    strSQL = strSQL & "SELECT "
'    strSQL = strSQL & "EP.Pkid "
'    strSQL = strSQL & "From "
'    strSQL = strSQL & gstrProtocolizacaoProcesso & " PP, "
'    strSQL = strSQL & gstrEmpenho & " EP, "
'    strSQL = strSQL & gstrSubempenho & " SP "
'    strSQL = strSQL & "WHERE "
'    strSQL = strSQL & "EP.Pkid = " & lvw_Empenho.SelectedItem.Tag & " AND "
'    strSQL = strSQL & "EP.Pkid = sp.intempenho AND "
'    strSQL = strSQL & "PP.strcodigo " & strOUTJSQLServer & "=" & " EP.Strcodigo " & strOUTJOracle & " AND "
'    strSQL = strSQL & "PP.bitdigito " & strOUTJSQLServer & "=" & " EP.Bitdigito " & strOUTJOracle & " AND "
'    strSQL = strSQL & "pp.intexercicio " & strOUTJSQLServer & "=" & " EP.Intexercicio " & strOUTJOracle & " AND "
'    strSQL = strSQL & "PP.bitdigito " & IIf(Digito <> "", "=" & Digito, "IS NULL") & " AND "
'    strSQL = strSQL & "pp.intexercicio " & IIf(Exercicio <> "", "=" & Exercicio, "IS NULL") & " AND "
'    strSQL = strSQL & "PP.strCodigo " & IIf(Codigo <> "", "='" & Codigo & "'", "IS NULL")
    
    
    strSQL = ""
    
    strSQL = strSQL & " SELECT  EP.strcodigo, EP.bitdigito, EP.intexercicio "
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrEmpenho & " EP"
    strSQL = strSQL & " WHERE "
    strSQL = strSQL & " EP.intNumero = " & dcbEmpenho.BoundText
    strSQL = strSQL & " UNION"
    strSQL = strSQL & " SELECT  EP.strcodigo, EP.bitdigito, EP.intexercicio "
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrEmpenho & " EP"
    strSQL = strSQL & " WHERE "
    strSQL = strSQL & " EP.intNumero = " & intPKIDEmpenho
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoTemp) Then
            If adoTemp.RecordCount > 1 Then
                VerificaEmpenhoProcesso = False
            Else
                VerificaEmpenhoProcesso = True
            End If
    End If
    
            
End Function

Private Function PreencheDados()
    Dim adoTemp As ADODB.Recordset
            Set gobjBanco = New clsBanco
            If gobjBanco.CriaADO(strQueryEmpenho, 5, adoTemp) Then
                Codigo = gstrENulo(adoTemp.Fields("strcodigo"))
                Digito = gstrENulo(adoTemp.Fields("Bitdigito"))
                Exercicio = gstrENulo(adoTemp.Fields("IntExercicio"))
            End If
End Function

Private Function proximoCodigoOP() As String
    txt_tmp = ""
     proximoCodigoOP = gstrProximoCodigo(txt_tmp, gstrOrdemPagamento, "intNumero", gintCodSeguranca, "intExercicio", Val(gintExercicio), , True, , , "intExercicio", Val(gintExercicio))
     'proximoCodigoOP = txt_tmp
     
End Function

Private Function LeCDCCredor(Optional strPKId As String, Optional strCDC As String)
    Dim strSQL          As String
    Dim adoResultado    As ADODB.Recordset
    
    If Trim(strPKId) = "" And Trim(strCDC) = "" Then
        LeCDCCredor = ""
        Exit Function
    End If
    
    strSQL = ""
    strSQL = strSQL & "SELECT CDC , PKID"
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrContribuinte
    If strPKId <> "" Then
        strSQL = strSQL & " WHERE PKID = " & strPKId
    ElseIf strCDC <> "" Then
        strSQL = strSQL & " WHERE CDC = " & strCDC
    End If
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        If adoResultado.EOF = False Then
            If strPKId <> "" Then
                LeCDCCredor = gstrENulo(adoResultado!CDC)
            ElseIf strCDC <> "" Then
                LeCDCCredor = gstrENulo(adoResultado!Pkid)
            End If
        End If
        
    End If
End Function

Private Function blnProcuraParcela() As Boolean
    Dim i As Integer
    blnProcuraParcela = False
    
    For i = 1 To lvw_Empenho.ListItems.Count
        If (lvw_Empenho.ListItems(i).SubItems(2) = dcbParcela And lvw_Empenho.ListItems(i).SubItems(1) = dcbEmpenho) Then
            blnProcuraParcela = True
            Exit Function
        End If
    Next

End Function

Private Function blnVerificaDataDaOrdem() As Boolean
    Dim i As Integer
    blnVerificaDataDaOrdem = False
    
    For i = 1 To lvw_Empenho.ListItems.Count
        If txtdtmData < CDate(lvw_Empenho.ListItems(i).SubItems(4)) Then
            blnVerificaDataDaOrdem = True
            Exit Function
        End If
    Next

End Function

Private Function dblValorTotalDescontoEmpenho() As Double
Dim intCont As Integer

    dblValorTotalDescontoEmpenho = 0
    For intCont = 1 To lvw_Empenho.ListItems.Count
        dblValorTotalDescontoEmpenho = dblValorTotalDescontoEmpenho + lvw_Empenho.ListItems(intCont).SubItems(6)
    Next intCont

End Function
Private Function dblValorTotalDescontoRestos() As Double
Dim intCont As Integer

    dblValorTotalDescontoRestos = 0
    For intCont = 1 To lvw_Resto.ListItems.Count
        dblValorTotalDescontoRestos = dblValorTotalDescontoRestos + lvw_Resto.ListItems(intCont).SubItems(6)
    Next intCont

End Function

Private Function dblValorTotalDescontoDoExtra() As Double
Dim intCont As Integer

    dblValorTotalDescontoDoExtra = 0
    For intCont = 1 To lvw_Despesa.ListItems.Count
        dblValorTotalDescontoDoExtra = dblValorTotalDescontoDoExtra + lvw_Despesa.ListItems(intCont).SubItems(3)
    Next intCont

End Function

Private Function VerificaAnulacaoReceitaProcesso() As Boolean
   Dim strSQL As String
   Dim adoResultado As New ADODB.Recordset
   
   strSQL = "SELECT PP.PKID FROM " & gstrProtocolizacaoProcesso & " PP "
   strSQL = strSQL & "WHERE RTRIM(LTRIM(PP.strCodigo)) = '" & Trim(txt_strCodigo) & "' AND "
   strSQL = strSQL & IIf(Len(Trim(txt_bitDigito)) > 0, "RTRIM(LTRIM(PP.bitDigito)) =  " & Trim(txt_bitDigito), "bitDigito IS NULL") & " AND "
   strSQL = strSQL & IIf(Len(Trim(txt_intExercicioProcesso)) > 0, "RTRIM(LTRIM(PP.intExercicio)) =  " & Trim(txt_intExercicioProcesso), "intExercicio IS NULL")
   
   Set gobjBanco = New clsBanco
   
   If gobjBanco.CriaADO(strSQL, 10, adoResultado) Then
      If Not adoResultado.EOF Then
         VerificaAnulacaoReceitaProcesso = True
      End If
   End If
   
End Function

Private Function dblQueryCredor() As Double
   Dim strSQL As String
   Dim adoResultado As New ADODB.Recordset
   
   If dcbEmpenho.MatchedWithList Then
        strSQL = "SELECT intCredor FROM " & gstrEmpenho
        strSQL = strSQL & " WHERE PKID = " & Val(dcbEmpenho.BoundText)
   Else
       'Set gobjBanco = New clsBanco
       'strSql = "SELECT intCredor FROM tblEmpenho WHERE intNumero = " & Trim(dcbEmpenho.Text)
       'If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
           'If Not adoResultado.EOF Then
        strSQL = "SELECT intCredor FROM " & gstrEmpenho
        strSQL = strSQL & " WHERE intNumero = " & Trim(dcbEmpenho.Text)
        
        If bytDBType = Oracle Then
            strSQL = strSQL & " AND TO_CHAR(dtmData,'YYYY') = "
        Else
            strSQL = strSQL & " AND YEAR(dtmData) = "
        End If
        strSQL = strSQL & gintExercicio
           'End If
       'End If
   End If
   Set gobjBanco = New clsBanco
   If gobjBanco.CriaADO(strSQL, 10, adoResultado) Then
      If Not adoResultado.EOF Then
         dblQueryCredor = gstrENulo(adoResultado!intCredor)
      Else
         dblQueryCredor = 0
      End If
   Else
      dblQueryCredor = 0
   End If
   
End Function
Private Function dblQueryCredorResto() As Double
   Dim strSQL As String
   Dim adoResultado As New ADODB.Recordset
   
   strSQL = "SELECT intCredor FROM " & gstrEmpenho
   strSQL = strSQL & " WHERE PKID = " & Val(dcbResto.BoundText)
   
   Set gobjBanco = New clsBanco
   If gobjBanco.CriaADO(strSQL, 10, adoResultado) Then
      If Not adoResultado.EOF Then
         dblQueryCredorResto = gstrENulo(adoResultado!intCredor)
      Else
         dblQueryCredorResto = 0
      End If
   Else
      dblQueryCredorResto = 0
   End If
   
End Function

Private Function dblQueryCredorDespesa(intPkidCredor As Integer) As Double
   
Dim strSQL As String
Dim adoResultado As New ADODB.Recordset

   strSQL = "SELECT intContribuinte FROM " & gstrDespesaExtraOrcamentaria
   strSQL = strSQL & " WHERE PKID = " & intPkidCredor

   Set gobjBanco = New clsBanco
   If gobjBanco.CriaADO(strSQL, 10, adoResultado) Then
      If Not adoResultado.EOF Then
         dblQueryCredorDespesa = gstrENulo(adoResultado!intcontribuinte)
      Else
         dblQueryCredorDespesa = 0
      End If
   Else
      dblQueryCredorDespesa = 0
   End If

End Function

Private Function strQuery() As String

Dim strSQL               As String
Dim strParametrosBusca   As String
Dim strTotal As String

    strParametrosBusca = ParametrosPesquisaOrdemPagamento
    
    strSQL = " SELECT "
    strSQL = strSQL & " OPS.* "
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrVw_Ordem_Pagamento & " OPS "
            
    If Trim(strParametrosBusca) <> "" Then
        strSQL = strSQL & " WHERE "
        strSQL = strSQL & Mid(strParametrosBusca, 5)
    End If
            
    strSQL = strSQL & " ORDER BY OPS.intNumero "
           
    strQuery = strSQL
    
End Function

Function ParametrosPesquisaOrdemPagamento() As String

Dim strSQL As String

    'ACRESCENTA NUMERO DA OP A PESQUISA
    strSQL = ""
    
    If Trim(txtintProcesso.Text) <> "" Then
        strSQL = strSQL & " AND OPS.intNumero = " & txtintProcesso.Text
    End If
    
    'ACRESCENTA EXERCICIO A PESQUISA
    If Trim(txtintExercicio.Text) <> "" And Len(Trim(txtintExercicio.Text)) = 4 Then
        strSQL = strSQL & " AND OPS.intExercicio = " & txtintExercicio.Text
    End If
    
    'ACRESCENTA DATA A PESQUISA
    If Trim(txtdtmData.Text) <> "" Then
        If gblnDataValida(txtdtmData.Text) Then
            strSQL = strSQL & " AND  OPS.dtmData = " & gstrConvDtParaSql(txtdtmData.Text)
        End If
    End If
    
    'ACRESCENTA DATA VENCIMENTO A PESQUISA
    If Trim(txtdtmDataVencimento.Text) <> "" Then
        If gblnDataValida(txtdtmDataVencimento.Text) Then
            strSQL = strSQL & " AND  OPS.dtmDataVencimento = " & gstrConvDtParaSql(txtdtmDataVencimento.Text)
        End If
    End If
    
    'ACRESCENTA NUMERO DO CREDOR A PESQUISA
    If Trim(txt_intNContribuinte.Text) <> "" Then
        strSQL = strSQL & " AND OPS.CDC = " & txt_intNContribuinte.Text
    End If
     
    'ACRESCENTA NOME DO CREDOR A PESQUISA
    If Trim(dbcintCredor.Text) <> "" And dbcintCredor.BoundText <> "" Then
        strSQL = strSQL & " AND UPPER(OPS.strNome) LIKE '" & UCase(Trim(dbcintCredor.Text)) & "%'"
    End If
    
    ParametrosPesquisaOrdemPagamento = strSQL

End Function


Private Sub DataOpPadrao()
Dim strSQL             As String
Dim adoResultado       As ADODB.Recordset
Dim dtmDataFechamento  As Date
Dim dtmDataUltimaOp    As Date

strSQL = ""
strSQL = "SELECT"
    strSQL = strSQL & " dtmFechamento"
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrFechamentoContabil
    strSQL = strSQL & " WHERE strCodigo = 'EC'"
    strSQL = strSQL & " AND intExercicio = " & gintExercicio

    Set gobjBanco = New clsBanco

    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
       If Not adoResultado.EOF Then
          dtmDataFechamento = gstrDataFormatada(adoResultado!dtmFechamento)
       End If
    End If

strSQL = ""
strSQL = "SELECT"
    strSQL = strSQL & " OP.dtmData"
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrOrdemPagamento & " OP"
    strSQL = strSQL & " WHERE OP.intExercicio = " & gintExercicio
    strSQL = strSQL & " AND OP.intNumero = (SELECT MAX(OPX.intNumero) FROM " & gstrOrdemPagamento & " OPX WHERE OPX.INTEXERCICIO = " & gintExercicio & ")"

    Set gobjBanco = New clsBanco

    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
       If Not adoResultado.EOF Then
          dtmDataUltimaOp = gstrDataFormatada(adoResultado!DTMDATA)
       End If
       
       If dtmDataUltimaOp > dtmDataFechamento Then
          txtdtmData.Text = gstrDataFormatada(dtmDataUltimaOp)
       Else
          txtdtmData.Text = DateAdd("d", 1, Format(dtmDataFechamento, "DD/MM/YY"))
       End If
    End If
    
End Sub


Private Sub Filtrar_dbcintCredor(strPKId As String)

   Dim strSQL As String

   strSQL = "SELECT CO.PKID,"
   strSQL = strSQL & " CO.STRNOME"
   strSQL = strSQL & " FROM "
   strSQL = strSQL & gstrContribuinte & " CO, "
   strSQL = strSQL & gstrItens & " IT, "
   strSQL = strSQL & gstrModuloContribuinte & " MC"
   strSQL = strSQL & " WHERE CO.PKID = " & strPKId & "AND"
   strSQL = strSQL & " IT.PKId = MC.intItem AND"
   strSQL = strSQL & " MC.intContribuinte = CO.Pkid AND"
   strSQL = strSQL & " IT.Pkid =" & gintModulo & " AND CO.BLNINATIVO = 0"

   'Cláudio
   'LeDaTabelaParaObj gstrContribuinte, dbcintCredor, "SELECT PKID, strNome FROM " & gstrContribuinte & _
                                                  " WHERE PKID = " & txt_intNContribuinte
                                                  
   LeDaTabelaParaObj gstrContribuinte, dbcintCredor, strSQL
   dbcintCredor.BoundText = strPKId


End Sub

Private Sub PreencheProcesso(lngPkid As Long, blnOP As Boolean)
    Dim strSQL  As String
    Dim adoTemp As ADODB.Recordset
    
    If Not blnOP Then
        If Trim(txtstrCodigoProcesso) = "" Then
            strSQL = ""
            strSQL = strSQL & " SELECT  EP.strcodigo, EP.bitdigito, EP.intexercicio "
            strSQL = strSQL & " FROM "
            strSQL = strSQL & gstrEmpenho & " EP"
            strSQL = strSQL & " WHERE "
            strSQL = strSQL & " EP.Pkid = " & lngPkid
            
            Set gobjBanco = New clsBanco
            If gobjBanco.CriaADO(strSQL, 5, adoTemp) Then
                If Not adoTemp.EOF Then
                    txtstrCodigoProcesso = gstrENulo(adoTemp!strCodigo)
                    txtintExercicioProcesso = gstrENulo(adoTemp!intExercicio)
                    txtbitDigitoProcesso = gstrENulo(adoTemp!bitDigito)
                Else
                    txtstrCodigoProcesso = ""
                    txtintExercicioProcesso = ""
                    txtbitDigitoProcesso = ""
                End If
            End If
        End If
    Else
        strSQL = ""
        strSQL = strSQL & " SELECT  strcodigoprocesso, bitdigitoprocesso, intexercicioprocesso "
        strSQL = strSQL & " FROM "
        strSQL = strSQL & gstrOrdemPagamento & " "
        strSQL = strSQL & " WHERE "
        strSQL = strSQL & " Pkid = " & lngPkid
        
        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strSQL, 5, adoTemp) Then
            If Not adoTemp.EOF Then
                txtstrCodigoProcesso = gstrENulo(adoTemp!strCodigoProcesso)
                txtintExercicioProcesso = gstrENulo(adoTemp!intExercicioProcesso)
                txtbitDigitoProcesso = gstrENulo(adoTemp!bitDigitoProcesso)
            End If
        End If
    End If

End Sub
