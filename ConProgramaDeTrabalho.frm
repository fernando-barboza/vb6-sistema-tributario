VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "TODG7.OCX"
Begin VB.Form frmConProgramaDeTrabalho 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Programas de Trabalho"
   ClientHeight    =   6765
   ClientLeft      =   1380
   ClientTop       =   2445
   ClientWidth     =   9690
   HelpContextID   =   15
   Icon            =   "ConProgramaDeTrabalho.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   9690
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   6705
      Left            =   30
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   30
      Width           =   9645
      _ExtentX        =   17013
      _ExtentY        =   11827
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Programas de Trabalho"
      TabPicture(0)   =   "ConProgramaDeTrabalho.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblintPrograma"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblintSubPrograma"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblintFuncao"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblintOrgao"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblintProjetoAtividade"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblintElementoDespesa"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lbl_CodigoReduzido"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblstrCodigo"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblintUnidadeOrcamentaria"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lbldblvalor"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblintSubfuncao"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblintTipoCredito"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lblintSubunidade"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lbl_Total"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "lblintVinculo"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "lblintFonteRecurso"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "lblGFRecurso"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtPKId"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtstrCodigo"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtdblValor"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txt_Total"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txtintCodigoReduzido"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "tdb_Lista"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txtstrGrupo"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txtintExercicio"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txtstrOrgao"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txtstrUnidade"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "txtstrSubunidade"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txtstrTipoCredito"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "txtstrFuncao"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "txtstrPrograma"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "txtstrProjetoAtividade"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "txtstrVinculo"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "txtstrSubfuncao"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "txtstrSubprograma"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "txtstrElemento"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "txtstrFonte"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "Frame1"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).ControlCount=   38
      Begin VB.Frame Frame1 
         Height          =   915
         Left            =   60
         TabIndex        =   39
         Top             =   2850
         Width           =   9495
         Begin VB.TextBox txt_Bloqueado 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
            CausesValidation=   0   'False
            Enabled         =   0   'False
            Height          =   285
            Left            =   930
            MultiLine       =   -1  'True
            OLEDragMode     =   1  'Automatic
            OLEDropMode     =   1  'Manual
            TabIndex        =   45
            Top             =   210
            Width           =   1545
         End
         Begin VB.TextBox txt_Empenhado 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
            CausesValidation=   0   'False
            Enabled         =   0   'False
            Height          =   285
            Left            =   4275
            MultiLine       =   -1  'True
            OLEDragMode     =   1  'Automatic
            OLEDropMode     =   1  'Manual
            TabIndex        =   44
            Top             =   540
            Width           =   1545
         End
         Begin VB.TextBox txt_Saldo 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
            CausesValidation=   0   'False
            Enabled         =   0   'False
            Height          =   285
            Left            =   7845
            MultiLine       =   -1  'True
            OLEDragMode     =   1  'Automatic
            OLEDropMode     =   1  'Manual
            TabIndex        =   43
            Top             =   540
            Width           =   1545
         End
         Begin VB.TextBox txt_Suplementado 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
            CausesValidation=   0   'False
            Enabled         =   0   'False
            Height          =   285
            Left            =   7845
            MultiLine       =   -1  'True
            OLEDragMode     =   1  'Automatic
            OLEDropMode     =   1  'Manual
            TabIndex        =   42
            Top             =   210
            Width           =   1545
         End
         Begin VB.TextBox txt_Reduzido 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
            CausesValidation=   0   'False
            Enabled         =   0   'False
            Height          =   285
            Left            =   945
            MultiLine       =   -1  'True
            OLEDragMode     =   1  'Automatic
            OLEDropMode     =   1  'Manual
            TabIndex        =   41
            Top             =   540
            Width           =   1545
         End
         Begin VB.TextBox txt_Reservado 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000000&
            CausesValidation=   0   'False
            Enabled         =   0   'False
            Height          =   285
            Left            =   4275
            MultiLine       =   -1  'True
            OLEDragMode     =   1  'Automatic
            OLEDropMode     =   1  'Manual
            TabIndex        =   40
            Top             =   210
            Width           =   1545
         End
         Begin VB.Label lbl_Bloqueado 
            AutoSize        =   -1  'True
            Caption         =   "Bloqueado"
            Height          =   195
            Left            =   120
            TabIndex        =   51
            ToolTipText     =   "bloqueado"
            Top             =   240
            Width           =   765
         End
         Begin VB.Label lbl_Emepnhado 
            AutoSize        =   -1  'True
            Caption         =   "Empenhado"
            Height          =   195
            Left            =   3375
            TabIndex        =   50
            Top             =   570
            Width           =   855
         End
         Begin VB.Label lbl_Saldo 
            AutoSize        =   -1  'True
            Caption         =   "Saldo"
            Height          =   195
            Left            =   7350
            TabIndex        =   49
            Top             =   570
            Width           =   405
         End
         Begin VB.Label lbl_Suplementado 
            AutoSize        =   -1  'True
            Caption         =   "Suplementado"
            Height          =   195
            Left            =   6735
            TabIndex        =   48
            ToolTipText     =   "suplementado"
            Top             =   240
            Width           =   1020
         End
         Begin VB.Label lbl_Reduzido 
            AutoSize        =   -1  'True
            Caption         =   "Reduzido"
            Height          =   195
            Left            =   210
            TabIndex        =   47
            Top             =   570
            Width           =   675
         End
         Begin VB.Label lbl_Reservado 
            AutoSize        =   -1  'True
            Caption         =   "Reservado"
            Height          =   195
            Left            =   3450
            TabIndex        =   46
            Top             =   240
            Width           =   780
         End
      End
      Begin VB.TextBox txtstrFonte 
         BackColor       =   &H80000000&
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   285
         Left            =   5820
         MultiLine       =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   37
         Top             =   2238
         Width           =   3735
      End
      Begin VB.TextBox txtstrElemento 
         BackColor       =   &H80000000&
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   285
         Left            =   5820
         MultiLine       =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   36
         Top             =   1930
         Width           =   3735
      End
      Begin VB.TextBox txtstrSubprograma 
         BackColor       =   &H80000000&
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   285
         Left            =   5820
         MultiLine       =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   35
         Top             =   1622
         Width           =   3735
      End
      Begin VB.TextBox txtstrSubfuncao 
         BackColor       =   &H80000000&
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   285
         Left            =   5820
         MultiLine       =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   34
         Top             =   1314
         Width           =   3735
      End
      Begin VB.TextBox txtstrVinculo 
         BackColor       =   &H80000000&
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   285
         Left            =   855
         MultiLine       =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   33
         Top             =   2238
         Width           =   3735
      End
      Begin VB.TextBox txtstrProjetoAtividade 
         BackColor       =   &H80000000&
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   285
         Left            =   855
         MultiLine       =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   32
         Top             =   1930
         Width           =   3735
      End
      Begin VB.TextBox txtstrPrograma 
         BackColor       =   &H80000000&
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   285
         Left            =   855
         MultiLine       =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   31
         Top             =   1622
         Width           =   3735
      End
      Begin VB.TextBox txtstrFuncao 
         BackColor       =   &H80000000&
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   285
         Left            =   855
         MultiLine       =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   30
         Top             =   1314
         Width           =   3735
      End
      Begin VB.TextBox txtstrTipoCredito 
         BackColor       =   &H80000000&
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   285
         Left            =   5820
         MultiLine       =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   29
         Top             =   1006
         Width           =   3735
      End
      Begin VB.TextBox txtstrSubunidade 
         BackColor       =   &H80000000&
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   285
         Left            =   855
         MultiLine       =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   28
         Top             =   1006
         Width           =   3735
      End
      Begin VB.TextBox txtstrUnidade 
         BackColor       =   &H80000000&
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   285
         Left            =   5820
         MultiLine       =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   27
         Top             =   698
         Width           =   3735
      End
      Begin VB.TextBox txtstrOrgao 
         BackColor       =   &H80000000&
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   285
         Left            =   855
         MultiLine       =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   26
         Top             =   698
         Width           =   3735
      End
      Begin VB.TextBox txtintExercicio 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000000&
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   285
         Left            =   7740
         MultiLine       =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   25
         Top             =   30
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.TextBox txtstrGrupo 
         BackColor       =   &H80000000&
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   285
         Left            =   5820
         MultiLine       =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   24
         Top             =   2550
         Width           =   3735
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_Lista 
         Height          =   2745
         Left            =   60
         Negotiate       =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   3870
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   4842
         _LayoutType     =   4
         _RowHeight      =   13
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "PKId"
         Columns(0).DataField=   "PKId"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Cod.Reduzido"
         Columns(1).DataField=   "intCodigoReduzido"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Órgão"
         Columns(2).DataField=   "Orgao"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "U.Orçamentária"
         Columns(3).DataField=   "UnidadeOrcamentaria"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).Caption=   "Subunidade"
         Columns(4).DataField=   "Subunidade"
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(5)._VlistStyle=   0
         Columns(5)._MaxComboItems=   5
         Columns(5).Caption=   "Proj/Ativ"
         Columns(5).DataField=   "ProjetoAtividade"
         Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(6)._VlistStyle=   0
         Columns(6)._MaxComboItems=   5
         Columns(6).Caption=   "Cod.E.Despesa"
         Columns(6).DataField=   "strCodigoElementoDespesa"
         Columns(6).NumberFormat=   "FormatText Event"
         Columns(6).EditMaskUpdate=   -1  'True
         Columns(6).EditMaskRight=   -1  'True
         Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(7)._VlistStyle=   0
         Columns(7)._MaxComboItems=   5
         Columns(7).Caption=   "Elemento de Despesa"
         Columns(7).DataField=   "Elemento"
         Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(8)._VlistStyle=   0
         Columns(8)._MaxComboItems=   5
         Columns(8).Caption=   "Valor"
         Columns(8).DataField=   "Valor"
         Columns(8).NumberFormat=   "Standard"
         Columns(8).EditMaskUpdate=   -1  'True
         Columns(8).EditMaskRight=   -1  'True
         Columns(8).ConvertEmptyCell=   1
         Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(9)._VlistStyle=   0
         Columns(9)._MaxComboItems=   5
         Columns(9).Caption=   "Total"
         Columns(9).DataField=   "dblTotal"
         Columns(9).NumberFormat=   "Standard"
         Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   10
         Splits(0)._UserFlags=   0
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).ScrollBars=   2
         Splits(0).AllowColSelect=   0   'False
         Splits(0).DividerColor=   13160664
         Splits(0).FilterBar=   -1  'True
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=10"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(8)=   "Column(1).Width=1931"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1852"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=260"
         Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(14)=   "Column(2).Width=1244"
         Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=1164"
         Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=260"
         Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(20)=   "Column(3).Width=2143"
         Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=2064"
         Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=260"
         Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(26)=   "Column(4).Width=1667"
         Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=1588"
         Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
         Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=260"
         Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
         Splits(0)._ColumnProps(32)=   "Column(5).Width=1482"
         Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
         Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=1402"
         Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
         Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=260"
         Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
         Splits(0)._ColumnProps(38)=   "Column(6).Width=1349"
         Splits(0)._ColumnProps(39)=   "Column(6).DividerColor=0"
         Splits(0)._ColumnProps(40)=   "Column(6)._WidthInPix=1270"
         Splits(0)._ColumnProps(41)=   "Column(6)._EditAlways=0"
         Splits(0)._ColumnProps(42)=   "Column(6).Order=7"
         Splits(0)._ColumnProps(43)=   "Column(7).Width=3942"
         Splits(0)._ColumnProps(44)=   "Column(7).DividerColor=0"
         Splits(0)._ColumnProps(45)=   "Column(7)._WidthInPix=3863"
         Splits(0)._ColumnProps(46)=   "Column(7)._EditAlways=0"
         Splits(0)._ColumnProps(47)=   "Column(7)._ColStyle=260"
         Splits(0)._ColumnProps(48)=   "Column(7).Order=8"
         Splits(0)._ColumnProps(49)=   "Column(8).Width=2461"
         Splits(0)._ColumnProps(50)=   "Column(8).DividerColor=0"
         Splits(0)._ColumnProps(51)=   "Column(8)._WidthInPix=2381"
         Splits(0)._ColumnProps(52)=   "Column(8)._EditAlways=0"
         Splits(0)._ColumnProps(53)=   "Column(8)._ColStyle=514"
         Splits(0)._ColumnProps(54)=   "Column(8).Order=9"
         Splits(0)._ColumnProps(55)=   "Column(9).Width=2725"
         Splits(0)._ColumnProps(56)=   "Column(9).DividerColor=0"
         Splits(0)._ColumnProps(57)=   "Column(9)._WidthInPix=2646"
         Splits(0)._ColumnProps(58)=   "Column(9)._EditAlways=0"
         Splits(0)._ColumnProps(59)=   "Column(9).AllowSizing=0"
         Splits(0)._ColumnProps(60)=   "Column(9).Visible=0"
         Splits(0)._ColumnProps(61)=   "Column(9).Order=10"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowUpdate     =   0   'False
         DefColWidth     =   0
         EditDropDown    =   0   'False
         HeadLines       =   1
         FootLines       =   1
         MarqueeUnique   =   0   'False
         TabAction       =   2
         MultipleLines   =   0
         EmptyRows       =   -1  'True
         CellTips        =   1
         CellTipsWidth   =   0
         InsertMode      =   0   'False
         MultiSelect     =   0
         DeadAreaBackColor=   13160664
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
         _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
         _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
         _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33"
         _StyleDefs(7)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
         _StyleDefs(8)   =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=3"
         _StyleDefs(9)   =   "FooterStyle:id=3,.parent=1,.namedParent=35"
         _StyleDefs(10)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(11)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
         _StyleDefs(12)  =   "EditorStyle:id=7,.parent=1"
         _StyleDefs(13)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
         _StyleDefs(14)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
         _StyleDefs(15)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
         _StyleDefs(16)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
         _StyleDefs(17)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
         _StyleDefs(18)  =   "Splits(0).Style:id=87,.parent=1"
         _StyleDefs(19)  =   "Splits(0).CaptionStyle:id=96,.parent=4"
         _StyleDefs(20)  =   "Splits(0).HeadingStyle:id=88,.parent=2"
         _StyleDefs(21)  =   "Splits(0).FooterStyle:id=89,.parent=3"
         _StyleDefs(22)  =   "Splits(0).InactiveStyle:id=90,.parent=5"
         _StyleDefs(23)  =   "Splits(0).SelectedStyle:id=92,.parent=6"
         _StyleDefs(24)  =   "Splits(0).EditorStyle:id=91,.parent=7"
         _StyleDefs(25)  =   "Splits(0).HighlightRowStyle:id=93,.parent=8"
         _StyleDefs(26)  =   "Splits(0).EvenRowStyle:id=94,.parent=9"
         _StyleDefs(27)  =   "Splits(0).OddRowStyle:id=95,.parent=10"
         _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=97,.parent=11"
         _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=98,.parent=12"
         _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=16,.parent=87"
         _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=13,.parent=88"
         _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=14,.parent=89"
         _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=15,.parent=91"
         _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=102,.parent=87"
         _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=99,.parent=88,.alignment=0"
         _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=100,.parent=89"
         _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=101,.parent=91"
         _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=118,.parent=87"
         _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=115,.parent=88,.alignment=0"
         _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=116,.parent=89"
         _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=117,.parent=91"
         _StyleDefs(42)  =   "Splits(0).Columns(3).Style:id=122,.parent=87"
         _StyleDefs(43)  =   "Splits(0).Columns(3).HeadingStyle:id=119,.parent=88,.alignment=0"
         _StyleDefs(44)  =   "Splits(0).Columns(3).FooterStyle:id=120,.parent=89"
         _StyleDefs(45)  =   "Splits(0).Columns(3).EditorStyle:id=121,.parent=91"
         _StyleDefs(46)  =   "Splits(0).Columns(4).Style:id=126,.parent=87"
         _StyleDefs(47)  =   "Splits(0).Columns(4).HeadingStyle:id=123,.parent=88,.alignment=0"
         _StyleDefs(48)  =   "Splits(0).Columns(4).FooterStyle:id=124,.parent=89"
         _StyleDefs(49)  =   "Splits(0).Columns(4).EditorStyle:id=125,.parent=91"
         _StyleDefs(50)  =   "Splits(0).Columns(5).Style:id=142,.parent=87"
         _StyleDefs(51)  =   "Splits(0).Columns(5).HeadingStyle:id=139,.parent=88,.alignment=0"
         _StyleDefs(52)  =   "Splits(0).Columns(5).FooterStyle:id=140,.parent=89"
         _StyleDefs(53)  =   "Splits(0).Columns(5).EditorStyle:id=141,.parent=91"
         _StyleDefs(54)  =   "Splits(0).Columns(6).Style:id=20,.parent=87"
         _StyleDefs(55)  =   "Splits(0).Columns(6).HeadingStyle:id=17,.parent=88"
         _StyleDefs(56)  =   "Splits(0).Columns(6).FooterStyle:id=18,.parent=89"
         _StyleDefs(57)  =   "Splits(0).Columns(6).EditorStyle:id=19,.parent=91"
         _StyleDefs(58)  =   "Splits(0).Columns(7).Style:id=146,.parent=87"
         _StyleDefs(59)  =   "Splits(0).Columns(7).HeadingStyle:id=143,.parent=88,.alignment=0"
         _StyleDefs(60)  =   "Splits(0).Columns(7).FooterStyle:id=144,.parent=89"
         _StyleDefs(61)  =   "Splits(0).Columns(7).EditorStyle:id=145,.parent=91"
         _StyleDefs(62)  =   "Splits(0).Columns(8).Style:id=150,.parent=87,.alignment=1"
         _StyleDefs(63)  =   "Splits(0).Columns(8).HeadingStyle:id=147,.parent=88,.alignment=2"
         _StyleDefs(64)  =   "Splits(0).Columns(8).FooterStyle:id=148,.parent=89"
         _StyleDefs(65)  =   "Splits(0).Columns(8).EditorStyle:id=149,.parent=91"
         _StyleDefs(66)  =   "Splits(0).Columns(9).Style:id=24,.parent=87"
         _StyleDefs(67)  =   "Splits(0).Columns(9).HeadingStyle:id=21,.parent=88"
         _StyleDefs(68)  =   "Splits(0).Columns(9).FooterStyle:id=22,.parent=89"
         _StyleDefs(69)  =   "Splits(0).Columns(9).EditorStyle:id=23,.parent=91"
         _StyleDefs(70)  =   "Named:id=33:Normal"
         _StyleDefs(71)  =   ":id=33,.parent=0"
         _StyleDefs(72)  =   "Named:id=34:Heading"
         _StyleDefs(73)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(74)  =   ":id=34,.wraptext=-1"
         _StyleDefs(75)  =   "Named:id=35:Footing"
         _StyleDefs(76)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(77)  =   "Named:id=36:Selected"
         _StyleDefs(78)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(79)  =   "Named:id=37:Caption"
         _StyleDefs(80)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(81)  =   "Named:id=38:HighlightRow"
         _StyleDefs(82)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
         _StyleDefs(83)  =   "Named:id=39:EvenRow"
         _StyleDefs(84)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(85)  =   "Named:id=40:OddRow"
         _StyleDefs(86)  =   ":id=40,.parent=33"
         _StyleDefs(87)  =   "Named:id=41:RecordSelector"
         _StyleDefs(88)  =   ":id=41,.parent=34"
         _StyleDefs(89)  =   "Named:id=42:FilterBar"
         _StyleDefs(90)  =   ":id=42,.parent=33"
      End
      Begin VB.TextBox txtintCodigoReduzido 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         Enabled         =   0   'False
         Height          =   285
         Left            =   855
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   390
         Width           =   1005
      End
      Begin VB.TextBox txt_Total 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000000&
         CausesValidation=   0   'False
         Enabled         =   0   'False
         Height          =   285
         Left            =   3045
         MultiLine       =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   5
         Top             =   2550
         Width           =   1545
      End
      Begin VB.TextBox txtdblValor 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   285
         Left            =   855
         MultiLine       =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   2
         Top             =   2550
         Width           =   1545
      End
      Begin VB.TextBox txtstrCodigo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         CausesValidation=   0   'False
         Height          =   285
         Left            =   5820
         LinkItem        =   "0"
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   4
         TabStop         =   0   'False
         Tag             =   "2"
         Top             =   390
         Width           =   3735
      End
      Begin VB.TextBox txtPKId 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         Enabled         =   0   'False
         Height          =   315
         Left            =   8610
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label lblGFRecurso 
         AutoSize        =   -1  'True
         Caption         =   "G.F.de Recurso"
         Height          =   195
         Left            =   4635
         TabIndex        =   23
         ToolTipText     =   "Grupo da Fonte de Recurso"
         Top             =   2580
         Width           =   1125
      End
      Begin VB.Label lblintFonteRecurso 
         AutoSize        =   -1  'True
         Caption         =   "F.de Recurso"
         Height          =   195
         Left            =   4800
         TabIndex        =   22
         ToolTipText     =   "Fonte de Recurso"
         Top             =   2274
         Width           =   960
      End
      Begin VB.Label lblintVinculo 
         AutoSize        =   -1  'True
         Caption         =   "Vínculo"
         Height          =   195
         Left            =   255
         TabIndex        =   21
         Top             =   2274
         Width           =   555
      End
      Begin VB.Label lbl_Total 
         AutoSize        =   -1  'True
         Caption         =   "Total"
         Height          =   195
         Left            =   2610
         TabIndex        =   20
         Top             =   2580
         Width           =   360
      End
      Begin VB.Label lblintSubunidade 
         AutoSize        =   -1  'True
         Caption         =   "Subunid."
         Height          =   195
         Left            =   180
         TabIndex        =   19
         ToolTipText     =   "Unidade orçamentária"
         Top             =   1058
         Width           =   630
      End
      Begin VB.Label lblintTipoCredito 
         AutoSize        =   -1  'True
         Caption         =   "Tipo do Crédito"
         Height          =   195
         Left            =   4680
         TabIndex        =   18
         Top             =   1058
         Width           =   1080
      End
      Begin VB.Label lblintSubfuncao 
         AutoSize        =   -1  'True
         Caption         =   "Subfunções"
         Height          =   195
         Left            =   4905
         TabIndex        =   17
         ToolTipText     =   "Subfunções de governo"
         Top             =   1362
         Width           =   855
      End
      Begin VB.Label lbldblvalor 
         AutoSize        =   -1  'True
         Caption         =   "Valor"
         Height          =   195
         Left            =   450
         TabIndex        =   16
         Top             =   2580
         Width           =   360
      End
      Begin VB.Label lblintUnidadeOrcamentaria 
         AutoSize        =   -1  'True
         Caption         =   "U.Orçamentária"
         Height          =   195
         Left            =   4650
         TabIndex        =   15
         ToolTipText     =   "Unidade orçamentária"
         Top             =   754
         Width           =   1110
      End
      Begin VB.Label lblstrCodigo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Funcional Programática"
         Height          =   195
         Left            =   4095
         TabIndex        =   14
         Top             =   450
         Width           =   1665
      End
      Begin VB.Label lbl_CodigoReduzido 
         AutoSize        =   -1  'True
         Caption         =   "Reduzido"
         Height          =   195
         Left            =   135
         TabIndex        =   13
         Top             =   450
         Width           =   675
      End
      Begin VB.Label lblintElementoDespesa 
         AutoSize        =   -1  'True
         Caption         =   "E.da Despesa"
         Height          =   195
         Left            =   4755
         TabIndex        =   12
         ToolTipText     =   "Elemento da despesa"
         Top             =   1970
         Width           =   1005
      End
      Begin VB.Label lblintProjetoAtividade 
         AutoSize        =   -1  'True
         Caption         =   "Proj/Ativ"
         Height          =   195
         Left            =   195
         TabIndex        =   11
         ToolTipText     =   "Projeto/Atividade"
         Top             =   1970
         Width           =   615
      End
      Begin VB.Label lblintOrgao 
         AutoSize        =   -1  'True
         Caption         =   "Orgão"
         Height          =   195
         Left            =   375
         TabIndex        =   10
         Top             =   754
         Width           =   435
      End
      Begin VB.Label lblintFuncao 
         AutoSize        =   -1  'True
         Caption         =   "Funções"
         Height          =   195
         Left            =   195
         TabIndex        =   9
         ToolTipText     =   "Funções de governo"
         Top             =   1362
         Width           =   615
      End
      Begin VB.Label lblintSubPrograma 
         AutoSize        =   -1  'True
         Caption         =   "Subprograma"
         Height          =   195
         Left            =   4815
         TabIndex        =   8
         Top             =   1666
         Width           =   945
      End
      Begin VB.Label lblintPrograma 
         AutoSize        =   -1  'True
         Caption         =   "Programa"
         Height          =   195
         Left            =   135
         TabIndex        =   7
         Top             =   1666
         Width           =   675
      End
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Bloqueado"
      Height          =   195
      Left            =   7230
      TabIndex        =   38
      Top             =   3420
      Width           =   765
   End
End
Attribute VB_Name = "frmConProgramaDeTrabalho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mblnClickOk             As Boolean
    Dim mblnSelecionou          As Boolean
    Dim mobjAux                 As Object
    Dim mblocalizar             As Boolean

Private Sub Form_Activate()
    gintCodSeguranca = 866
    VirificaGradeListView Me
    txtintExercicio = gintExercicio
    If mblnSelecionou Then
        HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
    Else
        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
    End If
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar, gstrNovo, gstrSalvar
End Sub

Private Sub Form_Deactivate()
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, _
                             gstrAplicar, _
                             gstrDeletar, _
                             gstrSalvar, _
                             gstrNovo
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = vbKeyF1 Then
        Call_HtmlHelp Me.HelpContextID
    End If
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    VerificaListaAutomatica "", tdb_Lista, strQuery
    VerificaObjParaAplicar mobjAux
    Screen.MousePointer = vbDefault
    mblocalizar = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form_Deactivate
End Sub

Private Sub tdb_Lista_Click()
mblocalizar = True
End Sub

Private Sub tdb_Lista_DataSourceChanged()
    txt_Total = tdb_Lista.Columns("dblTotal")
End Sub

Private Sub tdb_Lista_DblClick()
    MantemForm gstrAplicar
    mblocalizar = True
End Sub

Private Sub tdb_Lista_FilterChange()
    gblnFilraCampos tdb_Lista
End Sub

Private Sub tdb_Lista_FormatText(ByVal ColIndex As Integer, Value As Variant, Bookmark As Variant)
    Value = gvntFormatacaoEspecifica(Value, 3)
End Sub

Private Sub tdb_Lista_KeyDown(KeyCode As Integer, Shift As Integer)
    'mblnClickOk = True
    mblocalizar = True
End Sub

Private Sub tdb_Lista_KeyPress(KeyAscii As Integer)
mblocalizar = True
End Sub

Private Sub tdb_Lista_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mblnClickOk = True
    mblocalizar = True
End Sub

Private Sub tdb_Lista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    With tdb_Lista
        If (Not .EOF And Not .BOF) And mblnClickOk Then
            mblnClickOk = False
            txtPKId = tdb_Lista.Columns(0).Value
            LeProgramaTrabalho tdb_Lista, tdb_Lista, txtstrOrgao, txtstrSubunidade, _
                               txtstrFuncao, txtstrPrograma, txtstrProjetoAtividade, _
                               txtstrUnidade, txtstrTipoCredito, txtstrSubfuncao, _
                               txtstrSubprograma, txtstrElemento, txt_Saldo, _
                               txt_Empenhado, txtdblValor, txt_Bloqueado, _
                               txt_Suplementado, txt_Reduzido, txt_Reservado, _
                               txtintCodigoReduzido, txtstrCodigo, txtstrVinculo, _
                               txtstrFonte, txtstrGrupo
            gCorLinhaSelecionada tdb_Lista
            If mobjAux Is Nothing Then
                HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
            Else
                HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
            End If
        End If
    End With
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
    
     If strModoOperacao = gstrLocalizar Then
        If mblocalizar = True Then
            Exit Sub
        End If
    End If
    ToolBarGeral strModoOperacao, gstrProgramaDeTrabalho, False, _
                 tdb_Lista, Me, mobjAux, strQuery, strQueryAplicar, _
                 rptProgramaDeTrabalho, strQueryRelatorio
                 
    If strModoOperacao = gstrNovo Then
        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar, gstrNovo, gstrSalvar
    End If

End Sub

Private Function strQueryAplicar() As String
    Dim strSql  As String
    strSql = ""
    strSql = strSql & "SELECT PKId, strCodigo "
    strSql = strSql & "FROM " & gstrProgramaDeTrabalho & " "
    strSql = strSql & "WHERE bytSituacao = 1 "
    strSql = strSql & "AND intExercicio = " & Val(txtintExercicio)
    strQueryAplicar = strSql
End Function

Private Function strQuery() As String

'******************************************************************************************
' Data: 09/06/2003
' Alteração: - Substituição do comando nativo ISNULL() do SQL Server pela função gstrISNULL
' Responsável: Everton Bianchini
'------------------------------------------------------------------------------------------
' Data: 09/06/2003
' Alteração: - Substituição do comando nativo de outer join (*= ou =*) do SQL Server pela
'            variável strOUTJSQLServer e inclusão do comando de outer join ((+)) do Oracle,
'            representado pela variável strOUTJOracle.
' Responsável: Everton Bianchini
'******************************************************************************************

    Dim strSql  As String
    strSql = ""
    strSql = strSql & "SELECT PT.PKId, PT.intCodigoReduzido, OG.strCodigo AS Orgao, "
    strSql = strSql & "UO.strCodigo AS UnidadeOrcamentaria, SU.strCodigo AS SubUnidade, "
    strSql = strSql & "PA.strCodigo AS ProjetoAtividade, "
    strSql = strSql & "ED.strCodigoElementoDespesa, ED.strDescricao AS Elemento, "
'    strSql = strSql & "PT.dblValor AS Valor, (SELECT ISNULL(SUM(dblValor), 0) "
    strSql = strSql & "PT.dblValor AS Valor, (SELECT " & gstrISNULL("SUM(dblValor)", "0")
    strSql = strSql & "FROM " & gstrProgramaDeTrabalho & " WHERE intExercicio = " & gintExercicio & " ) AS dblTotal "
    strSql = strSql & "FROM "
    strSql = strSql & gstrProgramaDeTrabalho & " PT, " & gstrOrgao & " OG, "
    strSql = strSql & gstrUnidadeOrcamentaria & " UO, " & gstrSubUnidade & " SU, "
    strSql = strSql & gstrProjeto & " PA, " & gstrElementoDespesa & " ED "
'    strSql = strSql & "WHERE PT.intOrgao *= OG.PKId "
    strSql = strSql & "WHERE PT.intOrgao " & strOUTJSQLServer & "= OG.PKId " & strOUTJOracle
'    strSql = strSql & "AND PT.intUnidadeOrcamentaria *= UO.PKId "
    strSql = strSql & "AND PT.intUnidadeOrcamentaria " & strOUTJSQLServer & "= UO.PKId " & strOUTJOracle
'    strSql = strSql & "AND PT.intProjetoAtividade *= PA.PKId "
    strSql = strSql & "AND PT.intProjetoAtividade " & strOUTJSQLServer & "= PA.PKId " & strOUTJOracle
'    strSql = strSql & "AND PT.intSubunidade *= SU.PKId "
    strSql = strSql & "AND PT.intSubunidade " & strOUTJSQLServer & "= SU.PKId " & strOUTJOracle
'    strSql = strSql & "AND PT.intElementoDespesa *= ED.PKId "
    strSql = strSql & "AND PT.intElementoDespesa " & strOUTJSQLServer & "= ED.PKId " & strOUTJOracle
    strSql = strSql & "AND PT.intExercicio = " & gintExercicio & " "
    strSql = strSql & "AND PT.bytSituacao = 1 "
    strSql = strSql & "ORDER BY PT.intCodigoReduzido, OG.strCodigo, "
    strSql = strSql & "UO.strCodigo, SU.strCodigo, PA.strCodigo, "
    strSql = strSql & "ED.strCodigoElementoDespesa"
    strQuery = strSql
End Function

Private Sub txtdblValor_GotFocus()
    MarcaCampo txtdblValor
End Sub

Public Function strQueryRelatorio()

'******************************************************************************************
' Data: 09/06/2003
' Alteração: - Substituição do comando nativo de outer join (*= ou =*) do SQL Server pela
'            variável strOUTJSQLServer e inclusão do comando de outer join ((+)) do Oracle,
'            representado pela variável strOUTJOracle.
' Responsável: Everton Bianchini
'******************************************************************************************

    Dim strSql  As String
    strSql = ""
    strSql = strSql & "SELECT PT.intCodigoReduzido AS CodigoReduzido, "
    strSql = strSql & "PT.strCodigo AS CódigoProgramaTrabalho, OG.strDescricao AS Orgao, "
    strSql = strSql & "UO.strDescricao AS UnidadeOrcamentaria, "
    strSql = strSql & "SU.strDescricao AS Subunidade, TC.strDescricao AS TipoCredito, "
    strSql = strSql & "FG.strDescricao AS FuncaoGoverno, "
    strSql = strSql & "SF.strDescricao AS Subfuncao, PG.strDescricao AS Programa, "
    strSql = strSql & "SP.strDescricao AS Subprograma, PJ.strDescricao AS ProjetoAtividade, "
    strSql = strSql & "ED.strDescricao AS ElementoDespesa, "
    strSql = strSql & "PT.dblValor AS Valor "
    strSql = strSql & "FROM "
    strSql = strSql & gstrProgramaDeTrabalho & " PT, "
    strSql = strSql & gstrOrgao & " OG, "
    strSql = strSql & gstrUnidadeOrcamentaria & " UO, "
    strSql = strSql & gstrSubUnidade & " SU, "
    strSql = strSql & gstrTipoCredito & " TC, "
    strSql = strSql & gstrFuncaoDoGoverno & " FG, "
    strSql = strSql & gstrSubFuncaoGoverno & " SF, "
    strSql = strSql & gstrPrograma & " PG, "
    strSql = strSql & gstrSubPrograma & " SP, "
    strSql = strSql & gstrProjeto & " PJ, "
    strSql = strSql & gstrElementoDespesa & " ED "
'    strSql = strSql & "WHERE PT.intOrgao *= OG.PKId "
    strSql = strSql & "WHERE PT.intOrgao " & strOUTJSQLServer & "= OG.PKId " & strOUTJOracle
'    strSql = strSql & "AND PT.intUnidadeOrcamentaria *= UO.PKId "
    strSql = strSql & "AND PT.intUnidadeOrcamentaria " & strOUTJSQLServer & "= UO.PKId " & strOUTJOracle
'    strSql = strSql & "AND PT.intSubunidade *= SU.PKId "
    strSql = strSql & "AND PT.intSubunidade " & strOUTJSQLServer & "= SU.PKId " & strOUTJOracle
'    strSql = strSql & "AND PT.intTipoCredito *= TC.PKId "
    strSql = strSql & "AND PT.intTipoCredito " & strOUTJSQLServer & "= TC.PKId " & strOUTJOracle
'    strSql = strSql & "AND PT.intFuncao *= FG.PKId "
    strSql = strSql & "AND PT.intFuncao " & strOUTJSQLServer & "= FG.PKId " & strOUTJOracle
'    strSql = strSql & "AND PT.intSubFuncao *= SF.PKId "
    strSql = strSql & "AND PT.intSubFuncao " & strOUTJSQLServer & "= SF.PKId " & strOUTJOracle
'    strSql = strSql & "AND PT.intPrograma *= PG.PKId "
    strSql = strSql & "AND PT.intPrograma " & strOUTJSQLServer & "= PG.PKId " & strOUTJOracle
'    strSql = strSql & "AND PT.intSubPrograma *= SP.PKId "
    strSql = strSql & "AND PT.intSubPrograma " & strOUTJSQLServer & "= SP.PKId " & strOUTJOracle
'    strSql = strSql & "AND PT.intProjetoAtividade *= PJ.PKId "
    strSql = strSql & "AND PT.intProjetoAtividade " & strOUTJSQLServer & "= PJ.PKId " & strOUTJOracle
'    strSql = strSql & "AND PT.intElementoDespesa *= ED.PKId "
    strSql = strSql & "AND PT.intElementoDespesa " & strOUTJSQLServer & "= ED.PKId " & strOUTJOracle
    strSql = strSql & "AND bytSituacao = 1 "
    strSql = strSql & "AND PT.intExercicio = " & Val(txtintExercicio)
    strSql = strSql & " ORDER BY PT.intCodigoReduzido, PT.strCodigo, OG.strDescricao, "
    strSql = strSql & "UO.strDescricao, SU.strDescricao, TC.strDescricao, "
    strSql = strSql & "FG.strDescricao, SF.strDescricao, PG.strDescricao, "
    strSql = strSql & "SP.strDescricao, PJ.strDescricao, ED.strDescricao, PT.dblValor "
    strQueryRelatorio = strSql
End Function
