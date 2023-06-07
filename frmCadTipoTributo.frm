VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmCadTipoTributo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro Tipo de Tributos"
   ClientHeight    =   3735
   ClientLeft      =   1020
   ClientTop       =   2610
   ClientWidth     =   8805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   8805
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   3690
      Left            =   30
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   30
      Width           =   8745
      _ExtentX        =   15425
      _ExtentY        =   6509
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tipos de Tributo"
      TabPicture(0)   =   "frmCadTipoTributo.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblstrDescricao"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "dbcintReceita"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "tdb_ListaTipoTributo"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtPKId"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txt_strTipoReceita"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "optbytTipo(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "optbytTipo(1)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "optbytTipo(2)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "optbytTipo(3)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtstrDescricao"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "cmd_Receita"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "optbytTipo(4)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      Begin VB.OptionButton optbytTipo 
         Caption         =   "Horário Especial"
         Height          =   330
         Index           =   4
         Left            =   6240
         TabIndex        =   14
         Top             =   1425
         Width           =   1485
      End
      Begin VB.CommandButton cmd_Receita 
         Height          =   300
         Left            =   8310
         Picture         =   "frmCadTipoTributo.frx":001C
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         Tag             =   "438"
         ToolTipText     =   "Ativa Cadastro de Grupo de Assunto"
         Top             =   630
         Width           =   330
      End
      Begin VB.TextBox txtstrDescricao 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   900
         MaxLength       =   40
         TabIndex        =   8
         Top             =   1020
         Width           =   7725
      End
      Begin VB.OptionButton optbytTipo 
         Caption         =   "Outros"
         Height          =   330
         Index           =   3
         Left            =   7770
         TabIndex        =   7
         Top             =   1425
         Width           =   765
      End
      Begin VB.OptionButton optbytTipo 
         Caption         =   "Ocupação"
         Height          =   330
         Index           =   2
         Left            =   5040
         TabIndex        =   6
         Top             =   1410
         Width           =   1230
      End
      Begin VB.OptionButton optbytTipo 
         Caption         =   "Feiras"
         Height          =   330
         Index           =   1
         Left            =   4245
         TabIndex        =   5
         Top             =   1410
         Width           =   780
      End
      Begin VB.OptionButton optbytTipo 
         Caption         =   "Publicidade"
         Height          =   330
         Index           =   0
         Left            =   3060
         TabIndex        =   4
         Top             =   1410
         Value           =   -1  'True
         Width           =   1155
      End
      Begin VB.TextBox txt_strTipoReceita 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   900
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1425
         Width           =   2100
      End
      Begin VB.TextBox txtPKId 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         Locked          =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   2
         Top             =   0
         Visible         =   0   'False
         Width           =   855
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_ListaTipoTributo 
         Height          =   1755
         Left            =   60
         TabIndex        =   1
         Top             =   1860
         Width           =   8580
         _ExtentX        =   15134
         _ExtentY        =   3096
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
         Columns(1).Caption=   "Descrição"
         Columns(1).DataField=   "Descr"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Descrição Receita"
         Columns(2).DataField=   "DescrRec"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   16
         Columns(3)._MaxComboItems=   5
         Columns(3).ValueItems(0)._DefaultItem=   0
         Columns(3).ValueItems(0).Value=   "1"
         Columns(3).ValueItems(0).Value.vt=   8
         Columns(3).ValueItems(0).DisplayValue=   "Convênio"
         Columns(3).ValueItems(0).DisplayValue.vt=   8
         Columns(3).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
         Columns(3).ValueItems(1)._DefaultItem=   0
         Columns(3).ValueItems(1).Value=   "2"
         Columns(3).ValueItems(1).Value.vt=   8
         Columns(3).ValueItems(1).DisplayValue=   "Imposto"
         Columns(3).ValueItems(1).DisplayValue.vt=   8
         Columns(3).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
         Columns(3).ValueItems(2)._DefaultItem=   0
         Columns(3).ValueItems(2).Value=   "3"
         Columns(3).ValueItems(2).Value.vt=   8
         Columns(3).ValueItems(2).DisplayValue=   "Taxa"
         Columns(3).ValueItems(2).DisplayValue.vt=   8
         Columns(3).ValueItems(2)._PropDict=   "_DefaultItem,517,2"
         Columns(3).ValueItems(3)._DefaultItem=   0
         Columns(3).ValueItems(3).Value=   "4"
         Columns(3).ValueItems(3).Value.vt=   8
         Columns(3).ValueItems(3).DisplayValue=   "Tarifa"
         Columns(3).ValueItems(3).DisplayValue.vt=   8
         Columns(3).ValueItems(3)._PropDict=   "_DefaultItem,517,2"
         Columns(3).ValueItems(4)._DefaultItem=   0
         Columns(3).ValueItems(4).Value=   "5"
         Columns(3).ValueItems(4).Value.vt=   8
         Columns(3).ValueItems(4).DisplayValue=   "Repasse Governamental"
         Columns(3).ValueItems(4).DisplayValue.vt=   8
         Columns(3).ValueItems(4)._PropDict=   "_DefaultItem,517,2"
         Columns(3).ValueItems(5)._DefaultItem=   0
         Columns(3).ValueItems(5).Value=   "6"
         Columns(3).ValueItems(5).Value.vt=   8
         Columns(3).ValueItems(5).DisplayValue=   "Repasse não Governamental"
         Columns(3).ValueItems(5).DisplayValue.vt=   8
         Columns(3).ValueItems(5)._PropDict=   "_DefaultItem,517,2"
         Columns(3).ValueItems.Count=   6
         Columns(3).Caption=   "Tipo da Receita"
         Columns(3).DataField=   "Tipo"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   4
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
         Splits(0)._ColumnProps(0)=   "Columns.Count=4"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).AllowSizing=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(8)=   "Column(1).Width=6271"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=6191"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(13)=   "Column(2).Width=4524"
         Splits(0)._ColumnProps(14)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(15)=   "Column(2)._WidthInPix=4445"
         Splits(0)._ColumnProps(16)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(18)=   "Column(3).Width=2805"
         Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=2725"
         Splits(0)._ColumnProps(21)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
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
         _StyleDefs(30)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
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
         _StyleDefs(41)  =   "Splits(0).Columns(1).Style:id=54,.parent=13"
         _StyleDefs(42)  =   "Splits(0).Columns(1).HeadingStyle:id=51,.parent=14"
         _StyleDefs(43)  =   "Splits(0).Columns(1).FooterStyle:id=52,.parent=15"
         _StyleDefs(44)  =   "Splits(0).Columns(1).EditorStyle:id=53,.parent=17"
         _StyleDefs(45)  =   "Splits(0).Columns(2).Style:id=50,.parent=13"
         _StyleDefs(46)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
         _StyleDefs(47)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
         _StyleDefs(48)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
         _StyleDefs(49)  =   "Splits(0).Columns(3).Style:id=32,.parent=13"
         _StyleDefs(50)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
         _StyleDefs(51)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
         _StyleDefs(52)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
         _StyleDefs(53)  =   "Named:id=33:Normal"
         _StyleDefs(54)  =   ":id=33,.parent=0"
         _StyleDefs(55)  =   "Named:id=34:Heading"
         _StyleDefs(56)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(57)  =   ":id=34,.wraptext=-1"
         _StyleDefs(58)  =   "Named:id=35:Footing"
         _StyleDefs(59)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(60)  =   "Named:id=36:Selected"
         _StyleDefs(61)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(62)  =   "Named:id=37:Caption"
         _StyleDefs(63)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(64)  =   "Named:id=38:HighlightRow"
         _StyleDefs(65)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(66)  =   "Named:id=39:EvenRow"
         _StyleDefs(67)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(68)  =   "Named:id=40:OddRow"
         _StyleDefs(69)  =   ":id=40,.parent=33"
         _StyleDefs(70)  =   "Named:id=41:RecordSelector"
         _StyleDefs(71)  =   ":id=41,.parent=34"
         _StyleDefs(72)  =   "Named:id=42:FilterBar"
         _StyleDefs(73)  =   ":id=42,.parent=33"
      End
      Begin MSDataListLib.DataCombo dbcintReceita 
         Height          =   315
         Left            =   900
         TabIndex        =   10
         Top             =   630
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin VB.Label lblstrDescricao 
         AutoSize        =   -1  'True
         Caption         =   "Descrição "
         Height          =   195
         Left            =   90
         TabIndex        =   13
         Top             =   1080
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descrição:"
         Height          =   195
         Left            =   90
         TabIndex        =   12
         Top             =   705
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
         Height          =   195
         Left            =   90
         TabIndex        =   11
         Top             =   1425
         Width           =   360
      End
   End
End
Attribute VB_Name = "frmCadTipoTributo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mobjAux         As Object
Dim blnOrdenacaoAsc As Boolean
Dim bytOrdenacao    As Byte
Dim mblnSelecionou  As Boolean
Dim mblnAlterando   As Boolean
Dim mobjLista       As Boolean
Dim strSql          As String
Dim adoResultado    As ADODB.Recordset
Dim mblnClickOk    As Boolean
    
'************************************************************************************************
'OBS: Foram criadas constantes para os Tipos de Tributos no ModTributario - TIPO_PUBLICIDADE ...
'************************************************************************************************

Private Function strQueryRelatorio() As String
    
    strSql = ""
    strSql = strSql & "select TT.strdescricao DescricaoTributo , RE.strdescricao DescricaoReceita, "
    strSql = strSql & gstrCASEWHEN("RE.Byttipo", "1,'Convênio',2,'Imposto',3,'Taxa',4,'Tarifa',5,'Repasse Governamental',6,'Repasse não Governamental'") & "TipoReceita "
    strSql = strSql & " FROM " & gstrTributoTipo & " TT, " & gstrReceita & " RE"
    strSql = strSql & " ORDER BY "
    strSql = strSql & "TT.strdescricao"
    
    strQueryRelatorio = strSql
    
End Function

Private Function DescricaoTipo(bytTipo As Integer) As String
    Select Case bytTipo
        Case 1
            DescricaoTipo = "Convênio"
        Case 2
            DescricaoTipo = "Imposto"
        Case 3
            DescricaoTipo = "Taxa"
        Case 4
            DescricaoTipo = "Tarifa"
        Case 5
            DescricaoTipo = "Repasse Governamental"
        Case 6
            DescricaoTipo = "Repasse não Governamental"
    End Select
End Function
    
Private Sub cmd_Receita_Click()
    CarregaForm frmCadReceita
End Sub

Private Sub dbcintReceita_Click(Area As Integer)
    If Area = 2 Then
        If dbcintReceita.MatchedWithList Then
        Set gobjBanco = New clsBanco
            If gobjBanco.CriaADO(strQueryTipo, 5, adoResultado) Then
                If Not adoResultado.EOF Then txt_strTipoReceita.Text = DescricaoTipo(gstrENulo(adoResultado!bytTipo))
           End If
        End If
    End If
End Sub

Private Function strQueryTipo() As String
    strSql = ""
    strSql = strSql & "SELECT "
    strSql = strSql & "Re.pkid, "
    strSql = strSql & "RE.bytTipo "
    strSql = strSql & "FROM "
    strSql = strSql & gstrReceita & " RE "
    strSql = strSql & "WHERE "
    strSql = strSql & "RE.Pkid = " & dbcintReceita.BoundText
    strQueryTipo = strSql
End Function

Private Sub dbcintReceita_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintReceita, Me, , KeyCode, Shift
End Sub

Private Sub dbcintReceita_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintReceita
End Sub

Private Sub dbcintReceita_GotFocus()
    MarcaCampo dbcintReceita
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 1173
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
    mblnAlterando = False
    bytOrdenacao = 2: blnOrdenacaoAsc = True
    dbcintReceita.Tag = strQueryReceita & ";strDescricao"
    VerificaObjParaAplicar mobjAux
    TrocaCorObjeto txt_strTipoReceita, True, True
End Sub

Private Function strQueryReceita() As String
Dim strSql As String

    strSql = ""
    strSql = strSql & "SELECT "
    strSql = strSql & "Pkid, "
    strSql = strSql & "strDescricao "
    strSql = strSql & "FROM "
    strSql = strSql & gstrReceita
    
    strQueryReceita = strSql

End Function

Private Function strQuery() As String
    
    strSql = ""
    strSql = strSql & "Select TT.pkid, "
    strSql = strSql & "TT.strdescricao Descr, "
    strSql = strSql & "RE.strdescricao DescrRec, "
    strSql = strSql & "RE.bytTipo Tipo "
    strSql = strSql & "from "
    strSql = strSql & gstrTributoTipo & " TT, "
    strSql = strSql & gstrReceita & " RE "
    strSql = strSql & "Where "
    strSql = strSql & "TT.INTRECEITA = Re.Pkid "
    
    Select Case bytOrdenacao
        Case 0
            strSql = strSql & "ORDER BY TT.strdescricao " & IIf(blnOrdenacaoAsc, " ASC", " DESC")
        Case 1
            strSql = strSql & "ORDER BY RE.strDescricao " & IIf(blnOrdenacaoAsc, " ASC", " DESC")
        Case 2
            strSql = strSql & "ORDER BY RE.bytTipo " & IIf(blnOrdenacaoAsc, " ASC", " DESC")
    End Select
    
    strQuery = strSql

End Function

Private Sub tab_3dPasta_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
End Sub

Private Function strQueryTipoTributo()
Dim strSql As String
    
    strSql = ""
    strSql = strSql & "SELECT PKID, strdescricao "
    strSql = strSql & "FROM "
    strSql = strSql & gstrTributoTipo
    strQueryTipoTributo = strSql

End Function

Private Function strQueryAplicar()
Dim strSql As String
    
    strSql = ""
    strSql = strSql & "SELECT PKID, strdescricao "
    strSql = strSql & "FROM "
    strSql = strSql & gstrTributoTipo
    strQueryAplicar = strSql

End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrImprimir
End Sub

Private Sub tdb_ListaTipoTributo_Click()
    mblnClickOk = True
End Sub

Private Sub tdb_ListaTipoTributo_DblClick()
    MantemForm gstrAplicar
End Sub

Private Sub tdb_ListaTipoTributo_FilterChange()
    mblnClickOk = False
    gblnFilraCampos tdb_ListaTipoTributo
End Sub

Private Sub tdb_ListaTipoTributo_HeadClick(ByVal ColIndex As Integer)
    gOrdenaGrid tdb_ListaTipoTributo, ColIndex
End Sub

Private Sub tdb_ListaTipoTributo_KeyDown(KeyCode As Integer, Shift As Integer)
    mblnClickOk = True
End Sub

Private Sub tdb_ListaTipoTributo_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    With tdb_ListaTipoTributo
    If (Not .EOF And Not .BOF) And mblnClickOk Then
        mblnClickOk = False
        txtPKId.Text = .Columns("PKID").Value
        txt_strTipoReceita.Text = DescricaoTipo(gstrENulo(.Columns("Tipo").Value))
        LeDaTabelaParaObj gstrTributoTipo, Me
        gCorLinhaSelecionada tdb_ListaTipoTributo
        If mobjAux Is Nothing Then
            HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
        Else
            HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
        End If
        mblnSelecionou = True
        mblnAlterando = True
    End If
    End With
End Sub

Private Sub txtstrDescricao_GotFocus()
    MarcaCampo txtstrDescricao
End Sub

Private Sub txtstrDescricao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txtstrDescricao
End Sub

Private Function blnDadosOk() As Boolean
    
    blnDadosOk = False
    
    If txtstrDescricao.Text = Empty Then
        ExibeMensagem "Descrição do Tipo de Tributo Inválida"
        txtstrDescricao.SetFocus
        Exit Function
    ElseIf optbytTipo(0).Value = False And optbytTipo(1).Value = False And optbytTipo(2).Value = False And optbytTipo(3).Value = False And optbytTipo(4).Value = False Then
        ExibeMensagem "Selecione um Tipo Válido."
        optbytTipo(0).SetFocus
        Exit Function
    ElseIf Not dbcintReceita.MatchedWithList Then
        ExibeMensagem "Descrição da Receita Inválida"
        dbcintReceita.SetFocus
        Exit Function
    End If
    
    blnDadosOk = True
    
End Function

Public Sub MantemForm(ByVal strModoOperacao As String)
strSql = strQuery
    If strModoOperacao = gstrNovo Then
        mblnSelecionou = False
        mblnAlterando = False
        LimpaObjeto Me
        txtstrDescricao.Text = ""
        dbcintReceita.Text = ""
        txt_strTipoReceita.Text = ""
    Else
        If UCase(strModoOperacao) = "SALVAR" Or UCase(strModoOperacao) = "DELETAR" Then
            If Not blnDadosOk Then
                Exit Sub
            Else
                strSql = strQueryTipoTributo
                ToolBarGeral strModoOperacao, gstrTributoTipo, mblnAlterando, tdb_ListaTipoTributo, _
                             Me, mobjAux, strSql, strQueryTipoTributo
                LimpaObjeto Me
                txtstrDescricao.Text = ""
                dbcintReceita.Text = ""
                txt_strTipoReceita.Text = ""
                Exit Sub
            End If
        Else
            If UCase(strModoOperacao) = gstrPreencherLista Then
                PreencherListaDeOpcoes Me.ActiveControl
                Exit Sub
            End If
        End If
    End If
    If strModoOperacao = UCase(gstrImprimir) Then
        ImprimeRelatorio rptTipoTributo, strQueryRelatorio
        Exit Sub
    End If
    ToolBarGeral strModoOperacao, gstrTributoTipo, mblnAlterando, tdb_ListaTipoTributo, Me, mobjAux, strSql, strQueryAplicar

End Sub


