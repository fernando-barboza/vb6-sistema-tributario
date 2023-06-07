VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmDescontosProvisorios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Descontos Provisórios"
   ClientHeight    =   6135
   ClientLeft      =   1440
   ClientTop       =   2280
   ClientWidth     =   8670
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   8670
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   6015
      Left            =   90
      TabIndex        =   12
      Top             =   45
      Width           =   8490
      _ExtentX        =   14975
      _ExtentY        =   10610
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Descontos Provisórios"
      TabPicture(0)   =   "frmDescontosProvisorios.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "tdb_Lista"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fra_Descontos"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fra_Periodo"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fra_Condicoes"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtPKId"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      Begin VB.TextBox txtPKId 
         Height          =   270
         Left            =   2070
         TabIndex        =   26
         Top             =   0
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Frame fra_Condicoes 
         Caption         =   "Condições"
         Height          =   1275
         Left            =   135
         TabIndex        =   22
         Top             =   2385
         Width           =   8205
         Begin VB.TextBox txtintParcelaInicialJuros 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   6690
            MaxLength       =   4
            TabIndex        =   11
            Top             =   540
            Width           =   1245
         End
         Begin VB.TextBox txtdblJurosParcelamento 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   6690
            TabIndex        =   10
            Top             =   225
            Width           =   1245
         End
         Begin VB.TextBox txtdtmdtVencimento 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2370
            TabIndex        =   9
            Top             =   855
            Width           =   1350
         End
         Begin VB.TextBox txtdblValorMinimo 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2370
            TabIndex        =   8
            Top             =   540
            Width           =   2685
         End
         Begin VB.TextBox txtintParcela 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2370
            MaxLength       =   3
            TabIndex        =   7
            Top             =   225
            Width           =   630
         End
         Begin VB.Label Label14 
            Caption         =   "%"
            Height          =   225
            Left            =   7980
            TabIndex        =   33
            Top             =   300
            Width           =   195
         End
         Begin VB.Label lblintParcelaInicialJuros 
            Caption         =   "A partir da parcela"
            Height          =   285
            Left            =   5310
            TabIndex        =   32
            Top             =   570
            Width           =   1785
         End
         Begin VB.Label lblinJjurosParcelamento 
            Caption         =   "Juros mensais de"
            Height          =   225
            Left            =   5370
            TabIndex        =   31
            Top             =   270
            Width           =   1305
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Vencimento até"
            Height          =   195
            Left            =   180
            TabIndex        =   25
            Top             =   900
            Width           =   1110
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Valor minímo por parcela"
            Height          =   195
            Left            =   180
            TabIndex        =   24
            Top             =   585
            Width           =   1755
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Quantidade max. de parcelas"
            Height          =   195
            Left            =   180
            TabIndex        =   23
            Top             =   270
            Width           =   2070
         End
      End
      Begin VB.Frame fra_Periodo 
         Caption         =   "Período"
         Height          =   1275
         Left            =   135
         TabIndex        =   17
         Top             =   405
         Width           =   8205
         Begin VB.TextBox txtstrLegislacao 
            Height          =   285
            Left            =   1065
            MaxLength       =   50
            TabIndex        =   3
            Top             =   855
            Width           =   7020
         End
         Begin VB.TextBox txtstrDescricao 
            Height          =   285
            Left            =   1065
            MaxLength       =   70
            TabIndex        =   2
            Top             =   540
            Width           =   7020
         End
         Begin VB.TextBox txtdtmdtInicial 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1065
            TabIndex        =   0
            Top             =   225
            Width           =   1305
         End
         Begin VB.TextBox txtdtmdtFinal 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   3330
            TabIndex        =   1
            Top             =   225
            Width           =   1305
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Data inicial"
            Height          =   195
            Left            =   135
            TabIndex        =   21
            Top             =   270
            Width           =   780
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Legislação"
            Height          =   195
            Left            =   135
            TabIndex        =   20
            Top             =   900
            Width           =   765
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Descrição"
            Height          =   195
            Left            =   135
            TabIndex        =   19
            Top             =   585
            Width           =   720
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Data final"
            Height          =   195
            Left            =   2580
            TabIndex        =   18
            Top             =   270
            Width           =   675
         End
      End
      Begin VB.Frame fra_Descontos 
         Caption         =   "Descontos"
         Height          =   645
         Left            =   135
         TabIndex        =   13
         Top             =   1710
         Width           =   8205
         Begin VB.TextBox txtdblJuros 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   6885
            TabIndex        =   6
            Top             =   225
            Width           =   1035
         End
         Begin VB.TextBox txtdblMulta 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   4320
            TabIndex        =   5
            Top             =   225
            Width           =   1035
         End
         Begin VB.TextBox txtdblValorOriginal 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1800
            TabIndex        =   4
            Top             =   225
            Width           =   1035
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "%"
            Height          =   195
            Left            =   7965
            TabIndex        =   29
            Top             =   270
            Width           =   120
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "%"
            Height          =   195
            Left            =   5400
            TabIndex        =   28
            Top             =   270
            Width           =   120
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "%"
            Height          =   195
            Left            =   2880
            TabIndex        =   27
            Top             =   270
            Width           =   120
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Juros desconto"
            Height          =   195
            Left            =   5715
            TabIndex        =   16
            Top             =   270
            Width           =   1080
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Multa desconto"
            Height          =   195
            Left            =   3150
            TabIndex        =   15
            Top             =   270
            Width           =   1095
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Valor original desconto"
            Height          =   195
            Left            =   135
            TabIndex        =   14
            Top             =   270
            Width           =   1605
         End
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_Lista 
         Height          =   2145
         Left            =   135
         TabIndex        =   30
         Top             =   3750
         Width           =   8190
         _ExtentX        =   14446
         _ExtentY        =   3784
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
         Columns(1).Caption=   "Data inicial"
         Columns(1).DataField=   "dtmdtInicial"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Data final"
         Columns(2).DataField=   "dtmdtFinal"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Descrição"
         Columns(3).DataField=   "strDescricao"
         Columns(3).NumberFormat=   "Standard"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   4
         Splits(0)._UserFlags=   0
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
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
         Splits(0)._ColumnProps(8)=   "Column(1).Width=1984"
         Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1905"
         Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=2"
         Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(14)=   "Column(2).Width=1958"
         Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=1879"
         Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=2"
         Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(20)=   "Column(3).Width=10345"
         Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=10266"
         Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=0"
         Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
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
         TabAction       =   1
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
         _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=46,.parent=13"
         _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
         _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
         _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
         _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=13,.alignment=1"
         _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
         _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
         _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
         _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=13,.alignment=1"
         _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
         _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13,.alignment=0"
         _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
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
         _StyleDefs(64)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
         _StyleDefs(65)  =   "Named:id=39:EvenRow"
         _StyleDefs(66)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(67)  =   "Named:id=40:OddRow"
         _StyleDefs(68)  =   ":id=40,.parent=33"
         _StyleDefs(69)  =   "Named:id=41:RecordSelector"
         _StyleDefs(70)  =   ":id=41,.parent=34"
         _StyleDefs(71)  =   "Named:id=42:FilterBar"
         _StyleDefs(72)  =   ":id=42,.parent=33"
      End
   End
End
Attribute VB_Name = "frmDescontosProvisorios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mblnAlterando       As Boolean
    Dim mobjAux             As Object
    Dim mblnPrimeiraVez     As Boolean
    Dim blnOrdenacaoAsc     As Boolean
    Dim bytOrdenacao        As Byte
    Dim mblnSelecionou      As Boolean

Private Sub tdb_Lista_Click()
    mblnPrimeiraVez = True
    If glngQtdLinhaTDBGrid(tdb_Lista) = 1 Then
        tdb_Lista_RowColChange 0, 0
    End If
End Sub

Private Sub tdb_Lista_FilterChange()
    mblnPrimeiraVez = False
    gblnFilraCampos tdb_Lista
End Sub

Private Sub tdb_Lista_HeadClick(ByVal ColIndex As Integer)
    gOrdenaGrid tdb_Lista, ColIndex
End Sub

Private Sub tdb_Lista_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    
    With tdb_Lista
        If Not .EOF And Not .BOF Then
            txtPKId.Text = .Columns("PKID").Value
            If mblnPrimeiraVez Then
                mblnAlterando = True
                LeDaTabelaParaObj gstrDescontosProvisorios, Me
                
                
                
                gCorLinhaSelecionada tdb_Lista
                HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
                
                If mobjAux Is Nothing Then
                    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
                Else
                    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
                End If
                mblnSelecionou = True
            End If
        End If
    End With

End Sub

Private Sub txtdblJuros_GotFocus()
    MarcaCampo txtdblJuros
End Sub

Private Sub txtdblJuros_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txtdblJuros
End Sub

Private Sub txtdblMulta_GotFocus()
    MarcaCampo txtdblMulta
End Sub

Private Sub txtdblMulta_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txtdblMulta
End Sub

Private Sub txtdblValorMinimo_GotFocus()
    MarcaCampo txtdblValorMinimo
End Sub

Private Sub txtdblValorMinimo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txtdblValorMinimo
End Sub

Private Sub txtdblValorOriginal_GotFocus()
    MarcaCampo txtdblValorOriginal
End Sub

Private Sub txtdblValorOriginal_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txtdblValorOriginal
End Sub

Private Sub txtdtmdtFinal_GotFocus()
    MarcaCampo txtdtmdtFinal
End Sub

Private Sub txtdtmdtFinal_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtdtmdtFinal
End Sub

Private Sub txtdtmdtFinal_LostFocus()
    txtdtmdtFinal.Text = gstrDataFormatada(txtdtmdtFinal.Text)
End Sub

Private Sub txtdtmdtInicial_GotFocus()
    MarcaCampo txtDtmDtInicial
End Sub

Private Sub txtdtmdtInicial_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtDtmDtInicial
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
Dim strSql As String

    strSql = "SELECT * FROM " & gstrDescontosProvisorios

    If UCase(strModoOperacao) = "SALVAR" Then
       mblnPrimeiraVez = False
       If Not blnDadosOk Then
          Exit Sub
       End If
    End If
     
    If UCase(strModoOperacao) = "DELETAR" Then
       mblnPrimeiraVez = False
    End If
    
    'If UCase(strModoOperacao) = "LOCALIZAR" Then
    '   If Trim(txtintCodigoDaTestada.Text) = "" And Trim(txtstrNomeDaTestada.Text) = "" Then
    '      txtPKId.Text = ""
    '   End If
    'End If
    
    ToolBarGeral strModoOperacao, gstrDescontosProvisorios, mblnAlterando, tdb_Lista, Me, mobjAux, strSql, strSql
    
    If (UCase(strModoOperacao) = "SALVAR" Or UCase(strModoOperacao) = "DELETAR") And _
        gblnCancelarInclusao = False Then
        LeDaTabelaParaObj gstrDescontosProvisorios, tdb_Lista, strSql
        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
        Limpa_Controles frmDescontosProvisorios, True, False, False, False, False
        txtDtmDtInicial.SetFocus
    End If
    
    If UCase(strModoOperacao) = "NOVO" Then
       HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
       Limpa_Controles frmDescontosProvisorios, True, False, False, False, False
       txtDtmDtInicial.SetFocus
    End If
  
End Sub

Private Function blnDadosOk() As Boolean
    If Not gblnDataValida(txtDtmDtInicial.Text) Then
       ExibeMensagem "A data inicial informada não é válida."
       txtDtmDtInicial.SetFocus
       Exit Function
    ElseIf Not gblnDataValida(txtdtmdtFinal.Text) Then
       ExibeMensagem "A data final informada não é válida."
       txtdtmdtFinal.SetFocus
       Exit Function
    ElseIf CDate(txtDtmDtInicial.Text) > CDate(txtdtmdtFinal.Text) Then
       ExibeMensagem "A data inicial deve ser menor que a data final."
       txtDtmDtInicial.SetFocus
       Exit Function
    ElseIf Trim(txtstrdescricao.Text) = "" Then
       ExibeMensagem "A descrição deve ser informada."
       txtstrdescricao.SetFocus
       Exit Function
    ElseIf Trim(txtstrLegislacao.Text) = "" Then
       ExibeMensagem "A legislação deve ser informada."
       txtstrLegislacao.SetFocus
       Exit Function
    ElseIf Val(txtintParcela.Text) <= 0 Then
       ExibeMensagem "A quantidade máxima de parcelas deve ser informada."
       txtintParcela.SetFocus
       Exit Function
    ElseIf Not gblnDataValida(txtdtmDtVencimento.Text) Then
       ExibeMensagem "A data de vencimento deve ser informada."
       txtdtmDtVencimento.SetFocus
       Exit Function
    End If
    
    blnDadosOk = True
End Function

Private Sub Form_Activate()
    gintCodSeguranca = 1299
    If mblnSelecionou Then
        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
    Else
        HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
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
    bytOrdenacao = 2: blnOrdenacaoAsc = True
    'VerificaListaAutomatica gstrTipoDeTestada, tdb_Lista, "PKId, intCodigoDaTestada, strNomeDaTestada"
    VerificaObjParaAplicar mobjAux
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar, gstrDeletar
    mblnSelecionou = False
    mblnPrimeiraVez = False
End Sub

Private Sub txtdtmdtInicial_LostFocus()
    txtDtmDtInicial.Text = gstrDataFormatada(txtDtmDtInicial.Text)
End Sub

Private Sub txtdtmdtVencimento_GotFocus()
    MarcaCampo txtdtmDtVencimento
End Sub

Private Sub txtdtmdtVencimento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txtdtmDtVencimento
End Sub

Private Sub txtdtmdtVencimento_LostFocus()
    txtdtmDtVencimento.Text = gstrDataFormatada(txtdtmDtVencimento.Text)
End Sub

Private Sub txtdblJurosParcelamento_GotFocus()
    MarcaCampo txtdblJurosParcelamento
End Sub

Private Sub txtdblJurosParcelamento_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txtdblJurosParcelamento
End Sub

Private Sub txtdblJurosParcelamento_LostFocus()
    txtdblJurosParcelamento.Text = gstrConvVrDoSql(txtdblJurosParcelamento.Text, 2)
End Sub

Private Sub txtintParcelaInicialJuros_GotFocus()
    MarcaCampo txtintParcelaInicialJuros
End Sub

Private Sub txtintParcelaInicialJuros_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintParcelaInicialJuros
End Sub

Private Sub txtintParcela_GotFocus()
    MarcaCampo txtintParcela
End Sub

Private Sub txtintParcela_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintParcela
End Sub

Private Sub txtstrDescricao_GotFocus()
    MarcaCampo txtstrdescricao
End Sub

Private Sub txtstrLegislacao_GotFocus()
    MarcaCampo txtstrLegislacao
End Sub
