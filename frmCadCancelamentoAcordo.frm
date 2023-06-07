VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmCadCancelamentoAcordo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cancelamento de Acordo"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7470
   Icon            =   "frmCadCancelamentoAcordo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   7470
   Begin TabDlg.SSTab tab_3DPasta 
      Height          =   4875
      Left            =   120
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   90
      Width           =   7245
      _ExtentX        =   12779
      _ExtentY        =   8599
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Cancelamento de Acordo"
      TabPicture(0)   =   "frmCadCancelamentoAcordo.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl_dtmVencimento"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl_QtdeParcelasConsecutivas"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblAcordoInicial"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblAcordoFinal"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lbl_QtdeParcelasAlternadas"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "mskstrInscricaoFinal"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "mskstrInscricaoInicial"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "tdb_Acordos"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txt_dtmDataBase"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txt_QtdeParcelasConsecutivas"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtintExercicioInicial"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtintExercicioFinal"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txt_QtdeParcelasAlternadas"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).ControlCount=   13
      Begin VB.TextBox txt_QtdeParcelasAlternadas 
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
         Left            =   6210
         MaxLength       =   2
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   1560
         Width           =   705
      End
      Begin VB.TextBox txtintExercicioFinal 
         Height          =   285
         Left            =   6210
         MaxLength       =   4
         TabIndex        =   4
         Top             =   960
         Width           =   705
      End
      Begin VB.TextBox txtintExercicioInicial 
         Height          =   285
         Left            =   2850
         MaxLength       =   4
         TabIndex        =   2
         Top             =   960
         Width           =   705
      End
      Begin VB.TextBox txt_QtdeParcelasConsecutivas 
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
         Left            =   2850
         MaxLength       =   2
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   1530
         Width           =   705
      End
      Begin VB.TextBox txt_dtmDataBase 
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
         Left            =   1470
         MaxLength       =   10
         TabIndex        =   0
         Top             =   510
         Width           =   1245
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_Acordos 
         Height          =   2535
         Left            =   180
         TabIndex        =   7
         Top             =   2130
         Width           =   6885
         _ExtentX        =   12144
         _ExtentY        =   4471
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "PKID"
         Columns(0).DataField=   "pkid"
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Acordo"
         Columns(1).DataField=   "strinscricao"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Valor"
         Columns(2).DataField=   "dblvalor"
         Columns(2).NumberFormat=   "Standard"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   0
         Columns(3)._MaxComboItems=   5
         Columns(3).Caption=   "Vencimento"
         Columns(3).DataField=   "dtmdtVencimento"
         Columns(3).NumberFormat=   "FormatText Event"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(4)._VlistStyle=   0
         Columns(4)._MaxComboItems=   5
         Columns(4).DataField=   ""
         Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   5
         Splits(0)._UserFlags=   0
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   12632256
         Splits(0).FilterBar=   -1  'True
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=5"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=5186"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=5106"
         Splits(0)._ColumnProps(9)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(10)=   "Column(2).Width=3545"
         Splits(0)._ColumnProps(11)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(12)=   "Column(2)._WidthInPix=3466"
         Splits(0)._ColumnProps(13)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(14)=   "Column(3).Width=2725"
         Splits(0)._ColumnProps(15)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(16)=   "Column(3)._WidthInPix=2646"
         Splits(0)._ColumnProps(17)=   "Column(3).Order=4"
         Splits(0)._ColumnProps(18)=   "Column(4).Width=2725"
         Splits(0)._ColumnProps(19)=   "Column(4).DividerColor=0"
         Splits(0)._ColumnProps(20)=   "Column(4)._WidthInPix=2646"
         Splits(0)._ColumnProps(21)=   "Column(4).Order=5"
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=80,.bold=0,.fontsize=825,.italic=0"
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
         _StyleDefs(23)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
         _StyleDefs(24)  =   "Splits(0).EditorStyle:id=17,.parent=7"
         _StyleDefs(25)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
         _StyleDefs(26)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
         _StyleDefs(27)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
         _StyleDefs(28)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
         _StyleDefs(29)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
         _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
         _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
         _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
         _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
         _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=46,.parent=13"
         _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=43,.parent=14"
         _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=44,.parent=15"
         _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=45,.parent=17"
         _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=32,.parent=13"
         _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
         _StyleDefs(42)  =   "Splits(0).Columns(3).Style:id=50,.parent=13"
         _StyleDefs(43)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
         _StyleDefs(46)  =   "Splits(0).Columns(4).Style:id=54,.parent=13"
         _StyleDefs(47)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
         _StyleDefs(48)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
         _StyleDefs(49)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
         _StyleDefs(50)  =   "Named:id=33:Normal"
         _StyleDefs(51)  =   ":id=33,.parent=0"
         _StyleDefs(52)  =   "Named:id=34:Heading"
         _StyleDefs(53)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(54)  =   ":id=34,.wraptext=-1"
         _StyleDefs(55)  =   "Named:id=35:Footing"
         _StyleDefs(56)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(57)  =   "Named:id=36:Selected"
         _StyleDefs(58)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(59)  =   "Named:id=37:Caption"
         _StyleDefs(60)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(61)  =   "Named:id=38:HighlightRow"
         _StyleDefs(62)  =   ":id=38,.parent=33,.bgcolor=&H80000014&,.fgcolor=&H80000012&"
         _StyleDefs(63)  =   "Named:id=39:EvenRow"
         _StyleDefs(64)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(65)  =   "Named:id=40:OddRow"
         _StyleDefs(66)  =   ":id=40,.parent=33"
         _StyleDefs(67)  =   "Named:id=41:RecordSelector"
         _StyleDefs(68)  =   ":id=41,.parent=34"
         _StyleDefs(69)  =   "Named:id=42:FilterBar"
         _StyleDefs(70)  =   ":id=42,.parent=33"
      End
      Begin MSMask.MaskEdBox mskstrInscricaoInicial 
         Height          =   285
         Left            =   1470
         TabIndex        =   1
         Top             =   960
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   7
         Mask            =   "###.###"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskstrInscricaoFinal 
         Height          =   285
         Left            =   4830
         TabIndex        =   3
         Top             =   960
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   7
         Mask            =   "###.###"
         PromptChar      =   " "
      End
      Begin VB.Label lbl_QtdeParcelasAlternadas 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Quantidade Parcelas Alternadas"
         Height          =   195
         Left            =   3720
         TabIndex        =   13
         Top             =   1620
         Width           =   2280
      End
      Begin VB.Label lblAcordoFinal 
         AutoSize        =   -1  'True
         Caption         =   "Acordo Final"
         Height          =   195
         Left            =   3750
         TabIndex        =   12
         Top             =   1080
         Width           =   885
      End
      Begin VB.Label lblAcordoInicial 
         AutoSize        =   -1  'True
         Caption         =   "Acordo Inicial"
         Height          =   195
         Left            =   270
         TabIndex        =   11
         Top             =   1080
         Width           =   960
      End
      Begin VB.Label lbl_QtdeParcelasConsecutivas 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Quantidade Parcelas Consecutivas"
         Height          =   195
         Left            =   270
         TabIndex        =   10
         Top             =   1620
         Width           =   2490
      End
      Begin VB.Label lbl_dtmVencimento 
         AutoSize        =   -1  'True
         Caption         =   "Data Base"
         Height          =   195
         Left            =   270
         TabIndex        =   9
         Top             =   585
         Width           =   750
      End
   End
End
Attribute VB_Name = "frmCadCancelamentoAcordo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strInscricaoInicial         As String
Dim strInscricaoFinal           As String
Dim mblnAlterando               As Boolean

Private Sub Form_Activate()
    
    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrNovo
    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrLocalizar
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrImprimir

End Sub

Private Sub Form_Load()
    
    strInscricaoInicial = ""
    strInscricaoFinal = ""
    
End Sub

Private Sub mskstrInscricaoFinal_GotFocus()
    'If Len(Trim(mskstrInscricaoFinal)) = 0 And blnAutoNumeracao = True Then
    If Len(Trim(mskstrInscricaoFinal)) = 0 Then
        Screen.MousePointer = vbArrowHourglass
        'mskstrInscricaoFinal = ProximaInscricaoAcordo
        txtintExercicioFinal = Year(gstrDataDoSistema)
        Screen.MousePointer = vbDefault
    End If
    tab_3DPasta.Tab = 0
    MarcaCampo mskstrInscricaoFinal
End Sub

Private Sub mskstrInscricaoFinal_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", mskstrInscricaoFinal
End Sub

Private Sub mskstrInscricaoInicial_GotFocus()
'    If Len(Trim(mskstrInscricaoInicial)) = 0 And blnAutoNumeracao = True Then
    If Len(Trim(mskstrInscricaoInicial)) = 0 Then
        Screen.MousePointer = vbArrowHourglass
        'mskstrInscricaoInicial = ProximaInscricaoAcordo
        txtintExercicioInicial = Year(gstrDataDoSistema)
        Screen.MousePointer = vbDefault
    End If
    tab_3DPasta.Tab = 0
    MarcaCampo mskstrInscricaoInicial
End Sub

Private Sub mskstrInscricaoInicial_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", mskstrInscricaoInicial
End Sub


Private Sub txt_dtmDataBase_GotFocus()
    MarcaCampo txt_dtmDataBase
    If Len(txt_dtmDataBase) = 0 Then
        txt_dtmDataBase.Text = gstrDataDoSistema
    End If
End Sub

Private Sub txt_dtmDataBase_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txt_dtmDataBase
End Sub

Private Sub txt_dtmDataBase_LostFocus()
    txt_dtmDataBase = gstrDataFormatada(txt_dtmDataBase)
End Sub

Private Sub txt_QtdeParcelasAlternadas_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_QtdeParcelasAlternadas
End Sub

Private Sub txt_QtdeParcelasAlternadas_LostFocus()
    If Val(txt_QtdeParcelasAlternadas.Text) = 0 Then txt_QtdeParcelasAlternadas.Text = 0
    'PegaParcelasNaoPagas
End Sub

Private Sub txt_QtdeParcelasConsecutivas_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", txt_QtdeParcelasConsecutivas
End Sub

Private Sub txt_QtdeParcelasConsecutivas_LostFocus()
    If Val(txt_QtdeParcelasConsecutivas.Text) = 0 Then txt_QtdeParcelasConsecutivas.Text = 0
End Sub

Private Sub txtintExercicioFinal_GotFocus()
    MarcaCampo txtintExercicioFinal
End Sub

Private Sub txtintExercicioFinal_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintExercicioFinal
End Sub

Private Sub txtintExercicioFinal_LostFocus()
    
    If Len(Trim(Me.mskstrInscricaoInicial & Me.txtintExercicioInicial)) > 0 And Len(Trim(Me.mskstrInscricaoFinal & Me.txtintExercicioFinal)) = 0 Then
        Me.mskstrInscricaoFinal = Me.mskstrInscricaoInicial
        Me.txtintExercicioFinal = Me.txtintExercicioInicial
    End If
    
    If Me.mskstrInscricaoInicial & Me.txtintExercicioInicial.Text > Me.mskstrInscricaoFinal & Me.txtintExercicioFinal.Text Then
        MsgBox "Periodo entre Acordos invalidos.", 4096 + vbInformation, "Atenção"
        LocalLimpaMaskEdit
        Me.txtintExercicioInicial.Text = ""
        Me.txtintExercicioFinal.Text = ""
        Me.mskstrInscricaoInicial.SetFocus
        Exit Sub
    End If
    
End Sub

Private Sub txtintExercicioInicial_GotFocus()
    MarcaCampo txtintExercicioInicial
End Sub

Private Sub txtintExercicioInicial_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintExercicioInicial
End Sub
Sub PegaParcelasNaoPagas()
    Dim strSql                      As String
    Dim ADOTemp                     As New ADODB.Recordset
    
    Set gobjBanco = New clsBanco
    
    Me.MousePointer = 11
    
    strInscricaoInicial = Format(Me.mskstrInscricaoInicial & Me.txtintExercicioInicial, "00000000000000000000")
    
    strInscricaoFinal = Format(Me.mskstrInscricaoFinal & Me.txtintExercicioFinal, "00000000000000000000")
  
    'Monto a Query para Consulta
    strSql = ""
    strSql = "SELECT tbllancamentoalfa.pkid,tbllancamentoalfa.strinscricao,tbllancamentovalor.dblvalor,tbllancamentovalor.dtmdtVencimento "
    strSql = strSql & "FROM tbllancamentoalfa,tbllancamentovalor WHERE tbllancamentoalfa.strInscricao >= '" & strInscricaoInicial & "' AND tbllancamentoalfa.strInscricao <= '" & strInscricaoFinal & "' "
    strSql = strSql & "AND tbllancamentoalfa.pkid = tbllancamentovalor.intlancamentoalfa  AND tbllancamentoalfa.dtmdtcancelamento IS NULL AND tbllancamentovalor.dtmdtvencimento <= " & gstrConvDtParaSql(Me.txt_dtmDataBase, False) & " AND "
    strSql = strSql & "tbllancamentovalor.PkId NOT IN (SELECT intLancamentoValor FROM tblLancamentoPagamento)"
    
    If gobjBanco.CriaADO(strSql, 5, ADOTemp) Then
        If Not ADOTemp.EOF Then
            With ADOTemp
               'If Val(Me.txt_QtdeParcelasAlternadas.Text) >= .RecordCount Or Val(Me.txt_QtdeParcelasConsecutivas) >= .RecordCount Then
                If .RecordCount >= Val(Me.txt_QtdeParcelasAlternadas.Text) Or .RecordCount >= Val(Me.txt_QtdeParcelasConsecutivas) Then
                    'Prencho TDBGRID com as informações consultadas acima
                    Set tdb_Acordos.DataSource = ADOTemp
                    tdb_Acordos.Refresh
                Else
                    Set tdb_Acordos.DataSource = Nothing
                    tdb_Acordos.Refresh
                End If
            End With
        End If
    End If
    Me.MousePointer = 0
End Sub
Private Sub MarcaInadimplencia()
    Dim strSql              As String
    Dim lngPkId_Inicial     As Long
    Dim lngPkId_Final       As Long
    
    lngPkId_Inicial = Val(PkIdtblLancamentoAlfa(strInscricaoInicial))
    lngPkId_Final = Val(PkIdtblLancamentoAlfa(strInscricaoFinal))
    If Val(lngPkId_Inicial) > 0 And Val(lngPkId_Final) > 0 Then
        strSql = "UPDATE tblacordo SET dtmdtCancelamento = " & gstrConvDtParaSql(Me.txt_dtmDataBase.Text, False) & " "
        strSql = strSql & "WHERE intLancamentoAlfa >= " & lngPkId_Inicial & " AND intLancamentoAlfa <= " & lngPkId_Final & " "
        strSql = strSql & "AND dtmdtCancelamento IS NULL"
        If gblnExclusaoGravacaoOk("A", "", False) Then
            Set gobjBanco = New clsBanco
            gobjBanco.ExecutaBeginTrans
            
            Set gobjBanco = New clsBanco
            
            If Not gobjBanco.Execute(strSql, False) Then
                Set gobjBanco = New clsBanco
                gobjBanco.ExecutaRollbackTrans
            Else
                Set gobjBanco = New clsBanco
                gobjBanco.ExecutaCommitTrans
            End If
        End If
    Else
        MsgBox "Informe os Acordos a serem cancelados!", 4096 + vbInformation, " Atenção"
    End If
    
End Sub
Public Sub MantemForm(ByVal strModoOperacao As String)
    Dim varBookMark     As Variant
    Dim strSql          As String
    
    strSql = ""
    mblnAlterando = True
    
    If UCase(strModoOperacao) = UCase(gstrSalvar) Then
        If blnDadosOk Then
            MarcaInadimplencia
            Limpa_Controles Me, True, True, True, True, True
            LocalLimpaMaskEdit
        End If
    End If
    
    If UCase(strModoOperacao) = UCase(gstrLocalizar) Then
        PegaParcelasNaoPagas '-> Consulta as parcelas não pagas
    End If
   
    If UCase(strModoOperacao) = UCase(gstrNovo) Then
        Limpa_Controles Me, True, True, True, True, True
        LocalLimpaMaskEdit
        Set tdb_Acordos.DataSource = Nothing
        tdb_Acordos.Refresh
        Me.txt_dtmDataBase.SetFocus
    End If
        
End Sub
Private Function blnDadosOk() As Boolean
    
    blnDadosOk = True
    
    'Validadando a data base
    If Len(Trim(Me.txt_dtmDataBase.Text)) = 0 Then
        MsgBox "Informe a Data Base !", 4096 + vbInformation, "Aviso"
        blnDadosOk = False
        Me.txt_dtmDataBase.SetFocus
        Exit Function
    Else
        If CVDate(Me.txt_dtmDataBase.Text) = "" Then
            MsgBox "Data Base incorreta !", 4096 + vbInformation, "Atenção"
            blnDadosOk = False
            Me.txt_dtmDataBase.SetFocus
            Exit Function
        End If
    End If
    
End Function
Sub LocalLimpaMaskEdit()
    Dim strMasKaraMaskEdit      As String
    
    On Local Error Resume Next
    
    strMasKaraMaskEdit = Me.mskstrInscricaoInicial.Mask
    Me.mskstrInscricaoInicial.Mask = ""
    Me.mskstrInscricaoFinal.Mask = ""
    Me.mskstrInscricaoInicial.Text = ""
    Me.mskstrInscricaoFinal.Text = ""
    Me.mskstrInscricaoInicial.Mask = strMasKaraMaskEdit
    Me.mskstrInscricaoFinal.Mask = strMasKaraMaskEdit
    Set tdb_Acordos.DataSource = Nothing
    tdb_Acordos.Refresh
    
    
    Err.Clear
    
End Sub

Private Function PkIdtblLancamentoAlfa(strInscricaoParaConsultar As String) As String
Dim strSql      As String
Dim ADOTemp     As New ADODB.Recordset

    strSql = ""
    strSql = "SELECT PkId,strInscricao FROM tblLancamentoAlfa WHERE strInscricao = '" & strInscricaoParaConsultar & "'"
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, ADOTemp) Then
        If Not ADOTemp.EOF Then
            PkIdtblLancamentoAlfa = ADOTemp.Fields("PkId").Value
        End If
    End If
    ADOTemp.Close
    
End Function
'Private Function IsNullCampoDB(ByVal QualCampo As Variant) As Variant
'
'    Select Case bytDBType
'
'    Case EDatabases.SQLServer
'        IsNullCampoDB = " ISNULL("
'
'    Case EDatabases.Oracle
'
'        IsNullCampoDB = " NVL("
'
'    End Select
'
'    IsNullCampoDB = IsNullCampoDB & QualCampo & ", " & strTruePart & ") "
'
'End Function
