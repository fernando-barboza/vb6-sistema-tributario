VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmCadValoresDasFaixas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Faixa de Valores"
   ClientHeight    =   5745
   ClientLeft      =   4005
   ClientTop       =   4605
   ClientWidth     =   6360
   HelpContextID   =   27
   Icon            =   "CadValoresDasFaixas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   6360
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   5610
      Left            =   90
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   90
      Width           =   6225
      _ExtentX        =   10980
      _ExtentY        =   9895
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Valores"
      TabPicture(0)   =   "CadValoresDasFaixas.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblintFaixaDeValores"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblintTabelaDeValores"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblintReferenciaTributo"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "dbcintReferenciaTributo"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "dbcintFaixaDeValores"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "dbcintTabelasDeValores"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "grd_Valores"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmd_FaixaDeValores"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      Begin VB.CommandButton cmd_FaixaDeValores 
         Height          =   315
         Left            =   5700
         Picture         =   "CadValoresDasFaixas.frx":105E
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         Tag             =   "594"
         ToolTipText     =   "Ativa Cadastro de Faixa de Valores"
         Top             =   600
         Width           =   360
      End
      Begin TrueOleDBGrid70.TDBGrid grd_Valores 
         Height          =   3570
         Left            =   150
         Negotiate       =   -1  'True
         TabIndex        =   8
         Top             =   1890
         Width           =   5925
         _ExtentX        =   10451
         _ExtentY        =   6297
         _LayoutType     =   4
         _RowHeight      =   -2147483647
         _WasPersistedAsPixels=   0
         Columns(0)._VlistStyle=   0
         Columns(0)._MaxComboItems=   5
         Columns(0).Caption=   "Faixa Inicial"
         Columns(0).DataField=   ""
         Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(1)._VlistStyle=   0
         Columns(1)._MaxComboItems=   5
         Columns(1).Caption=   "Faixa Final"
         Columns(1).DataField=   ""
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "Valor"
         Columns(2).DataField=   ""
         Columns(2).DropDown=   "tdd_Valores"
         Columns(2).DropDown.vt=   8
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   3
         Splits(0)._UserFlags=   0
         Splits(0).MarqueeStyle=   2
         Splits(0).RecordSelectors=   0   'False
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=3"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2963"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2884"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(6)=   "Column(1).Width=2937"
         Splits(0)._ColumnProps(7)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(8)=   "Column(1)._WidthInPix=2858"
         Splits(0)._ColumnProps(9)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(10)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(11)=   "Column(2).Width=4022"
         Splits(0)._ColumnProps(12)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(13)=   "Column(2)._WidthInPix=3942"
         Splits(0)._ColumnProps(14)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(15)=   "Column(2).Order=3"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   3
         PrintInfos(0).Name=   "piInternal 0"
         PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
         PrintInfos(0).PageHeaderHeight=   0
         PrintInfos(0).PageFooterHeight=   0
         PrintInfos.Count=   1
         AllowDelete     =   -1  'True
         AllowAddNew     =   -1  'True
         DataMode        =   4
         DefColWidth     =   0
         HeadLines       =   1
         FootLines       =   1
         MarqueeUnique   =   0   'False
         TabAction       =   2
         MultipleLines   =   0
         CellTips        =   2
         CellTipsWidth   =   0
         MultiSelect     =   0
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
         _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7,.namedParent=38"
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
         _StyleDefs(60)  =   ":id=38,.parent=33,.bgcolor=&H80000002&,.fgcolor=&H80000014&"
         _StyleDefs(61)  =   "Named:id=39:EvenRow"
         _StyleDefs(62)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(63)  =   "Named:id=40:OddRow"
         _StyleDefs(64)  =   ":id=40,.parent=33"
         _StyleDefs(65)  =   "Named:id=41:RecordSelector"
         _StyleDefs(66)  =   ":id=41,.parent=34"
         _StyleDefs(67)  =   "Named:id=42:FilterBar"
         _StyleDefs(68)  =   ":id=42,.parent=33"
      End
      Begin MSDataListLib.DataCombo dbcintTabelasDeValores 
         Height          =   315
         Left            =   1215
         TabIndex        =   2
         Top             =   990
         Width           =   4875
         _ExtentX        =   8599
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbcintFaixaDeValores 
         Height          =   315
         Left            =   1215
         TabIndex        =   1
         Top             =   600
         Width           =   4485
         _ExtentX        =   7911
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbcintReferenciaTributo 
         Height          =   315
         Left            =   1215
         TabIndex        =   3
         Top             =   1395
         Width           =   4890
         _ExtentX        =   8625
         _ExtentY        =   556
         _Version        =   393216
         IntegralHeight  =   0   'False
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label lblintReferenciaTributo 
         AutoSize        =   -1  'True
         Caption         =   "Tipo"
         Height          =   195
         Left            =   795
         TabIndex        =   7
         Top             =   1500
         Width           =   315
      End
      Begin VB.Label lblintTabelaDeValores 
         AutoSize        =   -1  'True
         Caption         =   "Utilização"
         Height          =   195
         Left            =   450
         TabIndex        =   6
         Top             =   1110
         Width           =   690
      End
      Begin VB.Label lblintFaixaDeValores 
         AutoSize        =   -1  'True
         Caption         =   "Nome da Faixa"
         Height          =   195
         Left            =   90
         TabIndex        =   4
         Top             =   720
         Width           =   1065
      End
   End
End
Attribute VB_Name = "frmCadValoresDasFaixas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mblnAlterando           As Boolean
Dim mobjAux                 As ComboBox
Dim adoResultado            As ADODB.Recordset
Dim strSQL                  As String
Dim mblnSelecionou          As Boolean
Dim adoRec                  As ADODB.Recordset
Dim adoTdb                  As ADODB.Recordset
Dim x                       As XArrayDB
Dim Y                       As New XArrayDB

Private Function strQueryRelatorio() As String
Dim strSQL As String
    strSQL = ""
    strSQL = strSQL & " select PKID, dblfaixainicial FI, dblFaixaFinal FF, dblValor VL from " & gstrValorDaFaixa
    strSQL = strSQL & " ORdER BY PKID"
strQueryRelatorio = strSQL
End Function

Function Pesquisafaixa()
        If dbcintFaixaDeValores.MatchedWithList Then
            LimpaGrid
            strSQL = ""
            strSQL = strSQL & "SELECT VF.dblFaixaInicial, VF.dblFaixaFinal, "
            strSQL = strSQL & "VF.dblValor, intUtilizacao "
            strSQL = strSQL & "FROM " & gstrValorDaFaixa & " VF "
            strSQL = strSQL & "WHERE VF.intFaixaDeValores = " & dbcintFaixaDeValores.BoundText & " "
            strSQL = strSQL & "ORDER BY VF.dblFaixaInicial"
            
            Set gobjBanco = New clsBanco
            gobjBanco.CriaADO strSQL, 5, adoRec
                                            
            If adoRec.EOF Then
                dbcintTabelasDeValores.BoundText = ""
                'dbcintTabelasDeValores.Locked = False
            Else
                PreencherListaDeOpcoes dbcintTabelasDeValores, adoRec!intUtilizacao
                dbcintTabelasDeValores_Click 2
            End If
            MontaArray
            
            If dbcintTabelasDeValores.BoundText = "" Then
                dbcintTabelasDeValores.SetFocus
                Exit Function
            Else
            End If
                        
            grd_Valores.SetFocus
        Else
            LimpaGrid
        End If
End Function

Private Sub dbcintFaixaDeValores_Change()
    LimpaGrid
    dbcintTabelasDeValores.BoundText = ""
    dbcintReferenciaTributo.BoundText = ""
End Sub

Private Sub dbcintFaixaDeValores_Click(Area As Integer)
    DropDownDataCombo dbcintFaixaDeValores, Me, Area
    If Area = 2 Then
        If dbcintFaixaDeValores.MatchedWithList Then
            strSQL = ""
            strSQL = strSQL & "SELECT PkId, strNomeDaUtilizacao "
            strSQL = strSQL & "FROM " & gstrUtilizacaoDaTabelaDeValor & " "
            strSQL = strSQL & "ORDER BY strNomeDaUtilizacao "
            dbcintTabelasDeValores.Tag = strSQL & ";strNomeDaUtilizacao"
        End If
    End If
End Sub

Private Sub dbcintFaixaDeValores_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintFaixaDeValores, Me, , KeyCode, Shift
End Sub

Private Sub dbcintFaixaDeValores_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Not dbcintFaixaDeValores.MatchedWithList Then
            dbcintFaixaDeValores.BoundText = Trim(dbcintFaixaDeValores.Text)
            If dbcintFaixaDeValores.MatchedWithList Then
                dbcintFaixaDeValores_Click 2
            Else
                LimpaGrid
                dbcintTabelasDeValores.BoundText = ""
            End If
        Else
            dbcintFaixaDeValores_Click 2
        End If
    End If
End Sub

Private Sub dbcintFaixaDeValores_LostFocus()
'    Pesquisafaixa
    
End Sub

Private Sub dbcintTabelasDeValores_Click(Area As Integer)
    DropDownDataCombo dbcintTabelasDeValores, Me, Area
    If Area = 2 Then
        If dbcintTabelasDeValores.MatchedWithList And Len(Trim(dbcintFaixaDeValores.BoundText)) <> 0 Then
            LimpaGrid
            strSQL = ""
            strSQL = strSQL & "SELECT VF.dblFaixaInicial, VF.dblFaixaFinal, "
            strSQL = strSQL & "VF.dblValor, intUtilizacao, intReferenciaTributo "
            strSQL = strSQL & "FROM " & gstrValorDaFaixa & " VF "
            strSQL = strSQL & "WHERE VF.intFaixaDeValores = " & dbcintFaixaDeValores.BoundText
            strSQL = strSQL & " AND VF.intUtilizacao = " & dbcintTabelasDeValores.BoundText
            
            strSQL = strSQL & " ORDER BY VF.dblFaixaInicial"
            
            Set gobjBanco = New clsBanco
            If gobjBanco.CriaADO(strSQL, 5, adoRec) Then
                ' If adoRec.EOF Then
                '    dbcintTabelasDeValores.BoundText = ""
    '               dbcintTabelasDeValores.Locked = False
                'Else
                '    PreencherListaDeOpcoes dbcintTabelasDeValores, adoRec!intUtilizacao
                    'dbcintTabelasDeValores_Click 2
                    'dbcintTabelasDeValores.Locked = True
                'End If
                If Not adoRec.EOF Then
                    dbcintReferenciaTributo.BoundText = IIf(adoRec("intReferenciaTributo").Value > 0, adoRec("intReferenciaTributo").Value, "")
                End If
                                        
                MontaArray
                'If dbcintTabelasDeValores.BoundText = "" Then
                '    dbcintTabelasDeValores.SetFocus
                '    Exit Sub
                'Else
                'End If
                        
                grd_Valores.SetFocus
            End If
        Else
        '   Éder ...Pendencias tri0796
            ExibeMensagem "Selecione o Nome da Faixa !"
            Set dbcintTabelasDeValores.RowSource = Nothing
            dbcintTabelasDeValores.BoundText = ""
            dbcintFaixaDeValores.SetFocus
        End If
    End If
End Sub

Private Sub dbcintTabelasDeValores_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintTabelasDeValores, Me, , KeyCode, Shift
End Sub

Private Sub dbcintTabelasDeValores_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If Not dbcintTabelasDeValores.MatchedWithList Then
            dbcintTabelasDeValores.BoundText = Trim(dbcintTabelasDeValores.Text)
            If dbcintTabelasDeValores.MatchedWithList Then
                grd_Valores.SetFocus
            End If
        Else
            grd_Valores.SetFocus
        End If
    End If
End Sub

Private Sub cmd_FaixaDeValores_Click()
    ChamaFormCadastro frmCadFaixaDeValores, dbcintFaixaDeValores
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 594
     VirificaGradeListView Me
    
'=============
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
'=============
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim ShiftDown, AltDown, CtrlDown
    Select Case KeyCode
        Case vbKeyEscape
            SendKeys "{RIGHT}"
            Exit Sub
    End Select
    ShiftDown = (Shift And vbShiftMask) > 0
    AltDown = (Shift And vbAltMask) > 0
    CtrlDown = (Shift And vbCtrlMask) > 0

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
  If KeyCode = vbKeyF1 Then
     Call_HtmlHelp Me.HelpContextID
  End If
End Sub

Private Sub Form_Load()
    strSQL = ""
    strSQL = strSQL & "SELECT PkId, strNomeDaFaixa "
    strSQL = strSQL & "FROM " & gstrFaixaDeValor & " "
    strSQL = strSQL & "ORDER BY strNomeDaFaixa "
    'LeDaTabelaParaObj gstrFaixaDeValor, dbcintFaixaDeValores, strSql
    LeDaTabelaParaObj "", dbcintReferenciaTributo, strQueryDataComboReferenciaTributo
    dbcintFaixaDeValores.Tag = strSQL & ";strNomeDaFaixa"
End Sub

Private Sub grd_Valores_AfterColEdit(ByVal ColIndex As Integer)
If ColIndex = 0 Then
    grd_Valores.Columns("Faixa Inicial").Value = gstrConvVrDoSql(grd_Valores.Columns("Faixa Inicial").Value, 2)
ElseIf ColIndex = 1 Then
    grd_Valores.Columns("Faixa Final").Value = gstrConvVrDoSql(grd_Valores.Columns("Faixa Final").Value, 2)
ElseIf ColIndex = 2 Then
    grd_Valores.Columns("Valor").Value = gstrConvVrDoSql(grd_Valores.Columns("Valor").Value, 4)
End If

End Sub

Private Sub grd_Valores_KeyPress(KeyAscii As Integer)
    Select Case grd_Valores.Col
        Case 2
            If KeyAscii = vbKeyReturn Then
                KeyAscii = 0
                'SendKeys "%{DOWN}"
            End If
    End Select
    CaracterValido KeyAscii, "V", grd_Valores
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
Select Case UCase(strModoOperacao)
        Case Is = UCase(gstrNovo)
            LimpaGrid
            dbcintFaixaDeValores.BoundText = ""
            dbcintTabelasDeValores.BoundText = ""
            dbcintReferenciaTributo.BoundText = ""
'            dbcintTabelasDeValores.Locked = False
        Case Is = UCase(gstrPreencherLista)
            PreencherListaDeOpcoes Me.ActiveControl
        Case Is = UCase(gstrSalvar)
            If blnDadosOK Then
                GravaValores
            End If
        Case Is = UCase(gstrDeletar)
            DeletaValores
        Case Is = UCase(gstrFechar)
            Unload Me
        Case Is = UCase(gstrImprimir)
            ImprimeRelatorio rptCadValoresFaixas, strQueryRelatorio
    End Select
End Sub

Private Sub MontaArray()
    Dim varAux As Variant
    
    Set x = New XArrayDB
    x.Clear
    
    With adoRec
        If Not .EOF Then
            x.ReDim 0, .RecordCount - 1, 0, 2
            Do While Not .EOF
                varAux = gstrConvVrDoSql(.Fields(0), 2)
                x(.AbsolutePosition - 1, 0) = varAux
                varAux = gstrConvVrDoSql(.Fields(1), 2)
                x(.AbsolutePosition - 1, 1) = varAux
                varAux = gstrConvVrDoSql(.Fields(2), 4)
                x(.AbsolutePosition - 1, 2) = varAux
                .MoveNext
            Loop
        Else
            x.ReDim 0, 0, 0, 2
            x(0, 0) = ""
            x(0, 1) = ""
            x(0, 2) = ""
'            X(0, 3) = ""
        End If
    End With
    
    Set grd_Valores.Array = x
    grd_Valores.ReBind
    grd_Valores.Refresh
End Sub

Sub DeletaValores()
    Dim strSQL As String
    
    If dbcintFaixaDeValores.BoundText = "" Then
        ExibeMensagem "Faixa de valores tem que ser selecionada."
        Exit Sub
    End If
    
    If MsgBox("Confirma exclusão de todas estas faixa de valores?", vbQuestion + vbYesNo) = vbYes Then
        strSQL = ""
        strSQL = strSQL & "DELETE FROM " & gstrValorDaFaixa & " "
        strSQL = strSQL & "WHERE intFaixaDeValores = " & dbcintFaixaDeValores.BoundText
        
        Set gobjBanco = New clsBanco
        gobjBanco.Execute strSQL
        
        LimpaGrid
        dbcintFaixaDeValores.BoundText = ""
        dbcintTabelasDeValores.BoundText = ""
        dbcintReferenciaTributo.BoundText = ""
        dbcintTabelasDeValores.Locked = False
    End If
End Sub

Sub GravaValores()

'******************************************************************************************
' Data: 02/05/2003
' Alteração: - Substituição do comando nativo GETDATE() do SQL Server pela variável
'            strGETDATE.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Dim strSQL As String
    Dim strMsg As String
    Dim i      As Integer
        
    strMsg = "Confirma gravação destes valores?"
    
    If gblnExclusaoGravacaoOk("", strMsg, True) Then
         Set gobjBanco = New clsBanco
         gobjBanco.ExecutaBeginTrans

        strSQL = ""
        strSQL = strSQL & "DELETE FROM " & gstrValorDaFaixa & " "
        strSQL = strSQL & "WHERE intFaixaDeValores = " & dbcintFaixaDeValores.BoundText
        gobjBanco.Execute strSQL

        grd_Valores.MoveFirst
'
'        If dbcintReferenciaTributo.MatchedWithList Then
'            If blnTipoJaCadastrado(dbcintReferenciaTributo.BoundText) Then
'                ExibeMensagem "Este Tipo já está relacionado."
'                gobjBanco.ExecutaRollbackTrans
'                Exit Sub
'            End If
'        End If
        
        For i = 0 To x.Count(1) - 1
            strSQL = ""
            strSQL = strSQL & "INSERT INTO " & gstrValorDaFaixa & " "
            strSQL = strSQL & "(intFaixaDeValores, intUtilizacao, intReferenciaTributo, "
            strSQL = strSQL & "dblFaixaInicial, dblFaixaFinal, dblValor, dtmDtAtualizacao, lngCodUsr"
            strSQL = strSQL & ") Values ("
            strSQL = strSQL & dbcintFaixaDeValores.BoundText & ", "
            strSQL = strSQL & dbcintTabelasDeValores.BoundText & ", "
            strSQL = strSQL & gstrENulo(dbcintReferenciaTributo.BoundText, , True) & ", "
            strSQL = strSQL & gstrConvVrParaSql(x(i, 0)) & ", "
            strSQL = strSQL & gstrConvVrParaSql(x(i, 1)) & ", "
            If x(i, 2) = "" Or IsNull(x(i, 2)) Or x(i, 2) = Empty Or x(i, 2) = "0" Then
                strSQL = strSQL & "NULL, "
            Else
                strSQL = strSQL & gstrConvVrParaSql(x(i, 2)) & ", "
            End If
            strSQL = strSQL & strGETDATE & ", "
            strSQL = strSQL & glngCodUsr
            strSQL = strSQL & ")"
        
            If Not gobjBanco.Execute(strSQL, False) Then
                gobjBanco.ExecutaRollbackTrans
            End If
        Next i
        gobjBanco.ExecutaCommitTrans
        LimpaGrid
        dbcintFaixaDeValores.BoundText = ""
        dbcintTabelasDeValores.BoundText = ""
        dbcintReferenciaTributo.BoundText = ""
        dbcintTabelasDeValores.Locked = False
    End If
End Sub

Private Function blnDadosOK() As Boolean
    Dim i As Integer
    
    If Not dbcintFaixaDeValores.MatchedWithList Then
        ExibeMensagem "A nome da faixa tem que ser selecionado."
        dbcintFaixaDeValores.SetFocus
        Exit Function
    ElseIf Not dbcintTabelasDeValores.MatchedWithList Then
        ExibeMensagem "A utilização tem que ser selecionada."
        dbcintTabelasDeValores.SetFocus
        Exit Function
    End If
    
    If Not dbcintReferenciaTributo.MatchedWithList Then
        ExibeMensagem "O tipo tem que ser selecionado."
        dbcintReferenciaTributo.SetFocus
        Exit Function
    End If
    
    grd_Valores.MoveFirst
    For i = 0 To x.Count(1) - 1
        If x(i, 0) = "" Or x(i, 1) = "" Then
            MsgBox "A faixa inicial tem que ser digitada."
            grd_Valores.Row = i
            grd_Valores.SetFocus
            Exit Function
        ElseIf x(i, 1) = "" Then
            MsgBox "A faixa final tem que ser digitada."
            grd_Valores.Row = i
            grd_Valores.SetFocus
            Exit Function
        ElseIf CDbl(x(i, 0)) >= CDbl(x(i, 1)) Then
            MsgBox "A faixa inicial não pode ser maior ou igual que à faixa final."
            grd_Valores.Row = i
            grd_Valores.SetFocus
            Exit Function
        Else
        '    If (i + 1) < (X.Count(1) - 1) Then
'                If CDbl(X(i, 1)) >= CDbl(X(i + 1, 0)) Then
'                    MsgBox "A faixa final de um registro não pode ser maior ou igual à faixa inicial do registro posterior."
'                    grd_Valores.Row = i
'                    grd_Valores.SetFocus
'                    Exit Function
'                End If
'            End If
'            If i > 0 Then
'                If CDbl(X(i, 0)) <= CDbl(X(i - 1, 1)) Then
'                    MsgBox "A faixa inicial de um registro não pode ser menor ou igual à faixa final do registro anterior."
'                    grd_Valores.Row = i
'                    grd_Valores.SetFocus
'                    Exit Function
'                End If
'            End If
        End If
    Next
    blnDadosOK = True
End Function

Private Sub LimpaGrid()
    Set x = New XArrayDB
    
    x.Clear
    
    Set grd_Valores.Array = x
    grd_Valores.ReBind
    grd_Valores.Refresh
End Sub

Private Function blnValorCadastrado(dblValor As Variant) As Boolean
    Dim i As Integer
    
    For i = 0 To Y.Count(1) - 1
        If gvntConvVrDoSql(Y(i, 1)) = gvntConvVrDoSql(dblValor) Then
            blnValorCadastrado = True
            Exit Function
        End If
    Next
    blnValorCadastrado = False
End Function

Private Function strQueryDataComboReferenciaTributo() As String
    Dim strSQL As String
    strSQL = ""
    strSQL = strSQL & "SELECT PKId, strDescricao "
    strSQL = strSQL & "FROM " & gstrReferenciasDeTributos & " "
    strSQL = strSQL & "WHERE bytExibir = 1 AND bytGrupo = " & GRUPO_IMOB_TERRENO_APURADO & " "
    strSQL = strSQL & "ORDER BY strDescricao"
    strQueryDataComboReferenciaTributo = strSQL
End Function

Private Function blnTipoJaCadastrado(lngTipo As Long) As Boolean
Dim adoConsulta  As New ADODB.Recordset
Dim strSQL       As String

    strSQL = "SELECT Pkid FROM " & gstrValorDaFaixa & " WHERE intReferenciaTributo = " & lngTipo & " AND intFaixaDeValores <> " & dbcintFaixaDeValores.BoundText
    
    If gobjBanco.CriaADO(strSQL, 5, adoConsulta) Then
        blnTipoJaCadastrado = Not adoConsulta.EOF
    End If
    
End Function

