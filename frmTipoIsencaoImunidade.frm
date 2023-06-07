VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmTipoIsencaoImunidade 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tipos de isenção/imunidade"
   ClientHeight    =   4395
   ClientLeft      =   4035
   ClientTop       =   3360
   ClientWidth     =   5220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   5220
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtPkId 
      Height          =   285
      Left            =   4200
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   4170
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   5100
      _ExtentX        =   8996
      _ExtentY        =   7355
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Tipos de Isenção / Imunudade"
      TabPicture(0)   =   "frmTipoIsencaoImunidade.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl_strDescricao"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "tdb_Tipos"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtstrDescricao"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fra_Tipo"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtintTipo"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      Begin VB.TextBox txtintTipo 
         Height          =   285
         Left            =   3000
         TabIndex        =   9
         Top             =   0
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Frame fra_Tipo 
         Height          =   615
         Left            =   1200
         TabIndex        =   3
         Top             =   1080
         Width           =   3615
         Begin VB.OptionButton opt_Tipo 
            Caption         =   "Não incidente"
            Height          =   195
            Index           =   2
            Left            =   2160
            TabIndex        =   6
            Top             =   250
            Width           =   1285
         End
         Begin VB.OptionButton opt_Tipo 
            Caption         =   "Imune"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   5
            Top             =   220
            Width           =   735
         End
         Begin VB.OptionButton opt_Tipo 
            Caption         =   "Isento"
            Height          =   375
            Index           =   1
            Left            =   1200
            TabIndex        =   4
            Top             =   160
            Width           =   735
         End
      End
      Begin VB.TextBox txtstrDescricao 
         Height          =   285
         Left            =   1200
         MaxLength       =   35
         TabIndex        =   2
         Top             =   720
         Width           =   3615
      End
      Begin TrueOleDBGrid70.TDBGrid tdb_Tipos 
         Height          =   1935
         Left            =   240
         TabIndex        =   7
         Top             =   1920
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   3413
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
         Columns(1).Caption=   "Descrição"
         Columns(1).DataField=   "strDescricao"
         Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(2)._VlistStyle=   0
         Columns(2)._MaxComboItems=   5
         Columns(2).Caption=   "intTipo"
         Columns(2).DataField=   "intTipo"
         Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns(3)._VlistStyle=   16
         Columns(3)._MaxComboItems=   5
         Columns(3).ValueItems(0)._DefaultItem=   0
         Columns(3).ValueItems(0).Value=   "0"
         Columns(3).ValueItems(0).Value.vt=   8
         Columns(3).ValueItems(0).DisplayValue=   "Imune"
         Columns(3).ValueItems(0).DisplayValue.vt=   8
         Columns(3).ValueItems(0)._PropDict=   "_DefaultItem,517,2"
         Columns(3).ValueItems(1)._DefaultItem=   0
         Columns(3).ValueItems(1).Value=   "1"
         Columns(3).ValueItems(1).Value.vt=   8
         Columns(3).ValueItems(1).DisplayValue=   "Isento"
         Columns(3).ValueItems(1).DisplayValue.vt=   8
         Columns(3).ValueItems(1)._PropDict=   "_DefaultItem,517,2"
         Columns(3).ValueItems(2)._DefaultItem=   0
         Columns(3).ValueItems(2).Value=   "2"
         Columns(3).ValueItems(2).Value.vt=   8
         Columns(3).ValueItems(2).DisplayValue=   "Não incidente"
         Columns(3).ValueItems(2).DisplayValue.vt=   8
         Columns(3).ValueItems(2)._PropDict=   "_DefaultItem,517,2"
         Columns(3).ValueItems.Count=   3
         Columns(3).Caption=   "Tipo"
         Columns(3).DataField=   "intTipo"
         Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
         Columns.Count   =   4
         Splits(0)._UserFlags=   0
         Splits(0).MarqueeStyle=   3
         Splits(0).RecordSelectorWidth=   503
         Splits(0)._SavedRecordSelectors=   0   'False
         Splits(0).DividerColor=   12632256
         Splits(0).SpringMode=   0   'False
         Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
         Splits(0)._ColumnProps(0)=   "Columns.Count=4"
         Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
         Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
         Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
         Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
         Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
         Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
         Splits(0)._ColumnProps(7)=   "Column(1).Width=4789"
         Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
         Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=4710"
         Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
         Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
         Splits(0)._ColumnProps(12)=   "Column(2).Width=212"
         Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
         Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=132"
         Splits(0)._ColumnProps(15)=   "Column(2)._EditAlways=0"
         Splits(0)._ColumnProps(16)=   "Column(2).Visible=0"
         Splits(0)._ColumnProps(17)=   "Column(2).Order=3"
         Splits(0)._ColumnProps(18)=   "Column(3).Width=2143"
         Splits(0)._ColumnProps(19)=   "Column(3).DividerColor=0"
         Splits(0)._ColumnProps(20)=   "Column(3)._WidthInPix=2064"
         Splits(0)._ColumnProps(21)=   "Column(3)._EditAlways=0"
         Splits(0)._ColumnProps(22)=   "Column(3).Order=4"
         Splits.Count    =   1
         PrintInfos(0)._StateFlags=   0
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
         _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=28,.bold=0,.fontsize=825,.italic=0"
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
         _StyleDefs(30)  =   "Splits(0).Columns(0).Style:id=46,.parent=13"
         _StyleDefs(31)  =   "Splits(0).Columns(0).HeadingStyle:id=43,.parent=14"
         _StyleDefs(32)  =   "Splits(0).Columns(0).FooterStyle:id=44,.parent=15"
         _StyleDefs(33)  =   "Splits(0).Columns(0).EditorStyle:id=45,.parent=17"
         _StyleDefs(34)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
         _StyleDefs(35)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
         _StyleDefs(36)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
         _StyleDefs(37)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
         _StyleDefs(38)  =   "Splits(0).Columns(2).Style:id=50,.parent=13"
         _StyleDefs(39)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
         _StyleDefs(40)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
         _StyleDefs(41)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
         _StyleDefs(42)  =   "Splits(0).Columns(3).Style:id=32,.parent=13"
         _StyleDefs(43)  =   "Splits(0).Columns(3).HeadingStyle:id=29,.parent=14"
         _StyleDefs(44)  =   "Splits(0).Columns(3).FooterStyle:id=30,.parent=15"
         _StyleDefs(45)  =   "Splits(0).Columns(3).EditorStyle:id=31,.parent=17"
         _StyleDefs(46)  =   "Named:id=33:Normal"
         _StyleDefs(47)  =   ":id=33,.parent=0"
         _StyleDefs(48)  =   "Named:id=34:Heading"
         _StyleDefs(49)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(50)  =   ":id=34,.wraptext=-1"
         _StyleDefs(51)  =   "Named:id=35:Footing"
         _StyleDefs(52)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
         _StyleDefs(53)  =   "Named:id=36:Selected"
         _StyleDefs(54)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(55)  =   "Named:id=37:Caption"
         _StyleDefs(56)  =   ":id=37,.parent=34,.alignment=2"
         _StyleDefs(57)  =   "Named:id=38:HighlightRow"
         _StyleDefs(58)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
         _StyleDefs(59)  =   "Named:id=39:EvenRow"
         _StyleDefs(60)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
         _StyleDefs(61)  =   "Named:id=40:OddRow"
         _StyleDefs(62)  =   ":id=40,.parent=33"
         _StyleDefs(63)  =   "Named:id=41:RecordSelector"
         _StyleDefs(64)  =   ":id=41,.parent=34"
         _StyleDefs(65)  =   "Named:id=42:FilterBar"
         _StyleDefs(66)  =   ":id=42,.parent=33"
      End
      Begin VB.Label lbl_strDescricao 
         AutoSize        =   -1  'True
         Caption         =   "Descrição"
         Height          =   195
         Left            =   360
         TabIndex        =   1
         Top             =   765
         Width           =   720
      End
   End
End
Attribute VB_Name = "frmTipoIsencaoImunidade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public intCodSeguranca As Integer
Dim intIndice      As Integer
Dim blnPrimeiraVez As Boolean
Dim mblnAlterando  As Boolean
Dim mobjAux        As Object

Private Sub Form_Activate()
    VerificaObjParaAplicar mobjAux
    gintCodSeguranca = 1148
    If mblnAlterando Then
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


Private Sub opt_Tipo_Click(Index As Integer)
intIndice = Index
txtintTipo = intIndice
End Sub

Private Sub tdb_Tipos_Click()
    blnPrimeiraVez = False
End Sub

Private Sub tdb_Tipos_DblClick()
    If MDIMenu.actBarra.Tools(gstrAplicar).Enabled = True Then
        MantemForm gstrAplicar
    End If
End Sub
    
Private Sub tdb_Tipos_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If blnPrimeiraVez = False Then
    
        With tdb_Tipos
        
        If Not .BOF Or .EOF Then
        
            mblnAlterando = True
            
            HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
            
            txtPkId = .Columns("PKID").Text
            txtstrDescricao.Text = .Columns("strDescricao")
                
                Select Case .Columns("intTipo")
                    Case Is = 0
                        opt_Tipo(0).Value = True
                        txtintTipo = opt_Tipo(0).Index
                    Case Is = 1
                        opt_Tipo(1).Value = True
                        txtintTipo = opt_Tipo(1).Index
                    Case Is = 2
                        opt_Tipo(2).Value = True
                        txtintTipo = opt_Tipo(2).Index
                End Select
                
            End If
            
        End With
    End If
End Sub
Private Sub txtstrDescricao_GotFocus()
    MarcaCampo txtstrDescricao
End Sub


Public Sub MantemForm(ByVal strModoOperacao As String)
    Dim strSql As String
    strSql = strQueryAplicar
    
    Select Case strModoOperacao
    
        Case Is = UCase(gstrNovo)
            mblnAlterando = False
            txtPkId.Text = ""
            txtstrDescricao = ""
            txtintTipo = ""
            opt_Tipo(0).Value = False
            opt_Tipo(1).Value = False
            opt_Tipo(2).Value = False
            txtstrDescricao.SetFocus
            
        Case Is = UCase(gstrSalvar)
            If Not blnDadosOk Then
                Exit Sub
            End If
        Case Is = UCase(gstrLocalizar)
            txtPkId = ""
            txtintTipo = ""
        Case Is = UCase(gstrImprimir)
            ImprimeRelatorio rptTipoIsencaoImunidade, strQueryRelatorio
            Exit Sub
    End Select
    
    ToolBarGeral strModoOperacao, gstrTipoIsencaoImunidade, mblnAlterando, tdb_Tipos, Me, mobjAux, , strQueryAplicar
    
'        ToolBarGeral strModoOperacao, gstrMoedas, mblnAlterando, tdb_Lista, _
'                 Me, mobjAux, strQuery, strQueryAplicar
'
    If (UCase(strModoOperacao) = gstrSalvar) Or (UCase(strModoOperacao) = gstrDeletar) Then
        LeDaTabelaParaObj "", tdb_Tipos, "SELECT * FROM " & gstrTipoIsencaoImunidade
        MantemForm gstrNovo
    End If
        
End Sub
Private Function strQueryAplicar() As String

Dim strSql As String

    strSql = "SELECT pkid, strDescricao, intTipo "
    strSql = strSql & "FROM " & gstrTipoIsencaoImunidade
    
strQueryAplicar = strSql

End Function

Private Function strQueryRelatorio() As String
Dim strSql As String
strSql = ""
    strSql = strSql & "select PKID, STRDEsCRiCAO, "
    strSql = strSql & gstrCASEWHEN("INTTIPO", "0,'Imune',1,'Isento',2,'Não Incidente'") & " TIPO from "
    strSql = strSql & gstrTipoIsencaoImunidade
strQueryRelatorio = strSql
End Function
Private Function blnDadosOk() As Boolean

    blnDadosOk = False

    If Len(Trim(txtstrDescricao.Text)) = 0 Then
        ExibeMensagem "Digite uma descrição."
        txtstrDescricao.SetFocus
        Exit Function
    End If
        
    If (Not opt_Tipo(0).Value) And (Not opt_Tipo(1).Value) And (Not opt_Tipo(2).Value) Then
        ExibeMensagem "Selecione um tipo."
        Exit Function
    End If
    
    If mblnAlterando = True Then
       If txtstrDescricao.Text <> tdb_Tipos.Columns("Descrição") Or _
          txtintTipo.Text <> tdb_Tipos.Columns("intTipo") Then
          
          If gblnExisteCodigo(1, gstrTipoIsencaoImunidade, "strDescricao", txtstrDescricao.Text, _
             "intTipo", txtintTipo.Text) = True Then
             ExibeMensagem "O tipo '" & opt_Tipo(Val(txtintTipo.Text)).Caption & _
             "' referente à '" & txtstrDescricao.Text & "' já existe."
             Exit Function
          End If
          
       End If
    Else
       If gblnExisteCodigo(1, gstrTipoIsencaoImunidade, "strDescricao", txtstrDescricao.Text, _
          "intTipo", txtintTipo.Text) = True Then
          ExibeMensagem "O tipo '" & opt_Tipo(Val(txtintTipo.Text)).Caption & _
          "' referente à '" & txtstrDescricao.Text & "' já existe."
          Exit Function
       End If
    End If

    blnDadosOk = True

End Function
