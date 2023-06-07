VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRelatorioDeParcelasLancadas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relação das Parcelas Lançadas"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7410
   HelpContextID   =   699
   Icon            =   "RelatorioDeParcelasLancadas.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   7410
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   2865
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   5054
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Relação das Parcelas Lançadas"
      TabPicture(0)   =   "RelatorioDeParcelasLancadas.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame Frame1 
         Height          =   2265
         Left            =   180
         TabIndex        =   1
         Top             =   360
         Width           =   6795
         Begin VB.TextBox txtDtInicial 
            Height          =   285
            Left            =   1590
            MaxLength       =   4
            TabIndex        =   6
            Top             =   1470
            Width           =   1065
         End
         Begin VB.ComboBox cbo_intUtilizacaoDebito 
            Height          =   315
            ItemData        =   "RelatorioDeParcelasLancadas.frx":105E
            Left            =   1590
            List            =   "RelatorioDeParcelasLancadas.frx":1060
            TabIndex        =   2
            Top             =   420
            Width           =   5055
         End
         Begin VB.TextBox txt_strInscricaoCadastral 
            Alignment       =   1  'Right Justify
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
            Left            =   4095
            MaxLength       =   15
            TabIndex        =   4
            Top             =   780
            Width           =   1935
         End
         Begin VB.TextBox txt_strCodigo 
            Alignment       =   1  'Right Justify
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
            Left            =   1590
            MaxLength       =   6
            TabIndex        =   3
            Top             =   780
            Width           =   975
         End
         Begin MSDataListLib.DataCombo dbc_intContribuinte 
            Height          =   315
            Left            =   1590
            TabIndex        =   5
            Top             =   1110
            Width           =   5055
            _ExtentX        =   8916
            _ExtentY        =   556
            _Version        =   393216
            IntegralHeight  =   0   'False
            Text            =   ""
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Exercício"
            Height          =   195
            Left            =   855
            TabIndex        =   11
            Top             =   1560
            Width           =   675
         End
         Begin VB.Label lbl_intUtilizacaoDebito 
            AutoSize        =   -1  'True
            Caption         =   "Utilização"
            Height          =   195
            Left            =   840
            TabIndex        =   10
            Top             =   540
            Width           =   690
         End
         Begin VB.Label lbl_strInscricao 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Inscrição Cadastral"
            Height          =   195
            Left            =   2640
            TabIndex        =   9
            Top             =   870
            Width           =   1350
         End
         Begin VB.Label lbl_strCodigo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Código"
            Height          =   195
            Left            =   1035
            TabIndex        =   8
            Top             =   870
            Width           =   495
         End
         Begin VB.Label lbl_Label1 
            AutoSize        =   -1  'True
            Caption         =   "Código/Contribuinte"
            Height          =   195
            Left            =   120
            TabIndex        =   7
            Top             =   1230
            Width           =   1410
         End
      End
   End
End
Attribute VB_Name = "frmRelatorioDeParcelasLancadas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub dbc_intContribuinte_Click(Area As Integer)
    DropDownDataCombo dbc_intContribuinte, Me, Area
End Sub

Private Sub dbc_intContribuinte_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_intContribuinte, Me, , KeyCode, Shift
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 699
    
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrSalvar
    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrImprimir
End Sub

Private Sub Form_Deactivate()
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
End Sub

Private Sub Form_Load()
cbo_intUtilizacaoDebito.AddItem "Imobiliárias "
cbo_intUtilizacaoDebito.ItemData(cbo_intUtilizacaoDebito.NewIndex) = "1"
cbo_intUtilizacaoDebito.AddItem "Econômicas"
cbo_intUtilizacaoDebito.ItemData(cbo_intUtilizacaoDebito.NewIndex) = "2"
cbo_intUtilizacaoDebito.AddItem "Fiscalização"
cbo_intUtilizacaoDebito.ItemData(cbo_intUtilizacaoDebito.NewIndex) = "3"
cbo_intUtilizacaoDebito.AddItem "Outras Receitas"
cbo_intUtilizacaoDebito.ItemData(cbo_intUtilizacaoDebito.NewIndex) = "4"

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)

'******************************************************************************************
' Data: 09/05/2003
' Alteração: - Substituição da chamada direta à stored procedure pela função
'            gstrStoredProcedure.
' Responsável: Everton Bianchini
'******************************************************************************************

Dim strSQL As String

Dim strParameters As String

If strModoOperacao = gstrImprimir Then
'    strSql = ""
'    strSql = strSql & " SELECT DISTINCT"
'    strSql = strSql & " CO.PKId AS PKIdContribuinte, CO.strCodigoAnterior, CO.strNome,"
'    strSql = strSql & " IM.PKId AS PKIdImobiliario, IM.strCodigo, IM.strInscricaoAnterior AS strInscricao,"
'    strSql = strSql & " PT.intNumeroParcela, CR.strDescricao, PT.dblValorParcela, PT.dtmDataVencimento, "
'    strSql = strSql & " LC.PKId AS PKIdLancamentoCalculo, "
'    strSql = strSql & " CASE PT.strSituacao"
'    strSql = strSql & " WHEN 'A' THEN 'Aberto'"
'    strSql = strSql & " WHEN 'P' THEN 'Pago'"
'    strSql = strSql & " WHEN 'E' THEN 'Eliminado'"
'    strSql = strSql & " ELSE 'Aberto'  END AS strSituacao "
'    strSql = strSql & " FROM "
'    strSql = strSql & gstrContribuinte & " CO, "
'    strSql = strSql & gstrImobiliario & " IM, "
'    strSql = strSql & gstrReceita & " CR, "
'    strSql = strSql & gstrParcelaTaxa & " PT, "
'    strSql = strSql & gstrLancamentoCalculo & " LC "
'    strSql = strSql & " WHERE CO.PKId = IM.intContribuinte"
'    strSql = strSql & " AND CO.PKId = LC.intContribuinte"
'    strSql = strSql & " AND IM.strInscricaoAnterior = LC.strInscricaoCadastral"
'    strSql = strSql & " AND CR.PKID = PT.intReceita"
'    strSql = strSql & " AND LC.PKId = PT.intLancamentoCalculo"

    If cbo_intUtilizacaoDebito.ListIndex = 0 Then
'        strSQL = "sp_ParcelaLancada 0" 'Imobiliário
        strSQL = "sp_ParcelaLancada"
        
        strParameters = "0" 'Imobiliário
        
        If Trim(txt_strInscricaoCadastral.Text) <> "" Then
'            strSQL = strSQL & ", '" & txt_strInscricaoCadastral.Text & "'"
            strParameters = strParameters & ", '" & txt_strInscricaoCadastral.Text & "'"
        Else
'            strSQL = strSQL & ", NULL"
            strParameters = strParameters & ", NULL"
        End If
        If Trim(txt_strCodigo.Text) <> "" Then
'            strSQL = strSQL & ", '" & txt_strCodigo.Text & "'"
            strParameters = strParameters & ", '" & txt_strCodigo.Text & "'"
        Else
'            strSQL = strSQL & ", NULL"
            strParameters = strParameters & ", NULL"
        End If
    ElseIf cbo_intUtilizacaoDebito.ListIndex = 1 Then
'        strSQL = "sp_ParcelaLancada 1" 'Econômico
        strSQL = "sp_ParcelaLancada"
        
        strParameters = "1" 'Econômico
        
        If Trim(txt_strInscricaoCadastral.Text) <> "" Then
'            strSQL = strSQL & ", '" & txt_strInscricaoCadastral.Text & "'"
            strParameters = strParameters & ", '" & txt_strInscricaoCadastral.Text & "'"
        Else
'            strSQL = strSQL & ", NULL"
            strParameters = strParameters & ", NULL"
        End If
        
        strParameters = strParameters & IIf((bytDBType = EDatabases.Oracle), ", NULL", "")
        
    End If
    
    strSQL = gstrStoredProcedure(strSQL, strParameters, True)
    
    ImprimeRelatorio rptRelatorioDeParcelasLancadas, strSQL, "Relação das Parcelas Lançadas dos Imóveis"
    
End If
End Sub

