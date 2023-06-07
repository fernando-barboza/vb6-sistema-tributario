VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCadInconsistenciaImobiliaria 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inconsistências Imobiliarias"
   ClientHeight    =   1755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4845
   Icon            =   "CadInconsistenciaImobiliaria.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   4845
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   1485
      Left            =   150
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   150
      Width           =   4590
      _ExtentX        =   8096
      _ExtentY        =   2619
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Inconsistências Imobiliárias"
      TabPicture(0)   =   "CadInconsistenciaImobiliaria.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl_bunda"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl_Inicial"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "mskstrInscricaoAnterior2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "mskstrInscricaoAnterior"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      Begin MSMask.MaskEdBox mskstrInscricaoAnterior 
         Height          =   285
         Left            =   2130
         TabIndex        =   0
         Top             =   540
         Width           =   2250
         _ExtentX        =   3969
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   24
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskstrInscricaoAnterior2 
         Height          =   285
         Left            =   2130
         TabIndex        =   1
         Top             =   900
         Width           =   2250
         _ExtentX        =   3969
         _ExtentY        =   503
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   24
         PromptChar      =   " "
      End
      Begin VB.Label lbl_Inicial 
         AutoSize        =   -1  'True
         Caption         =   "Inscrição cadastral inicial"
         Height          =   195
         Left            =   210
         TabIndex        =   4
         Top             =   630
         Width           =   1770
      End
      Begin VB.Label lbl_bunda 
         AutoSize        =   -1  'True
         Caption         =   "Inscrição cadastral final"
         Height          =   195
         Left            =   315
         TabIndex        =   3
         Top             =   990
         Width           =   1665
      End
   End
End
Attribute VB_Name = "frmCadInconsistenciaImobiliaria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim mblnAlterando    As Boolean
    Dim mobjAux          As Object
    Dim mblnSelecionou   As Boolean
    Dim mblnPrimeiraVez  As Boolean
    Dim adoResultado As ADODB.Recordset

Private Sub Form_Activate()
    gintCodSeguranca = 688
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
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrSalvar, gstrDeletar
    VerificaMascaraInscricao
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
    mblnSelecionou = False
    mblnPrimeiraVez = False
End Sub
Public Sub MantemForm(ByVal strModoOperacao As String)
Dim adoRelatorio   As ADODB.Recordset
On Error GoTo ErroImprimeRelatorio
        If UCase(strModoOperacao) = "IMPRIMIR" Then
            If mskstrInscricaoAnterior.Text = "" Then
                ExibeMensagem "A inscrição inicial tem que ser digitada."
                mskstrInscricaoAnterior.SetFocus
                Exit Sub
            End If
            If mskstrInscricaoAnterior2.Text = "" Then
                ExibeMensagem "A inscrição final tem que ser digitada."
                mskstrInscricaoAnterior2.SetFocus
                Exit Sub
            End If
        End If
        If UCase(strModoOperacao) = "IMPRIMIR" Then
            If strQuerryTemOuNaoTem = False Then
                ExibeMensagem "Não foi encontrado nenhum cadastro com estas inscrições."
                mskstrInscricaoAnterior.SetFocus
                Exit Sub
            End If
        End If
        If UCase(strModoOperacao) = "IMPRIMIR" Then
            Screen.MousePointer = vbHourglass
            Set gobjBanco = New clsBanco
            If gobjBanco.CriaADO(strQuerryRelatorio, 5, adoRelatorio) Then
                Set rptInconsistenciaImobiliario.adoDataControl.Recordset = adoRelatorio
                rptInconsistenciaImobiliario.Show
            End If
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    If UCase(strModoOperacao) = "NOVO" Then
        LimpaObjetos
    End If
Screen.MousePointer = vbDefault
ErroImprimeRelatorio:
Resume FimImprimeRelatorio
            
FimImprimeRelatorio:
End Sub

Private Function strQuerryTemOuNaoTem() As Boolean

'******************************************************************************************
' Data: 02/05/2003
' Alteração: - Substituição do comando nativo CONVERT do SQL Server pela função gstrCONVERT
' Responsável: Everton Bianchini
'******************************************************************************************
    
Dim strSql     As String
Dim strInicial As String
Dim strFinal   As String
strInicial = mskstrInscricaoAnterior.Text
strFinal = mskstrInscricaoAnterior2.Text
    strSql = ""
    strSql = strSql & " SELECT COUNT(*) as Contador FROM " & gstrImobiliario
'    strSql = strSql & " WHERE CONVERT(NUMERIC(30),strInscricaoAnterior) BETWEEN " & gstrConvVrParaSql(Val(strInicial)) & " AND " & gstrConvVrParaSql(Val(strFinal))
    strSql = strSql & " WHERE " & gstrCONVERT(CDT_NUMERIC, "strInscricaoAnterior") & " BETWEEN " & gstrConvVrParaSql(Val(strInicial)) & " AND " & gstrConvVrParaSql(Val(strFinal))
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        With adoResultado
            Do While Not .EOF
                If !Contador = 0 Then
                    strQuerryTemOuNaoTem = False
                    Exit Function
                End If
                .MoveNext
            Loop
        End With
    End If
strQuerryTemOuNaoTem = True
End Function

Private Function strQuerryRelatorio() As String
Dim strSql     As String
strSql = ""
strSql = strSql & " SELECT count(*) as Contador from " & gstrImobiliario
strQuerryRelatorio = strSql
End Function

Sub VerificaMascaraInscricao()
Dim strSql As String
Dim adoResultado As ADODB.Recordset
Dim strMascara   As String
strMascara = ""
    strSql = ""
    strSql = strSql & "Select * From " & gstrCampoDeInscricao & " "
    strSql = strSql & "Where intTipoDeInscricao = " & TYP_IMOBILIARIA
    strSql = strSql & "Order By intSequencia"

    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        With adoResultado
            Do While Not .EOF
                strMascara = strMascara & String(!intTamanho, "#") & gstrVerificaCampoNulo(!strSeparador)
                .MoveNext
            Loop
        End With
    End If
    mskstrInscricaoAnterior.Mask = strMascara
    mskstrInscricaoAnterior2.Mask = strMascara
End Sub

Sub LimpaObjetos()
mskstrInscricaoAnterior.Text = ""
mskstrInscricaoAnterior2.Text = ""
mskstrInscricaoAnterior.SetFocus
End Sub

'''Private Function strQuerryAnalitico2() As String
'''Dim strSql As String
'''Dim dtInicial  As Date
'''Dim dtFinal    As Date
'''Dim codInicial As Double
'''Dim codFinal   As Double
'''dtInicial = CVDate(txt_Devolucao)
'''dtFinal = CVDate(txt_Ate)
'''codInicial = Val(txt_Inicial)
'''codFinal = Val(txt_Final)
'''
'''    strSql = ""
'''    strSql = strSql & " SELECT COUNT(*) as TotalDocs , DV.intDocumentosEmitidos Inteiro, DE.strDescricao DocNome "
'''    strSql = strSql & " FROM " & gstrDevolucao & " DV, "
'''    strSql = strSql & gstrDocumentoEmitido & " DE "
'''    strSql = strSql & " WHERE DV.intDocumentosEmitidos = DE.PKId "
'''    strSql = strSql & " AND DV.dtmDevolucao BETWEEN " & gstrConvDtParaSql(dtInicial) & " AND " & gstrConvDtParaSql(dtFinal)
'''    strSql = strSql & " AND DV.intContribuinte BETWEEN " & codInicial & " AND " & codFinal
'''    strSql = strSql & " GROUP BY DV.intDocumentosEmitidos, DE.strDescricao "
'''
'''strQuerryAnalitico2 = strSql
'''End Function
'''
'''
'''Private Function strQuerryAnalitico() As String
'''Dim strSql     As String
'''Dim dtInicial  As Date
'''Dim dtFinal    As Date
'''Dim codInicial As Double
'''Dim codFinal   As Double
'''dtInicial = CVDate(txt_Devolucao)
'''dtFinal = CVDate(txt_Ate)
'''codInicial = Val(txt_Inicial)
'''codFinal = Val(txt_Final)
'''    strSql = ""
'''
'''    strSql = strSql & " SELECT COUNT(*) as TOTAL, DV.strInscricao, DV.intContribuinte, "
'''    strSql = strSql & " CO.strNome, DE.strDescricao Documento, OC.strDescricao Ocorrencia,"
'''    strSql = strSql & " DV.intDocumentosEmitidos "
'''
'''    strSql = strSql & " FROM "
'''    strSql = strSql & gstrDevolucao & " DV,"
'''    strSql = strSql & gstrContribuinte & " CO,"
'''    strSql = strSql & gstrOcorrencia & " OC,"
'''    strSql = strSql & gstrDocumentoEmitido & " DE "
'''
'''    strSql = strSql & " WHERE DV.dtmDevolucao BETWEEN " & gstrConvDtParaSql(dtInicial) & " AND " & gstrConvDtParaSql(dtFinal)
'''    strSql = strSql & " AND DV.intContribuinte BETWEEN " & codInicial & " AND " & codFinal
'''    strSql = strSql & " AND DV.intContribuinte = CO.PKId "
'''    strSql = strSql & " AND DV.intOcorrencia = OC.PKId "
'''    strSql = strSql & " AND DV.intDocumentosEmitidos = DE.PKId "
'''
'''    strSql = strSql & " GROUP BY "
'''    strSql = strSql & " DV.strInscricao , DV.intContribuinte, CO.strNome, "
'''    strSql = strSql & " DE.strDescricao , OC.strDescricao ,"
'''    strSql = strSql & " DV.intDocumentosEmitidos "
'''
'''strQuerryAnalitico = strSql
'''End Function
'''
'''Private Function strQuerrySintetico() As String
'''Dim strSql    As String
'''Dim Inicial   As Date
'''Dim Final     As Date
'''    Inicial = CVDate(txt_Devolucao)
'''    Final = CVDate(txt_Ate)
'''    strSql = ""
'''    strSql = strSql & " SELECT COUNT(*) as qtdDocumento, DV.intOcorrencia, OC.strDescricao "
'''    strSql = strSql & " FROM " & gstrDevolucao & " DV, "
'''    strSql = strSql & gstrOcorrencia & " OC "
'''    strSql = strSql & " WHERE DV.intOcorrencia = OC.PKId "
'''    strSql = strSql & " AND DV.dtmDevolucao BETWEEN " & gstrConvDtParaSql(Inicial) & " AND " & gstrConvDtParaSql(Final)
'''    strSql = strSql & " GROUP BY DV.intOcorrencia, OC.strDescricao "
'''strQuerrySintetico = strSql
'''End Function


Private Sub mskstrInscricaoAnterior_GotFocus()
    MarcaCampo mskstrInscricaoAnterior
End Sub

Private Sub mskstrInscricaoAnterior_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", mskstrInscricaoAnterior
End Sub

Private Sub mskstrInscricaoAnterior2_GotFocus()
    MarcaCampo mskstrInscricaoAnterior2
End Sub

Private Sub mskstrInscricaoAnterior2_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", mskstrInscricaoAnterior2
End Sub
