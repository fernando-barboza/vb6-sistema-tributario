VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptDocumentoDiverso 
   Caption         =   "Documentos Diversos"
   ClientHeight    =   8370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9210
   MDIChild        =   -1  'True
   WindowState     =   2  'Maximized
   _ExtentX        =   16245
   _ExtentY        =   14764
   SectionData     =   "DocumentoDiverso.dsx":0000
End
Attribute VB_Name = "rptDocumentoDiverso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adoResultado As ADODB.Recordset
Private Sub ActiveReport_Activate()
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, True
End Sub

Private Sub ActiveReport_Deactivate()
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, True
End Sub


Private Sub ActiveReport_QueryClose(Cancel As Integer, CloseMode As Integer)
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, True
End Sub

Private Sub ActiveReport_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
End Sub

Private Sub ActiveReport_ReportStart()
    On Error Resume Next
    PadronizaToolBarRelatorio Me, lblExercicio
    LeImagemLogotipo imgBrasao, imgLogotipo, txtNomeFantasia, txtEstado
    lblRelatorio = Me.Caption
End Sub

Private Sub ActiveReport_ToolbarClick(ByVal Tool As DDActiveReports2.DDTool)
    Dim vnt As Variant
    If Tool.ID = 14 Then
        ActiveReport_KeyPress 27
    ElseIf Tool.ID = 15 Then
        AbreOpcoesExportacao Me
    ElseIf Tool.ID = 16 Then
        Configura_Relatorio Me, True
    End If
End Sub

Private Function strQuerryContribuinte() As String
Dim strSql As String
    strSql = ""
    strSql = strSql & " SELECT "
    strSql = strSql & " strNome "
    strSql = strSql & " FROM "
    strSql = strSql & gstrContribuinte
    strSql = strSql & " WHERE "
    strSql = strSql & " PKId = " & Val(txtintContribuinte.Text)
strQuerryContribuinte = strSql
End Function

Private Function strQueryCidade()
Dim strSql  As String
    strSql = ""
    strSql = strSql & " SELECT MU.PKId, MU.strDescricao "
    strSql = strSql & " FROM "
    strSql = strSql & gstrCidade & " MU, "
    strSql = strSql & gstrContribuinte & " CO "
    strSql = strSql & "WHERE "
    strSql = strSql & " MU.PKId = CO.intMunicipio "
    strSql = strSql & " AND CO.PKID = " & Val(gstrENulo(txtintContribuinte))
    strSql = strSql & " OR MU.intUF IS NULL "
strQueryCidade = strSql
End Function

Private Function strQueryBairroEstado()
Dim strSql As String
    strSql = ""
    strSql = strSql & "SELECT UF.PKId, UF.strEstado, UF.strSigla, CO.intCEP, CO.intBairro, BA.strDescricao "
    strSql = strSql & "FROM " & gstrUF & " UF, "
    strSql = strSql & gstrContribuinte & " CO, "
    strSql = strSql & gstrBairro & " BA "
    strSql = strSql & "WHERE "
    strSql = strSql & " UF.PKId = CO.intUF "
    strSql = strSql & " AND BA.PKId = CO.intBairro "
    strSql = strSql & " AND CO.PKID = " & Val(gstrENulo(txtintContribuinte))

strQueryBairroEstado = strSql
End Function

Private Sub Detail_Format()
Dim strSql As String
    
    If txtintContribuinte.Text <> "" Then
        
        Set gobjBanco = New clsBanco
        txt_strNomeDoCara.Text = ""
        If gobjBanco.CriaADO(strQuerryContribuinte, 5, adoResultado) Then
            With adoResultado
                Do While Not .EOF
                    txt_strNomeDoCara.Text = gstrENulo(!strNome)
                    .MoveNext
                Loop
            End With
        End If
        Set gobjBanco = New clsBanco
        txt_Cidade.Text = ""
        If gobjBanco.CriaADO(strQueryCidade, 5, adoResultado) Then
            With adoResultado
                Do While Not .EOF
                    txt_Cidade.Text = gstrENulo(!strDescricao)
                    .MoveNext
                Loop
            End With
        End If
        Set gobjBanco = New clsBanco
        txt_Endereco.Text = ""
        strSql = ""
        strSql = gstrQueryLogradouro(gstrContribuinte, _
                                     "PKId = " & Val(gstrENulo(txtintContribuinte)), _
                                     "intLogradouro")
        If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
            With adoResultado
                Do While Not .EOF
                     If IsNull(!strComplemento) Then
                        txt_Endereco = gstrENulo(!Logradouro) + ", " + gstrENulo(!intNumero)
                     Else
                        txt_Endereco = gstrENulo(!Logradouro) + ", " + gstrENulo(!intNumero) + " - " + gstrENulo(!strComplemento)
                     End If
                    .MoveNext
                Loop
            End With
        End If

        Set gobjBanco = New clsBanco
        txt_BairroEstado.Text = ""
        If gobjBanco.CriaADO(strQueryBairroEstado, 5, adoResultado) Then
            With adoResultado
                Do While Not .EOF
                    txt_BairroEstado = gstrENulo(!strDescricao) + " - " + Trim(txt_Cidade.Text) + " - " + gstrENulo(!strEstado) + " - " + gstrENulo(!strSigla)
                    txt_CEP.Text = gstrCEPFormatado(gstrENulo(!intCep))
                    .MoveNext
                Loop
            End With
        End If
    End If
End Sub

Private Sub PageFooter_Format()
    lblPagina.Caption = "Página " & pageNumber
End Sub

Private Sub PageHeader_Format()
    lblDataHora = gstrDataDoSistema(True, , True)
End Sub

Private Sub ReportFooter_Format()
    MostraEmissorRelatorio Me
End Sub
