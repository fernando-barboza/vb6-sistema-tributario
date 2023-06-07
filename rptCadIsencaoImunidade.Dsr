VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptCadIsencaoImunidade 
   Caption         =   "Tributario - rptCadIsencaoImunidade (ActiveReport)"
   ClientHeight    =   10815
   ClientLeft      =   0
   ClientTop       =   285
   ClientWidth     =   15240
   MDIChild        =   -1  'True
   _ExtentX        =   26882
   _ExtentY        =   19076
   SectionData     =   "rptCadIsencaoImunidade.dsx":0000
End
Attribute VB_Name = "rptCadIsencaoImunidade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ActiveReport_Activate()
  If adoDataControl.Recordset.RecordCount = 0 Then
     ExibeMensagem "Não existe(m) registro(s) com os dados solicitados."
     Unload Me
     Exit Sub
  End If
  HabilitaDesabilitaBotao1 False, gstrBtnArquivo, True
End Sub

Private Sub ActiveReport_Deactivate()
  HabilitaDesabilitaBotao1 True, gstrBtnArquivo, True
End Sub

Private Sub ActiveReport_QueryClose(Cancel As Integer, CloseMode As Integer)
  HabilitaDesabilitaBotao1 True, gstrBtnArquivo, True
End Sub

Private Sub ActiveReport_ReportStart()
  PadronizaToolBarRelatorio Me, lblExercicio
  LeImagemLogotipo imgBrasao, imgLogotipo, txtNomeFantasia, txtEstado
End Sub

Private Sub ActiveReport_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyEscape Then
     Unload Me
  End If
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

Private Sub Detail_Format()
  PreencheReceitas adoDataControl.Recordset("pkID")
  PreenchePeriodo adoDataControl.Recordset("pkID")
End Sub

Private Sub GroupHeader2_Format()
  txtstrInscricao.Text = gstrFormataInscricao(txtstrInscricao.Text, adoDataControl.Recordset("intUtilizacao"))
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

Private Sub PreenchePeriodo(lngPkid As Long)
Dim strSql          As String
Dim adoResultado    As ADODB.Recordset
    
  strSql = ""
  strSql = strSql & "Select "
  strSql = strSql & "IP.Pkid, "
  strSql = strSql & "IP.dtmData, "
  strSql = strSql & "IP.dtmInicial, "
  strSql = strSql & "IP.dtmFinal, "
  strSql = strSql & "IP.bytPosicao, "
  'Alterações Fernanda
  strSql = strSql & gstrISNULL("IP.strCodigoProcesso", "''") & strCONCAT & "'/'" & strCONCAT & gstrCONVERT(CDT_VARCHAR, (gstrISNULL("IP.intExercicioProcesso", "''"))) & strCONCAT & "'-'" & strCONCAT & gstrCONVERT(CDT_VARCHAR, (gstrISNULL("IP.bitDigitoProcesso", "''"))) & " strProcesso , "
  'Fim Alterações Fernanda
  strSql = strSql & "IP.bytCancelamento "
  'ACRESCENTAR OBSERVAÇÃO ???
  strSql = strSql & "FROM "
  strSql = strSql & gstrIsencaoImunidade & " I, "
  strSql = strSql & gstrIsencaoPeriodo & " IP "
  strSql = strSql & "WHERE "
  strSql = strSql & "I.Pkid = IP.Intisencaoimunidade AND "
  strSql = strSql & "I.Pkid = " & lngPkid & " "
  strSql = strSql & "Order by IP.DTMINICIAL DESC"
    
  Set gobjBanco = New clsBanco
  Set adoResultado = New ADODB.Recordset
  
  If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
     If Not adoResultado.EOF Then
        
        With rptSubIsencaoImuPeriodo
          If bytDBType = EDatabases.SQLServer Then
             .adoDataControl.ConnectionString = "driver={SQL Server};Server=" & gstrServidor & ";database=" & gstrDatabase & ";uid=" & gstrUsername & ";pwd=" & gstrPassword & ";"
          Else
             .adoDataControl.ConnectionString = "Provider=MSDAORA.1;Password=" & gstrPassword & ";User ID=" & gstrUsername & ";Data Source=" & gstrServidor & ";Persist Security Info=True"
            End If
            .adoDataControl.Source = strSql
          Set .adoDataControl.Recordset = adoResultado
           
        End With
        
        Set subPeriodo.object = rptSubIsencaoImuPeriodo
    End If
  End If
End Sub

Private Sub PreencheReceitas(lngPkid As Long)
Dim strSql          As String
Dim adoResultado    As ADODB.Recordset
  
  strSql = ""
  strSql = strSql & "Select "
  strSql = strSql & "R.PKID intReceita, "
  strSql = strSql & "R.Strdescricao strReceita, "
  strSql = strSql & "IR.DBLALIQUOTA "
  strSql = strSql & "From "
  strSql = strSql & gstrIsencaoImunidade & " I, "
  strSql = strSql & gstrIsencaoReceita & " IR, "
  strSql = strSql & gstrReceita & " R "
  strSql = strSql & "Where "
  strSql = strSql & "I.Pkid = IR.Intisencaoimunidade AND "
  strSql = strSql & "R.Pkid = IR.Intreceita AND "
  strSql = strSql & "I.Pkid = " & lngPkid
    
  Set gobjBanco = New clsBanco
  Set adoResultado = New ADODB.Recordset
    
  If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
     If Not adoResultado.EOF Then
        
        With rptSubIsencaoImuReceita
          If bytDBType = EDatabases.SQLServer Then
             .adoDataControl.ConnectionString = "driver={SQL Server};Server=" & gstrServidor & ";database=" & gstrDatabase & ";uid=" & gstrUsername & ";pwd=" & gstrPassword & ";"
          Else
             .adoDataControl.ConnectionString = "Provider=MSDAORA.1;Password=" & gstrPassword & ";User ID=" & gstrUsername & ";Data Source=" & gstrServidor & ";Persist Security Info=True"
          End If
          Set .adoDataControl.Recordset = adoResultado
        End With
        
        Set subReceita.object = rptSubIsencaoImuReceita
    End If
  End If
End Sub



