VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmDocCertidaoValorVenal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Certidão de Valor Venal"
   ClientHeight    =   2070
   ClientLeft      =   4200
   ClientTop       =   3825
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2070
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtstrRequerente 
      Height          =   315
      Left            =   1410
      MaxLength       =   50
      TabIndex        =   3
      Top             =   960
      Width           =   3795
   End
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   2025
      Left            =   60
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   45
      Width           =   5310
      _ExtentX        =   9366
      _ExtentY        =   3572
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Certidão de Valor Venal"
      TabPicture(0)   =   "frmDocCertidaoValorVenal.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblInscricaoInicial"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblstrDescricao"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblintexrcicio"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "dbcintInscricaoInicial"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtbitDigito"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtintExercicio"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtstrCodigo"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txt_intExercicio"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      Begin VB.TextBox txt_intExercicio 
         Alignment       =   1  'Right Justify
         CausesValidation=   0   'False
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
         Left            =   1350
         MaxLength       =   4
         TabIndex        =   4
         Top             =   1260
         Width           =   585
      End
      Begin VB.TextBox txtstrCodigo 
         Alignment       =   1  'Right Justify
         CausesValidation=   0   'False
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
         Left            =   1350
         MaxLength       =   15
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   1590
         Width           =   825
      End
      Begin VB.TextBox txtintExercicio 
         CausesValidation=   0   'False
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
         Left            =   2190
         MaxLength       =   4
         TabIndex        =   6
         Top             =   1590
         Width           =   465
      End
      Begin VB.TextBox txtbitDigito 
         Alignment       =   1  'Right Justify
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
         Left            =   2670
         MaxLength       =   2
         TabIndex        =   7
         Top             =   1590
         Width           =   285
      End
      Begin MSDataListLib.DataCombo dbcintInscricaoInicial 
         Height          =   315
         Left            =   1350
         TabIndex        =   2
         Top             =   540
         Width           =   2190
         _ExtentX        =   3863
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Requerente"
         Height          =   315
         Left            =   150
         TabIndex        =   10
         Top             =   990
         Width           =   1125
      End
      Begin VB.Label lblintexrcicio 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Exercício"
         Height          =   195
         Left            =   525
         TabIndex        =   8
         Top             =   1305
         Width           =   675
      End
      Begin VB.Label lblstrDescricao 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Processo"
         Height          =   195
         Left            =   555
         TabIndex        =   9
         Top             =   1635
         Width           =   660
      End
      Begin VB.Label lblInscricaoInicial 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Inscrição"
         Height          =   195
         Left            =   585
         TabIndex        =   1
         Top             =   660
         Width           =   645
      End
   End
End
Attribute VB_Name = "frmDocCertidaoValorVenal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const MODELO            As String = "CERTIDÃO DE VALOR VENAL"
Dim Documentos()        As cWordWrapper
Dim XArrayTabela        As XArrayDB
Dim XArrayAlinhaColunas As XArrayDB

Private Sub dbcintInscricaoInicial_Click(Area As Integer)
    DropDownDataCombo dbcintInscricaoInicial, Me, Area
End Sub

Private Sub dbcintInscricaoInicial_GotFocus()
    MarcaCampo dbcintInscricaoInicial
End Sub

Private Sub dbcintInscricaoInicial_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintInscricaoInicial, Me, , KeyCode, Shift
End Sub

Private Sub dbcintInscricaoInicial_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintInscricaoInicial
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 1156
    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
End Sub

Private Sub Form_Load()
    dbcintInscricaoInicial.Tag = strQueryInscricao(dbcintInscricaoInicial) & ";strInscricao"
End Sub

Private Function strQueryInscricao(objDataCombo As DataCombo)
Dim strSQL As String

    strSQL = "SELECT " & _
                    gintPkidFixo & " Pkid, " & _
                    gstrRIGHT("LA.strInscricao", gintRetornaTamanhoMascara(TYP_IMOBILIARIA)) & " strInscricao " & _
             " FROM " & _
                      gstrLancamentoAlfa & " LA " & _
             " WHERE " & _
                     " LA.Intutilizacao = " & TYP_IMOBILIARIA
             If Len(Trim(objDataCombo.Text)) > 0 Then
                 strSQL = strSQL & " AND LA.strInscricao Like '" & String(gintLenInscricao - gintRetornaTamanhoMascara(TYP_IMOBILIARIA), "0") & objDataCombo.Text & "%'"
             End If
             strSQL = strSQL & " GROUP BY LA.strInscricao" & _
                " ORDER BY LA.strInscricao"
        
    strQueryInscricao = strSQL

End Function

Public Sub MantemForm(ByVal strModoOperacao As String)
        
    Select Case UCase(strModoOperacao)
    
        Case Is = UCase(gstrPreencherLista)
            If TypeOf Me.ActiveControl Is DataCombo Then
                LeDaTabelaParaObj "", Me.ActiveControl, strQueryInscricao(Me.ActiveControl)
            End If
        Case Is = UCase(gstrImprimir)
            If blnDadosOK Then
                OpenWordDocumentOverTemplate
            End If
        Case Is = UCase(gstrNovo)
            Novo
    End Select
                    
End Sub

Private Sub Novo()
  dbcintInscricaoInicial.Text = ""
  Set dbcintInscricaoInicial.RowSource = Nothing
  txtbitDigito.Text = ""
  txtintExercicio.Text = ""
  txtstrCodigo.Text = ""
  txt_intExercicio.Text = ""
  dbcintInscricaoInicial.SetFocus
End Sub

Private Function blnDadosOK() As Boolean

    If Not dbcintInscricaoInicial.MatchedWithList Then
        ExibeMensagem "Selecione uma Inscrição Inicial válida."
        If dbcintInscricaoInicial.Enabled Then dbcintInscricaoInicial.SetFocus
        Exit Function
    End If
    
    If Trim(txt_intExercicio.Text) = "" Then
        ExibeMensagem "Favor informar o exercício."
        If txt_intExercicio.Enabled Then txt_intExercicio.SetFocus
        Exit Function
    End If
            
    If Len(Trim(txtstrCodigo.Text)) > 0 Or Len(Trim(txtintExercicio.Text)) > 0 Or Len(Trim(txtbitDigito.Text)) > 0 Then
    
        If Trim(txtstrCodigo.Text) = "" Then
            ExibeMensagem "Favor informar o Código do Processo."
            If txtstrCodigo.Enabled Then txtstrCodigo.SetFocus
            Exit Function
        End If
        
        If Trim(txtintExercicio.Text) = "" Then
            ExibeMensagem "Favor informar o Exercício do Processo."
            If txtintExercicio.Enabled Then txtintExercicio.SetFocus
            Exit Function
        End If
        
        If Trim(txtbitDigito.Text) = "" Then
            ExibeMensagem "Favor informar o Dígito do Processo."
            If txtbitDigito.Enabled Then txtbitDigito.SetFocus
            Exit Function
        End If
        
        If Not VerificaEmpenhoProcesso(txtstrCodigo, txtbitDigito, txtintExercicio) Then
            ExibeMensagem "O Processo informado não é válido."
            txtstrCodigo.SetFocus
            Exit Function
        End If
    
    End If
    
    blnDadosOK = True

End Function

Private Function strQuery() As String

Dim strSQL As String

    strSQL = "SELECT LA.Pkid, " & gstrRIGHT("LA.strInscricao", gintRetornaTamanhoMascara(TYP_IMOBILIARIA)) & " Inscrição, "
    strSQL = strSQL & " la.strnomeproprietario As Proprietario,"
    strSQL = strSQL & " la.strpromissario As Promissario,"
    strSQL = strSQL & " LI.dblAreaterreno Area,"
    strSQL = strSQL & " Predios.MedidaDaArea,"
    strSQL = strSQL & " LA.strInscricaoAuxiliar InscricaoAux,"
    strSQL = strSQL & " LA.intExercicio Exercício,"
    strSQL = strSQL & " LA.strLogradouro Logradouro ,"
    strSQL = strSQL & " LA.strNumero Número,"
    strSQL = strSQL & " LA.strComplemento Complemento ,"
    strSQL = strSQL & " LA.strBairro Bairro,"
    strSQL = strSQL & " LA.intCep CEP,"
    strSQL = strSQL & " LI.dblValorVenalTerreno ValorVenalTerreno,"
    strSQL = strSQL & " LI.dblValorTerrenoExcedente ValorTerrenoExcedente, "
    strSQL = strSQL & " LI.DBLVALORVENALTERRENO + LI.DBLVALORTERRENOEXCEDENTE ValorVenalFinal,"
    strSQL = strSQL & gstrISNULL("LI.dblValorReferencia", "0") & " ValorReferencia,"
    strSQL = strSQL & " LI.strIndexador ,"
    strSQL = strSQL & " LI.strLote Lote,"
    strSQL = strSQL & " LI.strQuadra Quadra,"
    strSQL = strSQL & " LI.strLoteamento Loteamento,"
    strSQL = strSQL & " LI.dblAreaTerreno AreaTerreno,"
    strSQL = strSQL & " LI.dblAreaExcedente AreaExcedente, "
    strSQL = strSQL & gstrISNULL("Predios.Registros", "0") & " Registros, "
    strSQL = strSQL & gstrISNULL("Predios.ValorVenalPredio", "0") & " ValorVenalPredios, "
    strSQL = strSQL & gstrISNULL("Predios.MedidaDaArea", "0") & " MedidaDaArea,"
    strSQL = strSQL & gstrISNULL("LI.dblValorVenalTerreno", "0") & " + " & gstrISNULL("LI.dblValorTerrenoExcedente", "0") & " + " & gstrISNULL("Predios.ValorVenalPredio", "0") & " ValorVenalTotal"
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrLancamentoAlfa & " LA, "
    strSQL = strSQL & gstrLancamentoIPTU & " LI,"
    strSQL = strSQL & " (SELECT COUNT(LP.Pkid) Registros,"
    strSQL = strSQL & " LP.intLancamentoIPTU,"
    strSQL = strSQL & " SUM(LP.dblValorVenalPredio) ValorVenalPredio,"
    strSQL = strSQL & " SUM(LP.dblMedidaDaArea) MedidaDaArea"
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrLancamentoPredioIPTU & " LP, "
    strSQL = strSQL & gstrLancamentoAlfa & " LA2, "
    strSQL = strSQL & gstrLancamentoIPTU & " LI2 "
    strSQL = strSQL & " WHERE LA2.Pkid = LI2.intLancamentoAlfa AND "
    strSQL = strSQL & " LA2.strInscricao = '"
    strSQL = strSQL & String(gintLenInscricao - Len(dbcintInscricaoInicial.Text), "0") & dbcintInscricaoInicial.Text & "'"
    strSQL = strSQL & " AND LA2.dtmDtCancelamento Is Null "
    strSQL = strSQL & " AND LI2.Pkid = LP.intlancamentoiptu "
    strSQL = strSQL & " GROUP BY LP.intLancamentoIPTU) Predios"
    strSQL = strSQL & " WHERE  LA.strInscricao = '"
    strSQL = strSQL & String(gintLenInscricao - Len(dbcintInscricaoInicial.Text), "0") & dbcintInscricaoInicial.Text & "' AND"
    strSQL = strSQL & " LA.Pkid = LI.intLancamentoAlfa AND"
    strSQL = strSQL & " LA.intExercicio = " & txt_intExercicio & " AND"
    strSQL = strSQL & " Predios.intLancamentoIPTU " & strOUTJOracle & "=" & strOUTJSQLServer & " LI.Pkid AND "
    strSQL = strSQL & " LA.dtmDtCancelamento Is Null "
    
    strQuery = strSQL

End Function

Private Sub OpenWordDocumentOverTemplate()
Dim blpMsg          As Boolean
Dim intFor          As Integer
Dim stpDocument     As String
Dim stpTemplate     As String
Dim objFileSystem   As Scripting.FileSystemObject
Dim stpTemplatePath As String
Dim stpDocumentPath As String
Dim adoResultado    As ADODB.Recordset
Dim adoTabela       As ADODB.Recordset
'------------------------------------
Dim lngNumero       As Long
Dim adoNumero       As ADODB.Recordset
        
    Screen.MousePointer = vbHourglass
    
    stpTemplate = gstrDirDocumentos & "Documentos\" & App.ProductName & "\WordModelos\" & MODELO & ".dot"
    stpTemplatePath = gstrDirDocumentos & "Documentos\" & App.ProductName & "\WordModelos\"
    stpDocumentPath = gstrDirDocumentos & "Documentos\" & App.ProductName & "\WordGravados\"
   
    Set objFileSystem = New Scripting.FileSystemObject
    
    If objFileSystem.FolderExists(stpTemplatePath) Then
        If objFileSystem.FileExists(stpTemplate) Then
            If objFileSystem.FolderExists(stpDocumentPath) Then
                Set gobjBanco = New clsBanco
                If gobjBanco.CriaADO(strQuery, 5, adoResultado) Then
                    If Not adoResultado.EOF Then
                        adoResultado.MoveFirst
                        ReDim Documentos(1 To adoResultado.RecordCount)
                        Do While Not adoResultado.EOF
                            
                            Set gobjBanco = New clsBanco
                            gobjBanco.ExecutaBeginTrans
                            
                            lngNumero = CLng(glngRetornaProximoNumeroGuia(gstrEmpresa, "intnumerocertidaovalorvenal"))
                            If Val(lngNumero) = 0 Then
                                Screen.MousePointer = vbDefault
                                Exit Sub
                            Else
                                gobjBanco.ExecutaCommitTrans
                            End If
                            
                            'Set gobjBanco = New clsBanco
                            'If gobjBanco.CriaADO("Select Max(intnumerocertidaovalorvenal) + 1 as Numero From " & gstrEmpresa, 5, adoNumero) Then
                            '    If Not adoNumero.EOF Then
                            '        If Val(gstrENulo(adoNumero!Numero)) > "0" Then
                            '            lngNumero = CLng(gstrENulo(adoNumero!Numero))
                            '        Else
                            '            lngNumero = 1
                            '        End If
                            '        gobjBanco.Execute ("Update " & gstrEmpresa & " Set intnumerocertidaovalorvenal = " & lngNumero)
                            '    Else
                            '        Screen.MousePointer = vbDefault
                            '        Exit Sub
                            '    End If
                            'Else
                            '    Screen.MousePointer = vbDefault
                            '    Exit Sub
                            'End If
                            
                            
                            stpDocument = stpDocumentPath & MODELO & "_" & gstrFormataInscricao(adoResultado("Inscrição"), 1) & "_" & adoResultado("Exercício") & "_" & Format$(lngNumero, "0000000000") & ".doc"
                            MontaArray adoResultado("Pkid"), adoResultado
                            blpMsg = True
                            If objFileSystem.FileExists(stpDocument) Then
                                If blpMsg Then objFileSystem.DeleteFile stpDocument, True
                            End If
                        
                            Set Documentos(adoResultado.AbsolutePosition) = New cWordWrapper
                    
                            Documentos(adoResultado.AbsolutePosition).GetContainer
                            Documentos(adoResultado.AbsolutePosition).DocumentTemplatePath = stpTemplate
                            Documentos(adoResultado.AbsolutePosition).DocumentPath = stpDocument
                            Documentos(adoResultado.AbsolutePosition).DocumentFormat = WORDOPENFORMATDOCUMENT
                            Documentos(adoResultado.AbsolutePosition).DocumentOpen
                                                            
                            'Substituição dos Campos
                            
                            'Monta a forma de alinhamento das colunas
                            
                            MontaAlinhamento
                            Documentos(adoResultado.AbsolutePosition).DocumentInsert "|Tabela1|", , XArrayTabela, XArrayAlinhaColunas
                            Documentos(adoResultado.AbsolutePosition).DocumentReplaceField "|Inscrição|", gstrFormataInscricao(adoResultado!Inscrição, TYP_IMOBILIARIA)
                            Documentos(adoResultado.AbsolutePosition).DocumentReplaceField "|Inscriçãoauxiliar|", gstrENulo(adoResultado!InscricaoAux)
                            
                            Documentos(adoResultado.AbsolutePosition).DocumentReplaceField "|Proprietario|", gstrENulo(adoResultado!Proprietario)
                            Documentos(adoResultado.AbsolutePosition).DocumentReplaceField "|Promissario|", gstrENulo(adoResultado!Promissario)
                            Documentos(adoResultado.AbsolutePosition).DocumentReplaceField "|Areadoterreno|", IIf(IsNull(adoResultado!Area), "0", adoResultado!Area)
                            Documentos(adoResultado.AbsolutePosition).DocumentReplaceField "|Areadaconstrução|", IIf(IsNull(adoResultado!MedidaDaArea), "0", adoResultado!MedidaDaArea)
                            Documentos(adoResultado.AbsolutePosition).DocumentReplaceField "|Exercício|", gstrENulo(adoResultado!Exercício)
                            Documentos(adoResultado.AbsolutePosition).DocumentReplaceField "|Processo|", txtstrCodigo.Text & "/" & txtintExercicio.Text & "-" & txtbitDigito.Text
                            Documentos(adoResultado.AbsolutePosition).DocumentReplaceField "|Logradouro|", gstrENulo(adoResultado!Logradouro) & ", " & gstrENulo(adoResultado!Número) & IIf(IsNull(adoResultado!Complemento), "", ", " & gstrENulo(adoResultado!Complemento))
                            Documentos(adoResultado.AbsolutePosition).DocumentReplaceField "|Bairro|", gstrENulo(adoResultado!Bairro)
                            Documentos(adoResultado.AbsolutePosition).DocumentReplaceField "|CEP|", gstrCEPFormatado(gstrENulo(adoResultado!CEP))
                            Documentos(adoResultado.AbsolutePosition).DocumentReplaceField "|Lote|", IIf(IsNull(adoResultado!Lote) = True, "", "Lote: " & gstrENulo(adoResultado!Lote) & " ")
                            Documentos(adoResultado.AbsolutePosition).DocumentReplaceField "|Quadra|", IIf(IsNull(adoResultado!Quadra) = True, "", "Quadra: " & gstrENulo(adoResultado!Quadra) & " ")
                            Documentos(adoResultado.AbsolutePosition).DocumentReplaceField "|Complemento|", UCase(gstrENulo(adoResultado!Complemento))
                            Documentos(adoResultado.AbsolutePosition).DocumentReplaceField "|Loteamento|", IIf(IsNull(adoResultado!Loteamento) = True, "", "Loteamento: " & gstrENulo(adoResultado!Loteamento) & " ")
                            Documentos(adoResultado.AbsolutePosition).DocumentReplaceField "|Área de Terreno|", gstrConvVrDoSql(adoResultado!AreaTerreno + IIf(IsNull(adoResultado!AreaExcedente) = True, 0, adoResultado!AreaExcedente), 2) & " M² "
                            Documentos(adoResultado.AbsolutePosition).DocumentReplaceField "|Total Área Construída|", IIf((adoResultado!MedidaDaArea = 0), "", "e área total construída de " & gstrConvVrDoSql(adoResultado!MedidaDaArea, 2) & " M² ")
                            Documentos(adoResultado.AbsolutePosition).DocumentReplaceField "|Qtde|", IIf(IsNull(adoResultado!Registros) = True, "", "(" & adoResultado!Registros & " Unidade(s))")
                            Documentos(adoResultado.AbsolutePosition).DocumentReplaceField "|Valor Venal|", gstrExtenso(adoResultado!ValorVenalTotal, 0)
                            Documentos(adoResultado.AbsolutePosition).DocumentReplaceField "|Data|", gstrDataPorExtenso(gstrDataDoSistema)
                            Documentos(adoResultado.AbsolutePosition).DocumentReplaceField "|Valor Venal Terreno|", gstrConvVrDoSql(gstrENulo(adoResultado!ValorVenalTerreno))
                            Documentos(adoResultado.AbsolutePosition).DocumentReplaceField "|Extenso Valor Terreno|", UCase(gstrExtenso(adoResultado!ValorVenalTerreno))
                            Documentos(adoResultado.AbsolutePosition).DocumentReplaceField "|Valor Venal Construcao|", gstrConvVrDoSql(gstrENulo(adoResultado!ValorVenalPredios))
                            Documentos(adoResultado.AbsolutePosition).DocumentReplaceField "|Extenso Construcao|", UCase(gstrExtenso(adoResultado!ValorVenalPredios))
                            Documentos(adoResultado.AbsolutePosition).DocumentReplaceField "|Codigo|", lngNumero & "/" & gstrENulo(adoResultado!Exercício)
                            Documentos(adoResultado.AbsolutePosition).DocumentReplaceField "|Usuario|", UCase(gstrNomeUsuario)
                            Documentos(adoResultado.AbsolutePosition).DocumentReplaceField "|Requerente|", UCase(txtstrRequerente)
                            Documentos(adoResultado.AbsolutePosition).DocumentReplaceField "|CodigoBarras|", gstrENulo(adoResultado!Inscrição) & CStr(lngNumero) & gstrENulo(adoResultado!Exercício)
                            Documentos(adoResultado.AbsolutePosition).DocumentSave
                            adoResultado.MoveNext
                        Loop
                    Else
                        ExibeMensagem "Não existem dados para a Automação."
                    End If
                End If
            Else
                MsgBox "A pasta de documentos : " & stpDocumentPath & " não foi localizada. A operação não pode ser realizada.", vbOKOnly + vbInformation, "Mensagem ao usuário"
            End If
        
        Else
            MsgBox "O modelo de documento do Microsoft Word : " & stpTemplate & " não foi localizado. A operação não pode ser realizada.", vbOKOnly + vbInformation, "Mensagem ao usuário"
        End If
    
    Else
       MsgBox "A pasta de modelos de documentos : " & stpTemplatePath & " não foi localizada. A operação não pode ser realizada.", vbOKOnly + vbInformation, "Mensagem ao usuário"
    End If
    
    Set objFileSystem = Nothing
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
End Sub

Private Sub MontaArray(intLancamento As Long, adoTotal As ADODB.Recordset)
Dim varAux          As Variant
Dim strSQL          As String
Dim adoValores      As New ADODB.Recordset
Dim varBookMark     As Variant

    'strSQL = strQuery & " AND LA.strInscricao =" & String(gintLenInscricao - Len(strInscricao), "0") & strInscricao
    'Set gobjBanco = New clsBanco
    
    'If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        
    Set adoValores = adoTotal
        
    'Vamos identificar o registro atual, pois se nao fizermos ele se perde ao retirar o filtro
    varBookMark = adoValores.Bookmark
    
    adoValores.Filter = "Pkid = " & intLancamento
    
        Set XArrayTabela = New XArrayDB
        XArrayTabela.Clear
        XArrayTabela.ReDim 0, 3, 0, 1
        If Not adoValores.EOF Then
            'adoResultado.MoveFirst
            'Nome dos Campos (coluna 0)
            varAux = ""
            XArrayTabela(0, 0) = varAux
            varAux = "Valor Venal Terreno:"
            XArrayTabela(1, 0) = varAux
            varAux = "Valor Venal Prédio(s):"
            XArrayTabela(2, 0) = varAux
            varAux = "Valor Venal Total:"
            XArrayTabela(3, 0) = varAux
    
            'Preenche com os Valores dos respectivos campos(coluna 1)
            varAux = "R$"
            XArrayTabela(0, 1) = varAux
            varAux = gstrConvVrDoSql(adoValores!ValorVenalFinal, 2)
            XArrayTabela(1, 1) = varAux
            varAux = gstrConvVrDoSql(adoValores!ValorVenalPredios, 2)
            XArrayTabela(2, 1) = varAux
            varAux = gstrConvVrDoSql(adoValores!ValorVenalTotal, 2)
            XArrayTabela(3, 1) = varAux
            
            'Preenche os FMPS dos rescpectivos campos(coluna 2)
            
            If adoValores!ValorReferencia <> 0 Then
                XArrayTabela.ReDim 0, 3, 0, 2
                varAux = adoValores!Strindexador
                XArrayTabela(0, 2) = varAux
                varAux = gstrConvVrDoSql((adoValores!ValorVenalTerreno + adoValores!ValorTerrenoExcedente) _
                                            / adoValores!ValorReferencia, 2)
                XArrayTabela(1, 2) = varAux
                
                varAux = gstrConvVrDoSql(adoValores!ValorVenalPredios / adoValores!ValorReferencia, 2)
                XArrayTabela(2, 2) = varAux
                
                varAux = gstrConvVrDoSql((adoValores!ValorVenalTerreno + adoValores!ValorTerrenoExcedente + _
                                            adoValores!ValorVenalPredios) / adoValores!ValorReferencia, 2)
                XArrayTabela(3, 2) = varAux
            End If
        End If

    adoValores.Filter = adFilterNone
    
    'Vamos retornar ao registro antes do filtro
    adoValores.Bookmark = varBookMark
    
    'End If
    

End Sub

Private Sub MontaAlinhamento()
'Array que contém as Colunas da Tabela
    
    Set XArrayAlinhaColunas = New XArrayDB
    With XArrayAlinhaColunas
        .Clear
        .ReDim 0, 0, 0, 2
        .Value(0, 0) = WORDALIGNPARAGRAPHLEFT
        .Value(0, 1) = WORDALIGNPARAGRAPHRIGHT
        .Value(0, 2) = WORDALIGNPARAGRAPHRIGHT
    End With
End Sub

Private Sub txt_intExercicio_GotFocus()
    MarcaCampo txt_intExercicio
End Sub

Private Sub txt_intExercicio_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_intExercicio
End Sub

Private Sub txtstrCodigo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtstrCodigo
End Sub
