VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmArquivoDistribuidor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Arquivo Distribuidor"
   ClientHeight    =   3825
   ClientLeft      =   4170
   ClientTop       =   4530
   ClientWidth     =   5670
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   3735
      Left            =   60
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   45
      Width           =   5520
      _ExtentX        =   9737
      _ExtentY        =   6588
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Arquivo Distribuidor"
      TabPicture(0)   =   "frmArquivoDistribuidor.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl_Potocolo"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl_Exercicio"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label4"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label5"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lbl_strCaminho(3)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "prg_Status"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cdl_Abrir"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "dbc_intLote"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "dbc_intAdvogado"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txt_intExercicio"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txt_intProtocolo"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txt_intQuantidade"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txt_intQntdePorArquivo"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txt_intArquivoInicial"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "cmd_Caminho"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txt_strCaminho"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).ControlCount=   19
      Begin VB.TextBox txt_strCaminho 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#.##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   1560
         TabIndex        =   8
         Top             =   2865
         Width           =   3135
      End
      Begin VB.CommandButton cmd_Caminho 
         Height          =   315
         Left            =   4725
         Picture         =   "frmArquivoDistribuidor.frx":001C
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         Tag             =   "585"
         ToolTipText     =   "Ativa Cadastro de Bairro"
         Top             =   2880
         Width           =   360
      End
      Begin VB.TextBox txt_intArquivoInicial 
         Height          =   315
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   6
         Top             =   1920
         Width           =   1035
      End
      Begin VB.TextBox txt_intQntdePorArquivo 
         Height          =   315
         Left            =   3990
         MaxLength       =   10
         TabIndex        =   5
         Top             =   1470
         Width           =   1035
      End
      Begin VB.TextBox txt_intQuantidade 
         Height          =   315
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   4
         Top             =   1470
         Width           =   1035
      End
      Begin VB.TextBox txt_intProtocolo 
         Height          =   315
         Left            =   1560
         MaxLength       =   10
         TabIndex        =   2
         Top             =   1035
         Width           =   1035
      End
      Begin VB.TextBox txt_intExercicio 
         Height          =   315
         Left            =   3990
         MaxLength       =   4
         TabIndex        =   3
         Top             =   1035
         Width           =   525
      End
      Begin MSDataListLib.DataCombo dbc_intAdvogado 
         Height          =   315
         Left            =   1560
         TabIndex        =   7
         Top             =   2400
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   556
         _Version        =   393216
         IntegralHeight  =   0   'False
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbc_intLote 
         Height          =   315
         Left            =   1560
         TabIndex        =   0
         Top             =   570
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         _Version        =   393216
         IntegralHeight  =   0   'False
         Text            =   ""
      End
      Begin MSComDlg.CommonDialog cdl_Abrir 
         Left            =   0
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComctlLib.ProgressBar prg_Status 
         Height          =   195
         Left            =   60
         TabIndex        =   18
         Top             =   3450
         Visible         =   0   'False
         Width           =   5385
         _ExtentX        =   9499
         _ExtentY        =   344
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label lbl_strCaminho 
         AutoSize        =   -1  'True
         Caption         =   "Caminho"
         Height          =   195
         Index           =   3
         Left            =   870
         TabIndex        =   17
         Top             =   2970
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Lote"
         Height          =   195
         Left            =   1170
         TabIndex        =   16
         Top             =   690
         Width           =   315
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Advogado"
         Height          =   195
         Left            =   750
         TabIndex        =   15
         Top             =   2520
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Nº Arquivo Inicial"
         Height          =   195
         Left            =   270
         TabIndex        =   14
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Qtde.por Arquivo"
         Height          =   195
         Left            =   2715
         TabIndex        =   13
         Top             =   1590
         Width           =   1200
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Qtde. Total"
         Height          =   195
         Left            =   690
         TabIndex        =   12
         Top             =   1590
         Width           =   795
      End
      Begin VB.Label lbl_Exercicio 
         AutoSize        =   -1  'True
         Caption         =   "Exercício"
         Height          =   195
         Left            =   3180
         TabIndex        =   11
         Top             =   1155
         Width           =   675
      End
      Begin VB.Label lbl_Potocolo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Protocolo"
         Height          =   195
         Left            =   810
         TabIndex        =   10
         Top             =   1155
         Width           =   675
      End
   End
End
Attribute VB_Name = "frmArquivoDistribuidor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim intProtocolo    As Long



Private Sub cmd_Caminho_Click()

On Error GoTo Problema_Na_Rotina

    
    cdl_Abrir.DialogTitle = "Informe o arquivo de destino"
    cdl_Abrir.flags = cdlOFNExplorer
    
    cdl_Abrir.CancelError = True
    cdl_Abrir.Filter = "Arquivo de texto (*.txt)|*.txt"
    cdl_Abrir.filename = "expref"
    cdl_Abrir.ShowSave
    

    If Err.Number = 0 Then
        txt_strCaminho.Text = Left$(cdl_Abrir.filename, InStrRev(cdl_Abrir.filename, "\"))
    End If

    cdl_Abrir.filename = Space$(0)
    
    Exit Sub
    
Problema_Na_Rotina:
    If Err.Number = cdlCancel Then
        txt_strCaminho.Text = Space$(0)
        cdl_Abrir.filename = Space$(0)
    End If

End Sub

Private Sub dbc_intAdvogado_GotFocus()
    MarcaCampo dbc_intAdvogado
End Sub

Private Sub dbc_intLote_GotFocus()
    MarcaCampo dbc_intLote
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 1377
    
    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrSalvar, gstrNovo
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar, gstrAplicar, gstrImprimir
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
        
    Select Case UCase(strModoOperacao)
        Case Is = UCase(gstrSalvar)
            If blnDadosOk Then
                GravarArquivo
            End If
        Case Is = UCase(gstrNovo)
            LimpaObjeto Me
        Case Is = UCase(gstrPreencherLista)
            ToolBarGeral strModoOperacao, gstrExecutivo, False, , Me
    End Select
    
                    
End Sub

Private Function blnDadosOk() As Boolean
Dim objFileSystem
    
    blnDadosOk = False
    
    If Not dbc_intLote.MatchedWithList Then
        ExibeMensagem "Selecione um lote válido."
        dbc_intLote.SetFocus
        Exit Function
    End If
    
    If Not IsNumeric(txt_intProtocolo) Then
        ExibeMensagem "Informe um Número de Protocolo válido."
        txt_intProtocolo.SetFocus
        Exit Function
    End If
    
    If Not IsNumeric(txt_intExercicio) Then
        ExibeMensagem "Informe um Exercicio válido."
        txt_intExercicio.SetFocus
        Exit Function
    End If
    
    If Not IsNumeric(txt_intQuantidade) Then
        ExibeMensagem "Informe uma Quantidade Total válida."
        txt_intQuantidade.SetFocus
        Exit Function
    End If
    
    If Not IsNumeric(txt_intQntdePorArquivo) Then
        ExibeMensagem "Informe uma Quantidade por Arquivo válida."
        txt_intQntdePorArquivo.SetFocus
        Exit Function
    End If
    
    If Not IsNumeric(txt_intArquivoInicial) Then
        ExibeMensagem "Informe um Número de Arquivo Inicial válido."
        txt_intArquivoInicial.SetFocus
        Exit Function
    End If
    
    If Not dbc_intAdvogado.MatchedWithList Then
        ExibeMensagem "Selecione um advogado válido."
        dbc_intAdvogado.SetFocus
        Exit Function
    End If
    
    Set objFileSystem = New Scripting.FileSystemObject
    
    If Not objFileSystem.FolderExists(txt_strCaminho) Then
        ExibeMensagem "Informe um caminho válido."
        txt_strCaminho.SetFocus
        Exit Function
    End If
    
    blnDadosOk = True

End Function


Private Sub GravarArquivo()
Dim strWord             As String
Dim intCont             As Integer
Dim intNumeroArquivo    As Integer
Dim strNomeArquivo      As String
Dim adoResultado        As ADODB.Recordset
Dim strSQL              As String
Dim strPrefeitura       As String
Dim varAux
Dim strNumeroOAB        As String
Dim strUFAdvogado       As String




    Set gobjBanco = New clsBanco
    
    gobjBanco.ExecutaBeginTrans
    
    strSQL = ""
    strSQL = strSQL & "SELECT * FROM " & gstrExecutivoAdvogados & " EA, "
    strSQL = strSQL & gstrUF & " UF "
    strSQL = strSQL & " WHERE EA.pkId = " & dbc_intAdvogado.BoundText
    strSQL = strSQL & " AND EA.intUF = UF.pkId "
        
    If Not gobjBanco.CriaADO(strSQL, 10, adoResultado) Then
        ExibeMensagem "Erro ao consultar os dados do Advogado selecionado."
        gobjBanco.ExecutaRollbackTrans
        Exit Sub
    Else
        strNumeroOAB = adoResultado!strOABNumero
        strUFAdvogado = adoResultado!strsigla
    End If
    
    
    
    strSQL = ""
    strSQL = strSQL & "SELECT strNomeFantasia FROM " & gstrEmpresa
    
    Set gobjBanco = New clsBanco
    
    If Not gobjBanco.CriaADO(strSQL, 10, adoResultado) Then
        ExibeMensagem "Erro ao consultar as informações sobre a prefeitura"
        gobjBanco.ExecutaRollbackTrans
        Exit Sub
    End If
    
    strPrefeitura = adoResultado!strNomeFantasia
    
    strSQL = ""
    strSQL = strSQL & "SELECT "
    
    If bytDBType = SQLServer Then
        strSQL = strSQL & " TOP " & CInt(txt_intQuantidade) & " "
    End If
    
    strSQL = strSQL & "EX.pkId intExecutivo, "
    strSQL = strSQL & "EX.intNumSeq, "
    strSQL = strSQL & "EX.strExecutadoNome, "
    strSQL = strSQL & "EM.intCodigo intMoeda, "
    strSQL = strSQL & "MO.strAbreviatura, "
    strSQL = strSQL & "EX.dblVltotTotal, "
    strSQL = strSQL & "EX.Strexecutadotplognotif strTipoLog, "
    strSQL = strSQL & "EX.Strexecutadotitlognotif strTitLog, "
    strSQL = strSQL & "EX.Strexecutadonomelognotif strLogradouro, "
    strSQL = strSQL & "EX.Strexecutadonumlognotif strNumero, "
    strSQL = strSQL & "EX.Strexecutadocomplnotif strComplemento, "
    strSQL = strSQL & "EX.Strexecutadobairronotif strBairro, "
    strSQL = strSQL & "EX.Strexecutadocidnotif strMunicipio, "
    strSQL = strSQL & "EX.Strexecutadoufnotif strUF, "
    strSQL = strSQL & "EX.Intexecutadocepnotif intCEP, "
    strSQL = strSQL & "EX.strExecutadoIdentidade, "
    strSQL = strSQL & "EX.strExecutadoCnpjCpf "
    strSQL = strSQL & "FROM "
    If bytDBType = Oracle Then
        strSQL = strSQL & gstrExecutivo & " EX, "
        strSQL = strSQL & gstrMoedas & " MO, "
        strSQL = strSQL & gstrExecutivoMoedas & " EM "
        strSQL = strSQL & "WHERE "
        strSQL = strSQL & "EX.Intmoeda = MO.pkID " & strOUTJOracle & " "
        strSQL = strSQL & "and MO.pkId = EM.Intmoedas " & strOUTJOracle & " "
        strSQL = strSQL & "and EX.intLoteExecutivo = " & dbc_intLote.Text
    Else
        strSQL = strSQL & gstrExecutivo & " EX LEFT JOIN " & gstrMoedas & " MO ON EX.Intmoeda = MO.pkID "
        strSQL = strSQL & "LEFT JOIN " & gstrExecutivoMoedas & " EM ON MO.pkId = EM.Intmoedas "
        strSQL = strSQL & "WHERE EX.intLoteExecutivo = " & dbc_intLote.Text
    End If
    
    If bytDBType = Oracle Then
        strSQL = strSQL & " and ROWNUM <= " & CInt(txt_intQuantidade) & " "
    End If
    strSQL = strSQL & " ORDER BY EX.intNumSeq"
    
    Set gobjBanco = New clsBanco
    
    If Not gobjBanco.CriaADO(strSQL, 10, adoResultado) Then
        ExibeMensagem "Erro ao consultar os registros"
        gobjBanco.ExecutaRollbackTrans
        Exit Sub
    End If
    
    If adoResultado.EOF Then
        ExibeMensagem "Não foram encontrados registros de Executivos Fiscais dentro das especificações estabelecidas."
        gobjBanco.ExecutaRollbackTrans
        Exit Sub
    End If
    
    intProtocolo = txt_intProtocolo - 1
    prg_Status.Min = 0
    prg_Status.Value = 0
    prg_Status.Max = adoResultado.RecordCount
    prg_Status.Visible = True
    
    
    intNumeroArquivo = txt_intArquivoInicial
    strNomeArquivo = txt_strCaminho & "expre" & String(3 - Len(CStr(intNumeroArquivo)), "0") & intNumeroArquivo & ".txt"
    intCont = 0
    Open strNomeArquivo For Output As #1
    
    With adoResultado
    
        While Not .EOF
        
            intCont = intCont + 1
            
            prg_Status.Value = prg_Status.Value + 1
            
            strWord = ""
        
            '1   001 Caractere   001 001 ‘[‘
            strWord = strWord & "["
            
            '2   002 Numérico    002 003 00 – Código do tipo de linha
            strWord = strWord & "00"
            
            '3   001 Caractere   004 004 ‘]‘
            strWord = strWord & "]"
            
            '4   001 Caractere   005 005 ‘[‘
            strWord = strWord & "["
            
            '5   020 Numérico    006 025 Dívida Ativa = num seq da tblexecutivo
            strWord = strWord & PreparaCampo(!intNumSeq, 2, 20)
            
            '6   001 Caractere   026 026 ‘]‘
            strWord = strWord & "]"
            
            '7   001 Caractere   027 027 ‘[‘
            strWord = strWord & "["
            
            '8   006 Numérico    028 033 Nº do protocolo = desenvolver do formulario + 1
            If AtualizaProtocolo(!intExecutivo) Then
                strWord = strWord & PreparaCampo(intProtocolo, 2, 6)
            Else
                ApagaProgress
                gobjBanco.ExecutaRollbackTrans
                Close #1
                Exit Sub
            End If
            '9   001 Caractere   034 034 ‘]‘
            strWord = strWord & "]"
            
            '10  001 Caractere   035 035 ‘[‘
            strWord = strWord & "["
            
            '11  004 Numérico    036 039 Ano do protocolo = do formulario
            strWord = strWord & PreparaCampo(txt_intExercicio, 2, 4)
            
            '12  001 Caractere   040 040 ‘]‘
            strWord = strWord & "]"
            
            '13  001 Caractere   041 041 ‘[‘
            strWord = strWord & "["
            
            '14  003 Numérico    042 044 Quantidade de requeridos = "01"
            strWord = strWord & "001"
            
            '15  001 Caractere   045 045 ‘]‘
            strWord = strWord & "]"
            
            '16  001 Caractere   046 046 ‘[‘
            strWord = strWord & "["
            
            '17  001 Numérico    047 047 Tipo da moeda = de tblexecutivo campo moeda - fkbuscar em tblexecutivomoeda -fk
            If IsNull(!intMoeda) Then
                strWord = strWord & "9"
            Else
                strWord = strWord & PreparaCampo(!intMoeda, 2, 1)
            End If
            
            '18  001 Caractere   048 048 ‘]‘
            strWord = strWord & "]"
            
            '19  001 Caractere   049 049 ‘[‘
            strWord = strWord & "["
            
            '20  015 Numérico    050  06  Valor da causa = tblexecutivo valor total.
            strWord = strWord & PreparaCampo(!dblVlTotTotal, 0, 15)
            
            '21  001 Caractere   065 065 ‘]‘
            strWord = strWord & "]"
            
            '22  001 Caractere   066 066 ‘[‘
            strWord = strWord & "["
            
            '23  066 Caractere   067 132 Prefeitura
            strWord = strWord & PreparaCampo(strPrefeitura, 1, 66)
            
            '24  001 Caractere   133 133 ‘]‘
            strWord = strWord & "]"
            
            '25  001 Caractere   134 134 ‘[‘
            strWord = strWord & "["
            
            '26  007 Caractere   135 141 Número da OAB = tblexecutivosadvogados.
            strWord = strWord & PreparaCampo(strNumeroOAB, 1, 7)
            
            '27  001 Caractere   142 142 ‘]‘
            strWord = strWord & "]"
            
            '28  001 Caractere   143 143 ‘[‘
            strWord = strWord & "["
            
            '29  002 Caractere   144 145 Unidade Federativa da OAB = tblexecutivosadvogados.
            strWord = strWord & PreparaCampo(strUFAdvogado, 1, 2)
            
            '30  001 Caractere   146 146 ‘]‘
            strWord = strWord & "]"
            
            '31  001 Caractere   147 147 ‘[‘
            strWord = strWord & "["
            
            '32  040 Caractere   148 187 Nome do Advogado = tblexecutivos advogados.
            strWord = strWord & PreparaCampo(dbc_intAdvogado.Text, 1, 40)
            
            '33  001 Caractere   188 188 ‘]‘
            strWord = strWord & "]"
            
            '34  001 Caractere   189 189 ‘[‘
            strWord = strWord & "["
            
            '35  003 Caractere   190 192 Filler (espaços em branco)
            strWord = strWord & "   "
            
            '36  001 Caractere   193 193 ‘]‘
            strWord = strWord & "]"
            
            
            
'--------------------------------------------------------------------------------------
            strWord = strWord & vbNewLine
'--------------------------------------------------------------------------------------

            
            '1 001 Caractere 001 001 ‘[‘
            strWord = strWord & "["
            
            '2 002 Numérico 002 003 10 – Código do tipo de linha
            strWord = strWord & "10"
            
            '3 001 Caractere 004 004 ‘]‘
            strWord = strWord & "]"
            
            '4 001 Caractere 005 005 ‘[‘
            strWord = strWord & "["
            
            '5 020 Numérico 006 025 Dívida Ativa = o mesmo que registro "00"
            strWord = strWord & PreparaCampo(!intNumSeq, 2, 20)
            
            '6 001 Caractere 026 026 ‘]‘
            strWord = strWord & "]"
            
            '7 001 Caractere 027 027 ‘[‘
            strWord = strWord & "["
            
            '8 006 Numérico 028 033 Nº do protocolo = o mesmo que registro "00"
            strWord = strWord & PreparaCampo(intProtocolo, 2, 6)
                
            '9 001 Caractere 034 034 ‘]‘
            strWord = strWord & "]"
            
            '10 001 Caractere 035 035 ‘[‘
            strWord = strWord & "["
            
            '11 004 Numérico 036 039 Ano do protocolo = o mesmo que registro "00"
            strWord = strWord & PreparaCampo(txt_intExercicio, 2, 4)
            
            '12 001 Caractere 040 040 ‘]‘
            strWord = strWord & "]"
            
            '13 001 Caractere 041 041 ‘[‘
            strWord = strWord & "["
            
            '14 003 Numérico 042 044 Seqüência da parte = "01"
            strWord = strWord & "001"
            
            '15 001 Caractere 045 045 ‘]‘
            strWord = strWord & "]"
            
            '16 001 Caractere 046 046 ‘[‘
            strWord = strWord & "["
            
            '17 066 Caractere 047 112 Nome da parte = tblexecutivo executadonome.
            strWord = strWord & PreparaCampo(!strExecutadoNome, 1, 66)
            
            '18 001 Caractere 113 113 ‘]‘
            strWord = strWord & "]"
            
            '19 001 Caractere 114 114 ‘[‘
            strWord = strWord & "["
            
            '20 040 Caractere 115 154 Endereço = tblexecutivo executado nomelog,numlog, compl,cep.
            
            varAux = ""
            
            If gstrENulo(!strTipoLog) <> "" Then
                varAux = varAux & gstrENulo(!strTipoLog) & " "
            End If
            
            If gstrENulo(!strTitLog) <> "" Then
                varAux = varAux & gstrENulo(!strTitLog) & " "
            End If
            
            If gstrENulo(!strLogradouro) <> "" Then
                varAux = varAux & gstrENulo(!strLogradouro)
            End If
            
            If gstrENulo(!strNumero) <> "" Then
                varAux = varAux & ", " & gstrENulo(!strNumero)
            End If
            
            If gstrENulo(!STRCOMPLEMENTO) <> "" Then
                varAux = varAux & ", " & gstrENulo(!STRCOMPLEMENTO)
            End If
            
            If gstrENulo(!STRBAIRRO) <> "" Then
                varAux = varAux & ", " & gstrENulo(!STRBAIRRO)
            End If
            
            If gstrENulo(!STRMUNICIPIO) <> "" Then
                varAux = varAux & ", " & gstrENulo(!STRMUNICIPIO)
            End If
            
            If gstrENulo(!STRUF) <> "" Then
                varAux = varAux & " -" & gstrENulo(!STRUF)
            End If
            
            If gstrENulo(!INTCEP) <> "" And gstrENulo(!INTCEP) <> "0" Then
                varAux = varAux & " - " & gstrENulo(!INTCEP)
            End If
            
            strWord = strWord & PreparaCampo(varAux, 1, 40)
            
            '21 001 Caractere 155 155 ‘]‘
            strWord = strWord & "]"
            
            '22 001 Caractere 156 156 ‘[‘
            strWord = strWord & "["
            
            '23 001 Numérico 157 157 Tipo do documento nº 1
                
            varAux = ""
            
            If gstrENulo(!StrExecutadoCnpjCpf) = "" Then
                varAux = "0"
            Else
                If Len(gstrENulo(!StrExecutadoCnpjCpf)) = 11 Then
                    varAux = "2"
                Else
                    varAux = "4"
                End If
            End If
            
            strWord = strWord & PreparaCampo(varAux, 2, 1)
            
            '24 001 Caractere 158 158 ‘]‘
            strWord = strWord & "]"
            
            '25 001 Caractere 159 159 ‘[‘
            strWord = strWord & "["
            
            '26 014 Caractere 160 173 Número do documento nº 1
            strWord = strWord & PreparaCampo(!StrExecutadoCnpjCpf, 2, 14)
            
            '27 001 Caractere 174 174 ‘]‘
            strWord = strWord & "]"
            
            '28 001 Caractere 175 175 ‘[‘
            strWord = strWord & "["
            
            '29 001 Numérico 176 176 Tipo do documento nº 2
            
            varAux = ""
            
            If gstrENulo(!strExecutadoIdentidade) = "" Then
                varAux = "0"
            Else
                varAux = "1"
            End If
            
            strWord = strWord & PreparaCampo(varAux, 2, 1)
            
            '30 001 Caractere 177 177 ‘]‘
            strWord = strWord & "]"
            
            '31 001 Caractere 178 178 ‘[‘
            strWord = strWord & "["
            
            '32 014 Caractere 179 192 Número do documento nº 2
            strWord = strWord & PreparaCampo(!strExecutadoIdentidade, 2, 14)
            
            '33 001 Caractere 193 193 ‘]‘
            strWord = strWord & "]"
            
        
            Print #1, strWord
            
            .MoveNext
            
            If intCont = txt_intQntdePorArquivo And Not adoResultado.EOF Then
                intCont = 0
                Close #1
                intNumeroArquivo = intNumeroArquivo + 1
                strNomeArquivo = txt_strCaminho & "expre" & String(3 - Len(CStr(intNumeroArquivo)), "0") & intNumeroArquivo & ".txt"
                Open strNomeArquivo For Output As #1
            End If
        
        Wend
        
    End With
    
    gobjBanco.ExecutaCommitTrans
    ExibeMensagem "Arquivos gerados com sucesso."
    ApagaProgress
    
    Close #1
    
    Exit Sub

End Sub

Private Sub ApagaProgress()
    prg_Status.Value = 0
    prg_Status.Visible = False
End Sub

Private Function PreparaCampo(Campo, Tipo As Integer, Tamanho As Integer)
Dim varAux

    Select Case Tipo
        Case 0 ' Valor
            
            varAux = gstrConvVrDoSql(Campo)
            varAux = Replace(varAux, ",", "")
            varAux = Replace(varAux, ".", "")
            
            varAux = String(Tamanho - Len(varAux), "0") & varAux
        
        Case 1 ' Texto
            
            varAux = gstrENulo(Campo)
            
            If Len(varAux) > Tamanho Then
                varAux = Left(varAux, Tamanho)
            End If
            
            varAux = varAux & String(Tamanho - Len(varAux), " ")
        
        Case 2 ' Número
            
            varAux = gstrENulo(Campo)
            varAux = String(Tamanho - Len(varAux), "0") & varAux
        
    End Select
        
    PreparaCampo = varAux

End Function


Private Function AtualizaProtocolo(intExecutivo As Long) As Boolean
Dim strSQL As String
Dim adoResultado As ADODB.Recordset

    AtualizaProtocolo = False
    
    intProtocolo = intProtocolo + 1
    
    strSQL = ""
    strSQL = strSQL & "SELECT * FROM " & gstrExecutivo
    strSQL = strSQL & " WHERE intNumeroProtocolo = " & intProtocolo
    strSQL = strSQL & " and intProtocoloAno = " & txt_intExercicio
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSQL, 10, adoResultado) Then
        If Not adoResultado.EOF Then
            ExibeMensagem "Já existe um protocolo " & intProtocolo & " no ano " & txt_intExercicio & " cadastrado. A operação será abortada. "
            Exit Function
        End If
    Else
        ExibeMensagem "Erro ao gravar o número do protocolo."
        Exit Function
    End If
    
    strSQL = ""
    strSQL = strSQL & "UPDATE " & gstrExecutivo & " SET intExecutivoAdvogados = " & dbc_intAdvogado.BoundText
    strSQL = strSQL & " WHERE pkId = " & intExecutivo
    
    Set gobjBanco = New clsBanco
    
    If Not gobjBanco.Execute(strSQL) Then
        ExibeMensagem "Erro ao gravar o Advogado selecionado."
        Exit Function
    End If
    
    strSQL = ""
    strSQL = strSQL & "UPDATE " & gstrExecutivo & " SET intNumeroProtocolo = " & intProtocolo
    strSQL = strSQL & ", intProtocoloAno = " & txt_intExercicio
    strSQL = strSQL & " WHERE pkId = " & intExecutivo
    
    Set gobjBanco = New clsBanco
    
    If gobjBanco.Execute(strSQL) Then
        AtualizaProtocolo = True
    Else
        ExibeMensagem "Ocorreu um erro ao atualizar o Numero do Protocolo."
        Exit Function
    End If

End Function

Private Sub Form_Load()
Dim strSQL As String

    strSQL = ""
    strSQL = strSQL & "Select EA.pkId, strNome "
    strSQL = strSQL & "from Tblexecutivoadvogados EA, Tblcontribuinte CO "
    strSQL = strSQL & "Where EA.intContribuinte = CO.Pkid;strNome "

    dbc_intAdvogado.Tag = strSQL
    
    
    strSQL = ""
    strSQL = strSQL & "SELECT Max(pkId) pkId, intLoteExecutivo FROM "
    strSQL = strSQL & gstrExecutivo
    strSQL = strSQL & " WHERE bitDistribuicaoEletronica = 1 "
    strSQL = strSQL & " AND (bitDistribuido = 0 OR bitDistribuido is null) "
    strSQL = strSQL & " AND intNumeroProtocolo is null group by intLoteExecutivo;intLoteExecutivo "
    
    dbc_intLote.Tag = strSQL
    
    
        
End Sub

Private Sub txt_intArquivoInicial_GotFocus()
    MarcaCampo txt_intArquivoInicial
End Sub

Private Sub txt_intArquivoInicial_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_intArquivoInicial
End Sub

Private Sub txt_intExercicio_GotFocus()
    MarcaCampo txt_intExercicio
End Sub

Private Sub txt_intExercicio_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_intExercicio
End Sub

Private Sub txt_intProtocolo_GotFocus()
    MarcaCampo txt_intProtocolo
End Sub

Private Sub txt_intProtocolo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_intProtocolo
End Sub

Private Sub txt_intQntdePorArquivo_GotFocus()
    MarcaCampo txt_intQntdePorArquivo
End Sub

Private Sub txt_intQntdePorArquivo_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_intQntdePorArquivo
End Sub

Private Sub txt_intQuantidade_GotFocus()
    MarcaCampo txt_intQuantidade
End Sub

Private Sub txt_intQuantidade_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_intQuantidade
End Sub

Private Sub txt_strCaminho_GotFocus()
    MarcaCampo txt_strCaminho
End Sub
