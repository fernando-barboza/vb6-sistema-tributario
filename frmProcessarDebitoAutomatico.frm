VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmProcessarDebitoAutomatico 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importação de arquivos"
   ClientHeight    =   1710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   5520
   Begin VB.CommandButton cmdProcessar 
      Caption         =   "Processar"
      Height          =   375
      Left            =   90
      TabIndex        =   5
      Top             =   1230
      Width           =   1035
   End
   Begin VB.Frame fraSelecione 
      Caption         =   "Localizar"
      Height          =   1005
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   5355
      Begin VB.CommandButton cmdArquivo 
         Caption         =   "..."
         Height          =   375
         Left            =   4860
         TabIndex        =   3
         Top             =   390
         Width           =   435
      End
      Begin VB.TextBox txtArquivo 
         Height          =   360
         Left            =   840
         TabIndex        =   2
         Top             =   397
         Width           =   3975
      End
      Begin VB.Label lblArquivo 
         AutoSize        =   -1  'True
         Caption         =   "Arquivo"
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   540
      End
   End
   Begin MSComDlg.CommonDialog dlgArquivo 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblProgresso 
      AutoSize        =   -1  'True
      Caption         =   "Aguarde..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4515
      TabIndex        =   4
      Top             =   1320
      Visible         =   0   'False
      Width           =   900
   End
End
Attribute VB_Name = "frmProcessarDebitoAutomatico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private objText  As TextStream
Private objFiles As FileSystemObject

Private Function ConverteDataDoArquivo(strData As String, blnLayoutNovo As Boolean) As Date
Dim strAux As String
    If blnLayoutNovo Then
        strAux = Right(strData, 2) & "/" & Mid(strData, 5, 2) & "/" & Left(strData, 4)
    Else
        strAux = Right(strData, 2) & "/" & Mid(strData, 3, 2) & "/" & Left(strData, 2)
    End If
    
    ConverteDataDoArquivo = CDate(strAux)
    
End Function

Private Function ConverteDataDoBanco(dtmData As Date) As String
Dim strAux As String
        
    strAux = Right(dtmData, 4) & Mid(dtmData, 4, 2) & Left(dtmData, 2)
    
    ConverteDataDoBanco = strAux
    
End Function

Private Sub cmdArquivo_Click()
    
    dlgArquivo.CancelError = True
    dlgArquivo.DialogTitle = "Selecione o arquivo"
'    dlgArquivo.InitDir = strUltimoCaminho
    dlgArquivo.Filter = "Todos arquivos (*.*)|*.*"
    dlgArquivo.flags = cdlOFNOverwritePrompt + cdlOFNHideReadOnly
    
    On Error GoTo err_cmd_Arquivo_Click
    
    dlgArquivo.ShowOpen
    txtArquivo = dlgArquivo.Filename
'    strUltimoCaminho = Replace(dlgArquivo.Filename, dlgArquivo.FileTitle, "")
    Exit Sub

err_cmd_Arquivo_Click:
    If Err.Number = 32755 Then
        txtArquivo = ""
    End If

End Sub

Private Sub cmdProcessar_Click()
Dim strSQL                  As String
Dim strLinha                As String
Dim intBanco                As Integer
Dim strAgencia              As String
Dim strInscricao            As String
Dim adoResultado            As ADODB.Recordset
Dim strDataOpcao            As String
Dim strCodigoBarras         As String
Dim strIdentificacaoBanco   As String
Dim strIdentificacaoDebAut  As String
Dim intComposicaoDaReceita  As Long

On Error GoTo Problema_Na_Rotina

    lblProgresso.Visible = True: Me.MousePointer = vbHourglass
    
    DoEvents
    
    If Trim$(txtArquivo.Text) = Space$(0) Then
        ExibeMensagem "Indique a localização do arquivo."
        Exit Sub
    ElseIf Dir$(txtArquivo.Text) = Space$(0) Then
        ExibeMensagem "Arquivo não encontrado no local especificado."
        Exit Sub
    End If
    
    Set objFiles = New FileSystemObject
    
    Set objText = objFiles.OpenTextFile(txtArquivo.Text)
    
    Do While Not objText.AtEndOfLine
        
        strLinha = objText.ReadLine
        
        If Left$(strLinha, 1) = "A" Then
            intBanco = Trim$(Mid$(strLinha, 43, 3))
        Else
        
            If Left$(strLinha, 1) = "B" Then
            
                strInscricao = String$(20 - Len(Left$(Trim$(Mid$(strLinha, 2, 25)), Len(Trim$(Mid$(strLinha, 2, 25))) - 5)), "0") & Left$(Trim$(Mid$(strLinha, 2, 25)), Len(Trim$(Mid$(strLinha, 2, 25))) - 5)
                strIdentificacaoDebAut = Trim$(Mid$(strLinha, 2, 25))
                
                If gobjBanco.CriaADO("select intcomposicaodareceita from " & gstrLancamentoAlfa & " where intexercicio = 2005 and strinscricao = '" & strInscricao & "'", 10, adoResultado) Then
                
                    If Not adoResultado.EOF Then
                        
                        intComposicaoDaReceita = adoResultado("intcomposicaodareceita")
                    
                        strAgencia = Trim$(Mid$(strLinha, 27, 4))
                        strIdentificacaoBanco = Trim$(Mid$(strLinha, 31, 14))
                        strDataOpcao = gstrConvDtParaSql(ConverteDataDoArquivo(Trim$(Mid$(strLinha, 45, 8)), True))
                        strCodigoBarras = Trim$(Mid$(strLinha, 2, 25)) & Trim$(Mid$(strLinha, 27, 4)) '& Trim$(Mid$(strLinha, 45, 8))
                        
                        adoResultado.Close
                        
                        If gobjBanco.CriaADO("select count(*) nreg from " & gstrDebitoAutomatico & " where strIdentificacaoDebAut = '" & strIdentificacaoDebAut & "'", 10, adoResultado) Then
                        
                            If Not adoResultado.EOF Then
                            
                                If adoResultado("nreg") = 0 Then
                                    gobjBanco.Execute "insert into " & gstrDebitoAutomatico & " (strInscricaoCadastral, intComposicaoDaReceita, strIdentificacaoDebAut, strAgencia, strIdentificacaoBanco, dtmDtOpcao, intBanco, dtmDtAtualizacao, lngCodUsr) values" & _
                                                                                              " ('" & strInscricao & "'," & intComposicaoDaReceita & ",'" & strIdentificacaoDebAut & "','" & strAgencia & "','" & strIdentificacaoBanco & "'," & strDataOpcao & "," & intBanco & "," & strGETDATE & "," & glngCodUsr & ")"
                                    
                                    adoResultado.Close
                                    
                                    strSQL = "select " & _
                                                "lv.pkid, " & _
                                                "lv.dtmdtvencimento, " & _
                                                "lv.dblValor " & _
                                             "from " & _
                                                gstrDebitoAutomatico & " da, " & _
                                                gstrLancamentoAlfa & " la, " & _
                                                gstrLancamentoValor & " lv " & _
                                             "Where " & _
                                                "la.strinscricao = da.strinscricaocadastral and " & _
                                                "lv.intlancamentoalfa = la.pkid and " & _
                                                "da.strinscricaocadastral = '" & strInscricao & "' and " & _
                                                "la.intExercicio = 2005 and lv.intparcela in(9,10,11,12)"
                                                
                                    If gobjBanco.CriaADO(strSQL, 10, adoResultado) Then
                                        
                                        gobjBanco.ExecutaBeginTrans
                                        
                                        On Error GoTo Problema_Na_Rotina_Insert
                                        
                                        Do While Not adoResultado.EOF
                                            
                                            gobjBanco.Execute "insert into " & gstrGuias & " (intnumero, dtmdtemissao, dtmdtvencimento, dblvalor, strcodbarra, dtmdtatualizacao, lngcodusr) values" & _
                                                                                           " (" & glngRetornaProximoNumeroGuia & "," & strGETDATE & "," & gstrConvDtParaSql(adoResultado("dtmdtvencimento")) & "," & gstrConvVrParaSql(adoResultado("dblValor")) & ",'" & strCodigoBarras & ConverteDataDoBanco(adoResultado("dtmdtvencimento")) & "'," & strGETDATE & "," & glngCodUsr & ")"
                                        
                                            gobjBanco.Execute "insert into " & gstrLancamentoGuias & " (intlancamentovalor, intguias, dblvalorprincipal, dtmdtatualizacao, lngcodusr) values" & _
                                                                                                     " (" & adoResultado("pkid") & "," & glngRetornaPkidTabelaPai("seqtblGuias", "tblGuias") & "," & gstrConvVrParaSql(adoResultado("dblValor")) & "," & strGETDATE & "," & glngCodUsr & ")"
                                        
                                            adoResultado.MoveNext
                                            
                                        Loop
                                        
                                        gobjBanco.ExecutaCommitTrans
                                        
                                    End If
                                    
                                End If
                                
                                adoResultado.Close
                                
                            End If
                            
                        End If
                    
                    End If
                    
                End If
                
            Else
            
                If IsNumeric(Mid$(strLinha, 1, 1)) Then
                
                    strInscricao = String$(9, "0") & Trim$(Left$(strLinha, 11))
                    strIdentificacaoDebAut = Trim$(Left$(strLinha, 16))
                    
                    If gobjBanco.CriaADO("select intcomposicaodareceita from " & gstrLancamentoAlfa & " where intexercicio = 2005 and strinscricao = '" & strInscricao & "'", 10, adoResultado) Then
                    
                        If Not adoResultado.EOF Then
                            
                            intComposicaoDaReceita = adoResultado("intcomposicaodareceita")
                        
                            strAgencia = Trim$(Mid$(strLinha, 25, 4))
                            strIdentificacaoBanco = Trim$(Mid$(strLinha, 29, 8))
                            strDataOpcao = gstrConvDtParaSql(Replace(Trim$(Mid$(strLinha, 53, 10)), ".", "/"))
                            strCodigoBarras = Trim$(Mid$(strLinha, 1, 16)) & Trim$(Mid$(strLinha, 25, 4)) '& Trim$(Mid$(strLinha, 59, 4)) & Trim$(Mid$(strLinha, 56, 2)) & Trim$(Mid$(strLinha, 53, 2))
                            
                            adoResultado.Close
                            
                            If gobjBanco.CriaADO("select count(*) nreg from " & gstrDebitoAutomatico & " where strIdentificacaoDebAut = '" & strIdentificacaoDebAut & "'", 10, adoResultado) Then
                            
                                If Not adoResultado.EOF Then
                                
                                    If adoResultado("nreg") = 0 Then
                                        gobjBanco.Execute "insert into " & gstrDebitoAutomatico & " (strInscricaoCadastral, intComposicaoDaReceita, strIdentificacaoDebAut, strAgencia, strIdentificacaoBanco, dtmDtOpcao, intBanco, dtmDtAtualizacao, lngCodUsr) values" & _
                                                                                                  " ('" & strInscricao & "'," & intComposicaoDaReceita & ",'" & strIdentificacaoDebAut & "','" & strAgencia & "','" & strIdentificacaoBanco & "'," & strDataOpcao & ", 341," & strGETDATE & "," & glngCodUsr & ")"
                                        
                                        adoResultado.Close
                                        
                                        strSQL = "select " & _
                                                    "lv.pkid, " & _
                                                    "lv.dtmdtvencimento, " & _
                                                    "lv.dblValor " & _
                                                 "from " & _
                                                    gstrDebitoAutomatico & " da, " & _
                                                    gstrLancamentoAlfa & " la, " & _
                                                    gstrLancamentoValor & " lv " & _
                                                 "Where " & _
                                                    "la.strinscricao = da.strinscricaocadastral and " & _
                                                    "lv.intlancamentoalfa = la.pkid and " & _
                                                    "da.strinscricaocadastral = '" & strInscricao & "' and " & _
                                                    "la.intExercicio = 2005 and lv.intparcela in(9,10,11,12)"
                                                    
                                        If gobjBanco.CriaADO(strSQL, 10, adoResultado) Then
                                            
                                            gobjBanco.ExecutaBeginTrans
                                            
                                            On Error GoTo Problema_Na_Rotina_Insert
                                            
                                            Do While Not adoResultado.EOF
                                                
                                                gobjBanco.Execute "insert into " & gstrGuias & " (intnumero, dtmdtemissao, dtmdtvencimento, dblvalor, strcodbarra, dtmdtatualizacao, lngcodusr) values" & _
                                                                                               " (" & glngRetornaProximoNumeroGuia & "," & strGETDATE & "," & gstrConvDtParaSql(adoResultado("dtmdtvencimento")) & "," & gstrConvVrParaSql(adoResultado("dblValor")) & ",'" & strCodigoBarras & ConverteDataDoBanco(adoResultado("dtmdtvencimento")) & "'," & strGETDATE & "," & glngCodUsr & ")"
                                            
                                                gobjBanco.Execute "insert into " & gstrLancamentoGuias & " (intlancamentovalor, intguias, dblvalorprincipal, dtmdtatualizacao, lngcodusr) values" & _
                                                                                                         " (" & adoResultado("pkid") & "," & glngRetornaPkidTabelaPai("seqtblGuias", "tblGuias") & "," & gstrConvVrParaSql(adoResultado("dblValor")) & "," & strGETDATE & "," & glngCodUsr & ")"
                                            
                                                adoResultado.MoveNext
                                                
                                            Loop
                                            
                                            gobjBanco.ExecutaCommitTrans
                                            
                                        End If
                                        
                                    End If
                                    
                                    adoResultado.Close
                                    
                                End If
                                
                            End If
                        
                        End If
                    
                    End If
                    
                End If
                        
            End If
        
        End If
        
    Loop
    
    objText.Close
    
    Set objText = Nothing
    
    Set objFiles = Nothing
    
'    ExibeMensagem "Importação efetuada com sucesso."
    
    lblProgresso.Visible = False: Me.MousePointer = vbNormal
    
    Exit Sub
    
Problema_Na_Rotina:

    ExibeMensagem "Erro nº " & Err.Number & " - " & Err.Description
    
    lblProgresso.Visible = False: Me.MousePointer = vbNormal

    Exit Sub
    
Problema_Na_Rotina_Insert:
    
    gobjBanco.ExecutaRollbackTrans
    
    lblProgresso.Visible = False: Me.MousePointer = vbNormal
    
End Sub

