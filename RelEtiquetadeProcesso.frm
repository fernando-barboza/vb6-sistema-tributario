VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmRelEtiquetadeProcesso 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Etiquetas de Processo"
   ClientHeight    =   3045
   ClientLeft      =   1905
   ClientTop       =   2115
   ClientWidth     =   6195
   Icon            =   "RelEtiquetadeProcesso.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   6195
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   2790
      Left            =   150
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   90
      Width           =   5985
      _ExtentX        =   10557
      _ExtentY        =   4921
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Comprovante"
      TabPicture(0)   =   "RelEtiquetadeProcesso.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "chkImprimeTitulo"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fra_FaixaNum"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fra_FaixaData"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fraVolume"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "chkTodosVolumes"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      Begin VB.CheckBox chkTodosVolumes 
         Caption         =   "Todos os Volumes"
         Height          =   315
         Left            =   2400
         TabIndex        =   6
         Top             =   2400
         Value           =   1  'Checked
         Width           =   1605
      End
      Begin VB.Frame fraVolume 
         Height          =   645
         Left            =   180
         TabIndex        =   16
         Top             =   1950
         Width           =   1695
         Begin VB.TextBox txtVolume 
            Height          =   315
            Left            =   1080
            TabIndex        =   5
            Top             =   210
            Width           =   375
         End
         Begin VB.Label lblVolume 
            AutoSize        =   -1  'True
            Caption         =   "Volume"
            Height          =   195
            Left            =   300
            TabIndex        =   17
            Top             =   240
            Width           =   525
         End
      End
      Begin VB.Frame fra_FaixaData 
         Height          =   1455
         Left            =   3030
         TabIndex        =   13
         Top             =   510
         Width           =   2625
         Begin VB.TextBox txt_dtmDataInicial 
            Height          =   285
            Left            =   1230
            MaxLength       =   10
            TabIndex        =   3
            Top             =   210
            Width           =   1005
         End
         Begin VB.TextBox txt_dtmDataFinal 
            Height          =   285
            Left            =   1200
            MaxLength       =   10
            TabIndex        =   4
            Top             =   990
            Width           =   1005
         End
         Begin VB.Label lbl_dtmDataFinal 
            AutoSize        =   -1  'True
            Caption         =   "Final"
            Height          =   195
            Left            =   435
            TabIndex        =   15
            Top             =   1020
            Width           =   330
         End
         Begin VB.Label lbl_dtmDataInicial 
            AutoSize        =   -1  'True
            Caption         =   "Inicial"
            Height          =   195
            Left            =   420
            TabIndex        =   14
            Top             =   240
            Width           =   405
         End
      End
      Begin VB.Frame fra_FaixaNum 
         Height          =   1425
         Left            =   180
         TabIndex        =   9
         Top             =   510
         Width           =   2445
         Begin VB.TextBox txtintExercicio 
            Height          =   285
            Left            =   1050
            MaxLength       =   4
            TabIndex        =   2
            Top             =   990
            Width           =   1005
         End
         Begin VB.TextBox txtintProtocolizacaoFinal 
            Height          =   285
            Left            =   1050
            MaxLength       =   10
            TabIndex        =   1
            Top             =   600
            Width           =   1005
         End
         Begin VB.TextBox txtintProtocolizacaoInicial 
            Height          =   285
            Left            =   1050
            MaxLength       =   10
            TabIndex        =   0
            Top             =   240
            Width           =   1005
         End
         Begin VB.Label lbl_intExercicio 
            AutoSize        =   -1  'True
            Caption         =   "Exercício"
            Height          =   195
            Left            =   300
            TabIndex        =   12
            Top             =   1020
            Width           =   675
         End
         Begin VB.Label lbl_intRegistroAtendInicial 
            AutoSize        =   -1  'True
            Caption         =   "Inicial"
            Height          =   195
            Left            =   270
            TabIndex        =   11
            Top             =   240
            Width           =   405
         End
         Begin VB.Label lbl_intRegistroAtendFinal 
            AutoSize        =   -1  'True
            Caption         =   "Final"
            Height          =   195
            Left            =   285
            TabIndex        =   10
            Top             =   630
            Width           =   330
         End
      End
      Begin VB.CheckBox chkImprimeTitulo 
         Caption         =   "Imprimir Títulos"
         Height          =   285
         Left            =   4290
         TabIndex        =   7
         Top             =   2400
         Width           =   1425
      End
   End
End
Attribute VB_Name = "frmRelEtiquetadeProcesso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mstrOpcao           As String
Dim mblnClickOk         As Boolean
Dim intCodSeguranca     As Integer
Dim blnPorData          As Boolean
Public blnImprimeTitulo As Boolean
Public blnCadastroPP    As Boolean


Private Sub chkTodosVolumes_Click()
If chkTodosVolumes.Value = 1 Then
    TrocaCorObjeto txtVolume, True
    txtVolume.Text = ""
Else
    TrocaCorObjeto txtVolume, False
    'txtVolume.SetFocus
End If

End Sub

Private Sub Form_Activate()
    gintCodSeguranca = intCodSeguranca
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrDeletar, gstrAplicar
    HabilitaDesabilitaBotao1 True, gstrBtnArquivo, gstrMarcaTudo, gstrDesmarcaTudo
End Sub

Private Sub Form_Deactivate()
    HabilitaDesabilitaBotao1 False, gstrBtnArquivo, gstrMarcaTudo, gstrDesmarcaTudo
End Sub

Private Sub Form_Load()
    blnCadastroPP = False
    TrocaCorObjeto txtVolume, True
    ConfiguraAmbiente False
    txt_dtmDataFinal.Text = Format(gstrDataDoSistema, "dd/mm/") & gintExercicio
    txt_dtmDataInicial = "02/01/" & gintExercicio
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form_Deactivate
End Sub


Public Sub MantemForm(ByVal strModoOperacao As String)
    Select Case UCase(strModoOperacao)
    Case UCase(gstrImprimir)
        VerificaOpcao
    Case UCase(gstrFechar)
        Unload Me
    End Select
End Sub

Private Sub VerificaOpcao()
    If blnDadosOk Then
        ImprimeComprovante
        If blnCadastroPP Then Unload Me
    End If
End Sub

Private Sub ImprimeComprovante()

'******************************************************************************************
' Data: 12/03/2003
' Alteração: - Alteração da string de execução da Stored Procedure.
' Responsável: Everton Bianchini
'******************************************************************************************
    
    Dim strSQL  As String
    strSQL = ""
    If blnPorData Then
    
        strSQL = gstrStoredProcedure("sp_EtiquetaProcessoMod2", _
             "' ',' ','0'," & gstrConvDtParaSql(txt_dtmDataInicial) & ", " & _
             gstrConvDtParaSql(txt_dtmDataFinal) & ", " & Val(txtVolume), True)
    
    Else
        strSQL = gstrStoredProcedure("sp_EtiquetaProcessoMod2", _
             "'" & txtintProtocolizacaoInicial & "'" & ", " & "'" & txtintProtocolizacaoFinal & "'" & ", " & "'" & txtintExercicio & "'" & ", " & gstrConvDtParaSql(txt_dtmDataInicial) & ", " & _
             gstrConvDtParaSql(txt_dtmDataFinal) & ", " & Val(txtVolume), True)
             
    End If
    
    blnImprimeTitulo = chkImprimeTitulo.Value
    
    ImprimeRelatorio rptRelEtiquetasProcessoMod2, strSQL, "Etiquetas de Processo"
    
End Sub
Private Function blnDadosOk() As Boolean
    
    blnDadosOk = False
    
    If blnPorData Then
    
        If gblnDataValida(txt_dtmDataInicial) = False Then
            ExibeMensagem "Data inicial deve ser informada corretamente."
            txt_dtmDataInicial.SetFocus
            Exit Function
        End If
        If gblnDataValida(txt_dtmDataFinal) = False Then
            ExibeMensagem "Data final deve ser informada corretamente."
            txt_dtmDataFinal.SetFocus
            Exit Function
        End If
        If CVDate(txt_dtmDataInicial) > CVDate(txt_dtmDataFinal) Then
            ExibeMensagem "Data inicial tem que ser anterior à data final."
            txt_dtmDataInicial.SetFocus
            Exit Function
        End If
    Else
        If txtintProtocolizacaoInicial.Text = "" Then
            ExibeMensagem "O código inicial deve ser informado."
            txtintProtocolizacaoInicial.SetFocus
            Exit Function
        End If
        
        If txtintProtocolizacaoFinal.Text = "" Then
            ExibeMensagem "O código final deve ser informado."
            txtintProtocolizacaoFinal.SetFocus
            Exit Function
        End If
        
        If txtintExercicio.Text = "" Then
            ExibeMensagem "O Exercício deve ser informado."
            txtintExercicio.SetFocus
            Exit Function
        End If
        
        If Int(txtintProtocolizacaoInicial) > Int(txtintProtocolizacaoFinal) Then
            ExibeMensagem "O Código inicial deve ser maior do que o código final."
            txtintProtocolizacaoInicial.SetFocus
            Exit Function
        End If
    End If
    If chkTodosVolumes.Value = 0 And txtVolume.Text = "" Or txtVolume.Text = "0" Then
        ExibeMensagem "O Volume deve ser informado."
        txtVolume.SetFocus
        Exit Function
    End If
    
    
    
    blnDadosOk = True
End Function

Private Sub txt_dtmDataFinal_GotFocus()
    MarcaCampo txt_dtmDataFinal
    ConfiguraAmbiente True
End Sub

Private Sub txt_dtmDataFinal_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txt_dtmDataFinal
End Sub

Private Sub txt_dtmDataFinal_LostFocus()
    txt_dtmDataFinal = gstrDataFormatada(txt_dtmDataFinal)
End Sub

Private Sub txt_dtmDataInicial_GotFocus()
    MarcaCampo txt_dtmDataInicial
    ConfiguraAmbiente True
End Sub

Private Sub txt_dtmDataInicial_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "D", txt_dtmDataInicial
End Sub

Private Sub txt_dtmDataInicial_LostFocus()
    txt_dtmDataInicial = gstrDataFormatada(txt_dtmDataInicial)
End Sub

Private Sub txtintExercicio_GotFocus()
    txtintExercicio = gintExercicio
    MarcaCampo txtintExercicio
    ConfiguraAmbiente False
End Sub

Private Sub txtintExercicio_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N"
End Sub

Private Sub txtintProtocolizacaoFinal_GotFocus()
    MarcaCampo txtintProtocolizacaoFinal
    ConfiguraAmbiente False
End Sub

Private Sub txtintProtocolizacaoFinal_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii
    CaracterValido KeyAscii, "N", txtintProtocolizacaoFinal
End Sub

Private Sub txtintProtocolizacaoInicial_GotFocus()
    MarcaCampo txtintProtocolizacaoInicial
    ConfiguraAmbiente False
End Sub

Private Sub txtintProtocolizacaoInicial_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintProtocolizacaoInicial
End Sub
Private Sub ConfiguraAmbiente(blpFaixaData As Boolean)
      
   blnPorData = blpFaixaData
   
   If blpFaixaData Then
      
      lbl_intRegistroAtendInicial.ForeColor = &H8000000C
      txtintProtocolizacaoInicial.ForeColor = &H8000000C
      lbl_intRegistroAtendFinal.ForeColor = &H8000000C
      txtintProtocolizacaoFinal.ForeColor = &H8000000C
      lbl_intExercicio.ForeColor = &H8000000C
      txtintExercicio.ForeColor = &H8000000C
      
      txtintProtocolizacaoInicial.BackColor = gvntFundoObjInacessivel
      txtintProtocolizacaoFinal.BackColor = gvntFundoObjInacessivel
      txtintExercicio.BackColor = gvntFundoObjInacessivel
      
      lbl_dtmDataInicial.ForeColor = vbBlack: txt_dtmDataInicial.ForeColor = vbBlack
      lbl_dtmDataFinal.ForeColor = vbBlack: txt_dtmDataFinal.ForeColor = vbBlack
      
      txt_dtmDataInicial.BackColor = &H80000005
      txt_dtmDataFinal.BackColor = &H80000005
      
      
      
   Else
      
      lbl_dtmDataInicial.ForeColor = &H8000000C
      txt_dtmDataInicial.ForeColor = &H8000000C
      lbl_dtmDataFinal.ForeColor = &H8000000C
      txt_dtmDataFinal.ForeColor = &H8000000C
      
      txt_dtmDataInicial.BackColor = gvntFundoObjInacessivel
      txt_dtmDataFinal.BackColor = gvntFundoObjInacessivel
      
      
      
      lbl_intRegistroAtendInicial.ForeColor = vbBlack: txtintProtocolizacaoInicial.ForeColor = vbBlack
      lbl_intRegistroAtendFinal.ForeColor = vbBlack: txtintProtocolizacaoFinal.ForeColor = vbBlack
      lbl_intExercicio.ForeColor = vbBlack: txtintExercicio.ForeColor = vbBlack
      
      txtintProtocolizacaoInicial.BackColor = &H80000005
      txtintProtocolizacaoFinal.BackColor = &H80000005
      txtintExercicio.BackColor = &H80000005
        
   
   End If
   
End Sub


Private Sub txtVolume_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtVolume
End Sub
