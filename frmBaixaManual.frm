VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmBaixaManual 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Baixa Manual"
   ClientHeight    =   3960
   ClientLeft      =   3060
   ClientTop       =   3990
   ClientWidth     =   8460
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   8460
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   3855
      Left            =   60
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   60
      Width           =   8340
      _ExtentX        =   14711
      _ExtentY        =   6800
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Movimento Bancário"
      TabPicture(0)   =   "frmBaixaManual.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblTributo"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblDataDoPagamento"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblAviso"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblcodigoBaixa"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblExercicio"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblContaBancaria"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "dbc_strComposicaoDaReceita"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "dbc_strDescricaoConta"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "dbc_intContaBancaria"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "dbcintLancamentoValor"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "dbcintcodigobaixa"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "dbc_intComposicaoDaReceita"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cmd_Composicao"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txt_dtmDtBaixa"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtPKId"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtintDigito"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "fra_Valores"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txt_intExercicio"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "dbc_intNumeroAviso"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtstrObservacao"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "cmd_ContaCorrente"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "chk_BaixaTotal"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).ControlCount=   23
      TabCaption(1)   =   "Baixa Automática"
      TabPicture(1)   =   "frmBaixaManual.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lbl_Status"
      Tab(1).Control(1)=   "dlgArquivo"
      Tab(1).Control(2)=   "pgr_Status"
      Tab(1).Control(3)=   "fra_Arquivo"
      Tab(1).Control(4)=   "cmdBaixar"
      Tab(1).Control(5)=   "optDebitoAutomatico"
      Tab(1).Control(6)=   "optGrem"
      Tab(1).Control(7)=   "cmd_Importar"
      Tab(1).ControlCount=   8
      Begin VB.CommandButton cmd_Importar 
         Caption         =   "Importar Acordos"
         Height          =   405
         Left            =   -68070
         TabIndex        =   45
         Top             =   3390
         Width           =   1335
      End
      Begin VB.OptionButton optGrem 
         Caption         =   "GREM"
         Height          =   345
         Left            =   -70320
         TabIndex        =   44
         Top             =   2130
         Width           =   1665
      End
      Begin VB.OptionButton optDebitoAutomatico 
         Caption         =   "Débito Automático"
         Height          =   345
         Left            =   -72180
         TabIndex        =   43
         Top             =   2130
         Value           =   -1  'True
         Width           =   1665
      End
      Begin VB.CommandButton cmdBaixar 
         Caption         =   "Baixar"
         Height          =   435
         Left            =   -74700
         TabIndex        =   42
         Top             =   2100
         Width           =   2145
      End
      Begin VB.Frame fra_Arquivo 
         Caption         =   " Arquivo de leitura "
         Height          =   840
         Left            =   -74760
         TabIndex        =   37
         Top             =   540
         Width           =   6000
         Begin VB.CommandButton cmd_Arquivo 
            Caption         =   "..."
            Height          =   300
            Left            =   5460
            Picture         =   "frmBaixaManual.frx":0038
            TabIndex        =   39
            TabStop         =   0   'False
            ToolTipText     =   "Localiza Arquivo de Baixa Automática"
            Top             =   315
            Width           =   345
         End
         Begin VB.TextBox txt_Arquivo 
            Height          =   285
            Left            =   1035
            TabIndex        =   38
            Top             =   315
            Width           =   4410
         End
         Begin VB.Label lbl_Arquivo 
            AutoSize        =   -1  'True
            Caption         =   "Localização"
            Height          =   195
            Left            =   120
            TabIndex        =   40
            Top             =   405
            Width           =   855
         End
      End
      Begin VB.CheckBox chk_BaixaTotal 
         Caption         =   "Todas"
         Height          =   195
         Left            =   4260
         TabIndex        =   13
         Top             =   1830
         Width           =   765
      End
      Begin VB.CommandButton cmd_ContaCorrente 
         Height          =   315
         Left            =   7890
         Picture         =   "frmBaixaManual.frx":0156
         Style           =   1  'Graphical
         TabIndex        =   33
         TabStop         =   0   'False
         Tag             =   "585"
         ToolTipText     =   "Ativa Cadastro de Conta Bancária"
         Top             =   885
         Width           =   360
      End
      Begin VB.TextBox txtstrObservacao 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   1
         EndProperty
         Height          =   285
         Left            =   1170
         TabIndex        =   18
         Top             =   2115
         Width           =   7065
      End
      Begin MSDataListLib.DataCombo dbc_intNumeroAviso 
         Height          =   315
         HelpContextID   =   1
         Left            =   2100
         TabIndex        =   11
         Top             =   1725
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.TextBox txt_intExercicio 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   1170
         MaxLength       =   4
         TabIndex        =   9
         Top             =   1725
         Width           =   465
      End
      Begin VB.Frame fra_Valores 
         Caption         =   "Valores"
         Height          =   1335
         Left            =   90
         TabIndex        =   19
         Top             =   2415
         Width           =   7635
         Begin VB.TextBox txtdblCorreto 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   750
            TabIndex        =   29
            Top             =   855
            Width           =   1215
         End
         Begin VB.TextBox txt_dblTotal 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "#.##0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
            Height          =   285
            Left            =   6285
            TabIndex        =   31
            Top             =   840
            Width           =   1215
         End
         Begin VB.TextBox txtdblCorrecao 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   6300
            TabIndex        =   27
            Top             =   300
            Width           =   1215
         End
         Begin VB.TextBox txtdblJuros 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   4290
            TabIndex        =   25
            Top             =   300
            Width           =   1215
         End
         Begin VB.TextBox txtdblMulta 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   2505
            TabIndex        =   23
            Top             =   300
            Width           =   1215
         End
         Begin VB.TextBox txtdblPrincipal 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0,00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   1
            EndProperty
            Height          =   285
            Left            =   750
            TabIndex        =   21
            Top             =   300
            Width           =   1215
         End
         Begin VB.Label lbldblCorreto 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Correto"
            Height          =   195
            Left            =   180
            TabIndex        =   28
            Top             =   915
            Width           =   510
         End
         Begin VB.Label lblTotal 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Total"
            Height          =   195
            Left            =   5805
            TabIndex        =   30
            Top             =   945
            Width           =   360
         End
         Begin VB.Label lbldblCorrecao 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Correção"
            Height          =   195
            Left            =   5610
            TabIndex        =   26
            Top             =   360
            Width           =   645
         End
         Begin VB.Label lbldblJuros 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Juros"
            Height          =   195
            Left            =   3840
            TabIndex        =   24
            Top             =   360
            Width           =   375
         End
         Begin VB.Label lbldblMulta 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Multa"
            Height          =   195
            Left            =   2055
            TabIndex        =   22
            Top             =   360
            Width           =   390
         End
         Begin VB.Label lbldblPrincipal 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Principal"
            Height          =   195
            Left            =   90
            TabIndex        =   20
            Top             =   360
            Width           =   600
         End
      End
      Begin VB.TextBox txtintDigito 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   5010
         MaxLength       =   1
         TabIndex        =   14
         Top             =   1725
         Width           =   270
      End
      Begin VB.TextBox txtPKId 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   3945
         Locked          =   -1  'True
         OLEDragMode     =   1  'Automatic
         OLEDropMode     =   1  'Manual
         TabIndex        =   32
         Top             =   -285
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txt_dtmDtBaixa 
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
         Left            =   1170
         MaxLength       =   10
         TabIndex        =   2
         Top             =   495
         Width           =   1125
      End
      Begin VB.CommandButton cmd_Composicao 
         Height          =   300
         Left            =   7890
         Picture         =   "frmBaixaManual.frx":0274
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         Tag             =   "590"
         ToolTipText     =   "Ativa Cadastro de Composição da Receita"
         Top             =   1305
         Width           =   360
      End
      Begin MSDataListLib.DataCombo dbc_intComposicaoDaReceita 
         Height          =   315
         HelpContextID   =   1
         Left            =   1830
         TabIndex        =   6
         Top             =   1305
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbcintcodigobaixa 
         Height          =   315
         HelpContextID   =   1
         Left            =   6540
         TabIndex        =   16
         Top             =   1725
         Width           =   1710
         _ExtentX        =   3016
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbcintLancamentoValor 
         Height          =   315
         HelpContextID   =   1
         Left            =   3540
         TabIndex        =   12
         Top             =   1725
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbc_intContaBancaria 
         Height          =   315
         HelpContextID   =   1
         Left            =   1170
         TabIndex        =   4
         Top             =   885
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbc_strDescricaoConta 
         Height          =   315
         Left            =   3630
         TabIndex        =   34
         Top             =   885
         Width           =   4230
         _ExtentX        =   7461
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dbc_strComposicaoDaReceita 
         Height          =   315
         HelpContextID   =   1
         Left            =   2805
         TabIndex        =   35
         Top             =   1305
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSComctlLib.ProgressBar pgr_Status 
         Height          =   165
         Left            =   -74760
         TabIndex        =   36
         Top             =   1440
         Visible         =   0   'False
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   291
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin MSComDlg.CommonDialog dlgArquivo 
         Left            =   -75000
         Top             =   2640
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label lbl_Status 
         Alignment       =   2  'Center
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   -74730
         TabIndex        =   41
         Top             =   1620
         Visible         =   0   'False
         Width           =   5955
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Observação:"
         Height          =   195
         Left            =   195
         TabIndex        =   17
         Top             =   2145
         Width           =   915
      End
      Begin VB.Label lblContaBancaria 
         AutoSize        =   -1  'True
         Caption         =   "Conta Corrente"
         Height          =   195
         Left            =   60
         TabIndex        =   3
         Top             =   990
         Width           =   1065
      End
      Begin VB.Label lblExercicio 
         AutoSize        =   -1  'True
         Caption         =   "Exercício"
         Height          =   195
         Left            =   450
         TabIndex        =   8
         Top             =   1830
         Width           =   675
      End
      Begin VB.Label lblcodigoBaixa 
         AutoSize        =   -1  'True
         Caption         =   "Código da Baixa"
         Height          =   195
         Left            =   5355
         TabIndex        =   15
         Top             =   1815
         Width           =   1155
      End
      Begin VB.Label lblAviso 
         AutoSize        =   -1  'True
         Caption         =   "Aviso"
         Height          =   195
         Left            =   1680
         TabIndex        =   10
         Top             =   1815
         Width           =   390
      End
      Begin VB.Label lblDataDoPagamento 
         AutoSize        =   -1  'True
         Caption         =   "Data de Baixa"
         Height          =   195
         Left            =   90
         TabIndex        =   1
         Top             =   555
         Width           =   1005
      End
      Begin VB.Label lblTributo 
         AutoSize        =   -1  'True
         Caption         =   "Composição da Receita"
         Height          =   195
         Left            =   60
         TabIndex        =   5
         Top             =   1380
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frmBaixaManual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Dim blnAlterando        As Boolean
    Dim bytOrdenacao        As Byte
    Dim blnOrdenacaoAsc     As Boolean
    Dim blnPrimeiraVez      As Boolean
    
    Dim strUltimoCaminho    As String

Private Sub chk_BaixaTotal_Click()
    If chk_BaixaTotal.Value Then
        txtdblPrincipal = "0,00"
        txtdblMulta = "0,00"
        txtdblJuros = "0,00"
        txtdblCorrecao = "0,00"
        txt_dblTotal = "0,00"
        txtdblCorreto = "0,00"
        txtintDigito.Text = ""
        Set dbcintcodigobaixa.RowSource = Nothing
            dbcintcodigobaixa.Text = ""
        Set dbcintLancamentoValor.RowSource = Nothing
            dbcintLancamentoValor.Text = ""
        TrocaCorObjeto dbcintLancamentoValor, True
        TrocaCorObjeto txtintDigito, True
        LeDaTabelaParaObj "", dbcintcodigobaixa, "Select Min(Pkid) as Pkid, Strabreviatura From " & gstrCodigoDeBaixa & " Where Byttipo = 2 Group By Strabreviatura"
    Else
        Set dbcintcodigobaixa.RowSource = Nothing
            dbcintcodigobaixa.Text = ""
        TrocaCorObjeto dbcintLancamentoValor, False
        TrocaCorObjeto txtintDigito, False
    End If
End Sub

Private Sub cmd_arquivo_Click()
    
    dlgArquivo.CancelError = True
    dlgArquivo.DialogTitle = "Selecione o arquivo"
    dlgArquivo.InitDir = strUltimoCaminho
    dlgArquivo.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
    dlgArquivo.flags = cdlOFNOverwritePrompt + cdlOFNHideReadOnly
    
    On Error GoTo err_cmd_Arquivo_Click
    
    dlgArquivo.ShowOpen
    txt_Arquivo = dlgArquivo.Filename
    strUltimoCaminho = Replace(dlgArquivo.Filename, dlgArquivo.FileTitle, "")
    Exit Sub

err_cmd_Arquivo_Click:
    If Err.Number = 32755 Then
        txt_Arquivo = ""
    End If

End Sub

Private Sub cmd_Composicao_Click()
    CarregaForm frmCadComposicaoDaReceita, dbc_intComposicaoDaReceita
End Sub

Private Sub cmd_ContaCorrente_Click()
    CarregaForm frmCadContasBancarias, dbc_intContaBancaria
End Sub

Private Sub cmd_Importar_Click()
    ImportaAcordos
End Sub

Private Sub cmdBaixar_Click()
Dim strLinha As String

    If Not blnDadosOkDebitoAutomatico Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    'Vamos identificar se Debito Automatico
    Open txt_Arquivo For Input As #1
    
ProximaLinha:
    
    Line Input #1, strLinha
    
    If Len(strLinha) = 0 Then
        GoTo ProximaLinha
    End If
    
    If optDebitoAutomatico.Value = True Then
        If UCase(Mid(strLinha, 1, 1)) = "A" Then
            
            Line Input #1, strLinha
            Close #1
            
            If UCase(Mid(strLinha, 1, 1)) = "B" Or UCase(Mid(strLinha, 1, 1)) = "F" Then
                BaixaMovimentosDebitoAutomatico
            Else
                ExibeMensagem "Este arquivo não é de Débito Automático."
            End If
        Else
            ExibeMensagem "Este arquivo não é de Débito Automático."
            Close #1
        End If
    Else
        If UCase(Mid(strLinha, 1, 1)) = "0" Then
            Close #1
            
            If Mid(strLinha, 80, 15) = "BANCO DO BRASIL" Then
                ExibeMensagem "Este arquivo não é GREM."
            ElseIf Len(strLinha) = 400 Then
                ExibeMensagem "Este arquivo não é GREM."
            Else
                BaixaMovimentoBancarioFichaCompensacaoGREM
            End If
        Else
            ExibeMensagem "Este arquivo não é de GREM."
            Close #1
        End If
    End If
    
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub dbc_intComposicaoDaReceita_Change()
    
    If dbc_intComposicaoDaReceita.MatchedWithList Then
        If dbc_strComposicaoDaReceita.BoundText <> dbc_intComposicaoDaReceita.BoundText Then
            PreencherListaDeOpcoes dbc_strComposicaoDaReceita, dbc_intComposicaoDaReceita.BoundText
            
            txt_intExercicio.Text = ""
            dbc_intNumeroAviso.BoundText = ""
            Set dbc_intNumeroAviso.RowSource = Nothing
            dbcintLancamentoValor.BoundText = ""
            Set dbcintLancamentoValor.RowSource = Nothing
            txtintDigito.Text = ""

        End If
     End If

End Sub

Private Sub dbc_intComposicaoDaReceita_Click(Area As Integer)
    DropDownDataCombo dbc_intComposicaoDaReceita, Me, Area
End Sub

Private Sub dbc_intComposicaoDaReceita_GotFocus()
    MarcaCampo dbc_intComposicaoDaReceita
End Sub

Private Sub dbc_intComposicaoDaReceita_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_intComposicaoDaReceita, Me, , KeyCode, Shift
End Sub

Private Sub dbc_intComposicaoDaReceita_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", dbc_intComposicaoDaReceita
End Sub

Private Sub dbc_intComposicaoDaReceita_LostFocus()
    LeDaTabelaParaObj "", dbc_intComposicaoDaReceita, strQueryComposicao
    If Not dbc_intComposicaoDaReceita.MatchedWithList Then
        dbc_strComposicaoDaReceita.BoundText = ""
        Set dbc_strComposicaoDaReceita.RowSource = Nothing
    End If
End Sub

Private Sub dbc_intNumeroAviso_Click(Area As Integer)
    DropDownDataCombo dbc_intNumeroAviso, Me, Area
End Sub

Private Sub dbc_intNumeroAviso_GotFocus()
    MarcaCampo dbc_intNumeroAviso
End Sub

Private Sub dbc_intNumeroAviso_KeyDown(KeyCode As Integer, Shift As Integer)
     DropDownDataCombo dbc_intNumeroAviso, Me, , KeyCode, Shift
End Sub

Private Sub dbc_intNumeroAviso_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", dbc_intNumeroAviso
End Sub

Private Sub dbc_intNumeroAviso_LostFocus()
    If dbc_intComposicaoDaReceita.MatchedWithList And Len(Trim(txt_intExercicio)) = 4 And Trim(dbc_intNumeroAviso.Text) <> "" Then
        LeDaTabelaParaObj "", dbc_intNumeroAviso, strQueryAviso
        If dbc_intNumeroAviso.MatchedWithList = False Then
            dbc_intNumeroAviso.SetFocus
        End If
    Else
        Set dbc_intNumeroAviso.RowSource = Nothing
        dbc_intNumeroAviso.Text = ""
    End If
End Sub

Private Sub dbcintcodigobaixa_Change()
Dim adoResultado    As New ADODB.Recordset
Dim blnCancelamento As Boolean
    
    blnCancelamento = False
    
    If dbcintcodigobaixa.MatchedWithList Then
        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO("SELECT bytTipo FROM " & gstrCodigoDeBaixa & " WHERE Pkid = " & dbcintcodigobaixa.BoundText, 5, adoResultado) Then
            If Not adoResultado.EOF Then
                blnCancelamento = adoResultado("bytTipo").Value = 2
            End If
        End If
    End If
    
    txtdblPrincipal = "0,00"
    txtdblMulta = "0,00"
    txtdblJuros = "0,00"
    txtdblCorrecao = "0,00"
    txt_dblTotal = "0,00"
    txtdblCorreto = "0,00"
    
    TrocaCorObjeto txtdblPrincipal, blnCancelamento
    TrocaCorObjeto txtdblMulta, blnCancelamento
    TrocaCorObjeto txtdblJuros, blnCancelamento
    TrocaCorObjeto txtdblCorrecao, blnCancelamento
    TrocaCorObjeto txt_dblTotal, blnCancelamento
    TrocaCorObjeto txtdblCorreto, True
    
End Sub

Private Sub dbcintcodigobaixa_Click(Area As Integer)
    DropDownDataCombo dbcintcodigobaixa, Me, Area
End Sub

Private Sub dbcintcodigobaixa_GotFocus()
    MarcaCampo dbcintcodigobaixa
End Sub

Private Sub dbcintcodigobaixa_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_intContaBancaria, Me, , KeyCode, Shift
End Sub

Private Sub dbcintcodigobaixa_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintcodigobaixa
End Sub

Private Sub dbcintLancamentoValor_Click(Area As Integer)
    DropDownDataCombo dbcintLancamentoValor, Me, Area
End Sub

Private Sub dbcintLancamentoValor_GotFocus()
    MarcaCampo dbcintLancamentoValor
End Sub

Private Sub dbcintLancamentoValor_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbcintLancamentoValor, Me, , KeyCode, Shift
End Sub

Private Sub dbcintLancamentoValor_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbcintLancamentoValor
End Sub

Private Sub dbcintLancamentoValor_LostFocus()
Dim adoResultado As New ADODB.Recordset

    If dbc_intNumeroAviso.MatchedWithList And Trim(dbcintLancamentoValor.Text) <> "" And Not blnAlterando Then
        
        LeDaTabelaParaObj "", dbcintLancamentoValor, strQueryParcela
        If Not dbcintLancamentoValor.MatchedWithList And Trim(dbcintLancamentoValor.Text) = "" Then
            txtdblPrincipal = "0,00"
            txtdblMulta = "0,00"
            txtdblJuros = "0,00"
            txtdblCorrecao = "0,00"
            txt_dblTotal = "0,00"
            txtdblCorreto = "0,00"
            txtintDigito.Text = ""
            Set dbcintcodigobaixa.RowSource = Nothing
                dbcintcodigobaixa.Text = ""
            Set dbcintLancamentoValor.RowSource = Nothing
                dbcintLancamentoValor.Text = ""
            Exit Sub
        End If
        txtintDigito = gstrCalculaDigitoModulo10(Trim(dbc_intNumeroAviso.Text) & Format$(Trim(dbcintLancamentoValor.Text), "000"))
        
        'Vamos carregar a combo de tipos de baixa de acordo com o vencimento da parcela
        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strQueryParcela, 5, adoResultado) Then
            If Not adoResultado.EOF Then
                LeDaTabelaParaObj "", dbcintcodigobaixa, strQueryCodigoBaixa(adoResultado("dtmDtVencimento").Value)
            End If
        End If

        If dbcintLancamentoValor.MatchedWithList = False Then
            dbcintLancamentoValor.SetFocus
        End If
        
        'Ja vamos fazer o calculo dos valores
        If Not blnDadosCalculoOk Then Exit Sub
        
        CalculaReajuste
    Else
        txtdblPrincipal = "0,00"
        txtdblMulta = "0,00"
        txtdblJuros = "0,00"
        txtdblCorrecao = "0,00"
        txt_dblTotal = "0,00"
        txtdblCorreto = "0,00"
        txtintDigito.Text = ""
        Set dbcintcodigobaixa.RowSource = Nothing
            dbcintcodigobaixa.Text = ""
        Set dbcintLancamentoValor.RowSource = Nothing
            dbcintLancamentoValor.Text = ""
    End If
End Sub

Private Sub dbc_strComposicaoDaReceita_Change()
    If dbc_strComposicaoDaReceita.MatchedWithList Then
        If dbc_strComposicaoDaReceita.BoundText <> dbc_intComposicaoDaReceita.BoundText Then
            PreencherListaDeOpcoes dbc_intComposicaoDaReceita, dbc_strComposicaoDaReceita.BoundText
            
            txt_intExercicio.Text = ""
            dbc_intNumeroAviso.BoundText = ""
            Set dbc_intNumeroAviso.RowSource = Nothing
            dbcintLancamentoValor.BoundText = ""
            Set dbcintLancamentoValor.RowSource = Nothing
            txtintDigito.Text = ""

        End If
    End If
End Sub

Private Sub dbc_strComposicaoDaReceita_Click(Area As Integer)
    DropDownDataCombo dbc_strComposicaoDaReceita, Me, Area
End Sub

Private Sub dbc_strComposicaoDaReceita_GotFocus()
    MarcaCampo dbc_strComposicaoDaReceita
End Sub

Private Sub dbc_strComposicaoDaReceita_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_strComposicaoDaReceita, Me, , KeyCode, Shift
End Sub

Private Sub dbc_strComposicaoDaReceita_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_strComposicaoDaReceita
End Sub

Private Sub dbc_strDescricaoConta_Change()
    If dbc_strDescricaoConta.MatchedWithList Then
        If dbc_strDescricaoConta.BoundText <> dbc_intContaBancaria.BoundText Then
            PreencherListaDeOpcoes dbc_intContaBancaria, dbc_strDescricaoConta.BoundText
        End If
    End If
End Sub

Private Sub dbc_strDescricaoConta_Click(Area As Integer)
    DropDownDataCombo dbc_strDescricaoConta, Me, Area
End Sub

Private Sub dbc_strDescricaoConta_GotFocus()
    MarcaCampo dbc_strDescricaoConta
End Sub

Private Sub dbc_strDescricaoConta_LostFocus()
    If Not dbc_strDescricaoConta.MatchedWithList Then dbc_intContaBancaria.BoundText = ""
End Sub

Private Sub dbc_strDescricaoConta_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_strDescricaoConta, Me, , KeyCode, Shift
End Sub

Private Sub dbc_strDescricaoConta_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_strDescricaoConta
End Sub

Private Sub dbc_intContaBancaria_Change()
    If dbc_intContaBancaria.MatchedWithList Then
        If dbc_strDescricaoConta.BoundText <> dbc_intContaBancaria.BoundText Then
            PreencherListaDeOpcoes dbc_strDescricaoConta, dbc_intContaBancaria.BoundText
        End If
    End If
End Sub

Private Sub dbc_intContaBancaria_LostFocus()
    LeDaTabelaParaObj "", dbc_intContaBancaria, strQueryContaCorrente
    If Not dbc_intContaBancaria.MatchedWithList Then
        dbc_strDescricaoConta.BoundText = ""
        Set dbc_strDescricaoConta.RowSource = Nothing
    End If
End Sub

Private Sub dbc_intContaBancaria_Click(Area As Integer)
    DropDownDataCombo dbc_intContaBancaria, Me, Area
End Sub

Private Sub dbc_intContaBancaria_KeyDown(KeyCode As Integer, Shift As Integer)
    DropDownDataCombo dbc_intContaBancaria, Me, , KeyCode, Shift
End Sub

Public Sub dbc_intContaBancaria_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "A", dbc_intContaBancaria
    If KeyAscii = 0 Then
            Set dbc_intComposicaoDaReceita.RowSource = Nothing
            Set dbc_intNumeroAviso.RowSource = Nothing
            Set dbcintLancamentoValor.RowSource = Nothing
            Set dbc_intContaBancaria.RowSource = Nothing
            Set dbc_strDescricaoConta.RowSource = Nothing
        dbc_intContaBancaria.Text = ""
        dbc_strDescricaoConta.Text = ""
        dbc_intComposicaoDaReceita.Text = ""
        dbc_strComposicaoDaReceita.Text = ""
    End If
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 1144
    HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrDeletar
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrImprimir
End Sub

Private Sub Form_Deactivate()
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
End Sub

Private Sub Form_Load()
    dbc_intComposicaoDaReceita.Tag = strQueryComposicao(True) & ";strDescricao"
    dbc_strComposicaoDaReceita.Tag = strQueryComposicaoDescricao & ";strdescricao"
    dbcintcodigobaixa.Tag = strQueryCodigoBaixa & ";strAbreviatura"
    dbc_intContaBancaria.Tag = strQueryContaCorrente(True) & ";intNumeroConta"
    dbc_strDescricaoConta.Tag = strQueryContaDescricao & ";CB.strdescricao"
    TrocaCorObjeto txtdblCorreto, True
    
    tab_3dPasta.TabVisible(1) = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar, gstrDeletar
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
Dim strSQL          As String
Dim adoResultado    As ADODB.Recordset
    Select Case UCase(strModoOperacao)
    
        Case Is = UCase(gstrSalvar)
            If Not blnDadosOk Then Exit Sub
            
            If gblnExclusaoGravacaoOk("I") Then
                Screen.MousePointer = vbArrowHourglass
                If GravaBaixaMaunual Then
                    MantemForm gstrNovo
                End If
                Screen.MousePointer = vbDefault
            End If
        Case Is = UCase(gstrPreencherLista)
        
            If Me.ActiveControl.Name = "dbc_intNumeroAviso" Then
                If dbc_intComposicaoDaReceita.MatchedWithList And Len(Trim(txt_intExercicio)) = 4 Then
                    LeDaTabelaParaObj "", dbc_intNumeroAviso, strQueryAviso(True)
                End If
            ElseIf Me.ActiveControl.Name = "dbcintLancamentoValor" Then
                If dbc_intNumeroAviso.MatchedWithList Then
                    LeDaTabelaParaObj "", dbcintLancamentoValor, strQueryParcela(True)
                End If
            Else
                PreencherListaDeOpcoes Me.ActiveControl
            End If
        Case Is = UCase(gstrNovo)
        
            Limpa_Controles Me, True, True, False, True, False
            Set dbc_intComposicaoDaReceita.RowSource = Nothing
            Set dbc_intNumeroAviso.RowSource = Nothing
            Set dbcintLancamentoValor.RowSource = Nothing
            Set dbc_intContaBancaria.RowSource = Nothing
            Set dbc_strDescricaoConta.RowSource = Nothing
            txt_dblTotal.Text = ""
            txt_dtmDtBaixa = gstrDataDoSistema
            txtstrObservacao.Text = ""
            txt_dtmDtBaixa.SetFocus
        Case Else
    End Select
End Sub

Private Function blnDadosOk()
    blnDadosOk = False
    
    If Not gblnDataValida(txt_dtmDtBaixa) Then
        ExibeMensagem "A Data informada não é válida."
        txt_dtmDtBaixa.SetFocus
        Exit Function
    ElseIf dbc_intComposicaoDaReceita.MatchedWithList = False Then
        ExibeMensagem "A Composição da Recita deve ser preenchida corretamente."
        dbc_intComposicaoDaReceita.SetFocus
        Exit Function
    ElseIf Trim(txt_intExercicio) = "" Then
        ExibeMensagem "O Exercício deve ser preenchido corretamente."
        txt_intExercicio.SetFocus
        Exit Function
    ElseIf dbc_intNumeroAviso.MatchedWithList = False Then
        ExibeMensagem "O número do aviso deve ser preenchido corretamente."
        If dbc_intNumeroAviso.Enabled Then dbc_intNumeroAviso.SetFocus
        Exit Function
    ElseIf chk_BaixaTotal.Value = 0 Then
        If dbcintLancamentoValor.MatchedWithList = False Then
            ExibeMensagem "A parcela deve ser preenchida corretamente."
            dbcintLancamentoValor.SetFocus
            Exit Function
        ElseIf Trim(txtintDigito) = "" Then
            ExibeMensagem "O Digito deve ser preenchido corretamente."
            txtintDigito.SetFocus
            Exit Function
        ElseIf dbcintcodigobaixa.MatchedWithList = False Then
            ExibeMensagem "O código da baixa deve ser preenchido corretamente."
            dbcintcodigobaixa.SetFocus
            Exit Function
        ElseIf VerificaDuplicado = True Then
            Exit Function
        End If
    End If
    
    If dbcintcodigobaixa.MatchedWithList = False Then
        ExibeMensagem "O código da baixa deve ser preenchido corretamente."
        dbcintcodigobaixa.SetFocus
        Exit Function
    End If
    
    If Weekday(CDate(txt_dtmDtBaixa)) = 7 Then
        ExibeMensagem "Baixa manual cancelada pois a data informada cai no Sábado."
        txt_dtmDtBaixa.SetFocus
        Exit Function
    ElseIf Weekday(CDate(txt_dtmDtBaixa)) = 1 Then
        ExibeMensagem "Baixa manual cancelada pois a data informada cai no Domingo."
        txt_dtmDtBaixa.SetFocus
        Exit Function
    End If
    
    If Val(dbc_intContaBancaria.BoundText) > 0 Then
        If CDate(txt_dtmDtBaixa) <= VerificaDataEncerramento("EF", Year(txt_dtmDtBaixa)) Then
            ExibeMensagem "Baixa inválida!" & Chr(13) & "Data menor ou igual que data do fechamento."
            txt_dtmDtBaixa.SetFocus
            Exit Function
        End If
    End If
    
    
    
    blnDadosOk = True
    
End Function

Private Function blnDadosCalculoOk()

    blnDadosCalculoOk = False
    
    If dbc_intComposicaoDaReceita.MatchedWithList = False Then
        Exit Function
    ElseIf Trim(txt_intExercicio) = "" Then
        Exit Function
    ElseIf dbc_intNumeroAviso.MatchedWithList = False Then
        Exit Function
    ElseIf dbcintLancamentoValor.MatchedWithList = False Then
        Exit Function
    ElseIf Not gblnDataValida(txt_dtmDtBaixa.Text) Then
        Exit Function
    End If
    
    blnDadosCalculoOk = True
    
End Function

Private Sub tab_3dPasta_DblClick()
'    tab_3dPasta.TabVisible(1) = True
End Sub

Private Sub txt_dblTotal_GotFocus()
    MarcaCampo txt_dblTotal
End Sub

Private Sub txt_dblTotal_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txt_dblTotal
End Sub

Private Sub txt_dblTotal_LostFocus()
    txt_dblTotal = gstrConvVrDoSql(txt_dblTotal, 2)
    dblValorPrincipal
End Sub

Private Sub txtintDigito_GotFocus()
    MarcaCampo txtintDigito
End Sub

Private Sub txtintDigito_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txtintDigito
End Sub

Private Sub txtintDigito_LostFocus()
    If dbc_intNumeroAviso.MatchedWithList And dbcintLancamentoValor.MatchedWithList Then
        If gstrCalculaDigitoModulo10(Trim(dbc_intNumeroAviso.Text) & Format$(Trim(dbcintLancamentoValor.Text), "000")) <> Trim(txtintDigito) Then
            ExibeMensagem "Aviso inválido."
            txtintDigito = ""
        End If
    End If
End Sub

Private Sub txt_dtmDtBaixa_LostFocus()
    txt_dtmDtBaixa = gstrDataFormatada(txt_dtmDtBaixa)
End Sub

Private Sub txt_intExercicio_Change()
    Set dbc_intNumeroAviso.RowSource = Nothing
    dbc_intNumeroAviso.Text = ""
    Set dbcintLancamentoValor.RowSource = Nothing
    dbcintLancamentoValor.Text = ""
    txtintDigito.Text = ""
End Sub

Private Sub txt_intExercicio_GotFocus()
    MarcaCampo txt_intExercicio
End Sub

Private Sub txt_intExercicio_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_intExercicio
End Sub

Private Sub txtdblCorrecao_GotFocus()
    MarcaCampo txtdblCorrecao
End Sub

Private Sub txtdblCorrecao_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txtdblCorrecao
End Sub

Private Sub txtdblCorrecao_LostFocus()
    txtdblCorrecao = gstrConvVrDoSql(txtdblCorrecao, 2)
    dblValorTotal
End Sub

Private Sub txtdblJuros_GotFocus()
    MarcaCampo txtdblJuros
End Sub

Private Sub txtdblJuros_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txtdblJuros
End Sub

Private Sub txtdblJuros_LostFocus()
    txtdblJuros = gstrConvVrDoSql(txtdblJuros, 2)
    dblValorTotal
End Sub

Private Sub txtdblMulta_GotFocus()
    MarcaCampo txtdblMulta
End Sub

Private Sub txtdblMulta_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txtdblMulta
End Sub

Private Sub txtdblMulta_LostFocus()
    txtdblMulta = gstrConvVrDoSql(txtdblMulta, 2)
    dblValorTotal
End Sub


Private Sub txtdblPrincipal_GotFocus()
    MarcaCampo txtdblPrincipal
End Sub

Private Sub txtdblPrincipal_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "V", txtdblPrincipal
End Sub

Private Sub txtdblPrincipal_LostFocus()
    txtdblPrincipal = gstrConvVrDoSql(txtdblPrincipal, 2)
    dblValorTotal
End Sub

Private Sub txt_dtmDtBaixa_GotFocus()
    If txt_dtmDtBaixa.Text = "" Then txt_dtmDtBaixa = gstrDataDoSistema
    MarcaCampo txt_dtmDtBaixa
End Sub

Private Sub txt_dtmDtBaixa_KeyPress(KeyAscii As Integer)

    CaracterValido KeyAscii, "D", txt_dtmDtBaixa
End Sub

Private Function strCompletaNumero(strValor As String, intNumeroCasas As Integer) As String
    Dim intI       As Integer
    Dim strDigito  As String
    
    For intI = 1 To gstrENulo(intNumeroCasas) - Len(strValor)
        strDigito = strDigito & "0"
    Next intI
    strCompletaNumero = strDigito
    
End Function

Private Function strNumeroAviso(lngPkidMovimentoBancario As Long) As String

Dim strSQL          As String
Dim adoResultado    As ADODB.Recordset

    strSQL = "SELECT " & gstrCONVERT(CDT_numeric, "LA.strNumeroAviso") & " strNumeroAviso, LV.intParcela"
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrLancamentoAlfa & " LA,"
    strSQL = strSQL & gstrLancamentoValor & " LV,"
    strSQL = strSQL & gstrMovimentoBancario & " MB"
    strSQL = strSQL & " WHERE "
    strSQL = strSQL & "MB.Pkid =" & lngPkidMovimentoBancario & " AND "
    strSQL = strSQL & "MB.intLancamentoValor " & strOUTJSQLServer & "= LV.Pkid " & strOUTJOracle & " AND"
    strSQL = strSQL & " LV.intLancamentoAlfa " & strOUTJSQLServer & "= LA.Pkid " & strOUTJOracle

    Set gobjBanco = New clsBanco
    
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        If Not adoResultado.EOF Then
            strNumeroAviso = strCompletaNumero(gstrENulo(adoResultado!strNumeroAviso), 6) & gstrENulo(adoResultado!strNumeroAviso) & strCompletaNumero(gstrENulo(adoResultado!intParcela), 3) & gstrENulo(adoResultado!intParcela) & "0"
        Else
            strNumeroAviso = ""
        End If
    End If
    
End Function

Private Function strQueryComposicao(Optional blnF5 As Boolean) As String
    Dim strSQL As String
    
    strSQL = "SELECT Pkid,"
    strSQL = strSQL & "intCodigo "
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrComposicaoDaReceita
    If blnF5 = False Then
        strSQL = strSQL & " WHERE intCodigo = " & Val(dbc_intComposicaoDaReceita.Text)
    End If
    strSQL = strSQL & " ORDER BY intCodigo"
    
    strQueryComposicao = strSQL

End Function

Private Sub dblValorTotal()
    Dim dblValorTotal As Variant
    
    dblValorTotal = CDbl(gstrConvVrDoSql(txtdblPrincipal.Text, 2, , True)) + _
                    CDbl(gstrConvVrDoSql(txtdblMulta.Text, 2, , True)) + _
                    CDbl(gstrConvVrDoSql(txtdblJuros.Text, 2, , True)) + _
                    CDbl(gstrConvVrDoSql(txtdblCorrecao.Text, 2, , True))
    
    txt_dblTotal = gstrConvVrDoSql(dblValorTotal, 2)
    
End Sub

Private Function strQueryAviso(Optional blnF5 As Boolean) As String
    Dim strSQL As String
    Dim strAux As String
    
    
    strSQL = ""
    strSQL = strSQL & "Select "
    strSQL = strSQL & "LA.Pkid, "
    strSQL = strSQL & gstrCONVERT(CDT_numeric, "LA.strNumeroAviso") & " strNumeroAviso "
    strSQL = strSQL & "From "
    strSQL = strSQL & gstrLancamentoAlfa & " LA, "
    strSQL = strSQL & gstrComposicaoDaReceita & " CR "
    strSQL = strSQL & "Where "
    strSQL = strSQL & "CR.Pkid = LA.Intcomposicaodareceita AND "
    strSQL = strSQL & "CR.Pkid = " & dbc_intComposicaoDaReceita.BoundText & " AND "
    strSQL = strSQL & "LA.intExercicio = " & Trim(txt_intExercicio)
    If blnF5 = False Or Trim(dbc_intNumeroAviso) <> "" Then
        strAux = "'" & String(gintLenNumAviso - Len(Trim(dbc_intNumeroAviso.Text)), "0") & Val(dbc_intNumeroAviso.Text) & "'"
        strSQL = strSQL & " AND LA.strNumeroAviso = " & strSUBSTRING & "(" & strAux & ", " & strLen & "(" & strAux & ") - " & strLen & "(LA.strNumeroAviso) + 1, " & strLen & "(LA.strNumeroAviso))"
    End If
    strSQL = strSQL & " Order By " & gstrCONVERT(CDT_numeric, "LA.strNumeroAviso")
    
    strQueryAviso = strSQL

End Function

Private Function strQueryCodigoBaixa(Optional dtmVencimento As Date) As String
    Dim strSQL As String
    
    strSQL = ""
    strSQL = strSQL & "Select "
    strSQL = strSQL & "Pkid, "
    strSQL = strSQL & "strAbreviatura "
    strSQL = strSQL & "From "
    strSQL = strSQL & gstrCodigoDeBaixa & " "
    If Trim(txt_dtmDtBaixa) <> "" And Trim(dtmVencimento) <> "" Then
        If CDate(dtmVencimento) < CDate(txt_dtmDtBaixa) Then
            strSQL = strSQL & " WHERE BytTipo = 4 "
        Else
            strSQL = strSQL & " WHERE BytTipo = 0 "
        End If
        strSQL = strSQL & "Order By Pkid "
    Else
        strSQL = strSQL & "Order By strabreviatura "
    End If
    
    strQueryCodigoBaixa = strSQL
    
End Function

Private Function strQueryParcela(Optional blnF5 As Boolean) As String
    Dim strSQL As String
    
    strSQL = ""
    strSQL = strSQL & "Select "
    strSQL = strSQL & "LV.Pkid, "
    strSQL = strSQL & "LV.Intparcela, "
    strSQL = strSQL & "LV.dtmDtVencimento "
    strSQL = strSQL & "From "
    strSQL = strSQL & gstrLancamentoValor & " LV, "
    strSQL = strSQL & gstrLancamentoPagamento & " LP "
    strSQL = strSQL & "Where "
    strSQL = strSQL & "LV.pkid " & strOUTJSQLServer & "= LP.intLancamentoValor" & strOUTJOracle & " And "
    strSQL = strSQL & "LP.intLancamentoValor Is Null And "
    strSQL = strSQL & "LV.Intlancamentoalfa = " & dbc_intNumeroAviso.BoundText
    
    If blnF5 = False Then
        strSQL = strSQL & " AND LV.Intparcela = " & dbcintLancamentoValor.Text
    End If
    
    strQueryParcela = strSQL
    
End Function

Private Function GravaPagamento() As String
    Dim strSQL As String
    
    strSQL = ""
    strSQL = strSQL & "Insert Into "
    strSQL = strSQL & gstrLancamentoPagamento & " ("
    strSQL = strSQL & "intlancamentovalor, "
    strSQL = strSQL & "dblvalorprincipal, "
    strSQL = strSQL & "dblvalormulta, "
    strSQL = strSQL & "dblvalorjuros, "
    strSQL = strSQL & "dblvalorcorrecao, "
    strSQL = strSQL & "dtmdtpagamento, "
    strSQL = strSQL & "intcodigobaixa, "
    strSQL = strSQL & "strObservacao, "
    strSQL = strSQL & "dtmDtAtualizacao, lngCodUsr) "
    strSQL = strSQL & "VaLues( "
    
    strSQL = strSQL & dbcintLancamentoValor.BoundText & ", "
    strSQL = strSQL & gstrConvVrParaSql(IIf(Trim(txtdblPrincipal) = "", 0, txtdblPrincipal)) & ", "
    strSQL = strSQL & gstrConvVrParaSql(IIf(Trim(txtdblMulta) = "", 0, txtdblMulta)) & ", "
    strSQL = strSQL & gstrConvVrParaSql(IIf(Trim(txtdblJuros) = "", 0, txtdblJuros)) & ", "
    strSQL = strSQL & gstrConvVrParaSql(IIf(Trim(txtdblCorrecao) = "", 0, txtdblCorrecao)) & ", "
    strSQL = strSQL & gstrConvDtParaSql(Trim(txt_dtmDtBaixa)) & ", "
    strSQL = strSQL & dbcintcodigobaixa.BoundText & ", '"
    strSQL = strSQL & txtstrObservacao.Text & "', "
    strSQL = strSQL & strGETDATE & ", "
    strSQL = strSQL & glngCodUsr
    strSQL = strSQL & ") "
    
    GravaPagamento = strSQL
    
End Function

Private Function VerificaDuplicado() As Boolean
    Dim strSQL As String
    Dim adoResultado As ADODB.Recordset
    
    VerificaDuplicado = False
    
    strSQL = " "
    strSQL = strSQL & "Select * from "
    strSQL = strSQL & gstrLancamentoPagamento & " Where "
    strSQL = strSQL & "intlancamentovalor = " & dbcintLancamentoValor.BoundText & " AND "
    strSQL = strSQL & "dblvalorprincipal = " & gstrConvVrParaSql(IIf(Trim(txtdblPrincipal) = "", 0, txtdblPrincipal)) & " AND "
    strSQL = strSQL & "dblvalormulta = " & gstrConvVrParaSql(IIf(Trim(txtdblMulta) = "", 0, txtdblMulta)) & " AND "
    strSQL = strSQL & "dblvalorjuros = " & gstrConvVrParaSql(IIf(Trim(txtdblJuros) = "", 0, txtdblJuros)) & " AND "
    strSQL = strSQL & "dblvalorcorrecao = " & gstrConvVrParaSql(IIf(Trim(txtdblCorrecao) = "", 0, txtdblCorrecao)) & " AND "
    strSQL = strSQL & "dtmdtpagamento = " & gstrConvDtParaSql(Trim(txt_dtmDtBaixa)) & " AND "
    strSQL = strSQL & "intcodigobaixa= " & dbcintcodigobaixa.BoundText
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        If adoResultado.RecordCount >= 1 Then
            If MsgBox("Baixa já efetuada. " & Chr(13) & "Deseja executar nova baixa ?", vbYesNo + vbQuestion) = vbNo Then
                VerificaDuplicado = True
            End If
        End If
    End If
    
End Function

Private Function strQueryContaCorrente(Optional blnF5 As Boolean) As String
    Dim strSQL As String

    strSQL = "SELECT CB.Pkid, "
    strSQL = strSQL & "intNumeroConta ContaCorrente"
    strSQL = strSQL & " FROM " & gstrContaBancaria & " CB, "
    strSQL = strSQL & gstrPlanoConta & " PC"
    strSQL = strSQL & " Where"
    strSQL = strSQL & " CB.Pkid = PC.Intcontabancaria"
    If blnF5 = False Then
        strSQL = strSQL & " AND CB.intNumeroConta = " & Val(dbc_intContaBancaria.Text)
    End If
    strSQL = strSQL & " ORDER BY intNumeroConta, strDigitoVerificador"
    
    strQueryContaCorrente = strSQL

End Function

Private Function strQueryContaDescricao() As String
    Dim strSQL As String

    strSQL = strSQL & "SELECT "
    strSQL = strSQL & "CB.Pkid, "
    strSQL = strSQL & "CB.strdescricao"
    strSQL = strSQL & " FROM " & gstrContaBancaria & " CB, "
    strSQL = strSQL & gstrPlanoConta & " PC"
    strSQL = strSQL & " Where"
    strSQL = strSQL & " CB.Pkid = PC.Intcontabancaria"
    strSQL = strSQL & " ORDER BY CB.strdescricao"
    
    strQueryContaDescricao = strSQL

End Function

Private Sub CalculaReajuste()
    Dim strSQL       As String
    Dim adoResultado As New ADODB.Recordset
    Dim adoParcelas  As New ADODB.Recordset
    
    Set gobjBanco = New clsBanco
    
    strSQL = "SELECT LV.dblValor ValorOrig, LV.dtmDtVencimento, LV.intMoeda " & _
             "FROM " & gstrLancamentoValor & " LV, " & gstrLancamentoAlfa & " LA " & _
             "WHERE LV.intLancamentoAlfa = LA.pkid " & _
             " AND LV.Pkid not in(SELECT Intlancamentovalor FROM " & gstrLancamentoPagamento & ") AND LA.Pkid = " & dbc_intNumeroAviso.BoundText & _
             " AND LV.intParcela = " & dbcintLancamentoValor.Text

    If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
        With adoResultado
            
        If Not adoResultado.EOF Then

            strSQL = gstrStoredProcedure("sp_AtualizaParcela", dbc_intComposicaoDaReceita.BoundText & ", " & txt_intExercicio & ", " & dbcintLancamentoValor.Text & ", " & gstrConvDtParaSql(!Dtmdtvencimento) & ", " & gstrConvDtParaSql(txt_dtmDtBaixa) & ", " & gstrConvVrParaSql(!ValorOrig) & ", " & !intMoeda, True)
            If gobjBanco.CriaADO(strSQL, 80, adoParcelas) Then
                txtdblPrincipal = Space$(0) & gstrConvVrDoSql(adoParcelas("dblValorPrincipal").Value)
                txtdblMulta = Space$(0) & gstrConvVrDoSql(adoParcelas("dblValorMulta").Value)
                txtdblJuros = Space$(0) & gstrConvVrDoSql(adoParcelas("dblValorJuros").Value)
                txtdblCorrecao = Space$(0) & gstrConvVrDoSql(adoParcelas("dblValorCorrecao").Value)
                txtdblCorreto = Space$(0) & gstrConvVrDoSql(Val(gstrConvVrParaSql(adoParcelas("dblValorPrincipal").Value)) + Val(gstrConvVrParaSql(adoParcelas("dblValorMulta").Value)) + Val(gstrConvVrParaSql(adoParcelas("dblValorJuros").Value)) + Val(gstrConvVrParaSql(adoParcelas("dblValorCorrecao").Value)))
                dblValorTotal
            Else
                txtdblPrincipal = Space$(0)
                txtdblMulta = Space$(0)
                txtdblJuros = Space$(0)
                txtdblCorrecao = Space$(0)
                txtdblCorreto = Space$(0)
            End If
        
        Else
            ExibeMensagem "Não foram encontrados lançamentos para esta Inscrição."
            Exit Sub
        End If
        
        End With
    End If
    
End Sub

Private Sub dblValorPrincipal()
    Dim dblValorPrincipal As Variant
    
    dblValorPrincipal = CDbl(gstrConvVrDoSql(txt_dblTotal.Text, 2, , True)) - _
                        CDbl(gstrConvVrDoSql(txtdblMulta.Text, 2, , True)) - _
                        CDbl(gstrConvVrDoSql(txtdblJuros.Text, 2, , True)) - _
                        CDbl(gstrConvVrDoSql(txtdblCorrecao.Text, 2, , True))
    
    txtdblPrincipal = gstrConvVrDoSql(dblValorPrincipal, 2)
    
End Sub

Private Function strQueryComposicaoDescricao() As String
    Dim strSQL As String
    
    strSQL = "SELECT Pkid,"
    strSQL = strSQL & " strDescricao Descricao "
    strSQL = strSQL & " FROM "
    strSQL = strSQL & gstrComposicaoDaReceita
    strSQL = strSQL & " ORDER BY strDescricao"
    strQueryComposicaoDescricao = strSQL

End Function

Private Function GravaBaixaMaunual() As Boolean
    Dim strSQL          As String
    Dim adoResultado    As ADODB.Recordset
    Dim adoParcelas     As ADODB.Recordset
    Dim dblPrincipal    As Double
    Dim dblMulta        As Double
    Dim dblJuros        As Double
    Dim dblCorrecao     As Double
    
    GravaBaixaMaunual = False
    
    Set gobjBanco = New clsBanco
    gobjBanco.ExecutaBeginTrans
    
    If chk_BaixaTotal.Value = 0 Then
        If gobjBanco.Execute(GravaPagamento) Then
        
            strSQL = "SELECT CO.intUtilizacao, "
            strSQL = strSQL & " (SELECT intLancamentoAlfa FROM " & gstrLancamentoValor & " WHERE Pkid = " & dbcintLancamentoValor.BoundText & ") intLancamentoAlfa "
            strSQL = strSQL & " FROM "
            strSQL = strSQL & gstrComposicaoDaReceita & " CO "
            strSQL = strSQL & " WHERE "
            strSQL = strSQL & "CO.Pkid =" & dbc_intComposicaoDaReceita.BoundText
        
            Set gobjBanco = New clsBanco
            If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
                If Not adoResultado.EOF Then
                   'Vamos verificar se é um Acordo
                    If adoResultado("intUtilizacao").Value = TYP_ACORDO Then
                        gQuitacaoDeAcordos adoResultado("intLancamentoAlfa").Value, txt_dtmDtBaixa.Text
                    End If
                End If
            End If
         
            If gblnBaixaCancelamento(dbc_intNumeroAviso.BoundText, dbc_intComposicaoDaReceita.BoundText, Trim(txt_intExercicio), dbcintLancamentoValor.Text, txt_dtmDtBaixa.Text, True, False) = True Then
                If gblnAnaliseDaReceita(dbcintLancamentoValor.BoundText, Val(dbc_intContaBancaria.BoundText), dbc_intComposicaoDaReceita.BoundText, Val(gstrConvVrParaSql(txtdblPrincipal)), Val(gstrConvVrParaSql(txtdblMulta)), Val(gstrConvVrParaSql(txtdblJuros)), Val(gstrConvVrParaSql(txtdblCorrecao)), txt_dtmDtBaixa, dbc_intNumeroAviso.BoundText, False, True, True) = True Then
                    gobjBanco.ExecutaCommitTrans
                Else
                    gobjBanco.ExecutaRollbackTrans
                    Set aAnaliseReceita = New XArrayDB
                    Set aAnaliseReceita = Nothing
                    Exit Function
                End If
                
                Set aAnaliseReceita = New XArrayDB
                Set aAnaliseReceita = Nothing
            Else
                gobjBanco.ExecutaRollbackTrans
                ExibeMensagem "Não foi possivel efetuar baixa manual devido às inconsistências das parcelas canceladas."
                Exit Function
            End If
        Else
            gobjBanco.ExecutaRollbackTrans
            ExibeMensagem "Não foi possivel efetuar baixa manual devido às inconsistências das parcelas canceladas."
            Exit Function
        End If
    Else
        strSQL = ""
        strSQL = strSQL & "Select "
        strSQL = strSQL & "LV.Pkid, "
        strSQL = strSQL & "LV.Intparcela, "
        strSQL = strSQL & "LV.DBLVALOR ValorOrig, "
        strSQL = strSQL & "LV.dtmDtVencimento, LV.intMoeda "
        strSQL = strSQL & "From "
        strSQL = strSQL & gstrLancamentoValor & " LV, "
        strSQL = strSQL & gstrLancamentoPagamento & " LP "
        strSQL = strSQL & "Where "
        strSQL = strSQL & "LV.pkid " & strOUTJSQLServer & "= LP.intLancamentoValor" & strOUTJOracle & " And "
        strSQL = strSQL & "LP.intLancamentoValor Is Null And "
        strSQL = strSQL & "LV.Bitparcelavalida = 1 And "
        strSQL = strSQL & "LV.Intlancamentoalfa = " & dbc_intNumeroAviso.BoundText
        
        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
            If Not adoResultado.EOF Then
                Do While Not adoResultado.EOF
                    If Not txtdblPrincipal.Enabled Then
'                        strSql = gstrStoredProcedure("sp_AtualizaParcela", dbc_intComposicaoDaReceita.BoundText & ", " & txt_intExercicio & ", " & adoResultado("Pkid").Value & ", " & gstrConvDtParaSql(adoResultado("Dtmdtvencimento").Value) & ", " & gstrConvDtParaSql(gstrDataDoSistema) & ", " & gstrConvVrParaSql(adoResultado("ValorOrig").Value) & ", " & adoResultado("intMoeda").Value, True)
'                        Set gobjBanco = New clsBanco
'                        If gobjBanco.CriaADO(strSql, 5, adoParcelas) Then
'                            dblPrincipal = gstrConvVrDoSql(gstrENulo(adoParcelas("dblValorPrincipal").Value), , , True)
'                            dblMulta = gstrConvVrDoSql(gstrENulo(adoParcelas("dblValorMulta").Value), , , True)
'                            dblJuros = gstrConvVrDoSql(gstrENulo(adoParcelas("dblValorJuros").Value), , , True)
'                            dblCorrecao = gstrConvVrDoSql(gstrENulo(adoParcelas("dblValorCorrecao").Value), , , True)
'                        End If
                        dblPrincipal = CDbl(gstrConvVrDoSql(txtdblPrincipal, , , True))
                        dblMulta = CDbl(gstrConvVrDoSql(txtdblMulta, , , True))
                        dblJuros = CDbl(gstrConvVrDoSql(txtdblJuros, , , True))
                        dblCorrecao = CDbl(gstrConvVrDoSql(txtdblCorrecao, , , True))
                    Else
                        dblPrincipal = CDbl(gstrConvVrDoSql(txtdblPrincipal, , , True)) / adoResultado.RecordCount
                        dblMulta = CDbl(gstrConvVrDoSql(txtdblMulta, , , True)) / adoResultado.RecordCount
                        dblJuros = CDbl(gstrConvVrDoSql(txtdblJuros, , , True)) / adoResultado.RecordCount
                        dblCorrecao = CDbl(gstrConvVrDoSql(txtdblCorrecao, , , True)) / adoResultado.RecordCount
                    End If
                    
                    strSQL = ""
                    strSQL = strSQL & "Insert Into "
                    strSQL = strSQL & gstrLancamentoPagamento & " ("
                    strSQL = strSQL & "intlancamentovalor, "
                    strSQL = strSQL & "dblvalorprincipal, "
                    strSQL = strSQL & "dblvalormulta, "
                    strSQL = strSQL & "dblvalorjuros, "
                    strSQL = strSQL & "dblvalorcorrecao, "
                    strSQL = strSQL & "dtmdtpagamento, "
                    strSQL = strSQL & "intcodigobaixa, "
                    strSQL = strSQL & "strObservacao, "
                    strSQL = strSQL & "dtmDtAtualizacao, lngCodUsr) "
                    strSQL = strSQL & "VaLues( "
                    
                    strSQL = strSQL & adoResultado("Pkid").Value & ", "
                    strSQL = strSQL & gstrConvVrParaSql(IIf(Trim(dblPrincipal) = "", 0, dblPrincipal)) & ", "
                    strSQL = strSQL & gstrConvVrParaSql(IIf(Trim(dblMulta) = "", 0, dblMulta)) & ", "
                    strSQL = strSQL & gstrConvVrParaSql(IIf(Trim(dblJuros) = "", 0, dblJuros)) & ", "
                    strSQL = strSQL & gstrConvVrParaSql(IIf(Trim(dblCorrecao) = "", 0, dblCorrecao)) & ", "
                    strSQL = strSQL & gstrConvDtParaSql(Trim(txt_dtmDtBaixa)) & ", "
                    strSQL = strSQL & dbcintcodigobaixa.BoundText & ", '"
                    strSQL = strSQL & txtstrObservacao.Text & "', "
                    strSQL = strSQL & strGETDATE & ", "
                    strSQL = strSQL & glngCodUsr
                    strSQL = strSQL & ") "
                    
                    If gobjBanco.Execute(strSQL) Then
                        strSQL = "SELECT CO.intUtilizacao, "
                        strSQL = strSQL & " (SELECT intLancamentoAlfa FROM " & gstrLancamentoValor & " WHERE Pkid = " & adoResultado("Pkid").Value & ") intLancamentoAlfa "
                        strSQL = strSQL & " FROM "
                        strSQL = strSQL & gstrComposicaoDaReceita & " CO "
                        strSQL = strSQL & " WHERE "
                        strSQL = strSQL & "CO.Pkid =" & dbc_intComposicaoDaReceita.BoundText
                    
                        Set gobjBanco = New clsBanco
                        If gobjBanco.CriaADO(strSQL, 5, adoParcelas) Then
                            If Not adoResultado.EOF Then
                               'Vamos verificar se é um Acordo
                                If adoParcelas("intUtilizacao").Value = TYP_ACORDO Then
                                    gQuitacaoDeAcordos adoParcelas("intLancamentoAlfa").Value, txt_dtmDtBaixa.Text
                                End If
                            End If
                        End If
                        Screen.MousePointer = vbArrowHourglass
                        If gblnBaixaCancelamento(dbc_intNumeroAviso.BoundText, dbc_intComposicaoDaReceita.BoundText, Trim(txt_intExercicio), adoResultado("PKID").Value, txt_dtmDtBaixa.Text, True, False) = True Then
                            If gblnAnaliseDaReceita(adoResultado("PKID").Value, Val(dbc_intContaBancaria.BoundText), dbc_intComposicaoDaReceita.BoundText, Val(gstrConvVrParaSql(dblPrincipal)), Val(gstrConvVrParaSql(dblMulta)), Val(gstrConvVrParaSql(dblJuros)), Val(gstrConvVrParaSql(dblCorrecao)), txt_dtmDtBaixa, dbc_intNumeroAviso.BoundText, False, True, True) = True Then
                                gobjBanco.ExecutaCommitTrans
                            Else
                                gobjBanco.ExecutaRollbackTrans
                                Set aAnaliseReceita = New XArrayDB
                                Set aAnaliseReceita = Nothing
                                Exit Function
                            End If
                            
                            Set aAnaliseReceita = New XArrayDB
                            Set aAnaliseReceita = Nothing
                        Else
                            gobjBanco.ExecutaRollbackTrans
                            ExibeMensagem "Não foi possivel efetuar baixa manual devido às inconsistências das parcelas canceladas."
                            Exit Function
                        End If
                    Else
                        gobjBanco.ExecutaRollbackTrans
                        ExibeMensagem "Não foi possivel efetuar baixa manual devido às inconsistências das parcelas canceladas."
                        Exit Function
                    End If
                    adoResultado.MoveNext
                Loop
            End If
        End If
    End If
    
    GravaBaixaMaunual = True
    
End Function

Private Sub BaixaMovimentosDebitoAutomatico()
Dim strLinha               As String
Dim lngLinha               As Long
Dim lngSize                As Long
Dim lngParcelas            As Long
Dim lngParcelasJaBaixadas  As Long
Dim dtmDataMovimento       As Date
Dim strInscricaoSemParcela As Variant
Dim strInscricaoParcelaPaga As Variant

Dim blnTipoBJaSomado       As Boolean
Dim blnTipoFJaSomado       As Boolean
Dim strSQL                 As String

Dim adoResultado           As New ADODB.Recordset
Dim adoConsulta            As New ADODB.Recordset

Dim bytResposta            As Byte
Dim strMes                 As String
Dim strBanco               As String
Dim bytQtdeTipos           As Byte ' Tipos (A,B,F,Z)

    On Error GoTo err_BaixaAutomatica
    
    bytQtdeTipos = 0
    lngLinha = 0
    blnTipoBJaSomado = False
    blnTipoFJaSomado = False
    
    Open txt_Arquivo For Input As #1
    
    lngSize = FileLen(txt_Arquivo) / 150
    
    pgr_Status.Value = 0
    pgr_Status.Visible = True
    pgr_Status.Max = Abs(lngSize)
    lbl_Status.Visible = True
    
    Set gobjBanco = New clsBanco
    
    gobjBanco.ExecutaBeginTrans
    
    lngParcelas = 0
    lngParcelasJaBaixadas = 0
    strInscricaoSemParcela = ""
    
    Do While Not EOF(1)

        Line Input #1, strLinha

        If Len(strLinha) = 0 Then
            GoTo ProximaLinha
        End If
        
        'Vamos verificar se a linha contem 150 posicoes
        If Len(strLinha) <> 150 Then
            ExibeMensagem "Ocorreu um Erro na leitura do arquivo. A operação não foi concluída."
            Screen.MousePointer = vbDefault
            pgr_Status.Visible = False
            lbl_Status.Visible = False
            gobjBanco.ExecutaRollbackTrans
            Close #1
            Exit Sub
        End If
        
        If Mid(strLinha, 1, 1) = "A" Then
            
            bytQtdeTipos = bytQtdeTipos + 1
            dtmDataMovimento = ConverteDataDoArquivo(Mid(strLinha, 66, 8), True)
            strBanco = Trim(Mid(strLinha, 46, 20))
            
            bytResposta = MsgBox("O Mês referente ao movimento é " & Month(dtmDataMovimento) & " (" & gstrNomeDoMes(Month(dtmDataMovimento)) & "). Deseja especificar outro?", vbYesNo)
            If bytResposta = vbYes Then
DigitarNovamente:
                strMes = InputBox("Digite o Mês desejado de 1 a 12.", "Alteração de Mês do movimento")
                If Not IsNumeric(strMes) Then
                    ExibeMensagem "Mês inválido. Digite novamente."
                    GoTo DigitarNovamente
                ElseIf Val(strMes) < 1 Or Val(strMes) > 12 Then
                    ExibeMensagem "Mês inválido. Digite novamente."
                    GoTo DigitarNovamente
                End If
            Else
                strMes = Month(dtmDataMovimento)
            End If
            
        ElseIf Mid(strLinha, 1, 1) = "B" Then
            
            If Not blnTipoBJaSomado Then
                bytQtdeTipos = bytQtdeTipos + 1
                blnTipoBJaSomado = True
            End If
            
        ElseIf Mid(strLinha, 1, 1) = "F" Then
        
            If Not blnTipoFJaSomado Then
                bytQtdeTipos = bytQtdeTipos + 1
                blnTipoFJaSomado = True
            End If
            
            lngParcelas = lngParcelas + 1
            
            'Vamos obter a parcela a ser baixada
            strSQL = "SELECT LV.Pkid, LV.dblValor FROM " & gstrLancamentoAlfa & " LA, " & gstrLancamentoValor & " LV WHERE LA.strInscricao = '" & Format(Mid(strLinha, 2, 11), "00000000000000000000") & "' AND LA.intExercicio = 2005 AND LA.intComposicaoDaReceita in (41,42) AND LA.Pkid = LV.intLancamentoAlfa AND MONTH(LV.dtmDtVencimento) = " & strMes
            If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
                If adoResultado.EOF Then
                    strInscricaoSemParcela = strInscricaoSemParcela & gstrFormataInscricao(Mid(strLinha, 2, 11), TYP_IMOBILIARIA) & Chr(13) & Chr(10)
                    lngParcelasJaBaixadas = lngParcelasJaBaixadas + 1
                Else
                    'Vamos verificar se ja existe pagamento para esta parcela
                    strSQL = "SELECT LP.Pkid FROM " & gstrLancamentoPagamento & " LP WHERE LP.intLancamentoValor = " & adoResultado("Pkid").Value
                    If gobjBanco.CriaADO(strSQL, 5, adoConsulta) Then
                        If Not adoConsulta.EOF Then
                            strInscricaoParcelaPaga = strInscricaoParcelaPaga & gstrFormataInscricao(Mid(strLinha, 2, 11), TYP_IMOBILIARIA) & Chr(13) & Chr(10)
                            lngParcelasJaBaixadas = lngParcelasJaBaixadas + 1
                        End If
                        strSQL = "INSERT INTO " & gstrLancamentoPagamento & " (intLancamentoValor, dblValorPrincipal, dblValorMulta, dblValorJuros, dblValorCorrecao, dblValorCorreto, dtmDtPagamento, dtmDtMovimento, intCodigoBaixa, strObservacao, dtmDtAtualizacao, lngCodUsr) " & _
                                 " VALUES (" & adoResultado("Pkid").Value & ", " & gstrConvVrParaSql(Mid(strLinha, 53, 15) / 100) & ",0,0,0,0," & gstrConvDtParaSql(ConverteDataDoArquivo(Mid(strLinha, 45, 8), True)) & ", " & gstrConvDtParaSql(Day(dtmDataMovimento) & "/" & strMes & "/" & Year(dtmDataMovimento)) & ",11,'Débito Automático'," & gstrConvDtParaSql(gstrDataDoSistema) & ",1)"
                        gobjBanco.Execute strSQL
                    End If
                    
                End If
            End If
                        
             
        ElseIf Mid(strLinha, 1, 1) = "Z" Then
            bytQtdeTipos = bytQtdeTipos + 1
            
            'Vamos verificar se a qtde gravada é igual ao total de registros
            If lngLinha + 1 <> Val(Mid(strLinha, 2, 6)) Then
                ExibeMensagem "A quantidade de registros processada não coincide com a informada no arquivo. A operação não concluída."
                Screen.MousePointer = vbDefault
                pgr_Status.Visible = False
                lbl_Status.Visible = False
                gobjBanco.ExecutaRollbackTrans
                Close #1
                Exit Sub
            End If
            
        End If
        
        lngLinha = lngLinha + 1
        
        DoEvents
        lbl_Status.Caption = lngLinha & " de " & lngSize
        Me.Refresh
        
        pgr_Status.Value = lngLinha
        
ProximaLinha:

    Loop
        
    Close #1
        
    pgr_Status.Visible = False
    lbl_Status.Visible = False
    
    If Len(strInscricaoSemParcela) > 0 Then
        ExibeMensagem "Não foram encontradas parcelas para as Inscrições: " & strInscricaoSemParcela
        Screen.MousePointer = vbDefault
        Close #1
    End If
    
    'Vamos verificar se todas as parcelas do movimento ja estao baixadas
    If lngParcelas = lngParcelasJaBaixadas Then
        ExibeMensagem "Todas as parcelas deste movimento já se encontram baixadas ou não encontradas."
        Screen.MousePointer = vbDefault
        gobjBanco.ExecutaRollbackTrans
        Exit Sub
    End If
    
    If Len(strInscricaoParcelaPaga) > 0 Then
        ExibeMensagem "Foram encontradas parcelas pagas para as Inscrições: " & strInscricaoParcelaPaga
        Screen.MousePointer = vbDefault
        Close #1
    End If
    
    'Vamos gerar arquivo de criticas
    If Len(strInscricaoSemParcela) > 0 Or Len(strInscricaoParcelaPaga) > 0 Then
        Open "C:\Criticas" & Replace(strBanco, "/", "") & Replace(dtmDataMovimento, "/", "") & ".txt" For Output As #1
        If Len(strInscricaoSemParcela) > 0 Then
            Print #1, "Não foram encontradas parcelas para as Inscrições: " & Chr(13) & Chr(10) & strInscricaoSemParcela
        End If
        If Len(strInscricaoParcelaPaga) > 0 Then
            Print #1, "Foram encontradas parcelas pagas para as Inscrições: " & Chr(13) & Chr(10) & strInscricaoParcelaPaga
        End If
        Close #1
    End If
    
    'Vamos verificar se foram lidos todos os tipos de registro
    If bytQtdeTipos < 3 Then
        ExibeMensagem "O arquivo de leitura está incorreto."
        Screen.MousePointer = vbDefault
        gobjBanco.ExecutaRollbackTrans
        Close #1
        Exit Sub
    End If
    
    gobjBanco.ExecutaCommitTrans
    
    Screen.MousePointer = vbDefault
    
    Exit Sub

err_BaixaAutomatica:
    ExibeMensagem "Ocorreu um Erro na leitura do arquivo. A operação não foi concluída."
    Screen.MousePointer = vbDefault
    pgr_Status.Visible = False
    lbl_Status.Visible = False
    gobjBanco.ExecutaRollbackTrans
    Close #1
    
End Sub

Private Sub BaixaMovimentoBancarioFichaCompensacaoGREM()
Dim strLinha               As String
Dim lngLinha               As Long
Dim lngSize                As Long

Dim lngParcelas             As Long
Dim lngParcelasJaBaixadas   As Long
Dim strInscricaoSemParcela  As String
Dim strInscricaoParcelaPaga As String
Dim dtmDataMovimento        As Date
Dim lngPkidContaBancariaMov As Long

Dim blnTipoJaSomado        As Boolean
Dim strSQL                 As String

Dim adoResultado           As New ADODB.Recordset
Dim adoConsulta            As New ADODB.Recordset

Dim bytQtdeTipos           As Byte ' Tipos (0,1,9)
Dim dblValorDoArquivo      As Double
Dim dblValorTarifa         As Double

    On Error GoTo err_BaixaAutomatica
    
    bytQtdeTipos = 0
    lngLinha = 0
    blnTipoJaSomado = False
    
    lngParcelas = 0
    lngParcelasJaBaixadas = 0
    strInscricaoSemParcela = ""
    
    Open txt_Arquivo For Input As #1
    
    lngSize = FileLen(txt_Arquivo) / 120
    
    pgr_Status.Value = 0
    pgr_Status.Visible = True
    pgr_Status.Max = Abs(lngSize)
    lbl_Status.Visible = True
    
    Set gobjBanco = New clsBanco
    
    gobjBanco.ExecutaBeginTrans
    
    Do While Not EOF(1)

        Line Input #1, strLinha

        If Len(strLinha) = 0 Then
            GoTo ProximaLinha
        End If

        'Vamos verificar se a linha contem 120 posicoes
        If Len(strLinha) <> 120 Then
            ExibeMensagem "Ocorreu um Erro na leitura do arquivo. A operação não foi concluída."
            Screen.MousePointer = vbDefault
            pgr_Status.Visible = False
            lbl_Status.Visible = False
            gobjBanco.ExecutaRollbackTrans
            Close #1
            Exit Sub
        End If

        If Mid(strLinha, 1, 1) = "0" Then
            bytQtdeTipos = bytQtdeTipos + 1
            dtmDataMovimento = Mid(strLinha, 67, 2) & "/" & Mid(strLinha, 69, 2) & "/" & Mid(strLinha, 71, 2)
            
            'Vamos buscar o Pkid referente à Conta em tblContaBancaria
            strSQL = "SELECT Pkid FROM " & gstrContaBancaria & " WHERE strConta = '" & Trim(Mid(strLinha, 23, 7)) & "'"
            If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
                If adoResultado.EOF Then
                    ExibeMensagem "Não foi encontrada Conta do movimento. A operação não concluída."
                    Screen.MousePointer = vbDefault
                    pgr_Status.Visible = False
                    lbl_Status.Visible = False
                    gobjBanco.ExecutaRollbackTrans
                    Close #1
                    Exit Sub
                Else
                    lngPkidContaBancariaMov = adoResultado("Pkid").Value
                End If
            End If

        ElseIf Mid(strLinha, 1, 1) = "1" Then
        
            lngParcelas = lngParcelas + 1
            
            If Not blnTipoJaSomado Then
                bytQtdeTipos = bytQtdeTipos + 1
                blnTipoJaSomado = True
            End If
            
            dblValorDoArquivo = Mid(strLinha, 76, 13) / 100
            dblValorTarifa = Mid(strLinha, 37, 13) / 100
            
            'Vamos verificar se é guia gerada pelo winpublic
            If Mid(strLinha, 16, 2) <> "00" Then
                'Vamos obter a parcela a ser baixada
                strSQL = "SELECT LV.Pkid, LV.dblValor FROM " & gstrLancamentoAlfa & " LA, " & gstrLancamentoValor & " LV WHERE LA.strInscricaoAuxiliar = '" & Mid(strLinha, 16, 7) & "' AND LA.intExercicio = 2005 AND LA.intComposicaoDaReceita = 39 AND LA.Pkid = LV.intLancamentoAlfa "
                If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
                    If adoResultado.EOF Then
                        strInscricaoSemParcela = strInscricaoSemParcela & gstrFormataInscricao(Format(Mid(strLinha, 18, 5) & "2005", "0000000000"), TYP_PRECO_PUBLICO) & Chr(13) & Chr(10)
                        lngParcelasJaBaixadas = lngParcelasJaBaixadas + 1
                    Else
                        'Vamos verificar se ja existe pagamento para esta parcela
                        strSQL = "SELECT LP.Pkid FROM " & gstrLancamentoPagamento & " LP WHERE LP.intLancamentoValor = " & adoResultado("Pkid").Value
                        If gobjBanco.CriaADO(strSQL, 5, adoConsulta) Then
                            If Not adoConsulta.EOF Then
                                strInscricaoParcelaPaga = strInscricaoParcelaPaga & gstrFormataInscricao(Mid(strLinha, 18, 5) & "2005", TYP_PRECO_PUBLICO) & Chr(13) & Chr(10)
                                lngParcelasJaBaixadas = lngParcelasJaBaixadas + 1
                            End If
                            If dblValorDoArquivo + dblValorTarifa <> adoResultado("dblvalor").Value Then
                                MsgBox "Guia: " & Mid(strLinha, 16, 7) & " Valor Guia:" & adoResultado("dblvalor").Value & "  Valor Pago:" & dblValorDoArquivo + dblValorTarifa & Chr(13) & Chr(10)
                            End If
                            strSQL = "INSERT INTO " & gstrLancamentoPagamento & " (intLancamentoValor, dblValorPrincipal, dblValorMulta, dblValorJuros, dblValorCorrecao, dblValorCorreto, dtmDtPagamento, dtmDtMovimento, intCodigoBaixa, strObservacao, dtmDtAtualizacao, lngCodUsr) " & _
                                     " VALUES (" & adoResultado("Pkid").Value & ", " & gstrConvVrParaSql(dblValorDoArquivo + dblValorTarifa) & ",0,0,0,0," & gstrConvDtParaSql(Mid(strLinha, 26, 2) & "/" & Mid(strLinha, 28, 2) & "/" & Mid(strLinha, 30, 2)) & ", " & gstrConvDtParaSql(dtmDataMovimento) & ",11,'Baixa Automatica GREM'," & gstrConvDtParaSql(gstrDataDoSistema) & ",1)"
                            gobjBanco.Execute strSQL
                        End If
                    End If
                End If
            Else
                'Vamos obter a parcela a ser baixada
                strSQL = "SELECT LG.intLancamentoValor, LG.dblValorPrincipal FROM " & gstrGuias & " G, " & gstrLancamentoGuias & " LG WHERE G.intNumero = " & Mid(strLinha, 16, 7) & " AND G.intContaBancaria = " & lngPkidContaBancariaMov & " AND G.Pkid = LG.intGuias "
                If gobjBanco.CriaADO(strSQL, 5, adoResultado) Then
                    If adoResultado.EOF Then
                        strInscricaoSemParcela = strInscricaoSemParcela & gstrFormataInscricao(Format(Mid(strLinha, 18, 5) & "2005", "0000000000"), TYP_PRECO_PUBLICO) & Chr(13) & Chr(10)
                        lngParcelasJaBaixadas = lngParcelasJaBaixadas + 1
                    Else
                        'Vamos verificar se ja existe pagamento para esta parcela
                        strSQL = "SELECT LP.Pkid FROM " & gstrLancamentoPagamento & " LP WHERE LP.intLancamentoValor = " & adoResultado("intLancamentoValor").Value
                        If gobjBanco.CriaADO(strSQL, 5, adoConsulta) Then
                            If Not adoConsulta.EOF Then
                                strInscricaoParcelaPaga = strInscricaoParcelaPaga & gstrFormataInscricao(Mid(strLinha, 18, 5) & "2005", TYP_PRECO_PUBLICO) & Chr(13) & Chr(10)
                                lngParcelasJaBaixadas = lngParcelasJaBaixadas + 1
                            End If
                            If dblValorDoArquivo + dblValorTarifa <> adoResultado("dblValorPrincipal").Value Then
                                MsgBox "Valor Guia:" & adoResultado("dblvalorPrincipal").Value & "  Valor Pago:" & dblValorDoArquivo + dblValorTarifa
                            End If
                            strSQL = "INSERT INTO " & gstrLancamentoPagamento & " (intLancamentoValor, dblValorPrincipal, dblValorMulta, dblValorJuros, dblValorCorrecao, dblValorCorreto, dtmDtPagamento, dtmDtMovimento, intCodigoBaixa, strObservacao, dtmDtAtualizacao, lngCodUsr) " & _
                                     " VALUES (" & adoResultado("intLancamentoValor").Value & ", " & gstrConvVrParaSql(dblValorDoArquivo + dblValorTarifa) & ",0,0,0,0," & gstrConvDtParaSql(Mid(strLinha, 26, 2) & "/" & Mid(strLinha, 28, 2) & "/" & Mid(strLinha, 30, 2)) & ", " & gstrConvDtParaSql(dtmDataMovimento) & ",11,'Baixa Automatica GREM'," & gstrConvDtParaSql(gstrDataDoSistema) & ",1)"
                            gobjBanco.Execute strSQL
                        End If
                    End If
                End If
            End If
            
        ElseIf Mid(strLinha, 1, 1) = "9" Then
            bytQtdeTipos = bytQtdeTipos + 1

            'Vamos verificar se a qtde gravada é igual ao total de registros
            If lngLinha - 1 <> Val(Mid(strLinha, 11, 6)) Then
                ExibeMensagem "A quantidade de registros processada não coincide com a informada no arquivo. A operação não concluída."
                Screen.MousePointer = vbDefault
                pgr_Status.Visible = False
                lbl_Status.Visible = False
                gobjBanco.ExecutaRollbackTrans
                Close #1
                Exit Sub
            End If

        End If

        lngLinha = lngLinha + 1

        DoEvents
        lbl_Status.Caption = lngLinha & " de " & lngSize
        Me.Refresh

        pgr_Status.Value = lngLinha

ProximaLinha:

    Loop
    Close #1
    
    pgr_Status.Visible = False
    lbl_Status.Visible = False
        
    'Vamos verificar se foram lidos todos os tipos de registro
    If bytQtdeTipos < 3 Then
        ExibeMensagem "O arquivo de leitura está incorreto."
        Screen.MousePointer = vbDefault
        gobjBanco.ExecutaRollbackTrans
        Close #1
        Exit Sub
    End If
    
    If Len(strInscricaoSemParcela) > 0 Then
        ExibeMensagem "Não foram encontradas parcelas para as Inscrições: " & strInscricaoSemParcela
        Screen.MousePointer = vbDefault
        Close #1
    End If
    
    'Vamos verificar se todas as parcelas do movimento ja estao baixadas
    If lngParcelas = lngParcelasJaBaixadas Then
        ExibeMensagem "Todas as parcelas deste movimento já se encontram baixadas ou não encontradas."
        Screen.MousePointer = vbDefault
        gobjBanco.ExecutaRollbackTrans
        Exit Sub
    End If
    
    If Len(strInscricaoParcelaPaga) > 0 Then
        ExibeMensagem "Foram encontradas parcelas pagas para as Inscrições: " & strInscricaoParcelaPaga
        Screen.MousePointer = vbDefault
        Close #1
    End If
    
    'Vamos gerar arquivo de criticas
    If Len(strInscricaoSemParcela) > 0 Or Len(strInscricaoParcelaPaga) > 0 Then
        Open "C:\CriticasGREM" & Replace(dtmDataMovimento, "/", "") & ".txt" For Output As #1
        If Len(strInscricaoSemParcela) > 0 Then
            Print #1, "Não foram encontradas parcelas para as Inscrições: " & Chr(13) & Chr(10) & strInscricaoSemParcela
        End If
        If Len(strInscricaoParcelaPaga) > 0 Then
            Print #1, "Foram encontradas parcelas pagas para as Inscrições: " & Chr(13) & Chr(10) & strInscricaoParcelaPaga
        End If
        Close #1
    End If
    
    'Vamos verificar se foram lidos todos os tipos de registro
    If bytQtdeTipos < 3 Then
        ExibeMensagem "O arquivo de leitura está incorreto."
        Screen.MousePointer = vbDefault
        gobjBanco.ExecutaRollbackTrans
        Close #1
        Exit Sub
    End If
    
    gobjBanco.ExecutaCommitTrans
    
    Screen.MousePointer = vbDefault
    
    Exit Sub

err_BaixaAutomatica:
    ExibeMensagem "Ocorreu um Erro na leitura do arquivo. A operação não foi concluída."
    Screen.MousePointer = vbDefault
    pgr_Status.Visible = False
    lbl_Status.Visible = False
    gobjBanco.ExecutaRollbackTrans
    Close #1
    
End Sub

Private Function blnDadosOkDebitoAutomatico() As Boolean

    On Error GoTo err_blnDadosOK
    
    If Trim(txt_Arquivo) = "" Then
        ExibeMensagem "Indique a localização do arquivo de retorno."
        Exit Function
    ElseIf Dir(txt_Arquivo) = "" Then
        ExibeMensagem "Arquivo não encontrado no local especificado."
        Exit Function
    'ElseIf Not gblnDataValida(txt_DtMovimento) Then
    '    ExibeMensagem "A Data deve ser preenchida corretamente."
    '    Exit Function
    End If
    
    blnDadosOkDebitoAutomatico = True
    Exit Function
    
err_blnDadosOK:
    blnDadosOkDebitoAutomatico = False

End Function

Private Function ConverteDataDoArquivo(strData As String, blnLayoutNovo As Boolean) As Date
Dim strAux As String
    If blnLayoutNovo Then
        strAux = Right(strData, 2) & "/" & Mid(strData, 5, 2) & "/" & Left(strData, 4)
    Else
        strAux = Right(strData, 2) & "/" & Mid(strData, 3, 2) & "/" & Left(strData, 2)
    End If
    
    ConverteDataDoArquivo = CDate(strAux)
    
End Function

Private Sub ImportaAcordos()
Dim strLinha As Variant
Dim strSQL   As String
Dim lngSize  As Long
Dim lngLinha As Long

    Set gobjBanco = New clsBanco
    
    Open txt_Arquivo For Input As #1
    
    lngSize = FileLen(txt_Arquivo) / 310
    lbl_Status.Visible = True
    Do While Not EOF(1)

        Line Input #1, strLinha
        
        strLinha = Replace(strLinha, "'", " ")
        
        'Vamos pular o cabeçalho
        If lngLinha > 1545130 Then
            'Acordos.txt
            'strSql = "INSERT INTO admgrjorigem.dbo.acordos(ACORDO,NUM_ACORDO,ANO_ACORDO,SACADO,CNPJ_CPF,RG,END_SACADO,END_NUM_SACADO,END_CEP_SACADO,END_BAIRRO_SACADO,END_CIDADE_SACADO,END_UF_SACADO,DDD1,TEL1,DDD2,TEL2,OCUP_SACADO,DT_CANCELAMENTO,GERADO,ACORDO_PM,LANCAMENTOALFA,NUM_PARCELAS,DT_CADASTRO,VL_PRINCIPAL,VL_DESCONTO,VL_SUCUMBENCIA,VL_CUSTAS,VL_MULTA,VL_JUROS,VL_PARCELA,VL_CORRECAO,VL_HONORARIOS,VL_TOTAL,PRINC_TITULO,SUC_TITULO,MULTA_TITULO,JUROS_TITULO,CORR_TITULO,TOTAL_TITULO,PRINC_CREDITO,CORR_CREDITO,TOT_CREDITO,STATUS,OBS) " & _
            '         " VALUES (" & Mid(strLinha, 1, 12) & ",'" & Mid(strLinha, 13, 10) & "'," & Val(Mid(strLinha, 24, 4)) & ",'" & Replace(Mid(strLinha, 64, 81), "'", " ") & "','" & Mid(strLinha, 145, 16) & "','" & Replace(Mid(strLinha, 161, 21), "'", " ") & "','" & Replace(Mid(strLinha, 182, 30), "'", " ") & "','" & Mid(strLinha, 213, 50) & "'," & Val(Mid(strLinha, 263, 11)) & ",'" & Replace(Mid(strLinha, 274, 41), "'", " ") & "','" & Mid(strLinha, 315, 51) & "','" & Replace(Mid(strLinha, 366, 2), "'", "S") & "','" & Mid(strLinha, 376, 10) & "','" & Mid(strLinha, 386, 21) & "','" & Mid(strLinha, 407, 10) & "','" & Mid(strLinha, 417, 21) & "','" & Mid(strLinha, 438, 41) & "','" & Mid(strLinha, 479, 18) & "','" & Mid(strLinha, 730, 12) & "','" & Mid(strLinha, 742, 12) & "','" & Mid(strLinha, 754, 12) & "','" & Mid(strLinha, 766, 12) & "','" & Mid(strLinha, 778, 23) & "','" & Mid(strLinha, 833, 13) & "','" & Mid(strLinha, 846, 13) & "','" & Mid(strLinha, 859, 15) & "','" & Mid(strLinha, 874, 13) & "','" & _
            '         Mid(strLinha, 887, 13) & "','" & Mid(strLinha, 900, 13) & "','" & Mid(strLinha, 913, 13) & "','" & Mid(strLinha, 926, 13) & "','" & Mid(strLinha, 939, 14) & "','" & Mid(strLinha, 953, 13) & "','" & Mid(strLinha, 966, 54) & "','" & Mid(strLinha, 1020, 54) & "','" & Mid(strLinha, 1074, 54) & "','" & Mid(strLinha, 1128, 54) & "','" & Mid(strLinha, 1182, 54) & "','" & Mid(strLinha, 1236, 54) & "','" & Mid(strLinha, 1290, 54) & "','" & Mid(strLinha, 1345, 53) & "','" & Mid(strLinha, 1398, 54) & "','" & Mid(strLinha, 1452, 21) & "','" & Mid(strLinha, 1473, 10) & "')"
            'Acordos_Titulos.txt
            'strSql = "INSERT INTO admgrjorigem.dbo.acordos_titulos(ACORDO,TITULO,TIPO_CALCULO,DT_ATUALIZACAO,CONTRIBUINTE,VL_PRINCIPAL,VL_SUCUMBENCIA,VL_MULTA,VL_JUROS,VL_CORRECAO,VL_TOTAL,VL_COTACAO) " & _
            '         " VALUES (" & Mid(strLinha, 1, 12) & "," & Mid(strLinha, 13, 12) & ",'" & Mid(strLinha, 25, 13) & "','" & Mid(strLinha, 38, 23) & "','" & Mid(strLinha, 93, 13) & "','" & Mid(strLinha, 106, 18) & "','" & Mid(strLinha, 160, 18) & "','" & Mid(strLinha, 214, 18) & "','" & Mid(strLinha, 268, 18) & "','" & Mid(strLinha, 322, 18) & "','" & Mid(strLinha, 376, 18) & "','" & Mid(strLinha, 430, 18) & "')"
            'Imoveis.txt
            'strSql = "INSERT INTO admgrjorigem.dbo.imoveis(COD_IMOVEL,TIPO_CONSTRUCAO,TIPO_IMOVEL,END_IMOVEL,END_COMPLEMENTO,END_CEP,END_BAIRRO,END_CIDADE,END_UF,QUADRA,DESC_LOTE,LOTE,SUBLOTE,INSCRICAO,SETOR,SUBQUADRA,LOTE_LOTE,SUB_SUBLOTE,LANCAMENTOALFA,TIPO_DIVIDA) " & _
            '         " VALUES (" & Mid(strLinha, 1, 12) & "," & IIf(Len(Trim(Mid(strLinha, 13, 16))) = 0, "NULL", Mid(strLinha, 13, 16)) & "," & IIf(Len(Trim(Mid(strLinha, 29, 12))) = 0, "NULL", Mid(strLinha, 29, 12)) & ",'" & Mid(strLinha, 41, 81) & "','" & Mid(strLinha, 122, 41) & "'," & IIf(Len(Trim(Mid(strLinha, 163, 11))) = 0, "NULL", Mid(strLinha, 163, 11)) & ",'" & Mid(strLinha, 174, 41) & "','" & Mid(strLinha, 215, 51) & "','" & Mid(strLinha, 266, 2) & "'," & IIf(Len(Trim(Mid(strLinha, 276, 14))) = 0, "NULL", Mid(strLinha, 276, 14)) & ",'" & Mid(strLinha, 290, 21) & "'," & IIf(Len(Trim(Mid(strLinha, 311, 21))) = 0, "NULL", Mid(strLinha, 311, 21)) & "," & IIf(Len(Trim(Mid(strLinha, 332, 15))) = 0, "NULL", Mid(strLinha, 332, 15)) & "," & IIf(Len(Trim(Mid(strLinha, 347, 12))) = 0, "NULL", Mid(strLinha, 347, 12)) & "," & IIf(Len(Trim(Mid(strLinha, 359, 13))) = 0, "NULL", Mid(strLinha, 359, 13)) & "," & IIf(Len(Trim(Mid(strLinha, 372, 17))) = 0, "NULL", Mid(strLinha, 372, 17)) & "," & _
            '         IIf(Len(Trim(Mid(strLinha, 389, 17))) = 0, "NULL", Mid(strLinha, 389, 17)) & "," & IIf(Len(Trim(Mid(strLinha, 406, 20))) = 0, "NULL", Mid(strLinha, 406, 20)) & "," & IIf(Len(Trim(Mid(strLinha, 426, 12))) = 0, "NULL", Mid(strLinha, 426, 12)) & "," & IIf(Len(Trim(Mid(strLinha, 438, 7))) = 0, "NULL", Mid(strLinha, 438, 7)) & ")"
            'Proprietarios_Imoveis.txt
            'strSql = "INSERT INTO admgrjorigem.dbo.proprietarios_imoveis(COD_IMOVEL,TIPO_PROP,CONTRIBUINTE,NOME,TIPO,CPF,RG,LANCAMENTOALFA) " & _
            '         " VALUES (" & Mid(strLinha, 1, 21) & "," & IIf(Len(Trim(Mid(strLinha, 22, 12))) = 0, "NULL", Mid(strLinha, 22, 12)) & "," & IIf(Len(Trim(Mid(strLinha, 34, 13))) = 0, "NULL", Mid(strLinha, 34, 13)) & ",'" & Mid(strLinha, 47, 81) & "'," & IIf(Len(Trim(Mid(strLinha, 128, 17))) = 0, "NULL", Mid(strLinha, 128, 17)) & ",'" & Mid(strLinha, 146, 17) & "','" & Mid(strLinha, 163, 16) & "'," & IIf(Len(Trim(Mid(strLinha, 179, 17))) = 0, "NULL", Mid(strLinha, 179, 17)) & ")"
            'Tipo_Divida.txt
            'strSql = "INSERT INTO admgrjorigem.dbo.tipo_divida(TIPO_DIVIDA,NOME,SIGLA) " & _
            '         " VALUES (" & Mid(strLinha, 1, 12) & ",'" & Mid(strLinha, 13, 201) & "','" & Mid(strLinha, 214, 26) & "')"
            'Acordo_Parcelas.txt
            'strSql = "INSERT INTO admgrjorigem.dbo.acordos_parcelas(ACORDO,BANCO,DT_CAD_PARCELA,NU_PARCELA,NU_PARCELA_ANISTIA,NU_TOT_PARCELA,VL_PARCELA,VL_PAGO_PARCELA,VL_JUROS_PARCELA,VL_MULTA_PARCELA,VL_DESC_PARCELA,VL_SUC_PARCELA,VL_CORR_PARCELA,BOLETO,DT_EMISSAO,DT_PARCELA,DT_PAGO,DT_CANC,TIPO_PARCELA,STATUS,SEQ_PARCELA,VL_PRINC_ACORDO_PARC,VL_SUC_ACORDO_PARC,VL_MULTA_ACORDO_PARC,VL_JUROS_ACORDO_PARC,VL_CORR_ACORDO_PARC,VL_TOT_ACORDO_PARC) " & _
            '         " VALUES (" & Mid(strLinha, 1, 12) & "," & IIf(Len(Trim(Mid(strLinha, 13, 12))) = 0, "NULL", Mid(strLinha, 13, 12)) & ",'" & Trim(Mid(strLinha, 25, 55)) & "'," & IIf(Len(Trim(Mid(strLinha, 80, 12))) = 0, "NULL", Mid(strLinha, 80, 12)) & "," & IIf(Len(Trim(Mid(strLinha, 92, 19))) = 0, "NULL", Mid(strLinha, 92, 19)) & "," & IIf(Len(Trim(Mid(strLinha, 111, 15))) = 0, "NULL", Mid(strLinha, 111, 15)) & ",'" & Mid(strLinha, 126, 13) & "','" & Mid(strLinha, 139, 16) & "','" & Mid(strLinha, 155, 17) & "','" & Mid(strLinha, 172, 17) & "','" & Mid(strLinha, 189, 16) & "','" & Mid(strLinha, 205, 15) & "','" & Mid(strLinha, 220, 20) & "'," & IIf(Len(Trim(Mid(strLinha, 241, 21))) = 0, "NULL", Mid(strLinha, 241, 21)) & ",'" & _
            '         IIf(Trim(Mid(strLinha, 262, 55)) = "NULL", "", Trim(Mid(strLinha, 262, 55))) & "','" & Trim(Mid(strLinha, 317, 51)) & "','" & IIf(Trim(Mid(strLinha, 368, 55)) = "NULL", "", Trim(Mid(strLinha, 368, 55))) & "','" & Mid(strLinha, 423, 55) & "','" & Trim(Mid(strLinha, 478, 21)) & "','" & Mid(strLinha, 499, 21) & "'," & IIf(Len(Trim(Mid(strLinha, 520, 12))) = 0, "NULL", Mid(strLinha, 520, 12)) & ",'" & Trim(Mid(strLinha, 532, 54)) & "','" & Trim(Mid(strLinha, 586, 54)) & "','" & Trim(Mid(strLinha, 640, 54)) & "','" & Trim(Mid(strLinha, 694, 54)) & "','" & Trim(Mid(strLinha, 748, 54)) & "','" & Trim(Mid(strLinha, 802, 54)) & "')"
            'Titulos.txt
            'strSql = "INSERT INTO admgrjorigem.dbo.titulos(TITULO,CONTRIBUINTE,TIPO_DIVIDA,ACORDO,MOEDA,IMOVEL,CERTIDAO,LIVRO,FOLHA,PROCESSO,VARA,EXERC_TITULO,VL_TITULO,VL_BASE_TITULO,VL_ORIG_TITULO,VL_CORR_TITULO,VL_JUROS_TITULO,VL_MULTA_TITULO,VL_CUSTAS_TITULO,VL_SUC_TITULO,VL_HONOR_TITULO,VL_COR_TITULO,VL_CORR_AT_TITULO,VL_JUROS_AT_TITULO,VL_MULTA_AT_TITULO,VL_CUSTAS_AT_TITULO,VL_SUC_AT_TITULO,DT_TITULO,NU_TITULO,OBS,LANCAMENTOALFA,DT_INSCRICAO,DT_BAIXA,VL_UFM,NU_EXERC_PARCELA) " & _
            '         " VALUES (" & Mid(strLinha, 1, 12) & "," & IIf(Len(Trim(Mid(strLinha, 13, 13))) = 0, "NULL", Mid(strLinha, 13, 13)) & "," & IIf(Len(Trim(Mid(strLinha, 26, 12))) = 0, "NULL", Mid(strLinha, 26, 12)) & "," & IIf(Len(Trim(Mid(strLinha, 38, 12))) = 0, "NULL", Mid(strLinha, 38, 12)) & "," & IIf(Len(Trim(Mid(strLinha, 50, 12))) = 0, "NULL", Mid(strLinha, 50, 12)) & "," & IIf(Len(Trim(Mid(strLinha, 62, 21))) = 0, "NULL", Mid(strLinha, 62, 21)) & "," & IIf(Len(Trim(Mid(strLinha, 83, 21))) = 0, "NULL", Mid(strLinha, 83, 21)) & "," & IIf(Len(Trim(Mid(strLinha, 104, 21))) = 0, "NULL", Mid(strLinha, 104, 21)) & "," & IIf(Len(Trim(Mid(strLinha, 125, 21))) = 0, "NULL", Mid(strLinha, 125, 21)) & ",'" & Trim(Mid(strLinha, 146, 21)) & "'," & IIf(Len(Trim(Mid(strLinha, 167, 12))) = 0, "NULL", Mid(strLinha, 167, 12)) & "," & _
            '         IIf(Len(Trim(Mid(strLinha, 179, 13))) = 0, "NULL", Mid(strLinha, 179, 13)) & ",'" & Trim(Mid(strLinha, 192, 15)) & "','" & Trim(Mid(strLinha, 207, 15)) & "','" & Mid(strLinha, 222, 15) & "','" & Trim(Mid(strLinha, 237, 15)) & "','" & Trim(Mid(strLinha, 252, 15)) & "','" & Trim(Mid(strLinha, 268, 21)) & "','" & Trim(Mid(strLinha, 289, 17)) & "','" & Trim(Mid(strLinha, 306, 21)) & "','" & Trim(Mid(strLinha, 327, 16)) & "','" & Trim(Mid(strLinha, 343, 15)) & "','" & Trim(Mid(strLinha, 358, 18)) & "','" & Trim(Mid(strLinha, 376, 19)) & "','" & Trim(Mid(strLinha, 395, 19)) & "','" & Trim(Mid(strLinha, 414, 21)) & "','" & Trim(Mid(strLinha, 435, 21)) & "','" & IIf(Trim(Mid(strLinha, 456, 23)) = "NULL", "", Trim(Mid(strLinha, 456, 23))) & "','" & Trim(Mid(strLinha, 511, 11)) & "','" & Trim(Mid(strLinha, 522, 257)) & "'," & IIf(Len(Trim(Mid(strLinha, 779, 12))) = 0, "NULL", Mid(strLinha, 779, 12)) & ",'" & IIf(Trim(Mid(strLinha, 791, 23)) = "NULL", "", Trim(Mid(strLinha, 791, 23))) & "','" & _
            '         IIf(Trim(Mid(strLinha, 846, 23)) = "NULL", "", Trim(Mid(strLinha, 846, 23))) & "','" & Trim(Mid(strLinha, 901, 54)) & "','" & Trim(Mid(strLinha, 955, 15)) & "')"
            'Titulos_Parcelas.txt
            strSQL = "INSERT INTO admgrjorigem.dbo.titulos_parcelas(PARCELA,TITULO,MOEDA,NU_PARCELA,VENC_PARCELA,VL_PARCELA,BOLETO,LANCAMENTOALFA,VL_CORR_PARCELA,VL_MULTA_PARCELA,VL_JUROS_PARCELA) " & _
                     " VALUES (" & Mid(strLinha, 1, 12) & "," & IIf(Len(Trim(Mid(strLinha, 13, 12))) = 0, "NULL", Mid(strLinha, 13, 12)) & "," & IIf(Len(Trim(Mid(strLinha, 25, 12))) = 0, "NULL", Mid(strLinha, 25, 12)) & "," & IIf(Len(Trim(Mid(strLinha, 37, 11))) = 0, "NULL", Mid(strLinha, 37, 11)) & ",'" & IIf(Trim(Mid(strLinha, 48, 23)) = "NULL", "", Trim(Mid(strLinha, 48, 23))) & "','" & Trim(Mid(strLinha, 103, 13)) & "','" & Trim(Mid(strLinha, 116, 21)) & "'," & _
                     IIf(Len(Trim(Mid(strLinha, 137, 12))) = 0, "NULL", Mid(strLinha, 137, 12)) & ",'" & Trim(Mid(strLinha, 149, 54)) & "','" & Trim(Mid(strLinha, 203, 54)) & "','" & Mid(strLinha, 257, 530) & "')"
            
            gobjBanco.Execute strSQL
        End If
        
        lngLinha = lngLinha + 1
        
        lbl_Status.Caption = lngLinha & " de " & lngSize
        Me.Refresh
        
    Loop
    
    Close #1
    
    lbl_Status.Visible = False
    
End Sub
