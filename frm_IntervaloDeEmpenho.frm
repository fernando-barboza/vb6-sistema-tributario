VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_IntervaloDeEmpenho 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nota de Empenho"
   ClientHeight    =   3360
   ClientLeft      =   3615
   ClientTop       =   1995
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   5040
   Begin VB.Frame fraExercicio 
      Caption         =   " Exercício "
      Height          =   645
      Left            =   1650
      TabIndex        =   5
      Top             =   1410
      Width           =   1095
      Begin VB.TextBox txtintExercicio 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   180
         TabIndex        =   6
         Top             =   240
         Width           =   705
      End
   End
   Begin VB.Frame fra_Guias 
      Caption         =   " Vias "
      Height          =   825
      Left            =   180
      TabIndex        =   13
      Top             =   2040
      Width           =   4635
      Begin VB.CheckBox chk_Compras 
         Caption         =   "Compras"
         Height          =   255
         Left            =   1860
         TabIndex        =   17
         Top             =   510
         Width           =   1455
      End
      Begin VB.CheckBox chk_Almoxarifado 
         Caption         =   "Almoxarifado"
         Height          =   255
         Left            =   1860
         TabIndex        =   16
         Top             =   210
         Width           =   1425
      End
      Begin VB.CheckBox chk_Tesouraria 
         Caption         =   "Tesouraria"
         Height          =   255
         Left            =   3420
         TabIndex        =   18
         Top             =   300
         Width           =   1065
      End
      Begin VB.CheckBox chk_Processo 
         Caption         =   "Processo"
         Height          =   255
         Left            =   180
         TabIndex        =   15
         Top             =   510
         Width           =   1815
      End
      Begin VB.CheckBox chk_Fornecedor 
         Caption         =   "Fornecedor"
         Height          =   255
         Left            =   180
         TabIndex        =   14
         Top             =   210
         Width           =   1875
      End
   End
   Begin VB.Frame fra_Final 
      Caption         =   " Final"
      Height          =   855
      Left            =   2520
      TabIndex        =   10
      Top             =   540
      Width           =   2265
      Begin VB.TextBox txt_ParcFinal 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1590
         TabIndex        =   4
         Top             =   420
         Width           =   585
      End
      Begin VB.TextBox txt_EmpFinal 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   420
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Parcela"
         Height          =   165
         Left            =   1590
         TabIndex        =   12
         Top             =   210
         Width           =   645
      End
      Begin VB.Label Label4 
         Caption         =   "Empenho"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   210
         Width           =   735
      End
   End
   Begin VB.Frame fra_Inicial 
      Caption         =   " Inicial "
      Height          =   855
      Left            =   150
      TabIndex        =   7
      Top             =   540
      Width           =   2265
      Begin VB.TextBox txt_EmpInicial 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   420
         Width           =   1215
      End
      Begin VB.TextBox txt_ParcInicial 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1590
         TabIndex        =   2
         Top             =   420
         Width           =   585
      End
      Begin VB.Label Label1 
         Caption         =   "Empenho"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   210
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Parcela"
         Height          =   165
         Left            =   1590
         TabIndex        =   8
         Top             =   210
         Width           =   645
      End
   End
   Begin VB.CheckBox chk_Estorno 
      Caption         =   "Só Estorno"
      Height          =   345
      Left            =   2370
      TabIndex        =   19
      Top             =   2910
      Width           =   1215
   End
   Begin MSComctlLib.TabStrip tbs_Intervalo 
      Height          =   3225
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   5689
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Intervalo de Empenho"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm_IntervaloDeEmpenho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
   frmCadEmpenho.blnAtivaFormImprime = True
   
   If Not frmCadEmpenho.mblnRestosAPagar Then
      txtintExercicio = gintExercicio
      TrocaCorObjeto txtintExercicio, True
   End If
    LerConfiguracaoDeImpressao
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmCadEmpenho.blnAtivaFormImprime = False
End Sub

Private Sub txt_EmpInicial_GotFocus()
    MarcaCampo txt_EmpInicial
End Sub

Private Sub txt_EmpInicial_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_EmpInicial
End Sub

Private Sub txt_EmpFinal_GotFocus()
    MarcaCampo txt_EmpFinal
End Sub

Private Sub txt_EmpFinal_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_EmpFinal
End Sub
Private Sub txt_ParcInicial_GotFocus()
    MarcaCampo txt_ParcInicial
End Sub

Private Sub txt_ParcInicial_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_ParcInicial
End Sub

Private Sub txt_ParcFinal_GotFocus()
    MarcaCampo txt_ParcFinal
End Sub

Private Sub txt_ParcFinal_KeyPress(KeyAscii As Integer)
    CaracterValido KeyAscii, "N", txt_ParcFinal
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
    If UCase(strModoOperacao) = UCase(gstrImprimir) Then
        If blnDadosOk Then
            frmCadEmpenho.strEmpInicial = txt_EmpInicial
            frmCadEmpenho.strEmpFinal = txt_EmpFinal
            frmCadEmpenho.strParcInicial = txt_ParcInicial
            frmCadEmpenho.strParcFinal = txt_ParcFinal
            frmCadEmpenho.intExercicioEmpenho = txtintExercicio
            frmCadEmpenho.blnSoEstorno = chk_Estorno
            frmCadEmpenho.blnAlmoxarifado = chk_Almoxarifado
            frmCadEmpenho.blnCompras = chk_Compras
            frmCadEmpenho.blnFornecedor = chk_Fornecedor
            frmCadEmpenho.blnProcesso = chk_Processo
            frmCadEmpenho.blnTesouraria = chk_Tesouraria
            'Salvando as ultimas configurações de Impressao
            SalvarConfiguracao "Fornecedor", Me.chk_Fornecedor.Value
            SalvarConfiguracao "Almoxarifado", Me.chk_Almoxarifado.Value
            SalvarConfiguracao "Processo", Me.chk_Processo.Value
            SalvarConfiguracao "Compras", Me.chk_Compras.Value
            SalvarConfiguracao "Tesouraria", Me.chk_Tesouraria.Value
            frmCadEmpenho.MantemForm gstrImprimir
            Unload Me
        Else
            frmCadEmpenho.blnAtivaFormImprime = True
        End If
    End If
End Sub

Private Function blnDadosOk()
   
   If Len(Trim(txt_EmpInicial)) = 0 Or Len(Trim(txt_EmpFinal)) = 0 Or Len(Trim(txt_ParcInicial)) = 0 Or Len(Trim(txt_ParcFinal)) = 0 Then
      ExibeMensagem "Preencha os dados corretamente."
      txt_EmpInicial.SetFocus
      Exit Function
   End If
   
   If (Val(txt_EmpInicial) > Val(txt_EmpFinal)) Then
      ExibeMensagem "O Empenho Inicial não pode ser maior que o empenho final."
      txt_EmpInicial.SetFocus
      Exit Function
   End If
   If Val(txtintExercicio) = 0 Then
      ExibeMensagem "O exercício do Empenho deve ser informado corretamente."
      txtintExercicio.SetFocus
      Exit Function
   End If
   blnDadosOk = True
End Function

Private Sub LerConfiguracaoDeImpressao()
    Dim test As String
    'Fornecedor
    If gstrLeValorRegister(HKEY_CURRENT_USER, "SOFTWARE\CPD\AdGover\Parâmetros\Orçamentários\", "Fornecedor") = "" Then
        SalvarConfiguracao "Fornecedor", Me.chk_Fornecedor.Value
    Else
        Me.chk_Fornecedor.Value = Val(gstrLeValorRegister(HKEY_CURRENT_USER, "SOFTWARE\CPD\AdGover\Parâmetros\Orçamentários\", "Fornecedor"))
    End If
    'Almoxarifado
    If gstrLeValorRegister(HKEY_CURRENT_USER, "SOFTWARE\CPD\AdGover\Parâmetros\Orçamentários\", "Almoxarifado") = "" Then
        SalvarConfiguracao "Almoxarifado", Me.chk_Almoxarifado.Value
    Else
        Me.chk_Almoxarifado.Value = Val(gstrLeValorRegister(HKEY_CURRENT_USER, "SOFTWARE\CPD\AdGover\Parâmetros\Orçamentários\", "Almoxarifado"))
    End If
    
    'Processo
    If gstrLeValorRegister(HKEY_CURRENT_USER, "SOFTWARE\CPD\AdGover\Parâmetros\Orçamentários\", "Processo") = "" Then
        SalvarConfiguracao "Processo", Me.chk_Processo.Value
    Else
        Me.chk_Processo.Value = Val(gstrLeValorRegister(HKEY_CURRENT_USER, "SOFTWARE\CPD\AdGover\Parâmetros\Orçamentários\", "Processo"))
    End If
    
    'Compras
    If gstrLeValorRegister(HKEY_CURRENT_USER, "SOFTWARE\CPD\AdGover\Parâmetros\Orçamentários\", "Compras") = "" Then
        SalvarConfiguracao "Compras", Me.chk_Compras.Value
    Else
        Me.chk_Compras.Value = Val(gstrLeValorRegister(HKEY_CURRENT_USER, "SOFTWARE\CPD\AdGover\Parâmetros\Orçamentários\", "Compras"))
    End If
    
    'Tesouraria
    If gstrLeValorRegister(HKEY_CURRENT_USER, "SOFTWARE\CPD\AdGover\Parâmetros\Orçamentários\", "Tesouraria") = "" Then
        SalvarConfiguracao "Tesouraria", Me.chk_Tesouraria.Value
    Else
        Me.chk_Tesouraria.Value = Val(gstrLeValorRegister(HKEY_CURRENT_USER, "SOFTWARE\CPD\AdGover\Parâmetros\Orçamentários\", "Tesouraria"))
    End If
End Sub

Private Sub SalvarConfiguracao(strNomedaChave As String, intCheckBoxValor As Integer)
    SetRegString HKEY_CURRENT_USER, "SOFTWARE\CPD\AdGover\Parâmetros\Orçamentários\", strNomedaChave, Trim(Str(intCheckBoxValor))
    gGravaValorRegister HKEY_CURRENT_USER, "SOFTWARE\CPD\AdGover\Parâmetros\Orçamentários\", strNomedaChave, Trim(Str(intCheckBoxValor))
End Sub
