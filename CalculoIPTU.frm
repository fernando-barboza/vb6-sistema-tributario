VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form frmCalculoIPTU 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cálculo do IPTU"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8520
   Icon            =   "CalculoIPTU.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   8520
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab tab_3DPasta 
      Height          =   2145
      Left            =   120
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   510
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   3784
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Cálculo"
      TabPicture(0)   =   "CalculoIPTU.frx":1042
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fra_Progressao"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmd_Iniciar"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmd_Cancelar"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin VB.CommandButton cmd_Cancelar 
         Caption         =   "&Cancelar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   7050
         TabIndex        =   8
         Top             =   750
         Width           =   1110
      End
      Begin VB.CommandButton cmd_Iniciar 
         Caption         =   "&Iniciar"
         Height          =   375
         Left            =   5820
         TabIndex        =   7
         Top             =   750
         Width           =   1110
      End
      Begin VB.Frame fra_Progressao 
         Caption         =   " Progressão "
         ClipControls    =   0   'False
         Height          =   735
         Left            =   180
         TabIndex        =   3
         Top             =   1170
         Width           =   7965
         Begin Threed.SSPanel ssp_Progressao 
            Height          =   345
            Left            =   480
            TabIndex        =   4
            Top             =   255
            Width           =   6855
            _Version        =   65536
            _ExtentX        =   12091
            _ExtentY        =   609
            _StockProps     =   15
            BackColor       =   12632256
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.26
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            BevelInner      =   1
            FloodType       =   1
         End
         Begin VB.Label lbl_PorCento 
            AutoSize        =   -1  'True
            Caption         =   "100%"
            Height          =   195
            Index           =   1
            Left            =   7365
            TabIndex        =   6
            Top             =   330
            Width           =   390
         End
         Begin VB.Label lbl_PorCento 
            AutoSize        =   -1  'True
            Caption         =   "0%"
            Height          =   195
            Index           =   0
            Left            =   210
            TabIndex        =   5
            Top             =   345
            Width           =   210
         End
      End
   End
   Begin VB.TextBox txtPKId 
      Height          =   270
      Left            =   -30
      TabIndex        =   0
      Top             =   960
      Visible         =   0   'False
      Width           =   645
   End
   Begin MSComctlLib.ImageList img_Arquivo 
      Left            =   -30
      Top             =   1290
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CalculoIPTU.frx":105E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CalculoIPTU.frx":11BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CalculoIPTU.frx":131E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CalculoIPTU.frx":147E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CalculoIPTU.frx":15DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CalculoIPTU.frx":173A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "CalculoIPTU.frx":1896
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlb_BarraFermta 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8520
      _ExtentX        =   15028
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "img_Arquivo"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Novo"
            Object.ToolTipText     =   "Novo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Salvar"
            Object.ToolTipText     =   "Salvar"
            ImageIndex      =   2
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Imprimir"
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Deletar"
            Object.ToolTipText     =   "Deletar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Aplicar"
            Object.ToolTipText     =   "Aplicar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "Grade"
            Object.ToolTipText     =   "Grade"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Fechar"
            Object.ToolTipText     =   "Fechar"
            ImageIndex      =   7
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "frmCalculoIPTU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim adoResultado               As ADODB.Recordset
Dim adoImovel                  As ADODB.Recordset

Dim adoAcidenteGeografico      As ADODB.Recordset
Dim adoRedutorDeProfundidade   As ADODB.Recordset
Dim adoPlantaDeValores         As ADODB.Recordset

Dim adoPadraoDeAcabamento      As ADODB.Recordset
Dim adoConstrucaoPorUtilizacao As ADODB.Recordset
Dim adoLocalizacaoPorEspecie   As ADODB.Recordset
Dim adoDepreciacaoPorIdade     As ADODB.Recordset
Dim adoReducaoPorArea          As ADODB.Recordset
Dim adoEquipamento             As ADODB.Recordset
Dim adoAliquota                As ADODB.Recordset

Dim adoIndicesDiversos         As ADODB.Recordset
Dim adoIndicesEconomicos       As ADODB.Recordset
Dim adoTaxaLimpezaPublica      As ADODB.Recordset

Dim strSql                     As String
Dim lngCodImovel               As Long
Dim strTipoImovel              As String
Dim dblUFIR                    As Double

'Valor Venal do Terreno
Dim dblValorVenalTerreno       As Double
Dim dblFatorDepreciacaoT       As Double
Dim dblRedutor1T               As Double
Dim dblRedutor2T               As Double
Dim dblAreaTributavel          As Double
Dim dblFracaoIdeal             As Double
Dim dblValorM2DoTerreno        As Double
Dim dblAreaUnidadeConstruida   As Double
Dim dblAreaTotalConstruida     As Double
Dim dblProfundidade            As Double
Dim strTipoDeSolo              As String
Dim strTopografia              As String
Dim strFormato                 As String

'Valor Venal da Construção
Dim dblValorVenalConstrucao    As Double
Dim dblValorM2Construcao       As Double
Dim dblAreaDaUnidade           As Double
Dim dblFatorDepreciacaoC       As Double
Dim dblRedutor1C               As Double
Dim dblRedutor2C               As Double
Dim dblRedutor3C               As Double
Dim strEspecie                 As String
Dim strLocalizacao             As String
Dim intIdadeImovel             As Integer

'Imposto Sobre Valor Venal
Dim dblImpostoSobreValorVenal  As Double
Dim dblAliquota                As Double

'Taxas
Dim dblTaxaDeConservacao       As Double
Dim dblTaxaDeIluminacao        As Double
Dim dblTaxaDeLimpeza           As Double
Dim dblTestadaDoLote           As Double
Dim dblIndiceConservacao       As Double
Dim dblIndiceIluminacao        As Double
Dim dblQuantidadeUFIR          As Double
Dim strEquipamentos            As String

Private Sub cmd_Cancelar_Click()
    gblnCancelar = True
End Sub

Private Sub cmd_Cancelar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Screen.MousePointer = 0
End Sub

Private Sub cmd_Iniciar_Click()
    Dim lngCont        As Long
    Dim bytTipoImposto As Byte
    
    If MsgBox("Confirma início do cálculo de IPTU?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If
    
    Screen.MousePointer = 11
    cmd_Cancelar.Enabled = True
    cmd_Iniciar.Enabled = False
    DoEvents
    
    'Redutores de Profundidade
    strSql = ""
    strSql = strSql & "Select * From " & gstrRedutorDeProfundidade & " "
    strSql = strSql & "Order By dblAte"
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoRedutorDeProfundidade) Then
        If adoRedutorDeProfundidade.EOF Then
            ExibeMensagem "Tabela de Redutores de Profundidade não possui valores."
            Exit Sub
        End If
    End If
    
    'Planta de Valores
    strSql = ""
    strSql = strSql & "Select * From " & gstrPlantaDeValor & " "
    strSql = strSql & "Order By intSetor"
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoPlantaDeValores) Then
        If adoPlantaDeValores.EOF Then
            ExibeMensagem "Tabela de Planta de Valores não possui valores."
            Exit Sub
        End If
    End If
    
    'Acidentes Geográficos
    strSql = ""
    strSql = strSql & "Select * From " & gstrAcidenteGeografico & " "
    strSql = strSql & "Order By dblAte"
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoAcidenteGeografico) Then
        If adoAcidenteGeografico.EOF Then
            ExibeMensagem "Tabela de Acidentes Geográficos não possui valores."
            Exit Sub
        End If
    End If
    
    'Padrão de Acabamento
    strSql = ""
    strSql = strSql & "Select * From " & gstrPadraoDeAcabamento & " "
    strSql = strSql & "Order By dblAte"
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoPadraoDeAcabamento) Then
        If adoPadraoDeAcabamento.EOF Then
            ExibeMensagem "Tabela de Padrões de Acabamento não possui valores."
            Exit Sub
        End If
    End If
    
    'Construção por Utilização
    strSql = ""
    strSql = strSql & "Select * From " & gstrConstrucaoPorUtilizacao & " "
    strSql = strSql & "Order By strUtilizacao"
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoConstrucaoPorUtilizacao) Then
        If adoConstrucaoPorUtilizacao.EOF Then
            ExibeMensagem "Tabela de M2 de Construção por Utilização não possui valores."
            Exit Sub
        End If
    End If
    
    'Localização por Espécie
    strSql = ""
    strSql = strSql & "Select * From " & gstrLocalizacaoPorEspecie & " "
    strSql = strSql & "Order By strEspecie"
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoLocalizacaoPorEspecie) Then
        If adoLocalizacaoPorEspecie.EOF Then
            ExibeMensagem "Tabela de localização por espécie não possui valores."
            Exit Sub
        End If
    End If
    
    'Depreciação por Idade
    strSql = ""
    strSql = strSql & "Select * From " & gstrDepreciacaoPorIdade & " "
    strSql = strSql & "Order By intAte"
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoDepreciacaoPorIdade) Then
        If adoDepreciacaoPorIdade.EOF Then
            ExibeMensagem "Tabela de depreciação por idade não possui valores."
            Exit Sub
        End If
    End If
    
    'Redução por Área
    strSql = ""
    strSql = strSql & "Select * From " & gstrReducaoPorArea & " "
    strSql = strSql & "Order By strEspecie"
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoReducaoPorArea) Then
        If adoReducaoPorArea.EOF Then
            ExibeMensagem "Tabela de redução por área não possui valores."
            Exit Sub
        End If
    End If
    
    'Equipamentos (Serviços do local)
    strSql = ""
    strSql = strSql & "Select * From " & gstrEquipamento & " "
    strSql = strSql & "Order By PKId"
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoEquipamento) Then
        If adoEquipamento.EOF Then
            ExibeMensagem "Tabela de equipamentos não possui valores."
            Exit Sub
        End If
    End If
    
    'Alíquota
    strSql = ""
    strSql = strSql & "Select * From " & gstrAliquota & " "
    strSql = strSql & "Order By intPontos"
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoAliquota) Then
        If adoAliquota.EOF Then
            ExibeMensagem "Tabela de alíquotas não possui valores."
            Exit Sub
        End If
    End If
    
    'Índices Diversos
    strSql = ""
    strSql = strSql & "Select * From " & gstrIndicesDiversos & " "
    strSql = strSql & "Order By dblValor"
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoIndicesDiversos) Then
        If adoIndicesDiversos.EOF Then
            ExibeMensagem "Tabela de índices diversos não possui valores."
            Exit Sub
        Else
            adoIndicesDiversos.MoveFirst
            adoIndicesDiversos.Find "PKId = " & 2
            If adoIndicesDiversos.EOF Then
                ExibeMensagem "Não foi encontrado índice de conservação do imóvel correspondente."
                Exit Sub
            End If
            dblIndiceConservacao = gvntConvVrDoSql(adoIndicesDiversos!dblValor)
            
            adoIndicesDiversos.MoveFirst
            adoIndicesDiversos.Find "PKId = " & 1
            If adoIndicesDiversos.EOF Then
                ExibeMensagem "Não foi encontrado índice de iluminação do imóvel correspondente."
                Exit Sub
            End If
            dblIndiceIluminacao = gvntConvVrDoSql(adoIndicesDiversos!dblValor)
        End If
    End If
    
    'UFIR
    strSql = ""
    strSql = strSql & "Select IE.PKId, IE.dblValor, IE.dtmData, "
    strSql = strSql & "E.strSiglaIndexador Sigla "
    strSql = strSql & "From " & gstrIndiceEconomico & " IE, " & gstrIndexadorEconomico & " E "
    strSql = strSql & "Where IE.intSiglaIndexador = E.PKId "
    strSql = strSql & "Order By IE.dtmData Desc"
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoIndicesEconomicos) Then
        If adoIndicesEconomicos.EOF Then
            ExibeMensagem "Tabela de índices econômicos não possui valores."
            Exit Sub
        Else
            adoIndicesEconomicos.MoveFirst
            adoIndicesEconomicos.Find "Sigla = 'UFIR'"
            If adoIndicesEconomicos.EOF Then
                ExibeMensagem "Não foi encontrado um valor de UFIR."
                Exit Sub
            End If
            dblUFIR = gvntConvVrDoSql(adoIndicesEconomicos!dblValor)
        End If
    End If
    
    'Taxa de Limpeza Púbica
    strSql = ""
    strSql = strSql & "Select * From " & gstrTaxaLimpezaPublica & " "
    strSql = strSql & "Order By strDescricao"
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoTaxaLimpezaPublica) Then
        If adoTaxaLimpezaPublica.EOF Then
            ExibeMensagem "Tabela de taxas de limpeza pública não possui valores."
            Exit Sub
        End If
    End If
    
    
    strSql = ""
    strSql = strSql & "Select * "
    strSql = strSql & "From " & gstrImovel & " IM "
    strSql = strSql & "" & ""
    
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoImovel) Then
        With adoImovel
            Do While .EOF = False
                dblValorVenalTerreno = 0
                dblValorVenalConstrucao = 0
                dblImpostoSobreValorVenal = 0
                dblTaxaDeConservacao = 0
                dblTaxaDeIluminacao = 0
                dblTaxaDeLimpeza = 0
            
                If Abs(adoImovel!ImpostoPredial) = 1 Then
                    bytTipoImposto = 0 'Predial
                    strTipoImovel = "Predial"
                Else
                    bytTipoImposto = 1 'Territorial
                    strTipoImovel = "Territorial"
                End If
                
                lngCodImovel = !PKId
                
                dblValorVenalTerreno = CalculoValorVenalTerreno(bytTipoImposto, adoImovel)
                
                dblValorVenalConstrucao = CalculoValorVenalConstrucao(adoImovel)
                
                dblImpostoSobreValorVenal = CalculoImpostoSobreValorVenal(bytTipoImposto, adoImovel)
                
                dblTaxaDeConservacao = TaxaDeConservacao(adoImovel)
                
                dblTaxaDeIluminacao = TaxaDeIluminacao(adoImovel)
                
                dblTaxaDeLimpeza = TaxaDeLimpeza(bytTipoImposto, adoImovel)
                
                Call GravaMemoriaDeCalculo
                
                lngCont = lngCont + 1
                ssp_Progressao.FloodPercent = Int(100 * lngCont / adoImovel.RecordCount)
                .MoveNext
            Loop
        End With
        Set gobjBanco = Nothing
        adoImovel.Close
        Set adoImovel = Nothing
    End If
    
    cmd_Cancelar.Enabled = False
    cmd_Iniciar.Enabled = True
    
    Screen.MousePointer = 0
    ExibeMensagem "Cálculo finalizado."
    ssp_Progressao.FloodPercent = 0
End Sub


Function CalculoValorVenalTerreno(bytTipo As Byte, ByVal adoTMPImovel As ADODB.Recordset) As Double
    dblFatorDepreciacaoT = 0
    dblRedutor1T = 0
    dblRedutor2T = 0
    dblAreaTributavel = 0
    dblFracaoIdeal = 0
    dblValorM2DoTerreno = 0
    dblAreaUnidadeConstruida = 0
    dblAreaTotalConstruida = 0
    dblProfundidade = 0
    
    strTipoDeSolo = ""
    strFormato = ""
    strTopografia = ""
    
    With adoTMPImovel
        dblAreaUnidadeConstruida = gvntConvVrDoSql(!AreaDoTerreno)
        dblAreaTotalConstruida = gvntConvVrDoSql(!AreaDeConstrucao)
        dblProfundidade = gvntConvVrDoSql(!Profundidade)
        
        If !AreaIndivisa Or dblAreaUnidadeConstruida > 2000 Then
            dblAreaTributavel = dblAreaUnidadeConstruida - ((dblAreaUnidadeConstruida / 500) * 140)
        Else
            dblAreaTributavel = dblAreaUnidadeConstruida
        End If
                
        Select Case bytTipo
            Case 0  'Predial
                dblFracaoIdeal = dblAreaUnidadeConstruida / dblAreaTotalConstruida
            
            Case 1  'Territorial
                dblFracaoIdeal = 1
        End Select
        
        If dblProfundidade = 0 Then
            dblRedutor1T = 1
        Else
            dblRedutor1T = (dblProfundidade + 30) / (dblProfundidade * 2)
            adoRedutorDeProfundidade.MoveFirst
            adoRedutorDeProfundidade.Find "dblAte >= " & dblRedutor1T
            If adoRedutorDeProfundidade.EOF Then
                ExibeMensagem "Não foi encontrado um redutor de profundidade correspondente."
            End If
            dblRedutor1T = gvntConvVrDoSql(adoRedutorDeProfundidade!dblValor)
        End If
    
        If !Firme Then
            dblRedutor2T = dblRedutor2T + 1
            strTipoDeSolo = "Firme"
        ElseIf !Arenoso Then
                dblRedutor2T = dblRedutor2T + 2
                strTipoDeSolo = "Arenoso"
        ElseIf !Pantanoso Then
                dblRedutor2T = dblRedutor2T + 3
                strTipoDeSolo = "Pantanoso"
        ElseIf !Rochoso Then
                dblRedutor2T = dblRedutor2T + 3
                strTipoDeSolo = "Rochoso"
        End If
        
        If !Regular Then
            dblRedutor2T = dblRedutor2T + 1
            strFormato = "Regular"
        ElseIf !Irregular Then
                dblRedutor2T = dblRedutor2T + 2
                strFormato = "Irregular"
        End If
        
        If !Plano Then
            dblRedutor2T = dblRedutor2T + 1
            strTopografia = "Plano"
        ElseIf !AcimaNivel Then
                dblRedutor2T = dblRedutor2T + 2
                strTopografia = "Acima do Nível"
        ElseIf !AbaixoNivel Then
                dblRedutor2T = dblRedutor2T + 2
                strTopografia = "Abaixo do Nível"
        ElseIf !Acidentado Then
                dblRedutor2T = dblRedutor2T + 3
                strTopografia = "Acidentado"
        End If
        
        adoAcidenteGeografico.MoveFirst
        adoAcidenteGeografico.Find "dblAte >= " & dblRedutor2T
        If adoAcidenteGeografico.EOF Then
            ExibeMensagem "Não foi encontrado um acidente geográfico correspondente."
        End If
        dblRedutor2T = gvntConvVrDoSql(adoAcidenteGeografico!dblValor)
        
        adoPlantaDeValores.MoveFirst
        adoPlantaDeValores.Find "intSetor = " & !Setor
        If adoPlantaDeValores.EOF Then
            ExibeMensagem "Não foi encontrado uma planta de valor para o imóvel correspondente."
        End If
        dblValorM2DoTerreno = adoPlantaDeValores!dblValorTerreno
        
    End With
    
    dblFatorDepreciacaoT = dblRedutor1T * dblRedutor2T
    
    CalculoValorVenalTerreno = dblAreaTributavel * dblFracaoIdeal * dblValorM2DoTerreno * dblFatorDepreciacaoT
End Function

Function CalculoValorVenalConstrucao(ByVal adoTMPImovel As ADODB.Recordset) As Double
    Dim intPadraoDeAcabamento As Integer
    Dim strCodigoPadrao       As String
    Dim strUtilizacao         As String
    
    
    dblValorM2Construcao = 0
    dblAreaDaUnidade = 0
    dblFatorDepreciacaoC = 0
    dblRedutor1C = 0
    dblRedutor2C = 0
    dblRedutor3C = 0
    strEspecie = ""
    strLocalizacao = ""
    intIdadeImovel = 0
    
    With adoTMPImovel
        dblAreaDaUnidade = !AreaDeConstrucao
        
        intPadraoDeAcabamento = 0
        
        'Conservaçao
        If !Pessima Then
            intPadraoDeAcabamento = 1
        ElseIf !Ma Then
                intPadraoDeAcabamento = 2
        ElseIf !Regular1 Then
                intPadraoDeAcabamento = 3
        ElseIf !Boa Then
                intPadraoDeAcabamento = 4
        ElseIf !Otima Then
                intPadraoDeAcabamento = 5
        End If
        
        'Estrutura
        If !Adobe Then
            intPadraoDeAcabamento = intPadraoDeAcabamento + 1
        ElseIf !Madeira Then
                intPadraoDeAcabamento = intPadraoDeAcabamento + 2
        ElseIf !Alvenaria Then
                intPadraoDeAcabamento = intPadraoDeAcabamento + 3
        ElseIf !Concreto Then
                intPadraoDeAcabamento = intPadraoDeAcabamento + 4
        ElseIf !Metalica Then
                intPadraoDeAcabamento = intPadraoDeAcabamento + 5
        End If
        
        'Cobertura
        If !Barro Then
            intPadraoDeAcabamento = intPadraoDeAcabamento + 1
        ElseIf !Amianto Then
                intPadraoDeAcabamento = intPadraoDeAcabamento + 2
        ElseIf !AluminioZinco Then
                intPadraoDeAcabamento = intPadraoDeAcabamento + 3
        ElseIf !Ceramica Then
                intPadraoDeAcabamento = intPadraoDeAcabamento + 4
        ElseIf !CobertLage Then
                intPadraoDeAcabamento = intPadraoDeAcabamento + 5
        End If
        
        'Acabamento Externo
        If !AcabExtSem Then
            intPadraoDeAcabamento = intPadraoDeAcabamento + 1
        ElseIf !AcabExtCaiacao Then
                intPadraoDeAcabamento = intPadraoDeAcabamento + 2
        ElseIf !AcabExtPinturaSimples Then
                intPadraoDeAcabamento = intPadraoDeAcabamento + 3
        ElseIf !AcabExtArgamassaPintada Then
                intPadraoDeAcabamento = intPadraoDeAcabamento + 4
        ElseIf !AcabExtRevestimentoEspecial Then
                intPadraoDeAcabamento = intPadraoDeAcabamento + 5
        End If
        
        'Piso
        If !SemPiso Then
            intPadraoDeAcabamento = intPadraoDeAcabamento + 1
        ElseIf !Cimento Then
                intPadraoDeAcabamento = intPadraoDeAcabamento + 2
        ElseIf !Taco Then
                intPadraoDeAcabamento = intPadraoDeAcabamento + 3
        ElseIf !Ceramica1 Then
                intPadraoDeAcabamento = intPadraoDeAcabamento + 4
        ElseIf !Tabua Then
                intPadraoDeAcabamento = intPadraoDeAcabamento + 5
        End If
        
        'Acabamento Interno
        If !Sem Then
            intPadraoDeAcabamento = intPadraoDeAcabamento + 1
        ElseIf !Caiacao Then
                intPadraoDeAcabamento = intPadraoDeAcabamento + 2
        ElseIf !PinturaSimples Then
                intPadraoDeAcabamento = intPadraoDeAcabamento + 3
        ElseIf !ArgamassaPintada Then
                intPadraoDeAcabamento = intPadraoDeAcabamento + 4
        ElseIf !RevestimentoEspecial Then
                intPadraoDeAcabamento = intPadraoDeAcabamento + 5
        End If
        
        'Forro
        If !SemForro Then
            intPadraoDeAcabamento = intPadraoDeAcabamento + 1
        ElseIf !ForroMadeira Then
                intPadraoDeAcabamento = intPadraoDeAcabamento + 2
        ElseIf !Lage Then
                intPadraoDeAcabamento = intPadraoDeAcabamento + 3
        ElseIf !ForroEspecial Then
                intPadraoDeAcabamento = intPadraoDeAcabamento + 4
        End If
        
        'Janelas
        If !Basculante Then
            intPadraoDeAcabamento = intPadraoDeAcabamento + 1
        ElseIf !Madeira1 Then
                intPadraoDeAcabamento = intPadraoDeAcabamento + 2
        ElseIf !Metalon Then
                intPadraoDeAcabamento = intPadraoDeAcabamento + 3
        ElseIf !MadeiraTrabalhada Then
                intPadraoDeAcabamento = intPadraoDeAcabamento + 4
        ElseIf !EsquadriaAluminio Then
                intPadraoDeAcabamento = intPadraoDeAcabamento + 5
        End If
        
        adoPadraoDeAcabamento.MoveFirst
        adoPadraoDeAcabamento.Find "dblAte >= " & intPadraoDeAcabamento
        If adoPadraoDeAcabamento.EOF Then
            ExibeMensagem "Não foi encontrado um padrão de acabamento correspondente."
            Exit Function
        Else
            strCodigoPadrao = Trim(adoPadraoDeAcabamento!strCodigo)
        End If
        
        'Utilização
        If !Residencial Then
            strUtilizacao = "Residencial"
        ElseIf !Comercial Then
                strUtilizacao = "Comercial"
        ElseIf !Industrial Then
                strUtilizacao = "Industrial"
        ElseIf !Hospital Then
                strUtilizacao = "Hospital"
        ElseIf !Escola Then
                strUtilizacao = "Escola"
        ElseIf !Hotel Then
                strUtilizacao = "Hotel"
        ElseIf !Igreja Then
                strUtilizacao = "Igreja"
        ElseIf !ServicoPublico Then
                strUtilizacao = "Serviço Público"
        ElseIf !OutrosServicos Then
                strUtilizacao = "Outros Serviços"
        End If
                
        adoConstrucaoPorUtilizacao.MoveFirst
        adoConstrucaoPorUtilizacao.Find "strUtilizacao = '" & strUtilizacao & "'"
        If adoConstrucaoPorUtilizacao.EOF Then
            ExibeMensagem "Não foi encontrado valor de M2 por espécie do imóvel correspondente."
            Exit Function
        Else
            If strCodigoPadrao = "A" Then
                dblValorM2Construcao = gvntConvVrDoSql(adoConstrucaoPorUtilizacao!dblPadraoA)
            ElseIf strCodigoPadrao = "B" Then
                    dblValorM2Construcao = gvntConvVrDoSql(adoConstrucaoPorUtilizacao!dblPadraoB)
            ElseIf strCodigoPadrao = "C" Then
                    dblValorM2Construcao = gvntConvVrDoSql(adoConstrucaoPorUtilizacao!dblPadraoC)
            ElseIf strCodigoPadrao = "D" Then
                    dblValorM2Construcao = gvntConvVrDoSql(adoConstrucaoPorUtilizacao!dblPadraoD)
            End If
        End If
        
        'Fator - Depreciação(Predial)
        
        'Espécie
        If !CasaDeVila Then
            strEspecie = "Casa de Vila"
        ElseIf !Barracao Then
                strEspecie = "Barracão"
        ElseIf !CasaConjugada Then
                strEspecie = "Casa Conjugada"
        ElseIf !Casa Then
                strEspecie = "Casa"
        ElseIf !Apartamento Then
                strEspecie = "Apartamento"
        ElseIf !Garagem Then
                strEspecie = "Garagem"
        ElseIf !Galpao Then
                strEspecie = "Galpão"
        ElseIf !Sala Then
                strEspecie = "Sala"
        ElseIf !Loja Then
                strEspecie = "Loja"
        End If
        
        'Fator - Localização
        adoLocalizacaoPorEspecie.MoveFirst
        adoLocalizacaoPorEspecie.Find "strEspecie = '" & strEspecie & "'"
        If adoLocalizacaoPorEspecie.EOF Then
            ExibeMensagem "Não foi encontrado fator de localização para o imóvel correspondente correspondente."
            Exit Function
        Else
            If !NoAlinhamento Then
                dblRedutor1C = gvntConvVrDoSql(adoLocalizacaoPorEspecie!dblAlinhamento)
                strLocalizacao = "No Alinhamento"
            ElseIf !Lateral Then
                    dblRedutor1C = gvntConvVrDoSql(adoLocalizacaoPorEspecie!dblEsquina)
                    strLocalizacao = "Esquina"
            ElseIf !Recuado Then
                    dblRedutor1C = gvntConvVrDoSql(adoLocalizacaoPorEspecie!dblRecuado)
                    strLocalizacao = "Recuado"
            ElseIf !DeFundo Then
                    dblRedutor1C = gvntConvVrDoSql(adoLocalizacaoPorEspecie!dblFundo)
                    strLocalizacao = "De Fundo"
            End If
        End If
        
        
        'Fator - Tempo
        intIdadeImovel = Year(Format(Date, "dd/mm/yyyy")) - !AnoConstrucao
        
        adoDepreciacaoPorIdade.MoveFirst
        adoDepreciacaoPorIdade.Find "intAte >= " & intIdadeImovel
        If adoDepreciacaoPorIdade.EOF Then
            ExibeMensagem "Não foi encontrado um redutor de depreciação por idade correspondente."
            Exit Function
        Else
            dblRedutor2C = gvntConvVrDoSql(adoDepreciacaoPorIdade!dblRedutor)
        End If
        
        'Fator - Redução - Por - Área
        
        If !CasaDeVila Then
            dblRedutor3C = 1
        ElseIf !Barracao Then
                dblRedutor3C = 1
        ElseIf !CasaConjugada Then
                dblRedutor3C = 1
        ElseIf !Casa Then
                dblRedutor3C = 1
        ElseIf !Apartamento Then
                dblRedutor3C = 1
        Else
            If dblAreaDaUnidade >= 1000 Then
                dblRedutor3C = "0,70"
            Else
                adoReducaoPorArea.MoveFirst
                adoReducaoPorArea.Find "strEspecie = '" & strEspecie & "'"
                If adoReducaoPorArea.EOF Then
                    ExibeMensagem "Não foi encontrado valor de redução por área do imóvel correspondente."
                    Exit Function
                Else
                    If dblAreaDaUnidade <= 150 Then
                        dblRedutor3C = gvntConvVrDoSql(adoReducaoPorArea!dbl150)
                    ElseIf dblAreaDaUnidade <= 200 Then
                            dblRedutor3C = gvntConvVrDoSql(adoReducaoPorArea!dbl200)
                    ElseIf dblAreaDaUnidade <= 500 Then
                            dblRedutor3C = gvntConvVrDoSql(adoReducaoPorArea!dbl500)
                    ElseIf dblAreaDaUnidade <= 1000 Then
                            dblRedutor3C = gvntConvVrDoSql(adoReducaoPorArea!dbl100)
                    ElseIf dblAreaDaUnidade > 1000 Then
                            dblRedutor3C = gvntConvVrDoSql(adoReducaoPorArea!dblAcima1000)
                    End If
                End If
            End If
        End If
        
        dblFatorDepreciacaoC = dblRedutor1C * dblRedutor2C * dblRedutor3C
        
        CalculoValorVenalConstrucao = dblAreaDaUnidade * dblValorM2Construcao * dblFatorDepreciacaoC
    End With
End Function

Function CalculoImpostoSobreValorVenal(ByVal bytTipo As Byte, _
                                       ByVal adoTMPImovel As ADODB.Recordset) As Double
    dblAliquota = 0
    strEquipamentos = ""
    
    With adoTMPImovel
        If !EquipaPavimentacao Then
            adoEquipamento.MoveFirst
            adoEquipamento.Find "strDescricao Like '*Pavimentação*'"
            If adoEquipamento.EOF Then
                ExibeMensagem "Não foi encontrado ponto do equipamento do imóvel correspondente."
                Exit Function
            End If
            dblAliquota = dblAliquota + adoEquipamento!intPonto
            strEquipamentos = "Pavimentação" & Chr(13) & Chr(10)
        End If
        If !EquipaEletricidade Then
            adoEquipamento.MoveFirst
            adoEquipamento.Find "strDescricao Like '*Elétrica*'"
            If adoEquipamento.EOF Then
                ExibeMensagem "Não foi encontrado ponto do equipamento do imóvel correspondente."
                Exit Function
            End If
            dblAliquota = dblAliquota + adoEquipamento!intPonto
            strEquipamentos = strEquipamentos & "Rede Elétrica" & Chr(13) & Chr(10)
        End If
        If !EquipaEsgoto Then
            adoEquipamento.MoveFirst
            adoEquipamento.Find "strDescricao Like '*Esgoto*'"
            If adoEquipamento.EOF Then
                ExibeMensagem "Não foi encontrado ponto do equipamento do imóvel correspondente."
                Exit Function
            End If
            dblAliquota = dblAliquota + adoEquipamento!intPonto
            strEquipamentos = strEquipamentos & "Rede De Esgoto" & Chr(13) & Chr(10)
        End If
        If !EquipaAgua Then
            adoEquipamento.MoveFirst
            adoEquipamento.Find "strDescricao Like '*Água*'"
            If adoEquipamento.EOF Then
                ExibeMensagem "Não foi encontrado ponto do equipamento do imóvel correspondente."
                Exit Function
            End If
            dblAliquota = dblAliquota + adoEquipamento!intPonto
            strEquipamentos = strEquipamentos & "Rede de Água" & Chr(13) & Chr(10)
        End If
        
        Select Case bytTipo
            Case 0  'Predial
                adoAliquota.MoveFirst
                adoAliquota.Find "intPontos = " & dblAliquota
                If adoAliquota.EOF Then
                    ExibeMensagem "Não foi encontrada alíquota do imóvel correspondente."
                    Exit Function
                End If
                
                If !Residencial Then
                    dblAliquota = adoAliquota!dblPredialResidencial
                Else
                    dblAliquota = adoAliquota!dblPredialNaoResidencial
                End If
                
            Case 1  'Territorial
                
                adoAliquota.MoveFirst
                adoAliquota.Find "intPontos = " & dblAliquota
                If adoAliquota.EOF Then
                    ExibeMensagem "Não foi encontrada alíquota do imóvel correspondente."
                    Exit Function
                End If
                dblAliquota = adoAliquota!dblTerritorial
                
        End Select
        
        CalculoImpostoSobreValorVenal = (dblValorVenalConstrucao + dblValorVenalTerreno) * (dblAliquota / 100)
        
    End With
End Function

Function TaxaDeConservacao(ByVal adoTMPImovel As ADODB.Recordset) As Double
    dblTestadaDoLote = 0
    With adoTMPImovel
        If !EquipaPavimentacao Then
            dblTestadaDoLote = gvntConvVrDoSql(!TestadaReal)
            TaxaDeConservacao = dblTestadaDoLote * dblIndiceConservacao
        End If
    End With
End Function

Function TaxaDeIluminacao(ByVal adoTMPImovel As ADODB.Recordset) As Double
    With adoTMPImovel
        If !EquipaEletricidade Then
            TaxaDeIluminacao = dblIndiceIluminacao * dblUFIR
        End If
    End With
End Function

Function TaxaDeLimpeza(bytTipo As Byte, ByVal adoTMPImovel As ADODB.Recordset) As Double
    With adoTMPImovel
        If !EquipaPavimentacao Then
            Select Case bytTipo
                Case 0  'Predial
                    adoTaxaLimpezaPublica.MoveFirst
                    adoTaxaLimpezaPublica.Find "PKId = 1"
                    If adoTaxaLimpezaPublica.EOF Then
                        ExibeMensagem "Não foi encontrada taxa de limpeza predial correspondente."
                        Exit Function
                    End If
                    dblQuantidadeUFIR = adoTaxaLimpezaPublica!dblQtdeUFIR
                
                Case 1  'Territorial
                    adoTaxaLimpezaPublica.MoveFirst
                    adoTaxaLimpezaPublica.Find "PKId = 2"
                    If adoTaxaLimpezaPublica.EOF Then
                        ExibeMensagem "Não foi encontrada taxa de limpeza territorial correspondente."
                        Exit Function
                    End If
                    dblQuantidadeUFIR = adoTaxaLimpezaPublica!dblQtdeUFIR
            End Select
        
            TaxaDeLimpeza = dblQuantidadeUFIR * dblUFIR
        End If
    End With
End Function

Sub GravaMemoriaDeCalculo()
    strSql = ""
    strSql = strSql & "Insert Into " & gstrMemoriaDeCalculoIPTU & " ("
    strSql = strSql & "intCodImovel, strTipoImovel, dblValorVenalTerreno, dblAreaTributavel, "
    strSql = strSql & "dblFracaoIdeal, dblValorM2Terreno, dblFatorDepreciacaoT, "
    strSql = strSql & "dblAreaUnidadeConstruida, dblAreaTotalConstruida, dblProfundidade, dblRedutor1T, "
    strSql = strSql & "dblRedutor2T, strTipoDeSolo, strFormato, strTopografia, "
    strSql = strSql & "dblValorVenalConstrucao, dblAreaUnidade, dblValorM2Construcao, "
    strSql = strSql & "dblFatorDepreciacaoC, dblRedutor1C, dblRedutor2C, dblRedutor3C, "
    strSql = strSql & "strEspecie, strLocalizacao, intIdadeImovel, dblAreaConstruida, "
    strSql = strSql & "dblImpostoSobreValorVenal, dblAliquota, "
    strSql = strSql & "dblTaxaConservacao, dblTaxaIluminacao, dblTaxaLimpeza,"
    strSql = strSql & "strEquipamentos,"
    strSql = strSql & "dblTestadaReal, dblUFIR, dblIndiceConservacao, dblIndiceIluminacao"
    strSql = strSql & ") Values ("
    strSql = strSql & lngCodImovel & ", '"
    strSql = strSql & strTipoImovel & "', "
    strSql = strSql & gstrConvVrParaSql(dblValorVenalTerreno) & ", "
    strSql = strSql & gstrConvVrParaSql(dblAreaTributavel) & ", "
    strSql = strSql & gstrConvVrParaSql(dblFracaoIdeal) & ", "
    strSql = strSql & gstrConvVrParaSql(dblValorM2DoTerreno) & ", "
    strSql = strSql & gstrConvVrParaSql(dblFatorDepreciacaoT) & ", "
    strSql = strSql & gstrConvVrParaSql(dblAreaUnidadeConstruida) & ", "
    strSql = strSql & gstrConvVrParaSql(dblAreaTotalConstruida) & ", "
    strSql = strSql & gstrConvVrParaSql(dblProfundidade) & ", "
    strSql = strSql & gstrConvVrParaSql(dblRedutor1T) & ", "
    strSql = strSql & gstrConvVrParaSql(dblRedutor2T) & ", '"
    strSql = strSql & strTipoDeSolo & "', '"
    strSql = strSql & strFormato & "', '"
    strSql = strSql & strTopografia & "', "
    strSql = strSql & gstrConvVrParaSql(dblValorVenalConstrucao) & ", "
    strSql = strSql & gstrConvVrParaSql(dblAreaDaUnidade) & ", "
    strSql = strSql & gstrConvVrParaSql(dblValorM2Construcao) & ", "
    strSql = strSql & gstrConvVrParaSql(dblFatorDepreciacaoC) & ", "
    strSql = strSql & gstrConvVrParaSql(dblRedutor1C) & ", "
    strSql = strSql & gstrConvVrParaSql(dblRedutor2C) & ", "
    strSql = strSql & gstrConvVrParaSql(dblRedutor3C) & ", '"
    strSql = strSql & strEspecie & "', '"
    strSql = strSql & strLocalizacao & "', "
    strSql = strSql & intIdadeImovel & ", "
    strSql = strSql & gstrConvVrParaSql(dblAreaDaUnidade) & ", "
    strSql = strSql & gstrConvVrParaSql(dblImpostoSobreValorVenal) & ", "
    strSql = strSql & gstrConvVrParaSql(dblAliquota) & ", "
    strSql = strSql & gstrConvVrParaSql(dblTaxaDeConservacao) & ", "
    strSql = strSql & gstrConvVrParaSql(dblTaxaDeIluminacao) & ", "
    strSql = strSql & gstrConvVrParaSql(dblTaxaDeLimpeza) & ", '"
    strSql = strSql & strEquipamentos & "', "
    strSql = strSql & gstrConvVrParaSql(dblTestadaDoLote) & ", "
    strSql = strSql & gstrConvVrParaSql(dblUFIR) & ", "
    strSql = strSql & gstrConvVrParaSql(dblIndiceConservacao) & ", "
    strSql = strSql & gstrConvVrParaSql(dblIndiceIluminacao) & " "
    strSql = strSql & ")"
    
    If Not gobjBanco.Execute(strSql) Then
        ExibeMensagem ""
    End If
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 664
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
  If KeyCode = vbKeyF1 Then
     Call_HtmlHelp Me.HelpContextID
  End If
End Sub

Private Sub tlb_BarraFermta_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case UCase(Button.Key)
        Case gstrFechar
            Unload Me
    End Select
End Sub



