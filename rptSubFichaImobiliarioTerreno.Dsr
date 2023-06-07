VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptSubFichaImobiliarioTerreno 
   Caption         =   "Tributario - rptSubFichaImobiliarioTerreno (ActiveReport)"
   ClientHeight    =   6525
   ClientLeft      =   1125
   ClientTop       =   2175
   ClientWidth     =   11985
   MDIChild        =   -1  'True
   _ExtentX        =   21140
   _ExtentY        =   11509
   SectionData     =   "rptSubFichaImobiliarioTerreno.dsx":0000
End
Attribute VB_Name = "rptSubFichaImobiliarioTerreno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Detail_Format()
    Line28.Y2 = Detail.Height
    Line29.Y2 = Detail.Height
    txt_dblValor = gstrConvVrDoSql(txt_dblValor, 1)
    
End Sub

Private Sub GHGleba_Format()
    FatorGleba
End Sub

Private Sub GHTerreno_Format()
    FatorProfundidade
End Sub

Private Sub FatorProfundidade()
Dim strSql As String
Dim dblValor As Double
Dim adoResultado As New ADODB.Recordset

    'Vamos pegar o fator de profundidade
    If Val(txt_dblValorProfundidade) > 0 Then
        dblValor = IIf(txt_dblAreaTerreno <= "", 0, txt_dblAreaTerreno) / txt_dblValorProfundidade
    Else
        dblValor = 0
    End If
    strSql = ""
    strSql = strSql & "Select "
    strSql = strSql & "dblfaixainicial, "
    strSql = strSql & "dblfaixafinal, "
    strSql = strSql & "dblValor AS Indice "
    strSql = strSql & "From "
    strSql = strSql & gstrValorDaFaixa
    strSql = strSql & " Where "
    strSql = strSql & "intfaixadevalores = " & FATOR_ZONEAMENTO

    Set adoResultado = Nothing
    Set gobjBanco = New clsBanco
    If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
        With adoResultado
            If Not .EOF Then
                Do While .EOF = False
                    If dblValor >= gstrENulo(!dblFaixaInicial) And dblValor <= gstrENulo(!dblFaixaFinal) Then
                        GHTerreno.Visible = True
                        GHTerreno.Height = 180
                        txt_strFatorTerreno = "Fator Profundidade"
                        txt_strtDetalheTerreno = dblValor
                        txt_dblValorTerreno.Text = gstrENulo(!Indice)
                        Exit Do
                    Else
                        GHTerreno.Visible = False
                        GHTerreno.Height = 0
                    End If
                    .MoveNext
                Loop
            End If
        End With
    End If
    Set adoResultado = Nothing
    Set gobjBanco = Nothing

End Sub

Private Function FatorGleba()

Dim strSql As String
Dim adoResultado As New ADODB.Recordset
Dim dblValorGleba

    
    dblValorGleba = IIf(txt_dblAreaTerreno <= "", 0, txt_dblAreaTerreno)
    
    If dblValorGleba >= CDbl("10000") Then
        'Vamos pegar o fator de Gleba
        strSql = ""
        strSql = strSql & "Select "
        strSql = strSql & "dblfaixainicial, "
        strSql = strSql & "dblfaixafinal, "
        strSql = strSql & "dblValor AS Indice "
        strSql = strSql & "From "
        strSql = strSql & gstrValorDaFaixa
        strSql = strSql & " Where "
        strSql = strSql & "intfaixadevalores = " & FATOR_SITUACAO
        Set gobjBanco = New clsBanco
        If gobjBanco.CriaADO(strSql, 5, adoResultado) Then
            With adoResultado
                If Not .EOF Then
                    Do While .EOF = False
                        If dblValorGleba >= gstrConvVrDoSql(!dblFaixaInicial) And dblValorGleba <= gstrConvVrDoSql(!dblFaixaFinal) Then
                            GHGleba.Visible = True
                            GHGleba.Height = 180
                            txt_strFatorGleba = "Fator Gleba"
                            txt_strDetalheGleba = dblValorGleba
                            txt_dblValorGleba = gstrENulo(!Indice)
                            Exit Do
                        End If
                        .MoveNext
                        
                    Loop
                End If
            End With
        End If
        Set adoResultado = Nothing
        Set gobjBanco = Nothing
    Else
        GHGleba.Visible = False
        GHGleba.Height = 0
    End If

End Function


Private Sub GroupHeader1_Format()

    If Not adoDataControl.Recordset.BOF And Not adoDataControl.Recordset.EOF Then
        Field52 = gstrConvVrDoSql(gstrENulo(adoDataControl.Recordset!strTestadaValor))
    End If
End Sub
