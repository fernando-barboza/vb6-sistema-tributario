VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} rptSubFichaLancamentoImobiliario 
   Caption         =   "rptSubFichaLancamentoImobiliario (ActiveReport)"
   ClientHeight    =   10815
   ClientLeft      =   30
   ClientTop       =   285
   ClientWidth     =   15360
   MDIChild        =   -1  'True
   _ExtentX        =   27093
   _ExtentY        =   19076
   SectionData     =   "rptSubFichaLancamentoImobiliario.dsx":0000
End
Attribute VB_Name = "rptSubFichaLancamentoImobiliario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public lngIDLancamentoContabil As Long

Private Sub Detail_Format()
    txtdblFatorObsolencia = gstrConvVrDoSql(txtdblFatorObsolencia, 2)
End Sub

Private Sub GroupFooter2_Format()
    Dim strSql      As String
    Dim adoResultado As ADODB.Recordset

    strSql = ""
    strSql = strSql & "Select "
    strSql = strSql & "LV.Intparcela, "
    strSql = strSql & "LV.Dblvalor, "
    strSql = strSql & "LV.Dtmdtvencimento "
    strSql = strSql & "From "
    strSql = strSql & "tbllancamentoalfa LA, "
    strSql = strSql & "tbllancamentovalor LV "
    strSql = strSql & "Where "
    strSql = strSql & "LA.Pkid = LV.Intlancamentoalfa AND "
    strSql = strSql & "LA.Pkid = " & lngIDLancamentoContabil

    Set adoResultado = New ADODB.Recordset
    Set gobjBanco = New clsBanco

    With rptSubFichaLancamentoImobiliarioParcelas

        If gobjBanco.CriaADO(strSql, 5, adoResultado) Then

            If adoResultado.EOF Then
                SubReport1.Visible = False
                Exit Sub
            Else
                SubReport1.Visible = True
            End If

            If bytDBType = EDatabases.SQLServer Then
                .adoDataControl.ConnectionString = "driver={SQL Server};Server=" & gstrServidor & ";database=" & gstrDatabase & ";uid=" & gstrUsername & ";pwd=" & gstrPassword & ";"
            Else
                .adoDataControl.ConnectionString = "Provider=MSDAORA.1;Password=" & gstrPassword & ";User ID=" & gstrUsername & ";Data Source=" & gstrServidor & ";Persist Security Info=True"
            End If
            Set .adoDataControl.Recordset = adoResultado
            Set SubReport1.object = rptSubFichaLancamentoImobiliarioParcelas

        End If

    End With

End Sub
