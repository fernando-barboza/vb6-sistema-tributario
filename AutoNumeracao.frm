VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmAutoNumeracao 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Auto Numeração"
   ClientHeight    =   4710
   ClientLeft      =   2130
   ClientTop       =   2460
   ClientWidth     =   7155
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   7155
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fra_Legenda 
      BackColor       =   &H8000000A&
      Caption         =   "Legenda"
      Height          =   1155
      Left            =   375
      TabIndex        =   2
      Top             =   3330
      Width           =   6405
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   3600
         Picture         =   "AutoNumeracao.frx":0000
         ScaleHeight     =   240
         ScaleWidth      =   2460
         TabIndex        =   5
         Top             =   360
         Width           =   2460
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   225
         Picture         =   "AutoNumeracao.frx":0778
         ScaleHeight     =   240
         ScaleWidth      =   2025
         TabIndex        =   4
         Top             =   660
         Width           =   2025
      End
      Begin VB.PictureBox Picture3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   240
         Picture         =   "AutoNumeracao.frx":0EFE
         ScaleHeight     =   240
         ScaleWidth      =   1875
         TabIndex        =   3
         Top             =   360
         Width           =   1875
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         Height          =   780
         Left            =   105
         Top             =   270
         Width           =   6150
      End
   End
   Begin TabDlg.SSTab tab_3dPasta 
      Height          =   4530
      Left            =   105
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   105
      Width           =   6945
      _ExtentX        =   12250
      _ExtentY        =   7990
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Auto Numeração"
      TabPicture(0)   =   "AutoNumeracao.frx":1618
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "TreeModulos"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin MSComctlLib.TreeView TreeModulos 
         Height          =   2535
         Left            =   195
         TabIndex        =   1
         Top             =   540
         Width           =   6555
         _ExtentX        =   11562
         _ExtentY        =   4471
         _Version        =   393217
         LineStyle       =   1
         Style           =   6
         Checkboxes      =   -1  'True
         Appearance      =   1
      End
   End
End
Attribute VB_Name = "frmAutoNumeracao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CarregaModulo()

    Dim strSistema      As String
    Dim strSql          As String
    Dim adoResultado    As ADODB.Recordset
    Dim intRelacao      As Integer
    
    Select Case UCase(App.ProductName)
        Case "TRIBUTARIO"
            strSistema = "J"
        Case "ORCAMENTARIO"
            strSistema = "F"
        Case "FROTA"
            strSistema = "B"
        Case "RH"
            strSistema = "I"
        Case "LEGISLACAO"
            strSistema = "C"
        Case "OUVIDORIA"
            strSistema = "M"
        Case "COMPRAS"
            strSistema = "A"
        Case "PATRIMONIO"
            strSistema = "G"
        Case "MATERIAL"
            strSistema = "E"
        Case "PROTOCOLO"
            strSistema = "H"
        Case "SEGURANCA"
            strSistema = "L"
        Case "MENOR"
            strSistema = "D"
        Case "GERENCIAL"
            strSistema = "N"
    End Select
    
    strSql = ""
    strSql = strSql & "SELECT strCodItem, strItem, bitAutoNumeracao "
    strSql = strSql & "FROM " & gstrItens
    strSql = strSql & " WHERE UPPER(" & strSUBSTRING & "(strCodItem,1,1)) = '" & strSistema & "' AND "
    strSql = strSql & " (bitAutoNumeracao <> 3 OR blnPermissao = 0)"
    strSql = strSql & "ORDER BY strCodItem"
    
    Set gobjBanco = New clsBanco
    
    TreeModulos.Nodes.Clear
    
    If gobjBanco.CriaADO(strSql, 0, adoResultado) Then
        While Not adoResultado.EOF
            
            If Len(adoResultado("strCodItem")) = 1 Then
                TreeModulos.Nodes.Add , intRelacao, adoResultado("strCodItem"), adoResultado("strItem")
                TreeModulos.Nodes(TreeModulos.Nodes.Count).Expanded = True
            Else
                If Len(adoResultado("strCodItem")) <> Len(TreeModulos.Nodes(TreeModulos.Nodes.Count).Key) Then
                    TreeModulos.Nodes.Add Left(Trim(adoResultado("strCodItem")), Len(Trim(adoResultado("strCodItem"))) - 1), tvwChild, Trim(adoResultado("strCodItem")), adoResultado("strItem")
                Else
                    TreeModulos.Nodes.Add TreeModulos.Nodes(TreeModulos.Nodes.Count).Key, tvwNext, adoResultado("strCodItem"), adoResultado("strItem")
                    
                End If
                If adoResultado("bitAutoNumeracao") = 1 Then
                    TreeModulos.Nodes(TreeModulos.Nodes.Count).Checked = True
                    TreeModulos_NodeCheck TreeModulos.Nodes(TreeModulos.Nodes.Count)
                    TreeModulos.Nodes(TreeModulos.Nodes.Count).BackColor = vbWhite
                ElseIf adoResultado("bitAutoNumeracao") = 2 Then
                    TreeModulos.Nodes(TreeModulos.Nodes.Count).Checked = True
                    TreeModulos.Nodes(TreeModulos.Nodes.Count).BackColor = vbMenuBar
                    TreeModulos_NodeCheck TreeModulos.Nodes(TreeModulos.Nodes.Count)
                Else
                    TreeModulos.Nodes(TreeModulos.Nodes.Count).BackColor = vbWhite
                End If
            End If
            adoResultado.MoveNext
        Wend
    End If
    
End Sub

Private Sub Form_Activate()
    gintCodSeguranca = 1007
     
'        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrNovo, gstrSalvar, gstrDeletar, gstrAplicar
'    Else
'        HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrDeletar
'    End If''

    'If mobjAux Is Nothing Then
    '    HabilitaDesabilitaBotao1 False, gstrMnuArquivo, gstrAplicar
    'Else'

        'HabilitaDesabilitaBotao1 True, gstrMnuArquivo, gstrAplicar
    'End If

End Sub

Private Sub Form_Load()
    CarregaModulo
End Sub

Private Sub TreeModulos_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    Debug.Print TreeModulos.SelectedItem.Key
    
    If Button = vbRightButton And Len(TreeModulos.SelectedItem.Key) > 2 Then
        Set TreeModulos.SelectedItem = TreeModulos.HitTest(x, y)
        
        If TreeModulos.SelectedItem.BackColor = vbWhite Then
            TreeModulos.SelectedItem.BackColor = vbMenuBar
            TreeModulos.SelectedItem.Checked = True
        Else
            TreeModulos.SelectedItem.BackColor = vbWhite
            TreeModulos.SelectedItem.Checked = False
        End If
    End If
    
End Sub

Private Sub TreeModulos_NodeCheck(ByVal Node As MSComctlLib.Node)
    Dim objNode         As Node
    Dim intParentCount  As Integer
    
    If Not Node.Checked Then Node.BackColor = vbWhite
    
    For Each objNode In TreeModulos.Nodes
        If InStr(objNode.FullPath, Node.FullPath) Then objNode.Checked = Node.Checked
    Next
        
    If Not Node.Parent Is Nothing Then
        For Each objNode In TreeModulos.Nodes
            If (objNode.FullPath <> Node.Parent.FullPath) And InStr(objNode.FullPath, Node.Parent.FullPath) And objNode.Checked = True Then
                intParentCount = intParentCount + 1
                Exit For
            End If
        Next
        
        If intParentCount = 0 Then
            Node.Parent.Checked = False
            
            intParentCount = 0
            
            If Not Node.Parent.Parent Is Nothing Then
                For Each objNode In TreeModulos.Nodes
                    If (objNode.FullPath <> Node.Parent.Parent.FullPath) And InStr(objNode.FullPath, Node.Parent.Parent.FullPath) And objNode.Checked = True Then
                        intParentCount = intParentCount + 1
                        Exit For
                    End If
                Next
                
                If intParentCount = 0 Then Node.Parent.Parent.Checked = False
            End If
        Else
            Node.Parent.Checked = True
            If Not Node.Parent.Parent Is Nothing Then
                Node.Parent.Parent.Checked = True
            End If
        End If
    End If
End Sub

Public Sub MantemForm(ByVal strModoOperacao As String)
    
    Dim objNode         As Node
    Dim strSql          As String

    If strModoOperacao = gstrSalvar Then
        For Each objNode In TreeModulos.Nodes
            
            If Len(objNode.Key) > 2 Then
            
                Set gobjBanco = New clsBanco
                
                strSql = "UPDATE " & gstrItens
                strSql = strSql & " SET bitAutoNumeracao = " & IIf(objNode.BackColor = vbWhite, Abs(CInt(objNode.Checked)), 2)
                strSql = strSql & " WHERE strCodItem ='" & objNode.Key & "'"
                
                gobjBanco.Execute strSql
                
            End If
        Next
    End If
    
End Sub


