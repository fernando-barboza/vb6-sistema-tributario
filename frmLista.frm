VERSION 5.00
Begin VB.Form frmLista 
   BorderStyle     =   0  'None
   ClientHeight    =   3345
   ClientLeft      =   3030
   ClientTop       =   2730
   ClientWidth     =   5610
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3345
   ScaleWidth      =   5610
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOpcao1 
      Height          =   405
      Left            =   3930
      TabIndex        =   2
      Top             =   2850
      Width           =   1575
   End
   Begin VB.ListBox lstLista 
      Height          =   2205
      Left            =   60
      TabIndex        =   1
      Top             =   540
      Width           =   5475
   End
   Begin VB.Shape Shape1 
      Height          =   3345
      Left            =   0
      Top             =   0
      Width           =   5595
   End
   Begin VB.Label lblTitulo 
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   5445
   End
End
Attribute VB_Name = "frmLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOpcao1_Click()

    If lstLista.SelCount > 0 Then
        
        gintCodigoItemLista = lstLista.ItemData(lstLista.ListIndex)
        gstrDescricaoItemLista = lstLista.List(lstLista.ListIndex)
            
        Unload Me
        
    Else
        MsgBox "É necessário selecionar algum item da lista.", vbOKOnly, "Mensagem ao Usuário"
    End If
    
End Sub

Private Sub Form_Load()
    
    gintCodigoItemLista = 0
    gstrDescricaoItemLista = Space$(0)
    
End Sub

