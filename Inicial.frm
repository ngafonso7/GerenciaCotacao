VERSION 5.00
Begin VB.Form Inicial 
   Caption         =   "Gerencia Cota��o Online"
   ClientHeight    =   2835
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   8055
   LinkTopic       =   "Form1"
   ScaleHeight     =   2835
   ScaleWidth      =   8055
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Gerencia Cota��o Online"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   29.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1455
      Left            =   1200
      TabIndex        =   0
      Top             =   600
      Width           =   5535
   End
   Begin VB.Menu mnuusuarios 
      Caption         =   "Usu�rios"
      Begin VB.Menu mnucadastrousuario 
         Caption         =   "Cadastro"
      End
      Begin VB.Menu mnuedicaoexclusaousuario 
         Caption         =   "Edi��o/Exclus�o"
      End
   End
   Begin VB.Menu mnucotacao 
      Caption         =   "Cota��o"
      Begin VB.Menu mnugerenciarcotacao 
         Caption         =   "Gerenciar"
      End
      Begin VB.Menu mnuUploadPedidos 
         Caption         =   "Upload de Pedidos"
      End
   End
   Begin VB.Menu mnusair 
      Caption         =   "Sair"
   End
End
Attribute VB_Name = "Inicial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ConfigGerencia As database


Private Sub Form_Load()
    
    Inicial.Show
    
    DoEvents
    'ChDir "c:\vbronny1"
    CaminhoConfigGerencia = CurDir() & "\ConfigGerenciaAvance.mdb"
    'CaminhoConfigGerencia = "C:\Gerencia Avance\ConfigGerenciaAvance.mdb"
    'CaminhoConfigGerencia = "C:\Program Files (x86)\Microsoft Visual Studio\VB98\ConfigGerenciaAvance.mdb"
    Set ConfigGerencia = OpenDatabase(CaminhoConfigGerencia)
    Set RegistroConfigura��oGerencia = ConfigGerencia.OpenRecordset("Configura��es")
    RegistroConfigura��oGerencia.MoveFirst
    caminhodados = RegistroConfigura��oGerencia.Fields("CaminhoDados")
    caminhofornecedor = RegistroConfigura��oGerencia.Fields("CaminhoNotaFiscal")
    
    
    
    host = RegistroConfigura��oGerencia.Fields("HostMysqlCotacao")
    username = RegistroConfigura��oGerencia.Fields("UsernameMysqlCotacao")
    password = RegistroConfigura��oGerencia.Fields("PasswordMysqlCotacao")
    database = RegistroConfigura��oGerencia.Fields("DatabaseMysqlCotacao")
    
    
    
    
    
End Sub

Private Sub mnucadastrousuario_Click()
    CadastroUsuario.Show 1
End Sub

Private Sub mnuedicaoexclusaousuario_Click()
    EdicaoExclusaoUsuario.Show 1
End Sub

Private Sub mnugerenciarcotacao_Click()
    Cotacao.Show 1
End Sub

Private Sub mnusair_Click()
    Unload Me
End Sub

Private Sub mnuUploadPedidos_Click()
    Load UploadPedidos
    UploadPedidos.Show 1
End Sub
