VERSION 5.00
Begin VB.Form CadastroUsuario 
   Caption         =   "Cadastro Usuário"
   ClientHeight    =   4770
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7065
   LinkTopic       =   "Form1"
   ScaleHeight     =   4770
   ScaleWidth      =   7065
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton botaoSair 
      Caption         =   "Sair"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   2445
      TabIndex        =   10
      Top             =   3915
      Width           =   2175
   End
   Begin VB.CommandButton botaoCancelar 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3727
      TabIndex        =   9
      Top             =   3120
      Width           =   2175
   End
   Begin VB.CommandButton botaoGravar 
      Caption         =   "Gravar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   1162
      TabIndex        =   8
      Top             =   3120
      Width           =   2175
   End
   Begin VB.ComboBox listaEmpresa 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2250
      TabIndex        =   3
      Text            =   "Selecione"
      Top             =   2110
      Width           =   4500
   End
   Begin VB.TextBox txtRepresentante 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2250
      TabIndex        =   2
      Top             =   1490
      Width           =   4500
   End
   Begin VB.TextBox txtSenha 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2250
      TabIndex        =   1
      Top             =   870
      Width           =   2000
   End
   Begin VB.TextBox txtUsuario 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2250
      TabIndex        =   0
      Top             =   250
      Width           =   2000
   End
   Begin VB.Label Label4 
      Caption         =   "Empresa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   370
      Left            =   150
      TabIndex        =   7
      Top             =   2110
      Width           =   2000
   End
   Begin VB.Label Label3 
      Caption         =   "Representante"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   370
      Left            =   150
      TabIndex        =   6
      Top             =   1490
      Width           =   2000
   End
   Begin VB.Label Label2 
      Caption         =   "Senha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   370
      Left            =   150
      TabIndex        =   5
      Top             =   870
      Width           =   2000
   End
   Begin VB.Label Label1 
      Caption         =   "Usuario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   370
      Left            =   150
      TabIndex        =   4
      Top             =   250
      Width           =   2000
   End
End
Attribute VB_Name = "CadastroUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim dbFornecedores As database
Dim tabelaFornecedores As Recordset
Dim consultaIdUsuario As ADODB.Recordset
Dim inserirEmpresa As ADODB.Recordset
Dim Fornecedores
Dim idFornecedores
Dim idUsuario As Integer

Private Sub botaoCancelar_Click()
    res = MsgBox("Deseja mesmo cancelar?", vbYesNo, "Cancelar")
    If res = vbYes Then
        txtUsuario.Text = ""
        txtSenha.Text = ""
        txtRepresentante.Text = ""
        listaEmpresa.ListIndex = -1
        listaEmpresa.Text = "Selecione"
        botaoSair.SetFocus
    Else
        txtUsuario.SetFocus
    End If
End Sub

Private Sub botaoGravar_Click()
    If (txtUsuario.Text <> "" And txtSenha.Text <> "" And txtRepresentante.Text <> "" And listaEmpresa.ListIndex <> -1) Then
        Usuario = txtUsuario.Text
        senha = txtSenha.Text
        rep = txtRepresentante.Text
        Forn = Fornecedores(listaEmpresa.ListIndex)
        idforn = idFornecedores(listaEmpresa.ListIndex)
        'Mysql
        '****************************************
        Set ConexaoMysql = New ADODB.Connection
        ConexaoMysql.ConnectionTimeout = 60
        ConexaoMysql.CommandTimeout = 400
        ConexaoMysql.CursorLocation = adUseClient
        ConexaoMysql.Open "Driver={MySQL ODBC 3.51 Driver};user=" + username + ";password=" + password + ";database=" + database + ";server=" + host + ";option=" & (1 + 2 + 8 + 32 + 2048 + 16384)
        Set inserirFornecedores = New ADODB.Recordset
        consulta = "INSERT INTO usuarios values(" & idUsuario & ",'" + Usuario + "','" + senha + "','" + rep + "',1);"
        inserirFornecedores.Open consulta, ConexaoMysql, adOpenStatic, adLockOptimistic, adCmdText
        
        Set inserirEmpresa = New ADODB.Recordset
        consulta = "INSERT INTO empresa values(null,'" + Forn + "','" & idUsuario & "','" & idforn & "');"
        inserirEmpresa.Open consulta, ConexaoMysql, adOpenStatic, adLockOptimistic, adCmdText
        
        
        ConexaoMysql.Close
        
        idUsuario = idUsuario + 1
        
        MsgBox "Usuário inserido com sucesso !!!", , "Cadastro de Usuário"
        
        txtUsuario.Text = ""
        txtSenha.Text = ""
        txtRepresentante.Text = ""
        listaEmpresa.ListIndex = -1
        listaEmpresa.Text = "Selecione"
    
        botaoSair.SetFocus
        
    End If
    
End Sub

Private Sub botaoSair_Click()
    Unload Me
End Sub



Private Sub Form_Load()

    Set dbFornecedores = OpenDatabase(caminhofornecedor)
    consulta = "Select [Razão Social],[Código] from Fornecedores order by [Razão Social];"
    Set tabelaFornecedores = dbFornecedores.OpenRecordset(consulta)
    If (tabelaFornecedores.RecordCount > 0) Then
        tabelaFornecedores.MoveLast
        tabelaFornecedores.MoveFirst
        ReDim Fornecedores(tabelaFornecedores.RecordCount)
        ReDim idFornecedores(tabelaFornecedores.RecordCount)
        For laço = 0 To tabelaFornecedores.RecordCount - 1
            listaEmpresa.AddItem (tabelaFornecedores.Fields("Razão Social"))
            Fornecedores(laço) = tabelaFornecedores.Fields("Razão Social")
            idFornecedores(laço) = tabelaFornecedores.Fields("Código")
            tabelaFornecedores.MoveNext
        Next
        
    End If
    
    Set ConexaoMysql = New ADODB.Connection
    ConexaoMysql.ConnectionTimeout = 60
    ConexaoMysql.CommandTimeout = 400
    ConexaoMysql.CursorLocation = adUseClient
    ConexaoMysql.Open "Driver={MySQL ODBC 3.51 Driver};user=" + username + ";password=" + password + ";database=" + database + ";server=" + host + ";option=" & (1 + 2 + 8 + 32 + 2048 + 16384)
    Set consultaIdUsuario = New ADODB.Recordset
    consulta = "SELECT Max(id) as novoID from usuarios"
    consultaIdUsuario.Open consulta, ConexaoMysql, adOpenStatic, adLockOptimistic, adCmdText
    If (consultaIdUsuario.RecordCount > 0) Then
        If (consultaIdUsuario.Fields("novoID") <> Empty) Then
            idUsuario = consultaIdUsuario.Fields("novoID") + 1
        Else
            idUsuario = 1
        End If
    Else
        idUsuario = 1
    End If
    ConexaoMysql.Close
    
    
        
    
End Sub

Private Sub listaEmpresa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If listaEmpresa.ListIndex <> -1 Then
            botaoGravar.SetFocus
        End If
    End If
End Sub

Private Sub txtRepresentante_Change()
    If txtUsuario.Text <> "" Or txtSenha.Text <> "" Or txtRepresentante.Text <> "" Then
        botaoSair.Enabled = False
    Else
        botaoSair.Enabled = True
    End If
End Sub

Private Sub txtRepresentante_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtRepresentante.Text <> "" Then
            listaEmpresa.SetFocus
        End If
    End If
End Sub

Private Sub txtSenha_Change()
    If txtUsuario.Text <> "" Or txtSenha.Text <> "" Or txtRepresentante.Text <> "" Then
        botaoSair.Enabled = False
    Else
        botaoSair.Enabled = True
    End If
End Sub

Private Sub txtSenha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtSenha.Text <> "" Then
            txtRepresentante.SetFocus
        End If
    End If
End Sub

Private Sub txtUsuario_Change()
    If txtUsuario.Text <> "" Or txtSenha.Text <> "" Or txtRepresentante.Text <> "" Then
        botaoSair.Enabled = False
    Else
        botaoSair.Enabled = True
    End If
End Sub

Private Sub txtUsuario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtUsuario.Text <> "" Then
            txtSenha.SetFocus
        End If
    End If
    
    
End Sub
