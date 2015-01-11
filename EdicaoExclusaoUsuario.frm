VERSION 5.00
Begin VB.Form EdicaoExclusaoUsuario 
   Caption         =   "Edição/Exclusão Usuarios"
   ClientHeight    =   6735
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10245
   LinkTopic       =   "Form1"
   ScaleHeight     =   6735
   ScaleWidth      =   10245
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox checkLiberado 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4000
      TabIndex        =   16
      Top             =   3435
      Width           =   2415
   End
   Begin VB.CommandButton botaoEditar 
      Caption         =   "Editar"
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
      Left            =   2955
      TabIndex        =   5
      Top             =   4545
      Width           =   2175
   End
   Begin VB.CommandButton botaoExcluir 
      Caption         =   "Excluir"
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
      Left            =   5335
      TabIndex        =   8
      Top             =   4545
      Width           =   2175
   End
   Begin VB.ComboBox listaRepresentantes 
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
      Left            =   4000
      TabIndex        =   0
      Text            =   "Selecione"
      Top             =   200
      Width           =   4500
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
      Left            =   4000
      TabIndex        =   1
      Top             =   960
      Width           =   2000
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
      Left            =   4000
      TabIndex        =   2
      Top             =   1575
      Width           =   2000
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
      Left            =   4000
      TabIndex        =   3
      Top             =   2205
      Width           =   4500
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
      Left            =   4000
      TabIndex        =   4
      Text            =   "Selecione"
      Top             =   2820
      Width           =   4500
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
      Left            =   2955
      TabIndex        =   6
      Top             =   5265
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
      Left            =   5355
      TabIndex        =   7
      Top             =   5265
      Width           =   2175
   End
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
      Left            =   4155
      TabIndex        =   9
      Top             =   5985
      Width           =   2175
   End
   Begin VB.Label Label6 
      Caption         =   "Liberado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   15
      Top             =   3435
      Width           =   1995
   End
   Begin VB.Label Label5 
      Caption         =   "Selecione o Representante"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   150
      TabIndex        =   14
      Top             =   200
      Width           =   3795
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
      Height          =   375
      Left            =   1800
      TabIndex        =   13
      Top             =   960
      Width           =   1995
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
      Height          =   375
      Left            =   1800
      TabIndex        =   12
      Top             =   1575
      Width           =   1995
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
      Height          =   375
      Left            =   1800
      TabIndex        =   11
      Top             =   2205
      Width           =   1995
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
      Height          =   375
      Left            =   1800
      TabIndex        =   10
      Top             =   2820
      Width           =   1995
   End
End
Attribute VB_Name = "EdicaoExclusaoUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim consultaIdUsuario As ADODB.Recordset
Dim consultaEmpresa As ADODB.Recordset
Dim exclusaoUsuario As ADODB.Recordset
Dim edicaoUsuario As ADODB.Recordset
Dim edicaoEmpresa As ADODB.Recordset

Dim idRepresentante
Dim usuarioRepresentante
Dim senhaRepresentante
Dim idEmpresa
Dim nomeEmpresa
Dim liberado

Dim empresaSelecionada



Private Sub botaoCancelar_Click()

    txtUsuario.Text = ""
    txtSenha.Text = ""
    txtRepresentante.Text = ""
    checkLiberado.Enabled = False
    listaEmpresa.ListIndex = -1
    listaEmpresa.Text = "Selecione"
    txtUsuario.Enabled = False
    txtSenha.Enabled = False
    txtRepresentante.Enabled = False
    listaEmpresa.Enabled = False
    listaRepresentantes.Enabled = True
    botaoGravar.Enabled = False
    botaoCancelar.Enabled = False
    botaoEditar.Enabled = False
    botaoExcluir.Enabled = False
    listaRepresentantes.SetFocus
    
End Sub

Private Sub botaoEditar_Click()
    txtUsuario.Enabled = True
    txtSenha.Enabled = True
    txtRepresentante.Enabled = True
    listaEmpresa.Enabled = True
    checkLiberado.Enabled = True
    listaRepresentantes.Enabled = False
    botaoGravar.Enabled = True
    botaoCancelar.Enabled = True
    botaoEditar.Enabled = False
    botaoExcluir.Enabled = False
    txtUsuario.SetFocus
    
End Sub

Private Sub botaoExcluir_Click()
    res = MsgBox("Deseja mesmo excluir?", vbYesNo)
    If (res = vbYes) Then
        Set ConexaoMysql = New ADODB.Connection
        ConexaoMysql.ConnectionTimeout = 60
        ConexaoMysql.CommandTimeout = 400
        ConexaoMysql.CursorLocation = adUseClient
        ConexaoMysql.Open "Driver={MySQL ODBC 3.51 Driver};user=" + username + ";password=" + password + ";database=" + database + ";server=" + host + ";option=" & (1 + 2 + 8 + 32 + 2048 + 16384)
        Set exclusaoUsuario = New ADODB.Recordset
        consulta = "DELETE FROM usuarios WHERE id = " & idRepresentante(listaRepresentantes.ListIndex) & ";"
        exclusaoUsuario.Open consulta, ConexaoMysql, adOpenStatic, adLockOptimistic, adCmdText
        'consulta = "DELETE * FROM usuario WHERE id = " & idRepresentante & ";"
        'exclusaoUsuario.Open consulta, ConexaoMysql, adOpenStatic, adLockOptimistic, adCmdText
        MsgBox "Usuário excluido com sucesso !"
        ConexaoMysql.Close
        
        txtUsuario.Text = ""
        txtSenha.Text = ""
        txtRepresentante.Text = ""
        listaEmpresa.Clear
        listaRepresentantes.Clear
        Call Form_Activate
    End If
    
End Sub

Private Sub botaoGravar_Click()
    If (txtUsuario.Text <> "" And txtSenha.Text <> "" And txtRepresentante.Text <> "" And listaEmpresa.ListIndex <> -1) Then
        usu = txtUsuario.Text
        sen = txtSenha.Text
        rep = txtRepresentante.Text
        emp = idEmpresa(listaEmpresa.ListIndex)
        If checkLiberado.Value = 1 Then
            lib = "1"
        Else
            lib = "0"
        End If
        Set ConexaoMysql = New ADODB.Connection
        ConexaoMysql.ConnectionTimeout = 60
        ConexaoMysql.CommandTimeout = 400
        ConexaoMysql.CursorLocation = adUseClient
        ConexaoMysql.Open "Driver={MySQL ODBC 3.51 Driver};user=" + username + ";password=" + password + ";database=" + database + ";server=" + host + ";option=" & (1 + 2 + 8 + 32 + 2048 + 16384)
        Set edicaoUsuario = New ADODB.Recordset
        consulta = "UPDATE usuarios SET usuario ='" + usu + "',senha = '" + sen + "',nomeRep='" + rep + "',liberado = '" + lib + "' WHERE id =" & idRepresentante(listaRepresentantes.ListIndex) & ";"
        edicaoUsuario.Open consulta, ConexaoMysql, adOpenStatic, adLockOptimistic, adCmdText
        
        Set edicaoEmpresa = New ADODB.Recordset
        consulta = "DELETE FROM empresa WHERE idEmpresa =" & empresaSelecionada & ";"
        edicaoEmpresa.Open consulta, ConexaoMysql, adOpenStatic, adLockOptimistic, adCmdText
        
        consulta = "INSERT INTO empresa values (null,'" + nomeEmpresa(listaEmpresa.ListIndex) + "'," & idRepresentante(listaRepresentantes.ListIndex) & "," & idEmpresa(listaEmpresa.ListIndex) & ");"
        edicaoEmpresa.Open consulta, ConexaoMysql, adOpenStatic, adLockOptimistic, adCmdText
        
        MsgBox "Usuario alterado com sucesso !"
        
        txtUsuario.Text = ""
        txtSenha.Text = ""
        txtRepresentante.Text = ""
        listaEmpresa.Clear
        listaRepresentantes.Enabled = True
        listaRepresentantes.Clear
        
        ConexaoMysql.Close
        
        Call Form_Activate
        
    Else
        MsgBox "Preencha todos os campos !"
    End If
End Sub

Private Sub botaoSair_Click()
    Unload Me
End Sub

Private Sub checkLiberado_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        botaoGravar.SetFocus
    End If
End Sub

Private Sub Form_Activate()

    Set ConexaoMysql = New ADODB.Connection
    ConexaoMysql.ConnectionTimeout = 60
    ConexaoMysql.CommandTimeout = 400
    ConexaoMysql.CursorLocation = adUseClient
    ConexaoMysql.Open "Driver={MySQL ODBC 3.51 Driver};user=" + username + ";password=" + password + ";database=" + database + ";server=" + host + ";option=" & (1 + 2 + 8 + 32 + 2048 + 16384)
    Set consultaIdUsuario = New ADODB.Recordset
    consulta = "SELECT * from usuarios order by nomeRep"
    consultaIdUsuario.Open consulta, ConexaoMysql, adOpenStatic, adLockOptimistic, adCmdText
    If (consultaIdUsuario.RecordCount > 0) Then
        consultaIdUsuario.MoveLast
        consultaIdUsuario.MoveFirst
        ReDim idRepresentante(consultaIdUsuario.RecordCount)
        ReDim usuarioRepresentante(consultaIdUsuario.RecordCount)
        ReDim senhaRepresentante(consultaIdUsuario.RecordCount)
        ReDim liberado(consultaIdUsuario.RecordCount)
        For laço = 0 To consultaIdUsuario.RecordCount - 1
            idRepresentante(laço) = consultaIdUsuario.Fields("id")
            usuarioRepresentante(laço) = consultaIdUsuario.Fields("usuario")
            senhaRepresentante(laço) = consultaIdUsuario.Fields("senha")
            If consultaIdUsuario.Fields("liberado") = "1" Then
                liberado(laço) = True
            Else
                liberado(laço) = False
            End If
            listaRepresentantes.AddItem (consultaIdUsuario.Fields("nomeRep"))
            consultaIdUsuario.MoveNext
        Next
        listaRepresentantes.ListIndex = -1
        listaRepresentantes.Text = "Selecione"
        listaRepresentantes.SetFocus
    Else
        MsgBox "Não existe(m) usuario(s) cadastrado(s) !"
        Unload Me
    End If
    txtUsuario.Enabled = False
    txtSenha.Enabled = False
    txtRepresentante.Enabled = False
    listaEmpresa.Enabled = False
    checkLiberado.Enabled = False
    
    botaoGravar.Enabled = False
    botaoCancelar.Enabled = False
    botaoEditar.Enabled = False
    botaoExcluir.Enabled = False
    ConexaoMysql.Close
End Sub


Private Sub listaEmpresa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If listaEmpresa.ListIndex <> -1 Then
            checkLiberado.SetFocus
        End If
    End If
End Sub

Private Sub listaRepresentantes_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        If listaRepresentantes.ListIndex <> -1 Then
            posEmpresa = -1
            Set ConexaoMysql = New ADODB.Connection
            ConexaoMysql.ConnectionTimeout = 60
            ConexaoMysql.CommandTimeout = 400
            ConexaoMysql.CursorLocation = adUseClient
            ConexaoMysql.Open "Driver={MySQL ODBC 3.51 Driver};user=" + username + ";password=" + password + ";database=" + database + ";server=" + host + ";option=" & (1 + 2 + 8 + 32 + 2048 + 16384)
            Set consultaEmpresa = New ADODB.Recordset
            consulta = "SELECT * FROM empresa WHERE usuario_idusuario = " & idRepresentante(listaRepresentantes.ListIndex) & ";"
            consultaEmpresa.Open consulta, ConexaoMysql, adOpenStatic, adLockOptimistic, adCmdText
            If (consultaEmpresa.RecordCount > 0) Then
                empr = consultaEmpresa.Fields("Nome")
            Else
                empr = ""
            End If
            Set dbFornecedores = OpenDatabase(caminhofornecedor)
            consulta = "Select [Razão Social],[Código] from Fornecedores order by [Razão Social];"
            Set tabelaFornecedores = dbFornecedores.OpenRecordset(consulta)
            If (tabelaFornecedores.RecordCount > 0) Then
                tabelaFornecedores.MoveLast
                tabelaFornecedores.MoveFirst
                ReDim nomeEmpresa(tabelaFornecedores.RecordCount)
                ReDim idEmpresa(tabelaFornecedores.RecordCount)
                For laço = 0 To tabelaFornecedores.RecordCount - 1
                    listaEmpresa.AddItem (tabelaFornecedores.Fields("Razão Social"))
                    nomeEmpresa(laço) = tabelaFornecedores.Fields("Razão Social")
                    idEmpresa(laço) = tabelaFornecedores.Fields("Código")
                    If empr = tabelaFornecedores.Fields("Razão Social") Then
                        posEmpresa = laço
                    End If
                    tabelaFornecedores.MoveNext
                Next
            End If
            txtUsuario.Text = usuarioRepresentante(listaRepresentantes.ListIndex)
            txtSenha.Text = senhaRepresentante(listaRepresentantes.ListIndex)
            txtRepresentante = listaRepresentantes.Text
            If liberado(listaRepresentantes.ListIndex) Then
                checkLiberado.Value = 1
            Else
                checkLiberado.Value = 0
            End If
            
            If (posEmpresa = -1) Then
                listaEmpresa.ListIndex = -1
                listaEmpresa.Text = "Selecione"
                empresaSelecionada = 0
            Else
                listaEmpresa.ListIndex = posEmpresa
                empresaSelecionada = idEmpresa(posEmpresa)
            End If
            botaoEditar.Enabled = True
            botaoExcluir.Enabled = True
            botaoEditar.SetFocus
            ConexaoMysql.Close
        End If
    End If
    
    
End Sub


Private Sub txtRepresentante_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtRepresentante.Text <> Empty Then
            listaEmpresa.SetFocus
        End If
    End If
End Sub

Private Sub txtSenha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtSenha.Text <> Empty Then
            txtRepresentante.SetFocus
            txtRepresentante.SelStart = 0
            txtRepresentante.SelLength = Len(txtSenha.Text)
        End If
    End If
End Sub
Private Sub txtUsuario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtUsuario.Text <> Empty Then
            txtSenha.SetFocus
            txtSenha.SelStart = 0
            txtSenha.SelLength = Len(txtSenha.Text)
        End If
    End If
    
End Sub
