VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form StatusCotacao 
   Caption         =   "Status da Cotação"
   ClientHeight    =   5820
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16875
   LinkTopic       =   "Form1"
   ScaleHeight     =   5820
   ScaleWidth      =   16875
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton botaoSair 
      Caption         =   "Sair"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7470
      TabIndex        =   1
      Top             =   5040
      Width           =   1935
   End
   Begin MSFlexGridLib.MSFlexGrid GradeEmpresa 
      Height          =   4455
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   16095
      _ExtentX        =   28390
      _ExtentY        =   7858
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "StatusCotacaO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim consultaPaginasCotacao As ADODB.Recordset
Dim consultaUsuario As ADODB.Recordset
Dim usuarios()
Dim id As Integer
Dim pagina As Integer

Private Sub botaoSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()

    

    ConexaoMysql.Open "Driver={MySQL ODBC 3.51 Driver};user=" + username + ";password=" + password + ";database=" + database + ";server=" + host + ";option=" & (1 + 2 + 8 + 32 + 2048 + 16384)
    Set consultaPaginasCotacao = New ADODB.Recordset
    
    consulta = "SELECT pagina From produtosRep Where id_cotacao = " & numEdicaoCotacao & " GROUP BY pagina ORDER BY pagina;"
    consultaPaginasCotacao.Open consulta, ConexaoMysql, adOpenStatic, adLockOptimistic, adCmdText
    
    If consultaPaginasCotacao.RecordCount > 0 Then
        
        consultaPaginasCotacao.MoveLast
        consultaPaginasCotacao.MoveFirst
        
        quant = consultaPaginasCotacao.RecordCount
        GradeEmpresa.Cols = quant + 1
        cabecalho = "|^Usuario                                          "
        For i = 1 To quant
            cabecalho = cabecalho + "|^" & i & "     "
        Next
        
        GradeEmpresa.FormatString = cabecalho
        
        consultaPaginasCotacao.Close
    
        'consulta = "Select u.usuario, u.id from usuarios u where"
        consulta = "SELECT usuarios.usuario,usuarios.id From produtosRep,usuarios Where id_cotacao = " & numEdicaoCotacao & " and usuarios.id = idRepresentante GROUP BY idRepresentante ORDER BY usuarios.usuario"
        Set consultaUsuario = New ADODB.Recordset
        consultaUsuario.Open consulta, ConexaoMysql, adOpenStatic, adLockOptimistic, adCmdText
        
        If consultaUsuario.RecordCount > 0 Then
        
            ReDim usuarios(consultaUsuario.RecordCount)
            
            GradeEmpresa.Rows = consultaUsuario.RecordCount + 1
            
            For laco = 0 To consultaUsuario.RecordCount - 1
                usuarios(laco) = consultaUsuario.Fields("id")
                GradeEmpresa.TextMatrix(laco + 1, 1) = consultaUsuario.Fields("usuario")
                consultaUsuario.MoveNext
            Next
        
        
        End If
        
        
        consultaUsuario.Close
        
        For coluna = 0 To GradeEmpresa.Cols - 3
        
            For linha = 0 To GradeEmpresa.Rows - 2
                GradeEmpresa.TextMatrix(linha + 1, coluna + 2) = " "
            Next
        
        Next
    
        Set consultaPaginasCotacao = New ADODB.Recordset
        'consulta = "SELECT idrepresentante, pagina From produtosRep Where id_cotacao = " & numEdicaoCotacao & " GROUP BY idrepresentante, pagina ORDER BY pagina;"
        consulta = "SELECT idrepresentante, pagina From produtosRep Where id_cotacao = " & numEdicaoCotacao & " AND precoProduto <> 0  GROUP BY idrepresentante, pagina ORDER BY pagina"
        consultaPaginasCotacao.Open consulta, ConexaoMysql, adOpenStatic, adLockOptimistic, adCmdText
        
        If consultaPaginasCotacao.RecordCount > 0 Then
            
            consultaPaginasCotacao.MoveLast
            consultaPaginasCotacao.MoveFirst
            
            For laco = 0 To consultaPaginasCotacao.RecordCount - 1
                
                id = consultaPaginasCotacao.Fields("idrepresentante")
                pagina = consultaPaginasCotacao.Fields("pagina")
                For i = 0 To GradeEmpresa.Rows - 1
                    If (id = usuarios(i)) Then
                        GradeEmpresa.TextMatrix(i + 1, pagina + 1) = "X"
                        i = GradeEmpresa.Rows
                    End If
                Next
                
                consultaPaginasCotacao.MoveNext
            Next
        End If
    End If

End Sub
