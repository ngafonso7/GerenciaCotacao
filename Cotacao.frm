VERSION 5.00
Begin VB.Form Cotacao 
   Caption         =   "Gerencia Cotação"
   ClientHeight    =   4665
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   ScaleHeight     =   4665
   ScaleWidth      =   7455
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton botaoStatus 
      Caption         =   "Status"
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
      Left            =   2820
      TabIndex        =   7
      Top             =   2280
      Width           =   1815
   End
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
      Left            =   2820
      TabIndex        =   6
      Top             =   3720
      Width           =   1815
   End
   Begin VB.CommandButton botaoEditar 
      Caption         =   "Editar"
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
      Left            =   2820
      TabIndex        =   5
      Top             =   3000
      Width           =   1815
   End
   Begin VB.CommandButton botaoEncerrar 
      Caption         =   "Encerrar"
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
      Left            =   3840
      TabIndex        =   4
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CommandButton botaoUpload 
      Caption         =   "Upload"
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
      Left            =   1800
      TabIndex        =   3
      Top             =   1560
      Width           =   1815
   End
   Begin VB.ComboBox listaCotacoes 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   3500
      TabIndex        =   1
      Text            =   "Carregando..."
      Top             =   200
      Width           =   2200
   End
   Begin VB.Label lblstatus 
      Alignment       =   2  'Center
      BackColor       =   &H80000007&
      Caption         =   "Status: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   200
      TabIndex        =   2
      Top             =   900
      Width           =   7100
   End
   Begin VB.Label Label1 
      Caption         =   "Numero da Cotação"
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
      Left            =   200
      TabIndex        =   0
      Top             =   200
      Width           =   3200
   End
End
Attribute VB_Name = "Cotacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dbCotacao As database
Dim consultaCotacao As Recordset
Dim inserirCotacao As ADODB.Recordset
Dim editarCotacao As Recordset

Dim usuariosCotacao As ADODB.Recordset

Dim encerraCotacao As ADODB.Recordset
Dim tabelaAuxiliar As Recordset

Dim numeroCotacao

Dim consultaIdCotacao As ADODB.Recordset


Private Sub botaoEditar_Click()
    res = MsgBox("Deseja mesmo editar a cotação nº " & numeroCotacao(listaCotacoes.ListIndex) & " ?", vbYesNo, "Edição de Cotação")
    If res = vbYes Then
        numEdicaoCotacao = numeroCotacao(listaCotacoes.ListIndex)
        EdicaoCotacao.Show 1
    End If
End Sub

Private Sub botaoEncerrar_Click()
    cont_erros = 0
    res = MsgBox("Deseja mesmo fazer o encerramento da cotação nº " & numeroCotacao(listaCotacoes.ListIndex) & ". Esta operação é irrevesível !", vbYesNo, "Encerramento de Cotação")
    If res = vbYes Then
    
        Path = "C:\Gerencia Avance\log_sql_" & numeroCotacao(listaCotacoes.ListIndex) & ".txt"
        
        Open Path For Output As #1

        
    
        Cotacao.MousePointer = vbHourglass
        
        numCot = numeroCotacao(listaCotacoes.ListIndex)
        
        ConexaoMysql.Open "Driver={MySQL ODBC 3.51 Driver};user=" + username + ";password=" + password + ";database=" + database + ";server=" + host + ";option=" & (1 + 2 + 8 + 32 + 2048 + 16384)
        
        Set encerraCotacao = New ADODB.Recordset
        consulta = "Delete From resultadoCotacao"
        encerraCotacao.Open consulta, ConexaoMysql, adOpenStatic, adLockOptimistic, adCmdText
        
        Print #1, "Apagando ultimo resultado da cotacao"
        
        consulta = "Call encerraCotacao(" & numCot & ");"
        encerraCotacao.Open consulta, ConexaoMysql, adOpenStatic, adLockOptimistic, adCmdText
        
        Print #1, "Preparando as informacoes para download"
        
        consulta = "Select * FROM resultadoCotacao;"
        encerraCotacao.Open consulta, ConexaoMysql, adOpenStatic, adLockOptimistic, adCmdText
        
        Print #1, "Download iniciado"
        Print #1, "#############################################"
        
        Set dbCotacao = OpenDatabase(caminhodados)
        
        If (encerraCotacao.RecordCount > 0) Then
            encerraCotacao.MoveLast
            encerraCotacao.MoveFirst
            consulta = "Delete From [Auxiliar Cotação];"
            dbCotacao.Execute (consulta)
            Set tabelaAuxiliar = dbCotacao.OpenRecordset("Auxiliar Cotação")
            Dim cod As Integer
            Dim emb As Integer
            Dim outra As String
            Dim emp As Integer
            For laço = 0 To encerraCotacao.RecordCount - 1
                tabelaAuxiliar.AddNew
                cod = encerraCotacao.Fields("codProduto")
                preco = encerraCotacao.Fields("precoProduto")
                emb = encerraCotacao.Fields("embProduto")
                If encerraCotacao.Fields("outraProduto") <> Empty Then
                    outra = encerraCotacao.Fields("outraProduto")
                Else
                    outra = ""
                End If
                emp = encerraCotacao.Fields("idEmpresa")
                Print #1, cod & " - " & preco & " - " & emb & " - " & outra & " - " & emp
                tabelaAuxiliar.Fields("codProduto") = cod
                tabelaAuxiliar.Fields("precoProduto") = preco
                tabelaAuxiliar.Fields("embProduto") = emb
                tabelaAuxiliar.Fields("outraProduto") = outra
                tabelaAuxiliar.Fields("idEmpresa") = emp
                tabelaAuxiliar.Update
                encerraCotacao.MoveNext
            Next
            tabelaAuxiliar.Close
        End If
        
        Print #1, "##############################################"
        Print #1, "Download de Dados concluido"
        Print #1, "##############################################"
        Print #1, "Iniciando processamento local"

        MsgBox "Cotação Importada. Processando os dados"
        encerraCotacao.Close
        Set tabelaAuxiliar = dbCotacao.OpenRecordset("Auxiliar Cotação")
        If (tabelaAuxiliar.RecordCount > 0) Then
            tabelaAuxiliar.MoveLast
            tabelaAuxiliar.MoveFirst
            For laço = 0 To tabelaAuxiliar.RecordCount - 1
                cod = tabelaAuxiliar.Fields("codProduto")
                preco = tabelaAuxiliar.Fields("precoProduto")
                emb = tabelaAuxiliar.Fields("embProduto")
                outra = tabelaAuxiliar.Fields("outraProduto")
                emp = tabelaAuxiliar.Fields("idEmpresa")
                
                consulta = "SELECT Cotação.[Preço à vista], Cotação.[Quantidade do produto por unidade de fornecimento], Cotação.[Marca do Produto] From Cotação WHERE (((Cotação.[Número da Cotação])=""" & Format(numCot, "00000") & """) AND ((Cotação.[Código do Produto])=""" & Format(cod, "00000") & """) AND ((Cotação.[Código do Fornecedor])=" & emp & "));"
                Set editarCotacao = dbCotacao.OpenRecordset(consulta)
                Print #1, consulta
                
                If (editarCotacao.RecordCount >= 1) Then
                
                    Print #1, "Produto encontrado. " & editarCotacao.Fields("[Preço à vista]") & " -> " & preco
                    editarCotacao.MoveLast
                    editarCotacao.MoveFirst
                    
                    editarCotacao.Edit
                    editarCotacao.Fields("[Preço à vista]") = preco
                    editarCotacao.Fields("[Quantidade do produto por unidade de fornecimento]") = emb
                    editarCotacao.Fields("[Marca do Produto]") = outra
                    editarCotacao.Update
                Else
                    Print #1, "Erro. Produto não encontrado"
                    cont_erros = cont_erros + 1
                End If
                tabelaAuxiliar.MoveNext
            Next
            If (ConexaoMysql.State = 1) Then
                ConexaoMysql.Close
                ConexaoMysql.Open "Driver={MySQL ODBC 3.51 Driver};user=" + username + ";password=" + password + ";database=" + database + ";server=" + host + ";option=" & (1 + 2 + 8 + 32 + 2048 + 16384)
            End If
            consulta = "UPDATE cotacoes SET encerrada=1 WHERE idCotacao = " & numCot & ";"
            encerraCotacao.Open consulta, ConexaoMysql, adOpenStatic, adLockOptimistic, adCmdText
            
            If (cont_erros > 0) Then
                MsgBox "Foram dectados " & cont_erros & "erro(s)."
            End If
            MsgBox "Dados processados, cotação encerrada com sucesso !"
            
            
            dataPC = InputBox("Informe a data da próxima cotação !", "Próxima cotação")
            If (ConexaoMysql.State = 1) Then
                ConexaoMysql.Close
                ConexaoMysql.Open "Driver={MySQL ODBC 3.51 Driver};user=" + username + ";password=" + password + ";database=" + database + ";server=" + host + ";option=" & (1 + 2 + 8 + 32 + 2048 + 16384)
            End If
            consulta = "DELETE FROM proximaCotacao"
            encerraCotacao.Open consulta, ConexaoMysql, adOpenStatic, adLockOptimistic, adCmdText
            
            consulta = "INSERT INTO proximaCotacao (data) value ('" & dataPC & "');"
            encerraCotacao.Open consulta, ConexaoMysql, adOpenStatic, adLockOptimistic, adCmdText
            
            MsgBox ("Cotação Encerrada")
            
            Set consultaIdCotacao = New ADODB.Recordset
            
            
            consulta = "SELECT * from cotacoes where idCotacao = " & numCot & ";"
            consultaIdCotacao.Open consulta, ConexaoMysql, adOpenStatic, adLockOptimistic, adCmdText
            If consultaIdCotacao.RecordCount > 0 Then
                If consultaIdCotacao.Fields("encerrada") Then
                    lblstatus.Caption = "Status: Cotação nº " & numCot & " está online e encerrada"
                    lblstatus.ForeColor = &HFF00&
                    lblstatus.Tag = "encerrada"
                    botaoUpload.Enabled = False
                    botaoEncerrar.Enabled = False
                Else
                    lblstatus.Caption = "Status: Cotação nº " & numCot & " está online e aberta"
                    lblstatus.ForeColor = &HFFFF&
                    lblstatus.Tag = "aberta"
                    botaoUpload.Enabled = True
                    botaoEncerrar.Enabled = True
                End If
            Else
                lblstatus.Caption = "Status: Cotação nº " & numCot & " não está online"
                lblstatus.ForeColor = &HFF&
                lblstatus.Tag = "nula"
                botaoUpload.Enabled = True
                botaoEncerrar.Enabled = False
            End If
            ConexaoMysql.Close
            
        Else
            MsgBox "Erro ao processar os dados !", vbCritical
            Print #1, "Erro ao processar os dados ! Tabela Auxiliar resultou em 0 registros"
        End If
        
        Cotacao.MousePointer = vbDefault
        dbCotacao.Close
        
        Close #1
        
    End If
End Sub

Private Sub botaoSair_Click()

    Unload Me
End Sub

Private Sub botaoStatus_Click()

    numEdicaoCotacao = numeroCotacao(listaCotacoes.ListIndex)
    Load StatusCotacaO
    StatusCotacaO.Show 1
    
End Sub

Private Sub botaoUpload_Click()
    res = MsgBox("Deseja mesmo fazer o upload da cotação nº " & numeroCotacao(listaCotacoes.ListIndex) & ". Isso apagará qualquer informação que estiver online !", vbYesNo, "Upload de Cotação")
    If res = vbYes Then
        Cotacao.MousePointer = 11
        numCot = numeroCotacao(listaCotacoes.ListIndex)
        ConexaoMysql.Open "Driver={MySQL ODBC 3.51 Driver};user=" + username + ";password=" + password + ";database=" + database + ";server=" + host + ";option=" & (1 + 2 + 8 + 32 + 2048 + 16384)
        
        Set limpaTudo = New ADODB.Recordset
        consulta = "Delete from produtos Where id_Cotacao = " & numCot & ";"
        limpaTudo.Open consulta, ConexaoMysql, adOpenStatic, adLockOptimistic, adCmdText
        
        consulta = "Delete from produtosRep ;"
        limpaTudo.Open consulta, ConexaoMysql, adOpenStatic, adLockOptimistic, adCmdText
        
        ConexaoMysql.Close
        
        Set dbCotacao = OpenDatabase(caminhodados)
        consulta = "Select [Número da cotação],[Código do produto],[Código de barras],[Descrição do produto],[Quantidade do produto por unidade de fornecimento] FROM Cotação Where [Número da cotação]=""" & Format(numeroCotacao(listaCotacoes.ListIndex), "00000") & """ GROUP BY [Número da cotação], [Código do produto], [Código de barras], [Descrição do produto],[Quantidade do produto por unidade de fornecimento] HAVING (((Cotação.[Código de barras]) Is Not Null)) ORDER BY Cotação.[Descrição do produto];"
        Set consultaCotacao = dbCotacao.OpenRecordset(consulta)
        If (consultaCotacao.RecordCount > 0) Then
            consultaCotacao.MoveLast
            consultaCotacao.MoveFirst
            
            numPaginas = Int(consultaCotacao.RecordCount / 50)
            If (numPaginas * 50) < consultaCotacao.RecordCount Then
                numPaginas = numPaginas + 1
            End If
            
            Set inserirCotacao = New ADODB.Recordset
            
            If (lblstatus.Tag = "nula") Then
                ConexaoMysql.Open "Driver={MySQL ODBC 3.51 Driver};user=" + username + ";password=" + password + ";database=" + database + ";server=" + host + ";option=" & (1 + 2 + 8 + 32 + 2048 + 16384)
                consulta = "INSERT INTO cotacoes values(" & Int(numCot) & ",null,0,0)"
                inserirCotacao.Open consulta, ConexaoMysql, adOpenStatic, adLockOptimistic, adCmdText
            End If
            pagina = 1
            cont = 1
            For produto = 1 To consultaCotacao.RecordCount
                If (cont = 51) Then
                    pagina = pagina + 1
                    cont = 1
                End If
                
                'usuariosCotacao.MoveLast
                'usuariosCotacao.MoveFirst
                'If (consultaCotacao.Fields("Código de barras") <> Empty) Then
                    cod = consultaCotacao.Fields("Código do produto")
                    codB = consultaCotacao.Fields("Código de barras")
                    descr = consultaCotacao.Fields("Descrição do produto")
                    emb = consultaCotacao.Fields("Quantidade do produto por unidade de fornecimento")
                    consulta = "INSERT INTO produtos values(null," & cod & ",'" + descr + "','" & codB & "'," & Int(numCot) & "," & emb & "," & pagina & ");"
                    inserirCotacao.Open consulta, ConexaoMysql, adOpenStatic, adLockOptimistic, adCmdText
                    
                'End If
                cont = cont + 1
                consultaCotacao.MoveNext
            Next
            ConexaoMysql.Close
            ConexaoMysql.Open "Driver={MySQL ODBC 3.51 Driver};user=" + username + ";password=" + password + ";database=" + database + ";server=" + host + ";option=" & (1 + 2 + 8 + 32 + 2048 + 16384)
            Set usuariosCotacao = New ADODB.Recordset
            consulta = "Select * From usuarios Where liberado = 1"
            usuariosCotacao.Open consulta, ConexaoMysql, adOpenStatic, adLockOptimistic, adCmdText
            If (usuariosCotacao.RecordCount > 0) Then
                For usuario = 1 To usuariosCotacao.RecordCount
                    rep = Int(usuariosCotacao.Fields("id"))
                    consulta = "Call replicaCotacao(" & rep & "," & numCot & ");"
                    inserirCotacao.Open consulta, ConexaoMysql, adOpenStatic, adLockOptimistic, adCmdText
                    usuariosCotacao.MoveNext
                Next
            End If
            Cotacao.MousePointer = 0
            ConexaoMysql.Close
            MsgBox "Cotação foi carregada com sucesso !"
            Set consultaIdCotacao = New ADODB.Recordset
            ConexaoMysql.Open "Driver={MySQL ODBC 3.51 Driver};user=" + username + ";password=" + password + ";database=" + database + ";server=" + host + ";option=" & (1 + 2 + 8 + 32 + 2048 + 16384)
            consulta = "SELECT * from cotacoes where idCotacao = " & numCot & ";"
            consultaIdCotacao.Open consulta, ConexaoMysql, adOpenStatic, adLockOptimistic, adCmdText
            If consultaIdCotacao.RecordCount > 0 Then
                If consultaIdCotacao.Fields("encerrada") Then
                    lblstatus.Caption = "Status: Cotação nº " & numCot & " está online e encerrada"
                    lblstatus.ForeColor = &HFF00&
                    lblstatus.Tag = "encerrada"
                    botaoUpload.Enabled = False
                    botaoEncerrar.Enabled = False
                Else
                    lblstatus.Caption = "Status: Cotação nº " & numCot & " está online e aberta"
                    lblstatus.ForeColor = &HFFFF&
                    lblstatus.Tag = "aberta"
                    botaoUpload.Enabled = True
                    botaoEncerrar.Enabled = True
                End If
            Else
                lblstatus.Caption = "Status: Cotação nº " & numCot & " não está online"
                lblstatus.ForeColor = &HFF&
                lblstatus.Tag = "nula"
                botaoUpload.Enabled = True
                botaoEncerrar.Enabled = False
            End If
            ConexaoMysql.Close
        End If
    End If
End Sub

Private Sub Form_Activate()
    listaCotacoes.Clear
    Set dbCotacao = OpenDatabase(caminhodados)
    'consulta = "Select [Número da Cotação] from [Cotação] group by [Número da Cotação] order by [Número da Cotação];"
    consulta = "SELECT numeroCotacao FROM NumeracaoCotacao ORDER BY numeroCotacao"
    Set consultaCotacao = dbCotacao.OpenRecordset(consulta)
    If consultaCotacao.RecordCount > 0 Then
        consultaCotacao.MoveLast
        consultaCotacao.MoveFirst
        ReDim numeroCotacao(consultaCotacao.RecordCount)
        For laço = 0 To consultaCotacao.RecordCount - 1
            'numeroCotacao(laço) = consultaCotacao.Fields("Número da Cotação")
            'listaCotacoes.AddItem (consultaCotacao.Fields("Número da Cotação"))
            numeroCotacao(laço) = consultaCotacao.Fields("numeroCotacao")
            listaCotacoes.AddItem (consultaCotacao.Fields("numeroCotacao"))
            consultaCotacao.MoveNext
        Next
        listaCotacoes.ListIndex = listaCotacoes.ListCount - 1
        
    End If
    botaoUpload.Enabled = False
    botaoEncerrar.Enabled = False
    Set ConexaoMysql = New ADODB.Connection
    ConexaoMysql.ConnectionTimeout = 600
    ConexaoMysql.CommandTimeout = 4000
    ConexaoMysql.CursorLocation = adUseClient
End Sub

Private Sub listaCotacoes_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13 Or KeyAscii = 32) And listaCotacoes.ListIndex <> -1 Then
        Cotacao.MousePointer = 11
        numCot = numeroCotacao(listaCotacoes.ListIndex)
        ConexaoMysql.Open "Driver={MySQL ODBC 3.51 Driver};user=" + username + ";password=" + password + ";database=" + database + ";server=" + host + ";option=" & (1 + 2 + 8 + 32 + 2048 + 16384)
        Set consultaIdCotacao = New ADODB.Recordset
        consulta = "SELECT * from cotacoes where idCotacao = " & numCot & ";"
        consultaIdCotacao.Open consulta, ConexaoMysql, adOpenStatic, adLockOptimistic, adCmdText
        If consultaIdCotacao.RecordCount > 0 Then
            If consultaIdCotacao.Fields("encerrada") Then
                lblstatus.Caption = "Status: Cotação nº " & numCot & " está online e encerrada"
                lblstatus.ForeColor = &HFF00&
                lblstatus.Tag = "encerrada"
                botaoUpload.Enabled = False
                botaoEncerrar.Enabled = False
                botaoStatus.Enabled = True
                
            Else
                lblstatus.Caption = "Status: Cotação nº " & numCot & " está online e aberta"
                lblstatus.ForeColor = &HFFFF&
                lblstatus.Tag = "aberta"
                botaoUpload.Enabled = True
                botaoEncerrar.Enabled = True
                botaoStatus.Enabled = True
            End If
        Else
            lblstatus.Caption = "Status: Cotação nº " & numCot & " não está online"
            lblstatus.ForeColor = &HFF&
            lblstatus.Tag = "nula"
            botaoUpload.Enabled = True
            botaoEncerrar.Enabled = False
            botaoStatus.Enabled = False
        End If
        Cotacao.MousePointer = 0
        ConexaoMysql.Close
        
    End If
    
End Sub
