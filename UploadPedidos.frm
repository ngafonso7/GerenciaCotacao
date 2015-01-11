VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form UploadPedidos 
   Caption         =   "Upload de Pedidos"
   ClientHeight    =   5490
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10305
   LinkTopic       =   "Form1"
   ScaleHeight     =   5490
   ScaleWidth      =   10305
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton botaoLiberarPedidos 
      Caption         =   "Liberar Pedidos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6720
      TabIndex        =   8
      Top             =   4560
      Width           =   2655
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
      Left            =   4920
      TabIndex        =   7
      Text            =   "Combo1"
      Top             =   240
      Width           =   4575
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
      Height          =   615
      Left            =   2880
      TabIndex        =   6
      Top             =   4560
      Width           =   2415
   End
   Begin VB.FileListBox File1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2730
      Left            =   4920
      Pattern         =   "*.pdf"
      TabIndex        =   3
      Top             =   840
      Width           =   4575
   End
   Begin VB.DirListBox Dir1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2730
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   4575
   End
   Begin VB.CommandButton botaoUpload 
      Caption         =   "Enviar Pedido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   4560
      Width           =   2415
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   9600
      Top             =   4680
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label lblarquivo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   5
      Top             =   3720
      Width           =   6495
   End
   Begin VB.Label Label1 
      Caption         =   "Arquivo Selecionado : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   3720
      Width           =   2895
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
      Left            =   360
      TabIndex        =   1
      Top             =   240
      Width           =   3795
   End
End
Attribute VB_Name = "UploadPedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim consultaIdUsuario As ADODB.Recordset
Dim regUploadPedido As ADODB.Recordset
Dim regDisponibilizar As ADODB.Recordset

Dim arquivos()

Dim idRepresentante()
Dim usuarioRepresentante()

Dim menor As Integer

Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Private Sub botaoConectar_Click()
    With Inet1
         .URL = "ftp://ftp.supermercadomaryse.com.br"
         .username = "supermerca110"
         .password = "maryse1"
         .Execute , "CD WEB/PDFs/"
         
         Do While .StillExecuting
            Sleep 100
            DoEvents
         Loop
         
         .Execute , "MKDIR 130"
         
         Do While .StillExecuting
            Sleep 100
            DoEvents
         Loop
         
         .Execute , "PUT C:\logNF\43.txt 130/log43.txt"
         
         Do While .StillExecuting
            Sleep 100
            DoEvents
        Loop
        
    End With
End Sub

Private Sub botaoLiberarPedidos_Click()
    res = MsgBox("Deseja mesmo disponibilizar os pedidos para download?", vbYesNo, "Disponibilizar Pedidos")
    If (res = vbYes) Then
        ConexaoMysql.Open "Driver={MySQL ODBC 3.51 Driver};user=" + username + ";password=" + password + ";database=" + database + ";server=" + host + ";option=" & (1 + 2 + 8 + 32 + 2048 + 16384)
        
        Set regDisponibilizar = New ADODB.Recordset
        consulta = "SELECT max(idcotacao) as cot from cotacoes"
        regDisponibilizar.Open consulta, ConexaoMysql, adOpenStatic, adLockOptimistic, adCmdText
        cot = regDisponibilizar.Fields("cot")
        
        Set regDisponibilizar = New ADODB.Recordset
        consulta = "UPDATE cotacoes set pedidosLiberados = 1 where idcotacao =" & cot & ";"
        regDisponibilizar.Open consulta, ConexaoMysql, adOpenStatic, adLockOptimistic, adCmdText
        
        MsgBox ("Pedidos Disponiveis para download !")
        
        ConexaoMysql.Close
    End If
End Sub

Private Sub botaoSair_Click()
    Unload Me
End Sub

Private Sub botaoUpload_Click()
    If (listaRepresentantes.ListIndex <> -1 And File1.FileName <> "") Then
    
        On Error GoTo erro
        MsgBox ("Upload iniciado")
        
        
        UploadPedidos.MousePointer = vbHourglass
        arquivo = Replace(File1.FileName, " ", "")
    
        ConexaoMysql.Open "Driver={MySQL ODBC 3.51 Driver};user=" + username + ";password=" + password + ";database=" + database + ";server=" + host + ";option=" & (1 + 2 + 8 + 32 + 2048 + 16384)
        
        Set regUploadPedido = New ADODB.Recordset
        consulta = "DELETE FROM pedidos WHERE idRepresentante = " & idRepresentante(listaRepresentantes.ListIndex)
        regUploadPedido.Open consulta, ConexaoMysql, adOpenStatic, adLockOptimistic, adCmdText
        
        Set regUploadPedido = New ADODB.Recordset
        consulta = "INSERT INTO pedidos (idRepresentante,arquivoPedido) values (" & idRepresentante(listaRepresentantes.ListIndex) & ",""" & arquivo & """);"
        regUploadPedido.Open consulta, ConexaoMysql, adOpenStatic, adLockOptimistic, adCmdText
        
        Set regUploadPedido = New ADODB.Recordset
        consulta = "SELECT Max(idCotacao) as cot from cotacoes"
        regUploadPedido.Open consulta, ConexaoMysql, adOpenStatic, adLockOptimistic, adCmdText
        
        If (regUploadPedido.RecordCount > 0) Then
            nomePasta = regUploadPedido.Fields("cot")
            With Inet1
                 .URL = "ftp://ftp.supermercadomaryse.com.br"
                 .username = "supermerca110"
                 .password = "maryse1"
                 .Execute , "CD /supermercadomaryse.com.br/web/PDFs"
                 
                 Do While .StillExecuting
                    Sleep 100
                    DoEvents
                 Loop
                 
                 .Execute , "MKDIR " & nomePasta
                 
                 Do While .StillExecuting
                    Sleep 100
                    DoEvents
                 Loop
                 
                 FileCopy File1.Path & "\" & File1.FileName, "C:\tempPedidos\" & arquivo
                 
                 '.Execute , "CD " & nomePasta
                 
                '.Execute , "PUT C:\tempPedidos" & "\" & arquivo & " " & """\supermercadomaryse.com.br\web\PDFs\" & nomePasta & "\" & arquivo & """"
                '.Execute , "PUT C:\tempPedidos" & "\" & arquivo & " /supermercadomaryse.com.br/web/PDFs/" & nomePasta & "/" & arquivo
                 .Execute , "PUT C:\tempPedidos" & "\" & arquivo & " " & nomePasta & "/" & arquivo
                 
                Do While .StillExecuting
                    Sleep 100
                    DoEvents
                Loop
                
                'MsgBox "Response: " & .ResponseCode & " - " & .ResponseInfo
                
                If (.ResponseCode = 0) Then
                    MsgBox ("Arquivo carregado com sucesso !")
                Else
                    MsgBox ("Mensagem ao suporte: " & "Response Code: " & .ResponseCode & " - Mensage: " & .ResponseInfo)
                    
                End If
                
                
                .Cancel
                
                
            End With
            If (ConexaoMysql.State = 1) Then
                ConexaoMysql.Close
            End If
            
            GoTo sair
erro:
            MsgBox ("Erro ao realizar upload. Tente novamente !")
            
sair:
            UploadPedidos.MousePointer = vbDefault
            listaRepresentantes.ListIndex = -1
            listaRepresentantes.Text = "Selecione"
            listaRepresentantes.SetFocus
            listaRepresentantes.Locked = False
            botaoLiberarPedidos.Enabled = True

            
            File1.Path = ""
            lblarquivo.Caption = ""
            
        End If
    
    End If
    
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
        
        
End Sub

Private Sub Dir1_Click()
    File1.Path = Dir1.Path
End Sub

Private Sub File1_Click()
    lblarquivo.Caption = File1.FileName
    
End Sub

Private Sub File1_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13 And File1.FileName <> "") Then
        botaoUpload.SetFocus
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
        For laço = 0 To consultaIdUsuario.RecordCount - 1
            idRepresentante(laço) = consultaIdUsuario.Fields("id")
            usuarioRepresentante(laço) = consultaIdUsuario.Fields("usuario")
            listaRepresentantes.AddItem (consultaIdUsuario.Fields("nomeRep"))
            consultaIdUsuario.MoveNext
        Next
        listaRepresentantes.ListIndex = -1
        listaRepresentantes.Text = "Selecione"
        listaRepresentantes.SetFocus
        
        Set regDisponibilizar = New ADODB.Recordset
        consulta = "SELECT pedidosLiberados from cotacoes WHERE pedidosLiberados=false and idcotacao in(SELECT max(idcotacao) from cotacoes)"
        regDisponibilizar.Open consulta, ConexaoMysql, adOpenStatic, adLockOptimistic, adCmdText
        
        If (regDisponibilizar.RecordCount = 0) Then
            botaoLiberarPedidos.Enabled = False
        End If
        
        ConexaoMysql.Close
    Else
        MsgBox "Não existe(m) usuario(s) cadastrado(s) !"
        Unload Me
    End If
    tam = Len(caminhodados)
    Dir1.Path = Left$(caminhodados, tam - 9) & "pedidos"
    
End Sub

Private Sub listaRepresentantes_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13 And listaRepresentantes.ListIndex <> -1) Then
        File1.SetFocus
    End If
   
    
End Sub
