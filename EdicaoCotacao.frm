VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form EdicaoCotacao 
   Caption         =   " Edita Cotação"
   ClientHeight    =   5820
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15315
   LinkTopic       =   "Form1"
   ScaleHeight     =   5820
   ScaleWidth      =   15315
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
      Left            =   6360
      TabIndex        =   1
      Top             =   5040
      Width           =   1935
   End
   Begin MSFlexGridLib.MSFlexGrid GradeEmpresa 
      Height          =   4455
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   9495
      _ExtentX        =   16748
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
Attribute VB_Name = "EdicaoCotacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim consultaEmpresaCotacao As ADODB.Recordset
Dim consultaEmpresa As ADODB.Recordset
Dim Empresas



Private Sub botaoSair_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    ConexaoMysql.Open "Driver={MySQL ODBC 3.51 Driver};user=" + username + ";password=" + password + ";database=" + database + ";server=" + host + ";option=" & (1 + 2 + 8 + 32 + 2048 + 16384)
    Set consultaEmpresa = New ADODB.Recordset
    
    consulta = "Select u.usuario, e.nome, u.id,e.idEmpresa from usuarios u, empresa e where u.id=usuario_idUsuario;"
    consultaEmpresa.Open consulta, ConexaoMysql, adOpenStatic, adLockOptimistic, adCmdText
    If (consultaEmpresa.RecordCount > 0) Then
        consultaEmpresa.MoveLast
        consultaEmpresa.MoveFirst
        
        ReDim Empresas(consultaEmpresa.RecordCount)
        
        GradeEmpresa.Rows = 1
        GradeEmpresa.Cols = 4
        GradeEmpresa.FormatString = " |^  Usuario  |^                              Empresa                                      |^    Participa"
        
        
        For emp = 0 To consultaEmpresa.RecordCount - 1
            part = " "
            
            consulta = "Select idRepresentante from produtosRep where idRepresentante = " & consultaEmpresa.Fields("id") & " and id_cotacao = " & numEdicaoCotacao & ";"
            Set consultaEmpresaCotacao = New ADODB.Recordset
            consultaEmpresaCotacao.Open consulta, ConexaoMysql, adOpenStatic, adLockOptimistic, adCmdText
            If (consultaEmpresaCotacao.RecordCount <> 0) Then
                part = "X"
            End If
            consultaEmpresaCotacao.Close
            Empresas(emp) = consultaEmpresa.Fields("idEmpresa")
            GradeEmpresa.AddItem (vbTab & consultaEmpresa.Fields("usuario") & vbTab & consultaEmpresa.Fields("nome") & vbTab & part)
            
            
            consultaEmpresa.MoveNext
        Next
    End If
    
End Sub
