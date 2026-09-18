VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPrincipal 
   Caption         =   "Mini Gestor de Despesas"
   ClientHeight    =   5910
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8760
   LinkTopic       =   "Form1"
   ScaleHeight     =   5910
   ScaleWidth      =   8760
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid gridDespesas 
      Height          =   3500
      Left            =   240
      TabIndex        =   7
      Top             =   1680
      Width           =   5900
      _ExtentX        =   10398
      _ExtentY        =   6165
      _Version        =   393216
      Cols            =   3
   End
   Begin VB.TextBox txtValor 
      Height          =   495
      Left            =   6360
      TabIndex        =   6
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox txtDescricao 
      Height          =   495
      Left            =   6360
      TabIndex        =   4
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton btnSair 
      Caption         =   "&Sair"
      Height          =   495
      Left            =   3240
      TabIndex        =   2
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton btnLimpar 
      Caption         =   "&Limpar"
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton btnSalvar 
      Caption         =   "&Salvar"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Valor:"
      Height          =   495
      Left            =   4920
      TabIndex        =   5
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Descrição:"
      Height          =   495
      Left            =   4920
      TabIndex        =   3
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    ' Abre a conexão; se falhar, fecha o formulário
    If Not AbrirConexao() Then
        Unload Me
        Exit Sub
    End If

    ' Define explicitamente que a grelha tem 3 colunas
    gridDespesas.Cols = 3

    ' Configura os cabeçalhos da grelha de despesas
    gridDespesas.TextMatrix(0, 0) = "ID"
    gridDespesas.TextMatrix(0, 1) = "Descrição"
    gridDespesas.TextMatrix(0, 2) = "Valor (R$)"

    ' Ajusta a largura das colunas
    gridDespesas.ColWidth(0) = 800
    gridDespesas.ColWidth(1) = 3200
    gridDespesas.ColWidth(2) = 1500

    ' Carrega os dados existentes ao abrir o formulário
    Call CarregarDespesas
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call FecharConexao
End Sub

Private Sub btnSalvar_Click()
    On Error GoTo ErroSalvar

    ' Validação simples dos campos
    If Trim(txtDescricao.Text) = "" Or Trim(txtValor.Text) = "" Then
        MsgBox "Por favor, preencha a descrição e o valor da despesa!", vbExclamation, "Atenção"
        Exit Sub
    End If

    If Not IsNumeric(txtValor.Text) Then
        MsgBox "O valor informado não é um número válido!", vbExclamation, "Atenção"
        txtValor.SetFocus
        Exit Sub
    End If

    ' Instrução SQL para inserir a despesa (aspas simples tratadas)
    Dim sql As String
    sql = "INSERT INTO despesas (descricao, valor) VALUES ('" & _
          Replace(txtDescricao.Text, "'", "''") & "', " & _
          Replace(txtValor.Text, ",", ".") & ")"

    ' Executa o comando na conexão aberta
    conn.Execute sql

    ' Limpa os campos após salvar
    txtDescricao.Text = ""
    txtValor.Text = ""
    txtDescricao.SetFocus

    ' Atualiza a grelha para mostrar o novo registo
    Call CarregarDespesas

    MsgBox "Despesa salva com sucesso!", vbInformation, "Sucesso"
    Exit Sub

ErroSalvar:
    MsgBox "Erro ao salvar a despesa: " & Err.Description, vbCritical, "Erro"
End Sub

Private Sub btnLimpar_Click()
    txtDescricao.Text = ""
    txtValor.Text = ""
    txtDescricao.SetFocus
End Sub

Private Sub btnSair_Click()
    Unload Me
End Sub

Private Sub CarregarDespesas()
    On Error GoTo ErroCarregar
    Dim rs As ADODB.Recordset
    Dim sql As String

    ' Garante 3 colunas antes de carregar
    gridDespesas.Cols = 3

    ' Consulta todas as despesas na base de dados
    sql = "SELECT id, descricao, valor FROM despesas ORDER BY id"

    Set rs = New ADODB.Recordset
    rs.Open sql, conn, adOpenStatic, adLockReadOnly, adCmdText

    ' Limpa a grelha mantendo apenas a linha de cabeçalho
    gridDespesas.Rows = 1

    ' Percorre os registos e adiciona na grelha
    Do While Not rs.EOF
        gridDespesas.AddItem rs.Fields("id").Value & vbTab & _
            rs.Fields("descricao").Value & vbTab & _
            Format(rs.Fields("valor").Value, "Standard")
        rs.MoveNext
    Loop

    rs.Close
    Set rs = Nothing
    Exit Sub

ErroCarregar:
    MsgBox "Erro " & Err.Number & ": " & Err.Description, vbCritical, "Erro"
End Sub

