VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPrincipal
   Caption = "Mini Gestor de Despesas"
   ClientHeight = 8700
   ClientWidth = 15600
   StartUpPosition = 2
   KeyPreview = -1
   Begin VB.Label pnlTopo
      Left = 0
      Top = 0
      Width = 15600
      Height = 1200
      TabIndex = 0
   End
   Begin VB.Label lblTitulo
      Left = 360
      Top = 200
      Width = 10000
      Height = 480
      TabIndex = 1
      Caption = "MINI GESTOR | Despesas"
   End
   Begin VB.Label lblSubtitulo
      Left = 360
      Top = 720
      Width = 10000
      Height = 300
      TabIndex = 2
      Caption = "Controle seus gastos com clareza"
   End
   Begin VB.Label lblEditor
      Left = 360
      Top = 1560
      Width = 4200
      Height = 480
      TabIndex = 3
      Caption = "Nova despesa"
   End
   Begin VB.Label lblDescricao
      Left = 360
      Top = 2280
      Width = 4200
      Height = 300
      TabIndex = 4
      Caption = "Descrição"
   End
   Begin VB.TextBox txtDescricao
      Left = 360
      Top = 2640
      Width = 4200
      Height = 480
      TabIndex = 5
      Text = ""
   End
   Begin VB.Label lblValor
      Left = 360
      Top = 3360
      Width = 2000
      Height = 300
      TabIndex = 6
      Caption = "Valor (R$)"
   End
   Begin VB.TextBox txtValor
      Left = 360
      Top = 3720
      Width = 4200
      Height = 480
      TabIndex = 7
      Text = ""
   End
   Begin VB.Label lblData
      Left = 360
      Top = 4440
      Width = 4200
      Height = 300
      TabIndex = 8
      Caption = "Data da despesa (dd/mm/aaaa)"
   End
   Begin VB.TextBox txtData
      Left = 360
      Top = 4800
      Width = 4200
      Height = 480
      TabIndex = 9
      Text = ""
   End
   Begin VB.CommandButton btnSalvar
      Left = 360
      Top = 5640
      Width = 4200
      Height = 600
      TabIndex = 10
      Caption = "&Salvar despesa"
      Style = 1
      BackColor = &H97C520&
   End
   Begin VB.CommandButton btnLimpar
      Left = 360
      Top = 6420
      Width = 4200
      Height = 540
      TabIndex = 11
      Caption = "&Limpar campos"
   End
   Begin VB.Label lblLista
      Left = 5040
      Top = 1560
      Width = 9000
      Height = 480
      TabIndex = 12
      Caption = "Despesas registradas"
   End
   Begin MSFlexGridLib.MSFlexGrid gridDespesas
      Left = 5040
      Top = 2280
      Width = 10200
      Height = 4200
      TabIndex = 13
      Rows = 2
      Cols = 4
      FixedRows = 1
      FixedCols = 0
   End
   Begin VB.CommandButton btnEditar
      Left = 5040
      Top = 6720
      Width = 2700
      Height = 540
      TabIndex = 14
      Caption = "&Editar selecionada"
   End
   Begin VB.CommandButton btnExcluir
      Left = 7920
      Top = 6720
      Width = 2700
      Height = 540
      TabIndex = 15
      Caption = "E&xcluir selecionada"
   End
   Begin VB.CommandButton btnAtualizar
      Left = 10800
      Top = 6720
      Width = 2100
      Height = 540
      TabIndex = 16
      Caption = "&Atualizar"
   End
   Begin VB.Label lblTotal
      Left = 5040
      Top = 7620
      Width = 10200
      Height = 540
      TabIndex = 17
      Caption = "Total das despesas: R$ 0,00"
   End
   Begin VB.Label lblStatus
      Left = 360
      Top = 8160
      Width = 12000
      Height = 300
      TabIndex = 18
      Caption = "Selecione uma despesa para editar ou excluir."
   End
   Begin VB.CommandButton btnSair
      Left = 13500
      Top = 8040
      Width = 1740
      Height = 480
      TabIndex = 19
      Caption = "Sai&r"
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private idEdicao As Long
Private descricaoOriginal As String
Private valorOriginal As Currency
Private dataOriginal As Date
Private carregando As Boolean

Private Sub Form_Load()
    AplicarVisual
    If Not AbrirConexao() Then
        Unload Me
        Exit Sub
    End If
    LimparEdicao
    CarregarDespesas
End Sub

Private Sub AplicarVisual()
    Dim controle As Control
    Me.BackColor = RGB(243, 246, 250)
    For Each controle In Me.Controls
        controle.Font.Name = "Segoe UI"
        controle.Font.Size = 11
        If TypeOf controle Is Label Then
            controle.BackStyle = 0
            controle.ForeColor = RGB(32, 49, 71)
        End If
        If TypeOf controle Is CommandButton Then controle.Font.Bold = True
    Next
    pnlTopo.BackStyle = 1
    pnlTopo.BackColor = RGB(24, 40, 64)
    pnlTopo.ZOrder 1
    lblTitulo.ForeColor = vbWhite
    lblTitulo.Font.Size = 22
    lblTitulo.Font.Bold = True
    lblSubtitulo.ForeColor = RGB(112, 226, 201)
    lblEditor.Font.Size = 16
    lblEditor.Font.Bold = True
    lblLista.Font.Size = 16
    lblLista.Font.Bold = True
    lblTotal.Font.Size = 18
    lblTotal.Font.Bold = True
    lblTotal.ForeColor = RGB(0, 116, 91)
    txtDescricao.MaxLength = 100
    txtValor.MaxLength = 12
    txtData.MaxLength = 10
    txtDescricao.TabIndex = 0
    txtValor.TabIndex = 1
    txtData.TabIndex = 2
    btnSalvar.TabIndex = 3
    btnLimpar.TabIndex = 4
    gridDespesas.TabIndex = 5
    btnEditar.TabIndex = 6
    btnExcluir.TabIndex = 7
    btnAtualizar.TabIndex = 8
    btnSair.TabIndex = 9
    With gridDespesas
        .BackColor = vbWhite
        .BackColorBkg = vbWhite
        .BackColorFixed = RGB(24, 40, 64)
        .ForeColorFixed = vbWhite
        .ForeColor = RGB(32, 49, 71)
        .BackColorSel = RGB(214, 242, 234)
        .ForeColorSel = RGB(24, 40, 64)
        .GridColor = RGB(226, 232, 240)
        .SelectionMode = 1
        .FocusRect = 0
        .AllowUserResizing = 1
        .RowHeightMin = 420
        .FormatString = "ID|Data|Descrição|Valor (R$)"
        .ColAlignment(3) = 7
    End With
    Form_Resize
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    If Me.Width < 14500 Then Me.Width = 14500
    If Me.Height < 8900 Then Me.Height = 8900
    pnlTopo.Width = Me.ScaleWidth
    gridDespesas.Width = Me.ScaleWidth - gridDespesas.Left - 360
    gridDespesas.Height = Me.ScaleHeight - gridDespesas.Top - 2340
    gridDespesas.ColWidth(0) = 0
    gridDespesas.ColWidth(1) = 1680
    gridDespesas.ColWidth(3) = 1920
    gridDespesas.ColWidth(2) = gridDespesas.Width - 3960
    btnEditar.Top = gridDespesas.Top + gridDespesas.Height + 240
    btnExcluir.Top = btnEditar.Top
    btnAtualizar.Top = btnEditar.Top
    lblTotal.Top = btnEditar.Top + 840
    lblTotal.Width = gridDespesas.Width
    lblStatus.Top = Me.ScaleHeight - 480
    lblStatus.Width = Me.ScaleWidth - 2520
    btnSair.Move Me.ScaleWidth - 2100, Me.ScaleHeight - 600
End Sub

Private Function NovoComando(ByVal nome As String) As ADODB.Command
    Dim comando As New ADODB.Command
    Set comando.ActiveConnection = conn
    comando.CommandType = adCmdStoredProc
    comando.CommandText = nome
    comando.CommandTimeout = 30
    Set NovoComando = comando
End Function

Private Sub ParametroValor(ByVal comando As ADODB.Command, ByVal nome As String, ByVal valor As Currency)
    Dim parametro As ADODB.Parameter
    Set parametro = comando.CreateParameter(nome, adDecimal, adParamInput)
    parametro.Precision = 10
    parametro.NumericScale = 2
    parametro.Value = valor
    comando.Parameters.Append parametro
End Sub

Private Sub ParametrosOriginais(ByVal comando As ADODB.Command)
    comando.Parameters.Append comando.CreateParameter("@DescricaoOriginal", adVarChar, adParamInput, 100, descricaoOriginal)
    ParametroValor comando, "@ValorOriginal", valorOriginal
    comando.Parameters.Append comando.CreateParameter("@DataOriginal", adDBDate, adParamInput, , dataOriginal)
End Sub

Private Sub CarregarDespesas()
    On Error GoTo Falha
    Dim comando As ADODB.Command
    Dim registros As ADODB.Recordset
    Dim total As Variant
    Dim quantidade As Long
    carregando = True
    btnEditar.Enabled = False
    btnExcluir.Enabled = False
    Set comando = NovoComando("dbo.usp_DespesasListar")
    Set registros = comando.Execute
    gridDespesas.Redraw = False
    gridDespesas.Rows = 2
    gridDespesas.Row = 1
    gridDespesas.Clear
    gridDespesas.FormatString = "ID|Data|Descrição|Valor (R$)"
    total = CDec(0)
    Do While Not registros.EOF
        quantidade = quantidade + 1
        If gridDespesas.Rows <= quantidade Then gridDespesas.Rows = quantidade + 1
        gridDespesas.TextMatrix(quantidade, 0) = CStr(registros!ID)
        gridDespesas.TextMatrix(quantidade, 1) = Format$(registros!DataDespesa, "dd/mm/yyyy")
        gridDespesas.TextMatrix(quantidade, 2) = CStr(registros!Descricao)
        gridDespesas.TextMatrix(quantidade, 3) = Format$(registros!Valor, "0.00")
        total = total + CDec(registros!Valor)
        registros.MoveNext
    Loop
    registros.Close
    gridDespesas.Redraw = True
    carregando = False
    lblTotal.Caption = "Total das despesas: " & FormatCurrency(total, 2)
    lblStatus.Caption = CStr(quantidade) & " despesa(s) | F2: nova | F5: atualizar | Duplo clique: editar"
    AtualizarSelecao
    Exit Sub
Falha:
    gridDespesas.Redraw = True
    carregando = False
    lblTotal.Caption = "Total indisponível"
    MsgBox "Não foi possível atualizar a lista: " & Err.Description, vbExclamation
End Sub

Private Sub AtualizarSelecao()
    If carregando Then Exit Sub
    Dim selecionado As Boolean
    selecionado = (Len(gridDespesas.TextMatrix(gridDespesas.Row, 0)) > 0 And gridDespesas.Row > 0)
    btnEditar.Enabled = selecionado And idEdicao = 0
    btnExcluir.Enabled = selecionado And idEdicao = 0
End Sub

Private Sub gridDespesas_RowColChange()
    AtualizarSelecao
End Sub

Private Sub gridDespesas_Click()
    AtualizarSelecao
End Sub

Private Sub gridDespesas_DblClick()
    If btnEditar.Enabled Then btnEditar_Click
End Sub

Private Sub LerSelecionado()
    idEdicao = CLng(gridDespesas.TextMatrix(gridDespesas.Row, 0))
    descricaoOriginal = gridDespesas.TextMatrix(gridDespesas.Row, 2)
    valorOriginal = CCur(gridDespesas.TextMatrix(gridDespesas.Row, 3))
    Dim texto As String
    texto = gridDespesas.TextMatrix(gridDespesas.Row, 1)
    dataOriginal = DateSerial(CInt(Right$(texto, 4)), CInt(Mid$(texto, 4, 2)), CInt(Left$(texto, 2)))
End Sub

Private Sub btnEditar_Click()
    If Not btnEditar.Enabled Then Exit Sub
    LerSelecionado
    txtDescricao.Text = descricaoOriginal
    txtValor.Text = Format$(valorOriginal, "0.00")
    txtData.Text = Format$(dataOriginal, "dd/mm/yyyy")
    lblEditor.Caption = "Editar despesa #" & CStr(idEdicao)
    btnSalvar.Caption = "&Salvar alterações"
    btnLimpar.Caption = "&Cancelar edição"
    btnEditar.Enabled = False
    btnExcluir.Enabled = False
    txtDescricao.SetFocus
End Sub

Private Sub LimparEdicao()
    idEdicao = 0
    txtDescricao.Text = ""
    txtValor.Text = ""
    txtData.Text = Format$(Date, "dd/mm/yyyy")
    lblEditor.Caption = "Nova despesa"
    btnSalvar.Caption = "&Salvar despesa"
    btnLimpar.Caption = "&Limpar campos"
    AtualizarSelecao
End Sub

Private Sub btnSalvar_Click()
    On Error GoTo Falha
    Dim valor As Currency
    Dim dataDespesa As Date
    Dim texto As String
    Dim comando As ADODB.Command
    If Len(Trim$(txtDescricao.Text)) = 0 Then
        MsgBox "Informe a descrição.", vbExclamation
        txtDescricao.SetFocus
        Exit Sub
    End If
    If Not IsNumeric(txtValor.Text) Then GoTo ValorInvalido
    If CDec(txtValor.Text) <= 0 Or CDec(txtValor.Text) > 99999999.99@ Then GoTo ValorInvalido
    valor = CCur(txtValor.Text)
    If CDec(txtValor.Text) <> CDec(valor) Or CDec(valor) * 100 <> Fix(CDec(valor) * 100) Then GoTo ValorInvalido
    texto = Trim$(txtData.Text)
    If Not texto Like "##/##/####" Then GoTo DataInvalida
    On Error GoTo DataInvalida
    dataDespesa = DateSerial(CInt(Right$(texto, 4)), CInt(Mid$(texto, 4, 2)), CInt(Left$(texto, 2)))
    If Format$(dataDespesa, "dd/mm/yyyy") <> texto Then GoTo DataInvalida
    On Error GoTo Falha
    btnSalvar.Enabled = False
    If idEdicao = 0 Then
        Set comando = NovoComando("dbo.usp_DespesasInserir")
    Else
        Set comando = NovoComando("dbo.usp_DespesasEditar")
        comando.Parameters.Append comando.CreateParameter("@ID", adInteger, adParamInput, , idEdicao)
    End If
    comando.Parameters.Append comando.CreateParameter("@DataDespesa", adDBDate, adParamInput, , dataDespesa)
    ParametroValor comando, "@Valor", valor
    comando.Parameters.Append comando.CreateParameter("@Descricao", adVarChar, adParamInput, 100, Trim$(txtDescricao.Text))
    If idEdicao <> 0 Then ParametrosOriginais comando
    comando.Execute , , adExecuteNoRecords
    LimparEdicao
    CarregarDespesas
    btnSalvar.Enabled = True
    txtDescricao.SetFocus
    Exit Sub
ValorInvalido:
    MsgBox "Informe um valor positivo, até 99.999.999,99, com no máximo duas casas decimais. Use o separador decimal do Windows.", vbExclamation
    txtValor.SetFocus
    Exit Sub
DataInvalida:
    MsgBox "Informe uma data válida no formato dd/mm/aaaa.", vbExclamation
    txtData.SetFocus
    Exit Sub
Falha:
    btnSalvar.Enabled = True
    MsgBox "Não foi possível salvar: " & Err.Description, vbExclamation
End Sub

Private Sub btnExcluir_Click()
    On Error GoTo Falha
    If Not btnExcluir.Enabled Then Exit Sub
    LerSelecionado
    If MsgBox("Excluir definitivamente esta despesa?" & vbCrLf & descricaoOriginal & vbCrLf & FormatCurrency(valorOriginal, 2), vbQuestion Or vbYesNo Or vbDefaultButton2, "Confirmar exclusão") <> vbYes Then
        idEdicao = 0
        Exit Sub
    End If
    Dim comando As ADODB.Command
    Set comando = NovoComando("dbo.usp_DespesasExcluir")
    comando.Parameters.Append comando.CreateParameter("@ID", adInteger, adParamInput, , idEdicao)
    ParametrosOriginais comando
    comando.Execute , , adExecuteNoRecords
    LimparEdicao
    CarregarDespesas
    Exit Sub
Falha:
    idEdicao = 0
    MsgBox "Não foi possível excluir: " & Err.Description, vbExclamation
    AtualizarSelecao
End Sub

Private Sub btnLimpar_Click()
    LimparEdicao
    txtDescricao.SetFocus
End Sub

Private Sub btnAtualizar_Click()
    If idEdicao <> 0 Then
        MsgBox "Salve ou cancele a edição antes de atualizar.", vbInformation
        Exit Sub
    End If
    CarregarDespesas
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF2 Then btnLimpar_Click
    If KeyCode = vbKeyF5 Then btnAtualizar_Click
End Sub

Private Sub btnSair_Click()
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FecharConexao
End Sub
