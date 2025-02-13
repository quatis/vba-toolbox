Private Sub CheckBox1_Click()
    If CheckBox1.Value = True Then
        DesmarcarOutrasCaixas "CheckBox1"
        Sheets("Saldo diário").Range("C1").Value = "Saldo Total"
    End If
End Sub

Private Sub CheckBox2_Click()
    If CheckBox2.Value = True Then
        DesmarcarOutrasCaixas "CheckBox2"
        Sheets("Saldo diário").Range("C1").Value = "Saldo Aplicado"
    End If
End Sub

Private Sub CheckBox3_Click()
    If CheckBox3.Value = True Then
        DesmarcarOutrasCaixas "CheckBox3"
        Sheets("Saldo diário").Range("C1").Value = "Delta de Saldos"
    End If
End Sub
Private Sub CheckBox4_Click()
    If CheckBox4.Value = True Then
        DesmarcarOutrasCaixas "CheckBox4"
        Range("C1").Value = "Saldos Negativos"
        ' Limpar e popular a ComboBox com empresas negativas
        PopularComboBoxNegativo
        ' Tornar a CheckBox5 visível
        CheckBox5.Visible = True
    Else
        PreencherComboBox
        ' Ocultar e desmarcar a CheckBox5
        CheckBox5.Visible = False
        CheckBox5.Value = False
        Range("C1").Value = ""
    End If
End Sub

Private Sub CheckBox5_Click()
    If CheckBox5.Value = True Then
        Range("C1").Value = "Saldo Aplicado Negativo"
    Else
        Range("C1").Value = "Saldos Negativos"
    End If
End Sub
Public Sub PopularComboBoxNegativo()
    Dim wsDados As Worksheet
    Dim rng As Range
    Dim cell As Range
    Dim i As Integer
    Dim empresas() As Empresa

    ' Inicializar empresas
    empresas = inicializarEmpresas()

    Set wsDados = ThisWorkbook.Sheets("Dados") ' Planilha com os dados das empresas negativas
    Set rng = wsDados.Range("R2:R50") ' Definindo o intervalo específico

    ' Limpar a ComboBox
    ComboBox1.Clear

    ' Popular a ComboBox com a lista de empresas da tabela do Power Query
    For Each cell In rng
        If cell.Value <> "" Then
            For i = LBound(empresas) To UBound(empresas)
                If empresas(i).codigo Like "*" & cell.Value & "*" Then
                    ComboBox1.AddItem empresas(i).nome
                    Exit For
                End If
            Next i
        End If
    Next cell
    ComboBox1.ListIndex = 0
End Sub
'sub que faz com que as checkboxes funcionem como optionbutton (radio) por questao estetica
Private Sub DesmarcarOutrasCaixas(except As String)
    Dim ctrl As OLEObject
    Dim checkBoxNames As Variant
    Dim i As Integer

    checkBoxNames = Array("CheckBox1", "CheckBox2", "CheckBox3", "CheckBox4")

    For i = LBound(checkBoxNames) To UBound(checkBoxNames)
        If checkBoxNames(i) <> except Then
            Sheets("Saldo diário").OLEObjects(checkBoxNames(i)).Object.Value = False
        End If
    Next i
End Sub
'sub executado toda vez que o arquivo e aberto, popula a combobox1 com os nomes de empresa
Public Sub PreencherComboBox()
        Dim empresas() As Empresa
        Dim i As Integer
        
        empresas = inicializarEmpresas()
        ComboBox1.Clear
    
        For i = LBound(empresas) To UBound(empresas)
            ComboBox1.AddItem empresas(i).nome
        Next i
    ComboBox1.ListIndex = 0
End Sub
'sub que quando alteramos o valor da combobox ele procura por esse valor na lista de empresas e chama o AtualizarCodigos do modulo1 passando como parametro a empresa
Private Sub ComboBox1_Change()
    Dim empresas() As Empresa
    Dim i As Integer
    empresas = inicializarEmpresas()  ' Inicializa a lista de empresas
    
    ' Verifica se o valor da ComboBox corresponde ao nome de alguma empresa na lista
    For i = 1 To UBound(empresas)
        If ComboBox1.Value = empresas(i).nome Then
            Call AtualizarCodigos(Me, empresas(i))
            Exit For  ' Encerra o loop após encontrar a empresa correspondente
        End If
    Next i
End Sub

Option Explicit
Public Type Empresa
    codigo As String
    nome As String
End Type

Public Sub AtualizarCodigos(ws As Worksheet, empresaSelecionada As Empresa)
'Sub chamado pela combobox da planilha Saldo Diario, altera as celulas AP1 e AP2 com o codigo e nome da empresa respectivamente.
    With ws
        .Range("AP1").Value = removerHifen(empresaSelecionada.codigo)
        .Range("AP2").Value = empresaSelecionada.nome
        .OLEObjects("ComboBox1").Object.Value = empresaSelecionada.nome
    End With
End Sub


Public Function inicializarEmpresas() As Empresa()
'Funcao que inicializa a matriz bidimensional de empresas, a variavel .nome popula a combo box toda vez que a planilha e aberta.
    Dim empresas(1 To 65) As Empresa
    
    With empresas(1)
        .codigo = "1000": .nome = "company1"
    End With
    With empresas(2)
        .codigo = "1600-": .nome = "company2"
    End With
    With empresas(3)
        .codigo = "7000-7100-7200-7300-"
        .nome = "complex1"
    End With
    
    inicializarEmpresas = empresas
End Function

Private Function removerHifen(ByVal codigo As String) As String
'Funcao para remover o ultimo hifen dos codigos de empresa, utilizado para calcular complexos.
    If Len(codigo) > 0 Then
        removerHifen = Left(codigo, Len(codigo) - 1)
    Else
        removerHifen = ""
    End If
End Function

Function SomarSaldos(codigos As String, data As Date, tipoSaldo As String) As Double
   Dim codigoArray() As String
   Dim i As Integer
   Dim saldoTotal As Double
   Dim codigo As String
   Dim linha As Long
   Dim coluna As Range
   Dim ws As Worksheet
   Set ws = Sheets("Dados")

   Select Case tipoSaldo
    Case "Saldo Total"
           Set coluna = ws.Range("F:F")
    Case "Saldo Aplicado"
           Set coluna = ws.Range("E:E")
    Case "Delta de Saldos"
           Set coluna = ws.Range("D:D")
    Case "Saldos Negativos"
           Set coluna = ws.Range("N:N")
    Case "Saldo Aplicado Negativo"
           Set coluna = ws.Range("M:M")
    Case Else
           SomarSaldos = 0
'funcao que seleciona a coluna de acordo com o tipo de saldo (saldo total = coluna F, aplicado = E, delta = D, negativo = N e apl negativo = M)
           Exit Function
   End Select
   codigoArray = Split(codigos, "-")
   saldoTotal = 0

   For i = LBound(codigoArray) To UBound(codigoArray)
       codigo = codigoArray(i)
       'configura a concatenacao dos codigos de empresa, aceitando variaveis do tipo codigo separadas por - para calcular os complexos
       If tipoSaldo = "Saldos Negativos" Or tipoSaldo = "Saldo Aplicado Negativo" Then
           ' PROCURA PELA CHAVE NA COLUNA O SE FOR SALDO NEGATIVO
           linha = Application.WorksheetFunction.Match(codigo & "-" & Format(data, "dd/mm/yyyy"), ws.Range("O:O"), 0)
       Else
           ' PROCURA NA COLUNA G PARA OUTROS TIPOS
           linha = Application.WorksheetFunction.Match(codigo & "-" & Format(data, "dd/mm/yyyy"), ws.Range("G:G"), 0)
       End If
       
           saldoTotal = saldoTotal + Application.WorksheetFunction.Index(coluna, linha)
   Next i
   'metodo que soma os saldos das empresas concatenadas por dia
'For loop percorre da LBound (Lower Bound - limite inferior) ate o Ubound (Upper Bound - limite superior)
'a funcao e executada de acordo com o numero de codigos passado
   
   SomarSaldos = saldoTotal
End Function
