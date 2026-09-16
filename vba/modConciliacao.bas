Attribute VB_Name = "modConciliacao"
'==============================================================================
' modConciliacao
'
' Motor de conciliacao de folha. Ponto de entrada: ExecutarConciliacao.
'
' Desenho: tudo entra em array de uma vez, o processamento acontece em memoria
' e a escrita volta em um unico bloco. Percorrer 10 mil linhas com Cells(i, j)
' leva minutos; em array leva menos de um segundo. Em automacao de Excel essa
' escolha e a diferenca entre uma ferramenta que o time usa e uma que o time
' abandona.
'
' Dependencias: modRegras, modUtil.
'==============================================================================
Option Explicit

Public Const VERSAO As String = "1.0.0"

' Colunas da aba Conciliacao
Private Const C_MATRICULA As Long = 1
Private Const C_RUBRICA As Long = 2
Private Const C_DESCRICAO As Long = 3
Private Const C_TIPO As Long = 4
Private Const C_VAL_ANT As Long = 5
Private Const C_VAL_ATU As Long = 6
Private Const C_QTD As Long = 7
Private Const C_VAR_ABS As Long = 8
Private Const C_VAR_PCT As Long = 9
Private Const C_SITUACAO As Long = 10

' Colunas da aba Excecoes
Private Const E_COLUNAS As Long = 14


Public Sub ExecutarConciliacao()
    Dim inicio As Single
    inicio = Timer

    On Error GoTo TratarErro
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Dim cfg As Object:    Set cfg = modUtil.LerConfig()
    Dim regras As Object: Set regras = modUtil.LerRegras()

    If Not modUtil.ValidarAmbiente(cfg) Then GoTo Finalizar

    Dim conc As Variant
    conc = modUtil.LerIntervalo(ThisWorkbook.Worksheets("Conciliacao"))
    If IsEmpty(conc) Then
        MsgBox "A aba Conciliacao esta vazia. Atualize as consultas antes " & _
               "de executar (Dados > Atualizar Tudo).", vbExclamation, "Conciliador"
        GoTo Finalizar
    End If

    Dim cadastro As Object: Set cadastro = modUtil.LerCadastro()
    Dim agregados As Object: Set agregados = AgregarPorMatricula(conc)

    ' Matriculas cuja ausencia ou entrada ja tem explicacao propria. Sem esta
    ' lista, uma pessoa que saiu da folha dispara uma exceção por rubrica,
    ' enterrando o alerta relevante em dezenas de linhas de ruido.
    Dim explicadas As Object: Set explicadas = CreateObject("Scripting.Dictionary")

    Dim saida() As Variant
    Dim total As Long
    ReDim saida(1 To 20000, 1 To E_COLUNAS)
    total = 0

    ' Ordem importa: as regras de movimentacao rodam primeiro e marcam as
    ' matriculas que as regras por rubrica devem ignorar.
    modRegras.AplicarR01 cfg, regras, cadastro, agregados, explicadas, saida, total
    modRegras.AplicarR02 cfg, regras, cadastro, agregados, explicadas, saida, total
    modRegras.AplicarR09 cfg, regras, cadastro, agregados, explicadas, saida, total
    modRegras.AplicarR03 cfg, regras, cadastro, agregados, explicadas, saida, total
    modRegras.AplicarPorRubrica cfg, regras, cadastro, conc, explicadas, saida, total

    EscreverExcecoes saida, total
    modUtil.RegistrarLog cfg, UBound(conc, 1), total, ContarCriticas(saida, total), _
                         Timer - inicio

Finalizar:
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.Calculate

    If total > 0 Then
        ThisWorkbook.Worksheets("Resumo").Activate
        MsgBox total & " excecao(oes) encontrada(s) em " & _
               Format(Timer - inicio, "0.0") & " segundo(s)." & vbCrLf & vbCrLf & _
               "Detalhe na aba Excecoes.", vbInformation, "Conciliador"
    End If
    Exit Sub

TratarErro:
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Erro " & Err.Number & " em " & Err.Source & vbCrLf & _
           Err.Description, vbCritical, "Conciliador"
End Sub


'------------------------------------------------------------------------------
' Agrega a conciliacao por matricula. Cada item guarda:
'   0 provento_anterior   1 provento_atual
'   2 desconto_anterior   3 desconto_atual
'   4 salario_base_atual  5 horas_extras_atual
'   6 linhas_com_valor_atual
'------------------------------------------------------------------------------
Private Function AgregarPorMatricula(conc As Variant) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim i As Long, chave As String, v As Variant

    For i = 1 To UBound(conc, 1)
        chave = CStr(conc(i, C_MATRICULA))
        If Not d.Exists(chave) Then
            d.Add chave, Array(0#, 0#, 0#, 0#, 0#, 0#, 0&)
        End If
        v = d(chave)

        If conc(i, C_TIPO) = "Provento" Then
            v(0) = v(0) + CDbl(conc(i, C_VAL_ANT))
            v(1) = v(1) + CDbl(conc(i, C_VAL_ATU))
        Else
            v(2) = v(2) + CDbl(conc(i, C_VAL_ANT))
            v(3) = v(3) + CDbl(conc(i, C_VAL_ATU))
        End If

        Select Case CStr(conc(i, C_RUBRICA))
            Case "001": v(4) = CDbl(conc(i, C_VAL_ATU))
            Case "002": v(5) = CDbl(conc(i, C_VAL_ATU))
        End Select

        If CDbl(conc(i, C_VAL_ATU)) <> 0 Then v(6) = v(6) + 1
        d(chave) = v
    Next i

    Set AgregarPorMatricula = d
End Function


Private Sub EscreverExcecoes(saida() As Variant, total As Long)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Excecoes")

    ' Limpa mantendo cabecalho e formatacao condicional.
    Dim ultima As Long
    ultima = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If ultima > 1 Then ws.Range("A2:N" & ultima).ClearContents

    If total = 0 Then
        ws.Range("A2").Value = "Nenhuma excecao encontrada nesta competencia."
        Exit Sub
    End If

    Dim bloco() As Variant
    ReDim bloco(1 To total, 1 To E_COLUNAS)

    Dim i As Long, j As Long
    For i = 1 To total
        For j = 1 To E_COLUNAS
            bloco(i, j) = saida(i, j)
        Next j
    Next i

    ws.Range("A2").Resize(total, E_COLUNAS).Value = bloco
    ws.Range("I2:K" & total + 1).NumberFormat = "R$ #,##0.00;[Red](R$ #,##0.00)"
    ws.Range("A2:N" & total + 1).Font.Name = "Arial"
    ws.Range("A2:N" & total + 1).Font.Size = 10
    ws.Range("L2:M" & total + 1).WrapText = False
    ws.AutoFilterMode = False
    ws.Range("A1:N" & total + 1).AutoFilter
    ws.Sort.SortFields.Clear
End Sub


Private Function ContarCriticas(saida() As Variant, total As Long) As Long
    Dim i As Long, n As Long
    For i = 1 To total
        If saida(i, 2) = "Critica" Or saida(i, 2) = "Crítica" Then n = n + 1
    Next i
    ContarCriticas = n
End Function


'------------------------------------------------------------------------------
' Conferencia contra o gabarito dos dados de demonstracao.
' Rode depois da conciliacao para validar o motor de regras.
'------------------------------------------------------------------------------
Public Sub ConferirGabarito()
    Dim wsG As Worksheet, wsE As Worksheet
    Set wsG = ThisWorkbook.Worksheets("Gabarito_Demo")
    Set wsE = ThisWorkbook.Worksheets("Excecoes")

    Dim encontradas As Object: Set encontradas = CreateObject("Scripting.Dictionary")
    Dim i As Long, ultima As Long

    ultima = wsE.Cells(wsE.Rows.Count, 1).End(xlUp).Row
    For i = 2 To ultima
        encontradas(CStr(wsE.Cells(i, 1).Value) & "|" & _
                    CStr(wsE.Cells(i, 3).Value)) = True
    Next i

    Dim achou As Long, faltou As Long, detalhe As String
    ultima = wsG.Cells(wsG.Rows.Count, 1).End(xlUp).Row
    For i = 2 To ultima
        If encontradas.Exists(CStr(wsG.Cells(i, 1).Value) & "|" & _
                              CStr(wsG.Cells(i, 2).Value)) Then
            achou = achou + 1
        Else
            faltou = faltou + 1
            If faltou <= 10 Then
                detalhe = detalhe & vbCrLf & "  " & wsG.Cells(i, 1).Value & _
                          " matricula " & wsG.Cells(i, 2).Value
            End If
        End If
    Next i

    MsgBox "Gabarito: " & achou & " de " & (achou + faltou) & " detectadas." & _
           IIf(faltou > 0, vbCrLf & vbCrLf & "Nao detectadas:" & detalhe, ""), _
           IIf(faltou = 0, vbInformation, vbExclamation), "Conferencia"
End Sub
