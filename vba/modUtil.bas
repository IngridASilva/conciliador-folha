Attribute VB_Name = "modUtil"
'==============================================================================
' modUtil
'
' Leitura de configuracao, acesso a dados e registro de execucao.
' Nenhuma regra de negocio mora aqui.
'==============================================================================
Option Explicit

Private Const E_COLUNAS As Long = 14


'------------------------------------------------------------------------------
' Le a aba Config pelos nomes definidos, nao por endereco de celula.
' Inserir uma linha em Config nao pode quebrar o codigo.
'------------------------------------------------------------------------------
Public Function LerConfig() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("Config")
    Dim ultima As Long, i As Long, nome As String

    ultima = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To ultima
        nome = Trim$(CStr(ws.Cells(i, 3).Value))
        If Len(nome) > 0 Then d(nome) = ws.Cells(i, 2).Value
    Next i

    Set LerConfig = d
End Function


'------------------------------------------------------------------------------
' Le a aba Regras. Item: Array(nome, severidade, ativa, acao_sugerida)
'------------------------------------------------------------------------------
Public Function LerRegras() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("Regras")
    Dim ultima As Long, i As Long, cod As String

    ultima = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To ultima
        cod = Trim$(CStr(ws.Cells(i, 1).Value))
        If Len(cod) > 0 Then
            d(cod) = Array(CStr(ws.Cells(i, 2).Value), _
                           CStr(ws.Cells(i, 3).Value), _
                           UCase$(Trim$(CStr(ws.Cells(i, 4).Value))) = "SIM", _
                           CStr(ws.Cells(i, 6).Value))
        End If
    Next i

    Set LerRegras = d
End Function


Public Function RegraAtiva(regras As Object, cod As String) As Boolean
    If Not regras.Exists(cod) Then
        RegraAtiva = False
    Else
        RegraAtiva = CBool(regras(cod)(2))
    End If
End Function


'------------------------------------------------------------------------------
' Le a aba Cadastro. Item: Array(nome, admissao, desligamento, gerencia, cc)
'------------------------------------------------------------------------------
Public Function LerCadastro() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("Cadastro")
    Dim dados As Variant
    dados = LerIntervalo(ws)
    If IsEmpty(dados) Then Set LerCadastro = d: Exit Function

    Dim i As Long
    For i = 1 To UBound(dados, 1)
        d(CStr(dados(i, 1))) = Array(CStr(dados(i, 2)), _
                                     dados(i, 4), dados(i, 5), _
                                     CStr(dados(i, 7)), CStr(dados(i, 8)))
    Next i

    Set LerCadastro = d
End Function


'------------------------------------------------------------------------------
' Le o bloco de dados de uma aba (sem cabecalho) em um array 2D.
' Devolve Empty quando so existe o cabecalho.
'------------------------------------------------------------------------------
Public Function LerIntervalo(ws As Worksheet) As Variant
    Dim ultimaLinha As Long, ultimaColuna As Long
    ultimaLinha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    ultimaColuna = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    If ultimaLinha < 2 Then
        LerIntervalo = Empty
    ElseIf ultimaLinha = 2 Then
        ' Uma unica linha nao vira array 2D automaticamente.
        Dim umaLinha() As Variant
        ReDim umaLinha(1 To 1, 1 To ultimaColuna)
        Dim c As Long
        For c = 1 To ultimaColuna
            umaLinha(1, c) = ws.Cells(2, c).Value
        Next c
        LerIntervalo = umaLinha
    Else
        LerIntervalo = ws.Range(ws.Cells(2, 1), _
                                ws.Cells(ultimaLinha, ultimaColuna)).Value
    End If
End Function


Public Function PrimeiroDiaDaCompetencia(competencia As Variant) As Date
    Dim t As String: t = CStr(competencia)
    PrimeiroDiaDaCompetencia = DateSerial(CInt(Left$(t, 4)), CInt(Right$(t, 2)), 1)
End Function


'------------------------------------------------------------------------------
' Acrescenta uma excecao ao buffer de saida.
'------------------------------------------------------------------------------
Public Sub AdicionarExcecao(saida() As Variant, total As Long, cod As String, _
                            regras As Object, cadastro As Object, _
                            matricula As String, rubrica As String, _
                            descricao As String, valorAnterior As Double, _
                            valorAtual As Double, detalhe As String)
    If total >= UBound(saida, 1) Then
        ReDim Preserve saida(1 To UBound(saida, 1) + 20000, 1 To E_COLUNAS)
    End If

    total = total + 1

    Dim r As Variant: r = regras(cod)
    Dim nome As String, gerencia As String, cc As String
    If cadastro.Exists(matricula) Then
        nome = cadastro(matricula)(0)
        gerencia = cadastro(matricula)(3)
        cc = cadastro(matricula)(4)
    End If

    saida(total, 1) = cod
    saida(total, 2) = r(1)
    saida(total, 3) = CLng(Val(matricula))
    saida(total, 4) = nome
    saida(total, 5) = gerencia
    saida(total, 6) = cc
    saida(total, 7) = rubrica
    saida(total, 8) = descricao
    saida(total, 9) = valorAnterior
    saida(total, 10) = valorAtual
    saida(total, 11) = valorAtual - valorAnterior
    saida(total, 12) = detalhe
    saida(total, 13) = r(3)
    saida(total, 14) = "Pendente"
End Sub


'------------------------------------------------------------------------------
' Checagens que evitam rodar sobre dado errado.
'------------------------------------------------------------------------------
Public Function ValidarAmbiente(cfg As Object) As Boolean
    ValidarAmbiente = False

    Dim obrigatorios As Variant
    obrigatorios = Array("MES_ATUAL", "MES_ANTERIOR", "LIM_VAR_LIQUIDO", _
                         "LIM_RUBRICA_NOVA", "LIM_HE_PCT", "SALARIO_MINIMO")

    Dim i As Long
    For i = LBound(obrigatorios) To UBound(obrigatorios)
        If Not cfg.Exists(CStr(obrigatorios(i))) Then
            MsgBox "Parametro ausente na aba Config: " & obrigatorios(i), _
                   vbCritical, "Conciliador"
            Exit Function
        End If
        If Len(Trim$(CStr(cfg(CStr(obrigatorios(i)))))) = 0 Then
            MsgBox "Parametro sem valor na aba Config: " & obrigatorios(i), _
                   vbCritical, "Conciliador"
            Exit Function
        End If
    Next i

    If CLng(cfg("MES_ATUAL")) <= CLng(cfg("MES_ANTERIOR")) Then
        MsgBox "A competencia atual precisa ser posterior a anterior." & vbCrLf & _
               "Atual: " & cfg("MES_ATUAL") & "   Anterior: " & cfg("MES_ANTERIOR"), _
               vbCritical, "Conciliador"
        Exit Function
    End If

    ' Limite muito alto silencia a ferramenta sem avisar ninguem.
    If CDbl(cfg("LIM_VAR_LIQUIDO")) > 0.5 Then
        If MsgBox("O limite de variacao esta em " & _
                  Format(cfg("LIM_VAR_LIQUIDO"), "0%") & _
                  ", alto o bastante para deixar passar erro relevante." & _
                  vbCrLf & vbCrLf & "Executar assim mesmo?", _
                  vbYesNo + vbExclamation, "Conciliador") = vbNo Then Exit Function
    End If

    ValidarAmbiente = True
End Function


Public Sub RegistrarLog(cfg As Object, linhasConciliacao As Long, _
                        excecoes As Long, criticas As Long, segundos As Single)
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("Log")
    Dim linha As Long
    linha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1

    Dim responsavel As String
    responsavel = CStr(cfg("RESPONSAVEL"))
    If Len(Trim$(responsavel)) = 0 Or responsavel = "preencher" Then
        responsavel = Environ$("USERNAME")
    End If

    ws.Cells(linha, 1).Value = Now
    ws.Cells(linha, 1).NumberFormat = "dd/mm/yyyy hh:mm"
    ws.Cells(linha, 2).Value = responsavel
    ws.Cells(linha, 3).Value = cfg("MES_ATUAL")
    ws.Cells(linha, 4).Value = linhasConciliacao
    ws.Cells(linha, 5).Value = excecoes
    ws.Cells(linha, 6).Value = criticas
    ws.Cells(linha, 7).Value = Round(segundos, 1)
    ws.Cells(linha, 8).Value = modConciliacao.VERSAO
    ws.Range(ws.Cells(linha, 1), ws.Cells(linha, 8)).Font.Name = "Arial"
    ws.Range(ws.Cells(linha, 1), ws.Cells(linha, 8)).Font.Size = 10
End Sub
