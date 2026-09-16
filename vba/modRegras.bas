Attribute VB_Name = "modRegras"
'==============================================================================
' modRegras
'
' Uma rotina por regra. Cada uma le o proprio limite da aba Config e o proprio
' texto de acao da aba Regras, entao mudar o comportamento nao exige mexer em
' codigo: exige mexer na planilha, que e onde a pessoa de RH consegue mexer.
'
' Regra desativada na aba Regras simplesmente nao roda. E o mecanismo para
' calibrar a ferramenta nos primeiros fechamentos sem apagar nada.
'==============================================================================
Option Explicit

' Colunas da aba Conciliacao
Private Const C_MATRICULA As Long = 1
Private Const C_RUBRICA As Long = 2
Private Const C_DESCRICAO As Long = 3
Private Const C_TIPO As Long = 4
Private Const C_VAL_ANT As Long = 5
Private Const C_VAL_ATU As Long = 6
Private Const C_QTD As Long = 7
Private Const C_SITUACAO As Long = 10


'------------------------------------------------------------------------------
' R01 - Admitido na competencia sem nenhum provento lancado.
'------------------------------------------------------------------------------
Public Sub AplicarR01(cfg As Object, regras As Object, cadastro As Object, _
                      agregados As Object, explicadas As Object, _
                      saida() As Variant, total As Long)
    If Not modUtil.RegraAtiva(regras, "R01") Then Exit Sub

    Dim primeiroDia As Date
    primeiroDia = modUtil.PrimeiroDiaDaCompetencia(cfg("MES_ATUAL"))

    Dim chave As Variant, info As Variant, agg As Variant
    For Each chave In cadastro.Keys
        info = cadastro(chave)
        If IsDate(info(1)) Then
            If CDate(info(1)) >= primeiroDia Then
                agg = Array(0#, 0#, 0#, 0#, 0#, 0#, 0&)
                If agregados.Exists(CStr(chave)) Then agg = agregados(CStr(chave))
                If agg(1) = 0 Then
                    explicadas(CStr(chave)) = True
                    modUtil.AdicionarExcecao saida, total, "R01", regras, cadastro, _
                        CStr(chave), "", "", 0, 0, _
                        "Admitido em " & Format(CDate(info(1)), "dd/mm/yyyy") & _
                        " e sem provento na competencia"
                End If
            End If
        End If
    Next chave
End Sub


'------------------------------------------------------------------------------
' R02 - Desligado em competencia anterior recebendo provento.
'------------------------------------------------------------------------------
Public Sub AplicarR02(cfg As Object, regras As Object, cadastro As Object, _
                      agregados As Object, explicadas As Object, _
                      saida() As Variant, total As Long)
    If Not modUtil.RegraAtiva(regras, "R02") Then Exit Sub

    Dim primeiroDia As Date
    primeiroDia = modUtil.PrimeiroDiaDaCompetencia(cfg("MES_ATUAL"))

    Dim chave As Variant, info As Variant, agg As Variant
    For Each chave In agregados.Keys
        If cadastro.Exists(CStr(chave)) Then
            info = cadastro(CStr(chave))
            If IsDate(info(2)) Then
                If CDate(info(2)) < primeiroDia Then
                    agg = agregados(CStr(chave))
                    If agg(1) > 0 Then
                        explicadas(CStr(chave)) = True
                        modUtil.AdicionarExcecao saida, total, "R02", regras, cadastro, _
                            CStr(chave), "", "", 0, agg(1), _
                            "Desligado em " & Format(CDate(info(2)), "dd/mm/yyyy") & _
                            " com " & Format(agg(1), "R$ #,##0.00") & " de provento"
                    End If
                End If
            End If
        End If
    Next chave
End Sub


'------------------------------------------------------------------------------
' R03 - Variacao do liquido acima do limite configurado.
'------------------------------------------------------------------------------
Public Sub AplicarR03(cfg As Object, regras As Object, cadastro As Object, _
                      agregados As Object, explicadas As Object, _
                      saida() As Variant, total As Long)
    If Not modUtil.RegraAtiva(regras, "R03") Then Exit Sub

    Dim limite As Double: limite = CDbl(cfg("LIM_VAR_LIQUIDO"))
    Dim chave As Variant, agg As Variant
    Dim liqAnt As Double, liqAtu As Double, variacao As Double

    For Each chave In agregados.Keys
        If Not explicadas.Exists(CStr(chave)) Then
            agg = agregados(CStr(chave))
            liqAnt = agg(0) + agg(2)
            liqAtu = agg(1) + agg(3)

            ' Exige as duas competencias com valor: admissao e desligamento
            ' geram variacao de 100% que ja tem regra propria.
            If liqAnt > 0 And liqAtu > 0 Then
                variacao = liqAtu / liqAnt - 1
                If Abs(variacao) > limite Then
                    modUtil.AdicionarExcecao saida, total, "R03", regras, cadastro, _
                        CStr(chave), "", "Liquido do mes", liqAnt, liqAtu, _
                        "Liquido variou " & Format(variacao, "0.0%") & _
                        " (limite " & Format(limite, "0.0%") & ")"
                End If
            End If
        End If
    Next chave
End Sub


'------------------------------------------------------------------------------
' R09 - Estava na folha anterior, sumiu agora e segue ativo no cadastro.
'------------------------------------------------------------------------------
Public Sub AplicarR09(cfg As Object, regras As Object, cadastro As Object, _
                      agregados As Object, explicadas As Object, _
                      saida() As Variant, total As Long)
    If Not modUtil.RegraAtiva(regras, "R09") Then Exit Sub

    Dim chave As Variant, agg As Variant, info As Variant
    For Each chave In agregados.Keys
        agg = agregados(CStr(chave))
        If agg(0) > 0 And agg(1) = 0 Then
            If cadastro.Exists(CStr(chave)) Then
                info = cadastro(CStr(chave))
                If Not IsDate(info(2)) Then
                    explicadas(CStr(chave)) = True
                    modUtil.AdicionarExcecao saida, total, "R09", regras, cadastro, _
                        CStr(chave), "", "", agg(0), 0, _
                        "Recebia " & Format(agg(0), "R$ #,##0.00") & _
                        " e nao consta na folha atual, sem desligamento no cadastro"
                End If
            End If
        End If
    Next chave
End Sub


'------------------------------------------------------------------------------
' Regras por linha de rubrica: R04, R05, R06, R07, R08.
'
' Percorre a conciliacao uma unica vez. Somar uma varredura por regra seria
' mais legivel e cinco vezes mais lento; com 10 mil linhas a diferenca ainda
' e aceitavel, com 200 mil nao.
'------------------------------------------------------------------------------
Public Sub AplicarPorRubrica(cfg As Object, regras As Object, cadastro As Object, _
                             conc As Variant, explicadas As Object, _
                             saida() As Variant, total As Long)
    Dim limRubrica As Double: limRubrica = CDbl(cfg("LIM_RUBRICA_NOVA"))
    Dim limHE As Double:      limHE = CDbl(cfg("LIM_HE_PCT"))
    Dim salMin As Double:     salMin = CDbl(cfg("SALARIO_MINIMO"))

    Dim r04 As Boolean: r04 = modUtil.RegraAtiva(regras, "R04")
    Dim r05 As Boolean: r05 = modUtil.RegraAtiva(regras, "R05")
    Dim r06 As Boolean: r06 = modUtil.RegraAtiva(regras, "R06")
    Dim r07 As Boolean: r07 = modUtil.RegraAtiva(regras, "R07")
    Dim r08 As Boolean: r08 = modUtil.RegraAtiva(regras, "R08")

    ' R08 precisa do salario base da propria pessoa, entao indexa antes.
    Dim salarios As Object: Set salarios = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = 1 To UBound(conc, 1)
        If CStr(conc(i, C_RUBRICA)) = "001" Then
            salarios(CStr(conc(i, C_MATRICULA))) = CDbl(conc(i, C_VAL_ATU))
        End If
    Next i

    Dim mat As String, rub As String, desc As String
    Dim vAnt As Double, vAtu As Double, situacao As String
    Dim salBase As Double

    For i = 1 To UBound(conc, 1)
        mat = CStr(conc(i, C_MATRICULA))
        rub = CStr(conc(i, C_RUBRICA))
        desc = CStr(conc(i, C_DESCRICAO))
        vAnt = CDbl(conc(i, C_VAL_ANT))
        vAtu = CDbl(conc(i, C_VAL_ATU))
        situacao = CStr(conc(i, C_SITUACAO))

        ' R06 - duplicidade. Vale mesmo para quem ja tem outra excecao:
        ' lancamento duplicado e erro de sistema, nao consequencia de
        ' movimentacao.
        If r06 And CLng(conc(i, C_QTD)) > 1 Then
            modUtil.AdicionarExcecao saida, total, "R06", regras, cadastro, mat, _
                rub, desc, vAnt, vAtu, _
                CLng(conc(i, C_QTD)) & " lancamentos da mesma rubrica na competencia"
        End If

        ' R07 - salario base abaixo do minimo.
        If r07 And rub = "001" And vAtu > 0 And vAtu < salMin Then
            modUtil.AdicionarExcecao saida, total, "R07", regras, cadastro, mat, _
                rub, desc, vAnt, vAtu, _
                "Salario base de " & Format(vAtu, "R$ #,##0.00") & _
                " contra minimo de " & Format(salMin, "R$ #,##0.00")
        End If

        ' R08 - horas extras acima do limite.
        If r08 And rub = "002" And vAtu > 0 Then
            salBase = 0
            If salarios.Exists(mat) Then salBase = salarios(mat)
            If salBase > 0 Then
                If vAtu / salBase > limHE Then
                    modUtil.AdicionarExcecao saida, total, "R08", regras, cadastro, _
                        mat, rub, desc, vAnt, vAtu, _
                        "Horas extras representam " & Format(vAtu / salBase, "0.0%") & _
                        " do salario base (limite " & Format(limHE, "0.0%") & ")"
                End If
            End If
        End If

        ' As duas regras abaixo so fazem sentido para quem nao teve
        ' movimentacao. Sem esse filtro, um desligamento gera seis linhas de
        ' rubrica removida e enterra o alerta que importa.
        If Not explicadas.Exists(mat) Then
            If r04 And situacao = "Nova" And Abs(vAtu) >= limRubrica Then
                modUtil.AdicionarExcecao saida, total, "R04", regras, cadastro, mat, _
                    rub, desc, vAnt, vAtu, _
                    "Rubrica inexistente na competencia anterior, valor " & _
                    Format(Abs(vAtu), "R$ #,##0.00")
            End If

            If r05 And situacao = "Removida" And Abs(vAnt) > 0 Then
                modUtil.AdicionarExcecao saida, total, "R05", regras, cadastro, mat, _
                    rub, desc, vAnt, vAtu, _
                    "Rubrica de " & Format(Abs(vAnt), "R$ #,##0.00") & _
                    " presente na competencia anterior e ausente agora"
            End If
        End If
    Next i
End Sub
