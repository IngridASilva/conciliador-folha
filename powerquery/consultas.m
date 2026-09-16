// ===========================================================================
// Consultas do Power Query para o conciliador de folha.
//
// Cole cada bloco em uma consulta nova (Dados > Obter Dados > Consulta em
// Branco > Editor Avançado).
//
// A pasta de origem vem da aba Config, não de caminho fixo no código. Trocar
// de servidor não pode exigir editar consulta.
// ===========================================================================

// === 1. Parâmetro de origem ===============================================
// Lê o caminho direto da célula nomeada PASTA_FOLHA.
pPastaFolha =
    let
        Config = Excel.CurrentWorkbook(){[Name = "PASTA_FOLHA"]}[Content],
        Caminho = Text.From(Config{0}[Column1])
    in
        Caminho;


pCompetenciaAtual =
    let
        C = Excel.CurrentWorkbook(){[Name = "MES_ATUAL"]}[Content]
    in
        Number.From(C{0}[Column1]);


pCompetenciaAnterior =
    let
        C = Excel.CurrentWorkbook(){[Name = "MES_ANTERIOR"]}[Content]
    in
        Number.From(C{0}[Column1]);


// === 2. Função de carga ===================================================
// Recebe a competência e devolve a folha já normalizada.
//
// O sistema de folha costuma exportar com cabeçalho em linha variável, coluna
// de valor como texto com vírgula decimal e desconto ora positivo ora
// negativo. Essa normalização é a maior parte do trabalho e é justamente o
// que não deve viver em VBA: Power Query faz isso melhor e deixa rastro.
fnCarregarFolha = (competencia as number) as table =>
    let
        Arquivo  = pPastaFolha & "\FOLHA_" & Text.From(competencia) & ".csv",
        Origem   = Csv.Document(
            File.Contents(Arquivo),
            [Delimiter = ";", Encoding = 1252, QuoteStyle = QuoteStyle.Csv]
        ),
        Cabecalho = Table.PromoteHeaders(Origem, [PromoteAllScalars = true]),

        // Descarta linhas de totalização que muitos sistemas incluem no meio
        // do arquivo. Sem isso o total dobra e ninguém percebe até o fim.
        SemTotais = Table.SelectRows(
            Cabecalho,
            each [matricula] <> null
                and not Text.StartsWith(Text.From([matricula]), "TOTAL")
        ),

        Tipos = Table.TransformColumnTypes(SemTotais, {
            {"matricula", Int64.Type},
            {"cod_rubrica", type text},
            {"descricao_rubrica", type text},
            {"tipo", type text},
            {"valor", type number}
        }),

        // Padroniza o sinal: desconto sempre negativo. Sistema que exporta
        // desconto positivo faz a conciliação somar o que deveria subtrair.
        Sinal = Table.AddColumn(Tipos, "valor_ajustado",
            each if [tipo] = "Desconto" and [valor] > 0
                 then -[valor] else [valor], type number),

        Final = Table.SelectColumns(
            Table.RenameColumns(
                Table.RemoveColumns(Sinal, {"valor"}),
                {{"valor_ajustado", "valor"}}
            ),
            {"matricula", "cod_rubrica", "descricao_rubrica", "tipo", "valor"}
        ),
        ComCompetencia = Table.AddColumn(Final, "id_mes",
            each competencia, Int64.Type)
    in
        ComCompetencia;


// === 3. As duas competências ==============================================
Folha_Atual     = fnCarregarFolha(pCompetenciaAtual);
Folha_Anterior  = fnCarregarFolha(pCompetenciaAnterior);


// === 4. Conciliação =======================================================
// Junção externa completa por matrícula e rubrica. É a etapa que o VBA não
// deve fazer: agrupamento e junção em Power Query são declarativos e
// auditáveis; em VBA viram trezentas linhas que ninguém revisa.
Conciliacao =
    let
        ChaveAtual = Table.Group(
            Folha_Atual,
            {"matricula", "cod_rubrica", "descricao_rubrica", "tipo"},
            {
                {"valor_atual", each List.Sum([valor]), type number},
                {"qtd_lancamentos", each Table.RowCount(_), Int64.Type}
            }
        ),
        ChaveAnterior = Table.Group(
            Folha_Anterior,
            {"matricula", "cod_rubrica", "descricao_rubrica", "tipo"},
            {{"valor_anterior", each List.Sum([valor]), type number}}
        ),

        Juntado = Table.NestedJoin(
            ChaveAnterior,
            {"matricula", "cod_rubrica", "descricao_rubrica", "tipo"},
            ChaveAtual,
            {"matricula", "cod_rubrica", "descricao_rubrica", "tipo"},
            "atual", JoinKind.FullOuter
        ),
        Expandido = Table.ExpandTableColumn(
            Juntado, "atual",
            {"matricula", "cod_rubrica", "descricao_rubrica", "tipo",
             "valor_atual", "qtd_lancamentos"},
            {"m2", "r2", "d2", "t2", "valor_atual", "qtd_lancamentos"}
        ),

        // Junção externa completa deixa a chave nula do lado que não existe.
        Chaves = Table.AddColumn(Expandido, "chave_matricula",
            each if [matricula] = null then [m2] else [matricula], Int64.Type),
        Chaves2 = Table.AddColumn(Chaves, "chave_rubrica",
            each if [cod_rubrica] = null then [r2] else [cod_rubrica], type text),
        Chaves3 = Table.AddColumn(Chaves2, "chave_descricao",
            each if [descricao_rubrica] = null then [d2]
                 else [descricao_rubrica], type text),
        Chaves4 = Table.AddColumn(Chaves3, "chave_tipo",
            each if [tipo] = null then [t2] else [tipo], type text),

        Zeros = Table.TransformColumns(Chaves4, {
            {"valor_anterior", each if _ = null then 0 else _, type number},
            {"valor_atual", each if _ = null then 0 else _, type number},
            {"qtd_lancamentos", each if _ = null then 0 else _, Int64.Type}
        }),

        Calculos = Table.AddColumn(Zeros, "variacao_abs",
            each [valor_atual] - [valor_anterior], type number),
        Calculos2 = Table.AddColumn(Calculos, "variacao_pct",
            each if Number.Abs([valor_anterior]) = 0 then null
                 else Number.Abs([valor_atual])
                      / Number.Abs([valor_anterior]) - 1, type number),
        Calculos3 = Table.AddColumn(Calculos2, "situacao",
            each if [valor_anterior] = 0 then "Nova"
                 else if [valor_atual] = 0 then "Removida"
                 else if [variacao_abs] = 0 then "Inalterada"
                 else "Alterada", type text),

        Final = Table.SelectColumns(Calculos3, {
            "chave_matricula", "chave_rubrica", "chave_descricao", "chave_tipo",
            "valor_anterior", "valor_atual", "qtd_lancamentos",
            "variacao_abs", "variacao_pct", "situacao"
        }),
        Renomeado = Table.RenameColumns(Final, {
            {"chave_matricula", "matricula"},
            {"chave_rubrica", "cod_rubrica"},
            {"chave_descricao", "descricao_rubrica"},
            {"chave_tipo", "tipo"}
        }),
        Ordenado = Table.Sort(Renomeado, {
            {"matricula", Order.Ascending}, {"cod_rubrica", Order.Ascending}
        })
    in
        Ordenado;


// === 5. Cadastro ==========================================================
// Ajuste a origem conforme o extrato do seu sistema de RH.
Cadastro =
    let
        Arquivo = pPastaFolha & "\CADASTRO.csv",
        Origem = Csv.Document(
            File.Contents(Arquivo),
            [Delimiter = ";", Encoding = 1252, QuoteStyle = QuoteStyle.Csv]
        ),
        Cabecalho = Table.PromoteHeaders(Origem, [PromoteAllScalars = true]),
        Tipos = Table.TransformColumnTypes(Cabecalho, {
            {"matricula", Int64.Type},
            {"nome", type text},
            {"status", type text},
            {"data_admissao", type date},
            {"data_desligamento", type date},
            {"gerencia", type text},
            {"centro_custo", type text}
        })
    in
        Tipos;
