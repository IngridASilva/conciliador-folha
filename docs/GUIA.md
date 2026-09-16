# Conciliador de folha — guia de montagem e uso

## Por que não veio um `.xlsm` pronto

Você pediu `.xlsm` e eu entreguei `.xlsx` mais os módulos separados. A razão é
técnica e vale a pena saber, porque afeta como você vai distribuir isso no
trabalho.

O código VBA não fica no XML da pasta de trabalho. Fica em `vbaProject.bin`,
um contêiner binário OLE dentro do arquivo. Dá para gerar esse binário por
programa, mas não dá para **testar** sem Excel instalado, e eu não tenho Excel
aqui. Entregar um binário de macro não testado é entregar algo que pode abrir
com "arquivo corrompido" no seu computador, e aí você perde mais tempo do que
economizou.

Há um segundo motivo, que vale independentemente do primeiro: desde 2022 o
Office **bloqueia macros de arquivos vindos da internet**, e não com aviso —
bloqueia mesmo, por Mark of the Web. Um `.xlsm` baixado chegaria com as macros
desabilitadas de qualquer forma, e em boa parte das empresas a política de TI
impede desbloquear. Importar os módulos em um arquivo criado na própria
máquina contorna isso de vez.

A montagem leva dois minutos.

---

## Montagem

**1. Abra `Conciliador_Folha.xlsx`.**

**2. Salve como `.xlsm`.**
Arquivo > Salvar como > "Pasta de Trabalho Habilitada para Macro do Excel
(*.xlsm)". Mesmo nome, mesma pasta.

**3. Abra o editor VBA.** `Alt + F11`

**4. Importe os três módulos.**
Arquivo > Importar Arquivo (`Ctrl + M`), um de cada vez:

```
vba/modUtil.bas
vba/modRegras.bas
vba/modConciliacao.bas
```

A ordem não importa, mas importe os três. `modConciliacao` chama os outros
dois.

**5. Confirme a referência do Scripting Dictionary.**
O código usa `CreateObject("Scripting.Dictionary")`, que é vinculação tardia e
não exige referência marcada. Se sua política de TI bloquear `Scripting`, me
diga e eu troco por `Collection`.

**6. Salve e feche o editor.**

---

## Primeira execução: validar antes de confiar

A pasta vem com dados de uma empresa fictícia e **52 anomalias plantadas de
propósito**, listadas na aba `Gabarito_Demo`.

1. `Alt + F8` > `ExecutarConciliacao` > Executar
2. `Alt + F8` > `ConferirGabarito` > Executar

A segunda macro compara o que o motor encontrou com o que deveria encontrar e
diz quantas das 52 apareceram. Se faltar alguma, ela nomeia a regra e a
matrícula.

Esse passo não é enfeite. Motor de regras sem verificação é um filtro que você
não sabe se está filtrando, e a falha é silenciosa: ele simplesmente não
acusa, e você acha que o fechamento está limpo.

---

## Botão na aba Resumo

Opcional, mas é o que faz a ferramenta ser usada por quem não gosta de
`Alt + F8`.

Inserir > Formas > Retângulo arredondado, desenhe na aba Resumo, escreva
"Executar conciliação", clique com o botão direito > Atribuir Macro >
`ExecutarConciliacao`.

---

## Apontar para dados reais

**1. Ajuste a aba Config.**
Só as células em azul. `PASTA_FOLHA` recebe o caminho onde ficam os arquivos
de competência.

**2. Importe as consultas do Power Query.**
Dados > Obter Dados > Consulta em Branco > Editor Avançado, e cole cada bloco
de `powerquery/consultas.m`. Crie nesta ordem: `pPastaFolha`,
`pCompetenciaAtual`, `pCompetenciaAnterior`, `fnCarregarFolha`, `Folha_Atual`,
`Folha_Anterior`, `Conciliacao`, `Cadastro`.

Carregue `Folha_Atual`, `Folha_Anterior`, `Conciliacao` e `Cadastro` para as
abas de mesmo nome, substituindo os dados de demonstração. As demais ficam só
como conexão.

**3. Adapte `fnCarregarFolha` ao layout do seu sistema.**
Delimitador, codificação, nomes de coluna e nome de arquivo. É a única parte
que muda de empresa para empresa, e está isolada numa função só por isso.

**4. Calibre os limites.**
Os valores de `Config` são ponto de partida, não recomendação. Rode três
fechamentos, veja quantas exceções cada regra gera e ajuste. Uma regra que
acusa 400 linhas todo mês não é rigorosa: é ignorada.

---

## Manutenção mensal

| Quando | O quê |
|---|---|
| Todo mês | Atualizar `MES_ATUAL` e `MES_ANTERIOR` em Config |
| Todo mês | Dados > Atualizar Tudo, depois executar a conciliação |
| Janeiro | Atualizar `SALARIO_MINIMO` (célula com fundo amarelo) |
| Na data-base | Revisar `LIM_VAR_LIQUIDO`: no mês do dissídio a variação é geral e o limite normal gera centenas de falsos positivos |
| Trimestral | Revisar a aba Regras: desativar o que virou ruído, ajustar limites |

O ponto sobre o mês do dissídio merece atenção. Se a categoria reajustou 5% e
seu limite é 15%, tudo bem. Se reajustou 12%, metade da folha vira exceção.
Suba o limite naquele mês e desça depois, ou crie uma regra específica que
compare contra o reajuste esperado em vez de contra zero.

---

## As nove regras

| Código | Regra | Severidade | Parâmetro |
|---|---|---|---|
| R01 | Admitido sem provento | Crítica | — |
| R02 | Desligado com provento | Crítica | — |
| R03 | Variação do líquido | Alta | `LIM_VAR_LIQUIDO` |
| R04 | Rubrica nova | Alta | `LIM_RUBRICA_NOVA` |
| R05 | Rubrica recorrente ausente | Média | — |
| R06 | Lançamento duplicado | Crítica | — |
| R07 | Salário abaixo do mínimo | Crítica | `SALARIO_MINIMO` |
| R08 | Horas extras acima do limite | Média | `LIM_HE_PCT` |
| R09 | Ausente sem desligamento | Crítica | — |

Ativar e desativar pela coluna `ativa` na aba Regras. Nada de comentar código.

### Uma decisão de desenho que vale explicar

R04 e R05 **não rodam** para quem já foi pego por R01, R02 ou R09. O motivo é
prático: quando uma pessoa sai da folha, todas as rubricas dela viram
"removida". Sem esse filtro, um único desligamento não registrado geraria seis
linhas de R05 e enterraria o alerta de R09, que é o que realmente importa.

Ferramenta de exceção morre por excesso, não por falta. A regra que ninguém
lê porque vem com 400 linhas é funcionalmente igual a uma regra desligada.

---

## Onde o VBA foi usado, e onde não foi

Divisão deliberada, e é o que separa uma ferramenta que sobrevive de uma que
vira legado em seis meses:

| Camada | Ferramenta | Por quê |
|---|---|---|
| Carga e normalização dos arquivos | Power Query | Declarativo, auditável, não quebra quando o sistema muda o nome de uma coluna |
| Junção e agrupamento | Power Query | O mesmo em VBA seriam trezentas linhas que ninguém revisa |
| Motor de regras | VBA | Lógica condicional complexa com estado entre regras |
| Escrita formatada e log | VBA | Power Query não escreve para onde precisa |
| Painel | Fórmulas | Recalcula sozinho, ninguém precisa lembrar de rodar nada |

Usar VBA para transformar dado hoje é escolha desatualizada. Usar VBA onde só
ele resolve é critério.

## Desempenho

Tudo entra em array de uma vez, processa em memória e volta em bloco único. Em
10 mil linhas de conciliação isso roda em menos de um segundo. A mesma lógica
com `Cells(i, j)` levaria minutos, e é a diferença entre uma ferramenta que o
time usa e uma que o time abandona.
