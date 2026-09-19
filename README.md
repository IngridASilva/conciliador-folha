# payroll-reconciler

Conciliador de folha de pagamento em Excel, Power Query e VBA. Compara duas
competências, aplica nove regras de exceção e devolve a lista do que precisa
ser conferido antes do fechamento.

Dados de demonstração vindos de
[hr-synthetic-data-br](../hr-synthetic-data-br), com **52 anomalias plantadas
de propósito** para validar o motor de regras.

## Começar

1. Abra `Conciliador_Folha.xlsx` e salve como `.xlsm`
2. `Alt + F11`, importe `vba/modUtil.bas`, `vba/modRegras.bas`,
   `vba/modConciliacao.bas`
3. `Alt + F8` > `ExecutarConciliacao`
4. `Alt + F8` > `ConferirGabarito` — confere se as 52 anomalias apareceram

Passo a passo completo em [`docs/GUIA.md`](docs/GUIA.md).

> O arquivo vem como `.xlsx` com os módulos separados, e não como `.xlsm`
> pronto. O motivo está no guia: o Office bloqueia macro de arquivo baixado da
> internet desde 2022, então importar os módulos num arquivo criado na própria
> máquina é o único caminho que funciona de verdade.

## As nove regras

| Código | Regra | Severidade |
|---|---|---|
| R01 | Admitido sem provento | Crítica |
| R02 | Desligado com provento | Crítica |
| R03 | Variação do líquido acima do limite | Alta |
| R04 | Rubrica nova de valor relevante | Alta |
| R05 | Rubrica recorrente ausente | Média |
| R06 | Lançamento duplicado | Crítica |
| R07 | Salário abaixo do mínimo | Crítica |
| R08 | Horas extras acima do limite | Média |
| R09 | Ausente na folha sem desligamento | Crítica |

Ativar e desativar pela coluna `ativa` na aba Regras. Nenhum código é
comentado para mudar comportamento.

## A divisão entre Power Query e VBA

| Camada | Ferramenta | Por quê |
|---|---|---|
| Carga e normalização | Power Query | Declarativo, auditável, não quebra quando o sistema renomeia uma coluna |
| Junção e agrupamento | Power Query | O mesmo em VBA seriam trezentas linhas que ninguém revisa |
| Motor de regras | VBA | Lógica condicional com estado entre regras |
| Escrita formatada e log | VBA | Power Query não escreve para onde precisa |
| Painel | Fórmulas | Recalcula sozinho |

Usar VBA para transformar dado hoje é escolha desatualizada. Usar VBA onde só
ele resolve é critério.

## Uma decisão de desenho

R04 e R05 não rodam para quem já foi pego por R01, R02 ou R09. Quando alguém
sai da folha, todas as rubricas viram "removida" — sem esse filtro, um
desligamento não registrado geraria seis linhas de ruído e enterraria o alerta
que importa.

Ferramenta de exceção morre por excesso, não por falta.

## Desempenho

Tudo entra em array de uma vez, processa em memória e volta em bloco único.
Dez mil linhas de conciliação rodam em menos de um segundo. A mesma lógica com
`Cells(i, j)` levaria minutos.

## Estrutura

```
Conciliador_Folha.xlsx   pasta de trabalho com dados de demonstração
vba/                     três módulos para importar
powerquery/consultas.m   consultas de carga, para apontar aos dados reais
build/                   scripts que geraram os dados de demonstração
docs/GUIA.md             montagem, calibração e manutenção mensal
```
