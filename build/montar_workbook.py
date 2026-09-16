"""Monta a pasta de trabalho do conciliador de folha."""

from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

B = Path("/home/claude/conciliador-folha/build")
SAIDA = Path("/home/claude/conciliador-folha/Conciliador_Folha.xlsx")

FONTE = "Arial"
AZUL = "1F3864"      # cabeçalhos
AZUL_INPUT = "0000FF"  # texto de célula de entrada
CINZA = "F2F2F2"
AMARELO = "FFF2CC"

MOEDA = 'R$ #,##0.00;[Red](R$ #,##0.00);"-"'
PCT = "0.0%"

wb = Workbook()
wb.remove(wb.active)


def estilo_cabecalho(ws, linha=1, n_col=None):
    n_col = n_col or ws.max_column
    for c in range(1, n_col + 1):
        cel = ws.cell(row=linha, column=c)
        cel.font = Font(name=FONTE, bold=True, color="FFFFFF", size=10)
        cel.fill = PatternFill("solid", fgColor=AZUL)
        cel.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.freeze_panes = ws.cell(row=linha + 1, column=1)
    ws.row_dimensions[linha].height = 30


def largura(ws, larguras: dict):
    for col, w in larguras.items():
        ws.column_dimensions[col].width = w


def escrever_df(ws, df: pd.DataFrame, formatos: dict | None = None,
                nome_tabela: str | None = None):
    ws.append(list(df.columns))
    for linha in df.itertuples(index=False):
        ws.append(list(linha))
    estilo_cabecalho(ws, 1, len(df.columns))

    for c in range(1, len(df.columns) + 1):
        letra = get_column_letter(c)
        fmt = (formatos or {}).get(df.columns[c - 1])
        for r in range(2, len(df) + 2):
            cel = ws.cell(row=r, column=c)
            cel.font = Font(name=FONTE, size=10)
            if fmt:
                cel.number_format = fmt

    if nome_tabela and len(df):
        ref = f"A1:{get_column_letter(len(df.columns))}{len(df) + 1}"
        t = Table(displayName=nome_tabela, ref=ref)
        t.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2", showRowStripes=True
        )
        ws.add_table(t)


# =========================================================================
# 1. Leia-me
# =========================================================================
ws = wb.create_sheet("Leia-me")
conteudo = [
    ("CONCILIADOR DE FOLHA DE PAGAMENTO", "titulo"),
    ("", None),
    ("O que faz", "h"),
    ("Compara duas competências da folha, aplica um conjunto de regras de exceção "
     "e devolve a lista do que precisa ser conferido antes do fechamento.", None),
    ("", None),
    ("Como usar", "h"),
    ("1. Ajuste os parâmetros na aba Config. Só as células em azul são editáveis.", None),
    ("2. Ative ou desative regras na aba Regras, coluna Ativa.", None),
    ("3. Atualize as consultas (Dados > Atualizar Tudo) para carregar os arquivos "
     "do sistema de folha. A pasta de origem fica em Config.", None),
    ("4. Execute a conciliação: Alt+F8 > ExecutarConciliacao, ou o botão na aba Resumo.", None),
    ("5. O resultado aparece em Excecoes e o painel em Resumo.", None),
    ("", None),
    ("Antes de rodar em dados reais", "h"),
    ("Esta pasta vem com dados de demonstração de uma empresa fictícia e com 52 "
     "anomalias injetadas de propósito, listadas na aba Gabarito_Demo.", None),
    ("Rode a conciliação uma vez e confira se as 52 aparecem. É a forma de validar "
     "que o motor de regras está funcionando antes de confiar nele.", None),
    ("", None),
    ("Legenda de cores", "h"),
    ("Azul: célula de entrada, pode editar.", "input"),
    ("Preto: resultado de fórmula ou carga, não edite.", None),
    ("Fundo amarelo: premissa que exige revisão a cada competência.", "premissa"),
    ("", None),
    ("Abas", "h"),
    ("Config            parâmetros e limites", None),
    ("Regras            catálogo das regras de exceção", None),
    ("Cadastro          situação cadastral por matrícula", None),
    ("Folha_Anterior    competência de referência", None),
    ("Folha_Atual       competência em fechamento", None),
    ("Conciliacao       cruzamento por matrícula e rubrica", None),
    ("Excecoes          saída do motor de regras", None),
    ("Resumo            painel de acompanhamento", None),
    ("Log               histórico de execuções", None),
    ("Gabarito_Demo     anomalias plantadas nos dados de exemplo", None),
]
for i, (texto, tipo) in enumerate(conteudo, start=1):
    cel = ws.cell(row=i, column=1, value=texto)
    if tipo == "titulo":
        cel.font = Font(name=FONTE, bold=True, size=16, color=AZUL)
    elif tipo == "h":
        cel.font = Font(name=FONTE, bold=True, size=11, color=AZUL)
    elif tipo == "input":
        cel.font = Font(name=FONTE, size=10, color=AZUL_INPUT)
    elif tipo == "premissa":
        cel.font = Font(name=FONTE, size=10)
        cel.fill = PatternFill("solid", fgColor=AMARELO)
    else:
        cel.font = Font(name=FONTE, size=10)
largura(ws, {"A": 100})

# =========================================================================
# 2. Config
# =========================================================================
ws = wb.create_sheet("Config")
params = [
    ("Parâmetro", "Valor", "Nome definido", "Observação"),
    ("Competência anterior", 202511, "MES_ANTERIOR", "Formato AAAAMM"),
    ("Competência atual", 202512, "MES_ATUAL", "Formato AAAAMM"),
    ("Limite de variação do líquido", 0.15, "LIM_VAR_LIQUIDO",
     "Acima disso vira exceção. 15% é ponto de partida; calibre com 2 ou 3 fechamentos."),
    ("Valor mínimo de rubrica nova", 500.0, "LIM_RUBRICA_NOVA",
     "Rubrica nova abaixo deste valor é ignorada, para não gerar ruído."),
    ("Limite de horas extras sobre salário", 0.40, "LIM_HE_PCT",
     "Percentual do salário base a partir do qual o lançamento é revisado."),
    ("Salário mínimo vigente", 1518.0, "SALARIO_MINIMO",
     "ATUALIZAR EM JANEIRO. Valor de referência nacional; piso de categoria pode ser maior."),
    ("Meses para considerar rubrica recorrente", 3, "MIN_MESES_RECORRENTE",
     "Usado pela regra de rubrica que desapareceu."),
    ("Pasta dos arquivos de folha", r"C:\Folha\Competencias", "PASTA_FOLHA",
     "Origem das consultas do Power Query."),
    ("Responsável pela conciliação", "preencher", "RESPONSAVEL",
     "Registrado no log de cada execução."),
]
for linha in params:
    ws.append(list(linha))
estilo_cabecalho(ws, 1, 4)

formatos_config = {4: PCT, 5: MOEDA, 6: PCT, 7: MOEDA}
for r in range(2, len(params) + 1):
    for c in range(1, 5):
        ws.cell(row=r, column=c).font = Font(name=FONTE, size=10)
    cel = ws.cell(row=r, column=2)
    cel.font = Font(name=FONTE, size=10, color=AZUL_INPUT, bold=True)
    cel.alignment = Alignment(horizontal="center")
    if r in formatos_config:
        cel.number_format = formatos_config[r]
    ws.cell(row=r, column=3).font = Font(name=FONTE, size=9, italic=True)
    ws.cell(row=r, column=4).alignment = Alignment(wrap_text=True, vertical="top")

# Premissa que exige revisão anual.
ws.cell(row=7, column=2).fill = PatternFill("solid", fgColor=AMARELO)
largura(ws, {"A": 40, "B": 18, "C": 24, "D": 62})

for r in range(2, len(params) + 1):
    nome = ws.cell(row=r, column=3).value
    wb.defined_names.add(
        __import__("openpyxl").workbook.defined_name.DefinedName(
            nome, attr_text=f"Config!$B${r}"
        )
    )

# =========================================================================
# 3. Regras
# =========================================================================
ws = wb.create_sheet("Regras")
regras = pd.DataFrame([
    ("R01", "Admitido sem provento", "Crítica", "SIM",
     "Admitido na competência sem nenhum provento lançado",
     "Verificar se a admissão foi integrada ao sistema de folha"),
    ("R02", "Desligado com provento", "Crítica", "SIM",
     "Desligado em competência anterior recebendo provento",
     "Bloquear pagamento e apurar origem do lançamento"),
    ("R03", "Variação do líquido", "Alta", "SIM",
     "Variação do líquido acima de LIM_VAR_LIQUIDO",
     "Conferir movimentação, afastamento ou lançamento retroativo"),
    ("R04", "Rubrica nova", "Alta", "SIM",
     "Rubrica inexistente na competência anterior, acima de LIM_RUBRICA_NOVA",
     "Validar autorização do lançamento"),
    ("R05", "Rubrica recorrente ausente", "Média", "SIM",
     "Rubrica presente nos últimos MIN_MESES_RECORRENTE meses e ausente agora",
     "Confirmar se houve cancelamento formal do benefício"),
    ("R06", "Lançamento duplicado", "Crítica", "SIM",
     "Mesma matrícula e rubrica lançadas mais de uma vez",
     "Estornar a duplicidade antes do fechamento"),
    ("R07", "Salário abaixo do mínimo", "Crítica", "SIM",
     "Salário base inferior a SALARIO_MINIMO",
     "Corrigir. Risco trabalhista direto"),
    ("R08", "Horas extras acima do limite", "Média", "SIM",
     "Horas extras acima de LIM_HE_PCT do salário base",
     "Conferir apontamento e alertar a gestão da área"),
    ("R09", "Ausente sem desligamento", "Crítica", "SIM",
     "Estava na folha anterior, sumiu agora e segue ativo no cadastro",
     "Risco de falta de pagamento. Conferir antes de transmitir"),
], columns=["cod_regra", "nome", "severidade", "ativa", "descricao", "acao_sugerida"])
escrever_df(ws, regras, nome_tabela="tblRegras")
for r in range(2, len(regras) + 2):
    ws.cell(row=r, column=4).font = Font(name=FONTE, size=10, color=AZUL_INPUT, bold=True)
    ws.cell(row=r, column=4).alignment = Alignment(horizontal="center")
    for c in (5, 6):
        ws.cell(row=r, column=c).alignment = Alignment(wrap_text=True, vertical="top")
largura(ws, {"A": 11, "B": 30, "C": 12, "D": 9, "E": 58, "F": 52})

# =========================================================================
# 4-6. Dados
# =========================================================================
cad = pd.read_parquet(B / "cadastro.parquet")
cad["data_admissao"] = pd.to_datetime(cad["data_admissao"]).dt.date
cad["data_desligamento"] = pd.to_datetime(cad["data_desligamento"]).dt.date
cad = cad[["matricula", "nome", "status", "data_admissao", "data_desligamento",
           "id_area", "gerencia", "centro_custo", "grade"]]
ws = wb.create_sheet("Cadastro")
escrever_df(ws, cad, {"data_admissao": "DD/MM/AAAA", "data_desligamento": "DD/MM/AAAA"},
            "tblCadastro")
largura(ws, {"A": 12, "B": 26, "C": 12, "D": 14, "E": 16, "F": 9, "G": 24, "H": 14, "I": 8})

for nome_aba, arq in [("Folha_Anterior", "folha_anterior.parquet"),
                      ("Folha_Atual", "folha_atual.parquet")]:
    df = pd.read_parquet(B / arq)[
        ["id_mes", "matricula", "cod_rubrica", "descricao_rubrica", "tipo", "valor"]
    ]
    ws = wb.create_sheet(nome_aba)
    escrever_df(ws, df, {"valor": MOEDA}, f"tbl{nome_aba.replace('_', '')}")
    largura(ws, {"A": 11, "B": 12, "C": 12, "D": 34, "E": 12, "F": 16})

# =========================================================================
# 7. Conciliacao (produzida pelo Power Query; materializada aqui)
# =========================================================================
ant = pd.read_parquet(B / "folha_anterior.parquet")
atu = pd.read_parquet(B / "folha_atual.parquet")
a = ant.groupby(["matricula", "cod_rubrica", "descricao_rubrica", "tipo"], as_index=False)[
    "valor"].sum().rename(columns={"valor": "valor_anterior"})
b = atu.groupby(["matricula", "cod_rubrica", "descricao_rubrica", "tipo"], as_index=False).agg(
    valor_atual=("valor", "sum"), qtd_lancamentos=("valor", "size"))
conc = a.merge(b, on=["matricula", "cod_rubrica", "descricao_rubrica", "tipo"], how="outer")
conc["valor_anterior"] = conc["valor_anterior"].fillna(0.0).round(2)
conc["valor_atual"] = conc["valor_atual"].fillna(0.0).round(2)
conc["qtd_lancamentos"] = conc["qtd_lancamentos"].fillna(0).astype(int)
conc["variacao_abs"] = (conc["valor_atual"] - conc["valor_anterior"]).round(2)
import numpy as np
base_ant = conc["valor_anterior"].abs().to_numpy(dtype=float)
conc["variacao_pct"] = np.where(
    base_ant > 0, conc["valor_atual"].abs().to_numpy(dtype=float) / np.where(base_ant > 0, base_ant, 1) - 1, np.nan
).round(4)
conc["situacao"] = "Alterada"
conc.loc[conc["valor_anterior"] == 0, "situacao"] = "Nova"
conc.loc[conc["valor_atual"] == 0, "situacao"] = "Removida"
conc.loc[conc["variacao_abs"] == 0, "situacao"] = "Inalterada"
conc = conc.sort_values(["matricula", "cod_rubrica"]).reset_index(drop=True)

ws = wb.create_sheet("Conciliacao")
escrever_df(ws, conc, {"valor_anterior": MOEDA, "valor_atual": MOEDA,
                       "variacao_abs": MOEDA, "variacao_pct": PCT}, "tblConciliacao")
largura(ws, {"A": 12, "B": 12, "C": 34, "D": 11, "E": 16, "F": 16,
             "G": 14, "H": 14, "I": 13, "J": 13})

# =========================================================================
# 8. Excecoes (preenchida pelo VBA)
# =========================================================================
ws = wb.create_sheet("Excecoes")
cabecalhos = ["cod_regra", "severidade", "matricula", "nome", "gerencia",
              "centro_custo", "cod_rubrica", "descricao_rubrica", "valor_anterior",
              "valor_atual", "variacao", "detalhe", "acao_sugerida", "status_tratativa"]
ws.append(cabecalhos)
estilo_cabecalho(ws, 1, len(cabecalhos))
largura(ws, {"A": 11, "B": 12, "C": 12, "D": 26, "E": 24, "F": 14, "G": 11,
             "H": 32, "I": 15, "J": 15, "K": 15, "L": 52, "M": 46, "N": 18})
ws.auto_filter.ref = f"A1:{get_column_letter(len(cabecalhos))}1"
ws.conditional_formatting.add(
    "B2:B20000",
    CellIsRule(operator="equal", formula=['"Crítica"'],
               fill=PatternFill("solid", fgColor="F8CBAD"), font=Font(bold=True)),
)
ws.conditional_formatting.add(
    "B2:B20000",
    CellIsRule(operator="equal", formula=['"Alta"'],
               fill=PatternFill("solid", fgColor="FFE699")),
)

# =========================================================================
# 9. Resumo (fórmulas sobre Excecoes)
# =========================================================================
ws = wb.create_sheet("Resumo")
fina = Side(style="thin", color="BFBFBF")
borda = Border(left=fina, right=fina, top=fina, bottom=fina)

ws["A1"] = "CONCILIAÇÃO DE FOLHA"
ws["A1"].font = Font(name=FONTE, bold=True, size=16, color=AZUL)
ws["A2"] = "=\"Competência \" & TEXT(MES_ATUAL,\"0000-00\") & \" contra \" & TEXT(MES_ANTERIOR,\"0000-00\")"
ws["A2"].font = Font(name=FONTE, size=10, italic=True)

blocos = [
    ("A4", "Exceções encontradas", "=COUNTA(Excecoes!A2:A20000)", "#,##0"),
    ("C4", "Críticas", '=COUNTIF(Excecoes!B2:B20000,"Crítica")', "#,##0"),
    ("E4", "Altas", '=COUNTIF(Excecoes!B2:B20000,"Alta")', "#,##0"),
    ("G4", "Médias", '=COUNTIF(Excecoes!B2:B20000,"Média")', "#,##0"),
]
for celula, rotulo, formula, fmt in blocos:
    col = celula[0]
    ws[f"{col}4"] = rotulo
    ws[f"{col}4"].font = Font(name=FONTE, size=9, bold=True, color="595959")
    ws[f"{col}5"] = formula
    ws[f"{col}5"].font = Font(name=FONTE, size=20, bold=True, color=AZUL)
    ws[f"{col}5"].number_format = fmt
    ws[f"{col}5"].alignment = Alignment(horizontal="left")

ws["A7"] = "Pessoas envolvidas"
ws["A8"] = "=IFERROR(SUMPRODUCT((Excecoes!C2:C20000<>\"\")/COUNTIF(Excecoes!C2:C20000,Excecoes!C2:C20000&\"\")),0)"
ws["C7"] = "Valor financeiro em exceção"
ws["C8"] = "=SUMPRODUCT(ABS(Excecoes!J2:J20000))"
ws["E7"] = "Total da folha atual"
ws["E8"] = "=SUMIF(Folha_Atual!E2:E20000,\"Provento\",Folha_Atual!F2:F20000)"
ws["G7"] = "Variação contra anterior"
ws["G8"] = "=IFERROR(E8/SUMIF(Folha_Anterior!E2:E20000,\"Provento\",Folha_Anterior!F2:F20000)-1,0)"
for col in "ACEG":
    ws[f"{col}7"].font = Font(name=FONTE, size=9, bold=True, color="595959")
    ws[f"{col}8"].font = Font(name=FONTE, size=13, bold=True)
ws["A8"].number_format = "#,##0"
ws["C8"].number_format = MOEDA
ws["E8"].number_format = MOEDA
ws["G8"].number_format = PCT

ws["A10"] = "Exceções por regra"
ws["A10"].font = Font(name=FONTE, bold=True, size=12, color=AZUL)
ws.append([])
linha = 11
ws.cell(row=linha, column=1, value="Regra")
ws.cell(row=linha, column=2, value="Nome")
ws.cell(row=linha, column=3, value="Severidade")
ws.cell(row=linha, column=4, value="Quantidade")
ws.cell(row=linha, column=5, value="Valor envolvido")
estilo_cabecalho(ws, linha, 5)
ws.freeze_panes = None

for i, r in regras.iterrows():
    lin = linha + 1 + i
    ws.cell(row=lin, column=1, value=r["cod_regra"])
    ws.cell(row=lin, column=2, value=r["nome"])
    ws.cell(row=lin, column=3, value=r["severidade"])
    ws.cell(row=lin, column=4,
            value=f'=COUNTIF(Excecoes!$A$2:$A$20000,$A{lin})').number_format = "#,##0"
    ws.cell(row=lin, column=5,
            value=f'=SUMPRODUCT((Excecoes!$A$2:$A$20000=$A{lin})*ABS(Excecoes!$J$2:$J$20000))'
            ).number_format = MOEDA
    for c in range(1, 6):
        ws.cell(row=lin, column=c).font = Font(name=FONTE, size=10)
        ws.cell(row=lin, column=c).border = borda

total = linha + 1 + len(regras)
ws.cell(row=total, column=2, value="Total").font = Font(name=FONTE, bold=True, size=10)
ws.cell(row=total, column=4,
        value=f"=SUM(D{linha + 1}:D{total - 1})").number_format = "#,##0"
ws.cell(row=total, column=5,
        value=f"=SUM(E{linha + 1}:E{total - 1})").number_format = MOEDA
for c in (4, 5):
    ws.cell(row=total, column=c).font = Font(name=FONTE, bold=True, size=10)

ws.cell(row=total + 2, column=1,
        value="Execute Alt+F8 > ExecutarConciliacao para preencher. "
              "Os números acima recalculam sozinhos.").font = Font(
    name=FONTE, size=9, italic=True, color="808080")
largura(ws, {"A": 14, "B": 32, "C": 13, "D": 14, "E": 20, "F": 14, "G": 22})

# =========================================================================
# 10. Log
# =========================================================================
ws = wb.create_sheet("Log")
ws.append(["data_hora", "responsavel", "competencia", "linhas_conciliacao",
           "excecoes_geradas", "criticas", "segundos", "versao"])
estilo_cabecalho(ws, 1, 8)
largura(ws, {"A": 20, "B": 24, "C": 14, "D": 20, "E": 18, "F": 12, "G": 11, "H": 12})

# =========================================================================
# 11. Gabarito_Demo
# =========================================================================
gab = pd.read_parquet(B / "gabarito.parquet")
ws = wb.create_sheet("Gabarito_Demo")
escrever_df(ws, gab, nome_tabela="tblGabarito")
largura(ws, {"A": 12, "B": 13, "C": 62})
ws.cell(row=len(gab) + 3, column=1,
        value="Anomalias plantadas de propósito nos dados de demonstração. "
              "Rode a conciliação e confira se todas aparecem em Excecoes.").font = Font(
    name=FONTE, size=9, italic=True, color="808080")

ordem = ["Leia-me", "Resumo", "Config", "Regras", "Excecoes", "Conciliacao",
         "Folha_Atual", "Folha_Anterior", "Cadastro", "Log", "Gabarito_Demo"]
wb._sheets = [wb[n] for n in ordem]
SAIDA.parent.mkdir(parents=True, exist_ok=True)
wb.save(SAIDA)
print("Salvo:", SAIDA, "| abas:", wb.sheetnames)
