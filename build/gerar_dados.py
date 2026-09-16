"""Gera dois meses de 'arquivo do sistema de folha' com anomalias injetadas.

Os dados saem da mesma empresa sintética usada nos projetos de Power BI, o que
mantém o portfólio coerente. As anomalias são deliberadas e documentadas na aba
Gabarito_Demo: sem gabarito não dá para provar que o motor de regras funciona.
"""
from pathlib import Path

import numpy as np
import pandas as pd

BASE = Path("/home/claude/hr-synthetic-data-br/data/synthetic")
rng = np.random.default_rng(2024)

folha = pd.read_parquet(BASE / "fato_folha_mensal.parquet")
snaps = pd.read_parquet(BASE / "fato_headcount_mensal.parquet")
colab = pd.read_parquet(BASE / "dim_colaborador.parquet")
areas = pd.read_parquet(BASE / "dim_area.parquet")

MES_ANT, MES_ATU = 202511, 202512
SALARIO_MINIMO = 1518.00

RUBRICAS = [
    ("001", "Salário base", "Provento"),
    ("002", "Horas extras 50%", "Provento"),
    ("003", "DSR sobre horas extras", "Provento"),
    ("101", "INSS", "Desconto"),
    ("102", "IRRF", "Desconto"),
    ("104", "Plano de saúde - coparticipação", "Desconto"),
]


def inss(base):
    return np.clip(base * 0.09 + (base > 7500) * 180, 0, 951.63)


def irrf(base):
    return np.where(
        base > 4664, (base - 4664) * 0.275 * 0.8,
        np.where(base > 2826, (base - 2826) * 0.15, 0.0),
    )


def montar(id_mes):
    f = folha[folha["id_mes"] == id_mes][
        ["matricula", "salario", "valor_horas_extras", "id_area"]
    ].copy()
    linhas = []
    for _, r in f.iterrows():
        sal, he = float(r["salario"]), float(r["valor_horas_extras"])
        dsr = he * 0.1818
        bruto = sal + he + dsr
        vals = {
            "001": sal, "002": he, "003": dsr,
            "101": -float(inss(bruto)),
            "102": -float(irrf(bruto - float(inss(bruto)))),
            "104": -round(float(rng.uniform(60, 260)), 2),
        }
        for cod, desc, tipo in RUBRICAS:
            v = round(vals[cod], 2)
            if cod in ("002", "003") and abs(v) < 1:
                continue
            linhas.append({
                "id_mes": id_mes, "matricula": int(r["matricula"]),
                "cod_rubrica": cod, "descricao_rubrica": desc, "tipo": tipo,
                "valor": v, "id_area": int(r["id_area"]),
            })
    return pd.DataFrame(linhas)


ant, atu = montar(MES_ANT), montar(MES_ATU)
gabarito = []


def marcar(regra, matriculas, obs):
    for m in matriculas:
        gabarito.append({
            "cod_regra": regra, "matricula": int(m), "observacao": obs
        })


ativos_atu, ativos_ant = set(atu["matricula"]), set(ant["matricula"])
comuns = sorted(ativos_atu & ativos_ant)
usadas: set[int] = set()


def sortear(n):
    disponiveis = [m for m in comuns if m not in usadas]
    escolhidos = [int(x) for x in rng.choice(disponiveis, n, replace=False)]
    usadas.update(escolhidos)
    return escolhidos


# --- R01: admitido no mês sem nenhum provento ------------------------------
novos = sorted(ativos_atu - ativos_ant)[:5]
atu = atu[~((atu["matricula"].isin(novos)) & (atu["tipo"] == "Provento"))]
marcar("R01", novos, "Admitido no mês, sem lançamento de provento")

# --- R02: desligado antes do mês com provento ------------------------------
deslig = colab.dropna(subset=["data_desligamento"])
deslig = deslig[deslig["data_desligamento"] < "2025-12-01"]["matricula"].tolist()
alvo = [m for m in deslig if m not in ativos_atu][:4]
atu = pd.concat([atu, pd.DataFrame([{
    "id_mes": MES_ATU, "matricula": int(m), "cod_rubrica": "001",
    "descricao_rubrica": "Salário base", "tipo": "Provento",
    "valor": round(float(rng.uniform(3000, 9000)), 2), "id_area": 1,
} for m in alvo])], ignore_index=True)
marcar("R02", alvo, "Desligado em competência anterior com provento lançado")

# --- R03: variação de líquido acima do limite ------------------------------
alvo = sortear(12)
mask = (atu["matricula"].isin(alvo)) & (atu["cod_rubrica"] == "001")
atu.loc[mask, "valor"] = (
    atu.loc[mask, "valor"] * rng.uniform(1.22, 1.45, int(mask.sum()))
).round(2)
marcar("R03", alvo, "Variação do líquido acima do limite configurado")

# --- R04: rubrica nova de valor relevante ----------------------------------
alvo = sortear(8)
atu = pd.concat([atu, pd.DataFrame([{
    "id_mes": MES_ATU, "matricula": int(m), "cod_rubrica": "205",
    "descricao_rubrica": "Prêmio por desempenho", "tipo": "Provento",
    "valor": round(float(rng.uniform(1800, 6500)), 2), "id_area": 1,
} for m in alvo])], ignore_index=True)
marcar("R04", alvo, "Rubrica inexistente na competência anterior")

# --- R05: rubrica recorrente que desapareceu -------------------------------
alvo = sortear(6)
atu = atu[~((atu["matricula"].isin(alvo)) & (atu["cod_rubrica"] == "104"))]
marcar("R05", alvo, "Rubrica recorrente ausente na competência atual")

# --- R06: lançamento duplicado ---------------------------------------------
alvo = sortear(3)
dup = atu[(atu["matricula"].isin(alvo)) & (atu["cod_rubrica"] == "001")].copy()
atu = pd.concat([atu, dup], ignore_index=True)
marcar("R06", alvo, "Lançamento duplicado da mesma rubrica")

# --- R07: salário base abaixo do mínimo ------------------------------------
alvo = sortear(2)
mask = (atu["matricula"].isin(alvo)) & (atu["cod_rubrica"] == "001")
atu.loc[mask, "valor"] = round(SALARIO_MINIMO - 210.0, 2)
marcar("R07", alvo, "Salário base abaixo do mínimo vigente")

# --- R08: horas extras acima do limite -------------------------------------
alvo = sortear(5)
for m in alvo:
    base = atu[(atu.matricula == m) & (atu.cod_rubrica == "001")]["valor"]
    if base.empty:
        continue
    v = round(float(base.iloc[0]) * 0.62, 2)
    idx = atu[(atu.matricula == m) & (atu.cod_rubrica == "002")].index
    if len(idx):
        atu.loc[idx, "valor"] = v
    else:
        atu = pd.concat([atu, pd.DataFrame([{
            "id_mes": MES_ATU, "matricula": int(m), "cod_rubrica": "002",
            "descricao_rubrica": "Horas extras 50%", "tipo": "Provento",
            "valor": v, "id_area": 1,
        }])], ignore_index=True)
marcar("R08", alvo, "Horas extras acima do percentual limite do salário")

# --- R09: sumiu da folha sem desligamento registrado -----------------------
alvo = sortear(7)
atu = atu[~atu["matricula"].isin(alvo)]
marcar("R09", alvo, "Ausente na folha atual sem desligamento no cadastro")

cad = colab[["matricula", "nome", "data_admissao", "data_desligamento", "status"]].copy()
cad = cad[cad["matricula"].isin(sorted(set(atu.matricula) | set(ant.matricula)))]
cad = cad.merge(
    snaps[snaps.id_mes == MES_ATU][["matricula", "id_area", "grade"]],
    on="matricula", how="left",
).merge(areas[["id_area", "diretoria", "gerencia", "centro_custo"]],
        on="id_area", how="left")

out = Path("/home/claude/conciliador-folha/build")
ant.sort_values(["matricula", "cod_rubrica"]).to_parquet(out / "folha_anterior.parquet", index=False)
atu.sort_values(["matricula", "cod_rubrica"]).to_parquet(out / "folha_atual.parquet", index=False)
cad.to_parquet(out / "cadastro.parquet", index=False)
pd.DataFrame(gabarito).to_parquet(out / "gabarito.parquet", index=False)
print("anterior:", len(ant), "| atual:", len(atu), "| cadastro:", len(cad),
      "| anomalias:", len(gabarito))
