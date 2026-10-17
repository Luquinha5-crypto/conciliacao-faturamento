"""Conciliação de faturamento.

Pergunta de negócio: dos CT-e faturados na quinzena (planilha B), para as placas
de uma filial e um cliente específico, quais NÃO aparecem na devolutiva do
cliente (planilha A)?

Regras principais
-----------------
* A unidade de análise é o **CT-e** (uma linha da planilha B).
* Como a planilha A não tem o CT-e, a ponte é feita por **nota fiscal + placa**:
  uma nota da B é considerada paga quando a A traz a mesma nota com uma das
  placas do CT-e (cavalo ou carreta).
* Um CT-e é `paga` quando todas as suas notas estão na A, `nao_paga` quando
  nenhuma está e `parcial` quando só algumas estão. `parcial` conta como pendente.
* O período vem da emissão da NF (`Emissão NF`). Linhas sem data NÃO somem:
  ficam num grupo próprio.
"""

from __future__ import annotations

import re
from dataclasses import dataclass
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

MESES = [
    "JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO",
    "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO",
]
STATUS = ["paga", "nao_paga", "parcial", "sem_nota"]
PENDENTES = ["nao_paga", "parcial"]


@dataclass
class Colunas:
    """Nomes das colunas das planilhas. Altere aqui se o sistema mudar um nome."""

    # Planilha A (devolutiva do cliente)
    nota_a: str = "Nota Fiscal"
    placa_a: str = "Veículo"
    # Planilha B (faturamento)
    nota_b: str = "N.º - Nota Fiscal"
    ctrc: str = "N.º do CTRC"
    chave_cte: str = "Chave CT-e"
    chave_nfe: str = "Chave NF-e"
    data_b: str = "Emissão NF"
    destinatario: str = "Razão Social - Destinatário"
    remetente: str = "Razão Social - Remetente"
    produto: str = "Produto - Nota Fiscal"
    volume: str = "Quantidade - Nota Fiscal"
    tarifa: str = "Tarifa Final - Frete Empresa"
    tarifa_real: str = "Tarifa Real - Frete Empresa"
    frete: str = "Valor Frete"
    icms: str = "ICMS - Composição do Frete"
    total: str = "Total da Prestação - Composição do Frete"
    prefixo_placa: str = "Placa"  # toda coluna da B que começa assim entra no filtro


# --------------------------------------------------------------------------
# Período
# --------------------------------------------------------------------------
def periodo_quinzena(ano: int, mes: int, quinzena: int) -> tuple[pd.Timestamp, pd.Timestamp]:
    """1ª quinzena: dias 1 a 15. 2ª quinzena: dia 16 até o último dia do mês (30 ou 31)."""
    if quinzena == 1:
        return pd.Timestamp(ano, mes, 1), pd.Timestamp(ano, mes, 15)
    if quinzena == 2:
        inicio = pd.Timestamp(ano, mes, 16)
        return inicio, inicio + pd.offsets.MonthEnd(0)
    raise ValueError("quinzena deve ser 1 ou 2")


def rotulo_quinzena(mes: int, quinzena: int) -> str:
    """Ex.: '2Q SETEMBRO' (nome da aba e do período)."""
    return f"{quinzena}Q {MESES[mes - 1]}"


# --------------------------------------------------------------------------
# Limpeza
# --------------------------------------------------------------------------
def normaliza_placa(serie: pd.Series) -> pd.Series:
    """Maiúsculas, sem espaços e sem hífens. Vazio vira NA."""
    s = serie.astype("string").str.upper().str.replace(r"[\s\-]", "", regex=True)
    return s.where(s != "", pd.NA)


def ler_placas(arquivo: str | Path) -> set[str]:
    """Lê placas separadas por linha, espaço, vírgula ou ponto e vírgula.

    Usa utf-8-sig: o Bloco de Notas às vezes grava uma marca invisível (BOM) no
    início do arquivo, que faria a primeira placa nunca casar.
    """
    texto = Path(arquivo).read_text(encoding="utf-8-sig")
    brutas = [p for p in re.split(r"[\s,;]+", texto) if p]
    return set(normaliza_placa(pd.Series(brutas, dtype="string")).dropna())


def explode_notas(df: pd.DataFrame, coluna: str) -> pd.DataFrame:
    """Uma nota por linha (o índice original é mantido).

    Separa células com várias notas e remove espaços, '.0' final, zeros à
    esquerda e o sufixo após hífen (1234-1 -> 1234). Vazios são descartados:
    um vazio nunca pode virar "nota paga".
    """
    t = df[[coluna]].dropna().copy()
    t["nf"] = t[coluna].astype(str).str.split(r"[,;\s]+", regex=True)
    t = t.explode("nf")
    t["nf"] = (
        t["nf"].str.strip()
        .str.replace(r"\.0$", "", regex=True)
        .str.lstrip("0")
        .str.split("-").str[0]
    )
    return t[t["nf"].notna() & (t["nf"] != "")]


def numero(serie: pd.Series) -> pd.Series:
    """Texto -> número, aceitando '1.234,56' (Brasil) e '1234.56'. O que não for número vira NaN."""
    s = serie.astype("string").str.strip()
    br = s.str.contains(",", na=False)
    s = s.where(~br, s.str.replace(".", "", regex=False).str.replace(",", ".", regex=False))
    return pd.to_numeric(s, errors="coerce")


def numero_soma(serie: pd.Series) -> pd.Series:
    """Como `numero`, mas soma as partes de células com vários valores ('10;20' -> 30)."""
    partes = serie.astype("string").str.split(";").explode()
    return numero(partes).groupby(level=0).sum(min_count=1)


def unico(serie: pd.Series) -> pd.Series:
    """Remove repetições dentro da célula ('A;A;B' -> 'A; B'). Vazio vira ''."""
    def limpa(x):
        if pd.isna(x):
            return ""
        partes = [p.strip() for p in str(x).split(";") if p.strip()]
        return "; ".join(dict.fromkeys(partes))
    return serie.map(limpa)


def checar_colunas(df: pd.DataFrame, nomes: list[str], origem: str) -> None:
    """Erro claro quando uma coluna esperada não existe (nome mudou, encoding errado...)."""
    faltam = [n for n in nomes if n not in df.columns]
    if faltam:
        raise ValueError(
            f"Planilha {origem}: colunas não encontradas {faltam}. "
            f"Confira nomes, separador e encoding. Colunas lidas: {list(df.columns)[:8]}..."
        )


# --------------------------------------------------------------------------
# Leitura e filtros
# --------------------------------------------------------------------------
def ler_planilhas(arq_a, arq_b, cols: Colunas, sep: str = ";", encoding: str = "utf-8"):
    """Lê A (Excel) e B (CSV). Tudo como texto: o Pandas não pode transformar
    1135370 em 1135370.0 nem apagar zeros à esquerda."""
    df_a = pd.read_excel(arq_a, dtype=str)
    df_b = pd.read_csv(arq_b, sep=sep, encoding=encoding, dtype=str)
    checar_colunas(df_a, [cols.nota_a, cols.placa_a], "A")
    checar_colunas(
        df_b,
        [cols.nota_b, cols.ctrc, cols.data_b, cols.destinatario, cols.frete, cols.icms, cols.total],
        "B",
    )
    return df_a, df_b


def filtrar_b(df_b: pd.DataFrame, placas: set[str], cols: Colunas, cliente_regex: str):
    """Mantém as linhas da B com destino ao cliente E com alguma placa da lista
    (em qualquer coluna de placa: cavalo ou carreta).

    Devolve (b_filtrada, colunas_de_placa, diagnostico). O diagnóstico lista
    placas da lista que não aparecem na B, e as que aparecem mas nunca com o cliente.
    """
    cols_placa = [c for c in df_b.columns if c.startswith(cols.prefixo_placa)]
    placas_norm = pd.concat([normaliza_placa(df_b[c]) for c in cols_placa], axis=1)

    mask_placa = placas_norm.isin(list(placas)).any(axis=1)
    mask_cliente = df_b[cols.destinatario].str.contains(cliente_regex, case=False, na=False)

    def conjunto(tabela):
        return set(pd.Series(tabela.to_numpy().ravel()).dropna())

    diag = {
        "linhas_cliente": int(mask_cliente.sum()),
        "linhas_placas": int(mask_placa.sum()),
        "placas_fora_da_b": sorted(placas - conjunto(placas_norm)),
        "placas_sem_cliente": sorted(
            (placas & conjunto(placas_norm)) - conjunto(placas_norm[mask_cliente])
        ),
    }
    return df_b[mask_placa & mask_cliente].copy(), cols_placa, diag


# --------------------------------------------------------------------------
# Classificação (nota + placa)
# --------------------------------------------------------------------------
def pares_pagos(df_a: pd.DataFrame, cols: Colunas):
    """Pares (nota, placa) da devolutiva e o conjunto de notas, sem considerar placa."""
    t = explode_notas(df_a, cols.nota_a)
    t["placa"] = normaliza_placa(df_a.loc[t.index, cols.placa_a]).to_numpy()
    return set(zip(t["nf"], t["placa"])), set(t["nf"])


def classificar(b: pd.DataFrame, df_a: pd.DataFrame, cols: Colunas, cols_placa: list[str]) -> pd.DataFrame:
    """Acrescenta à B: status, notas_faltantes e notas_placa_diferente.

    Uma nota é paga se a A tem (mesma nota, uma das placas do CT-e).
    `notas_placa_diferente` lista notas que existem na A, mas com outra placa:
    não contam como pagas, mas merecem olhar humano.
    """
    b = b.drop(columns=["status", "notas_faltantes", "notas_placa_diferente"], errors="ignore")
    if b.empty:
        return b.assign(status="", notas_faltantes="", notas_placa_diferente="")

    pares, notas_a = pares_pagos(df_a, cols)
    b_nf = explode_notas(b, cols.nota_b)

    placas_norm = pd.concat([normaliza_placa(b[c]) for c in cols_placa], axis=1)
    placas_linha = placas_norm.apply(lambda r: [p for p in r if pd.notna(p)], axis=1)

    b_nf["paga"] = [
        any((nf, p) in pares for p in placas_linha[i]) for i, nf in zip(b_nf.index, b_nf["nf"])
    ]
    b_nf["nota_na_a"] = b_nf["nf"].isin(notas_a)

    agr = b_nf.groupby(level=0).agg(total=("paga", "size"), pagas=("paga", "sum"))
    agr["status"] = "parcial"
    agr.loc[agr["pagas"] == 0, "status"] = "nao_paga"
    agr.loc[agr["pagas"] == agr["total"], "status"] = "paga"

    sem_par = b_nf[~b_nf["paga"]]
    agr["notas_faltantes"] = sem_par.groupby(level=0)["nf"].agg("; ".join)
    agr["notas_placa_diferente"] = sem_par[sem_par["nota_na_a"]].groupby(level=0)["nf"].agg("; ".join)

    saida = b.join(agr[["status", "notas_faltantes", "notas_placa_diferente"]])
    saida["status"] = saida["status"].fillna("sem_nota")  # CT-e sem nota informada
    for c in ["notas_faltantes", "notas_placa_diferente"]:
        saida[c] = saida[c].fillna("")
    return saida


def separar_periodo(b: pd.DataFrame, inicio: pd.Timestamp, fim: pd.Timestamp, cols: Colunas):
    """Divide a B em três grupos que NÃO se sobrepõem e cobrem tudo:
    no_periodo, fora_periodo e sem_data. Devolve (grupos, n_celulas_com_datas_diferentes)."""
    datas = b[cols.data_b].astype(str).str.split(r"[;,]")
    divergentes = int(datas.dropna().apply(lambda x: len({i.strip() for i in x}) > 1).sum())

    b = b.copy()
    b["data"] = pd.to_datetime(datas.str[0].str.strip(), format="%d/%m/%Y", errors="coerce")

    dentro = (b["data"] >= inicio) & (b["data"] <= fim)
    sem = b["data"].isna()
    grupos = {"no_periodo": b[dentro], "fora_periodo": b[~dentro & ~sem], "sem_data": b[sem]}
    assert sum(len(g) for g in grupos.values()) == len(b), "grupos não fecham com o total"
    return grupos, divergentes


# --------------------------------------------------------------------------
# Resumo, alertas e validação
# --------------------------------------------------------------------------
def montar_resumo(resultados: dict) -> pd.DataFrame:
    """Contagem por filial, grupo e status. A coluna total permite conferir se nada sumiu."""
    linhas = []
    for filial, grupos in resultados.items():
        for grupo, tab in grupos.items():
            cont = tab["status"].value_counts()
            linha = {"filial": filial, "grupo": grupo}
            linha.update({s: int(cont.get(s, 0)) for s in STATUS})
            linha["total"] = len(tab)
            linhas.append(linha)
    return pd.DataFrame(linhas)


def montar_alertas(tab: pd.DataFrame, cols: Colunas, tol_soma: float = 0.05, tol_tarifa: float = 0.01) -> pd.DataFrame:
    """Problemas que pedem olhar humano, sem alterar o resultado."""
    achados = []

    def add(mask, tipo, detalhe=""):
        for idx in tab.index[mask.fillna(False)]:
            achados.append({"CT-e": tab.at[idx, cols.ctrc], "tipo": tipo, "detalhe": detalhe})

    frete, icms, total = numero(tab[cols.frete]), numero(tab[cols.icms]), numero(tab[cols.total])
    add(tab[cols.ctrc].duplicated(keep=False), "CT-e repetido na planilha B")
    add(tab["status"] == "sem_nota", "CT-e sem nota fiscal informada")
    add(frete.isna() | icms.isna() | total.isna(), "valor vazio ou não numérico (frete, ICMS ou total)")
    add(((frete + icms) - total).abs() > tol_soma, "frete + ICMS difere do total da prestação")
    tf, tr = numero(tab[cols.tarifa]), numero(tab[cols.tarifa_real])
    add((tf - tr).abs() > tol_tarifa * tf.abs(), "tarifa final difere da tarifa real")
    return pd.DataFrame(achados, columns=["CT-e", "tipo", "detalhe"])


def validar_pendencias(df_a: pd.DataFrame, cols: Colunas, pendentes: pd.DataFrame) -> dict:
    """Segunda opinião, independente da limpeza: procura cada nota pendente como
    TRECHO de texto na coluna de notas da A.

    * `explicadas`: achadas na A, mas já sinalizadas como "placa diferente".
    * `inesperadas`: achadas na A e NÃO explicadas. Olhar à mão: pode ser um
      formato de nota novo (ou só um trecho de um número maior, em notas curtas).
    """
    texto = df_a[cols.nota_a].astype(str)
    faltantes = set(explode_notas(pendentes.assign(f=pendentes["notas_faltantes"]), "f")["nf"])
    explicadas_base = set(
        explode_notas(pendentes.assign(f=pendentes["notas_placa_diferente"]), "f")["nf"]
    )
    achadas = {n for n in faltantes if texto.str.contains(n, regex=False, na=False).any()}
    return {"explicadas": sorted(achadas & explicadas_base), "inesperadas": sorted(achadas - explicadas_base)}


# --------------------------------------------------------------------------
# Exportação (modelo do cliente + abas de apoio)
# --------------------------------------------------------------------------
COR_CABECALHO = "781E77"
_FINA = Side(style="thin")
_BORDA = Border(left=_FINA, right=_FINA, top=_FINA, bottom=_FINA)
CABECALHO_MODELO = [
    "QZNA", "Data emissão NFe", "Nº NFe", "Chave da NFe", "Nº CTe/NFSe", "Chave CTe",
    "Produto", "Volume", "ORIGEM (BASE/USINA)", "DESTINO (BASE/USINA/CLIENTE)",
    "TARIFA M³/KM/KG", "Valor Frete", "Valor Imposto", "Valor de Frete + Impostos",
]
_LARGURAS = [8, 16, 20, 52, 16, 52, 28, 14, 40, 40, 16, 16, 16, 24]
_FORMATOS = {2: "dd/mm/yyyy", 4: "@", 6: "@", 8: "#,##0", 11: "#,##0.00", 12: "#,##0.00", 13: "#,##0.00", 14: "#,##0.00"}


def tabela_modelo(df: pd.DataFrame, cols: Colunas, qzna: str) -> pd.DataFrame:
    """Uma linha por CT-e, nas 14 colunas do modelo do cliente."""
    t = pd.DataFrame({
        "QZNA": qzna,
        "Data emissão NFe": df["data"],
        "Nº NFe": unico(df[cols.nota_b]),
        "Chave da NFe": unico(df[cols.chave_nfe]),
        "Nº CTe/NFSe": df[cols.ctrc],
        "Chave CTe": df[cols.chave_cte],
        "Produto": unico(df[cols.produto]),
        "Volume": numero_soma(df[cols.volume]),
        "ORIGEM (BASE/USINA)": unico(df[cols.remetente]),
        "DESTINO (BASE/USINA/CLIENTE)": unico(df[cols.destinatario]),
        "TARIFA M³/KM/KG": numero(df[cols.tarifa]),
        "Valor Frete": numero(df[cols.frete]),
        "Valor Imposto": numero(df[cols.icms]),
        "Valor de Frete + Impostos": numero(df[cols.total]),
    }, index=df.index)
    return t.sort_values(["Data emissão NFe", "Nº CTe/NFSe"], kind="stable")


def _valor(v):
    """Converte valores do Pandas/NumPy para algo que o openpyxl grava."""
    if v is pd.NA or (not isinstance(v, (list, tuple)) and pd.isna(v)):
        return None
    if isinstance(v, pd.Timestamp):
        return v.to_pydatetime()
    if hasattr(v, "item"):
        return v.item()
    return v


def _celula(ws, linha, col, valor, fonte, alinhamento, formato=None):
    c = ws.cell(row=linha, column=col, value=valor)
    c.font, c.alignment, c.border = fonte, alinhamento, _BORDA
    if formato:
        c.number_format = formato
    return c


def _escrever_bloco_modelo(ws, linha: int, filial: str, tabela: pd.DataFrame) -> int:
    """Título, cabeçalho e linhas de uma filial. Devolve a próxima linha livre."""
    ws.cell(row=linha, column=2, value=f"FILIAL: {filial}").font = Font(name="Arial", size=16, bold=True)
    ws.merge_cells(start_row=linha, start_column=2, end_row=linha, end_column=3)
    ws.row_dimensions[linha].height = 23.25

    centro = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cab = Font(name="Arial", size=12, bold=True, color="FFFFFF")
    for j, nome in enumerate(CABECALHO_MODELO, start=1):
        c = _celula(ws, linha + 1, j, nome, cab, centro)
        c.fill = PatternFill("solid", fgColor=COR_CABECALHO)

    corpo, negrito = Font(name="Arial", size=12), Font(name="Arial", size=12, bold=True)
    if tabela.empty:
        ws.cell(row=linha + 2, column=2, value="Nenhuma pendência no período").font = corpo
        return linha + 4

    for i, (_, reg) in enumerate(tabela.iterrows(), start=linha + 2):
        for j, nome in enumerate(CABECALHO_MODELO, start=1):
            v = _valor(reg[nome])
            if nome == "Nº CTe/NFSe" and isinstance(v, str) and v.isdigit():
                v = int(v)
            _celula(ws, i, j, v, negrito if j == 1 else corpo, centro, _FORMATOS.get(j))
    return linha + 2 + len(tabela) + 1


def _escrever_tabela(ws, df: pd.DataFrame) -> None:
    """Tabela simples com o mesmo cabeçalho roxo, para as abas de apoio."""
    centro = Alignment(horizontal="center", vertical="center", wrap_text=True)
    corpo = Font(name="Arial", size=11)
    for j, nome in enumerate(df.columns, start=1):
        c = _celula(ws, 1, j, str(nome), Font(name="Arial", size=11, bold=True, color="FFFFFF"), centro)
        c.fill = PatternFill("solid", fgColor=COR_CABECALHO)
        maior = max([len(str(nome))] + [len(str(x)) for x in df[nome].head(200)])
        ws.column_dimensions[get_column_letter(j)].width = min(max(maior + 2, 10), 60)
    for i, (_, reg) in enumerate(df.iterrows(), start=2):
        for j, nome in enumerate(df.columns, start=1):
            v = _valor(reg[nome])
            _celula(ws, i, j, v, corpo, centro, "dd/mm/yyyy" if hasattr(v, "year") else None)
    ws.freeze_panes = "A2"


def exportar_excel(caminho, rotulo: str, cols: Colunas, resultados: dict, resumo: pd.DataFrame, alertas: pd.DataFrame) -> None:
    """Gera o Excel final.

    Abas: <rotulo> (pendências no modelo do cliente, um bloco por filial), resumo,
    pagas_periodo, detalhes_pendencias, e, só quando houver, sem_data e alertas.
    """
    qzna = rotulo.split()[0]
    wb = Workbook()
    ws = wb.active
    ws.title = rotulo
    for j, larg in enumerate(_LARGURAS, start=1):
        ws.column_dimensions[get_column_letter(j)].width = larg
    ws.freeze_panes = "B1"

    linha = 1
    pagas, detalhes, sem_data = [], [], []
    for filial, grupos in resultados.items():
        no = grupos["no_periodo"]
        pend = no[no["status"].isin(PENDENTES)]
        linha = _escrever_bloco_modelo(ws, linha, filial, tabela_modelo(pend, cols, qzna))

        pg = tabela_modelo(no[no["status"] == "paga"], cols, qzna)
        pagas.append(pg.assign(Filial=filial)[["Filial"] + CABECALHO_MODELO])
        det = pend[["status", "notas_faltantes", "notas_placa_diferente", cols.ctrc, cols.data_b, cols.nota_b]].copy()
        det.insert(0, "Filial", filial)
        detalhes.append(det)
        sd = grupos["sem_data"]
        if not sd.empty:
            sem_data.append(sd[[cols.ctrc, cols.data_b, cols.nota_b, "status"]].assign(Filial=filial))

    def nova_aba(nome, df):
        _escrever_tabela(wb.create_sheet(nome), df)

    nova_aba("resumo", resumo)
    nova_aba("pagas_periodo", pd.concat(pagas, ignore_index=True))
    nova_aba("detalhes_pendencias", pd.concat(detalhes, ignore_index=True))
    if sem_data:
        nova_aba("sem_data", pd.concat(sem_data, ignore_index=True))
    if not alertas.empty:
        nova_aba("alertas", alertas)
    wb.save(caminho)
