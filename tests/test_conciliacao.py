import pandas as pd
from src.conciliacao import (Colunas, explode_notas, normaliza_placa, numero, periodo_quinzena,
                             ler_planilhas, ler_placas, filtrar_b, classificar, separar_periodo)

def test_explode_notas_limpa():
    df = pd.DataFrame({"n": ["001000008", "1000003;1000004", "1000002-1", "1135370.0", None, ""]})
    nfs = explode_notas(df, "n")["nf"].tolist()
    assert nfs == ["1000008", "1000003", "1000004", "1000002", "1135370"]

def test_placa():
    assert normaliza_placa(pd.Series(["def-4567", " abc1d23 ", ""])).tolist() == ["DEF4567", "ABC1D23", pd.NA]

def test_numero_br():
    assert numero(pd.Series(["1.234,56", "1234.5", "x"])).tolist()[:2] == [1234.56, 1234.5]

def test_quinzenas():
    assert periodo_quinzena(2026, 9, 2) == (pd.Timestamp(2026, 9, 16), pd.Timestamp(2026, 9, 30))
    assert periodo_quinzena(2026, 10, 2)[1] == pd.Timestamp(2026, 10, 31)
    assert periodo_quinzena(2026, 9, 1)[1] == pd.Timestamp(2026, 9, 15)

def test_fluxo_exemplo():
    c = Colunas()
    a, b = ler_planilhas("exemplos/devolutiva_exemplo.xlsx", "exemplos/faturamento_exemplo.csv", c)
    bf, cp, _ = filtrar_b(b, ler_placas("exemplos/placas_exemplo.txt"), c, r"RA[IÍ]ZEN")
    g, _ = separar_periodo(classificar(bf, a, c, cp), *periodo_quinzena(2026, 9, 2), c)
    assert sum(len(x) for x in g.values()) == len(bf)
    assert g["no_periodo"]["status"].value_counts().to_dict() == {"paga": 4, "nao_paga": 2, "parcial": 1, "sem_nota": 1}
