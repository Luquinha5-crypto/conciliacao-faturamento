"""Roda a conciliação pela linha de comando.

Exemplo (dados fictícios):
    python executar.py --ano 2026 --mes 9 --quinzena 2 --filial LONDRINA
"""
import argparse
from pathlib import Path

from src.conciliacao import (Colunas, ler_placas, ler_planilhas, filtrar_b, classificar,
                             separar_periodo, periodo_quinzena, rotulo_quinzena,
                             montar_resumo, montar_alertas, validar_pendencias, exportar_excel)

p = argparse.ArgumentParser()
p.add_argument("--a", default="exemplos/devolutiva_exemplo.xlsx")
p.add_argument("--b", default="exemplos/faturamento_exemplo.csv")
p.add_argument("--placas", default="exemplos/placas_exemplo.txt")
p.add_argument("--sep", default=";")
p.add_argument("--encoding", default="utf-8", help="CSV real costuma ser latin-1")
p.add_argument("--cliente", default=r"RA[IÍ]ZEN")
p.add_argument("--filial", default="EXEMPLO")
p.add_argument("--ano", type=int, required=True)
p.add_argument("--mes", type=int, required=True)
p.add_argument("--quinzena", type=int, required=True, choices=[1, 2])
p.add_argument("--saida", default="saida/conciliacao.xlsx")
args = p.parse_args()

cols = Colunas()
inicio, fim = periodo_quinzena(args.ano, args.mes, args.quinzena)
rotulo = rotulo_quinzena(args.mes, args.quinzena)

df_a, df_b = ler_planilhas(args.a, args.b, cols, args.sep, args.encoding)
placas = ler_placas(args.placas)
b, cols_placa, diag = filtrar_b(df_b, placas, cols, args.cliente)
print(f"Linhas cliente: {diag['linhas_cliente']} | linhas placas: {diag['linhas_placas']} | após filtros: {len(b)}")
if diag["placas_fora_da_b"]:
    print("ATENÇÃO placas sem nenhuma viagem na B:", diag["placas_fora_da_b"])
if diag["placas_sem_cliente"]:
    print("ATENÇÃO placas que nunca aparecem com o cliente:", diag["placas_sem_cliente"])

b = classificar(b, df_a, cols, cols_placa)
grupos, divergentes = separar_periodo(b, inicio, fim, cols)
if divergentes:
    print(f"ATENÇÃO {divergentes} célula(s) com datas diferentes; usada a primeira.")

resumo = montar_resumo({args.filial: grupos})
print(resumo.to_string(index=False))
pend = grupos["no_periodo"][grupos["no_periodo"]["status"].isin(["nao_paga", "parcial"])]
val = validar_pendencias(df_a, cols, pend)
print("Validação independente -> inesperadas (olhar à mão):", val["inesperadas"], "| explicadas:", val["explicadas"])
alertas = montar_alertas(grupos["no_periodo"], cols)

Path(args.saida).parent.mkdir(parents=True, exist_ok=True)
exportar_excel(args.saida, rotulo, cols, {args.filial: grupos}, resumo, alertas)
print("Gerado:", args.saida)
