"""Gera dados FICTÍCIOS de exemplo (nada aqui vem de dados reais)."""
from pathlib import Path
import pandas as pd

pasta = Path(__file__).parent
DEST = "RAIZEN FICTICIA S.A."

def cte(n, nota, data, placa, carreta="", dest=DEST, frete="1.000,00", icms="120,00", total="1.120,00"):
    chave = f"35260900000000000100570010000{n:05d}1000000000"[:44].ljust(44, "0")
    return {
        "N.º do CTRC": str(n), "Chave CT-e": chave,
        "Emissão NF": data, "N.º - Nota Fiscal": nota,
        "Chave NF-e": ";".join(f"3526090000000000019955001{int(x):09d}1000000000"[:44] for x in nota.replace(";", " ").split()) if nota else "",
        "Razão Social - Remetente": "BASE FICTICIA LTDA", "Razão Social - Destinatário": dest,
        "Produto - Nota Fiscal": "ETANOL", "Quantidade - Nota Fiscal": "30000",
        "Tarifa Final - Frete Empresa": "33,33", "Tarifa Real - Frete Empresa": "33,33",
        "Valor Frete": frete, "ICMS - Composição do Frete": icms, "Total da Prestação - Composição do Frete": total,
        "Placa Veículo  - Nota Fiscal": placa, "Placa Central/Carreta - Nota Fiscal": carreta,
        "Placa Central/Carreta - Nota Fiscal.1": "",
    }

linhas = [
    cte(5001, "1000001", "17/09/2026", "ABC1D23"),                       # paga
    cte(5002, "1000002", "18/09/2026", "XYZ9K88"),                       # paga (A tem 1000002-1)
    cte(5003, "1000003;1000004", "19/09/2026;19/09/2026", "DEF-4567"),   # parcial
    cte(5004, "1000005", "20/09/2026", "ABC1D23"),                       # nao_paga
    cte(5005, "1000006", "21/09/2026", "GHI2J34", "ZZZ0Z00"),            # paga pela carreta
    cte(5006, "1000007", "22/09/2026", "XYZ9K88"),                       # nota na A com outra placa
    cte(5007, "001000008", "23/09/2026", "DEF4567"),                     # paga (zero à esquerda)
    cte(5008, "1000009", "10/09/2026", "ABC1D23"),                       # fora do período
    cte(5009, "1000010", "", "ABC1D23"),                                 # sem data
    cte(5010, "", "24/09/2026", "ABC1D23"),                              # sem nota
    cte(5011, "1000012", "25/09/2026", "ABC1D23", dest="OUTRO CLIENTE LTDA"),  # outro cliente
    cte(5012, "1000013", "25/09/2026", "QQQ1Q11"),                       # placa fora da lista
]
pd.DataFrame(linhas).to_csv(pasta / "faturamento_exemplo.csv", sep=";", index=False, encoding="utf-8")

a = pd.DataFrame({
    "Nota Fiscal": ["1000001", "1000002-1", "1000003", "1000006", "1000007", "1000008", "1000099"],
    "Veículo": ["ABC1D23", "XYZ9K88", "DEF4567", "ZZZ0Z00", "ABC1D23", "DEF4567", "ABC1D23"],
    "Custo Viagem": [1000, 980, 500, 900, 700, 800, 0],
})
a.to_excel(pasta / "devolutiva_exemplo.xlsx", index=False)
(pasta / "placas_exemplo.txt").write_text("ABC1D23\nXYZ9K88\nDEF-4567\nGHI2J34\n", encoding="utf-8")
print("ok")
