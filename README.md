# Conciliação de faturamento (logística)

Projeto de análise de dados em Python/Pandas que cruza duas planilhas para descobrir **quais viagens faturadas não foram reconhecidas/pagas pelo cliente**.

> Todos os dados deste repositório são **fictícios**. Nomes, placas e notas foram inventados.

## Pergunta de negócio

Para uma lista de veículos e um cliente específico, quais linhas do faturamento da quinzena **não aparecem** na planilha de devolutiva do cliente?

| Planilha | Papel |
|---|---|
| **B – faturamento geral** | Fonte dos dados; recebe os filtros (placas, cliente, período) |
| **A – devolutiva do cliente** | Usada para verificar se a nota foi reconhecida |

## Como funciona

1. **Filtra a B** por placas da lista (cavalo ou carreta) e pelo cliente (busca por trecho no destinatário).
2. **Limpa a chave** (nota fiscal) nas duas planilhas.
3. **Verifica** se cada nota da B existe na A, com `isin` (e não `merge`, para não multiplicar linhas).
4. **Classifica** cada linha: `paga`, `nao_paga` ou `parcial` (linha com várias notas, só algumas na A).
5. **Filtra o período** da quinzena e exporta um Excel com as abas `nao_pagas_periodo` e `pagas_periodo`.
6. **Valida** o resultado por um segundo caminho, independente da limpeza (busca por trecho).

## Problemas de qualidade de dados tratados

- Coluna de nota com **tipos misturados** (número e texto).
- **Várias notas na mesma célula** (`1000003;1000004`).
- **Zeros à esquerda** (`001000003`).
- **Sufixo após hífen** (`1000002-1`, `1000011-00`).
- Placas com **hífen, minúsculas ou espaços**.
- **Várias datas na mesma célula**, com teste de que são idênticas antes de usar a primeira.
- Coluna duplicada no CSV (o Pandas renomeia a segunda com `.1`).

## Como rodar

```bash
python -m venv .venv
.venv\Scripts\activate        # Windows  (Mac/Linux: source .venv/bin/activate)
python -m pip install -r requirements.txt
python -m notebook
Testado com **Python 3.13**. Versões muito recentes (como 3.14) podem falhar ao instalar o Jupyter no Windows por falta de pacotes pré-compilados.
py -3.13 -m venv venv
venv\Scripts\activate
python -m pip install --upgrade pip
python -m pip install pandas openpyxl jupyter
```

Abra `conciliacao.ipynb` e use **Kernel → Restart & Run All**. O resultado sai em `saida/resultado_conciliacao.xlsx`.

Resultado esperado com os dados de exemplo: 8 linhas no período, sendo 4 `paga`, 3 `nao_paga` e 1 `parcial`.

### Usar com outros dados

Altere apenas o bloco **1. Parâmetros**: arquivos, separador/encoding do CSV, período (`INICIO`/`FIM`), regex do cliente e nomes das colunas. Para a lista de placas reais, crie um `placas.txt` (uma placa por linha); ele é ignorado pelo Git.

## Estrutura

```
├── conciliacao.ipynb
├── requirements.txt
├── README.md
├── .gitignore
└── exemplos/
    ├── faturamento_exemplo.csv
    ├── devolutiva_exemplo.xlsx
    └── placas_exemplo.txt
```

## Limitações conhecidas

- Só enxerga notas **registradas no sistema de origem**. Notas emitidas fora dele exigem conferência manual (e-mails e relatórios). O resultado não é a lista completa de pendências.
- Compara a **existência** da nota, não os **valores** pagos.
- O sufixo após hífen é ignorado: assume-se que `1234-1` e `1234` são a mesma nota. Foi verificado nos dados originais que nenhuma nota base tinha mais de um sufixo, mas isso deve ser reconferido a cada rodada.
- A data usada para a quinzena é a emissão da nota fiscal; outra regra de negócio muda o resultado.
- Linhas fora do período são descartadas, não analisadas.

## Próximos passos

- Converter o notebook em script (`python conciliacao.py --inicio ... --fim ...`).
- Comparar também os valores (frete × custo da viagem).
- Testes automatizados para as funções de limpeza.

## Habilidades demonstradas

Pandas (filtros, `explode`, `groupby`, datas), limpeza de dados reais, validação independente de resultados, parametrização, documentação e boas práticas de versionamento.
