# Conciliação de faturamento (logística)

Projeto de análise de dados em Python/Pandas que cruza duas planilhas para descobrir **quais viagens faturadas não foram reconhecidas/pagas pelo cliente**.

> Todos os dados deste repositório são **fictícios**. Nomes, placas e notas foram inventados.

## Pergunta de negócio

Para uma lista de veículos e um cliente específico, quais linhas do faturamento da quinzena **não aparecem** na planilha de devolutiva do cliente?

| Planilha | Papel |
|---|---|
| **B – faturamento geral** | Fonte dos dados; recebe os filtros (placas, cliente, período) |
| **A – devolutiva do cliente** | Usada para verificar se a nota foi reconhecida |

## Como funciona (v2)

A unidade de análise é o **CT-e** (uma linha da planilha B). Como a A não tem CT-e, a ponte é **nota fiscal + placa**.

1. **Filtra a B** por placas (cavalo ou carreta, qualquer coluna `Placa...`) e cliente (regex no destinatário), avisando placas sem viagem.
2. **Limpa** notas e placas nas duas planilhas (tudo lido como texto).
3. **Classifica** cada CT-e: `paga` (todas as notas na A com placa do CT-e), `nao_paga`, `parcial` (pendente) e `sem_nota`. Notas que existem na A com outra placa aparecem em `notas_placa_diferente` para revisão humana.
4. **Separa o período** (quinzena 1 = dias 1-15; 2 = dia 16 ao fim do mês) em `no_periodo`, `fora_periodo` e `sem_data`, que somam o total.
5. **Exporta** um Excel: aba no modelo do cliente (`2Q SETEMBRO`) com as pendências, mais `resumo`, `pagas_periodo`, `detalhes_pendencias`, `sem_data` e `alertas`.
6. **Valida** por um segundo caminho (busca por trecho) e confere frete + ICMS contra o total.

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
python -m venv venv
venv\Scripts\activate        # Windows
python -m pip install -r requirements.txt
python executar.py --ano 2026 --mes 9 --quinzena 2 --filial LONDRINA
```

Usa os dados de `exemplos/` por padrão. Com dados reais: `--a`, `--b`, `--placas`, `--encoding latin-1`, `--cliente`. Resultado em `saida/conciliacao.xlsx`.
Esperado nos exemplos: 8 CT-e no período = 4 paga, 2 nao_paga, 1 parcial, 1 sem_nota.

Testes: `python -m pip install -r requirements-dev.txt` e `python -m pytest`.

## Estrutura

```
├── executar.py        # linha de comando
├── src/conciliacao.py # funções (limpeza, classificação, exportação)
├── tests/
├── exemplos/          # dados fictícios + gerar_exemplos.py
├── conciliacao.ipynb  # versão didática (v1)
└── README.md
```

## Limitações conhecidas

- Só enxerga notas **registradas no sistema de origem**. Notas emitidas fora dele exigem conferência manual (e-mails e relatórios). O resultado não é a lista completa de pendências.
- Compara a **existência** da nota, não os **valores** pagos.
- O sufixo após hífen é ignorado: assume-se que `1234-1` e `1234` são a mesma nota. Foi verificado nos dados originais que nenhuma nota base tinha mais de um sufixo, mas isso deve ser reconferido a cada rodada.
- A data usada para a quinzena é a emissão da nota fiscal; outra regra de negócio muda o resultado.
- Números com ponto sem vírgula (`30.000`) são lidos como decimal; confira o formato do CSV real.

## Próximos passos

- Leitura automática dos XML (MDF-e, CT-e, NF-e) recebidos por e-mail.
- Comparar também os valores (frete × custo da viagem).
- Testes automatizados para as funções de limpeza.

## Habilidades demonstradas

Pandas (filtros, `explode`, `groupby`, datas), limpeza de dados reais, validação independente de resultados, parametrização, documentação e boas práticas de versionamento.
