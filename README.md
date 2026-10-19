# Extração de I.S.S. de Autônomos

Lê um relatório PDF de **Pagamentos – Autônomos**, identifica os lançamentos com I.S.S. do mês anterior e gera uma planilha Excel (`.xlsx`).

## Instalação

Requer Python 3.10+ e um PDF com texto selecionável (PDF escaneado não é suportado).

```bash
pip install -r requirements.txt
```

## Uso

```bash
# PDF único na pasta do script (detecção automática)
python extrair_iss_autonomos_automatico.py

# PDF específico, saída e mês de referência opcionais
python extrair_iss_autonomos_automatico.py relatorio.pdf -o resultado.xlsx --data-referencia 01/10/2026
```

| Opção | Descrição | Padrão |
|---|---|---|
| `pdf` | Caminho do PDF | único `.pdf` da pasta do script |
| `-o`, `--saida` | Excel de saída | `autonomos_iss.xlsx` |
| `--data-referencia` | `DD/MM/AAAA`; processa o **mês anterior** a essa data | data atual |

## Testar com o exemplo

O PDF em `exemplo/` contém apenas dados fictícios:

```bash
python extrair_iss_autonomos_automatico.py exemplo/Relatorio_Autonomos_EXEMPLO.pdf --data-referencia 01/10/2026
```

Resultado esperado: 5 registros e total de I.S.S. de R$ 650,00.

## Regras

- O código do condomínio vem do cabeçalho `Empresa: 00XX`.
- Entram só pagamentos do mês anterior à data de referência, com I.S.S. maior que zero.
- Cada lançamento vira uma linha, ordenada por código e nome.

## Planilha gerada (aba `I.S.S.`)

Código · Nome Completo do Autônomo · Data da Prestação · Remuneração Bruta · Valor ISS · Pagamento Programado? (lista Sim/Não, padrão "Não")

## Problemas comuns

- **Planilha só com cabeçalho:** nenhum I.S.S. no período, ou o PDF não segue o layout esperado. Confira o período no resumo do terminal.
- **Vários PDFs na pasta:** informe o arquivo no comando.

## Privacidade

Não versione relatórios reais: eles contêm CPF/CNPJ e nomes. O `.gitignore` já bloqueia `*.pdf` e `*.xlsx`, exceto o PDF de exemplo.
