"""
Extrai os valores de I.S.S. do Relatório de Pagamentos e gera uma planilha Excel consolidada.

Excel gerado (colunas):
    Código do Condomínio | Nome Completo do Autônomo | Data de ocorrência |
    Remuneração Bruta | Valor ISS | Pagamento Programado?

Regras de negócio:
- "Código do Condomínio" é o código informado no cabeçalho de cada bloco do
  relatório: "Empresa: 00XX - ...". O código do autônomo (ex.: 000001) NÃO é
  usado como código da planilha.
- Só entram registros com I.S.S. preenchido e maior que zero.
- IMPORTANTE: Só entram pagamentos com data dentro do mês anterior ao mês de execução do
  script (calculado dinamicamente a partir da data atual, não da data do
  relatório).

Dependências Python (requirements.txt):
    pdfplumber
    openpyxl
"""

from __future__ import annotations

import argparse
import logging
import re
import sys
from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path

try:
    import pdfplumber
except ImportError as exc:  # pragma: no cover - mensagem amigável de setup
    raise SystemExit(
        "A biblioteca 'pdfplumber' não está instalada.\n"
        "Instale as dependências com: pip install -r requirements.txt"
    ) from exc

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

logging.basicConfig(level=logging.INFO, format="%(message)s")
logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Expressões regulares
# ---------------------------------------------------------------------------

# Exemplo de linha no relatório (dados fictícios):
#   "Empresa: 0001 - CONDOMINIO EDIFICIO EXEMPLO ALFA CNPJ: 00.000.000/0001-00"
#   "Empresa: 0002 - ED. EXEMPLO BETA - FULANO DE TAL CPF: 000.000.000-00"
# Só nos interessa o código de 4 dígitos; o nome da empresa não é usado na
# planilha final, então evitamos capturá-lo (isso também evita o bug de
# capturar "CNPJ: ..." / "CPF: ..." junto do nome).
EMPRESA_RE = re.compile(r"Empresa\s*:\s*(\d{4})\b", re.IGNORECASE)

# Exemplo de linha de autônomo (dados fictícios):
#   "000001 - FULANO DE TAL 1001 05/09/2026 3.000,00 330,00 150,00 2.520,00 600,00"
# Grupos: (código autônomo) - (nome) (contrato) (data) (valores monetários...)
AUTONOMO_RE = re.compile(
    r"^\s*(\d{6})\s*-\s*(.+?)\s+(\d+)\s+"
    r"(\d{2}/\d{2}/\d{4})\s+(.*)$"
)

MONEY_RE = re.compile(r"\d{1,3}(?:\.\d{3})*,\d{2}")


@dataclass(frozen=True)
class Registro:
    """Um lançamento de I.S.S. de um autônomo em um condomínio/empresa."""

    codigo_empresa: str
    nome_autonomo: str
    data: date
    remuneracao_bruta: float
    iss: float


# ---------------------------------------------------------------------------
# Utilidades de valores/datas
# ---------------------------------------------------------------------------


def valor_brasileiro_para_float(valor: str) -> float:
    """Converte '1.234,56' -> 1234.56. Lança ValueError se o formato for inválido."""
    try:
        return float(valor.replace(".", "").replace(",", "."))
    except ValueError as exc:
        raise ValueError(f"Valor monetário inválido: {valor!r}") from exc


def primeiro_e_ultimo_dia_mes_anterior(data_referencia: date) -> tuple[date, date]:
    """Retorna (primeiro dia, último dia) do mês anterior ao de `data_referencia`."""
    primeiro_mes_atual = data_referencia.replace(day=1)
    ultimo_mes_anterior = primeiro_mes_atual - timedelta(days=1)
    primeiro_mes_anterior = ultimo_mes_anterior.replace(day=1)
    return primeiro_mes_anterior, ultimo_mes_anterior


# ---------------------------------------------------------------------------
# Extração do PDF
# ---------------------------------------------------------------------------


def extrair_linhas_pdf(pdf_path: Path) -> list[str]:
    """
    Extrai o texto do PDF diretamente da camada de texto (sem OCR).

    O relatório é um PDF nativo (gerado por sistema, não escaneado), então a
    extração direta é exata e muito mais rápida/confiável do que OCR.
    """
    linhas: list[str] = []
    try:
        with pdfplumber.open(pdf_path) as pdf:
            if len(pdf.pages) == 0:
                raise ValueError("O PDF não contém páginas.")
            for pagina in pdf.pages:
                texto = pagina.extract_text() or ""
                linhas.extend(texto.splitlines())
    except Exception as exc:
        raise RuntimeError(f"Falha ao ler o PDF '{pdf_path}': {exc}") from exc

    if not linhas:
        logger.warning(
            "Nenhum texto foi extraído do PDF. Verifique se o arquivo não é "
            "um PDF escaneado (imagem) sem camada de texto."
        )
    return linhas


def extrair_registros(linhas: list[str]) -> list[Registro]:
    """
    Percorre as linhas do relatório e monta a lista de lançamentos de I.S.S.

    Ao encontrar "Empresa: 00XX", guarda esse código como a empresa/condomínio
    atual. Todas as linhas de autônomos seguintes pertencem a essa empresa até
    a próxima linha "Empresa: 00XX" (o cabeçalho se repete a cada página,
    inclusive quando um mesmo condomínio continua na página seguinte).
    """
    registros: list[Registro] = []
    empresa_atual: str | None = None
    linhas_ignoradas_sem_empresa = 0

    for numero_linha, linha_original in enumerate(linhas, start=1):
        linha = " ".join(linha_original.strip().split())
        if not linha:
            continue

        # 1) Cabeçalho da empresa/condomínio.
        empresa_match = EMPRESA_RE.search(linha)
        if empresa_match:
            empresa_atual = empresa_match.group(1).zfill(4)
            continue

        # 2) Linha de autônomo — só processamos se já sabemos a empresa atual.
        match = AUTONOMO_RE.match(linha)
        if not match:
            continue

        if empresa_atual is None:
            linhas_ignoradas_sem_empresa += 1
            continue

        nome = " ".join(match.group(2).split())
        data_texto = match.group(4)
        resto = match.group(5)

        try:
            data_pagamento = datetime.strptime(data_texto, "%d/%m/%Y").date()
        except ValueError:
            logger.warning(
                "Linha %d: data inválida '%s' — registro ignorado.",
                numero_linha, data_texto,
            )
            continue

        valores = MONEY_RE.findall(resto)

        # Ordem das colunas monetárias no relatório, após a data:
        #   Remuneração | I.R.R.F | I.N.S.S | I.S.S | Outros Desc. | Valor líquido | INSS Empresa
        # Em geral, I.R.R.F e Outros Desc. vêm em branco. Quando o I.S.S. tem
        # valor, aparecem 5 números (Remuneração, I.N.S.S, I.S.S, Valor
        # líquido, INSS Empresa) e o I.S.S. é o 3º (índice 2). Quando o I.S.S.
        # está vazio, aparecem só 4 números e não há I.S.S. a extrair.
        if len(valores) < 5:
            continue

        try:
            remuneracao_bruta = valor_brasileiro_para_float(valores[0])
            iss = valor_brasileiro_para_float(valores[2])
        except ValueError:
            logger.warning(
                "Linha %d: remuneração bruta ou valor de I.S.S. inválido — ignorado.",
                numero_linha,
            )
            continue

        if iss <= 0:
            continue

        registros.append(
            Registro(
                codigo_empresa=empresa_atual,
                nome_autonomo=nome,
                data=data_pagamento,
                remuneracao_bruta=remuneracao_bruta,
                iss=iss,
            )
        )

    if linhas_ignoradas_sem_empresa:
        logger.warning(
            "%d linha(s) de autônomo foram ignoradas por aparecerem antes de "
            "qualquer cabeçalho 'Empresa:' reconhecido.",
            linhas_ignoradas_sem_empresa,
        )

    return registros


def processar_pdf(
    pdf_path: Path,
    data_referencia: date | None = None,
) -> tuple[list[Registro], date, date]:
    """Lê o PDF e devolve os registros filtrados pelo mês anterior + o período usado."""
    if data_referencia is None:
        data_referencia = date.today()

    inicio, fim = primeiro_e_ultimo_dia_mes_anterior(data_referencia)

    linhas = extrair_linhas_pdf(pdf_path)
    registros = extrair_registros(linhas)

    registros_filtrados = [r for r in registros if inicio <= r.data <= fim]

    return registros_filtrados, inicio, fim


# ---------------------------------------------------------------------------
# Geração do Excel
# ---------------------------------------------------------------------------


def gerar_excel(registros: list[Registro], caminho_saida: Path) -> None:
    """Gera uma planilha com um lançamento por linha e marcação de pagamento."""
    wb = Workbook()
    ws = wb.active
    ws.title = "I.S.S."

    cabecalho = [
        "Código",
        "Nome Completo do Autônomo",
        "Data da Prestação",
        "Remuneração Bruta",
        "Valor ISS",
        "Pagamento Programado?",
    ]
    for coluna, valor in enumerate(cabecalho, start=1):
        celula = ws.cell(row=1, column=coluna, value=valor)
        celula.font = Font(bold=True)
        celula.fill = PatternFill(fill_type="solid", fgColor="D9EAF7")
        celula.alignment = Alignment(horizontal="center", vertical="center")

    itens = sorted(
        registros,
        key=lambda r: (r.codigo_empresa, r.nome_autonomo.upper(), r.data, r.remuneracao_bruta),
    )

    for linha, registro in enumerate(itens, start=2):
        ws.cell(linha, 1, registro.codigo_empresa)
        ws.cell(linha, 2, registro.nome_autonomo)
        celula_data = ws.cell(linha, 3, registro.data)
        celula_data.number_format = "DD/MM/YYYY"
        celula_remuneracao = ws.cell(linha, 4, round(registro.remuneracao_bruta, 2))
        celula_remuneracao.number_format = 'R$ #,##0.00'
        celula_iss = ws.cell(linha, 5, round(registro.iss, 2))
        celula_iss.number_format = 'R$ #,##0.00'
        ws.cell(linha, 6, "Não")

    # Lista suspensa para marcar facilmente se o pagamento foi programado.
    from openpyxl.worksheet.datavalidation import DataValidation
    validacao_pagamento = DataValidation(type="list", formula1='"Sim,Não"', allow_blank=False)
    validacao_pagamento.error = "Selecione Sim ou Não."
    validacao_pagamento.errorTitle = "Opção inválida"
    validacao_pagamento.prompt = "Marque Sim quando o pagamento já estiver programado."
    validacao_pagamento.promptTitle = "Pagamento programado?"
    ws.add_data_validation(validacao_pagamento)
    if len(itens) > 0:
        validacao_pagamento.add(f"F2:F{len(itens) + 1}")

    larguras = {1: 20, 2: 45, 3: 18, 4: 20, 5: 16, 6: 24}
    for coluna, largura in larguras.items():
        ws.column_dimensions[get_column_letter(coluna)].width = largura
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    ws.row_dimensions[1].height = 30

    try:
        caminho_saida.parent.mkdir(parents=True, exist_ok=True)
        wb.save(caminho_saida)
    except OSError as exc:
        raise RuntimeError(f"Falha ao salvar o Excel em '{caminho_saida}': {exc}") from exc


# ---------------------------------------------------------------------------
# Resumo em tela
# ---------------------------------------------------------------------------


def mostrar_resumo(
    registros: list[Registro], inicio: date, fim: date, caminho_saida: Path
) -> None:
    total_geral = sum(registro.iss for registro in registros)
    total_formatado = (
        f"R$ {total_geral:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    )

    logger.info("=" * 65)
    logger.info("EXTRAÇÃO DE I.S.S. - AUTÔNOMOS")
    logger.info("=" * 65)
    logger.info("Período considerado: %s a %s", inicio.strftime("%d/%m/%Y"), fim.strftime("%d/%m/%Y"))
    logger.info("Registros com I.S.S.: %d", len(registros))
    logger.info("Linhas finais na planilha: %d", len(registros))
    logger.info("Total de I.S.S.: %s", total_formatado)
    logger.info("Arquivo gerado: %s", caminho_saida)
    logger.info("=" * 65)


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


def parse_data_referencia(texto: str) -> date:
    try:
        return datetime.strptime(texto, "%d/%m/%Y").date()
    except ValueError as exc:
        raise argparse.ArgumentTypeError(
            f"Data inválida '{texto}'. Use o formato DD/MM/AAAA."
        ) from exc


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Extrai data, remuneração bruta e I.S.S. de pagamentos de autônomos (PDF) para Excel."
    )
    parser.add_argument(
        "pdf",
        nargs="?",
        help=("Caminho do relatório PDF. Se omitido, o programa procura "
              "automaticamente um único PDF na mesma pasta do script."),
    )
    parser.add_argument(
        "-o", "--saida", default="autonomos_iss.xlsx",
        help="Caminho/nome do Excel de saída (padrão: autonomos_iss.xlsx).",
    )
    parser.add_argument(
        "--data-referencia",
        type=parse_data_referencia,
        help="Data de referência DD/MM/AAAA. O programa processará o mês "
             "anterior a essa data. Padrão: data atual do sistema.",
    )
    args = parser.parse_args()

    # Se o PDF não foi informado, procura automaticamente um único PDF
    # na mesma pasta onde está este script.
    if args.pdf:
        pdf_path = Path(args.pdf)
    else:
        pasta_script = Path(__file__).resolve().parent
        pdfs = sorted(pasta_script.glob("*.pdf"))

        if len(pdfs) == 0:
            logger.error(
                "Nenhum arquivo PDF foi encontrado na pasta do script: %s",
                pasta_script,
            )
            return 1

        if len(pdfs) > 1:
            logger.error(
                "Foram encontrados %d arquivos PDF na pasta do script. "
                "Informe qual PDF deseja processar usando: python %s <arquivo.pdf>",
                len(pdfs),
                Path(__file__).name,
            )
            for pdf in pdfs:
                logger.error(" - %s", pdf.name)
            return 1

        pdf_path = pdfs[0]
        logger.info("PDF encontrado automaticamente: %s", pdf_path.name)

    if not pdf_path.exists():
        logger.error("PDF não encontrado: %s", pdf_path)
        return 1
    if pdf_path.suffix.lower() != ".pdf":
        logger.warning("O arquivo informado não tem extensão .pdf: %s", pdf_path)

    try:
        registros, inicio, fim = processar_pdf(pdf_path, args.data_referencia)
    except (RuntimeError, ValueError) as exc:
        logger.error("Erro ao processar o PDF: %s", exc)
        return 1

    if not registros:
        logger.warning(
            "Nenhum registro de I.S.S. encontrado no período %s a %s. "
            "A planilha será gerada apenas com o cabeçalho.",
            inicio.strftime("%d/%m/%Y"), fim.strftime("%d/%m/%Y"),
        )

    caminho_saida = Path(args.saida)
    try:
        gerar_excel(registros, caminho_saida)
    except RuntimeError as exc:
        logger.error(str(exc))
        return 1

    mostrar_resumo(registros, inicio, fim, caminho_saida)
    return 0


if __name__ == "__main__":
    sys.exit(main())