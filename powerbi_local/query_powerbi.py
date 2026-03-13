"""
Interface de linha de comando para consultar o Power BI Desktop local.

Uso:
    python query_powerbi.py --list-tables
    python query_powerbi.py --list-measures
    python query_powerbi.py --preview "NomeDaTabela"
    python query_powerbi.py --dax "EVALUATE TOPN(5, 'Vendas')"
    python query_powerbi.py --port 12345 --dax "EVALUATE 'Clientes'"

Requisitos (Windows):
    pip install pyodbc
    Driver: Microsoft OLE DB Provider for Analysis Services (incluso no Power BI Desktop)
"""

from __future__ import annotations

import argparse
import json
import sys

# Garante importação relativa quando executado diretamente
import os
sys.path.insert(0, os.path.dirname(__file__))

from connection import PowerBILocalConnection
from discover_port import get_powerbi_port


def print_table(rows: list[dict]) -> None:
    """Imprime uma lista de dicts no formato de tabela simples."""
    if not rows:
        print("(sem resultados)")
        return

    headers = list(rows[0].keys())
    col_widths = {h: len(h) for h in headers}
    for row in rows:
        for h in headers:
            col_widths[h] = max(col_widths[h], len(str(row.get(h, ""))))

    sep = "+-" + "-+-".join("-" * col_widths[h] for h in headers) + "-+"
    header_line = "| " + " | ".join(h.ljust(col_widths[h]) for h in headers) + " |"

    print(sep)
    print(header_line)
    print(sep)
    for row in rows:
        line = "| " + " | ".join(str(row.get(h, "")).ljust(col_widths[h]) for h in headers) + " |"
        print(line)
    print(sep)
    print(f"  {len(rows)} linha(s)")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Consulta o modelo de dados do Power BI Desktop local.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument(
        "--port",
        type=int,
        default=None,
        help="Porta do Analysis Services local (auto-detectada se omitida).",
    )
    parser.add_argument(
        "--database",
        default="",
        help="Nome do banco/modelo (auto-detectado se omitido).",
    )
    parser.add_argument(
        "--list-tables",
        action="store_true",
        help="Lista todas as tabelas do modelo.",
    )
    parser.add_argument(
        "--list-measures",
        action="store_true",
        help="Lista todas as medidas (measures) do modelo.",
    )
    parser.add_argument(
        "--preview",
        metavar="TABELA",
        help="Exibe as primeiras 10 linhas de uma tabela.",
    )
    parser.add_argument(
        "--top",
        type=int,
        default=10,
        help="Número de linhas para --preview (default: 10).",
    )
    parser.add_argument(
        "--dax",
        metavar="QUERY",
        help="Executa uma expressão DAX personalizada.",
    )
    parser.add_argument(
        "--json",
        action="store_true",
        dest="output_json",
        help="Saída em formato JSON.",
    )

    args = parser.parse_args()

    # Verifica se pelo menos uma ação foi solicitada
    if not any([args.list_tables, args.list_measures, args.preview, args.dax]):
        parser.print_help()
        sys.exit(0)

    # Determina a porta
    try:
        port = args.port if args.port else get_powerbi_port()
    except RuntimeError as e:
        print(f"[ERRO] {e}")
        sys.exit(1)

    # Conecta e executa
    try:
        with PowerBILocalConnection(port=port, database=args.database) as pbi:

            if args.list_tables:
                tables = pbi.list_tables()
                print(f"\n=== Tabelas no modelo ({len(tables)}) ===")
                if args.output_json:
                    print(json.dumps(tables, ensure_ascii=False, indent=2))
                else:
                    for t in tables:
                        print(f"  - {t}")

            if args.list_measures:
                measures = pbi.list_measures()
                print(f"\n=== Medidas no modelo ({len(measures)}) ===")
                if args.output_json:
                    print(json.dumps(measures, ensure_ascii=False, indent=2))
                else:
                    print_table(measures)

            if args.preview:
                rows = pbi.preview_table(args.preview, top=args.top)
                print(f"\n=== Preview: '{args.preview}' (top {args.top}) ===")
                if args.output_json:
                    print(json.dumps(rows, ensure_ascii=False, indent=2, default=str))
                else:
                    print_table(rows)

            if args.dax:
                rows = pbi.execute_dax(args.dax)
                print(f"\n=== Resultado DAX ===")
                if args.output_json:
                    print(json.dumps(rows, ensure_ascii=False, indent=2, default=str))
                else:
                    print_table(rows)

    except RuntimeError as e:
        print(f"\n[ERRO] {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
