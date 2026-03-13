"""
powerbi_local — Conexão com Power BI Desktop aberto localmente.

Uso básico:
    from powerbi_local.connection import PowerBILocalConnection

    with PowerBILocalConnection() as pbi:
        tabelas = pbi.list_tables()
        dados = pbi.preview_table("Vendas", top=20)
        resultado = pbi.execute_dax("EVALUATE SUMMARIZE('Vendas', 'Vendas'[Região])")
"""

from .connection import PowerBILocalConnection
from .discover_port import get_powerbi_port

__all__ = ["PowerBILocalConnection", "get_powerbi_port"]
