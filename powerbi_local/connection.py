"""
Conexão com o Power BI Desktop local via Analysis Services (OLEDB/XMLA).

Dependências (instalar no Windows):
    pip install pyodbc
    # Requer também o driver "Microsoft OLE DB Provider for Analysis Services"
    # instalado junto com o Power BI Desktop ou SSAS.

Alternativa leve (sem driver OLEDB):
    pip install requests  # para consultas via HTTP XMLA (se habilitado)
"""

from __future__ import annotations

import sys
from dataclasses import dataclass, field
from typing import Any

from discover_port import get_powerbi_port


@dataclass
class PowerBILocalConnection:
    """
    Representa uma conexão com o modelo de dados do Power BI Desktop local.

    Attributes:
        port: Porta TCP do Analysis Services local.
        database: Nome do banco/modelo (deixe vazio para auto-detectar).
        timeout: Timeout de conexão em segundos.
    """

    port: int = field(default_factory=get_powerbi_port)
    database: str = ""
    timeout: int = 30

    # Conexão interna (pyodbc ou adodbapi)
    _conn: Any = field(default=None, init=False, repr=False)

    # -----------------------------------------------------------------
    # Propriedades
    # -----------------------------------------------------------------

    @property
    def connection_string(self) -> str:
        """Connection string OLEDB para Analysis Services local."""
        parts = [
            "Provider=MSOLAP",
            f"Data Source=localhost:{self.port}",
        ]
        if self.database:
            parts.append(f"Initial Catalog={self.database}")
        return ";".join(parts)

    # -----------------------------------------------------------------
    # Ciclo de vida
    # -----------------------------------------------------------------

    def connect(self) -> "PowerBILocalConnection":
        """
        Abre a conexão com o Power BI Desktop local.

        Returns:
            Self (para uso em cadeia / context manager).

        Raises:
            RuntimeError: Se a conexão falhar.
        """
        if sys.platform != "win32":
            raise RuntimeError(
                "O Power BI Desktop roda apenas em Windows. "
                "Execute este script na mesma máquina Windows onde o Power BI está aberto."
            )

        try:
            import pyodbc  # type: ignore
        except ImportError:
            raise RuntimeError(
                "pyodbc não está instalado. Execute: pip install pyodbc\n"
                "Também é necessário o driver 'Microsoft OLE DB Provider for Analysis Services'."
            )

        try:
            self._conn = pyodbc.connect(self.connection_string, timeout=self.timeout)
            print(f"[OK] Conectado ao Power BI Desktop na porta {self.port}.")

            # Auto-detecta o banco se não especificado
            if not self.database:
                self.database = self._detect_database()

        except pyodbc.Error as e:
            raise RuntimeError(
                f"Falha ao conectar ao Power BI Desktop (porta {self.port}).\n"
                f"Detalhes: {e}\n\n"
                "Verifique:\n"
                "  1. O Power BI Desktop está aberto com um arquivo .pbix?\n"
                "  2. O driver MSOLAP está instalado?\n"
                "  3. A porta está correta?"
            ) from e

        return self

    def disconnect(self) -> None:
        """Fecha a conexão."""
        if self._conn:
            self._conn.close()
            self._conn = None
            print("[OK] Conexão encerrada.")

    def __enter__(self) -> "PowerBILocalConnection":
        return self.connect()

    def __exit__(self, *_) -> None:
        self.disconnect()

    # -----------------------------------------------------------------
    # Consultas DAX / MDX
    # -----------------------------------------------------------------

    def execute_dax(self, query: str) -> list[dict]:
        """
        Executa uma consulta DAX no modelo do Power BI.

        Args:
            query: Expressão DAX (ex: "EVALUATE <tabela>").

        Returns:
            Lista de dicionários com os resultados.
        """
        self._ensure_connected()
        cursor = self._conn.cursor()
        cursor.execute(query)
        columns = [col[0] for col in cursor.description]
        return [dict(zip(columns, row)) for row in cursor.fetchall()]

    def list_tables(self) -> list[str]:
        """
        Lista todas as tabelas disponíveis no modelo do Power BI.

        Returns:
            Lista com os nomes das tabelas.
        """
        results = self.execute_dax(
            "SELECT [TABLE_NAME] FROM $SYSTEM.DBSCHEMA_TABLES "
            "WHERE TABLE_TYPE = 'TABLE'"
        )
        return [r["TABLE_NAME"] for r in results]

    def list_measures(self) -> list[dict]:
        """
        Lista todas as medidas (measures) do modelo.

        Returns:
            Lista de dicts com 'table', 'name' e 'expression'.
        """
        results = self.execute_dax(
            "SELECT [MEASUREGROUP_NAME], [MEASURE_NAME], [EXPRESSION] "
            "FROM $SYSTEM.MDSCHEMA_MEASURES "
            "WHERE MEASURE_IS_VISIBLE"
        )
        return [
            {
                "table": r["MEASUREGROUP_NAME"],
                "name": r["MEASURE_NAME"],
                "expression": r.get("EXPRESSION", ""),
            }
            for r in results
        ]

    def preview_table(self, table_name: str, top: int = 10) -> list[dict]:
        """
        Retorna as primeiras linhas de uma tabela do modelo.

        Args:
            table_name: Nome da tabela.
            top: Número de linhas (default 10).

        Returns:
            Lista de dicts com os dados.
        """
        query = f"EVALUATE TOPN({top}, '{table_name}')"
        return self.execute_dax(query)

    # -----------------------------------------------------------------
    # Internos
    # -----------------------------------------------------------------

    def _ensure_connected(self) -> None:
        if self._conn is None:
            raise RuntimeError("Conexão não estabelecida. Chame .connect() primeiro.")

    def _detect_database(self) -> str:
        """Auto-detecta o nome do primeiro banco de dados disponível."""
        try:
            cursor = self._conn.cursor()
            cursor.execute(
                "SELECT [CATALOG_NAME] FROM $SYSTEM.DBSCHEMA_CATALOGS"
            )
            rows = cursor.fetchall()
            if rows:
                db = rows[0][0]
                print(f"[OK] Modelo detectado: '{db}'")
                return db
        except Exception:
            pass
        return ""
