"""
Descobre a porta local do Power BI Desktop (Analysis Services embutido).

O Power BI Desktop, quando aberto, inicia uma instância local do
Analysis Services (SSAS) em uma porta dinâmica. Este módulo
encontra essa porta automaticamente.
"""

import subprocess
import re
import os
import sys


def find_powerbi_port() -> int | None:
    """
    Encontra a porta TCP em que o Power BI Desktop está escutando
    via Analysis Services local.

    Funciona em Windows (onde o Power BI Desktop roda).

    Returns:
        Número da porta (int) ou None se não encontrado.
    """
    if sys.platform != "win32":
        print(
            "[AVISO] Detecção de porta local só funciona em Windows, "
            "onde o Power BI Desktop está instalado."
        )
        return None

    try:
        # Procura pelo processo msmdsrv.exe (Analysis Services) filho do Power BI
        result = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq msmdsrv.exe", "/FO", "CSV"],
            capture_output=True,
            text=True,
        )
        if "msmdsrv.exe" not in result.stdout:
            print("[ERRO] Power BI Desktop não está aberto ou não foi encontrado.")
            return None

        # Obtém o PID do msmdsrv.exe
        lines = result.stdout.strip().splitlines()
        pid = None
        for line in lines[1:]:  # pula cabeçalho
            parts = line.strip('"').split('","')
            if parts[0].lower() == "msmdsrv.exe":
                pid = parts[1]
                break

        if not pid:
            return None

        # Encontra a porta TCP que esse PID está escutando
        netstat = subprocess.run(
            ["netstat", "-ano"],
            capture_output=True,
            text=True,
        )

        pattern = re.compile(
            r"TCP\s+127\.0\.0\.1:(\d+)\s+0\.0\.0\.0:0\s+LISTENING\s+" + re.escape(pid)
        )
        match = pattern.search(netstat.stdout)
        if match:
            port = int(match.group(1))
            print(f"[OK] Power BI Desktop encontrado na porta: {port}")
            return port

        print("[ERRO] Porta do Power BI não encontrada via netstat.")
        return None

    except FileNotFoundError as e:
        print(f"[ERRO] Comando não encontrado: {e}")
        return None


def find_powerbi_port_from_workspaces() -> int | None:
    """
    Alternativa: lê a porta diretamente dos arquivos de workspace
    do Power BI Desktop no perfil do usuário.

    Returns:
        Número da porta (int) ou None se não encontrado.
    """
    if sys.platform != "win32":
        return None

    local_app_data = os.environ.get("LOCALAPPDATA", "")
    workspace_root = os.path.join(
        local_app_data, "Microsoft", "Power BI Desktop", "AnalysisServicesWorkspaces"
    )

    if not os.path.isdir(workspace_root):
        print(f"[AVISO] Diretório de workspaces não encontrado: {workspace_root}")
        return None

    # Cada workspace tem um arquivo msmdsrv.port.txt
    for entry in os.scandir(workspace_root):
        if entry.is_dir():
            port_file = os.path.join(entry.path, "Data", "msmdsrv.port.txt")
            if os.path.isfile(port_file):
                with open(port_file) as f:
                    port_str = f.read().strip()
                    if port_str.isdigit():
                        port = int(port_str)
                        print(f"[OK] Porta lida do arquivo de workspace: {port}")
                        return port

    print("[ERRO] Nenhum arquivo de porta encontrado nos workspaces.")
    return None


def get_powerbi_port() -> int:
    """
    Tenta ambas as estratégias para encontrar a porta do Power BI Desktop.

    Returns:
        Porta encontrada.

    Raises:
        RuntimeError: Se não conseguir encontrar a porta.
    """
    port = find_powerbi_port_from_workspaces()
    if port:
        return port

    port = find_powerbi_port()
    if port:
        return port

    raise RuntimeError(
        "Não foi possível encontrar a porta do Power BI Desktop.\n"
        "Verifique se o Power BI Desktop está aberto e com um arquivo .pbix carregado."
    )


if __name__ == "__main__":
    try:
        p = get_powerbi_port()
        print(f"Porta do Power BI Desktop: {p}")
    except RuntimeError as e:
        print(e)
        sys.exit(1)
