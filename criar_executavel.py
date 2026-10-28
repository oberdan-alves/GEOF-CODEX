from pathlib import Path
import subprocess
import sys
import tempfile


def gerar_executavel(pasta_projeto=None):
    if sys.platform != "win32":
        raise RuntimeError("Execute este gerador no Windows.")

    if pasta_projeto is None:
        padrao = r"G:\Meu Drive\python\GeradorCotacao"
        resposta = input(
            f"Pasta do projeto [Enter para {padrao}]: "
        ).strip().strip('"')
        pasta_projeto = resposta or padrao

    pasta = Path(pasta_projeto).expanduser().resolve()
    principal = pasta / "GeradorCotacao.py"

    if not principal.is_file():
        raise FileNotFoundError(f"Arquivo não encontrado: {principal}")

    modulos = [
        "Cotacao",
        "Emails",
        "Mapa",
        "Mensagens",
        "AdjudicacaoOrdem",
    ]

    comando = [
        sys.executable, "-m", "PyInstaller",
        "--noconfirm",
        "--clean",
        "--onefile",
        "--console",
        "--name", "GeradorCotacao",
        "--paths", str(pasta),
        "--distpath", str(pasta / "dist"),
    ]

    for modulo in modulos:
        comando.extend(["--hidden-import", modulo])

    ignorados = {
        "dist", "build", ".venv", "venv", "env",
        "__pycache__", ".git", "node_modules",
    }

    for arquivo in pasta.rglob("*.json"):
        relativo = arquivo.relative_to(pasta)

        if any(parte in ignorados for parte in relativo.parts):
            continue

        comando.extend([
            "--add-data",
            f"{arquivo};{relativo.parent}",
        ])

    with tempfile.TemporaryDirectory(prefix="cotacao_build_") as temporaria:
        comando.extend([
            "--workpath", str(Path(temporaria) / "build"),
            "--specpath", temporaria,
            str(principal),
        ])

        subprocess.run(comando, cwd=str(pasta), check=True)

    executavel = pasta / "dist" / "GeradorCotacao.exe"
    print(f"\nExecutável gerado em: {executavel}")
    return executavel


if __name__ == "__main__":
    gerar_executavel()