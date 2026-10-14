
import subprocess
import sys
from pathlib import Path

REPO = Path(r"G:\Meu Drive\python\GeradorCotacao")


def executar(comando, verificar=True):
    print(f"\n> {' '.join(comando)}")
    return subprocess.run(
        comando,
        cwd=REPO,
        check=verificar
    )


def main():
    if not (REPO / ".git").exists():
        print(f"ERRO: repositório Git não encontrado em:\n{REPO}")
        sys.exit(1)

    print("GITHUB -> VS CODE")
    print(f"Projeto: {REPO}")
    print("\nATENÇÃO: alterações locais podem ser perdidas.")

    confirmar = input(
        "\nDeseja substituir os arquivos locais pela versão do GitHub? (s/n): "
    ).strip().lower()

    if confirmar != "s":
        print("Operação cancelada.")
        return

    executar(["git", "fetch", "origin"])
    executar(["git", "reset", "--hard", "origin/main"])

    print("\nArquivos não rastreados que poderiam ser removidos:")
    executar(["git", "clean", "-nd"])

    limpar = input(
        "\nDeseja remover também esses arquivos não rastreados? (s/n): "
    ).strip().lower()

    if limpar == "s":
        executar(["git", "clean", "-fd"])

    print("\nSincronização concluída.")
    executar(["git", "status"])


if __name__ == "__main__":
    main()