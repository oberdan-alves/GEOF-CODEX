
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

    print("VS CODE -> GITHUB")
    print(f"Projeto: {REPO}")
    print("\nATENÇÃO: o envio forçado pode substituir commits do GitHub.")

    confirmar = input(
        "\nDeseja publicar todos os arquivos locais? (s/n): "
    ).strip().lower()

    if confirmar != "s":
        print("Operação cancelada.")
        return

    executar(["git", "add", "-A"])

    resultado = subprocess.run(
        ["git", "diff", "--cached", "--quiet"],
        cwd=REPO
    )

    if resultado.returncode == 0:
        print("\nNenhuma alteração nova para registrar.")
    elif resultado.returncode == 1:
        executar([
            "git", "commit",
            "-m", "Atualiza projeto pela versao local"
        ])
    else:
        print("ERRO ao verificar as alterações preparadas.")
        sys.exit(1)

    # Atualiza a referência remota antes de verificar o envio.
    executar(["git", "fetch", "origin"])

    # Tenta primeiro o envio normal, sem sobrescrever o histórico remoto.
    envio = executar(
        ["git", "push", "origin", "HEAD:main"],
        verificar=False
    )

    if envio.returncode == 0:
        print("\nPublicação concluída com sucesso.")
        return

    print("\nO envio normal falhou.")
    print("Pode haver divergência entre o histórico local e o remoto.")

    forcar = input(
        "\nDeseja tentar substituir o histórico remoto? (s/n): "
    ).strip().lower()

    if forcar != "s":
        print("Envio forçado cancelado. Verifique o erro acima.")
        return

    executar([
        "git", "push",
        "--force-with-lease",
        "origin", "HEAD:main"
    ])

    print("\nPublicação concluída.")
    executar(["git", "status"])


if __name__ == "__main__":
    main()