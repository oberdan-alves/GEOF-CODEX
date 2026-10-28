from pathlib import Path
import sys


def caminho_recurso(nome):
    """Localiza arquivos incluídos no executável."""
    if getattr(sys, "frozen", False):
        base = Path(sys._MEIPASS)
    else:
        base = Path(__file__).resolve().parent

    return base / nome


def pasta_programa():
    """Localiza a pasta onde está o executável."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent

    return Path(__file__).resolve().parent


def caminho_configuracao(nome):
    """
    Usa primeiro o arquivo ao lado do executável.
    Se não existir, usa o arquivo incluído no pacote.
    """
    externo = pasta_programa() / nome

    if externo.is_file():
        return externo

    return caminho_recurso(nome)