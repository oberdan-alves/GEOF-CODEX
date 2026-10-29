from pathlib import Path
import os
import ntpath
import sys
import shutil
import json
import tempfile

SUBPASTA_DADOS = os.path.join("Planilhas Geradoras Cotação Pagamento Prestação", "GeradorCotacao")
_pasta_compartilhada = None


def desativar_pasta_compartilhada():
    global _pasta_compartilhada
    _pasta_compartilhada = None


def pasta_dados_compartilhados(pasta_geof):
    pasta = normalizar_caminho(pasta_geof)
    if not pasta:
        raise ValueError("Escolha a PASTA DE TRABALHO (até GEOF) antes de salvar dados.")
    modulo = ntpath if os.name == "nt" else os.path
    return modulo.join(modulo.dirname(pasta), "Planilhas Geradoras Cotação Pagamento Prestação", "GeradorCotacao")


def configurar_pasta_compartilhada(pasta_geof):
    """Inicializa os dados compartilhados sem substituir arquivos já existentes."""
    global _pasta_compartilhada
    destino = Path(pasta_dados_compartilhados(pasta_geof))
    if destino == _pasta_compartilhada:
        destino.mkdir(parents=True, exist_ok=True)
        return destino
    _pasta_compartilhada = None
    destino.mkdir(parents=True, exist_ok=True)
    origens = [pasta_programa(), _pasta_recursos_empacotados()]
    for origem in origens:
        arquivos = [origem / nome for nome in ("config_padrao.json", "feriados.json")]
        assets = origem / "assets"
        if assets.is_dir():
            arquivos.extend(p for p in assets.rglob("*") if p.is_file())
        for arquivo in arquivos:
            if not arquivo.is_file():
                continue
            alvo = destino / arquivo.relative_to(origem)
            alvo.parent.mkdir(parents=True, exist_ok=True)
            try:
                with alvo.open("xb") as saida, arquivo.open("rb") as entrada:
                    shutil.copyfileobj(entrada, saida)
            except FileExistsError:
                pass
    _pasta_compartilhada = destino
    return destino


def salvar_json_compartilhado(nome, dados):
    if _pasta_compartilhada is None:
        raise ValueError("Selecione a PASTA DE TRABALHO (até GEOF) antes de salvar.")
    temporario = None
    try:
        with tempfile.NamedTemporaryFile(mode="w", encoding="utf-8", dir=_pasta_compartilhada,
                                         prefix=nome + ".", suffix=".tmp", delete=False) as arquivo:
            temporario = arquivo.name
            json.dump(dados, arquivo, ensure_ascii=False, indent=2)
            arquivo.flush()
            os.fsync(arquivo.fileno())
        os.replace(temporario, _pasta_compartilhada / nome)
    finally:
        if temporario and os.path.exists(temporario):
            os.unlink(temporario)


def normalizar_caminho(caminho):
    """Preserva UNC/unidades e usa os separadores nativos antes de abrir no Windows."""
    texto = str(caminho).strip() if caminho is not None else ""
    if len(texto) >= 2 and texto[0] == texto[-1] and texto[0] in ('"', "'"):
        texto = texto[1:-1].strip()
    if not texto:
        return ""
    return ntpath.normpath(texto) if os.name == "nt" else os.path.normpath(texto)


def _pasta_recursos_empacotados():
    if getattr(sys, "frozen", False):
        base = Path(sys._MEIPASS)
    else:
        base = Path(__file__).resolve().parent

    return base


def caminho_recurso(nome):
    """Após escolher GEOF, todos os recursos são lidos da pasta compartilhada."""
    return (_pasta_compartilhada or _pasta_recursos_empacotados()) / nome


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
    if _pasta_compartilhada is not None:
        return _pasta_compartilhada / nome
    externo = pasta_programa() / nome

    if externo.is_file():
        return externo

    return caminho_recurso(nome)
