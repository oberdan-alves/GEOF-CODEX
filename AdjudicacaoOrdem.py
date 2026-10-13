# -*- coding: utf-8 -*-
"""
AdjudicacaoOrdem.py
====================
MÓDULO DO 3º PASSO — Adjudicação e Ordem de Compra
-------------------------------------------------------
Responsável pelos botões do grupo "3º Passo – Adjudicação e Ordem":
    - ADJUDICAÇÃO  (monta o Mapa Comparativo: 1º/2º/3º colocado por item,
                     escolha do vencedor, itens cancelados, PDF + JPG)
    - ORDEM        (gera uma Ordem de Compra em PDF + JPG para cada
                     fornecedor vencedor, com os itens que ele ganhou)
    - SALVAR       (placeholder — salva os defaults da tela)

Consome o cadastro de propostas por fornecedor gerado no 2º Passo
(Mapa.py), lido do arquivo COTACAO\\{código}\\MAPA_EMPRESAS_{código}.json,
e grava nele também as decisões de vencedor/cancelamento tomadas aqui, para
que o botão ORDEM não precise perguntar de novo.

gerar_objeto_automatico() — usada aqui para montar o título do Mapa
Comparativo — vive em GeradorCotacao.py (função compartilhada por todos os
módulos, inclusive Emails.py, que a usa para montar o OBJETO de cada
cotação no E-mail Cotação); aqui só é importada, não redefinida.
"""

import os
import json
import tkinter as tk
from tkinter import messagebox

import customtkinter as ctk

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT

from GeradorCotacao import (
    COR_FUNDO_APP, COR_CARD, COR_TEXTO, COR_DESTAQUE,
    FONTE_BOTAO, FONTE_SECAO, FONTE_TABELA,
    COR_CINZA_ADJ, COR_CIANO_ADJ, COR_VERDE_ADJ, COR_AMARELO_ADJ,
    CONTATOS_ORDEM_PADRAO,
    formatar_moeda, normalizar_texto, extrair_nome_codigo_fornecedor, truncar_em_palavra,
    gerar_objeto_automatico,
    converter_pdf_para_jpg, exibir_preview_pdf_jpg, montar_cabecalho_institucional,
)
from Mapa import carregar_fornecedores_completo

TEXTO_OBSERVACAO_PADRAO = (
    "A cotação foi finalizada com menos de três propostas para o(s) item(ns) {itens}, uma vez "
    "que, realizada a pesquisa de mercado, não foram apresentadas outras propostas, e conforme "
    "pesquisas realizadas no Painel de Preços Federal, Mapa de Preços do DF o preço apresentado "
    "encontra-se em consonância com a realidade de mercado atual e dentro do limite estabelecido "
    "na tabela CMED.\n\n"
    "Foram apresentadas notas fiscais emitidas para outros contratantes, comprovando previamente "
    "que os preços estão em conformidade com os praticados semelhantes de objetos de mesma "
    "natureza, conforme a Lei 14.133/2021 Art. 23 §4."
)


# ----------------------------------------------------------------------
# MONTAGEM DOS DADOS DO MAPA COMPARATIVO
# ----------------------------------------------------------------------
def montar_mapa_itens_completo(linhas_resultado, mapa_fornecedores, fornecedores_texto_completo,
                                saldos_bd):
    """Monta a lista COMPLETA de itens da cotação (na ordem original do
    Controle), cada um com suas propostas ordenadas por menor preço — usada
    para o Mapa Comparativo, que precisa mostrar até itens sem proposta
    ('DESERTO'), diferente do ranking do Passo 26 (Mapa.py)."""
    resultado = []
    for item in linhas_resultado:
        codigo = str(item.get("CODIGO") or "")
        propostas = []
        for nome_forn, dados in mapa_fornecedores.items():
            texto_completo = fornecedores_texto_completo.get(normalizar_texto(nome_forn), nome_forn)
            for it in dados.get("itens", []):
                if it.get("codigo") != codigo:
                    continue
                unitario = it.get("unitario")
                if unitario is None or unitario <= 0:
                    continue
                propostas.append({
                    "fornecedor": texto_completo,
                    "quantidade": it.get("quantidade"),
                    "unitario": unitario,
                    "total": it.get("total"),
                    "prazo": it.get("prazo"),
                    "marca": it.get("marca"),
                })
        propostas.sort(key=lambda p: p["unitario"])
        resultado.append({
            "numero_item": item.get("ITEM"),
            "codigo": codigo,
            "descricao": item.get("DESCRICAO") or "",
            "un": item.get("UN") or "",
            "propostas": propostas,
            "saldo": saldos_bd.get(normalizar_texto(codigo)),
            "vencedor_idx": 0 if propostas else None,  # padrão = menor preço; pode ser alterado
        })
    return resultado


def montar_declaracao_vencedores(mapa_itens):
    """Agrupa TODOS os itens do mesmo vencedor numa única linha (não só os
    consecutivos!) — por padrão o 1º colocado, mas respeita a escolha manual
    em item['vencedor_idx']. Os números dos itens saem em ordem crescente
    ('1, 3, 6' -> ATS), e cada empresa aparece uma única vez na tabela,
    ordenada pela primeira aparição, como no modelo de Adjudicação."""
    numeros_por_vencedor = {}
    ordem_primeira_aparicao = []
    for item in mapa_itens:
        if item.get("cancelado"):
            continue
        propostas = item["propostas"]
        if not propostas:
            continue
        idx = min(item.get("vencedor_idx", 0) or 0, len(propostas) - 1)
        vencedor = propostas[idx]["fornecedor"]
        numero = item["numero_item"]
        if vencedor not in numeros_por_vencedor:
            numeros_por_vencedor[vencedor] = []
            ordem_primeira_aparicao.append(vencedor)
        numeros_por_vencedor[vencedor].append(numero)

    grupos = []
    for vencedor in ordem_primeira_aparicao:
        numeros_ordenados = sorted(numeros_por_vencedor[vencedor],
                                   key=lambda n: (n is None, n))
        grupos.append((numeros_ordenados, vencedor))
    return grupos


# ----------------------------------------------------------------------
# GERAÇÃO DO PDF — MAPA COMPARATIVO / SOLICITAÇÃO DE ADJUDICAÇÃO
# ----------------------------------------------------------------------
def gerar_pdf_mapa_comparativo(caminho_pdf, codigo_cotacao, ano, mapa_itens, objeto_titulo="",
                                observacoes="", diretor_nome="", diretor_matricula="",
                                sei_publicacao="", tipo_planilha="HRG",
                                publicacao_rotulo="1ª PUBLICAÇÃO"):
    """Gera o Mapa Comparativo / Solicitação de Adjudicação replicando o
    layout usado pela SES/PDPAS (cabeçalho, barra de título, tabela de
    1º/2º/3º colocado por item com saldo remanescente e descrição, tabela de
    empresas vencedoras e o texto de adjudicação/homologação). Formato
    retrato, igual ao modelo original (uma única página).

    As cores das barras de título seguem o tipo de planilha (extraídas dos
    modelos reais): APS = verde claro; APS INV = verde claro + amarelo;
    HRG = azul claro; HRG INV = azul claro + amarelo. Cada item pode ter um
    'vencedor_idx' (0=1º colocado, 1=2º, 2=3º) definido manualmente — o
    colocado escolhido recebe um destaque cinza na tabela, e é ele (não
    necessariamente o de menor preço) que aparece na tabela de empresas
    vencedoras."""
    is_hrg = "HRG" in tipo_planilha.upper()
    is_inv = "INV" in tipo_planilha.upper()
    cor_base = COR_CIANO_ADJ if is_hrg else COR_VERDE_ADJ
    sufixo_letterhead = "/HRG" if is_hrg else ""
    sigla_comissao = "HRG" if is_hrg else "SRSSU"

    doc = SimpleDocTemplate(caminho_pdf, pagesize=A4,
                             topMargin=1 * cm, bottomMargin=1 * cm,
                             leftMargin=1.2 * cm, rightMargin=1.2 * cm)
    styles = getSampleStyleSheet()
    largura_util = A4[0] - 2.4 * cm

    estilo_barra = ParagraphStyle("Barra", parent=styles["Normal"], fontSize=12,
                                  alignment=TA_CENTER, fontName="Helvetica-Bold")
    estilo_objeto = ParagraphStyle("Objeto", parent=styles["Normal"], fontSize=10.5,
                                   alignment=TA_CENTER, fontName="Helvetica-Bold", leading=13)
    estilo_mapa_titulo = ParagraphStyle("MapaTitulo", parent=styles["Normal"], fontSize=14,
                                        alignment=TA_CENTER, fontName="Helvetica-Bold")
    estilo_cel_item = ParagraphStyle("CelItem", parent=styles["Normal"], fontSize=8.5,
                                     alignment=TA_CENTER, fontName="Helvetica-Bold", leading=10)
    estilo_cel_normal = ParagraphStyle("CelNormal", parent=styles["Normal"], fontSize=7.5,
                                       alignment=TA_CENTER, leading=9)
    estilo_cel_bold = ParagraphStyle("CelBold", parent=styles["Normal"], fontSize=7.5,
                                     alignment=TA_CENTER, fontName="Helvetica-Bold", leading=9)
    estilo_desc = ParagraphStyle("Desc", parent=styles["Normal"], fontSize=7.5, leading=9,
                                 fontName="Helvetica-Bold", alignment=TA_LEFT)
    estilo_corpo = ParagraphStyle("Corpo", parent=styles["Normal"], fontSize=9, leading=12,
                                  alignment=TA_JUSTIFY, firstLineIndent=26)
    estilo_corpo_centro = ParagraphStyle("CorpoCentro", parent=estilo_corpo, alignment=TA_CENTER,
                                         firstLineIndent=0)
    estilo_sei = ParagraphStyle("Sei", parent=styles["Normal"], fontSize=9, leading=12,
                                textColor=colors.red, fontName="Helvetica-Bold", alignment=TA_CENTER)
    estilo_rodape = ParagraphStyle("Rodape", parent=styles["Normal"], fontSize=7.5, leading=9.5,
                                   alignment=TA_CENTER, fontName="Helvetica-Bold")

    story = [
        montar_cabecalho_institucional(largura_util, [
            "GOVERNO DO DISTRITO FEDERAL",
            "SECRETARIA DE ESTADO DE SAÚDE",
            "SUPERINTENDÊNCIA DA REGIÃO DE SAÚDE SUL",
            "DIRETORIA ADMINISTRATIVA",
            f"PROGRAMA DE DESCENTRALIZAÇÃO PROGRESSIVA DE AÇÕES DE SAÚDE - PDPAS{sufixo_letterhead}",
        ]),
        Spacer(1, 8),
    ]

    barra1 = Table([[Paragraph("SOLICITAÇÃO DE ADJUDICAÇÃO PARA O PEDIDO DE COTAÇÃO PDPAS",
                               estilo_barra)]], colWidths=[largura_util])
    barra1.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), cor_base),
                                ("BOX", (0, 0), (-1, -1), 0.5, colors.black),
                                ("TOPPADDING", (0, 0), (-1, -1), 3),
                                ("BOTTOMPADDING", (0, 0), (-1, -1), 3)]))
    story.append(barra1)

    # A segunda linha (cotação/ano - publicação) só ganha fundo amarelo nas
    # variantes "...INV"; nas variantes normais (APS/HRG) fica sem destaque,
    # igual aos modelos reais.
    texto_barra2 = f"{codigo_cotacao}/{ano} – {publicacao_rotulo}"
    if is_inv:
        barra2 = Table([[Paragraph(texto_barra2, estilo_barra)]], colWidths=[largura_util])
        barra2.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), COR_AMARELO_ADJ),
                                    ("TOPPADDING", (0, 0), (-1, -1), 3),
                                    ("BOTTOMPADDING", (0, 0), (-1, -1), 3)]))
        story.append(barra2)
    else:
        story.append(Paragraph(texto_barra2, estilo_barra))
    story.append(Spacer(1, 8))

    if objeto_titulo:
        story.append(Paragraph(objeto_titulo, estilo_objeto))
        story.append(Spacer(1, 6))

    story.append(Paragraph(f"MAPA COMPARATIVO {codigo_cotacao}/{ano}", estilo_mapa_titulo))
    story.append(Spacer(1, 8))

    frac = [0.09, 0.10, 0.10, 0.09, 0.10, 0.10, 0.09, 0.10, 0.10, 0.09]
    col_widths = [largura_util * f for f in frac]

    linhas = [[
        Paragraph("Item /<br/>Código", estilo_cel_bold),
        Paragraph("1º Colocado /<br/>Unitário / Total", estilo_cel_bold), "",
        Paragraph("Qtd./Und./<br/>Marca", estilo_cel_bold),
        Paragraph("2º Colocado /<br/>Unitário / Total", estilo_cel_bold), "",
        Paragraph("Qtd./Und./<br/>Marca", estilo_cel_bold),
        Paragraph("3º Colocado /<br/>Unitário / Total", estilo_cel_bold), "",
        Paragraph("Qtd./Und./<br/>Marca", estilo_cel_bold),
    ]]
    estilos_extra = [
        ("SPAN", (1, 0), (2, 0)), ("SPAN", (4, 0), (5, 0)), ("SPAN", (7, 0), (8, 0)),
        ("BACKGROUND", (0, 0), (-1, 0), COR_CINZA_ADJ),
    ]

    COR_VENCEDOR_ADJ = colors.HexColor("#D9D9D9")   # cinza claro — destaque do vencedor
    CORES_ALTERNADAS = [colors.HexColor("#D9D9D9"), colors.HexColor("#A5A5A5")]  # cinza claro/escuro

    linha_idx = 1
    for indice_item, item in enumerate(mapa_itens):
        cancelado = bool(item.get("cancelado"))
        propostas = [] if cancelado else item["propostas"]
        top3 = propostas[:3]
        cor_bloco = CORES_ALTERNADAS[indice_item % 2]
        linha_topo_bloco = linha_idx

        # linha grossa separando este bloco do item anterior (exceto no 1º)
        if indice_item > 0:
            estilos_extra.append(("LINEABOVE", (0, linha_topo_bloco), (-1, linha_topo_bloco),
                                  1.6, colors.black))

        # ---- linha 1: nome do colocado / quantidade ----
        linha1 = [Paragraph(str(item["numero_item"] or ""), estilo_cel_item)]
        for pos in range(3):
            if pos < len(top3):
                p = top3[pos]
                linha1 += [Paragraph(p["fornecedor"], estilo_cel_bold), "",
                           Paragraph(str(p["quantidade"] or ""), estilo_cel_normal)]
            else:
                if cancelado:
                    texto_vazio = "CANCELADO"
                else:
                    texto_vazio = "DESERTO" if not item["propostas"] else "PROPOSTA NÃO APRESENTADA"
                linha1 += [Paragraph(texto_vazio, estilo_cel_bold), "", ""]
        linhas.append(linha1)
        estilos_extra += [
            ("SPAN", (1, linha_idx), (2, linha_idx)), ("SPAN", (4, linha_idx), (5, linha_idx)),
            ("SPAN", (7, linha_idx), (8, linha_idx)),
            ("BACKGROUND", (0, linha_idx), (0, linha_idx), cor_bloco),
        ]
        linha_idx += 1

        # ---- linha 2: unitário / total / UN ----
        linha2 = [""]
        for pos in range(3):
            if pos < len(top3):
                p = top3[pos]
                linha2 += [Paragraph(formatar_moeda(p["unitario"]), estilo_cel_normal),
                           Paragraph(formatar_moeda(p["total"]), estilo_cel_normal),
                           Paragraph(item["un"], estilo_cel_normal)]
            else:
                linha2 += ["", "", ""]
        linhas.append(linha2)
        estilos_extra.append(("BACKGROUND", (0, linha_idx), (0, linha_idx), cor_bloco))
        linha_idx += 1

        # ---- linha 3: código / prazo / marca ----
        linha3 = [Paragraph(item["codigo"], estilo_cel_bold)]
        for pos in range(3):
            linha3.append(Paragraph("Prazo (dias)", estilo_cel_bold))
            if pos < len(top3):
                p = top3[pos]
                linha3 += [Paragraph(str(p["prazo"] or ""), estilo_cel_normal),
                           Paragraph(p["marca"] or "", estilo_cel_normal)]
            else:
                linha3 += ["", ""]
        linhas.append(linha3)
        estilos_extra += [
            ("BACKGROUND", (0, linha_idx), (0, linha_idx), cor_bloco),
            ("BACKGROUND", (1, linha_idx), (1, linha_idx), cor_bloco),
            ("BACKGROUND", (4, linha_idx), (4, linha_idx), cor_bloco),
            ("BACKGROUND", (7, linha_idx), (7, linha_idx), cor_bloco),
        ]
        linha_idx += 1

        # ---- linha 4: saldo remanescente / descrição (1 linha só) ----
        saldo_texto = formatar_moeda(item["saldo"]) if item["saldo"] is not None else "-"
        descricao_1_linha = truncar_em_palavra((item["descricao"] or "").strip(), 60)
        if descricao_1_linha and len(descricao_1_linha) < len((item["descricao"] or "").strip()):
            descricao_1_linha += "..."
        linha4 = [Paragraph("Saldo remanescente", estilo_cel_bold), "", "",
                  Paragraph(saldo_texto, estilo_cel_bold),
                  Paragraph(descricao_1_linha, estilo_desc), "", "", "", "", ""]
        linhas.append(linha4)
        estilos_extra += [
            ("SPAN", (0, linha_idx), (2, linha_idx)), ("SPAN", (4, linha_idx), (9, linha_idx)),
            ("BACKGROUND", (0, linha_idx), (2, linha_idx), cor_bloco),
        ]

        # ---- destaque do vencedor (não é sempre o 1º colocado — pode ter
        # sido alterado manualmente, ex.: desistência do menor preço) ----
        vencedor_pos = item.get("vencedor_idx")
        if vencedor_pos is not None and vencedor_pos < len(top3):
            col_inicio = 1 + 3 * vencedor_pos
            estilos_extra.append(
                ("BACKGROUND", (col_inicio, linha_topo_bloco), (col_inicio + 2, linha_idx - 1),
                 COR_VENCEDOR_ADJ))

        linha_idx += 1

    tabela = Table(linhas, colWidths=col_widths)
    tabela.setStyle(TableStyle(estilos_extra + [
        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 2), ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]))
    story.append(tabela)
    story.append(Spacer(1, 10))

    story.append(Paragraph(
        "Em função do que se apresenta na planilha acima, das necessidades expostas no Termo de "
        "Referência e que o(s) produto(s) ofertado(s) atende(m) ao(s) requisito(s) de melhor "
        "proposta, prevista no edital referente à presente cotação, solicitamos de V.Sª o respaldo "
        "administrativo para declarar vencedora, por terem apresentado o menor preço e/ou menor "
        "prazo de entrega, a(s) seguinte(s) empresa(s) para o(s) respectivo(s) item(ns):",
        estilo_corpo))
    story.append(Spacer(1, 6))

    grupos_vencedores = montar_declaracao_vencedores(mapa_itens)
    if grupos_vencedores:
        dados_decl = [["Item(ns)", "Empresa"]]
        for numeros, vencedor in grupos_vencedores:
            dados_decl.append([", ".join(str(n) for n in numeros), vencedor])
        tabela_decl = Table(dados_decl, colWidths=[largura_util * 0.15, largura_util * 0.85])
        tabela_decl.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), COR_CINZA_ADJ),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
            ("FONTSIZE", (0, 0), (-1, -1), 9),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ]))
        story.append(tabela_decl)
        story.append(Spacer(1, 10))

    if observacoes:
        for paragrafo in observacoes.split("\n\n"):
            paragrafo = paragrafo.strip()
            if paragrafo:
                story.append(Paragraph(paragrafo, estilo_corpo))
                story.append(Spacer(1, 4))
        story.append(Spacer(1, 2))

    story.append(Paragraph(
        "Aceita essa solicitação, procederemos à etapa do envio de Ordem de Compra do material "
        "cotado.", estilo_corpo))
    story.append(Spacer(1, 8))
    story.append(Paragraph("Atenciosamente,", estilo_corpo_centro))
    story.append(Paragraph(f"COMISSÃO DO PDPAS/{sigla_comissao}.", estilo_corpo_centro))
    story.append(Spacer(1, 12))

    story.append(Paragraph("ADJUDICAÇÃO E HOMOLOGAÇÃO", estilo_mapa_titulo))
    story.append(Spacer(1, 4))
    story.append(Paragraph("De acordo, em função do exposto acima.", estilo_corpo_centro))
    story.append(Spacer(1, 6))
    story.append(Paragraph(
        "<b>Adjudico e Homologo</b> o presente. Proceda-se à desaplicação dos recursos necessários "
        "para fazer face ao pagamento da presente aquisição, assim que os materiais forem entregues "
        "e a NOTA FISCAL for atestada.", estilo_corpo))
    story.append(Spacer(1, 6))
    story.append(Paragraph(
        "Em tempo, solicitamos envio ao <b>NPDOC</b> para acompanhamento do prazo de entrega.",
        estilo_corpo))
    story.append(Spacer(1, 10))
    story.append(Paragraph(
        "Atendendo ao art. 16. Inciso IV da Portaria SES 473, de 04/12/2023 Declaro que há "
        "<b>disponibilidade financeira</b> para cobrir a presente contratação.", estilo_corpo_centro))
    if diretor_nome:
        story.append(Paragraph(
            f"Diretor(a) Administrativo(a)/SRSSU: {diretor_nome}, Matrícula: {diretor_matricula}",
            estilo_corpo_centro))
    story.append(Spacer(1, 10))
    if sei_publicacao:
        story.append(Paragraph(
            f"A Publicação no Diário Oficial do DF será realizada pelo SEI {sei_publicacao}",
            estilo_sei))

    story.append(Spacer(1, 14))
    if is_hrg:
        linha_rodape2 = "PDPAS/HRG - GEOF / Diretoria Administrativa / SRSSU / SES"
        linha_rodape3 = "Área Especial Nº 01, Lote Único - Setor Central GAMA/DF. CEP: 72.405-901"
        linha_rodape4 = "CNPJ: 00.394.700/0006-12 - Fone: (61) 3449-7105"
    else:
        linha_rodape2 = "PDPAS / APS-SUL - GEOF / Diretoria Administrativa / SRSSU / SES"
        linha_rodape3 = ""
        linha_rodape4 = "CNPJ: 00.394.700/0051-77 - Fone: (61) 3449-7105"
    linhas_rodape = ["Brasília – Patrimônio Cultural da Humanidade", linha_rodape2]
    if linha_rodape3:
        linhas_rodape.append(linha_rodape3)
    linhas_rodape.append(linha_rodape4)
    story.append(Paragraph("<br/>".join(linhas_rodape), estilo_rodape))

    doc.build(story)


# ----------------------------------------------------------------------
# GERAÇÃO DO PDF — ORDEM DE COMPRA
# ----------------------------------------------------------------------
def gerar_pdf_ordem_compra(caminho_pdf, codigo_cotacao, ano, tr_numero, origem, fornecedor_info,
                            itens_vencidos, tipo_planilha, sei_texto, diretor_nome, diretor_matricula,
                            contatos=None):
    """Gera a ORDEM DE COMPRA de um fornecedor vencedor, com os itens que
    ele ganhou. Usa o mesmo esquema de cores da Adjudicação (APS=verde,
    HRG=azul/ciano, "...INV"=+amarelo), mas no Ordem a cor amarela fica na
    PRIMEIRA barra (ao contrário da Adjudicação), conforme os modelos reais.

    'contatos' é o dict configurável (ver CONTATOS_ORDEM_PADRAO) com os
    nomes/telefones do GERENTE, NFL, NUAL, NECFM, NPDOC e NFH — o bloco de
    "Informações técnicas/Aceite NF" muda de acordo com o tipo_planilha
    exato (APS / APS INV / HRG / HRG INV), cada um com seu próprio
    conjunto de contatos."""
    contatos = contatos or CONTATOS_ORDEM_PADRAO
    tipo_norm = tipo_planilha.strip().upper()
    is_hrg = "HRG" in tipo_norm
    is_inv = "INV" in tipo_norm
    cor_base = COR_CIANO_ADJ if is_hrg else COR_VERDE_ADJ
    cor_barra1 = COR_AMARELO_ADJ if is_inv else cor_base
    cor_barra2 = cor_base

    doc = SimpleDocTemplate(caminho_pdf, pagesize=A4,
                             topMargin=1 * cm, bottomMargin=1 * cm,
                             leftMargin=1.2 * cm, rightMargin=1.2 * cm)
    styles = getSampleStyleSheet()
    largura_util = A4[0] - 2.4 * cm

    estilo_barra = ParagraphStyle("BarraOrd", parent=styles["Normal"], fontSize=12,
                                  alignment=TA_CENTER, fontName="Helvetica-Bold")
    estilo_valor = ParagraphStyle("ValorOrd", parent=styles["Normal"], fontSize=9, leading=12)
    estilo_cel = ParagraphStyle("CelOrd", parent=styles["Normal"], fontSize=8.5, leading=10.5)
    estilo_cel_center = ParagraphStyle("CelCenterOrd", parent=estilo_cel, alignment=TA_CENTER)
    estilo_corpo = ParagraphStyle("CorpoOrd", parent=styles["Normal"], fontSize=8.5, leading=11)
    estilo_corpo_centro = ParagraphStyle("CorpoCentroOrd", parent=estilo_corpo, alignment=TA_CENTER)
    estilo_corpo_centro_bold = ParagraphStyle("CorpoCentroBoldOrd", parent=estilo_corpo_centro,
                                              fontName="Helvetica-Bold")
    estilo_rodape = ParagraphStyle("RodapeOrd", parent=styles["Normal"], fontSize=7.5, leading=9.5,
                                   alignment=TA_CENTER, fontName="Helvetica-Bold")

    story = [
        montar_cabecalho_institucional(largura_util, [
            "GOVERNO DO DISTRITO FEDERAL",
            "SECRETARIA DE ESTADO DE SAÚDE",
            "SUPERINTENDÊNCIA DA REGIÃO DE SAÚDE SUL",
            "DIRETORIA ADMINISTRATIVA",
            "GERÊNCIA DE ORÇAMENTO E FINANÇAS/SRSSU - PDPAS",
        ]),
        Spacer(1, 8),
    ]

    barra1 = Table([[Paragraph(f"ORDEM DE COMPRA - COTAÇÃO {codigo_cotacao}/{ano}", estilo_barra)]],
                    colWidths=[largura_util])
    barra1.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), cor_barra1),
                                ("BOX", (0, 0), (-1, -1), 0.5, colors.black),
                                ("TOPPADDING", (0, 0), (-1, -1), 3),
                                ("BOTTOMPADDING", (0, 0), (-1, -1), 3)]))
    story.append(barra1)

    # O TR já vem como "033/2026" (com o ano embutido) direto da aba Controle,
    # então só concatena "/{ano}" se ele ainda não tiver uma barra.
    if tr_numero and "/" in str(tr_numero):
        texto_barra2 = f"TR Nº {tr_numero} - {origem}"
    elif tr_numero:
        texto_barra2 = f"TR Nº {tr_numero}/{ano} - {origem}"
    else:
        texto_barra2 = origem or ""
    if is_inv:
        barra2 = Table([[Paragraph(texto_barra2, estilo_barra)]], colWidths=[largura_util])
        barra2.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), cor_barra2),
                                    ("TOPPADDING", (0, 0), (-1, -1), 3),
                                    ("BOTTOMPADDING", (0, 0), (-1, -1), 3)]))
        story.append(barra2)
    else:
        story.append(Paragraph(texto_barra2, estilo_barra))
    story.append(Spacer(1, 10))

    # ---- bloco com dados do fornecedor ----
    linhas_forn = [
        [Paragraph(f"<b>FORNECEDOR:</b> {fornecedor_info.get('razao_social') or fornecedor_info.get('nome', '')}",
                   estilo_valor),
         Paragraph(f"<b>CÓDIGO SES:</b> {fornecedor_info.get('codigo_ses', '')}", estilo_valor)],
        [Paragraph(f"<b>CONTATO:</b> {fornecedor_info.get('contato', '')}", estilo_valor),
         Paragraph(f"<b>SEI:</b> {sei_texto or ''}", estilo_valor)],
        [Paragraph(f"<b>FONE:</b> {fornecedor_info.get('telefone', '')}", estilo_valor), ""],
        [Paragraph(f"<b>CNPJ:</b> {fornecedor_info.get('cnpj', '')}", estilo_valor), ""],
        [Paragraph(f"<b>BANCO:</b> {fornecedor_info.get('banco', '')}", estilo_valor),
         Paragraph(f"<b>AGÊNCIA:</b> {fornecedor_info.get('agencia', '')} &nbsp;&nbsp; "
                   f"<b>CONTA:</b> {fornecedor_info.get('conta', '')}", estilo_valor)],
    ]
    tabela_forn = Table(linhas_forn, colWidths=[largura_util * 0.62, largura_util * 0.38])
    tabela_forn.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING", (0, 0), (-1, -1), 2), ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]))
    story.append(tabela_forn)
    story.append(Spacer(1, 8))

    story.append(Paragraph(
        "Solicitamos os seguintes materiais a serem entregues conforme cotação. A compra só será "
        "reconhecida com o atesto da Nota Fiscal, que deverá acompanhar os materiais abaixo "
        "solicitados:", estilo_corpo))
    story.append(Spacer(1, 6))

    # ---- tabela de itens vencidos por este fornecedor ----
    dados_tabela = [["Item", "Cod.", "Descrição", "Qtd.", "Und.", "Valor Unitário",
                     "Valor Total", "Prazo de Entrega"]]
    total_geral = 0
    for item, proposta in itens_vencidos:
        dados_tabela.append([
            str(item["numero_item"] or ""), item["codigo"],
            Paragraph(item["descricao"] or "", estilo_cel),
            str(proposta["quantidade"] or ""), item["un"],
            formatar_moeda(proposta["unitario"]), formatar_moeda(proposta["total"]),
            f'{proposta["prazo"]} dias úteis.' if proposta["prazo"] else "",
        ])
        total_geral += proposta["total"] or 0

    tabela_itens = Table(
        dados_tabela,
        colWidths=[1.2 * cm, 1.8 * cm, largura_util - 12.4 * cm, 1.3 * cm, 1.3 * cm,
                   2.4 * cm, 2.4 * cm, 2.4 * cm],
        repeatRows=1)
    tabela_itens.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#D9D9D9")),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
        ("FONTSIZE", (0, 0), (-1, -1), 8.5),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (0, 0), (1, -1), "CENTER"), ("ALIGN", (3, 0), (-1, -1), "CENTER"),
    ]))
    story.append(tabela_itens)

    tabela_total = Table([["TOTAL", "", formatar_moeda(total_geral)]],
                         colWidths=[3 * cm, largura_util - 3 * cm - 3.5 * cm, 3.5 * cm])
    tabela_total.setStyle(TableStyle([
        ("SPAN", (0, 0), (1, 0)),
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#D9D9D9")),
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("ALIGN", (0, 0), (0, 0), "CENTER"), ("ALIGN", (2, 0), (2, 0), "CENTER"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
    ]))
    story.append(tabela_total)
    story.append(Spacer(1, 10))

    # ---- dados para nota fiscal / entrega ----
    barra_nf = Table([[Paragraph("Dados para emissão de Nota Fiscal e local para entrega:",
                                 estilo_barra)]], colWidths=[largura_util])
    barra_nf.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), cor_barra1),
                                  ("TOPPADDING", (0, 0), (-1, -1), 3),
                                  ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
                                  ("ALIGN", (0, 0), (-1, -1), "LEFT"),
                                  ("LEFTPADDING", (0, 0), (-1, -1), 6)]))
    story.append(barra_nf)
    story.append(Spacer(1, 4))

    if is_hrg:
        nome_unidade = "HOSPITAL REGIONAL DO GAMA - DF"
        cnpj_unidade = "00.394.700/0006-12"
        insc_municipal = "0732950500936"
        endereco_unidade = "ÁREA ESPECIAL Nº 01 - Setor Central GAMA/DF. CEP: 72.405-901"
    else:
        nome_unidade = "SUPERINTENDENCIA DA REGIAO DE SAUDE SUL - APS"
        cnpj_unidade = "00.394.700/0051-77"
        insc_municipal = "0732950502475"
        endereco_unidade = ("GAOAPS-SUL no Setor Oeste, SOE Q. 12, Praça 1, Área Especial - "
                            "GAMA OESTE, CEP: 72.425-120.")

    story.append(Paragraph(f"<b>{nome_unidade}</b>", estilo_corpo))
    story.append(Paragraph(f"<b>CNPJ:</b> {cnpj_unidade} &nbsp;&nbsp; "
                           f"<b>Inscrição Municipal:</b> {insc_municipal}", estilo_corpo))
    story.append(Paragraph(f"<b>Endereço:</b> {endereco_unidade}", estilo_corpo))
    story.append(Spacer(1, 6))

    if not is_hrg:
        story.append(Paragraph("***LOCAL DE ENTREGA***", estilo_corpo_centro_bold))
        story.append(Paragraph(endereco_unidade, estilo_corpo_centro))
        story.append(Paragraph("Ao lado da Igreja Nossa Senhora Aparecida. Coordenadas no Maps "
                               "-16.007452594187583, -48.075027885099104", estilo_corpo_centro))
        story.append(Paragraph("Plus Code do Google Maps XWVF+2X9 Gama, Brasília - DF",
                               estilo_corpo_centro))
        story.append(Spacer(1, 6))

    story.append(Paragraph(
        "Os pagamentos da Região Sul são efetuados por TED, entre os dias 20 e 30 de cada mês. "
        "NÃO EMITIR BOLETO!", estilo_corpo_centro_bold))
    story.append(Spacer(1, 4))

    if tipo_norm == "APS":
        story.append(Paragraph(
            f"<b>Informações técnicas/Aceite NF:</b> {contatos.get('GERENTE', '')} - GERENTE;",
            estilo_corpo))
        story.append(Paragraph(f"Farmácia-NFL - {contatos.get('NFL', '')}", estilo_corpo))
        story.append(Paragraph(f"Almoxarifado-NUAL - {contatos.get('NUAL_APS', '')}", estilo_corpo))
        story.append(Paragraph(
            f"ENGENHARIA CLINICA/NECFM - {contatos.get('NECFM', '')}", estilo_corpo))
    elif tipo_norm == "APS INV":
        story.append(Paragraph(
            f"<b>Informações técnicas/Aceite NF:</b> {contatos.get('GERENTE', '')} - GERENTE;",
            estilo_corpo))
        story.append(Paragraph("Entrega no Núcleo de Patrimônio (NPDOC/APS) de segunda a QUINTA de "
                               "08:00 as 11:00 e 14:00 as 17:00", estilo_corpo))
        story.append(Paragraph(f"Chefe do NPDOC - {contatos.get('NPDOC_APSINV', '')}", estilo_corpo))
    elif tipo_norm == "HRG":
        story.append(Paragraph("<b>Informações técnicas/Aceite NF:</b>", estilo_corpo))
        story.append(Paragraph(f"Farmácia/NFH - {contatos.get('NFH', '')}", estilo_corpo))
        story.append(Paragraph(f"Almoxarifado/NUAL - {contatos.get('NUAL_HRG', '')}", estilo_corpo))
    else:  # HRG INV — atenção: aqui NÃO entra NECFM, só o Chefe do NPDOC
        story.append(Paragraph("<b>Informações técnicas/Aceite NF:</b>", estilo_corpo))
        story.append(Paragraph("Entrega no Núcleo de Patrimônio (NPDOC/HRG) de segunda a QUINTA de "
                               "08:00 as 11:00 e 14:00 as 17:00", estilo_corpo))
        story.append(Paragraph(f"Chefe do NPDOC - {contatos.get('NPDOC_HRGINV', '')}", estilo_corpo))
    story.append(Spacer(1, 8))

    story.append(Paragraph(
        "*Favor discriminar na Nota o valor UNITÁRIO e a QUANTIDADE de cada item, conforme esta "
        "Ordem de Compra.", estilo_corpo))
    story.append(Paragraph(
        "*Nota Fiscal empresa Não enquadrada Simples Nacional terá Imposto de Renda retido no "
        "percentual: 1,2% p/material; e 4,8% p/serviço, IN RFB nº 2145 26/06/2023", estilo_corpo))
    story.append(Paragraph(
        "*PAGAMENTO de empresas do DF, acima de R$ 5.000,00 somente em conta do BRB. Decreto "
        "Distrital Nº 32.767, de 17/02/2011 artigo 6º caput.", estilo_corpo))
    story.append(Paragraph(
        "*Caso seja necessário a instalação e/ou montagem deve ser agendada com o Chefe do "
        "Núcleo Citado acima", estilo_corpo))
    story.append(Spacer(1, 14))

    if diretor_nome:
        story.append(Paragraph(diretor_nome, estilo_corpo_centro_bold))
        story.append(Paragraph("Diretor(a) Administrativo(a)/SRSSU", estilo_corpo_centro))
        story.append(Paragraph(f"Matrícula: {diretor_matricula}", estilo_corpo_centro))

    story.append(Spacer(1, 12))
    story.append(Paragraph(
        "Brasília-Patrimônio Cultural da Humanidade<br/>"
        "GEOF / Diretoria Administrativa / SRSSU / SES<br/>"
        "AREA ESPECIAL Nº 01 - Setor Central GAMA/DF. CEP: 72.405-901<br/>"
        "Contato PREFERENCIALMENTE por e-mails, FoneVOIP: (61)3449-7105", estilo_rodape))

    doc.build(story)


# ----------------------------------------------------------------------
# MIXIN — métodos de ação ligados aos botões do 3º Passo
# ----------------------------------------------------------------------
class AdjudicacaoOrdemMixin:
    """Métodos herdados por AppCotacao. Usa o cadastro de propostas
    (self.mapa_fornecedores / arquivo MAPA_EMPRESAS_*.json) preparado no
    2º Passo (Mapa.py) para montar o Mapa Comparativo e as Ordens de Compra."""

    def acao_adjudicacao(self):
        if not self.linhas_resultado:
            messagebox.showwarning(
                "Atenção",
                "Nenhum item de cotação carregado. Clique em CONSULTAR COTAÇÃO antes.")
            return

        contexto = self._preparar_contexto_wizard()
        if contexto is None:
            return
        pasta_propostas, codigo_cotacao, ano = contexto

        caminho_json = self._caminho_arquivo_mapa_empresas(self._pasta_cotacao_atual, codigo_cotacao)
        if not os.path.exists(caminho_json):
            messagebox.showwarning(
                "Nenhum cadastro encontrado",
                f"Não há propostas cadastradas em:\n{caminho_json}\n\n"
                "Use o botão INCLUIR EMPRESAS MAPA primeiro.")
            return
        try:
            with open(caminho_json, "r", encoding="utf-8") as f:
                dados_salvos = json.load(f)
        except Exception as e:
            messagebox.showerror("Erro ao carregar cadastro", str(e))
            return

        self.mapa_fornecedores = dados_salvos.get("fornecedores", {})
        if not self.mapa_fornecedores:
            messagebox.showwarning("Atenção", "O arquivo salvo não contém fornecedores.")
            return

        self._gerar_mapa_comparativo()

    def _montar_dados_ordens_ativas(self):
        """Monta os dados necessários para gerar as Ordens de Compra da
        cotação ativa — em PDF (acao_ordem, aqui) ou por e-mail
        (acao_email_ordem, em Emails.py): consulta o cadastro de propostas
        já salvo, aplica as decisões de vencedor/cancelamento definidas na
        Adjudicação e agrupa os itens vencidos por fornecedor.

        Devolve um dict com 'ordens_por_fornecedor' (nome_normalizado ->
        (texto_completo, [(item, proposta), ...])), 'fornecedores_completo',
        'codigo_cotacao', 'ano', 'tr_numero', 'origem' e 'sei_texto' — ou
        None se algo impedir a montagem (já mostrando o aviso/erro
        correspondente nesse caso)."""
        if not self.linhas_resultado:
            messagebox.showwarning(
                "Atenção", "Nenhum item de cotação carregado. Clique em CONSULTAR COTAÇÃO antes.")
            return None

        contexto = self._preparar_contexto_wizard()
        if contexto is None:
            return None
        pasta_propostas, codigo_cotacao, ano = contexto

        caminho_json = self._caminho_arquivo_mapa_empresas(self._pasta_cotacao_atual, codigo_cotacao)
        if not os.path.exists(caminho_json):
            messagebox.showwarning(
                "Nenhum cadastro encontrado",
                f"Não há propostas cadastradas em:\n{caminho_json}\n\n"
                "Use INCLUIR EMPRESAS MAPA e depois ADJUDICAÇÃO primeiro.")
            return None
        try:
            with open(caminho_json, "r", encoding="utf-8") as f:
                dados_salvos = json.load(f)
        except Exception as e:
            messagebox.showerror("Erro ao carregar cadastro", str(e))
            return None

        mapa_fornecedores = dados_salvos.get("fornecedores", {})
        decisoes = dados_salvos.get("decisoes_itens", {})
        if not mapa_fornecedores:
            messagebox.showwarning("Atenção", "O arquivo salvo não contém fornecedores.")
            return None
        if not decisoes:
            messagebox.showwarning(
                "Atenção",
                "Ainda não há vencedores definidos para esta cotação.\n"
                "Clique em ADJUDICAÇÃO primeiro para escolher os vencedores de cada item.")
            return None

        try:
            fornecedores_completo = carregar_fornecedores_completo(self._obter_caminho_arquivo())
        except Exception as e:
            messagebox.showerror("Erro ao ler aba Fornecedores", str(e))
            return None
        fornecedores_texto = {chave: dados["texto_completo"]
                              for chave, dados in fornecedores_completo.items()}

        mapa_itens = montar_mapa_itens_completo(
            self.linhas_resultado, mapa_fornecedores, fornecedores_texto, self.saldos_bd_atual)

        # aplica as decisões (vencedor escolhido / cancelado) salvas na Adjudicação
        for item in mapa_itens:
            decisao = decisoes.get(str(item["numero_item"]))
            if decisao:
                item["vencedor_idx"] = decisao.get("vencedor_idx")
                item["cancelado"] = bool(decisao.get("cancelado"))

        # agrupa os itens vencidos por fornecedor (só o vencedor de cada item)
        ordens_por_fornecedor = {}   # nome_normalizado -> (texto_completo, [ (item, proposta), ... ])
        for item in mapa_itens:
            if item.get("cancelado"):
                continue
            propostas = item["propostas"]
            if not propostas:
                continue
            idx = min(item.get("vencedor_idx", 0) or 0, len(propostas) - 1)
            proposta_vencedora = propostas[idx]
            texto_vencedor = proposta_vencedora["fornecedor"]
            nome_vencedor, _ = extrair_nome_codigo_fornecedor(texto_vencedor)
            chave = normalizar_texto(nome_vencedor)
            if chave not in ordens_por_fornecedor:
                ordens_por_fornecedor[chave] = (texto_vencedor, [])
            ordens_por_fornecedor[chave][1].append((item, proposta_vencedora))

        if not ordens_por_fornecedor:
            messagebox.showinfo(
                "Nenhum vencedor", "Nenhum item tem vencedor definido (todos cancelados ou sem "
                                   "propostas).")
            return None

        tr_numero = self.linhas_resultado[0].get("TR") if self.linhas_resultado else ""
        origem = self.linhas_resultado[0].get("ORIGEM") if self.linhas_resultado else ""
        sei_texto = self.entry_sei.get().strip()

        return {
            "ordens_por_fornecedor": ordens_por_fornecedor,
            "fornecedores_completo": fornecedores_completo,
            "codigo_cotacao": codigo_cotacao,
            "ano": ano,
            "tr_numero": tr_numero,
            "origem": origem,
            "sei_texto": sei_texto,
        }

    def acao_ordem(self):
        dados = self._montar_dados_ordens_ativas()
        if dados is None:
            return
        ordens_por_fornecedor = dados["ordens_por_fornecedor"]
        fornecedores_completo = dados["fornecedores_completo"]
        codigo_cotacao = dados["codigo_cotacao"]
        ano = dados["ano"]
        tr_numero = dados["tr_numero"]
        origem = dados["origem"]
        sei_texto = dados["sei_texto"]
        diretor_nome = self.entry_da_nome.get().strip()
        diretor_matricula = self.entry_da_matricula.get().strip()

        arquivos_gerados = []
        for chave, (texto_vencedor, itens_vencidos) in ordens_por_fornecedor.items():
            info_fornecedor = fornecedores_completo.get(chave, {"nome": texto_vencedor})
            nome_arquivo_seguro = re_sub_nome_arquivo(info_fornecedor.get("nome", chave))
            base = f"ORDEM {codigo_cotacao} {nome_arquivo_seguro}"
            caminho_pdf = os.path.join(self._pasta_cotacao_atual, base + ".pdf")
            caminho_jpg_base = os.path.join(self._pasta_cotacao_atual, base)
            try:
                gerar_pdf_ordem_compra(
                    caminho_pdf, codigo_cotacao, ano, tr_numero, origem, info_fornecedor,
                    itens_vencidos, self._sufixo_planilha_atual, sei_texto,
                    diretor_nome, diretor_matricula,
                    contatos=self.config_dados.get("contatos_ordem", CONTATOS_ORDEM_PADRAO))
                caminho_jpg = converter_pdf_para_jpg(caminho_pdf, caminho_jpg_base)
                arquivos_gerados.append((info_fornecedor.get("nome", texto_vencedor),
                                         caminho_pdf, caminho_jpg))
            except Exception as e:
                messagebox.showerror("Erro ao gerar Ordem de Compra",
                                     f"Fornecedor {texto_vencedor}:\n{e}")
                return

        resumo = "\n".join(f"• {nome} → {os.path.basename(pdf)}"
                           for nome, pdf, _jpg in arquivos_gerados)
        messagebox.showinfo(
            "Ordens de Compra geradas",
            f"{len(arquivos_gerados)} ordem(ns) de compra gerada(s) em:\n"
            f"{self._pasta_cotacao_atual}\n\n{resumo}")

        # pré-visualiza a primeira ordem gerada
        if arquivos_gerados:
            _nome, _pdf, caminho_jpg = arquivos_gerados[0]
            exibir_preview_pdf_jpg(self, caminho_jpg, "Ordem de Compra")

    def acao_salvar_passo3(self):
        self._salvar_defaults_atuais()
        self._stub("SALVAR")

    # ==================================================================
    # PASSO 27 — Mapa Comparativo (tela + PDF + JPG)
    # ==================================================================
    def _gerar_mapa_comparativo(self):
        if not self.mapa_fornecedores:
            return
        fornecedores_texto = self._obter_fornecedores_texto_completo()
        mapa_itens = montar_mapa_itens_completo(
            self.linhas_resultado, self.mapa_fornecedores, fornecedores_texto, self.saldos_bd_atual)
        if not mapa_itens:
            return

        # 1) marca quais itens (se houver) devem aparecer como CANCELADO
        mapa_itens = self._perguntar_itens_cancelados(mapa_itens)

        # 2) confirma/ajusta manualmente quem é o vencedor de cada item
        #    (por padrão é o de menor preço, mas pode ter desistido etc.)
        mapa_itens = self._perguntar_vencedores(mapa_itens)

        # salva as decisões (vencedor/cancelado) para o botão ORDEM poder
        # reutilizá-las depois, sem precisar perguntar de novo
        self._salvar_decisoes_mapa(mapa_itens)

        # 3) monta os textos automáticos (objeto e observações padrão)
        objeto_auto = gerar_objeto_automatico(mapa_itens)
        itens_poucas_propostas = [str(it["numero_item"]) for it in mapa_itens
                                   if not it.get("cancelado") and 0 < len(it["propostas"]) < 3]
        observacoes_auto = (TEXTO_OBSERVACAO_PADRAO.format(itens=", ".join(itens_poucas_propostas))
                             if itens_poucas_propostas else "")

        dados_extra = self._perguntar_dados_finalizacao(objeto_auto, observacoes_auto)

        base = f"MAPA COMPARATIVO {self._codigo_cotacao_atual}"
        caminho_pdf = os.path.join(self._pasta_cotacao_atual, base + ".pdf")
        caminho_jpg_base = os.path.join(self._pasta_cotacao_atual, base)
        try:
            gerar_pdf_mapa_comparativo(
                caminho_pdf, self._codigo_cotacao_atual, self._ano_atual_wizard, mapa_itens,
                objeto_titulo=dados_extra.get("objeto", ""),
                observacoes=dados_extra.get("observacoes", ""),
                diretor_nome=self.entry_da_nome.get().strip(),
                diretor_matricula=self.entry_da_matricula.get().strip(),
                sei_publicacao=self.entry_sei.get().strip(),  # já vem preenchido na tela principal
                tipo_planilha=self._sufixo_planilha_atual)
            caminho_jpg = converter_pdf_para_jpg(caminho_pdf, caminho_jpg_base)
        except Exception as e:
            messagebox.showerror("Erro ao gerar Mapa Comparativo", str(e))
            return

        messagebox.showinfo(
            "Mapa Comparativo gerado",
            f"Arquivos salvos em:\n{caminho_pdf}\n{caminho_jpg}")

        exibir_preview_pdf_jpg(self, caminho_jpg, "Mapa Comparativo")

    def _perguntar_itens_cancelados(self, mapa_itens):
        """Tela simples para marcar quais itens devem aparecer como
        CANCELADO no Mapa Comparativo — mesmo que já tenham valores e
        fornecedores lançados, eles ficam suprimidos (igual ao DESERTO, mas
        com o texto 'CANCELADO' nas 3 colunas de colocado)."""
        janela = ctk.CTkToplevel(self)
        janela.title("Itens cancelados")
        janela.configure(fg_color=COR_FUNDO_APP)
        janela.geometry("900x600")
        janela.grab_set()

        ctk.CTkLabel(
            janela,
            text="Marque os itens que devem aparecer como CANCELADO no Mapa Comparativo "
                 "(os valores e fornecedores desse item ficam suprimidos, mesmo que já tenham "
                 "sido lançados). Deixe todos desmarcados se nenhum item foi cancelado.",
            font=FONTE_SECAO, text_color=COR_TEXTO, wraplength=850, justify="left"
        ).pack(anchor="w", padx=16, pady=(16, 8))

        scroll = ctk.CTkScrollableFrame(janela, fg_color="#161616", height=460)
        scroll.pack(fill="both", expand=True, padx=16, pady=(0, 10))

        variaveis = {}
        for item in mapa_itens:
            var = tk.BooleanVar(value=bool(item.get("cancelado", False)))
            variaveis[item["numero_item"]] = var
            texto = f"Item {item['numero_item']} — {item['codigo']} — {item['descricao'][:75]}"
            ctk.CTkCheckBox(scroll, text=texto, variable=var, font=FONTE_TABELA,
                            text_color=COR_TEXTO).pack(anchor="w", padx=10, pady=5)

        def confirmar():
            for item in mapa_itens:
                item["cancelado"] = variaveis[item["numero_item"]].get()
            janela.destroy()

        ctk.CTkButton(janela, text="Confirmar e continuar", font=FONTE_BOTAO, height=36,
                      fg_color=COR_DESTAQUE, command=confirmar).pack(pady=(0, 16))
        janela.protocol("WM_DELETE_WINDOW", confirmar)
        self.wait_window(janela)
        return mapa_itens

    def _perguntar_vencedores(self, mapa_itens):
        """Mostra os até 3 colocados de cada item e deixa o usuário confirmar
        (ou trocar) qual deles é o vencedor de fato — por padrão já vem
        marcado o de menor preço, mas pode ter desistido/sido desclassificado
        e o 2º ou 3º colocado assume."""
        itens_com_proposta = [it for it in mapa_itens if it["propostas"] and not it.get("cancelado")]
        if not itens_com_proposta:
            return mapa_itens

        janela = ctk.CTkToplevel(self)
        janela.title("Confirmar empresa vencedora por item")
        janela.configure(fg_color=COR_FUNDO_APP)
        janela.geometry("980x620")
        janela.grab_set()

        ctk.CTkLabel(
            janela,
            text="Por padrão, o vencedor é o de menor preço. Altere se necessário "
                 "(ex.: desistência — o 2º ou 3º colocado assume):",
            font=FONTE_SECAO, text_color=COR_TEXTO, wraplength=930, justify="left"
        ).pack(anchor="w", padx=16, pady=(16, 8))

        scroll = ctk.CTkScrollableFrame(janela, fg_color="#161616", height=460)
        scroll.pack(fill="both", expand=True, padx=16, pady=(0, 10))

        variaveis = {}
        for item in itens_com_proposta:
            frame_item = ctk.CTkFrame(scroll, fg_color="#1f1f1f", corner_radius=8)
            frame_item.pack(fill="x", pady=4, padx=2)
            ctk.CTkLabel(
                frame_item,
                text=f"Item {item['numero_item']} — {item['codigo']} — {item['descricao'][:70]}",
                font=("Segoe UI", 12, "bold"), text_color=COR_TEXTO, anchor="w"
            ).pack(anchor="w", padx=10, pady=(8, 4))

            var = tk.IntVar(value=item.get("vencedor_idx", 0) or 0)
            variaveis[item["numero_item"]] = var
            linha_opcoes = ctk.CTkFrame(frame_item, fg_color="transparent")
            linha_opcoes.pack(fill="x", padx=10, pady=(0, 8))
            for idx, p in enumerate(item["propostas"][:3]):
                texto_opcao = f"{idx + 1}º: {p['fornecedor']} — {formatar_moeda(p['unitario'])}"
                ctk.CTkRadioButton(linha_opcoes, text=texto_opcao, variable=var, value=idx,
                                   font=FONTE_TABELA, text_color=COR_TEXTO
                                   ).pack(side="left", padx=(0, 20))

        def confirmar():
            for item in itens_com_proposta:
                item["vencedor_idx"] = variaveis[item["numero_item"]].get()
            janela.destroy()

        ctk.CTkButton(janela, text="Confirmar vencedores e continuar", font=FONTE_BOTAO,
                      height=36, fg_color=COR_DESTAQUE, command=confirmar).pack(pady=(0, 16))
        janela.protocol("WM_DELETE_WINDOW", confirmar)
        self.wait_window(janela)
        return mapa_itens

    def _perguntar_dados_finalizacao(self, objeto_auto="", observacoes_auto=""):
        """Pergunta/confirma o objeto (título do mapa) e as observações —
        ambos já vêm pré-preenchidos automaticamente, mas continuam
        editáveis. O SEI de publicação é lido direto do campo SEI da tela
        principal, sem precisar perguntar de novo."""
        janela = ctk.CTkToplevel(self)
        janela.title("Dados do Mapa Comparativo")
        janela.configure(fg_color=COR_FUNDO_APP)
        janela.geometry("680x420")
        janela.grab_set()

        ctk.CTkLabel(janela, text="OBJETO da cotação (título do mapa):", font=FONTE_SECAO,
                     text_color=COR_TEXTO).pack(anchor="w", padx=16, pady=(16, 4))
        entry_objeto = ctk.CTkTextbox(janela, height=70, font=("Segoe UI", 14))
        entry_objeto.pack(fill="x", padx=16)
        if objeto_auto:
            entry_objeto.insert("1.0", objeto_auto)

        ctk.CTkLabel(janela, text="Observações (editável — já preenchido quando aplicável):",
                     font=FONTE_SECAO, text_color=COR_TEXTO).pack(anchor="w", padx=16, pady=(12, 4))
        entry_obs = ctk.CTkTextbox(janela, height=140, font=("Segoe UI", 14))
        entry_obs.pack(fill="x", padx=16)
        if observacoes_auto:
            entry_obs.insert("1.0", observacoes_auto)

        resultado = {}

        def confirmar():
            resultado["objeto"] = entry_objeto.get("1.0", "end").strip()
            resultado["observacoes"] = entry_obs.get("1.0", "end").strip()
            janela.destroy()

        ctk.CTkButton(janela, text="Gerar Mapa Comparativo", font=FONTE_BOTAO, height=36,
                      fg_color=COR_DESTAQUE, command=confirmar).pack(pady=18)
        janela.protocol("WM_DELETE_WINDOW", confirmar)
        self.wait_window(janela)
        return resultado

    def _salvar_decisoes_mapa(self, mapa_itens):
        """Grava no mesmo MAPA_EMPRESAS_{cotação}.json quem foi o vencedor
        escolhido (ou item cancelado) para cada item, para o botão ORDEM
        poder reaproveitar sem perguntar de novo."""
        caminho_json = self._caminho_arquivo_mapa_empresas(
            self._pasta_cotacao_atual, self._codigo_cotacao_atual)
        conteudo = {}
        if os.path.exists(caminho_json):
            try:
                with open(caminho_json, "r", encoding="utf-8") as f:
                    conteudo = json.load(f)
            except Exception:
                conteudo = {}
        conteudo["cotacao"] = self._codigo_cotacao_atual
        conteudo["ano"] = self._ano_atual_wizard
        conteudo["fornecedores"] = self.mapa_fornecedores
        conteudo["decisoes_itens"] = {
            str(item["numero_item"]): {
                "vencedor_idx": item.get("vencedor_idx"),
                "cancelado": bool(item.get("cancelado")),
            }
            for item in mapa_itens
        }
        try:
            with open(caminho_json, "w", encoding="utf-8") as f:
                json.dump(conteudo, f, ensure_ascii=False, indent=2)
        except Exception as e:
            messagebox.showerror("Erro ao salvar", f"Não foi possível salvar as decisões:\n{e}")


def re_sub_nome_arquivo(nome):
    """Remove caracteres inválidos para nome de arquivo do Windows."""
    import re
    return re.sub(r"[^\w\- ]", "", nome or "").strip()